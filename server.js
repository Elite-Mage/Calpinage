const express = require('express');
const http = require('http');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { execFile } = require('child_process');
const { WebSocketServer } = require('ws');
const Database = require('better-sqlite3');
const { v4: uuidv4 } = require('uuid');
const multer = require('multer');

const PORT = 3000;
const DB_PATH = path.join(__dirname, 'data.db');

// ─── DATABASE ────────────────────────────────────────────────────────────────
const db = new Database(DB_PATH);
db.pragma('journal_mode = WAL');
db.exec(`
  CREATE TABLE IF NOT EXISTS projects (
    id TEXT PRIMARY KEY,
    name TEXT NOT NULL DEFAULT 'Sans titre',
    data TEXT NOT NULL DEFAULT '{}',
    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
  );
`);

// ─── EXPRESS ─────────────────────────────────────────────────────────────────
const app = express();
app.use(express.json({ limit: '5mb' }));
app.use(express.static(__dirname));

// List projects
app.get('/api/projects', (req, res) => {
  const rows = db.prepare(
    'SELECT id, name, updated_at FROM projects ORDER BY updated_at DESC'
  ).all();
  res.json(rows);
});

// Create project
app.post('/api/projects', (req, res) => {
  const id = uuidv4();
  const name = req.body.name || 'Nouveau chantier';
  const data = JSON.stringify(req.body.data || {});
  db.prepare(
    "INSERT INTO projects (id, name, data, updated_at) VALUES (?, ?, ?, datetime('now'))"
  ).run(id, name, data);
  res.json({ id, name });
});

// Get project
app.get('/api/projects/:id', (req, res) => {
  const row = db.prepare('SELECT * FROM projects WHERE id = ?').get(req.params.id);
  if (!row) return res.status(404).json({ error: 'Not found' });
  res.json({ ...row, data: JSON.parse(row.data) });
});

// Save project (full replace)
app.put('/api/projects/:id', (req, res) => {
  const name = req.body.name || 'Sans titre';
  const data = JSON.stringify(req.body.data || {});
  const info = db.prepare(
    "UPDATE projects SET name=?, data=?, updated_at=datetime('now') WHERE id=?"
  ).run(name, data, req.params.id);
  if (info.changes === 0) return res.status(404).json({ error: 'Not found' });
  // Broadcast to all WS clients watching this project
  broadcastToProject(req.params.id, { type: 'saved', id: req.params.id, name });
  res.json({ ok: true });
});

// Delete project
app.delete('/api/projects/:id', (req, res) => {
  db.prepare('DELETE FROM projects WHERE id = ?').run(req.params.id);
  res.json({ ok: true });
});

// ─── IMPORT DXF / DWG ────────────────────────────────────────────────────────
const upload = multer({
  dest: os.tmpdir(),
  limits: { fileSize: 20 * 1024 * 1024 }, // 20 Mo max
  fileFilter: (_req, file, cb) => {
    const ok = /\.(dxf|dwg)$/i.test(file.originalname);
    cb(ok ? null : new Error('Seuls les fichiers .dxf et .dwg sont acceptés'), ok);
  },
});

// Dossier temporaire pour les fichiers Excel générés
const EXCEL_TMP = path.join(os.tmpdir(), 'calpinage-excel');
if (!fs.existsSync(EXCEL_TMP)) fs.mkdirSync(EXCEL_TMP, { recursive: true });

app.post('/api/parse-dxf', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Aucun fichier reçu' });

  // Multer supprime l'extension → on renomme le fichier temporaire avec l'extension originale
  const ext = path.extname(req.file.originalname).toLowerCase(); // .dxf ou .dwg
  const tmpIn = req.file.path + ext;
  fs.renameSync(req.file.path, tmpIn);

  const script = path.join(__dirname, 'parse_dxf.py');
  const baseName = path.basename(req.file.originalname, ext);
  const excelId = uuidv4();
  const excelPath = path.join(EXCEL_TMP, excelId + '.xlsx');

  execFile('python3', [script, tmpIn, '--excel', excelPath], { timeout: 60000 }, (err, stdout, stderr) => {
    fs.unlink(tmpIn, () => {});
    if (err) {
      console.error('[parse-dxf]', stderr);
      return res.status(500).json({ error: stderr || err.message });
    }
    try {
      const data = JSON.parse(stdout);
      data.chantier.nom = baseName;
      // Inclure l'URL de téléchargement Excel dans la réponse
      data._excelUrl = `/api/download-excel/${excelId}/${encodeURIComponent(baseName)}.xlsx`;
      res.json(data);
    } catch (e) {
      res.status(500).json({ error: 'Réponse invalide du parser : ' + e.message });
    }
  });
});

// Téléchargement du fichier Excel généré
app.get('/api/download-excel/:id/:filename', (req, res) => {
  const filePath = path.join(EXCEL_TMP, req.params.id + '.xlsx');
  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ error: 'Fichier Excel introuvable ou expiré' });
  }
  res.download(filePath, req.params.filename, (err) => {
    // Nettoyer le fichier après téléchargement (ou erreur)
    fs.unlink(filePath, () => {});
  });
});

// ─── HTTP + WS SERVER ────────────────────────────────────────────────────────
const server = http.createServer(app);
const wss = new WebSocketServer({ server });

// Map: projectId -> Set<ws>
const rooms = new Map();

function broadcastToProject(projectId, msg, exclude) {
  const room = rooms.get(projectId);
  if (!room) return;
  const payload = JSON.stringify(msg);
  for (const client of room) {
    if (client !== exclude && client.readyState === client.OPEN) {
      client.send(payload);
    }
  }
}

wss.on('connection', (ws) => {
  let currentProjectId = null;

  ws.on('message', (raw) => {
    let msg;
    try { msg = JSON.parse(raw); } catch { return; }

    if (msg.type === 'join') {
      // Leave previous room
      if (currentProjectId) {
        const prev = rooms.get(currentProjectId);
        if (prev) { prev.delete(ws); if (prev.size === 0) rooms.delete(currentProjectId); }
      }
      currentProjectId = msg.projectId;
      if (!rooms.has(currentProjectId)) rooms.set(currentProjectId, new Set());
      rooms.get(currentProjectId).add(ws);

      // Send current viewers count
      broadcastToProject(currentProjectId, {
        type: 'viewers',
        count: rooms.get(currentProjectId).size,
      });
    }

    if (msg.type === 'patch' && currentProjectId) {
      // Relay the patch to all other clients in the room
      broadcastToProject(currentProjectId, msg, ws);
    }
  });

  ws.on('close', () => {
    if (currentProjectId) {
      const room = rooms.get(currentProjectId);
      if (room) {
        room.delete(ws);
        if (room.size === 0) rooms.delete(currentProjectId);
        else broadcastToProject(currentProjectId, { type: 'viewers', count: room.size });
      }
    }
  });
});

server.listen(PORT, () =>
  console.log(`Calepinage running → http://localhost:${PORT}`)
);
