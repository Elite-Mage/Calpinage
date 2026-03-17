const express = require('express');
const http = require('http');
const path = require('path');
const fs = require('fs');
const { WebSocketServer } = require('ws');
const Database = require('better-sqlite3');
const { v4: uuidv4 } = require('uuid');

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
