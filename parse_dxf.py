#!/usr/bin/env python3
"""
parse_dxf.py — Convertit un fichier DXF (ou DWG via dwg2dxf) en JSON + Excel
              structurés pour l'appli Calepinage.

Usage :
  python3 parse_dxf.py fichier.dxf                       → JSON sur stdout
  python3 parse_dxf.py fichier.dxf --excel sortie.xlsx   → JSON stdout + Excel
  python3 parse_dxf.py fichier.dxf --out sortie.json     → JSON fichier + Excel auto
  python3 parse_dxf.py fichier.dwg --excel sortie.xlsx   → conversion DWG→DXF puis pareil

Dépendances : ezdxf, openpyxl  (pip install ezdxf openpyxl)
Pour DWG   : LibreDWG (dwg2dxf) doit être installé (apt install libredwg-tools)
"""

import sys
import json
import math
import os
import subprocess
import tempfile
import argparse
import datetime
from collections import Counter, defaultdict

try:
    import ezdxf
except ImportError:
    print("ERREUR: ezdxf manquant. Lancez: pip install ezdxf", file=sys.stderr)
    sys.exit(1)

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# ─── Palette ACI → informations groupe ───────────────────────────────────────
# Clé = numéro de couleur ACI DXF
ACI_COLOR_MAP = {
    30:  {"name": "Panneaux Orange (N1)",   "color": "#fd9a51", "stock_w": 3650, "niveau": "N1"},
    25:  {"name": "Panneaux Marron (RDC)",  "color": "#8B5E3C", "stock_w": 2550, "niveau": "RDC"},
    1:   {"name": "Panneaux Rouge",         "color": "#e63946", "stock_w": 3650, "niveau": ""},
    2:   {"name": "Panneaux Jaune",         "color": "#f4d03f", "stock_w": 3650, "niveau": ""},
    3:   {"name": "Panneaux Vert",          "color": "#2ecc71", "stock_w": 3650, "niveau": ""},
    4:   {"name": "Panneaux Cyan",          "color": "#1abc9c", "stock_w": 3650, "niveau": ""},
    5:   {"name": "Panneaux Bleu",          "color": "#3498db", "stock_w": 3650, "niveau": ""},
    6:   {"name": "Panneaux Magenta",       "color": "#9b59b6", "stock_w": 3650, "niveau": ""},
    7:   {"name": "Panneaux Blanc/Noir",    "color": "#aaaaaa", "stock_w": 3650, "niveau": ""},
}

def round_mm(v):
    """Arrondit à l'entier millimètre le plus proche."""
    return int(round(float(v)))

def normalize_rect(w, h):
    """Retourne (w, h) avec w ≥ h (grande dimension en premier)."""
    w, h = round_mm(w), round_mm(h)
    return (w, h) if w >= h else (h, w)

def classify_subtype(w, h, stock_w):
    """
    Classifie un panneau en sous-type selon ses dimensions.
    - Bandeau Haut : h ≈ 1064 mm  (valeur typique, ajustable)
    - Plein        : w ≈ stock_w  (largeur du stock, ex. 3650 ou 2550)
    - Pièce spéciale : tout le reste
    """
    if abs(h - 1064) <= 8:
        return "Bandeau Haut"
    if abs(w - stock_w) <= 15:
        return "Plein"
    return "Pièce spéciale"

def _dwg_to_dxf_libredwg(dwg_path):
    """Conversion DWG→DXF via dwg2dxf (LibreDWG) si disponible."""
    for cmd in ["dwg2dxf", "/usr/bin/dwg2dxf", "/usr/local/bin/dwg2dxf"]:
        result = subprocess.run(["which", cmd.split("/")[-1]], capture_output=True, text=True)
        if result.returncode == 0:
            tmpdir = tempfile.mkdtemp()
            base = os.path.splitext(os.path.basename(dwg_path))[0]
            out_dxf = os.path.join(tmpdir, base + ".dxf")
            subprocess.run([cmd, dwg_path, "-o", out_dxf], check=True,
                           capture_output=True, text=True)
            return out_dxf
    return None


def _dwg_to_dxf_aspose(dwg_path):
    """Conversion DWG→DXF via aspose-cad (fallback)."""
    try:
        import aspose.cad as cad
        from aspose.cad.imageoptions import DxfOptions
        image = cad.Image.load(dwg_path)
        opts = DxfOptions()
        tmpdir = tempfile.mkdtemp()
        base = os.path.splitext(os.path.basename(dwg_path))[0]
        out_dxf = os.path.join(tmpdir, base + ".dxf")
        image.save(out_dxf, opts)
        return out_dxf
    except ImportError:
        return None
    except Exception as e:
        print(f"aspose-cad conversion failed: {e}", file=sys.stderr)
        return None


def dwg_to_dxf(dwg_path):
    """
    Convertit DWG→DXF en essayant d'abord dwg2dxf (LibreDWG),
    puis aspose-cad en fallback.
    Retourne le chemin du DXF temporaire ou lève une exception.
    """
    # Essai 1 : LibreDWG (dwg2dxf)
    result = _dwg_to_dxf_libredwg(dwg_path)
    if result:
        return result

    # Essai 2 : aspose-cad
    result = _dwg_to_dxf_aspose(dwg_path)
    if result:
        return result

    raise RuntimeError(
        "Aucun convertisseur DWG→DXF disponible.\n"
        "Installez l'un des deux :\n"
        "  pip install aspose-cad\n"
        "  ou compilez LibreDWG depuis https://github.com/LibreDWG/libredwg"
    )

# ─── Calcul ossature par façade (analyse spatiale) ────────────────────────────

def calc_ossature_facades(rects_spatial, facade_labels, entraxe_max=600):
    """
    Analyse la disposition spatiale des panneaux pour calculer l'ossature
    (Oméga et Zed) par façade.

    Logique :
    - Les panneaux sont regroupés en colonnes verticales (même position X).
    - OMÉGA (jonction entre 2 panneaux côte à côte) :
      - Bords extrêmes de façade (gauche/droite)
      - Entre deux colonnes adjacentes là où les deux ont des panneaux
    - ZED (bords d'ouverture + entraxe) :
      - Bord d'ouverture : là où une seule colonne a un panneau (fenêtre/porte)
      - Entraxe : support intermédiaire quand la largeur du panneau > entraxe_max

    Retourne un dict {facade_name: {"omega_mm", "zed_mm", "omega_ml", "zed_ml",
                                     "omega_details": {h: qty}, "zed_details": {h: qty}}}
    """
    def nearest_facade(xcenter):
        return min(facade_labels, key=lambda lbl: abs(lbl[0] - xcenter))[1]

    by_facade = defaultdict(list)
    for r in rects_spatial:
        fname = nearest_facade((r["xmin"] + r["xmax"]) / 2)
        by_facade[fname].append(r)

    result = {}
    for fname in sorted(by_facade.keys()):
        panels = by_facade[fname]

        # Regrouper en colonnes par position X (tolérance 10mm)
        columns = defaultdict(list)
        for p in panels:
            col_key = round(p["xmin"] / 10) * 10
            columns[col_key].append(p)

        # Dédupliquer les panneaux par colonne (couches DXF superposées)
        # et retirer les sous-panneaux englobés par un panneau plus grand
        for col_key in columns:
            seen = set()
            unique = []
            for p in columns[col_key]:
                sig = (round(p["ymin"]), round(p["ymax"]))
                if sig not in seen:
                    seen.add(sig)
                    unique.append(p)
            # Retirer les panneaux strictement englobés par un autre
            filtered = []
            unique.sort(key=lambda p: p["ymax"] - p["ymin"], reverse=True)
            for p in unique:
                englobed = False
                for bigger in filtered:
                    if p["ymin"] >= bigger["ymin"] - 1 and p["ymax"] <= bigger["ymax"] + 1:
                        englobed = True
                        break
                if not englobed:
                    filtered.append(p)
            columns[col_key] = filtered

        sorted_cols = sorted(columns.items())
        total_omega_mm = 0
        total_zed_mm = 0
        omega_details = defaultdict(int)  # hauteur → quantité
        zed_details = defaultdict(int)

        for i, (col_x, col_panels) in enumerate(sorted_cols):
            col_panels_sorted = sorted(col_panels, key=lambda p: p["ymin"])

            # ── ZED d'entraxe : support intermédiaire par panneau ──
            for p in col_panels_sorted:
                pw = round(p["xmax"] - p["xmin"])
                ph = round(p["ymax"] - p["ymin"])
                nb_zed = max(0, math.ceil(pw / entraxe_max) - 1)
                if nb_zed > 0:
                    total_zed_mm += nb_zed * ph
                    zed_details[ph] += nb_zed

            # ── OMÉGA / ZED aux jonctions entre colonnes ──
            if i == 0:
                # Bord gauche de la façade → oméga par panneau
                for p in col_panels_sorted:
                    ph = round(p["ymax"] - p["ymin"])
                    total_omega_mm += ph
                    omega_details[ph] += 1

            if i < len(sorted_cols) - 1:
                _, next_col_panels = sorted_cols[i + 1]
                next_sorted = sorted(next_col_panels, key=lambda p: p["ymin"])

                # Pour chaque panneau de la colonne courante :
                # vérifier le recouvrement avec les panneaux de la colonne suivante
                for p in col_panels_sorted:
                    p_ymin, p_ymax = p["ymin"], p["ymax"]
                    overlap = 0
                    for np in next_sorted:
                        os = max(p_ymin, np["ymin"])
                        oe = min(p_ymax, np["ymax"])
                        if oe > os:
                            overlap += oe - os
                    non_overlap = (p_ymax - p_ymin) - overlap
                    if overlap > 0:
                        h = round(overlap)
                        total_omega_mm += h
                        omega_details[h] += 1
                    if non_overlap > 10:
                        h = round(non_overlap)
                        total_zed_mm += h
                        zed_details[h] += 1

                # Panneaux de la colonne suivante non couverts par la courante
                for np in next_sorted:
                    np_ymin, np_ymax = np["ymin"], np["ymax"]
                    overlap = 0
                    for p in col_panels_sorted:
                        os = max(np_ymin, p["ymin"])
                        oe = min(np_ymax, p["ymax"])
                        if oe > os:
                            overlap += oe - os
                    non_overlap = (np_ymax - np_ymin) - overlap
                    if non_overlap > 10:
                        h = round(non_overlap)
                        total_zed_mm += h
                        zed_details[h] += 1
            else:
                # Bord droit de la façade → oméga par panneau
                for p in col_panels_sorted:
                    ph = round(p["ymax"] - p["ymin"])
                    total_omega_mm += ph
                    omega_details[ph] += 1

        result[fname] = {
            "omega_mm": round(total_omega_mm),
            "zed_mm": round(total_zed_mm),
            "omega_ml": round(total_omega_mm / 1000, 2),
            "zed_ml": round(total_zed_mm / 1000, 2),
            "omega_details": dict(omega_details),
            "zed_details": dict(zed_details),
        }

    return result


def parse_dxf_file(filepath):
    """
    Lit un fichier DXF et retourne les données structurées sous forme de dict
    compatible avec le format JSON de l'appli Calepinage.
    """
    doc = ezdxf.readfile(filepath)
    msp = doc.modelspace()

    # 1. Récupère les labels de façades (entités TEXT)
    facade_labels = []
    for e in msp:
        if e.dxftype() in ("TEXT", "MTEXT"):
            try:
                txt = e.dxf.text if e.dxftype() == "TEXT" else e.text
                x = e.dxf.insert[0]
                facade_labels.append((x, txt.strip()))
            except Exception:
                pass

    facade_labels.sort(key=lambda t: t[0])

    if not facade_labels:
        facade_labels = [(0.0, "Façade")]

    # 2. Construit un dict layer → couleur ACI pour résoudre BYLAYER
    layer_colors = {}
    for layer in doc.layers:
        try:
            c = layer.dxf.color
            if c is not None:
                layer_colors[layer.dxf.name] = abs(c)
        except Exception:
            pass

    def resolve_color(entity):
        """Résout la couleur ACI : entité > calque (BYLAYER)."""
        color = entity.dxf.get("color", 256)
        if color == 256:  # BYLAYER
            layer_name = entity.dxf.get("layer", "0")
            color = layer_colors.get(layer_name, 256)
        return color

    # 3. Extrait tous les rectangles (LWPOLYLINE + POLYLINE)
    rects = []
    rects_spatial = []  # Garde les positions complètes pour l'analyse ossature

    def add_rect(xmin, xmax, ymin, ymax, color_aci):
        """Ajoute un rectangle avec dédoublonnage spatial (tolérance 5mm)."""
        raw_w = xmax - xmin
        raw_h = ymax - ymin
        if raw_w < 10 or raw_h < 10:
            return
        # Dédoublonnage spatial
        TOL = 5
        for ex in rects_spatial:
            if (abs(ex["xmin"] - xmin) < TOL and abs(ex["xmax"] - xmax) < TOL and
                    abs(ex["ymin"] - ymin) < TOL and abs(ex["ymax"] - ymax) < TOL):
                return  # doublon
        w, h = normalize_rect(raw_w, raw_h)
        xcenter = (xmin + xmax) / 2.0
        rects.append({"xcenter": xcenter, "w": w, "h": h, "color": color_aci})
        rects_spatial.append({"xmin": xmin, "xmax": xmax, "ymin": ymin, "ymax": ymax, "color": color_aci})

    def process_entity(e, offset_x=0.0, offset_y=0.0, parent_color=None, parent_layer=None):
        """Traite une entité DXF (supporte offset pour les INSERT/blocks)."""
        def _resolve(e):
            c = resolve_color(e)
            if c == 0 and parent_color:  # BYBLOCK
                c = parent_color
            if c == 256 and parent_layer:
                c = layer_colors.get(parent_layer, 256)
            return c

        if e.dxftype() == "LWPOLYLINE":
            pts = [(p[0] + offset_x, p[1] + offset_y) for p in e.get_points()]
            if len(pts) >= 3:
                xs = [float(p[0]) for p in pts]
                ys = [float(p[1]) for p in pts]
                add_rect(min(xs), max(xs), min(ys), max(ys), _resolve(e))
        elif e.dxftype() == "POLYLINE":
            try:
                pts = [(v.dxf.location.x + offset_x, v.dxf.location.y + offset_y) for v in e.vertices]
                if len(pts) >= 3:
                    xs = [float(p[0]) for p in pts]
                    ys = [float(p[1]) for p in pts]
                    add_rect(min(xs), max(xs), min(ys), max(ys), _resolve(e))
            except Exception:
                pass
        elif e.dxftype() == "INSERT":
            try:
                block_name = e.dxf.name
                block = doc.blocks.get(block_name)
                if block is None:
                    return
                ins_x = float(e.dxf.get("insert", (0, 0, 0))[0]) + offset_x
                ins_y = float(e.dxf.get("insert", (0, 0, 0))[1]) + offset_y
                ins_color = resolve_color(e)
                ins_layer = e.dxf.get("layer", "0")
                for be in block:
                    process_entity(be, ins_x, ins_y, ins_color, ins_layer)
            except Exception:
                pass

    for e in msp:
        process_entity(e)

    # 2b. Assembler les LINE isolées en rectangles (4 lignes → 1 rectangle)
    lines_by_key = defaultdict(list)
    for e in msp:
        if e.dxftype() == "LINE":
            color_aci = resolve_color(e)
            layer = e.dxf.get("layer", "0")
            key = f"{color_aci}|{layer}"
            x1, y1 = float(e.dxf.start.x), float(e.dxf.start.y)
            x2, y2 = float(e.dxf.end.x), float(e.dxf.end.y)
            lines_by_key[key].append((x1, y1, x2, y2, color_aci, layer))

    for key, segs in lines_by_key.items():
        h_segs = [(s[0], s[1], s[2], s[3]) for s in segs if abs(s[1] - s[3]) < 1]
        v_segs = [(s[0], s[1], s[2], s[3]) for s in segs if abs(s[0] - s[2]) < 1]
        color_aci = segs[0][4]
        for hi in range(len(h_segs)):
            h1 = h_segs[hi]
            h1xmin, h1xmax = min(h1[0], h1[2]), max(h1[0], h1[2])
            for hj in range(hi + 1, len(h_segs)):
                h2 = h_segs[hj]
                h2xmin, h2xmax = min(h2[0], h2[2]), max(h2[0], h2[2])
                if abs(h1xmin - h2xmin) > 2 or abs(h1xmax - h2xmax) > 2:
                    continue
                ymin = min(h1[1], h2[1])
                ymax = max(h1[1], h2[1])
                if ymax - ymin < 10:
                    continue
                has_left = has_right = False
                for v in v_segs:
                    vx = (v[0] + v[2]) / 2
                    vymin, vymax = min(v[1], v[3]), max(v[1], v[3])
                    if abs(vymin - ymin) < 2 and abs(vymax - ymax) < 2:
                        if abs(vx - h1xmin) < 2:
                            has_left = True
                        if abs(vx - h1xmax) < 2:
                            has_right = True
                if has_left and has_right:
                    add_rect(h1xmin, h1xmax, ymin, ymax, color_aci)

    # 3. Assigne chaque rectangle à la façade la plus proche (par X)
    def nearest_facade(xcenter):
        return min(facade_labels, key=lambda lbl: abs(lbl[0] - xcenter))[1]

    # 4. Regroupe par (couleur ACI, façade) → compteur (w, h)
    data = defaultdict(lambda: defaultdict(Counter))
    for r in rects:
        facade = nearest_facade(r["xcenter"])
        data[r["color"]][facade][(r["w"], r["h"])] += 1

    # 5. Construit la structure JSON de l'appli
    groups = []
    gid = 1
    for color_aci in sorted(data.keys()):
        info = ACI_COLOR_MAP.get(color_aci, {
            "name": f"Panneaux (ACI {color_aci})",
            "color": "#888888",
            "stock_w": 3650,
            "niveau": "",
        })
        stock_w = info["stock_w"]

        subsections = []
        ssid = 1
        for facade_name in sorted(data[color_aci].keys()):
            dim_counter = data[color_aci][facade_name]

            # Regroupe par sous-type
            subtype_pieces = defaultdict(list)
            for (w, h), qty in sorted(dim_counter.items()):
                st = classify_subtype(w, h, stock_w)
                subtype_pieces[st].append({"w": w, "h": h, "qty": qty})

            panel_subtypes = []
            stid = 1
            for st_name in ["Bandeau Haut", "Plein", "Pièce spéciale"]:
                pieces_list = subtype_pieces.get(st_name, [])
                if not pieces_list:
                    continue
                pieces = [
                    {"id": pid, "w": p["w"], "h": p["h"], "qty": p["qty"]}
                    for pid, p in enumerate(pieces_list, 1)
                ]
                panel_subtypes.append({
                    "id": stid,
                    "name": st_name,
                    "pieces": pieces,
                    "nextPieceId": len(pieces) + 1,
                })
                stid += 1

            if panel_subtypes:
                subsections.append({
                    "id": ssid,
                    "name": facade_name,
                    "panelSubtypes": panel_subtypes,
                    "nextSubtypeId": stid,
                    "activeSubtypeId": 1,
                })
                ssid += 1

        if subsections:
            groups.append({
                "id": gid,
                "name": info["name"],
                "color": info["color"],
                "reference": "",
                "coloris": "",
                "epaisseur": 18,
                "prixM2": "",
                "subsections": subsections,
                "nextSubsectionId": ssid,
                "activeSubsectionId": 1,
            })
            gid += 1

    today = datetime.date.today().isoformat()
    base_name = os.path.splitext(os.path.basename(filepath))[0]

    # Calcul ossature par façade (analyse spatiale)
    ossature = calc_ossature_facades(rects_spatial, facade_labels)

    return {
        "version": "6.0",
        "chantier": {"nom": base_name, "date": today},
        "stockPanels": [
            {"id": 1, "w": 3650, "h": 1860, "active": True},
            {"id": 2, "w": 2550, "h": 1860, "active": True},
            {"id": 3, "w": 4270, "h": 2130, "active": True},
        ],
        "nextStockId": 4,
        "groups": groups,
        "nextGroupId": gid,
        "activeGroupId": 1,
        "ossature_facades": ossature,
    }

# ─── Génération Excel ─────────────────────────────────────────────────────────

def hex_to_argb(hex_color):
    """'#fd9a51' → 'FFFD9A51'"""
    h = hex_color.lstrip("#")
    return "FF" + h.upper()

def _border():
    thin = Side(style="thin", color="FFD0D0D0")
    return Border(left=thin, right=thin, top=thin, bottom=thin)

def _header_fill(hex_color):
    return PatternFill("solid", fgColor=hex_to_argb(hex_color))

def generate_excel(data, out_path):
    """Génère un fichier Excel détaillé à partir du dict data (format JSON appli)."""
    if not HAS_OPENPYXL:
        raise RuntimeError("openpyxl manquant. Lancez: pip install openpyxl")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # retire la feuille par défaut

    facades_order = []
    for g in data["groups"]:
        for ss in g["subsections"]:
            if ss["name"] not in facades_order:
                facades_order.append(ss["name"])

    # ── 1. Feuille Récapitulatif ────────────────────────────────────────────
    ws_recap = wb.create_sheet("Récapitulatif")
    title = data["chantier"]["nom"] + " — Récapitulatif"
    ws_recap.append([title])
    ws_recap["A1"].font = Font(bold=True, size=13)
    ws_recap.append(["Catégorie", "Larg. (mm)", "Haut. (mm)"] + facades_order)
    for col in range(1, 4 + len(facades_order)):
        ws_recap.cell(2, col).font = Font(bold=True)
        ws_recap.cell(2, col).fill = PatternFill("solid", fgColor="FFE0E0E0")
        ws_recap.cell(2, col).border = _border()

    # Collecte toutes les lignes du récap
    recap_rows = defaultdict(lambda: {f: 0 for f in facades_order})
    recap_keys = []  # (categorie, w, h) dans l'ordre de rencontre

    for g in data["groups"]:
        niv = ACI_COLOR_MAP.get(
            next((k for k, v in ACI_COLOR_MAP.items() if v["color"] == g["color"]), None),
            {}
        ).get("niveau", "")
        for ss in g["subsections"]:
            for st in ss["panelSubtypes"]:
                for p in st["pieces"]:
                    if st["name"] == "Plein":
                        cat = f"Plein {niv}".strip()
                        col1, col2 = p["h"], p["w"]  # façade h, stock w
                    elif st["name"] == "Bandeau Haut":
                        cat = "Bandeau Haut"
                        col1, col2 = p["w"], p["h"]
                    else:
                        cat = f"{st['name']} {niv}".strip()
                        col1, col2 = min(p["w"], p["h"]), max(p["w"], p["h"])
                    key = (cat, col1, col2)
                    if key not in recap_keys:
                        recap_keys.append(key)
                    recap_rows[key][ss["name"]] += p["qty"]

    for cat, c1, c2 in recap_keys:
        row = [cat, c1, c2] + [recap_rows[(cat, c1, c2)].get(f, 0) for f in facades_order]
        ws_recap.append(row)
        r = ws_recap.max_row
        for col in range(1, 4 + len(facades_order)):
            ws_recap.cell(r, col).border = _border()

    ws_recap.column_dimensions["A"].width = 28
    ws_recap.column_dimensions["B"].width = 14
    ws_recap.column_dimensions["C"].width = 14
    for i in range(len(facades_order)):
        ws_recap.column_dimensions[get_column_letter(4 + i)].width = 14

    # ── 2. Feuille par façade ───────────────────────────────────────────────
    for facade_name in facades_order:
        ws = wb.create_sheet(facade_name)
        ws.append([facade_name])
        ws["A1"].font = Font(bold=True, size=12)

        headers = ["Catégorie", "Largeur façade (mm)", "Hauteur panneau (mm)",
                   "Quantité", "Surface unitaire (m²)", "Surface totale (m²)"]
        ws.append(headers)
        for col in range(1, len(headers) + 1):
            ws.cell(2, col).font = Font(bold=True)
            ws.cell(2, col).fill = PatternFill("solid", fgColor="FFE8E8E8")
            ws.cell(2, col).border = _border()

        total_surface = 0.0

        for g in data["groups"]:
            ss = next((s for s in g["subsections"] if s["name"] == facade_name), None)
            if ss is None:
                continue
            niv = ACI_COLOR_MAP.get(
                next((k for k, v in ACI_COLOR_MAP.items() if v["color"] == g["color"]), None),
                {}
            ).get("niveau", "")
            fill = _header_fill(g["color"])

            for st in ss["panelSubtypes"]:
                for p in st["pieces"]:
                    if st["name"] == "Plein":
                        cat = f"Plein {niv}".strip()
                        c1, c2 = p["h"], p["w"]
                    elif st["name"] == "Bandeau Haut":
                        cat = "Bandeau Haut"
                        c1, c2 = p["w"], p["h"]
                    else:
                        cat = f"{st['name']} {niv}".strip()
                        c1, c2 = min(p["w"], p["h"]), max(p["w"], p["h"])
                    s_unit = round(c1 * c2 / 1_000_000, 4)
                    s_tot = round(s_unit * p["qty"], 4)
                    total_surface += s_tot
                    ws.append([cat, c1, c2, p["qty"], s_unit, s_tot])
                    r = ws.max_row
                    for col in range(1, 7):
                        ws.cell(r, col).border = _border()
                    ws.cell(r, 1).fill = fill

        # Ligne total
        ws.append(["TOTAL FAÇADE", None, None, None, None, round(total_surface, 4)])
        r = ws.max_row
        for col in range(1, 7):
            ws.cell(r, col).font = Font(bold=True)
            ws.cell(r, col).fill = PatternFill("solid", fgColor="FFDDDDDD")
            ws.cell(r, col).border = _border()

        ws.column_dimensions["A"].width = 28
        for col_letter in ["B", "C", "D", "E", "F"]:
            ws.column_dimensions[col_letter].width = 20

    # ── 3. Feuille Ossature ────────────────────────────────────────────────
    ossature = data.get("ossature_facades", {})
    if ossature:
        ws_oss = wb.create_sheet("Ossature")
        title = data["chantier"]["nom"] + " — Ossature"
        ws_oss.append([title])
        ws_oss["A1"].font = Font(bold=True, size=13)
        ws_oss.append([])  # ligne vide

        # En-têtes
        ws_oss.append(["Façade", "Oméga (ml)", "Zed (ml)", "Total (ml)"])
        r = ws_oss.max_row
        for col in range(1, 5):
            ws_oss.cell(r, col).font = Font(bold=True)
            ws_oss.cell(r, col).fill = PatternFill("solid", fgColor="FFE0E0E0")
            ws_oss.cell(r, col).border = _border()

        grand_omega = 0
        grand_zed = 0

        for facade_name in facades_order:
            oss = ossature.get(facade_name, {})
            omega_ml = oss.get("omega_ml", 0)
            zed_ml = oss.get("zed_ml", 0)
            total_ml = round(omega_ml + zed_ml, 2)
            grand_omega += omega_ml
            grand_zed += zed_ml

            ws_oss.append([facade_name, omega_ml, zed_ml, total_ml])
            r = ws_oss.max_row
            for col in range(1, 5):
                ws_oss.cell(r, col).border = _border()
            ws_oss.cell(r, 2).number_format = '0.00'
            ws_oss.cell(r, 3).number_format = '0.00'
            ws_oss.cell(r, 4).number_format = '0.00'

        # Ligne total
        ws_oss.append(["TOTAL", round(grand_omega, 2), round(grand_zed, 2),
                        round(grand_omega + grand_zed, 2)])
        r = ws_oss.max_row
        for col in range(1, 5):
            ws_oss.cell(r, col).font = Font(bold=True)
            ws_oss.cell(r, col).fill = PatternFill("solid", fgColor="FFDDDDDD")
            ws_oss.cell(r, col).border = _border()
            if col >= 2:
                ws_oss.cell(r, col).number_format = '0.00'

        # Section détaillée par hauteur de profil
        for profile_type, label, color_hex, fill_hex in [
            ("omega_details", "DÉTAIL OMÉGA PAR HAUTEUR", "FF2563EB", "FFD4E6FF"),
            ("zed_details", "DÉTAIL ZED PAR HAUTEUR", "FFCA8A04", "FFFFF3CD"),
        ]:
            ws_oss.append([])
            ws_oss.append([label])
            r = ws_oss.max_row
            ws_oss.cell(r, 1).font = Font(bold=True, size=11, color=color_hex)

            ws_oss.append(["Hauteur (mm)", "Quantité", "Total (ml)"])
            r = ws_oss.max_row
            for col in range(1, 4):
                ws_oss.cell(r, col).font = Font(bold=True)
                ws_oss.cell(r, col).fill = PatternFill("solid", fgColor=fill_hex)
                ws_oss.cell(r, col).border = _border()

            # Agréger toutes les façades
            all_details = defaultdict(int)
            for facade_name in facades_order:
                oss = ossature.get(facade_name, {})
                for h_str, qty in oss.get(profile_type, {}).items():
                    all_details[int(h_str)] += qty

            for h in sorted(all_details.keys(), reverse=True):
                qty = all_details[h]
                ml = round(h * qty / 1000, 2)
                ws_oss.append([h, qty, ml])
                r = ws_oss.max_row
                for col in range(1, 4):
                    ws_oss.cell(r, col).border = _border()

        ws_oss.column_dimensions["A"].width = 22
        ws_oss.column_dimensions["B"].width = 16
        ws_oss.column_dimensions["C"].width = 16
        ws_oss.column_dimensions["D"].width = 16

    wb.save(out_path)
    return out_path

# ─── Arrondissement d'un JSON existant ───────────────────────────────────────

def round_json(data):
    """Arrondit toutes les dimensions w/h au mm le plus proche dans un JSON existant."""
    import copy
    d = copy.deepcopy(data)
    for g in d.get("groups", []):
        for ss in g.get("subsections", []):
            for st in ss.get("panelSubtypes", []):
                for p in st.get("pieces", []):
                    p["w"] = round_mm(p["w"])
                    p["h"] = round_mm(p["h"])
    return d

# ─── CLI ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Parse DXF/DWG → JSON + Excel Calepinage")
    parser.add_argument("input", help="Fichier DXF ou DWG en entrée")
    parser.add_argument("--excel", "-e", help="Chemin de sortie Excel (.xlsx)", default=None)
    parser.add_argument("--out", "-o", help="Chemin de sortie JSON", default=None)
    parser.add_argument("--round-json", help="Arrondir un JSON existant (ne parse pas de DXF)")
    args = parser.parse_args()

    # Mode arrondi d'un JSON existant
    if args.round_json:
        with open(args.round_json, encoding="utf-8") as f:
            data = json.load(f)
        data = round_json(data)
        out = args.out or args.round_json
        with open(out, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"JSON arrondi sauvegardé → {out}", file=sys.stderr)
        return

    filepath = args.input
    tmp_dxf = None

    # Si DWG : conversion
    if filepath.lower().endswith(".dwg"):
        print("Conversion DWG → DXF...", file=sys.stderr)
        try:
            tmp_dxf = dwg_to_dxf(filepath)
            filepath = tmp_dxf
        except RuntimeError as e:
            print(f"ERREUR conversion DWG : {e}", file=sys.stderr)
            sys.exit(1)

    # Parse DXF
    print(f"Lecture de {filepath}...", file=sys.stderr)
    data = parse_dxf_file(filepath)

    # JSON
    json_str = json.dumps(data, ensure_ascii=False, indent=2)
    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(json_str)
        print(f"JSON sauvegardé → {args.out}", file=sys.stderr)
    else:
        print(json_str)

    # Excel
    excel_path = args.excel
    if excel_path is None and args.out:
        excel_path = os.path.splitext(args.out)[0] + ".xlsx"
    if excel_path:
        generate_excel(data, excel_path)
        print(f"Excel sauvegardé → {excel_path}", file=sys.stderr)

    # Nettoyage DXF temporaire
    if tmp_dxf and os.path.exists(tmp_dxf):
        os.unlink(tmp_dxf)

if __name__ == "__main__":
    main()
