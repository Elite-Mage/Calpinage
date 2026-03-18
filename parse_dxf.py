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

def _merge_y_intervals(panels):
    """Fusionne les intervalles Y des panneaux d'une colonne.
    Retourne une liste triée de (ymin, ymax)."""
    intervals = [(p["ymin"], p["ymax"]) for p in panels]
    intervals.sort()
    merged = []
    for s, e in intervals:
        if merged and s <= merged[-1][1] + 20:  # fusionne si gap < 20mm
            merged[-1] = (merged[-1][0], max(merged[-1][1], e))
        else:
            merged.append((s, e))
    return merged


def _interval_overlap(a_intervals, b_intervals):
    """Longueur totale de recouvrement entre deux ensembles d'intervalles Y."""
    overlap = 0
    for a_s, a_e in a_intervals:
        for b_s, b_e in b_intervals:
            o_start = max(a_s, b_s)
            o_end = min(a_e, b_e)
            if o_end > o_start:
                overlap += o_end - o_start
    return overlap


def _interval_total(intervals):
    """Longueur totale couverte par les intervalles."""
    return sum(e - s for s, e in intervals)


def _interval_diff(a_intervals, b_intervals):
    """Longueur de A non couverte par B."""
    return _interval_total(a_intervals) - _interval_overlap(a_intervals, b_intervals)


def calc_ossature_facades(rects_spatial, facade_labels):
    """
    Analyse la disposition spatiale des panneaux pour calculer l'ossature
    (Oméga et Zed) par façade.

    Logique :
    - Les panneaux sont regroupés en colonnes verticales (même position X).
    - Entre deux colonnes adjacentes :
      - Là où les deux ont des panneaux (joint 8mm) → OMÉGA (jonction)
      - Là où une seule colonne a un panneau (bord d'ouverture) → ZED
    - Bords extrêmes de façade (gauche/droite) → OMÉGA

    Retourne un dict {facade_name: {"omega_mm": int, "zed_mm": int, "omega_ml": float, "zed_ml": float}}
    """
    def nearest_facade(xcenter):
        return min(facade_labels, key=lambda lbl: abs(lbl[0] - xcenter))[1]

    # Regrouper par façade
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

        sorted_cols = sorted(columns.items())
        total_omega_mm = 0
        total_zed_mm = 0

        for i, (col_x, col_panels) in enumerate(sorted_cols):
            col_intervals = _merge_y_intervals(col_panels)
            col_height = _interval_total(col_intervals)

            # Bord gauche de la façade → oméga
            if i == 0:
                total_omega_mm += col_height

            if i < len(sorted_cols) - 1:
                _, next_col_panels = sorted_cols[i + 1]
                next_intervals = _merge_y_intervals(next_col_panels)

                # Là où les deux colonnes ont des panneaux → oméga
                overlap = _interval_overlap(col_intervals, next_intervals)
                # Là où seule la colonne courante a un panneau → zed
                only_current = _interval_diff(col_intervals, next_intervals)
                # Là où seule la colonne suivante a un panneau → zed
                only_next = _interval_diff(next_intervals, col_intervals)

                total_omega_mm += overlap
                total_zed_mm += only_current + only_next
            else:
                # Bord droit de la façade → oméga
                total_omega_mm += col_height

        result[fname] = {
            "omega_mm": round(total_omega_mm),
            "zed_mm": round(total_zed_mm),
            "omega_ml": round(total_omega_mm / 1000, 2),
            "zed_ml": round(total_zed_mm / 1000, 2),
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

    # 2. Extrait tous les rectangles (LWPOLYLINE avec 4 points)
    rects = []
    rects_spatial = []  # Garde les positions complètes pour l'analyse ossature
    for e in msp:
        if e.dxftype() == "LWPOLYLINE":
            pts = list(e.get_points())
            if len(pts) < 3:
                continue
            xs = [float(p[0]) for p in pts]
            ys = [float(p[1]) for p in pts]
            raw_w = max(xs) - min(xs)
            raw_h = max(ys) - min(ys)
            if raw_w < 10 or raw_h < 10:
                continue  # ignore lignes dégénérées
            w, h = normalize_rect(raw_w, raw_h)
            xcenter = (min(xs) + max(xs)) / 2.0
            color_aci = e.dxf.get("color", 256)  # 256 = BYLAYER
            rects.append({"xcenter": xcenter, "w": w, "h": h, "color": color_aci})
            rects_spatial.append({
                "xmin": min(xs), "xmax": max(xs),
                "ymin": min(ys), "ymax": max(ys),
                "color": color_aci,
            })

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

        # Section détaillée : Oméga par façade puis Zed par façade
        ws_oss.append([])
        ws_oss.append(["DÉTAIL OMÉGA PAR FAÇADE"])
        r = ws_oss.max_row
        ws_oss.cell(r, 1).font = Font(bold=True, size=11, color="FF2563EB")

        ws_oss.append(["Façade", "Oméga (mm)", "Oméga (ml)"])
        r = ws_oss.max_row
        for col in range(1, 4):
            ws_oss.cell(r, col).font = Font(bold=True)
            ws_oss.cell(r, col).fill = PatternFill("solid", fgColor="FFD4E6FF")
            ws_oss.cell(r, col).border = _border()

        for facade_name in facades_order:
            oss = ossature.get(facade_name, {})
            ws_oss.append([facade_name, oss.get("omega_mm", 0), oss.get("omega_ml", 0)])
            r = ws_oss.max_row
            for col in range(1, 4):
                ws_oss.cell(r, col).border = _border()

        ws_oss.append([])
        ws_oss.append(["DÉTAIL ZED PAR FAÇADE"])
        r = ws_oss.max_row
        ws_oss.cell(r, 1).font = Font(bold=True, size=11, color="FFCA8A04")

        ws_oss.append(["Façade", "Zed (mm)", "Zed (ml)"])
        r = ws_oss.max_row
        for col in range(1, 4):
            ws_oss.cell(r, col).font = Font(bold=True)
            ws_oss.cell(r, col).fill = PatternFill("solid", fgColor="FFFFF3CD")
            ws_oss.cell(r, col).border = _border()

        for facade_name in facades_order:
            oss = ossature.get(facade_name, {})
            ws_oss.append([facade_name, oss.get("zed_mm", 0), oss.get("zed_ml", 0)])
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
