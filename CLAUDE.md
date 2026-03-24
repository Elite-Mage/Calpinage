# Calpinage — Référence rapide

## Fichiers principaux
- `parse_dxf.py` — parser Python (parse_dxf, calc_ossature_facades)
- `index.html` — parser JS (parseDxfClientSide, calcOssatureFacades, calcOssature)
- Les deux parsers doivent rester **strictement alignés**.

## Règles critiques
- Couleurs ACI : 1-6=RdC→R+5 (stock 3650), 7=blanc (3650), 25=spéciaux (2550), 30=Attique (4270)
- Sous-types : Bandeau Haut (h≈1064±8), Plein (w≈stock±15), sinon Pièce spéciale
- Dimensions normalisées : w=max côté, h=min côté, arrondi mm
- Filtrage : min 10mm, dédup spatiale 5mm
- Ossature : Oméga aux jonctions (gap≤20mm), Zed aux bords libres + entraxe (600mm régulier, 800mm bandeau)

## Détails complets
Voir le code source directement — les commentaires et la logique font référence.
