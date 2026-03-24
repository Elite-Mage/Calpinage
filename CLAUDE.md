# Calpinage — Référence rapide

## Fichiers principaux
- `parse_dxf.py` — parser Python (parse_dxf, calc_ossature_facades)
- `index.html` — parser JS (parseDxfClientSide, calcOssatureFacades, calcOssature)
- Les deux parsers doivent rester **strictement alignés**.

## Règles critiques
- Couleurs ACI : identifient le type de panneau par couleur, pas de largeur stock associée
- Dimensions : L (largeur) = horizontal, H (hauteur) = vertical, arrondi mm
- Sous-types : Bandeau Toiture (dernier panneau, rien au-dessus, h<1600mm), Plein (h≥1600mm), sinon Pièce spéciale
- Séparation étages : joint horizontal 5-20mm entre panneaux = changement d'étage
- Façades : numérotation automatique (Façade 1, 2, etc.) si pas de texte dans le DXF
- Filtrage : min 10mm, dédup spatiale 5mm
- Ossature : Oméga aux jonctions (gap≤20mm, rectangle 80mm, toujours continu), Zed aux bords libres + entraxe (600mm régulier, 800mm bandeau toiture, rectangle 40mm)

## Détails complets
Voir le code source directement — les commentaires et la logique font référence.
