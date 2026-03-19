# Calpinage — Documentation technique

## Méthode de lecture DXF et métré des PANNEAUX

### Entités DXF supportées
1. **LWPOLYLINE** (>= 3 vertices) → bounding box = rectangle du panneau
2. **POLYLINE** (>= 3 vertices) → bounding box = rectangle du panneau
3. **LINE** — 4 lignes (2 horizontales + 2 verticales) de même couleur/calque formant un rectangle fermé (tolérance 2mm sur les jonctions)
4. **INSERT/BLOCK** — expansion récursive : on entre dans le bloc référencé, on applique l'offset (insert point), et on traite chaque entité du bloc comme si elle était dans le modelspace. Héritage couleur BYBLOCK (code 0).

### Filtrage et déduplication
- **Taille minimum** : largeur ET hauteur >= 10mm (élimine traits, cotes, etc.)
- **Dédoublonnage spatial** : tolérance 5mm sur les 4 coordonnées (xmin, xmax, ymin, ymax). Si un rectangle identique existe déjà, on l'ignore. Évite les doublons quand un panneau est à la fois en entité directe ET dans un bloc INSERT.

### Normalisation des dimensions
- `w = max(largeur, hauteur)` — toujours le plus grand côté
- `h = min(largeur, hauteur)` — toujours le plus petit côté
- Arrondi au mm entier (`Math.round` / `round`)

### Résolution des couleurs ACI
Priorité :
1. Couleur de l'entité elle-même (code DXF 62)
2. BYLAYER (code 256) → on va chercher la couleur du calque dans la table LAYER
3. BYBLOCK (code 0) → on hérite de la couleur de l'INSERT parent

### Correspondance couleur ACI → type de panneau
```
ACI 1 (rouge)    → "Panneaux RdC"        stock_w=3650
ACI 2 (jaune)    → "Panneaux R+1"        stock_w=3650
ACI 3 (vert)     → "Panneaux R+2"        stock_w=3650
ACI 4 (cyan)     → "Panneaux R+3"        stock_w=3650
ACI 5 (bleu)     → "Panneaux R+4"        stock_w=3650
ACI 6 (magenta)  → "Panneaux R+5"        stock_w=3650
ACI 7 (blanc)    → "Panneaux (blanc)"    stock_w=3650
ACI 25           → "Panneaux spéciaux"   stock_w=2550
ACI 30           → "Panneaux Attique"    stock_w=4270
```

### Classification des sous-types de pièces
- **"Bandeau Haut"** : `|h - 1064| <= 8` (hauteur ~1064mm)
- **"Plein"** : `|w - stock_w| <= 15` (largeur = panneau stock à 15mm près)
- **"Pièce spéciale"** : tout le reste

### Affectation aux façades
- Les labels de façade sont extraits depuis les entités TEXT/MTEXT du DXF (position X)
- Chaque panneau est affecté à la façade dont le label est le plus proche de son `xcenter = (xmin + xmax) / 2`
- Si aucun label trouvé → façade par défaut "Façade"

### Calcul de la surface utile
- `surface_pièce = w × h / 1 000 000` (mm² → m²)
- `surface_totale = Σ(surface_pièce × quantité)` pour toutes les pièces

### Fichiers implémentant cette logique
- **Python** : `parse_dxf.py` — fonction `parse_dxf()`
- **JavaScript** : `index.html` — fonction `parseDxfClientSide()`
- Les deux parsers sont **strictement alignés** et doivent produire des résultats identiques.

---

## Méthode de calcul OSSATURE (Oméga et Zed)

### Principes
- Tous les profils (Oméga et Zed) sont **verticaux** — jamais d'ossature horizontale.
- Joint de **8mm** entre tous les panneaux sur la façade.
- **Pas de ZED/OMÉGA aux bords de façade** (gauche/droite extrêmes).
- Le **bandeau haut** (~1064mm) a son ossature intégrée (omegas aux jonctions entre panneaux bandeau adjacents, ZED entraxe 800mm).

### OMÉGA (jonction entre 2 panneaux)
- Un oméga est posé **uniquement** à une jonction entre **2 panneaux adjacents** (gap X ≤ 20mm).
- Il est 2× plus épais qu'un Zed → permet de fixer **2 vis** (un panneau à gauche, un panneau à droite).
- **Séparation par étage** : quand 2 panneaux plein (RDC 2550mm + N1 3650mm) se touchent avec un joint de 8mm, l'oméga est séparé en 2 barres distinctes (une par étage). Les panneaux plein ne fusionnent PAS à travers le joint inter-étage.
- **Fusion fenêtre** : les pièces de fenêtre (< hauteur d'étage) fusionnent à travers le joint inter-étage. Ex: 750mm + 8mm + 822mm = 1580mm → un seul oméga.
- **Bandeau** : un oméga bandeau de hauteur 1064mm est ajouté uniquement aux jonctions entre 2 panneaux bandeau adjacents (gap ≤ 20mm).

### ZED (bord libre ou renfort d'entraxe)
Un Zed est utilisé dans 2 cas :
1. **Bord libre** d'ouverture (fenêtre, porte) — là où un seul côté a un panneau → Zed.
   - Pas de ZED aux bords extrêmes de façade (le user ne les compte pas).
   - Gap ZED : à une jonction fenêtre, les segments Y où un seul côté a du panneau (ex: en dessous de la fenêtre, ou dans l'ouverture fenêtre) → Zed de la hauteur du gap.
2. **Renfort d'entraxe** :
   - Panneaux réguliers : `entraxe_max = 600mm` → `nbZed = ceil(largeur / 600) - 1`
   - Panneaux bandeau : `entraxe_bandeau = 800mm` → `nbZed = ceil(largeur / 800) - 1`
   - Hauteur du Zed = hauteur du panneau sur la façade (ymax - ymin).

### Analyse spatiale
L'algorithme utilise les positions réelles des panneaux (`rectsSpatial`) :
1. Détecter les hauteurs d'étage (RDC/N1) à partir des couleurs ACI (25=RDC, 30=N1).
2. Séparer le **bandeau** (hauteur façade ~1064mm ±15mm) des panneaux réguliers.
3. Construire des **colonnes** par position X (tolérance 10mm), dédupliquer.
4. Pour chaque paire de colonnes adjacentes (gap X ≤ 20mm) :
   - Calculer les **overlaps individuels** entre chaque panneau gauche et droit.
   - **Fusionner** les overlaps séparés par ≤ 10mm (joint), SAUF si les deux overlaps sont des panneaux plein d'étage → **séparer** (un oméga par étage).
   - Les segments fusionnés → **OMÉGA**.
   - Les **gaps** (un seul côté a du panneau) → **ZED** bord libre (si hauteur ≥ 100mm).
5. Si gap X > 20mm → **ouverture** → ZED bord libre sur les panneaux adjacents.
6. ZED d'entraxe calculés par panneau (régulier: 600mm, bandeau: 800mm).
7. **OMÉGA bandeau** : uniquement aux jonctions entre 2 panneaux bandeau adjacents (gap ≤ 20mm), PAS à chaque jonction régulière sous le bandeau.

### Fichiers implémentant cette logique
- **Python** : `parse_dxf.py` — fonction `calc_ossature_facades()`
- **JavaScript** : `index.html` — fonctions `calcOssatureFacades()` (parser) et `calcOssature()` (rendu)
- Les deux parsers spatiaux sont **strictement alignés** et doivent produire des résultats identiques.
