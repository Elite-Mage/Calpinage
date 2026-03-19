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
- L'ossature du **bandeau haut** est comptée **séparément** (joint entre bandeau et panneaux du dessous, pas alignés).

### OMÉGA (jonction entre 2 panneaux)
- Un oméga est posé **uniquement** à une jonction entre **2 panneaux adjacents**.
- Il est 2× plus épais qu'un Zed → permet de fixer **2 vis** (un panneau à gauche, un panneau à droite).
- Hauteur de l'oméga = hauteur de la zone de recouvrement (overlap Y) entre les 2 panneaux.

### ZED (bord libre ou renfort d'entraxe)
Un Zed est utilisé dans 2 cas :
1. **Bord libre** d'un panneau (pas de panneau voisin de l'autre côté) :
   - Bords extrêmes de façade (gauche/droite)
   - Bords d'ouverture (fenêtre, porte) — un seul bord de panneau → Zed
2. **Renfort d'entraxe** : quand la largeur d'un panneau dépasse `entraxe_max` (600mm par défaut), des Zed intermédiaires sont ajoutés.
   - `nbZed = ceil(largeur_panneau / entraxe_max) - 1`
   - Hauteur du Zed = hauteur du montant (hauteur du panneau)

### Analyse spatiale
L'algorithme utilise les positions réelles des panneaux (`rectsSpatial`) :
1. Les panneaux sont regroupés en **colonnes** par position X (tolérance 10mm).
2. Le bandeau haut (hauteur ~1064mm ±8mm) est séparé des panneaux réguliers.
3. Pour chaque paire de colonnes adjacentes :
   - **Y-overlap** (les 2 colonnes ont des panneaux à la même hauteur) → **OMÉGA** (jonction)
   - **Pas de overlap** (une seule colonne a un panneau) → **ZED** (bord libre)
4. Bords extrêmes de façade → **ZED** (bord libre)
5. ZED d'entraxe calculés par panneau (dimensionnel, indépendant de la position).

### Fichiers implémentant cette logique
- **Python** : `parse_dxf.py` — fonction `calc_ossature_facades()`
- **JavaScript** : `index.html` — fonctions `calcOssatureFacades()` (parser) et `calcOssature()` (rendu)
- Les deux parsers spatiaux sont **strictement alignés** et doivent produire des résultats identiques.
