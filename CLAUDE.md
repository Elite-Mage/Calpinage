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

(À documenter — en cours de vérification)
