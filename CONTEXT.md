# Calpinage — Contexte projet pour VS Code

## Projet
Application web de **calpinage** (optimisation de découpe de panneaux de bardage) pour le bâtiment.
Elle lit des fichiers DXF de plans de façade, extrait les panneaux, les classe et optimise leur découpe dans des panneaux fournisseur.

## Architecture

### Fichiers principaux
| Fichier | Rôle |
|---------|------|
| `index.html` | Application web complète (HTML + CSS + JS inline, ~4500 lignes). Contient le parser DXF côté client, l'optimisation 2D, le rendu, l'export Excel/DXF, la gestion des projets. |
| `parse_dxf.py` | Parser DXF côté serveur (Python). Fonctions : `parse_dxf_file()`, `calc_ossature_facades()`, `generate_excel()`. |
| `server.js` | Serveur Node.js Express. API REST pour les projets (`/api/projects`) et parsing DXF/DWG (`/api/parse-dxf`). WebSocket pour sync temps réel. |
| `CLAUDE.md` | Documentation technique détaillée des algorithmes (lecture DXF, ossature, classification). |

### Les deux parsers (JS + Python) doivent rester strictement alignés
Toute modification du parser JS dans `index.html` doit être répliquée dans `parse_dxf.py` et vice-versa.

## Structure de données

### Hiérarchie (ordre de tri)
```
Groupes (par couleur ACI, tri numérique)
  └─ Façades (tri alphabétique)
       └─ Sous-types (ordre fixe : Bandeau Haut → Plein → Pièce spéciale)
            └─ Pièces (tri par surface décroissante)
```

### Couleurs ACI → Types de panneaux
```
ACI 1  (rouge)   → Panneaux Rouge       stock_w=3650
ACI 2  (jaune)   → Panneaux Jaune       stock_w=3650
ACI 3  (vert)    → Panneaux Vert        stock_w=3650
ACI 4  (cyan)    → Panneaux Cyan        stock_w=3650
ACI 5  (bleu)    → Panneaux Bleu        stock_w=3650
ACI 6  (magenta) → Panneaux Magenta     stock_w=3650
ACI 7  (blanc)   → Panneaux Blanc/Noir  stock_w=3650
ACI 25 (marron)  → Panneaux Marron RDC  stock_w=2550
ACI 30 (orange)  → Panneaux Orange N1   stock_w=3650
ACI 114          → fusionné avec ACI 7
```

### Classification des sous-types
- **Bandeau Haut** : hauteur ~1064mm (±8mm)
- **Plein** : largeur = stock_w (±15mm)
- **Pièce spéciale** : tout le reste

## Fonctions clés dans index.html

| Fonction | Ligne ~ | Rôle |
|----------|---------|------|
| `parseDxfClientSide()` | ~2090 | Parse un fichier DXF côté navigateur |
| `renderGroupTabs()` | ~909 | Rendu des onglets de groupes (couleurs) |
| `renderActiveGroup()` | ~926 | Rendu du groupe actif (façades, sous-types, pièces) |
| `migrateGroups()` | ~901 | Migration + tri des données au chargement |
| `optimize()` | ~1367 | Lance l'optimisation de découpe 2D |
| `renderAllResults()` | ~1397 | Affiche les résultats d'optimisation |
| `exportExcel()` | ~1726 | Export Excel (récap + découpe + ossature) |
| `calcOssature()` | ~3270 | Calcul de l'ossature (oméga + zed) |
| `calcOssatureFacades()` | ~2552 | Calcul ossature spatiale dans le parser |
| `applyProjectData()` | ~1908 | Chargement d'un projet sauvegardé |
| `exportDxfPanneaux()` | ~4170 | Export DXF des panneaux |
| `exportDxfOssature()` | ~4240 | Export DXF de l'ossature |

## Fonctions clés dans parse_dxf.py

| Fonction | Ligne ~ | Rôle |
|----------|---------|------|
| `parse_dxf_file()` | ~366 | Parse un DXF, retourne le JSON complet |
| `calc_ossature_facades()` | ~135 | Calcul ossature par analyse spatiale |
| `generate_excel()` | ~655 | Génère un fichier Excel (.xlsx) |
| `classify_subtype()` | ~64 | Classe une pièce en sous-type |

## Règles importantes

1. **Pas de dimensions stock par défaut** — la liste `stockPanels` est vide à l'import DXF. L'utilisateur ajoute ses dimensions manuellement.
2. **Tri obligatoire** partout (UI + Excel) : Couleur → Façade → Sous-type → Pièces par surface décroissante.
3. **Excel** : feuille récap + une feuille par couleur (pas par façade). Dans chaque feuille : sections par façade puis sous-types.
4. **Ossature** : oméga aux jonctions entre 2 panneaux adjacents, zed aux bords libres et en renfort d'entraxe (600mm régulier, 800mm bandeau).
5. **Déduplication** : tolérance 5mm sur les coordonnées spatiales pour éviter les doublons.

## Stack technique
- Frontend : HTML/CSS/JS vanilla (pas de framework)
- Backend : Node.js + Express + WebSocket
- Python : ezdxf + openpyxl
- Stockage : API REST ou localStorage (fallback)

## Branche de développement
```
claude/fix-dxf-panel-export-TVSMG
```

## Commandes utiles
```bash
node server.js          # Lancer le serveur
python3 parse_dxf.py "fichier.dxf"   # Parser un DXF en ligne de commande
```
