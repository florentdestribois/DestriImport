# Export Optiplanning & SWOOD - Destribois

Outil d'export des materiaux et chants depuis `Outil_Material_Import.xlsm` vers les formats Optiplanning (TXT) et SWOOD (XML).

Developpe pour **Destribois** - Agencement & Menuiserie.

---

## Contenu du dossier

| Fichier | Role |
|---|---|
| `Export_Optiplanning.exe` | Executable Windows (double-clic pour lancer) |
| `export_optiplanning.py` | Code source Python (GUI + CLI) |
| `Outil_Material_Import.xlsm` | Fichier Excel SWOOD source (pages Materials + EdgeBands) |
| `charte_graphique/` | Polices, logo et couleurs Destribois (embarques dans l'EXE) |

---

## 4 Exports disponibles

### 1. TXT Optiplanning

| | |
|---|---|
| **Source** | Page Materials - colonnes A, D-F, AR-AT |
| **Sortie** | `Materiaux_a_importer_Optiplanning_YYYYMMDD_HHMMSS.txt` |
| **Format** | 8 colonnes separees par tabulation |

**Colonnes exportees :**

| # | Colonne | Description |
|---|---|---|
| 1 | SawReference | Reference de sciage (nom + epaisseur si melamine) |
| 2 | BOARDL | Longueur de la plaque (mm) |
| 3 | BOARDW | Largeur de la plaque (mm) |
| 4 | Thickness | Epaisseur (mm) |
| 5 | FiberMaterial | Sens du fil (1 = horizontal, 0 = aucun) |
| 6 | Cost | Cout unitaire en euros/m2 |
| 7 | Parametres | "Destribois" ou "Destribois 5m" si BOARDL > 3200 mm |
| 8 | Ref Fournisseur | Reference du fournisseur |

---

### 2. XML Plaques Nesting

| | |
|---|---|
| **Source** | Page Materials - 49 colonnes |
| **Sortie** | `Plaques_Nesting_YYYYMMDD_HHMMSS.xml` |
| **Format** | `<SWOODMat><Boards><Board .../></Boards></SWOODMat>` |
| **Usage** | Import des plaques dans SWOOD Nesting |

**Particularites :**
- Les dimensions (Length, Width, Thickness) sont converties de **mm en metres** (division par 1000) car SWOOD Nesting multiplie par 1000 a l'import.
- Le **cout par plaque** est calcule automatiquement : `surface_m2 x cout_euro_m2`.
- L'attribut **Path** est renseigne depuis la colonne correspondante du XLSM.
- L'attribut **SupplierReference** est renseigne depuis la Reference Fournisseur.
- Le format XML est genere en texte brut avec indentation par tabulations, identique au format de la macro VBA pour une compatibilite maximale avec l'import SWOOD.

---

### 3. XML Materiaux SWOOD

| | |
|---|---|
| **Source** | Pages Materials (49 col.) + EdgeBands (23 col.) |
| **Sortie** | `Import_Swood_Materiaux_YYYYMMDD_HHMMSS.xml` |
| **Format** | `<SWOODMat><Materials>...</Materials><EdgeBands>...</EdgeBands></SWOODMat>` |
| **Usage** | Import complet des materiaux et chants dans SWOOD |

**Particularites :**
- Reproduit **fidelement la macro VBA** du fichier Excel (SaveTextToFile).
- Lit la ligne 3 (tags de structure) et la ligne 4 (noms des attributs) pour construire dynamiquement le XML.
- Gere les blocs `<Properties>`, `<Layers>` et les attributs simples.
- Combine les 2 feuilles (Materials + EdgeBands) dans un seul fichier XML.

---

### 4. XML Chants (EdgeBands)

| | |
|---|---|
| **Source** | Page EdgeBands - 23 colonnes |
| **Sortie** | `Import_Swood_Chants_YYYYMMDD_HHMMSS.xml` |
| **Format** | `<SWOODMat><EdgeBands><EdgeBand .../></EdgeBands></SWOODMat>` |
| **Usage** | Import des chants seuls dans SWOOD |

**Particularites :**
- Meme logique que l'export Materiaux (reproduction de la macro VBA).
- Exporte uniquement la feuille EdgeBands.

---

## Utilisation

### Mode GUI (recommande)

Lancer `Export_Optiplanning.exe` (double-clic).

1. Le fichier source `Outil_Material_Import.xlsm` est detecte automatiquement s'il se trouve dans le meme dossier que l'executable.
2. Choisir un **dossier de destination** (optionnel - par defaut : meme dossier que le XLSM).
3. Cliquer sur l'un des **4 boutons d'export**.
4. Le **journal** en bas de fenetre affiche le detail de l'operation.
5. La **barre de statut** indique le resultat (vert = succes, rouge = erreur).

### Mode CLI (ligne de commande)

```bash
python export_optiplanning.py <fichier.xlsm> <type>
```

Types disponibles :

| Type | Description |
|---|---|
| `txt` | Export TXT Optiplanning |
| `nesting` | Export XML Plaques Nesting |
| `materials` | Export XML Materiaux SWOOD (Materials + EdgeBands) |
| `edgebands` | Export XML Chants seuls |

**Exemples :**
```bash
python export_optiplanning.py Outil_Material_Import.xlsm txt
python export_optiplanning.py Outil_Material_Import.xlsm nesting
python export_optiplanning.py Outil_Material_Import.xlsm materials
python export_optiplanning.py Outil_Material_Import.xlsm edgebands
```

---

## Structure du fichier Excel source

### Page Materials (49 colonnes)

| Lignes | Contenu |
|---|---|
| A1 | En-tete XML ligne 1 (`<?xml version="1.0" ...?>`) |
| A2 | En-tete XML ligne 2 (`<SWOODMat xmlns:xsd=...>`) |
| Ligne 3 | Tags de structure XML (vide, Properties, Property, Layers, Layer, etc.) |
| Ligne 4 | Noms des attributs XML (Name, Description, Path, Thickness, etc.) |
| Ligne 5+ | Donnees des materiaux |

**Colonnes principales :**

| Col. | Attribut | Exemple |
|---|---|---|
| A (1) | Name | Melamine-F186-Beton Chicago gris clair-ST9 |
| B (2) | Description | 7786359 |
| C (3) | Path | Melamine 19 mm |
| D (4) | Thickness | 19 |
| E (5) | FiberMaterial | 1 |
| F (6) | Cost (euros/m2) | 15.79 |
| AR (44) | BOARDL (mm) | 2790 |
| AS (45) | BOARDW (mm) | 2070 |
| AT (46) | Reference Fournisseur | 7786359 |
| AU (47) | Fournisseur | Dispano |
| AV (48) | Finish | 0 |
| AW (49) | Glass | 0 |

### Page EdgeBands (23 colonnes)

| Col. | Attribut | Exemple |
|---|---|---|
| A (1) | Name | F186 ST9 - 1 mm |
| B (2) | ID | 13 |
| C (3) | Description | |
| D (4) | Path | Chants 1mm |
| E (5) | Cost | 2 |
| F (6) | Reference | test |
| G (7) | Thickness | 1 |

---

## Charte graphique Destribois

L'interface utilise la charte graphique officielle Destribois :

| Element | Valeur |
|---|---|
| Couleur primaire | `#2E3544` (bleu-gris fonce) |
| Couleur secondaire | `#AE9367` (dore) |
| Fond clair | `#EDE6DC` (beige) |
| Police titres | Abhaya Libre SemiBold |
| Police corps | Roboto Regular / Medium |
| Logo | Logo_Destribois_seul.png |

Le dossier `charte_graphique/` contient les polices, le logo et les references couleurs. Ces fichiers sont embarques dans l'executable.

---

## Developpement

### Prerequis

- Python 3.8+
- `pip install openpyxl pillow pyinstaller`

### Generer l'executable

```bash
pyinstaller --onefile --windowed --name Export_Optiplanning --distpath . --clean --noconfirm --add-data "charte_graphique;charte_graphique" --add-data "Y:/01_EURL Destribois/10_Communication/01_Charte_graphique/Logo/Logo_Destribois_seul.png;." export_optiplanning.py
```

### Architecture du code

```
export_optiplanning.py
|
|-- Dataclasses
|   |-- MaterialSWOOD      (49 champs - page Materials)
|   |-- EdgeBandSWOOD       (23 champs - page EdgeBands)
|
|-- Lecture XLSM
|   |-- read_all_materials_from_xlsm()   (49 colonnes)
|   |-- read_materials_from_xlsm()       (colonnes essentielles - TXT)
|   |-- read_edgebands_from_xlsm()       (23 colonnes)
|
|-- Exports
|   |-- export_optiplanning_txt()        (Export 1 - TXT)
|   |-- export_xml_boards_nesting()      (Export 2 - XML Nesting)
|   |-- export_xml_materials()           (Export 3 - XML Materiaux)
|   |-- export_xml_edgebands()           (Export 4 - XML Chants)
|   |-- _export_vba_xml_sheet()          (Moteur XML generique - macro VBA)
|
|-- Interface GUI
|   |-- App                              (Tkinter - theme Destribois)
```

---

## Notes techniques

- Le format XML des exports 3 et 4 (Materiaux et Chants) est genere en **reproduisant fidelement la macro VBA** du fichier Excel. La ligne 3 du XLSM contient les tags de structure (`Properties`, `Layers`, etc.) et la ligne 4 contient les noms d'attributs.
- L'export Nesting utilise le meme format texte brut avec tabulations pour garantir la compatibilite avec l'import SWOOD.
- Les fichiers XML sont encodes en **UTF-8** avec retours a la ligne **CRLF** (`\r\n`).
- Le **cout par plaque** (Nesting) est calcule : `(longueur_mm / 1000) x (largeur_mm / 1000) x cout_euro_m2`.
- Les **dimensions Nesting** sont en metres (SWOOD multiplie par 1000 a l'import).
