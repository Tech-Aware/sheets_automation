# sheets_automation

Ce dépôt contenait initialement les scripts Google Apps Script pilotant le fichier Sheets/Excel « Prerelease ». Il inclut désormais un premier squelette Python/CustomTkinter qui charge ce fichier Excel et expose une interface bureau multi-onglets.

## Application CustomTkinter
- Code localisé dans `python_app/`.
- Chargement des onglets Achats/Stock/Ventes/Compta directement depuis `Prerelease 1.2.xlsx` via `openpyxl`.
- Tableau de bord : cartes KPI calculées à l'aide de `services.summaries.build_inventory_snapshot`.
- Onglets tables : `ui.tables.ScrollableTable` encapsule un `ttk.Treeview` scrollable pour naviguer dans les données (10 premières colonnes affichées pour rester lisible). Un double-clic sur une cellule permet désormais de modifier sa valeur directement depuis l'application.
- Calendrier : liste des noms de mois français définis dans `config.MONTH_NAMES_FR`.
- Workflow Achats → Stock → Ventes : un nouvel onglet "Workflow" orchestre l'ajout d'une commande Achats, sa mise en stock, l'enregistrement de la vente (avec écriture comptable) et la gestion des retours depuis l'interface.

### Démarrage
```bash
pip install customtkinter openpyxl
python -m python_app.main  # charge automatiquement « Prerelease 1.2.xlsx » depuis la racine du dépôt
```

Vous pouvez passer un autre chemin vers un fichier Excel compatible ou remplacer l'onglet Achats par une base SQLite générée localement :
```bash
python -m python_app.seed_purchases_db  # produit python_app/data/achats.db
python -m python_app.main /chemin/vers/mon_fichier.xlsx --achats-db python_app/data/achats.db
```

## Structure Python
- `python_app/config.py` : transcription Python des constantes `config.gs` (HEADERS, noms de mois, etc.).
- `python_app/datasources/workbook.py` : dépôt pour charger les feuilles Excel (`TableData`).
- `python_app/datasources/sqlite_purchases.py` : mini dépôt SQLite pour l'onglet Achats (création du schéma, conversion bidirectionnelle en `TableData`).
- `python_app/services/summaries.py` : calcul du snapshot inventaire (stock vs ventes, marges moyennes).
- `python_app/services/workflow.py` : réplique le flux Achats > Stock > Ventes/Compta et expose une API manipulée par l'onglet Workflow.
- `python_app/ui/tables.py` : composant table scrollable.
- `python_app/main.py` : point d'entrée CustomTkinter (`VintageErpApp`). Il persiste automatiquement l'onglet Achats dans `python_app/data/achats.db` lors de la fermeture, ce qui évite de perdre les modifications entre deux sessions. Tant que cette base reste vide, l'application continue de charger Achats depuis le classeur Excel afin d'éviter d'afficher un onglet vide au premier lancement.

Ce socle permet d'itérer vers une reproduction complète des workflows Apps Script dans une application desktop Python.

## Base Achats SQLite
- Les données fournies dans la feuille Achats ont été converties en fixtures `python_app/data/purchases_seed.py` (texte).
- Générez localement `python_app/data/achats.db` via `python -m python_app.seed_purchases_db` : le fichier n'est pas versionné pour éviter d'inclure un binaire lourd dans Git.
- L'application charge automatiquement `python_app/data/achats.db` s'il existe et enregistre toutes les modifications Achats dans ce fichier lorsqu'on ferme l'application (le fichier est créé au besoin). Vous pouvez passer un autre chemin via `--achats-db` si vous souhaitez stocker l'état ailleurs.
