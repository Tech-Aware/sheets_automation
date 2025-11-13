# sheets_automation

Ce dépôt contenait initialement les scripts Google Apps Script pilotant le fichier Sheets/Excel « Prerelease ». Il inclut désormais un premier squelette Python/CustomTkinter qui charge ce fichier Excel et expose une interface bureau multi-onglets.

## Application CustomTkinter
- Code localisé dans `python_app/`.
- Chargement des onglets Achats/Stock/Ventes/Compta directement depuis `Prerelease 1.2.xlsx` via `openpyxl`.
- Tableau de bord : cartes KPI calculées à l'aide de `services.summaries.build_inventory_snapshot`.
- Onglets tables : `ui.tables.ScrollableTable` encapsule un `ttk.Treeview` scrollable pour naviguer dans les données (10 premières colonnes affichées pour rester lisible).
- Calendrier : liste des noms de mois français définis dans `config.MONTH_NAMES_FR`.

### Démarrage
```bash
pip install customtkinter openpyxl
python -m python_app.main  # charge automatiquement « Prerelease 1.2.xlsx » depuis la racine du dépôt
```

Vous pouvez passer un autre chemin vers un fichier Excel compatible :
```bash
python -m python_app.main /chemin/vers/mon_fichier.xlsx
```

## Structure Python
- `python_app/config.py` : transcription Python des constantes `config.gs` (HEADERS, noms de mois, etc.).
- `python_app/datasources/workbook.py` : dépôt pour charger les feuilles Excel (`TableData`).
- `python_app/services/summaries.py` : calcul du snapshot inventaire (stock vs ventes, marges moyennes).
- `python_app/ui/tables.py` : composant table scrollable.
- `python_app/main.py` : point d'entrée CustomTkinter (`VintageErpApp`).

Ce socle permet d'itérer vers une reproduction complète des workflows Apps Script dans une application desktop Python.
