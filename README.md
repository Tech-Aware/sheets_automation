# sheets_automation

Ce dépôt contenait initialement les scripts Google Apps Script pilotant le fichier Sheets/Excel « Prerelease ». Il inclut désormais un premier squelette Python/CustomTkinter qui embarque un snapshot de ce classeur (onglets Achats/Stock/Ventes/Compta) et expose une interface bureau multi-onglets sans dépendre d'Excel au démarrage.

## Application CustomTkinter
- Code localisé dans `python_app/`.
- Chargement des onglets Achats/Stock/Ventes/Compta directement depuis les données packagées (`python_app/data/workbook_snapshot.py`).
- Tableau de bord : cartes KPI calculées à l'aide de `services.summaries.build_inventory_snapshot`.
- Onglets tables : `ui.tables.ScrollableTable` encapsule un `ttk.Treeview` scrollable pour naviguer dans les données (10 premières colonnes affichées pour rester lisible). Un double-clic sur une cellule permet désormais de modifier sa valeur directement depuis l'application.
- Calendrier : liste des noms de mois français définis dans `config.MONTH_NAMES_FR`.
- Workflow Achats → Stock → Ventes : un nouvel onglet "Workflow" orchestre l'ajout d'une commande Achats, sa mise en stock, l'enregistrement de la vente (avec écriture comptable) et la gestion des retours depuis l'interface.

### Démarrage
```bash
pip install customtkinter
python -m python_app.main  # recharge les données packagées et synchronise python_app/data/achats.db
```

Vous pouvez passer un autre chemin de base Achats pour persister l'état ailleurs :
```bash
python -m python_app.main --achats-db mon_autre_fichier.db
```

## Structure Python
- `python_app/config.py` : transcription Python des constantes `config.gs` (HEADERS, noms de mois, etc.).
- `python_app/datasources/workbook.py` : dépôt pour charger ponctuellement des feuilles Excel (`TableData`).
- `python_app/data/workbook_snapshot.py` : snapshot des onglets Achats/Stock/Ventes/Compta extrait de « Prerelease 1.2.xlsx » et injecté dans l'application.
- `python_app/datasources/sqlite_purchases.py` : mini dépôt SQLite pour l'onglet Achats (création du schéma, conversion bidirectionnelle en `TableData`).
- `python_app/services/summaries.py` : calcul du snapshot inventaire (stock vs ventes, marges moyennes).
- `python_app/services/workflow.py` : réplique le flux Achats > Stock > Ventes/Compta et expose une API manipulée par l'onglet Workflow.
- `python_app/ui/tables.py` : composant table scrollable.
- `python_app/main.py` : point d'entrée CustomTkinter (`VintageErpApp`). Il persiste automatiquement l'onglet Achats dans `python_app/data/achats.db` lors de la fermeture, ce qui évite de perdre les modifications entre deux sessions.

Ce socle permet d'itérer vers une reproduction complète des workflows Apps Script dans une application desktop Python.

## Base Achats SQLite
- Les données du classeur (Achats/Stock/Ventes/Compta) sont packagées dans `python_app/data/workbook_snapshot.py` et utilisées pour initialiser automatiquement la base SQLite Achats/Stock.
- Générez localement `python_app/data/achats.db` via `python -m python_app.seed_purchases_db` : le fichier n'est pas versionné pour éviter d'inclure un binaire lourd dans Git.
- L'application charge automatiquement `python_app/data/achats.db` s'il existe (sinon elle le crée depuis le snapshot) et enregistre toutes les modifications Achats dans ce fichier lorsqu'on ferme l'application. Vous pouvez passer un autre chemin via `--achats-db` si vous souhaitez stocker l'état ailleurs.
