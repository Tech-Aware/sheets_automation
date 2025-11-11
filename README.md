# sheets_automation

Ce dépôt contient désormais deux approches complémentaires :

* Le script Google Apps Script historique (`onEdit_Main.gs`).
* Une application bureau Python/Tkinter (`inventory_app`) qui reprend les métriques clés
  des feuilles Achats/Stock/Ventes et les présente dans une interface fluide.

## Pré-requis

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Lancer l'application bureau

```bash
python -m inventory_app.app
```

Par défaut, l'application charge le classeur d'exemple `Copie de TEST (3).xlsx`. Vous
pouvez préciser un autre fichier :

```bash
python -m inventory_app.app /chemin/vers/votre_fichier.xlsx
```

L'interface offre :

* Un tableau de bord synthétique (valeur de stock, achats cumulés, ventes, délai moyen).
* Des onglets Achats, Stock et Ventes avec filtrage instantané et vues tabulaires modernes.
* Un module de reporting qui met en avant les articles générant le plus de valeur, la
  répartition par lots et les alertes opérationnelles (achats en attente, articles sans prix).

La lecture du classeur est effectuée via `openpyxl` en mode `data_only`, ce qui permet de
récupérer les valeurs calculées sans dépendre des fonctions personnalisées Google Sheets.
