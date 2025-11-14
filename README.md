# sheets_automation

## Workflow des retours

- Chaque export depuis la feuille **Stock** crée une ligne dans **Ventes** avec la colonne `RETOUR` décochée.
- Dès qu'une cellule de la colonne `RETOUR` est cochée dans **Ventes**, l'automatisation `handleVentesReturn` est déclenchée automatiquement par `onEdit`.
- Le script supprime l'écriture comptable correspondante dans la feuille mensuelle, recrée l'article dans **Stock** avec une nouvelle date de mise en stock (date du clic) et réinitialise les statuts `MIS EN LIGNE`, `PUBLIÉ`, `VENDU`.
- Lorsque la remise en stock est confirmée, la ligne initiale est supprimée de **Ventes** et un toast confirme le retour.
- Si la compta ne peut pas être ajustée ou si la remise en stock échoue, la case `RETOUR` est automatiquement réinitialisée et aucune donnée n'est supprimée.

## Menus utiles

- **Actions Ventes → Recopier les ventes en compta** reste disponible pour recalculer entièrement les feuilles mensuelles si nécessaire (par exemple après modification historique).
- Les actions groupées de la feuille **Stock** (publication, mise en ligne, etc.) continuent de fonctionner comme auparavant; la remise en stock automatique ne modifie que les lignes marquées comme retour dans **Ventes**.
