# sheets_automation

## Retours de vente instantanés

La colonne **RETOUR** de la feuille `Ventes` déclenche désormais tout le
workflow de reprise de stock :

1. Dès qu'une case est cochée, le script vérifie la ligne puis supprime
   automatiquement la vente du journal comptable mensuel correspondant.
2. Si la suppression est validée, la ligne `Ventes` est effacée et une
   nouvelle ligne est ajoutée dans la feuille `Stock` avec l'ID, le SKU,
   le libellé, le prix et les informations de lot/taille disponibles.
   Les colonnes statutaires (mise en stock, mise en ligne, publication,
   vendu, etc.) sont remises à zéro et la date de retour est inscrite dans
   la colonne « DATE DE MISE EN STOCK » pour tracer l'opération.
3. En cas d'échec (journal comptable introuvable, ligne déjà retraitée,
   etc.), la case est automatiquement décochée et un toast précise la
   raison pour éviter tout retrait multiple.

Il n'est donc plus nécessaire de passer par des menus manuels : cocher la
colonne **RETOUR** suffit pour remettre un article en stock immédiatement.
