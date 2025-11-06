# sheets_automation

## Gestion des validations "Stock"

Les colonnes **TAILLE DE COLIS** et **LOT** reçoivent automatiquement les listes
déroulantes définies sur les premières lignes de la feuille *Stock*. Lors de
l'édition, le script conserve toutes les validations existantes : seules les
cellules sans menu ou avec une règle incompatible sont mises à jour, ce qui
évite d'écraser les menus personnalisés configurés manuellement par les
utilisateurs.
