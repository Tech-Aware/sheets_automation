"""Python representation of the Apps Script configuration constants."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Mapping, Sequence


@dataclass(frozen=True)
class HeaderGroup:
    """Container used to mimic the nested HEADERS structure."""

    entries: Mapping[str, str]

    def __getattr__(self, name: str) -> str:
        try:
            return self.entries[name]
        except KeyError as exc:  # pragma: no cover - attribute forwarding
            raise AttributeError(name) from exc

    def __iter__(self):
        return iter(self.entries.items())


HEADERS = {
    "ACHATS": HeaderGroup(
        {
            "ID": "ID",
            "ARTICLE": "ARTICLES",
            "ARTICLE_ALT": "ARTICLE",
            "MARQUE": "MARQUE",
            "DATE_ACHAT": "DATE D'ACHAT",
            "GENRE_DATA": "GENRE(data)",
            "GENRE_DATA_ALT": "GENRE(DATA)",
            "GENRE_LEGACY": "GENRE",
            "REFERENCE": "REFERENCE",
            "GRADE": "GRADE",
            "FOURNISSEUR": "FOURNISSEUR/CODE",
            "MOIS": "MOIS",
            "MOIS_NUM": "MOIS NUM",
            "DATE_LIVRAISON": "DATE DE LIVRAISON",
            "QUANTITE_RECUE": "QUANTITÉ RECUE",
            "QUANTITE_RECUE_ALT": "QUANTITE RECUE",
            "QUANTITE_COMMANDEE": "QUANTITÉ COMMANDÉE",
            "DELAI_LIVRAISON": "DELAI DE LIVRAISON",
            "PRET_STOCK": "PRËT POUR MISE EN STOCK",
            "PRET_STOCK_ALT": "PRÊT POUR MISE EN STOCK",
            "PRET_STOCK_COMBINED": "PRÊT POUR MISE EN STOCK | DATE",
            "DATE_MISE_EN_STOCK": "MIS EN STOCK LE",
            "DATE_MISE_EN_STOCK_ALT": "DATE DE MISE EN STOCK",
            "FRAIS_COLISSAGE": "FRAIS DE COLISSAGE",
            "PRIX_ACHAT_SHIP": "PRIX D'ACHAT SHIP INCLUS",
            "PRIX_UNITAIRE_TTC": "PRIX UNITAIRE TTC",
            "PRIX_UNITAIRE_BRUT": "PRIX UNITAIRE BRUTE",
            "TOTAL_TTC": "TOTAL TTC",
            "FRAIS_LAVAGE": "FRAIS DE LAVAGE",
            "TRACKING": "TRACKING",
        }
    ),
    "STOCK": HeaderGroup(
        {
            "ID": "ID",
            "SKU": "SKU",
            "REFERENCE": "REFERENCE",
            "LIBELLE": "LIBELLÉ",
            "LIBELLE_ALT": "LIBELLE",
            "ARTICLE": "ARTICLES",
            "ARTICLE_ALT": "ARTICLE",
            "MARQUE": "MARQUE",
            "PRIX_VENTE": "PRIX DE VENTE",
            "TAILLE_COLIS": "TAILLE DU COLIS",
            "TAILLE_COLIS_ALT": "TAILLE COLIS",
            "TAILLE": "TAILLE",
            "LOT": "LOT",
            "LOT_ALT": "LOTS",
            "DATE_LIVRAISON": "DATE DE LIVRAISON",
            "DATE_MISE_EN_STOCK": "DATE DE MISE EN STOCK",
            "MIS_EN_LIGNE": "MIS EN LIGNE | DATE DE MISE EN LIGNE",
            "MIS_EN_LIGNE_ALT": "MIS EN LIGNE",
            "DATE_MISE_EN_LIGNE": "MIS EN LIGNE | DATE DE MISE EN LIGNE",
            "DATE_MISE_EN_LIGNE_ALT": "DATE DE MISE EN LIGNE",
            "PUBLIE": "PUBLIÉ | DATE DE PUBLICATION",
            "PUBLIE_ALT": "PUBLIÉ",
            "DATE_PUBLICATION": "PUBLIÉ | DATE DE PUBLICATION",
            "DATE_PUBLICATION_ALT": "DATE DE PUBLICATION",
            "VENDU": "VENDU | DATE DE VENTE",
            "VENDU_ALT": "VENDU",
            "DATE_VENTE": "VENDU | DATE DE VENTE",
            "DATE_VENTE_ALT": "DATE DE VENTE",
            "VENTE_EXPORTEE_LE": "VENTE EXPORTEE LE",
            "VALIDER_SAISIE": "VALIDER LA SAISIE",
            "VALIDER_SAISIE_ALT": "VALIDER",
        }
    ),
    "VENTES": HeaderGroup(
        {
            "ID": "ID",
            "DATE_VENTE": "DATE DE VENTE",
            "ARTICLE": "ARTICLE",
            "ARTICLE_ALT": "ARTICLES",
            "SKU": "SKU",
            "PRIX_VENTE": "PRIX VENTE",
            "PRIX_VENTE_ALT": "PRIX DE VENTE",
            "FRAIS_COLISSAGE": "FRAIS DE COLISSAGE",
            "TAILLE_COLIS": "TAILLE DU COLIS",
            "TAILLE": "TAILLE",
            "LOT": "LOT",
            "DELAI_IMMOBILISATION": "DÉLAI D'IMMOBILISATION",
            "DELAI_MISE_EN_LIGNE": "DELAI DE MISE EN LIGNE",
            "DELAI_PUBLICATION": "DELAI DE PUBLICATION",
            "DELAI_VENTE": "DELAI DE VENTE",
            "RETOUR": "RETOUR",
        }
    ),
}

DEFAULT_VENTES_HEADERS: Sequence[str] = (
    HEADERS["VENTES"].ID,
    HEADERS["VENTES"].DATE_VENTE,
    HEADERS["VENTES"].ARTICLE,
    HEADERS["VENTES"].SKU,
    HEADERS["VENTES"].PRIX_VENTE,
    HEADERS["VENTES"].DELAI_IMMOBILISATION,
    HEADERS["VENTES"].DELAI_MISE_EN_LIGNE,
    HEADERS["VENTES"].DELAI_PUBLICATION,
    HEADERS["VENTES"].DELAI_VENTE,
    HEADERS["VENTES"].FRAIS_COLISSAGE,
    HEADERS["VENTES"].TAILLE,
    HEADERS["VENTES"].LOT,
)

MONTHLY_LEDGER_HEADERS: Sequence[str] = (
    "ID",
    "SKU",
    "LIBELLÉS",
    "DATE DE VENTE",
    "MARGE BRUTE",
    "COEFF MARGE",
    "NBR PCS VENDU",
)

DEFAULT_STOCK_HEADERS: Sequence[str] = (
    HEADERS["STOCK"].ID,
    HEADERS["STOCK"].SKU,
    HEADERS["STOCK"].DATE_LIVRAISON,
    HEADERS["STOCK"].DATE_MISE_EN_STOCK,
    HEADERS["STOCK"].MIS_EN_LIGNE,
    HEADERS["STOCK"].PUBLIE,
    HEADERS["STOCK"].VENDU,
    HEADERS["STOCK"].PRIX_VENTE,
    HEADERS["STOCK"].TAILLE,
    HEADERS["STOCK"].LOT_ALT,
    HEADERS["STOCK"].VALIDER_SAISIE_ALT,
)

MONTH_NAMES_FR: Sequence[str] = (
    "Janvier",
    "Février",
    "Mars",
    "Avril",
    "Mai",
    "Juin",
    "Juillet",
    "Août",
    "Septembre",
    "Octobre",
    "Novembre",
    "Décembre",
)
