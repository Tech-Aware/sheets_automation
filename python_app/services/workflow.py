"""High level helpers that reproduce the Achats → Stock → Ventes workflow."""
from __future__ import annotations

from dataclasses import dataclass as _dataclass
from datetime import date
from sys import version_info
from typing import Optional

from ..config import HEADERS
from ..datasources.workbook import TableData


# ``slots`` support for :func:`dataclasses.dataclass` only arrived in Python 3.10.
# The UI is expected to run on Python 3.9 (32-bit), so we silently drop the
# argument on older interpreters while keeping the memory optimisation available
# on newer ones.
_DATACLASS_KWARGS = {"slots": True} if version_info >= (3, 10) else {}


@_dataclass(**_DATACLASS_KWARGS)
class PurchaseInput:
    article: str
    marque: str
    reference: str
    quantite: int
    prix_unitaire: float
    frais_colissage: float = 0.0
    date_livraison: Optional[str] = None


@_dataclass(**_DATACLASS_KWARGS)
class StockInput:
    purchase_id: str
    sku: str
    prix_vente: float
    lot: str = ""
    taille: str = ""


@_dataclass(**_DATACLASS_KWARGS)
class SaleInput:
    sku: str
    prix_vente: float
    frais_colissage: float = 0.0
    date_vente: Optional[str] = None
    lot: str = ""
    taille: str = ""


class WorkflowCoordinator:
    """Mutable helper that wires the in-memory tables together."""

    def __init__(self, achats: TableData, stock: TableData, ventes: TableData, compta: TableData):
        self.achats = achats
        self.stock = stock
        self.ventes = ventes
        self.compta = compta

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def create_purchase(self, data: PurchaseInput) -> dict:
        purchase_id = str(self._next_numeric_id(self.achats.rows, HEADERS["ACHATS"].ID))
        livraison = data.date_livraison or self._today()
        row = {
            HEADERS["ACHATS"].ID: purchase_id,
            HEADERS["ACHATS"].ARTICLE: data.article,
            HEADERS["ACHATS"].MARQUE: data.marque,
            HEADERS["ACHATS"].REFERENCE: data.reference,
            HEADERS["ACHATS"].QUANTITE_RECUE: data.quantite,
            HEADERS["ACHATS"].DATE_LIVRAISON: livraison,
            HEADERS["ACHATS"].PRIX_UNITAIRE_TTC: round(data.prix_unitaire, 2),
            HEADERS["ACHATS"].FRAIS_COLISSAGE: round(data.frais_colissage, 2),
            HEADERS["ACHATS"].TOTAL_TTC: round((data.prix_unitaire * data.quantite) + data.frais_colissage, 2),
            HEADERS["ACHATS"].PRET_STOCK_COMBINED: livraison,
        }
        self.achats.rows.append(row)
        return row

    def transfer_to_stock(self, data: StockInput) -> dict:
        purchase = self._find_row(self.achats.rows, HEADERS["ACHATS"].ID, data.purchase_id)
        if purchase is None:
            raise ValueError(f"Achat {data.purchase_id} introuvable")
        stock_id = str(self._next_numeric_id(self.stock.rows, HEADERS["STOCK"].ID))
        date_stock = self._today()
        libelle = purchase.get(HEADERS["ACHATS"].ARTICLE) or purchase.get(HEADERS["ACHATS"].ARTICLE_ALT)
        row = {
            HEADERS["STOCK"].ID: stock_id,
            HEADERS["STOCK"].SKU: data.sku,
            HEADERS["STOCK"].LIBELLE: libelle,
            HEADERS["STOCK"].ARTICLE: libelle,
            HEADERS["STOCK"].PRIX_VENTE: round(data.prix_vente, 2),
            HEADERS["STOCK"].LOT: data.lot,
            HEADERS["STOCK"].TAILLE: data.taille,
            HEADERS["STOCK"].DATE_LIVRAISON: purchase.get(HEADERS["ACHATS"].DATE_LIVRAISON),
            HEADERS["STOCK"].DATE_MISE_EN_STOCK: date_stock,
        }
        self.stock.rows.append(row)
        return row

    def register_sale(self, data: SaleInput) -> dict:
        stock_row = self._find_row(self.stock.rows, HEADERS["STOCK"].SKU, data.sku)
        if stock_row is None:
            raise ValueError(f"Article {data.sku} introuvable dans le stock")
        sale_date = data.date_vente or self._today()
        stock_row[HEADERS["STOCK"].VENDU_ALT] = sale_date
        stock_row[HEADERS["STOCK"].DATE_VENTE_ALT] = sale_date
        stock_row[HEADERS["STOCK"].PRIX_VENTE] = round(data.prix_vente, 2)
        sale_id = str(self._next_numeric_id(self.ventes.rows, HEADERS["VENTES"].ID))
        libelle = stock_row.get(HEADERS["STOCK"].LIBELLE) or stock_row.get(HEADERS["STOCK"].ARTICLE)
        sale_row = {
            HEADERS["VENTES"].ID: sale_id,
            HEADERS["VENTES"].DATE_VENTE: sale_date,
            HEADERS["VENTES"].ARTICLE: libelle,
            HEADERS["VENTES"].SKU: data.sku,
            HEADERS["VENTES"].PRIX_VENTE: round(data.prix_vente, 2),
            HEADERS["VENTES"].FRAIS_COLISSAGE: round(data.frais_colissage, 2),
            HEADERS["VENTES"].TAILLE: data.taille or stock_row.get(HEADERS["STOCK"].TAILLE),
            HEADERS["VENTES"].LOT: data.lot or stock_row.get(HEADERS["STOCK"].LOT),
        }
        self.ventes.rows.append(sale_row)
        ledger_row = self._build_compta_row(sale_row)
        if ledger_row:
            self.compta.rows.append(ledger_row)
        return sale_row

    def register_return(self, sku: str, note: str) -> dict:
        sale_row = self._find_row(self.ventes.rows, HEADERS["VENTES"].SKU, sku)
        if sale_row is None:
            raise ValueError(f"Aucune vente trouvée pour le SKU {sku}")
        sale_row[HEADERS["VENTES"].RETOUR] = note or "Retour client"
        stock_row = self._find_row(self.stock.rows, HEADERS["STOCK"].SKU, sku)
        if stock_row is not None:
            stock_row[HEADERS["STOCK"].VENDU_ALT] = ""
            stock_row[HEADERS["STOCK"].DATE_VENTE_ALT] = ""
        return sale_row

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------
    @staticmethod
    def _next_numeric_id(rows: list[dict], column: str) -> int:
        max_value = 0
        for row in rows:
            try:
                value = int(row.get(column) or 0)
            except (TypeError, ValueError):
                continue
            max_value = max(max_value, value)
        return max_value + 1

    @staticmethod
    def _find_row(rows: list[dict], column: str, value) -> dict | None:
        for row in rows:
            if row.get(column) == value:
                return row
        return None

    @staticmethod
    def _today() -> str:
        return date.today().isoformat()

    def _build_compta_row(self, sale_row: dict) -> dict:
        if not self.compta.headers:
            return {}
        price = float(sale_row.get(HEADERS["VENTES"].PRIX_VENTE) or 0)
        frais = float(sale_row.get(HEADERS["VENTES"].FRAIS_COLISSAGE) or 0)
        margin = price - frais
        coeff = round(margin / price, 2) if price else 0.0
        ledger = {}
        for header in self.compta.headers:
            if header == "ID":
                ledger[header] = sale_row.get(HEADERS["VENTES"].ID)
            elif header == "SKU":
                ledger[header] = sale_row.get(HEADERS["VENTES"].SKU)
            elif header == "LIBELLÉS":
                ledger[header] = sale_row.get(HEADERS["VENTES"].ARTICLE)
            elif header == "DATE DE VENTE":
                ledger[header] = sale_row.get(HEADERS["VENTES"].DATE_VENTE)
            elif header == "MARGE BRUTE":
                ledger[header] = round(margin, 2)
            elif header == "COEFF MARGE":
                ledger[header] = coeff
            elif header == "NBR PCS VENDU":
                ledger[header] = 1
        return ledger


__all__ = [
    "PurchaseInput",
    "SaleInput",
    "StockInput",
    "WorkflowCoordinator",
]
