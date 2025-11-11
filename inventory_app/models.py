"""Domain models representing the different workbook entities."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Any, Dict, Iterable, List, Optional

from inventory_app.data_loader import SheetData


@dataclass(slots=True)
class PurchaseRecord:
    id_achat: str
    designation: str
    quantite: float
    prix_unitaire_ttc: float
    total_ttc: float
    date_achat: Optional[datetime]
    fournisseur: Optional[str]
    statut: str

    @classmethod
    def from_row(cls, row: Dict[str, Any]) -> "PurchaseRecord":
        return cls(
            id_achat=str(row.get("ID ACHAT", "")),
            designation=str(row.get("DESIGNATION", "")),
            quantite=_to_float(row.get("QUANTITE", 0)),
            prix_unitaire_ttc=_to_float(row.get("PRIX UNITAIRE TTC", 0)),
            total_ttc=_to_float(row.get("TOTAL TTC", 0)),
            date_achat=_to_datetime(row.get("DATE D'ACHAT")),
            fournisseur=_optional_str(row.get("FOURNISSEUR")),
            statut=_optional_str(row.get("STATUT")) or "Inconnu",
        )


@dataclass(slots=True)
class StockRecord:
    sku: str
    designation: str
    disponibilite: str
    prix_vente: float
    date_publication: Optional[datetime]
    date_vente: Optional[datetime]
    lot: Optional[str]

    @classmethod
    def from_row(cls, row: Dict[str, Any]) -> "StockRecord":
        return cls(
            sku=str(row.get("SKU", "")),
            designation=_optional_str(row.get("DESIGNATION")) or "",
            disponibilite=_optional_str(row.get("STATUT")) or "Inconnu",
            prix_vente=_to_float(row.get("PRIX", 0)),
            date_publication=_to_datetime(row.get("DATE PUBLICATION")),
            date_vente=_to_datetime(row.get("DATE VENTE")),
            lot=_optional_str(row.get("LOT")),
        )


@dataclass(slots=True)
class SaleRecord:
    sku: str
    designation: str
    date_vente: Optional[datetime]
    prix_vente: float
    frais_port: float
    marge: float
    delai: Optional[int]

    @classmethod
    def from_row(cls, row: Dict[str, Any]) -> "SaleRecord":
        prix = _to_float(row.get("PRIX", 0))
        frais = _to_float(row.get("FRAIS", 0))
        return cls(
            sku=_optional_str(row.get("SKU")) or "",
            designation=_optional_str(row.get("ARTICLE")) or "",
            date_vente=_to_datetime(row.get("DATE VENTE")),
            prix_vente=prix,
            frais_port=frais,
            marge=prix - frais,
            delai=_to_int(row.get("DELAI")),
        )


@dataclass(slots=True)
class DashboardSnapshot:
    total_stock_value: float
    total_purchases_value: float
    total_sales_value: float
    average_delay: float
    available_stock: int
    pending_purchases: int


def build_purchases(sheet: SheetData) -> List[PurchaseRecord]:
    return [PurchaseRecord.from_row(row) for row in sheet.rows]


def build_stock(sheet: SheetData) -> List[StockRecord]:
    return [StockRecord.from_row(row) for row in sheet.rows]


def build_sales(sheet: SheetData) -> List[SaleRecord]:
    return [SaleRecord.from_row(row) for row in sheet.rows]


def summarise(purchases: Iterable[PurchaseRecord], stock: Iterable[StockRecord], sales: Iterable[SaleRecord]) -> DashboardSnapshot:
    purchases_list = list(purchases)
    stock_list = list(stock)
    sales_list = list(sales)

    total_stock_value = sum(item.prix_vente for item in stock_list if item.disponibilite.lower() != "vendu")
    total_purchases_value = sum(item.total_ttc for item in purchases_list)
    total_sales_value = sum(item.prix_vente for item in sales_list)

    delays = [sale.delai for sale in sales_list if sale.delai is not None]
    average_delay = (sum(delays) / len(delays)) if delays else 0.0

    available_stock = sum(1 for item in stock_list if item.disponibilite.lower() == "disponible")
    pending_purchases = sum(1 for item in purchases_list if item.statut.lower() not in {"recu", "livre"})

    return DashboardSnapshot(
        total_stock_value=total_stock_value,
        total_purchases_value=total_purchases_value,
        total_sales_value=total_sales_value,
        average_delay=average_delay,
        available_stock=available_stock,
        pending_purchases=pending_purchases,
    )


def _to_float(value: Any) -> float:
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0


def _to_int(value: Any) -> Optional[int]:
    try:
        if value is None or value == "":
            return None
        return int(value)
    except (TypeError, ValueError):
        return None


def _optional_str(value: Any) -> Optional[str]:
    if value in (None, ""):
        return None
    return str(value)


def _to_datetime(value: Any) -> Optional[datetime]:
    if isinstance(value, datetime):
        return value
    return None
