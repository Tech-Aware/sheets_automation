"""Utility helpers that reproduce the dashboard logic in Python."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, Mapping

from ..config import HEADERS


@dataclass
class InventorySnapshot:
    """Lightweight summary derived from Achats/Stock/Ventes tables."""

    stock_pieces: int
    stock_value: float
    unique_references: int
    reference_stock_value: float
    sold_pieces: int
    sold_value: float
    average_margin: float | None

    def as_dict(self) -> Mapping[str, float | int | None]:
        return {
            "stock_pieces": self.stock_pieces,
            "unique_references": self.unique_references,
            "stock_value": self.stock_value,
            "reference_stock_value": self.reference_stock_value,
            "sold_pieces": self.sold_pieces,
            "sold_value": self.sold_value,
            "average_margin": self.average_margin,
        }


def _safe_float(value) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _base_reference_from_stock(row: Mapping) -> str | None:
    sku = row.get(HEADERS["STOCK"].SKU)
    reference = row.get(HEADERS["STOCK"].REFERENCE)

    if isinstance(sku, str) and sku:
        return sku.split("-", 1)[0].strip() or None
    if isinstance(reference, str) and reference:
        return reference.strip() or None
    return None


def _build_reference_unit_price_index(achats_rows: Iterable[Mapping]) -> dict[str, float]:
    index: dict[str, float] = {}
    for row in achats_rows:
        reference = row.get(HEADERS["ACHATS"].REFERENCE)
        if not reference:
            continue
        quantity = _safe_float(row.get(HEADERS["ACHATS"].QUANTITE_COMMANDEE))
        total_ttc = _safe_float(row.get(HEADERS["ACHATS"].TOTAL_TTC))
        if quantity > 0:
            index[str(reference)] = total_ttc / quantity
    return index


def build_inventory_snapshot(
    stock_rows: Iterable[Mapping], ventes_rows: Iterable[Mapping], achats_rows: Iterable[Mapping] | None = None
) -> InventorySnapshot:
    """Compute the high-level KPIs from the Excel sheets."""

    reference_unit_price = _build_reference_unit_price_index(achats_rows or [])
    stock_value = 0.0
    stock_count = 0
    base_counts: dict[str, int] = {}
    for row in stock_rows:
        price = row.get(HEADERS["STOCK"].PRIX_VENTE) or 0
        vendu = row.get(HEADERS["STOCK"].VENDU_ALT) or row.get(HEADERS["STOCK"].VENDU)
        if vendu:
            continue
        stock_count += 1
        stock_value += _safe_float(price)
        base = _base_reference_from_stock(row)
        if base:
            base_counts[base] = base_counts.get(base, 0) + 1

    reference_stock_value = 0.0
    for base, count in base_counts.items():
        unit_price = reference_unit_price.get(base, 0.0)
        reference_stock_value += unit_price * count

    sold_value = 0.0
    sold_count = 0
    total_margin = 0.0
    for row in ventes_rows:
        sold_count += 1
        sold_value += _safe_float(row.get(HEADERS["VENTES"].PRIX_VENTE))
        prix = _safe_float(row.get(HEADERS["VENTES"].PRIX_VENTE))
        frais = _safe_float(row.get(HEADERS["VENTES"].FRAIS_COLISSAGE))
        total_margin += prix - frais

    average_margin = (total_margin / sold_count) if sold_count else None

    return InventorySnapshot(
        stock_pieces=stock_count,
        stock_value=round(stock_value, 2),
        unique_references=len(base_counts),
        reference_stock_value=round(reference_stock_value, 2),
        sold_pieces=sold_count,
        sold_value=round(sold_value, 2),
        average_margin=round(average_margin, 2) if average_margin is not None else None,
    )


__all__ = ["InventorySnapshot", "build_inventory_snapshot"]
