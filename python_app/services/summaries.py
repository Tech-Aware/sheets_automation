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


class InventoryCache:
    """Incremental cache for inventory KPIs.

    This cache keeps lightweight aggregates in memory so that callers can
    refresh the dashboard without rebuilding the whole snapshot from scratch.
    """

    def __init__(
        self,
        *,
        reference_totals: dict[str, tuple[float, float]],
        stock_count: int,
        stock_value: float,
        base_counts: dict[str, int],
        sold_count: int,
        sold_value: float,
        total_margin: float,
    ) -> None:
        self._reference_totals = reference_totals
        self._stock_count = stock_count
        self._stock_value = stock_value
        self._base_counts = base_counts
        self._sold_count = sold_count
        self._sold_value = sold_value
        self._total_margin = total_margin

    @classmethod
    def from_tables(
        cls,
        stock_rows: Iterable[Mapping],
        ventes_rows: Iterable[Mapping],
        achats_rows: Iterable[Mapping] | None = None,
    ) -> "InventoryCache":
        reference_totals = _build_reference_totals(achats_rows or [])
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

        sold_value = 0.0
        sold_count = 0
        total_margin = 0.0
        for row in ventes_rows:
            sold_count += 1
            sold_value += _safe_float(row.get(HEADERS["VENTES"].PRIX_VENTE))
            prix = _safe_float(row.get(HEADERS["VENTES"].PRIX_VENTE))
            frais = _safe_float(row.get(HEADERS["VENTES"].FRAIS_COLISSAGE))
            total_margin += prix - frais

        return cls(
            reference_totals=reference_totals,
            stock_count=stock_count,
            stock_value=stock_value,
            base_counts=base_counts,
            sold_count=sold_count,
            sold_value=sold_value,
            total_margin=total_margin,
        )

    def snapshot(self) -> InventorySnapshot:
        reference_unit_price = _unit_price_index(self._reference_totals)
        reference_stock_value = 0.0
        for base, count in self._base_counts.items():
            unit_price = reference_unit_price.get(base, 0.0)
            reference_stock_value += unit_price * count

        average_margin = (self._total_margin / self._sold_count) if self._sold_count else None

        return InventorySnapshot(
            stock_pieces=self._stock_count,
            stock_value=round(self._stock_value, 2),
            unique_references=len(self._base_counts),
            reference_stock_value=round(reference_stock_value, 2),
            sold_pieces=self._sold_count,
            sold_value=round(self._sold_value, 2),
            average_margin=round(average_margin, 2) if average_margin is not None else None,
        )

    def on_purchase_added(self, row: Mapping) -> None:
        reference = row.get(HEADERS["ACHATS"].REFERENCE)
        if not reference:
            return
        quantity = _safe_float(row.get(HEADERS["ACHATS"].QUANTITE_COMMANDEE))
        total_ttc = _safe_float(row.get(HEADERS["ACHATS"].TOTAL_TTC))
        if quantity <= 0:
            return
        total, qty = self._reference_totals.get(str(reference), (0.0, 0.0))
        self._reference_totals[str(reference)] = (total + total_ttc, qty + quantity)

    def on_purchase_removed(self, row: Mapping) -> None:
        reference = row.get(HEADERS["ACHATS"].REFERENCE)
        if not reference:
            return
        quantity = _safe_float(row.get(HEADERS["ACHATS"].QUANTITE_COMMANDEE))
        total_ttc = _safe_float(row.get(HEADERS["ACHATS"].TOTAL_TTC))
        if quantity <= 0:
            return
        total, qty = self._reference_totals.get(str(reference), (0.0, 0.0))
        new_total = total - total_ttc
        new_qty = qty - quantity
        if new_qty <= 0 or new_total <= 0:
            self._reference_totals.pop(str(reference), None)
        else:
            self._reference_totals[str(reference)] = (new_total, new_qty)

    def on_stock_added(self, row: Mapping) -> None:
        vendu = row.get(HEADERS["STOCK"].VENDU_ALT) or row.get(HEADERS["STOCK"].VENDU)
        if vendu:
            return
        self._stock_count += 1
        self._stock_value += _safe_float(row.get(HEADERS["STOCK"].PRIX_VENTE))
        base = _base_reference_from_stock(row)
        if base:
            self._base_counts[base] = self._base_counts.get(base, 0) + 1

    def on_stock_removed(self, row: Mapping) -> None:
        vendu = row.get(HEADERS["STOCK"].VENDU_ALT) or row.get(HEADERS["STOCK"].VENDU)
        if vendu:
            return
        self._stock_count -= 1
        self._stock_value -= _safe_float(row.get(HEADERS["STOCK"].PRIX_VENTE))
        base = _base_reference_from_stock(row)
        if base and base in self._base_counts:
            self._base_counts[base] = max(0, self._base_counts.get(base, 0) - 1)
            if self._base_counts[base] == 0:
                del self._base_counts[base]

    def on_stock_sold(
        self,
        row: Mapping,
        *,
        sale_price: float,
        frais: float = 0.0,
        was_sold: bool = False,
    ) -> None:
        price_value = _safe_float(sale_price)
        if not was_sold:
            self.on_stock_removed(row)
        self._sold_count += 1
        self._sold_value += price_value
        self._total_margin += price_value - _safe_float(frais)

    def on_stock_return(self, row: Mapping, *, was_sold: bool) -> None:
        if not was_sold:
            return
        self.on_stock_added(row)


def _build_reference_totals(achats_rows: Iterable[Mapping]) -> dict[str, tuple[float, float]]:
    totals: dict[str, tuple[float, float]] = {}
    for row in achats_rows:
        reference = row.get(HEADERS["ACHATS"].REFERENCE)
        if not reference:
            continue
        quantity = _safe_float(row.get(HEADERS["ACHATS"].QUANTITE_COMMANDEE))
        total_ttc = _safe_float(row.get(HEADERS["ACHATS"].TOTAL_TTC))
        if quantity <= 0:
            continue
        prev_total, prev_qty = totals.get(str(reference), (0.0, 0.0))
        totals[str(reference)] = (prev_total + total_ttc, prev_qty + quantity)
    return totals


def _unit_price_index(reference_totals: Mapping[str, tuple[float, float]]) -> dict[str, float]:
    index: dict[str, float] = {}
    for reference, (total_ttc, quantity) in reference_totals.items():
        if quantity > 0:
            index[reference] = total_ttc / quantity
    return index


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
    totals = _build_reference_totals(achats_rows)
    return _unit_price_index(totals)


def build_inventory_snapshot(
    stock_rows: Iterable[Mapping], ventes_rows: Iterable[Mapping], achats_rows: Iterable[Mapping] | None = None
) -> InventorySnapshot:
    """Compute the high-level KPIs from the Excel sheets."""
    cache = InventoryCache.from_tables(stock_rows, ventes_rows, achats_rows)
    return cache.snapshot()


__all__ = ["InventoryCache", "InventorySnapshot", "build_inventory_snapshot"]
