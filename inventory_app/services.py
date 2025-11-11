"""Business logic layer bridging workbook data and the UI."""
from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from typing import Dict, Iterable, List, Sequence

from .models import DashboardSnapshot, PurchaseRecord, SaleRecord, StockRecord, summarise


@dataclass(slots=True)
class ReportRow:
    label: str
    value: str


@dataclass(slots=True)
class ReportingBundle:
    dashboard: DashboardSnapshot
    top_products: List[ReportRow]
    lots: List[ReportRow]
    alerts: List[str]


def build_reporting(purchases: Sequence[PurchaseRecord], stock: Sequence[StockRecord], sales: Sequence[SaleRecord]) -> ReportingBundle:
    dashboard = summarise(purchases, stock, sales)

    top_products = _top_stock_by_value(stock)
    lots = _group_lots(stock)
    alerts = _collect_alerts(purchases, stock)

    return ReportingBundle(dashboard=dashboard, top_products=top_products, lots=lots, alerts=alerts)


def _top_stock_by_value(stock: Sequence[StockRecord]) -> List[ReportRow]:
    totals: Dict[str, float] = defaultdict(float)
    for item in stock:
        if item.disponibilite.lower() == "vendu":
            continue
        totals[item.designation] += item.prix_vente
    top = sorted(totals.items(), key=lambda pair: pair[1], reverse=True)[:5]
    return [ReportRow(label=label, value=f"{value:,.2f} €") for label, value in top]


def _group_lots(stock: Sequence[StockRecord]) -> List[ReportRow]:
    totals: Dict[str, int] = defaultdict(int)
    for item in stock:
        lot = item.lot or "(non défini)"
        totals[lot] += 1
    ordered = sorted(totals.items(), key=lambda pair: pair[1], reverse=True)
    return [ReportRow(label=label, value=str(value)) for label, value in ordered]


def _collect_alerts(purchases: Sequence[PurchaseRecord], stock: Sequence[StockRecord]) -> List[str]:
    alerts: List[str] = []

    no_price = [item.sku for item in stock if item.disponibilite.lower() == "disponible" and item.prix_vente == 0]
    if no_price:
        alerts.append(
            "Articles disponibles sans prix défini : " + ", ".join(no_price[:5]) + ("…" if len(no_price) > 5 else "")
        )

    pending = [purchase.id_achat for purchase in purchases if purchase.statut.lower() in {"en cours", "à préparer"}]
    if pending:
        alerts.append(
            "Achats à finaliser : " + ", ".join(pending[:5]) + ("…" if len(pending) > 5 else "")
        )

    return alerts


def filter_records(records: Iterable, search: str, *attributes: str):
    """Yield records for which one of the ``attributes`` matches ``search``."""

    needle = search.lower().strip()
    if not needle:
        yield from records
        return

    for record in records:
        for attribute in attributes:
            value = getattr(record, attribute, "")
            if value is None:
                continue
            if needle in str(value).lower():
                yield record
                break
