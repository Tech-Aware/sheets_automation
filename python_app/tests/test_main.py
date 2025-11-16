import pytest

pytest.importorskip("customtkinter")

from python_app.config import HEADERS
from python_app.main import StockSummaryPanel


def test_stock_summary_counts_base_reference_from_sku():
    rows = [
        {
            HEADERS["STOCK"].SKU: "ABC-1",
            HEADERS["STOCK"].PRIX_VENTE: "10",
            HEADERS["STOCK"].VENDU: "",
        },
        {
            HEADERS["STOCK"].SKU: "ABC-2",
            HEADERS["STOCK"].PRIX_VENTE: "15",
            HEADERS["STOCK"].VENDU_ALT: None,
        },
        {  # should count as a distinct reference even without a SKU
            HEADERS["STOCK"].REFERENCE: "XYZ",
            HEADERS["STOCK"].PRIX_VENTE: "20",
            HEADERS["STOCK"].VENDU: None,
        },
    ]

    stats = StockSummaryPanel._compute_stats(rows)

    assert stats["reference_count"] == 2
    assert stats["stock_value"] == 45.0
    assert stats["value_per_reference"] == 22.5
