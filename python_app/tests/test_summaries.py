from python_app.config import HEADERS
from python_app.services.summaries import InventoryCache, build_inventory_snapshot


def test_inventory_snapshot_counts_unique_references_and_values():
    stock_rows = [
        {  # two SKUs share the same base "ABC"
            HEADERS["STOCK"].SKU: "ABC-1",
            HEADERS["STOCK"].PRIX_VENTE: "12",
            HEADERS["STOCK"].VENDU: "",
        },
        {
            HEADERS["STOCK"].SKU: "ABC-2",
            HEADERS["STOCK"].PRIX_VENTE: "15",
            HEADERS["STOCK"].VENDU: None,
        },
        {  # sold item should be ignored
            HEADERS["STOCK"].SKU: "DEF-1",
            HEADERS["STOCK"].PRIX_VENTE: "9",
            HEADERS["STOCK"].VENDU: "01/01/2024",
        },
        {  # falls back to the reference column when SKU is missing
            HEADERS["STOCK"].REFERENCE: "XYZ",
            HEADERS["STOCK"].PRIX_VENTE: "5",
            HEADERS["STOCK"].VENDU: "",
        },
    ]

    ventes_rows = []

    achats_rows = [
        {
            HEADERS["ACHATS"].REFERENCE: "ABC",
            HEADERS["ACHATS"].TOTAL_TTC: 200,
            HEADERS["ACHATS"].QUANTITE_COMMANDEE: 4,
        },
        {
            HEADERS["ACHATS"].REFERENCE: "XYZ",
            HEADERS["ACHATS"].TOTAL_TTC: 50,
            HEADERS["ACHATS"].QUANTITE_COMMANDEE: 5,
        },
    ]

    snapshot = build_inventory_snapshot(stock_rows, ventes_rows, achats_rows)

    assert snapshot.stock_pieces == 3
    assert snapshot.unique_references == 2
    assert snapshot.stock_value == 32.0  # 12 + 15 + 5
    assert snapshot.reference_stock_value == 110.0  # (200/4 * 2 ABC) + (50/5 * 1 XYZ)


def test_inventory_cache_tracks_incremental_updates():
    stock_rows = [
        {HEADERS["STOCK"].SKU: "ABC-1", HEADERS["STOCK"].PRIX_VENTE: 10, HEADERS["STOCK"].VENDU: ""}
    ]
    ventes_rows = []
    achats_rows = [
        {HEADERS["ACHATS"].REFERENCE: "ABC", HEADERS["ACHATS"].TOTAL_TTC: 50, HEADERS["ACHATS"].QUANTITE_COMMANDEE: 5}
    ]

    cache = InventoryCache.from_tables(stock_rows, ventes_rows, achats_rows)

    new_stock = {HEADERS["STOCK"].SKU: "ABC-2", HEADERS["STOCK"].PRIX_VENTE: 20, HEADERS["STOCK"].VENDU: ""}
    cache.on_stock_added(new_stock)
    assert cache.snapshot().stock_pieces == 2

    cache.on_stock_sold(new_stock, sale_price=25, frais=5)
    sold_snapshot = cache.snapshot()
    assert sold_snapshot.stock_pieces == 1
    assert sold_snapshot.sold_pieces == 1
    assert sold_snapshot.sold_value == 25.0
    assert sold_snapshot.average_margin == 20.0

    cache.on_stock_return(new_stock, was_sold=True)
    returned = cache.snapshot()
    assert returned.stock_pieces == 2
