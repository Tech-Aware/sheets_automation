import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parents[2]))

from python_app.config import HEADERS
from python_app.datasources.sqlite_purchases import (
    PurchaseDatabase,
    table_to_purchase_records,
    table_to_stock_records,
)
from python_app.datasources.workbook import TableData


def make_table(values):
    return TableData(headers=["ID"], rows=[{"ID": value} for value in values])


def extract_ids(values):
    table = make_table(values)
    return [record.id for record in table_to_purchase_records(table)]


def test_table_to_purchase_records_drops_blank_ids():
    ids = extract_ids(["", "   ", None, False])
    assert ids == [None, None, None, None]


def test_table_to_purchase_records_normalizes_numeric_like_ids():
    ids = extract_ids([1, 2.0, "3", "004", "5.0", " 6 "])
    assert ids == [1, 2, 3, 4, 5, 6]


def test_table_to_purchase_records_rejects_fractional_ids():
    ids = extract_ids([1.5, "2.7", "abc"])
    assert ids == [None, None, None]


def test_table_to_stock_records_handles_missing_table():
    assert table_to_stock_records(None) == []


def test_purchase_database_persists_stock_rows(tmp_path):
    db_path = tmp_path / "achats.db"
    db = PurchaseDatabase(db_path)
    achats_headers = [HEADERS["ACHATS"].ID, HEADERS["ACHATS"].ARTICLE]
    stock_headers = [HEADERS["STOCK"].ID, HEADERS["STOCK"].SKU, HEADERS["STOCK"].PRIX_VENTE]
    achats_rows = [
        {HEADERS["ACHATS"].ID: "1", HEADERS["ACHATS"].ARTICLE: "Test"},
    ]
    stock_rows = [
        {
            HEADERS["STOCK"].ID: "1",
            HEADERS["STOCK"].SKU: "SKU-1",
            HEADERS["STOCK"].PRIX_VENTE: 42.0,
        }
    ]
    achats_table = TableData(headers=achats_headers, rows=achats_rows)
    stock_table = TableData(headers=stock_headers, rows=stock_rows)
    db.replace_all(table_to_purchase_records(achats_table), table_to_stock_records(stock_table))

    loaded_stock = db.load_stock_table()
    assert loaded_stock.rows[0][HEADERS["STOCK"].SKU] == "SKU-1"
    assert loaded_stock.rows[0][HEADERS["STOCK"].PRIX_VENTE] == 42.0


def test_purchase_database_round_trip_preserves_achats_rows(tmp_path):
    db_path = tmp_path / "achats.db"
    db = PurchaseDatabase(db_path)
    achats_headers = [HEADERS["ACHATS"].ID, HEADERS["ACHATS"].ARTICLE]
    achats_rows = [
        {HEADERS["ACHATS"].ID: "10", HEADERS["ACHATS"].ARTICLE: "Pantalon"},
        {HEADERS["ACHATS"].ID: "11", HEADERS["ACHATS"].ARTICLE: "Veste"},
    ]
    achats_table = TableData(headers=achats_headers, rows=achats_rows)

    db.replace_all(table_to_purchase_records(achats_table), [])

    loaded = db.load_table()
    assert len(loaded.rows) == 2
    assert loaded.rows[0][HEADERS["ACHATS"].ID] == 10
    assert loaded.rows[1][HEADERS["ACHATS"].ARTICLE] == "Veste"
