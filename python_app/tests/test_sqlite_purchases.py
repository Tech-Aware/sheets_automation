import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parents[2]))

from python_app.datasources.sqlite_purchases import table_to_purchase_records
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
