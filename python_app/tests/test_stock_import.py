from python_app.datasources.workbook import TableData
from python_app.services.stock_import import merge_stock_table


def make_table(headers, rows):
    return TableData(headers=headers, rows=rows)


def test_merge_stock_table_adds_only_new_rows():
    headers = ("ID", "SKU", "LIBELLÉ", "LOT")
    target = make_table(headers, [
        {"ID": "100", "SKU": "SKU-1", "LIBELLÉ": "Alpha", "LOT": "A"},
    ])
    source = make_table(headers, [
        {"ID": "100", "SKU": "SKU-1", "LIBELLÉ": "Alpha", "LOT": "A"},
        {"ID": "101", "SKU": "SKU-2", "LIBELLÉ": "Bravo", "LOT": "B"},
        {"ID": "", "SKU": "SKU-3", "LIBELLÉ": "Charlie", "LOT": "C"},
    ])

    added = merge_stock_table(target, source)

    assert added == 2
    assert len(target.rows) == 3
    assert any(row["SKU"] == "SKU-2" for row in target.rows)
    assert any(row["SKU"] == "SKU-3" for row in target.rows)


def test_merge_stock_table_skips_rows_without_identifiers():
    headers = ("ID", "SKU", "LIBELLÉ")
    target = make_table(headers, [])
    source = make_table(headers, [
        {"ID": "", "SKU": "", "LIBELLÉ": "Sans ID"},
    ])

    added = merge_stock_table(target, source)

    assert added == 0
    assert target.rows == []


def test_merge_stock_table_handles_numeric_signatures():
    headers = ("ID", "SKU", "LIBELLÉ")
    target = make_table(headers, [{"ID": "200", "SKU": "SKU-5", "LIBELLÉ": "Delta"}])
    source = make_table(headers, [
        {"ID": 200.0, "SKU": "SKU-5", "LIBELLÉ": "Duplicate"},
        {"ID": 201, "SKU": "SKU-6", "LIBELLÉ": "Echo"},
    ])

    added = merge_stock_table(target, source)

    assert added == 1
    assert any(row["ID"] == 201 for row in target.rows)
