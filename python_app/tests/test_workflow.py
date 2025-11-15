import math
import sys
from pathlib import Path

# Ensure the repository root is on sys.path so that ``python_app`` can be imported
sys.path.append(str(Path(__file__).resolve().parents[2]))

from python_app.config import HEADERS
from python_app.datasources.workbook import TableData
from python_app.services.workflow import WorkflowCoordinator


def test_next_numeric_id_accepts_float_strings():
    rows = [
        {"ID": "1.0"},
        {"ID": 2},
        {"ID": "003"},
        {"ID": " 4 "},
    ]
    assert WorkflowCoordinator._next_numeric_id(rows, "ID") == 5


def test_next_numeric_id_ignores_invalid_values():
    rows = [
        {"ID": "abc"},
        {"ID": math.nan},
        {"ID": ""},
        {},
    ]
    assert WorkflowCoordinator._next_numeric_id(rows, "ID") == 1


def test_next_numeric_id_supports_float_instances():
    rows = [
        {"ID": 5.0},
        {"ID": "6.0"},
    ]
    assert WorkflowCoordinator._next_numeric_id(rows, "ID") == 7


def test_delete_purchases_removes_rows_and_stock():
    achats = TableData(
        headers=[HEADERS["ACHATS"].ID, HEADERS["ACHATS"].ARTICLE],
        rows=[
            {HEADERS["ACHATS"].ID: "100", HEADERS["ACHATS"].ARTICLE: "Chemise"},
            {HEADERS["ACHATS"].ID: "101", HEADERS["ACHATS"].ARTICLE: "Pantalon"},
        ],
    )
    stock = TableData(
        headers=[HEADERS["STOCK"].ID, HEADERS["STOCK"].SKU],
        rows=[
            {HEADERS["STOCK"].ID: "100", HEADERS["STOCK"].SKU: "CHE-1"},
            {HEADERS["STOCK"].ID: "999", HEADERS["STOCK"].SKU: "OLD-1"},
        ],
    )
    ventes = TableData(headers=[], rows=[])
    compta = TableData(headers=[], rows=[])
    workflow = WorkflowCoordinator(achats, stock, ventes, compta)

    removed, stock_removed = workflow.delete_purchases([0])

    assert removed == 1
    assert stock_removed == 1
    assert len(achats.rows) == 1
    assert all(row[HEADERS["STOCK"].ID] != "100" for row in stock.rows)
