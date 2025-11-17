"""CLI helper to create the Achats SQLite database from the seed dataset."""
from __future__ import annotations

from pathlib import Path
import sys

try:  # pragma: no cover - entrypoint convenience
    from .data.workbook_snapshot import snapshot_tables
    from .datasources.sqlite_purchases import PurchaseDatabase
    from .datasources.sqlite_purchases import table_to_purchase_records, table_to_stock_records
except ImportError:  # pragma: no cover - executed when run as a script
    package_root = Path(__file__).resolve().parent.parent
    if str(package_root) not in sys.path:
        sys.path.append(str(package_root))
    from python_app.data.workbook_snapshot import snapshot_tables
    from python_app.datasources.sqlite_purchases import PurchaseDatabase
    from python_app.datasources.sqlite_purchases import table_to_purchase_records, table_to_stock_records

DEFAULT_DB = Path(__file__).resolve().parent / "data" / "achats.db"


def main(db_path: Path | None = None) -> Path:
    target = db_path or DEFAULT_DB
    db = PurchaseDatabase(target)
    tables = snapshot_tables()
    purchases = table_to_purchase_records(tables["Achats"])
    stock_rows = table_to_stock_records(tables.get("Stock"))
    db.replace_all(purchases, stock_rows)
    return target


if __name__ == "__main__":  # pragma: no cover - manual helper
    created_path = main()
    print(f"Base Achats initialisée dans {created_path}")
