"""CLI helper to create the Achats SQLite database from the seed dataset."""
from __future__ import annotations

from pathlib import Path
import sys

try:  # pragma: no cover - entrypoint convenience
    from .data.purchases_seed import SEED_PURCHASES
    from .datasources.sqlite_purchases import PurchaseDatabase, PurchaseRecord
except ImportError:  # pragma: no cover - executed when run as a script
    package_root = Path(__file__).resolve().parent.parent
    if str(package_root) not in sys.path:
        sys.path.append(str(package_root))
    from python_app.data.purchases_seed import SEED_PURCHASES
    from python_app.datasources.sqlite_purchases import PurchaseDatabase, PurchaseRecord

DEFAULT_DB = Path(__file__).resolve().parent / "data" / "achats.db"


def main(db_path: Path | None = None) -> Path:
    target = db_path or DEFAULT_DB
    db = PurchaseDatabase(target)
    records = [PurchaseRecord.from_mapping(entry) for entry in SEED_PURCHASES]
    db.replace_all(records)
    return target


if __name__ == "__main__":  # pragma: no cover - manual helper
    created_path = main()
    print(f"Base Achats initialisée dans {created_path}")
