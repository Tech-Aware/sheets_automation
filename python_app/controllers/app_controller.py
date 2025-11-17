from __future__ import annotations

from pathlib import Path
from tkinter import messagebox
from typing import Mapping

from ..config import HEADERS
from ..data.workbook_snapshot import snapshot_tables
from ..datasources.sqlite_purchases import PurchaseDatabase, table_to_purchase_records, table_to_stock_records
from ..datasources.workbook import TableData
from ..services.workflow import WorkflowCoordinator
from ..ui.app import VintageErpApp


# Base SQLite par défaut situées dans python_app/data
DEFAULT_ACHATS_DB = Path(__file__).resolve().parent.parent / "data" / "achats.db"


class AppController:
    """Bootstrapper responsible for wiring data sources, workflow and UI."""

    def __init__(self, achats_db_path: Path | None = None):
        self.custom_db_requested = achats_db_path is not None
        self.achats_db_path = Path(achats_db_path) if achats_db_path is not None else DEFAULT_ACHATS_DB

    def run(self) -> int:
        tables = snapshot_tables()
        self._seed_default_database(tables)
        try:
            achats_table, stock_table = self._load_persisted_tables()
        except FileNotFoundError:
            messagebox.showerror("Base Achats introuvable", f"Impossible d'ouvrir {self.achats_db_path!s}")
            return 1
        self._apply_overrides(tables, achats_table, stock_table)
        self._normalize_tables(tables)

        workflow = WorkflowCoordinator(
            tables["Achats"],
            tables["Stock"],
            tables["Ventes"],
            tables["Compta 09-2025"],
        )
        app = VintageErpApp(tables, workflow, achats_db_path=self.achats_db_path)
        app.mainloop()
        return 0

    def _load_persisted_tables(self) -> tuple[TableData | None, TableData | None]:
        achats_table: TableData | None = None
        stock_table: TableData | None = None
        if self.achats_db_path.exists():
            db = PurchaseDatabase(self.achats_db_path)
            loaded_table = db.load_table()
            if loaded_table.rows:
                achats_table = loaded_table
            loaded_stock = db.load_stock_table()
            if loaded_stock.rows:
                stock_table = loaded_stock
        elif self.custom_db_requested:
            raise FileNotFoundError(self.achats_db_path)
        return achats_table, stock_table

    def _seed_default_database(self, tables: Mapping[str, TableData]) -> None:
        """Ensure the packaged workbook data is present in the SQLite file."""

        if self.achats_db_path is None:
            return
        if self.custom_db_requested and not self.achats_db_path.exists():
            # Custom path must be provided by the user.
            return
        db = PurchaseDatabase(self.achats_db_path)
        if self.achats_db_path.exists():
            existing_achats = db.load_table()
            existing_stock = db.load_stock_table()
            if existing_achats.rows or existing_stock.rows:
                return
        purchases = table_to_purchase_records(tables.get("Achats"))
        stock_records = table_to_stock_records(tables.get("Stock"))
        db.replace_all(purchases, stock_records)

    def _apply_overrides(
        self,
        tables: Mapping[str, TableData],
        achats_table: TableData | None,
        stock_table: TableData | None,
    ) -> None:
        if achats_table is not None:
            _ensure_purchase_ready_dates(achats_table)
            tables["Achats"] = achats_table
        else:
            _ensure_purchase_ready_dates(tables["Achats"])
        if stock_table is not None:
            tables["Stock"] = stock_table

    def _normalize_tables(self, tables: Mapping[str, TableData]) -> None:
        _normalize_id_column(tables.get("Achats"), HEADERS["ACHATS"].ID)
        _normalize_id_column(tables.get("Stock"), HEADERS["STOCK"].ID)
        _normalize_id_column(tables.get("Ventes"), HEADERS["VENTES"].ID)
        _normalize_id_column(tables.get("Compta 09-2025"), "ID")

    def persist_tables(self, tables: Mapping[str, TableData]) -> None:
        if self.achats_db_path is None:
            return
        records = table_to_purchase_records(tables.get("Achats"))
        stock_records = table_to_stock_records(tables.get("Stock"))
        PurchaseDatabase(self.achats_db_path).replace_all(records, stock_records)


def _has_ready_date(value) -> bool:
    if value in (None, ""):
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, str) and value.strip().upper() == "FALSE":
        return False
    return True


def _ensure_purchase_ready_dates(table: TableData | None) -> None:
    if table is None:
        return
    ready_header = HEADERS["ACHATS"].DATE_MISE_EN_STOCK
    fallback_headers = (
        HEADERS["ACHATS"].PRET_STOCK_COMBINED,
        HEADERS["ACHATS"].PRET_STOCK,
        HEADERS["ACHATS"].PRET_STOCK_ALT,
    )
    headers = list(table.headers)
    if ready_header not in headers:
        headers.append(ready_header)
        table.headers = tuple(headers) if isinstance(table.headers, tuple) else headers
    for row in table.rows:
        current = row.get(ready_header)
        if _has_ready_date(current):
            continue
        for fallback in fallback_headers:
            candidate = row.get(fallback)
            if _has_ready_date(candidate):
                row[ready_header] = candidate
                break
        else:
            row.setdefault(ready_header, "")


def _normalize_integer_value(value) -> int | None:
    if value in (None, "") or isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        if not value.is_integer():
            return None
        return int(value)
    text = str(value).strip()
    if not text:
        return None
    try:
        number = float(text)
    except ValueError:
        return None
    if not number.is_integer():
        return None
    return int(number)


def _normalize_id_column(table: TableData | None, id_header: str) -> None:
    if table is None:
        return
    for row in table.rows:
        normalized = _normalize_integer_value(row.get(id_header))
        if normalized is not None:
            row[id_header] = normalized
