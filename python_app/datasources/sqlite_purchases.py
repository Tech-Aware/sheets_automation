"""SQLite persistence layer for the Achats table."""
from __future__ import annotations

from dataclasses import MISSING, dataclass, fields
from pathlib import Path
import sqlite3
from typing import Iterable, Mapping, Sequence

from ..config import HEADERS
from .workbook import TableData


# ``slots=True`` is only supported on Python 3.10+.  The UI is often executed
# with Python 3.9 (e.g., on Windows 32-bit installations), so we stick to the
# broadest compatible dataclass definition.
@dataclass
class PurchaseRecord:
    """Structured representation of a purchase row."""

    id: int | None = None
    articles: str = ""
    marque: str = ""
    genre: str = ""
    genre_data: str = ""
    date_achat: str | None = None
    reference: str = ""
    grade: str = ""
    fournisseur_code: str = ""
    mois: str = ""
    mois_num: str = ""
    date_livraison: str | None = None
    delai_livraison: int | None = None
    prix_achat_ship: float | None = None
    quantite_commandee: int | None = None
    quantite_recue: int | None = None
    prix_unitaire_brut: float | None = None
    frais_lavage: float | None = None
    frais_colissage: float | None = None
    total_ttc: float | None = None
    prix_unitaire_ttc: float | None = None
    tracking: str = ""
    pret_pour_stock_date: str | None = None
    date_mise_en_stock: str | None = None

    @classmethod
    def from_mapping(cls, payload: Mapping[str, object]) -> "PurchaseRecord":
        values = {}
        for field in fields(cls):
            value = payload.get(field.name)
            if value is None and field.default is not MISSING:
                value = field.default
            values[field.name] = value
        return cls(**values)

    def as_sql_params(self) -> tuple:
        return tuple(getattr(self, column) for column, _ in _COLUMN_TO_HEADER)


_COLUMN_TO_HEADER: Sequence[tuple[str, str]] = (
    ("id", HEADERS["ACHATS"].ID),
    ("articles", HEADERS["ACHATS"].ARTICLE),
    ("marque", HEADERS["ACHATS"].MARQUE),
    ("genre", HEADERS["ACHATS"].GENRE_LEGACY),
    ("genre_data", HEADERS["ACHATS"].GENRE_DATA),
    ("date_achat", HEADERS["ACHATS"].DATE_ACHAT),
    ("reference", HEADERS["ACHATS"].REFERENCE),
    ("grade", HEADERS["ACHATS"].GRADE),
    ("fournisseur_code", HEADERS["ACHATS"].FOURNISSEUR),
    ("mois", HEADERS["ACHATS"].MOIS),
    ("mois_num", HEADERS["ACHATS"].MOIS_NUM),
    ("date_livraison", HEADERS["ACHATS"].DATE_LIVRAISON),
    ("delai_livraison", HEADERS["ACHATS"].DELAI_LIVRAISON),
    ("prix_achat_ship", HEADERS["ACHATS"].PRIX_ACHAT_SHIP),
    ("quantite_commandee", HEADERS["ACHATS"].QUANTITE_COMMANDEE),
    ("quantite_recue", HEADERS["ACHATS"].QUANTITE_RECUE),
    ("prix_unitaire_brut", HEADERS["ACHATS"].PRIX_UNITAIRE_BRUT),
    ("frais_lavage", HEADERS["ACHATS"].FRAIS_LAVAGE),
    ("frais_colissage", HEADERS["ACHATS"].FRAIS_COLISSAGE),
    ("total_ttc", HEADERS["ACHATS"].TOTAL_TTC),
    ("prix_unitaire_ttc", HEADERS["ACHATS"].PRIX_UNITAIRE_TTC),
    ("tracking", HEADERS["ACHATS"].TRACKING),
    ("pret_pour_stock_date", HEADERS["ACHATS"].PRET_STOCK_COMBINED),
    ("date_mise_en_stock", HEADERS["ACHATS"].DATE_MISE_EN_STOCK),
)

ACHATS_TABLE_HEADERS: Sequence[str] = tuple(header for _, header in _COLUMN_TO_HEADER)


class PurchaseDatabase:
    """Simple helper around the SQLite file that stores Achats."""

    def __init__(self, db_path: str | Path):
        self.path = Path(db_path)

    def ensure_schema(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with sqlite3.connect(self.path) as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS achats (
                    row_id INTEGER PRIMARY KEY AUTOINCREMENT,
                    id INTEGER,
                    articles TEXT NOT NULL DEFAULT '',
                    marque TEXT NOT NULL DEFAULT '',
                    genre TEXT,
                    genre_data TEXT,
                    date_achat TEXT,
                    reference TEXT,
                    grade TEXT,
                    fournisseur_code TEXT,
                    mois TEXT,
                    mois_num TEXT,
                    date_livraison TEXT,
                    delai_livraison INTEGER,
                    prix_achat_ship REAL,
                    quantite_commandee INTEGER,
                    quantite_recue INTEGER,
                    prix_unitaire_brut REAL,
                    frais_lavage REAL,
                    frais_colissage REAL,
                    total_ttc REAL,
                    prix_unitaire_ttc REAL,
                    tracking TEXT,
                    pret_pour_stock_date TEXT,
                    date_mise_en_stock TEXT
                )
                """
            )
            conn.execute(
                "CREATE UNIQUE INDEX IF NOT EXISTS idx_achats_id ON achats(id) WHERE id IS NOT NULL"
            )

    def replace_all(self, records: Iterable[PurchaseRecord | Mapping[str, object]]) -> None:
        self.ensure_schema()
        normalized = [self._ensure_record(record) for record in records]
        with sqlite3.connect(self.path) as conn:
            conn.execute("DELETE FROM achats")
            conn.executemany(
                """
                INSERT INTO achats (
                    id, articles, marque, genre, genre_data, date_achat, reference, grade,
                    fournisseur_code, mois, mois_num, date_livraison, delai_livraison,
                    prix_achat_ship, quantite_commandee, quantite_recue, prix_unitaire_brut,
                    frais_lavage, frais_colissage, total_ttc, prix_unitaire_ttc, tracking,
                    pret_pour_stock_date, date_mise_en_stock
                ) VALUES (
                    ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
                )
                """,
                [record.as_sql_params() for record in normalized],
            )
            conn.commit()

    def load_table(self) -> TableData:
        if not self.path.exists():  # pragma: no cover - defensive guard
            raise FileNotFoundError(self.path)
        with sqlite3.connect(self.path) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.execute(
                """
                SELECT id, articles, marque, genre, genre_data, date_achat, reference, grade,
                       fournisseur_code, mois, mois_num, date_livraison, delai_livraison,
                       prix_achat_ship, quantite_commandee, quantite_recue, prix_unitaire_brut,
                       frais_lavage, frais_colissage, total_ttc, prix_unitaire_ttc, tracking,
                       pret_pour_stock_date, date_mise_en_stock
                FROM achats
                ORDER BY CASE WHEN id IS NULL THEN 1 ELSE 0 END, id, row_id
                """
            )
            rows = cursor.fetchall()
        table_rows: list[dict[str, object]] = []
        for db_row in rows:
            row_dict: dict[str, object] = {}
            for column, header in _COLUMN_TO_HEADER:
                value = db_row[column]
                row_dict[header] = value if value is not None else ""
            table_rows.append(row_dict)
        return TableData(headers=ACHATS_TABLE_HEADERS, rows=table_rows)

    @staticmethod
    def _ensure_record(record: PurchaseRecord | Mapping[str, object]) -> PurchaseRecord:
        if isinstance(record, PurchaseRecord):
            return record
        return PurchaseRecord.from_mapping(record)


def table_to_purchase_records(table: TableData) -> list[PurchaseRecord]:
    """Convert a :class:`TableData` Achats table to structured records."""

    records: list[PurchaseRecord] = []
    for row in table.rows:
        payload: dict[str, object] = {}
        for column, header in _COLUMN_TO_HEADER:
            payload[column] = row.get(header)
        records.append(PurchaseRecord.from_mapping(payload))
    return records


__all__ = ["PurchaseDatabase", "PurchaseRecord", "ACHATS_TABLE_HEADERS", "table_to_purchase_records"]
