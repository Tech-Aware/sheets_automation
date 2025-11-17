"""SQLite persistence layer for the Achats table."""
from __future__ import annotations

from dataclasses import MISSING, dataclass, fields
import math
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


@dataclass
class StockRecord:
    """Structured representation of a stock row."""

    id: str = ""
    sku: str = ""
    reference: str = ""
    libelle: str = ""
    article: str = ""
    marque: str = ""
    prix_vente: float | None = None
    taille_colis: str = ""
    taille: str = ""
    lot: str = ""
    date_livraison: str = ""
    date_mise_en_stock: str = ""
    mis_en_ligne: str = ""
    date_mise_en_ligne: str = ""
    publie: str = ""
    date_publication: str = ""
    vendu: str = ""
    date_vente: str = ""
    vente_exportee_le: str = ""
    valider_saisie: str = ""

    @classmethod
    def from_mapping(cls, payload: Mapping[str, object]) -> "StockRecord":
        values = {}
        for field in fields(cls):
            value = payload.get(field.name)
            if value is None and field.default is not MISSING:
                value = field.default
            values[field.name] = value
        return cls(**values)

    def as_sql_params(self) -> tuple:
        return tuple(getattr(self, column) for column, _ in _STOCK_COLUMN_TO_HEADER)


_STOCK_COLUMN_TO_HEADER: Sequence[tuple[str, str]] = (
    ("id", HEADERS["STOCK"].ID),
    ("sku", HEADERS["STOCK"].SKU),
    ("reference", HEADERS["STOCK"].REFERENCE),
    ("libelle", HEADERS["STOCK"].LIBELLE),
    ("article", HEADERS["STOCK"].ARTICLE),
    ("marque", HEADERS["STOCK"].MARQUE),
    ("prix_vente", HEADERS["STOCK"].PRIX_VENTE),
    ("taille_colis", HEADERS["STOCK"].TAILLE_COLIS),
    ("taille", HEADERS["STOCK"].TAILLE),
    ("lot", HEADERS["STOCK"].LOT),
    ("date_livraison", HEADERS["STOCK"].DATE_LIVRAISON),
    ("date_mise_en_stock", HEADERS["STOCK"].DATE_MISE_EN_STOCK),
    ("mis_en_ligne", HEADERS["STOCK"].MIS_EN_LIGNE),
    ("date_mise_en_ligne", HEADERS["STOCK"].DATE_MISE_EN_LIGNE),
    ("publie", HEADERS["STOCK"].PUBLIE),
    ("date_publication", HEADERS["STOCK"].DATE_PUBLICATION),
    ("vendu", HEADERS["STOCK"].VENDU),
    ("date_vente", HEADERS["STOCK"].DATE_VENTE),
    ("vente_exportee_le", HEADERS["STOCK"].VENTE_EXPORTEE_LE),
    ("valider_saisie", HEADERS["STOCK"].VALIDER_SAISIE),
)

STOCK_TABLE_HEADERS: Sequence[str] = tuple(header for _, header in _STOCK_COLUMN_TO_HEADER)

_STOCK_ALIASES: Mapping[str, Sequence[str]] = {
    HEADERS["STOCK"].LIBELLE: (HEADERS["STOCK"].LIBELLE_ALT,),
    HEADERS["STOCK"].ARTICLE: (HEADERS["STOCK"].ARTICLE_ALT,),
    HEADERS["STOCK"].TAILLE_COLIS: (HEADERS["STOCK"].TAILLE_COLIS_ALT,),
    HEADERS["STOCK"].LOT: (HEADERS["STOCK"].LOT_ALT,),
    HEADERS["STOCK"].MIS_EN_LIGNE: (HEADERS["STOCK"].MIS_EN_LIGNE_ALT,),
    HEADERS["STOCK"].DATE_MISE_EN_LIGNE: (HEADERS["STOCK"].DATE_MISE_EN_LIGNE_ALT,),
    HEADERS["STOCK"].PUBLIE: (HEADERS["STOCK"].PUBLIE_ALT,),
    HEADERS["STOCK"].DATE_PUBLICATION: (HEADERS["STOCK"].DATE_PUBLICATION_ALT,),
    HEADERS["STOCK"].VENDU: (HEADERS["STOCK"].VENDU_ALT,),
    HEADERS["STOCK"].DATE_VENTE: (HEADERS["STOCK"].DATE_VENTE_ALT,),
    HEADERS["STOCK"].VALIDER_SAISIE: (HEADERS["STOCK"].VALIDER_SAISIE_ALT,),
}


class PurchaseDatabase:
    """Simple helper around the SQLite file that stores Achats."""

    def __init__(self, db_path: str | Path):
        self.path = Path(db_path)

    def ensure_schema(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with sqlite3.connect(self.path) as conn:
            conn.row_factory = sqlite3.Row
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
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS stock (
                    row_id INTEGER PRIMARY KEY AUTOINCREMENT,
                    id TEXT,
                    sku TEXT,
                    reference TEXT,
                    libelle TEXT,
                    article TEXT,
                    marque TEXT,
                    prix_vente REAL,
                    taille_colis TEXT,
                    taille TEXT,
                    lot TEXT,
                    date_livraison TEXT,
                    date_mise_en_stock TEXT,
                    mis_en_ligne TEXT,
                    date_mise_en_ligne TEXT,
                    publie TEXT,
                    date_publication TEXT,
                    vendu TEXT,
                    date_vente TEXT,
                    vente_exportee_le TEXT,
                    valider_saisie TEXT
                )
                """
            )
            self._ensure_column(conn, "stock", "marque", "TEXT")

    def replace_all(
        self,
        records: Iterable[PurchaseRecord | Mapping[str, object]],
        stock_records: Iterable[StockRecord | Mapping[str, object]] | None = None,
    ) -> None:
        self.ensure_schema()
        normalized = [self._ensure_purchase_record(record) for record in records]
        normalized_stock = []
        if stock_records is not None:
            normalized_stock = [self._ensure_stock_record(record) for record in stock_records]
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
            conn.execute("DELETE FROM stock")
            if normalized_stock:
                conn.executemany(
                    """
                    INSERT INTO stock (
                        id, sku, reference, libelle, article, marque, prix_vente, taille_colis, taille,
                        lot, date_livraison, date_mise_en_stock, mis_en_ligne, date_mise_en_ligne, publie,
                        date_publication, vendu, date_vente, vente_exportee_le, valider_saisie
                    ) VALUES (
                        ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
                    )
                    """,
                    [record.as_sql_params() for record in normalized_stock],
                )
            conn.commit()

    def load_table(self) -> TableData:
        if not self.path.exists():  # pragma: no cover - defensive guard
            raise FileNotFoundError(self.path)
        self.ensure_schema()
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

    def load_stock_table(self) -> TableData:
        if not self.path.exists():  # pragma: no cover - defensive guard
            raise FileNotFoundError(self.path)
        self.ensure_schema()
        with sqlite3.connect(self.path) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.execute(
                """
                SELECT id, sku, reference, libelle, article, marque, prix_vente, taille_colis, taille,
                       lot, date_livraison, date_mise_en_stock, mis_en_ligne, date_mise_en_ligne, publie,
                       date_publication, vendu, date_vente, vente_exportee_le, valider_saisie
                FROM stock
                ORDER BY row_id
                """
            )
            rows = cursor.fetchall()
        table_rows: list[dict[str, object]] = []
        for db_row in rows:
            row_dict: dict[str, object] = {}
            for column, header in _STOCK_COLUMN_TO_HEADER:
                value = db_row[column]
                row_dict[header] = value if value is not None else ""
            table_rows.append(row_dict)
        return TableData(headers=STOCK_TABLE_HEADERS, rows=table_rows)

    @staticmethod
    def _ensure_purchase_record(record: PurchaseRecord | Mapping[str, object]) -> PurchaseRecord:
        if isinstance(record, PurchaseRecord):
            return record
        return PurchaseRecord.from_mapping(record)

    @staticmethod
    def _ensure_stock_record(record: StockRecord | Mapping[str, object]) -> StockRecord:
        if isinstance(record, StockRecord):
            return record
        return StockRecord.from_mapping(record)

    @staticmethod
    def _ensure_column(conn: sqlite3.Connection, table: str, column: str, definition: str) -> None:
        cursor = conn.execute(f"PRAGMA table_info({table})")
        existing_columns = {row[1] for row in cursor.fetchall()}
        if column not in existing_columns:
            conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")


def _normalize_purchase_id(value) -> int | None:
    """Convert any spreadsheet value to a clean integer ID."""

    if value in (None, ""):
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        if not math.isfinite(value) or not value.is_integer():
            return None
        return int(value)
    text = str(value).strip()
    if not text:
        return None
    try:
        return int(text)
    except (TypeError, ValueError):
        try:
            number = float(text)
        except (TypeError, ValueError):
            return None
        if not number.is_integer():
            return None
        return int(number)


def table_to_purchase_records(table: TableData) -> list[PurchaseRecord]:
    """Convert a :class:`TableData` Achats table to structured records."""

    records: list[PurchaseRecord] = []
    for row in table.rows:
        payload: dict[str, object] = {}
        for column, header in _COLUMN_TO_HEADER:
            value = row.get(header)
            if column == "id":
                value = _normalize_purchase_id(value)
            payload[column] = value
        records.append(PurchaseRecord.from_mapping(payload))
    return records


def _resolve_stock_value(row: Mapping[str, object], header: str):
    """Return the value for ``header`` falling back to known aliases."""

    aliases = _STOCK_ALIASES.get(header, ())
    for candidate in (header, *aliases):
        value = row.get(candidate)
        if value not in (None, ""):
            return value
    return row.get(header)


def table_to_stock_records(table: TableData | None) -> list[StockRecord]:
    """Convert a :class:`TableData` Stock table to structured records."""

    if table is None:
        return []
    records: list[StockRecord] = []
    for row in table.rows:
        payload: dict[str, object] = {}
        for column, header in _STOCK_COLUMN_TO_HEADER:
            payload[column] = _resolve_stock_value(row, header)
        records.append(StockRecord.from_mapping(payload))
    return records


__all__ = [
    "PurchaseDatabase",
    "PurchaseRecord",
    "StockRecord",
    "ACHATS_TABLE_HEADERS",
    "STOCK_TABLE_HEADERS",
    "table_to_purchase_records",
    "table_to_stock_records",
]
