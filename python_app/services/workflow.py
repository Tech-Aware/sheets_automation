"""High level helpers that reproduce the Achats → Stock → Ventes workflow."""
from __future__ import annotations

from dataclasses import dataclass as _dataclass
from datetime import date
import math
import re
import unicodedata
from threading import RLock
from sys import version_info
from typing import Mapping, Optional, Sequence

from ..config import HEADERS, MONTH_NAMES_FR
from ..datasources.workbook import TableData
from ..services.summaries import InventoryCache
from ..utils.datefmt import format_display_date, parse_date_value
from ..utils.perf import performance_monitor


# ``slots`` support for :func:`dataclasses.dataclass` only arrived in Python 3.10.
# The UI is expected to run on Python 3.9 (32-bit), so we silently drop the
# argument on older interpreters while keeping the memory optimisation available
# on newer ones.
_DATACLASS_KWARGS = {"slots": True} if version_info >= (3, 10) else {}


@_dataclass(**_DATACLASS_KWARGS)
class PurchaseInput:
    article: str
    marque: str
    reference: str = ""
    quantite: int = 0
    prix_unitaire: float = 0.0
    frais_colissage: float = 0.0
    date_livraison: Optional[str] = None
    genre: str = ""
    date_achat: Optional[str] = None
    grade: str = ""
    fournisseur: str = ""
    quantite_commandee: Optional[int] = None
    quantite_recue: Optional[int] = None
    prix_achat_total: Optional[float] = None
    frais_lavage: float = 0.0
    tracking: str = ""


@_dataclass(**_DATACLASS_KWARGS)
class StockInput:
    purchase_id: str
    sku: str
    prix_vente: float
    lot: str = ""
    taille: str = ""


@_dataclass(**_DATACLASS_KWARGS)
class SaleInput:
    sku: str
    prix_vente: float
    frais_colissage: float = 0.0
    date_vente: Optional[str] = None
    lot: str = ""
    taille: str = ""


class WorkflowCoordinator:
    """Mutable helper that wires the in-memory tables together."""

    def __init__(self, achats: TableData, stock: TableData, ventes: TableData, compta: TableData):
        self.achats = achats
        self.stock = stock
        self.ventes = ventes
        self.compta = compta
        self._lock = RLock()
        self._achats_header_map = self._build_header_map(self.achats.headers)
        self._stock_header_map = self._build_header_map(self.stock.headers)
        self._purchase_by_id: dict[str, dict] = {}
        self._stock_by_sku: dict[str, dict] = {}
        self._sales_by_sku: dict[str, dict] = {}
        self._sku_suffix_index: dict[str, int] = {}
        self._max_purchase_id = 0
        self._max_sale_id = 0
        self._inventory_cache = InventoryCache.from_tables(self.stock.rows, self.ventes.rows, self.achats.rows)
        with self._lock:
            self._rebuild_indexes()

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def create_purchase(self, data: PurchaseInput) -> dict:
        with self._lock, performance_monitor.track("workflow.create_purchase"):
            purchase_id = str(self._next_purchase_id())
            achat_date = self._parse_date_string(data.date_achat) or self._today_date()
            livraison_date = self._parse_date_string(data.date_livraison) or achat_date
            mois_label, mois_num = self._month_info(achat_date)

            qty_received = data.quantite_recue if data.quantite_recue is not None else data.quantite
            if qty_received is None or qty_received <= 0:
                qty_received = data.quantite_commandee if data.quantite_commandee else data.quantite
            qty_received = qty_received or 0
            qty_ordered = data.quantite_commandee if data.quantite_commandee is not None else qty_received

            prix_achat = data.prix_achat_total
            if prix_achat is None:
                prix_achat = round((data.prix_unitaire or 0) * qty_received, 2)
            prix_unitaire_brut = round(prix_achat / qty_received, 2) if qty_received else round(data.prix_unitaire, 2)
            total_ttc = round(prix_achat + data.frais_colissage + data.frais_lavage, 2)
            prix_unitaire_ttc = round(total_ttc / qty_received, 2) if qty_received else total_ttc
            delai = (livraison_date - achat_date).days if achat_date and livraison_date else ""
            reference = self._normalize_reference(data.reference) or self._generate_reference(
                data.article, data.marque, data.genre
            )

            row: dict = {}
            self._set_purchase_value(row, HEADERS["ACHATS"].ID, purchase_id)
            self._set_purchase_value(row, HEADERS["ACHATS"].ARTICLE, data.article)
            self._set_purchase_value(row, HEADERS["ACHATS"].MARQUE, data.marque)
            self._set_purchase_value(row, HEADERS["ACHATS"].REFERENCE, reference)
            self._set_purchase_value(row, HEADERS["ACHATS"].GENRE_DATA, data.genre)
            self._set_purchase_value(row, HEADERS["ACHATS"].GENRE_LEGACY, data.genre)
            self._set_purchase_value(row, HEADERS["ACHATS"].DATE_ACHAT, self._format_date(achat_date))
            self._set_purchase_value(row, HEADERS["ACHATS"].GRADE, data.grade)
            self._set_purchase_value(row, HEADERS["ACHATS"].FOURNISSEUR, data.fournisseur)
            self._set_purchase_value(row, HEADERS["ACHATS"].MOIS, mois_label)
            self._set_purchase_value(row, HEADERS["ACHATS"].MOIS_NUM, mois_num)
            self._set_purchase_value(row, HEADERS["ACHATS"].DATE_LIVRAISON, self._format_date(livraison_date))
            self._set_purchase_value(row, HEADERS["ACHATS"].DELAI_LIVRAISON, delai)
            self._set_purchase_value(row, HEADERS["ACHATS"].QUANTITE_COMMANDEE, qty_ordered)
            self._set_purchase_value(row, HEADERS["ACHATS"].QUANTITE_RECUE, qty_received)
            self._set_purchase_value(row, HEADERS["ACHATS"].PRIX_ACHAT_SHIP, round(prix_achat, 2))
            self._set_purchase_value(row, HEADERS["ACHATS"].PRIX_UNITAIRE_BRUT, prix_unitaire_brut)
            self._set_purchase_value(row, HEADERS["ACHATS"].FRAIS_LAVAGE, round(data.frais_lavage, 2))
            self._set_purchase_value(row, HEADERS["ACHATS"].FRAIS_COLISSAGE, round(data.frais_colissage, 2))
            self._set_purchase_value(row, HEADERS["ACHATS"].TOTAL_TTC, total_ttc)
            self._set_purchase_value(row, HEADERS["ACHATS"].PRIX_UNITAIRE_TTC, prix_unitaire_ttc)
            self._set_purchase_value(row, HEADERS["ACHATS"].TRACKING, data.tracking)
            self._set_purchase_value(row, HEADERS["ACHATS"].PRET_STOCK_COMBINED, "")
            self._set_purchase_value(row, HEADERS["ACHATS"].DATE_MISE_EN_STOCK, "")
            self.achats.rows.append(row)
            self._index_purchase(row)
            self._inventory_cache.on_purchase_added(row)
            return row

    def delete_purchases(self, row_indices: Sequence[int]) -> tuple[int, int]:
        """Remove Achats rows by index and drop matching stock entries."""

        if not row_indices:
            return 0, 0
        with self._lock, performance_monitor.track(
            "workflow.delete_purchases", metadata={"count": len(row_indices)}
        ):
            removed_purchase_ids: list[str] = []
            removed = 0
            for idx in sorted(set(row_indices), reverse=True):
                if idx < 0 or idx >= len(self.achats.rows):
                    continue
                row = self.achats.rows.pop(idx)
                removed += 1
                purchase_id = self._get_purchase_value(row, HEADERS["ACHATS"].ID)
                if purchase_id not in (None, ""):
                    removed_purchase_ids.append(str(purchase_id).strip())
                self._purchase_by_id.pop(self._normalize_id_key(purchase_id), None)
                self._inventory_cache.on_purchase_removed(row)
            stock_removed = self._remove_stock_rows(removed_purchase_ids)
            return removed, stock_removed

    def transfer_to_stock(self, data: StockInput) -> dict:
        with self._lock, performance_monitor.track("workflow.transfer_to_stock"):
            purchase = self._purchase_by_id.get(self._normalize_id_key(data.purchase_id))
            if purchase is None:
                raise ValueError(f"Achat {data.purchase_id} introuvable")
            stock_id_value = self._get_purchase_value(purchase, HEADERS["ACHATS"].ID) or data.purchase_id
            stock_id = str(stock_id_value)
            date_stock = self._today()
            article = self._get_purchase_value(purchase, HEADERS["ACHATS"].ARTICLE) or self._get_purchase_value(
                purchase, HEADERS["ACHATS"].ARTICLE_ALT
            )
            marque = self._get_purchase_value(purchase, HEADERS["ACHATS"].MARQUE) or ""
            libelle = " ".join(part for part in (article, marque) if part).strip()
            row: dict = {}
            self._set_stock_value(row, HEADERS["STOCK"].ID, stock_id)
            self._set_stock_value(row, HEADERS["STOCK"].SKU, data.sku)
            self._set_stock_value(row, HEADERS["STOCK"].LIBELLE, libelle)
            self._set_stock_value(row, HEADERS["STOCK"].ARTICLE, libelle)
            self._set_stock_value(row, HEADERS["STOCK"].MARQUE, marque)
            self._set_stock_value(row, HEADERS["STOCK"].PRIX_VENTE, round(data.prix_vente, 2))
            self._set_stock_value(row, HEADERS["STOCK"].LOT, data.lot)
            self._set_stock_value(row, HEADERS["STOCK"].TAILLE, data.taille)
            self._set_stock_value(
                row,
                HEADERS["STOCK"].DATE_LIVRAISON,
                self._get_purchase_value(purchase, HEADERS["ACHATS"].DATE_LIVRAISON),
            )
            self._set_stock_value(row, HEADERS["STOCK"].DATE_MISE_EN_STOCK, date_stock)
            self.stock.rows.append(row)
            self._index_stock_row(row)
            self._inventory_cache.on_stock_added(row)
            return row

    def register_sale(self, data: SaleInput) -> dict:
        with self._lock, performance_monitor.track("workflow.register_sale"):
            stock_row = self._stock_by_sku.get(self._normalize_sku_key(data.sku))
            if stock_row is None:
                raise ValueError(f"Article {data.sku} introuvable dans le stock")
            sale_date_value = self._parse_date_string(data.date_vente) or self._today_date()
            sale_date = self._format_date(sale_date_value)
            was_sold = self._is_stock_sold(stock_row)
            stock_row[self._stock_column(HEADERS["STOCK"].VENDU_ALT)] = sale_date
            stock_row[self._stock_column(HEADERS["STOCK"].DATE_VENTE_ALT)] = sale_date
            stock_row[self._stock_column(HEADERS["STOCK"].PRIX_VENTE)] = round(data.prix_vente, 2)
            sale_id = str(self._next_sale_id())
            libelle = self._get_stock_value(stock_row, HEADERS["STOCK"].LIBELLE) or self._get_stock_value(
                stock_row, HEADERS["STOCK"].ARTICLE
            )
            sale_row = {
                HEADERS["VENTES"].ID: sale_id,
                HEADERS["VENTES"].DATE_VENTE: sale_date,
                HEADERS["VENTES"].ARTICLE: libelle,
                HEADERS["VENTES"].SKU: data.sku,
                HEADERS["VENTES"].PRIX_VENTE: round(data.prix_vente, 2),
                HEADERS["VENTES"].FRAIS_COLISSAGE: round(data.frais_colissage, 2),
                HEADERS["VENTES"].TAILLE: data.taille or self._get_stock_value(stock_row, HEADERS["STOCK"].TAILLE),
                HEADERS["VENTES"].LOT: data.lot or self._get_stock_value(stock_row, HEADERS["STOCK"].LOT),
            }
            self.ventes.rows.append(sale_row)
            self._sales_by_sku[self._normalize_sku_key(data.sku)] = sale_row
            self._inventory_cache.on_stock_sold(
                stock_row,
                sale_price=data.prix_vente,
                frais=data.frais_colissage,
                was_sold=was_sold,
            )
            ledger_row = self._build_compta_row(sale_row)
            if ledger_row:
                self.compta.rows.append(ledger_row)
            return sale_row

    def register_return(self, sku: str, note: str) -> dict:
        with self._lock, performance_monitor.track("workflow.register_return"):
            sale_row = self._sales_by_sku.get(self._normalize_sku_key(sku))
            if sale_row is None:
                raise ValueError(f"Aucune vente trouvée pour le SKU {sku}")
            sale_row[HEADERS["VENTES"].RETOUR] = note or "Retour client"
            stock_row = self._stock_by_sku.get(self._normalize_sku_key(sku))
            if stock_row is not None:
                was_sold = self._is_stock_sold(stock_row)
                stock_row[self._stock_column(HEADERS["STOCK"].VENDU_ALT)] = ""
                stock_row[self._stock_column(HEADERS["STOCK"].DATE_VENTE_ALT)] = ""
                self._inventory_cache.on_stock_return(stock_row, was_sold=was_sold)
            return sale_row

    def build_sku_base(self, article: str, marque: str, genre: str = "") -> str:
        return self._generate_reference(article, marque, genre)

    def inventory_snapshot(self) -> "InventorySnapshot":
        with self._lock:
            return self._inventory_cache.snapshot()

    def prepare_stock_from_purchase(self, purchase_id: str, ready_date: str | None = None) -> list[dict]:
        with self._lock, performance_monitor.track("workflow.prepare_stock_from_purchase"):
            purchase = self._purchase_by_id.get(self._normalize_id_key(purchase_id))
            if purchase is None:
                raise ValueError(f"Achat {purchase_id} introuvable")
            already_ready = self._get_purchase_value(purchase, HEADERS["ACHATS"].PRET_STOCK_COMBINED)
            if isinstance(already_ready, str):
                already_ready = already_ready.strip()
            if already_ready:
                raise ValueError("Cette commande a déjà été validée pour la mise en stock")
            qty = self._safe_int(
                self._get_purchase_value(purchase, HEADERS["ACHATS"].QUANTITE_RECUE)
                or self._get_purchase_value(purchase, HEADERS["ACHATS"].QUANTITE_RECUE_ALT)
            )
            if qty <= 0:
                raise ValueError("La quantité reçue est manquante pour cette commande")
            base = self._normalize_reference(self._get_purchase_value(purchase, HEADERS["ACHATS"].REFERENCE))
            article = self._get_purchase_value(purchase, HEADERS["ACHATS"].ARTICLE) or ""
            marque = self._get_purchase_value(purchase, HEADERS["ACHATS"].MARQUE) or ""
            if not base:
                base = self._generate_reference(article, marque, genre)
                self._set_purchase_value(purchase, HEADERS["ACHATS"].REFERENCE, base)
            ready_value = self._parse_date_string(ready_date)
            ready_stamp = self._format_date(ready_value or self._today_date())
            self._set_purchase_value(purchase, HEADERS["ACHATS"].PRET_STOCK_COMBINED, ready_stamp)
            self._set_purchase_value(purchase, HEADERS["ACHATS"].DATE_MISE_EN_STOCK, ready_stamp)
            livraison_raw = self._get_purchase_value(purchase, HEADERS["ACHATS"].DATE_LIVRAISON)
            livraison_date = self._parse_date_string(livraison_raw) or self._today_date()
            livraison_str = self._format_date(livraison_date)
            genre = self._get_purchase_value(purchase, HEADERS["ACHATS"].GENRE_DATA) or self._get_purchase_value(
                purchase, HEADERS["ACHATS"].GENRE_LEGACY)
            libelle = " ".join(part for part in (article, marque, genre) if part).strip()
            next_suffix = self._next_sku_suffix(base)
            purchase_id_value = self._get_purchase_value(purchase, HEADERS["ACHATS"].ID) or purchase_id
            stock_id = str(purchase_id_value)
            created: list[dict] = []
            for idx in range(qty):
                stock_row: dict = {}
                suffix = next_suffix + idx + 1
                sku = f"{base}-{suffix}"
                self._set_stock_value(stock_row, HEADERS["STOCK"].ID, stock_id)
                self._set_stock_value(stock_row, HEADERS["STOCK"].SKU, sku)
                self._set_stock_value(stock_row, HEADERS["STOCK"].LIBELLE, libelle)
                self._set_stock_value(stock_row, HEADERS["STOCK"].ARTICLE, libelle)
                self._set_stock_value(stock_row, HEADERS["STOCK"].MARQUE, marque)
                self._set_stock_value(stock_row, HEADERS["STOCK"].REFERENCE, base)
                self._set_stock_value(stock_row, HEADERS["STOCK"].PRIX_VENTE, 0.0)
                self._set_stock_value(stock_row, HEADERS["STOCK"].LOT, "")
                self._set_stock_value(stock_row, HEADERS["STOCK"].TAILLE, "")
                self._set_stock_value(stock_row, HEADERS["STOCK"].DATE_LIVRAISON, livraison_str)
                self._set_stock_value(stock_row, HEADERS["STOCK"].DATE_MISE_EN_STOCK, ready_stamp)
                self.stock.rows.append(stock_row)
                self._index_stock_row(stock_row, base=base, suffix=next_suffix + idx + 1)
                self._inventory_cache.on_stock_added(stock_row)
                created.append(stock_row)
            return created

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------
    @staticmethod
    def _next_numeric_id(rows: list[dict], column: str) -> int:
        max_value = 0
        for row in rows:
            value = WorkflowCoordinator._coerce_numeric_id(row.get(column))
            if value is None:
                continue
            max_value = max(max_value, value)
        return max_value + 1

    @staticmethod
    def _coerce_numeric_id(value) -> int | None:
        if value in (None, ""):
            return None
        if isinstance(value, bool):
            return None
        if isinstance(value, int):
            return value
        if isinstance(value, float):
            if not math.isfinite(value):
                return None
            return int(value)
        text = str(value).strip()
        if not text:
            return None
        try:
            return int(text)
        except (TypeError, ValueError):
            try:
                return int(float(text))
            except (TypeError, ValueError):
                return None

    @staticmethod
    def _find_row(rows: list[dict], column: str, value) -> dict | None:
        for row in rows:
            if row.get(column) == value:
                return row
        return None

    def rebuild_indexes(self) -> None:
        with self._lock, performance_monitor.track(
            "workflow.rebuild_indexes",
            metadata={
                "achats": len(self.achats.rows),
                "stock": len(self.stock.rows),
                "ventes": len(self.ventes.rows),
            },
        ):
            self._rebuild_indexes()

    def _today(self) -> str:
        return self._format_date(self._today_date())

    @staticmethod
    def _today_date() -> date:
        return date.today()

    def _build_header_map(self, headers: list[str]) -> dict[str, str]:
        mapping: dict[str, str] = {}
        for header in headers:
            if not isinstance(header, str):
                continue
            normalized = self._normalize_header_name(header)
            if normalized:
                mapping[normalized] = header
        return mapping

    @staticmethod
    def _normalize_header_name(value: str) -> str:
        if not value:
            return ""
        return "".join(ch for ch in value.lower() if ch.isalnum())

    def _purchase_column(self, header_key: str) -> str:
        return self._achats_header_map.get(self._normalize_header_name(header_key), header_key)

    def _stock_column(self, header_key: str) -> str:
        return self._stock_header_map.get(self._normalize_header_name(header_key), header_key)

    def _set_purchase_value(self, row: dict, header_key: str, value):
        row[self._purchase_column(header_key)] = value

    def _get_purchase_value(self, row: dict, header_key: str):
        return row.get(self._purchase_column(header_key))

    def _set_stock_value(self, row: dict, header_key: str, value):
        row[self._stock_column(header_key)] = value

    def _get_stock_value(self, row: dict, header_key: str):
        return row.get(self._stock_column(header_key))

    @staticmethod
    def _safe_int(value) -> int:
        try:
            return int(float(value))
        except (TypeError, ValueError):
            return 0

    @staticmethod
    def _safe_float(value) -> float:
        try:
            return float(value)
        except (TypeError, ValueError):
            return 0.0

    def _parse_date_string(self, value) -> date | None:
        return parse_date_value(value)

    @staticmethod
    def _format_date(value: date | None) -> str:
        return format_display_date(value)

    def _month_info(self, value: date | None) -> tuple[str, str]:
        if not value:
            return "", ""
        month_name = MONTH_NAMES_FR[value.month - 1] if 1 <= value.month <= 12 else ""
        return month_name.upper(), f"{value.month:02d}"

    def _normalize_reference(self, value: str | None) -> str:
        if not value:
            return ""
        return self._clean_letters(value)

    def _generate_reference(self, article: str, marque: str, genre: str = "") -> str:
        article_code = self._build_initials(article, 2)
        marque_code = self._build_initials(marque, 3)
        genre_code = self._genre_initial(genre)
        base = f"{article_code}{marque_code}{genre_code}".strip()
        return base or "SKU"

    def _build_initials(self, value: str, limit: int) -> str:
        cleaned = self._clean_letters(value)
        if not cleaned:
            return ""
        tokens = re.findall(r"[A-Z0-9]+", cleaned)
        letters: list[str] = []
        for token in tokens or [cleaned]:
            letters.append(token[0])
            if len(letters) >= limit:
                break
        return "".join(letters)

    def _remove_stock_rows(self, purchase_ids: Sequence[str]) -> int:
        if not purchase_ids:
            return 0
        normalized = {pid.strip() for pid in purchase_ids if pid and pid.strip()}
        if not normalized:
            return 0
        removed = 0
        for idx in range(len(self.stock.rows) - 1, -1, -1):
            value = self._get_stock_value(self.stock.rows[idx], HEADERS["STOCK"].ID)
            if value is None:
                continue
            sku = self._get_stock_value(self.stock.rows[idx], HEADERS["STOCK"].SKU)
            if str(value).strip() in normalized:
                row = self.stock.rows.pop(idx)
                self._stock_by_sku.pop(self._normalize_sku_key(sku), None)
                self._inventory_cache.on_stock_removed(row)
                removed += 1
        return removed

    @staticmethod
    def _clean_letters(value: str) -> str:
        normalized = unicodedata.normalize("NFKD", value or "")
        ascii_value = "".join(ch for ch in normalized if ch.isalnum())
        return ascii_value.upper()

    def _genre_initial(self, genre: str) -> str:
        code = self._clean_letters(genre)
        if code.startswith("H"):
            return "H"
        if code.startswith("F"):
            return "F"
        return ""

    def _next_sku_suffix(self, base: str) -> int:
        return self._sku_suffix_index.get(base, 0)

    def _normalize_id_key(self, value) -> str:
        return str(value).strip() if value not in (None, "") else ""

    def _normalize_sku_key(self, sku) -> str:
        return str(sku).strip().upper() if sku not in (None, "") else ""

    def _index_purchase(self, row: dict) -> None:
        purchase_id = self._normalize_id_key(self._get_purchase_value(row, HEADERS["ACHATS"].ID))
        if purchase_id:
            self._purchase_by_id[purchase_id] = row
            numeric = self._coerce_numeric_id(purchase_id)
            if numeric is not None:
                self._max_purchase_id = max(self._max_purchase_id, numeric)

    def _index_sale(self, row: dict) -> None:
        sku = self._normalize_sku_key(row.get(HEADERS["VENTES"].SKU))
        if sku:
            self._sales_by_sku[sku] = row
        numeric = self._coerce_numeric_id(row.get(HEADERS["VENTES"].ID))
        if numeric is not None:
            self._max_sale_id = max(self._max_sale_id, numeric)

    def _extract_sku_base(self, sku: str | None) -> str:
        if not sku:
            return ""
        parts = str(sku).strip().upper().rsplit("-", 1)
        return parts[0] if parts else ""

    def _extract_sku_suffix(self, sku: str | None) -> int:
        if not sku:
            return 0
        parts = str(sku).strip().upper().rsplit("-", 1)
        if len(parts) != 2:
            return 0
        try:
            return int(parts[1])
        except ValueError:
            return 0

    def _index_stock_row(self, row: dict, *, base: str | None = None, suffix: int | None = None) -> None:
        sku = self._normalize_sku_key(self._get_stock_value(row, HEADERS["STOCK"].SKU))
        if sku:
            self._stock_by_sku[sku] = row
        sku_base = base or self._extract_sku_base(sku) or self._normalize_sku_key(
            self._get_stock_value(row, HEADERS["STOCK"].REFERENCE)
        )
        if sku_base:
            current_suffix = suffix if suffix is not None else self._extract_sku_suffix(sku)
            if current_suffix:
                self._sku_suffix_index[sku_base] = max(self._sku_suffix_index.get(sku_base, 0), current_suffix)

    def _is_stock_sold(self, row: Mapping) -> bool:
        vendu = row.get(self._stock_column(HEADERS["STOCK"].VENDU_ALT))
        vendu = vendu or row.get(self._stock_column(HEADERS["STOCK"].VENDU))
        return bool(vendu)

    def _rebuild_indexes(self) -> None:
        self._purchase_by_id.clear()
        self._stock_by_sku.clear()
        self._sales_by_sku.clear()
        self._sku_suffix_index.clear()
        self._max_purchase_id = 0
        self._max_sale_id = 0
        for row in self.achats.rows:
            self._index_purchase(row)
        for row in self.stock.rows:
            self._index_stock_row(row)
        for row in self.ventes.rows:
            self._index_sale(row)
        self._inventory_cache = InventoryCache.from_tables(self.stock.rows, self.ventes.rows, self.achats.rows)

    def _next_purchase_id(self) -> int:
        self._max_purchase_id += 1
        return self._max_purchase_id

    def _next_sale_id(self) -> int:
        self._max_sale_id += 1
        return self._max_sale_id

    def _build_compta_row(self, sale_row: dict) -> dict:
        if not self.compta.headers:
            return {}
        price = float(sale_row.get(HEADERS["VENTES"].PRIX_VENTE) or 0)
        frais = float(sale_row.get(HEADERS["VENTES"].FRAIS_COLISSAGE) or 0)
        margin = price - frais
        coeff = round(margin / price, 2) if price else 0.0
        ledger = {}
        for header in self.compta.headers:
            if header == "ID":
                ledger[header] = sale_row.get(HEADERS["VENTES"].ID)
            elif header == "SKU":
                ledger[header] = sale_row.get(HEADERS["VENTES"].SKU)
            elif header == "LIBELLÉS":
                ledger[header] = sale_row.get(HEADERS["VENTES"].ARTICLE)
            elif header == "DATE DE VENTE":
                ledger[header] = sale_row.get(HEADERS["VENTES"].DATE_VENTE)
            elif header == "MARGE BRUTE":
                ledger[header] = round(margin, 2)
            elif header == "COEFF MARGE":
                ledger[header] = coeff
            elif header == "NBR PCS VENDU":
                ledger[header] = 1
        return ledger


__all__ = [
    "PurchaseInput",
    "SaleInput",
    "StockInput",
    "WorkflowCoordinator",
]
