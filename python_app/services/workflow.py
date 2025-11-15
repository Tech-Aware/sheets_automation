"""High level helpers that reproduce the Achats → Stock → Ventes workflow."""
from __future__ import annotations

from dataclasses import dataclass as _dataclass
from datetime import date, datetime, timedelta
import re
import unicodedata
from sys import version_info
from typing import Optional

from ..config import HEADERS, MONTH_NAMES_FR
from ..datasources.workbook import TableData


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
        self._achats_header_map = self._build_header_map(self.achats.headers)
        self._stock_header_map = self._build_header_map(self.stock.headers)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def create_purchase(self, data: PurchaseInput) -> dict:
        purchase_id = str(self._next_numeric_id(self.achats.rows, HEADERS["ACHATS"].ID))
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
        reference = self._normalize_reference(data.reference) or self._generate_reference(data.article, data.marque, data.genre)

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
        return row

    def transfer_to_stock(self, data: StockInput) -> dict:
        purchase = self._find_row(self.achats.rows, HEADERS["ACHATS"].ID, data.purchase_id)
        if purchase is None:
            raise ValueError(f"Achat {data.purchase_id} introuvable")
        stock_id = str(self._next_numeric_id(self.stock.rows, HEADERS["STOCK"].ID))
        date_stock = self._today()
        libelle = self._get_purchase_value(purchase, HEADERS["ACHATS"].ARTICLE) or self._get_purchase_value(
            purchase, HEADERS["ACHATS"].ARTICLE_ALT
        )
        row: dict = {}
        self._set_stock_value(row, HEADERS["STOCK"].ID, stock_id)
        self._set_stock_value(row, HEADERS["STOCK"].SKU, data.sku)
        self._set_stock_value(row, HEADERS["STOCK"].LIBELLE, libelle)
        self._set_stock_value(row, HEADERS["STOCK"].ARTICLE, libelle)
        self._set_stock_value(row, HEADERS["STOCK"].PRIX_VENTE, round(data.prix_vente, 2))
        self._set_stock_value(row, HEADERS["STOCK"].LOT, data.lot)
        self._set_stock_value(row, HEADERS["STOCK"].TAILLE, data.taille)
        self._set_stock_value(row, HEADERS["STOCK"].DATE_LIVRAISON, self._get_purchase_value(purchase, HEADERS["ACHATS"].DATE_LIVRAISON))
        self._set_stock_value(row, HEADERS["STOCK"].DATE_MISE_EN_STOCK, date_stock)
        self.stock.rows.append(row)
        return row

    def register_sale(self, data: SaleInput) -> dict:
        stock_row = self._find_row(self.stock.rows, HEADERS["STOCK"].SKU, data.sku)
        if stock_row is None:
            raise ValueError(f"Article {data.sku} introuvable dans le stock")
        sale_date = data.date_vente or self._today()
        stock_row[self._stock_column(HEADERS["STOCK"].VENDU_ALT)] = sale_date
        stock_row[self._stock_column(HEADERS["STOCK"].DATE_VENTE_ALT)] = sale_date
        stock_row[self._stock_column(HEADERS["STOCK"].PRIX_VENTE)] = round(data.prix_vente, 2)
        sale_id = str(self._next_numeric_id(self.ventes.rows, HEADERS["VENTES"].ID))
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
        ledger_row = self._build_compta_row(sale_row)
        if ledger_row:
            self.compta.rows.append(ledger_row)
        return sale_row

    def register_return(self, sku: str, note: str) -> dict:
        sale_row = self._find_row(self.ventes.rows, HEADERS["VENTES"].SKU, sku)
        if sale_row is None:
            raise ValueError(f"Aucune vente trouvée pour le SKU {sku}")
        sale_row[HEADERS["VENTES"].RETOUR] = note or "Retour client"
        stock_row = self._find_row(self.stock.rows, HEADERS["STOCK"].SKU, sku)
        if stock_row is not None:
            stock_row[self._stock_column(HEADERS["STOCK"].VENDU_ALT)] = ""
            stock_row[self._stock_column(HEADERS["STOCK"].DATE_VENTE_ALT)] = ""
        return sale_row

    def build_sku_base(self, article: str, marque: str, genre: str = "") -> str:
        return self._generate_reference(article, marque, genre)

    def prepare_stock_from_purchase(self, purchase_id: str, ready_date: str | None = None) -> list[dict]:
        purchase = self._find_row(self.achats.rows, HEADERS["ACHATS"].ID, purchase_id)
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
        ready_stamp = ready_date or self._today()
        self._set_purchase_value(purchase, HEADERS["ACHATS"].PRET_STOCK_COMBINED, ready_stamp)
        self._set_purchase_value(purchase, HEADERS["ACHATS"].DATE_MISE_EN_STOCK, ready_stamp)
        livraison_raw = self._get_purchase_value(purchase, HEADERS["ACHATS"].DATE_LIVRAISON)
        livraison_date = self._parse_date_string(livraison_raw) or self._today_date()
        livraison_str = self._format_date(livraison_date)
        genre = self._get_purchase_value(purchase, HEADERS["ACHATS"].GENRE_DATA) or self._get_purchase_value(
            purchase, HEADERS["ACHATS"].GENRE_LEGACY)
        libelle = " ".join(part for part in (article, marque, genre) if part).strip()
        next_suffix = self._next_sku_suffix(base)
        next_stock_id = self._next_numeric_id(self.stock.rows, HEADERS["STOCK"].ID)
        created: list[dict] = []
        for idx in range(qty):
            stock_row: dict = {}
            stock_id = str(next_stock_id + idx)
            suffix = next_suffix + idx + 1
            sku = f"{base}-{suffix}"
            self._set_stock_value(stock_row, HEADERS["STOCK"].ID, stock_id)
            self._set_stock_value(stock_row, HEADERS["STOCK"].SKU, sku)
            self._set_stock_value(stock_row, HEADERS["STOCK"].LIBELLE, libelle)
            self._set_stock_value(stock_row, HEADERS["STOCK"].ARTICLE, libelle)
            self._set_stock_value(stock_row, HEADERS["STOCK"].REFERENCE, base)
            self._set_stock_value(stock_row, HEADERS["STOCK"].OLD_SKU, "")
            self._set_stock_value(stock_row, HEADERS["STOCK"].PRIX_VENTE, 0.0)
            self._set_stock_value(stock_row, HEADERS["STOCK"].LOT, "")
            self._set_stock_value(stock_row, HEADERS["STOCK"].TAILLE, "")
            self._set_stock_value(stock_row, HEADERS["STOCK"].DATE_LIVRAISON, livraison_str)
            self._set_stock_value(stock_row, HEADERS["STOCK"].DATE_MISE_EN_STOCK, ready_stamp)
            self.stock.rows.append(stock_row)
            created.append(stock_row)
        return created

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------
    @staticmethod
    def _next_numeric_id(rows: list[dict], column: str) -> int:
        max_value = 0
        for row in rows:
            try:
                value = int(row.get(column) or 0)
            except (TypeError, ValueError):
                continue
            max_value = max(max_value, value)
        return max_value + 1

    @staticmethod
    def _find_row(rows: list[dict], column: str, value) -> dict | None:
        for row in rows:
            if row.get(column) == value:
                return row
        return None

    def _today(self) -> str:
        return self._today_date().isoformat()

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

    def _parse_date_string(self, value) -> date | None:
        if value is None or value == "":
            return None
        if isinstance(value, date):
            return value
        if isinstance(value, (int, float)):
            return date(1899, 12, 30) + timedelta(days=int(value))
        text = str(value).strip()
        if not text:
            return None
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y"):
            try:
                return datetime.strptime(text, fmt).date()
            except ValueError:
                continue
        try:
            serial = float(text)
        except ValueError:
            return None
        return date(1899, 12, 30) + timedelta(days=int(serial))

    @staticmethod
    def _format_date(value: date | None) -> str:
        return value.isoformat() if isinstance(value, date) else ""

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
        sku_column = self._stock_column(HEADERS["STOCK"].SKU)
        prefix = f"{base}-"
        max_suffix = 0
        for row in self.stock.rows:
            raw = row.get(sku_column)
            if raw is None:
                continue
            text = str(raw).strip().upper()
            if not text.startswith(prefix):
                continue
            suffix = text[len(prefix) :]
            try:
                max_suffix = max(max_suffix, int(suffix))
            except ValueError:
                continue
        return max_suffix

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
