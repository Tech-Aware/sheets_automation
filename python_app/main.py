"""Entry point for the CustomTkinter application."""
from __future__ import annotations

import argparse
from datetime import date
import math
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Mapping, Sequence

import customtkinter as ctk

# ``python python_app/main.py`` was failing because ``__package__`` is empty when a
# module is executed as a script.  To keep relative imports working both for
# ``python -m python_app.main`` and direct execution, fall back to absolute
# imports after adding the repository root to ``sys.path``.
try:  # pragma: no cover - defensive import path configuration
    from .config import DEFAULT_STOCK_HEADERS, HEADERS, MONTH_NAMES_FR
    from .datasources.workbook import TableData, WorkbookRepository
    from .datasources.sqlite_purchases import (
        PurchaseDatabase,
        table_to_purchase_records,
        table_to_stock_records,
    )
    from .services.stock_import import merge_stock_table
    from .services.summaries import (
        _base_reference_from_stock,
        _build_reference_unit_price_index,
        build_inventory_snapshot,
    )
    from .services.workflow import PurchaseInput, SaleInput, StockInput, WorkflowCoordinator
    from .ui.tables import ScrollableTable
    from .ui.widgets import DatePickerEntry
    from .utils.datefmt import format_display_date, parse_date_value
except ImportError:  # pragma: no cover - executed when run as a script
    package_root = Path(__file__).resolve().parent.parent
    if str(package_root) not in sys.path:
        sys.path.append(str(package_root))
    from python_app.config import DEFAULT_STOCK_HEADERS, HEADERS, MONTH_NAMES_FR
    from python_app.datasources.workbook import TableData, WorkbookRepository
    from python_app.datasources.sqlite_purchases import (
        PurchaseDatabase,
        table_to_purchase_records,
        table_to_stock_records,
    )
    from python_app.services.stock_import import merge_stock_table
    from python_app.services.summaries import (
        _base_reference_from_stock,
        _build_reference_unit_price_index,
        build_inventory_snapshot,
    )
    from python_app.services.workflow import PurchaseInput, SaleInput, StockInput, WorkflowCoordinator
    from python_app.ui.tables import ScrollableTable
    from python_app.ui.widgets import DatePickerEntry
    from python_app.utils.datefmt import format_display_date, parse_date_value


def _has_ready_date(value) -> bool:
    """Return ``True`` when ``value`` looks like a real ready-date stamp."""

    if value in (None, ""):
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, str) and value.strip().upper() == "FALSE":
        return False
    return True


def _ensure_purchase_ready_dates(table: TableData | None) -> None:
    """Populate ``MIS EN STOCK LE`` from the "Prêt" columns when missing."""

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
    """Return ``value`` as a clean integer when possible."""

    if value in (None, "") or isinstance(value, bool):
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
        number = float(text)
    except ValueError:
        return None
    if not math.isfinite(number) or not number.is_integer():
        return None
    return int(number)


def _normalize_id_column(table: TableData | None, id_header: str) -> None:
    """Ensure ID columns are stored as integers instead of floats."""

    if table is None:
        return
    for row in table.rows:
        normalized = _normalize_integer_value(row.get(id_header))
        if normalized is not None:
            row[id_header] = normalized


def _safe_float(value) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


class VintageErpApp(ctk.CTk):
    """Simple multipage CustomTkinter application."""

    def __init__(
        self,
        repository: WorkbookRepository,
        achats_table: TableData | None = None,
        stock_table: TableData | None = None,
        *,
        achats_db_path: Path | None = None,
    ):
        super().__init__()
        self.title("Vintage ERP (Prerelease 1.2)")
        self.geometry("1200x800")
        self.minsize(1024, 720)
        self.repository = repository
        self.achats_db_path = Path(achats_db_path) if achats_db_path is not None else None
        self.tables = self.repository.load_many("Achats", "Stock", "Ventes", "Compta 09-2025")
        if achats_table is not None:
            _ensure_purchase_ready_dates(achats_table)
            self.tables["Achats"] = achats_table
        else:
            _ensure_purchase_ready_dates(self.tables["Achats"])
        if stock_table is not None:
            self.tables["Stock"] = stock_table
        _normalize_id_column(self.tables.get("Achats"), HEADERS["ACHATS"].ID)
        _normalize_id_column(self.tables.get("Stock"), HEADERS["STOCK"].ID)
        _normalize_id_column(self.tables.get("Ventes"), HEADERS["VENTES"].ID)
        _normalize_id_column(self.tables.get("Compta 09-2025"), "ID")
        self.workflow = WorkflowCoordinator(
            self.tables["Achats"],
            self.tables["Stock"],
            self.tables["Ventes"],
            self.tables["Compta 09-2025"],
        )
        self.tabview = ctk.CTkTabview(self)
        self.tabview.pack(fill="both", expand=True, padx=16, pady=16)
        self.table_views: dict[str, TableView] = {}
        self.purchase_view: PurchasesView | None = None
        self._build_tabs()
        self.protocol("WM_DELETE_WINDOW", self._handle_close)

    def _build_tabs(self):
        summary = build_inventory_snapshot(
            self.tables["Stock"].rows, self.tables["Ventes"].rows, self.tables["Achats"].rows
        )

        dashboard_tab = self.tabview.add("Dashboard")
        self.dashboard_view = DashboardView(dashboard_tab, summary)

        for sheet, label in (
            ("Achats", "Achats"),
            ("Stock", "Stock"),
            ("Ventes", "Ventes"),
            ("Compta 09-2025", "Compta"),
        ):
            tab = self.tabview.add(label)
            if sheet == "Achats":
                view = PurchasesView(tab, self.tables[sheet], self.workflow, self.refresh_views)
                self.purchase_view = view
            elif sheet == "Stock":
                view = StockTableView(
                    tab,
                    self.tables[sheet],
                    workflow=self.workflow,
                    on_table_changed=self.refresh_views,
                )
            else:
                view = TableView(tab, self.tables[sheet], on_table_changed=self.refresh_views)
            self.table_views[sheet] = view

        months_tab = self.tabview.add("Calendrier")
        CalendarView(months_tab)

        stock_options_tab = self.tabview.add("Options")
        StockOptionsView(stock_options_tab, self.tables["Stock"], self.refresh_views)

    def refresh_views(self):
        summary = build_inventory_snapshot(
            self.tables["Stock"].rows, self.tables["Ventes"].rows, self.tables["Achats"].rows
        )
        self.dashboard_view.refresh(summary)
        if self.purchase_view is not None:
            self.purchase_view.refresh()
        for view in self.table_views.values():
            view.refresh()

    def _handle_close(self):
        try:
            self._persist_tables()
        except Exception as exc:  # pragma: no cover - UI safeguard
            messagebox.showerror("Sauvegarde des données", f"Échec de l'enregistrement des données : {exc}")
            return
        self.destroy()

    def _persist_tables(self):
        if self.achats_db_path is None:
            return
        records = table_to_purchase_records(self.tables["Achats"])
        stock_records = table_to_stock_records(self.tables.get("Stock"))
        PurchaseDatabase(self.achats_db_path).replace_all(records, stock_records)


class DashboardView(ctk.CTkFrame):
    def __init__(self, master, snapshot):
        super().__init__(master)
        self.pack(fill="both", expand=True)
        title = ctk.CTkLabel(self, text="Vue d'ensemble", font=ctk.CTkFont(size=24, weight="bold"))
        title.pack(pady=(16, 8))
        grid = ctk.CTkFrame(self)
        grid.pack(padx=16, pady=16, fill="x")
        self.value_labels: dict[str, ctk.CTkLabel] = {}
        stats = snapshot.as_dict()
        for idx, (label, value) in enumerate(stats.items()):
            card = ctk.CTkFrame(grid)
            card.grid(row=0, column=idx, padx=8, pady=8, sticky="nsew")
            grid.grid_columnconfigure(idx, weight=1)
            ctk.CTkLabel(card, text=label.replace("_", " ").title()).pack(padx=12, pady=(12, 4))
            value_label = ctk.CTkLabel(card, text=str(value))
            value_label.pack(padx=12, pady=(0, 12))
            self.value_labels[label] = value_label

    def refresh(self, snapshot):
        stats = snapshot.as_dict()
        for key, value in stats.items():
            label = self.value_labels.get(key)
            if label is not None:
                label.configure(text=str(value))


class PurchasesView(ctk.CTkFrame):
    SUMMARY_HEADERS = (
        HEADERS["ACHATS"].ID,
        HEADERS["ACHATS"].ARTICLE,
        HEADERS["ACHATS"].REFERENCE,
        HEADERS["ACHATS"].DATE_ACHAT,
        HEADERS["ACHATS"].DATE_MISE_EN_STOCK,
        HEADERS["ACHATS"].TOTAL_TTC,
    )

    def __init__(self, master, table, workflow: WorkflowCoordinator, refresh_callback):
        super().__init__(master)
        self.pack(fill="both", expand=True)
        self.table = table
        self.workflow = workflow
        self.refresh_callback = refresh_callback
        self.log_var = tk.StringVar(value="Consultez les commandes ou ajoutez-en une nouvelle.")
        self.status_var = tk.StringVar(value="Double-cliquez sur une commande pour voir les détails.")
        self.ready_checkbox_var = tk.BooleanVar(value=False)
        self.ready_id_entry: ctk.CTkEntry | None = None
        self.ready_date_entry: ctk.CTkEntry | None = None
        self.table_widget: ScrollableTable | None = None
        self.add_dialog: AddPurchaseDialog | None = None
        self.detail_dialog: PurchaseDetailDialog | None = None
        self.layout = ctk.CTkFrame(self)
        self.layout.pack(fill="both", expand=True, padx=16, pady=16)
        self.layout.grid_rowconfigure(0, weight=1)
        self.layout.grid_columnconfigure(0, weight=3)
        self.layout.grid_columnconfigure(1, weight=2)
        self._build_table(self.layout)
        self._build_ready_section(self.layout)

    def _build_table(self, parent):
        frame = ctk.CTkFrame(parent)
        frame.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        helper = ctk.CTkFrame(frame)
        helper.pack(fill="x", padx=12, pady=(12, 0))
        ctk.CTkButton(helper, text="Ajouter une commande", command=self._open_add_dialog).pack(side="left")
        ctk.CTkButton(helper, text="Supprimer la sélection", command=self._delete_selected_rows).pack(
            side="left", padx=(8, 0)
        )
        ctk.CTkLabel(helper, textvariable=self.status_var, anchor="w").pack(side="right", fill="x", expand=True, padx=(12, 0))
        self.table_widget = ScrollableTable(
            frame,
            self.SUMMARY_HEADERS,
            self._build_summary_rows(),
            height=18,
            column_width=135,
            column_widths={
                HEADERS["ACHATS"].ID: 21,
                HEADERS["ACHATS"].REFERENCE: 92,
                HEADERS["ACHATS"].TOTAL_TTC: 110,
            },
            enable_inline_edit=False,
            on_row_activated=self._open_detail_dialog,
        )
        self.table_widget.pack(fill="both", expand=True, padx=12, pady=12)

    def _build_ready_section(self, parent):
        frame = ctk.CTkFrame(parent)
        frame.grid(row=0, column=1, sticky="nsew")
        ctk.CTkLabel(frame, text="Valider la mise en stock", font=ctk.CTkFont(size=18, weight="bold"), anchor="w").pack(
            fill="x", padx=12, pady=(12, 4)
        )
        self.ready_id_entry = self._entry_row(frame, "ID achat")
        self.ready_date_entry = self._entry_row(frame, "Date de mise en stock (JJ/MM/AAAA)", date_picker=True)
        checkbox_row = ctk.CTkFrame(frame)
        checkbox_row.pack(fill="x", padx=12, pady=(0, 4))
        ctk.CTkCheckBox(
            checkbox_row,
            text="Cocher pour valider aujourd'hui",
            variable=self.ready_checkbox_var,
            command=self._handle_ready_checkbox,
        ).pack(anchor="w")
        ctk.CTkButton(frame, text="Créer les SKU", command=self._handle_ready).pack(padx=12, pady=(8, 4))
        ctk.CTkLabel(frame, textvariable=self.log_var, anchor="w").pack(fill="x", padx=12, pady=(0, 12))

    def _entry_row(self, parent, label: str, *, date_picker: bool = False):
        row = ctk.CTkFrame(parent)
        row.pack(fill="x", padx=12, pady=2)
        ctk.CTkLabel(row, text=label, width=220, anchor="w").pack(side="left")
        entry = DatePickerEntry(row) if date_picker else ctk.CTkEntry(row)
        entry.pack(side="left", fill="x", expand=True, padx=(8, 0))
        return entry

    def _build_summary_rows(self):
        summary_rows = []
        for row in self.table.rows:
            summary: dict[str, str] = {}
            for header in self.SUMMARY_HEADERS:
                value = row.get(header, "")
                if header == HEADERS["ACHATS"].TOTAL_TTC and isinstance(value, (int, float)):
                    summary[header] = f"{value:.2f}"
                else:
                    summary[header] = value if value is not None else ""
            summary_rows.append(summary)
        return summary_rows

    def _open_add_dialog(self):
        if self.add_dialog is not None and self.add_dialog.winfo_exists():
            self.add_dialog.focus()
            return
        self.add_dialog = AddPurchaseDialog(
            self,
            self._handle_add_submission,
            self.workflow.build_sku_base,
            self._on_dialog_closed,
        )

    def _on_dialog_closed(self):
        self.add_dialog = None

    def _open_detail_dialog(self, row_index: int):
        if row_index < 0 or row_index >= len(self.table.rows):
            return
        if self.detail_dialog is not None and self.detail_dialog.winfo_exists():
            self.detail_dialog.focus()
            return
        row = self.table.rows[row_index]
        self.detail_dialog = PurchaseDetailDialog(
            self,
            row_index,
            self.table.headers,
            row,
            lambda data: self._handle_detail_save(row_index, data),
            self._on_detail_dialog_closed,
        )

    def _handle_detail_save(self, row_index: int, payload: dict[str, str]):
        if row_index < 0 or row_index >= len(self.table.rows):
            return
        row = self.table.rows[row_index]
        row.update(payload)
        self.table_widget.refresh(self._build_summary_rows())
        purchase_id = row.get(HEADERS["ACHATS"].ID, row_index + 1)
        self.status_var.set(f"Commande {purchase_id} mise à jour")

    def _on_detail_dialog_closed(self):
        self.detail_dialog = None

    def _handle_add_submission(self, form_data: dict[str, str]) -> tuple[bool, str]:
        article = form_data.get("article", "").strip()
        marque = form_data.get("marque", "").strip()
        if not article or not marque:
            message = "Article et marque sont obligatoires"
            self._log(message)
            return False, message
        genre = form_data.get("genre", "").strip()
        date_achat = form_data.get("date_achat", "").strip()
        grade = form_data.get("grade", "").strip()
        fournisseur = form_data.get("fournisseur", "").strip()
        date_livraison = form_data.get("date_livraison", "").strip()
        tracking = form_data.get("tracking", "").strip()
        try:
            qty_cmd = self._parse_int(form_data.get("quantite_commandee", ""), "Quantité commandée")
            qty_recue = self._parse_int(form_data.get("quantite_recue", ""), "Quantité reçue") or qty_cmd
            prix_total = self._parse_float(form_data.get("prix_achat", ""), "Prix d'achat TTC")
            frais_lavage = self._parse_float(form_data.get("frais_lavage", ""), "Frais de lavage")
        except ValueError as exc:
            message = str(exc)
            self._log(message)
            return False, message
        if qty_recue <= 0:
            message = "La quantité reçue doit être supérieure à 0"
            self._log(message)
            return False, message
        unit_price = prix_total / qty_recue if qty_recue else 0.0
        data = PurchaseInput(
            article=article,
            marque=marque,
            genre=genre,
            date_achat=date_achat or None,
            grade=grade,
            fournisseur=fournisseur,
            quantite_commandee=qty_cmd,
            quantite_recue=qty_recue,
            quantite=qty_recue,
            prix_achat_total=prix_total,
            prix_unitaire=unit_price,
            frais_lavage=frais_lavage,
            date_livraison=date_livraison or None,
            tracking=tracking,
        )
        row = self.workflow.create_purchase(data)
        message = f"Commande {row.get(HEADERS['ACHATS'].ID)} ajoutée"
        self._log(message)
        self.refresh_callback()
        return True, message

    @staticmethod
    def _parse_int(value: str, label: str) -> int:
        if not value:
            return 0
        try:
            return int(value)
        except ValueError as exc:
            raise ValueError(f"{label} doit être un nombre entier") from exc

    @staticmethod
    def _parse_float(value: str, label: str) -> float:
        if not value:
            return 0.0
        try:
            return float(value)
        except ValueError as exc:
            raise ValueError(f"{label} doit être un nombre") from exc

    def _handle_ready_checkbox(self):
        if self.ready_date_entry is None:
            return
        if self.ready_checkbox_var.get():
            self.ready_date_entry.delete(0, tk.END)
            self.ready_date_entry.insert(0, format_display_date(date.today()))
        else:
            self.ready_date_entry.delete(0, tk.END)

    def _handle_ready(self):
        if self.ready_id_entry is None or self.ready_date_entry is None:
            return
        purchase_id = self.ready_id_entry.get().strip()
        if not purchase_id:
            self._log("Veuillez saisir l'ID de la commande à mettre en stock")
            return
        ready_date_raw = self.ready_date_entry.get().strip()
        if self.ready_checkbox_var.get() and not ready_date_raw:
            ready_date_raw = format_display_date(date.today())
        ready_value = parse_date_value(ready_date_raw)
        ready_stamp = format_display_date(ready_value) if ready_value else (ready_date_raw or None)
        try:
            created = self.workflow.prepare_stock_from_purchase(purchase_id, ready_stamp)
        except ValueError as exc:
            self._log(str(exc))
            return
        self._log(f"{len(created)} SKU créés pour l'achat {purchase_id}")
        self.ready_checkbox_var.set(False)
        self.ready_id_entry.delete(0, tk.END)
        self.ready_date_entry.delete(0, tk.END)
        self.refresh_callback()

    def refresh(self):
        if self.table_widget is not None:
            self.table_widget.refresh(self._build_summary_rows())

    def _log(self, message: str):
        self.log_var.set(message)

    def _delete_selected_rows(self):
        if self.table_widget is None:
            return
        indices = self.table_widget.get_selected_indices()
        if not indices:
            self.status_var.set("Sélectionnez au moins une commande à supprimer.")
            return
        count = len(indices)
        if not messagebox.askyesno(
            "Confirmer la suppression",
            f"Supprimer définitivement {count} commande(s) ?",
        ):
            return
        removed, stock_removed = self.workflow.delete_purchases(indices)
        if removed:
            self.table_widget.refresh(self._build_summary_rows())
            self.refresh_callback()
            status = f"{removed} commande(s) supprimée(s)."
            if stock_removed:
                status += f" {stock_removed} article(s) lié(s) retiré(s) du stock."
            self.status_var.set(status)
            self._log(status)
        else:
            self.status_var.set("Aucune commande supprimée.")


class AddPurchaseDialog(ctk.CTkToplevel):
    def __init__(self, master, on_submit, sku_builder, on_close):
        super().__init__(master)
        self.on_submit = on_submit
        self.on_close = on_close
        self.sku_builder = sku_builder
        self.entries: dict[str, ctk.CTkEntry | DatePickerEntry] = {}
        self.genre_var = tk.StringVar(value="Mixte")
        self.sku_preview_var = tk.StringVar(value="Préfixe SKU : --")
        self.log_var = tk.StringVar(value="")
        self.title("Nouvelle commande")
        self.geometry("520x700")
        self.minsize(460, 580)
        self.transient(master.winfo_toplevel())
        self.grab_set()
        self._build_form()
        self.protocol("WM_DELETE_WINDOW", self._close)
        self._update_sku_preview()

    def _build_form(self):
        frame = ctk.CTkScrollableFrame(self)
        frame.pack(fill="both", expand=True, padx=16, pady=16)
        ctk.CTkLabel(frame, text="Nouvelle commande", font=ctk.CTkFont(size=18, weight="bold"), anchor="w").pack(
            fill="x", padx=12, pady=(8, 12)
        )
        fields = [
            ("Article", "article", "", False),
            ("Marque", "marque", "", False),
            ("Date d'achat (JJ/MM/AAAA)", "date_achat", "", True),
            ("Grade", "grade", "", False),
            ("Fournisseur / Code", "fournisseur", "", False),
            ("Quantité commandée", "quantite_commandee", "1", False),
            ("Quantité reçue", "quantite_recue", "1", False),
            ("Prix d'achat TTC", "prix_achat", "0", False),
            ("Frais de lavage", "frais_lavage", "0", False),
            ("Date de livraison (JJ/MM/AAAA)", "date_livraison", "", True),
            ("Tracking", "tracking", "", False),
        ]
        article_entry = None
        marque_entry = None
        for label, key, default, is_date in fields[:2]:
            entry = self._field(frame, label, key, default, date_picker=is_date)
            if key == "article":
                article_entry = entry
            else:
                marque_entry = entry
        genre_row = ctk.CTkFrame(frame)
        genre_row.pack(fill="x", padx=12, pady=2)
        ctk.CTkLabel(genre_row, text="Genre", width=220, anchor="w").pack(side="left")
        genre_menu = ctk.CTkOptionMenu(
            genre_row,
            values=["Homme", "Femme", "Mixte"],
            variable=self.genre_var,
            command=lambda _value: self._update_sku_preview(),
        )
        genre_menu.pack(side="left", fill="x", expand=True, padx=(8, 0))
        for label, key, default, is_date in fields[2:]:
            self._field(frame, label, key, default, date_picker=is_date)
        if article_entry is not None:
            article_entry.bind("<KeyRelease>", lambda _event: self._update_sku_preview())
        if marque_entry is not None:
            marque_entry.bind("<KeyRelease>", lambda _event: self._update_sku_preview())
        ctk.CTkLabel(frame, textvariable=self.sku_preview_var, anchor="w").pack(fill="x", padx=12, pady=(6, 0))
        buttons = ctk.CTkFrame(frame)
        buttons.pack(fill="x", padx=12, pady=(12, 4))
        ctk.CTkButton(buttons, text="Annuler", command=self._close).pack(side="right", padx=(8, 0))
        ctk.CTkButton(buttons, text="Ajouter la commande", command=self._submit).pack(side="right")
        ctk.CTkLabel(frame, textvariable=self.log_var, anchor="w").pack(fill="x", padx=12, pady=(4, 0))

    def _field(self, parent, label: str, key: str, default: str, *, date_picker: bool = False):
        row = ctk.CTkFrame(parent)
        row.pack(fill="x", padx=12, pady=2)
        ctk.CTkLabel(row, text=label, width=220, anchor="w").pack(side="left")
        entry = DatePickerEntry(row) if date_picker else ctk.CTkEntry(row)
        if default:
            entry.insert(0, default)
        entry.pack(side="left", fill="x", expand=True, padx=(8, 0))
        self.entries[key] = entry
        return entry

    def _get_entry_value(self, key: str) -> str:
        entry = self.entries.get(key)
        return entry.get().strip() if entry else ""

    def _collect_payload(self) -> dict[str, str]:
        data = {key: self._get_entry_value(key) for key in self.entries}
        data["genre"] = self.genre_var.get()
        return data

    def _submit(self):
        self.log_var.set("")
        payload = self._collect_payload()
        success, message = self.on_submit(payload)
        if success:
            self._close()
        else:
            self.log_var.set(message)


    def _update_sku_preview(self):
        article = self._get_entry_value("article")
        marque = self._get_entry_value("marque")
        genre = self.genre_var.get()
        if not article and not marque:
            self.sku_preview_var.set("Préfixe SKU : --")
            return
        base = self.sku_builder(article, marque, genre)
        self.sku_preview_var.set(f"Préfixe SKU : {base}")

    def _close(self):
        try:
            self.grab_release()
        except tk.TclError:
            pass
        if callable(self.on_close):
            self.on_close()
        self.destroy()


class PurchaseDetailDialog(ctk.CTkToplevel):
    def __init__(self, master, row_index: int, headers, row_data: dict, on_submit, on_close):
        super().__init__(master)
        self.row_index = row_index
        self.headers = headers
        self.row_data = row_data
        self.entries: dict[str, ctk.CTkEntry | DatePickerEntry] = {}
        self.on_submit = on_submit
        self.on_close = on_close
        self.title("Détails de la commande")
        self.geometry("560x780")
        self.minsize(480, 640)
        self.transient(master.winfo_toplevel())
        self.grab_set()
        self._build_form()
        self.protocol("WM_DELETE_WINDOW", self._close)

    def _build_form(self):
        frame = ctk.CTkScrollableFrame(self)
        frame.pack(fill="both", expand=True, padx=16, pady=16)
        purchase_id = self.row_data.get(HEADERS["ACHATS"].ID, "--")
        title = f"Commande #{purchase_id}" if purchase_id not in (None, "") else "Commande"
        ctk.CTkLabel(frame, text=title, font=ctk.CTkFont(size=18, weight="bold"), anchor="w").pack(
            fill="x", padx=12, pady=(8, 12)
        )
        for header in self.headers:
            value = self.row_data.get(header, "")
            self._field(frame, header, value if value is not None else "")
        buttons = ctk.CTkFrame(frame)
        buttons.pack(fill="x", padx=12, pady=(12, 0))
        ctk.CTkButton(buttons, text="Fermer", command=self._close).pack(side="right", padx=(8, 0))
        ctk.CTkButton(buttons, text="Enregistrer", command=self._submit).pack(side="right")

    def _field(self, parent, label: str, value: str):
        row = ctk.CTkFrame(parent)
        row.pack(fill="x", padx=12, pady=2)
        ctk.CTkLabel(row, text=label, width=220, anchor="w").pack(side="left")
        entry = DatePickerEntry(row) if self._is_date_field(label) else ctk.CTkEntry(row)
        if value not in (None, ""):
            entry.insert(0, str(value))
        entry.pack(side="left", fill="x", expand=True, padx=(8, 0))
        self.entries[label] = entry

    @staticmethod
    def _is_date_field(label: str) -> bool:
        return "DATE" in label.upper()

    def _submit(self):
        payload = {key: widget.get().strip() for key, widget in self.entries.items()}
        self.on_submit(payload)
        self._close()

    def _close(self):
        try:
            self.grab_release()
        except tk.TclError:
            pass
        if callable(self.on_close):
            self.on_close()
        self.destroy()


class TableView(ctk.CTkFrame):
    def __init__(self, master, table, on_table_changed=None):
        super().__init__(master)
        self.pack(fill="both", expand=True)
        self.table = table
        self.on_table_changed = on_table_changed
        self.content = ctk.CTkFrame(self)
        self.content.pack(fill="both", expand=True)
        self.content.grid_rowconfigure(0, weight=1)
        self.content.grid_columnconfigure(0, weight=1)

        self.table_frame = ctk.CTkFrame(self.content)
        self.table_frame.grid(row=0, column=0, sticky="nsew")
        helper = ctk.CTkFrame(self.table_frame)
        helper.pack(fill="x", padx=12, pady=(12, 0))
        self.helper_frame = helper
        self.actions_frame = ctk.CTkFrame(helper)
        self.actions_frame.pack(side="left")
        self.status_var = tk.StringVar(value="Double-cliquez sur une cellule pour la modifier.")
        ctk.CTkLabel(helper, textvariable=self.status_var, anchor="w").pack(side="right", fill="x", expand=True, padx=(12, 0))
        self.table_widget: ScrollableTable | None = None
        if self._should_show_table():
            self.table_widget = self._create_table_widget(self.table_frame, self._visible_headers(), table.rows)
            self.table_widget.pack(fill="both", expand=True, padx=12, pady=12)
        self._build_extra_controls(self.content)

    def _create_table_widget(self, parent, headers: Sequence[str], rows: Sequence[dict]):
        return ScrollableTable(
            parent,
            headers,
            rows,
            height=20,
            on_cell_edited=self._on_cell_edit,
            column_width=160,
            column_widths={"ID": 34},
        )

    def _on_cell_edit(self, row_index: int, column: str, new_value: str):
        try:
            self.table.rows[row_index][column] = new_value
        except (IndexError, KeyError):
            pass
        self.status_var.set(f"Ligne {row_index + 1} – {column} mis à jour")

    def refresh(self):
        if self.table_widget is not None:
            self.table_widget.refresh(self.table.rows)

    def _visible_headers(self) -> Sequence[str]:
        return list(self.table.headers[:10])

    def _should_show_table(self) -> bool:
        return True

    def _build_extra_controls(self, parent):
        """Hook for subclasses to add contextual actions."""

    def _delete_rows_by_indices(self, indices: Sequence[int]) -> int:
        removed = 0
        if not indices:
            return removed
        valid_indices = sorted(set(idx for idx in indices if 0 <= idx < len(self.table.rows)), reverse=True)
        for idx in valid_indices:
            try:
                del self.table.rows[idx]
                removed += 1
            except IndexError:
                continue
        if removed:
            self._notify_data_changed()
        return removed

    def _notify_data_changed(self):
        if self.on_table_changed is not None:
            self.on_table_changed()
        else:
            self.refresh()


class StockCardList(ctk.CTkFrame):
    """Compact list of stock items rendered as tappable rectangles."""

    DEFAULT_COLOR = "#e5e7eb"
    SELECTED_COLOR = "#bfdbfe"
    BORDER_COLOR = "#cbd5e1"
    BORDER_COLOR_SELECTED = "#60a5fa"
    GRADIENT_START_COLOR = "#ffffff"
    GRADIENT_END_COLOR = "#8b5cf6"
    GRADIENT_OPACITY = 0.6

    def __init__(self, master, table, *, on_open_details, on_mark_sold, on_bulk_action, on_selection_change=None):
        super().__init__(master)
        self.table = table
        self.on_open_details = on_open_details
        self.on_mark_sold = on_mark_sold
        self.on_bulk_action = on_bulk_action
        self.on_selection_change = on_selection_change
        self._selected_indices: set[int] = set()
        self._cards: dict[int, ctk.CTkFrame] = {}
        self._card_canvases: dict[int, tk.Canvas] = {}
        ctk.CTkLabel(
            self,
            text=(
                "Vue vignettes (clic pour sélectionner, double-clic pour détailler, "
                "clic droit pour marquer vendu ou appliquer une action de groupe)"
            ),
            anchor="w",
            font=ctk.CTkFont(weight="bold"),
        ).pack(fill="x", padx=12, pady=(8, 4))
        self.container = ctk.CTkScrollableFrame(self, height=240)
        self.container.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self.refresh(self.table.rows)

    def refresh(self, rows: Sequence[dict]):
        for child in self.container.winfo_children():
            child.destroy()
        self._cards.clear()
        self._card_canvases.clear()
        self._selected_indices = {idx for idx in self._selected_indices if idx < len(rows)}
        for idx, row in enumerate(rows):
            self._add_card(idx, row)
        self._update_selection_display()

    def get_selected_indices(self) -> list[int]:
        return sorted(self._selected_indices)

    def _add_card(self, index: int, row: dict):
        sku = str(row.get(HEADERS["STOCK"].SKU, "")).strip()
        label = row.get(HEADERS["STOCK"].ARTICLE, "") or row.get(HEADERS["STOCK"].LIBELLE, "")
        status = "Vendu" if row.get(HEADERS["STOCK"].VENDU_ALT, "") else ""
        subtitle = f"{sku} – {label}" if label else sku
        card = ctk.CTkFrame(self.container, height=76, fg_color="transparent", border_width=1)
        card.grid_propagate(False)
        card.pack(fill="x", padx=4, pady=2)
        self._cards[index] = card

        gradient_canvas = tk.Canvas(card, highlightthickness=0, bd=0)
        gradient_canvas.place(relx=0, rely=0, relwidth=1, relheight=1)
        gradient_canvas.bind(
            "<Configure>",
            lambda event, canvas=gradient_canvas, idx=index: self._draw_card_gradient(
                canvas, idx in self._selected_indices, event.width, event.height
            ),
        )
        self._card_canvases[index] = gradient_canvas

        text_frame = ctk.CTkFrame(card, fg_color="transparent")
        text_frame.place(relx=0, rely=0, relwidth=1, relheight=1)
        text_frame.grid_propagate(False)

        content_frame = ctk.CTkFrame(text_frame, fg_color="transparent")
        content_frame.pack(fill="both", expand=True, padx=10, pady=8)

        title = ctk.CTkLabel(
            content_frame,
            text=subtitle or "(SKU manquant)",
            anchor="w",
            font=ctk.CTkFont(weight="bold"),
        )
        title.pack(fill="x")
        metadata_labels = []
        for line in self._build_metadata_lines(row):
            lbl = ctk.CTkLabel(content_frame, text=line, anchor="w", font=ctk.CTkFont(size=12))
            lbl.pack(fill="x")
            metadata_labels.append(lbl)

        if status:
            ctk.CTkLabel(card, text=status, text_color="#0f5132").pack(side="right", padx=8)

        for widget in (card, text_frame, content_frame, title, *metadata_labels):
            self._bind_card_events(widget, index)

    def _bind_card_events(self, widget, index: int):
        widget.bind("<Button-1>", lambda _e, idx=index: self._toggle_selection(idx))
        widget.bind("<Double-Button-1>", lambda _e, idx=index: self.on_open_details(idx))
        widget.bind("<Button-3>", lambda e, idx=index: self._handle_right_click(e, idx))

    def _toggle_selection(self, index: int):
        if index in self._selected_indices:
            self._selected_indices.remove(index)
        else:
            self._selected_indices.add(index)
        self._update_selection_display()

    def _update_selection_display(self):
        for idx in self._cards:
            self._apply_card_style(idx, idx in self._selected_indices)
        if self.on_selection_change is not None:
            self.on_selection_change(self.get_selected_indices())

    def _handle_right_click(self, event, index: int):
        if index not in self._selected_indices:
            self._selected_indices = {index}
            self._update_selection_display()
        click_date = date.today()
        if len(self._selected_indices) > 1:
            self._show_context_menu(event, click_date)
            return
        self.on_mark_sold(index, click_date)

    def _show_context_menu(self, event, click_date: date):
        menu = tk.Menu(self, tearoff=False)
        actions = (
            ("Définir la date de mise en ligne", "mise_en_ligne"),
            ("Définir la date de publication", "publication"),
            ("Définir la date de vente", "vente"),
            ("Déclarer les articles comme vendu", "vendu"),
        )
        for label, action in actions:
            menu.add_command(
                label=label,
                command=lambda act=action: self.on_bulk_action(self.get_selected_indices(), act, click_date),
            )
        menu.tk_popup(event.x_root, event.y_root)

    def _apply_card_style(self, index: int, selected: bool):
        border_color = self.BORDER_COLOR_SELECTED if selected else self.BORDER_COLOR
        card = self._cards.get(index)
        if card is not None:
            card.configure(border_color=border_color)
        canvas = self._card_canvases.get(index)
        if canvas is not None:
            self._draw_card_gradient(canvas, selected)

    def _draw_card_gradient(self, canvas: tk.Canvas, selected: bool, width: int | None = None, height: int | None = None):
        canvas.delete("gradient")
        width = width or canvas.winfo_width()
        height = height or canvas.winfo_height()
        if width <= 0 or height <= 0:
            return
        base_hex = self.SELECTED_COLOR if selected else self.DEFAULT_COLOR
        base_rgb = self._hex_to_rgb(base_hex)
        start_rgb = self._apply_opacity(self.GRADIENT_START_COLOR, base_rgb, self.GRADIENT_OPACITY)
        end_rgb = self._apply_opacity(self.GRADIENT_END_COLOR, base_rgb, self.GRADIENT_OPACITY)
        steps = max(int(height), 1)
        for i in range(steps):
            ratio = i / steps
            blended = self._blend_colors(start_rgb, end_rgb, ratio)
            color = f"#{blended[0]:02x}{blended[1]:02x}{blended[2]:02x}"
            canvas.create_rectangle(0, i, width, i + 1, outline="", fill=color, tags="gradient")

    @staticmethod
    def _hex_to_rgb(value: str) -> tuple[int, int, int]:
        value = value.lstrip("#")
        return tuple(int(value[i : i + 2], 16) for i in (0, 2, 4))

    @staticmethod
    def _apply_opacity(foreground_hex: str, background_rgb: tuple[int, int, int], opacity: float) -> tuple[int, int, int]:
        foreground_rgb = StockCardList._hex_to_rgb(foreground_hex)
        return tuple(
            int(foreground_rgb[i] * opacity + background_rgb[i] * (1 - opacity)) for i in range(3)
        )

    @staticmethod
    def _blend_colors(start: tuple[int, int, int], end: tuple[int, int, int], ratio: float) -> tuple[int, int, int]:
        return tuple(int(start[i] + (end[i] - start[i]) * ratio) for i in range(3))

    def _scroll_to_widget(self, widget: tk.Widget):
        container = self.container
        canvas = getattr(container, "_parent_canvas", None)
        if canvas is None:
            return
        container.update_idletasks()
        widget_y = widget.winfo_y()
        height = container.winfo_height()
        total_height = container.winfo_reqheight()
        if total_height <= height:
            return
        fraction = max(0.0, min(1.0, widget_y / (total_height - height)))
        canvas.yview_moveto(fraction)

    def focus_on_sku(self, query: str) -> bool:
        query_lower = query.lower()
        for idx, row in enumerate(self.table.rows):
            sku = str(row.get(HEADERS["STOCK"].SKU, "")).lower()
            if not sku:
                continue
            if query_lower in sku:
                self._selected_indices = {idx}
                self._update_selection_display()
                card = self._cards.get(idx)
                if card is not None:
                    self._scroll_to_widget(card)
                return True
        return False

    @staticmethod
    def _first_non_empty(row: dict, keys: tuple[str, ...]):
        for key in keys:
            value = row.get(key)
            if value not in (None, ""):
                return value
        return None

    def _format_card_date(self, value) -> str:
        parsed = parse_date_value(value)
        if parsed:
            return format_display_date(parsed)
        return str(value).strip() if value not in (None, "") else ""

    def _build_metadata_lines(self, row: dict) -> list[str]:
        lines: list[str] = []
        mise_date = self._format_card_date(
            self._first_non_empty(
                row,
                (
                    HEADERS["STOCK"].DATE_MISE_EN_LIGNE,
                    HEADERS["STOCK"].DATE_MISE_EN_LIGNE_ALT,
                    HEADERS["STOCK"].MIS_EN_LIGNE,
                    HEADERS["STOCK"].MIS_EN_LIGNE_ALT,
                ),
            )
        )
        if mise_date:
            lines.append(f"Mis en ligne le : {mise_date}")

        publication_date = self._format_card_date(
            self._first_non_empty(
                row,
                (
                    HEADERS["STOCK"].DATE_PUBLICATION,
                    HEADERS["STOCK"].DATE_PUBLICATION_ALT,
                    HEADERS["STOCK"].PUBLIE,
                    HEADERS["STOCK"].PUBLIE_ALT,
                ),
            )
        )
        if publication_date:
            lines.append(f"Publié le : {publication_date}")

        sale_date = self._format_card_date(
            self._first_non_empty(
                row,
                (
                    HEADERS["STOCK"].DATE_VENTE,
                    HEADERS["STOCK"].DATE_VENTE_ALT,
                    HEADERS["STOCK"].VENDU,
                    HEADERS["STOCK"].VENDU_ALT,
                ),
            )
        )
        if sale_date:
            lines.append(f"Vendu le : {sale_date}")

        prix = row.get(HEADERS["STOCK"].PRIX_VENTE)
        if prix not in (None, ""):
            try:
                prix_value = float(prix)
                prix_text = f"{prix_value:.2f} €"
            except (TypeError, ValueError):
                prix_text = str(prix).strip()
            if prix_text:
                lines.append(f"Prix : {prix_text}")

        taille_colis = row.get(HEADERS["STOCK"].TAILLE_COLIS_ALT) or row.get(HEADERS["STOCK"].TAILLE_COLIS, "")
        if taille_colis not in (None, ""):
            lines.append(f"Taille du colis : {taille_colis}")

        lot = row.get(HEADERS["STOCK"].LOT_ALT) or row.get(HEADERS["STOCK"].LOT, "")
        if lot not in (None, ""):
            lines.append(f"Lot : {lot}")

        return lines


class StockSummaryPanel(ctk.CTkFrame):
    """Right-hand column that surfaces key stock metrics."""

    _METRICS = (
        ("stock_pieces", "Articles en stock"),
        ("stock_value", "Valeur estimée du stock"),
        ("reference_count", "Références uniques"),
        ("value_per_reference", "Valeur moyenne / référence"),
        ("value_per_piece", "Prix moyen / article"),
    )

    def __init__(self, master, achats_rows: Sequence[Mapping] | None = None):
        super().__init__(master)
        self.achats_rows: Sequence[Mapping] = achats_rows or []
        ctk.CTkLabel(
            self,
            text="Indicateurs stock",
            font=ctk.CTkFont(size=16, weight="bold"),
            anchor="w",
        ).pack(fill="x", padx=12, pady=(12, 6))

        self.value_labels: dict[str, ctk.CTkLabel] = {}
        for key, label in self._METRICS:
            card = ctk.CTkFrame(self)
            card.pack(fill="x", padx=12, pady=6)
            ctk.CTkLabel(card, text=label, anchor="w").pack(side="left", padx=8, pady=10)
            value_label = ctk.CTkLabel(card, text="-", font=ctk.CTkFont(weight="bold"))
            value_label.pack(side="right", padx=8, pady=10)
            self.value_labels[key] = value_label

    def update(self, rows: Sequence[Mapping], achats_rows: Sequence[Mapping] | None = None):
        if achats_rows is not None:
            self.achats_rows = achats_rows
        stats = self._compute_stats(rows, self.achats_rows)
        self.value_labels["stock_pieces"].configure(text=str(stats["stock_pieces"]))
        self.value_labels["stock_value"].configure(text=f"{stats['stock_value']:.2f} €")
        self.value_labels["reference_count"].configure(text=str(stats["reference_count"]))
        self.value_labels["value_per_reference"].configure(text=f"{stats['value_per_reference']:.2f} €")
        self.value_labels["value_per_piece"].configure(text=f"{stats['value_per_piece']:.2f} €")

    @staticmethod
    def _compute_stats(
        rows: Sequence[Mapping], achats_rows: Sequence[Mapping] | None = None
    ) -> dict[str, float | int]:
        pieces = 0
        references: set[str] = set()

        for row in rows:
            vendu = row.get(HEADERS["STOCK"].VENDU_ALT) or row.get(HEADERS["STOCK"].VENDU)
            if vendu:
                continue
            pieces += 1
            base_reference = _base_reference_from_stock(row)
            if base_reference:
                references.add(base_reference)

        base_counts = {base: 0 for base in references}
        for row in rows:
            vendu = row.get(HEADERS["STOCK"].VENDU_ALT) or row.get(HEADERS["STOCK"].VENDU)
            if vendu:
                continue
            base_reference = _base_reference_from_stock(row)
            if base_reference and base_reference in base_counts:
                base_counts[base_reference] += 1

        reference_unit_price = _build_reference_unit_price_index(achats_rows or [])
        stock_value = sum(reference_unit_price.get(base, 0.0) * count for base, count in base_counts.items())

        reference_count = len(references)
        value_per_reference = stock_value / reference_count if reference_count else 0.0
        value_per_piece = stock_value / pieces if pieces else 0.0

        return {
            "stock_pieces": pieces,
            "stock_value": round(stock_value, 2),
            "reference_count": reference_count,
            "value_per_reference": round(value_per_reference, 2),
            "value_per_piece": round(value_per_piece, 2),
        }


class StockDetailDialog(ctk.CTkToplevel):
    """Lightweight popup used to enrich a stock item."""

    def __init__(self, master, row: dict, *, on_save):
        super().__init__(master)
        self.on_save = on_save
        self.title("Détails de l'article")
        self.geometry("420x360")
        self.resizable(False, False)
        owner = master.winfo_toplevel() if hasattr(master, "winfo_toplevel") else None
        if owner is not None:
            self.transient(owner)
        self.grab_set()
        self.lift()
        self.focus()
        self.row = row
        self._fields: dict[str, ctk.CTkEntry | DatePickerEntry] = {}
        form = ctk.CTkFrame(self)
        form.pack(fill="both", expand=True, padx=12, pady=12)
        self._add_field(
            form,
            "date_mise_en_ligne",
            "Date de mise en ligne",
            self._initial_date(
                row,
                StockTableView._DATE_COLUMNS,
                HEADERS["STOCK"].DATE_MISE_EN_LIGNE,
                HEADERS["STOCK"].MIS_EN_LIGNE,
            ),
            date_picker=True,
        )
        self._add_field(
            form,
            "date_publication",
            "Date de publication",
            self._initial_date(
                row,
                StockTableView._PUBLICATION_COLUMNS,
                HEADERS["STOCK"].DATE_PUBLICATION,
                HEADERS["STOCK"].PUBLIE,
            ),
            date_picker=True,
        )
        self._add_field(
            form,
            "date_vente",
            "Date de vente",
            self._initial_date(
                row,
                StockTableView._SALE_COLUMNS,
                HEADERS["STOCK"].DATE_VENTE,
                HEADERS["STOCK"].VENDU,
            ),
            date_picker=True,
        )
        self._add_field(
            form,
            "prix_vente",
            "Prix de vente",
            row.get(HEADERS["STOCK"].PRIX_VENTE, ""),
        )
        self._add_field(
            form,
            "taille_colis",
            "Taille du colis",
            row.get(HEADERS["STOCK"].TAILLE_COLIS) or row.get(HEADERS["STOCK"].TAILLE_COLIS_ALT, ""),
        )
        self._add_field(
            form,
            "lot",
            "Lot",
            row.get(HEADERS["STOCK"].LOT_ALT) or row.get(HEADERS["STOCK"].LOT, ""),
        )
        ctk.CTkButton(form, text="Enregistrer", command=self._save).pack(fill="x", pady=(12, 4))

    def _add_field(self, parent, key: str, label: str, value, *, date_picker: bool = False):
        row = ctk.CTkFrame(parent)
        row.pack(fill="x", pady=4)
        ctk.CTkLabel(row, text=label, width=160, anchor="w").pack(side="left")
        entry = DatePickerEntry(row) if date_picker else ctk.CTkEntry(row)
        if value not in (None, ""):
            entry.insert(0, str(value))
        entry.pack(side="left", fill="x", expand=True, padx=(8, 0))
        self._fields[key] = entry

    @staticmethod
    def _initial_date(row: dict, aliases: set[str], *keys: str) -> str:
        for key in (*keys, *aliases):
            value = row.get(key)
            if value not in (None, ""):
                return value
        return ""

    def _save(self):
        updates = {name: field.get().strip() for name, field in self._fields.items()}
        self.on_save(updates)
        self.destroy()


class StockTableView(TableView):
    DISPLAY_HEADERS: Sequence[str] = DEFAULT_STOCK_HEADERS

    SIZE_CHOICES: Sequence[str] = ("Petit", "Moyen", "Grand")
    LOT_CHOICES: Sequence[str] = ("2", "3", "4", "5")

    _DISPLAY_FALLBACKS = {
        HEADERS["STOCK"].LOT_ALT: HEADERS["STOCK"].LOT,
        HEADERS["STOCK"].VALIDER_SAISIE_ALT: HEADERS["STOCK"].VALIDER_SAISIE,
        HEADERS["STOCK"].PUBLIE_ALT: HEADERS["STOCK"].PUBLIE,
        HEADERS["STOCK"].DATE_PUBLICATION_ALT: HEADERS["STOCK"].DATE_PUBLICATION,
    }

    _DATE_COLUMNS = {
        HEADERS["STOCK"].MIS_EN_LIGNE,
        HEADERS["STOCK"].MIS_EN_LIGNE_ALT,
        HEADERS["STOCK"].DATE_MISE_EN_LIGNE,
        HEADERS["STOCK"].DATE_MISE_EN_LIGNE_ALT,
    }

    _PUBLICATION_COLUMNS = {
        HEADERS["STOCK"].PUBLIE,
        HEADERS["STOCK"].PUBLIE_ALT,
        HEADERS["STOCK"].DATE_PUBLICATION,
        HEADERS["STOCK"].DATE_PUBLICATION_ALT,
    }

    _SALE_COLUMNS = {
        HEADERS["STOCK"].VENDU,
        HEADERS["STOCK"].VENDU_ALT,
        HEADERS["STOCK"].DATE_VENTE,
        HEADERS["STOCK"].DATE_VENTE_ALT,
    }

    _CHECKBOX_STYLE_COLUMNS = (
        _DATE_COLUMNS
        | _PUBLICATION_COLUMNS
        | _SALE_COLUMNS
    )

    def __init__(self, master, table, *, workflow: WorkflowCoordinator | None = None, on_table_changed=None):
        self.workflow = workflow
        self.search_var: tk.StringVar | None = None
        self.search_entry: ctk.CTkEntry | None = None
        super().__init__(master, table, on_table_changed=on_table_changed)
        self.status_var.set(
            "Sélectionnez une vignette (clic), double-clic pour détailler, clic droit pour actions rapides."
        )

    def _create_table_widget(self, parent, headers: Sequence[str], rows: Sequence[dict]):
        return ScrollableTable(
            parent,
            headers,
            rows,
            height=20,
            on_cell_edited=self._on_cell_edit,
            on_cell_activated=self._handle_cell_activation,
            column_width=160,
            column_widths={"ID": 34},
            value_formatter=self._format_cell_value,
            dropdown_choices={
                HEADERS["STOCK"].TAILLE: self.SIZE_CHOICES,
                HEADERS["STOCK"].LOT_ALT: self.LOT_CHOICES,
            },
        )

    def _visible_headers(self) -> Sequence[str]:
        self._ensure_display_aliases()
        return self.DISPLAY_HEADERS

    def _should_show_table(self) -> bool:
        return False

    def refresh(self):
        self._ensure_display_aliases()
        super().refresh()
        if hasattr(self, "card_list"):
            self.card_list.refresh(self.table.rows)
        if hasattr(self, "summary_panel"):
            achats_rows = self.workflow.achats.rows if self.workflow is not None else None
            self.summary_panel.update(self.table.rows, achats_rows)

    def _build_extra_controls(self, parent):
        parent.grid_columnconfigure(0, weight=2, uniform="stock")
        parent.grid_columnconfigure(1, weight=1, uniform="stock")
        parent.grid_rowconfigure(0, weight=0)
        self.table_frame.grid_configure(columnspan=2, sticky="nsew")
        ctk.CTkButton(
            self.actions_frame,
            text="Supprimer la sélection",
            command=self._delete_selected_rows,
        ).pack(side="left")
        search_frame = ctk.CTkFrame(self.actions_frame, fg_color="transparent")
        search_frame.pack(side="left", padx=(12, 0))
        ctk.CTkLabel(search_frame, text="Rechercher SKU:").pack(side="left", padx=(0, 6))
        self.search_var = tk.StringVar()
        self.search_entry = ctk.CTkEntry(search_frame, textvariable=self.search_var, width=180)
        self.search_entry.pack(side="left")
        self.search_entry.bind("<Return>", lambda _e: self._handle_sku_search())
        ctk.CTkButton(search_frame, text="Go", width=56, command=self._handle_sku_search).pack(
            side="left", padx=(8, 0)
        )

        parent.grid_rowconfigure(1, weight=1)
        self.card_list = StockCardList(
            parent,
            self.table,
            on_open_details=self._open_card_details,
            on_mark_sold=self._handle_card_sale,
            on_bulk_action=self._handle_bulk_card_action,
            on_selection_change=self._on_card_selection_change,
        )
        self.card_list.grid(row=1, column=0, sticky="nsew", padx=(12, 6), pady=(0, 12))

        achats_rows = self.workflow.achats.rows if self.workflow is not None else None
        self.summary_panel = StockSummaryPanel(parent, achats_rows=achats_rows)
        self.summary_panel.grid(row=1, column=1, sticky="nsew", padx=(6, 12), pady=(0, 12))
        self.summary_panel.update(self.table.rows, achats_rows)

    def _delete_selected_rows(self):
        indices = self.table_widget.get_selected_indices() if self.table_widget is not None else []
        if not indices:
            self.status_var.set("Sélectionnez au moins une ligne avant de supprimer.")
            return
        count = len(indices)
        if not messagebox.askyesno("Confirmer la suppression", f"Supprimer {count} article(s) sélectionné(s) du stock ?"):
            return
        removed = self._delete_rows_by_indices(indices)
        if removed:
            self.status_var.set(f"{removed} article(s) supprimé(s) du stock.")
        else:
            self.status_var.set("Aucune ligne supprimée.")

    def _on_cell_edit(self, row_index: int, column: str, new_value: str):
        target_column = self._DISPLAY_FALLBACKS.get(column, column)
        try:
            self.table.rows[row_index][target_column] = new_value
            if target_column != column:
                self.table.rows[row_index][column] = new_value
        except (IndexError, KeyError):
            pass
        self.status_var.set(f"Ligne {row_index + 1} – {column} mis à jour")

    def _handle_cell_activation(self, row_index: int | None, column: str) -> bool:
        if row_index is None:
            return False
        if not (0 <= row_index < len(self.table.rows)):
            return False
        today = format_display_date(date.today())
        row = self.table.rows[row_index]
        if column in self._DATE_COLUMNS:
            columns = self._DATE_COLUMNS
            set_message = "Date de mise en ligne renseignée"
            clear_message = "Date de mise en ligne effacée"
        elif column in self._PUBLICATION_COLUMNS:
            columns = self._PUBLICATION_COLUMNS
            set_message = "Date de publication renseignée"
            clear_message = "Date de publication effacée"
        elif column in self._SALE_COLUMNS:
            columns = self._SALE_COLUMNS
            set_message = "Date de vente renseignée"
            clear_message = "Date de vente effacée"
        else:
            return False

        has_date = any(str(row.get(key, "")).strip() for key in columns)
        new_value = "" if has_date else today
        for key in columns:
            row[key] = new_value
        if self.table_widget is not None:
            self.table_widget.refresh(self.table.rows)
        status = clear_message if has_date else set_message
        self.status_var.set(f"{status} pour la ligne {row_index + 1}")
        self._notify_data_changed()
        return True

    def _handle_sku_search(self):
        if self.search_var is None:
            return
        query = self.search_var.get().strip()
        if not query:
            self.status_var.set("Saisissez un SKU à rechercher.")
            return
        found = self.card_list.focus_on_sku(query)
        if found:
            self.status_var.set(f"Vignette trouvée pour le SKU contenant '{query}'.")
        else:
            self.status_var.set(f"Aucun article ne correspond au SKU '{query}'.")

    def _ensure_display_aliases(self):
        for row in self.table.rows:
            for header in self.DISPLAY_HEADERS:
                row.setdefault(header, "")
            for alias, source in self._DISPLAY_FALLBACKS.items():
                if alias in row and row[alias] not in (None, ""):
                    continue
                row[alias] = row.get(source, "")

    def _format_cell_value(self, header: str, value: object) -> str:
        if header in self._CHECKBOX_STYLE_COLUMNS and not _has_ready_date(value):
            return "☐"
        return "" if value is None else str(value)

    def _open_card_details(self, row_index: int):
        if not (0 <= row_index < len(self.table.rows)):
            return
        row = self.table.rows[row_index]

        def _apply(updates: dict[str, str]):
            mise = self._normalize_date_input(updates.get("date_mise_en_ligne", ""))
            publication = self._normalize_date_input(updates.get("date_publication", ""))
            vente = self._normalize_date_input(updates.get("date_vente", ""))
            prix = updates.get("prix_vente", "")
            taille_colis = updates.get("taille_colis", "")
            lot = updates.get("lot", "")
            if mise is not None:
                for key in self._DATE_COLUMNS:
                    row[key] = mise
            if publication is not None:
                for key in self._PUBLICATION_COLUMNS:
                    row[key] = publication
            if vente is not None:
                for key in self._SALE_COLUMNS:
                    row[key] = vente
            if prix:
                try:
                    row[HEADERS["STOCK"].PRIX_VENTE] = round(float(prix), 2)
                except ValueError:
                    pass
            if taille_colis is not None:
                row[HEADERS["STOCK"].TAILLE_COLIS] = taille_colis
                row[HEADERS["STOCK"].TAILLE_COLIS_ALT] = taille_colis
            if lot is not None:
                row[HEADERS["STOCK"].LOT] = lot
                row[HEADERS["STOCK"].LOT_ALT] = lot
            self.status_var.set(f"Ligne {row_index + 1} mise à jour via vignette")
            self.refresh()
            self._notify_data_changed()

        StockDetailDialog(self, row, on_save=_apply)

    def _on_card_selection_change(self, indices: Sequence[int]):
        count = len(indices)
        self.status_var.set(f"{count} vignette(s) sélectionnée(s)")

    def _ensure_sale_requirements(self, row: dict) -> bool:
        price = row.get(HEADERS["STOCK"].PRIX_VENTE)
        package_size = row.get(HEADERS["STOCK"].TAILLE_COLIS) or row.get(HEADERS["STOCK"].TAILLE_COLIS_ALT)
        if price in (None, ""):
            return False
        try:
            float(price)
        except (TypeError, ValueError):
            return False
        return package_size not in (None, "")

    def _set_sale_columns(self, row: dict, value: str):
        for key in self._SALE_COLUMNS:
            row[key] = value

    def _handle_card_sale(self, row_index: int, click_date: date):
        if not (0 <= row_index < len(self.table.rows)):
            return
        row = self.table.rows[row_index]
        if not self._ensure_sale_requirements(row):
            self.status_var.set(
                "Impossible de déclarer la vente : prix de vente et taille du colis requis."
            )
            return
        date_text = format_display_date(click_date)
        sale_value = f"Vendu le {date_text}"
        self._set_sale_columns(row, sale_value)
        self.status_var.set(f"Article {row_index + 1} marqué vendu le {date_text}")
        self.refresh()
        self._notify_data_changed()

    def _apply_date_to_rows(self, indices: Sequence[int], columns: set[str], label: str, date_text: str):
        updated = 0
        for idx in indices:
            if not (0 <= idx < len(self.table.rows)):
                continue
            row = self.table.rows[idx]
            for key in columns:
                row[key] = date_text
            updated += 1
        if updated:
            self.status_var.set(f"{label} mise à jour pour {updated} article(s)")
            self.refresh()
            self._notify_data_changed()

    def _handle_bulk_card_action(self, indices: Sequence[int], action: str, click_date: date):
        if not indices:
            self.status_var.set("Aucune vignette sélectionnée")
            return
        date_text = format_display_date(click_date)
        if action == "mise_en_ligne":
            self._apply_date_to_rows(indices, self._DATE_COLUMNS, "Date de mise en ligne", date_text)
            return
        if action == "publication":
            self._apply_date_to_rows(indices, self._PUBLICATION_COLUMNS, "Date de publication", date_text)
            return
        if action == "vente":
            self._apply_date_to_rows(indices, self._SALE_COLUMNS, "Date de vente", date_text)
            return
        if action == "vendu":
            success = 0
            failures = 0
            sale_value = f"Vendu le {date_text}"
            for idx in indices:
                if not (0 <= idx < len(self.table.rows)):
                    continue
                row = self.table.rows[idx]
                if not self._ensure_sale_requirements(row):
                    failures += 1
                    continue
                self._set_sale_columns(row, sale_value)
                success += 1
            messages: list[str] = []
            if success:
                messages.append(f"Articles vendus mis à jour pour {success} vignette(s)")
                self.refresh()
                self._notify_data_changed()
            if failures:
                messages.append(
                    f"{failures} vignette(s) non mises à jour : prix de vente et taille du colis requis"
                )
            if messages:
                self.status_var.set(" ".join(messages))
            return

    @staticmethod
    def _normalize_date_input(value: str | None) -> str | None:
        if value is None:
            return None
        text = value.strip()
        if not text:
            return ""
        parsed = parse_date_value(text)
        return format_display_date(parsed) if parsed else text


class WorkflowView(ctk.CTkFrame):
    def __init__(self, master, coordinator: WorkflowCoordinator, refresh_callback):
        super().__init__(master)
        self.coordinator = coordinator
        self.refresh_callback = refresh_callback
        self.pack(fill="both", expand=True)
        self.log_var = tk.StringVar(value="Prêt à traiter vos flux Achats → Stock → Ventes.")

        title = ctk.CTkLabel(self, text="Automatiser les flux", font=ctk.CTkFont(size=20, weight="bold"))
        title.pack(pady=(16, 8))

        self._build_purchase_section()
        self._build_stock_section()
        self._build_sale_section()
        self._build_return_section()

        log_frame = ctk.CTkFrame(self)
        log_frame.pack(fill="x", padx=16, pady=(8, 16))
        ctk.CTkLabel(log_frame, textvariable=self.log_var, anchor="w").pack(fill="x", padx=12, pady=12)

    # ------------------------------------------------------------------
    # Sections builders
    # ------------------------------------------------------------------
    def _build_purchase_section(self):
        frame = self._section("1. Ajouter une commande dans Achats")
        self.purchase_article = self._field(frame, "Article")
        self.purchase_marque = self._field(frame, "Marque")
        self.purchase_reference = self._field(frame, "Référence")
        self.purchase_quantite = self._field(frame, "Quantité", default="1")
        self.purchase_prix = self._field(frame, "Prix unitaire TTC", default="0")
        self.purchase_frais = self._field(frame, "Frais de colissage", default="0")
        ctk.CTkButton(frame, text="Créer l'achat", command=self._handle_add_purchase).pack(pady=(8, 4))

    def _build_stock_section(self):
        frame = self._section("2. Générer les articles en stock")
        self.stock_purchase_id = self._field(frame, "ID Achat")
        self.stock_sku = self._field(frame, "SKU")
        self.stock_prix = self._field(frame, "Prix de vente", default="0")
        self.stock_lot = self._field(frame, "Lot")
        self.stock_taille = self._field(frame, "Taille")
        ctk.CTkButton(frame, text="Passer en stock", command=self._handle_stock_transfer).pack(pady=(8, 4))

    def _build_sale_section(self):
        frame = self._section("3. Basculer un article vendu vers Ventes/Compta")
        self.sale_sku = self._field(frame, "SKU")
        self.sale_prix = self._field(frame, "Prix de vente", default="0")
        self.sale_frais = self._field(frame, "Frais colissage", default="0")
        self.sale_date = self._field(frame, "Date de vente (JJ/MM/AAAA)", date_picker=True)
        self.sale_lot = self._field(frame, "Lot")
        self.sale_taille = self._field(frame, "Taille")
        ctk.CTkButton(frame, text="Enregistrer la vente", command=self._handle_sale).pack(pady=(8, 4))

    def _build_return_section(self):
        frame = self._section("4. Gérer un retour")
        self.return_sku = self._field(frame, "SKU")
        self.return_note = self._field(frame, "Motif")
        ctk.CTkButton(frame, text="Enregistrer le retour", command=self._handle_return).pack(pady=(8, 4))

    # ------------------------------------------------------------------
    # Actions
    # ------------------------------------------------------------------
    def _handle_add_purchase(self):
        try:
            data = PurchaseInput(
                article=self.purchase_article.get().strip(),
                marque=self.purchase_marque.get().strip(),
                reference=self.purchase_reference.get().strip(),
                quantite=int(self.purchase_quantite.get() or "0"),
                prix_unitaire=float(self.purchase_prix.get() or "0"),
                frais_colissage=float(self.purchase_frais.get() or "0"),
            )
        except ValueError:
            self._log("Valeurs numériques invalides pour l'achat")
            return
        if not data.article or not data.marque:
            self._log("Article et marque sont obligatoires")
            return
        row = self.coordinator.create_purchase(data)
        self._log(f"Achat {row.get(HEADERS['ACHATS'].ID)} ajouté")
        self.refresh_callback()

    def _handle_stock_transfer(self):
        try:
            prix = float(self.stock_prix.get() or "0")
        except ValueError:
            self._log("Prix de vente invalide")
            return
        purchase_id = self.stock_purchase_id.get().strip()
        sku = self.stock_sku.get().strip()
        if not purchase_id or not sku:
            self._log("L'ID achat et le SKU sont obligatoires")
            return
        data = StockInput(
            purchase_id=purchase_id,
            sku=sku,
            prix_vente=prix,
            lot=self.stock_lot.get().strip(),
            taille=self.stock_taille.get().strip(),
        )
        try:
            row = self.coordinator.transfer_to_stock(data)
        except ValueError as exc:
            self._log(str(exc))
            return
        self._log(f"SKU {row.get(HEADERS['STOCK'].SKU)} créé dans le stock")
        self.refresh_callback()

    def _handle_sale(self):
        try:
            prix = float(self.sale_prix.get() or "0")
            frais = float(self.sale_frais.get() or "0")
        except ValueError:
            self._log("Montants invalides pour la vente")
            return
        sku = self.sale_sku.get().strip()
        if not sku:
            self._log("Le SKU est obligatoire pour enregistrer une vente")
            return
        data = SaleInput(
            sku=sku,
            prix_vente=prix,
            frais_colissage=frais,
            date_vente=self.sale_date.get().strip() or None,
            lot=self.sale_lot.get().strip(),
            taille=self.sale_taille.get().strip(),
        )
        try:
            row = self.coordinator.register_sale(data)
        except ValueError as exc:
            self._log(str(exc))
            return
        self._log(f"Vente {row.get(HEADERS['VENTES'].ID)} envoyée vers Ventes/Compta")
        self.refresh_callback()

    def _handle_return(self):
        sku = self.return_sku.get().strip()
        if not sku:
            self._log("Le SKU est obligatoire pour un retour")
            return
        try:
            row = self.coordinator.register_return(sku, self.return_note.get().strip())
        except ValueError as exc:
            self._log(str(exc))
            return
        self._log(f"Retour enregistré pour le SKU {row.get(HEADERS['VENTES'].SKU)}")
        self.refresh_callback()

    # ------------------------------------------------------------------
    # UI helpers
    # ------------------------------------------------------------------
    def _section(self, title: str):
        frame = ctk.CTkFrame(self)
        frame.pack(fill="x", padx=16, pady=8)
        ctk.CTkLabel(frame, text=title, font=ctk.CTkFont(size=16, weight="bold"), anchor="w").pack(
            fill="x", padx=12, pady=(12, 4)
        )
        return frame

    def _field(self, parent, label: str, *, default: str | None = None, date_picker: bool = False):
        row = ctk.CTkFrame(parent)
        row.pack(fill="x", padx=12, pady=4)
        ctk.CTkLabel(row, text=label, width=180, anchor="w").pack(side="left")
        entry = DatePickerEntry(row) if date_picker else ctk.CTkEntry(row)
        if default is not None:
            entry.insert(0, default)
        entry.pack(side="left", fill="x", expand=True, padx=(8, 0))
        return entry

    def _log(self, message: str):
        self.log_var.set(message)


class CalendarView(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master)
        self.pack(fill="both", expand=True)
        label = ctk.CTkLabel(self, text="Calendrier comptable", font=ctk.CTkFont(size=18, weight="bold"))
        label.pack(pady=12)
        listbox = tk.Listbox(self)
        for month in MONTH_NAMES_FR:
            listbox.insert(tk.END, month)
        listbox.pack(fill="both", expand=True, padx=32, pady=16)


class StockOptionsView(ctk.CTkFrame):
    def __init__(self, master, table: TableData, refresh_callback):
        super().__init__(master)
        self.table = table
        self.refresh_callback = refresh_callback
        self.pack(fill="both", expand=True)

        self.status_var = tk.StringVar(value="Importer ou ajuster les données du stock.")

        title = ctk.CTkLabel(self, text="Options", font=ctk.CTkFont(size=20, weight="bold"))
        title.pack(pady=(16, 8))

        frame = ctk.CTkFrame(self)
        frame.pack(fill="x", padx=16, pady=12)

        ctk.CTkLabel(frame, text="Importer des articles", anchor="w", font=ctk.CTkFont(weight="bold")).pack(
            fill="x", pady=(12, 4)
        )
        ctk.CTkLabel(
            frame,
            text="Sélectionnez un export Excel pour ajouter les nouveaux articles (sans doublons).",
            anchor="w",
            wraplength=800,
            justify="left",
        ).pack(fill="x", padx=12)
        ctk.CTkButton(frame, text="Charger un XLSX pour le stock", command=self._handle_import_xlsx).pack(
            pady=(8, 12), padx=12
        )

        ctk.CTkLabel(self, textvariable=self.status_var, anchor="w").pack(fill="x", padx=16, pady=(4, 16))

    def _handle_import_xlsx(self):
        path = filedialog.askopenfilename(
            title="Importer le stock",
            filetypes=(
                ("Excel", "*.xlsx *.xlsm"),
                ("Tous les fichiers", "*.*"),
            ),
        )
        if not path:
            return
        try:
            repository = WorkbookRepository(path)
            sheet_name = self._resolve_stock_sheet(repository)
            if sheet_name is None:
                raise ValueError("Ce classeur ne contient aucun onglet exploitable.")
            source_table = repository.load_table(sheet_name)
        except Exception as exc:  # pragma: no cover - UI guard
            messagebox.showerror("Import du stock", f"Impossible de lire le fichier sélectionné : {exc}")
            return
        added = merge_stock_table(self.table, source_table)
        filename = Path(path).name
        if added:
            self.status_var.set(f"{added} article(s) importé(s) depuis {filename}.")
            self.refresh_callback()
        else:
            self.status_var.set(f"Aucun nouvel article à importer depuis {filename}.")

    @staticmethod
    def _resolve_stock_sheet(repository: WorkbookRepository) -> str | None:
        sheet_names = list(repository.available_tables())
        if not sheet_names:
            return None
        for name in sheet_names:
            if name.lower() == "stock":
                return name
        return sheet_names[0]


DEFAULT_WORKBOOK = Path(__file__).resolve().parent.parent / "Prerelease 1.2.xlsx"
DEFAULT_ACHATS_DB = Path(__file__).resolve().parent / "data" / "achats.db"


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Vintage ERP UI")
    parser.add_argument(
        "workbook",
        nargs="?",
        default=DEFAULT_WORKBOOK,
        type=Path,
        help="Path to the Excel workbook (defaults to Prerelease 1.2.xlsx located at the repo root)",
    )
    parser.add_argument(
        "--achats-db",
        type=Path,
        default=None,
        help=(
            "Chemin vers la base SQLite utilisée pour l'onglet Achats. "
            "Par défaut, python_app/data/achats.db est utilisé et créé si nécessaire."
        ),
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    workbook_path = Path(args.workbook)
    try:
        repo = WorkbookRepository(workbook_path)
    except FileNotFoundError:
        messagebox.showerror("Workbook introuvable", f"Impossible d'ouvrir {workbook_path!s}")
        return 1
    custom_db_requested = args.achats_db is not None
    achats_db_path = Path(args.achats_db) if custom_db_requested else DEFAULT_ACHATS_DB
    achats_table: TableData | None = None
    stock_table: TableData | None = None
    if achats_db_path.exists():
        db = PurchaseDatabase(achats_db_path)
        try:
            loaded_table = db.load_table()
        except FileNotFoundError:
            messagebox.showerror("Base Achats introuvable", f"Impossible d'ouvrir {achats_db_path!s}")
            return 1
        if loaded_table.rows:
            achats_table = loaded_table
        else:
            # L'application créait un fichier SQLite vide au premier lancement,
            # puis se contentait de charger ce contenu vide au démarrage
            # suivant.  Ignorez les bases sans lignes pour retomber sur les
            # données du classeur Excel et repeupler la base à la fermeture.
            print(
                f"Base Achats vide détectée ({achats_db_path!s}). "
                "Chargement des données depuis le classeur."
            )
        try:
            loaded_stock = db.load_stock_table()
        except FileNotFoundError:
            messagebox.showerror("Base Achats introuvable", f"Impossible d'ouvrir {achats_db_path!s}")
            return 1
        if loaded_stock.rows:
            stock_table = loaded_stock
    elif custom_db_requested:
        messagebox.showerror("Base Achats introuvable", f"Impossible d'ouvrir {achats_db_path!s}")
        return 1
    app = VintageErpApp(
        repo,
        achats_table=achats_table,
        stock_table=stock_table,
        achats_db_path=achats_db_path,
    )
    app.mainloop()
    return 0


if __name__ == "__main__":  # pragma: no cover - CLI entry point
    raise SystemExit(main())
