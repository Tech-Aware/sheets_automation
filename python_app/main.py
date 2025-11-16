"""Entry point for the CustomTkinter application."""
from __future__ import annotations

import argparse
from datetime import date
import math
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Sequence

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
    from .services.summaries import build_inventory_snapshot
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
    from python_app.services.summaries import build_inventory_snapshot
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
        summary = build_inventory_snapshot(self.tables["Stock"].rows, self.tables["Ventes"].rows)

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
                view = StockTableView(tab, self.tables[sheet], on_table_changed=self.refresh_views)
            else:
                view = TableView(tab, self.tables[sheet], on_table_changed=self.refresh_views)
            self.table_views[sheet] = view

        months_tab = self.tabview.add("Calendrier")
        CalendarView(months_tab)

        workflow_tab = self.tabview.add("Workflow")
        WorkflowView(workflow_tab, self.workflow, self.refresh_views)

    def refresh_views(self):
        summary = build_inventory_snapshot(self.tables["Stock"].rows, self.tables["Ventes"].rows)
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

        table_frame = ctk.CTkFrame(self.content)
        table_frame.grid(row=0, column=0, sticky="nsew")
        helper = ctk.CTkFrame(table_frame)
        helper.pack(fill="x", padx=12, pady=(12, 0))
        self.helper_frame = helper
        self.actions_frame = ctk.CTkFrame(helper)
        self.actions_frame.pack(side="left")
        self.status_var = tk.StringVar(value="Double-cliquez sur une cellule pour la modifier.")
        ctk.CTkLabel(helper, textvariable=self.status_var, anchor="w").pack(side="right", fill="x", expand=True, padx=(12, 0))
        self.table_widget = self._create_table_widget(table_frame, self._visible_headers(), table.rows)
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
        self.table_widget.refresh(self.table.rows)

    def _visible_headers(self) -> Sequence[str]:
        return list(self.table.headers[:10])

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


class StockTableView(TableView):
    DISPLAY_HEADERS: Sequence[str] = DEFAULT_STOCK_HEADERS

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

    def __init__(self, master, table, on_table_changed=None):
        super().__init__(master, table, on_table_changed=on_table_changed)

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
        )

    def _visible_headers(self) -> Sequence[str]:
        self._ensure_display_aliases()
        return self.DISPLAY_HEADERS

    def refresh(self):
        self._ensure_display_aliases()
        super().refresh()

    def _build_extra_controls(self, parent):
        parent.grid_columnconfigure(0, weight=5)
        parent.grid_columnconfigure(1, weight=0)
        ctk.CTkButton(
            self.actions_frame,
            text="Supprimer la sélection",
            command=self._delete_selected_rows,
        ).pack(side="left")

        panel = ctk.CTkFrame(parent)
        panel.grid(row=0, column=1, sticky="nsew", padx=(12, 0), pady=12)

        import_frame = ctk.CTkFrame(panel)
        import_frame.pack(fill="x", padx=12, pady=(0, 8))
        ctk.CTkLabel(import_frame, text="Importer des articles", anchor="w", font=ctk.CTkFont(weight="bold")).pack(
            fill="x", pady=(4, 4)
        )
        ctk.CTkButton(import_frame, text="Depuis un fichier XLSX", command=self._handle_import_xlsx).pack(fill="x")
        ctk.CTkLabel(
            panel,
            text="Sélectionnez un export Excel pour ajouter les nouveaux articles (sans doublons).",
            anchor="w",
            wraplength=260,
        ).pack(fill="x", padx=12, pady=(0, 8))

    def _delete_selected_rows(self):
        indices = self.table_widget.get_selected_indices()
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
            message_set = "Date de mise en ligne renseignée"
            message_cleared = "Statut de mise en ligne réinitialisé"
        elif column in self._PUBLICATION_COLUMNS:
            columns = self._PUBLICATION_COLUMNS
            message_set = "Date de publication renseignée"
            message_cleared = "Statut de publication réinitialisé"
        elif column in self._SALE_COLUMNS:
            columns = self._SALE_COLUMNS
            message_set = "Date de vente renseignée"
            message_cleared = "Statut de vente réinitialisé"
        else:
            return False

        has_date = any(_has_ready_date(row.get(key)) for key in columns)
        new_value = "" if has_date else today
        for key in columns:
            row[key] = new_value
        self.table_widget.refresh(self.table.rows)
        message = message_cleared if has_date else message_set
        self.status_var.set(f"{message} pour la ligne {row_index + 1}")
        self._notify_data_changed()
        return True

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
            self._notify_data_changed()
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

    def _ensure_display_aliases(self):
        for row in self.table.rows:
            for header in self.DISPLAY_HEADERS:
                row.setdefault(header, "")
            for alias, source in self._DISPLAY_FALLBACKS.items():
                if alias in row and row[alias] not in (None, ""):
                    continue
                row[alias] = row.get(source, "")

    def _format_cell_value(self, header: str, value: object) -> str:
        if header in (*self._DATE_COLUMNS, *self._PUBLICATION_COLUMNS, *self._SALE_COLUMNS) and not _has_ready_date(value):
            return "☐"
        return "" if value is None else str(value)


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
