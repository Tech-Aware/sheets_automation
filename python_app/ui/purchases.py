from __future__ import annotations

from datetime import date
import tkinter as tk
from typing import Mapping, Sequence

import customtkinter as ctk

from ..config import HEADERS
from ..services.workflow import PurchaseInput, WorkflowCoordinator
from ..ui.tables import ScrollableTable
from ..ui.widgets import DatePickerEntry
from ..utils.datefmt import format_display_date, parse_date_value
from ..utils.perf import performance_monitor


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
        with performance_monitor.track("ui.purchases.build_summary", metadata={"rows": len(self.table.rows)}):
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
        if self.table_widget is not None:
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
            with performance_monitor.track("ui.purchases.refresh_table"):
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
        removed = 0
        for idx in sorted(indices, reverse=True):
            if 0 <= idx < len(self.table.rows):
                del self.table.rows[idx]
                removed += 1
        if removed:
            self.status_var.set(f"{removed} commande(s) supprimée(s)")
            self.table_widget.refresh(self._build_summary_rows())
            self.refresh_callback()
        else:
            self.status_var.set("Aucune commande supprimée")


class AddPurchaseDialog(ctk.CTkToplevel):
    def __init__(self, master, on_submit, sku_builder, on_close):
        super().__init__(master)
        self.on_submit = on_submit
        self.sku_builder = sku_builder
        self.on_close = on_close
        self.title("Ajouter une commande")
        self.geometry("640x760")
        self.resizable(False, False)
        owner = master.winfo_toplevel() if hasattr(master, "winfo_toplevel") else None
        if owner is not None:
            self.transient(owner)
        self.grab_set()
        self.lift()
        self.focus()
        self._fields: dict[str, ctk.CTkEntry] = {}
        self._status = tk.StringVar(value="")
        self._build_form()
        self.protocol("WM_DELETE_WINDOW", self._close)

    def _build_form(self):
        frame = ctk.CTkScrollableFrame(self)
        frame.pack(fill="both", expand=True, padx=16, pady=16)
        ctk.CTkLabel(frame, text="Nouvelle commande", font=ctk.CTkFont(size=18, weight="bold"), anchor="w").pack(
            fill="x", padx=12, pady=(8, 12)
        )
        for label in (
            "Article",
            "Marque",
            "Genre",
            "Date achat",
            "Grade",
            "Fournisseur",
            "Quantité commandée",
            "Quantité reçue",
            "Prix d'achat TTC",
            "Frais de lavage",
            "Date livraison",
            "Tracking",
        ):
            self._field(frame, label)
        ctk.CTkButton(frame, text="Proposer un SKU", command=self._suggest_sku).pack(fill="x", padx=12, pady=(4, 2))
        for label in (
            "SKU",
            "Quantité utilisable",
            "Prix unitaire TTC",
        ):
            self._field(frame, label)
        ctk.CTkLabel(frame, textvariable=self._status, anchor="w").pack(fill="x", padx=12, pady=(4, 4))
        ctk.CTkButton(frame, text="Enregistrer", command=self._submit).pack(fill="x", padx=12, pady=(8, 4))

    def _field(self, parent, label: str):
        row = ctk.CTkFrame(parent)
        row.pack(fill="x", padx=12, pady=2)
        ctk.CTkLabel(row, text=label, width=220, anchor="w").pack(side="left")
        entry = ctk.CTkEntry(row)
        entry.pack(side="left", fill="x", expand=True, padx=(8, 0))
        self._fields[label] = entry

    def _suggest_sku(self):
        article = self._fields["Article"].get().strip()
        marque = self._fields["Marque"].get().strip()
        genre = self._fields["Genre"].get().strip()
        grade = self._fields["Grade"].get().strip()
        label, sku = self.sku_builder(article=article, marque=marque, genre=genre, grade=grade)
        self._fields["SKU"].delete(0, tk.END)
        self._fields["SKU"].insert(0, sku)
        self._status.set(label)

    def _submit(self):
        payload = {key: widget.get().strip() for key, widget in self._fields.items()}
        payload.update(
            {
                "quantite": payload.get("Quantité utilisable"),
                "prix_unitaire": payload.get("Prix unitaire TTC"),
                "sku": payload.get("SKU"),
                "date_achat": payload.get("Date achat"),
                "date_livraison": payload.get("Date livraison"),
                "quantite_commandee": payload.get("Quantité commandée"),
                "quantite_recue": payload.get("Quantité reçue"),
                "prix_achat": payload.get("Prix d'achat TTC"),
                "frais_lavage": payload.get("Frais de lavage"),
            }
        )
        success, message = self.on_submit(payload)
        self._status.set(message)
        if success:
            self._close()

    def _close(self):
        try:
            self.grab_release()
        except tk.TclError:
            pass
        if callable(self.on_close):
            self.on_close()
        self.destroy()


class PurchaseDetailDialog(ctk.CTkToplevel):
    def __init__(self, master, row_index: int, headers: Sequence[str], row: Mapping, on_submit, on_close):
        super().__init__(master)
        self.row_index = row_index
        self.headers = headers
        self.row_data = row
        self.on_submit = on_submit
        self.on_close = on_close
        self.title("Détails de la commande")
        self.geometry("720x820")
        self.resizable(True, True)
        owner = master.winfo_toplevel() if hasattr(master, "winfo_toplevel") else None
        if owner is not None:
            self.transient(owner)
        self.grab_set()
        self.lift()
        self.focus()
        self.entries: dict[str, ctk.CTkEntry | DatePickerEntry] = {}
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
