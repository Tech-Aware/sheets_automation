"""Entry point for the CustomTkinter application."""
from __future__ import annotations

import argparse
import sys
import tkinter as tk
from pathlib import Path
from tkinter import messagebox

import customtkinter as ctk

# ``python python_app/main.py`` was failing because ``__package__`` is empty when a
# module is executed as a script.  To keep relative imports working both for
# ``python -m python_app.main`` and direct execution, fall back to absolute
# imports after adding the repository root to ``sys.path``.
try:  # pragma: no cover - defensive import path configuration
    from .config import HEADERS, MONTH_NAMES_FR
    from .datasources.workbook import WorkbookRepository
    from .services.summaries import build_inventory_snapshot
    from .services.workflow import PurchaseInput, SaleInput, StockInput, WorkflowCoordinator
    from .ui.tables import ScrollableTable
except ImportError:  # pragma: no cover - executed when run as a script
    package_root = Path(__file__).resolve().parent.parent
    if str(package_root) not in sys.path:
        sys.path.append(str(package_root))
    from python_app.config import HEADERS, MONTH_NAMES_FR
    from python_app.datasources.workbook import WorkbookRepository
    from python_app.services.summaries import build_inventory_snapshot
    from python_app.services.workflow import PurchaseInput, SaleInput, StockInput, WorkflowCoordinator
    from python_app.ui.tables import ScrollableTable


class VintageErpApp(ctk.CTk):
    """Simple multipage CustomTkinter application."""

    def __init__(self, repository: WorkbookRepository):
        super().__init__()
        self.title("Vintage ERP (Prerelease 1.2)")
        self.geometry("1200x800")
        self.repository = repository
        self.tables = self.repository.load_many("Achats", "Stock", "Ventes", "Compta 09-2025")
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
            else:
                view = TableView(tab, self.tables[sheet])
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
    def __init__(self, master, table, workflow: WorkflowCoordinator, refresh_callback):
        super().__init__(master)
        self.pack(fill="both", expand=True)
        self.table = table
        self.workflow = workflow
        self.refresh_callback = refresh_callback
        self.entries: dict[str, ctk.CTkEntry] = {}
        self.log_var = tk.StringVar(value="Saisissez une commande pour commencer.")
        self.sku_preview_var = tk.StringVar(value="Préfixe SKU : --")
        self._build_form()
        self._build_ready_section()
        self._build_table()

    def _build_form(self):
        frame = ctk.CTkFrame(self)
        frame.pack(fill="x", padx=16, pady=(16, 8))
        ctk.CTkLabel(frame, text="Nouvelle commande", font=ctk.CTkFont(size=18, weight="bold"), anchor="w").pack(
            fill="x", padx=12, pady=(12, 4)
        )
        fields = [
            ("Article", "article", ""),
            ("Marque", "marque", ""),
            ("Genre", "genre", ""),
            ("Date d'achat (AAAA-MM-JJ)", "date_achat", ""),
            ("Grade", "grade", ""),
            ("Fournisseur / Code", "fournisseur", ""),
            ("Quantité commandée", "quantite_commandee", "1"),
            ("Quantité reçue", "quantite_recue", "1"),
            ("Prix d'achat TTC", "prix_achat", "0"),
            ("Frais de lavage", "frais_lavage", "0"),
            ("Date de livraison (AAAA-MM-JJ)", "date_livraison", ""),
            ("Tracking", "tracking", ""),
        ]
        article_entry = None
        marque_entry = None
        for label, key, default in fields:
            entry = self._field(frame, label, key, default=default)
            if key == "article":
                article_entry = entry
            elif key == "marque":
                marque_entry = entry
        if article_entry is not None:
            article_entry.bind("<KeyRelease>", self._update_sku_preview)
        if marque_entry is not None:
            marque_entry.bind("<KeyRelease>", self._update_sku_preview)
        ctk.CTkLabel(frame, textvariable=self.sku_preview_var, anchor="w").pack(fill="x", padx=12, pady=(4, 0))
        ctk.CTkButton(frame, text="Ajouter la commande", command=self._handle_add).pack(padx=12, pady=(8, 4))
        ctk.CTkLabel(frame, textvariable=self.log_var, anchor="w").pack(fill="x", padx=12, pady=(0, 12))

    def _build_ready_section(self):
        frame = ctk.CTkFrame(self)
        frame.pack(fill="x", padx=16, pady=(0, 8))
        ctk.CTkLabel(frame, text="Valider la mise en stock", font=ctk.CTkFont(size=18, weight="bold"), anchor="w").pack(
            fill="x", padx=12, pady=(12, 4)
        )
        self._field(frame, "ID achat", "ready_id")
        self._field(frame, "Date de mise en stock (AAAA-MM-JJ)", "ready_date")
        ctk.CTkButton(frame, text="Créer les SKU", command=self._handle_ready).pack(padx=12, pady=(8, 12))

    def _build_table(self):
        frame = ctk.CTkFrame(self)
        frame.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        helper = ctk.CTkFrame(frame)
        helper.pack(fill="x", padx=12, pady=(12, 0))
        self.status_var = tk.StringVar(value="Double-cliquez sur une cellule pour la modifier.")
        ctk.CTkLabel(helper, textvariable=self.status_var, anchor="w").pack(fill="x")
        self.table_widget = ScrollableTable(
            frame,
            self.table.headers,
            self.table.rows,
            height=18,
            column_width=150,
            on_cell_edited=self._on_cell_edit,
        )
        self.table_widget.pack(fill="both", expand=True, padx=12, pady=12)

    def _field(self, parent, label: str, key: str, *, default: str = ""):
        row = ctk.CTkFrame(parent)
        row.pack(fill="x", padx=12, pady=2)
        ctk.CTkLabel(row, text=label, width=220, anchor="w").pack(side="left")
        entry = ctk.CTkEntry(row)
        if default:
            entry.insert(0, default)
        entry.pack(side="left", fill="x", expand=True, padx=(8, 0))
        self.entries[key] = entry
        return entry

    def _handle_add(self):
        article = self._get_entry_value("article")
        marque = self._get_entry_value("marque")
        if not article or not marque:
            self._log("Article et marque sont obligatoires")
            return
        genre = self._get_entry_value("genre")
        date_achat = self._get_entry_value("date_achat")
        grade = self._get_entry_value("grade")
        fournisseur = self._get_entry_value("fournisseur")
        date_livraison = self._get_entry_value("date_livraison")
        tracking = self._get_entry_value("tracking")
        try:
            qty_cmd = self._parse_int_field("quantite_commandee", "Quantité commandée")
            qty_recue = self._parse_int_field("quantite_recue", "Quantité reçue") or qty_cmd
            prix_total = self._parse_float_field("prix_achat", "Prix d'achat TTC")
            frais_lavage = self._parse_float_field("frais_lavage", "Frais de lavage")
        except ValueError as exc:
            self._log(str(exc))
            return
        if qty_recue <= 0:
            self._log("La quantité reçue doit être supérieure à 0")
            return
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
        self._log(f"Commande {row.get(HEADERS['ACHATS'].ID)} ajoutée")
        self._clear_form()
        self.refresh_callback()

    def _handle_ready(self):
        purchase_id = self._get_entry_value("ready_id")
        if not purchase_id:
            self._log("Veuillez saisir l'ID de la commande à mettre en stock")
            return
        ready_date = self._get_entry_value("ready_date") or None
        try:
            created = self.workflow.prepare_stock_from_purchase(purchase_id, ready_date)
        except ValueError as exc:
            self._log(str(exc))
            return
        self._log(f"{len(created)} SKU créés pour l'achat {purchase_id}")
        self.refresh_callback()

    def _parse_int_field(self, key: str, label: str) -> int:
        value = self._get_entry_value(key)
        if not value:
            return 0
        try:
            return int(value)
        except ValueError as exc:
            raise ValueError(f"{label} doit être un nombre entier") from exc

    def _parse_float_field(self, key: str, label: str) -> float:
        value = self._get_entry_value(key)
        if not value:
            return 0.0
        try:
            return float(value)
        except ValueError as exc:
            raise ValueError(f"{label} doit être un nombre") from exc

    def _get_entry_value(self, key: str) -> str:
        entry = self.entries.get(key)
        return entry.get().strip() if entry else ""

    def _update_sku_preview(self, _event=None):
        article = self._get_entry_value("article")
        marque = self._get_entry_value("marque")
        if not article and not marque:
            self.sku_preview_var.set("Préfixe SKU : --")
            return
        base = self.workflow.build_sku_base(article, marque)
        self.sku_preview_var.set(f"Préfixe SKU : {base}")

    def _clear_form(self):
        for key in ("article", "marque", "genre", "date_achat", "grade", "fournisseur", "date_livraison", "tracking"):
            entry = self.entries.get(key)
            if entry:
                entry.delete(0, tk.END)
        for key, default in (("quantite_commandee", "1"), ("quantite_recue", "1"), ("prix_achat", "0"), ("frais_lavage", "0")):
            entry = self.entries.get(key)
            if entry:
                entry.delete(0, tk.END)
                entry.insert(0, default)
        self._update_sku_preview()

    def _on_cell_edit(self, row_index: int, column: str, new_value: str):
        try:
            self.table.rows[row_index][column] = new_value
        except (IndexError, KeyError):
            return
        self.status_var.set(f"Ligne {row_index + 1} – {column} mis à jour")

    def refresh(self):
        self.table_widget.refresh(self.table.rows)

    def _log(self, message: str):
        self.log_var.set(message)


class TableView(ctk.CTkFrame):
    def __init__(self, master, table):
        super().__init__(master)
        self.pack(fill="both", expand=True)
        self.table = table
        helper = ctk.CTkFrame(self)
        helper.pack(fill="x", padx=12, pady=(12, 0))
        self.status_var = tk.StringVar(value="Double-cliquez sur une cellule pour la modifier.")
        ctk.CTkLabel(helper, textvariable=self.status_var, anchor="w").pack(side="left")
        self.table_widget = ScrollableTable(
            self,
            table.headers[:10],
            table.rows,
            height=20,
            on_cell_edited=self._on_cell_edit,
            column_width=160,
        )
        self.table_widget.pack(fill="both", expand=True, padx=12, pady=12)

    def _on_cell_edit(self, row_index: int, column: str, new_value: str):
        try:
            self.table.rows[row_index][column] = new_value
        except (IndexError, KeyError):
            pass
        self.status_var.set(f"Ligne {row_index + 1} – {column} mis à jour")

    def refresh(self):
        self.table_widget.refresh(self.table.rows)


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
        self.sale_date = self._field(frame, "Date de vente (AAAA-MM-JJ)")
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

    def _field(self, parent, label: str, *, default: str | None = None):
        row = ctk.CTkFrame(parent)
        row.pack(fill="x", padx=12, pady=4)
        ctk.CTkLabel(row, text=label, width=180, anchor="w").pack(side="left")
        entry = ctk.CTkEntry(row)
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


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Vintage ERP UI")
    parser.add_argument(
        "workbook",
        nargs="?",
        default=DEFAULT_WORKBOOK,
        type=Path,
        help="Path to the Excel workbook (defaults to Prerelease 1.2.xlsx located at the repo root)",
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
    app = VintageErpApp(repo)
    app.mainloop()
    return 0


if __name__ == "__main__":  # pragma: no cover - CLI entry point
    raise SystemExit(main())
