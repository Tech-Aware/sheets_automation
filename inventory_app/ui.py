"""Tkinter interface used by the inventory application."""
from __future__ import annotations

import threading
from dataclasses import dataclass
from pathlib import Path
from tkinter import BOTH, END, StringVar, Tk, ttk
from typing import Callable, Dict, Iterable, Sequence, Tuple

from inventory_app.data_loader import WorkbookRepository, load_named_sheets
from inventory_app.models import (
    PurchaseRecord,
    SaleRecord,
    StockRecord,
    build_purchases,
    build_sales,
    build_stock,
)
from inventory_app.services import (
    ReportingBundle,
    ReportRow,
    build_reporting,
    filter_records,
)


@dataclass(slots=True)
class TableDefinition:
    headers: Sequence[Tuple[str, str]]  # (column id, displayed label)
    extract: Callable[[object], Sequence[str]]
    search_attributes: Sequence[str]
    cache_key: str


class InventoryApplication:
    """Main application class orchestrating the Tkinter widgets."""

    def __init__(self, master: Tk, workbook_path: Path) -> None:
        self.master = master
        self.repository = WorkbookRepository(workbook_path)

        self.master.title("Gestion des stocks - Console avancée")
        self.master.geometry("1200x720")

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", font=("Segoe UI", 10), rowheight=28)
        style.configure("Treeview.Heading", font=("Segoe UI", 11, "bold"))
        style.configure("Dashboard.TLabel", font=("Segoe UI", 14, "bold"))
        style.configure("Badge.TLabel", font=("Segoe UI", 12), background="#f0f0f0")

        self._notebook = ttk.Notebook(master)
        self._notebook.pack(fill=BOTH, expand=True)

        self._search_vars: Dict[str, StringVar] = {}
        self._tables: Dict[str, ttk.Treeview] = {}
        self._data_cache: Dict[str, Sequence] = {}
        self._table_defs: Dict[str, TableDefinition] = {}
        self._reporting: ReportingBundle | None = None

        self._create_dashboard_tab()
        self._create_table_tab(
            "Achats",
            TableDefinition(
                headers=[
                    ("id", "ID"),
                    ("designation", "Désignation"),
                    ("quantite", "Quantité"),
                    ("prix", "Prix TTC"),
                    ("total", "Total TTC"),
                    ("statut", "Statut"),
                ],
                extract=lambda item: (
                    item.id_achat,
                    item.designation,
                    f"{item.quantite:g}",
                    f"{item.prix_unitaire_ttc:,.2f}",
                    f"{item.total_ttc:,.2f}",
                    item.statut,
                ),
                search_attributes=("id_achat", "designation", "statut"),
                cache_key="purchases",
            ),
        )
        self._create_table_tab(
            "Stock",
            TableDefinition(
                headers=[
                    ("sku", "SKU"),
                    ("designation", "Désignation"),
                    ("statut", "Statut"),
                    ("prix", "Prix"),
                    ("publication", "Publication"),
                    ("vente", "Vente"),
                    ("lot", "Lot"),
                ],
                extract=lambda item: (
                    item.sku,
                    item.designation,
                    item.disponibilite,
                    f"{item.prix_vente:,.2f}",
                    item.date_publication.strftime("%d/%m/%Y") if item.date_publication else "",
                    item.date_vente.strftime("%d/%m/%Y") if item.date_vente else "",
                    item.lot or "",
                ),
                search_attributes=("sku", "designation", "disponibilite", "lot"),
                cache_key="stock",
            ),
        )
        self._create_table_tab(
            "Ventes",
            TableDefinition(
                headers=[
                    ("sku", "SKU"),
                    ("designation", "Désignation"),
                    ("date", "Date"),
                    ("prix", "Prix"),
                    ("frais", "Frais"),
                    ("marge", "Marge"),
                    ("delai", "Délai"),
                ],
                extract=lambda item: (
                    item.sku,
                    item.designation,
                    item.date_vente.strftime("%d/%m/%Y") if item.date_vente else "",
                    f"{item.prix_vente:,.2f}",
                    f"{item.frais_port:,.2f}",
                    f"{item.marge:,.2f}",
                    str(item.delai or ""),
                ),
                search_attributes=("sku", "designation"),
                cache_key="sales",
            ),
        )
        self._create_reporting_tab()

        self._load_data_async()

    # ------------------------------------------------------------------
    # UI construction helpers
    # ------------------------------------------------------------------
    def _create_dashboard_tab(self) -> None:
        frame = ttk.Frame(self._notebook, padding=16)
        self._notebook.add(frame, text="Tableau de bord")

        self._dashboard_labels = {
            "stock": ttk.Label(frame, style="Dashboard.TLabel"),
            "purchase": ttk.Label(frame, style="Dashboard.TLabel"),
            "sales": ttk.Label(frame, style="Dashboard.TLabel"),
            "delay": ttk.Label(frame, style="Dashboard.TLabel"),
            "available": ttk.Label(frame, style="Dashboard.TLabel"),
            "pending": ttk.Label(frame, style="Dashboard.TLabel"),
        }

        for idx, (key, label) in enumerate(self._dashboard_labels.items()):
            label.grid(row=idx, column=0, sticky="w", pady=6)

        refresh_btn = ttk.Button(frame, text="Actualiser", command=self._load_data_async)
        refresh_btn.grid(row=0, column=1, padx=20)

    def _create_table_tab(self, name: str, table_definition: TableDefinition) -> None:
        frame = ttk.Frame(self._notebook, padding=8)
        self._notebook.add(frame, text=name)

        search_var = StringVar()
        self._search_vars[name] = search_var

        search_entry = ttk.Entry(frame, textvariable=search_var)
        search_entry.pack(fill="x", padx=4, pady=4)

        tree = ttk.Treeview(frame, columns=[col for col, _ in table_definition.headers], show="headings")
        for column_id, label in table_definition.headers:
            tree.heading(column_id, text=label)
            tree.column(column_id, width=150, anchor="center")
        tree.pack(fill=BOTH, expand=True, padx=4, pady=4)

        def on_search(*_args):
            records = self._data_cache.get(table_definition.cache_key, [])
            filtered = list(filter_records(records, search_var.get(), *table_definition.search_attributes))
            _populate_table(tree, table_definition, filtered)

        search_var.trace_add("write", on_search)
        self._tables[name] = tree
        self._table_defs[name] = table_definition

    def _create_reporting_tab(self) -> None:
        frame = ttk.Frame(self._notebook, padding=12)
        self._notebook.add(frame, text="Reporting")

        self._top_products_box = _ReportBox(frame, "Top articles", row=0)
        self._lots_box = _ReportBox(frame, "Lots", row=0, column=1)
        self._alerts_box = _AlertBox(frame, row=1, column=0, columnspan=2)

    # ------------------------------------------------------------------
    # Data loading
    # ------------------------------------------------------------------
    def _load_data_async(self) -> None:
        threading.Thread(target=self._load_data, daemon=True).start()

    def _load_data(self) -> None:
        sheets = load_named_sheets(
            self.repository,
            {
                "Achats": [
                    "ID ACHAT",
                    "DESIGNATION",
                    "QUANTITE",
                    "PRIX UNITAIRE TTC",
                    "TOTAL TTC",
                    "DATE D'ACHAT",
                    "FOURNISSEUR",
                    "STATUT",
                ],
                "Stock": [
                    "SKU",
                    "DESIGNATION",
                    "STATUT",
                    "PRIX",
                    "DATE PUBLICATION",
                    "DATE VENTE",
                    "LOT",
                ],
                "Ventes": [
                    "SKU",
                    "ARTICLE",
                    "DATE VENTE",
                    "PRIX",
                    "FRAIS",
                    "MARGE",
                    "DELAI",
                ],
            },
        )

        purchases = build_purchases(sheets["Achats"])
        stock = build_stock(sheets["Stock"])
        sales = build_sales(sheets["Ventes"])

        self.master.after(0, lambda: self._on_data_loaded(purchases, stock, sales))

    def _on_data_loaded(self, purchases: Sequence[PurchaseRecord], stock: Sequence[StockRecord], sales: Sequence[SaleRecord]) -> None:
        self._data_cache = {"purchases": purchases, "stock": stock, "sales": sales}
        self._reporting = build_reporting(purchases, stock, sales)

        self._refresh_dashboard()
        self._refresh_tables()
        self._refresh_reporting()

    # ------------------------------------------------------------------
    # Refresh helpers
    # ------------------------------------------------------------------
    def _refresh_dashboard(self) -> None:
        if not self._reporting:
            return
        dashboard = self._reporting.dashboard
        self._dashboard_labels["stock"].configure(text=f"Valeur stock : {dashboard.total_stock_value:,.2f} €")
        self._dashboard_labels["purchase"].configure(text=f"Achats cumulés : {dashboard.total_purchases_value:,.2f} €")
        self._dashboard_labels["sales"].configure(text=f"Ventes : {dashboard.total_sales_value:,.2f} €")
        self._dashboard_labels["delay"].configure(text=f"Délai moyen : {dashboard.average_delay:.1f} j")
        self._dashboard_labels["available"].configure(text=f"Articles disponibles : {dashboard.available_stock}")
        self._dashboard_labels["pending"].configure(text=f"Achats à traiter : {dashboard.pending_purchases}")

    def _refresh_tables(self) -> None:
        for name, tree in self._tables.items():
            table_def = self._table_defs[name]
            search_value = self._search_vars[name].get()
            data = self._data_cache.get(table_def.cache_key, [])
            filtered = list(filter_records(data, search_value, *table_def.search_attributes))
            _populate_table(tree, table_def, filtered)
    def _refresh_reporting(self) -> None:
        if not self._reporting:
            return
        reporting = self._reporting
        self._top_products_box.update_rows(reporting.top_products)
        self._lots_box.update_rows(reporting.lots)
        self._alerts_box.update_alerts(reporting.alerts)



class _ReportBox(ttk.LabelFrame):
    def __init__(self, master, title: str, row: int, column: int = 0, columnspan: int = 1) -> None:
        super().__init__(master, text=title, padding=12)
        self.grid(row=row, column=column, columnspan=columnspan, sticky="nsew", padx=6, pady=6)
        master.grid_columnconfigure(column, weight=1)
        master.grid_rowconfigure(row, weight=1)
        self._tree = ttk.Treeview(self, columns=("label", "value"), show="headings", height=8)
        self._tree.heading("label", text="Libellé")
        self._tree.heading("value", text="Valeur")
        self._tree.column("label", width=240, anchor="w")
        self._tree.column("value", width=120, anchor="e")
        self._tree.pack(fill=BOTH, expand=True)

    def update_rows(self, rows: Iterable[ReportRow]) -> None:
        self._tree.delete(*self._tree.get_children())
        for row in rows:
            self._tree.insert("", END, values=(row.label, row.value))


class _AlertBox(ttk.LabelFrame):
    def __init__(self, master, row: int, column: int = 0, columnspan: int = 1) -> None:
        super().__init__(master, text="Alertes", padding=12)
        self.grid(row=row, column=column, columnspan=columnspan, sticky="nsew", padx=6, pady=6)
        master.grid_columnconfigure(column, weight=1)
        self._listbox = ttk.Treeview(self, columns=("alert",), show="headings", height=6)
        self._listbox.heading("alert", text="Message")
        self._listbox.column("alert", anchor="w")
        self._listbox.pack(fill=BOTH, expand=True)

    def update_alerts(self, alerts: Sequence[str]) -> None:
        self._listbox.delete(*self._listbox.get_children())
        if not alerts:
            self._listbox.insert("", END, values=("Aucune alerte",))
            return
        for alert in alerts:
            self._listbox.insert("", END, values=(alert,))


def _populate_table(tree: ttk.Treeview, definition: TableDefinition, records: Sequence) -> None:
    tree.delete(*tree.get_children())
    for record in records:
        tree.insert("", END, values=definition.extract(record))


def run_app(workbook_path: Path | str) -> None:
    root = Tk()
    app = InventoryApplication(root, Path(workbook_path))
    root.mainloop()
