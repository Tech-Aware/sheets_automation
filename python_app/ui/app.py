from __future__ import annotations

from pathlib import Path
from typing import Mapping
from tkinter import messagebox

import customtkinter as ctk

from ..datasources.sqlite_purchases import PurchaseDatabase, table_to_purchase_records, table_to_stock_records
from ..datasources.workbook import TableData
from ..services.workflow import WorkflowCoordinator
from .calendar import CalendarView
from .dashboard import DashboardView
from .purchases import PurchasesView
from .stock import StockOptionsView, StockTableView
from .table_view import TableView


class VintageErpApp(ctk.CTk):
    """Simple multipage CustomTkinter application."""

    def __init__(
        self,
        tables: Mapping[str, TableData],
        workflow: WorkflowCoordinator,
        *,
        achats_db_path: Path | None = None,
    ):
        super().__init__()
        self.title("Vintage ERP (Prerelease 1.2)")
        self.geometry("1200x800")
        self.minsize(1024, 720)
        self.tables = tables
        self.workflow = workflow
        self.achats_db_path = Path(achats_db_path) if achats_db_path is not None else None
        self.tabview = ctk.CTkTabview(self)
        self.tabview.pack(fill="both", expand=True, padx=16, pady=16)
        self.table_views: dict[str, TableView] = {}
        self.purchase_view: PurchasesView | None = None
        self.dashboard_view: DashboardView | None = None
        self._refresh_job: str | None = None
        self._build_tabs()
        self.protocol("WM_DELETE_WINDOW", self._handle_close)

    def _build_tabs(self):
        summary = self.workflow.inventory_snapshot()

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
        """Schedule a UI refresh while coalescing rapid successive calls."""

        if self._refresh_job is not None:
            return
        self._refresh_job = self.after_idle(self._perform_refresh)

    def _perform_refresh(self):
        self._refresh_job = None
        self.workflow.rebuild_indexes()
        summary = self.workflow.inventory_snapshot()
        if self.dashboard_view is not None:
            self.dashboard_view.refresh(summary)
        if self.purchase_view is not None:
            self.purchase_view.refresh()
        for view in self.table_views.values():
            view.refresh()

    def _handle_close(self):
        try:
            if self._refresh_job is not None:
                self.after_cancel(self._refresh_job)
                self._refresh_job = None
            self._persist_tables()
        except Exception as exc:  # pragma: no cover - UI safeguard
            messagebox.showerror("Sauvegarde des données", f"Échec de l'enregistrement des données : {exc}")
            return
        self.destroy()

    def _persist_tables(self):
        if self.achats_db_path is None:
            return
        records = table_to_purchase_records(self.tables.get("Achats"))
        stock_records = table_to_stock_records(self.tables.get("Stock"))
        PurchaseDatabase(self.achats_db_path).replace_all(records, stock_records)
