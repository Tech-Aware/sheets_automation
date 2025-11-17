from __future__ import annotations

from concurrent.futures import Future, ThreadPoolExecutor
from pathlib import Path
from typing import Mapping
from tkinter import messagebox

import customtkinter as ctk

from ..datasources.sqlite_purchases import PurchaseDatabase, table_to_purchase_records, table_to_stock_records
from ..datasources.workbook import TableData
from ..services.workflow import WorkflowCoordinator
from ..utils.perf import format_report, performance_monitor
from .loading import LoadingDialog
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
        self._refresh_executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="ui-worker")
        self._pending_refresh: Future | None = None
        self._loading_dialog: LoadingDialog | None = None
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

    def refresh_views(self, *, prepare_only: bool = False, cancel_only: bool = False):
        """Schedule a UI refresh while coalescing rapid successive calls.

        ``prepare_only`` allows callers to display the loading dialog before
        entering a potentially heavy workflow, ensuring the user sees feedback
        as soon as the process starts. ``cancel_only`` is a safeguard to close
        the dialog when a prepare step fails before the actual refresh begins.
        """

        if cancel_only:
            self._close_loading_dialog()
            return
        if prepare_only:
            self._show_loading_dialog()
            self._update_loading_progress(0.02)
            return
        if self._pending_refresh is not None and not self._pending_refresh.done():
            return
        self._show_loading_dialog()
        self._update_loading_progress(0.05)
        self._pending_refresh = self._refresh_executor.submit(self._compute_refresh_payload)
        self._pending_refresh.add_done_callback(lambda fut: self.after(0, self._apply_refresh, fut))

    def _compute_refresh_payload(self):
        with performance_monitor.track("ui.refresh.rebuild_indexes"):
            self.workflow.rebuild_indexes()
        with performance_monitor.track("ui.refresh.inventory_snapshot"):
            summary = self.workflow.inventory_snapshot()
        return summary

    def _apply_refresh(self, future: Future):
        self._pending_refresh = None
        try:
            summary = future.result()
        except Exception as exc:  # pragma: no cover - UI safeguard
            messagebox.showerror("Rafraîchissement UI", f"Échec du rafraîchissement : {exc}")
            self._close_loading_dialog()
            return
        self._update_loading_progress(0.45)
        table_row_counts = {
            name: len(view.table.rows) if hasattr(view, "table") and hasattr(view.table, "rows") else None
            for name, view in self.table_views.items()
        }
        total_steps = sum(
            step is not None
            for step in (
                self.dashboard_view,
                self.purchase_view,
                *[view for view in self.table_views.values() if view is not self.purchase_view],
            )
        )
        progress_cursor = 0
        progress_base = 0.45
        progress_step = 0.55 / max(total_steps, 1)
        with performance_monitor.track(
            "ui.refresh.widgets",
            metadata={
                "tables": len(self.table_views),
                "rows": sum(count or 0 for count in table_row_counts.values()),
            },
        ):
            if self.dashboard_view is not None:
                with performance_monitor.track("ui.refresh.dashboard"):
                    self.dashboard_view.refresh(summary)
                progress_cursor += 1
                self._update_loading_progress(progress_base + progress_step * progress_cursor)
            if self.purchase_view is not None:
                with performance_monitor.track(
                    "ui.refresh.purchases",
                    metadata={"rows": table_row_counts.get("Achats")},
                ):
                    self.purchase_view.refresh()
                progress_cursor += 1
                self._update_loading_progress(progress_base + progress_step * progress_cursor)
            for name, view in self.table_views.items():
                if view is self.purchase_view:
                    continue
                with performance_monitor.track(
                    "ui.refresh.table_view",
                    metadata={"sheet": name, "rows": table_row_counts.get(name)},
                ):
                    view.refresh()
                progress_cursor += 1
                self._update_loading_progress(progress_base + progress_step * progress_cursor)
        self._update_loading_progress(1.0)
        self.after(300, self._close_loading_dialog)

    def _handle_close(self):
        try:
            if self._pending_refresh is not None:
                self._pending_refresh.cancel()
            self._refresh_executor.shutdown(wait=False)
            self._persist_tables()
            slowest = performance_monitor.slowest()
            if slowest:
                print(format_report(slowest))
        except Exception as exc:  # pragma: no cover - UI safeguard
            messagebox.showerror("Sauvegarde des données", f"Échec de l'enregistrement des données : {exc}")
            return
        self._close_loading_dialog()
        self.destroy()

    def _persist_tables(self):
        if self.achats_db_path is None:
            return
        records = table_to_purchase_records(self.tables.get("Achats"))
        stock_records = table_to_stock_records(self.tables.get("Stock"))
        PurchaseDatabase(self.achats_db_path).replace_all(records, stock_records)

    def _show_loading_dialog(self):
        if self._loading_dialog is not None:
            try:
                self._loading_dialog.focus()
                return
            except Exception:
                self._loading_dialog = None
        dialog = LoadingDialog(self, message="Patientez pendant le chargement de vos données")
        dialog.grab_set()
        self._loading_dialog = dialog

    def _update_loading_progress(self, value: float):
        if self._loading_dialog is not None:
            try:
                self._loading_dialog.update_progress(value)
            except Exception:
                self._loading_dialog = None

    def _close_loading_dialog(self):
        if self._loading_dialog is not None:
            try:
                self._loading_dialog.close()
            finally:
                self._loading_dialog = None
