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
    from .ui.tables import ScrollableTable
except ImportError:  # pragma: no cover - executed when run as a script
    package_root = Path(__file__).resolve().parent.parent
    if str(package_root) not in sys.path:
        sys.path.append(str(package_root))
    from python_app.config import HEADERS, MONTH_NAMES_FR
    from python_app.datasources.workbook import WorkbookRepository
    from python_app.services.summaries import build_inventory_snapshot
    from python_app.ui.tables import ScrollableTable


class VintageErpApp(ctk.CTk):
    """Simple multipage CustomTkinter application."""

    def __init__(self, repository: WorkbookRepository):
        super().__init__()
        self.title("Vintage ERP (Prerelease 1.2)")
        self.geometry("1200x800")
        self.repository = repository
        self.tabview = ctk.CTkTabview(self)
        self.tabview.pack(fill="both", expand=True, padx=16, pady=16)
        self._build_tabs()

    def _build_tabs(self):
        tables = self.repository.load_many("Achats", "Stock", "Ventes", "Compta 09-2025")
        summary = build_inventory_snapshot(tables["Stock"].rows, tables["Ventes"].rows)

        dashboard_tab = self.tabview.add("Dashboard")
        DashboardView(dashboard_tab, summary)

        for sheet, label in (
            ("Achats", "Achats"),
            ("Stock", "Stock"),
            ("Ventes", "Ventes"),
            ("Compta 09-2025", "Compta"),
        ):
            tab = self.tabview.add(label)
            TableView(tab, tables[sheet])

        months_tab = self.tabview.add("Calendrier")
        CalendarView(months_tab)


class DashboardView(ctk.CTkFrame):
    def __init__(self, master, snapshot):
        super().__init__(master)
        self.pack(fill="both", expand=True)
        title = ctk.CTkLabel(self, text="Vue d'ensemble", font=ctk.CTkFont(size=24, weight="bold"))
        title.pack(pady=(16, 8))
        grid = ctk.CTkFrame(self)
        grid.pack(padx=16, pady=16, fill="x")
        stats = snapshot.as_dict()
        for idx, (label, value) in enumerate(stats.items()):
            card = ctk.CTkFrame(grid)
            card.grid(row=0, column=idx, padx=8, pady=8, sticky="nsew")
            grid.grid_columnconfigure(idx, weight=1)
            ctk.CTkLabel(card, text=label.replace("_", " ").title()).pack(padx=12, pady=(12, 4))
            ctk.CTkLabel(card, text=str(value)).pack(padx=12, pady=(0, 12))


class TableView(ctk.CTkFrame):
    def __init__(self, master, table):
        super().__init__(master)
        self.pack(fill="both", expand=True)
        self.table = table
        helper = ctk.CTkFrame(self)
        helper.pack(fill="x", padx=12, pady=(12, 0))
        self.status_var = tk.StringVar(value="Double-cliquez sur une cellule pour la modifier.")
        ctk.CTkLabel(helper, textvariable=self.status_var, anchor="w").pack(side="left")
        ScrollableTable(
            self,
            table.headers[:10],
            table.rows,
            height=20,
            on_cell_edited=self._on_cell_edit,
            column_width=160,
        ).pack(fill="both", expand=True, padx=12, pady=12)

    def _on_cell_edit(self, row_index: int, column: str, new_value: str):
        try:
            self.table.rows[row_index][column] = new_value
        except (IndexError, KeyError):
            pass
        self.status_var.set(f"Ligne {row_index + 1} – {column} mis à jour")


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
