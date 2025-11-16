from __future__ import annotations

import tkinter as tk
from typing import Sequence

import customtkinter as ctk

from .tables import ScrollableTable


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
