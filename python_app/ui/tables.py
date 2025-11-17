"""Reusable Tk table component used across tabs."""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from typing import Callable, Iterable, Mapping, Sequence

from ..utils.perf import performance_monitor


class ScrollableTable(ttk.Frame):
    """Small helper that wraps a Treeview with scrollbars and inline editing."""

    def __init__(
        self,
        master,
        headers: Sequence[str],
        rows: Iterable[dict],
        *,
        height: int = 15,
        column_width: int = 140,
        column_widths: Mapping[str, int] | None = None,
        on_cell_edited: Callable[[int, str, str], None] | None = None,
        on_cell_activated: Callable[[int | None, str], bool] | None = None,
        on_row_activated: Callable[[int], None] | None = None,
        enable_inline_edit: bool = True,
        value_formatter: Callable[[str, object], str] | None = None,
        dropdown_choices: Mapping[str, Sequence[str]] | None = None,
    ):
        super().__init__(master)
        self.on_cell_edited = on_cell_edited
        self.on_cell_activated = on_cell_activated
        self.on_row_activated = on_row_activated
        self.enable_inline_edit = enable_inline_edit
        self.value_formatter = value_formatter
        self._dropdown_choices = {key: tuple(values) for key, values in (dropdown_choices or {}).items()}
        self._vsb: ttk.Scrollbar | None = None
        self._hsb: ttk.Scrollbar | None = None
        self._editor: tk.Entry | ttk.Combobox | None = None
        self._editing_item: str | None = None
        self._editing_column: str | None = None
        self._item_to_row_index: dict[str, int] = {}
        self._headers = list(headers)
        self._column_widths = dict(column_widths or {})
        style = ttk.Style(self)
        try:  # ``clam`` provides better contrast for striped rows
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(
            "Scrollable.Treeview",
            font=("Segoe UI", 11),
            rowheight=28,
            background="#fcfcfc",
            fieldbackground="#fcfcfc",
        )
        style.configure("Scrollable.Treeview.Heading", font=("Segoe UI", 11, "bold"))
        style.map(
            "Scrollable.Treeview",
            background=[("selected", "#0F62FE")],
            foreground=[("selected", "white")],
        )
        self.tree = ttk.Treeview(
            self,
            columns=self._headers,
            show="headings",
            height=height,
            style="Scrollable.Treeview",
            selectmode="extended",
        )
        for header in self._headers:
            self.tree.heading(header, text=header)
            width = self._column_widths.get(header, column_width)
            self.tree.column(header, width=width, anchor=tk.W)
        self.tree.tag_configure("even", background="#f2f4f8")
        self._vsb = ttk.Scrollbar(self, orient="vertical", command=self._on_vscroll)
        self._hsb = ttk.Scrollbar(self, orient="horizontal", command=self._on_hscroll)
        self.tree.configure(
            yscrollcommand=self._on_tree_yscroll,
            xscrollcommand=self._on_tree_xscroll,
        )
        self.tree.grid(row=0, column=0, sticky="nsew")
        self._vsb.grid(row=0, column=1, sticky="ns")
        self._hsb.grid(row=1, column=0, sticky="ew")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self._insert_rows(rows)
        self.tree.bind("<Double-1>", self._handle_double_click)
        self.tree.bind("<Configure>", self._reposition_editor)

    def refresh(self, rows: Iterable[dict]):
        rows_list = list(rows)
        with performance_monitor.track("ui.table.refresh", metadata={"rows": len(rows_list)}):
            for child in self.tree.get_children():
                self.tree.delete(child)
            self._item_to_row_index.clear()
            self._insert_rows(rows_list)

    def get_selected_indices(self) -> list[int]:
        indices: list[int] = []
        for item in self.tree.selection():
            row_index = self._item_to_row_index.get(item)
            if row_index is not None:
                indices.append(row_index)
        return indices

    # ------------------------------------------------------------------
    # Inline editing helpers
    # ------------------------------------------------------------------
    def _insert_rows(self, rows: Iterable[dict]):
        with performance_monitor.track("ui.table.insert_rows"):
            for idx, row in enumerate(rows):
                values: list[str] = []
                for header in self._headers:
                    raw_value = row.get(header, "")
                    base_value = self.value_formatter(header, raw_value) if self.value_formatter is not None else raw_value
                    values.append(self._render_display_value(header, base_value))
                item = self.tree.insert(
                    "",
                    tk.END,
                    values=values,
                    tags=("even" if idx % 2 == 0 else "odd",),
                )
                self._item_to_row_index[item] = idx

    def _handle_double_click(self, event):
        item = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        if not item or column == "#0":
            return

        column_index = int(column.replace("#", "")) - 1
        if column_index < 0 or column_index >= len(self._headers):
            return
        column_id = self._headers[column_index]
        row_index = self._item_to_row_index.get(item)

        if self.on_cell_activated is not None and self.on_cell_activated(row_index, column_id):
            return

        if self.enable_inline_edit:
            self._begin_edit(item, column_id)
            return
        if self.on_row_activated is None:
            return
        if row_index is None:
            return
        self.on_row_activated(row_index)

    def _begin_edit(self, item: str, column_id: str):
        if self._editor is not None:
            self._finalize_edit()
        bbox = self.tree.bbox(item, column_id)
        if not bbox:
            return
        x, y, width, height = bbox
        value = self._clean_dropdown_value(column_id, self.tree.set(item, column_id))
        self._editing_item = item
        self._editing_column = column_id
        choices = self._dropdown_choices.get(column_id)
        if choices:
            editor = ttk.Combobox(self.tree, values=choices, state="readonly")
            if value:
                editor.set(value)
            editor.bind("<<ComboboxSelected>>", self._finalize_edit)
        else:
            editor = tk.Entry(self.tree)
            editor.insert(0, value)
            editor.select_range(0, tk.END)
            editor.bind("<Return>", self._finalize_edit)
            editor.bind("<Escape>", self._cancel_edit)
        editor.focus()
        editor.place(x=x, y=y, width=width, height=height)
        editor.bind("<FocusOut>", self._finalize_edit)
        self._editor = editor

    def _finalize_edit(self, _event=None):
        if not self._editor or self._editing_item is None or self._editing_column is None:
            return
        new_value = self._clean_dropdown_value(self._editing_column, self._editor.get())
        self.tree.set(
            self._editing_item,
            self._editing_column,
            self._render_display_value(self._editing_column, new_value),
        )
        row_index = self._item_to_row_index.get(self._editing_item)
        if row_index is not None and self.on_cell_edited is not None:
            self.on_cell_edited(row_index, self._editing_column, new_value)
        self._cleanup_editor()

    def _cancel_edit(self, _event=None):
        self._cleanup_editor()

    def _cleanup_editor(self):
        if self._editor is not None:
            self._editor.destroy()
        self._editor = None
        self._editing_item = None
        self._editing_column = None

    def _render_display_value(self, column_id: str, raw_value) -> str:
        display_value = "" if raw_value is None else str(raw_value)
        if column_id in self._dropdown_choices:
            indicator = "▼"
            return f"{display_value} {indicator}" if display_value else indicator
        return display_value

    def _clean_dropdown_value(self, column_id: str, value: str) -> str:
        if column_id not in self._dropdown_choices:
            return value
        return value.rstrip().removesuffix("▼").rstrip()

    def _reposition_editor(self, _event=None):
        if self._editor is None or self._editing_item is None or self._editing_column is None:
            return
        bbox = self.tree.bbox(self._editing_item, self._editing_column)
        if not bbox:
            self._cleanup_editor()
            return
        x, y, width, height = bbox
        self._editor.place(x=x, y=y, width=width, height=height)

    def _on_vscroll(self, *args):
        self.tree.yview(*args)
        self._reposition_editor()

    def _on_hscroll(self, *args):
        self.tree.xview(*args)
        self._reposition_editor()

    def _on_tree_yscroll(self, *args):
        if self._vsb is not None:
            self._vsb.set(*args)
        self._reposition_editor()

    def _on_tree_xscroll(self, *args):
        if self._hsb is not None:
            self._hsb.set(*args)
        self._reposition_editor()


__all__ = ["ScrollableTable"]
