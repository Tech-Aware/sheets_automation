"""Reusable Tk table component used across tabs."""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from typing import Iterable, Sequence


class ScrollableTable(ttk.Frame):
    """Very small helper that wraps a Treeview with scrollbars."""

    def __init__(self, master, headers: Sequence[str], rows: Iterable[dict], *, height: int = 15):
        super().__init__(master)
        self.tree = ttk.Treeview(self, columns=headers, show="headings", height=height)
        for header in headers:
            self.tree.heading(header, text=header)
            self.tree.column(header, width=140, anchor=tk.W)
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        for row in rows:
            values = [row.get(header, "") for header in headers]
            self.tree.insert("", tk.END, values=values)

    def refresh(self, rows: Iterable[dict]):
        for child in self.tree.get_children():
            self.tree.delete(child)
        for row in rows:
            values = [row.get(col, "") for col in self.tree["columns"]]
            self.tree.insert("", tk.END, values=values)


__all__ = ["ScrollableTable"]
