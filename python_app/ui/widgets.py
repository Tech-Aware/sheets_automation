"""Reusable UI helpers such as date pickers."""
from __future__ import annotations

import calendar
from datetime import date
from functools import partial
from typing import Callable

import customtkinter as ctk
import tkinter as tk

from ..config import MONTH_NAMES_FR
from ..utils.datefmt import format_display_date, parse_date_value


class CalendarPopup(ctk.CTkToplevel):
    """Lightweight calendar popup used by :class:`DatePickerEntry`."""

    def __init__(self, master, on_date_selected: Callable[[date], None], initial: date | None = None):
        super().__init__(master)
        self.on_date_selected = on_date_selected
        self.displayed_date = initial or date.today()
        self._year = self.displayed_date.year
        self._month = self.displayed_date.month
        self.title("Sélectionnez une date")
        self.resizable(False, False)
        self.attributes("-topmost", True)
        self.configure(padx=12, pady=12)
        self.body = ctk.CTkFrame(self)
        self.body.pack(fill="both", expand=True)
        self.header = ctk.CTkFrame(self.body)
        self.header.pack(fill="x", pady=(0, 8))
        self.calendar_frame = ctk.CTkFrame(self.body)
        self.calendar_frame.pack(fill="both", expand=True)
        self._build_header()
        self._render_calendar()

    def position_near(self, widget):
        self.update_idletasks()
        x = widget.winfo_rootx()
        y = widget.winfo_rooty() + widget.winfo_height()
        self.geometry(f"+{x}+{y}")

    def _build_header(self):
        for child in self.header.winfo_children():
            child.destroy()
        ctk.CTkButton(self.header, text="◀", width=32, command=self._prev_month).pack(side="left")
        month_label = MONTH_NAMES_FR[self._month - 1]
        title = ctk.CTkLabel(
            self.header,
            text=f"{month_label} {self._year}",
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        title.pack(side="left", expand=True)
        ctk.CTkButton(self.header, text="▶", width=32, command=self._next_month).pack(side="right")

    def _render_calendar(self):
        for child in self.calendar_frame.winfo_children():
            child.destroy()
        weekday_row = ctk.CTkFrame(self.calendar_frame)
        weekday_row.pack(fill="x", pady=(0, 4))
        for label in ("Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"):
            ctk.CTkLabel(weekday_row, text=label, width=40).pack(side="left", expand=True)
        cal = calendar.Calendar(firstweekday=0)
        for week in cal.monthdayscalendar(self._year, self._month):
            row = ctk.CTkFrame(self.calendar_frame)
            row.pack(fill="x", pady=1)
            for day in week:
                if day == 0:
                    ctk.CTkLabel(row, text="", width=40).pack(side="left", expand=True)
                    continue
                button = ctk.CTkButton(
                    row,
                    text=str(day),
                    width=40,
                    command=partial(self._select_day, day),
                )
                button.pack(side="left", expand=True, padx=1)

    def _select_day(self, day: int):
        chosen = date(self._year, self._month, day)
        self.on_date_selected(chosen)
        self.destroy()

    def _prev_month(self):
        if self._month == 1:
            self._month = 12
            self._year -= 1
        else:
            self._month -= 1
        self._build_header()
        self._render_calendar()

    def _next_month(self):
        if self._month == 12:
            self._month = 1
            self._year += 1
        else:
            self._month += 1
        self._build_header()
        self._render_calendar()


class DatePickerEntry(ctk.CTkFrame):
    """Entry widget that shows a calendar popup on demand."""

    def __init__(self, master, placeholder: str | None = None):
        super().__init__(master)
        self.entry = ctk.CTkEntry(self)
        if placeholder:
            self.entry.insert(0, placeholder)
        self.entry.pack(side="left", fill="x", expand=True)
        self.calendar_button = ctk.CTkButton(self, text="📅", width=42, command=self._open_calendar)
        self.calendar_button.pack(side="left", padx=(6, 0))
        self._popup: CalendarPopup | None = None

    # ------------------------------------------------------------------
    # Proxy methods so the widget mimics a regular entry
    # ------------------------------------------------------------------
    def get(self):
        return self.entry.get()

    def delete(self, start, end=None):
        self.entry.delete(start, end)

    def insert(self, index, value):
        self.entry.insert(index, value)

    def bind(self, sequence=None, func=None, add=None):
        return self.entry.bind(sequence, func, add=add)

    def focus(self):  # pragma: no cover - UI glue
        self.entry.focus()

    def _close_popup(self):
        if self._popup is not None:
            try:
                self._popup.destroy()
            except tk.TclError:
                pass
            self._popup = None

    def _open_calendar(self):
        if self._popup is not None and self._popup.winfo_exists():
            self._popup.focus()
            return
        current_value = self.get().strip()
        initial = None
        if current_value:
            initial = parse_date_value(current_value)
        popup = CalendarPopup(self, self._on_date_selected, initial)
        popup.position_near(self.entry)
        popup.bind("<Destroy>", lambda _event: setattr(self, "_popup", None))
        self._popup = popup

    def _on_date_selected(self, chosen: date):
        self.delete(0, tk.END)
        self.insert(0, format_display_date(chosen))
        self._close_popup()


__all__ = ["DatePickerEntry", "CalendarPopup"]
