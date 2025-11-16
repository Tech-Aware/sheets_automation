from __future__ import annotations

import tkinter as tk

import customtkinter as ctk

from ..config import MONTH_NAMES_FR


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
