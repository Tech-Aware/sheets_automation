from __future__ import annotations

import customtkinter as ctk


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
