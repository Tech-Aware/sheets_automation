from __future__ import annotations

import customtkinter as ctk


class LoadingDialog(ctk.CTkToplevel):
    """Modal-like dialog displayed while data is refreshing."""

    def __init__(self, master, *, message: str):
        super().__init__(master)
        self.title("Chargement en cours")
        self.resizable(False, False)
        self.attributes("-topmost", True)
        self.geometry("360x140")
        self.protocol("WM_DELETE_WINDOW", lambda: None)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        self.message_label = ctk.CTkLabel(
            self,
            text=message,
            wraplength=320,
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        self.message_label.grid(row=0, column=0, padx=16, pady=(16, 8), sticky="ew")

        self.progress = ctk.CTkProgressBar(self, width=280)
        self.progress.grid(row=1, column=0, padx=16, pady=(0, 8), sticky="ew")
        self.progress.set(0)

        self.percent_label = ctk.CTkLabel(self, text="0%", font=ctk.CTkFont(size=12))
        self.percent_label.grid(row=2, column=0, pady=(0, 12))

    def update_progress(self, value: float):
        bounded = max(0.0, min(1.0, value))
        self.progress.set(bounded)
        self.percent_label.configure(text=f"{int(bounded * 100)}%")
        self.update_idletasks()

    def close(self):
        try:
            self.destroy()
        except Exception:
            pass
