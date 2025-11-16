from __future__ import annotations

import tkinter as tk

import customtkinter as ctk

from ..config import HEADERS
from ..services.workflow import PurchaseInput, SaleInput, StockInput, WorkflowCoordinator
from ..ui.widgets import DatePickerEntry
from ..utils.datefmt import format_display_date


class WorkflowView(ctk.CTkFrame):
    def __init__(self, master, coordinator: WorkflowCoordinator, refresh_callback):
        super().__init__(master)
        self.coordinator = coordinator
        self.refresh_callback = refresh_callback
        self.pack(fill="both", expand=True)
        self.log_var = tk.StringVar(value="Prêt à traiter vos flux Achats → Stock → Ventes.")

        title = ctk.CTkLabel(self, text="Automatiser les flux", font=ctk.CTkFont(size=20, weight="bold"))
        title.pack(pady=(16, 8))

        self._build_purchase_section()
        self._build_stock_section()
        self._build_sale_section()
        self._build_return_section()

        log_frame = ctk.CTkFrame(self)
        log_frame.pack(fill="x", padx=16, pady=(8, 16))
        ctk.CTkLabel(log_frame, textvariable=self.log_var, anchor="w").pack(fill="x", padx=12, pady=12)

    # ------------------------------------------------------------------
    # Sections builders
    # ------------------------------------------------------------------
    def _build_purchase_section(self):
        frame = self._section("1. Ajouter une commande dans Achats")
        self.purchase_article = self._field(frame, "Article")
        self.purchase_marque = self._field(frame, "Marque")
        self.purchase_reference = self._field(frame, "Référence")
        self.purchase_quantite = self._field(frame, "Quantité", default="1")
        self.purchase_prix = self._field(frame, "Prix unitaire TTC", default="0")
        self.purchase_frais = self._field(frame, "Frais de colissage", default="0")
        ctk.CTkButton(frame, text="Créer l'achat", command=self._handle_add_purchase).pack(pady=(8, 4))

    def _build_stock_section(self):
        frame = self._section("2. Générer les articles en stock")
        self.stock_purchase_id = self._field(frame, "ID Achat")
        self.stock_sku = self._field(frame, "SKU")
        self.stock_prix = self._field(frame, "Prix de vente", default="0")
        self.stock_lot = self._field(frame, "Lot")
        self.stock_taille = self._field(frame, "Taille")
        ctk.CTkButton(frame, text="Passer en stock", command=self._handle_stock_transfer).pack(pady=(8, 4))

    def _build_sale_section(self):
        frame = self._section("3. Basculer un article vendu vers Ventes/Compta")
        self.sale_sku = self._field(frame, "SKU")
        self.sale_prix = self._field(frame, "Prix de vente", default="0")
        self.sale_frais = self._field(frame, "Frais colissage", default="0")
        self.sale_date = self._field(frame, "Date de vente (JJ/MM/AAAA)", date_picker=True)
        self.sale_lot = self._field(frame, "Lot")
        self.sale_taille = self._field(frame, "Taille")
        ctk.CTkButton(frame, text="Enregistrer la vente", command=self._handle_sale).pack(pady=(8, 4))

    def _build_return_section(self):
        frame = self._section("4. Gérer un retour")
        self.return_sku = self._field(frame, "SKU")
        self.return_note = self._field(frame, "Motif")
        ctk.CTkButton(frame, text="Enregistrer le retour", command=self._handle_return).pack(pady=(8, 4))

    # ------------------------------------------------------------------
    # Actions
    # ------------------------------------------------------------------
    def _handle_add_purchase(self):
        try:
            data = PurchaseInput(
                article=self.purchase_article.get().strip(),
                marque=self.purchase_marque.get().strip(),
                reference=self.purchase_reference.get().strip(),
                quantite=int(self.purchase_quantite.get() or "0"),
                prix_unitaire=float(self.purchase_prix.get() or "0"),
                frais_colissage=float(self.purchase_frais.get() or "0"),
            )
        except ValueError:
            self._log("Valeurs numériques invalides pour l'achat")
            return
        if not data.article or not data.marque:
            self._log("Article et marque sont obligatoires")
            return
        row = self.coordinator.create_purchase(data)
        self._log(f"Achat {row.get(HEADERS['ACHATS'].ID)} ajouté")
        self.refresh_callback()

    def _handle_stock_transfer(self):
        try:
            prix = float(self.stock_prix.get() or "0")
        except ValueError:
            self._log("Prix de vente invalide")
            return
        purchase_id = self.stock_purchase_id.get().strip()
        sku = self.stock_sku.get().strip()
        if not purchase_id or not sku:
            self._log("L'ID achat et le SKU sont obligatoires")
            return
        data = StockInput(
            purchase_id=purchase_id,
            sku=sku,
            prix_vente=prix,
            lot=self.stock_lot.get().strip(),
            taille=self.stock_taille.get().strip(),
        )
        try:
            row = self.coordinator.transfer_to_stock(data)
        except ValueError as exc:
            self._log(str(exc))
            return
        self._log(f"SKU {row.get(HEADERS['STOCK'].SKU)} créé dans le stock")
        self.refresh_callback()

    def _handle_sale(self):
        try:
            prix = float(self.sale_prix.get() or "0")
            frais = float(self.sale_frais.get() or "0")
        except ValueError:
            self._log("Montants invalides pour la vente")
            return
        sku = self.sale_sku.get().strip()
        if not sku:
            self._log("Le SKU est obligatoire pour enregistrer une vente")
            return
        data = SaleInput(
            sku=sku,
            prix_vente=prix,
            frais_colissage=frais,
            date_vente=self.sale_date.get().strip() or None,
            lot=self.sale_lot.get().strip(),
            taille=self.sale_taille.get().strip(),
        )
        try:
            row = self.coordinator.register_sale(data)
        except ValueError as exc:
            self._log(str(exc))
            return
        self._log(f"Vente {row.get(HEADERS['VENTES'].ID)} envoyée vers Ventes/Compta")
        self.refresh_callback()

    def _handle_return(self):
        sku = self.return_sku.get().strip()
        if not sku:
            self._log("Le SKU est obligatoire pour un retour")
            return
        try:
            row = self.coordinator.register_return(sku, self.return_note.get().strip())
        except ValueError as exc:
            self._log(str(exc))
            return
        self._log(f"Retour enregistré pour le SKU {row.get(HEADERS['VENTES'].SKU)}")
        self.refresh_callback()

    # ------------------------------------------------------------------
    # UI helpers
    # ------------------------------------------------------------------
    def _section(self, title: str):
        frame = ctk.CTkFrame(self)
        frame.pack(fill="x", padx=16, pady=8)
        ctk.CTkLabel(frame, text=title, font=ctk.CTkFont(size=16, weight="bold"), anchor="w").pack(
            fill="x", padx=12, pady=(12, 4)
        )
        return frame

    def _field(self, parent, label: str, *, default: str | None = None, date_picker: bool = False):
        row = ctk.CTkFrame(parent)
        row.pack(fill="x", padx=12, pady=4)
        ctk.CTkLabel(row, text=label, width=180, anchor="w").pack(side="left")
        entry = DatePickerEntry(row) if date_picker else ctk.CTkEntry(row)
        if default is not None:
            entry.insert(0, default)
        entry.pack(side="left", fill="x", expand=True, padx=(8, 0))
        return entry

    def _log(self, message: str):
        self.log_var.set(message)
