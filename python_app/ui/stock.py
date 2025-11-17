from __future__ import annotations

from datetime import date
import math
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import Mapping, Sequence

import customtkinter as ctk

from ..config import DEFAULT_STOCK_HEADERS, HEADERS
from ..datasources.workbook import WorkbookRepository
from ..services.stock_import import merge_stock_table
from ..services.summaries import _base_reference_from_stock, _build_reference_unit_price_index
from ..services.workflow import StockInput, WorkflowCoordinator
from ..ui.table_view import TableView
from ..ui.tables import ScrollableTable
from ..ui.widgets import DatePickerEntry
from ..utils.datefmt import format_display_date, parse_date_value


def _has_ready_date(value) -> bool:
    if value in (None, ""):
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, str) and value.strip().upper() == "FALSE":
        return False
    return True


class StockCardList(ctk.CTkFrame):
    """Compact list of stock items rendered as tappable rectangles."""

    DEFAULT_COLOR = "#e5e7eb"
    SELECTED_COLOR = "#bfdbfe"
    BORDER_COLOR = "#cbd5e1"
    BORDER_COLOR_SELECTED = "#60a5fa"

    def __init__(self, master, table, *, on_open_details, on_mark_sold, on_bulk_action, on_selection_change=None):
        super().__init__(master)
        self.table = table
        self.on_open_details = on_open_details
        self.on_mark_sold = on_mark_sold
        self.on_bulk_action = on_bulk_action
        self.on_selection_change = on_selection_change
        self._selected_indices: set[int] = set()
        self._cards: dict[int, ctk.CTkFrame] = {}
        ctk.CTkLabel(
            self,
            text=(
                "Vue vignettes (clic pour sélectionner, double-clic pour détailler, "
                "clic droit pour marquer vendu ou appliquer une action de groupe)"
            ),
            anchor="w",
            font=ctk.CTkFont(weight="bold"),
        ).pack(fill="x", padx=12, pady=(8, 4))
        self.container = ctk.CTkScrollableFrame(self, height=240)
        self.container.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self.refresh(self.table.rows)

    def refresh(self, rows: Sequence[dict]):
        indexed_rows = list(enumerate(rows))
        self._render_cards(indexed_rows)

    def filter_by_sku_prefix(self, prefix: str) -> int:
        query = prefix.strip().lower()
        if not query:
            self.refresh(self.table.rows)
            return len(self.table.rows)

        indexed_rows = [
            (idx, row)
            for idx, row in enumerate(self.table.rows)
            if str(row.get(HEADERS["STOCK"].SKU, "")).strip().lower().startswith(query)
        ]
        self._render_cards(indexed_rows)
        return len(indexed_rows)

    def _render_cards(self, indexed_rows: Sequence[tuple[int, dict]]):
        for child in self.container.winfo_children():
            child.destroy()
        self._cards.clear()
        visible_indices = {idx for idx, _ in indexed_rows}
        self._selected_indices = {idx for idx in self._selected_indices if idx in visible_indices}
        for idx, row in indexed_rows:
            self._add_card(idx, row)
        self._update_selection_display()

    def get_selected_indices(self) -> list[int]:
        return sorted(self._selected_indices)

    def _add_card(self, index: int, row: dict):
        sku = str(row.get(HEADERS["STOCK"].SKU, "")).strip()
        label = row.get(HEADERS["STOCK"].ARTICLE, "") or row.get(HEADERS["STOCK"].LIBELLE, "")
        status = "Vendu" if row.get(HEADERS["STOCK"].VENDU_ALT, "") else ""
        subtitle = f"{sku} – {label}" if label else sku
        card = ctk.CTkFrame(self.container, height=76, fg_color=self.DEFAULT_COLOR, border_width=1)
        card.grid_propagate(False)
        card.pack(fill="x", padx=4, pady=2)
        self._cards[index] = card

        text_frame = ctk.CTkFrame(card, fg_color="transparent")
        text_frame.place(relx=0, rely=0, relwidth=1, relheight=1)
        text_frame.grid_propagate(False)

        content_frame = ctk.CTkFrame(text_frame, fg_color="transparent")
        content_frame.pack(fill="both", expand=True, padx=10, pady=8)

        title = ctk.CTkLabel(
            content_frame,
            text=subtitle or "(SKU manquant)",
            anchor="w",
            font=ctk.CTkFont(weight="bold"),
            wraplength=560,
            justify="left",
        )
        title.pack(fill="x")
        metadata_labels = []
        for line in self._build_metadata_lines(row):
            lbl = ctk.CTkLabel(
                content_frame,
                text=line,
                anchor="w",
                font=ctk.CTkFont(size=12),
                wraplength=560,
                justify="left",
            )
            lbl.pack(fill="x")
            metadata_labels.append(lbl)

        if status:
            ctk.CTkLabel(card, text=status, text_color="#0f5132").pack(side="right", padx=8)

        for widget in (card, text_frame, content_frame, title, *metadata_labels):
            self._bind_card_events(widget, index)

    def _bind_card_events(self, widget, index: int):
        widget.bind("<Button-1>", lambda _e, idx=index: self._toggle_selection(idx))
        widget.bind("<Double-Button-1>", lambda _e, idx=index: self.on_open_details(idx))
        widget.bind("<Button-3>", lambda e, idx=index: self._handle_right_click(e, idx))

    def _toggle_selection(self, index: int):
        if index in self._selected_indices:
            self._selected_indices.remove(index)
        else:
            self._selected_indices.add(index)
        self._update_selection_display()

    def _update_selection_display(self):
        for idx in self._cards:
            self._apply_card_style(idx, idx in self._selected_indices)
        if self.on_selection_change is not None:
            self.on_selection_change(self.get_selected_indices())

    def _handle_right_click(self, event, index: int):
        if index not in self._selected_indices:
            self._selected_indices.add(index)
        self._update_selection_display()
        click_date = date.today()
        self._show_context_menu(event, click_date)

    def _show_context_menu(self, event, click_date: date):
        menu = tk.Menu(self, tearoff=False)
        menu.add_command(
            label="Ajouter des détails",
            command=lambda: self.on_open_details(self.get_selected_indices()),
        )
        menu.add_separator()
        menu.add_command(
            label="Déclarer les articles comme vendu",
            command=lambda: self.on_bulk_action(self.get_selected_indices(), "vendu", click_date),
        )
        menu.tk_popup(event.x_root, event.y_root)

    def _apply_card_style(self, index: int, selected: bool):
        fg_color = self.SELECTED_COLOR if selected else self.DEFAULT_COLOR
        border = self.BORDER_COLOR_SELECTED if selected else self.BORDER_COLOR
        card = self._cards.get(index)
        if card is not None:
            card.configure(fg_color=fg_color, border_color=border)

    def _build_metadata_lines(self, row: Mapping) -> list[str]:
        details: list[str] = []

        size_value = self._display_value(row.get(HEADERS["STOCK"].TAILLE))
        if size_value:
            details.append(f"Taille : {size_value}")

        listing_date = self._first_non_empty(
            row,
            (
                HEADERS["STOCK"].DATE_MISE_EN_LIGNE,
                HEADERS["STOCK"].MIS_EN_LIGNE,
                HEADERS["STOCK"].MIS_EN_LIGNE_ALT,
                HEADERS["STOCK"].DATE_MISE_EN_LIGNE_ALT,
            ),
        )
        if listing_date:
            details.append(f"Mis en ligne le {listing_date}")

        publication_date = self._first_non_empty(
            row,
            (
                HEADERS["STOCK"].DATE_PUBLICATION,
                HEADERS["STOCK"].PUBLIE,
                HEADERS["STOCK"].PUBLIE_ALT,
                HEADERS["STOCK"].DATE_PUBLICATION_ALT,
            ),
        )
        if publication_date:
            details.append(f"Publié le {publication_date}")

        lot_value = self._display_value(row.get(HEADERS["STOCK"].LOT_ALT)) or self._display_value(
            row.get(HEADERS["STOCK"].LOT)
        )
        if lot_value:
            details.append(f"Lot {lot_value}")

        return [" | ".join(details)] if details else []

    @staticmethod
    def _first_non_empty(row: Mapping, keys: Sequence[str]) -> str:
        for key in keys:
            value = row.get(key)
            display_value = StockCardList._display_value(value)
            if display_value:
                return display_value
        return ""

    @staticmethod
    def _display_value(value) -> str:
        if value is None:
            return ""
        if isinstance(value, str):
            return value.strip()
        if isinstance(value, float) and math.isnan(value):
            return ""
        return str(value)

    def _handle_right_click_sale(self, idx: int, date_label: str):
        date_text = format_display_date(date.today())
        self.on_bulk_action(self.get_selected_indices(), date_label, date_text)

    def _handle_bulk_sale(self, indices: Sequence[int]):
        if not indices:
            return
        date_text = format_display_date(date.today())
        self.on_bulk_action(indices, "sale", date_text)

    def _handle_bulk_publication(self, indices: Sequence[int]):
        if not indices:
            return
        date_text = format_display_date(date.today())
        self.on_bulk_action(indices, "publication", date_text)

    def _handle_bulk_listing(self, indices: Sequence[int]):
        if not indices:
            return
        date_text = format_display_date(date.today())
        self.on_bulk_action(indices, "listing", date_text)

    def _handle_right_click_mark_sold(self, idx: int):
        self.on_bulk_action(self.get_selected_indices(), "vendu", date.today())


class StockSummaryPanel(ctk.CTkFrame):
    def __init__(self, master, table, *, on_refresh=None):
        super().__init__(master, fg_color="transparent")
        self.table = table
        self.on_refresh = on_refresh
        self.total_pieces = tk.StringVar(value="0")
        self.total_value = tk.StringVar(value="0.0 €")
        self.reference_count = tk.StringVar(value="0")
        self.value_per_reference = tk.StringVar(value="0.0 €")
        self.value_per_piece = tk.StringVar(value="0.0 €")
        self._build()

    def _build(self):
        wrapper = ctk.CTkFrame(self, fg_color="transparent")
        wrapper.pack(fill="both", expand=True, pady=(0, 12))
        title = ctk.CTkLabel(wrapper, text="Indicateurs de stock", font=ctk.CTkFont(size=16, weight="bold"))
        title.pack(anchor="w", padx=12, pady=(0, 8))

        stats_frame = ctk.CTkFrame(wrapper, fg_color="transparent")
        stats_frame.pack(fill="x", padx=12)
        for label, var in (
            ("Pièces en stock", self.total_pieces),
            ("Valeur estimée", self.total_value),
            ("Références uniques", self.reference_count),
            ("Valeur / référence", self.value_per_reference),
            ("Valeur / pièce", self.value_per_piece),
        ):
            item = ctk.CTkFrame(stats_frame)
            item.pack(fill="x", pady=4)
            item.grid_columnconfigure(1, weight=1)
            ctk.CTkLabel(item, text=label, anchor="w").grid(row=0, column=0, sticky="w", padx=12, pady=8)
            ctk.CTkLabel(item, textvariable=var, font=ctk.CTkFont(size=15, weight="bold")).grid(
                row=0, column=1, sticky="e", padx=12, pady=8
            )

        ctk.CTkButton(wrapper, text="Recalculer", command=self._handle_refresh).pack(anchor="w", padx=12, pady=(4, 8))

    def update(self, rows: Sequence[Mapping], achats_rows: Sequence[Mapping] | None = None):
        stats = self._compute_stats(rows, achats_rows)
        self.total_pieces.set(str(stats["stock_pieces"]))
        self.total_value.set(f"{stats['stock_value']:.2f} €")
        self.reference_count.set(str(stats["reference_count"]))
        self.value_per_reference.set(f"{stats['value_per_reference']:.2f} €")
        self.value_per_piece.set(f"{stats['value_per_piece']:.2f} €")

    def _handle_refresh(self):
        if callable(self.on_refresh):
            self.on_refresh()

    @staticmethod
    def _compute_stats(
        rows: Sequence[Mapping], achats_rows: Sequence[Mapping] | None = None
    ) -> dict[str, float | int]:
        pieces = 0
        base_counts: dict[str, int] = {}

        for row in rows:
            vendu = row.get(HEADERS["STOCK"].VENDU_ALT) or row.get(HEADERS["STOCK"].VENDU)
            if vendu:
                continue
            base_reference = _base_reference_from_stock(row)
            pieces += 1
            if base_reference:
                base_counts[base_reference] = base_counts.get(base_reference, 0) + 1

        references = set(base_counts)

        reference_unit_price = _build_reference_unit_price_index(achats_rows or [])
        stock_value = sum(reference_unit_price.get(base, 0.0) * count for base, count in base_counts.items())

        reference_count = len(references)
        value_per_reference = stock_value / reference_count if reference_count else 0.0
        value_per_piece = stock_value / pieces if pieces else 0.0

        return {
            "stock_pieces": pieces,
            "stock_value": round(stock_value, 2),
            "reference_count": reference_count,
            "value_per_reference": round(value_per_reference, 2),
            "value_per_piece": round(value_per_piece, 2),
        }


class StockDetailDialog(ctk.CTkToplevel):
    """Lightweight popup used to enrich a stock item."""

    def __init__(self, master, row: dict, *, on_save):
        super().__init__(master)
        self.on_save = on_save
        self.title("Détails de l'article")
        self.geometry("420x360")
        self.resizable(False, False)
        owner = master.winfo_toplevel() if hasattr(master, "winfo_toplevel") else None
        if owner is not None:
            self.transient(owner)
        self.grab_set()
        self.lift()
        self.focus()
        self.row = row
        self._fields: dict[str, ctk.CTkEntry | ctk.CTkOptionMenu | DatePickerEntry] = {}
        form = ctk.CTkFrame(self)
        form.pack(fill="both", expand=True, padx=12, pady=12)
        self._add_field(
            form,
            "date_mise_en_ligne",
            "Date de mise en ligne",
            self._initial_date(
                row,
                StockTableView._DATE_COLUMNS,
                HEADERS["STOCK"].DATE_MISE_EN_LIGNE,
                HEADERS["STOCK"].MIS_EN_LIGNE,
            ),
            date_picker=True,
        )
        self._add_field(
            form,
            "date_publication",
            "Date de publication",
            self._initial_date(
                row,
                StockTableView._PUBLICATION_COLUMNS,
                HEADERS["STOCK"].DATE_PUBLICATION,
                HEADERS["STOCK"].PUBLIE,
            ),
            date_picker=True,
        )
        self._add_field(
            form,
            "date_vente",
            "Date de vente",
            self._initial_date(
                row,
                StockTableView._SALE_COLUMNS,
                HEADERS["STOCK"].DATE_VENTE,
                HEADERS["STOCK"].VENDU,
            ),
            date_picker=True,
        )
        self._add_field(
            form,
            "prix_vente",
            "Prix de vente",
            row.get(HEADERS["STOCK"].PRIX_VENTE, ""),
        )
        self._add_field(
            form,
            "taille",
            "Taille",
            row.get(HEADERS["STOCK"].TAILLE, ""),
            choices=StockTableView.SIZE_CHOICES,
        )
        self._add_field(
            form,
            "taille_colis",
            "Taille du colis",
            row.get(HEADERS["STOCK"].TAILLE_COLIS) or row.get(HEADERS["STOCK"].TAILLE_COLIS_ALT, ""),
        )
        self._add_field(
            form,
            "lot",
            "Lot",
            row.get(HEADERS["STOCK"].LOT_ALT) or row.get(HEADERS["STOCK"].LOT, ""),
            choices=StockTableView.LOT_CHOICES,
        )
        ctk.CTkButton(form, text="Enregistrer", command=self._save).pack(fill="x", pady=(12, 4))

    def _add_field(
        self,
        parent,
        key: str,
        label: str,
        value,
        *,
        date_picker: bool = False,
        choices: Sequence[str] | None = None,
    ):
        row = ctk.CTkFrame(parent)
        row.pack(fill="x", pady=4)
        ctk.CTkLabel(row, text=label, width=160, anchor="w").pack(side="left")
        if choices:
            available_values = ["", *choices]
            if value not in (None, "", *choices):
                available_values.insert(1, str(value))
            var = tk.StringVar(value=str(value or ""))
            entry: ctk.CTkOptionMenu | ctk.CTkEntry | DatePickerEntry = ctk.CTkOptionMenu(
                row, values=available_values, variable=var
            )
        else:
            entry = DatePickerEntry(row) if date_picker else ctk.CTkEntry(row)
            if value not in (None, ""):
                entry.insert(0, str(value))
        entry.pack(side="left", fill="x", expand=True, padx=(8, 0))
        self._fields[key] = entry

    @staticmethod
    def _initial_date(row: dict, aliases: set[str], *keys: str) -> str:
        for key in (*keys, *aliases):
            value = row.get(key)
            if value not in (None, ""):
                return value
        return ""

    def _save(self):
        updates = {name: field.get().strip() for name, field in self._fields.items()}
        self.on_save(updates)
        self.destroy()


class StockTableView(TableView):
    DISPLAY_HEADERS: Sequence[str] = DEFAULT_STOCK_HEADERS

    SIZE_CHOICES: Sequence[str] = ("Petit", "Moyen", "Grand")
    LOT_CHOICES: Sequence[str] = ("2", "3", "4", "5")

    _DISPLAY_FALLBACKS = {
        HEADERS["STOCK"].LOT_ALT: HEADERS["STOCK"].LOT,
        HEADERS["STOCK"].VALIDER_SAISIE_ALT: HEADERS["STOCK"].VALIDER_SAISIE,
        HEADERS["STOCK"].PUBLIE_ALT: HEADERS["STOCK"].PUBLIE,
        HEADERS["STOCK"].DATE_PUBLICATION_ALT: HEADERS["STOCK"].DATE_PUBLICATION,
    }

    _DATE_COLUMNS = {
        HEADERS["STOCK"].MIS_EN_LIGNE,
        HEADERS["STOCK"].MIS_EN_LIGNE_ALT,
        HEADERS["STOCK"].DATE_MISE_EN_LIGNE,
        HEADERS["STOCK"].DATE_MISE_EN_LIGNE_ALT,
    }

    _PUBLICATION_COLUMNS = {
        HEADERS["STOCK"].PUBLIE,
        HEADERS["STOCK"].PUBLIE_ALT,
        HEADERS["STOCK"].DATE_PUBLICATION,
        HEADERS["STOCK"].DATE_PUBLICATION_ALT,
    }

    _SALE_COLUMNS = {
        HEADERS["STOCK"].VENDU,
        HEADERS["STOCK"].VENDU_ALT,
        HEADERS["STOCK"].DATE_VENTE,
        HEADERS["STOCK"].DATE_VENTE_ALT,
    }

    _CHECKBOX_STYLE_COLUMNS = (
        _DATE_COLUMNS
        | _PUBLICATION_COLUMNS
        | _SALE_COLUMNS
    )

    def __init__(self, master, table, *, workflow: WorkflowCoordinator | None = None, on_table_changed=None):
        self.workflow = workflow
        self.search_var: tk.StringVar | None = None
        self.search_entry: ctk.CTkEntry | None = None
        super().__init__(master, table, on_table_changed=on_table_changed)
        self.status_var.set(
            "Sélectionnez une vignette (clic), double-clic pour détailler, clic droit pour actions rapides."
        )

    def _create_table_widget(self, parent, headers: Sequence[str], rows: Sequence[dict]):
        return ScrollableTable(
            parent,
            headers,
            rows,
            height=20,
            on_cell_edited=self._on_cell_edit,
            on_cell_activated=self._handle_cell_activation,
            column_width=160,
            column_widths={"ID": 34},
            value_formatter=self._format_cell_value,
            dropdown_choices={
                HEADERS["STOCK"].TAILLE: self.SIZE_CHOICES,
                HEADERS["STOCK"].LOT_ALT: self.LOT_CHOICES,
            },
        )

    def _visible_headers(self) -> Sequence[str]:
        self._ensure_display_aliases()
        return self.DISPLAY_HEADERS

    def _should_show_table(self) -> bool:
        return False

    def refresh(self):
        self._ensure_display_aliases()
        super().refresh()
        if hasattr(self, "card_list"):
            self.card_list.refresh(self.table.rows)
        if hasattr(self, "summary_panel"):
            achats_rows = self.workflow.achats.rows if self.workflow is not None else None
            self.summary_panel.update(self.table.rows, achats_rows)

    def _build_extra_controls(self, parent):
        parent.grid_columnconfigure(0, weight=3, uniform="stock")
        parent.grid_columnconfigure(1, weight=2, uniform="stock")
        parent.grid_rowconfigure(0, weight=0)
        self.table_frame.grid_configure(columnspan=2, sticky="nsew")
        ctk.CTkButton(
            self.actions_frame,
            text="Supprimer la sélection",
            command=self._delete_selected_rows,
        ).pack(side="left")
        search_frame = ctk.CTkFrame(self.actions_frame, fg_color="transparent")
        search_frame.pack(side="left", padx=(12, 0))
        ctk.CTkLabel(search_frame, text="Rechercher SKU:").pack(side="left", padx=(0, 6))
        self.search_var = tk.StringVar()
        self.search_entry = ctk.CTkEntry(search_frame, textvariable=self.search_var, width=180)
        self.search_entry.pack(side="left")
        self.search_entry.bind("<Return>", lambda _e: self._handle_sku_search())
        self.search_entry.bind("<KeyRelease>", self._handle_live_sku_filter)
        ctk.CTkButton(search_frame, text="Go", width=56, command=self._handle_sku_search).pack(
            side="left", padx=(8, 0)
        )

        parent.grid_rowconfigure(1, weight=1)
        self.card_list = StockCardList(
            parent,
            self.table,
            on_open_details=self._open_detail_dialog,
            on_mark_sold=self._handle_card_sale,
            on_bulk_action=self._handle_bulk_card_action,
            on_selection_change=self._handle_card_selection_change,
        )
        self.card_list.grid(row=1, column=0, sticky="nsew", padx=(0, 8))

        self.summary_panel = StockSummaryPanel(parent, self.table, on_refresh=self.refresh)
        self.summary_panel.grid(row=1, column=1, sticky="nsew", padx=(8, 0))

    def _handle_card_selection_change(self, indices: Sequence[int]):
        if not self.table_widget:
            return
        self.table_widget.select_rows(indices)

    def _handle_cell_activation(self, row_index: int, column: str):
        if row_index < 0 or row_index >= len(self.table.rows):
            return
        if column == HEADERS["STOCK"].VENDU_ALT:
            date_text = format_display_date(date.today())
            self._set_sale_columns(self.table.rows[row_index], f"Vendu le {date_text}")
            self.status_var.set(f"Article {row_index + 1} marqué vendu le {date_text}")
            self._notify_data_changed()
            return
        self._open_detail_dialog(row_index)

    def _open_detail_dialog(self, row_indices: int | Sequence[int]):
        indices = [row_indices] if isinstance(row_indices, int) else list(row_indices)
        valid_indices = [idx for idx in indices if 0 <= idx < len(self.table.rows)]
        if not valid_indices:
            return
        primary_index = valid_indices[0]
        dialog = StockDetailDialog(
            self,
            self.table.rows[primary_index],
            on_save=lambda updates: self._save_detail(valid_indices, updates),
        )
        dialog.focus()

    def _save_detail(self, row_indices: Sequence[int], updates: Mapping[str, str]):
        for row_index in row_indices:
            self._apply_detail_updates(row_index, updates)
        self._notify_data_changed()

    def _apply_detail_updates(self, row_index: int, updates: Mapping[str, str]):
        for name, value in updates.items():
            if name == "date_mise_en_ligne":
                self._apply_date(row_index, value, self._DATE_COLUMNS, HEADERS["STOCK"].DATE_MISE_EN_LIGNE)
            elif name == "date_publication":
                self._apply_date(row_index, value, self._PUBLICATION_COLUMNS, HEADERS["STOCK"].DATE_PUBLICATION)
            elif name == "date_vente":
                self._apply_date(row_index, value, self._SALE_COLUMNS, HEADERS["STOCK"].DATE_VENTE)
            elif name == "prix_vente":
                self.table.rows[row_index][HEADERS["STOCK"].PRIX_VENTE] = value
            elif name == "taille":
                self.table.rows[row_index][HEADERS["STOCK"].TAILLE] = value
            elif name == "taille_colis":
                self.table.rows[row_index][HEADERS["STOCK"].TAILLE_COLIS_ALT] = value
            elif name == "lot":
                self.table.rows[row_index][HEADERS["STOCK"].LOT_ALT] = value

    def _apply_date(self, row_index: int, value: str, columns: set[str], primary_key: str):
        if not (0 <= row_index < len(self.table.rows)):
            return
        normalized = self._normalize_date_input(value)
        for key in columns:
            self.table.rows[row_index][key] = normalized
        self.table.rows[row_index][primary_key] = normalized

    def _handle_live_sku_filter(self, _event=None):
        if not self.search_var:
            return
        self._apply_sku_filter(self.search_var.get())

    def _handle_sku_search(self):
        if not self.search_var:
            return
        self._apply_sku_filter(self.search_var.get())

    def _apply_sku_filter(self, prefix: str):
        results = self.card_list.filter_by_sku_prefix(prefix) if hasattr(self, "card_list") else 0
        query = prefix.strip()
        if not query:
            self.status_var.set("Toutes les vignettes sont affichées.")
            return

        if results:
            self.status_var.set(f"{results} résultat(s) pour le préfixe SKU '{query}'.")
        else:
            self.status_var.set("Aucun résultat pour ce préfixe SKU.")

    def _ensure_display_aliases(self):
        for row in self.table.rows:
            for header in self.DISPLAY_HEADERS:
                row.setdefault(header, "")
            for alias, source in self._DISPLAY_FALLBACKS.items():
                if alias in row and row[alias] not in (None, ""):
                    continue
                row[alias] = row.get(source, "")

    def _format_cell_value(self, header: str, value: object) -> str:
        if header in self._CHECKBOX_STYLE_COLUMNS and not _has_ready_date(value):
            return "☐"
        return "" if value is None else str(value)

    def _delete_selected_rows(self):
        indices = self.card_list.get_selected_indices()
        if not indices:
            self.status_var.set("Sélectionnez au moins une vignette à supprimer.")
            return
        removed = self._delete_rows_by_indices(indices)
        if removed:
            self.status_var.set(f"{removed} article(s) supprimé(s) du stock")
        else:
            self.status_var.set("Aucun article supprimé")

    def _ensure_sale_requirements(self, row: Mapping) -> bool:
        return bool(row.get(HEADERS["STOCK"].PRIX_VENTE) and row.get(HEADERS["STOCK"].TAILLE_COLIS_ALT))

    def _handle_card_sale(self, row_index: int, click_date: date):
        if not (0 <= row_index < len(self.table.rows)):
            return
        row = self.table.rows[row_index]
        if not self._ensure_sale_requirements(row):
            self.status_var.set(
                "Impossible de déclarer la vente : prix de vente et taille du colis requis."
            )
            return
        date_text = format_display_date(click_date)
        sale_value = f"Vendu le {date_text}"
        self._set_sale_columns(row, sale_value)
        self.status_var.set(f"Article {row_index + 1} marqué vendu le {date_text}")
        self._notify_data_changed()

    def _apply_date_to_rows(self, indices: Sequence[int], columns: set[str], label: str, date_text: str):
        updated = 0
        for idx in indices:
            if not (0 <= idx < len(self.table.rows)):
                continue
            row = self.table.rows[idx]
            for key in columns:
                row[key] = date_text
            updated += 1
        if updated:
            self.status_var.set(f"{label} mise à jour pour {updated} article(s)")
            self._notify_data_changed()

    def _handle_bulk_card_action(self, indices: Sequence[int], action: str, click_date: date):
        if not indices:
            self.status_var.set("Aucune vignette sélectionnée")
            return
        date_text = format_display_date(click_date)
        if action == "mise_en_ligne":
            self._apply_date_to_rows(indices, self._DATE_COLUMNS, "Date de mise en ligne", date_text)
            return
        if action == "publication":
            self._apply_date_to_rows(indices, self._PUBLICATION_COLUMNS, "Date de publication", date_text)
            return
        if action == "vente":
            self._apply_date_to_rows(indices, self._SALE_COLUMNS, "Date de vente", date_text)
            return
        if action == "vendu":
            success = 0
            failures = 0
            sale_value = f"Vendu le {date_text}"
            for idx in indices:
                if not (0 <= idx < len(self.table.rows)):
                    continue
                row = self.table.rows[idx]
                if not self._ensure_sale_requirements(row):
                    failures += 1
                    continue
                self._set_sale_columns(row, sale_value)
                success += 1
            messages: list[str] = []
            if success:
                messages.append(f"Articles vendus mis à jour pour {success} vignette(s)")
                self._notify_data_changed()
            if failures:
                messages.append(
                    f"{failures} vignette(s) non mises à jour : prix de vente et taille du colis requis"
                )
            if messages:
                self.status_var.set(" ".join(messages))
            return

    @staticmethod
    def _normalize_date_input(value: str | None) -> str | None:
        if value is None:
            return None
        text = value.strip()
        if not text:
            return ""
        parsed = parse_date_value(text)
        return format_display_date(parsed) if parsed else text

    def _set_sale_columns(self, row: Mapping, sale_value: str):
        for column in self._SALE_COLUMNS:
            row[column] = sale_value
        row[HEADERS["STOCK"].DATE_VENTE] = sale_value


class StockOptionsView(ctk.CTkFrame):
    def __init__(self, master, table, refresh_callback):
        super().__init__(master)
        self.table = table
        self.refresh_callback = refresh_callback
        self.status_var = tk.StringVar(value="")
        self.pack(fill="both", expand=True)
        self._build_import_section()

    def _build_import_section(self):
        frame = ctk.CTkFrame(self)
        frame.pack(fill="x", padx=16, pady=16)
        ctk.CTkLabel(
            frame,
            text="Importer un stock existant",
            font=ctk.CTkFont(size=16, weight="bold"),
            anchor="w",
        ).pack(fill="x", padx=12, pady=(12, 8))
        ctk.CTkLabel(
            frame,
            text=(
                "Utilisez cette option pour fusionner un XLSX de stock avec le tableau actuel. "
                "Les nouveaux SKU seront ajoutés, les existants laissés intacts."
            ),
            anchor="w",
            wraplength=800,
            justify="left",
        ).pack(fill="x", padx=12)
        ctk.CTkButton(frame, text="Charger un XLSX pour le stock", command=self._handle_import_xlsx).pack(
            pady=(8, 12), padx=12
        )

        ctk.CTkLabel(self, textvariable=self.status_var, anchor="w").pack(fill="x", padx=16, pady=(4, 16))

    def _handle_import_xlsx(self):
        path = filedialog.askopenfilename(
            title="Importer le stock",
            filetypes=(
                ("Excel", "*.xlsx *.xlsm"),
                ("Tous les fichiers", "*.*"),
            ),
        )
        if not path:
            return
        try:
            repository = WorkbookRepository(path)
            sheet_name = self._resolve_stock_sheet(repository)
            if sheet_name is None:
                raise ValueError("Ce classeur ne contient aucun onglet exploitable.")
            source_table = repository.load_table(sheet_name)
        except Exception as exc:  # pragma: no cover - UI guard
            messagebox.showerror("Import du stock", f"Impossible de lire le fichier sélectionné : {exc}")
            return
        added = merge_stock_table(self.table, source_table)
        filename = Path(path).name
        if added:
            self.status_var.set(f"{added} article(s) importé(s) depuis {filename}.")
            self.refresh_callback()
        else:
            self.status_var.set(f"Aucun nouvel article à importer depuis {filename}.")

    @staticmethod
    def _resolve_stock_sheet(repository: WorkbookRepository) -> str | None:
        sheet_names = list(repository.available_tables())
        if not sheet_names:
            return None
        for name in sheet_names:
            if name.lower() == "stock":
                return name
        return sheet_names[0]
