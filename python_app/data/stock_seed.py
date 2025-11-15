"""Seed data matching the Stock sheet."""
from __future__ import annotations


def _build_rows():
    rows = []

    def add_block(purchase_id: int, article: str, reference: str, sku_suffixes: list[int], *, date_livraison: str, date_stock: str):
        for suffix in sku_suffixes:
            rows.append(
                {
                    "id": str(purchase_id),
                    "libelle": article,
                    "article": article,
                    "reference": reference,
                    "sku": f"{reference}-{suffix}",
                    "date_livraison": date_livraison,
                    "date_mise_en_stock": date_stock,
                }
            )

    add_block(
        3,
        "VÊTEMENTS PERSO Homme/ Benoît",
        "VPHB",
        list(range(3, 11)),
        date_livraison="01/03/2020",
        date_stock="01/09/2020",
    )

    add_block(
        1,
        "SNEAKERS MIX Homme/ Femme",
        "SM",
        [15, 16, 17],
        date_livraison="29/08/2025",
        date_stock="30/08/2025",
    )

    add_block(
        4,
        "JEANS LEVIS Femme/ Kevin",
        "JLFK",
        [21, 27, 31, 32, 37, 44, 46, 55, 60, 65, 69, 70, 72, 74],
        date_livraison="04/09/2025",
        date_stock="05/09/2025",
    )

    add_block(
        5,
        "VÊTEMENTS TYMAN Homme",
        "VTH",
        [10, 11, 12, 13],
        date_livraison="06/09/2025",
        date_stock="07/09/2025",
    )

    return rows


SEED_STOCK = _build_rows()
