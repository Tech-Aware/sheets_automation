"""Entry point for the CustomTkinter application."""
from __future__ import annotations

import argparse
from pathlib import Path

from .controllers.app_controller import DEFAULT_WORKBOOK, AppController


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Vintage ERP UI")
    parser.add_argument(
        "workbook",
        nargs="?",
        default=DEFAULT_WORKBOOK,
        type=Path,
        help="Path to the Excel workbook (defaults to Prerelease 1.2.xlsx located at the repo root)",
    )
    parser.add_argument(
        "--achats-db",
        type=Path,
        default=None,
        help=(
            "Chemin vers la base SQLite utilisée pour l'onglet Achats. "
            "Par défaut, python_app/data/achats.db est utilisé et créé si nécessaire."
        ),
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    controller = AppController(Path(args.workbook), args.achats_db)
    return controller.run()


if __name__ == "__main__":  # pragma: no cover - CLI entry point
    raise SystemExit(main())
