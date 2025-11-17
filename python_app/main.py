"""Entry point for the CustomTkinter application."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

if __package__ is None or __package__ == "":  # pragma: no cover - convenience for direct execution
    sys.path.append(str(Path(__file__).resolve().parent.parent))

from python_app.controllers.app_controller import AppController


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Vintage ERP UI")
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
    controller = AppController(args.achats_db)
    return controller.run()


if __name__ == "__main__":  # pragma: no cover - CLI entry point
    raise SystemExit(main())
