"""Public entry points for launching the Tkinter programme."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

if __package__ in {None, ""}:  # pragma: no cover - executed for direct script runs
    sys.path.append(str(Path(__file__).resolve().parent.parent))

from inventory_app.ui import run_app


def main(argv: list[str] | None = None) -> None:
    """Launch the graphical application."""

    parser = argparse.ArgumentParser(description="Console de gestion des stocks")
    parser.add_argument(
        "workbook",
        nargs="?",
        default=str(Path(__file__).resolve().parent.parent / "Copie de TEST (3).xlsx"),
        help="Chemin du classeur Excel à analyser",
    )
    args = parser.parse_args(argv)

    run_app(Path(args.workbook))


if __name__ == "__main__":  # pragma: no cover - manual execution helper
    main()
