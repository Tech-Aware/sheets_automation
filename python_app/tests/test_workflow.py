import math
import sys
from pathlib import Path

# Ensure the repository root is on sys.path so that ``python_app`` can be imported
sys.path.append(str(Path(__file__).resolve().parents[2]))

from python_app.services.workflow import WorkflowCoordinator


def test_next_numeric_id_accepts_float_strings():
    rows = [
        {"ID": "1.0"},
        {"ID": 2},
        {"ID": "003"},
        {"ID": " 4 "},
    ]
    assert WorkflowCoordinator._next_numeric_id(rows, "ID") == 5


def test_next_numeric_id_ignores_invalid_values():
    rows = [
        {"ID": "abc"},
        {"ID": math.nan},
        {"ID": ""},
        {},
    ]
    assert WorkflowCoordinator._next_numeric_id(rows, "ID") == 1


def test_next_numeric_id_supports_float_instances():
    rows = [
        {"ID": 5.0},
        {"ID": "6.0"},
    ]
    assert WorkflowCoordinator._next_numeric_id(rows, "ID") == 7
