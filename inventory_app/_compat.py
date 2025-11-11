"""Compatibility helpers for older Python runtimes."""
from __future__ import annotations

import sys
from typing import Dict, Any

# ``dataclasses`` gained ``slots`` support in Python 3.10. Older runtimes, such as
# the Python 3.9 interpreter bundled with some Windows installations, do not
# accept the argument which previously caused the application to fail at import
# time.
DATACLASS_KWARGS: Dict[str, Any] = {"slots": True} if sys.version_info >= (3, 10) else {}
