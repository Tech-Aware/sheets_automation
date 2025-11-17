"""Lightweight performance instrumentation helpers.

The UI code frequently performs heavy computations (table rebuilds, KPI
aggregations, large Treeview refreshes).  This module centralises timing so
we can consistently identify the slowest user-facing actions and print them
to the console while the app runs.
"""

from __future__ import annotations

from collections import deque
from contextlib import contextmanager
from dataclasses import dataclass, field
from threading import Lock
from time import perf_counter
from typing import Callable, Iterable, Iterator, Mapping


@dataclass(frozen=True)
class TimedEvent:
    """Captured timing information for a single action."""

    name: str
    duration_ms: float
    metadata: Mapping[str, object] = field(default_factory=dict)


class PerformanceMonitor:
    """Collect and report the slowest UI actions.

    - ``track`` is used as a context manager around critical sections.
    - Recent events are stored in a bounded deque to avoid unbounded memory
      usage while still surfacing hot spots.
    - When a duration exceeds ``log_threshold_ms``, the event is immediately
      printed so that operators can see laggy interactions in real time.
    """

    def __init__(
        self,
        *,
        max_events: int = 200,
        log_threshold_ms: float = 50.0,
        sink: Callable[[TimedEvent], None] | None = None,
    ) -> None:
        self._events: deque[TimedEvent] = deque(maxlen=max_events)
        self._lock = Lock()
        self._log_threshold_ms = log_threshold_ms
        self._sink = sink or self._console_sink

    @contextmanager
    def track(self, name: str, metadata: Mapping[str, object] | None = None) -> Iterator[None]:
        start = perf_counter()
        try:
            yield
        finally:
            duration_ms = (perf_counter() - start) * 1000
            event = TimedEvent(name=name, duration_ms=duration_ms, metadata=metadata or {})
            self._record(event)

    def slowest(self, limit: int = 5) -> list[TimedEvent]:
        with self._lock:
            return sorted(self._events, key=lambda evt: evt.duration_ms, reverse=True)[:limit]

    def _record(self, event: TimedEvent) -> None:
        with self._lock:
            self._events.append(event)
        if event.duration_ms >= self._log_threshold_ms:
            self._sink(event)

    @staticmethod
    def _console_sink(event: TimedEvent) -> None:
        metadata = (
            " " + " ".join(f"{key}={value}" for key, value in event.metadata.items())
            if event.metadata
            else ""
        )
        print(f"[perf] {event.name} {event.duration_ms:.1f} ms{metadata}")


performance_monitor = PerformanceMonitor()


def format_report(events: Iterable[TimedEvent]) -> str:
    """Return a human-readable report of recorded timings."""

    lines = ["Actions utilisateur les plus lentes :"]
    for event in events:
        metadata = (
            " (" + ", ".join(f"{key}={value}" for key, value in event.metadata.items()) + ")"
            if event.metadata
            else ""
        )
        lines.append(f"- {event.name}: {event.duration_ms:.1f} ms{metadata}")
    return "\n".join(lines)


__all__ = ["performance_monitor", "TimedEvent", "format_report"]
