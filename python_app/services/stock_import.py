"""Helpers to merge imported stock rows into the in-memory table."""
from __future__ import annotations

from typing import Iterable, Mapping, Sequence

from ..config import HEADERS
from ..datasources.workbook import TableData

_ID_HEADER = HEADERS["STOCK"].ID
_SKU_HEADER = HEADERS["STOCK"].SKU


def merge_stock_table(target: TableData | None, source: TableData | None) -> int:
    """Append rows from ``source`` into ``target`` while preventing duplicates."""

    if target is None or source is None:
        return 0
    if not source.rows:
        return 0
    header_map = _build_header_map(source.headers)
    existing_signatures = _build_existing_signatures(target.rows)
    target_headers = list(target.headers)
    imported = 0
    for row in source.rows:
        signature = _row_signature(row, header_map)
        if signature is None or signature in existing_signatures:
            continue
        normalized = _normalize_row(row, target_headers, header_map)
        target.rows.append(normalized)
        existing_signatures.add(signature)
        imported += 1
    return imported


def _build_header_map(headers: Sequence[str]) -> dict[str, str]:
    mapping: dict[str, str] = {}
    for header in headers:
        if not isinstance(header, str):
            continue
        normalized = _normalize_header(header)
        if normalized:
            mapping[normalized] = header
    return mapping


def _normalize_header(value: str) -> str:
    return "".join(ch for ch in value.lower() if ch.isalnum()) if value else ""


def _build_existing_signatures(rows: Iterable[Mapping[str, object]]) -> set[tuple[str, str]]:
    signatures: set[tuple[str, str]] = set()
    for row in rows:
        signature = (
            _normalize_signature_value(row.get(_ID_HEADER)),
            _normalize_signature_value(row.get(_SKU_HEADER)),
        )
        if signature == ("", ""):
            continue
        signatures.add(signature)
    return signatures


def _row_signature(row: Mapping[str, object], header_map: Mapping[str, str]) -> tuple[str, str] | None:
    signature = (
        _normalize_signature_value(_extract_value(row, header_map, _ID_HEADER)),
        _normalize_signature_value(_extract_value(row, header_map, _SKU_HEADER)),
    )
    if signature == ("", ""):
        return None
    return signature


def _normalize_signature_value(value) -> str:
    if value in (None, ""):
        return ""
    if isinstance(value, (int, float)):
        if isinstance(value, float) and not value.is_integer():
            return str(value)
        return str(int(value))
    return str(value).strip().upper()


def _normalize_row(row: Mapping[str, object], headers: Sequence[str], header_map: Mapping[str, str]) -> dict:
    normalized: dict[str, object] = {}
    for header in headers:
        normalized[header] = _extract_value(row, header_map, header)
    return normalized


def _extract_value(row: Mapping[str, object], header_map: Mapping[str, str], header: str):
    normalized = _normalize_header(header)
    source_header = header_map.get(normalized, header)
    return row.get(source_header, "")


__all__ = ["merge_stock_table"]
