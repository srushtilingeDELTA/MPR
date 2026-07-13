"""Visible placeholders when Excel/template data is unavailable."""

from __future__ import annotations

MISSING_PREFIX = "[MISSING:"
EMPTY_CELL = "—"
DATA_NOT_FOUND_PREFIX = "Data not found for "
DNF = "DNF"


def chart_insert_message(chart_name: str) -> str:
    name = chart_name.strip() or "this metric"
    return f"Chart for {name} needs to be inserted here."


def data_not_found(what: str) -> str:
    return f"{DATA_NOT_FOUND_PREFIX}{what.strip()}"


def is_data_not_found(text: str) -> bool:
    return text.strip().startswith(DATA_NOT_FOUND_PREFIX)


def is_placeholder_text(text: str) -> bool:
    text = text.strip()
    return text.startswith(MISSING_PREFIX) or text.startswith(DATA_NOT_FOUND_PREFIX)


def missing_label(what: str, *, source: str = "") -> str:
    """Return a user-visible marker for data that must be supplied manually."""
    what = what.strip()
    if source:
        return f"{MISSING_PREFIX} {what} — add from {source}]"
    return f"{MISSING_PREFIX} {what}]"


def is_missing_text(text: str) -> bool:
    return is_placeholder_text(text)


def has_numeric_value(value) -> bool:
    if value is None:
        return False
    try:
        number = float(value)
    except (TypeError, ValueError):
        return False
    return not (number != number or number in (float("inf"), float("-inf")))


def empty_cell() -> str:
    """Return the standard empty table cell marker."""
    return EMPTY_CELL


def has_text_value(value) -> bool:
    if value is None:
        return False
    text = str(value).strip()
    return bool(text) and text.lower() not in {"nan", "none"}
