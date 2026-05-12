from __future__ import annotations
import datetime as _dt
from typing import Any, Iterable, Mapping
MIN_PERIOD_DATE = _dt.date(2026, 1, 1)
MIN_PERIOD_DATE_ISO = MIN_PERIOD_DATE.isoformat()
PERIOD_DATE_KEYS = ("period_date", "Week", "week", "date")
def parse_period_date(value: Any) -> _dt.date | None:
    if value is None:
        return None
    if isinstance(value, _dt.datetime):
        return value.date()
    if isinstance(value, _dt.date):
        return value
    text = str(value).strip()
    if not text:
        return None
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%m-%d-%Y", "%d-%b-%Y", "%d-%b-%y"):
        try:
            return _dt.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None
def is_on_or_after_min_period(value: Any) -> bool:
    period_date = parse_period_date(value)
    return period_date is not None and period_date >= MIN_PERIOD_DATE
def row_period_date(row: Mapping[str, Any], date_keys: Iterable[str] = PERIOD_DATE_KEYS) -> Any:
    for key in date_keys:
        if key in row and row.get(key) not in (None, ""):
            return row.get(key)
    return None
def should_keep_period_row(row: Mapping[str, Any], date_keys: Iterable[str] = PERIOD_DATE_KEYS) -> bool:
    return is_on_or_after_min_period(row_period_date(row, date_keys))
def filter_period_rows(
    rows: Iterable[Mapping[str, Any]],
    date_keys: Iterable[str] = PERIOD_DATE_KEYS,
) -> list[Mapping[str, Any]]:
    return [row for row in rows if should_keep_period_row(row, date_keys)]