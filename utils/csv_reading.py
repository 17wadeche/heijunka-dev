from __future__ import annotations
from typing import Any
import pandas as pd
def _rewind_if_possible(source: Any) -> None:
    try:
        source.seek(0)
    except (AttributeError, OSError):
        return
def read_csv_resilient(source: Any, **kwargs: Any) -> pd.DataFrame:
    attempts = [
        dict(kwargs),
        {**kwargs, "engine": "python", "sep": None, "on_bad_lines": "skip"},
        {**kwargs, "engine": "python", "sep": ",", "on_bad_lines": "skip"},
    ]
    last_error: Exception | None = None
    for options in attempts:
        _rewind_if_possible(source)
        try:
            return pd.read_csv(source, **options)
        except pd.errors.ParserError as exc:
            last_error = exc
    if last_error is not None:
        raise last_error
    raise pd.errors.ParserError("Unable to parse CSV")