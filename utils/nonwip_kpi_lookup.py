from __future__ import annotations
from typing import Any, Callable
import json
import re
import numpy as np
import pandas as pd
FORTY_HOUR_TEAMS = {
    "SVT", "PVH", "NV", "Enabling Technologies", "DBS", "PH", "Spine",
    "PSS", "SCS", "TDD", "ACM", "VSS", "Endoscopy", "Surgical AST-GST",
    "PH-NM MEIC", "TCT",
}
ENTERPRISE_PEOPLE_COUNT_FROM_NW_TEAMS = {
    "ENT", "DBS", "NV", "Enabling Technologies", "Spine", "PH", "SCS",
    "TDD", "ACM", "CPT", "DS", "CDS", "NI", "VSS", "Endoscopy",
    "Surgical AST-GST", "PH-NM MEIC", "TCT",
}
PERSON_CAPACITY_OVERRIDES = {
    "chelsey": 16.0,
    "mg": 36.0, 
    "lindsey": 32.0
}
def _to_float(value: Any, default: float = 0.0) -> float:
    try:
        out = pd.to_numeric(value, errors="coerce")
        if pd.isna(out):
            return default
        return float(out)
    except Exception:
        return default
def _sum_col(df: pd.DataFrame | None, col: str) -> float:
    if df is None or df.empty or col not in df.columns:
        return 0.0
    return float(pd.to_numeric(df[col], errors="coerce").fillna(0.0).sum())
def _capacity_from_count_with_person_overrides(
    count: float,
    default_hours: float,
    wk_people: pd.DataFrame | None,
) -> float:
    capacity_hours = float(count) * float(default_hours)
    if wk_people is None or wk_people.empty or "person" not in wk_people.columns:
        return capacity_hours
    people = (
        wk_people["person"]
        .dropna()
        .map(_person_key)
        .drop_duplicates()
    )
    for person in people:
        if person in PERSON_CAPACITY_OVERRIDES:
            capacity_hours += PERSON_CAPACITY_OVERRIDES[person] - float(default_hours)
    return capacity_hours
def _normalize_person_name(name: Any) -> str:
    return " ".join(str(name or "").strip().split())
def _person_key(name: Any) -> str:
    s = _normalize_person_name(name)
    s = re.sub(r"\s*\(\d+\)\s*$", "", str(s or "").strip())
    s = re.sub(r"\s+", " ", s).strip()
    return s.casefold()
def _coerce_week(value) -> pd.Timestamp | None:
    wk = pd.to_datetime(value, errors="coerce")
    if pd.isna(wk):
        return None
    return pd.Timestamp(wk).normalize()
def _names_from_non_wip_payload(payload: Any) -> set[str]:
    if payload is None:
        return set()
    if not isinstance(payload, (list, dict)) and pd.isna(payload):
        return set()
    obj = payload
    if isinstance(payload, str):
        try:
            obj = json.loads(payload)
        except Exception:
            parts = [p.strip() for p in re.split(r"[,;\n\r]+", payload) if p.strip()]
            return {_normalize_person_name(p) for p in parts if _normalize_person_name(p)}
    if isinstance(obj, dict):
        return {_normalize_person_name(k) for k in obj.keys() if _normalize_person_name(k)}
    if isinstance(obj, list):
        names = set()
        for item in obj:
            if isinstance(item, dict):
                name = item.get("name") or item.get("person") or item.get("Person")
            else:
                name = item
            norm = _normalize_person_name(name)
            if norm:
                names.add(norm)
        return names
    return set()
def _enterprise_people_count_from_sources(
    *,
    team: str,
    week,
    nw_frame: pd.DataFrame | None,
    person_hours: pd.DataFrame | None,
    people_in_wip: pd.DataFrame | None,
) -> float:
    wk = _coerce_week(week)
    if wk is None:
        return np.nan
    names: set[str] = set()
    if nw_frame is not None and not nw_frame.empty:
        raw_nw = nw_frame.copy()
        if "period_date" in raw_nw.columns:
            raw_nw["period_date"] = pd.to_datetime(raw_nw["period_date"], errors="coerce").dt.normalize()
            nw_match = raw_nw.loc[(raw_nw.get("team") == team) & (raw_nw["period_date"] == wk)]
            if "non_wip_by_person" in nw_match.columns:
                for payload in nw_match["non_wip_by_person"].dropna().tolist():
                    names.update(_names_from_non_wip_payload(payload))
    for df_, person_col in [(person_hours, "person"), (people_in_wip, "person")]:
        if df_ is None or df_.empty or person_col not in df_.columns:
            continue
        frame = df_.copy()
        if "period_date" not in frame.columns or "team" not in frame.columns:
            continue
        frame["period_date"] = pd.to_datetime(frame["period_date"], errors="coerce").dt.normalize()
        sub = frame.loc[(frame["team"] == team) & (frame["period_date"] == wk), person_col]
        names.update(_normalize_person_name(x) for x in sub.tolist() if _normalize_person_name(x))
    return float(len(names)) if names else np.nan
def _completed_hours_from_metrics(metrics_frame: pd.DataFrame | None, team: str, week) -> float:
    if metrics_frame is None or metrics_frame.empty:
        return np.nan
    if "team" not in metrics_frame.columns or "period_date" not in metrics_frame.columns:
        return np.nan
    completed_col = None
    for col in metrics_frame.columns:
        if str(col).strip().lower() in {"completed hours", "completed_hours", "wip_hours"}:
            completed_col = col
            break
    if completed_col is None:
        return np.nan
    wk = _coerce_week(week)
    if wk is None:
        return np.nan
    frame = metrics_frame.copy()
    frame["period_date"] = pd.to_datetime(frame["period_date"], errors="coerce").dt.normalize()
    vals = pd.to_numeric(
        frame.loc[(frame["team"].astype(str).str.strip() == team) & (frame["period_date"] == wk), completed_col],
        errors="coerce",
    ).fillna(0.0)
    return float(vals.sum()) if not vals.empty else np.nan
def _row_total_non_wip_hours(nw_row: pd.Series | None) -> float:
    if nw_row is None:
        return np.nan
    for col in ["non_wip_hours", "total_non_wip_hours", "total_non-wip_hours"]:
        if col in nw_row.index:
            val = _to_float(nw_row.get(col), np.nan)
            if pd.notna(val):
                return float(val)
    return np.nan
def _person_available_hours_for_week(
    person_hours: pd.DataFrame | None,
    team: str,
    week,
    person_key: str,
) -> float:
    if person_hours is None or person_hours.empty:
        return np.nan
    needed = {"team", "period_date", "person", "Available Hours"}
    if not needed.issubset(person_hours.columns):
        return np.nan
    wk = pd.to_datetime(week, errors="coerce")
    if pd.isna(wk):
        return np.nan
    ph = person_hours.copy()
    ph["period_date"] = pd.to_datetime(ph["period_date"], errors="coerce").dt.normalize()
    mask = (
        ph["team"].astype(str).str.strip().str.upper().eq(str(team).strip().upper())
        & ph["period_date"].eq(pd.Timestamp(wk).normalize())
        & ph["person"].map(_person_key).eq(_person_key(person_key))
    )
    vals = pd.to_numeric(
        ph.loc[mask, "Available Hours"],
        errors="coerce",
    ).dropna()
    return float(vals.sum()) if not vals.empty else np.nan
def enterprise_nonwip_kpi_lookup(
    *,
    team: str,
    week,
    nw_row: pd.Series,
    wk_people: pd.DataFrame,
    people_count: float | int | None,
    completed_hours: float,
    total_non_wip_hours: float,
    factor_out_ooo: bool = True,
    peter_available_hours: float | None = None,
    cpt_total_non_wip_hours: float | None = None,
    person_hours: pd.DataFrame | None = None,
    people_in_wip: pd.DataFrame | None = None,
    nw_frame: pd.DataFrame | None = None,
    metrics_frame: pd.DataFrame | None = None,
    ent_capacity_callback: Callable[..., float] | None = None,
    ent_capacity_kwargs: dict[str, Any] | None = None,
) -> dict[str, Any]:
    team_name = str(team or "").strip()
    wk_people = wk_people.copy() if wk_people is not None else pd.DataFrame()
    if not wk_people.empty:
        for col in ["Expected Hours", "OOO Hours", "Other Team WIP"]:
            if col in wk_people.columns:
                wk_people[col] = pd.to_numeric(wk_people[col], errors="coerce").fillna(0.0)
    row_people_count = _to_float(nw_row.get("people_count", np.nan), np.nan) if nw_row is not None else np.nan
    if team_name in ENTERPRISE_PEOPLE_COUNT_FROM_NW_TEAMS and pd.notna(row_people_count):
        count = float(row_people_count)
    else:
        source_count = _enterprise_people_count_from_sources(
            team=team_name,
            week=week,
            nw_frame=nw_frame,
            person_hours=person_hours,
            people_in_wip=people_in_wip,
        )
        if pd.notna(source_count):
            count = float(source_count)
        else:
            try:
                count = float(people_count) if people_count is not None and pd.notna(people_count) else 0.0
            except Exception:
                count = 0.0
    if count <= 0 and not wk_people.empty and "person" in wk_people.columns:
        count = float(wk_people["person"].astype(str).str.strip().replace("", pd.NA).dropna().nunique())
    if team_name in FORTY_HOUR_TEAMS:
        capacity_hours = _capacity_from_count_with_person_overrides(count, 40.0, wk_people)
    elif team_name in {"DS", "Lit & Letters"}:
        capacity_hours = _capacity_from_count_with_person_overrides(count, 37.5, wk_people)
    elif team_name == "CPT":
        capacity_hours = _sum_col(wk_people, "Available Hours")
        if capacity_hours <= 0.0:
            capacity_hours = _sum_col(wk_people, "Expected Hours")
    elif team_name in {"CDS", "NI"}:
        if peter_available_hours is None or pd.isna(peter_available_hours):
            peter_available_hours = _person_available_hours_for_week(
                person_hours=person_hours,
                team=team_name,
                week=week,
                person_key="peter mchugh",
            )
        peter_fallback_capacity = 10.0 if team_name == "CDS" else 27.75
        peter_capacity = (
            float(peter_available_hours)
            if peter_available_hours is not None
            and pd.notna(peter_available_hours)
            else peter_fallback_capacity
        )
        assigned_count = 1 if count > 0 else 0
        remaining_count = max(count - assigned_count, 0.0)
        capacity_hours = (assigned_count * peter_capacity) + (remaining_count * 37.75)
    elif team_name == "ENT" and ent_capacity_callback is not None:
        try:
            capacity_hours = float(ent_capacity_callback(**(ent_capacity_kwargs or {})))
        except Exception:
            capacity_hours = _sum_col(wk_people, "Expected Hours")
    else:
        capacity_hours = _sum_col(wk_people, "Expected Hours")
    ooo_hours = _sum_col(wk_people, "OOO Hours")
    other_team_wip_hours = _sum_col(wk_people, "Other Team WIP")
    row_total_non_wip_hours = _row_total_non_wip_hours(nw_row)
    source_total_non_wip_hours = total_non_wip_hours
    if team_name == "CPT" and cpt_total_non_wip_hours is not None:
        source_total_non_wip_hours = cpt_total_non_wip_hours
    elif pd.notna(row_total_non_wip_hours):
        source_total_non_wip_hours = row_total_non_wip_hours
    total_non_wip_hours = _to_float(source_total_non_wip_hours, 0.0)
    if team_name == "PH-NM MEIC":
        non_wip_hours = max(total_non_wip_hours - other_team_wip_hours - ooo_hours, 0.0)
    else:
        non_wip_hours = max(total_non_wip_hours - other_team_wip_hours, 0.0)
    metrics_completed_hours = _completed_hours_from_metrics(metrics_frame, team_name, week)
    completed_hours = _to_float(
        metrics_completed_hours if pd.notna(metrics_completed_hours) else completed_hours,
        0.0,
    )
    if completed_hours == 0.0 and not wk_people.empty and "Completed Hours" in wk_people.columns:
        completed_hours = _sum_col(wk_people, "Completed Hours")
    unaccounted_hours = max(
        capacity_hours - completed_hours - other_team_wip_hours - non_wip_hours - ooo_hours,
        0.0,
    )
    over_hours = max(
        completed_hours + other_team_wip_hours + non_wip_hours + ooo_hours - capacity_hours,
        0.0,
    )
    if factor_out_ooo:
        pct_denom = max(capacity_hours - ooo_hours, 0.0)
        ooo_pct = 0.0
    else:
        pct_denom = capacity_hours
        ooo_pct = (ooo_hours / pct_denom) if pct_denom > 0 else np.nan
    def pct(hours: float) -> float:
        return (hours / pct_denom) if pct_denom > 0 else np.nan
    return {
        "week_start": pd.to_datetime(week, errors="coerce"),
        "team": team_name,
        "people_count": count,
        "capacity_hours": capacity_hours,
        "pct_denom": pct_denom,
        "completed_hours": completed_hours,
        "other_team_wip_hours": other_team_wip_hours,
        "non_wip_hours": non_wip_hours,
        "ooo_hours": ooo_hours,
        "unaccounted_hours": unaccounted_hours,
        "over_hours": over_hours,
        "warning": f"Over {over_hours:.2f} hours" if over_hours > 0 else "",
        "wip_pct": pct(completed_hours),
        "other_team_wip_pct": pct(other_team_wip_hours),
        "non_wip_pct": pct(non_wip_hours),
        "ooo_pct": ooo_pct,
        "unaccounted_pct": pct(unaccounted_hours),
    }