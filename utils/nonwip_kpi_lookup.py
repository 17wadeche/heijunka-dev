from __future__ import annotations
from typing import Any, Callable
import numpy as np
import pandas as pd
FORTY_HOUR_TEAMS = {
    "SVT", "PVH", "NV", "Enabling Technologies", "DBS", "PH", "Spine",
    "PSS", "SCS", "TDD", "ACM", "VSS", "Endoscopy", "Surgical AST-GST",
    "PH-NM MEIC", "TCT",
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
    ent_capacity_callback: Callable[..., float] | None = None,
    ent_capacity_kwargs: dict[str, Any] | None = None,
) -> dict[str, Any]:
    team_name = str(team or "").strip()
    wk_people = wk_people.copy() if wk_people is not None else pd.DataFrame()
    if not wk_people.empty:
        for col in ["Expected Hours", "OOO Hours", "Other Team WIP"]:
            if col in wk_people.columns:
                wk_people[col] = pd.to_numeric(wk_people[col], errors="coerce").fillna(0.0)
    try:
        count = float(people_count) if people_count is not None and pd.notna(people_count) else 0.0
    except Exception:
        count = 0.0
    if count <= 0 and not wk_people.empty and "person" in wk_people.columns:
        count = float(wk_people["person"].astype(str).str.strip().replace("", pd.NA).dropna().nunique())
    if team_name in FORTY_HOUR_TEAMS:
        capacity_hours = count * 40.0
    elif team_name in {"DS", "Lit & Letters"}:
        capacity_hours = count * 37.5
    elif team_name == "CPT":
        cpt_31_count = 2
        cpt_30_2_count = 1
        cpt_37_5_count = 4
        assigned_count = cpt_31_count + cpt_30_2_count + cpt_37_5_count
        remaining_count = max(count - assigned_count, 0.0)
        capacity_hours = (
            (cpt_31_count * 31.0)
            + (cpt_30_2_count * 30.2)
            + (cpt_37_5_count * 37.5)
            + (remaining_count * 37.75)
        )
    elif team_name == "CDS":
        assigned_count = 1 if count > 0 else 0
        remaining_count = max(count - assigned_count, 0.0)
        capacity_hours = (assigned_count * 10.0) + (remaining_count * 37.75)
    elif team_name == "NI":
        assigned_count = 1
        remaining_count = max(count - assigned_count, 0.0)
        capacity_hours = (assigned_count * 27.75) + (remaining_count * 37.75)
    elif team_name == "ENT" and ent_capacity_callback is not None:
        try:
            capacity_hours = float(ent_capacity_callback(**(ent_capacity_kwargs or {})))
        except Exception:
            capacity_hours = _sum_col(wk_people, "Expected Hours")
    else:
        capacity_hours = _sum_col(wk_people, "Expected Hours")
    ooo_hours = _sum_col(wk_people, "OOO Hours")
    other_team_wip_hours = _sum_col(wk_people, "Other Team WIP")
    total_non_wip_hours = _to_float(total_non_wip_hours, 0.0)
    if team_name == "PH-NM MEIC":
        non_wip_hours = max(total_non_wip_hours - other_team_wip_hours - ooo_hours, 0.0)
    else:
        non_wip_hours = max(total_non_wip_hours - other_team_wip_hours, 0.0)
    completed_hours = _to_float(completed_hours, 0.0)
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
        "wip_pct": pct(completed_hours),
        "other_team_wip_pct": pct(other_team_wip_hours),
        "non_wip_pct": pct(non_wip_hours),
        "ooo_pct": ooo_pct,
        "unaccounted_pct": pct(unaccounted_hours),
    }