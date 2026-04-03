# pages/Enterprise.py
from __future__ import annotations
import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import pandas as pd
import streamlit as st
import numpy as np
import io
from utils.activity_map import ACTIVITY_MAP
def _candidate_repo_roots(start: Path) -> List[Path]:
    roots: List[Path] = []
    p = start.resolve()
    for parent in [p, *p.parents]:
        roots.append(parent)
    roots.extend(
        [
            Path("/mount/src/heijunka-dev"),
            Path("/mount/src/HEIJUNKA-DEV"),
            Path("/mount/src/heijunka-dev/..").resolve(),
            Path.cwd(),
        ]
    )
    seen = set()
    out: List[Path] = []
    for r in roots:
        try:
            rr = r.resolve()
        except Exception:
            rr = r
        if rr not in seen:
            seen.add(rr)
            out.append(rr)
    return out
NAME_ALIASES = {
    "mirlay": "Mirlay Morin",
    "nikita": "Nikita Schazenbach",
    "jacob": "Jacob Woolley",
    "madison": "Madison Moeller",
    "pavani uppari":"Uppari Pavani",
    "s, prabhu":"Prabhu S",
    "damahe, jagruti":"Jagruti Damahe",
    "kallagunta, malleshwari":"Malleshwari Kallagunta",
    "gopikalyani ijigiri":"Gopikalyani Iligiri",
    "dey, pranjal":"Pranjal Dey",
    "shanmugasundaram, naveen":"Naveen Shanmugasundaram",
    "shanmugasundaram, naveenkumar":"Naveen Shanmugasundaram",
    "s, giridhar":"Giridhar S",
    "surekha raju anantarapu":"Surekha Raju",
    "anwar, mohd faiz":"Mohd Faiz Anwar",
    "nath, koushik":"Koushik Nath",
    "iligiri, gopikalyani":"Gopikalyani Iligiri",
    "gowda, manjunath":"Manjunath Gowda",
    "andrew o":"Andrew",
    "kumar, shailesh":"Shailesh Kumar",
    "michael": "Michael F",
    "mani s.":"Mani",
    "kuche":"Ku Che",
    "goutham kumar, p":"P Goutham Kumar",
}
def normalize_person_name(name: str) -> str:
    s = str(name or "").strip()
    s = " ".join(s.split())
    key = s.lower()
    return NAME_ALIASES.get(key, s)
def find_org_config_path() -> Tuple[Optional[Path], List[Path]]:
    attempted: List[Path] = []
    start = Path(__file__).resolve()
    preferred_root = start.parents[1] if len(start.parents) >= 2 else start.parent
    preferred = preferred_root / "config" / "enterprise_org.json"
    attempted.append(preferred)
    if preferred.exists():
        return preferred, attempted
    for root in _candidate_repo_roots(start):
        cand = root / "config" / "enterprise_org.json"
        attempted.append(cand)
        if cand.exists():
            return cand, attempted
    for root in _candidate_repo_roots(start):
        config_dir = root / "config"
        if config_dir.is_dir():
            try:
                for f in config_dir.iterdir():
                    attempted.append(f)
                    if f.is_file() and f.name.lower() == "enterprise_org.json":
                        return f, attempted
            except Exception:
                pass
    return None, attempted
@st.cache_data(show_spinner=False)
def _build_export_lookup_tables_cached(
    metrics_df: Optional[pd.DataFrame],
    nonwip_df: Optional[pd.DataFrame],
    org,
    factor_out_ooo: bool,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    team_export = _weekly_team_export_df(
        metrics_df,
        nonwip_df,
        org,
        factor_out_ooo=factor_out_ooo,
    )
    if team_export is None or team_export.empty:
        empty = pd.DataFrame()
        return empty, empty, empty
    team_export = team_export.copy()
    today = pd.Timestamp.now().normalize()
    if "week_start" in team_export.columns:
        team_export["week_start"] = pd.to_datetime(
            team_export["week_start"], errors="coerce"
        ).dt.normalize()
        team_export = team_export[team_export["week_start"] <= today].copy()
    for col in ["completed_hours", "non_wip_hours", "ooo_hours", "unaccounted_hours"]:
        if col not in team_export.columns:
            team_export[col] = 0.0
    team_export = team_export[
        (
            pd.to_numeric(team_export["completed_hours"], errors="coerce").fillna(0.0)
            + pd.to_numeric(team_export["non_wip_hours"], errors="coerce").fillna(0.0)
        ) > 0
    ].reset_index(drop=True)
    if team_export.empty:
        empty = pd.DataFrame()
        return empty, empty, empty
    ou_export = _rollup_export_level(
        team_export,
        "ou",
        factor_out_ooo=factor_out_ooo,
    )
    portfolio_export = _rollup_export_level(
        team_export,
        "portfolio",
        factor_out_ooo=factor_out_ooo,
    )
    return team_export, ou_export, portfolio_export
@st.cache_data(show_spinner=False)
def _prepare_nonwip_activity_source(source_raw: pd.DataFrame) -> pd.DataFrame:
    if source_raw is None or source_raw.empty:
        return pd.DataFrame()
    source_df = _normalize_df_columns(source_raw.copy())
    dc = _get_date_col(source_df)
    json_col = _first_col(source_df, ["non_wip_activities", "non-wip_activities"])
    if not (dc and json_col):
        return pd.DataFrame()
    source_df[dc] = _safe_to_datetime(source_df, dc)
    source_df = source_df.dropna(subset=[dc]).sort_values(dc)
    rows: list[dict] = []
    for _, r in source_df.iterrows():
        wk = r[dc]
        payload = _loads_json_maybe(r[json_col])
        if not payload:
            continue
        if isinstance(payload, dict):
            payload = [payload]
        if not isinstance(payload, list):
            continue
        for item in payload:
            if not isinstance(item, dict):
                continue
            act = item.get("activity") or item.get("Activity") or item.get("type")
            hrs = item.get("hours") or item.get("Hours")
            if act is None or hrs is None:
                continue
            try:
                hrs_val = float(hrs)
            except Exception:
                hrs_val = 0.0
            rows.append(
                {
                    "week": wk,
                    "activity": str(act).strip(),
                    "hours": hrs_val,
                }
            )
    if not rows:
        return pd.DataFrame()
    act_df = pd.DataFrame(rows)
    act_df["week"] = pd.to_datetime(act_df["week"], errors="coerce")
    act_df = act_df.dropna(subset=["week"])
    act_df["week_start"] = _weekly_start(act_df["week"])
    return act_df
@st.cache_data(show_spinner=False)
def _cached_excel_bytes(
    team_export_display: pd.DataFrame,
    ou_export_display: pd.DataFrame,
    portfolio_export_display: pd.DataFrame,
) -> bytes:
    return _excel_bytes_from_export_dfs(
        team_export_display,
        ou_export_display,
        portfolio_export_display,
    )
@dataclass(frozen=True)
class TeamConfig:
    name: str
    enabled: bool = True
    meta: Dict[str, Any] = None  # extra fields like portfolio/ou/etc
@dataclass(frozen=True)
class OrgConfig:
    org_name: str
    teams: List[TeamConfig]
    raw: Dict[str, Any]
def _coerce_bool(v: Any, default: bool = True) -> bool:
    if v is None:
        return default
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return bool(v)
    if isinstance(v, str):
        return v.strip().lower() in {"1", "true", "yes", "y", "enabled", "on"}
    return default
def _add_avg_hours_day_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    pct_to_avg = {
        "wip_pct": "wip_avg_hours_day",
        "non_wip_pct": "non_wip_avg_hours_day",
        "ooo_pct": "ooo_avg_hours_day",
        "unaccounted_pct": "unaccounted_avg_hours_day",
    }
    for pct_col, avg_col in pct_to_avg.items():
        if pct_col in out.columns:
            out[avg_col] = pd.to_numeric(out[pct_col], errors="coerce") * 8.0
    return out
def _threshold_cell_style(val: Any, threshold: float, good_if_gte: bool = False) -> str:
    try:
        v = float(val)
    except Exception:
        return ""
    good = v >= threshold if good_if_gte else v < threshold
    if good:
        return "background-color: #d1fae5; color: #065f46;"
    return "background-color: #fee2e2; color: #991b1b;"
TEAMS_CONFIG_PATH = Path(__file__).resolve().parents[1] / "teams.json"
@st.cache_data(show_spinner=False)
def load_team_config(config_path: str | None = None) -> dict:
    p = Path(config_path) if config_path else TEAMS_CONFIG_PATH
    try:
        with open(p, "r", encoding="utf-8") as f:
            obj = json.load(f)
        return obj if isinstance(obj, dict) else {}
    except Exception:
        return {}
def irl_people_for_team(team: str, config: dict) -> set[str]:
    if not isinstance(config, dict):
        return set()
    team_cfg = config.get(str(team).strip(), {})
    if not isinstance(team_cfg, dict):
        return set()
    raw = team_cfg.get("irl_people", [])
    if not isinstance(raw, list):
        return set()
    return {str(x).strip() for x in raw if str(x).strip()}
def irl_people_for_team(team: str, config: dict) -> set[str]:
    if not isinstance(config, dict):
        return set()
    team_cfg = config.get(str(team).strip(), {})
    if not isinstance(team_cfg, dict):
        return set()
    raw = team_cfg.get("irl_people", [])
    if not isinstance(raw, list):
        return set()
    return {str(x).strip() for x in raw if str(x).strip()}
def explode_non_wip_by_person(nw: pd.DataFrame) -> pd.DataFrame:
    cols = ["team", "period_date", "person", "Non-WIP Hours"]
    if nw.empty or "non_wip_by_person" not in nw.columns:
        return pd.DataFrame(columns=cols)
    rows = []
    sub = nw[["team", "period_date", "non_wip_by_person"]].dropna(subset=["non_wip_by_person"]).copy()
    for _, r in sub.iterrows():
        wk = pd.to_datetime(r.get("period_date"), errors="coerce")
        if pd.isna(wk):
            continue
        wk = pd.Timestamp(wk).normalize()
        payload = r["non_wip_by_person"]
        try:
            obj = json.loads(payload) if isinstance(payload, str) else payload
            if not isinstance(obj, dict):
                continue
        except Exception:
            continue
        for person, hrs in obj.items():
            try:
                v = float(hrs)
            except Exception:
                v = np.nan
            person_name = normalize_person_name(str(person).strip())
            if not person_name:
                continue
            rows.append({
                "team": str(r["team"]).strip(),
                "period_date": wk,
                "person": person_name,
                "Non-WIP Hours": v,
            })
    out = pd.DataFrame(rows, columns=cols)
    if not out.empty:
        out["period_date"] = pd.to_datetime(out["period_date"], errors="coerce").dt.normalize()
    return out
def explode_person_hours(df: pd.DataFrame) -> pd.DataFrame:
    cols = ["team", "period_date", "person", "Actual Hours", "Available Hours", "Utilization"]
    if df is None or df.empty:
        return pd.DataFrame(columns=cols)
    temp = _normalize_df_columns(df.copy())
    if "person_hours" not in temp.columns:
        return pd.DataFrame(columns=cols)
    if "team" not in temp.columns or "period_date" not in temp.columns:
        return pd.DataFrame(columns=cols)
    BAD_NAMES = {"", "-", "–", "—", "nan", "NaN", "NAN", "n/a", "N/A", "na", "NA", "null", "NULL", "none", "None"}
    def _is_good_name(s: str) -> bool:
        t = str(s).strip()
        return t and t not in BAD_NAMES and t.lower() not in {b.lower() for b in BAD_NAMES}
    rows: list[dict] = []
    sub = temp.loc[:, ["team", "period_date", "person_hours"]].dropna(subset=["person_hours"]).copy()
    for _, r in sub.iterrows():
        wk = pd.to_datetime(r.get("period_date"), errors="coerce")
        if pd.isna(wk):
            continue
        wk = pd.Timestamp(wk).normalize()
        payload = r["person_hours"]
        try:
            obj = json.loads(payload) if isinstance(payload, str) else payload
            if not isinstance(obj, dict):
                continue
        except Exception:
            continue
        for person, vals in obj.items():
            if not _is_good_name(person):
                continue
            vals = vals if isinstance(vals, dict) else {}
            a = pd.to_numeric(vals.get("actual"), errors="coerce")
            t = pd.to_numeric(vals.get("available"), errors="coerce")
            a = float(a) if pd.notna(a) else 0.0
            t = float(t) if pd.notna(t) else 0.0
            if a == 0.0 and t == 0.0:
                continue
            util = (a / t) if t not in (0, 0.0) else np.nan
            person_name = normalize_person_name(str(person).strip())
            if not person_name:
                continue
            rows.append({
                "team": str(r["team"]).strip(),
                "period_date": wk,
                "person": person_name,
                "Actual Hours": a,
                "Available Hours": t,
                "Utilization": util,
            })
    out = pd.DataFrame(rows, columns=cols)
    if not out.empty:
        out["period_date"] = pd.to_datetime(out["period_date"], errors="coerce").dt.normalize()
    return out
def explode_people_in_wip(df: pd.DataFrame) -> pd.DataFrame:
    cols = ["team", "period_date", "person"]
    if df is None or df.empty:
        return pd.DataFrame(columns=cols)
    temp = _normalize_df_columns(df.copy())
    if "people_in_wip" not in temp.columns:
        return pd.DataFrame(columns=cols)
    if "team" not in temp.columns or "period_date" not in temp.columns:
        return pd.DataFrame(columns=cols)
    BAD_NAMES = {"", "-", "–", "—", "nan", "NaN", "NAN", "n/a", "N/A", "na", "NA", "null", "NULL", "none", "None"}
    def _is_good_name(s: str) -> bool:
        return str(s).strip() and str(s).strip() not in BAD_NAMES
    def _as_names(x) -> list[str]:
        if isinstance(x, list):
            return [normalize_person_name(str(s).strip()) for s in x if _is_good_name(str(s))]
        if isinstance(x, dict):
            return [normalize_person_name(str(k).strip()) for k in x.keys() if _is_good_name(str(k))]
        if isinstance(x, str):
            s = x.strip()
            try:
                obj = json.loads(s)
                if isinstance(obj, list):
                    return [normalize_person_name(str(v).strip()) for v in obj if _is_good_name(str(v))]
                if isinstance(obj, dict):
                    return [normalize_person_name(str(k).strip()) for k in obj.keys() if _is_good_name(str(k))]
            except Exception:
                pass
            parts = [p.strip() for p in re.split(r"[,;\n\r]+", s) if _is_good_name(p)]
            return [normalize_person_name(p) for p in parts]
        return []
    rows: list[dict] = []
    sub = temp.loc[:, ["team", "period_date", "people_in_wip"]].dropna(subset=["people_in_wip"]).copy()
    for _, r in sub.iterrows():
        wk = pd.to_datetime(r.get("period_date"), errors="coerce")
        if pd.isna(wk):
            continue
        wk = pd.Timestamp(wk).normalize()
        people = _as_names(r["people_in_wip"])
        for person in people:
            if not person:
                continue
            rows.append({
                "team": str(r["team"]).strip(),
                "period_date": wk,
                "person": person,
            })
    out = pd.DataFrame(rows, columns=cols)
    if not out.empty:
        out["period_date"] = pd.to_datetime(out["period_date"], errors="coerce").dt.normalize()
        out = out.drop_duplicates(subset=["team", "period_date", "person"])
    return out
def merged_people_count_for_week(
    team: str,
    week,
    nw_frame: pd.DataFrame,
    person_hours: pd.DataFrame,
    people_in_wip: pd.DataFrame,
) -> int:
    wk = pd.to_datetime(week, errors="coerce").normalize()
    if nw_frame is not None and not nw_frame.empty and team in {"ENT", "DBS", "NV", "Enabling Technologies", "Spine", "PH", "SCS", "TDD", "ACM"}:
        raw_nw = nw_frame.copy()
        if "period_date" in raw_nw.columns:
            raw_nw["period_date"] = pd.to_datetime(raw_nw["period_date"], errors="coerce").dt.normalize()
        if "people_count" in raw_nw.columns:
            team_match = raw_nw.loc[
                (raw_nw["team"] == team) & (raw_nw["period_date"] == wk),
                "people_count"
            ]
            team_match = pd.to_numeric(team_match, errors="coerce").dropna()
            if not team_match.empty:
                return int(team_match.iloc[0])
    names = set()
    long_nw = explode_non_wip_by_person(nw_frame)
    for df_, person_col in [
        (long_nw, "person"),
        (person_hours, "person"),
        (people_in_wip, "person"),
    ]:
        if df_ is None or df_.empty:
            continue
        sub = df_.loc[
            (df_["team"] == team) & (df_["period_date"] == wk),
            [person_col]
        ].copy()
        if not sub.empty:
            vals = sub[person_col].astype(str).map(normalize_person_name).str.strip()
            names.update(x for x in vals if x)
    return len(names)
def accounted_nonwip_by_person_from_row(row) -> tuple[dict[str, float], dict[str, float]]:
    payload = row.get("non_wip_activities", "[]")
    try:
        activities = json.loads(payload) if isinstance(payload, str) else payload
    except Exception:
        activities = []
    if not isinstance(activities, list) or not activities:
        return {}, {}
    import re
    accounted_other: dict[str, float] = {}
    accounted_nonother: dict[str, float] = {}
    for d in activities:
        name = normalize_person_name(str(d.get("name", "")).strip())
        if not name:
            continue
        act_raw = str(d.get("activity", ""))
        act_key = re.sub(r"[^A-Z0-9]", "", act_raw.upper())  # normalize
        if act_key == "OOO":
            continue
        try:
            hrs = float(d.get("hours", 0) or 0)
        except Exception:
            hrs = 0.0
        if hrs <= 0:
            continue
        else:
            accounted_nonother[name] = accounted_nonother.get(name, 0.0) + hrs
    accounted_other = {k: round(v, 2) for k, v in accounted_other.items()}
    accounted_nonother = {k: round(v, 2) for k, v in accounted_nonother.items()}
    return accounted_other, accounted_nonother
def build_person_weekly_accounting(
    team: str,
    week,
    nw_row,
    long_nw: pd.DataFrame,
    person_hours: pd.DataFrame,
    week_hours: float = 40.0,
    irl_people: set[str] | None = None,
) -> pd.DataFrame:
    wk = pd.to_datetime(week, errors="coerce").normalize()
    nw_people = long_nw.loc[
        (long_nw["team"] == team) & (long_nw["period_date"] == wk),
        ["person", "Non-WIP Hours"]
    ].copy()
    if nw_people.empty:
        nw_people = pd.DataFrame(columns=["person", "Non-WIP Hours"])
    nw_people["person"] = nw_people["person"].astype(str).str.strip()
    nw_people["Non-WIP Hours"] = pd.to_numeric(nw_people["Non-WIP Hours"], errors="coerce").fillna(0.0)
    wip_people = person_hours.loc[
        (person_hours["team"] == team) & (person_hours["period_date"] == wk),
        ["person", "Actual Hours"]
    ].copy()
    if wip_people.empty:
        wip_people = pd.DataFrame(columns=["person", "Actual Hours"])
    wip_people["person"] = wip_people["person"].astype(str).str.strip()
    wip_people["Completed Hours"] = pd.to_numeric(wip_people["Actual Hours"], errors="coerce").fillna(0.0)
    wip_people = wip_people.drop(columns=["Actual Hours"], errors="ignore")
    acct_other_map, acct_nonother_map = accounted_nonwip_by_person_from_row(nw_row)
    other_df = pd.DataFrame(
        [{"person": str(k).strip(), "Other Team WIP": float(v)} for k, v in acct_other_map.items()]
    )
    if other_df.empty:
        other_df = pd.DataFrame(columns=["person", "Other Team WIP"])
    acct_df = pd.DataFrame(
        [{"person": str(k).strip(), "Accounted Non-WIP": float(v)} for k, v in acct_nonother_map.items()]
    )
    if acct_df.empty:
        acct_df = pd.DataFrame(columns=["person", "Accounted Non-WIP"])
    payload = nw_row.get("non_wip_activities", "[]")
    try:
        activities = json.loads(payload) if isinstance(payload, str) else payload
    except Exception:
        activities = []
    ooo_by_person: dict[str, float] = {}
    if isinstance(activities, list):
        for item in activities:
            if not isinstance(item, dict):
                continue
            person = normalize_person_name(str(item.get("name", "")).strip())
            activity = str(item.get("activity", "")).strip().upper()
            try:
                hrs = float(item.get("hours", 0) or 0)
            except Exception:
                hrs = 0.0
            if not person or hrs <= 0:
                continue
            if activity == "OOO":
                ooo_by_person[person] = ooo_by_person.get(person, 0.0) + hrs
    ooo_df = pd.DataFrame(
        [{"person": k, "OOO Hours": round(v, 2)} for k, v in ooo_by_person.items()]
    )
    if ooo_df.empty:
        ooo_df = pd.DataFrame(columns=["person", "OOO Hours"])
    def _clean_person_col(df_in: pd.DataFrame, value_col: str) -> pd.DataFrame:
        if df_in.empty:
            return pd.DataFrame(columns=["person", value_col])
        out = df_in.copy()
        out["person"] = (
            out["person"]
            .astype("string")
            .fillna("")
            .map(lambda x: normalize_person_name(str(x).strip()))
        )
        out["person"] = out["person"].replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
        out = out.dropna(subset=["person"]).copy()
        out[value_col] = pd.to_numeric(out[value_col], errors="coerce").fillna(0.0)
        return out[["person", value_col]]
    nw_people = _clean_person_col(nw_people, "Non-WIP Hours")
    wip_people = _clean_person_col(wip_people, "Completed Hours")
    other_df = _clean_person_col(other_df, "Other Team WIP")
    acct_df = _clean_person_col(acct_df, "Accounted Non-WIP")
    ooo_df = _clean_person_col(ooo_df, "OOO Hours")
    nw_people = nw_people.groupby("person", as_index=False)["Non-WIP Hours"].sum()
    wip_people = wip_people.groupby("person", as_index=False)["Completed Hours"].sum()
    other_df = other_df.groupby("person", as_index=False)["Other Team WIP"].sum()
    acct_df = acct_df.groupby("person", as_index=False)["Accounted Non-WIP"].sum()
    ooo_df = ooo_df.groupby("person", as_index=False)["OOO Hours"].sum()
    all_people = sorted(
        set(nw_people["person"].tolist())
        | set(wip_people["person"].tolist())
        | set(other_df["person"].tolist())
        | set(acct_df["person"].tolist())
        | set(ooo_df["person"].tolist())
    )
    people = pd.DataFrame({"person": pd.Series(all_people, dtype="string")})
    out = (
        people
        .merge(nw_people.astype({"person": "string"}), on="person", how="left")
        .merge(wip_people.astype({"person": "string"}), on="person", how="left")
        .merge(other_df.astype({"person": "string"}), on="person", how="left")
        .merge(acct_df.astype({"person": "string"}), on="person", how="left")
        .merge(ooo_df.astype({"person": "string"}), on="person", how="left")
    )
    fill_zero_cols = [
        "Non-WIP Hours",
        "Completed Hours",
        "Other Team WIP",
        "Accounted Non-WIP",
        "OOO Hours",
    ]
    for c in fill_zero_cols:
        if c not in out.columns:
            out[c] = 0.0
        out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0.0)
    out["person_key"] = out["person"].astype(str).str.strip().str.lower()
    irl_people_norm = {str(x).strip().lower() for x in (irl_people or set())}
    out["Expected Hours"] = np.where(
        out["person_key"].isin(irl_people_norm),
        39.0,
        float(week_hours),
    )
    out["OOO Hours"] = pd.to_numeric(out["OOO Hours"], errors="coerce").fillna(0.0)
    out["Non-WIP Hours"] = pd.to_numeric(out["Non-WIP Hours"], errors="coerce").fillna(0.0)
    out["Completed Hours"] = pd.to_numeric(out["Completed Hours"], errors="coerce").fillna(0.0)
    out["Other Team WIP"] = pd.to_numeric(out["Other Team WIP"], errors="coerce").fillna(0.0)
    out["Accounted Non-WIP"] = pd.to_numeric(out["Accounted Non-WIP"], errors="coerce").fillna(0.0)
    non_ooo_total = out["Non-WIP Hours"].clip(lower=0.0)
    out["Other Team WIP"] = np.minimum(out["Other Team WIP"], non_ooo_total)
    remaining_nonwip = (non_ooo_total - out["Other Team WIP"]).clip(lower=0.0)
    out["Accounted Non-WIP"] = np.minimum(out["Accounted Non-WIP"], remaining_nonwip)
    out["Unaccounted"] = (
        out["Expected Hours"]
        - out["Completed Hours"]
        - out["OOO Hours"]
        - out["Other Team WIP"]
        - out["Accounted Non-WIP"]
    ).clip(lower=0.0)
    out["Total Used"] = (
        out["Completed Hours"]
        + out["OOO Hours"]
        + out["Other Team WIP"]
        + out["Accounted Non-WIP"]
    )
    out["period_date"] = wk
    out["team"] = team
    return out.sort_values(["person"]).reset_index(drop=True)
def _selected_nonwip_start_floor(df: Optional[pd.DataFrame]) -> Optional[pd.Timestamp]:
    if df is None or df.empty:
        return None
    team_col = _get_team_col(df)
    date_col = _get_date_col(df)
    if not team_col or not date_col:
        return None
    tmp = df.copy()
    tmp[date_col] = pd.to_datetime(tmp[date_col], errors="coerce")
    tmp = tmp.dropna(subset=[team_col, date_col])
    if tmp.empty:
        return None
    tmp = tmp[tmp[team_col].astype(str).isin(set(team_filter))]
    if tmp.empty:
        return None
    per_team_min = (
        tmp.groupby(team_col)[date_col]
        .min()
        .dropna()
    )
    if per_team_min.empty:
        return None
    return pd.to_datetime(per_team_min.max())
def parse_org_config(data: Dict[str, Any]) -> OrgConfig:
    org_name = (
        data.get("org_name")
        or (data.get("org") or {}).get("name")
        or data.get("name")
        or "Enterprise"
    )
    teams_raw = data.get("teams") or data.get("Teams") or []
    teams: List[TeamConfig] = []
    if isinstance(teams_raw, list):
        for t in teams_raw:
            if isinstance(t, str):
                teams.append(TeamConfig(name=t, enabled=True, meta={}))
            elif isinstance(t, dict):
                name = t.get("name") or t.get("team") or t.get("Team")
                if not name:
                    continue
                enabled = _coerce_bool(t.get("enabled"), True)
                meta = {
                    k: v
                    for k, v in t.items()
                    if k not in {"name", "team", "Team", "enabled"}
                }
                teams.append(TeamConfig(name=str(name), enabled=enabled, meta=meta))
    elif isinstance(teams_raw, dict):
        for name, tmeta in teams_raw.items():
            if isinstance(tmeta, dict):
                enabled = _coerce_bool(tmeta.get("enabled"), True)
                meta = {k: v for k, v in tmeta.items() if k != "enabled"}
            else:
                enabled, meta = True, {}
            teams.append(TeamConfig(name=str(name), enabled=enabled, meta=meta))
    return OrgConfig(org_name=str(org_name), teams=teams, raw=data)
def load_org_config() -> Tuple[Optional[OrgConfig], Optional[str], List[str], Optional[str]]:
    cfg_path, attempted = find_org_config_path()
    attempted_str = [str(p) for p in attempted]
    if cfg_path is None:
        return None, None, attempted_str, None
    try:
        text = cfg_path.read_text(encoding="utf-8")
        data = json.loads(text)
        org = parse_org_config(data)
        if not org.teams:
            return (
                None,
                f"Found config at:\n{cfg_path}\n\n…but it has no teams. "
                "Add a `teams` list (strings or objects with `name`).",
                attempted_str,
                str(cfg_path),
            )
        return org, None, attempted_str, str(cfg_path)
    except Exception as e:
        return None, f"Failed to read/parse config:\n{cfg_path}\n\n{e}", attempted_str, str(cfg_path)
def _repo_root_from_cfg_path_str(cfg_path_str: Optional[str]) -> Path:
    if cfg_path_str:
        p = Path(cfg_path_str)
        if p.exists():
            if p.parent.name.lower() == "config":
                return p.parent.parent
            return p.parent
    return Path(__file__).resolve().parents[1]
def _try_read_csv(path: Path) -> Optional[pd.DataFrame]:
    try:
        if path.exists() and path.is_file():
            return pd.read_csv(path)
    except Exception:
        return None
    return None
@st.cache_data(show_spinner=False)
def load_common_data(repo_root_str: str) -> Dict[str, pd.DataFrame]:
    repo_root = Path(repo_root_str)
    candidates = {
        "metrics": repo_root / "metrics.csv",
        "metrics_aggregate_dev": repo_root / "metrics_aggregate_dev.csv",
        "non_wip": repo_root / "non_wip.csv",
        "non_wip_activities": repo_root / "non_wip_activities.csv",
        "closures": repo_root / "closures.csv",
        "timeliness": repo_root / "timeliness.csv",
        "Timeliness": repo_root / "Timeliness.csv",
        "NS_WIP": repo_root / "NS_WIP.csv",
        "ns_non_wip_activities": repo_root / "ns_non_wip_activities.csv",
        "CRM_WIP": repo_root / "CRM_WIP.csv",
        "crm_non_wip_activities": repo_root / "crm_non_wip_activities.csv",
        "MS_WIP": repo_root / "MS_WIP.csv",
        "ms_non_wip_activities": repo_root / "ms_non_wip_activities.csv",
    }
    out: Dict[str, pd.DataFrame] = {}
    for key, p in candidates.items():
        df = _try_read_csv(p)
        if df is not None and not df.empty:
            out[key] = df
    return out
def _maybe_apply_styles():
    try:
        from utils.styles import apply_global_styles  # type: ignore
        apply_global_styles()
    except Exception:
        pass
def _norm(s: str) -> str:
    return str(s).strip().lower().replace(" ", "_")
def _first_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    cols = {_norm(c): c for c in df.columns}
    for cand in candidates:
        if cand in cols:
            return cols[cand]
    return None
def _get_team_col(df: pd.DataFrame) -> Optional[str]:
    return _first_col(df, ["team", "team_name", "org_team", "squad"])
def _get_date_col(df: pd.DataFrame) -> Optional[str]:
    return _first_col(df, ["week", "period_date", "date", "day", "as_of", "timestamp"])
def _safe_to_datetime(df: pd.DataFrame, col: str) -> pd.Series:
    return pd.to_datetime(df[col], errors="coerce")
def _weekly_start(s: pd.Series) -> pd.Series:
    return s.dt.to_period("W-MON").dt.start_time
def _loads_json_maybe(v: Any) -> Any:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (dict, list)):
        return v
    if isinstance(v, str):
        t = v.strip()
        if not t:
            return None
        try:
            return json.loads(t)
        except Exception:
            return None
    return None
def _canon_activity(label: str) -> str:
    s_orig = str(label or "").strip()
    if not s_orig:
        return s_orig
    s = re.sub(r"\s+", " ", s_orig).strip()
    s = re.sub(r"^[\.\,\;\:\-\–\—\s]+", "", s).strip()
    s = re.sub(r"[:\-\–\—]\s*\d+\s*$", "", s).strip()
    if not s:
        return s
    lower = s.lower()
    explicit_map = ACTIVITY_MAP
    if lower in explicit_map:
        return explicit_map[lower]
    acronym_tokens = {
        "im", "wip", "ooo", "sla", "qa", "hc", "pe", "wfh", "pto",
        "ri", "capa",
    }
    words = lower.split(" ")
    if len(words) == 1:
        w = words[0]
        if w.endswith("s") and not w.endswith("ss") and len(w) > 3:
            w = w[:-1]
        if w in acronym_tokens:
            return w.upper()
        return w.capitalize()
    last = words[-1]
    if last.endswith("s") and not last.endswith("ss") and len(last) > 3:
        words[-1] = last[:-1]
    pretty = []
    for w in words:
        if not w:
            continue
        if w in acronym_tokens:
            pretty.append(w.upper())
        else:
            pretty.append(w.capitalize())
    return " ".join(pretty)
def _normalize_df_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    norm_to_first: Dict[str, str] = {}
    rename: Dict[str, str] = {}
    for col in df.columns:
        n = _norm(col)
        if n not in norm_to_first:
            norm_to_first[n] = col
            rename[col] = n
        else:
            primary = norm_to_first[n]
            df[primary] = df[primary].where(
                df[primary].notna() & (df[primary].astype(str).str.strip() != ""),
                df[col]
            )
            rename[col] = None  # mark for drop
    drop_cols = [c for c, v in rename.items() if v is None]
    df = df.drop(columns=drop_cols)
    df = df.rename(columns={c: v for c, v in rename.items() if v is not None})
    return df
def split_nonwip_activity_minutes(cat: pd.DataFrame) -> pd.DataFrame:
    import numpy as np
    if cat.empty:
        return cat
    rows: list[dict] = []
    for _, r in cat.iterrows():
        activity_text = str(r["Activity"])
        total_hours_raw = pd.to_numeric(r["Hours"], errors="coerce")
        total_hours = float(total_hours_raw) if pd.notna(total_hours_raw) else 0.0
        s = activity_text.replace(";", " ").replace(",", " ").replace(":", " ")
        s = re.sub(r"\s+", " ", s).strip()
        if not s:
            rows.append({"Activity": _canon_activity(activity_text), "Hours": total_hours})
            continue
        pattern = re.compile(
            r"(?P<num>\d+)\s*(?P<unit>h|hr|hrs|hour|hours|m|min|mins|minute|minutes)?\b",
            re.IGNORECASE,
        )
        sub_acts: list[tuple[str, int]] = []
        prev_end = 0
        for m in pattern.finditer(s):
            num = int(m.group("num"))
            unit = (m.group("unit") or "").lower()
            mins = num * 60 if unit in ("h", "hr", "hrs", "hour", "hours") else num
            label = s[prev_end:m.start()]
            prev_end = m.end()
            label = label.strip()
            if not label:
                continue
            label = re.sub(r"\([^)]*$", "", label)
            label = re.sub(r"\(.*?\)", "", label)
            label = re.sub(r"[:\-–—]+$", "", label)
            label = label.strip(" ,;:()[]-–—")
            label = re.sub(r"\s+", " ", label).strip()
            label = _canon_activity(label)
            if label and mins > 0:
                sub_acts.append((label, mins))
        if sub_acts:
            has_delims = bool(re.search(r"[;,]", activity_text))
            if len(sub_acts) == 1 and not has_delims:
                rows.append({"Activity": _canon_activity(activity_text), "Hours": total_hours})
                continue
            total_minutes = sum(m for _, m in sub_acts)
            if total_hours <= 0 and total_minutes > 0:
                total_hours = total_minutes / 60.0
            if total_hours > 0 and total_minutes > 0:
                for label, mins in sub_acts:
                    h_sub = total_hours * (mins / total_minutes)
                    rows.append({"Activity": label, "Hours": h_sub})
            else:
                rows.append({"Activity": _canon_activity(activity_text), "Hours": total_hours})
        else:
            rows.append({"Activity": _canon_activity(activity_text), "Hours": total_hours})
    import numpy as np
    out = pd.DataFrame(rows)
    if out.empty:
        return cat
    out["Activity"] = out["Activity"].map(_canon_activity)
    return out.groupby("Activity", as_index=False)["Hours"].sum()
st.set_page_config(page_title="Enterprise Dashboard", layout="wide")
_maybe_apply_styles()
st.title("Enterprise Dashboard")
org, org_err, attempted_paths, cfg_path_str = load_org_config()
if org is None:
    st.error("No org config found.")
    expected_example = "/mount/src/heijunka-dev/config/enterprise_org.json"
    st.caption(
        "Expected file at: "
        f"`{expected_example}`\n\n"
        "Create `config/enterprise_org.json` and add teams.\n\n"
        "Tip: On Linux, folder names are case-sensitive. If your repo folder is `HEIJUNKA-DEV`, "
        "hard-coded paths to `heijunka-dev` will fail."
    )
    if org_err:
        st.code(org_err)
    with st.expander("Paths checked (debug)", expanded=False):
        st.write("\n".join(attempted_paths[:300]))
    st.stop()
repo_root = _repo_root_from_cfg_path_str(cfg_path_str)
data = load_common_data(str(repo_root))
enabled_teams = [t for t in org.teams if t.enabled]
all_team_names = [t.name for t in org.teams]
enabled_team_names = [t.name for t in enabled_teams] or all_team_names
with st.sidebar:
    st.subheader(org.org_name)
    all_portfolios = sorted(
        {
            str((t.meta or {}).get("portfolio")).strip()
            for t in org.teams
            if (t.meta or {}).get("portfolio") is not None
        }
    )
    portfolio_filter = st.multiselect(
        "Portfolio",
        options=all_portfolios,
        default=all_portfolios,
    )
    teams_after_portfolio = (
        [
            t
            for t in org.teams
            if str((t.meta or {}).get("portfolio")).strip() in set(portfolio_filter)
        ]
        if portfolio_filter
        else []
    )
    all_ous = sorted(
        {
            str((t.meta or {}).get("ou")).strip()
            for t in teams_after_portfolio
            if (t.meta or {}).get("ou") is not None
        }
    )
    ou_filter = st.multiselect(
        "OU",
        options=all_ous,
        default=all_ous,
    )
    teams_after_ou = (
        [
            t
            for t in teams_after_portfolio
            if str((t.meta or {}).get("ou")).strip() in set(ou_filter)
        ]
        if ou_filter
        else []
    )
    team_key = "enterprise_team_filter"
    team_options = [t.name for t in teams_after_ou]
    default_teams = [t for t in enabled_team_names if t in team_options] or team_options
    prev_options_key = "enterprise_prev_team_options"
    prev_team_options = st.session_state.get(prev_options_key)
    if prev_team_options != team_options:
        st.session_state[team_key] = default_teams
        st.session_state[prev_options_key] = team_options
    elif team_key not in st.session_state:
        st.session_state[team_key] = default_teams
    team_filter = st.multiselect(
        "Teams",
        options=team_options,
        key=team_key,
    )
def filter_by_team(df: pd.DataFrame) -> pd.DataFrame:
    if not team_filter:
        return df.iloc[0:0]
    team_cols = [
        c
        for c in df.columns
        if c.strip().lower() in {"team", "team_name", "squad", "org_team"}
    ]
    if not team_cols:
        tc = _get_team_col(df)
        if tc:
            team_cols = [tc]
    if not team_cols:
        return df
    col = team_cols[0]
    return df[df[col].astype(str).isin(set(team_filter))]
def _date_bounds_for_df(df: pd.DataFrame) -> tuple[Optional[pd.Timestamp], Optional[pd.Timestamp], Optional[str]]:
    dc = _get_date_col(df)
    if not dc:
        return None, None, None
    ser = pd.to_datetime(df[dc], errors="coerce").dropna()
    if ser.empty:
        return None, None, dc
    return ser.min(), ser.max(), dc
def section_date_range(
    label: str,
    df: Optional[pd.DataFrame],
    key: str,
    min_floor_ts: Optional[pd.Timestamp] = None,
    allow_future_dates: bool = False,
) -> tuple[Optional[pd.Timestamp], Optional[pd.Timestamp]]:
    if df is None or df.empty:
        return None, None
    mn, mx, _ = _date_bounds_for_df(df)
    if mn is None or mx is None:
        st.info("No date column detected for this section.")
        return None, None
    import datetime
    min_d = mn.date()
    if min_floor_ts is not None and pd.notna(min_floor_ts):
        min_d = max(min_d, pd.to_datetime(min_floor_ts).date())
    max_d = mx.date()
    today_d = datetime.date.today()
    if allow_future_dates:
        max_selectable = max_d
        preset_anchor_end = max_d
    else:
        max_selectable = today_d
        if max_selectable < min_d:
            max_selectable = min_d
            st.warning(
                f"{label}: all available data starts after today ({min_d}). "
                "Date range has been clamped to the first available date."
            )
        preset_anchor_end = today_d
    presets = [
        "Custom",
        "Past week",
        "Past month",
        "Past 3 months",
        "Past 6 months",
        "Past year",
        "Past 2 years",
    ]
    days_map = {
        "Past week": 7,
        "Past month": 30,
        "Past 3 months": 90,
        "Past 6 months": 180,
        "Past year": 365,
        "Past 2 years": 730,
    }
    preset_key = f"{key}_preset"
    dates_key = f"{key}_dates"
    last_preset_key = f"{key}_last_preset"
    preset = st.selectbox(
        f"{label} — quick range",
        options=presets,
        index=0,
        key=preset_key,
        help="Choose a preset, or pick Custom to select exact dates.",
    )
    if preset in days_map:
        end_default = preset_anchor_end
        start_default = max(
            min_d,
            (pd.to_datetime(preset_anchor_end) - pd.Timedelta(days=days_map[preset])).date(),
        )
    else:
        start_default = min_d
        end_default = max_selectable
    prev = st.session_state.get(last_preset_key)
    if prev != preset:
        st.session_state[dates_key] = (start_default, end_default)
        st.session_state[last_preset_key] = preset
        st.rerun()
    if dates_key in st.session_state:
        v = st.session_state[dates_key]
        if isinstance(v, (tuple, list)):
            vals = list(v)
            if len(vals) >= 2:
                s, e = vals[0], vals[1]
                if hasattr(s, "date"):
                    s = s.date()
                if hasattr(e, "date"):
                    e = e.date()
                if s is None:
                    s = min_d
                if e is None:
                    e = s
                s = min(max(s, min_d), max_selectable)
                e = min(max(e, min_d), max_selectable)
                if e < s:
                    e = s
                st.session_state[dates_key] = (s, e)
            elif len(vals) == 1:
                d = vals[0]
                if hasattr(d, "date"):
                    d = d.date()
                if d is not None:
                    d = min(max(d, min_d), max_selectable)
                    st.session_state[dates_key] = (d,)
            else:
                st.session_state[dates_key] = (start_default, end_default)
        else:
            d = v.date() if hasattr(v, "date") else v
            if d is not None:
                d = min(max(d, min_d), max_selectable)
                st.session_state[dates_key] = d
    dr = st.date_input(
        label,
        min_value=min_d,
        max_value=max_selectable,
        key=dates_key,
        help="Filters only this section.",
    )
    if isinstance(dr, tuple):
        if len(dr) == 2:
            start_d, end_d = dr
        elif len(dr) == 1:
            start_d, end_d = dr[0], dr[0]
        else:
            start_d, end_d = start_default, end_default
    else:
        start_d, end_d = dr, dr
    start_ts = pd.to_datetime(start_d)
    end_ts = pd.to_datetime(end_d) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    return start_ts, end_ts
def filter_by_date_range(df: pd.DataFrame, start_ts: Optional[pd.Timestamp], end_ts: Optional[pd.Timestamp]) -> pd.DataFrame:
    if start_ts is None or end_ts is None:
        return df
    dc = _get_date_col(df)
    if not dc:
        return df
    tmp = df.copy()
    tmp[dc] = pd.to_datetime(tmp[dc], errors="coerce")
    tmp = tmp.dropna(subset=[dc])
    return tmp[(tmp[dc] >= start_ts) & (tmp[dc] <= end_ts)]
def filter_df(df: pd.DataFrame, start_ts: Optional[pd.Timestamp], end_ts: Optional[pd.Timestamp]) -> pd.DataFrame:
    return filter_by_date_range(filter_by_team(df), start_ts, end_ts)
st.markdown(f"**Selected teams:** {len(team_filter)}")
if not team_filter:
    st.warning("No teams selected.")
    st.stop()
selected_nonwip_floor = None
for key in ["ns_non_wip_activities", "crm_non_wip_activities", "ms_non_wip_activities", "non_wip", "non_wip_activities"]:
    if key in data:
        floor_candidate = _selected_nonwip_start_floor(filter_by_team(data[key]))
        if floor_candidate is not None:
            selected_nonwip_floor = floor_candidate
            break
def _team_meta_lookup(org: OrgConfig) -> pd.DataFrame:
    rows = []
    for t in org.teams:
        meta = t.meta or {}
        rows.append({
            "team": t.name,
            "ou": meta.get("ou"),
            "portfolio": meta.get("portfolio"),
        })
    return pd.DataFrame(rows)
def _prepare_weekly_accounting_inputs(
    metrics_frame: pd.DataFrame,
    nw_frame: pd.DataFrame,
) -> dict[str, pd.DataFrame]:
    long_nw = explode_non_wip_by_person(nw_frame)
    person_hours = explode_person_hours(metrics_frame)
    people_in_wip = explode_people_in_wip(metrics_frame)
    if not long_nw.empty:
        long_nw["period_date"] = pd.to_datetime(long_nw["period_date"], errors="coerce").dt.normalize()
    if not person_hours.empty:
        person_hours["period_date"] = pd.to_datetime(person_hours["period_date"], errors="coerce").dt.normalize()
    if not people_in_wip.empty:
        people_in_wip["period_date"] = pd.to_datetime(people_in_wip["period_date"], errors="coerce").dt.normalize()
    return {
        "long_nw": long_nw,
        "person_hours": person_hours,
        "people_in_wip": people_in_wip,
    }
def ent_capacity_hours_for_week(
    team: str,
    week,
    nw_frame: pd.DataFrame,
    irl_people: set[str] | None = None,
) -> float:
    wk = pd.to_datetime(week, errors="coerce").normalize()
    irl_people_norm = {str(x).strip().lower() for x in (irl_people or set())}
    if nw_frame is None or nw_frame.empty:
        return 0.0
    raw_nw = nw_frame.copy()
    raw_nw["period_date"] = pd.to_datetime(raw_nw["period_date"], errors="coerce").dt.normalize()
    row = raw_nw.loc[
        (raw_nw["team"] == team) & (raw_nw["period_date"] == wk)
    ]
    if row.empty:
        return 0.0
    people_count_series = pd.to_numeric(row["people_count"], errors="coerce").dropna()
    people_count = int(people_count_series.iloc[0]) if not people_count_series.empty else 0
    irl_count = 0
    if "non_wip_by_person" in row.columns:
        payload = row.iloc[0].get("non_wip_by_person")
        try:
            obj = json.loads(payload) if isinstance(payload, str) else payload
        except Exception:
            obj = {}
        if isinstance(obj, dict):
            names = {
                normalize_person_name(str(k).strip()).strip().lower()
                for k in obj.keys()
                if str(k).strip()
            }
            irl_count = sum(1 for n in names if n in irl_people_norm)
    irl_count = min(irl_count, people_count)
    non_irl_count = max(people_count - irl_count, 0)
    return float((irl_count * 39.0) + (non_irl_count * 40.0))
def _weekly_team_export_df(
    dfm: Optional[pd.DataFrame],
    dfnw: Optional[pd.DataFrame],
    org: OrgConfig,
    factor_out_ooo: bool = False,
) -> pd.DataFrame:
    if dfnw is None or dfnw.empty:
        return pd.DataFrame()
    teams_cfg = load_team_config()
    meta = _team_meta_lookup(org)
    nw = _normalize_df_columns(dfnw.copy())
    if "period_date" not in nw.columns:
        dc = _get_date_col(nw)
        if dc is None:
            return pd.DataFrame()
        nw["period_date"] = pd.to_datetime(nw[dc], errors="coerce").dt.normalize()
    else:
        nw["period_date"] = pd.to_datetime(nw["period_date"], errors="coerce").dt.normalize()
    if "team" not in nw.columns:
        tc = _get_team_col(nw)
        if tc is None:
            return pd.DataFrame()
        nw["team"] = nw[tc].astype(str).str.strip()
    else:
        nw["team"] = nw["team"].astype(str).str.strip()
    nw = nw.dropna(subset=["period_date"])
    enabled_team_names = {t.name for t in org.teams if t.enabled} or {t.name for t in org.teams}
    nw = nw[nw["team"].isin(enabled_team_names)].copy()
    if nw.empty:
        return pd.DataFrame()
    metrics_frame = _normalize_df_columns(dfm.copy()) if dfm is not None and not dfm.empty else pd.DataFrame()
    prepared = _prepare_weekly_accounting_inputs(metrics_frame, nw)
    long_nw = prepared["long_nw"]
    person_hours = prepared["person_hours"]
    people_in_wip = prepared["people_in_wip"]
    metrics_team = pd.DataFrame(columns=["team", "week_start", "completed_hours"])
    if not metrics_frame.empty:
        if "period_date" not in metrics_frame.columns:
            dc = _get_date_col(metrics_frame)
            if dc is not None:
                metrics_frame["period_date"] = pd.to_datetime(metrics_frame[dc], errors="coerce").dt.normalize()
        else:
            metrics_frame["period_date"] = pd.to_datetime(metrics_frame["period_date"], errors="coerce").dt.normalize()
        if "team" not in metrics_frame.columns:
            tc = _get_team_col(metrics_frame)
            if tc is not None:
                metrics_frame["team"] = metrics_frame[tc].astype(str).str.strip()
        else:
            metrics_frame["team"] = metrics_frame["team"].astype(str).str.strip()
        completed_col = _first_col(metrics_frame, ["completed_hours", "completed hours", "wip_hours"])
        if completed_col and "team" in metrics_frame.columns and "period_date" in metrics_frame.columns:
            m = metrics_frame.dropna(subset=["team", "period_date"]).copy()
            m = m[m["team"].isin(enabled_team_names)].copy()
            m["week_start"] = pd.to_datetime(m["period_date"], errors="coerce").dt.normalize()
            m["completed_hours"] = pd.to_numeric(m[completed_col], errors="coerce").fillna(0.0)
            metrics_team = (
                m.groupby(["team", "week_start"], as_index=False)
                .agg(completed_hours=("completed_hours", "sum"))
            )
    nw["week_start"] = nw["period_date"]
    total_non_wip_col = _first_col(nw, ["total_non_wip_hours", "total_non-wip_hours"])
    if total_non_wip_col is None:
        nw["non_wip_hours"] = 0.0
    else:
        nw["non_wip_hours"] = pd.to_numeric(nw[total_non_wip_col], errors="coerce").fillna(0.0)
    rows: list[dict[str, Any]] = []
    for _, nw_row in nw.iterrows():
        team = str(nw_row.get("team", "")).strip()
        wk = pd.to_datetime(nw_row.get("week_start"), errors="coerce")
        if not team or pd.isna(wk):
            continue
        wk = pd.Timestamp(wk).normalize()
        team_irl_people = irl_people_for_team(team, teams_cfg)
        wk_people = build_person_weekly_accounting(
            team=team,
            week=wk,
            nw_row=nw_row,
            long_nw=long_nw,
            person_hours=person_hours,
            week_hours=40.0,
            irl_people=team_irl_people,
        )
        if wk_people.empty:
            continue
        wk_people = wk_people.copy()
        wk_people["Expected Hours"] = pd.to_numeric(wk_people["Expected Hours"], errors="coerce").fillna(0.0)
        wk_people["OOO Hours"] = pd.to_numeric(wk_people["OOO Hours"], errors="coerce").fillna(0.0)
        people_count = merged_people_count_for_week(
            team=team,
            week=wk,
            nw_frame=nw,
            person_hours=person_hours,
            people_in_wip=people_in_wip,
        )
        if people_count is None or float(people_count) <= 0:
            people_count = float(
                wk_people["person"].astype(str).str.strip().replace("", pd.NA).dropna().nunique()
            )
        if team in {"NV", "Enabling Technologies", "DBS", "PH", "Spine", "PSS", "SCS", "TDD","ACM"}:
            capacity_hours = float(people_count) * 40.0
        elif team == "ENT":
            capacity_hours = ent_capacity_hours_for_week(
                team=team,
                week=wk,
                nw_frame=nw,
            )
        else:
            capacity_hours = float(wk_people["Expected Hours"].sum())
        ooo_hours = float(wk_people["OOO Hours"].sum())
        non_wip_hours = float(pd.to_numeric(nw_row.get("non_wip_hours", 0.0), errors="coerce") or 0.0)
        completed_match = metrics_team[
            (metrics_team["team"] == team) &
            (pd.to_datetime(metrics_team["week_start"], errors="coerce").dt.normalize() == wk)
        ]
        completed_hours = float(completed_match["completed_hours"].sum()) if not completed_match.empty else 0.0
        unaccounted_hours = max(
            capacity_hours - completed_hours - non_wip_hours - ooo_hours,
            0.0,
        )
        if factor_out_ooo:
            pct_denom = max(capacity_hours - ooo_hours, 0.0)
            ooo_pct = 0.0
        else:
            pct_denom = capacity_hours
            ooo_pct = (ooo_hours / pct_denom) if pct_denom > 0 else pd.NA
        rows.append({
            "team": team,
            "week_start": wk,
            "people_count": float(people_count),
            "completed_hours": completed_hours,
            "non_wip_hours": non_wip_hours,
            "ooo_hours": ooo_hours,
            "capacity_hours": capacity_hours,
            "unaccounted_hours": unaccounted_hours,
            "wip_pct": (completed_hours / pct_denom) if pct_denom > 0 else pd.NA,
            "non_wip_pct": (non_wip_hours / pct_denom) if pct_denom > 0 else pd.NA,
            "ooo_pct": ooo_pct,
            "unaccounted_pct": (unaccounted_hours / pct_denom) if pct_denom > 0 else pd.NA,
        })
    if not rows:
        return pd.DataFrame()
    base = pd.DataFrame(rows)
    base = (
        base.groupby(["team", "week_start"], as_index=False)
        .agg(
            people_count=("people_count", "max"),
            completed_hours=("completed_hours", "sum"),
            non_wip_hours=("non_wip_hours", "sum"),
            ooo_hours=("ooo_hours", "sum"),
            capacity_hours=("capacity_hours", "sum"),
            unaccounted_hours=("unaccounted_hours", "sum"),
        )
    )
    if factor_out_ooo:
        pct_denom = (base["capacity_hours"] - base["ooo_hours"]).clip(lower=0.0)
        base["ooo_pct"] = 0.0
    else:
        pct_denom = base["capacity_hours"]
        base["ooo_pct"] = (base["ooo_hours"] / pct_denom).where(pct_denom > 0)
    base["wip_pct"] = (base["completed_hours"] / pct_denom).where(pct_denom > 0)
    base["non_wip_pct"] = (base["non_wip_hours"] / pct_denom).where(pct_denom > 0)
    base["unaccounted_pct"] = (base["unaccounted_hours"] / pct_denom).where(pct_denom > 0)
    base = base.merge(meta, on="team", how="left")
    base = _add_avg_hours_day_columns(base)
    return base.sort_values(["week_start", "portfolio", "ou", "team"]).reset_index(drop=True)
def _rollup_export_level(df: pd.DataFrame, level: str, factor_out_ooo: bool = False) -> pd.DataFrame:
    if df.empty:
        return df.copy()
    if level == "ou":
        group_cols = ["week_start", "portfolio", "ou"]
    elif level == "portfolio":
        group_cols = ["week_start", "portfolio"]
    else:
        raise ValueError("level must be 'ou' or 'portfolio'")
    out = (
        df.groupby(group_cols, as_index=False)
        .agg(
            people_count=("people_count", "sum"),
            completed_hours=("completed_hours", "sum"),
            non_wip_hours=("non_wip_hours", "sum"),
            ooo_hours=("ooo_hours", "sum"),
            capacity_hours=("capacity_hours", "sum"),
            unaccounted_hours=("unaccounted_hours", "sum"),
        )
    )
    if factor_out_ooo:
        pct_denom = (out["capacity_hours"] - out["ooo_hours"]).clip(lower=0.0)
        out["ooo_pct"] = 0.0
    else:
        pct_denom = out["capacity_hours"]
        out["ooo_pct"] = (out["ooo_hours"] / pct_denom).where(pct_denom > 0)
    out["wip_pct"] = (out["completed_hours"] / pct_denom).where(pct_denom > 0)
    out["non_wip_pct"] = (out["non_wip_hours"] / pct_denom).where(pct_denom > 0)
    out["unaccounted_pct"] = (out["unaccounted_hours"] / pct_denom).where(pct_denom > 0)
    out = _add_avg_hours_day_columns(out)
    if level == "portfolio":
        out["ou"] = pd.NA
    cols = [
        "week_start",
        "portfolio",
        "ou",
        "people_count",
        "completed_hours",
        "wip_pct",
        "wip_avg_hours_day",
        "non_wip_hours",
        "non_wip_pct",
        "non_wip_avg_hours_day",
        "ooo_hours",
        "ooo_pct",
        "ooo_avg_hours_day",
        "capacity_hours",
        "unaccounted_hours",
        "unaccounted_pct",
        "unaccounted_avg_hours_day",
    ]
    cols = [c for c in cols if c in out.columns]
    return out[cols].sort_values(group_cols).reset_index(drop=True)
def _display_export_team_df(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {
        "portfolio": "Portfolio",
        "ou": "OU",
        "team": "Team",
        "week_start": "Week Start",
        "completed_hours": "Completed Hours",
        "people_count": "People Count",
        "non_wip_hours": "Non-WIP Hours",
        "ooo_hours": "OOO Hours",
        "capacity_hours": "Capacity Hours",
        "unaccounted_hours": "Unaccounted Hours",
        "wip_pct": "WIP %",
        "wip_avg_hours_day": "WIP Avg. Hours/Day",
        "non_wip_pct": "Non-WIP %",
        "non_wip_avg_hours_day": "Non-WIP Avg. Hours/Day",
        "ooo_pct": "OOO %",
        "ooo_avg_hours_day": "OOO Avg. Hours/Day",
        "unaccounted_pct": "Unaccounted %",
        "unaccounted_avg_hours_day": "Unaccounted Avg. Hours/Day",
    }
    preferred_order = [
        "portfolio", "ou", "team", "week_start",
        "capacity_hours", "people_count",
        "completed_hours", "wip_pct", "wip_avg_hours_day",
        "non_wip_hours", "non_wip_pct", "non_wip_avg_hours_day",
        "ooo_hours", "ooo_pct", "ooo_avg_hours_day",
        "unaccounted_hours", "unaccounted_pct", "unaccounted_avg_hours_day",
    ]
    cols = [c for c in preferred_order if c in df.columns] + [c for c in df.columns if c not in preferred_order]
    out = df[cols].copy().rename(columns=rename_map)
    if "Week Start" in out.columns:
        out["Week Start"] = pd.to_datetime(out["Week Start"], errors="coerce").dt.date
    return out
def _display_export_ou_df(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {
        "portfolio": "Portfolio",
        "ou": "OU",
        "week_start": "Week Start",
        "completed_hours": "Completed Hours",
        "people_count": "People Count",
        "non_wip_hours": "Non-WIP Hours",
        "ooo_hours": "OOO Hours",
        "capacity_hours": "Capacity Hours",
        "unaccounted_hours": "Unaccounted Hours",
        "wip_pct": "WIP %",
        "wip_avg_hours_day": "WIP Avg. Hours/Day",
        "non_wip_pct": "Non-WIP %",
        "non_wip_avg_hours_day": "Non-WIP Avg. Hours/Day",
        "ooo_pct": "OOO %",
        "ooo_avg_hours_day": "OOO Avg. Hours/Day",
        "unaccounted_pct": "Unaccounted %",
        "unaccounted_avg_hours_day": "Unaccounted Avg. Hours/Day",
    }
    preferred_order = [
        "portfolio", "ou", "week_start",
        "capacity_hours", "people_count",
        "completed_hours", "wip_pct", "wip_avg_hours_day",
        "non_wip_hours", "non_wip_pct", "non_wip_avg_hours_day",
        "ooo_hours", "ooo_pct", "ooo_avg_hours_day",
        "unaccounted_hours", "unaccounted_pct", "unaccounted_avg_hours_day",
    ]
    cols = [c for c in preferred_order if c in df.columns] + [c for c in df.columns if c not in preferred_order]
    out = df[cols].copy().rename(columns=rename_map)
    if "Week Start" in out.columns:
        out["Week Start"] = pd.to_datetime(out["Week Start"], errors="coerce").dt.date
    return out
def _display_export_portfolio_df(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {
        "portfolio": "Portfolio",
        "week_start": "Week Start",
        "completed_hours": "Completed Hours",
        "people_count": "People Count",
        "non_wip_hours": "Non-WIP Hours",
        "ooo_hours": "OOO Hours",
        "capacity_hours": "Capacity Hours",
        "unaccounted_hours": "Unaccounted Hours",
        "wip_pct": "WIP %",
        "wip_avg_hours_day": "WIP Avg. Hours/Day",
        "non_wip_pct": "Non-WIP %",
        "non_wip_avg_hours_day": "Non-WIP Avg. Hours/Day",
        "ooo_pct": "OOO %",
        "ooo_avg_hours_day": "OOO Avg. Hours/Day",
        "unaccounted_pct": "Unaccounted %",
        "unaccounted_avg_hours_day": "Unaccounted Avg. Hours/Day",
    }
    preferred_order = [
        "portfolio", "week_start",
        "capacity_hours", "people_count",
        "completed_hours", "wip_pct", "wip_avg_hours_day",
        "non_wip_hours", "non_wip_pct", "non_wip_avg_hours_day",
        "ooo_hours", "ooo_pct", "ooo_avg_hours_day",
        "unaccounted_hours", "unaccounted_pct", "unaccounted_avg_hours_day",
    ]
    cols = [c for c in preferred_order if c in df.columns] + [
        c for c in df.columns if c not in preferred_order and c != "ou"
    ]
    out = df[cols].copy().rename(columns=rename_map)
    if "Week Start" in out.columns:
        out["Week Start"] = pd.to_datetime(out["Week Start"], errors="coerce").dt.date
    return out
def _excel_bytes_from_export_dfs(
    team_df: pd.DataFrame,
    ou_df: pd.DataFrame,
    portfolio_df: pd.DataFrame,
) -> bytes:
    last_err = None
    for engine in ("openpyxl", "xlsxwriter"):
        buf = io.BytesIO()
        try:
            with pd.ExcelWriter(buf, engine=engine) as writer:
                team_df.to_excel(writer, index=False, sheet_name="Team Weekly")
                ou_df.to_excel(writer, index=False, sheet_name="OU Weekly")
                portfolio_df.to_excel(writer, index=False, sheet_name="Portfolio Weekly")
            buf.seek(0)
            return buf.getvalue()
        except Exception as e:
            last_err = e
    raise RuntimeError(
        f"Excel export requires openpyxl or xlsxwriter to be installed. Last error: {last_err}"
    )
def _concat_frames(frames: list[pd.DataFrame]) -> Optional[pd.DataFrame]:
    if not frames:
        return None
    if len(frames) == 1:
        return frames[0]
    return pd.concat(frames, ignore_index=True, sort=False)
metrics_frames = []
for key in ["metrics", "metrics_aggregate_dev", "NS_WIP", "CRM_WIP", "MS_WIP"]:
    if key in data:
        d = data[key].copy()
        if not d.empty:
            metrics_frames.append(d)
nonwip_frames = []
for key in ["ns_non_wip_activities", "ms_non_wip_activities", "crm_non_wip_activities", "non_wip_activities", "non_wip"]:
    if key in data:
        d = _normalize_df_columns(data[key].copy())
        if not d.empty:
            nonwip_frames.append(d)
shared_metrics_df = _concat_frames(metrics_frames)
shared_nonwip_df = _concat_frames(nonwip_frames)
tabs = st.tabs(["Overview", "Non-WIP", "Export"])
@st.cache_data(show_spinner=False)
def _get_export_lookup_bundle(
    shared_metrics_df: Optional[pd.DataFrame],
    shared_nonwip_df: Optional[pd.DataFrame],
    org,
    factor_out_ooo: bool,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    return _build_export_lookup_tables_cached(
        shared_metrics_df,
        shared_nonwip_df,
        org,
        factor_out_ooo=factor_out_ooo,
    )
with tabs[0]:
    st.subheader("Summary")
    overview_factor_out_ooo = st.toggle(
        "Factor out OOO from overview calculations",
        value=False,
        key="overview_factor_out_ooo",
        help="When on, OOO is removed from the denominator for overview percentages, OOO Hours/OOO % are shown as 0, and Unaccounted is recalculated against capacity excluding OOO.",
    )
    overview_team_export, overview_ou_export, overview_portfolio_export = _get_export_lookup_bundle(
        shared_metrics_df,
        shared_nonwip_df,
        org,
        overview_factor_out_ooo,
    )
    team_lookup = overview_team_export
    ou_lookup = overview_ou_export
    portfolio_lookup = overview_portfolio_export
    if team_lookup.empty:
        st.info("No overview data available.")
    else:
        control_cols = st.columns([1.25, 1.0, 1.25])
        week_options = sorted(
            team_lookup["week_start"].dropna().unique(),
            reverse=True,
        )
        selected_week = control_cols[0].selectbox(
            "Week",
            options=week_options,
            index=0,
            format_func=lambda x: pd.Timestamp(x).strftime("%Y-%m-%d"),
            key="overview_selected_week",
        )
        filter_level = control_cols[1].radio(
            "Filter by",
            options=["Portfolio", "OU", "Team"],
            index=0,
            horizontal=True,
            key="overview_filter_level",
        )
        if filter_level == "Portfolio":
            lookup_df = portfolio_lookup.copy()
            filter_col = "portfolio"
            label = "Portfolio"
        elif filter_level == "OU":
            lookup_df = ou_lookup.copy()
            filter_col = "ou"
            label = "OU"
        else:
            lookup_df = team_lookup.copy()
            filter_col = "team"
            label = "Team"
        lookup_df["week_start"] = pd.to_datetime(lookup_df["week_start"], errors="coerce").dt.normalize()
        scoped_week = lookup_df[
            lookup_df["week_start"] == pd.Timestamp(selected_week).normalize()
        ].copy()
        options = sorted(
            x for x in scoped_week[filter_col].dropna().astype(str).unique()
            if str(x).strip()
        )
        if not options:
            st.info(f"No {label} values available for the selected week.")
        else:
            selected_value = control_cols[2].selectbox(
                label,
                options=options,
                index=0,
                key=f"overview_selected_{filter_col}",
            )
            scoped_df = scoped_week[
                scoped_week[filter_col].astype(str) == str(selected_value)
            ].copy()
            row = scoped_df.iloc[0] if len(scoped_df) == 1 else None
            def _safe_metric(v, pct: bool = False):
                if pd.isna(v):
                    return "—"
                return f"{float(v):.1%}" if pct else f"{float(v):.2f}"
            st.markdown("""
            <style>
            div[data-testid="stMetric"]{ text-align: center; }
            label[data-testid="stMetricLabel"]{ display: block; width: 100%; text-align: center; margin: 0; }
            label[data-testid="stMetricLabel"] p{ text-align: center !important; margin: 0 !important; }
            div[data-testid="stMetricValue"]{ text-align: center !important; width: 100%; }
            </style>
            """, unsafe_allow_html=True)
            _, c1, c2, _ = st.columns([1.2, 1.2, 1.2, 1.2])
            c1.metric("Avg Per Person **WIP** Daily Hours", _safe_metric(scoped_df["wip_avg_hours_day"].iloc[0]))
            c2.metric("Avg Per Person **Non-WIP** Daily Hours", _safe_metric(scoped_df["non_wip_avg_hours_day"].iloc[0]))
            _, p1, p2, _ = st.columns([1.2, 1.2, 1.2, 1.2])
            p1.metric("**WIP** Ratio", _safe_metric(scoped_df["wip_pct"].iloc[0], pct=True))
            p2.metric("**Non-WIP** Ratio", _safe_metric(scoped_df["non_wip_pct"].iloc[0], pct=True))
            st.divider()
            _, _, c3, c4, _, _, _ = st.columns([1.35, 1.2, 1.2, 1.2, 1.2, 1.0, 0.5])
            c3.metric("Avg **OOO** Weekly Hours", _safe_metric(scoped_df["ooo_hours"].iloc[0]))
            c4.metric("Avg **Unaccounted** Weekly Hours", _safe_metric(scoped_df["unaccounted_hours"].iloc[0]))
            _, _, p3, p4, _, _, _ = st.columns([1.35, 1.2, 1.2, 1.2, 1.2, 1.0, 0.5])
            p3.metric("**OOO** % of week", _safe_metric(scoped_df["ooo_pct"].iloc[0], pct=True))
            p4.metric("**Unaccounted** % remaining", _safe_metric(scoped_df["unaccounted_pct"].iloc[0], pct=True))
            st.divider()
            st.subheader("Selected rows")
            st.dataframe(scoped_df, use_container_width=True, hide_index=True)
EXCLUDED_NON_WIP = {"ooo", "non-wip", "non_wip", "other", "other team wip"}
def _norm_activity_name(val: Any) -> str:
    return str(val).strip().lower().replace("_", "-")
with tabs[1]:
    st.markdown("### Non-WIP activities")
    activity_keys = [
        "ns_non_wip_activities",
        "ms_non_wip_activities",
        "crm_non_wip_activities",
        "non_wip_activities",
    ]
    available_frames = []
    for key in activity_keys:
        if key in data:
            cand = filter_by_team(data[key])
            if not cand.empty:
                available_frames.append(_normalize_df_columns(cand.copy()))
    if not available_frames:
        st.info("No non-WIP activity CSVs found.")
        st.stop()
    source_raw = pd.concat(available_frames, ignore_index=True, sort=False).drop_duplicates()
    parsed_nonwip = _prepare_nonwip_activity_source(source_raw)
    top_n = st.number_input(
        "Number of activities to show",
        min_value=1,
        max_value=50,
        value=15,
        step=1,
        key="nonwip_top_n",
    )
    raw_dc = _get_date_col(source_raw)
    if not raw_dc:
        st.info("No date column found for Non-WIP activity data.")
        st.stop()
    source_raw = source_raw.copy()
    source_raw[raw_dc] = pd.to_datetime(source_raw[raw_dc], errors="coerce")
    source_raw = source_raw.dropna(subset=[raw_dc])
    if source_raw.empty:
        st.info("No dated Non-WIP activity data available.")
        st.stop()
    nonwip_min_d = source_raw[raw_dc].min().date()
    nonwip_max_d = source_raw[raw_dc].max().date()
    nw_dates = st.date_input(
        "Non-WIP date range",
        value=(nonwip_min_d, nonwip_max_d),
        min_value=nonwip_min_d,
        max_value=nonwip_max_d,
        key="dr_nonwip_dates_only",
        help="Filters only this section.",
    )
    if isinstance(nw_dates, tuple):
        if len(nw_dates) == 2:
            nw_start_d, nw_end_d = nw_dates
        elif len(nw_dates) == 1:
            nw_start_d, nw_end_d = nw_dates[0], nw_dates[0]
        else:
            nw_start_d, nw_end_d = nonwip_min_d, nonwip_max_d
    else:
        nw_start_d, nw_end_d = nw_dates, nw_dates
    nw_start = pd.to_datetime(nw_start_d)
    nw_end = pd.to_datetime(nw_end_d) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    source_df = filter_by_date_range(source_raw, nw_start, nw_end)
    if source_df.empty:
        st.info("No Non-WIP activity data available in this date range.")
        st.stop()
    source_df = _normalize_df_columns(source_df.copy())
    dc = _get_date_col(source_df)
    json_col = _first_col(source_df, ["non_wip_activities", "non-wip_activities"])
    if not (dc and json_col):
        st.info("Need `Week/period_date` and `Non-WIP Activities` (JSON list) to roll up activities.")
        st.stop()
    tmp = source_df.copy()
    tmp[dc] = _safe_to_datetime(tmp, dc)
    tmp = tmp.dropna(subset=[dc]).sort_values(dc)
    rows: List[Dict[str, Any]] = []
    for _, r in tmp.iterrows():
        wk = r[dc]
        payload = _loads_json_maybe(r[json_col])
        if not payload:
            continue
        if isinstance(payload, dict):
            payload = [payload]
        if not isinstance(payload, list):
            continue
        for item in payload:
            if not isinstance(item, dict):
                continue
            act = item.get("activity") or item.get("Activity") or item.get("type")
            hrs = item.get("hours") or item.get("Hours")
            if act is None or hrs is None:
                continue
            try:
                hrs_val = float(hrs)
            except Exception:
                hrs_val = 0.0
            rows.append(
                {
                    "week": wk,
                    "activity": str(act).strip(),
                    "hours": hrs_val,
                }
            )
    if not rows:
        st.info("No parsable activity rows found in the JSON column.")
        st.stop()
    act_df = pd.DataFrame(rows)
    act_df["week"] = pd.to_datetime(act_df["week"], errors="coerce")
    act_df = act_df.dropna(subset=["week"])
    act_df["week_start"] = _weekly_start(act_df["week"])
    weekly_raw = (
        act_df.groupby(["week_start", "activity"], as_index=False)
        .agg(hours=("hours", "sum"))
    )
    normalised_chunks: List[pd.DataFrame] = []
    for wk_val, grp in weekly_raw.groupby("week_start"):
        cat = grp[["activity", "hours"]].rename(columns={"activity": "Activity", "hours": "Hours"})
        cat_norm = split_nonwip_activity_minutes(cat)
        cat_norm["week_start"] = wk_val
        normalised_chunks.append(cat_norm)
    if not normalised_chunks:
        st.info("No activity data available after normalization.")
        st.stop()
    rolled = pd.concat(normalised_chunks, ignore_index=True)
    rolled = rolled.rename(columns={"Activity": "activity", "Hours": "hours"})
    rolled["activity_norm"] = rolled["activity"].map(_norm_activity_name)
    rolled = rolled[~rolled["activity_norm"].isin(EXCLUDED_NON_WIP)].copy()
    if rolled.empty:
        st.info('No activity data available after excluding "OOO" and "Non-WIP".')
        st.stop()
    weekly_by_activity = (
        rolled.groupby(["week_start", "activity"], as_index=False)
        .agg(hours=("hours", "sum"))
        .sort_values(["week_start", "hours"], ascending=[True, False])
    )
    total_hours = (
        weekly_by_activity.groupby("activity", as_index=False)
        .agg(total_hours=("hours", "sum"))
        .sort_values("total_hours", ascending=False)
        .head(int(top_n))
        .reset_index(drop=True)
    )
    if total_hours.empty:
        st.info("No chartable Non-WIP activity data available after exclusions.")
        st.stop()
    import matplotlib.pyplot as plt
    def _short_label(s: Any, max_len: int = 22) -> str:
        s = str(s).strip()
        return s if len(s) <= max_len else s[: max_len - 3] + "..."
    chart_df = total_hours.copy()
    chart_df["label"] = chart_df["activity"].map(lambda x: _short_label(x, 22))
    fig, ax = plt.subplots(figsize=(14, 5.5))
    bars = ax.bar(chart_df["label"], chart_df["total_hours"])
    ax.set_ylabel("Total Hours")
    ax.set_xlabel("Non-WIP Activity")
    ax.set_title(f"Top {int(top_n)} Non-WIP Activities by Total Hours")
    ax.tick_params(axis="x", rotation=45, labelsize=9)
    plt.setp(ax.get_xticklabels(), ha="right")
    for bar, val in zip(bars, chart_df["total_hours"]):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height(),
            f"{val:.1f}",
            ha="center",
            va="bottom",
            fontsize=9,
        )
    fig.tight_layout()
    st.pyplot(fig)
    st.caption(
        f"Top {int(top_n)} activities by total hours for the selected period, "
        "sorted highest to lowest from left to right."
    )
    st.divider()
    st.markdown("#### Activity breakdown — pie chart")
    pie_dates = st.date_input(
        "Pie chart date range",
        value=(nonwip_min_d, nonwip_max_d),
        min_value=nonwip_min_d,
        max_value=nonwip_max_d,
        key="dr_nonwip_pie_dates_only",
        help="Filters only the pie chart.",
    )
    if isinstance(pie_dates, tuple):
        if len(pie_dates) == 2:
            pie_start_d, pie_end_d = pie_dates
        elif len(pie_dates) == 1:
            pie_start_d, pie_end_d = pie_dates[0], pie_dates[0]
        else:
            pie_start_d, pie_end_d = nonwip_min_d, nonwip_max_d
    else:
        pie_start_d, pie_end_d = pie_dates, pie_dates
    pie_start = pd.to_datetime(pie_start_d)
    pie_end = pd.to_datetime(pie_end_d) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    pie_source_df = filter_by_date_range(source_raw, pie_start, pie_end)
    if pie_source_df.empty:
        st.info("No Non-WIP activity data in the selected pie chart date range.")
        st.stop()
    pie_source_df = _normalize_df_columns(pie_source_df.copy())
    pie_dc = _get_date_col(pie_source_df)
    pie_json_col = _first_col(pie_source_df, ["non_wip_activities", "non-wip_activities"])
    pie_rows: List[Dict[str, Any]] = []
    if pie_dc and pie_json_col:
        pie_tmp = pie_source_df.copy()
        pie_tmp[pie_dc] = _safe_to_datetime(pie_tmp, pie_dc)
        pie_tmp = pie_tmp.dropna(subset=[pie_dc]).sort_values(pie_dc)
        for _, r in pie_tmp.iterrows():
            payload = _loads_json_maybe(r[pie_json_col])
            if not payload:
                continue
            if isinstance(payload, dict):
                payload = [payload]
            if not isinstance(payload, list):
                continue
            for item in payload:
                if not isinstance(item, dict):
                    continue
                act = item.get("activity") or item.get("Activity") or item.get("type")
                hrs = item.get("hours") or item.get("Hours")
                if act is None or hrs is None:
                    continue
                try:
                    hrs_val = float(hrs)
                except Exception:
                    hrs_val = 0.0
                pie_rows.append(
                    {
                        "activity": str(act).strip(),
                        "hours": hrs_val,
                    }
                )
    if not pie_rows:
        st.info("No parsable activity rows found for the selected pie chart date range.")
        st.stop()
    pie_act_df = pd.DataFrame(pie_rows)
    pie_cat = pie_act_df.rename(columns={"activity": "Activity", "hours": "Hours"})
    pie_cat_norm = split_nonwip_activity_minutes(pie_cat)
    pie_rolled = pie_cat_norm.rename(columns={"Activity": "activity", "Hours": "hours"})
    pie_rolled = pie_rolled.groupby("activity", as_index=False).agg(hours=("hours", "sum"))
    pie_rolled = pie_rolled[
        ~pie_rolled["activity"].map(_norm_activity_name).isin(EXCLUDED_NON_WIP)
    ].sort_values("hours", ascending=False)
    if pie_rolled.empty:
        st.info('No pie chart data available after excluding "OOO" and "Non-WIP".')
        st.stop()
    pie_df = pie_rolled.head(int(top_n)).reset_index(drop=True)
    fig, ax = plt.subplots()
    ax.pie(
        pie_df["hours"],
        labels=pie_df["activity"],
        autopct="%1.0f%%",
        startangle=90,
    )
    ax.axis("equal")
    st.pyplot(fig)
with tabs[2]:
    st.subheader("Export")
    export_factor_out_ooo = st.toggle(
        "Factor out OOO from export calculations",
        value=False,
        key="export_factor_out_ooo",
        help="When on, OOO is removed from the denominator for export percentages, OOO Hours/OOO % are shown as 0, and Unaccounted is recalculated against capacity excluding OOO.",
    )
    team_export, ou_export, portfolio_export = _get_export_lookup_bundle(
        shared_metrics_df,
        shared_nonwip_df,
        org,
        export_factor_out_ooo,
    )
    def _format_export_display_team(df: pd.DataFrame) -> pd.io.formats.style.Styler:
        rename_map = {
            "portfolio": "Portfolio",
            "ou": "OU",
            "team": "Team",
            "week_start": "Week Start",
            "completed_hours": "Completed Hours",
            "people_count": "People Count",
            "non_wip_hours": "Non-WIP Hours",
            "ooo_hours": "OOO Hours",
            "capacity_hours": "Capacity Hours",
            "unaccounted_hours": "Unaccounted Hours",
            "wip_pct": "WIP %",
            "wip_avg_hours_day": "WIP Avg. Hours/Day",
            "non_wip_pct": "Non-WIP %",
            "non_wip_avg_hours_day": "Non-WIP Avg. Hours/Day",
            "ooo_pct": "OOO %",
            "ooo_avg_hours_day": "OOO Avg. Hours/Day",
            "unaccounted_pct": "Unaccounted %",
            "unaccounted_avg_hours_day": "Unaccounted Avg. Hours/Day",
        }
        preferred_order = [
            "portfolio", "ou", "team", "week_start",
            "capacity_hours", "people_count",
            "completed_hours", "wip_pct", "wip_avg_hours_day",
            "non_wip_hours", "non_wip_pct", "non_wip_avg_hours_day",
            "ooo_hours", "ooo_pct", "ooo_avg_hours_day",
            "unaccounted_hours", "unaccounted_pct", "unaccounted_avg_hours_day",
        ]
        cols = [c for c in preferred_order if c in df.columns] + [c for c in df.columns if c not in preferred_order]
        out = df[cols].copy().rename(columns=rename_map)
        if "Week Start" in out.columns:
            out["Week Start"] = pd.to_datetime(out["Week Start"], errors="coerce").dt.date
        fmt = {}
        for c in [
            "Completed Hours", "People Count", "Non-WIP Hours", "OOO Hours",
            "Capacity Hours", "Unaccounted Hours",
            "WIP Avg. Hours/Day", "Non-WIP Avg. Hours/Day",
            "OOO Avg. Hours/Day", "Unaccounted Avg. Hours/Day",
        ]:
            if c in out.columns:
                fmt[c] = "{:,.2f}"
        for c in ["WIP %", "Non-WIP %", "OOO %", "Unaccounted %"]:
            if c in out.columns:
                fmt[c] = "{:.1%}"
        styler = out.style.format(fmt)
        if "WIP %" in out.columns:
            styler = styler.map(lambda v: _threshold_cell_style(v, 0.80, good_if_gte=True), subset=["WIP %"])
        if "Non-WIP %" in out.columns:
            styler = styler.map(lambda v: _threshold_cell_style(v, 0.20), subset=["Non-WIP %"])
        return styler
    def _format_export_display_ou(df: pd.DataFrame) -> pd.io.formats.style.Styler:
        rename_map = {
            "portfolio": "Portfolio",
            "ou": "OU",
            "week_start": "Week Start",
            "completed_hours": "Completed Hours",
            "people_count": "People Count",
            "non_wip_hours": "Non-WIP Hours",
            "ooo_hours": "OOO Hours",
            "capacity_hours": "Capacity Hours",
            "unaccounted_hours": "Unaccounted Hours",
            "wip_pct": "WIP %",
            "wip_avg_hours_day": "WIP Avg. Hours/Day",
            "non_wip_pct": "Non-WIP %",
            "non_wip_avg_hours_day": "Non-WIP Avg. Hours/Day",
            "ooo_pct": "OOO %",
            "ooo_avg_hours_day": "OOO Avg. Hours/Day",
            "unaccounted_pct": "Unaccounted %",
            "unaccounted_avg_hours_day": "Unaccounted Avg. Hours/Day",
        }
        preferred_order = [
            "portfolio", "ou", "week_start",
            "capacity_hours", "people_count",
            "completed_hours", "wip_pct", "wip_avg_hours_day",
            "non_wip_hours", "non_wip_pct", "non_wip_avg_hours_day",
            "ooo_hours", "ooo_pct", "ooo_avg_hours_day",
            "unaccounted_hours", "unaccounted_pct", "unaccounted_avg_hours_day",
        ]
        cols = [c for c in preferred_order if c in df.columns] + [c for c in df.columns if c not in preferred_order]
        out = df[cols].copy().rename(columns=rename_map)
        if "Week Start" in out.columns:
            out["Week Start"] = pd.to_datetime(out["Week Start"], errors="coerce").dt.date
        fmt = {}
        for c in [
            "Completed Hours", "People Count", "Non-WIP Hours", "OOO Hours",
            "Capacity Hours", "Unaccounted Hours",
            "WIP Avg. Hours/Day", "Non-WIP Avg. Hours/Day",
            "OOO Avg. Hours/Day", "Unaccounted Avg. Hours/Day",
        ]:
            if c in out.columns:
                fmt[c] = "{:,.2f}"
        for c in ["WIP %", "Non-WIP %", "OOO %", "Unaccounted %"]:
            if c in out.columns:
                fmt[c] = "{:.1%}"
        styler = out.style.format(fmt)
        if "WIP %" in out.columns:
            styler = styler.map(lambda v: _threshold_cell_style(v, 0.80, good_if_gte=True), subset=["WIP %"])
        if "Non-WIP %" in out.columns:
            styler = styler.map(lambda v: _threshold_cell_style(v, 0.20), subset=["Non-WIP %"])
        return styler
    def _format_export_display_portfolio(df: pd.DataFrame) -> pd.io.formats.style.Styler:
        rename_map = {
            "portfolio": "Portfolio",
            "week_start": "Week Start",
            "completed_hours": "Completed Hours",
            "people_count": "People Count",
            "non_wip_hours": "Non-WIP Hours",
            "ooo_hours": "OOO Hours",
            "capacity_hours": "Capacity Hours",
            "unaccounted_hours": "Unaccounted Hours",
            "wip_pct": "WIP %",
            "wip_avg_hours_day": "WIP Avg. Hours/Day",
            "non_wip_pct": "Non-WIP %",
            "non_wip_avg_hours_day": "Non-WIP Avg. Hours/Day",
            "ooo_pct": "OOO %",
            "ooo_avg_hours_day": "OOO Avg. Hours/Day",
            "unaccounted_pct": "Unaccounted %",
            "unaccounted_avg_hours_day": "Unaccounted Avg. Hours/Day",
        }
        preferred_order = [
            "portfolio", "week_start",
            "capacity_hours", "people_count",
            "completed_hours", "wip_pct", "wip_avg_hours_day",
            "non_wip_hours", "non_wip_pct", "non_wip_avg_hours_day",
            "ooo_hours", "ooo_pct", "ooo_avg_hours_day",
            "unaccounted_hours", "unaccounted_pct", "unaccounted_avg_hours_day",
        ]
        cols = [c for c in preferred_order if c in df.columns] + [
            c for c in df.columns if c not in preferred_order and c != "ou"
        ]
        out = df[cols].copy().rename(columns=rename_map)
        if "Week Start" in out.columns:
            out["Week Start"] = pd.to_datetime(out["Week Start"], errors="coerce").dt.date
        fmt = {}
        for c in [
            "Completed Hours", "People Count", "Non-WIP Hours", "OOO Hours",
            "Capacity Hours", "Unaccounted Hours",
            "WIP Avg. Hours/Day", "Non-WIP Avg. Hours/Day",
            "OOO Avg. Hours/Day", "Unaccounted Avg. Hours/Day",
        ]:
            if c in out.columns:
                fmt[c] = "{:,.2f}"
        for c in ["WIP %", "Non-WIP %", "OOO %", "Unaccounted %"]:
            if c in out.columns:
                fmt[c] = "{:.1%}"
        styler = out.style.format(fmt)
        if "WIP %" in out.columns:
            styler = styler.map(lambda v: _threshold_cell_style(v, 0.80, good_if_gte=True), subset=["WIP %"])
        if "Non-WIP %" in out.columns:
            styler = styler.map(lambda v: _threshold_cell_style(v, 0.20), subset=["Non-WIP %"])
        return styler
    if team_export.empty:
        st.info("No exportable team/week data found.")
    else:
        st.markdown("#### Team weekly")
        st.dataframe(_format_export_display_team(team_export), width="stretch", hide_index=True)
        st.markdown("#### OU weekly")
        st.dataframe(_format_export_display_ou(ou_export), width="stretch", hide_index=True)
        st.markdown("#### Portfolio weekly")
        st.dataframe(_format_export_display_portfolio(portfolio_export), width="stretch", hide_index=True)
        try:
            team_export_display = _display_export_team_df(team_export)
            ou_export_display = _display_export_ou_df(ou_export)
            portfolio_export_display = _display_export_portfolio_df(portfolio_export)
            xlsx_bytes = _cached_excel_bytes(
                team_export_display,
                ou_export_display,
                portfolio_export_display,
            )
            st.download_button(
                label="Download Excel export",
                data=xlsx_bytes,
                file_name="enterprise_weekly_export.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"Excel export failed: {e}")