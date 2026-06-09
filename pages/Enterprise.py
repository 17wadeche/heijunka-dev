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
from datetime import datetime
import io
from utils.activity_map import ACTIVITY_MAP
from utils.csv_reading import read_csv_resilient
import altair as alt
from zoneinfo import ZoneInfo
def get_page_last_updated_label() -> str:
    try:
        ts = datetime.fromtimestamp(
            Path(__file__).stat().st_mtime,
            tz=ZoneInfo("America/Chicago"),
        )
        return f"Last updated: {ts.strftime('%Y-%m-%d %I:%M %p %Z')}"
    except Exception:
        return "Last updated: Unknown"
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
    "kuche":"Ku Che",
    "goutham kumar, p":"P Goutham Kumar",
}
PERSON_WEEKLY_HOURS = {
    "chelsey": 16.0,
    "mg": 36.0, 
    "lindsey": 32.0
}
def _capacity_from_count_with_person_overrides(
    count: float,
    default_hours: float,
    people: pd.DataFrame | None = None,
    person_col: str = "person",
) -> float:
    capacity_hours = float(count) * float(default_hours)
    if people is None or people.empty or person_col not in people.columns:
        return capacity_hours
    person_keys = people[person_col].dropna().map(person_key).drop_duplicates()
    for key in person_keys:
        if key in PERSON_WEEKLY_HOURS:
            capacity_hours += PERSON_WEEKLY_HOURS[key] - float(default_hours)
    return capacity_hours
def normalize_person_name(name: str) -> str:
    s = str(name or "").strip()
    s = " ".join(s.split())
    key = s.lower()
    return NAME_ALIASES.get(key, s)
def person_key(name: Any) -> str:
    s = normalize_person_name(str(name or ""))
    s = re.sub(r"\s*\(\d+\)\s*$", "", str(s or "").strip())
    s = re.sub(r"\s+", " ", s).strip()
    return s.casefold()
FY27_START = pd.Timestamp("2026-04-27").normalize()
FY27_END = pd.Timestamp("2027-04-25").normalize()
FY27_FISCAL_MONTHS = {
    "May '27":       ("2026-04-27", "2026-05-25"),
    "June '27":      ("2026-05-25", "2026-06-22"),
    "July '27":      ("2026-06-22", "2026-07-27"),
    "August '27":    ("2026-07-27", "2026-08-24"),
    "September '27": ("2026-08-24", "2026-09-21"),
    "October '27":   ("2026-09-21", "2026-10-26"),
    "November '27":  ("2026-10-26", "2026-11-23"),
    "December '27":  ("2026-11-23", "2026-12-21"),
    "January '28":   ("2026-12-21", "2027-01-25"),
    "February '28":  ("2027-01-25", "2027-02-22"),
    "March '28":     ("2027-02-22", "2027-03-22"),
    "April '28":     ("2027-03-22", "2027-04-26"),
}
FY27_FISCAL_MONTHS = {
    k: (pd.Timestamp(v[0]).normalize(), pd.Timestamp(v[1]).normalize())
    for k, v in FY27_FISCAL_MONTHS.items()
}
def _weeks_between(week_options, start_ts, end_ts) -> list[pd.Timestamp]:
    start_ts = pd.Timestamp(start_ts).normalize()
    end_ts = pd.Timestamp(end_ts).normalize()
    return [
        pd.Timestamp(w).normalize()
        for w in week_options
        if start_ts <= pd.Timestamp(w).normalize() <= end_ts
    ]
def _fy27_weeks(week_options) -> list[pd.Timestamp]:
    return _weeks_between(week_options, FY27_START, FY27_END)
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
def _apply_effective_capacity_for_export_display(
    df: pd.DataFrame,
    factor_out_ooo: bool,
) -> pd.DataFrame:
    out = df.copy()
    if (
        factor_out_ooo
        and "capacity_hours" in out.columns
        and "ooo_hours" in out.columns
    ):
        cap = pd.to_numeric(out["capacity_hours"], errors="coerce").fillna(0.0)
        ooo = pd.to_numeric(out["ooo_hours"], errors="coerce").fillna(0.0)
        out["capacity_hours"] = (cap - ooo).clip(lower=0.0)
    return out
@st.cache_data(show_spinner=False)
def _build_export_lookup_tables_cached(
    metrics_df: Optional[pd.DataFrame],
    nonwip_df: Optional[pd.DataFrame],
    _org,                    # underscore prefix → not hashed
    factor_out_ooo: bool,
    cache_key: str,          # stable hashable key for cache identity
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    team_export = _weekly_team_export_df(
        metrics_df,
        nonwip_df,
        _org,
        factor_out_ooo=factor_out_ooo,
    )
    if team_export is None or team_export.empty:
        empty = pd.DataFrame()
        return empty, empty, empty, empty
    team_export = team_export.copy()
    today = pd.Timestamp.now().normalize()
    if "week_start" in team_export.columns:
        team_export["week_start"] = pd.to_datetime(
            team_export["week_start"], errors="coerce"
        ).dt.normalize()
        team_export = team_export[team_export["week_start"] <= today].copy()
    for col in [
        "completed_hours",
        "non_wip_hours",
        "other_team_wip_hours",
        "ooo_hours",
        "unaccounted_hours",
        "capacity_hours",
    ]:
        if col not in team_export.columns:
            team_export[col] = 0.0
    team_export = team_export[
        (
            pd.to_numeric(team_export["completed_hours"], errors="coerce").fillna(0.0)
            + pd.to_numeric(team_export["non_wip_hours"], errors="coerce").fillna(0.0)
            + pd.to_numeric(team_export["other_team_wip_hours"], errors="coerce").fillna(0.0)
            + pd.to_numeric(team_export["unaccounted_hours"], errors="coerce").fillna(0.0)
            + pd.to_numeric(team_export["capacity_hours"], errors="coerce").fillna(0.0)
        ) > 0
    ].reset_index(drop=True)
    if team_export.empty:
        empty = pd.DataFrame()
        return empty, empty, empty, empty
    enterprise_export = _rollup_export_level(
        team_export,
        "enterprise",
        factor_out_ooo=factor_out_ooo,
    )
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
    team_export = _apply_effective_capacity_for_export_display(
        team_export,
        factor_out_ooo,
    )
    ou_export = _apply_effective_capacity_for_export_display(
        ou_export,
        factor_out_ooo,
    )
    portfolio_export = _apply_effective_capacity_for_export_display(
        portfolio_export,
        factor_out_ooo,
    )
    enterprise_export = _apply_effective_capacity_for_export_display(
        enterprise_export,
        factor_out_ooo,
    )
    return team_export, ou_export, portfolio_export, enterprise_export
@st.cache_data(show_spinner=False)
def load_precomputed(repo_root_str: str):
    data = load_common_data(repo_root_str)
    metrics = data["metrics"]
    nonwip = data["non_wip"]
    exploded = {
        "person_hours_long": explode_person_hours(metrics),
        "nonwip_by_person_long": explode_non_wip_by_person(nonwip),
        "people_in_wip_long": explode_people_in_wip(metrics),
        "nonwip_activity_long": _prepare_nonwip_activity_source(nonwip),
    }
    return data, exploded
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
    enterprise_export_display: pd.DataFrame,
    missing_teams_display: pd.DataFrame,
) -> bytes:
    return _excel_bytes_from_export_dfs(
        team_export_display,
        ou_export_display,
        portfolio_export_display,
        enterprise_export_display,
        missing_teams_display,
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
@st.cache_data(show_spinner=False)
def _build_missing_team_weeks_df(
    team_export: pd.DataFrame,
    org: OrgConfig,
    selected_weeks: list[pd.Timestamp] | None = None,
) -> pd.DataFrame:
    if org is None or not org.teams:
        return pd.DataFrame(columns=["Week Start", "Team", "Portfolio", "OU", "Status"])
    enabled_teams = [t for t in org.teams if t.enabled] or org.teams
    team_meta = _team_meta_lookup(org).copy()
    min_missing_week = pd.Timestamp("2026-04-13").normalize()
    weeks = []
    if selected_weeks:
        weeks = [pd.Timestamp(w).normalize() for w in selected_weeks if pd.notna(w)]
    elif team_export is not None and not team_export.empty and "week_start" in team_export.columns:
        weeks = sorted(
            pd.to_datetime(team_export["week_start"], errors="coerce")
            .dropna()
            .dt.normalize()
            .unique()
        )
    weeks = [w for w in weeks if w >= min_missing_week]
    if not weeks:
        return pd.DataFrame(columns=["Week Start", "Team", "Portfolio", "OU", "Status"])
    expected = pd.MultiIndex.from_product(
        [weeks, [t.name for t in enabled_teams]],
        names=["week_start", "team"],
    ).to_frame(index=False)
    actual = pd.DataFrame(columns=["week_start", "team"])
    if team_export is not None and not team_export.empty:
        actual_cols = ["week_start", "team"]
        if "missing_data" in team_export.columns:
            actual_cols.append("missing_data")
        actual = team_export.loc[:, actual_cols].copy()
        if "missing_data" in actual.columns:
            actual = actual[~actual["missing_data"].fillna(False).astype(bool)].copy()
        actual = actual.loc[:, ["week_start", "team"]]
        actual["week_start"] = pd.to_datetime(actual["week_start"], errors="coerce").dt.normalize()
        actual["team"] = actual["team"].astype(str).str.strip()
        actual = actual[
            actual["week_start"].notna() & (actual["week_start"] >= min_missing_week)
        ].dropna(subset=["week_start", "team"]).drop_duplicates()
    missing = (
        expected
        .merge(actual.assign(_present=1), on=["week_start", "team"], how="left")
        .loc[lambda d: d["_present"].isna(), ["week_start", "team"]]
        .merge(team_meta, on="team", how="left")
        .sort_values(["week_start", "portfolio", "ou", "team"])
        .reset_index(drop=True)
    )
    if missing.empty:
        return pd.DataFrame(columns=["Week Start", "Team", "Portfolio", "OU", "Status"])
    missing["Status"] = "Missing weekly data"
    missing = missing.rename(
        columns={
            "week_start": "Week Start",
            "team": "Team",
            "portfolio": "Portfolio",
            "ou": "OU",
        }
    )
    missing["Week Start"] = pd.to_datetime(missing["Week Start"], errors="coerce").dt.date
    return missing
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
@st.cache_data(show_spinner=False)
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
@st.cache_data(show_spinner=False)
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
            person_name = normalize_person_name(str(person).strip())
            if not person_name:
                continue
            team_name = str(r["team"]).strip().upper()
            keep_zero_capacity_person = (
                team_name in {"CDS", "NI"}
                and person_key(person_name) == "peter mchugh"
            )
            if a == 0.0 and t == 0.0 and not keep_zero_capacity_person:
                continue
            util = (a / t) if t not in (0, 0.0) else np.nan
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
@st.cache_data(show_spinner=False)
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
def _high_unaccounted_flag_series(df: pd.DataFrame, pct_col: str = "unaccounted_pct") -> pd.Series:
    if pct_col not in df.columns:
        return pd.Series([""] * len(df), index=df.index, dtype="string")
    vals = pd.to_numeric(df[pct_col], errors="coerce").fillna(0.0)
    return vals.map(lambda v: "❗" if float(v) > 0.25 else "").astype("string")
def _append_export_alert_column(
    df: pd.DataFrame,
    pct_col: str = "unaccounted_pct",
    alert_col_name: str = "Alert",
) -> pd.DataFrame:
    out = df.copy()
    out[alert_col_name] = _high_unaccounted_flag_series(out, pct_col=pct_col)
    return out
def merged_people_count_for_week(
    team: str,
    week,
    nw_frame: pd.DataFrame,
    person_hours: pd.DataFrame,
    people_in_wip: pd.DataFrame,
    long_nw: pd.DataFrame | None = None,
) -> int:
    wk = pd.to_datetime(week, errors="coerce").normalize()
    if nw_frame is not None and not nw_frame.empty and team in {"ENT", "DBS", "NV", "Enabling Technologies", "Lit & Letters", "Spine", "PSS", "PH", "SCS", "TDD", "ACM","CPT","DS","CDS","NI", "VSS","Endoscopy","Surgical AST-GST","PH-NM MEIC","TCT"}:
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
    if long_nw is None:
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
    def _canon_activity_for_bucket(label: str) -> str:
        raw = str(label or "").strip()
        if not raw:
            return ""
        lower = raw.lower().strip()
        mapped = ACTIVITY_MAP.get(lower, raw)
        return re.sub(r"[^A-Z0-9]", "", str(mapped).upper())
    other_team_key = "OTHERTEAMWIP"
    accounted_other: dict[str, float] = {}
    accounted_nonother: dict[str, float] = {}
    for d in activities:
        name = normalize_person_name(str(d.get("name", "")).strip())
        if not name:
            continue
        raw_act = str(d.get("activity", "")).strip().upper()
        act_key = _canon_activity_for_bucket(d.get("activity", ""))
        if act_key == "OOO" or raw_act in {"OOO", "OUT OF OFFICE", "HOLIDAY"}:
            continue
        try:
            hrs = float(d.get("hours", 0) or 0)
        except Exception:
            hrs = 0.0
        if hrs <= 0:
            continue
        if act_key == other_team_key:
            accounted_other[name] = accounted_other.get(name, 0.0) + hrs
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
        ["person", "Actual Hours", "Available Hours"]
    ].copy()
    if wip_people.empty:
        wip_people = pd.DataFrame(columns=["person", "Actual Hours", "Available Hours"])
    wip_people["person"] = wip_people["person"].astype(str).str.strip()
    wip_people["Completed Hours"] = pd.to_numeric(wip_people["Actual Hours"], errors="coerce").fillna(0.0)
    wip_people["Available Hours"] = pd.to_numeric(
        wip_people["Available Hours"],
        errors="coerce",
    )
    available_people = wip_people[["person", "Available Hours"]].copy()
    wip_people = wip_people[["person", "Completed Hours"]]
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
    OOO_LABELS = {"OOO", "OUT OF OFFICE", "HOLIDAY"}
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
            if activity in OOO_LABELS:
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
    available_people = _clean_person_col(available_people, "Available Hours")
    other_df = _clean_person_col(other_df, "Other Team WIP")
    acct_df = _clean_person_col(acct_df, "Accounted Non-WIP")
    ooo_df = _clean_person_col(ooo_df, "OOO Hours")
    nw_people = nw_people.groupby("person", as_index=False)["Non-WIP Hours"].sum()
    wip_people = wip_people.groupby("person", as_index=False)["Completed Hours"].sum()
    available_people = available_people.groupby("person", as_index=False)["Available Hours"].sum()
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
        .merge(available_people.astype({"person": "string"}), on="person", how="left")
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
    out["person_key"] = out["person"].map(person_key)
    irl_people_norm = {person_key(x) for x in (irl_people or set())}
    team_key = str(team).strip().upper()
    TEAM_WEEKLY_HOURS = {
        "CPT": 37.75,
        "CDS": 37.75,
        "NI": 37.75,
        "DS": 37.5,
        "LIT & LETTERS": 37.5,
    }
    PERSON_TEAM_WEEKLY_HOURS = {
        ("peter mchugh", "CDS"): 10.0,
        ("peter mchugh", "NI"): 27.75,
    }
    person_override_expected = out["person_key"].map(PERSON_WEEKLY_HOURS).astype("float64")
    base_expected = pd.Series(
        np.where(
            out["person_key"].isin(irl_people_norm),
            39.0,
            TEAM_WEEKLY_HOURS.get(team_key, float(week_hours)),
        ),
        index=out.index,
        dtype="float64",
    )
    team_override_expected = out["person_key"].map(
        lambda p: PERSON_TEAM_WEEKLY_HOURS.get((p, team_key), np.nan)
    ).astype("float64")
    out["Expected Hours"] = (
        team_override_expected
        .combine_first(person_override_expected)
        .combine_first(base_expected)
    )
    if team_key in {"CDS", "NI"} and "Available Hours" in out.columns:
        peter_available = pd.to_numeric(out["Available Hours"], errors="coerce")
        out.loc[
            out["person_key"].eq("peter mchugh") & peter_available.notna(),
            "Expected Hours",
        ] = peter_available
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
            return read_csv_resilient(path, encoding="utf-8-sig")
    except Exception:
        return None
    return None
@st.cache_data(show_spinner=False)
def load_common_data(repo_root_str: str) -> Dict[str, pd.DataFrame]:
    repo_root = Path(repo_root_str)
    candidates = {
        "metrics": repo_root / "IV_DATA" / "metrics.csv",
        "metrics_aggregate_dev": repo_root / "IV_DATA" / "metrics_aggregate_dev.csv",
        "non_wip": repo_root / "IV_DATA" / "non_wip.csv",
        "non_wip_activities": repo_root / "IV_DATA" / "non_wip_activities.csv",
        "closures": repo_root / "closures.csv",
        "NS_WIP": repo_root / "NS_DATA" / "NS_WIP.csv",
        "ns_non_wip_activities": repo_root / "NS_DATA" / "ns_non_wip_activities.csv",
        "CRM_WIP": repo_root / "CRM_DATA" / "CRM_WIP.csv",
        "crm_non_wip_activities": repo_root / "CRM_DATA" / "crm_non_wip_activities.csv",
        "MS_WIP": repo_root / "MS_DATA" / "MS_WIP.csv",
        "ms_non_wip_activities": repo_root / "MS_DATA" / "ms_non_wip_activities.csv",
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
    dt = pd.to_datetime(s, errors="coerce")
    return (dt - pd.to_timedelta(dt.dt.dayofweek, unit="D")).dt.normalize()
def _to_monday(series: pd.Series) -> pd.Series:
    return _weekly_start(series)
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
with st.sidebar:
    st.caption(get_page_last_updated_label())
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
data, exploded = load_precomputed(str(repo_root))
enabled_teams = [t for t in org.teams if t.enabled]
all_team_names = [t.name for t in org.teams]
enabled_team_names = [t.name for t in enabled_teams] or all_team_names
team_filter = enabled_team_names or all_team_names
org_cache_key = json.dumps(org.raw, sort_keys=True, default=str)
if not team_filter:
    st.warning("No teams selected.")
    st.stop()
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
def _is_truthy_meta(value: Any) -> bool:
    return _coerce_bool(value, default=False)
def _to_float_meta(value: Any, default: float = 0.0) -> float:
    try:
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return default
        return float(value)
    except Exception:
        return default
def _unaccounted_only_team_rows(
    org: OrgConfig,
    weeks: list[pd.Timestamp],
    existing_team_weeks: set[tuple[str, pd.Timestamp]],
    factor_out_ooo: bool = False,
) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    normalized_weeks = sorted({pd.Timestamp(w).normalize() for w in weeks if pd.notna(w)})
    if not normalized_weeks:
        return pd.DataFrame()
    for team_cfg in org.teams:
        if not team_cfg.enabled:
            continue
        meta = team_cfg.meta or {}
        if not _is_truthy_meta(meta.get("unaccounted_only")):
            continue
        people_count = _to_float_meta(
            meta.get("people_count", meta.get("headcount", meta.get("people"))),
            default=0.0,
        )
        weekly_hours = _to_float_meta(
            meta.get("weekly_hours_per_person", meta.get("hours_per_person", meta.get("week_hours"))),
            default=40.0,
        )
        capacity_hours = _to_float_meta(
            meta.get("weekly_capacity_hours", meta.get("capacity_hours")),
            default=people_count * weekly_hours,
        )
        if capacity_hours <= 0:
            continue
        for wk in normalized_weeks:
            key = (team_cfg.name, wk)
            if key in existing_team_weeks:
                continue
            pct_denom = capacity_hours
            rows.append({
                "team": team_cfg.name,
                "week_start": wk,
                "people_count": people_count,
                "completed_hours": 0.0,
                "other_team_wip_hours": 0.0,
                "non_wip_hours": 0.0,
                "ooo_hours": 0.0,
                "capacity_hours": capacity_hours,
                "unaccounted_hours": capacity_hours,
                "over_hours": 0.0,
                "warning": str(meta.get("warning") or "Missing weekly data"),
                "missing_data": True,
                "wip_pct": 0.0,
                "other_team_wip_pct": 0.0,
                "non_wip_pct": 0.0,
                "ooo_pct": 0.0,
                "unaccounted_pct": (capacity_hours / pct_denom) if pct_denom > 0 else pd.NA,
            })
    if not rows:
        return pd.DataFrame()
    return pd.DataFrame(rows)
def _prepare_weekly_accounting_inputs(
    metrics_frame: pd.DataFrame,
    nw_frame: pd.DataFrame,
) -> dict[str, pd.DataFrame]:
    long_nw = explode_non_wip_by_person(nw_frame)
    person_hours = explode_person_hours(metrics_frame)
    people_in_wip = explode_people_in_wip(metrics_frame)
    if not long_nw.empty:
        long_nw["period_date"] = _to_monday(long_nw["period_date"])
    if not person_hours.empty:
        person_hours["period_date"] = _to_monday(person_hours["period_date"])
    if not people_in_wip.empty:
        people_in_wip["period_date"] = _to_monday(people_in_wip["period_date"])
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
    raw_nw["period_date"] = _to_monday(raw_nw["period_date"])
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
    capacity_hours = float((irl_count * 39.0) + (non_irl_count * 40.0))
    if "non_wip_by_person" in row.columns:
        payload = row.iloc[0].get("non_wip_by_person")
        try:
            obj = json.loads(payload) if isinstance(payload, str) else payload
        except Exception:
            obj = {}
        if isinstance(obj, dict):
            for name in obj.keys():
                key = person_key(name)
                if key in PERSON_WEEKLY_HOURS:
                    default_hours = 39.0 if key in irl_people_norm else 40.0
                    capacity_hours += PERSON_WEEKLY_HOURS[key] - default_hours
    return capacity_hours
def _person_available_hours_for_week(
    person_hours: pd.DataFrame | None,
    team: str,
    week,
    person_key_value: str,
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
    ph["period_date"] = _to_monday(ph["period_date"])
    mask = (
        ph["team"].astype(str).str.strip().str.upper().eq(str(team).strip().upper())
        & ph["period_date"].eq(pd.Timestamp(wk).normalize())
        & ph["person"].map(person_key).eq(person_key(person_key_value))
    )
    vals = pd.to_numeric(
        ph.loc[mask, "Available Hours"],
        errors="coerce",
    ).dropna()
    return float(vals.sum()) if not vals.empty else np.nan
def _non_ooo_activity_hours_from_row(row: Any) -> float:
    payload = row.get("non_wip_activities", "[]")
    try:
        activities = json.loads(payload) if isinstance(payload, str) else payload
    except Exception:
        activities = []
    if not isinstance(activities, list) or not activities:
        return np.nan
    total = 0.0
    saw_non_ooo = False
    ooo_labels = {"OOO", "OUT OF OFFICE", "HOLIDAY"}
    for item in activities:
        if not isinstance(item, dict):
            continue
        raw_activity = str(
            item.get("activity") or item.get("Activity") or item.get("type") or ""
        ).strip()
        activity = raw_activity.upper()
        try:
            hours = float(item.get("hours", item.get("Hours", 0)) or 0.0)
        except Exception:
            hours = 0.0
        if hours <= 0 or activity in ooo_labels:
            continue
        total += hours
        saw_non_ooo = True
    return float(total) if saw_non_ooo else np.nan
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
        nw["period_date"] = _to_monday(nw[dc])
    else:
        nw["period_date"] = _to_monday(nw["period_date"])
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
                metrics_frame["period_date"] = _to_monday(metrics_frame[dc])
        else:
            metrics_frame["period_date"] = _to_monday(metrics_frame["period_date"])
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
    irl_lookup: dict[str, set[str]] = {
        t: irl_people_for_team(t, teams_cfg) for t in enabled_team_names
    }
    rows: list[dict[str, Any]] = []
    for _, nw_row in nw.iterrows():
        team = str(nw_row.get("team", "")).strip()
        wk = pd.to_datetime(nw_row.get("week_start"), errors="coerce")
        if not team or pd.isna(wk):
            continue
        wk = pd.Timestamp(wk).normalize()
        team_irl_people = irl_lookup.get(team, set())
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
            long_nw=long_nw,           
        )
        if people_count is None or float(people_count) <= 0:
            people_count = float(
                wk_people["person"].astype(str).str.strip().replace("", pd.NA).dropna().nunique()
            )
        if team in {"SVT", "PVH","NV", "Enabling Technologies", "DBS", "PH", "Spine", "PSS", "SCS", "TDD","ACM","ACM","VSS","Endoscopy","Surgical AST-GST", "PH-NM MEIC", "TCT"}:
            capacity_hours = _capacity_from_count_with_person_overrides(people_count, 40.0, wk_people)
        elif team == "DS":
            capacity_hours = _capacity_from_count_with_person_overrides(people_count, 37.5, wk_people)
        elif team == "Lit & Letters":
            capacity_hours = _capacity_from_count_with_person_overrides(people_count, 37.5, wk_people)
        elif team == "CPT":
            capacity_hours = (
                float(pd.to_numeric(wk_people["Available Hours"], errors="coerce").fillna(0.0).sum())
                if "Available Hours" in wk_people.columns
                else 0.0
            )
            if capacity_hours <= 0.0:
                capacity_hours = float(wk_people["Expected Hours"].sum())
        elif team in {"CDS", "NI"}:
            peter_available = _person_available_hours_for_week(
                person_hours=person_hours,
                team=team,
                week=wk,
                person_key_value="peter mchugh",
            )
            peter_fallback_capacity = 10.0 if team == "CDS" else 27.75
            peter_capacity = (
                float(peter_available)
                if pd.notna(peter_available)
                else peter_fallback_capacity
            )
            assigned_count = 1 if float(people_count) > 0 else 0
            remaining_count = max(float(people_count) - assigned_count, 0)
            capacity_hours = (
                (assigned_count * peter_capacity)
                + (remaining_count * 37.75)
            )
        elif team == "ENT":
            capacity_hours = ent_capacity_hours_for_week(
                team=team,
                week=wk,
                nw_frame=nw,
                irl_people=team_irl_people,
            )
        else:
            capacity_hours = float(wk_people["Expected Hours"].sum())
        ooo_hours = float(wk_people["OOO Hours"].sum())
        other_team_wip_hours = (
            float(pd.to_numeric(wk_people["Other Team WIP"], errors="coerce").fillna(0.0).sum())
            if not wk_people.empty and "Other Team WIP" in wk_people.columns
            else 0.0
        )
        if team == "CPT":
            cpt_nonwip = long_nw.loc[
                (long_nw["team"] == team) &
                (pd.to_datetime(long_nw["period_date"], errors="coerce").dt.normalize() == wk),
                "Non-WIP Hours"
            ]
            total_non_wip_hours = float(
                pd.to_numeric(cpt_nonwip, errors="coerce").fillna(0.0).sum()
            )
        else:
            total_non_wip_hours = float(
                pd.to_numeric(nw_row.get("non_wip_hours", 0.0), errors="coerce") or 0.0
            )
        if team == "PM-CTS":
            pm_cts_non_ooo_total = _non_ooo_activity_hours_from_row(nw_row)
            if pd.notna(pm_cts_non_ooo_total):
                total_non_wip_hours = float(pm_cts_non_ooo_total)
        if team == "PH-NM MEIC":
            non_wip_hours = max(total_non_wip_hours - other_team_wip_hours - ooo_hours, 0.0)
        else:
            non_wip_hours = max(total_non_wip_hours - other_team_wip_hours, 0.0)
        completed_match = metrics_team[
            (metrics_team["team"] == team) &
            (pd.to_datetime(metrics_team["week_start"], errors="coerce").dt.normalize() == wk)
        ]
        completed_hours = float(completed_match["completed_hours"].sum()) if not completed_match.empty else 0.0
        if completed_hours == 0.0 and not wk_people.empty and "Completed Hours" in wk_people.columns:
            completed_hours = float(pd.to_numeric(wk_people["Completed Hours"], errors="coerce").fillna(0.0).sum())
        unaccounted_hours = max(
            capacity_hours - completed_hours - other_team_wip_hours - non_wip_hours - ooo_hours,
            0.0,
        )
        over_hours = max(
            completed_hours + other_team_wip_hours + non_wip_hours + ooo_hours - capacity_hours,
            0.0,
        )
        warning = f"Over {over_hours:.2f} hours" if over_hours > 0 else ""
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
            "other_team_wip_hours": other_team_wip_hours,
            "non_wip_hours": non_wip_hours,
            "ooo_hours": ooo_hours,
            "capacity_hours": capacity_hours,
            "unaccounted_hours": unaccounted_hours,
            "over_hours": over_hours,
            "warning": warning,
            "missing_data": False,
            "wip_pct": (completed_hours / pct_denom) if pct_denom > 0 else pd.NA,
            "other_team_wip_pct": (other_team_wip_hours / pct_denom) if pct_denom > 0 else pd.NA,
            "non_wip_pct": (non_wip_hours / pct_denom) if pct_denom > 0 else pd.NA,
            "ooo_pct": ooo_pct,
            "unaccounted_pct": (unaccounted_hours / pct_denom) if pct_denom > 0 else pd.NA,
        })
    known_weeks = [pd.Timestamp(w).normalize() for w in nw["week_start"].dropna().unique()]
    if not metrics_team.empty and "week_start" in metrics_team.columns:
        known_weeks.extend(
            pd.Timestamp(w).normalize()
            for w in metrics_team["week_start"].dropna().unique()
        )
    existing_team_weeks = {
        (str(r.get("team", "")).strip(), pd.Timestamp(r.get("week_start")).normalize())
        for r in rows
        if str(r.get("team", "")).strip() and pd.notna(r.get("week_start"))
    }
    synthetic_rows = _unaccounted_only_team_rows(
        org,
        known_weeks,
        existing_team_weeks,
        factor_out_ooo=factor_out_ooo,
    )
    if not synthetic_rows.empty:
        rows.extend(synthetic_rows.to_dict("records"))
    if not rows:
        return pd.DataFrame()
    base = pd.DataFrame(rows)
    base = (
        base.groupby(["team", "week_start"], as_index=False)
        .agg(
            people_count=("people_count", "max"),
            completed_hours=("completed_hours", "sum"),
            other_team_wip_hours=("other_team_wip_hours", "sum"),
            non_wip_hours=("non_wip_hours", "sum"),
            ooo_hours=("ooo_hours", "sum"),
            capacity_hours=("capacity_hours", "sum"),
            over_hours=("over_hours", "sum"),
            missing_data=("missing_data", "all"),
        )
    )
    base["unaccounted_hours"] = (
        base["capacity_hours"]
        - base["completed_hours"]
        - base["other_team_wip_hours"]
        - base["non_wip_hours"]
        - base["ooo_hours"]
    ).clip(lower=0.0)
    base["warning"] = np.where(
        base["over_hours"] > 0,
        "Over " + base["over_hours"].round(2).astype(str) + " hours",
        "",
    )
    if factor_out_ooo:
        pct_denom = (base["capacity_hours"] - base["ooo_hours"]).clip(lower=0.0)
        base["ooo_pct"] = 0.0
    else:
        pct_denom = base["capacity_hours"]
        base["ooo_pct"] = (base["ooo_hours"] / pct_denom).where(pct_denom > 0)
    base["wip_pct"] = (base["completed_hours"] / pct_denom).where(pct_denom > 0)
    base["non_wip_pct"] = (base["non_wip_hours"] / pct_denom).where(pct_denom > 0)
    base["other_team_wip_pct"] = (base["other_team_wip_hours"] / pct_denom).where(pct_denom > 0)
    base["unaccounted_pct"] = (base["unaccounted_hours"] / pct_denom).where(pct_denom > 0)
    base = base.merge(meta, on="team", how="left")
    base = _add_avg_hours_day_columns(base)
    return base.sort_values(["week_start", "portfolio", "ou", "team"]).reset_index(drop=True)
def _rollup_export_level(df: pd.DataFrame, level: str, factor_out_ooo: bool = False) -> pd.DataFrame:
    if df.empty:
        return df.copy()
    if level == "enterprise":
        group_cols = ["week_start"]
    elif level == "ou":
        group_cols = ["week_start", "portfolio", "ou"]
    elif level == "portfolio":
        group_cols = ["week_start", "portfolio"]
    else:
        raise ValueError("level must be 'enterprise', 'ou', or 'portfolio'")
    out = (
        df.groupby(group_cols, as_index=False)
        .agg(
            people_count=("people_count", "sum"),
            completed_hours=("completed_hours", "sum"),
            other_team_wip_hours=("other_team_wip_hours", "sum"),
            non_wip_hours=("non_wip_hours", "sum"),
            ooo_hours=("ooo_hours", "sum"),
            capacity_hours=("capacity_hours", "sum"),
        )
    )
    out["unaccounted_hours"] = (
        out["capacity_hours"]
        - out["completed_hours"]
        - out["other_team_wip_hours"]
        - out["non_wip_hours"]
        - out["ooo_hours"]
    ).clip(lower=0.0)
    out["over_hours"] = (
        out["completed_hours"]
        + out["other_team_wip_hours"]
        + out["non_wip_hours"]
        + out["ooo_hours"]
        - out["capacity_hours"]
    ).clip(lower=0.0)
    out["warning"] = np.where(
        out["over_hours"] > 0,
        "Over " + out["over_hours"].round(2).astype(str) + " hours",
        "",
    )
    if factor_out_ooo:
        pct_denom = (out["capacity_hours"] - out["ooo_hours"]).clip(lower=0.0)
        out["ooo_pct"] = 0.0
    else:
        pct_denom = out["capacity_hours"]
        out["ooo_pct"] = (out["ooo_hours"] / pct_denom).where(pct_denom > 0)
    out["wip_pct"] = (out["completed_hours"] / pct_denom).where(pct_denom > 0)
    out["other_team_wip_pct"] = (out["other_team_wip_hours"] / pct_denom).where(pct_denom > 0)
    out["non_wip_pct"] = (out["non_wip_hours"] / pct_denom).where(pct_denom > 0)
    out["unaccounted_pct"] = (out["unaccounted_hours"] / pct_denom).where(pct_denom > 0)
    out = _add_avg_hours_day_columns(out)
    if level == "enterprise":
        out["enterprise"] = "Enterprise"
        out["portfolio"] = pd.NA
        out["ou"] = pd.NA
    elif level == "portfolio":
        out["ou"] = pd.NA
    cols = [
        "week_start",
        "enterprise",
        "portfolio",
        "ou",
        "people_count",
        "completed_hours",
        "wip_pct",
        "wip_avg_hours_day",
        "other_team_wip_hours",
        "other_team_wip_pct",
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
        "over_hours",
        "warning",
    ]
    cols = [c for c in cols if c in out.columns]
    return out[cols].sort_values(group_cols).reset_index(drop=True)
def _display_export_team_df(df: pd.DataFrame) -> pd.DataFrame:
    team_df = _append_alert_before_display(df, include_alert=True)
    rename_map = {
        "Alert": "Alert",
        "team": "Team",
        "week_start": "Week Start",
        "completed_hours": "Completed Hours",
        "people_count": "People",
        "non_wip_hours": "Non-WIP Hours",
        "ooo_hours": "OOO Hours",
        "capacity_hours": "Capacity",
        "unaccounted_hours": "Unaccounted Hours",
        "wip_pct": "WIP %",
        "non_wip_pct": "Non-WIP %",
        "ooo_pct": "OOO %",
        "unaccounted_pct": "Unaccounted %",
        "other_team_wip_hours": "Other Team WIP",
        "other_team_wip_pct": "Other Team WIP %",
        "over_hours": "Over Hours",
        "warning": "Warning",
    }
    preferred_order = [
        "Alert",
        "team", "week_start",
        "capacity_hours", "people_count",
        "completed_hours", "wip_pct",
        "other_team_wip_hours", "other_team_wip_pct",
        "non_wip_hours", "non_wip_pct",
        "ooo_hours", "ooo_pct",
        "unaccounted_hours", "unaccounted_pct",
        "over_hours", "warning",
    ]
    cols = [c for c in preferred_order if c in team_df.columns]
    out = team_df[cols].copy().rename(columns=rename_map)
    if "Week Start" in out.columns:
        out["Week Start"] = pd.to_datetime(out["Week Start"], errors="coerce").dt.date
    return out
def _display_export_ou_df(df: pd.DataFrame) -> pd.DataFrame:
    ou_df = _append_alert_before_display(df, include_alert=True)
    rename_map = {
        "Alert": "Alert",
        "ou": "OU",
        "week_start": "Week Start",
        "capacity_hours": "Capacity",
        "people_count": "People",
        "completed_hours": "Completed Hours",
        "wip_pct": "WIP %",
        "other_team_wip_hours": "Other Team WIP",
        "other_team_wip_pct": "Other Team WIP %",
        "non_wip_hours": "Non-WIP Hours",
        "non_wip_pct": "Non-WIP %",
        "ooo_hours": "OOO Hours",
        "ooo_pct": "OOO %",
        "unaccounted_hours": "Unaccounted Hours",
        "unaccounted_pct": "Unaccounted %",
        "over_hours": "Over Hours",
        "warning": "Warning",
    }
    preferred_order = [
        "Alert",
        "ou", "week_start",
        "capacity_hours", "people_count",
        "completed_hours", "wip_pct",
        "other_team_wip_hours", "other_team_wip_pct",
        "non_wip_hours", "non_wip_pct",
        "ooo_hours", "ooo_pct",
        "unaccounted_hours", "unaccounted_pct",
        "over_hours", "warning",
    ]
    cols = [c for c in preferred_order if c in ou_df.columns]
    out = ou_df[cols].copy().rename(columns=rename_map)
    if "Week Start" in out.columns:
        out["Week Start"] = pd.to_datetime(out["Week Start"], errors="coerce").dt.date
    return out
def _display_export_portfolio_df(df: pd.DataFrame) -> pd.DataFrame:
    portfolio_df = _append_alert_before_display(df, include_alert=True)
    rename_map = {
        "Alert": "Alert",
        "portfolio": "Portfolio",
        "week_start": "Week Start",
        "capacity_hours": "Capacity",
        "people_count": "People",
        "completed_hours": "Completed Hours",
        "wip_pct": "WIP %",
        "other_team_wip_hours": "Other Team WIP",
        "other_team_wip_pct": "Other Team WIP %",
        "non_wip_hours": "Non-WIP Hours",
        "non_wip_pct": "Non-WIP %",
        "ooo_hours": "OOO Hours",
        "ooo_pct": "OOO %",
        "unaccounted_hours": "Unaccounted Hours",
        "unaccounted_pct": "Unaccounted %",
        "over_hours": "Over Hours",
        "warning": "Warning",
    }
    preferred_order = [
        "Alert",
        "portfolio", "week_start",
        "capacity_hours", "people_count",
        "completed_hours", "wip_pct",
        "other_team_wip_hours", "other_team_wip_pct",
        "non_wip_hours", "non_wip_pct",
        "ooo_hours", "ooo_pct",
        "unaccounted_hours", "unaccounted_pct",
        "over_hours", "warning",
    ]
    cols = [c for c in preferred_order if c in portfolio_df.columns]
    out = portfolio_df[cols].copy().rename(columns=rename_map)
    if "Week Start" in out.columns:
        out["Week Start"] = pd.to_datetime(out["Week Start"], errors="coerce").dt.date
    return out
def _display_export_enterprise_df(df: pd.DataFrame) -> pd.DataFrame:
    enterprise_df = _append_alert_before_display(df, include_alert=True)
    rename_map = {
        "Alert": "Alert",
        "enterprise": "Enterprise",
        "week_start": "Week Start",
        "capacity_hours": "Capacity",
        "people_count": "People",
        "completed_hours": "Completed Hours",
        "wip_pct": "WIP %",
        "other_team_wip_hours": "Other Team WIP",
        "other_team_wip_pct": "Other Team WIP %",
        "non_wip_hours": "Non-WIP Hours",
        "non_wip_pct": "Non-WIP %",
        "ooo_hours": "OOO Hours",
        "ooo_pct": "OOO %",
        "unaccounted_hours": "Unaccounted Hours",
        "unaccounted_pct": "Unaccounted %",
        "over_hours": "Over Hours",
        "warning": "Warning",
    }
    preferred_order = [
        "Alert",
        "enterprise", "week_start",
        "capacity_hours", "people_count",
        "completed_hours", "wip_pct",
        "other_team_wip_hours", "other_team_wip_pct",
        "non_wip_hours", "non_wip_pct",
        "ooo_hours", "ooo_pct",
        "unaccounted_hours", "unaccounted_pct",
        "over_hours", "warning",
    ]
    cols = [c for c in preferred_order if c in enterprise_df.columns]
    out = enterprise_df[cols].copy().rename(columns=rename_map)
    if "Week Start" in out.columns:
        out["Week Start"] = pd.to_datetime(out["Week Start"], errors="coerce").dt.date
    return out
def _over_100_row_highlight_style(row: pd.Series) -> list[str]:
    required = ["WIP %", "Non-WIP %", "OOO %"]
    if not all(c in row.index for c in required):
        return [""] * len(row)
    vals = pd.to_numeric(
        pd.Series([row["WIP %"], row["Non-WIP %"], row["OOO %"]]),
        errors="coerce",
    ).fillna(0.0)
    is_over = float(vals.sum()) > 1.0
    if not is_over:
        return [""] * len(row)
    return ["background-color: #fef3c7;"] * len(row)
def _excel_col_name(idx: int) -> str:
    name = ""
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        name = chr(65 + rem) + name
    return name
def _apply_over_100_outline_xlsxwriter(writer, sheet_name: str, df: pd.DataFrame) -> None:
    if df is None or df.empty:
        return
    required_cols = ["WIP %", "Non-WIP %", "OOO %"]
    if not all(c in df.columns for c in required_cols):
        return
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    yellow_outline = workbook.add_format({
        "border": 2,
        "border_color": "#facc15",
    })
    wip_col = _excel_col_name(df.columns.get_loc("WIP %"))
    non_wip_col = _excel_col_name(df.columns.get_loc("Non-WIP %"))
    ooo_col = _excel_col_name(df.columns.get_loc("OOO %"))
    first_data_row = 2  # Excel row 1 is header; row 2 is first data row
    last_data_row = len(df) + 1
    last_col = _excel_col_name(len(df.columns) - 1)
    worksheet.conditional_format(
        f"A{first_data_row}:{last_col}{last_data_row}",
        {
            "type": "formula",
            "criteria": f"=(${wip_col}{first_data_row}+${non_wip_col}{first_data_row}+${ooo_col}{first_data_row})>1",
            "format": yellow_outline,
        },
    )
def _excel_bytes_from_export_dfs(
    team_df: pd.DataFrame,
    ou_df: pd.DataFrame,
    portfolio_df: pd.DataFrame,
    enterprise_df: pd.DataFrame,
    missing_teams_df: pd.DataFrame,
) -> bytes:
    last_err = None
    for engine in ("xlsxwriter", "openpyxl"):
        buf = io.BytesIO()
        try:
            with pd.ExcelWriter(buf, engine=engine) as writer:
                sheets = [
                    ("Team Weekly", team_df),
                    ("OU Weekly", ou_df),
                    ("Portfolio Weekly", portfolio_df),
                    ("Enterprise Weekly", enterprise_df),
                    ("Missing Teams", missing_teams_df),
                ]
                for sheet_name, df in sheets:
                    if df is None:
                        continue
                    safe_name = _safe_sheet_name(sheet_name)
                    df.to_excel(writer, index=False, sheet_name=safe_name)
                if engine == "xlsxwriter":
                    for sheet_name, df in sheets[:4]:
                        if df is not None and not df.empty:
                            _apply_over_100_outline_xlsxwriter(writer, _safe_sheet_name(sheet_name), df)
            buf.seek(0)
            return buf.getvalue()
        except Exception as e:
            last_err = e
    raise RuntimeError(
        f"Excel export requires openpyxl or xlsxwriter to be installed. Last error: {last_err}"
    )
def _safe_sheet_name(name: str) -> str:
    return re.sub(r"[\[\]\*\/\\\?\:]", "_", str(name))[:31]
@st.cache_data(show_spinner=False)
def _cached_custom_excel_bytes(
    sheet_items: tuple[tuple[str, pd.DataFrame], ...],
) -> bytes:
    buf = io.BytesIO()
    last_err = None
    for engine in ("openpyxl", "xlsxwriter"):
        buf = io.BytesIO()
        try:
            with pd.ExcelWriter(buf, engine=engine) as writer:
                wrote_any = False
                for sheet_name, df in sheet_items:
                    if df is None:
                        continue
                    safe_name = _safe_sheet_name(sheet_name)
                    df.to_excel(writer, index=False, sheet_name=safe_name)
                    wrote_any = True
                if not wrote_any:
                    pd.DataFrame({"Message": ["No rows selected for export."]}).to_excel(
                        writer,
                        index=False,
                        sheet_name="Export",
                    )
            buf.seek(0)
            return buf.getvalue()
        except Exception as e:
            last_err = e
    raise RuntimeError(f"Could not create Excel workbook: {last_err}")
def _apply_display_column_selection(
    df: pd.DataFrame,
    selected_columns: list[str],
) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    cols = [c for c in selected_columns if c in df.columns]
    if not cols:
        return df.iloc[:, 0:0].copy()
    return df[cols].copy()
def _append_alert_before_display(df: pd.DataFrame, include_alert: bool) -> pd.DataFrame:
    if df is None or df.empty or not include_alert:
        return df
    if "unaccounted_pct" in df.columns:
        return _append_export_alert_column(df, pct_col="unaccounted_pct")
    return df
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
factor_out_ooo = st.toggle(
    "Factor out OOO from calculations",
    value=True,
    key="factor_out_ooo",
)
page = st.segmented_control(
    "Tab:",
    options=["Overview", "Non-WIP", "Export"],
    default="Overview",
    key="enterprise_section",
)
@st.cache_data(show_spinner=False)
def _get_export_lookup_bundle(
    shared_metrics_df: Optional[pd.DataFrame],
    shared_nonwip_df: Optional[pd.DataFrame],
    _org,
    factor_out_ooo: bool,
    cache_key: str,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    return _build_export_lookup_tables_cached(
        shared_metrics_df,
        shared_nonwip_df,
        _org,
        factor_out_ooo=factor_out_ooo,
        cache_key=cache_key,
    )
EXCLUDED_NON_WIP = {"ooo", "non-wip", "non_wip", "other", "nan", "", "break", "other team wip", "extra wip", "see commercial tab","other (hours)", "used other", "used the other", "export"}
def _norm_activity_name(val: Any) -> str:
    return str(val).strip().lower().replace("_", "-")
def _is_training_or_mentoring_activity(val: Any) -> bool:
    return bool(re.search(r"\b(?:train\w*|mentor\w*)\b", str(val), flags=re.IGNORECASE))
@st.cache_data(show_spinner=False)
def build_training_mentoring_export(
    source_raw: pd.DataFrame,
    before_date: Any,
) -> pd.DataFrame:
    columns = ["Team", "Week Start", "Training/Mentoring Hours"]
    if source_raw is None or source_raw.empty:
        return pd.DataFrame(columns=columns)
    source_df = _normalize_df_columns(source_raw.copy())
    team_col = _get_team_col(source_df)
    date_col = _get_date_col(source_df)
    json_col = _first_col(source_df, ["non_wip_activities", "non-wip_activities"])
    if not (team_col and date_col and json_col):
        return pd.DataFrame(columns=columns)
    source_df[date_col] = _safe_to_datetime(source_df, date_col)
    cutoff = pd.Timestamp(before_date).normalize()
    source_df = source_df.dropna(subset=[date_col]).copy()
    source_df = source_df[source_df[date_col].dt.normalize().lt(cutoff)].copy()
    source_df["_week_start"] = _weekly_start(source_df[date_col])
    source_df["_team"] = source_df[team_col].astype(str).str.strip()
    source_df = source_df[source_df["_team"].ne("")].copy()
    if source_df.empty:
        return pd.DataFrame(columns=columns)
    team_weeks = (
        source_df[["_team", "_week_start"]]
        .drop_duplicates()
        .rename(columns={"_team": "Team", "_week_start": "Week Start"})
    )
    activity_rows: list[dict[str, Any]] = []
    for _, row in source_df.iterrows():
        payload = _loads_json_maybe(row[json_col])
        if isinstance(payload, dict):
            payload = [payload]
        if not isinstance(payload, list):
            continue
        for item in payload:
            if not isinstance(item, dict):
                continue
            activity = item.get("activity") or item.get("Activity") or item.get("type")
            hours = item.get("hours") or item.get("Hours")
            if activity is None or hours is None:
                continue
            if not _is_training_or_mentoring_activity(activity):
                continue
            activity_rows.append(
                {
                    "Team": row["_team"],
                    "Week Start": row["_week_start"],
                    "Activity": str(activity).strip(),
                    "Hours": hours,
                }
            )
    training_rows: list[dict[str, Any]] = []
    if activity_rows:
        activities = pd.DataFrame(activity_rows)
        for (team, week_start), group in activities.groupby(["Team", "Week Start"]):
            split_activities = split_nonwip_activity_minutes(group[["Activity", "Hours"]])
            training_hours = pd.to_numeric(
                split_activities.loc[
                    split_activities["Activity"].map(_is_training_or_mentoring_activity),
                    "Hours",
                ],
                errors="coerce",
            ).sum()
            training_rows.append(
                {
                    "Team": team,
                    "Week Start": week_start,
                    "Training/Mentoring Hours": float(training_hours),
                }
            )
    if training_rows:
        totals = pd.DataFrame(training_rows)
        export_df = team_weeks.merge(totals, on=["Team", "Week Start"], how="left")
    else:
        export_df = team_weeks.copy()
        export_df["Training/Mentoring Hours"] = 0.0
    export_df["Training/Mentoring Hours"] = (
        pd.to_numeric(export_df["Training/Mentoring Hours"], errors="coerce").fillna(0.0).round(2)
    )
    return export_df.loc[:, columns].sort_values(["Week Start", "Team"]).reset_index(drop=True)
def build_training_mentoring_totals_export(training_export: pd.DataFrame) -> pd.DataFrame:
    columns = ["Team", "Training/Mentoring Hours"]
    if training_export is None or training_export.empty:
        return pd.DataFrame(columns=columns)
    totals = training_export.copy()
    totals["Training/Mentoring Hours"] = pd.to_numeric(
        totals["Training/Mentoring Hours"], errors="coerce"
    ).fillna(0.0)
    totals = (
        totals.groupby("Team", as_index=False)["Training/Mentoring Hours"]
        .sum()
        .sort_values("Team")
        .reset_index(drop=True)
    )
    totals["Training/Mentoring Hours"] = totals["Training/Mentoring Hours"].round(2)
    return totals.loc[:, columns]
if page == "Overview":
    st.subheader("Summary")
    overview_team_export, overview_ou_export, overview_portfolio_export, overview_enterprise_export = _get_export_lookup_bundle(
        shared_metrics_df,
        shared_nonwip_df,
        org,
        factor_out_ooo,
        cache_key=org_cache_key,
    )
    team_lookup = overview_team_export
    ou_lookup = overview_ou_export
    portfolio_lookup = overview_portfolio_export
    enterprise_lookup = overview_enterprise_export
    if team_lookup.empty:
        st.info("No overview data available.")
    else:
        filter_card = st.container(border=True)
        with filter_card:
            st.markdown("#### Overview filters")
            control_cols = st.columns([1.15, 1.0, 1.25])
            week_options = sorted(
                [
                    pd.Timestamp(w).normalize()
                    for w in team_lookup["week_start"].dropna().unique()
                ],
                reverse=True,
            )
            if "overview_selected_weeks" in st.session_state:
                valid_weeks = set(week_options)
                st.session_state["overview_selected_weeks"] = [
                    pd.Timestamp(w).normalize()
                    for w in st.session_state["overview_selected_weeks"]
                    if pd.Timestamp(w).normalize() in valid_weeks
                ]
            @st.dialog("Choose fiscal month")
            def _overview_fiscal_month_dialog():
                fiscal_month = st.selectbox(
                    "Fiscal month",
                    options=list(FY27_FISCAL_MONTHS.keys()),
                    key="overview_fiscal_month_choice",
                )
                start_ts, end_ts = FY27_FISCAL_MONTHS[fiscal_month]
                st.caption(
                    f"{fiscal_month}: {start_ts.strftime('%Y-%m-%d')} through "
                    f"{end_ts.strftime('%Y-%m-%d')}"
                )
                if st.button("Apply", type="primary", key="overview_apply_fiscal_month"):
                    st.session_state["overview_selected_weeks"] = _weeks_between(
                        week_options,
                        start_ts,
                        end_ts,
                    )
                    st.rerun()
            shortcut_cols = control_cols[0].columns([1, 1])
            if shortcut_cols[0].button("FY27", key="overview_fy27_btn"):
                st.session_state["overview_selected_weeks"] = _fy27_weeks(week_options)
                st.rerun()
            if shortcut_cols[1].button("Fiscal month", key="overview_fiscal_month_btn"):
                _overview_fiscal_month_dialog()
            selected_weeks = control_cols[0].multiselect(
                "Weeks",
                options=week_options,
                default=week_options[:8] if len(week_options) > 8 else week_options,
                format_func=lambda x: pd.Timestamp(x).strftime("%Y-%m-%d"),
                key="overview_selected_weeks",
                placeholder="Select one or more weeks",
            )
            current_week_start = (
                pd.Timestamp.now(tz=ZoneInfo("America/Chicago"))
                .normalize()
                - pd.Timedelta(
                    days=pd.Timestamp.now(tz=ZoneInfo("America/Chicago")).weekday()
                )
            )
            selected_week_dates = {
                pd.Timestamp(w).date()
                for w in selected_weeks
            }
            if current_week_start.date() in selected_week_dates:
                st.warning("Current week is included; data may not be final yet.")
            filter_level = control_cols[1].radio(
                "Filter by",
                options=["Enterprise", "Portfolio", "OU", "Team"],
                index=0,
                horizontal=True,
                key="overview_filter_level",
            )
            if filter_level == "Enterprise":
                lookup_df = enterprise_lookup.copy()
                filter_col = "enterprise"
                label = "Enterprise"
                if filter_col not in lookup_df.columns:
                    lookup_df[filter_col] = getattr(org, "org_name", "Enterprise")
            elif filter_level == "Portfolio":
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
            lookup_df["week_start"] = pd.to_datetime(
                lookup_df["week_start"],
                errors="coerce",
            ).dt.normalize()
            selected_week_set = {
                pd.Timestamp(w).normalize()
                for w in selected_weeks
            }
            scoped_week = lookup_df[
                lookup_df["week_start"].isin(selected_week_set)
            ].copy()
            if filter_col not in scoped_week.columns:
                st.info(f"No {label} data available.")
            elif not selected_weeks:
                st.info("Select one or more weeks to view overview metrics.")
            else:
                options = sorted(
                    x
                    for x in scoped_week[filter_col].dropna().astype(str).unique()
                    if str(x).strip()
                )
                if not options:
                    st.info(f"No {label} values available for the selected weeks.")
                else:
                    selected_value = control_cols[2].selectbox(
                        label,
                        options=options,
                        index=0,
                        key=f"overview_selected_{filter_col}",
                        placeholder=f"Select {label.lower()}",
                    )
                    scoped_df = scoped_week[
                        scoped_week[filter_col].astype(str) == str(selected_value)
                    ].copy()
                    history_df = lookup_df[
                        lookup_df[filter_col].astype(str) == str(selected_value)
                    ].copy()
                    def _safe_metric(v, pct: bool = False):
                        if pd.isna(v):
                            return "—"
                        return f"{float(v):.1%}" if pct else f"{float(v):.2f}"
                    def _metric_from_df(df: pd.DataFrame, col: str, default=np.nan):
                        if df.empty or col not in df.columns:
                            return default
                        return df[col].iloc[0]
                    def _sum_col(df: pd.DataFrame, col: str) -> float:
                        if df.empty or col not in df.columns:
                            return 0.0
                        return float(
                            pd.to_numeric(df[col], errors="coerce")
                            .fillna(0.0)
                            .sum()
                        )
                    period_summary = pd.DataFrame()
                    if not scoped_df.empty:
                        totals = {
                            "completed_hours": _sum_col(scoped_df, "completed_hours"),
                            "other_team_wip_hours": _sum_col(scoped_df, "other_team_wip_hours"),
                            "non_wip_hours": _sum_col(scoped_df, "non_wip_hours"),
                            "ooo_hours": _sum_col(scoped_df, "ooo_hours"),
                            "capacity_hours": _sum_col(scoped_df, "capacity_hours"),
                            "unaccounted_hours": _sum_col(scoped_df, "unaccounted_hours"),
                        }
                        data_week_count = max(
                            scoped_df["week_start"].nunique(),
                            1,
                        )
                        pct_denom = totals["capacity_hours"]
                        period_summary = pd.DataFrame([{
                            **totals,
                            "avg_other_team_wip_hours": totals["other_team_wip_hours"] / data_week_count,
                            "avg_ooo_hours": totals["ooo_hours"] / data_week_count,
                            "avg_unaccounted_hours": totals["unaccounted_hours"] / data_week_count,
                            "wip_pct": (
                                totals["completed_hours"] / pct_denom
                                if pct_denom > 0
                                else pd.NA
                            ),
                            "other_team_wip_pct": (
                                totals["other_team_wip_hours"] / pct_denom
                                if pct_denom > 0
                                else pd.NA
                            ),
                            "non_wip_pct": (
                                totals["non_wip_hours"] / pct_denom
                                if pct_denom > 0
                                else pd.NA
                            ),
                            "ooo_pct": (
                                0.0
                                if factor_out_ooo
                                else (
                                    totals["ooo_hours"] / pct_denom
                                    if pct_denom > 0
                                    else pd.NA
                                )
                            ),
                            "unaccounted_pct": (
                                totals["unaccounted_hours"] / pct_denom
                                if pct_denom > 0
                                else pd.NA
                            ),
                        }])
                        period_summary = _add_avg_hours_day_columns(period_summary)
                    metric_df = period_summary if not period_summary.empty else scoped_df
                    warning_text = ""
                    if "warning" in scoped_df.columns and not scoped_df.empty:
                        warnings = [
                            str(x).strip()
                            for x in scoped_df["warning"].dropna().tolist()
                            if str(x).strip()
                        ]
                        warning_text = "; ".join(sorted(set(warnings)))
                    if warning_text:
                        st.warning(warning_text)
                    history_df["week_start"] = pd.to_datetime(
                        history_df["week_start"],
                        errors="coerce",
                    ).dt.normalize()
                    history_df["week_start"] = pd.to_datetime(
                        history_df["week_start"],
                        errors="coerce",
                    ).dt.normalize()
                    overview_trend_start = pd.Timestamp("2026-01-01").normalize()
                    history_df = (
                        history_df
                        .dropna(subset=["week_start"])
                        .loc[lambda d: d["week_start"] >= overview_trend_start]
                        .sort_values("week_start")
                    )
                    unaccounted_val = _metric_from_df(
                        metric_df,
                        "unaccounted_pct",
                        0.0,
                    )
                    try:
                        unaccounted_val = (
                            float(unaccounted_val)
                            if pd.notna(unaccounted_val)
                            else 0.0
                        )
                    except Exception:
                        unaccounted_val = 0.0
                    if unaccounted_val > 0.25:
                        st.error("❗ Unaccounted hours are high for this selection (>25%).")
                    st.markdown(
                        """
                        <style>
                        div[data-testid="stMetric"]{ text-align: center; }
                        label[data-testid="stMetricLabel"]{ display: block; width: 100%; text-align: center; margin: 0; }
                        label[data-testid="stMetricLabel"] p{ text-align: center !important; margin: 0 !important; }
                        div[data-testid="stMetricValue"]{ text-align: center !important; width: 100%; }
                        </style>
                        """,
                        unsafe_allow_html=True,
                    )
                    selected_week_count = len(selected_week_set)
                    data_week_count = (
                        scoped_df["week_start"].nunique()
                        if not scoped_df.empty and "week_start" in scoped_df.columns
                        else 0
                    )
                    st.caption(
                        f"Metric cards summarize {data_week_count} week(s)"
                    )
                    _, c1, c2, _ = st.columns([1.2, 1.2, 1.2, 1.2])
                    c1.metric(
                        "Avg Per Person **WIP** Daily Hours",
                        _safe_metric(_metric_from_df(metric_df, "wip_avg_hours_day")),
                    )
                    c2.metric(
                        "Avg Per Person **Non-WIP** Daily Hours",
                        _safe_metric(_metric_from_df(metric_df, "non_wip_avg_hours_day")),
                    )
                    _, p1, p2, _ = st.columns([1.2, 1.2, 1.2, 1.2])
                    p1.metric(
                        "**WIP** Ratio",
                        _safe_metric(_metric_from_df(metric_df, "wip_pct"), pct=True),
                    )
                    p2.metric(
                        "**Non-WIP** Ratio",
                        _safe_metric(_metric_from_df(metric_df, "non_wip_pct"), pct=True),
                    )
                    st.divider()
                    _, _, c3, c4, c5, _, _ = st.columns(
                        [.8, .8, 1.2, 1.2, 1.2, 1.0, 0.5]
                    )
                    c4.metric(
                        "Avg **OOO** Weekly Hours",
                        _safe_metric(_metric_from_df(metric_df, "avg_ooo_hours")),
                    )
                    c3.metric(
                        "Avg **Other Team WIP** Weekly Hours",
                        _safe_metric(_metric_from_df(metric_df, "avg_other_team_wip_hours", 0.0)),
                    )
                    c5.metric(
                        "Avg **Unaccounted** Weekly Hours",
                        _safe_metric(_metric_from_df(metric_df, "avg_unaccounted_hours")),
                    )
                    _, _, p3, p4, p5, _, _ = st.columns(
                        [.8, .8, 1.2, 1.2, 1.2, 1.0, 0.5]
                    )
                    p4.metric(
                        "**OOO** % of week",
                        _safe_metric(_metric_from_df(metric_df, "ooo_pct"), pct=True),
                    )
                    p3.metric(
                        "**Other Team WIP** % of week",
                        _safe_metric(_metric_from_df(metric_df, "other_team_wip_pct", 0.0), pct=True),
                    )
                    p5.metric(
                        "**Unaccounted** % remaining",
                        _safe_metric(_metric_from_df(metric_df, "unaccounted_pct"), pct=True),
                    )
                    st.divider()
                    st.subheader(f"{label} WIP % trend")
                    if history_df.empty:
                        st.info("No historical WIP % data available for this selection.")
                    else:
                        chart_df = history_df.loc[:, ["week_start", "wip_pct"]].copy()
                        chart_df["week_start"] = pd.to_datetime(
                            chart_df["week_start"],
                            errors="coerce",
                        )
                        chart_df["wip_pct"] = pd.to_numeric(
                            chart_df["wip_pct"],
                            errors="coerce",
                        )
                        chart_df = (
                            chart_df
                            .dropna(subset=["week_start", "wip_pct"])
                            .sort_values("week_start")
                        )
                        if chart_df.empty:
                            st.info("No historical WIP % data available for this selection.")
                        else:
                            chart_df["week_label"] = chart_df["week_start"].dt.strftime("%Y-%m-%d")
                            chart_df["wip_label"] = chart_df["wip_pct"].map(lambda v: f"{v:.1%}")
                            latest_week = chart_df["week_start"].max()
                            latest_df = chart_df[
                                chart_df["week_start"] == latest_week
                            ].copy()
                            base = alt.Chart(chart_df).encode(
                                x=alt.X(
                                    "week_start:T",
                                    title="Week of",
                                    axis=alt.Axis(
                                        format="%Y-%m-%d",
                                        labelAngle=-35,
                                        tickCount="week",
                                    ),
                                    timeUnit="yearmonthdate",
                                ),
                                y=alt.Y(
                                    "wip_pct:Q",
                                    title="WIP %",
                                    axis=alt.Axis(format=".0%"),
                                    scale=alt.Scale(
                                        domain=[
                                            0,
                                            max(
                                                1.0,
                                                float(chart_df["wip_pct"].max()) * 1.1,
                                            ),
                                        ]
                                    ),
                                ),
                                tooltip=[
                                    alt.Tooltip("week_label:N", title="Week"),
                                    alt.Tooltip("wip_pct:Q", title="WIP %", format=".1%"),
                                ],
                            )
                            line = base.mark_line(strokeWidth=3).encode()
                            points = base.mark_circle(size=80).encode()
                            latest_point = alt.Chart(latest_df).mark_circle(
                                size=180,
                                filled=True,
                            ).encode(
                                x="week_start:T",
                                y="wip_pct:Q",
                                tooltip=[
                                    alt.Tooltip("week_label:N", title="Week"),
                                    alt.Tooltip("wip_pct:Q", title="WIP %", format=".1%"),
                                ],
                            )
                            data_labels = base.mark_text(
                                align="center",
                                dy=-12,
                                fontSize=11,
                            ).encode(
                                text="wip_label:N",
                            )
                            rule = alt.Chart(
                                pd.DataFrame({"y": [0.80]})
                            ).mark_rule(
                                strokeDash=[6, 4]
                            ).encode(
                                y="y:Q"
                            )
                            trend_chart = (
                                alt.layer(rule, line, points, latest_point, data_labels)
                                .properties(height=360)
                                .interactive()
                            )
                            st.altair_chart(trend_chart, width="stretch")
elif page == "Non-WIP":
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
    team_meta = _team_meta_lookup(org).copy()
    if not team_meta.empty and "team" in source_raw.columns:
        source_raw = source_raw.copy()
        source_raw["team"] = source_raw["team"].astype(str).str.strip()
        team_meta["team"] = team_meta["team"].astype(str).str.strip()
        source_raw = source_raw.merge(
            team_meta[["team", "portfolio"]].rename(columns={"portfolio": "_portfolio"}),
            on="team",
            how="left",
        )
        portfolio_options = sorted(
            p
            for p in source_raw["_portfolio"].dropna().astype(str).unique()
            if p.strip()
        )
        selected_portfolio = st.selectbox(
            "Portfolio",
            options=["All portfolios", *portfolio_options],
            index=0,
            key="nonwip_portfolio_filter",
            help="Filters only the Non-WIP page.",
        )
        if selected_portfolio != "All portfolios":
            source_raw = source_raw[
                source_raw["_portfolio"].astype(str).eq(selected_portfolio)
            ].copy()
        if source_raw.empty:
            st.info("No Non-WIP activity data available for the selected portfolio.")
            st.stop()
    parsed_nonwip = _prepare_nonwip_activity_source(source_raw)
    top_n = st.number_input(
        "Number of activities to show",
        min_value=1,
        max_value=200,
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
    label_map = (
        rolled.assign(activity_len=rolled["activity"].astype(str).str.len())
        .sort_values(["activity_norm", "activity_len", "activity"], ascending=[True, False, True])
        .drop_duplicates(subset=["activity_norm"])
        .loc[:, ["activity_norm", "activity"]]
        .rename(columns={"activity": "display_activity"})
    )
    weekly_by_activity = (
        rolled.groupby(["week_start", "activity_norm"], as_index=False)
        .agg(hours=("hours", "sum"))
        .merge(label_map, on="activity_norm", how="left")
        .sort_values(["week_start", "hours"], ascending=[True, False])
    )
    total_hours = (
        weekly_by_activity.groupby(["activity_norm", "display_activity"], as_index=False)
        .agg(total_hours=("hours", "sum"))
        .sort_values("total_hours", ascending=False)
        .head(int(top_n))
        .reset_index(drop=True)
        .rename(columns={"display_activity": "activity"})
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
    st.divider()
    training_export = build_training_mentoring_export(
        source_raw,
        before_date=datetime.now(ZoneInfo("America/Chicago")).date(),
    )
    training_totals_export = build_training_mentoring_totals_export(training_export)
    training_export_bytes = _cached_custom_excel_bytes(
        (
            ("Team Weekly", training_export),
            ("Team Totals", training_totals_export),
        )
    )
    st.download_button(
        label="Export Training",
        data=training_export_bytes,
        file_name="enterprise_training_mentoring_by_team.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="export_training_mentoring",
        help=(
            "Exports each team's combined Training/Mentoring hours for every week before "
            "today in the selected portfolio, plus a Team Totals tab that sums all weeks."
        ),
    )
elif page == "Export":
    st.subheader("Export")
    team_export, ou_export, portfolio_export, enterprise_export = _get_export_lookup_bundle(
        shared_metrics_df, shared_nonwip_df, org, factor_out_ooo,
        cache_key=org_cache_key,
    )
    export_scope_df = enterprise_export.copy()
    export_filter_col = None
    export_filter_label = "Enterprise"
    export_filter_level = "Enterprise"
    export_selected_weeks = []
    export_selected_values = []
    if not team_export.empty:
        export_filter_card = st.container(border=True)
        with export_filter_card:
            st.markdown("#### Export filters")
            export_cols = st.columns([1.15, 1.0, 1.4])
            export_week_options = sorted(
                team_export["week_start"].dropna().unique(),
                reverse=True,
            )
            export_selected_weeks = export_cols[0].multiselect(
                "Weeks",
                options=export_week_options,
                default=export_week_options[:8] if len(export_week_options) > 8 else export_week_options,
                format_func=lambda x: pd.Timestamp(x).strftime("%Y-%m-%d"),
                key="export_selected_weeks",
                placeholder="Select one or more weeks",
            )
            export_filter_level = export_cols[1].radio(
                "Filter by",
                options=["Enterprise", "Portfolio", "OU", "Team"],
                index=0,
                horizontal=True,
                key="export_filter_level",
            )
            if export_filter_level == "Enterprise":
                export_scope_df = enterprise_export.copy()
                export_filter_col = None
                export_filter_label = "Enterprise"
            elif export_filter_level == "Portfolio":
                export_scope_df = portfolio_export.copy()
                export_filter_col = "portfolio"
                export_filter_label = "Portfolio"
            elif export_filter_level == "OU":
                export_scope_df = ou_export.copy()
                export_filter_col = "ou"
                export_filter_label = "OU"
            else:
                export_scope_df = team_export.copy()
                export_filter_col = "team"
                export_filter_label = "Team"
            export_scope_df["week_start"] = pd.to_datetime(
                export_scope_df["week_start"], errors="coerce"
            ).dt.normalize()
            selected_week_set = {
                pd.Timestamp(x).normalize()
                for x in export_selected_weeks
            }
            export_scoped_weeks = export_scope_df[
                export_scope_df["week_start"].isin(selected_week_set)
            ].copy()
            if export_filter_level == "Enterprise":
                export_options = ["Enterprise"] if not export_scoped_weeks.empty else []
            else:
                export_options = sorted(
                    x
                    for x in export_scoped_weeks[export_filter_col].dropna().astype(str).unique()
                    if str(x).strip()
                )
            export_selected_values = export_cols[2].multiselect(
                export_filter_label,
                options=export_options,
                default=export_options,
                key=f"export_selected_{export_filter_level.lower()}",
                placeholder=f"Select one or more {export_filter_label.lower()} values",
            )
        if export_selected_weeks and export_selected_values:
            export_scope_df = export_scope_df[
                export_scope_df["week_start"].isin(selected_week_set)
            ].copy()
            if export_filter_col is not None:
                export_scope_df = export_scope_df[
                    export_scope_df[export_filter_col].astype(str).isin(export_selected_values)
                ].copy()
        else:
            export_scope_df = export_scope_df.iloc[0:0].copy()
    missing_teams_display = _build_missing_team_weeks_df(
        team_export=team_export,
        org=org,
        selected_weeks=export_selected_weeks,
    )
    def _tree_safe_str(value: Any) -> str:
        if pd.isna(value):
            return ""
        return str(value).strip()
    def _tree_week_value(value: Any) -> str:
        if pd.isna(value):
            return ""
        return pd.Timestamp(value).strftime("%Y-%m-%d")
    def _rollup_row_key(level: str, row: pd.Series) -> str:
        week = _tree_week_value(row.get("week_start"))
        if level == "Enterprise":
            name = "Enterprise"
        elif level == "Portfolio":
            name = _tree_safe_str(row.get("portfolio"))
        elif level == "OU":
            name = f"{_tree_safe_str(row.get('portfolio'))}::{_tree_safe_str(row.get('ou'))}"
        else:
            name = _tree_safe_str(row.get("team"))
        return f"{level}|{week}|{name}"
    def _rows_for_week(df: pd.DataFrame, week_value: Any) -> pd.DataFrame:
        if df is None or df.empty or "week_start" not in df.columns:
            return pd.DataFrame()
        out = df.copy()
        out["week_start"] = pd.to_datetime(out["week_start"], errors="coerce").dt.normalize()
        week_value = pd.Timestamp(week_value).normalize()
        return out[out["week_start"] == week_value].copy()
    def _filter_value(df: pd.DataFrame, col: str, value: Any) -> pd.DataFrame:
        if df is None or df.empty or col not in df.columns or pd.isna(value):
            return pd.DataFrame()
        return df[df[col].astype(str) == str(value)].copy()
    def _tree_rollup_label(
        level: str,
        row: pd.Series,
        depth: int,
        has_children: bool,
        is_open: bool,
    ) -> str:
        icon = ""
        if has_children:
            icon = "▾ " if is_open else "▸ "
        indent = " " * depth
        if level == "Enterprise":
            label = "Enterprise"
        elif level == "Portfolio":
            label = _tree_safe_str(row.get("portfolio"))
        elif level == "OU":
            label = _tree_safe_str(row.get("ou"))
        else:
            label = _tree_safe_str(row.get("team"))
        return f"{indent}{icon}{label}"
    def _tree_alert(row: pd.Series) -> str:
        raw_unaccounted = row.get("unaccounted_pct", 0)
        try:
            return "❗" if pd.notna(raw_unaccounted) and float(raw_unaccounted) > 0.25 else ""
        except Exception:
            return ""
    def _append_tree_row(
        rows: list[dict],
        row: pd.Series,
        level: str,
        depth: int,
        has_children: bool,
        open_keys: set[str],
    ) -> None:
        key = _rollup_row_key(level, row)
        is_open = key in open_keys
        raw_over_hours = pd.to_numeric(row.get("over_hours", 0), errors="coerce")
        is_over_hours = bool(pd.notna(raw_over_hours) and float(raw_over_hours) > 0)
        rows.append({
            "_row_key": key,
            "_has_children": bool(has_children),
            "_level_depth": int(depth),
            "_is_over_hours": is_over_hours,
            "Open": bool(is_open) if has_children else False,
            "Level": level,
            "Roll-up": _tree_rollup_label(level, row, depth, has_children, is_open),
            "Alert": _tree_alert(row),
            "Week Start": pd.Timestamp(row.get("week_start")).date() if pd.notna(row.get("week_start")) else pd.NaT,
            "Portfolio": row.get("portfolio", ""),
            "OU": row.get("ou", ""),
            "Team": row.get("team", ""),
            "Capacity": row.get("capacity_hours"),
            "People": row.get("people_count"),
            "Completed Hours": row.get("completed_hours"),
            "WIP %": row.get("wip_pct"),
            "Other Team WIP": row.get("other_team_wip_hours"),
            "Other Team WIP %": row.get("other_team_wip_pct"),
            "Non-WIP Hours": row.get("non_wip_hours"),
            "Non-WIP %": row.get("non_wip_pct"),
            "OOO Hours": row.get("ooo_hours"),
            "OOO %": row.get("ooo_pct"),
            "Unaccounted Hours": row.get("unaccounted_hours"),
            "Unaccounted %": row.get("unaccounted_pct"),
            "Over Hours": row.get("over_hours"),
            "Warning": row.get("warning", ""),
        })
    def _build_export_tree_table(export_scope_df: pd.DataFrame) -> pd.DataFrame:
        open_keys = st.session_state.setdefault("export_rollup_open_keys", set())
        rows: list[dict] = []
        if export_scope_df is None or export_scope_df.empty:
            return pd.DataFrame()
        scoped = export_scope_df.copy()
        scoped["week_start"] = pd.to_datetime(scoped["week_start"], errors="coerce").dt.normalize()
        scoped = scoped.dropna(subset=["week_start"]).copy()
        if export_filter_level == "Enterprise":
            root_rows = scoped.sort_values(["week_start"]).reset_index(drop=True)
            for _, ent_row in root_rows.iterrows():
                wk = ent_row["week_start"]
                ent_key = _rollup_row_key("Enterprise", ent_row)
                _append_tree_row(rows, ent_row, "Enterprise", 0, True, open_keys)
                if ent_key not in open_keys:
                    continue
                portfolios = _rows_for_week(portfolio_export, wk)
                if not portfolios.empty:
                    portfolios = portfolios.sort_values(["portfolio"], na_position="last")
                for _, port_row in portfolios.iterrows():
                    port_key = _rollup_row_key("Portfolio", port_row)
                    _append_tree_row(rows, port_row, "Portfolio", 1, True, open_keys)
                    if port_key not in open_keys:
                        continue
                    ous = _rows_for_week(ou_export, wk)
                    ous = _filter_value(ous, "portfolio", port_row.get("portfolio"))
                    if not ous.empty:
                        ous = ous.sort_values(["ou"], na_position="last")
                    for _, ou_row in ous.iterrows():
                        ou_key = _rollup_row_key("OU", ou_row)
                        _append_tree_row(rows, ou_row, "OU", 2, True, open_keys)
                        if ou_key not in open_keys:
                            continue
                        teams = _rows_for_week(team_export, wk)
                        teams = _filter_value(teams, "portfolio", ou_row.get("portfolio"))
                        teams = _filter_value(teams, "ou", ou_row.get("ou"))
                        if not teams.empty:
                            teams = teams.sort_values(["team"], na_position="last")
                        for _, team_row in teams.iterrows():
                            _append_tree_row(rows, team_row, "Team", 3, False, open_keys)
        elif export_filter_level == "Portfolio":
            root_rows = scoped.sort_values(["week_start", "portfolio"], na_position="last").reset_index(drop=True)
            for _, port_row in root_rows.iterrows():
                wk = port_row["week_start"]
                port_key = _rollup_row_key("Portfolio", port_row)
                _append_tree_row(rows, port_row, "Portfolio", 0, True, open_keys)
                if port_key not in open_keys:
                    continue
                ous = _rows_for_week(ou_export, wk)
                ous = _filter_value(ous, "portfolio", port_row.get("portfolio"))
                if not ous.empty:
                    ous = ous.sort_values(["ou"], na_position="last")
                for _, ou_row in ous.iterrows():
                    ou_key = _rollup_row_key("OU", ou_row)
                    _append_tree_row(rows, ou_row, "OU", 1, True, open_keys)
                    if ou_key not in open_keys:
                        continue
                    teams = _rows_for_week(team_export, wk)
                    teams = _filter_value(teams, "portfolio", ou_row.get("portfolio"))
                    teams = _filter_value(teams, "ou", ou_row.get("ou"))
                    if not teams.empty:
                        teams = teams.sort_values(["team"], na_position="last")
                    for _, team_row in teams.iterrows():
                        _append_tree_row(rows, team_row, "Team", 2, False, open_keys)
        elif export_filter_level == "OU":
            root_rows = scoped.sort_values(["week_start", "portfolio", "ou"], na_position="last").reset_index(drop=True)
            for _, ou_row in root_rows.iterrows():
                wk = ou_row["week_start"]
                ou_key = _rollup_row_key("OU", ou_row)
                _append_tree_row(rows, ou_row, "OU", 0, True, open_keys)
                if ou_key not in open_keys:
                    continue
                teams = _rows_for_week(team_export, wk)
                teams = _filter_value(teams, "portfolio", ou_row.get("portfolio"))
                teams = _filter_value(teams, "ou", ou_row.get("ou"))
                if not teams.empty:
                    teams = teams.sort_values(["team"], na_position="last")
                for _, team_row in teams.iterrows():
                    _append_tree_row(rows, team_row, "Team", 1, False, open_keys)
        else:
            root_rows = scoped.sort_values(
                ["week_start", "portfolio", "ou", "team"],
                na_position="last",
            ).reset_index(drop=True)
            for _, team_row in root_rows.iterrows():
                _append_tree_row(rows, team_row, "Team", 0, False, open_keys)
        return pd.DataFrame(rows)
    def _format_tree_for_display(df: pd.DataFrame) -> pd.DataFrame:
        out = df.copy()
        for col in [
            "Capacity",
            "People",
            "Completed Hours",
            "Other Team WIP",
            "Non-WIP Hours",
            "OOO Hours",
            "Unaccounted Hours",
            "Over Hours",
        ]:
            if col in out.columns:
                out[col] = pd.to_numeric(out[col], errors="coerce")
        for col in [
            "WIP %",
            "Other Team WIP %",
            "Non-WIP %",
            "OOO %",
            "Unaccounted %",
        ]:
            if col in out.columns:
                out[col] = pd.to_numeric(out[col], errors="coerce") * 100
        return out
    def _normalize_tree_weeks(df: pd.DataFrame) -> pd.DataFrame:
        if df is None or df.empty or "week_start" not in df.columns:
            return pd.DataFrame()
        out = df.copy()
        out["week_start"] = pd.to_datetime(out["week_start"], errors="coerce").dt.normalize()
        return out.dropna(subset=["week_start"]).copy()
    def _all_expandable_tree_keys(export_scope_df: pd.DataFrame) -> set[str]:
        keys: set[str] = set()
        scoped = _normalize_tree_weeks(export_scope_df)
        if scoped.empty:
            return keys
        def _add_keys(df: pd.DataFrame, level: str) -> None:
            if df is None or df.empty:
                return
            for _, r in df.iterrows():
                keys.add(_rollup_row_key(level, r))
        if export_filter_level == "Enterprise":
            _add_keys(scoped, "Enterprise")
            weeks = scoped[["week_start"]].drop_duplicates()
            portfolios = _normalize_tree_weeks(portfolio_export)
            if not portfolios.empty:
                portfolios = portfolios.merge(weeks, on="week_start", how="inner")
            _add_keys(portfolios, "Portfolio")
            ous = _normalize_tree_weeks(ou_export)
            if not ous.empty:
                ous = ous.merge(weeks, on="week_start", how="inner")
            _add_keys(ous, "OU")
        elif export_filter_level == "Portfolio":
            _add_keys(scoped, "Portfolio")
            scope_pairs = scoped[["week_start", "portfolio"]].drop_duplicates()
            ous = _normalize_tree_weeks(ou_export)
            if not ous.empty:
                ous = ous.merge(scope_pairs, on=["week_start", "portfolio"], how="inner")
            _add_keys(ous, "OU")
        elif export_filter_level == "OU":
            _add_keys(scoped, "OU")
        return keys
    def _style_export_tree_row(row: pd.Series) -> list[str]:
        level = str(row.get("Level", "")).strip()
        is_over = bool(row.get("_is_over_hours", False))
        if is_over:
            base_style = "background-color: #fef08a; color: #713f12; font-weight: 600;"
        else:
            level_backgrounds = {
                "Enterprise": "#ffffff",
                "Portfolio": "#f8fafc",
                "OU": "#eef2f7",
                "Team": "#e5e7eb",
            }
            bg = level_backgrounds.get(level, "")
            base_style = f"background-color: {bg};" if bg else ""
        styles = [base_style] * len(row.index)
        if "WIP %" in row.index:
            idx = row.index.get_loc("WIP %")
            styles[idx] = _threshold_cell_style(
                row.get("WIP %"),
                threshold=80.0,
                good_if_gte=True,
            )
        if "Non-WIP %" in row.index:
            idx = row.index.get_loc("Non-WIP %")
            styles[idx] = _threshold_cell_style(
                row.get("Non-WIP %"),
                threshold=20.0,
                good_if_gte=False,
            )
        return styles
    if team_export.empty:
        st.info("No exportable team/week data found.")
    else:
        if export_scope_df.empty:
            st.info("No export rows match the selected filters.")
        else:
            st.markdown("#### Weekly roll-up")
            action_cols = st.columns([1, 1, 5])
            with action_cols[0]:
                if st.button("Collapse all", key="export_tree_collapse_all"):
                    st.session_state["export_rollup_open_keys"] = set()
                    st.rerun()
            with action_cols[1]:
                if st.button("Expand all", key="export_tree_expand_all"):
                    all_parent_keys = _all_expandable_tree_keys(export_scope_df)
                    st.session_state["export_rollup_open_keys"] = (
                        set(st.session_state.get("export_rollup_open_keys", set()))
                        | all_parent_keys
                    )
                    st.rerun()
            tree_df = _build_export_tree_table(export_scope_df)
            if tree_df.empty:
                st.info("No export rows match the selected filters.")
            else:
                display_tree_df = _format_tree_for_display(tree_df)
                styled_display_tree_df = display_tree_df.style.apply(
                    _style_export_tree_row,
                    axis=1,
                )
                edited_tree = st.data_editor(
                    styled_display_tree_df,
                    width="stretch",
                    hide_index=True,
                    num_rows="fixed",
                    disabled=[
                        c
                        for c in display_tree_df.columns
                        if c != "Open"
                    ],
                    column_config={
                        "_row_key": None,
                        "_has_children": None,
                        "_level_depth": None,
                        "_is_over_hours": None,
                        "Portfolio": None,
                        "OU": None,
                        "Team": None,
                        "Open": st.column_config.CheckboxColumn(
                            "Open",
                            help="Check to expand this roll-up row. Uncheck to collapse it.",
                            width=None,
                        ),
                        "Level": st.column_config.TextColumn(
                            "Level",
                            width=None,
                        ),
                        "Roll-up": st.column_config.TextColumn(
                            "Roll-up",
                            width=None,
                        ),
                        "Alert": st.column_config.TextColumn(
                            "Alert",
                            width=None,
                        ),
                        "Week Start": st.column_config.DateColumn(
                            "Week Start",
                            width=None,
                        ),
                        "Capacity": st.column_config.NumberColumn(
                            "Capacity",
                            format="%.2f",
                            width=None,
                        ),
                        "People": st.column_config.NumberColumn(
                            "People",
                            format="%.2f",
                            width=None,
                        ),
                        "Completed Hours": st.column_config.NumberColumn(
                            "WIP Hours",
                            format="%.2f",
                            width=None,
                        ),
                        "WIP %": st.column_config.NumberColumn(
                            "WIP %",
                            format="%.1f%%",
                            width=None,
                        ),
                        "Other Team WIP": st.column_config.NumberColumn(
                            "Other Team WIP",
                            format="%.2f",
                            width=None,
                        ),
                        "Other Team WIP %": st.column_config.NumberColumn(
                            "Other %",
                            format="%.1f%%",
                            width=None,
                        ),
                        "Non-WIP Hours": st.column_config.NumberColumn(
                            "Non-WIP Hours",
                            format="%.2f",
                            width=None,
                        ),
                        "Non-WIP %": st.column_config.NumberColumn(
                            "Non-WIP %",
                            format="%.1f%%",
                            width=None,
                        ),
                        "OOO Hours": st.column_config.NumberColumn(
                            "OOO Hours",
                            format="%.2f",
                            width=None,
                        ),
                        "OOO %": st.column_config.NumberColumn(
                            "OOO %",
                            format="%.1f%%",
                            width=None,
                        ),
                        "Unaccounted Hours": st.column_config.NumberColumn(
                            "Missing Hours",
                            format="%.2f",
                            width=None,
                        ),
                        "Unaccounted %": st.column_config.NumberColumn(
                            "Missing %",
                            format="%.1f%%",
                            width=None,
                        ),
                        "Over Hours": None,
                        "Warning": st.column_config.TextColumn(
                            "Warning",
                            width=None,
                        ),
                    },
                    key="export_rollup_tree_editor",
                )
                current_open_keys = set(st.session_state.get("export_rollup_open_keys", set()))
                visible_parent_keys = set(
                    edited_tree.loc[
                        edited_tree["_has_children"].fillna(False).astype(bool),
                        "_row_key",
                    ].astype(str)
                )
                checked_visible_keys = set(
                    edited_tree.loc[
                        edited_tree["Open"].fillna(False).astype(bool)
                        & edited_tree["_has_children"].fillna(False).astype(bool),
                        "_row_key",
                    ].astype(str)
                )
                next_open_keys = (
                    current_open_keys - visible_parent_keys
                ) | checked_visible_keys
                if next_open_keys != current_open_keys:
                    st.session_state["export_rollup_open_keys"] = next_open_keys
                    st.rerun()
            if export_filter_level == "Enterprise":
                team_export_display = _display_export_team_df(team_export.iloc[0:0].copy())
                ou_export_display = _display_export_ou_df(ou_export.iloc[0:0].copy())
                portfolio_export_display = _display_export_portfolio_df(portfolio_export.iloc[0:0].copy())
                enterprise_export_display = _display_export_enterprise_df(export_scope_df)
            elif export_filter_level == "Team":
                team_export_display = _display_export_team_df(export_scope_df)
                ou_export_display = _display_export_ou_df(ou_export.iloc[0:0].copy())
                portfolio_export_display = _display_export_portfolio_df(portfolio_export.iloc[0:0].copy())
                enterprise_export_display = _display_export_enterprise_df(enterprise_export.iloc[0:0].copy())
            elif export_filter_level == "OU":
                team_export_display = _display_export_team_df(team_export.iloc[0:0].copy())
                ou_export_display = _display_export_ou_df(export_scope_df)
                portfolio_export_display = _display_export_portfolio_df(portfolio_export.iloc[0:0].copy())
                enterprise_export_display = _display_export_enterprise_df(enterprise_export.iloc[0:0].copy())
            else:
                team_export_display = _display_export_team_df(team_export.iloc[0:0].copy())
                ou_export_display = _display_export_ou_df(ou_export.iloc[0:0].copy())
                portfolio_export_display = _display_export_portfolio_df(export_scope_df)
                enterprise_export_display = _display_export_enterprise_df(enterprise_export.iloc[0:0].copy())
            st.markdown("#### Teams missing weekly data")
            if missing_teams_display.empty:
                st.success("No missing teams for the selected week(s).")
            else:
                st.dataframe(
                    missing_teams_display,
                    width="stretch",
                    hide_index=True,
                )
            if st.button("Prepare Excel export", key="prepare_enterprise_excel_export"):
                try:
                    st.session_state["enterprise_excel_export_bytes"] = _cached_excel_bytes(
                        team_export_display,
                        ou_export_display,
                        portfolio_export_display,
                        enterprise_export_display,
                        missing_teams_display,
                    )
                except Exception as e:
                    st.session_state.pop("enterprise_excel_export_bytes", None)
                    st.error(f"Excel export failed: {e}")
            xlsx_bytes = st.session_state.get("enterprise_excel_export_bytes")
            if xlsx_bytes:
                st.download_button(
                    label="Download Excel export",
                    data=xlsx_bytes,
                    file_name="enterprise_weekly_export.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.caption("Excel bytes are generated only when requested so the dashboard can render faster.")