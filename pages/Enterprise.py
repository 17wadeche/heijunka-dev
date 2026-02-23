# pages/Enterprise.py
from __future__ import annotations
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import pandas as pd
import streamlit as st
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
        "Timeliness": repo_root / "Timeliness.csv",  # allow either casing
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
def _to_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce")
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
def _workdays_per_week_assumption() -> int:
    return 5
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
    team_options = [t.name for t in teams_after_ou]
    default_teams = [t for t in enabled_team_names if t in team_options]
    if not default_teams and team_options:
        default_teams = team_options
    team_filter = st.multiselect(
        "Teams",
        options=team_options,
        default=default_teams,
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
) -> tuple[Optional[pd.Timestamp], Optional[pd.Timestamp]]:
    if df is None or df.empty:
        return None, None
    mn, mx, _ = _date_bounds_for_df(df)
    if mn is None or mx is None:
        st.info("No date column detected for this section.")
        return None, None
    import datetime
    min_d = mn.date()
    max_d = mx.date()
    today_d = datetime.date.today()
    anchor_end = min(max(today_d, min_d), max_d)
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
        start_default = max(min_d, (pd.to_datetime(anchor_end) - pd.Timedelta(days=days_map[preset])).date())
        end_default = anchor_end
    else:
        start_default = min_d
        end_default = max_d
    prev = st.session_state.get(last_preset_key)
    if prev != preset:
        st.session_state[dates_key] = (start_default, end_default)
        st.session_state[last_preset_key] = preset
        st.rerun()
    dr = st.date_input(
        label,
        min_value=min_d,
        max_value=max_d,
        key=dates_key,
        help="Filters only this section.",
    )
    if isinstance(dr, tuple) and len(dr) == 2:
        start_d, end_d = dr
    else:
        start_d, end_d = start_default, end_default
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
tabs = st.tabs(["Overview", "Non-WIP"])
def _get_metrics_df() -> Optional[pd.DataFrame]:
    if "metrics" in data:
        d = filter_by_team(data["metrics"])
        if not d.empty:
            return d
    if "metrics_aggregate_dev" in data:
        d = filter_by_team(data["metrics_aggregate_dev"])
        if not d.empty:
            return d
    return None
def _get_nonwip_df() -> Optional[pd.DataFrame]:
    if "non_wip" in data:
        d = filter_by_team(data["non_wip"])
        if not d.empty:
            return d
    if "non_wip_activities" in data:
        d = filter_by_team(data["non_wip_activities"])
        if not d.empty:
            return d
    return None
def _metrics_cols(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    return {
        "date": _get_date_col(df),
        "wip_hours": _first_col(df, ["completed_hours", "wip_hours", "completedhours"]),
        "avail_hours": _first_col(df, ["total_available_hours", "person_hours", "total_available_hours."]),
        "hc_used": _first_col(df, ["actual_hc_used", "actual_hc_used."]),
        "hc_in_wip": _first_col(df, ["hc_in_wip"]),
        "people_in_wip": _first_col(df, ["people_in_wip"]),
        "person_hours_json": _first_col(df, ["person_hours"]),
    }
def _nonwip_cols(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    return {
        "date": _get_date_col(df),
        "total_nonwip": _first_col(df, ["total_non-wip_hours", "total_non_wip_hours", "total_non_wip_hours."]),
        "people_count": _first_col(df, ["people_count", "people_count."]),
        "pct_in_wip": _first_col(df, ["%_in_wip", "%_in_wip."]),
        "ooohours": _first_col(df, ["ooo_hours", "ooo_hours."]),
        "nonwip_by_person_json": _first_col(df, ["non-wip_by_person", "non_wip_by_person", "non_wip_by_person."]),
        "wip_workers_json": _first_col(df, ["wip_workers", "wip_workers."]),
        "wip_workers_count": _first_col(df, ["wip_workers_count", "wip_workers_count."]),
        "wip_workers_ooo": _first_col(df, ["wip_workers_ooo_hours", "wip_workers_ooo_hours."]),
        "activities_json": _first_col(df, ["non-wip_activities", "non_wip_activities"]),
    }
def _people_lookup_by_week(dfnw: pd.DataFrame) -> Dict[pd.Timestamp, float]:
    out: Dict[pd.Timestamp, float] = {}
    nwc = _nonwip_cols(dfnw)
    if not (nwc["date"] and nwc["people_count"]):
        return out
    tmp = dfnw.copy()
    tmp[nwc["date"]] = pd.to_datetime(tmp[nwc["date"]], errors="coerce")
    tmp = tmp.dropna(subset=[nwc["date"]])
    tmp["people_count"] = _to_num(tmp[nwc["people_count"]]).fillna(0.0)
    grp = tmp.groupby(nwc["date"], as_index=False)["people_count"].sum()
    for _, r in grp.iterrows():
        out[pd.to_datetime(r[nwc["date"]])] = float(r["people_count"])
    return out
def _weekly_rollup_summary(
    dfm: Optional[pd.DataFrame],
    dfnw: Optional[pd.DataFrame],
    dfnw_act: Optional[pd.DataFrame],
    denom_mode: str,
    wd: int,
) -> Tuple[
    Optional[float], Optional[float], Optional[float], Optional[float],  # avg daily wip pp, avg daily nonwip pp, avg weekly ooo, avg weekly unacct
    Optional[float], Optional[float], Optional[float], Optional[float],  # pct_wip, pct_nonwip, pct_ooo, pct_unacct
]:
    if dfm is None or dfm.empty or dfnw is None or dfnw.empty:
        return (None, None, None, None, None, None, None, None)
    mc = _metrics_cols(dfm)
    nwc = _nonwip_cols(dfnw)
    if not (mc["date"] and mc["wip_hours"] and nwc["date"]):
        return (None, None, None, None, None, None, None, None)
    m = dfm.copy()
    m[mc["date"]] = _safe_to_datetime(m, mc["date"])
    m = m.dropna(subset=[mc["date"]])
    m["week_start"] = _weekly_start(m[mc["date"]])
    m["completed"] = _to_num(m[mc["wip_hours"]]).fillna(0.0)
    m_wk = m.groupby("week_start", as_index=False).agg(completed=("completed", "sum"))
    n = dfnw.copy()
    n[nwc["date"]] = _safe_to_datetime(n, nwc["date"])
    n = n.dropna(subset=[nwc["date"]])
    n["week_start"] = _weekly_start(n[nwc["date"]])

    if nwc["people_count"] and nwc["people_count"] in n.columns:
        n["people_count"] = _to_num(n[nwc["people_count"]]).fillna(0.0)
    else:
        n["people_count"] = 0.0

    if nwc["total_nonwip"] and nwc["total_nonwip"] in n.columns:
        n["nonwip_total"] = _to_num(n[nwc["total_nonwip"]]).fillna(0.0)
    else:
        n["nonwip_total"] = 0.0

    if nwc["ooohours"] and nwc["ooohours"] in n.columns:
        n["ooo_total"] = _to_num(n[nwc["ooohours"]]).fillna(0.0)
    else:
        n["ooo_total"] = 0.0
    def _col_num(colkey: str) -> pd.Series:
        c = nwc.get(colkey)
        if c and c in n.columns:
            return _to_num(n[c]).fillna(0.0)
        return pd.Series([0.0] * len(n), index=n.index)
    n["wip_workers_count"] = _col_num("wip_workers_count")
    n["wip_workers_ooo"] = _col_num("wip_workers_ooo")
    wip_workers_by_week: Dict[pd.Timestamp, List[str]] = {}
    if nwc.get("wip_workers_json") and nwc["wip_workers_json"] in n.columns:
        for _, r in n.iterrows():
            wk = r["week_start"]
            payload = _loads_json_maybe(r[nwc["wip_workers_json"]])
            if isinstance(payload, list):
                wip_workers_by_week[pd.to_datetime(wk)] = [str(x).strip() for x in payload if str(x).strip()]
    n_wk = n.groupby("week_start", as_index=False).agg(
        people_count=("people_count", "sum"),
        nonwip_total=("nonwip_total", "sum"),
        ooo_total=("ooo_total", "sum"),
        wip_workers_count=("wip_workers_count", "sum"),
        wip_workers_ooo=("wip_workers_ooo", "sum"),
    )
    nonwip_wipworkers_by_week: Dict[pd.Timestamp, float] = {}
    if denom_mode == "HC that worked in WIP" and dfnw_act is not None and not dfnw_act.empty:
        act = dfnw_act.copy()
        adc = _get_date_col(act)
        if adc is not None:
            act[adc] = _safe_to_datetime(act, adc)
            act = act.dropna(subset=[adc])
            act["week_start"] = _weekly_start(act[adc])
            act_cols = {_norm(c): c for c in act.columns}
            by_person_col = act_cols.get("non_wip_by_person") or act_cols.get("non-wip_by_person") or act_cols.get("nonwip_by_person")
            if by_person_col:
                for _, r in act.iterrows():
                    wk = pd.to_datetime(r["week_start"])
                    wips = set(wip_workers_by_week.get(wk, []))
                    if not wips:
                        continue
                    dct = _loads_json_maybe(r[by_person_col])
                    if not isinstance(dct, dict):
                        continue
                    s = 0.0
                    for name, hrs in dct.items():
                        if str(name).strip() in wips:
                            try:
                                s += float(hrs)
                            except Exception:
                                pass
                    nonwip_wipworkers_by_week[wk] = nonwip_wipworkers_by_week.get(wk, 0.0) + s
    wk = m_wk.merge(n_wk, on="week_start", how="inner").sort_values("week_start")
    if wk.empty:
        return (None, None, None, None, None, None, None, None)
    if denom_mode == "Total HC on team":
        wk["denom_hc"] = wk["people_count"]
        wk["ooo_use"] = wk["ooo_total"]
        wk["nonwip_use"] = wk["nonwip_total"]
    else:
        wk["denom_hc"] = wk["wip_workers_count"]
        wk["ooo_use"] = wk["wip_workers_ooo"]
        wk["nonwip_use"] = wk["week_start"].map(lambda x: nonwip_wipworkers_by_week.get(pd.to_datetime(x), float("nan")))
        wk["nonwip_use"] = wk["nonwip_use"].where(wk["nonwip_use"].notna(), wk["nonwip_total"])
    wk.loc[wk["denom_hc"] <= 0, "denom_hc"] = pd.NA
    wk["wip_daily_pp"] = wk["completed"] / (wk["denom_hc"] * float(wd))
    wk["nonwip_daily_pp"] = wk["nonwip_use"] / (wk["denom_hc"] * float(wd))
    wk["capacity_weekly"] = wk["denom_hc"] * float(wd) * 8.0
    wk["unacct_weekly"] = wk["capacity_weekly"] - wk["ooo_use"] - wk["completed"] - wk["nonwip_use"]
    wk["pct_wip"] = 100.0 * (wk["wip_daily_pp"] / 8.0)
    wk["pct_nonwip"] = 100.0 * (wk["nonwip_daily_pp"] / 8.0)
    wk["pct_ooo"] = 100.0 * (wk["ooo_use"] / wk["capacity_weekly"])
    wk["pct_unacct"] = 100.0 - (wk["pct_wip"] + wk["pct_nonwip"] + wk["pct_ooo"])
    avg_wip_daily_pp = float(wk["wip_daily_pp"].dropna().mean()) if wk["wip_daily_pp"].notna().any() else None
    avg_nonwip_daily_pp = float(wk["nonwip_daily_pp"].dropna().mean()) if wk["nonwip_daily_pp"].notna().any() else None
    avg_ooo_weekly = float(wk["ooo_use"].dropna().mean()) if wk["ooo_use"].notna().any() else None
    avg_unacct_weekly = float(wk["unacct_weekly"].dropna().mean()) if wk["unacct_weekly"].notna().any() else None
    pct_wip = float(wk["pct_wip"].dropna().mean()) if wk["pct_wip"].notna().any() else None
    pct_nonwip = float(wk["pct_nonwip"].dropna().mean()) if wk["pct_nonwip"].notna().any() else None
    pct_ooo = float(wk["pct_ooo"].dropna().mean()) if wk["pct_ooo"].notna().any() else None
    pct_unacct = float(wk["pct_unacct"].dropna().mean()) if wk["pct_unacct"].notna().any() else None
    return (
        avg_wip_daily_pp, avg_nonwip_daily_pp, avg_ooo_weekly, avg_unacct_weekly,
        pct_wip, pct_nonwip, pct_ooo, pct_unacct
    )
with tabs[0]:
    dfm_raw = _get_metrics_df()
    dfnw_raw = _get_nonwip_df()
    bounds_df = dfm_raw if (dfm_raw is not None and not dfm_raw.empty) else dfnw_raw
    ov_start, ov_end = section_date_range("Overview date range", bounds_df, key="dr_overview")
    dfm = filter_by_date_range(dfm_raw, ov_start, ov_end) if dfm_raw is not None else None
    dfnw = filter_by_date_range(dfnw_raw, ov_start, ov_end) if dfnw_raw is not None else None
    people_by_week: Dict[pd.Timestamp, float] = {}
    if dfnw is not None and not dfnw.empty:
        people_by_week = _people_lookup_by_week(dfnw)
    wd = _workdays_per_week_assumption()
    st.subheader("Summary")
    denom_mode = st.radio(
        "Per-person denominator for WIP daily hours",
        options=["Total HC on team", "HC that worked in WIP"],
        index=0,
        horizontal=True,
        help="Team-total daily WIP = Completed Hours/5. Per-person daily WIP divides by headcount too.",
    )
    dfnw_act_raw = filter_by_team(data["non_wip_activities"]) if "non_wip_activities" in data else None
    dfnw_act = filter_by_date_range(dfnw_act_raw, ov_start, ov_end) if dfnw_act_raw is not None else None
    (
        avg_daily_wip_per_person,
        avg_daily_nonwip_per_person,
        avg_weekly_ooo_hours,
        avg_weekly_unacct_hours,
        pct_wip,
        pct_nonwip,
        pct_ooo,
        pct_unacct,
    ) = _weekly_rollup_summary(dfm, dfnw, dfnw_act, denom_mode=denom_mode, wd=wd)
    st.markdown("""
    <style>
    div[data-testid="stMetric"]{ text-align: center; }
    label[data-testid="stMetricLabel"]{ display: block; width: 100%; text-align: center; margin: 0; }
    label[data-testid="stMetricLabel"] p{ text-align: center !important; margin: 0 !important; }
    div[data-testid="stMetricValue"]{ text-align: center !important; width: 100%; }
    </style>
    """, unsafe_allow_html=True)
    _, c1, c2, _ = st.columns([1.2, 1.2, 1.2, 1.2])
    c1.metric("Avg Per Person **WIP** Daily Hours", f"{avg_daily_wip_per_person:.2f}" if avg_daily_wip_per_person is not None else "—")
    c2.metric("Avg Per Person **Non-WIP** Daily Hours", f"{avg_daily_nonwip_per_person:.2f}" if avg_daily_nonwip_per_person is not None else "—")
    _, p1, p2, _ = st.columns([1.2, 1.2, 1.2, 1.2])
    p1.metric("**WIP** Ratio", f"{pct_wip:.1f}%" if pct_wip is not None else "—")
    p2.metric("**Non-WIP** Ratio", f"{pct_nonwip:.1f}%" if pct_nonwip is not None else "—")
    st.divider()
    _, _, _, c3, c4,_, _, _ = st.columns([0.5, 1.0, 1.2, 1.2, 1.2, 1.2, 1.0, 0.5])
    c3.metric("Avg **OOO** Weekly Hours", f"{avg_weekly_ooo_hours:.2f}" if avg_weekly_ooo_hours is not None else "—")
    c4.metric("Avg **Unaccounted** Weekly Hours", f"{avg_weekly_unacct_hours:.2f}" if avg_weekly_unacct_hours is not None else "—")
    _, _, _, p3, p4, _, _, _ = st.columns([0.5, 1.0, 1.2, 1.2, 1.2, 1.2, 1.0, 0.5])
    p3.metric("**OOO** % of week", f"{pct_ooo:.1f}%" if pct_ooo is not None else "—")
    p4.metric("**Unaccounted** % remaining", f"{pct_unacct:.1f}%" if pct_unacct is not None else "—")
    st.divider()
    st.subheader("Trend: avg daily WIP hours (week over week)")
    if dfm is not None:
        mc = _metrics_cols(dfm)
        if mc["date"] and mc["wip_hours"]:
            temp = dfm.copy()
            temp[mc["date"]] = _safe_to_datetime(temp, mc["date"])
            temp = temp.dropna(subset=[mc["date"]]).sort_values(mc["date"])
            temp["wip_hours"] = _to_num(temp[mc["wip_hours"]]).fillna(0.0)
            temp["week_start"] = _weekly_start(temp[mc["date"]])
            today = pd.Timestamp.now()
            current_week_start = today.to_period("W-MON").start_time
            temp = temp[temp["week_start"] <= current_week_start]
            wk = (
                temp.groupby("week_start", as_index=False)
                .agg(wip_hours=("wip_hours", "sum"))
                .sort_values("week_start")
            )
            wk["team_total"] = wk["wip_hours"] / float(wd)
            wk["per_person"] = pd.NA
            if denom_mode == "Total HC on team" and people_by_week:
                people_by_week_start = {
                    pd.to_datetime(k).to_period("W-MON").start_time: v
                    for k, v in people_by_week.items()
                }
                wk["people_count"] = wk["week_start"].map(people_by_week_start)
                wk["per_person"] = wk["wip_hours"] / (float(wd) * wk["people_count"])
                wk.loc[wk["people_count"].fillna(0) <= 0, "per_person"] = pd.NA
            elif mc["hc_in_wip"] and mc["hc_in_wip"] in temp.columns:
                temp["hc_in_wip"] = _to_num(temp[mc["hc_in_wip"]]).fillna(0.0)
                wk_hc = (
                    temp.groupby("week_start", as_index=False)
                    .agg(hc_in_wip=("hc_in_wip", "sum"))
                    .sort_values("week_start")
                )
                wk = wk.merge(wk_hc, on="week_start", how="left")
                wk["per_person"] = wk["wip_hours"] / (float(wd) * wk["hc_in_wip"])
                wk.loc[wk["hc_in_wip"].fillna(0) <= 0, "per_person"] = pd.NA
            wk["team_total_up_only"] = wk["team_total"].cummax()
            if wk["per_person"].notna().any():
                wk["per_person_up_only"] = wk["per_person"].cummax()
            view = st.selectbox("Trend view", ["Group total", "Per person"], index=1)
            if view == "Group total":
                st.line_chart(wk.set_index("week_start")["team_total_up_only"])
            else:
                if "per_person_up_only" not in wk.columns or wk["per_person_up_only"].notna().sum() == 0:
                    st.info("Per-person trend not available (no People Count / HC in WIP found for selected range).")
                else:
                    st.line_chart(wk.set_index("week_start")["per_person_up_only"])
        else:
            st.info("Need Week/period_date + Completed Hours to show trend.")
    else:
        st.info("No metrics data loaded for selected teams.")
with tabs[1]:
    if "non_wip" not in data and "non_wip_activities" not in data:
        st.info("No non-WIP CSVs found (expected `non_wip.csv` and/or `non_wip_activities.csv`).")
        st.stop()
    st.markdown("### Non-WIP activities")
    source_raw = None
    if "non_wip" in data:
        cand = filter_by_team(data["non_wip"])
        if not cand.empty:
            source_raw = cand
    if source_raw is None and "non_wip_activities" in data:
        cand = filter_by_team(data["non_wip_activities"])
        if not cand.empty:
            source_raw = cand
    if source_raw is None or source_raw.empty:
        st.info("No Non-WIP activity data available after team filtering.")
        st.stop()
    nw_start, nw_end = section_date_range("Non-WIP date range", source_raw, key="dr_nonwip")
    source_df = filter_by_date_range(source_raw, nw_start, nw_end)
    if source_df.empty:
        st.info("No Non-WIP activity data available in this date range.")
        st.stop()
    dc = _get_date_col(source_df)
    json_col = None
    for c in source_df.columns:
        if _norm(c) in {"non-wip_activities", "non_wip_activities"}:
            json_col = c
            break
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
        if isinstance(payload, list):
            for item in payload:
                if not isinstance(item, dict):
                    continue
                act = item.get("activity") or item.get("Activity") or item.get("type")
                hrs = item.get("hours") or item.get("Hours")
                if act is None or hrs is None:
                    continue
                rows.append(
                    {
                        "week": wk,
                        "activity": str(act).strip(),
                        "hours": float(hrs) if str(hrs) != "" else 0.0,
                    }
                )
    if not rows:
        st.info("No parsable activity rows found in the JSON column.")
        st.stop()
    act_df = pd.DataFrame(rows)
    act_df["week"] = pd.to_datetime(act_df["week"], errors="coerce")
    act_df = act_df.dropna(subset=["week"])
    act_df["week_start"] = _weekly_start(act_df["week"])
    weekly_by_activity = (
        act_df.groupby(["week_start", "activity"], as_index=False)
        .agg(hours=("hours", "sum"))
        .sort_values(["week_start", "hours"], ascending=[True, False])
    )
    avg_weekly = (
        weekly_by_activity.groupby("activity", as_index=False)
        .agg(avg_weekly_hours=("hours", "mean"))
        .sort_values("avg_weekly_hours", ascending=False)
        .head(12)
    )
    st.bar_chart(avg_weekly.set_index("activity")["avg_weekly_hours"])
    st.caption("Top activities by **average weekly hours** (top 12).")
    last_week = weekly_by_activity["week_start"].max()
    last = weekly_by_activity[weekly_by_activity["week_start"] == last_week].copy()
    last = last.sort_values("hours", ascending=False)
    st.write(f"Most recent week starting: **{pd.to_datetime(last_week).date()}**")
    if len(last) > 9:
        top = last.head(8)
        other = pd.DataFrame(
            [
                {
                    "week_start": last_week,
                    "activity": "Other",
                    "hours": float(last["hours"].iloc[8:].sum()),
                }
            ]
        )
        pie_df = pd.concat([top, other], ignore_index=True)
    else:
        pie_df = last
    import matplotlib.pyplot as plt
    fig, ax = plt.subplots()
    ax.pie(
        pie_df["hours"],
        labels=pie_df["activity"],
        autopct="%1.0f%%",
        startangle=90,
    )
    ax.axis("equal")
    st.pyplot(fig)
