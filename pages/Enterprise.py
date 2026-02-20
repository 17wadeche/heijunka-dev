# pages/Enterprise.py
from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st


# -----------------------------
# Repo / config helpers
# -----------------------------
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
    meta: Dict[str, Any] = None


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
                meta = {k: v for k, v in t.items() if k not in {"name", "team", "Team", "enabled"}}
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


# -----------------------------
# Column helpers
# -----------------------------
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


def _week_key(s: pd.Series) -> pd.Series:
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
            # sometimes double quotes are doubled in CSVs; try a mild normalize
            try:
                return json.loads(t.replace('""', '"'))
            except Exception:
                return None
    return None


def _workdays_per_week_assumption() -> int:
    return 5


# -----------------------------
# Streamlit setup
# -----------------------------
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
    portfolio_filter = st.multiselect("Portfolio", options=all_portfolios, default=all_portfolios)

    teams_after_portfolio = (
        [t for t in org.teams if str((t.meta or {}).get("portfolio")).strip() in set(portfolio_filter)]
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
    ou_filter = st.multiselect("OU", options=all_ous, default=all_ous)

    teams_after_ou = (
        [t for t in teams_after_portfolio if str((t.meta or {}).get("ou")).strip() in set(ou_filter)]
        if ou_filter
        else []
    )

    team_options = [t.name for t in teams_after_ou]
    default_teams = [t for t in enabled_team_names if t in team_options]
    if not default_teams and team_options:
        default_teams = team_options

    team_filter = st.multiselect("Teams", options=team_options, default=default_teams)


def filter_by_team(df: pd.DataFrame) -> pd.DataFrame:
    if not team_filter:
        return df.iloc[0:0]
    team_cols = [c for c in df.columns if c.strip().lower() in {"team", "team_name", "squad", "org_team"}]
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


def section_date_range(label: str, df: Optional[pd.DataFrame], key: str) -> tuple[Optional[pd.Timestamp], Optional[pd.Timestamp]]:
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

    presets = ["Custom", "Past week", "Past month", "Past 3 months", "Past 6 months", "Past year", "Past 2 years"]
    days_map = {"Past week": 7, "Past month": 30, "Past 3 months": 90, "Past 6 months": 180, "Past year": 365, "Past 2 years": 730}

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

    dr = st.date_input(label, min_value=min_d, max_value=max_d, key=dates_key, help="Filters only this section.")
    if isinstance(dr, tuple) and len(dr) == 2:
        start_d, end_d = dr
    else:
        start_d, end_d = start_default, end_default

    start_ts = pd.to_datetime(start_d)
    end_ts = pd.to_datetime(end_d) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    return start_ts, end_ts


def filter_by_date_range(df: Optional[pd.DataFrame], start_ts: Optional[pd.Timestamp], end_ts: Optional[pd.Timestamp]) -> Optional[pd.DataFrame]:
    if df is None:
        return None
    if start_ts is None or end_ts is None:
        return df
    dc = _get_date_col(df)
    if not dc:
        return df
    tmp = df.copy()
    tmp[dc] = pd.to_datetime(tmp[dc], errors="coerce")
    tmp = tmp.dropna(subset=[dc])
    return tmp[(tmp[dc] >= start_ts) & (tmp[dc] <= end_ts)]


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


def _get_nonwip_totals_df() -> Optional[pd.DataFrame]:
    if "non_wip" in data:
        d = filter_by_team(data["non_wip"])
        return d if not d.empty else None
    return None


def _get_nonwip_activities_df() -> Optional[pd.DataFrame]:
    if "non_wip_activities" in data:
        d = filter_by_team(data["non_wip_activities"])
        return d if not d.empty else None
    return None


def _metrics_cols(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    return {
        "date": _get_date_col(df),
        "completed_hours": _first_col(df, ["completed_hours", "wip_hours", "completedhours", "completed_hours."]),
        "hc_in_wip": _first_col(df, ["hc_in_wip"]),
    }


def _nonwip_cols(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    return {
        "date": _get_date_col(df),
        "total_nonwip": _first_col(df, ["total_non-wip_hours", "total_non_wip_hours", "total_non_wip_hours."]),
        "people_count": _first_col(df, ["people_count", "people_count."]),
        "ooohours": _first_col(df, ["ooo_hours", "ooo_hours."]),
        "wip_workers": _first_col(df, ["wip_workers", "wip workers"]),
        "wip_workers_count": _first_col(df, ["wip_workers_count", "wip workers count"]),
        "wip_workers_ooo": _first_col(df, ["wip_workers_ooo_hours", "wip workers ooo hours"]),
    }


def _fmt(v: Optional[float], nd: int = 2) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "—"
    return f"{float(v):.{nd}f}"


def _fmt_pct(v: Optional[float]) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "—"
    return f"{float(v):.1f}%"


# -----------------------------
# Overview
# -----------------------------
with tabs[0]:
    dfm_raw = _get_metrics_df()
    dfnw_totals_raw = _get_nonwip_totals_df()
    dfnw_act_raw = _get_nonwip_activities_df()

    bounds_df = dfm_raw if (dfm_raw is not None and not dfm_raw.empty) else dfnw_totals_raw
    ov_start, ov_end = section_date_range("Overview date range", bounds_df, key="dr_overview")

    dfm = filter_by_date_range(dfm_raw, ov_start, ov_end)
    dfnw_totals = filter_by_date_range(dfnw_totals_raw, ov_start, ov_end)
    dfnw_act = filter_by_date_range(dfnw_act_raw, ov_start, ov_end)

    wd = _workdays_per_week_assumption()
    HOURS_PER_DAY = 8.0

    st.subheader("Summary")
    denom_mode = st.radio(
        "Per-person denominator for WIP daily hours",
        options=["Total HC on team", "HC that worked in WIP"],
        index=0,
        horizontal=True,
        help="Total HC uses People Count. HC that worked in WIP uses WIP Workers Count + worker-filtered non-wip-by-person.",
    )

    # ---- Build weekly summary table: wk_summary ----
    wk_m = None
    if dfm is not None and not dfm.empty:
        mc = _metrics_cols(dfm)
        if mc["date"] and mc["completed_hours"]:
            m = dfm.copy()
            m[mc["date"]] = _safe_to_datetime(m, mc["date"])
            m = m.dropna(subset=[mc["date"]])
            m["week_start"] = _week_key(m[mc["date"]])
            m["completed_hours"] = _to_num(m[mc["completed_hours"]]).fillna(0.0)
            wk_m = m.groupby("week_start", as_index=False).agg(completed_hours=("completed_hours", "sum"))

    wk_nw = None
    if dfnw_totals is not None and not dfnw_totals.empty:
        nwc = _nonwip_cols(dfnw_totals)
        if nwc["date"]:
            n = dfnw_totals.copy()
            n[nwc["date"]] = _safe_to_datetime(n, nwc["date"])
            n = n.dropna(subset=[nwc["date"]])
            n["week_start"] = _week_key(n[nwc["date"]])

            n["people_count"] = _to_num(n[nwc["people_count"]]).fillna(0.0) if nwc["people_count"] else 0.0
            n["nonwip_hours"] = _to_num(n[nwc["total_nonwip"]]).fillna(0.0) if nwc["total_nonwip"] else 0.0
            n["ooo_hours"] = _to_num(n[nwc["ooohours"]]).fillna(0.0) if nwc["ooohours"] else 0.0

            n["wip_workers_count"] = _to_num(n[nwc["wip_workers_count"]]).fillna(0.0) if nwc["wip_workers_count"] else 0.0
            n["wip_workers_ooo"] = _to_num(n[nwc["wip_workers_ooo"]]).fillna(0.0) if nwc["wip_workers_ooo"] else 0.0

            def _parse_workers(v: Any) -> List[str]:
                x = _loads_json_maybe(v)
                if isinstance(x, list):
                    return [str(z).strip() for z in x if str(z).strip()]
                return []

            n["wip_workers_list"] = (
                n[nwc["wip_workers"]].apply(_parse_workers) if nwc["wip_workers"] else [[]] * len(n)
            )

            def _union_lists(series: pd.Series) -> List[str]:
                s = set()
                for lst in series:
                    for name in (lst or []):
                        s.add(str(name).strip())
                return sorted([x for x in s if x])

            wk_nw = (
                n.groupby("week_start", as_index=False)
                .agg(
                    people_count=("people_count", "sum"),
                    nonwip_hours=("nonwip_hours", "sum"),
                    ooo_hours=("ooo_hours", "sum"),
                    wip_workers_count=("wip_workers_count", "sum"),
                    wip_workers_ooo=("wip_workers_ooo", "sum"),
                    wip_workers_list=("wip_workers_list", _union_lists),
                )
            )

    if wk_m is None and wk_nw is None:
        st.info("No weekly data available in the selected date range.")
        st.stop()

    if wk_m is None:
        wk_summary = wk_nw.copy()
    elif wk_nw is None:
        wk_summary = wk_m.copy()
    else:
        wk_summary = wk_nw.merge(wk_m, on="week_start", how="outer")

    # Ensure numeric columns exist
    for col in ["people_count", "nonwip_hours", "ooo_hours", "wip_workers_count", "wip_workers_ooo", "completed_hours"]:
        if col not in wk_summary.columns:
            wk_summary[col] = 0.0
    wk_summary = wk_summary.fillna({c: 0.0 for c in ["people_count", "nonwip_hours", "ooo_hours", "wip_workers_count", "wip_workers_ooo", "completed_hours"]})
    wk_summary = wk_summary.sort_values("week_start")

    # ---- Worker-filtered Non-WIP numerator (from non_wip_activities.csv) ----
    # For "HC that worked in WIP": sum non_wip_by_person for people in WIP Workers list for that week.
    workers_nonwip_by_week: Dict[pd.Timestamp, float] = {}
    if dfnw_act is not None and not dfnw_act.empty:
        dc = _get_date_col(dfnw_act)
        json_col = _first_col(dfnw_act, ["non_wip_by_person", "non-wip_by_person"])
        if dc and json_col:
            a = dfnw_act.copy()
            a[dc] = _safe_to_datetime(a, dc)
            a = a.dropna(subset=[dc])
            a["week_start"] = _week_key(a[dc])

            # Build quick lookup from week_start -> workers list
            wk_workers_lookup: Dict[pd.Timestamp, set] = {}
            if "wip_workers_list" in wk_summary.columns:
                for _, rr in wk_summary.iterrows():
                    wk_workers_lookup[pd.to_datetime(rr["week_start"])] = set(rr.get("wip_workers_list", []) or [])

            def _sum_workers_nonwip(row: pd.Series) -> float:
                wk0 = pd.to_datetime(row["week_start"])
                workers = wk_workers_lookup.get(wk0, set())
                if not workers:
                    return 0.0
                payload = _loads_json_maybe(row[json_col])
                if not isinstance(payload, dict):
                    return 0.0
                total = 0.0
                for name, hrs in payload.items():
                    nm = str(name).strip()
                    if nm in workers:
                        try:
                            total += float(hrs)
                        except Exception:
                            pass
                return total
            a["workers_nonwip_hours"] = a.apply(_sum_workers_nonwip, axis=1)
            tmp = a.groupby("week_start", as_index=False).agg(workers_nonwip_hours=("workers_nonwip_hours", "sum"))
            workers_nonwip_by_week = {pd.to_datetime(k): float(v) for k, v in zip(tmp["week_start"], tmp["workers_nonwip_hours"])}
    wk_summary["workers_nonwip_hours"] = wk_summary["week_start"].apply(lambda x: float(workers_nonwip_by_week.get(pd.to_datetime(x), 0.0)))
    def clamp_pct(x: Optional[float]) -> Optional[float]:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return None
        return max(0.0, min(100.0, float(x)))
    valid_totalHC = wk_summary["people_count"] > 0
    valid_workers = wk_summary["wip_workers_count"] > 0
    people_days_total = float((wk_summary.loc[valid_totalHC, "people_count"] * wd).sum())
    worker_days_total = float((wk_summary.loc[valid_workers, "wip_workers_count"] * wd).sum())
    completed_total_totalHC = float(wk_summary.loc[valid_totalHC, "completed_hours"].sum())
    nonwip_total_totalHC = float(wk_summary.loc[valid_totalHC, "nonwip_hours"].sum())
    ooo_total_totalHC = float(wk_summary.loc[valid_totalHC, "ooo_hours"].sum())
    capacity_total_totalHC = float((wk_summary.loc[valid_totalHC, "people_count"] * wd * HOURS_PER_DAY).sum())
    completed_total_workers = float(wk_summary.loc[valid_workers, "completed_hours"].sum())
    nonwip_total_workers = float(wk_summary.loc[valid_workers, "workers_nonwip_hours"].sum())
    ooo_total_workers = float(wk_summary.loc[valid_workers, "wip_workers_ooo"].sum())
    capacity_total_workers = float((wk_summary.loc[valid_workers, "wip_workers_count"] * wd * HOURS_PER_DAY).sum())
    if denom_mode == "Total HC on team":
        avg_wip_pp_daily = (completed_total_totalHC / people_days_total) if people_days_total > 0 else None
        avg_nonwip_pp_daily = (nonwip_total_totalHC / people_days_total) if people_days_total > 0 else None
        wip_pct = clamp_pct((avg_wip_pp_daily / HOURS_PER_DAY) * 100.0 if avg_wip_pp_daily is not None else None)
        nonwip_pct = clamp_pct((avg_nonwip_pp_daily / HOURS_PER_DAY) * 100.0 if avg_nonwip_pp_daily is not None else None)
        ooo_hours_display = float(wk_summary.loc[valid_totalHC, "ooo_hours"].mean()) if valid_totalHC.any() else None
        ooo_pct = clamp_pct((ooo_total_totalHC / capacity_total_totalHC) * 100.0 if capacity_total_totalHC > 0 else None)
        wk_summary["unacct_weekly_totalHC"] = (
            (wk_summary["people_count"] * wd * HOURS_PER_DAY)
            - wk_summary["ooo_hours"]
            - wk_summary["completed_hours"]
            - wk_summary["nonwip_hours"]
        )
        unacct_weekly_display = float(wk_summary.loc[valid_totalHC, "unacct_weekly_totalHC"].mean()) if valid_totalHC.any() else None
        unacct_pct = clamp_pct(100.0 - ((wip_pct or 0.0) + (nonwip_pct or 0.0) + (ooo_pct or 0.0))) if (wip_pct is not None and nonwip_pct is not None and ooo_pct is not None) else None
        disp_wip, disp_nonwip, disp_ooo, disp_unacct = avg_wip_pp_daily, avg_nonwip_pp_daily, ooo_hours_display, unacct_weekly_display
        disp_wip_pct, disp_nonwip_pct, disp_ooo_pct, disp_unacct_pct = wip_pct, nonwip_pct, ooo_pct, unacct_pct
    else:
        avg_wip_pp_daily = (completed_total_workers / worker_days_total) if worker_days_total > 0 else None
        avg_nonwip_pp_daily = (nonwip_total_workers / worker_days_total) if worker_days_total > 0 else None
        wip_pct = clamp_pct((avg_wip_pp_daily / HOURS_PER_DAY) * 100.0 if avg_wip_pp_daily is not None else None)
        nonwip_pct = clamp_pct((avg_nonwip_pp_daily / HOURS_PER_DAY) * 100.0 if avg_nonwip_pp_daily is not None else None)
        ooo_hours_display = float(wk_summary.loc[valid_workers, "wip_workers_ooo"].mean()) if valid_workers.any() else None
        ooo_pct = clamp_pct((ooo_total_workers / capacity_total_workers) * 100.0 if capacity_total_workers > 0 else None)
        wk_summary["unacct_weekly_workers"] = (
            (wk_summary["wip_workers_count"] * wd * HOURS_PER_DAY)
            - wk_summary["wip_workers_ooo"]
            - wk_summary["completed_hours"]
            - wk_summary["workers_nonwip_hours"]
        )
        unacct_weekly_display = float(wk_summary.loc[valid_workers, "unacct_weekly_workers"].mean()) if valid_workers.any() else None
        unacct_pct = clamp_pct(100.0 - ((wip_pct or 0.0) + (nonwip_pct or 0.0) + (ooo_pct or 0.0))) if (wip_pct is not None and nonwip_pct is not None and ooo_pct is not None) else None
        disp_wip, disp_nonwip, disp_ooo, disp_unacct = avg_wip_pp_daily, avg_nonwip_pp_daily, ooo_hours_display, unacct_weekly_display
        disp_wip_pct, disp_nonwip_pct, disp_ooo_pct, disp_unacct_pct = wip_pct, nonwip_pct, ooo_pct, unacct_pct
    st.markdown(
        """
        <style>
        div[data-testid="stMetric"]{ text-align:center; }
        label[data-testid="stMetricLabel"]{ display:block; width:100%; text-align:center; margin:0; }
        label[data-testid="stMetricLabel"] p{ text-align:center !important; margin:0 !important; }
        div[data-testid="stMetricValue"]{ text-align:center !important; width:100%; }
        </style>
        """,
        unsafe_allow_html=True,
    )
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Avg/Person **WIP** Daily Hours", _fmt(disp_wip, 2), _fmt_pct(disp_wip_pct))
    c2.metric("Avg/Person **Non-WIP** Daily Hours", _fmt(disp_nonwip, 2), _fmt_pct(disp_nonwip_pct))
    if denom_mode == "Total HC on team":
        c3.metric("**OOO** Hours (avg weekly)", _fmt(disp_ooo, 2), _fmt_pct(disp_ooo_pct))
    else:
        c3.metric("**OOO** Hours (WIP workers, avg weekly)", _fmt(disp_ooo, 2), _fmt_pct(disp_ooo_pct))
    c4.metric("Avg **Unaccounted** Weekly Hours", _fmt(disp_unacct, 2), _fmt_pct(disp_unacct_pct))
    st.caption("Assumes **5 workdays/week** and **8 hours/day**. Percents are based on an **8-hour day** (WIP/Non-WIP) and capacity share (OOO). Unaccounted % is remainder after WIP+Non-WIP+OOO.")
    st.divider()
    st.subheader("Trend: avg daily WIP hours (week over week)")
    if dfm is not None and not dfm.empty:
        mc = _metrics_cols(dfm)
        if mc["date"] and mc["completed_hours"]:
            temp = dfm.copy()
            temp[mc["date"]] = _safe_to_datetime(temp, mc["date"])
            temp = temp.dropna(subset=[mc["date"]]).sort_values(mc["date"])
            temp["wip_hours"] = _to_num(temp[mc["completed_hours"]]).fillna(0.0)
            temp["week_start"] = _weekly_start(temp[mc["date"]])
            today = pd.Timestamp.now()
            current_week_start = today.to_period("W-MON").start_time
            temp = temp[temp["week_start"] <= current_week_start]
            wk_trend = (
                temp.groupby("week_start", as_index=False)
                .agg(wip_hours=("wip_hours", "sum"))
                .sort_values("week_start")
            )
            wk_trend["team_total_daily"] = wk_trend["wip_hours"] / float(wd)
            wk_trend["per_person_daily"] = pd.NA
            if denom_mode == "Total HC on team":
                people_lookup = {pd.to_datetime(r["week_start"]): float(r["people_count"]) for _, r in wk_summary.iterrows()}
                wk_trend["people_count"] = wk_trend["week_start"].map(people_lookup).fillna(0.0)
                wk_trend["per_person_daily"] = wk_trend["wip_hours"] / (float(wd) * wk_trend["people_count"])
                wk_trend.loc[wk_trend["people_count"] <= 0, "per_person_daily"] = pd.NA
            else:
                workers_lookup = {pd.to_datetime(r["week_start"]): float(r["wip_workers_count"]) for _, r in wk_summary.iterrows()}
                wk_trend["wip_workers_count"] = wk_trend["week_start"].map(workers_lookup).fillna(0.0)
                wk_trend["per_person_daily"] = wk_trend["wip_hours"] / (float(wd) * wk_trend["wip_workers_count"])
                wk_trend.loc[wk_trend["wip_workers_count"] <= 0, "per_person_daily"] = pd.NA
            view = st.selectbox("Trend view", ["Group total (daily)", "Per person (daily)"], index=1)
            if view == "Group total (daily)":
                st.line_chart(wk_trend.set_index("week_start")["team_total_daily"])
            else:
                if wk_trend["per_person_daily"].notna().sum() == 0:
                    st.info("Per-person trend not available (no denominator values found for selected range).")
                else:
                    st.line_chart(wk_trend.set_index("week_start")["per_person_daily"])
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
    if source_df is None or source_df.empty:
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
                rows.append({"week": wk, "activity": str(act).strip(), "hours": float(hrs) if str(hrs) != "" else 0.0})
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
        other = pd.DataFrame([{"week_start": last_week, "activity": "Other", "hours": float(last["hours"].iloc[8:].sum())}])
        pie_df = pd.concat([top, other], ignore_index=True)
    else:
        pie_df = last
    import matplotlib.pyplot as plt
    fig, ax = plt.subplots()
    ax.pie(pie_df["hours"], labels=pie_df["activity"], autopct="%1.0f%%", startangle=90)
    ax.axis("equal")
    st.pyplot(fig)