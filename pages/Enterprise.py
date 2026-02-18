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


# ----------------------------
# Sidebar filters
# ----------------------------
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
        help="Filter the org by portfolio.",
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
        help="Filter the Teams list by OU (within selected portfolios).",
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
        help="Select teams to include in the dashboard.",
    )

    # --- NEW: Date filter ---
    st.divider()
    st.caption("Date filter (applies across all tabs)")

    # Build a global min/max from all loaded datasets that have a date column
    date_candidates: List[pd.Timestamp] = []
    for k, df in data.items():
        dc = _get_date_col(df)
        if not dc:
            continue
        ser = pd.to_datetime(df[dc], errors="coerce").dropna()
        if not ser.empty:
            date_candidates.append(ser.min())
            date_candidates.append(ser.max())

    if date_candidates:
        global_min = min(date_candidates).date()
        global_max = max(date_candidates).date()
        date_range = st.date_input(
            "Date range",
            value=(global_min, global_max),
            min_value=global_min,
            max_value=global_max,
            help="Filter rows to this date range (inclusive).",
        )
        if isinstance(date_range, tuple) and len(date_range) == 2:
            date_start, date_end = date_range
        else:
            date_start, date_end = global_min, global_max
    else:
        st.info("No date columns detected in CSVs.")
        date_start, date_end = None, None

    st.divider()
    show_raw = st.toggle("Show raw tables", value=False)


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


def filter_by_date(df: pd.DataFrame) -> pd.DataFrame:
    """Apply the global sidebar date filter if we can find a date column."""
    if date_start is None or date_end is None:
        return df
    dc = _get_date_col(df)
    if not dc:
        return df
    tmp = df.copy()
    tmp[dc] = pd.to_datetime(tmp[dc], errors="coerce")
    tmp = tmp.dropna(subset=[dc])
    start_ts = pd.to_datetime(date_start)
    end_ts = pd.to_datetime(date_end) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    return tmp[(tmp[dc] >= start_ts) & (tmp[dc] <= end_ts)]


def filter_df(df: pd.DataFrame) -> pd.DataFrame:
    return filter_by_date(filter_by_team(df))


st.markdown(f"**Selected teams:** {len(team_filter)}")
if not team_filter:
    st.warning("No teams selected.")
    st.stop()

tabs = st.tabs(["Overview", "Non-WIP"])

def _get_metrics_df() -> Optional[pd.DataFrame]:
    if "metrics" in data:
        d = filter_df(data["metrics"])
        if not d.empty:
            return d
    if "metrics_aggregate_dev" in data:
        d = filter_df(data["metrics_aggregate_dev"])
        if not d.empty:
            return d
    return None


def _get_nonwip_df() -> Optional[pd.DataFrame]:
    if "non_wip" in data:
        d = filter_df(data["non_wip"])
        if not d.empty:
            return d
    if "non_wip_activities" in data:
        d = filter_df(data["non_wip_activities"])
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
        "activities_json": _first_col(df, ["non-wip_activities", "non_wip_activities"]),
    }


def _people_lookup_by_week(dfnw: pd.DataFrame) -> Dict[pd.Timestamp, float]:
    """Map week(date) -> people_count for denominator."""
    out: Dict[pd.Timestamp, float] = {}
    nwc = _nonwip_cols(dfnw)
    if not (nwc["date"] and nwc["people_count"]):
        return out
    tmp = dfnw.copy()
    tmp[nwc["date"]] = pd.to_datetime(tmp[nwc["date"]], errors="coerce")
    tmp = tmp.dropna(subset=[nwc["date"]])
    tmp["people_count"] = _to_num(tmp[nwc["people_count"]]).fillna(0.0)
    # If multiple teams -> multiple rows per week; sum gives "total people" across selected teams
    grp = tmp.groupby(nwc["date"], as_index=False)["people_count"].sum()
    for _, r in grp.iterrows():
        out[pd.to_datetime(r[nwc["date"]])] = float(r["people_count"])
    return out


# ----------------------------
# Overview
# ----------------------------
with tabs[0]:
    dfm = _get_metrics_df()
    dfnw = _get_nonwip_df()
    wd = _workdays_per_week_assumption()

    st.subheader("Summary")

    # Denominator choice for per-person calcs
    denom_mode = st.radio(
        "Per-person denominator for WIP daily hours",
        options=["Total HC on team", "HC that worked in WIP"],
        index=0,
        horizontal=True,
        help="Team-total daily WIP = Completed Hours/5. Per-person daily WIP divides by headcount too.",
    )
    avg_daily_wip_team = None
    avg_daily_wip_per_person = None
    avg_daily_nonwip_team = None
    avg_daily_nonwip_per_person = None
    pct_wip = None
    people_by_week: Dict[pd.Timestamp, float] = {}
    if dfnw is not None and not dfnw.empty:
        people_by_week = _people_lookup_by_week(dfnw)
    hc_used_value: Optional[float] = None
    if dfm is not None:
        mc0 = _metrics_cols(dfm)
        if mc0["date"] and mc0["hc_used"] and mc0["hc_used"] in dfm.columns:
            tmp_hc = dfm.copy()
            tmp_hc[mc0["date"]] = _safe_to_datetime(tmp_hc, mc0["date"])
            tmp_hc = tmp_hc.dropna(subset=[mc0["date"]]).sort_values(mc0["date"])
            tmp_hc["hc_used"] = _to_num(tmp_hc[mc0["hc_used"]])
            if tmp_hc["hc_used"].notna().any():
                hc_used_value = float(tmp_hc["hc_used"].dropna().mean())
    if dfm is not None:
        mc = _metrics_cols(dfm)
        if mc["date"] and mc["wip_hours"]:
            temp = dfm.copy()
            temp[mc["date"]] = _safe_to_datetime(temp, mc["date"])
            temp = temp.dropna(subset=[mc["date"]]).sort_values(mc["date"])
            temp["wip_hours"] = _to_num(temp[mc["wip_hours"]]).fillna(0.0)
            temp["daily_wip_team"] = temp["wip_hours"] / float(wd)
            avg_daily_wip_team = float(temp["daily_wip_team"].mean())
            if denom_mode == "Total HC on team" and people_by_week:
                temp["people_count"] = temp[mc["date"]].map(people_by_week)
                temp["daily_wip_per_person"] = temp["wip_hours"] / (float(wd) * temp["people_count"])
                temp.loc[temp["people_count"].fillna(0) <= 0, "daily_wip_per_person"] = pd.NA
                if temp["daily_wip_per_person"].notna().any():
                    avg_daily_wip_per_person = float(temp["daily_wip_per_person"].dropna().mean())
            else:
                if mc["hc_in_wip"] and mc["hc_in_wip"] in temp.columns:
                    temp["hc_in_wip"] = _to_num(temp[mc["hc_in_wip"]]).fillna(0.0)
                    temp["daily_wip_per_person"] = temp["wip_hours"] / (float(wd) * temp["hc_in_wip"])
                    temp.loc[temp["hc_in_wip"] <= 0, "daily_wip_per_person"] = pd.NA
                    if temp["daily_wip_per_person"].notna().any():
                        avg_daily_wip_per_person = float(temp["daily_wip_per_person"].dropna().mean())
        else:
            st.info("Metrics data is missing a date column (Week/period_date) and/or Completed Hours.")
    if dfnw is not None:
        nwc = _nonwip_cols(dfnw)
        if nwc["date"] and nwc["total_nonwip"]:
            tempn = dfnw.copy()
            tempn[nwc["date"]] = _safe_to_datetime(tempn, nwc["date"])
            tempn = tempn.dropna(subset=[nwc["date"]]).sort_values(nwc["date"])
            tempn["nonwip_hours"] = _to_num(tempn[nwc["total_nonwip"]]).fillna(0.0)
            tempn["daily_nonwip_team"] = tempn["nonwip_hours"] / float(wd)
            avg_daily_nonwip_team = float(tempn["daily_nonwip_team"].mean())
            if nwc["people_count"]:
                tempn["people_count"] = _to_num(tempn[nwc["people_count"]]).fillna(0.0)
                tempn["daily_nonwip_per_person"] = tempn["nonwip_hours"] / (float(wd) * tempn["people_count"])
                tempn.loc[tempn["people_count"] <= 0, "daily_nonwip_per_person"] = pd.NA
                if tempn["daily_nonwip_per_person"].notna().any():
                    avg_daily_nonwip_per_person = float(tempn["daily_nonwip_per_person"].dropna().mean())
    if avg_daily_wip_team is not None and avg_daily_nonwip_team is not None and (avg_daily_wip_team + avg_daily_nonwip_team) > 0:
        pct_wip = 100.0 * avg_daily_wip_team / (avg_daily_wip_team + avg_daily_nonwip_team)
    elif (
        avg_daily_wip_per_person is not None
        and avg_daily_nonwip_per_person is not None
        and (avg_daily_wip_per_person + avg_daily_nonwip_per_person) > 0
    ):
        pct_wip = 100.0 * avg_daily_wip_per_person / (avg_daily_wip_per_person + avg_daily_nonwip_per_person)
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("Avg daily WIP (team total)", f"{avg_daily_wip_team:.2f}" if avg_daily_wip_team is not None else "—")
    k2.metric("Avg daily WIP (per person)", f"{avg_daily_wip_per_person:.2f}" if avg_daily_wip_per_person is not None else "—")
    k3.metric("Avg daily Non-WIP (team total)", f"{avg_daily_nonwip_team:.2f}" if avg_daily_nonwip_team is not None else "—")
    k4.metric("Avg daily Non-WIP (per person)", f"{avg_daily_nonwip_per_person:.2f}" if avg_daily_nonwip_per_person is not None else "—")
    k5.metric("% WIP (WIP / (WIP+Non-WIP))", f"{pct_wip:.1f}%" if pct_wip is not None else "—")
    k6.metric("Actual HC Used (6 hrs/day target)", f"{hc_used_value:.2f}" if hc_used_value is not None else "—")
    st.caption("Daily averages assume **5 workdays/week**. Per-person uses headcount where available.")
    st.divider()
    st.subheader("Trend: avg daily WIP hours (week over week)")
    if dfm is not None:
        mc = _metrics_cols(dfm)
        if mc["date"] and mc["wip_hours"]:
            temp = dfm.copy()
            temp[mc["date"]] = _safe_to_datetime(temp, mc["date"])
            temp = temp.dropna(subset=[mc["date"]]).sort_values(mc["date"])
            temp["wip_hours"] = _to_num(temp[mc["wip_hours"]]).fillna(0.0)

            # --- NEW: true WoW buckets + exclude future weeks ---
            temp["week_start"] = _weekly_start(temp[mc["date"]])

            # "Current week" boundary (Mon-start week)
            today = pd.Timestamp.now()
            current_week_start = today.to_period("W-MON").start_time
            temp = temp[temp["week_start"] <= current_week_start]

            # Weekly aggregation across selected teams
            wk = (
                temp.groupby("week_start", as_index=False)
                .agg(wip_hours=("wip_hours", "sum"))
                .sort_values("week_start")
            )
            wk["team_total"] = wk["wip_hours"] / float(wd)

            # build per-person weekly series (same logic as above, but weekly)
            wk["per_person"] = pd.NA
            if denom_mode == "Total HC on team" and people_by_week:
                # people_by_week keys may be exact dates; align by week_start
                people_by_week_start = {
                    pd.to_datetime(k).to_period("W-MON").start_time: v
                    for k, v in people_by_week.items()
                }
                wk["people_count"] = wk["week_start"].map(people_by_week_start)
                wk["per_person"] = wk["wip_hours"] / (float(wd) * wk["people_count"])
                wk.loc[wk["people_count"].fillna(0) <= 0, "per_person"] = pd.NA
            elif mc["hc_in_wip"] and mc["hc_in_wip"] in temp.columns:
                # Need weekly hc_in_wip sum too
                temp["hc_in_wip"] = _to_num(temp[mc["hc_in_wip"]]).fillna(0.0)
                wk_hc = (
                    temp.groupby("week_start", as_index=False)
                    .agg(hc_in_wip=("hc_in_wip", "sum"))
                    .sort_values("week_start")
                )
                wk = wk.merge(wk_hc, on="week_start", how="left")
                wk["per_person"] = wk["wip_hours"] / (float(wd) * wk["hc_in_wip"])
                wk.loc[wk["hc_in_wip"].fillna(0) <= 0, "per_person"] = pd.NA

            # --- NEW: "only go up" enforcement ---
            wk["team_total_up_only"] = wk["team_total"].cummax()
            if wk["per_person"].notna().any():
                wk["per_person_up_only"] = wk["per_person"].cummax()

            view = st.selectbox("Trend view", ["Team total", "Per person"], index=0)
            if view == "Team total":
                st.line_chart(wk.set_index("week_start")["team_total_up_only"])
                st.caption("WoW (weekly) and forced non-decreasing: **cummax(Completed Hours / 5)**. Future weeks excluded.")
            else:
                if "per_person_up_only" not in wk.columns or wk["per_person_up_only"].notna().sum() == 0:
                    st.info("Per-person trend not available (no People Count / HC in WIP found for selected range).")
                else:
                    st.line_chart(wk.set_index("week_start")["per_person_up_only"])
                    st.caption("WoW (weekly) and forced non-decreasing: **cummax(Completed Hours / (5 * headcount))**. Future weeks excluded.")
        else:
            st.info("Need Week/period_date + Completed Hours to show trend.")
    else:
        st.info("No metrics data loaded for selected teams.")

with tabs[1]:
    st.subheader("Non-WIP")
    if "non_wip" not in data and "non_wip_activities" not in data:
        st.info("No non-WIP CSVs found (expected `non_wip.csv` and/or `non_wip_activities.csv`).")
        st.stop()
    if "non_wip" in data:
        st.markdown("### Weekly Non-WIP summary (`non_wip.csv`)")
        dfn = filter_df(data["non_wip"])
        if dfn.empty:
            st.warning("No rows after filters.")
        else:
            dc = _get_date_col(dfn)
            total_col = _first_col(dfn, ["total_non-wip_hours", "total_non_wip_hours"])
            pct_col = _first_col(dfn, ["%_in_wip"])
            if dc:
                tmp = dfn.copy()
                tmp[dc] = _safe_to_datetime(tmp, dc)
                tmp = tmp.dropna(subset=[dc]).sort_values(dc)
                if total_col:
                    tmp["total_nonwip_hours"] = _to_num(tmp[total_col]).fillna(0.0)
                    tmp["avg_daily_nonwip"] = tmp["total_nonwip_hours"] / float(_workdays_per_week_assumption())
                    st.line_chart(tmp.set_index(dc)["avg_daily_nonwip"])
                if pct_col:
                    tmp["pct_in_wip"] = _to_num(tmp[pct_col])
                    if tmp["pct_in_wip"].notna().any():
                        st.line_chart(tmp.set_index(dc)["pct_in_wip"])
            else:
                st.caption("No Week/period_date column found for charts.")
            if show_raw:
                st.dataframe(dfn, use_container_width=True)
            st.download_button(
                "Download filtered non_wip as CSV",
                data=dfn.to_csv(index=False).encode("utf-8"),
                file_name="non_wip_filtered.csv",
                mime="text/csv",
            )
    st.divider()
    st.markdown("### Non-WIP activities")
    source_df = None
    if "non_wip" in data:
        cand = filter_df(data["non_wip"])
        if not cand.empty:
            source_df = cand
    if source_df is None and "non_wip_activities" in data:
        cand = filter_df(data["non_wip_activities"])
        if not cand.empty:
            source_df = cand
    if source_df is None:
        st.info("No Non-WIP activity data available after filtering.")
    else:
        dc = _get_date_col(source_df)
        json_col = None
        for c in source_df.columns:
            if _norm(c) in {"non-wip_activities", "non_wip_activities"}:
                json_col = c
                break
        if not (dc and json_col):
            st.info("Need `Week/period_date` and `Non-WIP Activities` (JSON list) to roll up activities.")
        else:
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
            else:
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
                if show_raw:
                    st.dataframe(weekly_by_activity, use_container_width=True)