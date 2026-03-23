# pages/Enterprise.py
from __future__ import annotations
import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import pandas as pd
import streamlit as st
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
def _coalesce_matching_cols(df: pd.DataFrame, candidates: List[str]) -> pd.Series:
    matches = []
    wanted = set(candidates)
    for c in df.columns:
        if _norm(c) in wanted:
            matches.append(c)
    if not matches:
        return pd.Series([pd.NA] * len(df), index=df.index)
    out = df[matches[0]]
    for c in matches[1:]:
        out = out.where(out.notna() & (out.astype(str).str.strip() != ""), df[c])
    return out
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
def _format_export_display(df: pd.DataFrame) -> pd.io.formats.style.Styler:
    rename_map = {
        "team": "Team",
        "week_start": "Week Start",
        "completed_hours": "Completed Hours",
        "people_count": "People Count",
        "non_wip_hours": "Non-WIP Hours",
        "ooo_hours": "OOO Hours",
        "ou": "OU",
        "portfolio": "Portfolio",
        "capacity_hours": "Capacity Hours",
        "unaccounted_hours": "Unaccounted Hours",
        "wip_pct": "WIP %",
        "non_wip_pct": "Non-WIP %",
        "ooo_pct": "OOO %",
        "unaccounted_pct": "Unaccounted %",
    }
    out = df.copy().rename(columns=rename_map)
    if "Week Start" in out.columns:
        out["Week Start"] = pd.to_datetime(out["Week Start"], errors="coerce").dt.date
    fmt = {}
    for c in [
        "Completed Hours",
        "People Count",
        "Non-WIP Hours",
        "OOO Hours",
        "Capacity Hours",
        "Unaccounted Hours",
    ]:
        if c in out.columns:
            fmt[c] = "{:,.2f}"
    for c in ["WIP %", "Non-WIP %", "OOO %", "Unaccounted %"]:
        if c in out.columns:
            fmt[c] = "{:.1%}"
    return out.style.format(fmt)
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
def _weekly_team_export_df(
    dfm: Optional[pd.DataFrame],
    dfnw: Optional[pd.DataFrame],
    org: OrgConfig,
) -> pd.DataFrame:
    metrics_team = pd.DataFrame(columns=["team", "week_start", "completed_hours"])
    nonwip_team = pd.DataFrame(columns=["team", "week_start", "people_count", "non_wip_hours", "ooo_hours"])
    if dfm is not None and not dfm.empty:
        m = dfm.copy()
        team_col_m = _get_team_col(m)
        date_ser = _coalesce_matching_cols(m, ["week", "period_date", "date", "day", "as_of", "timestamp"])
        wip_ser = _coalesce_matching_cols(m, ["completed_hours", "wip_hours", "completedhours"])
        if team_col_m is not None:
            m["__date__"] = pd.to_datetime(date_ser, errors="coerce")
            m["__completed_hours__"] = pd.to_numeric(wip_ser, errors="coerce")
            m = m.dropna(subset=["__date__"])
            m["week_start"] = _weekly_start(m["__date__"])
            m["__completed_hours__"] = m["__completed_hours__"].fillna(0.0)
            metrics_team = (
                m.groupby([team_col_m, "week_start"], as_index=False)
                .agg(completed_hours=("__completed_hours__", "sum"))
                .rename(columns={team_col_m: "team"})
            )
    if dfnw is not None and not dfnw.empty:
        nw = dfnw.copy()
        team_col_nw = _get_team_col(nw)
        date_ser = _coalesce_matching_cols(nw, ["week", "period_date", "date", "day", "as_of", "timestamp"])
        people_ser = _coalesce_matching_cols(nw, ["people_count", "people_count."])
        nonwip_ser = _coalesce_matching_cols(nw, ["total_non-wip_hours", "total_non_wip_hours", "total_non_wip_hours."])
        ooo_ser = _coalesce_matching_cols(nw, ["ooo_hours", "ooo_hours."])
        if team_col_nw is not None:
            nw["__date__"] = pd.to_datetime(date_ser, errors="coerce")
            nw = nw.dropna(subset=["__date__"])
            nw["week_start"] = _weekly_start(nw["__date__"])
            nw["people_count"] = pd.to_numeric(people_ser, errors="coerce").fillna(0.0)
            nw["non_wip_hours"] = pd.to_numeric(nonwip_ser, errors="coerce").fillna(0.0)
            nw["ooo_hours"] = pd.to_numeric(ooo_ser, errors="coerce").fillna(0.0)
            nonwip_team = (
                nw.groupby([team_col_nw, "week_start"], as_index=False)
                .agg(
                    people_count=("people_count", "sum"),
                    non_wip_hours=("non_wip_hours", "sum"),
                    ooo_hours=("ooo_hours", "sum"),
                )
                .rename(columns={team_col_nw: "team"})
            )
    if metrics_team.empty and nonwip_team.empty:
        return pd.DataFrame()
    base = metrics_team.merge(nonwip_team, on=["team", "week_start"], how="outer")
    for col in ["completed_hours", "people_count", "non_wip_hours", "ooo_hours"]:
        if col not in base.columns:
            base[col] = 0.0
        base[col] = pd.to_numeric(base[col], errors="coerce").fillna(0.0)
    meta = _team_meta_lookup(org)
    base = base.merge(meta, on="team", how="left")
    base["capacity_hours"] = base["people_count"] * 40.0
    base["unaccounted_hours"] = (
        base["capacity_hours"]
        - base["completed_hours"]
        - base["non_wip_hours"]
        - base["ooo_hours"]
    ).clip(lower=0.0)
    for src, pct_col in [
        ("completed_hours", "wip_pct"),
        ("non_wip_hours", "non_wip_pct"),
        ("ooo_hours", "ooo_pct"),
        ("unaccounted_hours", "unaccounted_pct"),
    ]:
        base[pct_col] = (base[src] / base["capacity_hours"]).where(base["capacity_hours"] > 0)
    base = _add_avg_hours_day_columns(base)
    return base.sort_values(["week_start", "portfolio", "ou", "team"]).reset_index(drop=True)
def _rollup_export_level(df: pd.DataFrame, level: str) -> pd.DataFrame:
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
        )
    )
    out["capacity_hours"] = out["people_count"] * 40.0
    out["unaccounted_hours"] = (
        out["capacity_hours"]
        - out["completed_hours"]
        - out["non_wip_hours"]
        - out["ooo_hours"]
    ).clip(lower=0.0)
    for src, pct_col in [
        ("completed_hours", "wip_pct"),
        ("non_wip_hours", "non_wip_pct"),
        ("ooo_hours", "ooo_pct"),
        ("unaccounted_hours", "unaccounted_pct"),
    ]:
        out[pct_col] = (out[src] / out["capacity_hours"]).where(out["capacity_hours"] > 0)
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
tabs = st.tabs(["Overview", "Non-WIP", "Export"])
def _get_metrics_df() -> Optional[pd.DataFrame]:
    frames = []
    for key in ["metrics", "metrics_aggregate_dev", "NS_WIP", "CRM_WIP", "MS_WIP"]:
        if key in data:
            d = filter_by_team(data[key])
            if not d.empty:
                frames.append(d.copy())
    if not frames:
        return None
    return pd.concat(frames, ignore_index=True, sort=False).drop_duplicates()
def _get_nonwip_df() -> Optional[pd.DataFrame]:
    frames = []
    for key in ["ns_non_wip_activities", "crm_non_wip_activities", "ms_non_wip_activities", "non_wip", "non_wip_activities"]:
        if key in data:
            d = filter_by_team(data[key])
            if not d.empty:
                frames.append(_normalize_df_columns(d.copy()))
    if not frames:
        return None
    return pd.concat(frames, ignore_index=True, sort=False).drop_duplicates()
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
    Optional[float], Optional[float], Optional[float], Optional[float],
    Optional[float], Optional[float], Optional[float], Optional[float],
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
    def _norm_person_name(v: Any) -> str:
        return str(v).strip().lower()
    wip_workers_by_week: Dict[pd.Timestamp, set[str]] = {}
    if nwc.get("wip_workers_json") and nwc["wip_workers_json"] in n.columns:
        for _, r in n.iterrows():
            wk = pd.to_datetime(r["week_start"])
            payload = _loads_json_maybe(r[nwc["wip_workers_json"]])
            if isinstance(payload, list):
                vals = {_norm_person_name(x) for x in payload if str(x).strip()}
                wip_workers_by_week[wk] = wip_workers_by_week.get(wk, set()) | vals
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
                    wips = wip_workers_by_week.get(wk, set())
                    if not wips:
                        continue
                    dct = _loads_json_maybe(r[by_person_col])
                    if not isinstance(dct, dict):
                        continue
                    s = 0.0
                    for name, hrs in dct.items():
                        if _norm_person_name(name) in wips:
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
    ov_start, ov_end = section_date_range(
        "Overview date range",
        bounds_df,
        key="dr_overview",
        min_floor_ts=selected_nonwip_floor,
    )
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
    def _get_nonwip_activity_detail_df() -> Optional[pd.DataFrame]:
        for key in ["ns_non_wip_activities", "crm_non_wip_activities","ms_non_wip_activities", "non_wip_activities", "non_wip"]:
            if key in data:
                d = filter_by_team(data[key])
                if not d.empty:
                    return d
        return None
    dfnw_act_raw = _get_nonwip_activity_detail_df()
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
    _, _, c3, c4,_, _, _ = st.columns([1.35, 1.2, 1.2, 1.2, 1.2, 1.0, 0.5])
    c3.metric("Avg **OOO** Weekly Hours", f"{avg_weekly_ooo_hours:.2f}" if avg_weekly_ooo_hours is not None else "—")
    c4.metric("Avg **Unaccounted** Weekly Hours", f"{avg_weekly_unacct_hours:.2f}" if avg_weekly_unacct_hours is not None else "—")
    _, _, p3, p4, _, _, _ = st.columns([1.35, 1.2, 1.2, 1.2, 1.2, 1.0, 0.5])
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
EXCLUDED_NON_WIP = {"ooo", "non-wip", "non_wip", "other", "other team wip"}
def _norm_activity_name(val: Any) -> str:
    return str(val).strip().lower().replace("_", "-")
with tabs[1]:
    if (
        "non_wip" not in data
        and "non_wip_activities" not in data
        and "ns_non_wip_activities" not in data
        and "crm_non_wip_activities" not in data
        and "ms_non_wip_activities" not in data
    ):
        st.info("No non-WIP CSVs found.")
    else:
        st.markdown("### Non-WIP activities")
        source_raw = None
        for key in ["ns_non_wip_activities","ms_non_wip_activities", "crm_non_wip_activities", "non_wip", "non_wip_activities"]:
            if key in data:
                cand = filter_by_team(data[key])
                if not cand.empty:
                    source_raw = cand
                    break
        if source_raw is None or source_raw.empty:
            st.info("No Non-WIP activity data available after team filtering.")
            st.stop()
        nw_start, nw_end = section_date_range(
            "Non-WIP date range",
            source_raw,
            key="dr_nonwip",
            min_floor_ts=selected_nonwip_floor,
        )
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
        weekly_raw = (
            act_df.groupby(["week_start", "activity"], as_index=False)
            .agg(hours=("hours", "sum"))
        )
        normalised_chunks: List[pd.DataFrame] = []
        for wk_val, grp in weekly_raw.groupby("week_start"):
            cat = grp[["activity", "hours"]].rename(
                columns={"activity": "Activity", "hours": "Hours"}
            )
            cat_norm = split_nonwip_activity_minutes(cat)
            cat_norm["week_start"] = wk_val
            normalised_chunks.append(cat_norm)
        if normalised_chunks:
            weekly_by_activity = (
                pd.concat(normalised_chunks, ignore_index=True)
                .rename(columns={"Activity": "activity", "Hours": "hours"})
                .sort_values(["week_start", "hours"], ascending=[True, False])
            )
        else:
            weekly_by_activity = weekly_raw.copy()
        weekly_by_activity = weekly_by_activity[
            ~weekly_by_activity["activity"].map(_norm_activity_name).isin(EXCLUDED_NON_WIP)
        ]
        if weekly_by_activity.empty:
            st.info("No Non-WIP activity data available after exclusions.")
            st.stop()
        total_hours = (
            weekly_by_activity.groupby("activity", as_index=False)
            .agg(total_hours=("hours", "sum"))
            .sort_values("total_hours", ascending=False)
            .head(15)
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
        ax.set_title("Top 15 Non-WIP Activities by Total Hours")
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
        st.caption("Top 15 activities by total hours for the selected period, sorted highest to lowest from left to right.")
        st.divider()
        st.markdown("#### Activity breakdown — pie chart")
        pie_start, pie_end = section_date_range(
            "Pie chart date range",
            source_raw,
            key="dr_nonwip_pie",
            min_floor_ts=selected_nonwip_floor,
        )
        pie_source_df = filter_by_date_range(source_raw, pie_start, pie_end)
        if pie_source_df.empty:
            st.info("No Non-WIP activity data in the selected pie chart date range.")
            st.stop()
        pie_dc = _get_date_col(pie_source_df)
        pie_json_col = None
        for c in pie_source_df.columns:
            if _norm(c) in {"non-wip_activities", "non_wip_activities"}:
                pie_json_col = c
                break
        pie_rows: List[Dict[str, Any]] = []
        if pie_dc and pie_json_col:
            pie_tmp = pie_source_df.copy()
            pie_tmp[pie_dc] = _safe_to_datetime(pie_tmp, pie_dc)
            pie_tmp = pie_tmp.dropna(subset=[pie_dc]).sort_values(pie_dc)
            for _, r in pie_tmp.iterrows():
                payload = _loads_json_maybe(r[pie_json_col])
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
                        pie_rows.append(
                            {
                                "activity": str(act).strip(),
                                "hours": float(hrs) if str(hrs) != "" else 0.0,
                            }
                        )
        if pie_rows:
            pie_act_df = pd.DataFrame(pie_rows)
            pie_cat = pie_act_df.rename(columns={"activity": "Activity", "hours": "Hours"})
            pie_cat_norm = split_nonwip_activity_minutes(pie_cat)
            pie_rolled = pie_cat_norm.rename(columns={"Activity": "activity", "Hours": "hours"})
            pie_rolled = pie_rolled.groupby("activity", as_index=False).agg(hours=("hours", "sum"))
            pie_rolled = pie_rolled[
                ~pie_rolled["activity"].map(_norm_activity_name).isin(EXCLUDED_NON_WIP)
            ]
            pie_rolled = pie_rolled.sort_values("hours", ascending=False)
            if pie_rolled.empty:
                st.info('No pie chart data available after excluding "OOO" and "Non-WIP".')
                st.stop()
            if len(pie_rolled) > 14:
                top_pie = pie_rolled.head(14)
                other_pie = pd.DataFrame(
                    [{
                        "activity": "Other",
                        "hours": float(pie_rolled["hours"].iloc[14:].sum()),
                    }]
                )
                pie_df = pd.concat([top_pie, other_pie], ignore_index=True)
            else:
                pie_df = pie_rolled
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
        else:
            st.info("No parsable activity rows found for the selected pie chart date range.")
with tabs[2]:
    st.subheader("Export")
    export_team_filter = st.multiselect(
        "Export — Teams",
        options=all_team_names,
        default=all_team_names,
        key="export_team_filter",
    )

    def filter_by_export_team(df: pd.DataFrame) -> pd.DataFrame:
        if not export_team_filter:
            return df.iloc[0:0]
        tc = _get_team_col(df)
        if not tc:
            return df
        return df[df[tc].astype(str).isin(set(export_team_filter))]

    def filter_by_export_date(df: pd.DataFrame, start_ts, end_ts) -> pd.DataFrame:
        if start_ts is None or end_ts is None:
            return df
        dc = _get_date_col(df)
        if not dc:
            return df
        tmp = df.copy()
        tmp[dc] = pd.to_datetime(tmp[dc], errors="coerce")
        tmp = tmp.dropna(subset=[dc])
        return tmp[(tmp[dc] >= start_ts) & (tmp[dc] <= end_ts)]
    export_metrics_frames = []
    for key in ["metrics", "metrics_aggregate_dev", "NS_WIP", "CRM_WIP", "MS_WIP"]:
        if key in data:
            d = filter_by_export_team(data[key].copy())
            if not d.empty:
                export_metrics_frames.append(d)
    export_nonwip_frames = []
    for key in ["ns_non_wip_activities", "ms_non_wip_activities","crm_non_wip_activities", "non_wip_activities", "non_wip"]:
        if key in data:
            d = filter_by_export_team(data[key].copy())
            if not d.empty:
                export_nonwip_frames.append(_normalize_df_columns(d))
    export_bounds_df = export_metrics_frames[0] if export_metrics_frames else (
        export_nonwip_frames[0] if export_nonwip_frames else None
    )
    ex_start, ex_end = section_date_range(
        "Export date range",
        export_bounds_df,
        key="dr_export",
        min_floor_ts=None,
        allow_future_dates=True,
    )
    def _concat_frames(frames):
        if not frames:
            return None
        if len(frames) == 1:
            return frames[0]
        base_cols = set(frames[0].columns)
        compatible = [f for f in frames if set(f.columns) == base_cols]
        other = [f for f in frames if set(f.columns) != base_cols]
        result_frames = []
        if compatible:
            result_frames.append(pd.concat(compatible, ignore_index=True))
        result_frames.extend(other)
        if len(result_frames) == 1:
            return result_frames[0]
        return pd.concat(result_frames, ignore_index=True, sort=False)
    export_metrics_filtered = _concat_frames(
        [filter_by_export_date(f, ex_start, ex_end) for f in export_metrics_frames]
    ) if export_metrics_frames else None
    export_nonwip_filtered = _concat_frames(
        [filter_by_export_date(f, ex_start, ex_end) for f in export_nonwip_frames]
    ) if export_nonwip_frames else None
    team_export = _weekly_team_export_df(export_metrics_filtered, export_nonwip_filtered, org)
    if not team_export.empty:
        team_export = team_export[
            (team_export["completed_hours"] > 0) & (team_export["people_count"] > 0)
        ].reset_index(drop=True)
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
            styler = styler.applymap(lambda v: _threshold_cell_style(v, 0.80, good_if_gte=True), subset=["WIP %"])
        if "Non-WIP %" in out.columns:
            styler = styler.applymap(lambda v: _threshold_cell_style(v, 0.20), subset=["Non-WIP %"])
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
            styler = styler.applymap(lambda v: _threshold_cell_style(v, 0.80, good_if_gte=True), subset=["WIP %"])
        if "Non-WIP %" in out.columns:
            styler = styler.applymap(lambda v: _threshold_cell_style(v, 0.20), subset=["Non-WIP %"])
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
            styler = styler.applymap(lambda v: _threshold_cell_style(v, 0.80, good_if_gte=True), subset=["WIP %"])
        if "Non-WIP %" in out.columns:
            styler = styler.applymap(lambda v: _threshold_cell_style(v, 0.20), subset=["Non-WIP %"])
        return styler
    if team_export.empty:
        st.info("No exportable team/week data found.")
    else:
        ou_export = _rollup_export_level(team_export, "ou")
        portfolio_export = _rollup_export_level(team_export, "portfolio")
        st.markdown("#### Team weekly")
        st.dataframe(_format_export_display_team(team_export), use_container_width=True, hide_index=True)
        st.markdown("#### OU weekly")
        st.dataframe(_format_export_display_ou(ou_export), use_container_width=True, hide_index=True)
        st.markdown("#### Portfolio weekly")
        st.dataframe(_format_export_display_portfolio(portfolio_export), use_container_width=True, hide_index=True)
        try:
            team_export_display = _display_export_team_df(team_export)
            ou_export_display = _display_export_ou_df(ou_export)
            portfolio_export_display = _display_export_portfolio_df(portfolio_export)
            xlsx_bytes = _excel_bytes_from_export_dfs(
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