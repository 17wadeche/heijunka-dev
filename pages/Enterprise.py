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
    meta: Dict[str, Any] = None  # any extra fields
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
        "non_wip": repo_root / "non_wip.csv",
        "non_wip_activities": repo_root / "non_wip_activities.csv",
        "closures": repo_root / "closures.csv",
        "timelines": repo_root / "timelines.csv",
        "metrics_aggregate_dev": repo_root / "metrics_aggregate_dev.csv",
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
    team_filter = st.multiselect(
        "Teams",
        options=all_team_names,
        default=enabled_team_names,
        help="Select teams to include in the dashboard.",
    )
    show_raw = st.toggle("Show raw tables", value=False)
    st.divider()
    st.caption("Detected data files")
    if data:
        for k in sorted(data.keys()):
            st.write(f"{k} ({len(data[k])} rows)")
    else:
        st.write("No CSVs found at repo root.")
def filter_by_team(df: pd.DataFrame) -> pd.DataFrame:
    if not team_filter:
        return df.iloc[0:0]
    team_cols = [c for c in df.columns if c.strip().lower() in {"team", "team_name", "squad", "org_team"}]
    if not team_cols:
        return df
    col = team_cols[0]
    return df[df[col].astype(str).isin(set(team_filter))]
st.markdown(f"**Org:** {org.org_name} &nbsp;&nbsp;|&nbsp;&nbsp; **Teams in config:** {len(org.teams)}")
if not team_filter:
    st.warning("No teams selected.")
    st.stop()
tabs = st.tabs(["Overview", "Metrics", "Timelines", "Closures", "Non-WIP", "Config"])
with tabs[0]:
    col1, col2, col3, col4 = st.columns(4)
    metrics_rows = len(filter_by_team(data["metrics"])) if "metrics" in data else 0
    timelines_rows = len(filter_by_team(data["timelines"])) if "timelines" in data else 0
    closures_rows = len(filter_by_team(data["closures"])) if "closures" in data else 0
    nonwip_rows = len(filter_by_team(data["non_wip"])) if "non_wip" in data else 0
    col1.metric("Metrics rows", f"{metrics_rows:,}")
    col2.metric("Timelines rows", f"{timelines_rows:,}")
    col3.metric("Closures rows", f"{closures_rows:,}")
    col4.metric("Non-WIP rows", f"{nonwip_rows:,}")
    st.divider()
    if not data:
        st.info(
            "Config loaded successfully, but no CSV data files were found at the repo root. "
            "If your CSVs live elsewhere, update your pipeline or adjust the loader paths in this page."
        )
    else:
        st.subheader("Available datasets")
        for key, df in sorted(data.items()):
            st.write(f"**{key}**")
            st.caption("Columns: " + ", ".join(df.columns.astype(str).tolist()[:40]))
with tabs[1]:
    st.subheader("Metrics")
    if "metrics" not in data and "metrics_aggregate_dev" not in data:
        st.info("No metrics CSV found (expected `metrics.csv` or `metrics_aggregate_dev.csv`).")
    else:
        dfm = data.get("metrics") or data.get("metrics_aggregate_dev")
        dfm = filter_by_team(dfm)
        if dfm.empty:
            st.warning("No rows after team filter.")
        else:
            date_cols = [c for c in dfm.columns if c.strip().lower() in {"date", "day", "as_of", "timestamp"}]
            if date_cols:
                dc = date_cols[0]
                dfm2 = dfm.copy()
                dfm2[dc] = pd.to_datetime(dfm2[dc], errors="coerce")
                dfm2 = dfm2.dropna(subset=[dc]).sort_values(dc)
                numeric_cols = dfm2.select_dtypes(include="number").columns.tolist()
                if numeric_cols:
                    metric_col = st.selectbox("Metric column", numeric_cols, index=0)
                    st.line_chart(dfm2.set_index(dc)[metric_col])
                else:
                    st.info("No numeric columns found to chart in metrics data.")
            else:
                st.info("No date-like column found to chart. Showing table instead.")
            if show_raw:
                st.dataframe(dfm, use_container_width=True)
            st.download_button(
                "Download filtered metrics as CSV",
                data=dfm.to_csv(index=False).encode("utf-8"),
                file_name="metrics_filtered.csv",
                mime="text/csv",
            )
with tabs[2]:
    st.subheader("Timelines")
    if "timelines" not in data:
        st.info("No timelines CSV found (expected `timelines.csv`).")
    else:
        dft = filter_by_team(data["timelines"])
        if dft.empty:
            st.warning("No rows after team filter.")
        else:
            start_cols = [c for c in dft.columns if c.strip().lower() in {"start", "start_date", "begin"}]
            end_cols = [c for c in dft.columns if c.strip().lower() in {"end", "end_date", "finish"}]
            if start_cols and end_cols:
                sc, ec = start_cols[0], end_cols[0]
                temp = dft.copy()
                temp[sc] = pd.to_datetime(temp[sc], errors="coerce")
                temp[ec] = pd.to_datetime(temp[ec], errors="coerce")
                temp["duration_days"] = (temp[ec] - temp[sc]).dt.days
                st.bar_chart(temp.dropna(subset=["duration_days"])["duration_days"])
            else:
                st.caption("No start/end columns detected for a simple duration view.")
            if show_raw:
                st.dataframe(dft, use_container_width=True)
            st.download_button(
                "Download filtered timelines as CSV",
                data=dft.to_csv(index=False).encode("utf-8"),
                file_name="timelines_filtered.csv",
                mime="text/csv",
            )
with tabs[3]:
    st.subheader("Closures")
    if "closures" not in data:
        st.info("No closures CSV found (expected `closures.csv`).")
    else:
        dfc = filter_by_team(data["closures"])
        if dfc.empty:
            st.warning("No rows after team filter.")
        else:
            status_cols = [c for c in dfc.columns if c.strip().lower() in {"status", "state", "outcome"}]
            if status_cols:
                sc = status_cols[0]
                st.bar_chart(dfc[sc].astype(str).value_counts())
            else:
                st.caption("No status/state/outcome column found for a breakdown.")
            if show_raw:
                st.dataframe(dfc, use_container_width=True)
            st.download_button(
                "Download filtered closures as CSV",
                data=dfc.to_csv(index=False).encode("utf-8"),
                file_name="closures_filtered.csv",
                mime="text/csv",
            )
with tabs[4]:
    st.subheader("Non-WIP")
    if "non_wip" not in data and "non_wip_activities" not in data:
        st.info("No non-WIP CSVs found (expected `non_wip.csv` and/or `non_wip_activities.csv`).")
    else:
        if "non_wip" in data:
            st.markdown("**non_wip.csv**")
            dfn = filter_by_team(data["non_wip"])
            if show_raw:
                st.dataframe(dfn, use_container_width=True)
            st.download_button(
                "Download filtered non_wip as CSV",
                data=dfn.to_csv(index=False).encode("utf-8"),
                file_name="non_wip_filtered.csv",
                mime="text/csv",
            )
        if "non_wip_activities" in data:
            st.markdown("**non_wip_activities.csv**")
            dfa = filter_by_team(data["non_wip_activities"])
            if show_raw:
                st.dataframe(dfa, use_container_width=True)
            st.download_button(
                "Download filtered non_wip_activities as CSV",
                data=dfa.to_csv(index=False).encode("utf-8"),
                file_name="non_wip_activities_filtered.csv",
                mime="text/csv",
            )
with tabs[5]:
    st.subheader("Org config")
    st.caption("This is the parsed view (supports teams as strings or objects).")
    cfg_table = pd.DataFrame(
        [
            {
                "team": t.name,
                "enabled": t.enabled,
                **(t.meta or {}),
            }
            for t in org.teams
        ]
    )
    st.dataframe(cfg_table, use_container_width=True)
    with st.expander("Raw JSON", expanded=False):
        st.json(org.raw)