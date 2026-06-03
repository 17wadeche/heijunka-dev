# pages/Neuroscience.py
import hmac
import os, sys
from pathlib import Path
import pandas as pd
import numpy as np
import streamlit as st
from utils.nonwip_kpi_lookup import enterprise_nonwip_kpi_lookup
import altair as alt
import json
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from utils.activity_map import ACTIVITY_MAP
from utils.styles import apply_global_styles
apply_global_styles()
NON_WIP_DEFAULT_PATH = Path(r"C:\heijunka-dev\MS_DATA\ms_non_wip_activities.csv")
WIP_GROUPS_DEFAULT_PATH = Path(r"C:\heijunka-dev\MS_DATA\MS_WIP_metrics.csv")
NONWIP_GROUPS_DEFAULT_PATH = Path(r"C:\heijunka-dev\MS_DATA\MS_NONWIP_METRICS.csv")
def _safe_secret(name: str, default=None):
    import os
    try:
        return st.secrets.get(name, os.environ.get(name, default))
    except Exception:
        return os.environ.get(name, default)
MS_WIP_GROUP_METRICS_URL = _safe_secret("MS_WIP_GROUP_METRICS_URL")
MS_NONWIP_GROUP_METRICS_URL = _safe_secret("MS_NONWIP_GROUP_METRICS_URL")
NON_WIP_DATA_URL = _safe_secret("MS_NON_WIP_DATA_URL")
DATA_URL = _safe_secret("MS_HEIJUNKA_DATA_URL")
def _fmt_hours_minutes(x) -> str:
    try:
        total_mins = int(round(float(x) * 60))
    except Exception:
        return "0m"
    h, m = divmod(total_mins, 60)
    if h and m:
        return f"{h}h {m:02d}m"
    if h and not m:
        return f"{h}h"
    return f"{m}m"
from pathlib import Path
TEAMS_CONFIG_PATH = Path(__file__).resolve().parents[1] / "teams.json"
@st.cache_data(show_spinner=False, ttl=15 * 60)
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
@st.cache_data(show_spinner=False, ttl=15 * 60)
def load_non_wip(
    nw_path: str | None = None,
    nw_url: str | None = None,
    cache_tag: str = "MS", 
) -> pd.DataFrame:
    if nw_url is None:
        nw_url = NON_WIP_DATA_URL
    if nw_url:
        try:
            df = pd.read_csv(nw_url, dtype=str, keep_default_na=False, encoding="utf-8-sig")
        except Exception:
            import io, requests
            r = requests.get(nw_url, timeout=20)
            r.raise_for_status()
            df = pd.read_csv(
                io.StringIO(r.content.decode("utf-8-sig", errors="replace")),
                dtype=str, keep_default_na=False
            )
    else:
        p = Path(nw_path or NON_WIP_DEFAULT_PATH)
        if not p.exists():
            return pd.DataFrame(columns=[
                "team","period_date","source_file","people_count",
                "total_non_wip_hours","% in WIP","non_wip_by_person"
            ])
        df = pd.read_csv(p, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    if "period_date" in df.columns:
        df["period_date"] = pd.to_datetime(df["period_date"], errors="coerce").dt.normalize()
    for c in ["people_count", "total_non_wip_hours", "% in WIP", "OOO Hours"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    if "% in WIP" in df.columns and "% Non-WIP" not in df.columns:
        s = pd.to_numeric(df["% in WIP"], errors="coerce")
        if pd.notna(s.max()):
            if float(s.max()) <= 1.5:
                pct_wip_0_100 = s * 100.0
            else:
                pct_wip_0_100 = s
            df["% Non-WIP"] = 100.0 - pct_wip_0_100
    return df
@st.cache_data(show_spinner=False, ttl=15 * 60)
def load_wip_group_metrics(path: str | None = None, url: str | None = None) -> pd.DataFrame:
    if url is None:
        url = MS_WIP_GROUP_METRICS_URL
    if url:
        df = pd.read_csv(
            url,
            engine="python",
            sep=None,
            encoding="utf-8-sig",
            on_bad_lines="skip",
            dtype=str,
        )
        return _postprocess(df)
    p = Path(path or WIP_GROUPS_DEFAULT_PATH)
    if not p.exists():
        return pd.DataFrame()
    df = pd.read_csv(
        p,
        engine="python",
        sep=None,
        encoding="utf-8-sig",
        on_bad_lines="skip",
        dtype=str,
    )
    return _postprocess(df)
@st.cache_data(show_spinner=False, ttl=15 * 60)
def load_nonwip_group_metrics(path: str | None = None, url: str | None = None) -> pd.DataFrame:
    if url is None:
        url = MS_NONWIP_GROUP_METRICS_URL
    if url:
        df = pd.read_csv(url, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    else:
        p = Path(path or NONWIP_GROUPS_DEFAULT_PATH)
        if not p.exists():
            return pd.DataFrame()
        df = pd.read_csv(p, dtype=str, keep_default_na=False, encoding="utf-8-sig")

    if "period_date" in df.columns:
        df["period_date"] = pd.to_datetime(df["period_date"], errors="coerce").dt.normalize()
    for c in ["people_count", "total_non_wip_hours", "% in WIP", "OOO Hours"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    if "% in WIP" in df.columns and "% Non-WIP" not in df.columns:
        s = pd.to_numeric(df["% in WIP"], errors="coerce")
        if pd.notna(s.max()):
            df["% Non-WIP"] = 100.0 - (s * 100.0 if float(s.max()) <= 1.5 else s)
    return df
TEAM_BREAKDOWN_RULES = {
    "ACM": ["All", "US", "MEIC", "CTS"],
    "Endoscopy": ["All", "US", "MEIC", "CTS"],
    "VSS": ["All", "US", "MEIC", "CTS"],
    "Surgical AST-GST": ["All", "US", "MEIC", "CTS"],
    "Surgical Robotics": ["All", "US", "MEIC"],
}
def _norm_team_text(x: str) -> str:
    return " ".join(str(x or "").strip().split())
def split_team_group(team_name: str) -> tuple[str, str]:
    raw = _norm_team_text(team_name)
    raw_lower = raw.lower()
    if not raw:
        return "", "All"
    explicit_map = {
        "endoscopy": ("Endoscopy", "US"),
        "endo us": ("Endoscopy", "US"),
        "endo meic": ("Endoscopy", "MEIC"),
        "cts-gis": ("Endoscopy", "CTS"),
        "cts-sibo": ("Surgical AST-GST", "CTS"),
        "cts-sinh": ("Surgical AST-GST", "CTS"),
        "cts-vents": ("VSS", "CTS"),
        "acm": ("ACM", "US"),
        "vss": ("VSS", "US"),
        "surgical ast-gst": ("Surgical AST-GST", "US"),
        "surgical robotics": ("Surgical Robotics", "US"),
    }
    if raw_lower in explicit_map:
        return explicit_map[raw_lower]
    for base, allowed in TEAM_BREAKDOWN_RULES.items():
        if raw == base:
            return base, "All"
        if raw in {f"CTS-{base}-RI", f"CTS-{base}-PM"}:
            return base, "CTS"
        for subgroup in allowed:
            if subgroup == "All":
                continue
            candidates = {
                f"{base} - {subgroup}",
                f"{base}-{subgroup}",
                f"{base}_{subgroup}",
                f"{base} {subgroup}",
                f"{base} ({subgroup})",
            }
            if raw in candidates:
                return base, subgroup
    return raw, "All"
def add_team_group_columns(frame: pd.DataFrame) -> pd.DataFrame:
    if frame is None or frame.empty or "team" not in frame.columns:
        return frame.copy()
    out = frame.copy()
    parsed = out["team"].map(split_team_group)
    out["team_group"] = parsed.map(lambda x: x[0])
    out["team_subgroup"] = parsed.map(lambda x: x[1])
    return out
def grouped_team_options(frame: pd.DataFrame) -> list[str]:
    if frame is None or frame.empty:
        return []
    if "team_group" not in frame.columns:
        frame = add_team_group_columns(frame)
    groups = sorted([t for t in frame["team_group"].dropna().unique() if str(t).strip()])
    return groups
def subgroup_options_for_team(team_group: str) -> list[str]:
    return TEAM_BREAKDOWN_RULES.get(team_group, ["All"])
def filter_team_view(
    frame: pd.DataFrame,
    team_group: str,
    subgroup: str = "All",
    fallback_to_all: bool = True,
) -> pd.DataFrame:
    if frame is None or frame.empty:
        return frame.copy()
    if "team_group" not in frame.columns:
        frame = add_team_group_columns(frame)
    sub = frame[frame["team_group"] == team_group].copy()
    if subgroup != "All":
        exact = sub[sub["team_subgroup"] == subgroup]
        if exact.empty and fallback_to_all:
            sub = sub[sub["team_subgroup"] == "All"].copy()
        else:
            sub = exact
    return sub
@st.cache_data(show_spinner=False, ttl=15 * 60)
def explode_non_wip_by_person(nw: pd.DataFrame) -> pd.DataFrame:
    cols = ["team","period_date","person","Non-WIP Hours"]
    if nw.empty or "non_wip_by_person" not in nw.columns:
        return pd.DataFrame(columns=cols)
    rows = []
    sub = nw[["team","period_date","non_wip_by_person"]].dropna(subset=["non_wip_by_person"])
    for _, r in sub.iterrows():
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
            rows.append({
                "team": r["team"],
                "period_date": pd.to_datetime(r["period_date"], errors="coerce").normalize(),
                "person": normalize_person_name(str(person).strip()),
                "Non-WIP Hours": v
            })
    out = pd.DataFrame(rows, columns=cols)
    if not out.empty:
        out["period_date"] = pd.to_datetime(out["period_date"], errors="coerce").dt.normalize()
    return out
def merged_people_count_for_week(
    team: str,
    week,
    nw_frame: pd.DataFrame,
    person_hours: pd.DataFrame,
    people_in_wip: pd.DataFrame,
) -> int:
    wk = pd.to_datetime(week, errors="coerce").normalize()
    if nw_frame is not None and not nw_frame.empty:
        raw_nw = nw_frame.copy()
        raw_nw["period_date"] = pd.to_datetime(raw_nw["period_date"], errors="coerce").dt.normalize()
        if "people_count" in raw_nw.columns:
            team_match = raw_nw.loc[
                (raw_nw["team"] == team) & (raw_nw["period_date"] == wk),
                "people_count",
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
            [person_col],
        ].copy()
        if not sub.empty:
            vals = sub[person_col].astype(str).str.strip()
            names.update(x for x in vals if x)
    return len(names)
DEFAULT_DATA_PATH = Path(r"C:\heijunka-dev\MS_DATA\MS_WIP.csv")
if hasattr(st, "autorefresh"):
    st.autorefresh(interval=60 * 60 * 1000, key="auto-refresh")
@st.cache_data(show_spinner=False, ttl=15 * 60)
def load_data(data_path: str | None, data_url: str | None):
    if data_url:
        try:
            lower = data_url.lower()
            if lower.endswith((".xlsx", ".xlsm", ".xls")):
                df = pd.read_excel(data_url, sheet_name="All Metrics")
            elif lower.endswith(".json"):
                df = pd.read_json(data_url)
            else:
                df = pd.read_csv(
                    data_url,
                    engine="python",      # enables sep=None sniffing
                    sep=None,             # auto-detect delimiter (comma, tab, semicolon…)
                    encoding="utf-8-sig", # handles BOM
                    on_bad_lines="skip",  # don't die on ragged rows
                    dtype=str,            # keep raw text; you coerce later in _postprocess
                )
        except pd.errors.ParserError:
            try:
                df = pd.read_csv(
                    data_url,
                    engine="python",
                    sep=";",
                    encoding="utf-8-sig",
                    on_bad_lines="skip",
                    dtype=str,
                )
            except Exception as e:
                st.error(f"Couldn't parse HEIJUNKA_DATA_URL as CSV: {e}")
                return pd.DataFrame()
        except Exception:
            import io, requests
            try:
                r = requests.get(data_url, timeout=20)
                r.raise_for_status()
                b = r.content
                head = b[:2048].lstrip()
                if head.startswith((b"{", b"[")):
                    df = pd.read_json(io.BytesIO(b))
                elif b[:2] == b"PK":
                    df = pd.read_excel(io.BytesIO(b), sheet_name="All Metrics")
                else:
                    df = pd.read_csv(
                        io.StringIO(b.decode("utf-8-sig", errors="replace")),
                        engine="python",
                        sep=None,
                        on_bad_lines="skip",
                        dtype=str,
                    )
            except Exception as e:
                st.error(f"Failed to fetch/parse HEIJUNKA_DATA_URL: {e}")
                return pd.DataFrame()
        return _postprocess(df)
    if not data_path:
        return pd.DataFrame()
    p = Path(data_path)
    if not p.exists():
        return pd.DataFrame()
    if p.suffix.lower() in (".xlsx", ".xlsm"):
        df = pd.read_excel(p, sheet_name="All Metrics")
    elif p.suffix.lower() == ".csv":
        df = pd.read_csv(p, engine="python", sep=None, encoding="utf-8-sig", on_bad_lines="skip", dtype=str)
    elif p.suffix.lower() == ".json":
        df = pd.read_json(p)
    else:
        return pd.DataFrame()
    return _postprocess(df)
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
def _postprocess(df: pd.DataFrame) -> pd.DataFrame:
    _NA_STRINGS = {
        "": np.nan, "-": np.nan, "–": np.nan, "—": np.nan,
        "nan": np.nan, "NaN": np.nan, "NAN": np.nan,
        "n/a": np.nan, "N/A": np.nan, "na": np.nan, "NA": np.nan, "null": np.nan, "NULL": np.nan
    }
    if df.empty:
        return df
    def _norm_name(x: str) -> str:
        s = str(x).strip()
        s = " ".join(s.split())
        return s
    df = df.rename(columns={c: _norm_name(c) for c in df.columns})
    canon_map = {}
    for c in list(df.columns):
        lc = c.lower()
        if lc == "hc in wip":
            canon_map[c] = "HC in WIP"
        elif lc in ("open complaint timeliness", "open complaints timeliness", "complaint timeliness"):
            canon_map[c] = "Open Complaint Timeliness"
        elif lc in ("actual hc used", "actual_hc_used", "actual-hc-used"):
            canon_map[c] = "Actual HC used"
        elif lc in ("people in wip", "people_wip", "people-in-wip", "people_wip_list"):
            canon_map[c] = "People in WIP"
    if canon_map:
        df = df.rename(columns=canon_map)
    if "period_date" in df.columns:
        df["period_date"] = pd.to_datetime(df["period_date"], errors="coerce").dt.normalize()
    if "Open Complaint Timeliness" in df.columns:
        s = (df["Open Complaint Timeliness"]
                .astype(str)
                .str.strip()
                .replace({"": np.nan, "—": np.nan, "-": np.nan}))
        s = s.str.replace("%", "", regex=False).str.replace(",", "", regex=False)
        v = pd.to_numeric(s, errors="coerce")
        if pd.notna(v.max()) and float(v.max()) > 1.5:
            v = v / 100.0
        df["Open Complaint Timeliness"] = v
    for col in ["Total Available Hours", "Completed Hours", "Target Output", "Actual Output",
                "Target UPLH", "Actual UPLH", "HC in WIP", "Actual HC used", "Closures", "Opened"]:
        if col in df.columns:
            s = (
                df[col]
                .astype(str)
                .str.strip()
                .replace(_NA_STRINGS)
            )
            df[col] = pd.to_numeric(s, errors="coerce")  
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    if {"Actual Output", "Target Output"}.issubset(df.columns):
        df["Efficiency vs Target"] = (df["Actual Output"] / df["Target Output"]).replace([np.inf, -np.inf], np.nan)
    if {"Completed Hours", "Total Available Hours"}.issubset(df.columns):
        df["Capacity Utilization"] = (df["Completed Hours"] / df["Total Available Hours"]).replace([np.inf, -np.inf], np.nan)
    return df
def accounted_nonwip_by_person_from_row(row) -> tuple[dict[str, float], dict[str, float]]:
    payload = row.get("non_wip_activities", "[]")
    try:
        activities = json.loads(payload) if isinstance(payload, str) else payload
    except Exception:
        activities = []
    if not isinstance(activities, list) or not activities:
        return {}, {}
    import re
    other_team_key = "OTHERTEAMWIP"
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
        if act_key == other_team_key:
            accounted_other[name] = accounted_other.get(name, 0.0) + hrs
        else:
            accounted_nonother[name] = accounted_nonother.get(name, 0.0) + hrs
    accounted_other = {k: round(v, 2) for k, v in accounted_other.items()}
    accounted_nonother = {k: round(v, 2) for k, v in accounted_nonother.items()}
    return accounted_other, accounted_nonother
def _json_payloads_from_series(series: pd.Series) -> list:
    payloads = []
    for payload in series.dropna().tolist():
        try:
            obj = json.loads(payload) if isinstance(payload, str) else payload
        except Exception:
            continue
        payloads.append(obj)
    return payloads
def _merge_non_wip_by_person_rows(rows: pd.DataFrame) -> str:
    totals: dict[str, float] = {}
    if rows is None or rows.empty or "non_wip_by_person" not in rows.columns:
        return json.dumps(totals)
    for obj in _json_payloads_from_series(rows["non_wip_by_person"]):
        if not isinstance(obj, dict):
            continue
        for person, hours in obj.items():
            name = normalize_person_name(str(person).strip())
            if not name:
                continue
            try:
                hrs = float(hours or 0)
            except Exception:
                hrs = 0.0
            totals[name] = totals.get(name, 0.0) + hrs
    return json.dumps({name: round(hours, 2) for name, hours in sorted(totals.items())})
def _merge_non_wip_activity_rows(rows: pd.DataFrame) -> str:
    activities: list[dict] = []
    if rows is None or rows.empty or "non_wip_activities" not in rows.columns:
        return json.dumps(activities)
    for obj in _json_payloads_from_series(rows["non_wip_activities"]):
        if not isinstance(obj, list):
            continue
        for item in obj:
            if isinstance(item, dict):
                activities.append(item)
    return json.dumps(activities)
def build_ooo_table_from_row(row) -> pd.DataFrame:
    payload = row.get("non_wip_activities", "[]")
    try:
        obj = json.loads(payload) if isinstance(payload, str) else payload
    except Exception:
        obj = []
    if not isinstance(obj, list) or not obj:
        return pd.DataFrame(columns=["Activity", "Day", "Name", "Time"])
    df = pd.DataFrame(obj)
    for c in ["activity", "day", "name", "hours"]:
        if c not in df.columns:
            df[c] = None
    if "days" not in df.columns:
        df["days"] = np.nan
    df["hours"] = pd.to_numeric(df["hours"], errors="coerce")
    df["days"]  = pd.to_numeric(df["days"], errors="coerce")
    df["day_norm"] = (
        df["day"]
        .astype(str)
        .str.strip()
        .replace({"": np.nan, "None": np.nan, "nan": np.nan})
    )
    grp = (
        df.groupby(["activity", "name"], as_index=False)
          .agg(
              hours=("hours", "sum"),
              days_known=("days", lambda s: pd.to_numeric(s, errors="coerce").sum(min_count=1)),
              day_values=("day_norm", lambda s: sorted(set([x for x in s.dropna().unique()]))),
          )
    )
    def _label_row(r):
        n = r["days_known"]
        dv = r["day_values"] or []
        try:
            n_int = int(n) if pd.notna(n) else None
        except Exception:
            n_int = None
        if n_int is not None:
            if n_int > 1:
                return f"{n_int} days"
            if n_int == 1:
                return dv[0] if len(dv) == 1 else "Week"
        if len(dv) > 1:
            return f"{len(dv)} days"
        if len(dv) == 1:
            return dv[0]
        return "Week"
    grp["Day"] = grp.apply(_label_row, axis=1)
    out = (
        grp.rename(columns={"activity": "Activity", "name": "Name", "hours": "HoursRaw"})
           [["Activity", "Day", "Name", "HoursRaw"]]
           .assign(
               Activity=lambda d: d["Activity"].astype(str).str.strip(),
               Name=lambda d: d["Name"].astype(str).map(normalize_person_name),
           )
    )
    out["Time"] = out["HoursRaw"].fillna(0).map(_fmt_hours_minutes)
    out = (
        out[["Activity", "Day", "Name", "Time", "HoursRaw"]]
           .sort_values(["Activity", "Name"])
           .reset_index(drop=True)
    )
    return out
def split_nonwip_activity_minutes(cat: pd.DataFrame) -> pd.DataFrame:
    import re
    import numpy as np
    if cat.empty:
        return cat
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
        compact = re.sub(r"[^a-z0-9]", "", lower)
        if re.fullmatch(r"email(s)?(&|and|/)?im", compact):
            return "Email & IM"
        key = lower
        explicit_map = ACTIVITY_MAP
        if key in explicit_map:
            return explicit_map[key]
        acronym_tokens = {
            "im", "wip", "ooo", "sla", "qa", "hc", "pe", "wfh", "pto",
            "ri", "capa",
        }
        words = lower.split(" ")
        if len(words) == 1:
            w = words[0]
            if w.endswith("s") and not w.endswith("ss") and len(w) > 3:
                w = w[:-1]  # emails -> email
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
            label = re.sub(r"\([^)]*$", "", label)       # half-open "( ... "
            label = re.sub(r"\(.*?\)", "", label)        # full "( ... )"
            label = re.sub(r"[:\-–—]+$", "", label)      # trailing punctuation
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
    out = pd.DataFrame(rows)
    if out.empty:
        return cat
    out["Activity"] = out["Activity"].map(_canon_activity)
    return out.groupby("Activity", as_index=False)["Hours"].sum()
def ahu_person_share_for_week(frame: pd.DataFrame, week, teams_in_view: list[str], people_df: pd.DataFrame) -> pd.DataFrame:
    if frame.empty or "Actual HC used" not in frame.columns:
        return pd.DataFrame(columns=["team", "period_date", "person", "percent"])
    wk = pd.to_datetime(week, errors="coerce").normalize()
    if pd.isna(wk):
        return pd.DataFrame(columns=["team", "period_date", "person", "percent"])
    ppl = explode_people_in_wip(frame)
    out_rows: list[dict] = []
    for team in teams_in_view:
        team_ahu_series = (
            frame.loc[(frame["team"] == team) & (frame["period_date"] == wk), "Actual HC used"]
            .dropna()
        )
        if team_ahu_series.empty:
            continue
        per_df = None
        if people_df is not None and not people_df.empty:
            teamw = people_df.loc[
                (people_df["team"] == team) & (people_df["period_date"] == wk)
            ]
            if not teamw.empty and teamw["Actual Hours"].notna().any():
                g = teamw.groupby("person", as_index=False)["Actual Hours"].sum()
                tot = float(g["Actual Hours"].sum())
                if tot > 0:
                    per_df = g.assign(percent=lambda d: d["Actual Hours"] / tot)[["person", "percent"]]
        if per_df is None:
            sub = ppl.loc[(ppl["team"] == team) & (ppl["period_date"] == wk)]
            if not sub.empty:
                unique_people = sub["person"].dropna().drop_duplicates().tolist()
                n = len(unique_people)
                if n > 0:
                    per_df = pd.DataFrame({"person": unique_people, "percent": [1.0 / n] * n})
        if per_df is None or per_df.empty:
            continue
        per_df = per_df.assign(team=team, period_date=wk)
        out_rows.append(per_df[["team", "period_date", "person", "percent"]])
    if not out_rows:
        return pd.DataFrame(columns=["team", "period_date", "person", "percent"])
    return pd.concat(out_rows, ignore_index=True)
@st.cache_data(show_spinner=False, ttl=15 * 60)
def explode_outputs_json(df: pd.DataFrame, col_name: str, key_label: str) -> pd.DataFrame:
    cols = ["team", "period_date", key_label, "Actual", "Target"]
    if df.empty or col_name not in df.columns:
        return pd.DataFrame(columns=cols)
    def _bad_key(k: str) -> bool:
        s = str(k).strip()
        return s in {"", "-", "–", "—"}
    rows: list[dict] = []
    sub = df.loc[:, ["team", "period_date", col_name]].dropna(subset=[col_name]).copy()
    for _, r in sub.iterrows():
        payload = r[col_name]
        try:
            obj = json.loads(payload) if isinstance(payload, str) else payload
        except Exception:
            continue
        if not isinstance(obj, dict):
            continue
        for k, vals in obj.items():
            if _bad_key(k):
                continue
            if isinstance(vals, dict):
                outv = (vals.get("output", None) if "output" in vals else
                        vals.get("actual", None) if "actual" in vals else
                        vals.get("Actual", None))
                tgtv = vals.get("target", vals.get("Target"))
            else:
                outv, tgtv = vals, np.nan
            outv = pd.to_numeric(outv, errors="coerce")
            tgtv = pd.to_numeric(tgtv, errors="coerce")
            if pd.isna(outv) and pd.isna(tgtv):
                continue
            val = str(k).strip()
            if key_label == "person":
                val = normalize_person_name(val)
            rows.append({
                "team": r["team"],
                "period_date": pd.to_datetime(r["period_date"], errors="coerce").normalize(),
                key_label: str(k).strip(),
                "Actual": outv,
                "Target": tgtv
            })
    out = pd.DataFrame(rows, columns=cols)
    if not out.empty:
        out["period_date"] = pd.to_datetime(out["period_date"], errors="coerce").dt.normalize()
    return out
@st.cache_data(show_spinner=False, ttl=15 * 60)
def explode_people_in_wip(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "People in WIP" not in df.columns:
        return pd.DataFrame(columns=["team", "period_date", "person"])
    sub = df.loc[:, ["team", "period_date", "People in WIP"]].dropna(subset=["People in WIP"]).copy()
    BAD_NAMES = {"", "-", "–", "—", "nan", "NaN", "NAN", "n/a", "N/A", "na", "NA", "null", "NULL", "none", "None"}
    def _is_good_name(s: str) -> bool:
        return s.strip() and s.strip() not in BAD_NAMES
    rows: list[dict] = []
    def _as_names(x) -> list[str]:
        if isinstance(x, list):
            return [str(s).strip() for s in x if _is_good_name(str(s))]
        if isinstance(x, str):
            s = x.strip()
            try:
                obj = json.loads(s)
                if isinstance(obj, list):
                    return [str(v).strip() for v in obj if _is_good_name(str(v))]
                if isinstance(obj, dict):
                    return [str(k).strip() for k, v in obj.items() if _is_good_name(str(k))]
            except Exception:
                pass
            import re
            parts = [p.strip() for p in re.split(r"[,;\n\r]+", s) if _is_good_name(p)]
            return parts
        if isinstance(x, dict):
            return [str(k).strip() for k in x.keys() if _is_good_name(str(k))]
        return []
    import re
    for _, r in sub.iterrows():
        people = _as_names(r["People in WIP"])
        for person in people:
            rows.append({
                "team": r["team"],
                "period_date": pd.to_datetime(r["period_date"], errors="coerce").normalize(),
                "person": normalize_person_name(person)
            })
    out = pd.DataFrame(rows, columns=["team", "period_date", "person"])
    if not out.empty:
        out["period_date"] = pd.to_datetime(out["period_date"], errors="coerce").dt.normalize()
    return out
@st.cache_data(show_spinner=False, ttl=15 * 60)
def explode_person_hours(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "Person Hours" not in df.columns:
        return pd.DataFrame(columns=[
            "team","period_date","person","Actual Hours","Available Hours","Utilization"
        ])
    BAD_NAMES = {"", "-", "–", "—", "nan", "NaN", "NAN", "n/a", "N/A", "na", "NA",
                 "null", "NULL", "none", "None"}
    def _is_good_name(s: str) -> bool:
        t = str(s).strip()
        return t and t not in BAD_NAMES and t.lower() not in {b.lower() for b in BAD_NAMES}
    rows: list[dict] = []
    sub = df.loc[:, ["team", "period_date", "Person Hours"]].dropna(subset=["Person Hours"]).copy()
    for _, r in sub.iterrows():
        payload = r["Person Hours"]
        try:
            obj = json.loads(payload) if isinstance(payload, str) else payload
            if not isinstance(obj, dict):
                continue
        except Exception:
            continue
        for person, vals in obj.items():
            if not _is_good_name(person):
                continue
            a = pd.to_numeric((vals or {}).get("actual"), errors="coerce")
            t = pd.to_numeric((vals or {}).get("available"), errors="coerce")
            a = float(a) if pd.notna(a) else 0.0
            t = float(t) if pd.notna(t) else 0.0
            if (a == 0.0) and (t == 0.0):
                continue
            util = (a / t) if t not in (0, 0.0) else np.nan
            rows.append({
                "team": r["team"],
                "period_date": pd.to_datetime(r["period_date"], errors="coerce").normalize(),
                "person": normalize_person_name(str(person).strip()),
                "Actual Hours": a,
                "Available Hours": t,
                "Utilization": util
            })
    out = pd.DataFrame(
        rows,
        columns=["team","period_date","person","Actual Hours","Available Hours","Utilization"]
    )
    if not out.empty:
        out["period_date"] = pd.to_datetime(out["period_date"], errors="coerce").dt.normalize()
    return out
def build_person_weekly_accounting(
    team: str,
    week,
    nw_row,
    metrics_frame: pd.DataFrame,
    nw_frame: pd.DataFrame,
    week_hours: float = 40.0,
    irl_people: set[str] | None = None,
) -> pd.DataFrame:
    wk = pd.to_datetime(week, errors="coerce").normalize()
    long_nw = explode_non_wip_by_person(nw_frame)
    nw_people = long_nw.loc[
        (long_nw["team"] == team) & (long_nw["period_date"] == wk),
        ["person", "Non-WIP Hours"]
    ].copy()
    if nw_people.empty:
        nw_people = pd.DataFrame(columns=["person", "Non-WIP Hours"])
    nw_people["person"] = nw_people["person"].astype(str).str.strip()
    nw_people["Non-WIP Hours"] = pd.to_numeric(nw_people["Non-WIP Hours"], errors="coerce").fillna(0.0)
    person_hours = explode_person_hours(metrics_frame)
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
        .fillna(0.0)
    )
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
def _find_first_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    for c in candidates:
        if c in df.columns:
            return c
    return None
@st.cache_data(show_spinner=False, ttl=15 * 60)
def explode_cell_station_hours(df: pd.DataFrame) -> pd.DataFrame:
    col = _find_first_col(
        df,
        ["Cell/Station Hours", "Cell Station Hours", "Hours by Cell/Station", "Cell Hours", "Station Hours"]
    )
    if not col or df.empty or col not in df.columns:
        return pd.DataFrame(columns=["team", "period_date", "cell_station", "Actual Hours", "Available Hours"])
    sub = df.loc[:, ["team", "period_date", col]].dropna(subset=[col]).copy()
    rows: list[dict] = []
    for _, r in sub.iterrows():
        payload = r[col]
        try:
            obj = json.loads(payload) if isinstance(payload, str) else payload
            if not isinstance(obj, dict):
                continue
        except Exception:
            continue
        for cell, vals in obj.items():
            if isinstance(vals, dict):
                a = pd.to_numeric((vals or {}).get("actual"), errors="coerce")
                t = pd.to_numeric((vals or {}).get("available"), errors="coerce")
            else:
                a = pd.to_numeric(vals, errors="coerce")
                t = np.nan
            rows.append({
                "team": r["team"],
                "period_date": pd.to_datetime(r["period_date"], errors="coerce").normalize(),
                "cell_station": str(cell).strip(),
                "Actual Hours": a,
                "Available Hours": t,
            })
    out = pd.DataFrame(rows)
    if not out.empty:
        out["period_date"] = pd.to_datetime(out["period_date"], errors="coerce").dt.normalize()
    return out
def _maybe_as_float(x):
    try:
        return float(x)
    except Exception:
        return np.nan
@st.cache_data(show_spinner=False, ttl=15 * 60)
def explode_outputs_by_cell_person(df: pd.DataFrame, team: str) -> pd.DataFrame:
    col = _find_first_col(
        df,
        [
            "Output by Cell/Station - by person",  # <- NEW: your per-person outputs column
            "Outputs by Cell/Station",             # station totals
        ]
    )
    cols = ["team","period_date","cell_station","person","Actual","Target"]
    if df.empty or col not in df.columns:
        return pd.DataFrame(columns=cols)
    sub = df.loc[df["team"] == team, ["team","period_date", col]].dropna(subset=[col]).copy()
    rows = []
    for _, r in sub.iterrows():
        payload = r[col]
        try:
            obj = json.loads(payload) if isinstance(payload, str) else payload
        except Exception:
            obj = None
        if not isinstance(obj, dict):
            continue
        for station, vals in obj.items():
            if isinstance(vals, dict) and any(isinstance(v, dict) for v in vals.values()):
                for person, pv in vals.items():
                    if not isinstance(pv, dict):
                        continue
                    a = _maybe_as_float(pv.get("actual", pv.get("output", pv.get("Actual"))))
                    t = _maybe_as_float(pv.get("target", pv.get("Target")))
                    if pd.isna(a) and pd.isna(t):
                        continue
                    rows.append({
                        "team": r["team"],
                        "period_date": pd.to_datetime(r["period_date"], errors="coerce").normalize(),
                        "cell_station": str(station).strip(),
                        "person": normalize_person_name(str(person).strip()),
                        "Actual": a,
                        "Target": t,
                    })
    out = pd.DataFrame(rows, columns=cols)
    if not out.empty:
        out["period_date"] = pd.to_datetime(out["period_date"], errors="coerce").dt.normalize()
    return out
@st.cache_data(show_spinner=False, ttl=15 * 60)
def explode_cell_person_hours(df: pd.DataFrame, team: str) -> pd.DataFrame:
    col = _find_first_col(
        df,
        [
            "Hours by Cell/Station - by person",   # <- NEW: your per-person column
            "Hours by Cell Station - by person",   # <- tolerant variant
            "Cell/Station Hours",                  # totals (no person breakdown)
            "Hours by Cell/Station",
            "Cell Station Hours",
            "Cell Hours",
            "Station Hours",
        ]
    )
    cols = ["team","period_date","cell_station","person","Actual Hours","Available Hours"]
    if not col or df.empty or col not in df.columns:
        return pd.DataFrame(columns=cols)
    sub = df.loc[df["team"] == team, ["team","period_date", col]].dropna(subset=[col]).copy()
    rows = []
    for _, r in sub.iterrows():
        payload = r[col]
        try:
            obj = json.loads(payload) if isinstance(payload, str) else payload
        except Exception:
            obj = None
        if not isinstance(obj, dict):
            continue
        for station, per in obj.items():
            if not isinstance(per, dict):
                continue
            if all(not isinstance(v, dict) for v in per.values()):
                for person, hours in per.items():
                    a = _maybe_as_float(hours)
                    if pd.isna(a):
                        continue
                    rows.append({
                        "team": r["team"],
                        "period_date": pd.to_datetime(r["period_date"], errors="coerce").normalize(),
                        "cell_station": str(station).strip(),
                        "person": normalize_person_name(str(person).strip()),
                        "Actual Hours": a,
                        "Available Hours": np.nan,
                    })
                continue
            for person, pv in per.items():
                if not isinstance(pv, dict):
                    continue
                a = _maybe_as_float(pv.get("actual", pv.get("hours")))
                t = _maybe_as_float(pv.get("available", pv.get("target")))
                if pd.isna(a) and pd.isna(t):
                    continue
                rows.append({
                    "team": r["team"],
                    "period_date": pd.to_datetime(r["period_date"], errors="coerce").normalize(),
                    "cell_station": str(station).strip(),
                    "person": normalize_person_name(str(person).strip()),
                    "Actual Hours": a,
                    "Available Hours": t,
                })
    out = pd.DataFrame(rows, columns=cols)
    if not out.empty:
        out["period_date"] = pd.to_datetime(out["period_date"], errors="coerce").dt.normalize()
    return out
@st.cache_data(show_spinner=False, ttl=15 * 60)
def build_person_station_hours_over_time(frame: pd.DataFrame, team: str, person: str) -> pd.DataFrame:
    hrs = explode_cell_person_hours(frame, team)
    if hrs.empty:
        return pd.DataFrame(columns=["period_date", "cell_station", "Actual Hours", "Available Hours"])
    sub = (
        hrs[(hrs["person"] == person)]
        .copy()
        .assign(
            **{
                "Actual Hours": pd.to_numeric(hrs.loc[hrs["person"] == person, "Actual Hours"], errors="coerce"),
                "Available Hours": pd.to_numeric(hrs.loc[hrs["person"] == person, "Available Hours"], errors="coerce"),
            }
        )
        .dropna(subset=["period_date", "cell_station"])
    )
    if not sub.empty:
        sub = (sub
               .groupby(["period_date", "cell_station"], as_index=False)
               .agg({"Actual Hours": "sum", "Available Hours": "sum"}))
    return sub.sort_values(["period_date", "cell_station"])
@st.cache_data(show_spinner=False, ttl=15 * 60)
def build_station_person_hours_over_time(frame: pd.DataFrame, team: str, station: str) -> pd.DataFrame:
    hrs = explode_cell_person_hours(frame, team)
    if hrs.empty:
        return pd.DataFrame(columns=["period_date","person","Actual Hours","Available Hours","Delta","LabelGroup"])
    sub = hrs[(hrs["cell_station"] == station)].copy()
    if sub.empty:
        return pd.DataFrame(columns=["period_date","person","Actual Hours","Available Hours","Delta","LabelGroup"])
    sub["Actual Hours"] = pd.to_numeric(sub["Actual Hours"], errors="coerce")
    sub["Available Hours"] = pd.to_numeric(sub["Available Hours"], errors="coerce")
    sub["Delta"] = sub["Actual Hours"] - sub["Available Hours"]
    sub["LabelGroup"] = np.where(
        sub["Available Hours"].notna(),
        np.where(sub["Delta"] >= 0, "pos", "neg"),
        "none"
    )
    keep = ["period_date","person","Actual Hours","Available Hours","Delta","LabelGroup"]
    return sub[keep].sort_values(["period_date","person"])
@st.cache_data(show_spinner=False, ttl=15 * 60)
def build_station_person_outputs_over_time(frame: pd.DataFrame, team: str, station: str) -> pd.DataFrame:
    cp = explode_outputs_by_cell_person(frame, team)
    if cp.empty:
        return pd.DataFrame(columns=["period_date","person","Actual","Target","Delta","LabelGroup"])
    sub = cp[(cp["cell_station"] == station)].copy()
    if sub.empty:
        return pd.DataFrame(columns=["period_date","person","Actual","Target","Delta","LabelGroup"])
    sub["Delta"] = sub["Actual"] - sub["Target"]
    sub["LabelGroup"] = np.where(
        sub["Target"].notna(),
        np.where(sub["Delta"] >= 0, "pos", "neg"),
        "none"
    )
    return sub.sort_values(["period_date","person"])
@st.cache_data(show_spinner=False, ttl=15 * 60)
def build_station_person_uplh_over_time(frame: pd.DataFrame, team: str, station: str) -> pd.DataFrame:
    outs = explode_outputs_by_cell_person(frame, team)
    hrs  = explode_cell_person_hours(frame, team)
    if outs.empty or hrs.empty:
        return pd.DataFrame(columns=[
            "period_date","person","Actual","Target","Actual Hours","Actual UPLH","Target UPLH","Delta","LabelGroup"
        ])
    m = (
        outs[outs["cell_station"] == station]
        .merge(hrs[hrs["cell_station"] == station][["period_date","person","Actual Hours"]],
               on=["period_date","person"], how="left")
        .dropna(subset=["Actual"])
    )
    if m.empty:
        return pd.DataFrame(columns=[
            "period_date","person","Actual","Target","Actual Hours","Actual UPLH","Target UPLH","Delta","LabelGroup"
        ])
    m["Actual UPLH"] = (m["Actual"] / m["Actual Hours"]).replace([np.inf, -np.inf], np.nan)
    m["Target UPLH"] = (m["Target"] / m["Actual Hours"]).replace([np.inf, -np.inf], np.nan)
    m["Delta"] = m["Actual UPLH"] - m["Target UPLH"]
    m["LabelGroup"] = np.where(
        m["Target UPLH"].notna(),
        np.where(m["Delta"] >= 0, "pos", "neg"),
        "none"
    )
    keep = ["period_date","person","Actual","Target","Actual Hours","Actual UPLH","Target UPLH","Delta","LabelGroup"]
    return m[keep].sort_values(["period_date","person"])
@st.cache_data(show_spinner=False, ttl=15 * 60)
def build_uplh_by_person_long(frame: pd.DataFrame, team: str) -> pd.DataFrame:
    outp = explode_outputs_json(frame[frame["team"] == team], "Outputs by Person", "person")
    if outp.empty:
        return pd.DataFrame(columns=[
            "team", "period_date", "person",
            "Actual Output", "Target Output", "Actual Hours",
            "Actual UPLH", "Target UPLH"
        ])
    hrs = explode_person_hours(frame[frame["team"] == team])[["period_date", "person", "Actual Hours"]]
    m = (
        outp.merge(hrs, on=["period_date", "person"], how="left")
            .rename(columns={"Actual": "Actual Output", "Target": "Target Output"})
    )
    m["Actual UPLH"] = (m["Actual Output"] / m["Actual Hours"]).replace([np.inf, -np.inf], np.nan)
    m["Target UPLH"] = (m["Target Output"] / m["Actual Hours"]).replace([np.inf, -np.inf], np.nan)
    m["team"] = team
    cols = ["team", "period_date", "person", "Actual Output", "Target Output", "Actual Hours", "Actual UPLH", "Target UPLH"]
    return m[cols].dropna(subset=["Actual Hours"]) 
@st.cache_data(show_spinner=False, ttl=15 * 60)
def build_uplh_by_cell_long(frame: pd.DataFrame, team: str) -> pd.DataFrame:
    outc = explode_outputs_json(frame[frame["team"] == team], "Outputs by Cell/Station", "cell_station")
    if outc.empty:
        return pd.DataFrame(columns=[
            "team", "period_date", "cell_station",
            "Actual Output", "Target Output", "Actual Hours",
            "Actual UPLH", "Target UPLH"
        ])
    hc = explode_cell_station_hours(frame[frame["team"] == team])[["period_date", "cell_station", "Actual Hours"]]
    m = (
        outc.merge(hc, on=["period_date", "cell_station"], how="left")
            .rename(columns={"Actual": "Actual Output", "Target": "Target Output"})
    )
    m["Actual UPLH"] = (m["Actual Output"] / m["Actual Hours"]).replace([np.inf, -np.inf], np.nan)
    m["Target UPLH"] = (m["Target Output"] / m["Actual Hours"]).replace([np.inf, -np.inf], np.nan)
    m["team"] = team
    cols = ["team", "period_date", "cell_station", "Actual Output", "Target Output", "Actual Hours", "Actual UPLH", "Target UPLH"]
    return m[cols].dropna(subset=["Actual Hours"])
data_path = None if DATA_URL else str(DEFAULT_DATA_PATH)
mtime_key = 0
if data_path:
    p = Path(data_path)
    mtime_key = p.stat().st_mtime if p.exists() else 0
df = load_data(data_path, DATA_URL)
def _first_valid_team(value, options):
    if value in options:
        return value
    return options[0] if options else None
def _first_valid_subgroup(value, team_group):
    opts = subgroup_options_for_team(team_group)
    if value in opts:
        return value
    return "All" if "All" in opts else (opts[0] if opts else "All")
if "selected_team_subgroup" not in st.session_state:
    st.session_state.selected_team_subgroup = st.session_state.get("nw_team_subgroup", "All")
_all_wip_teams = (
    sorted([t for t in df["team"].dropna().unique()])
    if not df.empty and "team" in df.columns
    else []
)
if "selected_team" not in st.session_state:
    existing_teams = st.session_state.get("teams_sel", [])
    if existing_teams:
        st.session_state.selected_team = existing_teams[0]
    elif _all_wip_teams:
        st.session_state.selected_team = _all_wip_teams[0]
wip_group_df = add_team_group_columns(load_wip_group_metrics())
nonwip_group_df = add_team_group_columns(load_nonwip_group_metrics())
df = add_team_group_columns(df)
def kpi_card(
    container,
    label: str,
    value,
    fmt: str | None = None,
    color: str | None = None,
    help: str | None = None,
    subtext: str | None = None,
):
    if pd.isna(value):
        val_html = "—"
    else:
        try:
            val_html = (fmt or "{}").format(value)
        except Exception:
            val_html = str(value)
    help_icon = f"""<span title="{help}" style="cursor:help;margin-left:6px;color:#9ca3af;">ⓘ</span>""" if help else ""
    value_color = color or "#111827"
    subtext_html = f"""<div style="font-size:12px;color:#6b7280;margin-top:4px;">{subtext}</div>""" if subtext else ""
    container.markdown(
        f"""
        <div style="padding:12px 16px;border-radius:10px;border:1px solid #eee;">
          <div style="font-size:12px;color:#6b7280;display:flex;align-items:center;gap:4px;">
            <span>{label}</span>{help_icon}
          </div>
          <div style="font-size:28px;font-weight:700;color:{value_color};">{val_html}</div>
          {subtext_html}
        </div>
        """,
        unsafe_allow_html=True,
    )
def _capacity_subtext(hours_val, capacity_val) -> str | None:
    if pd.isna(hours_val) or pd.isna(capacity_val) or float(capacity_val) <= 0:
        return None
    pct = float(hours_val) / float(capacity_val)
    hrs_per_day = pct * 8.0
    return f"{pct:.1%} of capacity • {hrs_per_day:.1f}h/day"
def merged_people_count_for_week(team: str, week, metrics_frame: pd.DataFrame, nw_frame: pd.DataFrame) -> int:
    wk = pd.to_datetime(week, errors="coerce").normalize()
    if nw_frame is not None and not nw_frame.empty:
        raw_nw = nw_frame.copy()
        raw_nw["period_date"] = pd.to_datetime(raw_nw["period_date"], errors="coerce").dt.normalize()
        if "people_count" in raw_nw.columns:
            team_match = raw_nw.loc[
                (raw_nw["team"] == team) & (raw_nw["period_date"] == wk),
                "people_count",
            ]
            team_match = pd.to_numeric(team_match, errors="coerce").dropna()
            if not team_match.empty:
                return int(team_match.iloc[0])
    a = explode_non_wip_by_person(nw_frame)
    b = explode_person_hours(metrics_frame)
    c = explode_people_in_wip(metrics_frame)
    names = set()
    for df_, person_col in [(a, "person"), (b, "person"), (c, "person")]:
        sub = df_.loc[
            (df_["team"] == team) & (df_["period_date"] == wk),
            [person_col]
        ].copy()
        if not sub.empty:
            vals = (
                sub[person_col]
                .astype(str)
                .map(normalize_person_name)
                .str.strip()
            )
            names.update(x for x in vals if x)
    return len(names)
def percent_color(v: float | None, threshold: float, invert: bool = False) -> str:
    if v is None or pd.isna(v):
        return "#111827"
    good = (v >= threshold) if not invert else (v <= threshold)
    return "#22c55e" if good else "#ef4444"
SURGICAL_OVERVIEW_TRIGGER_TEAMS = {
    "Surgical Legal",
    "Surgical AST-GST",
    "Surgical Robotics",
    "MEIC MIR",
}
SURGICAL_ROLLUPS = {
    "Surgical US": [
        ("Surgical Legal", "All"),
        ("Surgical Robotics", "US"),
        ("Surgical AST-GST", "US"),
    ],
    "Surgical MEIC": [
        ("MEIC MIR", "All"),
        ("Surgical Robotics", "MEIC"),
        ("Surgical AST-GST", "MEIC"),
    ],
}
SURGICAL_ROLLUP_FALLBACKS = {
    "Surgical US": {"people_count": 19.0, "capacity_hours": 760.0},
    "Surgical MEIC": {"people_count": 25.0, "capacity_hours": 1000.0},
}
def _build_rollup_kpi(
    rollup_name: str,
    rollup_parts: list[tuple[str, str]],
    week: pd.Timestamp,
    nw_group_frame: pd.DataFrame,
    wip_group_frame: pd.DataFrame,
    include_ooo_in_kpi_pct: bool,
) -> dict:
    nw_rows = []
    wip_rows = []
    for team_group, subgroup in rollup_parts:
        nw_rows.append(filter_team_view(nw_group_frame, team_group, subgroup, fallback_to_all=False))
        wip_rows.append(filter_team_view(wip_group_frame, team_group, subgroup, fallback_to_all=False))
    nw_all = pd.concat(nw_rows, ignore_index=True) if nw_rows else pd.DataFrame()
    wip_all = pd.concat(wip_rows, ignore_index=True) if wip_rows else pd.DataFrame()
    nw_week = nw_all[pd.to_datetime(nw_all["period_date"], errors="coerce").dt.normalize() == week].copy()
    wip_week = wip_all[pd.to_datetime(wip_all["period_date"], errors="coerce").dt.normalize() == week].copy()
    if nw_week.empty:
        return {}
    row = nw_week.iloc[0].copy()
    row["total_non_wip_hours"] = pd.to_numeric(nw_week.get("total_non_wip_hours"), errors="coerce").fillna(0.0).sum()
    row["OOO Hours"] = pd.to_numeric(nw_week.get("OOO Hours"), errors="coerce").fillna(0.0).sum()
    row["people_count"] = pd.to_numeric(nw_week.get("people_count"), errors="coerce").fillna(0.0).sum()
    if "% in WIP" in nw_week.columns:
        row["% in WIP"] = pd.to_numeric(nw_week["% in WIP"], errors="coerce").mean()
    row["non_wip_by_person"] = _merge_non_wip_by_person_rows(nw_week)
    row["non_wip_activities"] = _merge_non_wip_activity_rows(nw_week)
    teams_cfg = load_team_config()
    rollup_irl_people = set()
    for team_group, _sub in rollup_parts:
        rollup_irl_people.update(irl_people_for_team(team_group, teams_cfg))
    wk_people = build_person_weekly_accounting(
        team="Surgical Rollup",
        week=week,
        nw_row=row,
        metrics_frame=wip_week,
        nw_frame=nw_week,
        week_hours=40.0,
        irl_people=rollup_irl_people,
    )
    rolled_non_wip_hours = float(pd.to_numeric(nw_week.get("total_non_wip_hours"), errors="coerce").fillna(0.0).sum())
    rolled_ooo_hours = float(pd.to_numeric(nw_week.get("OOO Hours"), errors="coerce").fillna(0.0).sum())
    rolled_people_count = float(pd.to_numeric(nw_week.get("people_count"), errors="coerce").fillna(0.0).sum())
    fallback_cfg = SURGICAL_ROLLUP_FALLBACKS.get(rollup_name, {})
    fallback_people_count = float(fallback_cfg.get("people_count", 0.0)) if fallback_cfg else 0.0
    if fallback_people_count > 0 and rolled_people_count < fallback_people_count:
        rolled_people_count = fallback_people_count
    rolled_wip_hours = float(pd.to_numeric(wip_week.get("Completed Hours"), errors="coerce").fillna(0.0).sum())
    kpi = enterprise_nonwip_kpi_lookup(
        team="Surgical Rollup",
        week=week,
        nw_row=row,
        wk_people=wk_people,
        people_count=rolled_people_count,
        completed_hours=rolled_wip_hours,
        total_non_wip_hours=rolled_non_wip_hours,
        factor_out_ooo=not include_ooo_in_kpi_pct,
        nw_frame=nw_week,
        metrics_frame=wip_week,
    )
    capacity_hours = kpi.get("capacity_hours", np.nan)
    fallback_capacity_hours = float(fallback_cfg.get("capacity_hours", 0.0)) if fallback_cfg else 0.0
    if fallback_capacity_hours > 0:
        if pd.isna(capacity_hours) or float(capacity_hours) < fallback_capacity_hours:
            capacity_hours = fallback_capacity_hours
            kpi["capacity_hours"] = capacity_hours
    pct_denom = capacity_hours
    if not include_ooo_in_kpi_pct and pd.notna(capacity_hours):
        pct_denom = max(float(capacity_hours) - float(rolled_ooo_hours), 0.0)
    unaccounted_hours = (
        max(float(capacity_hours) - (rolled_wip_hours + rolled_non_wip_hours + rolled_ooo_hours), 0.0)
        if pd.notna(capacity_hours)
        else np.nan
    )
    kpi.update({
        "people_count": rolled_people_count,
        "completed_hours": rolled_wip_hours,
        "non_wip_hours": rolled_non_wip_hours,
        "ooo_hours": rolled_ooo_hours,
        "unaccounted_hours": unaccounted_hours,
        "pct_denom": pct_denom,
        "wip_pct": (rolled_wip_hours / pct_denom) if pd.notna(pct_denom) and pct_denom > 0 else np.nan,
        "non_wip_pct": (rolled_non_wip_hours / pct_denom) if pd.notna(pct_denom) and pct_denom > 0 else np.nan,
        "ooo_pct": ((rolled_ooo_hours / pct_denom) if include_ooo_in_kpi_pct and pd.notna(pct_denom) and pct_denom > 0 else 0.0),
        "unaccounted_pct": (unaccounted_hours / pct_denom) if pd.notna(unaccounted_hours) and pd.notna(pct_denom) and pct_denom > 0 else np.nan,
    })
    return kpi
st.markdown("<h1 style='text-align: center;'>MS Heijunka Metrics Dashboard</h1>", unsafe_allow_html=True)
label = "Show WIP view" if st.session_state.get("nonwip_mode", False) else "Show Non-WIP view"
nonwip_mode = st.toggle(
    label,
    value=st.session_state.get("nonwip_mode", False),
    key="nonwip_mode",
    help="Switch between WIP and Non-WIP metrics"
)
if nonwip_mode:
    nw_base = add_team_group_columns(load_non_wip())
    if nw_base.empty:
        st.info("No Non-WIP data found yet. Make sure non_wip_activities.csv exists.")
        st.stop()
    st.markdown("### Non-WIP Overview")
    teams_nw = grouped_team_options(nw_base)
    c_team, c_week = st.columns(2)
    preferred_team = _first_valid_team(
        st.session_state.get("selected_team"),
        teams_nw,
    )
    if preferred_team is not None:
        st.session_state["nw_team"] = preferred_team
    preferred_subgroup = _first_valid_subgroup(
        st.session_state.get("selected_team_subgroup", "All"),
        preferred_team,
    )
    st.session_state["nw_team_subgroup"] = preferred_subgroup
    def _sync_from_nonwip_team():
        team = st.session_state.get("nw_team")
        if team:
            st.session_state.selected_team = team
            st.session_state.teams_sel = [team]
        st.session_state.selected_team_subgroup = _first_valid_subgroup(
            st.session_state.get("nw_team_subgroup", "All"),
            team,
        )
        st.session_state.nw_team_subgroup = st.session_state.selected_team_subgroup
    def _sync_from_nonwip_subgroup():
        st.session_state.selected_team_subgroup = st.session_state.get("nw_team_subgroup", "All")
    with c_team:
        team_nw = st.selectbox(
            "Team",
            options=teams_nw,
            key="nw_team",
            on_change=_sync_from_nonwip_team,
        )
        subgroup_opts = subgroup_options_for_team(team_nw)
        has_extra_groups = len(subgroup_opts) > 1
        if has_extra_groups:
            st.session_state["nw_team_subgroup"] = _first_valid_subgroup(
                st.session_state.get("selected_team_subgroup", "All"),
                team_nw,
            )
            subgroup_nw = st.selectbox(
                "Group:",
                options=subgroup_opts,
                key="nw_team_subgroup",
                on_change=_sync_from_nonwip_subgroup,
            )
        else:
            subgroup_nw = "All"
            st.session_state.selected_team_subgroup = "All"
    st.session_state.selected_team = team_nw
    st.session_state.teams_sel = [team_nw]
    st.session_state.selected_team_subgroup = subgroup_nw
    nw_source = nw_base if subgroup_nw == "All" else nonwip_group_df
    nw_view = filter_team_view(nw_source, team_nw, subgroup_nw)
    nw = nw_source
    today_nw = pd.Timestamp.today().normalize()
    weeks_nw = sorted(
        [
            d for d in pd.to_datetime(nw_view["period_date"].dropna().unique())
            if pd.notna(d) and pd.to_datetime(d).normalize() <= today_nw
        ],
        reverse=True
    )
    if not weeks_nw:
        st.info("No weeks available for this team up to today.")
        st.stop()
    with c_week:
        week_nw = st.selectbox(
            "Week",
            options=weeks_nw,
            index=0,
            format_func=lambda d: pd.to_datetime(d).date().isoformat(),
            key="nw_week",
        )
    week_nw = pd.to_datetime(week_nw).normalize()
    sel = nw_view[nw_view["period_date"] == week_nw].copy()
    if sel.empty:
        st.info("No Non-WIP row for that team/week.")
        st.stop()
    if len(sel) > 1:
        agg_row = sel.iloc[0].copy()
        numeric_sum_cols = ["people_count", "total_non_wip_hours", "OOO Hours"]
        for col in numeric_sum_cols:
            if col in sel.columns:
                agg_row[col] = pd.to_numeric(sel[col], errors="coerce").sum()
        if "% in WIP" in sel.columns:
            pct_series = pd.to_numeric(sel["% in WIP"], errors="coerce")
            agg_row["% in WIP"] = pct_series.mean()
        if "% Non-WIP" in sel.columns:
            pct_nw_series = pd.to_numeric(sel["% Non-WIP"], errors="coerce")
            agg_row["% Non-WIP"] = pct_nw_series.mean()
        agg_row["non_wip_by_person"] = _merge_non_wip_by_person_rows(sel)
        agg_row["non_wip_activities"] = _merge_non_wip_activity_rows(sel)
        agg_row["team"] = team_nw if subgroup_nw == "All" else f"{team_nw} - {subgroup_nw}"
        agg_row["team_group"] = team_nw
        agg_row["team_subgroup"] = subgroup_nw
        sel = pd.DataFrame([agg_row])
    row = sel.iloc[0]
    if "% Non-WIP" in row.index and pd.notna(row["% Non-WIP"]):
        pct_non_wip = float(row["% Non-WIP"])
    else:
        pct_in_wip = float(row.get("% in WIP", np.nan))
        pct_non_wip = (100.0 - pct_in_wip) if pd.notna(pct_in_wip) else np.nan
    include_ooo_in_kpi_pct = st.toggle(
        "Include OOO Hours in KPI % of capacity",
        value=False,
        key="include_ooo_in_kpi_pct",
        help="When off, OOO Hours shows 0.0% of capacity and other KPI percentages are calculated against capacity excluding OOO hours.",
    )
    def colored_percent_metric(container, label: str, value: float | None, threshold=80.0):
        if pd.isna(value):
            container.metric(label, "—")
            return
        color = "#ef4444" if float(value) < threshold else "#22c55e"
        container.markdown(
            f"""
            <div style="padding:12px 16px;border-radius:10px;border:1px solid #eee;">
            <div style="font-size:12px;color:#6b7280;">{label}</div>
            <div style="font-size:28px;font-weight:700;color:{color};">{value:.2f}%</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    c1, c2, c3, c4 = st.columns(4)
    if subgroup_nw == "All":
        wip_match = df[(df["team"] == team_nw) & (df["period_date"] == week_nw)]
    else:
        wip_match = filter_team_view(wip_group_df, team_nw, subgroup_nw, fallback_to_all=False)
        wip_match = wip_match[wip_match["period_date"] == week_nw]
    wip_hours_val = (
        float(pd.to_numeric(wip_match["Completed Hours"], errors="coerce").sum())
        if not wip_match.empty and "Completed Hours" in wip_match.columns
        else np.nan
    )
    if subgroup_nw == "All":
        metrics_frame_for_accounting = df
    else:
        wip_group_filtered = filter_team_view(wip_group_df, team_nw, subgroup_nw, fallback_to_all=False).copy()
        metrics_frame_for_accounting = wip_group_filtered
        metrics_frame_for_accounting["team"] = team_nw
    if subgroup_nw == "All":
        nw_view_for_accounting = nw_base
    else:
        nw_view_for_accounting = nw_view.copy()
        nw_view_for_accounting["team"] = team_nw
    people_count_val = merged_people_count_for_week(team_nw, week_nw, df, nw)
    teams_cfg = load_team_config()
    team_irl_people = irl_people_for_team(team_nw, teams_cfg)
    wk_people_kpi = build_person_weekly_accounting(
        team=team_nw,
        week=week_nw,
        nw_row=row,
        metrics_frame=metrics_frame_for_accounting,
        nw_frame=nw_view_for_accounting, 
        week_hours=40.0,
        irl_people=team_irl_people,
    )
    if not wk_people_kpi.empty and "Expected Hours" in wk_people_kpi.columns:
        capacity_val = float(pd.to_numeric(wk_people_kpi["Expected Hours"], errors="coerce").fillna(0.0).sum())
    else:
        irl_count = len(team_irl_people)
        total_people = float(people_count_val) if pd.notna(people_count_val) and float(people_count_val) > 0 else np.nan
        if pd.notna(total_people):
            non_irl_count = max(total_people - irl_count, 0.0)
            capacity_val = (irl_count * 39.0) + (non_irl_count * 40.0)
        else:
            capacity_val = np.nan
    nonwip_hours_val = float(pd.to_numeric(row.get("total_non_wip_hours", np.nan), errors="coerce")) \
        if pd.notna(pd.to_numeric(row.get("total_non_wip_hours", np.nan), errors="coerce")) else np.nan
    ooo_hours_val = float(pd.to_numeric(row.get("OOO Hours", np.nan), errors="coerce")) \
        if pd.notna(pd.to_numeric(row.get("OOO Hours", np.nan), errors="coerce")) else 0.0
    used_hours = (
        (0.0 if pd.isna(wip_hours_val) else float(wip_hours_val))
        + (0.0 if pd.isna(nonwip_hours_val) else float(nonwip_hours_val))
        + (0.0 if pd.isna(ooo_hours_val) else float(ooo_hours_val))
    )
    unaccounted_hours_val = (
        max(float(capacity_val) - used_hours, 0.0)
        if pd.notna(capacity_val)
        else np.nan
    )
    capacity_pct_basis = capacity_val
    if not include_ooo_in_kpi_pct and pd.notna(capacity_val):
        capacity_pct_basis = max(float(capacity_val) - float(ooo_hours_val), 0.0)
    wip_pct = (
        wip_hours_val / capacity_pct_basis
        if pd.notna(wip_hours_val) and pd.notna(capacity_pct_basis) and capacity_pct_basis > 0
        else np.nan
    )
    nonwip_pct = (
        nonwip_hours_val / capacity_pct_basis
        if pd.notna(nonwip_hours_val) and pd.notna(capacity_pct_basis) and capacity_pct_basis > 0
        else np.nan
    )
    ooo_pct = (
        (ooo_hours_val / capacity_pct_basis)
        if include_ooo_in_kpi_pct and pd.notna(ooo_hours_val) and pd.notna(capacity_pct_basis) and capacity_pct_basis > 0
        else 0.0
    )
    unaccounted_pct = (
        unaccounted_hours_val / capacity_pct_basis
        if pd.notna(unaccounted_hours_val) and pd.notna(capacity_pct_basis) and capacity_pct_basis > 0
        else np.nan
    )
    _enterprise_kpi = enterprise_nonwip_kpi_lookup(
        team=team_nw,
        week=week_nw,
        nw_row=row,
        wk_people=wk_people_kpi,
        people_count=locals().get("people_count_merged", locals().get("people_count_val", np.nan)),
        completed_hours=wip_hours_val,
        total_non_wip_hours=locals().get("total_nonwip_hours_val", locals().get("nonwip_hours_val", np.nan)),
        factor_out_ooo=not include_ooo_in_kpi_pct,
        person_hours=locals().get("_ppl_hours_kpi"),
        people_in_wip=locals().get("_ppl_in_wip_kpi"),
        nw_frame=nw_view_for_accounting,
        metrics_frame=metrics_frame_for_accounting,
        ent_capacity_callback=globals().get("ent_capacity_hours_for_week"),
        ent_capacity_kwargs={
            "team": team_nw,
            "week": week_nw,
            "nw_frame": nw_view_for_accounting,
            "irl_people": locals().get("team_irl_people", set()),
        },
    )
    capacity_val = _enterprise_kpi["capacity_hours"]
    capacity_pct_basis = _enterprise_kpi["pct_denom"]
    wip_hours_val = _enterprise_kpi["completed_hours"]
    other_team_wip_hours_val = _enterprise_kpi["other_team_wip_hours"]
    nonwip_hours_val = _enterprise_kpi["non_wip_hours"]
    ooo_hours_val = _enterprise_kpi["ooo_hours"]
    unaccounted_hours_val = _enterprise_kpi["unaccounted_hours"]
    wip_pct = _enterprise_kpi["wip_pct"]
    other_team_wip_pct = _enterprise_kpi["other_team_wip_pct"]
    nonwip_pct = _enterprise_kpi["non_wip_pct"]
    ooo_pct = _enterprise_kpi["ooo_pct"]
    unaccounted_pct = _enterprise_kpi["unaccounted_pct"]
    kpi_card(
        c1,
        "WIP Hours",
        wip_hours_val,
        fmt="{:,.1f}",
        color=percent_color(wip_pct, threshold=0.80, invert=False),
        subtext=_capacity_subtext(wip_hours_val, capacity_pct_basis),
    )
    kpi_card(
        c2,
        "Non-WIP Hours",
        nonwip_hours_val,
        fmt="{:,.1f}",
        color=percent_color(nonwip_pct, threshold=0.20, invert=True),
        subtext=_capacity_subtext(nonwip_hours_val, capacity_pct_basis),
    )
    kpi_card(
        c3,
        "OOO Hours",
        ooo_hours_val,
        fmt="{:,.1f}",
        subtext=_capacity_subtext(
            0.0 if not include_ooo_in_kpi_pct else ooo_hours_val,
            capacity_pct_basis,
        ),
    )
    kpi_card(
        c4,
        "Unaccounted Hours",
        unaccounted_hours_val,
        fmt="{:,.1f}",
        subtext=_capacity_subtext(unaccounted_hours_val, capacity_pct_basis),
    )
    if team_nw in SURGICAL_OVERVIEW_TRIGGER_TEAMS:
        with st.popover("Get Regional Surgical Breakdown", use_container_width=False):
            st.markdown("### Regional Surgical Breakdown")
            st.caption(f"Week: {week_nw.date().isoformat()}")
            for rollup_name, rollup_parts in SURGICAL_ROLLUPS.items():
                rollup_kpi = _build_rollup_kpi(
                    rollup_name=rollup_name,
                    rollup_parts=rollup_parts,
                    week=week_nw,
                    nw_group_frame=nonwip_group_df,
                    wip_group_frame=wip_group_df,
                    include_ooo_in_kpi_pct=include_ooo_in_kpi_pct,
                )
                st.markdown(f"#### {rollup_name}")
                if not rollup_kpi:
                    st.info("No data available for this week.")
                    continue
                r1, r2, r3, r4 = st.columns(4)
                kpi_card(
                    r1, "WIP Hours", rollup_kpi["completed_hours"], fmt="{:,.1f}",
                    color=percent_color(rollup_kpi["wip_pct"], threshold=0.80, invert=False),
                    subtext=_capacity_subtext(rollup_kpi["completed_hours"], rollup_kpi["pct_denom"]),
                )
                kpi_card(
                    r2, "Non-WIP Hours", rollup_kpi["non_wip_hours"], fmt="{:,.1f}",
                    color=percent_color(rollup_kpi["non_wip_pct"], threshold=0.20, invert=True),
                    subtext=_capacity_subtext(rollup_kpi["non_wip_hours"], rollup_kpi["pct_denom"]),
                )
                kpi_card(
                    r3, "OOO Hours", rollup_kpi["ooo_hours"], fmt="{:,.1f}",
                    subtext=_capacity_subtext(
                        0.0 if not include_ooo_in_kpi_pct else rollup_kpi["ooo_hours"],
                        rollup_kpi["pct_denom"],
                    ),
                )
                kpi_card(
                    r4, "Unaccounted Hours", rollup_kpi["unaccounted_hours"], fmt="{:,.1f}",
                    subtext=_capacity_subtext(rollup_kpi["unaccounted_hours"], rollup_kpi["pct_denom"]),
                )
                st.markdown("---")
    st.markdown("---")
    st.markdown("#### Non-WIP Activities")
    if "non_wip_activities" not in sel.columns or sel.iloc[0].get("non_wip_activities", "") in ("", "[]", None):
        st.info("No Non-WIP activities recorded for this selection.")
    else:
        act_tbl = build_ooo_table_from_row(sel.iloc[0])
        if act_tbl.empty:
            st.info("No Non-WIP activities recorded for this selection.")
        else:
            display_tbl = act_tbl.drop(columns=["HoursRaw"], errors="ignore")
            st.dataframe(display_tbl, width="stretch", hide_index=True)
    teams_cfg = load_team_config()
    team_irl_people = irl_people_for_team(team_nw, teams_cfg)
    wk_people = build_person_weekly_accounting(
        team=team_nw,
        week=week_nw,
        nw_row=row,
        metrics_frame=metrics_frame_for_accounting,
        nw_frame=nw_view_for_accounting,
        week_hours=40.0,
        irl_people=team_irl_people,
    )
    if wk_people.empty:
        st.info("No per-person weekly breakdown for this selection.")
    else:
        wk_people = wk_people.rename(columns={
            "Other Team WIP": "Accounted_Other",
            "Accounted Non-WIP": "Accounted_NonOther",
        })
        wk_people = wk_people[
            ~wk_people["person"].str.contains("TM", case=False, na=False)
        ].copy()
        stack = (
            wk_people.melt(
                id_vars=["person", "period_date", "Non-WIP Hours", "Completed Hours"],
                value_vars=["OOO Hours","Accounted_Other", "Accounted_NonOther", "Unaccounted"],
                var_name="Category",
                value_name="Hours"
            )
            .dropna(subset=["Hours"])
        )
        stack = stack.merge(
            wk_people[[
                "person",
                "Completed Hours",
                "Non-WIP Hours",
                "OOO Hours",
                "Accounted_Other",
                "Accounted_NonOther",
                "Unaccounted"
            ]],
            on="person",
            how="left",
        )
        label_map = {
            "OOO Hours": "OOO",
            "Accounted_Other": "Other Team WIP",
            "Accounted_NonOther": "Accounted Non-WIP",
            "Unaccounted": "Unaccounted",
        }
        stack["CategoryLabel"] = stack["Category"].map(label_map)
        wk_people["StackTotal"] = (
            wk_people["OOO Hours"].fillna(0)
            + wk_people["Accounted_Other"].fillna(0)
            + wk_people["Accounted_NonOther"].fillna(0)
            + wk_people["Unaccounted"].fillna(0)
        )
        order_people = wk_people.sort_values("StackTotal", ascending=False)["person"].tolist()
        vmax = float(pd.to_numeric(wk_people["StackTotal"], errors="coerce").max())
        headroom = max(1.0, vmax * 0.18) if pd.notna(vmax) else 1.0
        y_scale = alt.Scale(domain=[0, vmax + headroom], nice=False, clamp=False)
        totals = (
            wk_people[["person", "period_date", "StackTotal"]]
            .rename(columns={"StackTotal": "Total"})
            .assign(Status=lambda d: np.where(d["Total"] <= 7.5, "Good (≤7.5)", "Over (>7.5)"))
        )
        outline = (
            alt.Chart(totals)
            .mark_bar(fillOpacity=0, strokeWidth=2)
            .encode(
                x=alt.X("person:N", sort=order_people),
                y=alt.Y("Total:Q", scale=y_scale),
                stroke=alt.Color(
                    "Status:N",
                    title="Total vs 7.5",
                    scale=alt.Scale(
                        domain=["Good (≤7.5)", "Over (>7.5)"],
                        range=["#22c55e", "#ef4444"],
                    ),
                ),
            )
        )
        bars = (
            alt.Chart(stack)
            .mark_bar(clip=False)
            .encode(
                x=alt.X(
                    "person:N",
                    title="Person",
                    axis=alt.Axis(
                        labelAngle=-45,
                        labelLimit=220,
                        labelOverlap=False,
                    ),
                ),
                y=alt.Y(
                    "Hours:Q",
                    title="Non-WIP Hours (week)",
                    stack="zero",
                    scale=y_scale,
                ),
                color=alt.Color(
                    "CategoryLabel:N",
                    title="Legend",
                    scale=alt.Scale(
                        domain=["OOO", "Other Team WIP", "Accounted Non-WIP", "Unaccounted"],
                        range=["#a855f7", "#2563eb", "#22c55e", "#9ca3af"],
                    ),
                ),
                tooltip=[
                    alt.Tooltip("person:N", title="Person"),
                    alt.Tooltip("Accounted_Other:Q", title="Other Team WIP Hours", format=",.2f"),
                    alt.Tooltip("Accounted_NonOther:Q", title="Accounted Non-WIP Hours", format=",.2f"),
                    alt.Tooltip("Unaccounted:Q", title="Unaccounted Hours", format=",.2f"),
                    alt.Tooltip("OOO Hours:Q", title="OOO Hours", format=",.2f"),
                    alt.Tooltip("period_date:T", title="Week"),
                ],
            )
        )
        ref = (
            alt.Chart(pd.DataFrame({"y": [7.5]}))
            .mark_rule(strokeDash=[4, 3], color="#6b7280")
            .encode(y=alt.Y("y:Q", scale=y_scale))
        )
        chart = (outline + ref + bars) \
            .properties(
                height=340,
                padding={"left": 8, "right": 12, "top": 36, "bottom": 64},
            ) \
            .configure_axis(labelOverlap=True) \
            .configure_view(stroke=None)
        st.altair_chart(chart, width="stretch")
        st.markdown("#### Non-WIP Activities")
        if "non_wip_activities" in sel.columns and sel.iloc[0].get("non_wip_activities", "") not in ("", "[]", None):
            act_tbl2 = build_ooo_table_from_row(sel.iloc[0])
            if not act_tbl2.empty and "HoursRaw" in act_tbl2.columns:
                cat = (
                    act_tbl2.groupby("Activity", as_index=False)["HoursRaw"]
                            .sum()
                            .rename(columns={"HoursRaw": "Hours"})
                )
                cat = cat[cat["Activity"].astype(str).str.strip().str.upper() != "OOO"].copy()
                cat = split_nonwip_activity_minutes(cat)
                if not cat.empty:
                    cat = cat.sort_values("Hours", ascending=False)
                    order_acts = cat["Activity"].tolist()
                    act_chart = (
                        alt.Chart(cat)
                        .mark_bar()
                        .encode(
                            x=alt.X(
                                "Activity:N",
                                title="Activity",
                                sort=order_acts,
                                axis=alt.Axis(labelAngle=-30, labelLimit=140),
                            ),
                            y=alt.Y("Hours:Q", title="Total Non-WIP Hours"),
                            tooltip=[
                                alt.Tooltip("Activity:N", title="Activity"),
                                alt.Tooltip("Hours:Q", title="Hours", format=",.2f"),
                            ],
                        )
                        .properties(
                            height=280
                        )
                    )
                    st.altair_chart(act_chart, width="stretch")
    st.markdown("#### Team Trends")
    team_hist = nw_view.copy()
    team_hist = team_hist.dropna(subset=["period_date"]).copy()
    if not team_hist.empty and len(team_hist) > 1:
        agg_map = {}
        if "people_count" in team_hist.columns:
            agg_map["people_count"] = "sum"
        if "total_non_wip_hours" in team_hist.columns:
            agg_map["total_non_wip_hours"] = "sum"
        if "OOO Hours" in team_hist.columns:
            agg_map["OOO Hours"] = "sum"
        if "% in WIP" in team_hist.columns:
            agg_map["% in WIP"] = "mean"
        if "% Non-WIP" in team_hist.columns:
            agg_map["% Non-WIP"] = "mean"
        if agg_map:
            team_hist = (
                team_hist.groupby("period_date", as_index=False)
                .agg(agg_map)
                .sort_values("period_date")
            )
            team_hist["team"] = team_nw if subgroup_nw == "All" else f"{team_nw} - {subgroup_nw}"
            team_hist["team_group"] = team_nw
            team_hist["team_subgroup"] = subgroup_nw
    if not team_hist.empty:
        team_hist["period_date"] = pd.to_datetime(team_hist["period_date"], errors="coerce")
        if "total_non_wip_hours" in team_hist.columns:
            team_hist["total_non_wip_hours"] = pd.to_numeric(team_hist["total_non_wip_hours"], errors="coerce")
        if "% Non-WIP" in team_hist.columns:
            team_hist["% Non-WIP"] = pd.to_numeric(team_hist["% Non-WIP"], errors="coerce")
        trend_hours = team_hist.dropna(subset=["period_date", "total_non_wip_hours"]).copy()
        trend_pct = team_hist.dropna(subset=["period_date", "% Non-WIP"]).copy()
        t1, t2 = st.columns(2)
        with t1:
            if trend_hours.empty:
                st.info("No trendable Non-WIP Hours data available.")
            else:
                ch1 = (
                    alt.Chart(trend_hours)
                    .mark_line(point=True)
                    .encode(
                        x=alt.X("period_date:T", title="Week"),
                        y=alt.Y("total_non_wip_hours:Q", title="Total Non-WIP Hours"),
                        tooltip=[
                            alt.Tooltip("period_date:T", title="Date"),
                            alt.Tooltip("total_non_wip_hours:Q", title="Non-WIP Hours", format=",.1f"),
                        ],
                    )
                    .properties(height=240, title="Total Non-WIP Hours")
                )
                st.altair_chart(ch1, width="stretch")
        with t2:
            if trend_pct.empty:
                st.info("No trendable % Non-WIP data available.")
            else:
                ch2 = (
                    alt.Chart(trend_pct)
                    .mark_line(point=True)
                    .encode(
                        x=alt.X("period_date:T", title="Week"),
                        y=alt.Y("% Non-WIP:Q", title="% Non-WIP"),
                        tooltip=[
                            alt.Tooltip("period_date:T", title="Date"),
                            alt.Tooltip("% Non-WIP:Q", title="% Non-WIP", format=",.2f"),
                        ],
                    )
                    .properties(height=240, title="% Non-WIP")
                )
                st.altair_chart(ch2, width="stretch")
    st.stop()
if df.empty:
    st.warning("No data found yet. Make sure metrics_aggregate_dev.csv exists and has the 'All Metrics' sheet.")
    st.stop()
def _get_qp_teams() -> list[str]:
    try:
        qp = st.query_params
        vals = qp.get_all("teams") if hasattr(qp, "get_all") else qp.get("teams", [])
    except Exception:
        qp = st.experimental_get_query_params()
        vals = qp.get("teams", [])
    if vals is None:
        return []
    if isinstance(vals, str):
        return [vals]
    return [str(v) for v in vals]
def _set_qp_teams(values: list[str]) -> None:
    try:
        st.query_params["teams"] = values
    except Exception:
        st.experimental_set_query_params(teams=values)
def _sets_equal(a, b) -> bool:
    return set(a) == set(b)
teams = grouped_team_options(df)
default_team = _first_valid_team(
    st.session_state.get("selected_team"),
    teams,
)
default_teams = [default_team] if default_team else []
if "teams_sel" not in st.session_state:
    saved = [t for t in teams if t in _get_qp_teams()]
    st.session_state.teams_sel = saved or default_teams
else:
    st.session_state.teams_sel = [
        t for t in st.session_state.teams_sel
        if t in teams
    ] or default_teams
has_dates = df["period_date"].notna().any()
min_date = pd.to_datetime(df["period_date"].min()).date() if has_dates else None
max_date_raw = pd.to_datetime(df["period_date"].max()).date() if has_dates else None
if has_dates and min_date and max_date_raw:
    today_date = pd.Timestamp.today().normalize().date()
    max_date = min(max_date_raw, today_date)
    default_start = pd.to_datetime("2025-10-27").date()
    if "ms_start_date" not in st.session_state:
        st.session_state["ms_start_date"] =max(min_date, default_start)
    if "ms_end_date" not in st.session_state:
        st.session_state["ms_end_date"] = max_date
    else:
        st.session_state["ms_end_date"] = min(st.session_state["ms_end_date"], max_date)
    start = st.session_state["ms_start_date"]
    end = st.session_state["ms_end_date"]
    if start > end:
        st.error("Start date cannot be after end date!")
col1, col2, col3 = st.columns([2, 2, 6], gap="large")
with col1:
    selected_teams = st.multiselect("Teams", teams, key="teams_sel")
if selected_teams:
    st.session_state.selected_team = selected_teams[0]
    st.session_state.selected_team_subgroup = _first_valid_subgroup(
        st.session_state.get("selected_team_subgroup", "All"),
        selected_teams[0],
    )
current_qp = _get_qp_teams()
if not _sets_equal(st.session_state.teams_sel, current_qp):
    _set_qp_teams(sorted(st.session_state.teams_sel))
f = df.copy()
if st.session_state.teams_sel:
    f = f[f["team"].isin(st.session_state.teams_sel)]
if start and end:
    f = f[(f["period_date"] >= pd.to_datetime(start)) & (f["period_date"] <= pd.to_datetime(end))]
if f.empty:
    st.info("No rows match your filters.")
    st.stop()
ppl_hours = explode_person_hours(f)
latest = (f.sort_values(["team", "period_date"])
            .groupby("team", as_index=False)
            .tail(1)
            .copy()
)
tot_target = latest["Target Output"].sum(skipna=True)
tot_actual = latest["Actual Output"].sum(skipna=True)
tot_tahl  = latest["Total Available Hours"].sum(skipna=True)
tot_chl   = latest["Completed Hours"].sum(skipna=True)
tot_hc_wip = latest["HC in WIP"].sum(skipna=True) if "HC in WIP" in latest.columns else np.nan
tot_hc_used = latest["Actual HC used"].sum(skipna=True) if "Actual HC used" in latest.columns else np.nan
target_uplh = (tot_target / tot_tahl) if tot_tahl else np.nan
actual_uplh = (tot_actual / tot_chl)  if tot_chl else np.nan
@st.cache_data(show_spinner=False, ttl=15 * 60)
def build_person_station_outputs_over_time(df, team_name, person):
    import json
    rows = []
    col = "Output by Cell/Station - by person"
    if col not in df.columns:
        return pd.DataFrame()
    sub = df.loc[df["team"] == team_name, ["period_date", col]].dropna(subset=[col]).copy()
    for _, r in sub.iterrows():
        pdte = pd.to_datetime(r["period_date"]).normalize()
        payload = r[col]
        try:
            data = json.loads(payload) if isinstance(payload, str) else payload
        except Exception:
            continue
        if not isinstance(data, dict):
            continue
        for station, person_map in data.items():
            if not isinstance(person_map, dict):
                continue
            if person in person_map and isinstance(person_map[person], dict):
                outv = pd.to_numeric(person_map[person].get("output"), errors="coerce")
                tgtv = pd.to_numeric(person_map[person].get("target"), errors="coerce")
                rows.append({
                    "period_date": pdte,
                    "person": person,
                    "cell_station": station,
                    "Actual": outv,
                    "Target": tgtv,
                })
    return pd.DataFrame(rows)
@st.cache_data(show_spinner=False, ttl=15 * 60)
def build_person_station_uplh_over_time(df: pd.DataFrame, team_name: str, person_name: str) -> pd.DataFrame:
    def _first_col(cands):
        for c in cands:
            if c in df.columns:
                return c
        return None
    col_uplh_pers   = _first_col(["UPLH by Cell/Station - by person"])  # nested {station:{person:{actual,target}}}
    col_out_pers    = _first_col(["Output by Cell/Station - by person", "Outputs by Cell/Station - by person"])  # {station:{person:{output,target}}}
    col_hrs_pers    = _first_col(["Hours by Cell/Station - by person"])  # {station:{person:{actual,target}}}
    d = df.loc[df["team"] == team_name].copy()
    if d.empty:
        return pd.DataFrame()
    d["period_date"] = pd.to_datetime(d["period_date"], errors="coerce").dt.normalize()
    rows = []
    for _, r in d.iterrows():
        wk = r.get("period_date", pd.NaT)
        if pd.isna(wk):
            continue
        def _parse_json_cell(value):
            if pd.isna(value):
                return {}
            if isinstance(value, (dict, list)):
                return value
            try:
                return json.loads(value)
            except Exception:
                return {}
        uplh_blob  = _parse_json_cell(r.get(col_uplh_pers))  if col_uplh_pers else {}
        out_blob   = _parse_json_cell(r.get(col_out_pers))   if col_out_pers  else {}
        hrs_blob   = _parse_json_cell(r.get(col_hrs_pers))   if col_hrs_pers  else {}
        stations = set()
        for blob in (uplh_blob, out_blob, hrs_blob):
            if isinstance(blob, dict):
                stations.update(blob.keys())
        for stn in sorted(stations):
            stn_key = str(stn)
            actual_uplh = target_uplh = np.nan
            actual_out = target_out = np.nan
            actual_hrs = target_hrs = np.nan
            if isinstance(uplh_blob.get(stn), dict):
                per_map = uplh_blob[stn]
                if person_name in per_map:
                    rec = per_map.get(person_name, {})
                else:
                    rec = next((per_map[k] for k in per_map if str(k).strip().lower() == str(person_name).strip().lower()), {})
                if isinstance(rec, dict):
                    actual_uplh = pd.to_numeric(rec.get("actual"), errors="coerce")
                    target_uplh = pd.to_numeric(rec.get("target"), errors="coerce")
            if isinstance(out_blob.get(stn), dict):
                per_map = out_blob[stn]
                if person_name in per_map:
                    rec = per_map.get(person_name, {})
                else:
                    rec = next((per_map[k] for k in per_map if str(k).strip().lower() == str(person_name).strip().lower()), {})
                if isinstance(rec, dict):
                    actual_out = pd.to_numeric(rec.get("output"), errors="coerce")
                    target_out = pd.to_numeric(rec.get("target"), errors="coerce")
            if isinstance(hrs_blob.get(stn), dict):
                per_map = hrs_blob[stn]
                if person_name in per_map:
                    rec = per_map.get(person_name, {})
                else:
                    rec = next((per_map[k] for k in per_map if str(k).strip().lower() == str(person_name).strip().lower()), {})
                if isinstance(rec, dict):
                    actual_hrs = pd.to_numeric(rec.get("actual"), errors="coerce")
                    target_hrs = pd.to_numeric(rec.get("target"), errors="coerce")
            if pd.isna(actual_uplh):
                if pd.notna(actual_out) and pd.notna(actual_hrs) and actual_hrs not in (0, 0.0):
                    actual_uplh = actual_out / actual_hrs
            if pd.isna(target_uplh):
                if pd.notna(target_out) and pd.notna(target_hrs) and target_hrs not in (0, 0.0):
                    target_uplh = target_out / target_hrs
            if (pd.isna(actual_out) or actual_out == 0) and (pd.isna(actual_hrs) or actual_hrs == 0) and pd.isna(actual_uplh):
                continue
            delta = np.nan
            if pd.notna(actual_uplh) and pd.notna(target_uplh):
                delta = actual_uplh - target_uplh
            rows.append({
                "period_date": wk,
                "team": team_name,
                "person": person_name,
                "cell_station": stn_key,
                "Actual": actual_out,
                "Target": target_out,
                "Actual Hours": actual_hrs,
                "Target Hours": target_hrs,
                "Actual UPLH": actual_uplh,
                "Target UPLH": target_uplh,
                "Delta": delta,
            })
    if not rows:
        return pd.DataFrame()
    out = pd.DataFrame(rows)
    num_cols = ["Actual", "Target", "Actual Hours", "Target Hours", "Actual UPLH", "Target UPLH", "Delta"]
    for c in num_cols:
        out[c] = pd.to_numeric(out[c], errors="coerce")
    out = out.sort_values(["cell_station", "period_date"]).reset_index(drop=True)
    return out
def _normalize_percent_value(v: float | int | np.floating | None) -> tuple[float, str]:
    if pd.isna(v):
        return np.nan, "{:.0%}"
    try:
        v = float(v)
    except Exception:
        return np.nan, "{:.0%}"
    if v <= 1.0:
        return v, "{:.0%}"
    return v / 100.0, "{:.0%}"
timeliness_avg_raw = latest["Open Complaint Timeliness"].dropna().mean() if "Open Complaint Timeliness" in latest.columns else np.nan
timeliness_avg, timeliness_fmt = _normalize_percent_value(timeliness_avg_raw)
nw_all = load_non_wip()
if not nw_all.empty:
    if "total_non_wip_hours" in nw_all.columns:
        nw_all = nw_all[["team", "period_date", "total_non_wip_hours"]].copy()
        nw_all["total_non_wip_hours"] = pd.to_numeric(nw_all["total_non_wip_hours"], errors="coerce")
    else:
        nw_all = nw_all.assign(total_non_wip_hours=np.nan)[["team", "period_date", "total_non_wip_hours"]]
else:
    nw_all = pd.DataFrame(columns=["team", "period_date", "total_non_wip_hours"])
latest_nw = latest.merge(nw_all, on=["team", "period_date"], how="left")
tot_nonwip = latest_nw["total_non_wip_hours"].sum(skipna=True) if "total_non_wip_hours" in latest_nw.columns else 0.0
tot_closures = latest["Closures"].sum(skipna=True) if "Closures" in latest.columns else np.nan
prod_den = (tot_chl or 0.0) + (tot_nonwip or 0.0)
productivity = (float(tot_closures) / prod_den) if (pd.notna(tot_closures) and prod_den) else np.nan
efficiency = (float(tot_closures) / float(tot_chl)) if (pd.notna(tot_closures) and tot_chl) else np.nan
kpi_cols = st.columns(2)
def kpi(col, label, value, fmt="{:,.2f}", help: str | None = None):
    if pd.isna(value):
        col.metric(label, "—", help=help)
    else:
        try:
            col.metric(label, fmt.format(value), help=help)
        except Exception:
            col.metric(label, str(value), help=help)
def kpi_vs_target(col, label, actual, target, fmt_val="{:,.2f}", help: str | None = None):
    if pd.isna(actual) or pd.isna(target) or not target:
        col.metric(label, "—", help=help)
        return
    try:
        value_str = fmt_val.format(actual)
    except Exception:
        value_str = str(actual)
    diff = (float(actual) - float(target)) / float(target)
    delta_str = f"{diff:+.0%} vs target"
    col.metric(label, value_str, delta=delta_str, delta_color="normal", help=help)
left, mid, right = st.columns(3)
base = alt.Chart(f).transform_calculate(
    week="toDate(datum.period_date)"
).encode(
    x=alt.X("period_date:T", title="Week")
)
teams_in_view = sorted([t for t in f["team"].dropna().unique()])
multi_team = len(teams_in_view) > 1
team_sel = alt.selection_point(fields=["team"], bind="legend")
mid2, right2 = st.columns(2) 
with mid2:
    st.subheader("Actual HC used Trend")
    if "Actual HC used" in f.columns and f["Actual HC used"].notna().any():
        ahu = f[["team", "period_date", "Actual HC used"]].dropna()
        base_ahu = alt.Chart(ahu).encode(
            x=alt.X("period_date:T", title="Week"),
            y=alt.Y("Actual HC used:Q", title="Actual HC used"),
            color=alt.Color("team:N", title="Team") if len(teams_in_view) > 1 else alt.value("indianred"),
            tooltip=["team:N", "period_date:T", alt.Tooltip("Actual HC used:Q", format=",.2f")]
        )
        st.altair_chart(
            base_ahu.mark_line(point=True).properties(height=280),
            width="stretch"
        )
        if len(teams_in_view) == 1:
            team_name = teams_in_view[0]
            if 'ppl_hours' not in locals():
                ppl_hours = explode_person_hours(f)
            team_people = ppl_hours.loc[ppl_hours["team"] == team_name].copy()
            if team_people.empty:
                st.info(f"No per-person data available for {team_name}.")
            else:
                all_weeks = sorted(
                    pd.to_datetime(team_people["period_date"].dropna().unique()),
                    reverse=True
                )
                picked_week = st.selectbox(
                    f"Week:",
                    options=all_weeks,
                    index=0,
                    format_func=lambda d: pd.to_datetime(d).date().isoformat(),
                    key="ahu_week_select_anyteam",
                )
                picked_week = pd.to_datetime(picked_week).normalize()
                wk_people = team_people.loc[team_people["period_date"] == picked_week].copy()
                if wk_people.empty:
                    st.info("No per-person data for the selected week.")
                else:
                    wk_people["Actual"] = pd.to_numeric(wk_people["Actual Hours"], errors="coerce")
                    wk_people = wk_people.loc[wk_people["Actual"].fillna(0) > 0].copy()
                    if wk_people.empty:
                        st.info("Nobody to show after filtering zero-hour entries.")
                    else:
                        wk_people["Avg Daily Hours"] = (wk_people["Actual"] / 5.0)
                        wk_people["OverUnder"] = np.where(
                            wk_people["Avg Daily Hours"] >= 6, "≥ 6 (Over)", "< 6 (Under)"
                        )
                        wk_people["Delta"] = wk_people["Avg Daily Hours"] - 6
                        wk_people["DeltaLabel"] = wk_people["Delta"].map(lambda x: f"{x:+.2f}")
                        vmax = float(pd.to_numeric(wk_people["Avg Daily Hours"], errors="coerce").max())
                        pad  = max(0.3, (max(vmax, 6) * 0.12))  # a little headroom
                        y_lo = 0.0
                        y_hi = max(vmax, 6) + pad
                        y_scale = alt.Scale(domain=[y_lo, y_hi], nice=False, clamp=False)
                        order_people = (
                            wk_people.sort_values("Avg Daily Hours", ascending=False)["person"].tolist()
                        )
                        color_enc = alt.Color(
                            "OverUnder:N",
                            title="vs 6",
                            scale=alt.Scale(
                                domain=["≥ 6 (Over)", "< 6 (Under)"],
                                range=["#22c55e", "#ef4444"]  # green / red
                            )
                        )
                        bars = (
                            alt.Chart(wk_people)
                            .mark_bar()
                            .encode(
                                x=alt.X("person:N", title="Person", sort=order_people),
                                y=alt.Y("Avg Daily Hours:Q", title="Avg Daily Hours (Actual/5)", scale=y_scale),
                                color=color_enc,
                                tooltip=[
                                    "period_date:T",
                                    "person:N",
                                    alt.Tooltip("Actual:Q", title="Actual Hours (week)", format=",.2f"),
                                    alt.Tooltip("Avg Daily Hours:Q", title="Avg Daily Hours", format=",.2f"),
                                    alt.Tooltip("Delta:Q", title="Over/Under vs 6", format="+.2f"),
                                ],
                            )
                            .properties(height=280)
                        )
                        label_pad = max(0.08, (y_hi - y_lo) * 0.03)
                        labels = (
                            alt.Chart(wk_people.assign(LabelY=lambda d: d["Avg Daily Hours"] + label_pad))
                            .mark_text(dy=-4)
                            .encode(
                                x="person:N",
                                y=alt.Y("LabelY:Q", scale=y_scale),
                                text="DeltaLabel:N",
                                color=alt.Color(
                                    "OverUnder:N",
                                    legend=None,
                                    scale=alt.Scale(
                                        domain=["≥ 6 (Over)", "< 6 (Under)"],
                                        range=["#22c55e", "#ef4444"]
                                    ),
                                ),
                            )
                        )
                        ref = alt.Chart(pd.DataFrame({"y": [6]})).mark_rule(strokeDash=[4, 3]).encode(y=alt.Y("y:Q", scale=y_scale))
                        st.altair_chart(bars + labels + ref, width="stretch")
        else:
            st.caption("Select exactly one team to drill into per-person daily hours.")
    else:
        st.info("No 'Actual HC used' data available in the selected range.")
with right2:
    st.subheader("Hours Trend")
    _nw = load_non_wip()
    teams_cfg = load_team_config()
    irl_lookup = {t: irl_people_for_team(t, teams_cfg) for t in teams_in_view}
    mix_rows = []
    nw_sub = _nw[_nw["team"].isin(teams_in_view)].copy()
    if not nw_sub.empty:
        for _, nw_row in nw_sub.iterrows():
            team = str(nw_row.get("team", "")).strip()
            wk = pd.to_datetime(nw_row.get("period_date"), errors="coerce")
            if not team or pd.isna(wk):
                continue
            wk = wk.normalize()
            wk_people = build_person_weekly_accounting(
                team=team,
                week=wk,
                nw_row=nw_row,
                metrics_frame=f,
                nw_frame=_nw,
                week_hours=40.0,
                irl_people=irl_lookup.get(team, set()),
            )
            if wk_people.empty:
                continue
            wk_people = wk_people.copy()
            wk_people["WIP"] = pd.to_numeric(wk_people["Completed Hours"], errors="coerce").fillna(0.0)
            wk_people["Other Team WIP"] = pd.to_numeric(wk_people["Other Team WIP"], errors="coerce").fillna(0.0)
            wk_people["Non-WIP"] = pd.to_numeric(wk_people["Accounted Non-WIP"], errors="coerce").fillna(0.0)
            wk_people["OOO"] = pd.to_numeric(wk_people["OOO Hours"], errors="coerce").fillna(0.0)
            wk_people["Unaccounted"] = pd.to_numeric(wk_people["Unaccounted"], errors="coerce").fillna(0.0)
            wk_people["Denom"] = (
                wk_people["WIP"]
                + wk_people["Other Team WIP"]
                + wk_people["Non-WIP"]
                + wk_people["OOO"]
                + wk_people["Unaccounted"]
            )
            long_df = wk_people.melt(
                id_vars=["team", "period_date", "person", "Denom"],
                value_vars=[
                    "WIP",
                    "Other Team WIP",
                    "Non-WIP",
                    "OOO",
                    "Unaccounted",
                ],
                var_name="Category",
                value_name="Hours",
            )
            long_df["Pct"] = np.where(
                long_df["Denom"] > 0,
                long_df["Hours"] / long_df["Denom"],
                np.nan,
            )
            mix_rows.append(
                long_df[["team", "period_date", "person", "Category", "Hours", "Pct"]]
            )
    person_mix = (
        pd.concat(mix_rows, ignore_index=True)
        if mix_rows
        else pd.DataFrame(columns=["team", "period_date", "person", "Category", "Hours", "Pct"])
    )
    person_mix = person_mix.dropna(subset=["period_date", "person", "Pct"]).copy()
    if person_mix.empty:
        st.info("No WIP vs Other Team WIP vs Non-WIP vs OOO vs Unaccounted data available.")
    else:
        mix_weeks = sorted(person_mix["period_date"].dropna().unique(), reverse=True)
        picked_mix_week = st.selectbox(
            "Week",
            options=mix_weeks,
            index=0,
            format_func=lambda d: pd.to_datetime(d).date().isoformat(),
            key="time_mix_week_right2",
        )
        week_mix = person_mix[
            person_mix["period_date"] == pd.to_datetime(picked_mix_week).normalize()
        ].copy()
        chosen_mix_teams = []
        if multi_team:
            teams_present = sorted(week_mix["team"].dropna().unique().tolist())
            chosen_mix_teams = st.multiselect(
                "Team(s)",
                options=teams_present,
                default=teams_present,
                key="time_mix_teams_right2",
            )
            if chosen_mix_teams:
                week_mix = week_mix[week_mix["team"].isin(chosen_mix_teams)].copy()
        if week_mix.empty:
            st.info("No person mix data for that selection.")
        else:
            category_domain = [
                "WIP",
                "Other Team WIP",
                "Non-WIP",
                "OOO",
                "Unaccounted",
            ]
            category_colors = [
                "#2563eb",  # WIP
                "#8b5cf6",  # Other Team WIP
                "#22c55e",  # Non-WIP
                "#f59e0b",  # OOO
                "#9ca3af",  # Unaccounted
            ]
            category_order_map = {
                "WIP": 0,
                "Other Team WIP": 1,
                "Non-WIP": 2,
                "OOO": 3,
                "Unaccounted": 4,
            }
            week_mix["Category"] = week_mix["Category"].astype(str).str.strip()
            week_mix = week_mix[week_mix["Category"].isin(category_domain)].copy()
            week_mix["CategoryOrder"] = week_mix["Category"].map(category_order_map)
            week_mix["Pct"] = pd.to_numeric(week_mix["Pct"], errors="coerce").fillna(0.0)
            week_mix["Hours"] = pd.to_numeric(week_mix["Hours"], errors="coerce").fillna(0.0)
            top_controls_left, top_controls_right = st.columns([1, 1])
            with top_controls_left:
                factor_out_ooo_top = st.toggle(
                    "Factor out OOO (top chart)",
                    value=True,
                    key="time_mix_factor_out_ooo_top_right2",
                )
            top_mix = week_mix.copy()
            if factor_out_ooo_top and not top_mix.empty:
                weekly_person_totals = (
                    top_mix.groupby(["period_date", "person"], as_index=False)["Hours"]
                    .sum()
                    .rename(columns={"Hours": "TotalHours"})
                )
                weekly_ooo = (
                    top_mix[top_mix["Category"] == "OOO"]
                    .groupby(["period_date", "person"], as_index=False)["Hours"]
                    .sum()
                    .rename(columns={"Hours": "OOOHours"})
                )
                weekly_base = weekly_person_totals.merge(
                    weekly_ooo,
                    on=["period_date", "person"],
                    how="left",
                )
                weekly_base["OOOHours"] = weekly_base["OOOHours"].fillna(0.0)
                weekly_base["AdjDenom"] = (
                    weekly_base["TotalHours"] - weekly_base["OOOHours"]
                ).clip(lower=0.0)
                top_mix = top_mix[top_mix["Category"] != "OOO"].copy()
                top_mix = top_mix.merge(
                    weekly_base[["period_date", "person", "AdjDenom"]],
                    on=["period_date", "person"],
                    how="left",
                )
                top_mix["Pct"] = np.where(
                    top_mix["AdjDenom"] > 0,
                    top_mix["Hours"] / top_mix["AdjDenom"],
                    np.nan,
                )
                top_mix = top_mix.dropna(subset=["Pct"]).copy()
            exclude_people = {
                "TM10", "TM11", "TM12", "TM13", "TM14", "TM15",
                "TM16", "TM5", "TM6", "TM7", "TM8", "TM9","TM3", "TM4", "#REF!"
            }
            top_mix["person"] = top_mix["person"].astype(str).str.strip()
            top_mix = top_mix[~top_mix["person"].isin(exclude_people)].copy()
            top_categories = [c for c in category_domain if (not factor_out_ooo_top or c != "OOO")]
            top_colors = [category_colors[category_domain.index(c)] for c in top_categories]
            person_order = (
                top_mix.groupby("person", as_index=False)["Hours"]
                .sum()
                .sort_values("Hours", ascending=False)["person"]
                .tolist()
            )
            label_src = top_mix.sort_values(["person", "CategoryOrder"]).copy()
            label_src["cum_pct"] = label_src.groupby("person")["Pct"].cumsum()
            label_src["y_mid"] = label_src["cum_pct"] - (label_src["Pct"] / 2.0)
            label_src = label_src[label_src["Pct"] >= 0.05].copy()
            bars = alt.Chart(top_mix).mark_bar().encode(
                x=alt.X(
                    "person:N",
                    title="Person",
                    sort=person_order,
                    axis=alt.Axis(
                        labelAngle=-90,
                        labelLimit=180,
                        labelOverlap=False, 
                    ),
                ),
                y=alt.Y(
                    "Pct:Q",
                    title="% of Time" if not factor_out_ooo_top else "% of Non-OOO Time",
                    stack="normalize",
                    axis=alt.Axis(format=".0%"),
                    scale=alt.Scale(domain=[0, 1]),
                ),
                color=alt.Color(
                    "Category:N",
                    title="Legend",
                    scale=alt.Scale(
                        domain=top_categories,
                        range=top_colors,
                    ),
                    sort=top_categories,
                    legend=alt.Legend(
                        orient="top",
                        direction="horizontal",
                        title=None,
                        labelLimit=200,
                    ),
                ),
                order=alt.Order("CategoryOrder:Q", sort="ascending"),
                tooltip=[
                    alt.Tooltip("team:N", title="Team"),
                    alt.Tooltip("person:N", title="Person"),
                    alt.Tooltip("Category:N", title="Category"),
                    alt.Tooltip("Hours:Q", title="Hours", format=",.2f"),
                    alt.Tooltip("Pct:Q", title="% of Time", format=".1%"),
                    alt.Tooltip("period_date:T", title="Week"),
                ],
            )
            labels = alt.Chart(label_src).mark_text(
                color="white",
                fontSize=11,
                fontWeight="bold",
                align="center",
                baseline="middle",
            ).encode(
                x=alt.X("person:N", sort=person_order),
                y=alt.Y(
                    "y_mid:Q",
                    scale=alt.Scale(domain=[0, 1]),
                    axis=None,
                ),
                detail="Category:N",
                text=alt.Text("Pct:Q", format=".0%"),
            )
            person_totals = (
                week_mix.groupby("person", as_index=False)["Hours"]
                .sum()
                .rename(columns={"Hours": "TotalHours"})
            )
            if "wk_people_kpi" in dir() and not wk_people_kpi.empty and "person" in wk_people_kpi.columns and "Expected Hours" in wk_people_kpi.columns:
                expected_hrs = wk_people_kpi[["person", "Expected Hours"]].copy()
                expected_hrs["person"] = expected_hrs["person"].astype(str).str.strip()
            else:
                expected_hrs = pd.DataFrame({"person": person_totals["person"], "Expected Hours": 40.0})
            person_totals["person"] = person_totals["person"].astype(str).str.strip()
            person_totals = person_totals.merge(expected_hrs, on="person", how="left")
            person_totals["Expected Hours"] = person_totals["Expected Hours"].fillna(40.0)
            overflow_df = person_totals[person_totals["TotalHours"] > person_totals["Expected Hours"]].copy()
            overflow_df["y_pos"] = 1.02  # just above the top of the 100% bar
            overflow_df["label"] = "⚠"
            overflow_layer = (
                alt.Chart(overflow_df)
                .mark_text(
                    fontSize=14,
                    fontWeight="bold",
                    color="#ef4444",
                    baseline="bottom",
                )
                .encode(
                    x=alt.X("person:N", sort=person_order),
                    y=alt.Y("y_pos:Q", scale=alt.Scale(domain=[0, 1.12]), axis=None),
                    text=alt.Text("label:N"),
                    tooltip=[
                        alt.Tooltip("person:N", title="Person"),
                        alt.Tooltip("TotalHours:Q", title="Total Hours", format=",.1f"),
                        alt.Tooltip("Expected Hours:Q", title="Expected Hours", format=",.1f"),
                    ],
                )
            )
            top_chart = (bars + labels + overflow_layer).properties(height=420)
            st.altair_chart(top_chart, width="stretch")
            st.markdown("##### Drill-down over time")
            people_for_drill = sorted(top_mix["person"].dropna().unique().tolist())
            picked_person_mix = st.selectbox(
                "Person",
                options=people_for_drill,
                key="time_mix_person_right2",
            )
            drill_controls_left, drill_controls_right = st.columns(2)
            with drill_controls_left:
                drill_window = st.segmented_control(
                    "Weeks",
                    options=[8, 12, 16],
                    default=16,
                    key="time_mix_window_right2",
                )
            with drill_controls_right:
                factor_out_ooo = st.toggle(
                    "Factor out OOO",
                    value=False,
                    key="time_mix_factor_out_ooo_right2",
                )
            drill_df = person_mix[person_mix["person"] == picked_person_mix].copy()
            if multi_team and chosen_mix_teams:
                drill_df = drill_df[drill_df["team"].isin(chosen_mix_teams)].copy()
            today = pd.Timestamp.today().normalize()
            cutoff = today - pd.Timedelta(weeks=drill_window)
            drill_df = drill_df[
                (drill_df["period_date"] >= cutoff) &
                (drill_df["period_date"] <= today)
            ].copy()
            if drill_df.empty:
                st.info("No over-time data for that person.")
            else:
                drill_df["Category"] = drill_df["Category"].astype(str).str.strip()
                drill_df = drill_df[drill_df["Category"].isin(category_domain)].copy()
                drill_df["CategoryOrder"] = drill_df["Category"].map(category_order_map)
                drill_df["Pct"] = pd.to_numeric(drill_df["Pct"], errors="coerce").fillna(0.0)
                drill_df["Hours"] = pd.to_numeric(drill_df["Hours"], errors="coerce").fillna(0.0)
                drill_df["period_date"] = pd.to_datetime(drill_df["period_date"], errors="coerce")
                latest_weeks = (
                    pd.Series(drill_df["period_date"].dropna().sort_values().unique()).tolist()
                )[-int(drill_window):]
                drill_df = drill_df[drill_df["period_date"].isin(latest_weeks)].copy()
                if factor_out_ooo and not drill_df.empty:
                    base_df = drill_df.copy()
                    weekly_person_totals = (
                        base_df.groupby(["period_date", "person"], as_index=False)["Hours"]
                        .sum()
                        .rename(columns={"Hours": "TotalHours"})
                    )
                    weekly_ooo = (
                        base_df[base_df["Category"] == "OOO"]
                        .groupby(["period_date", "person"], as_index=False)["Hours"]
                        .sum()
                        .rename(columns={"Hours": "OOOHours"})
                    )
                    weekly_base = weekly_person_totals.merge(
                        weekly_ooo,
                        on=["period_date", "person"],
                        how="left",
                    )
                    weekly_base["OOOHours"] = weekly_base["OOOHours"].fillna(0.0)
                    weekly_base["AdjDenom"] = (
                        weekly_base["TotalHours"] - weekly_base["OOOHours"]
                    ).clip(lower=0.0)
                    drill_df = drill_df[drill_df["Category"] != "OOO"].copy()
                    drill_df = drill_df.merge(
                        weekly_base[["period_date", "person", "AdjDenom"]],
                        on=["period_date", "person"],
                        how="left",
                    )
                    drill_df["Pct"] = np.where(
                        drill_df["AdjDenom"] > 0,
                        drill_df["Hours"] / drill_df["AdjDenom"],
                        np.nan,
                    )
                    drill_df = drill_df.dropna(subset=["Pct"]).copy()
                if drill_df.empty:
                    st.info("No over-time data for that person after applying filters.")
                else:
                    drill_categories = [c for c in category_domain if (not factor_out_ooo or c != "OOO")]
                    drill_colors = [
                        category_colors[category_domain.index(c)]
                        for c in drill_categories
                    ]
                    drill_df = drill_df.sort_values(["period_date", "CategoryOrder"]).copy()
                    drill_label_src = drill_df.copy()
                    drill_label_src["cum_pct"] = drill_label_src.groupby("period_date")["Pct"].cumsum()
                    drill_label_src["y_mid"] = drill_label_src["cum_pct"] - (drill_label_src["Pct"] / 2.0)
                    drill_label_src = drill_label_src[drill_label_src["Pct"] >= 0.05].copy()
                    week_count = max(len(latest_weeks), 1)
                    drill_width = max(380, min(900, week_count * 52))
                    drill_bars = (
                        alt.Chart(drill_df)
                        .mark_bar(size=28)
                        .encode(
                            x=alt.X(
                                "period_date:T",
                                title="Week",
                                axis=alt.Axis(format="%m/%d", labelAngle=0),
                            ),
                            y=alt.Y(
                                "Pct:Q",
                                title="% of Time" if not factor_out_ooo else "% of Non-OOO Time",
                                stack="normalize",
                                axis=alt.Axis(format=".0%"),
                                scale=alt.Scale(domain=[0, 1]),
                            ),
                            color=alt.Color(
                                "Category:N",
                                title="Legend",
                                scale=alt.Scale(
                                    domain=drill_categories,
                                    range=drill_colors,
                                ),
                                sort=drill_categories,
                                legend=alt.Legend(
                                    orient="top",
                                    direction="horizontal",
                                    title=None,
                                    labelLimit=200,
                                ),
                            ),
                            order=alt.Order("CategoryOrder:Q", sort="ascending"),
                            tooltip=[
                                alt.Tooltip("team:N", title="Team"),
                                alt.Tooltip("period_date:T", title="Week"),
                                alt.Tooltip("Category:N", title="Category"),
                                alt.Tooltip("Hours:Q", title="Hours", format=",.2f"),
                                alt.Tooltip("Pct:Q", title="% of Time", format=".1%"),
                            ],
                        )
                    )
                    drill_labels = alt.Chart(drill_label_src).mark_text(
                        color="white",
                        fontSize=10,
                        fontWeight="bold",
                        align="center",
                        baseline="middle",
                    ).encode(
                        x=alt.X("period_date:T"),
                        y=alt.Y(
                            "y_mid:Q",
                            scale=alt.Scale(domain=[0, 1]),
                            axis=None,
                        ),
                        detail="Category:N",
                        text=alt.Text("Pct:Q", format=".0%"),
                    )
                    drill_totals = (
                        drill_df.groupby("period_date", as_index=False)["Hours"]
                        .sum()
                        .rename(columns={"Hours": "TotalHours"})
                    )
                    person_key = str(picked_person_mix).strip().lower()
                    PERSON_WEEKLY_HOURS_DRILL = {"chelsey": 16.0, "mg": 36.0, "lindsey": 32.0}
                    if "wk_people_kpi" in dir() and not wk_people_kpi.empty and "person" in wk_people_kpi.columns and "Expected Hours" in wk_people_kpi.columns:
                        person_expected_match = wk_people_kpi.loc[
                            wk_people_kpi["person"].astype(str).str.strip().str.lower() == person_key,
                            "Expected Hours"
                        ]
                        drill_expected_hrs = float(person_expected_match.iloc[0]) if not person_expected_match.empty else PERSON_WEEKLY_HOURS_DRILL.get(person_key, 40.0)
                    else:
                        drill_expected_hrs = PERSON_WEEKLY_HOURS_DRILL.get(person_key, 40.0)
                    drill_totals["ExpectedHours"] = drill_expected_hrs
                    drill_overflow_df = drill_totals[drill_totals["TotalHours"] > drill_totals["ExpectedHours"]].copy()
                    drill_overflow_df["y_pos"] = 1.02
                    drill_overflow_df["label"] = "⚠"
                    drill_overflow_layer = (
                        alt.Chart(drill_overflow_df)
                        .mark_text(
                            fontSize=14,
                            fontWeight="bold",
                            color="#ef4444",
                            baseline="bottom",
                        )
                        .encode(
                            x=alt.X("period_date:T"),
                            y=alt.Y("y_pos:Q", scale=alt.Scale(domain=[0, 1.12]), axis=None),
                            text=alt.Text("label:N"),
                            tooltip=[
                                alt.Tooltip("period_date:T", title="Week"),
                                alt.Tooltip("TotalHours:Q", title="Total Hours", format=",.1f"),
                                alt.Tooltip("ExpectedHours:Q", title="Expected Hours", format=",.1f"),
                            ],
                        )
                    )
                    drill = (drill_bars + drill_labels + drill_overflow_layer).properties(
                        height=280,
                        width=drill_width,
                    )
                    st.altair_chart(drill, width="stretch")