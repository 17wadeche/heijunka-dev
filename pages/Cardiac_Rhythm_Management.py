# pages/Cardiac_Rhythm_Management.py
import hmac
import os, sys
from pathlib import Path
import pandas as pd
import numpy as np
import streamlit as st
import altair as alt
from utils.nonwip_kpi_lookup import enterprise_nonwip_kpi_lookup
import json
import unicodedata
import re
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from utils.activity_map import ACTIVITY_MAP
from utils.styles import apply_global_styles
apply_global_styles()
NON_WIP_DEFAULT_PATH = Path(r"C:\heijunka-dev\CRM_DATA\crm_non_wip_activities.csv")
def _safe_secret(name: str, default=None):
    import os
    try:
        return st.secrets.get(name, os.environ.get(name, default))
    except Exception:
        return os.environ.get(name, default)
NON_WIP_DATA_URL = _safe_secret("CRM_NON_WIP_DATA_URL")
DATA_URL = _safe_secret("CRM_HEIJUNKA_DATA_URL")
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
NAME_ALIASES = {
    "-": "-",
    "peter mchugh": "Peter McHugh",
    "peter mc hugh": "Peter McHugh",
}
def normalize_person_name(name: str) -> str:
    s = str(name or "")
    s = unicodedata.normalize("NFKC", s)
    s = "".join(ch for ch in s if unicodedata.category(ch)[0] != "C")
    s = s.replace("\u00a0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    key = s.casefold()
    if key in NAME_ALIASES:
        return NAME_ALIASES[key]
    def _title_token(t: str) -> str:
        if not t:
            return t
        for prefix in ("mc", "mac"):
            if t.lower().startswith(prefix) and len(t) > len(prefix):
                return prefix.capitalize() + t[len(prefix):].capitalize()
        return t.capitalize()
    tokens = s.split(" ")
    normalized = " ".join(_title_token(t) for t in tokens)
    if normalized.casefold() in NAME_ALIASES:
        return NAME_ALIASES[normalized.casefold()]
    return normalized
def person_key(name: str) -> str:
    s = normalize_person_name(name)
    s = re.sub(r"\s*\(\d+\)\s*$", "", str(s or "").strip())
    s = re.sub(r"\s+", " ", s).strip()
    return s.casefold()
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
    cache_tag: str = "CRM", 
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
                "period_date": safe_normalize_date(r["period_date"]),
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
DEFAULT_DATA_PATH = Path(r"C:\heijunka-dev\CRM_DATA\CRM_WIP.csv")
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
        name = normalize_person_name(d.get("name", ""))
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
               Name=lambda d: d["Name"].astype(str).str.strip(),
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
            rows.append({
                "team": r["team"],
                "period_date": safe_normalize_date(r["period_date"]),
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
    sub["period_date"] = sub["period_date"].map(safe_normalize_date)
    sub = sub.dropna(subset=["period_date"])
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
            parts = [p.strip() for p in re.split(r"[,;\n\r]+", s) if _is_good_name(p)]
            return parts
        if isinstance(x, dict):
            return [str(k).strip() for k in x.keys() if _is_good_name(str(k))]
        return []
    for _, r in sub.iterrows():
        people = _as_names(r["People in WIP"])
        for person in people:
            rows.append({
                "team": r["team"],
                "period_date": r["period_date"],
                "person": normalize_person_name(person),
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
            person_name = normalize_person_name(str(person).strip())
            team_name = str(r["team"]).strip().upper()
            keep_zero_capacity_person = (
                team_name in {"CDS", "NI"}
                and person_key(person_name) == "peter mchugh"
            )
            if (a == 0.0) and (t == 0.0) and not keep_zero_capacity_person:
                continue
            util = (a / t) if t not in (0, 0.0) else np.nan
            rows.append({
                "team": r["team"],
                "period_date": pd.to_datetime(r["period_date"], errors="coerce").normalize() if pd.notna(pd.to_datetime(r["period_date"], errors="coerce")) else pd.NaT,
                "person": person_name,
                "Actual Hours": a,
                "Available Hours": t,
                "Utilization": util
            })
    out = pd.DataFrame(rows)
    if not out.empty:
        out["period_date"] = pd.to_datetime(out["period_date"], errors="coerce").dt.normalize()
    return out
PERSON_WEEKLY_HOURS = {
    "colm larkin": 30.2,
    "megan mulligan": 31.0,
    "roland simpson": 37.5,
    "kara housmann": 37.5,
    "sarah korthauer": 37.5,
    "kyle mai": 37.5,
}
PERSON_TEAM_WEEKLY_HOURS = {
    ("peter mchugh", "NI"): 27.75,
}
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
        ["person", "Actual Hours", "Available Hours"]
    ].copy()
    if wip_people.empty:
        wip_people = pd.DataFrame(columns=["person", "Actual Hours", "Available Hours"])
    wip_people["person"] = wip_people["person"].astype(str).str.strip()
    wip_people["Completed Hours"] = pd.to_numeric(wip_people["Actual Hours"], errors="coerce").fillna(0.0)
    wip_people["Available Hours"] = pd.to_numeric(wip_people["Available Hours"], errors="coerce")
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
        .fillna(0.0)
    )
    out["person_key"] = out["person"].map(person_key)
    irl_people_norm = {person_key(x) for x in (irl_people or set())}
    base_expected = pd.Series(
        np.where(
            out["person_key"].isin(irl_people_norm),
            39.0,
            float(week_hours),
        ),
        index=out.index,
        dtype="float64",
    )
    team_key = str(team).strip().upper()
    team_override_expected = out["person_key"].map(
        lambda p: PERSON_TEAM_WEEKLY_HOURS.get((p, team_key), np.nan)
    ).astype("float64")
    person_override_expected = out["person_key"].map(PERSON_WEEKLY_HOURS).astype("float64")
    out["Expected Hours"] = (
        team_override_expected
        .combine_first(person_override_expected)
        .combine_first(base_expected)
    )
    if team_key in {"CDS", "NI"}:
        peter_available = pd.to_numeric(out["Available Hours"], errors="coerce")
        out.loc[
            out["person_key"].eq("peter mchugh") & peter_available.notna(),
            "Expected Hours"
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
                "period_date": safe_normalize_date(r["period_date"]),
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
            "Output by Cell/Station - by person",
            "Outputs by Cell/Station",   
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
                        "period_date": safe_normalize_date(r["period_date"]),
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
                        "period_date": safe_normalize_date(r["period_date"]),
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
                    "period_date": safe_normalize_date(r["period_date"]),
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
def safe_normalize_date(x):
    ts = pd.to_datetime(x, errors="coerce")
    return pd.NaT if pd.isna(ts) else ts.normalize()
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
    return m[cols].dropna(subset=["Actual Hours"])  # keep rows with hours; UPLH itself can be NaN if target missing
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
def percent_color(v: float | None, threshold: float, invert: bool = False) -> str:
    if v is None or pd.isna(v):
        return "#111827"
    good = (v >= threshold) if not invert else (v <= threshold)
    return "#22c55e" if good else "#ef4444"
st.markdown("<h1 style='text-align: center;'>CRM Heijunka Metrics Dashboard</h1>", unsafe_allow_html=True)
label = "Show WIP view" if st.session_state.get("nonwip_mode", False) else "Show Non-WIP view"
nonwip_mode = st.toggle(
    label,
    value=st.session_state.get("nonwip_mode", False),
    key="nonwip_mode",
    help="Switch between WIP and Non-WIP metrics"
)
if nonwip_mode:
    nw = load_non_wip()
    if nw.empty:
        st.info("No Non-WIP data found yet. Make sure non_wip_activities.csv exists.")
        st.stop()
    st.markdown("### Non-WIP Overview")
    teams_nw = sorted([t for t in nw["team"].dropna().unique()])
    c_team, c_week = st.columns(2)
    preferred_team = st.session_state.get("selected_team")
    if preferred_team not in teams_nw:
        preferred_team = _first_valid_team(preferred_team, teams_nw)
    if preferred_team is not None:
        st.session_state["nw_team"] = preferred_team
    def _sync_from_nonwip_team():
        team = st.session_state.get("nw_team")
        if team:
            st.session_state.selected_team = team
            st.session_state.teams_sel = [team]
    with c_team:
        team_nw = st.selectbox(
            "Team",
            options=teams_nw,
            key="nw_team",
            on_change=_sync_from_nonwip_team,
        )
    st.session_state.selected_team = team_nw
    st.session_state.teams_sel = [team_nw]
    weeks_nw = sorted(
        pd.to_datetime(nw.loc[nw["team"] == team_nw, "period_date"].dropna().unique()),
        reverse=True
    )
    if not weeks_nw:
        st.info("No weeks available for this team.")
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
    sel = nw[(nw["team"] == team_nw) & (nw["period_date"] == week_nw)]
    if sel.empty:
        st.info("No Non-WIP row for that team/week.")
        st.stop()
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
    wip_match = df[(df["team"] == team_nw) & (df["period_date"] == week_nw)]
    wip_hours_val = (
        float(pd.to_numeric(wip_match["Completed Hours"], errors="coerce").sum())
        if not wip_match.empty and "Completed Hours" in wip_match.columns
        else np.nan
    )
    TEAM_WEEKLY_HOURS = {
        "CPT": 37.75,
        "CDS": 37.75,
        "NI": 37.75,
        "DS": 37.5,
        "LIT & LETTERS": 37.5,
    }
    team_key = str(team_nw).strip().upper()
    team_week_hours = TEAM_WEEKLY_HOURS.get(team_key, 40.0)
    _ppl_hours_kpi = explode_person_hours(df)
    _ppl_in_wip_kpi = explode_people_in_wip(df)
    people_count_merged = merged_people_count_for_week(
        team=team_nw,
        week=week_nw,
        nw_frame=nw,
        person_hours=_ppl_hours_kpi,
        people_in_wip=_ppl_in_wip_kpi,
    )
    wk_people_kpi = build_person_weekly_accounting(
        team=team_nw,
        week=week_nw,
        nw_row=row,
        metrics_frame=df,
        nw_frame=nw,
        week_hours=team_week_hours,
        irl_people=set(),
    )
    if not wk_people_kpi.empty and "Expected Hours" in wk_people_kpi.columns:
        capacity_val = float(
            pd.to_numeric(wk_people_kpi["Expected Hours"], errors="coerce").fillna(0.0).sum()
        )
    elif people_count_merged > 0:
        capacity_val = float(people_count_merged * team_week_hours)
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
        nw_frame=nw,
        metrics_frame=df,
        ent_capacity_callback=globals().get("ent_capacity_hours_for_week"),
        ent_capacity_kwargs={
            "team": team_nw,
            "week": week_nw,
            "nw_frame": nw,
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
    TEAM_WEEKLY_HOURS = {
        "CPT": 37.75,
        "CDS": 37.75,
        "NI": 37.75,
        "DS": 37.5,
        "LIT & LETTERS": 37.5,
    }
    team_key = str(team_nw).strip().upper()
    team_week_hours = TEAM_WEEKLY_HOURS.get(team_key, 40.0)
    wk_people = build_person_weekly_accounting(
        team=team_nw,
        week=week_nw,
        nw_row=row,
        metrics_frame=df,
        nw_frame=nw,
        week_hours=team_week_hours,
        irl_people=set(),
    )
    if wk_people.empty:
        st.info("No per-person weekly breakdown for this selection.")
    else:
        wk_people = wk_people.rename(columns={
            "Other Team WIP": "Accounted_Other",
            "Accounted Non-WIP": "Accounted_NonOther",
        })
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
                    sort=order_people,
                    axis=alt.Axis(
                        labelAngle=-45,
                        labelOverlap=False,   # show all labels
                        labelLimit=0,         # don't truncate names
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
                height=300,
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
                                axis=alt.Axis(
                                    labelAngle=-45,
                                    labelLimit=200,
                                    labelOverlap=False,
                                ),
                            ),
                            y=alt.Y("Hours:Q", title="Total Non-WIP Hours"),
                            tooltip=[
                                alt.Tooltip("Activity:N", title="Activity"),
                                alt.Tooltip("Hours:Q", title="Hours", format=",.2f"),
                            ],
                        )
                        .properties(
                            height=420,
                            padding={"left": 8, "right": 12, "top": 16, "bottom": 80},
                        )
                    )
                    st.altair_chart(act_chart, width="stretch")
    st.markdown("#### Team Trends")
    team_hist = nw[nw["team"] == team_nw].dropna(subset=["period_date"]).sort_values("period_date")
    if not team_hist.empty:
        t1, t2 = st.columns(2)
        with t1:
            ch1 = (
                alt.Chart(team_hist)
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
            ch2 = (
                alt.Chart(team_hist)
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
teams = sorted([t for t in df["team"].dropna().unique()])
default_team = _first_valid_team(st.session_state.get("selected_team"), teams)
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
max_date = pd.to_datetime(df["period_date"].max()).date() if has_dates else None
if has_dates and min_date and max_date:
    if "start_date" not in st.session_state or st.session_state["start_date"] is None:
        st.session_state["start_date"] = min_date
    if "end_date" not in st.session_state or st.session_state["end_date"] is None:
        st.session_state["end_date"] = max_date
    st.session_state["start_date"] = min(max(st.session_state["start_date"], min_date), max_date)
    st.session_state["end_date"] = min(max(st.session_state["end_date"], min_date), max_date)
    if st.session_state["start_date"] > st.session_state["end_date"]:
        st.session_state["start_date"] = min_date
        st.session_state["end_date"] = max_date
    start = st.session_state["start_date"]
    end = st.session_state["end_date"]
else:
    start, end = None, None
col1, col2, col3 = st.columns([2, 2, 6], gap="large")
with col1:
    selected_teams = st.multiselect("Teams", teams, key="teams_sel")
if selected_teams:
    st.session_state.selected_team = selected_teams[0]
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
st.markdown("---")
left, right = st.columns(2)
base = alt.Chart(f).transform_calculate(
    week="toDate(datum.period_date)"
).encode(
    x=alt.X("period_date:T", title="Week")
)
teams_in_view = sorted([t for t in f["team"].dropna().unique()])
multi_team = len(teams_in_view) > 1
team_sel = alt.selection_point(fields=["team"], bind="legend")
with left:
    st.subheader("Actual WIP HC used Trend")
    if "Actual HC used" in f.columns and f["Actual HC used"].notna().any():
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
with right:
    st.subheader("Hours Trend")
    _nw = load_non_wip()
    TEAM_WEEKLY_HOURS = {
        "CPT": 37.75,
        "CDS": 37.75,
        "NI": 37.75,
        "DS": 37.5,
        "LIT & LETTERS": 37.5,
    }
    def canonical_person_label(name: str) -> str:
        s = normalize_person_name(name)
        s = re.sub(r"\s*\(\d+\)\s*$", "", str(s or "").strip())
        s = re.sub(r"\s+", " ", s).strip()
        return s
    mix_rows = []
    nw_sub = _nw[_nw["team"].isin(teams_in_view)].copy()
    if not nw_sub.empty:
        for _, nw_row in nw_sub.iterrows():
            team = str(nw_row.get("team", "")).strip()
            wk = pd.to_datetime(nw_row.get("period_date"), errors="coerce")
            if not team or pd.isna(wk):
                continue
            wk = wk.normalize()
            TEAM_WEEKLY_HOURS = {
                "CPT": 37.75,
                "CDS": 37.75,
                "NI": 37.75,
                "DS": 37.5,
                "LIT & LETTERS": 37.5,
            }
            team_week_hours = TEAM_WEEKLY_HOURS.get(team.upper(), 40.0)
            wk_people = build_person_weekly_accounting(
                team=team,
                week=wk,
                nw_row=nw_row,
                metrics_frame=f,
                nw_frame=_nw,
                week_hours=team_week_hours,
                irl_people=set(),   # do not use teams.json
            )
            if wk_people.empty:
                continue
            wk_people = wk_people.copy()
            wk_people["person"] = wk_people["person"].map(canonical_person_label)
            wk_people["Expected Hours"] = pd.to_numeric(
                wk_people["Expected Hours"],
                errors="coerce"
            ).fillna(0.0)
            wk_people["Completed Hours"] = pd.to_numeric(wk_people["Completed Hours"], errors="coerce").fillna(0.0)
            wk_people["Other Team WIP"] = pd.to_numeric(wk_people["Other Team WIP"], errors="coerce").fillna(0.0)
            wk_people["Accounted Non-WIP"] = pd.to_numeric(wk_people["Accounted Non-WIP"], errors="coerce").fillna(0.0)
            wk_people["OOO Hours"] = pd.to_numeric(wk_people["OOO Hours"], errors="coerce").fillna(0.0)
            wk_people["Unaccounted"] = (
                wk_people["Expected Hours"]
                - wk_people["Completed Hours"]
                - wk_people["Other Team WIP"]
                - wk_people["Accounted Non-WIP"]
                - wk_people["OOO Hours"]
            ).clip(lower=0.0)
            wk_people = wk_people.drop(columns=["person_key", "team_key"], errors="ignore")
            wk_people["WIP"] = wk_people["Completed Hours"]
            wk_people["Non-WIP"] = wk_people["Accounted Non-WIP"]
            wk_people["OOO"] = wk_people["OOO Hours"]
            wk_people = (
                wk_people.groupby(["team", "period_date", "person"], as_index=False)
                .agg({
                    "WIP": "sum",
                    "Other Team WIP": "sum",
                    "Non-WIP": "sum",
                    "OOO": "sum",
                    "Unaccounted": "sum",
                    "Expected Hours": "max",
                })
            )
            wk_people["Denom"] = wk_people["Expected Hours"]
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
            week_mix = week_mix.sort_values(
                ["team", "period_date", "person", "CategoryOrder"]
            ).copy()
            week_mix["TruePct"] = week_mix["Pct"]
            week_mix["CumPctBefore"] = (
                week_mix
                .groupby(["team", "period_date", "person"])["TruePct"]
                .cumsum()
                - week_mix["TruePct"]
            )
            week_mix["PlotPct"] = np.minimum(
                week_mix["TruePct"],
                (1.0 - week_mix["CumPctBefore"]).clip(lower=0.0),
            ).clip(lower=0.0)
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
                    "PlotPct:Q",
                    title="% of Time" if not factor_out_ooo_top else "% of Non-OOO Time",
                    stack="zero",
                    axis=alt.Axis(format=".0%"),
                    scale=alt.Scale(domain=[0, 1.08], nice=False, clamp=False),
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
                    alt.Tooltip("TruePct:Q", title="% of Time", format=".1%"),
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
            overflow_df = (
                week_mix.groupby(["team", "period_date", "person"], as_index=False)
                .agg(
                    TotalPct=("TruePct", "sum"),
                    TotalHours=("Hours", "sum"),
                )
            )
            overflow_df = overflow_df[overflow_df["TotalPct"] > 1.0].copy()
            overflow_df["y_pos"] = 1.025
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
                    y=alt.Y(
                        "y_pos:Q",
                        scale=alt.Scale(domain=[0, 1.08], nice=False, clamp=False),
                        axis=None,
                    ),
                    text="label:N",
                    tooltip=[
                        alt.Tooltip("person:N", title="Person"),
                        alt.Tooltip("TotalHours:Q", title="Total Hours Worked", format=",.2f"),
                        alt.Tooltip("TotalPct:Q", title="Total % of Time", format=".1%"),
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
                                stack="zero",
                                axis=alt.Axis(format=".0%"),
                                scale=alt.Scale(domain=[0, 1.12], nice=False, clamp=False),
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
                                alt.Tooltip("TruePct:Q", title="% of Time", format=".1%"),
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
                    picked_person_key = person_key(picked_person_mix)
                    PERSON_WEEKLY_HOURS = {
                        "colm larkin": 30.2,
                        "megan mulligan": 31.0,
                        "roland simpson": 37.5,
                        "kara housmann": 37.5,
                        "sarah korthauer": 37.5,
                        "kyle mai": 37.5,
                    }
                    PERSON_TEAM_WEEKLY_HOURS = {
                        ("peter mchugh", "CDS"): 10.0,
                        ("peter mchugh", "NI"): 27.75,
                    }
                    TEAM_WEEKLY_HOURS = {
                        "CPT": 37.75,
                        "CDS": 37.75,
                        "NI": 37.75,
                        "DS": 37.5,
                        "LIT & LETTERS": 37.5,
                    }
                    if multi_team and chosen_mix_teams:
                        teams_for_drill = {str(t).strip().upper() for t in chosen_mix_teams}
                    elif "team_name" in locals() and team_name:
                        teams_for_drill = {str(team_name).strip().upper()}
                    else:
                        teams_for_drill = set()
                    default_drill_expected = (
                        sum(TEAM_WEEKLY_HOURS.get(t, 40.0) for t in teams_for_drill)
                        if teams_for_drill
                        else 40.0
                    )
                    team_specific_expected = [
                        PERSON_TEAM_WEEKLY_HOURS[(picked_person_key, t)]
                        for t in teams_for_drill
                        if (picked_person_key, t) in PERSON_TEAM_WEEKLY_HOURS
                    ]
                    if team_specific_expected:
                        drill_expected_hrs = sum(team_specific_expected)
                    else:
                        drill_expected_hrs = PERSON_WEEKLY_HOURS.get(
                            picked_person_key,
                            default_drill_expected,
                        )
                    drill_totals["ExpectedHours"] = drill_expected_hrs
                    active_peter_teams = {"CDS", "NI"} & teams_for_drill
                    if picked_person_key == "peter mchugh" and active_peter_teams:
                        peter_avail = (
                            explode_person_hours(df)
                            .loc[
                                lambda d: (
                                    d["team"].astype(str).str.strip().str.upper().isin(active_peter_teams)
                                    & d["person"].map(person_key).eq("peter mchugh")
                                ),
                                ["period_date", "Available Hours"],
                            ]
                            .copy()
                        )
                        if not peter_avail.empty:
                            peter_avail["period_date"] = pd.to_datetime(
                                peter_avail["period_date"],
                                errors="coerce",
                            ).dt.normalize()
                            peter_avail["Available Hours"] = pd.to_numeric(
                                peter_avail["Available Hours"],
                                errors="coerce",
                            )
                            peter_avail = (
                                peter_avail
                                .dropna(subset=["period_date"])
                                .groupby("period_date", as_index=False)["Available Hours"]
                                .sum()
                            )
                            drill_totals["period_date"] = pd.to_datetime(
                                drill_totals["period_date"],
                                errors="coerce",
                            ).dt.normalize()
                            drill_totals = drill_totals.merge(
                                peter_avail.rename(columns={"Available Hours": "PeterAvailableHours"}),
                                on="period_date",
                                how="left",
                            )
                            peter_fallback_expected = sum(
                                PERSON_TEAM_WEEKLY_HOURS.get(("peter mchugh", t), TEAM_WEEKLY_HOURS.get(t, 40.0))
                                for t in active_peter_teams
                            )
                            drill_totals["ExpectedHours"] = np.where(
                                drill_totals["PeterAvailableHours"].notna(),
                                drill_totals["ExpectedHours"] - peter_fallback_expected + drill_totals["PeterAvailableHours"],
                                drill_totals["ExpectedHours"],
                            )
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