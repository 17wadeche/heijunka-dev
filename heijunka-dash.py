# heijunka-dash.py
import os
from pathlib import Path
import pandas as pd
import numpy as np
import streamlit as st
import altair as alt
import json
NON_WIP_DEFAULT_PATH = Path(r"C:\heijunka-dev\non_wip_activities.csv")
NON_WIP_DATA_URL = st.secrets.get("NON_WIP_DATA_URL", os.environ.get("NON_WIP_DATA_URL"))
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
@st.cache_data(show_spinner=False, ttl=15 * 60)
def load_non_wip(nw_path: str | None = None, nw_url: str | None = NON_WIP_DATA_URL) -> pd.DataFrame:
    if nw_url:
        try:
            df = pd.read_csv(nw_url, dtype=str, keep_default_na=False, encoding="utf-8-sig")
        except Exception:
            import io, requests
            r = requests.get(nw_url, timeout=20)
            r.raise_for_status()
            df = pd.read_csv(io.StringIO(r.content.decode("utf-8-sig", errors="replace")),
                             dtype=str, keep_default_na=False)
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
                "person": str(person).strip(),
                "Non-WIP Hours": v
            })
    out = pd.DataFrame(rows, columns=cols)
    if not out.empty:
        out["period_date"] = pd.to_datetime(out["period_date"], errors="coerce").dt.normalize()
    return out
DEFAULT_DATA_PATH = Path(r"C:\heijunka-dev\metrics_aggregate_dev.csv")
DATA_URL = st.secrets.get("HEIJUNKA_DATA_URL", os.environ.get("HEIJUNKA_DATA_URL"))
st.set_page_config(page_title="Heijunka Metrics", layout="wide")
hide_streamlit_style = """
    <style>
    [data-testid="stToolbar"] { display: none; }
    #MainMenu { visibility: hidden; }
    header { visibility: hidden; }
    footer { visibility: hidden; }
    </style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)
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
                "Target UPLH", "Actual UPLH", "HC in WIP", "Actual HC used", "Closures"]:
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
        name = str(d.get("name", "")).strip()
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
              day_values=("day_norm", lambda s: sorted(set([x for x in s.dropna().unique()])))
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
        grp.rename(columns={"activity": "Activity", "name": "Name", "hours": "Time"})
           [["Activity", "Day", "Name", "Time"]]
           .assign(
               Activity=lambda d: d["Activity"].astype(str).str.strip(),
               Name=lambda d: d["Name"].astype(str).str.strip(),
               Time=lambda d: d["Time"].fillna(0).map(_fmt_hours_minutes),   # <-- here
           )
           .sort_values(["Activity", "Name"])
           .reset_index(drop=True)                                           # <-- and here
    )
    return out
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
                "period_date": pd.to_datetime(r["period_date"], errors="coerce").normalize(),
                key_label: str(k).strip(),
                "Actual": outv,
                "Target": tgtv
            })
    out = pd.DataFrame(rows, columns=cols)
    if not out.empty:
        out["period_date"] = pd.to_datetime(out["period_date"], errors="coerce").dt.normalize()
    return out
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
                "person": person
            })
    out = pd.DataFrame(rows)
    if not out.empty:
        out["period_date"] = pd.to_datetime(out["period_date"], errors="coerce").dt.normalize()
    return out
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
                "person": str(person).strip(),
                "Actual Hours": a,
                "Available Hours": t,
                "Utilization": util
            })
    out = pd.DataFrame(rows)
    if not out.empty:
        out["period_date"] = pd.to_datetime(out["period_date"], errors="coerce").dt.normalize()
    return out
def _find_first_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    for c in candidates:
        if c in df.columns:
            return c
    return None
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
                        "person": str(person).strip(),
                        "Actual": a,
                        "Target": t,
                    })
    out = pd.DataFrame(rows, columns=cols)
    if not out.empty:
        out["period_date"] = pd.to_datetime(out["period_date"], errors="coerce").dt.normalize()
    return out
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
                        "person": str(person).strip(),
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
                    "person": str(person).strip(),
                    "Actual Hours": a,
                    "Available Hours": t,
                })
    out = pd.DataFrame(rows, columns=cols)
    if not out.empty:
        out["period_date"] = pd.to_datetime(out["period_date"], errors="coerce").dt.normalize()
    return out
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
def kpi_card(container, label: str, value, fmt: str | None = None, color: str | None = None, help: str | None = None):
    if pd.isna(value):
        val_html = "—"
    else:
        try:
            val_html = (fmt or "{}").format(value)
        except Exception:
            val_html = str(value)
    help_icon = f"""<span title="{help}" style="cursor:help;margin-left:6px;color:#9ca3af;">ⓘ</span>""" if help else ""
    value_color = color or "#111827"
    container.markdown(
        f"""
        <div style="padding:12px 16px;border-radius:10px;border:1px solid #eee;">
          <div style="font-size:12px;color:#6b7280;display:flex;align-items:center;gap:4px;">
            <span>{label}</span>{help_icon}
          </div>
          <div style="font-size:28px;font-weight:700;color:{value_color};">{val_html}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
def percent_color(v: float | None, threshold: float, invert: bool = False) -> str:
    if v is None or pd.isna(v):
        return "#111827"
    good = (v >= threshold) if not invert else (v <= threshold)
    return "#22c55e" if good else "#ef4444"
st.markdown("<h1 style='text-align: center;'>Heijunka Metrics Dashboard</h1>", unsafe_allow_html=True)
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
    with c_team:
        team_nw = st.selectbox("Team", options=teams_nw, index=0, key="nw_team")
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
    people_val = int(row["people_count"]) if pd.notna(row["people_count"]) else np.nan
    kpi_card(c1, "People Count", people_val, fmt="{:,}")
    hours_val = float(row["total_non_wip_hours"]) if pd.notna(row["total_non_wip_hours"]) else np.nan
    kpi_card(c2, "Total Non-WIP Hours", hours_val, fmt="{:,.1f}")
    ooo_val = float(row.get("OOO Hours", np.nan)) if "OOO Hours" in sel.columns else np.nan
    kpi_card(c3, "OOO Hours", ooo_val, fmt="{:,.1f}", help="8 hours per person per OOO day")
    nonwip_val = float(pct_non_wip) if pd.notna(pct_non_wip) else np.nan
    kpi_card(
        c4,
        "% Non-WIP",
        nonwip_val,
        fmt="{:.2f}%",
        color=percent_color(nonwip_val, threshold=25.0, invert=True),
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
            st.dataframe(act_tbl, use_container_width=True, hide_index=True)
    long_nw = explode_non_wip_by_person(nw)
    wk_people = long_nw[(long_nw["team"] == team_nw) & (long_nw["period_date"] == week_nw)].dropna(subset=["Non-WIP Hours"])
    if wk_people.empty:
        st.info("No per-person Non-WIP breakdown for this selection.")
    else:
        acct_other_map, acct_nonother_map = accounted_nonwip_by_person_from_row(row)
        wk_people = wk_people.assign(
            Accounted_Other=lambda d: d["person"].map(lambda p: float(acct_other_map.get(str(p).strip(), 0.0))),
            Accounted_NonOther=lambda d: d["person"].map(lambda p: float(acct_nonother_map.get(str(p).strip(), 0.0))),
        )
        wk_people["Accounted_Other"] = wk_people[["Accounted_Other", "Non-WIP Hours"]].min(axis=1)
        remaining = (wk_people["Non-WIP Hours"] - wk_people["Accounted_Other"]).clip(lower=0)
        wk_people["Accounted_NonOther"] = np.minimum(
            wk_people["Accounted_NonOther"].astype(float),
            remaining.astype(float)
        )
        wk_people["Unaccounted"] = (
            wk_people["Non-WIP Hours"] 
            - wk_people["Accounted_Other"] 
            - wk_people["Accounted_NonOther"]
        ).clip(lower=0)        
        stack = (
            wk_people.melt(
                id_vars=["person", "period_date"],
                value_vars=["Accounted_Other", "Accounted_NonOther", "Unaccounted"],
                var_name="Category",
                value_name="Hours"
            )
            .dropna(subset=["Hours"])
        )
        label_map = {
            "Accounted_Other": "Other Team WIP",
            "Accounted_NonOther": "Accounted",
            "Unaccounted": "Unaccounted",
        }
        stack["CategoryLabel"] = stack["Category"].map(label_map)
        order_people = wk_people.sort_values("Non-WIP Hours", ascending=False)["person"].tolist()
        vmax = float(pd.to_numeric(wk_people["Non-WIP Hours"], errors="coerce").max())
        headroom = max(1.0, vmax * 0.18) if pd.notna(vmax) else 1.0
        y_scale = alt.Scale(domain=[0, vmax + headroom], nice=False, clamp=False)
        totals = (
            wk_people[["person", "period_date", "Non-WIP Hours"]]
            .rename(columns={"Non-WIP Hours": "Total"})
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
                x=alt.X("person:N", title="Person", sort=order_people, axis=alt.Axis(labelAngle=-30, labelLimit=140)),
                y=alt.Y("Hours:Q", title="Non-WIP Hours (week)", stack="zero", scale=y_scale),
                color=alt.Color(
                    "CategoryLabel:N",
                    title="Legend",
                    scale=alt.Scale(
                        domain=["Other Team WIP", "Accounted", "Unaccounted"],
                        range=["#2563eb", "#22c55e", "#9ca3af"]
                    )
                ),
                tooltip=[
                    alt.Tooltip("person:N", title="Person"),
                    alt.Tooltip("CategoryLabel:N", title="Category"),
                    alt.Tooltip("Hours:Q", title="Hours", format=",.2f"),
                    alt.Tooltip("period_date:T", title="Date"),
                ],
            )
        )
        ref = (
            alt.Chart(pd.DataFrame({"y": [7.5]}))
            .mark_rule(strokeDash=[4, 3], color="#6b7280")
            .encode(y=alt.Y("y:Q", scale=y_scale))
        )
        chart = (bars + ref + outline) \
            .properties(
                height=300,
                title=f"{team_nw} • Per-person Non-WIP Hours (Accounted vs Unaccounted)",
                padding={"left": 8, "right": 12, "top": 36, "bottom": 64},
            ) \
            .configure_axis(labelOverlap=True) \
            .configure_view(stroke=None)
    st.altair_chart(chart, use_container_width=True)
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
            st.altair_chart(ch1, use_container_width=True)
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
            st.altair_chart(ch2, use_container_width=True)
    st.markdown("#### Weekly Non-WIP Rows")
    show_cols = ["team","period_date","people_count","total_non_wip_hours","% Non-WIP"]
    tbl = (
        team_hist[show_cols]
        .rename(columns={
            "team": "Team",
            "period_date": "Date",
            "people_count": "People Count",
            "total_non_wip_hours": "Non-WIP Hours",
            "% Non-WIP": "% Non-WIP",
        })
        .sort_values("Date", ascending=False)
    )
    if "Date" in tbl.columns:
        tbl["Date"] = pd.to_datetime(tbl["Date"], errors="coerce").dt.date
    st.dataframe(
        tbl.style.format({
            "People Count": "{:,.0f}",
            "Non-WIP Hours": "{:,.1f}",
            "% Non-WIP": "{:.2f}%",
        }),
        use_container_width=True,
        hide_index=True,
    )
    st.stop()
with st.expander("Glossary", expanded=False):
    st.markdown("""
- **Target UPLH** — Target Output ÷ Target Hours (i.e., **Total Available Hours**)
- **Actual UPLH** — Actual Output ÷ Actual Hours (i.e., **Completed Hours**)
- **Capacity Utilization** — Completed Hours ÷ Available Hours (Amount of the Capacity that has been used)
- **HC in WIP** — Number of **unique people** who logged any time in WIP during the week
- **Actual HC used** — Total actual hours worked ÷ **30**  
  <small>(assumes **6 hours in WIP per person per day × 5 days**)</small>
- **Closures** — Number of **PEs** closed during the week
- **Efficiency** — Closures ÷ Completed WIP Hours.
- **Productivity** — Closures ÷ (Completed WIP Hours + Non-WIP Hours)
- **Multi-Axis View tip** — If you select **only one** series, you can project the next **3 months**.
""", unsafe_allow_html=True)
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
default_teams = [teams[0]] if teams else []
if "teams_sel" not in st.session_state:
    saved = [t for t in teams if t in _get_qp_teams()]
    st.session_state.teams_sel = saved or default_teams
has_dates = df["period_date"].notna().any()
min_date = pd.to_datetime(df["period_date"].min()).date() if has_dates else None
max_date = pd.to_datetime(df["period_date"].max()).date() if has_dates else None
if has_dates and min_date and max_date:
    if "start_date" not in st.session_state:
        st.session_state["start_date"] = min_date
    if "end_date" not in st.session_state:
        st.session_state["end_date"] = max_date
    start = st.session_state["start_date"]
    end = st.session_state["end_date"]
    if start > end:
        st.error("Start date cannot be after end date!")
        start, end = min_date, max_date
        st.session_state["start_date"] = start
        st.session_state["end_date"] = end
else:
    start, end = None, None
col1, col2, col3 = st.columns([2, 2, 6], gap="large")
with col1:
    selected_teams = st.multiselect("Teams", teams, key="teams_sel")
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
kpi_cols = st.columns(4)
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
with kpi_cols[0]:
    st.subheader("Latest Week (Selected Teams)")
row1 = st.columns(7)
kpi(
    row1[0],
    "HC in WIP",
    tot_hc_wip,
    "{:,.0f}",
    help="Unique people with any time in WIP for the week",
)
kpi(
    row1[1],
    "Actual HC used",
    tot_hc_used,
    "{:,.2f}",
    help="Based on 6 hours per person in WIP per day",
)
kpi(
    row1[2],
    "Capacity Utilization",
    (tot_chl / tot_tahl if tot_tahl else np.nan),
    "{:.0%}",
    help="Completed vs Available hours",
)
kpi_vs_target(
    row1[3],
    "Open Complaint Timeliness",
    timeliness_avg,
    0.87,
    "{:.0%}",
)
row2 = st.columns(3)
kpi(
    row1[4],
    "Closures",
    tot_closures,
    "{:,.0f}",
    help="Events Closed",
)
kpi(
    row1[5],
    "Efficiency",
    efficiency,
    "{:.3f}",
    help="Closures ÷ Completed WIP Hours",
)
kpi(
    row1[6],
    "Productivity",
    productivity,
    "{:,.3f}",
    help="Closures ÷ All Available Hours",
)
st.markdown("---")
if has_dates and min_date and max_date:
    st.markdown("#### Date Range")
    date_col1, date_col2 = st.columns(2)
    with date_col1:
        st.date_input(
            "Start",
            value=st.session_state["start_date"],
            min_value=min_date,
            max_value=max_date,
            key="start_date",
        )
    with date_col2:
        st.date_input(
            "End",
            value=st.session_state["end_date"],
            min_value=min_date,
            max_value=max_date,
            key="end_date",
        )
left, mid, right = st.columns(3)
base = alt.Chart(f).transform_calculate(
    week="toDate(datum.period_date)"
).encode(
    x=alt.X("period_date:T", title="Week")
)
teams_in_view = sorted([t for t in f["team"].dropna().unique()])
multi_team = len(teams_in_view) > 1
team_sel = alt.selection_point(fields=["team"], bind="legend")
with left:
    st.subheader("WIP Hours Trend")
    have_hours = {"Total Available Hours", "Completed Hours"}.issubset(f.columns)
    teams_in_view = sorted([t for t in f["team"].dropna().unique()])
    single_team = (len(teams_in_view) == 1)
    if not have_hours:
        st.info("Hours columns not found (need 'Total Available Hours' and 'Completed Hours').")
    else:
        hrs_long = (
            f.melt(
                id_vars=["team", "period_date"],
                value_vars=["Total Available Hours", "Completed Hours"],
                var_name="Metric",
                value_name="Value"
            )
            .dropna(subset=["Value"])
            .assign(Metric=lambda d: d["Metric"].replace({
                "Total Available Hours": "Target Hours",
                "Completed Hours": "Actual Hours"
            }))
        )
        team_sel = alt.selection_point(fields=["team"], bind="legend")
        base_trend = alt.Chart(hrs_long).encode(
            x=alt.X("period_date:T", title="Week"),
            y=alt.Y("Value:Q", title="Hours"),
            color=alt.Color("Metric:N", title="Series"),
            tooltip=["team:N", "period_date:T", "Metric:N", alt.Tooltip("Value:Q", format=",.0f")],
        )
        line = base_trend.mark_line(point=False).encode(
            detail="team:N",
            opacity=alt.condition(team_sel, alt.value(1.0), alt.value(0.25))
            if len(teams_in_view) > 1 else alt.value(1.0)
        )
        pts = base_trend.mark_point().encode(
            shape=alt.Shape("team:N", title="Team") if len(teams_in_view) > 1 else alt.value("circle"),
            size=alt.value(45),
            opacity=alt.condition(team_sel, alt.value(1.0), alt.value(0.25))
            if len(teams_in_view) > 1 else alt.value(1.0)
        )
        chart_ph = st.empty()
        chart_ph.altair_chart(
            alt.layer(line, pts).properties(height=280).add_params(team_sel),
            use_container_width=True
        )
        team_name = teams_in_view[0] if len(teams_in_view) == 1 else None
        if team_name is None:
            if 'ppl_hours' in locals() and ppl_hours.empty:
                st.caption("Per-person drilldown not available (no 'Person Hours' found).")
            else:
                st.caption("Per-person & cell/station drilldowns are available when exactly one team is selected.")
        else:
            hours_by = st.selectbox(
                "Hours by:",
                options=["Person", "Cell/Station"],
                index=0,
                key="hours_by_select",
            )
            team_weeks = sorted(
                pd.to_datetime(f.loc[f["team"] == team_name, "period_date"].dropna().unique()),
                reverse=True
            )
            if not team_weeks:
                st.info("No weeks available for drilldown.")
            else:
                picked_week = st.selectbox(
                    f"Week:",
                    options=team_weeks,
                    index=0,
                    format_func=lambda d: pd.to_datetime(d).date().isoformat(),
                    key="hours_by_week_select",
                )
                picked_week = pd.to_datetime(picked_week).normalize()
                layers = [line, pts]
                if picked_week is not None:
                    rule_df = pd.DataFrame({"period_date": [picked_week]})
                    rule = alt.Chart(rule_df).mark_rule(strokeDash=[4, 3]).encode(x="period_date:T")
                    layers.append(rule)
                chart_ph.altair_chart(
                    alt.layer(*layers).properties(height=280).add_params(team_sel),
                    use_container_width=True
                )
                if hours_by == "Person":
                    if 'ppl_hours' not in locals() or ppl_hours.empty:
                        st.info(f"No per-person data available for {team_name}.")
                    else:
                        team_people = ppl_hours.loc[ppl_hours["team"] == team_name].copy()
                        wk_people = team_people.loc[team_people["period_date"] == picked_week].copy()
                        if wk_people.empty:
                            st.info("No per-person data for the selected week.")
                        else:
                            wk2 = (
                                wk_people.assign(
                                    Actual=lambda d: pd.to_numeric(d["Actual Hours"], errors="coerce"),
                                    Avail=lambda d: pd.to_numeric(d["Available Hours"], errors="coerce"),
                                )
                                .assign(Diff=lambda d: d["Actual"] - d["Avail"])
                                .assign(DiffRounded=lambda d: d["Diff"].round(1))
                            )
                            wk2 = wk2.loc[
                                ~(
                                    (wk2["Actual"].fillna(0) == 0) &
                                    (wk2["Avail"].fillna(0)  == 0)
                                )
                            ].assign(
                                DiffLabel=lambda d: d["DiffRounded"].map(lambda x: f"{x:+.1f}")
                            )
                            if wk2.empty:
                                st.info("Nobody to show after filtering zero-hour +0.0 entries.")
                            else:
                                bars = (
                                    alt.Chart(wk2)
                                    .mark_bar()
                                    .encode(
                                        x=alt.X("person:N", title="Person", sort=alt.Sort(field="person")),
                                        y=alt.Y("Actual:Q", title="Actual Hours"),
                                        tooltip=[
                                            "person:N",
                                            alt.Tooltip("Actual:Q", title="Actual Hours", format=",.1f"),
                                            alt.Tooltip("Avail:Q", title="Available Hours", format=",.1f"),
                                            alt.Tooltip("DiffRounded:Q", title="Over / Under", format="+.1f"),
                                            alt.Tooltip("period_date:T", title="Week"),
                                        ],
                                    )
                                    .properties(height=280)
                                )
                                labels = (
                                    alt.Chart(wk2)
                                    .mark_text(dy=-6)
                                    .encode(
                                        x="person:N",
                                        y="Actual:Q",
                                        text=alt.Text("DiffLabel:N"),
                                        color=alt.condition("datum.DiffRounded >= 0", alt.value("#22c55e"), alt.value("#ef4444")),
                                    )
                                )
                                st.altair_chart(bars + labels, use_container_width=True)
                                people_in_week = (
                                    wk2["person"].dropna().astype(str).str.strip().drop_duplicates().tolist()
                                    if "person" in wk2.columns and not wk2.empty else []
                                )
                                if people_in_week:
                                    picked_person_hours = st.selectbox(
                                        "Drill further: Person over time (per-station lines)",
                                        options=sorted(people_in_week),
                                        index=0,
                                        key="hours_person_over_time_select",
                                    )
                                    pt = build_person_station_hours_over_time(f, team_name, picked_person_hours)
                                    if pt.empty:
                                        st.caption("No nested per-station hours found for this person.")
                                    else:
                                        base_ts = alt.Chart(pt).encode(
                                            x=alt.X("period_date:T", title="Week"),
                                            y=alt.Y("Actual Hours:Q", title="Actual Hours"),
                                            color=alt.Color("cell_station:N", title="Cell/Station"),
                                            tooltip=[
                                                "period_date:T",
                                                alt.Tooltip("cell_station:N", title="Cell/Station"),
                                                alt.Tooltip("Actual Hours:Q", title="Actual", format=",.1f"),
                                                alt.Tooltip("Available Hours:Q", title="Available", format=",.1f"),
                                            ],
                                        )
                                        lines = base_ts.mark_line()
                                        pts   = base_ts.mark_point(size=70, filled=True)  # points same color as line
                                        st.altair_chart(
                                            (lines + pts).properties(
                                                height=280,
                                            ),
                                            use_container_width=True,
                                        )
                    pass
                else:
                    cells_hours = explode_cell_station_hours(f)
                    team_cells = cells_hours.loc[cells_hours["team"] == team_name].copy()
                    if team_cells.empty:
                        st.info(f"No cell/station data available for {team_name}.")
                    else:
                        wk_cells = team_cells.loc[team_cells["period_date"] == picked_week].copy()
                        if wk_cells.empty:
                            st.info("No cell/station data for the selected week.")
                        else:
                            wk_cells2 = (
                                wk_cells.assign(
                                    Actual=lambda d: pd.to_numeric(d["Actual Hours"], errors="coerce"),
                                    Avail=lambda d: pd.to_numeric(d["Available Hours"], errors="coerce"),
                                )
                                .assign(Diff=lambda d: d["Actual"] - d["Avail"])
                                .assign(DiffRounded=lambda d: d["Diff"].round(1))
                            )
                            wk_cells2 = wk_cells2.loc[
                                ~(
                                    (wk_cells2["Actual"].fillna(0) == 0) &
                                    (wk_cells2["Avail"].fillna(0)  == 0)
                                )
                            ].assign(
                                DiffLabel=lambda d: d["DiffRounded"].map(lambda x: f"{x:+.1f}")
                            )
                            if wk_cells2.empty:
                                st.info("Nothing to show after filtering zero-hour +0.0 entries.")
                            else:
                                order_cells = wk_cells2.sort_values("Actual", ascending=False)["cell_station"].tolist()
                                bars = (
                                    alt.Chart(wk_cells2)
                                    .mark_bar()
                                    .encode(
                                        x=alt.X("cell_station:N", title="Cell/Station", sort=order_cells),
                                        y=alt.Y("Actual:Q", title="Actual Hours"),
                                        tooltip=[
                                            alt.Tooltip("cell_station:N", title="Cell/Station"),
                                            alt.Tooltip("Actual:Q", title="Actual Hours", format=",.1f"),
                                            alt.Tooltip("period_date:T", title="Week"),
                                        ],
                                    )
                                    .properties(height=280)
                                )
                                labels = alt.Chart(wk_cells2).mark_text(dy=-6).encode(
                                    x="cell_station:N",
                                    y="Actual:Q",
                                )
                                st.altair_chart(bars + labels, use_container_width=True)
                            stations_in_week = (
                                wk_cells2["cell_station"].dropna().astype(str).str.strip().unique().tolist()
                                if ("cell_station" in wk_cells2.columns and not wk_cells2.empty)
                                else []
                            )
                            stations_in_week = (
                                wk_cells2["cell_station"].dropna().astype(str).str.strip().unique().tolist()
                                if ("cell_station" in wk_cells2.columns and not wk_cells2.empty)
                                else []
                            )
                            if stations_in_week:
                                picked_station_hours = st.selectbox(
                                    "Drill further: Station over time (per-person lines)",
                                    options=stations_in_week,
                                    index=0,
                                    key="hours_station_over_time_select",
                                )
                                ht = build_station_person_hours_over_time(f, team_name, picked_station_hours)
                                if ht.empty:
                                    st.caption("No nested per-person station-hours found. Showing station totals over time.")
                                    stn_hours_tot = (
                                        explode_cell_station_hours(f[f["team"] == team_name])
                                        .query("cell_station == @picked_station_hours")
                                        .dropna(subset=["period_date"])
                                        .rename(columns={"Actual Hours": "Actual", "Available Hours": "Target"})
                                    )
                                    if not stn_hours_tot.empty:
                                        base_ts = alt.Chart(stn_hours_tot).encode(
                                            x=alt.X("period_date:T", title="Week"),
                                            y=alt.Y("Actual:Q", title="Actual Hours"),
                                            tooltip=[
                                                "period_date:T",
                                                alt.Tooltip("Actual:Q", title="Actual", format=",.1f"),
                                                alt.Tooltip("Target:Q", title="Available", format=",.1f"),
                                            ],
                                        )
                                        line_a = base_ts.mark_line(point=True)
                                        line_t = (
                                            alt.Chart(stn_hours_tot)
                                            .mark_line(point=True, strokeDash=[4, 3], color="#6b7280")
                                            .encode(x="period_date:T", y=alt.Y("Target:Q", title="Available"))
                                        )
                                        st.altair_chart(
                                            (line_a + line_t).properties(
                                                height=280,
                                                title=f"{picked_station_hours} • Hours over time (station total)",
                                            ),
                                            use_container_width=True,
                                        )
                                else:
                                    base_ts = alt.Chart(ht).encode(
                                        x=alt.X("period_date:T", title="Week"),
                                        y=alt.Y("Actual Hours:Q", title="Actual Hours"),
                                        color=alt.Color("person:N", title="Person"),
                                        tooltip=[
                                            "period_date:T",
                                            "person:N",
                                            alt.Tooltip("Actual Hours:Q", title="Actual", format=",.1f"),
                                            alt.Tooltip("Available Hours:Q", title="Available", format=",.1f"),
                                        ],
                                    )
                                    lines = base_ts.mark_line()
                                    pts   = base_ts.mark_point(size=70, filled=True)  # points = line color
                                    st.altair_chart(
                                        (lines + pts).properties(
                                            height=280,
                                        ),
                                        use_container_width=True,
                                    )
with mid:
    st.subheader("Output Trend")
    out_long = (
        f.melt(
            id_vars=["team", "period_date"],
            value_vars=["Target Output", "Actual Output"],
            var_name="Metric", value_name="Value"
        ).dropna(subset=["Value"])
    )
    base = alt.Chart(out_long).encode(
        x=alt.X("period_date:T", title="Week"),
        y=alt.Y("Value:Q", title="Output"),
        color=alt.Color("Metric:N", title="Series"),
        tooltip=["team:N", "period_date:T", "Metric:N", alt.Tooltip("Value:Q", format=",.0f")]
    )
    line = base.mark_line().encode(
        detail="team:N",
        opacity=alt.condition(team_sel, alt.value(1.0), alt.value(0.25)) if multi_team else alt.value(1.0)
    )
    pts = base.mark_point().encode(
        shape=alt.Shape("team:N", title="Team") if multi_team else alt.value("circle"),
        size=alt.value(45),
        opacity=alt.condition(team_sel, alt.value(1.0), alt.value(0.25)) if multi_team else alt.value(1.0)
    )
    st.altair_chart((line + pts).properties(height=280).add_params(team_sel), use_container_width=True)
    if len(teams_in_view) != 1:
        st.caption("Select exactly one team to enable week drilldown.")
    else:
        team_name = teams_in_view[0]
        team_weeks = sorted(pd.to_datetime(f.loc[f["team"] == team_name, "period_date"].dropna().unique()), reverse=True)
        if not team_weeks:
            st.info("No weeks available for drilldown.")
        else:
            by_choice = st.selectbox(
                "Output by:",
                options=["Cell/Station", "Person"],
                index=0,
                key="output_by_select"
            )
            col_map = {
                "Person": ("Outputs by Person", "person"),
                "Cell/Station": ("Outputs by Cell/Station", "cell_station"),
            }
            col_name, key_label = col_map[by_choice]
            picked_week = st.selectbox(
                "Week:",
                options=team_weeks,
                index=0,
                format_func=lambda d: pd.to_datetime(d).date().isoformat(),
                key="output_by_week_select",
            )
            picked_week = pd.to_datetime(picked_week).normalize()
            if col_name not in f.columns:
                st.info(f"No '{col_name}' data available.")
            else:
                exploded = explode_outputs_json(f[f["team"] == team_name], col_name, key_label)
                if exploded.empty or "period_date" not in exploded.columns:
                    st.info("No drilldown records for the selected grouping.")
                else:
                    wk = exploded.loc[exploded["period_date"] == picked_week].copy()
                    if wk.empty:
                        st.info("No data for the selected week.")
                    else:
                        wk2 = (
                            wk.assign(
                                Actual=pd.to_numeric(wk["Actual"], errors="coerce"),
                                Target=pd.to_numeric(wk["Target"], errors="coerce"),
                            )
                            .dropna(subset=["Actual"])
                            .assign(
                                HasTarget=lambda d: d["Target"].notna(),
                                Diff=lambda d: d["Actual"] - d["Target"],
                            )
                        )
                        wk2["DiffRounded"] = np.where(wk2["HasTarget"], wk2["Diff"].round(1), np.nan)
                        wk2["DiffLabel"]   = np.where(wk2["HasTarget"], wk2["DiffRounded"].map(lambda x: f"{x:+.1f}"), "—")
                        wk2 = wk2.loc[~((wk2["Actual"].fillna(0) == 0) & (wk2["Target"].fillna(0) == 0))].copy()
                        order_keys = wk2.sort_values("Actual", ascending=False)[key_label].tolist()
                        if not wk2.empty:
                            vmax = float(pd.to_numeric(wk2["Actual"], errors="coerce").max())
                            rng  = max(0.0, vmax)
                            pad  = max(3.0, rng * 0.22)
                            y_scale = alt.Scale(domain=[0.0, vmax + pad], nice=False, clamp=False)
                            label_pad = max(1.0, (vmax + pad) * 0.04)
                            wk2["LabelY"] = wk2["Actual"] + np.where(wk2["DiffRounded"].fillna(-1) >= 0, label_pad, -label_pad)
                        else:
                            y_scale = alt.Scale()
                            wk2["LabelY"] = wk2["Actual"]
                        bars = (
                            alt.Chart(wk2)
                            .mark_bar()
                            .encode(
                                x=alt.X(f"{key_label}:N", title=by_choice, sort=order_keys),
                                y=alt.Y("Actual:Q", title="Actual Output", scale=y_scale),
                                tooltip=[
                                    alt.Tooltip(f"{key_label}:N", title=by_choice),
                                    alt.Tooltip("Actual:Q", title="Actual", format=",.0f"),
                                    alt.Tooltip("Target:Q", title="Target", format=",.0f"),
                                    alt.Tooltip("DiffLabel:N", title="Over / Under"),  # <- use string label, shows "—" if no target
                                    alt.Tooltip("period_date:T", title="Week"),
                                ],
                            )
                            .properties(height=280)
                        )
                        labels = (
                            alt.Chart(wk2)
                            .mark_text()
                            .encode(
                                x=f"{key_label}:N",
                                y=alt.Y("LabelY:Q", scale=y_scale),
                                text="DiffLabel:N",  # <- no "nan"
                                color=alt.condition("datum.DiffRounded >= 0", alt.value("#22c55e"), alt.value("#ef4444")),
                            )
                        )
                        st.altair_chart(bars + labels, use_container_width=True)
                        if by_choice == "Cell/Station":
                            stations_in_week = wk2["cell_station"].dropna().unique().tolist()
                            if stations_in_week:
                                picked_station = st.selectbox(
                                    "Drill further: Station over time (per-person lines)",
                                    options=stations_in_week,
                                    index=0,
                                    key="outputs_station_over_time_select",
                                )
                                ot = build_station_person_outputs_over_time(f, team_name, picked_station)
                                if ot.empty:
                                    st.caption("No nested per-person outputs found for this station. Showing station totals over time.")
                                    stn_long = (
                                        explode_outputs_json(
                                            f[f["team"] == team_name],
                                            "Outputs by Cell/Station",
                                            "cell_station"
                                        )
                                        .query("cell_station == @picked_station")
                                        .rename(columns={"Actual": "Value", "Target": "Target"})
                                        .dropna(subset=["period_date"])
                                    )
                                    if not stn_long.empty:
                                        base_ts = alt.Chart(stn_long).encode(
                                            x=alt.X("period_date:T", title="Week"),
                                            y=alt.Y("Value:Q", title="Actual Output"),
                                            tooltip=[
                                                "period_date:T",
                                                alt.Tooltip("Value:Q", title="Actual", format=",.0f"),
                                                alt.Tooltip("Target:Q", title="Target", format=",.0f"),
                                            ],
                                        )
                                        line_a = base_ts.mark_line(point=True)
                                        line_t = (
                                            alt.Chart(stn_long)
                                            .mark_line(point=True, strokeDash=[4, 3])
                                            .encode(x="period_date:T", y=alt.Y("Target:Q", title="Target"))
                                        )
                                        st.altair_chart(
                                            (line_a + line_t).properties(
                                                height=280,
                                                title=f"{picked_station} • Outputs over time (station total)"
                                            ),
                                            use_container_width=True
                                        )
                                else:
                                    ot = ot.assign(
                                        HasTarget=lambda d: d["Target"].notna(),
                                        Delta=lambda d: d["Actual"] - d["Target"],
                                    )
                                    ot = ot.assign(
                                        DiffLabel=lambda d: np.where(d["HasTarget"], d["Delta"].round(1).map(lambda x: f"{x:+.1f}"), "—"),
                                        PointFillGroup=lambda d: np.where(~d["HasTarget"], "none", np.where(d["Delta"] >= 0, "pos", "neg")),
                                    )
                                    base_ts = alt.Chart(ot).encode(
                                        x=alt.X("period_date:T", title="Week"),
                                        y=alt.Y("Actual:Q", title="Actual Output"),
                                        color=alt.Color("person:N", title="Person"),  # line/stroke color
                                        tooltip=[
                                            "period_date:T",
                                            "person:N",
                                            alt.Tooltip("Actual:Q", title="Actual", format=",.0f"),
                                            alt.Tooltip("Target:Q", title="Target", format=",.0f"),
                                            alt.Tooltip("DiffLabel:N", title="Over / Under"),
                                        ],
                                    )
                                    lines = base_ts.mark_line()
                                    pts = (
                                        alt.Chart(ot)
                                        .mark_point(size=85, filled=True, strokeWidth=1.2)
                                        .encode(
                                            x="period_date:T",
                                            y="Actual:Q",
                                            stroke=alt.Color("person:N", legend=None),
                                            fill=alt.Color(
                                                "PointFillGroup:N",
                                                legend=None,
                                                scale=alt.Scale(
                                                    domain=["pos", "neg", "none"],
                                                    range=["#22c55e", "#ef4444", "#9ca3af"]
                                                ),
                                            ),
                                            tooltip=[
                                                "period_date:T",
                                                "person:N",
                                                alt.Tooltip("Actual:Q", title="Actual", format=",.0f"),
                                                alt.Tooltip("Target:Q", title="Target", format=",.0f"),
                                                alt.Tooltip("DiffLabel:N", title="Over / Under"),
                                            ],
                                        )
                                    )
                                    st.altair_chart(
                                        (lines + pts)
                                            .properties(height=280),
                                        use_container_width=True
                                    )
                        elif by_choice == "Person":
                            people_in_week = wk2["person"].dropna().unique().tolist()
                            if people_in_week:
                                picked_person = st.selectbox(
                                    "Drill further: Person over time (per-station lines)",
                                    options=people_in_week,
                                    index=0,
                                    key="outputs_person_over_time_select",
                                )
                                pt = build_person_station_outputs_over_time(f, team_name, picked_person)
                                if pt.empty:
                                    st.caption("No nested per-station outputs found for this person.")
                                else:
                                    pt = pt.assign(
                                        HasTarget=lambda d: d["Target"].notna(),
                                        Delta=lambda d: d["Actual"] - d["Target"],
                                    )
                                    pt = pt.assign(
                                        DiffLabel=lambda d: np.where(d["HasTarget"], d["Delta"].round(1).map(lambda x: f"{x:+.1f}"), "—"),
                                        PointFillGroup=lambda d: np.where(~d["HasTarget"], "none", np.where(d["Delta"] >= 0, "pos", "neg")),
                                    )
                                    base_ts = alt.Chart(pt).encode(
                                        x=alt.X("period_date:T", title="Week"),
                                        y=alt.Y("Actual:Q", title="Actual Output"),
                                        color=alt.Color("cell_station:N", title="Cell/Station"),
                                        tooltip=[
                                            "period_date:T",
                                            alt.Tooltip("cell_station:N", title="Cell/Station"),
                                            alt.Tooltip("Actual:Q", title="Actual", format=",.0f"),
                                            alt.Tooltip("Target:Q", title="Target", format=",.0f"),
                                            alt.Tooltip("DiffLabel:N", title="Over / Under"),
                                        ],
                                    )
                                    lines = base_ts.mark_line()
                                    pts = (
                                        alt.Chart(pt)
                                        .mark_point(size=85, filled=True, strokeWidth=1.2)
                                        .encode(
                                            x="period_date:T",
                                            y="Actual:Q",
                                            stroke=alt.Color("cell_station:N", legend=None),  # stroke matches line color
                                            fill=alt.Color(
                                                "PointFillGroup:N",
                                                legend=None,
                                                scale=alt.Scale(
                                                    domain=["pos", "neg", "none"],
                                                    range=["#22c55e", "#ef4444", "#9ca3af"]
                                                ),
                                            ),
                                            tooltip=[
                                                "period_date:T",
                                                alt.Tooltip("cell_station:N", title="Cell/Station"),
                                                alt.Tooltip("Actual:Q", title="Actual", format=",.0f"),
                                                alt.Tooltip("Target:Q", title="Target", format=",.0f"),
                                                alt.Tooltip("DiffLabel:N", title="Over / Under"),
                                            ],
                                        )
                                    )
                                    st.altair_chart(
                                        (lines + pts)
                                            .properties(height=280),
                                        use_container_width=True
                                    )
with right:
    st.subheader("UPLH Trend")
    team_sel = alt.selection_point(fields=["team"], bind="legend")
    have_target_uplh = "Target UPLH" in f.columns
    uplh_vars = ["Actual UPLH"] + (["Target UPLH"] if have_target_uplh else [])
    uplh_long = (
        f.melt(
            id_vars=["team", "period_date"],
            value_vars=uplh_vars,
            var_name="Metric",
            value_name="Value",
        )
        .dropna(subset=["Value"])
    )
    if not uplh_long.empty:
        vmin = float(pd.to_numeric(uplh_long["Value"], errors="coerce").min())
        vmax = float(pd.to_numeric(uplh_long["Value"], errors="coerce").max())
        rng  = max(0.0, vmax - vmin)
        pad  = max(0.2, rng * 0.15)
        lo   = max(0.0, vmin - pad)
        hi   = vmax + pad
        y_scale = alt.Scale(domain=[lo, hi], nice=False, clamp=False)
    else:
        y_scale = alt.Scale()
    sel_wk = alt.selection_point(
        name="wk_uplh",
        fields=["period_date"],
        on="click",
        clear="dblclick",
        empty="none",
    )
    trend_base = (
        alt.Chart(uplh_long)
        .encode(
            x=alt.X("period_date:T", title="Week"),
            y=alt.Y("Value:Q", title="Actual UPLH", scale=y_scale),
            color=alt.Color("Metric:N", title="Series"),
            tooltip=["team:N", "period_date:T", "Metric:N", alt.Tooltip("Value:Q", format=",.2f")],
        )
    )
    line = trend_base.mark_line().encode(
        detail="team:N",
        opacity=alt.condition(team_sel, alt.value(1.0), alt.value(0.25)) if multi_team else alt.value(1.0),
    )
    pts = trend_base.mark_point(size=70).encode(
        shape=alt.Shape("team:N", title="Team") if multi_team else alt.value("circle"),
        opacity=alt.condition(team_sel, alt.value(1.0), alt.value(0.25)) if multi_team else alt.value(1.0),
    )
    rule = (
        alt.Chart(uplh_long)
        .transform_filter(sel_wk)
        .mark_rule(strokeDash=[4, 3])
        .encode(x="period_date:T")
    )
    top = alt.layer(line, pts, rule).properties(height=280).add_params(team_sel, sel_wk)
    def _find_wp_uplh_cols(df: pd.DataFrame) -> tuple[str | None, str | None]:
        wp1, wp2 = None, None
        for c in df.columns:
            lc = str(c).lower().replace(" ", "")
            if "uplh" not in lc:
                continue
            if any(tag in lc for tag in ("wp1", "wp01", "wp_1", "wp-1")) and wp1 is None:
                wp1 = c
            if any(tag in lc for tag in ("wp2", "wp02", "wp_2", "wp-2")) and wp2 is None:
                wp2 = c
        return wp1, wp2
    wp1_col, wp2_col = _find_wp_uplh_cols(f)
    team_for_drill = teams_in_view[0] if not multi_team and teams_in_view else None
    if (not multi_team) and team_for_drill == "PH" and wp1_col and wp2_col:
        wp_long = (
            f[["team", "period_date", wp1_col, wp2_col]]
            .rename(columns={wp1_col: "WP1", wp2_col: "WP2"})
            .melt(id_vars=["team", "period_date"], var_name="WP", value_name="UPLH")
            .dropna(subset=["UPLH"])
        )
        title_text = (
            alt.Chart(uplh_long)
            .transform_filter(sel_wk)
            .transform_aggregate(period_date="min(period_date)")
            .transform_calculate(label="'WP1 vs WP2 UPLH (' + timeFormat(datum.period_date, '%Y-%m-%d') + ')'")
            .mark_text(align="left", baseline="top")
            .encode(x=alt.value(0), y=alt.value(16), text="label:N")
            .properties(height=24)
        )
        base_wp = (
            alt.Chart(wp_long)
            .transform_filter(sel_wk)
            .transform_filter(team_sel)
        )
        wp_chart = (
            base_wp.mark_bar()
            .encode(
                x=alt.X("WP:N", title="WP"),
                y=alt.Y("UPLH:Q", title="Actual UPLH", axis=alt.Axis(titlePadding=12, labelPadding=6)),
                color=alt.Color("WP:N", legend=None),
                tooltip=["period_date:T", "WP:N", alt.Tooltip("UPLH:Q", format=",.2f")],
            )
            .properties(height=280)
        )
        combined = alt.vconcat(top, title_text, wp_chart, spacing=0).resolve_legend(color="independent").add_params(team_sel, sel_wk)
        st.altair_chart(combined, use_container_width=True)
    elif not multi_team and team_for_drill is not None:
        top_ph = st.empty()
        top_ph.altair_chart(top, use_container_width=True)
        by_choice = st.selectbox(
            "UPLH by:",
            options=["Person", "Cell/Station"],
            index=0,
            key="uplh_by_select",
        )
        team_weeks = sorted(
            pd.to_datetime(f.loc[f["team"] == team_for_drill, "period_date"].dropna().unique()),
            reverse=True
        )
        if team_weeks:
            picked_week = st.selectbox(
                "Week:",
                options=team_weeks,
                index=0,
                format_func=lambda d: pd.to_datetime(d).date().isoformat(),
                key="uplh_week_select",
            )
            picked_week = pd.to_datetime(picked_week).normalize()
            rule_df = pd.DataFrame({"period_date": [picked_week]})
            rule_week = alt.Chart(rule_df).mark_rule(strokeDash=[4, 3]).encode(x="period_date:T")
            top_ph.altair_chart(
                alt.layer(line, pts, rule_week).properties(height=280).add_params(team_sel),
                use_container_width=True
            )
        else:
            picked_week = None
            st.info("No weeks available for drilldown.")
        lower = None
        drill = None 
        lower_area   = st.container()
        controls_area = st.container()
        drill_area   = st.container()
        if picked_week is not None:
            if by_choice == "Person":
                uplh_person = build_uplh_by_person_long(f, team_for_drill)
                if uplh_person.empty:
                    st.info("No 'Outputs by Person' and/or 'Person Hours' data to compute UPLH.")
                else:
                    wk_p = uplh_person.loc[uplh_person["period_date"] == picked_week].copy()
                    if wk_p.empty:
                        st.info("No UPLH-by-person records for that week.")
                    else:
                        wk_p["HasTarget"]   = wk_p["Target UPLH"].notna()
                        wk_p["Delta"]       = wk_p["Actual UPLH"] - wk_p["Target UPLH"]
                        wk_p["DeltaRounded"]= wk_p["Delta"].round(2)
                        wk_p["DeltaLabel"]  = np.where(wk_p["HasTarget"], wk_p["DeltaRounded"].map(lambda x: f"{x:+.2f}"), "—")
                        wk_p["LabelGroup"]  = np.where(~wk_p["HasTarget"], "none", np.where(wk_p["Delta"] >= 0, "pos", "neg"))
                        order_people = wk_p.sort_values("Actual UPLH", ascending=False)["person"].tolist()
                        vmax = float(pd.to_numeric(wk_p["Actual UPLH"], errors="coerce").max())
                        pad  = max(0.1, vmax * 0.12) if pd.notna(vmax) else 0.3
                        y_scale = alt.Scale(domain=[0, (vmax + pad) if pd.notna(vmax) else 1.0], nice=False, clamp=False)
                        bars = (
                            alt.Chart(wk_p)
                            .mark_bar(color="#2563eb")
                            .encode(
                                x=alt.X("person:N", title="Person", sort=order_people),
                                y=alt.Y("Actual UPLH:Q", title="Actual UPLH", scale=y_scale),
                                tooltip=[
                                    "period_date:T", "person:N",
                                    alt.Tooltip("Actual Output:Q", title="Actual Output", format=",.0f"),
                                    alt.Tooltip("Target Output:Q", title="Target Output", format=",.0f"),
                                    alt.Tooltip("Actual Hours:Q", title="Hours (actual)", format=",.1f"),
                                    alt.Tooltip("Actual UPLH:Q", title="Actual UPLH", format=",.2f"),
                                    alt.Tooltip("Target UPLH:Q", title="Target UPLH", format=",.2f"),
                                    alt.Tooltip("DeltaRounded:Q", title="Δ vs Target", format="+.2f"),
                                ],
                            )
                            .properties(height=280)
                        )
                        label_pad = max(0.05, (vmax + pad) * 0.03) if pd.notna(vmax) else 0.08
                        labels = (
                            alt.Chart(
                                wk_p.assign(LabelY=lambda d: d["Actual UPLH"] + np.where(d["Delta"].fillna(-1) >= 0, label_pad, -label_pad))
                            )
                            .mark_text(dy=-10)
                            .encode(
                                x="person:N",
                                y=alt.Y("LabelY:Q", scale=y_scale),
                                text="DeltaLabel:N",
                                color=alt.Color(
                                    "LabelGroup:N",
                                    legend=None,
                                    scale=alt.Scale(domain=["pos","neg","none"], range=["#22c55e","#ef4444","#9ca3af"]),
                                ),
                            )
                        )
                        lower = bars + labels
                        people_in_week = wk_p["person"].dropna().unique().tolist()
                        if people_in_week:
                            picked_person_uplh = controls_area.selectbox(
                                "Drill further: Person UPLH over time (per-station lines)",
                                options=sorted(people_in_week),
                                index=0,
                                key="uplh_person_over_time_select",
                            )
                            pu = build_person_station_uplh_over_time(f, team_for_drill, picked_person_uplh)
                            if pu.empty:
                                st.caption("No nested per-station UPLH found for this person.")
                            else:
                                pu = pu.assign(
                                    HasTarget=lambda d: d["Target UPLH"].notna(),
                                    Delta=lambda d: d["Actual UPLH"] - d["Target UPLH"],
                                    DiffLabel=lambda d: np.where(d["HasTarget"], d["Delta"].round(2).map(lambda x: f"{x:+.2f}"), "—"),
                                    LabelGroup=lambda d: np.where(~d["HasTarget"], "none", np.where(d["Delta"] >= 0, "pos", "neg")),
                                )
                                base_ts = alt.Chart(pu).encode(
                                    x=alt.X("period_date:T", title="Week"),
                                    y=alt.Y("Actual UPLH:Q", title="Actual UPLH"),
                                    color=alt.Color("cell_station:N", title="Cell/Station"),
                                    tooltip=[
                                        "period_date:T",
                                        alt.Tooltip("cell_station:N", title="Cell/Station"),
                                        alt.Tooltip("Actual:Q", title="Actual Output", format=",.0f"),
                                        alt.Tooltip("Target:Q", title="Target Output", format=",.0f"),
                                        alt.Tooltip("Actual Hours:Q", title="Hours (actual)", format=",.2f"),
                                        alt.Tooltip("Actual UPLH:Q", title="Actual UPLH", format=",.2f"),
                                        alt.Tooltip("Target UPLH:Q", title="Target UPLH", format=",.2f"),
                                        alt.Tooltip("DiffLabel:N", title="Δ vs Target"),
                                    ],
                                )
                                lines = base_ts.mark_line()
                                pts   = (
                                    alt.Chart(pu)
                                    .mark_point(size=85, filled=True, strokeWidth=1.2)
                                    .encode(
                                        x="period_date:T",
                                        y="Actual UPLH:Q",
                                        stroke=alt.Color("cell_station:N", legend=None),
                                        fill=alt.Color(
                                            "LabelGroup:N",
                                            legend=None,
                                            scale=alt.Scale(domain=["pos","neg","none"], range=["#22c55e","#ef4444","#9ca3af"]),
                                        ),
                                    )
                                )
                                drill = (lines + pts).properties(
                                    height=280,
                                )
            else:  # by_choice == "Cell/Station"
                uplh_cell = build_uplh_by_cell_long(f, team_for_drill)
                if uplh_cell.empty:
                    st.info("No 'Outputs by Cell/Station' and/or 'Cell/Station Hours' data to compute UPLH.")
                else:
                    wk_c = uplh_cell.loc[uplh_cell["period_date"] == picked_week].copy()
                    if wk_c.empty:
                        st.info("No UPLH-by-cell/station records for that week.")
                    else:
                        wk_c["HasTarget"]   = wk_c["Target UPLH"].notna()
                        wk_c["Delta"]       = wk_c["Actual UPLH"] - wk_c["Target UPLH"]
                        wk_c["DeltaRounded"]= wk_c["Delta"].round(2)
                        wk_c["DeltaLabel"]  = np.where(wk_c["HasTarget"], wk_c["DeltaRounded"].map(lambda x: f"{x:+.2f}"), "—")
                        wk_c["LabelGroup"]  = np.where(~wk_c["HasTarget"], "none", np.where(wk_c["Delta"] >= 0, "pos", "neg"))
                        order_cells = wk_c.sort_values("Actual UPLH", ascending=False)["cell_station"].tolist()
                        vmax = float(pd.to_numeric(wk_c["Actual UPLH"], errors="coerce").max())
                        pad  = max(0.1, vmax * 0.12) if pd.notna(vmax) else 0.3
                        y_scale = alt.Scale(domain=[0, (vmax + pad) if pd.notna(vmax) else 1.0], nice=False, clamp=False)
                        label_pad = max(0.05, (vmax + pad) * 0.03) if pd.notna(vmax) else 0.08
                        wk_c["LabelY"] = wk_c["Actual UPLH"] + np.where(wk_c["Delta"].fillna(-1) >= 0, label_pad, -label_pad)
                        bars = (
                            alt.Chart(wk_c)
                            .mark_bar(color="#2563eb")
                            .encode(
                                x=alt.X(
                                    "cell_station:N",
                                    sort=order_cells,
                                    axis=alt.Axis(
                                        title="Cell/Station",
                                        labelAngle=-35,      # <- ensures the x-axis shows and is readable
                                        labelOverlap=False,
                                    ),
                                ),
                                y=alt.Y("Actual UPLH:Q", title="Actual UPLH", scale=y_scale),
                                tooltip=[
                                    alt.Tooltip("cell_station:N", title="Cell/Station"),
                                    alt.Tooltip("Actual UPLH:Q", title="Actual UPLH", format=",.2f"),
                                    alt.Tooltip("Target UPLH:Q", title="Target UPLH", format=",.2f"),
                                    alt.Tooltip("DeltaRounded:Q", title="Δ vs Target", format="+.2f"),
                                    "period_date:T",
                                ],
                            )
                            .properties(height=280)
                        )
                        labels = (
                            alt.Chart(wk_c)
                            .mark_text()
                            .encode(
                                x="cell_station:N",
                                y=alt.Y("LabelY:Q", scale=y_scale),
                                text="DeltaLabel:N",  # already formatted with +/-
                                color=alt.Color(
                                    "LabelGroup:N",
                                    legend=None,
                                    scale=alt.Scale(
                                        domain=["pos", "neg", "none"],
                                        range=["#22c55e", "#ef4444", "#9ca3af"],  # green / red / gray
                                    ),
                                ),
                            )
                        )
                        lower = bars + labels
                        stations_in_week = wk_c["cell_station"].dropna().astype(str).str.strip().unique().tolist()
                        if stations_in_week:
                            picked_station_uplh = controls_area.selectbox(
                                "Drill further: Station UPLH over time (per-person lines)",
                                options=stations_in_week,
                                index=0,
                                key="uplh_station_over_time_select",
                            )
                            ut = build_station_person_uplh_over_time(f, team_for_drill, picked_station_uplh)
                            if ut.empty:
                                st.caption("No nested per-person station-hours found. Showing station UPLH totals over time.")
                                stn_uplh_tot = (
                                    build_uplh_by_cell_long(f, team_for_drill)
                                    .query("cell_station == @picked_station_uplh")
                                    .dropna(subset=["period_date"])
                                )
                                if not stn_uplh_tot.empty:
                                    base_ts = alt.Chart(stn_uplh_tot).encode(
                                        x=alt.X("period_date:T", title="Week"),
                                        y=alt.Y("Actual UPLH:Q", title="Actual UPLH"),
                                        tooltip=[
                                            "period_date:T",
                                            alt.Tooltip("Actual UPLH:Q", title="Actual UPLH", format=",.2f"),
                                            alt.Tooltip("Target UPLH:Q", title="Target UPLH", format=",.2f"),
                                        ],
                                    )
                                    line_a = base_ts.mark_line(point=True, color="#2563eb")
                                    line_t = (
                                        alt.Chart(stn_uplh_tot)
                                        .mark_line(point=True, strokeDash=[4,3], color="#6b7280")
                                        .encode(x="period_date:T", y=alt.Y("Target UPLH:Q", title="Target UPLH"))
                                    )
                                    drill = (line_a + line_t).properties(
                                        height=280,
                                        title=f"{picked_station_uplh} • UPLH over time (station total)"
                                    )
                            else:
                                ut = ut.assign(
                                    HasTarget=lambda d: d["Target UPLH"].notna(),
                                    Delta=lambda d: d["Actual UPLH"] - d["Target UPLH"],
                                    DiffLabel=lambda d: np.where(d["HasTarget"], d["Delta"].round(2).map(lambda x: f"{x:+.2f}"), "—"),
                                    LabelGroup=lambda d: np.where(~d["HasTarget"], "none", np.where(d["Delta"] >= 0, "pos", "neg")),
                                )
                                base_ts = alt.Chart(ut).encode(
                                    x=alt.X("period_date:T", title="Week"),
                                    y=alt.Y("Actual UPLH:Q", title="Actual UPLH"),
                                    color=alt.Color("person:N", title="Person"),
                                    tooltip=[
                                        "period_date:T",
                                        "person:N",
                                        alt.Tooltip("Actual:Q", title="Actual Output", format=",.0f"),
                                        alt.Tooltip("Target:Q", title="Target Output", format=",.0f"),
                                        alt.Tooltip("Actual Hours:Q", title="Hours (actual)", format=",.2f"),
                                        alt.Tooltip("Actual UPLH:Q", title="Actual UPLH", format=",.2f"),
                                        alt.Tooltip("Target UPLH:Q", title="Target UPLH", format=",.2f"),
                                        alt.Tooltip("DiffLabel:N", title="Δ vs Target"),
                                    ],
                                )
                                lines = base_ts.mark_line()
                                pts = (
                                    alt.Chart(ut)
                                    .mark_point(size=85, filled=True, strokeWidth=1.2)
                                    .encode(
                                        x="period_date:T",
                                        y="Actual UPLH:Q",
                                        stroke=alt.Color("person:N", legend=None),
                                        fill=alt.Color(
                                            "LabelGroup:N",
                                            legend=None,
                                            scale=alt.Scale(domain=["pos","neg","none"], range=["#22c55e","#ef4444","#9ca3af"]),
                                        ),
                                    )
                                )
                                drill = (lines + pts).properties(
                                    height=280,
                                )
        if lower is not None:
            lower_area.altair_chart(lower, use_container_width=True)
        else:
            lower_area.altair_chart(top, use_container_width=True)
        if drill is not None:
            drill_area.altair_chart(drill, use_container_width=True)
st.markdown("---")
left2, mid2, right2 = st.columns(3) 
with left2:
    st.subheader("HC in WIP Trend")
    if "HC in WIP" in f.columns and f["HC in WIP"].notna().any():
        hc = f[["team", "period_date", "HC in WIP"]].dropna()
        base_hc = alt.Chart(hc).encode(
            x=alt.X("period_date:T", title="Week"),
            y=alt.Y("HC in WIP:Q", title="HC in WIP"),
            color=alt.Color("team:N", title="Team") if len(teams_in_view) > 1 else alt.value("steelblue"),
            tooltip=["team:N", "period_date:T", alt.Tooltip("HC in WIP:Q", format=",.0f")]
        )
        st.altair_chart(
            base_hc.mark_line(point=True).properties(height=280),
            use_container_width=True
        )
    else:
        st.info("No 'HC in WIP' data available in the selected range.")
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
            use_container_width=True
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

                        st.altair_chart(bars + labels + ref, use_container_width=True)
        else:
            st.caption("Select exactly one team to drill into per-person daily hours.")
    else:
        st.info("No 'Actual HC used' data available in the selected range.")
if len(teams_in_view) == 1:
    team_name = teams_in_view[0]
    st.subheader(f"{team_name} • Multi-Axis View")
    single = (
        f[f["team"] == team_name]
        .sort_values("period_date")
        .copy()
    )
    nw_all = load_non_wip()
    if not nw_all.empty:
        cols_needed = ["team", "period_date", "total_non_wip_hours"]
        for c in cols_needed:
            if c not in nw_all.columns:
                nw_all[c] = np.nan
        nw_all["total_non_wip_hours"] = pd.to_numeric(nw_all["total_non_wip_hours"], errors="coerce")
        single = single.merge(nw_all[cols_needed], on=["team", "period_date"], how="left")
    else:
        single["total_non_wip_hours"] = np.nan
    if "Closures" in single.columns:
        denom = (single["Completed Hours"].fillna(0).astype(float) + single["total_non_wip_hours"].fillna(0).astype(float))
        num   = single["Closures"].astype(float)
        single["Productivity"] = np.where(denom > 0, num / denom, np.nan)
    else:
        single["Productivity"] = np.nan
    if {"Closures", "Completed Hours"}.issubset(single.columns):
        denom = single["Completed Hours"].astype(float)
        single["Efficiency"] = np.where(denom > 0, single["Closures"].astype(float) / denom, np.nan)
    else:
        single["Efficiency"] = np.nan
    metric_options = [
        "HC in WIP",
        "Open Complaint Timeliness",
        "Actual UPLH",
        "Actual Output",
        "Actual Hours",
        "Actual HC used",
        "Closures",
        "Efficiency",
        "Productivity",
    ]
    available = []
    for opt in metric_options:
        if opt == "Actual Hours":
            if "Completed Hours" in single.columns:
                available.append(opt)
        elif opt in single.columns:
            available.append(opt)
    selected = st.multiselect(
        "Series",
        available,
        default=available,
        key="single_team_series",
        help="Tip: select exactly one series to enable a 3-month projection"  # NEW
    )
    if len(selected) != 1:
        st.caption("Select **one** series to enable the 3-month projection.")
    if selected:
        display_to_col = {
            "HC in WIP": "HC in WIP",
            "Open Complaint Timeliness": "Open Complaint Timeliness",
            "Actual UPLH": "Actual UPLH",
            "Actual Output": "Actual Output",
            "Actual Hours": "Completed Hours",
            "Actual HC used": "Actual HC used",
            "Closures": "Closures",
            "Efficiency": "Efficiency",
            "Productivity": "Productivity",
        }
        base = alt.Chart(single).encode(x=alt.X("period_date:T", title="Week"))
        def tooltip_for(metric: str):
            col = display_to_col[metric]
            if metric == "Open Complaint Timeliness":
                return ["period_date:T", "metric:N", alt.Tooltip(f"{col}:Q", format=".0%")]
            if metric in ("Actual UPLH", "Actual HC used"):
                return ["period_date:T", "metric:N", alt.Tooltip(f"{col}:Q", format=".2f")]
            if metric == "Actual UPLH":
                return ["period_date:T", "metric:N", alt.Tooltip(f"{col}:Q", format=".2f")]
            if metric == "Productivity":                          
                return ["period_date:T", "metric:N", alt.Tooltip(f"{col}:Q", format=".3f")]
            if metric in ("Efficiency", "Productivity"):
                return ["period_date:T", "metric:N", alt.Tooltip(f"{col}:Q", format=".3f")]
            if metric == "Closures":                         
                return ["period_date:T", "metric:N", alt.Tooltip(f"{col}:Q", format=",.0f")]
            return ["period_date:T", "metric:N", alt.Tooltip(f"{col}:Q", format=",.0f")]
        color_enc = alt.Color("metric:N", title="Series")
        single_sel = (len(selected) == 1)
        def axis_for(metric: str) -> alt.Axis:
            title = metric if single_sel else None
            show = single_sel
            kwargs = dict(title=title, labels=show, ticks=show, domain=show)
            if metric == "Open Complaint Timeliness":
                kwargs["format"] = "%"
            return alt.Axis(**kwargs)
        def y_enc(metric: str, field: str) -> alt.Y:
            ax = axis_for(metric)
            if metric == "Open Complaint Timeliness":
                col = display_to_col[metric]
                vals = single[col].dropna().astype(float)
                if len(vals):
                    vmin = float(vals.min())
                    vmax = float(vals.max())
                else:
                    vmin, vmax = 0.0, 1.0
                rng = max(0.0, vmax - vmin)
                pad = max(0.02, rng * 0.15)
                lo = max(0.0, vmin - pad)
                hi = min(1.0, vmax + pad)
                return alt.Y(f"{field}:Q", axis=ax, scale=alt.Scale(domain=[lo, hi], clamp=True, nice=False))
            else:
                return alt.Y(f"{field}:Q", axis=ax)
        layers = []
        for metric in selected:
            col = display_to_col.get(metric)
            if not col or col not in single.columns:
                continue
            layers.append(
                base.transform_calculate(metric=f'"{metric}"')
                    .mark_line(point=True)
                    .encode(
                        y=y_enc(metric, col),
                        color=color_enc,
                        tooltip=tooltip_for(metric),
                    )
            )
        shared_scale = single_sel
        if single_sel:
            metric = selected[0]
            col = display_to_col[metric]
            st.caption("Click **Show 3-month forecast** to project the selected series.")  # NEW
            if st.button("Show 3-month forecast"):
                df = single[["period_date", col]].dropna().sort_values("period_date").copy()
                if len(df) >= 3:
                    freq = pd.infer_freq(df["period_date"]) or "W"
                    last_date = df["period_date"].max()
                    end_date = last_date + pd.DateOffset(months=3)
                    future_index = pd.date_range(
                        start=last_date + pd.tseries.frequencies.to_offset(freq),
                        end=end_date, freq=freq
                    )
                    y = df[col].astype(float).values
                    alpha = float(np.clip(2.0 / (len(y) + 1), 0.2, 0.8))
                    beta  = alpha / 2.0
                    l, b = y[0], y[1] - y[0]
                    for t in range(1, len(y)):
                        prev_l = l
                        l = alpha * y[t] + (1 - alpha) * (l + b)
                        b = beta  * (l - prev_l) + (1 - beta)  * b
                    steps = np.arange(1, len(future_index) + 1)
                    ypred = l + steps * b
                    preds_in, lvl, tr = [], y[0], y[1] - y[0]
                    for t in range(1, len(y)):
                        preds_in.append(lvl + tr)
                        prev_lvl = lvl
                        lvl = alpha * y[t] + (1 - alpha) * (lvl + tr)
                        tr = beta  * (lvl - prev_lvl) + (1 - beta)  * tr
                    resid = y[1:] - np.array(preds_in)
                    resid_sd = float(np.std(resid, ddof=1)) if len(resid) > 2 else 0.0
                    lower = ypred - 1.96 * resid_sd
                    upper = ypred + 1.96 * resid_sd
                    if metric == "Open Complaint Timeliness":
                        ypred = np.clip(ypred, 0.0, 1.0)
                        lower = np.clip(lower, 0.0, 1.0)
                        upper = np.clip(upper, 0.0, 1.0)
                    forecast_df = pd.DataFrame({
                        "period_date": future_index,
                        col: ypred,
                        "lower": lower,
                        "upper": upper,
                        "metric": metric,
                    })
                    band = alt.Chart(forecast_df).mark_area(opacity=0.15).encode(
                        x=alt.X("period_date:T", title="Week"),
                        y=y_enc(metric, "lower"),
                        y2="upper:Q",
                        color=alt.Color("metric:N", legend=None),
                    )
                    f_line = alt.Chart(forecast_df).mark_line(point=True, strokeDash=[5, 5]).encode(
                        x="period_date:T",
                        y=y_enc(metric, col),
                        color=color_enc,
                        tooltip=tooltip_for(metric),
                    )
                    layers.extend([band, f_line])
                else:
                    st.info("Not enough historical points to forecast. Need at least 3.")
        if layers:
            if shared_scale:
                combo = alt.layer(*layers).properties(height=320)
            else:
                combo = alt.layer(*layers).resolve_scale(y="independent").properties(height=320)
            st.altair_chart(combo, use_container_width=True)
        else:
            st.info("Select at least one series to display.")
    else:
        st.info("Select at least one series to display.")
        layers = []
        def side(i: int) -> str:
            return "left" if (i % 2 == 0) else "right"
        i = 0
        if "HC in WIP" in selected and "HC in WIP" in single.columns:
            layers.append(
                base.mark_line(point=True).encode(
                    y=alt.Y("HC in WIP:Q", axis=alt.Axis(title=None, labels=False)),
                    color=alt.value("steelblue"),
                    tooltip=["period_date:T", alt.Tooltip("HC in WIP:Q", format=",.0f")]
                )
            )
            i += 1
        if "Open Complaint Timeliness" in selected and "Open Complaint Timeliness" in single.columns:
            layers.append(
                base.mark_line(point=True).encode(
                    y=alt.Y("Open Complaint Timeliness:Q", axis=alt.Axis(title=None, labels=False)),
                    color=alt.value("orange"),
                    tooltip=["period_date:T", alt.Tooltip("Open Complaint Timeliness:Q", format=".0%")]
                )
            )
            i += 1
        if "Actual UPLH" in selected and "Actual UPLH" in single.columns:
            layers.append(
                base.mark_line(point=True).encode(
                    y=alt.Y("Actual UPLH:Q", axis=alt.Axis(title=None, labels=False)),
                    color=alt.value("green"),
                    tooltip=["period_date:T", alt.Tooltip("Actual UPLH:Q", format=".2f")]
                )
            )
            i += 1
        if "Actual Output" in selected and "Actual Output" in single.columns:
            layers.append(
                base.mark_line(point=True).encode(
                    y=alt.Y("Actual Output:Q", axis=alt.Axis(title=None, labels=False)),
                    color=alt.value("red"),
                    tooltip=["period_date:T", alt.Tooltip("Actual Output:Q", format=",.0f")]
                )
            )
            i += 1
        if "Actual Hours" in selected and "Completed Hours" in single.columns:
            layers.append(
                base.mark_line(point=True).encode(
                    y=alt.Y("Completed Hours:Q", axis=alt.Axis(title=None, labels=False)),
                    color=alt.value("purple"),
                    tooltip=["period_date:T", alt.Tooltip("Completed Hours:Q", format=",.0f")]
                )
            )
            i += 1
        if "Actual HC used" in selected and "Actual HC used" in single.columns:
            layers.append(
                base.mark_line(point=True).encode(
                    y=alt.Y("Actual HC used:Q", axis=alt.Axis(title=None, labels=False)),
                    color=alt.value("indianred"),
                    tooltip=["period_date:T", alt.Tooltip("Actual HC used:Q", format=",.2f")]
                )
            )
            i += 1
        combo = alt.layer(*layers).resolve_scale(y="independent").properties(height=320)
        st.altair_chart(combo, use_container_width=True)
st.markdown("---")
st.subheader("Detailed Rows")
_nw = load_non_wip()
if not _nw.empty:
    needed = ["team", "period_date", "total_non_wip_hours"]
    for c in needed:
        if c not in _nw.columns:
            _nw[c] = np.nan
    _nw["total_non_wip_hours"] = pd.to_numeric(_nw["total_non_wip_hours"], errors="coerce")
    f_for_table = f.merge(_nw[needed], on=["team", "period_date"], how="left")
else:
    f_for_table = f.copy()
    f_for_table["total_non_wip_hours"] = np.nan
if {"Closures", "Completed Hours"}.issubset(f_for_table.columns):
    denom_prod = (
        f_for_table["Completed Hours"].fillna(0).astype(float)
        + f_for_table["total_non_wip_hours"].fillna(0).astype(float)
    )
    num_close = pd.to_numeric(f_for_table["Closures"], errors="coerce")
    f_for_table["Productivity"] = np.where(denom_prod > 0, num_close / denom_prod, np.nan)
    denom_eff = f_for_table["Completed Hours"].astype(float)
    f_for_table["Efficiency"] = np.where(denom_eff > 0, num_close / denom_eff, np.nan)
else:
    f_for_table["Productivity"] = np.nan
    f_for_table["Efficiency"] = np.nan
hide_cols = {"source_file", "fallback_used", "error", "Person Hours", "UPLH WP1", "UPLH WP2", "People in WIP", "Cell/Station Hours", "Outputs by Cell/Station", "Outputs by Person", "total_non_wip_hours", "Hours by Cell/Station - by person", "Output by Cell/Station - by person", "UPLH by Cell/Station - by person"}
drop_these = [c for c in f_for_table.columns if c in hide_cols or c.startswith("Unnamed:")]
f_table = (
    f_for_table.drop(columns=drop_these, errors="ignore")
    .sort_values(["team", "period_date"], ascending=[True, False])
)
fmt_map: dict[str, str] = {}
if "Open Complaint Timeliness" in f_table.columns:
    fmt_map["Open Complaint Timeliness"] = "{:.0%}"
if "Capacity Utilization" in f_table.columns:
    fmt_map["Capacity Utilization"] = "{:.2%}"
for col in ["Total Available Hours", "Completed Hours", "Target Output", "Actual Output"]:
    if col in f_table.columns:
        fmt_map[col] = "{:,.1f}"
for col in ["Target UPLH", "Actual UPLH", "Actual HC used", "Efficiency vs Target"]:
    if col in f_table.columns:
        fmt_map[col] = "{:,.2f}"
if "HC in WIP" in f_table.columns:
    fmt_map["HC in WIP"] = "{:,.0f}"
if "Closures" in f_table.columns:
    fmt_map["Closures"] = "{:,.0f}"
if "Productivity" in f_table.columns:
    fmt_map["Productivity"] = "{:.4f}"
f_table_display = f_table.rename(columns={"team": "Team", "period_date": "Date"})
if "Date" in f_table_display.columns:
    f_table_display["Date"] = pd.to_datetime(f_table_display["Date"], errors="coerce").dt.date
st.dataframe(
    f_table_display.style.format(fmt_map),  # keep your number formatting
    use_container_width=True,
    hide_index=True                          # hides the left index column
)