# heijunka-dash.py
import os
from pathlib import Path
import pandas as pd
import numpy as np
import streamlit as st
import altair as alt
import json
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
                "Target UPLH", "Actual UPLH", "HC in WIP", "Actual HC used"]:
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
def build_uplh_by_person_long(frame: pd.DataFrame, team: str) -> pd.DataFrame:
    outp = explode_outputs_json(frame[frame["team"] == team], "Outputs by Person", "person")
    if outp.empty:
        return pd.DataFrame(columns=["team", "period_date", "person", "Actual", "Actual Hours", "UPLH"])
    hrs = explode_person_hours(frame[frame["team"] == team])[["period_date", "person", "Actual Hours"]]
    m = (outp
         .merge(hrs, on=["period_date", "person"], how="left")
         .rename(columns={"Actual": "Actual Output"}))
    m["UPLH"] = (m["Actual Output"] / m["Actual Hours"]).replace([np.inf, -np.inf], np.nan)
    m["team"] = team
    return m[["team", "period_date", "person", "Actual Output", "Actual Hours", "UPLH"]].dropna(subset=["UPLH"])
def build_uplh_by_cell_long(frame: pd.DataFrame, team: str) -> pd.DataFrame:
    outc = explode_outputs_json(frame[frame["team"] == team], "Outputs by Cell/Station", "cell_station")
    if outc.empty:
        return pd.DataFrame(columns=["team", "period_date", "cell_station", "Actual", "Actual Hours", "UPLH"])
    hc = explode_cell_station_hours(frame[frame["team"] == team])[["period_date", "cell_station", "Actual Hours"]]
    m = (outc
         .merge(hc, on=["period_date", "cell_station"], how="left")
         .rename(columns={"Actual": "Actual Output"}))
    m["UPLH"] = (m["Actual Output"] / m["Actual Hours"]).replace([np.inf, -np.inf], np.nan)
    m["team"] = team
    return m[["team", "period_date", "cell_station", "Actual Output", "Actual Hours", "UPLH"]].dropna(subset=["UPLH"])
data_path = None if DATA_URL else str(DEFAULT_DATA_PATH)
mtime_key = 0
if data_path:
    p = Path(data_path)
    mtime_key = p.stat().st_mtime if p.exists() else 0
df = load_data(data_path, DATA_URL)
st.markdown("<h1 style='text-align: center;'>Heijunka Metrics Dashboard</h1>", unsafe_allow_html=True)
with st.expander("Glossary", expanded=False):
    st.markdown("""
- **Target UPLH** — Target Output ÷ Target Hours (i.e., **Total Available Hours**).
- **Actual UPLH** — Actual Output ÷ Actual Hours (i.e., **Completed Hours**).
- **HC in WIP** — Number of **unique people** who logged any time in WIP during the week.
- **Actual HC used** — Total actual hours worked ÷ **32.5**  
  <small>(assumes **6.5 hours in WIP per person per day × 5 days**)</small>
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
col1, col2, col3 = st.columns([2, 2, 6], gap="large")
with col1:
    selected_teams = st.multiselect("Teams", teams, key="teams_sel")
with col2:
    has_dates = df["period_date"].notna().any()
    min_date = pd.to_datetime(df["period_date"].min()).date() if has_dates else None
    max_date = pd.to_datetime(df["period_date"].max()).date() if has_dates else None
    if min_date and max_date:
        date_col1, date_col2 = st.columns(2)
        with date_col1:
            start = st.date_input("Start", value=min_date, min_value=min_date, max_value=max_date, key="start_date")
        with date_col2:
            end = st.date_input("End", value=max_date, min_value=min_date, max_value=max_date, key="end_date")
        if start > end:
            st.error("Start date cannot be after end date!")
            start, end = None, None
    else:
        start, end = None, None
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
            .tail(1))
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
tot_target = latest["Target Output"].sum(skipna=True)
tot_actual = latest["Actual Output"].sum(skipna=True)
tot_tahl  = latest["Total Available Hours"].sum(skipna=True)
tot_chl   = latest["Completed Hours"].sum(skipna=True)
tot_hc_wip = latest["HC in WIP"].sum(skipna=True) if "HC in WIP" in latest.columns else np.nan
tot_hc_used = latest["Actual HC used"].sum(skipna=True) if "Actual HC used" in latest.columns else np.nan
target_uplh = (tot_target / tot_tahl) if tot_tahl else np.nan
actual_uplh = (tot_actual / tot_chl)  if tot_chl else np.nan
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
with kpi_cols[0]:
    st.subheader("Latest (Selected Teams)")
kpi(kpi_cols[1], "Target Output", tot_target, "{:,.0f}")
kpi_vs_target(kpi_cols[2], "Actual Output", tot_actual, tot_target, "{:,.0f}")
kpi(kpi_cols[3], "Actual vs Target", (tot_actual/tot_target if tot_target else np.nan), "{:.2f}x")
kpi_cols2 = st.columns(4)
kpi(kpi_cols2[1],
    "Target UPLH",
    (tot_target/tot_tahl if tot_tahl else np.nan),
    "{:.2f}",
    help="Target Output ÷ Target Hours (Total Available Hours)"  # NEW
)
kpi_vs_target(
    kpi_cols2[2],
    "Actual UPLH",
    actual_uplh,
    target_uplh,
    "{:.2f}",
    help="Actual Output ÷ Actual Hours (Completed Hours)"  # NEW
)
kpi(kpi_cols2[3],
    "Capacity Utilization",
    (tot_chl/tot_tahl if tot_tahl else np.nan),
    "{:.0%}",
    help="Completed vs Available hours"
)
kpi_cols3 = st.columns(4)
kpi(
    kpi_cols3[1],
    "HC in WIP",
    tot_hc_wip,
    "{:,.0f}",
    help="Unique people with any time in WIP for the week"  # NEW
)
kpi(kpi_cols3[2],
    "Actual HC used",
    tot_hc_used,
    "{:,.2f}",
    help="Based on 6.5 hours per person in WIP per day"
)
kpi_vs_target(kpi_cols3[3], "Open Complaint Timeliness", timeliness_avg, 0.87, "{:.0%}")
st.markdown("---")
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
    st.subheader("Hours Trend")
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
        picked_week = None
        team_name = None
        team_people = None
        all_weeks = []
        if single_team and 'ppl_hours' in locals() and not ppl_hours.empty:
            team_name = teams_in_view[0]
            team_people = ppl_hours.loc[ppl_hours["team"] == team_name].copy()
            if not team_people.empty:
                all_weeks = sorted(pd.to_datetime(team_people["period_date"].dropna().unique()), reverse=True)
        if single_team and all_weeks:
            picked_week = st.selectbox(
                f"Pick a week for {team_name} drilldown:",
                options=all_weeks,
                index=0,
                format_func=lambda d: pd.to_datetime(d).date().isoformat(),
                key="per_person_week_select",
            )
            picked_week = pd.to_datetime(picked_week) if picked_week is not None else None
            layers = [line, pts]
            if picked_week is not None:
                rule_df = pd.DataFrame({"period_date": [picked_week]})
                rule = alt.Chart(rule_df).mark_rule(strokeDash=[4, 3]).encode(x="period_date:T")
                layers.append(rule)
            chart_ph.altair_chart(
                alt.layer(*layers).properties(height=280).add_params(team_sel),
                use_container_width=True
            )
        if single_team and (team_people is None or team_people.empty):
            st.caption(f"No per-person data available for {teams_in_view[0]}.")
        elif single_team and picked_week is not None and team_people is not None:
            wk_people = team_people.loc[team_people["period_date"] == picked_week]
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
                        .properties(height=260, title=f"{team_name} • Per-person Hours (labels show over/under vs available)")
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
        else:
            if not single_team:
                st.caption("Per-person drilldown is available when exactly one team is selected.")
            elif 'ppl_hours' in locals() and ppl_hours.empty:
                st.caption("Per-person drilldown not available (no 'Person Hours' found).")
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
                            .properties(height=260)
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
            y=alt.Y("Value:Q", title="UPLH", scale=y_scale),
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
            .properties(height=230)
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
        team_weeks = sorted(pd.to_datetime(f.loc[f["team"] == team_for_drill, "period_date"].dropna().unique()), reverse=True)
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
            top_ph.altair_chart(alt.layer(line, pts, rule_week).properties(height=280).add_params(team_sel), use_container_width=True)
        else:
            picked_week = None
            st.info("No weeks available for drilldown.")
        lower = None
        if picked_week is not None:
            if by_choice == "Person":
                uplh_person = build_uplh_by_person_long(f, team_for_drill)
                if uplh_person.empty:
                    st.info("No 'Outputs by Person' and/or 'Person Hours' data to compute UPLH.")
                else:
                    wk = uplh_person.loc[uplh_person["period_date"] == picked_week].copy()
                    if wk.empty:
                        st.info("No UPLH-by-person records for that week.")
                    else:
                        lower = (
                            alt.Chart(wk)
                            .mark_bar()
                            .encode(
                                x=alt.X("person:N", title="Person", sort="-y"),
                                y=alt.Y("UPLH:Q", title="UPLH"),
                                tooltip=[
                                    "period_date:T",
                                    "person:N",
                                    alt.Tooltip("Actual Output:Q", title="Actual Output", format=",.0f"),
                                    alt.Tooltip("Actual Hours:Q", title="Actual Hours", format=",.1f"),
                                    alt.Tooltip("UPLH:Q", title="UPLH", format=",.2f"),
                                ],
                            )
                            .properties(height=230)
                        )
            else:
                uplh_cell = build_uplh_by_cell_long(f, team_for_drill)
                if uplh_cell.empty:
                    st.info("No 'Outputs by Cell/Station' and/or 'Cell/Station Hours' data to compute UPLH.")
                else:
                    wk = uplh_cell.loc[uplh_cell["period_date"] == picked_week].copy()
                    if wk.empty:
                        st.info("No UPLH-by-cell/station records for that week.")
                    else:
                        lower = (
                            alt.Chart(wk)
                            .mark_bar()
                            .encode(
                                x=alt.X("cell_station:N", title="Cell/Station", sort="-y"),
                                y=alt.Y("UPLH:Q", title="UPLH"),
                                tooltip=[
                                    "period_date:T",
                                    alt.Tooltip("cell_station:N", title="Cell/Station"),
                                    alt.Tooltip("Actual Output:Q", title="Actual Output", format=",.0f"),
                                    alt.Tooltip("Actual Hours:Q", title="Actual Hours", format=",.1f"),
                                    alt.Tooltip("UPLH:Q", title="UPLH", format=",.2f"),
                                ],
                            )
                            .properties(height=230)
                        )
        if lower is not None:
            st.altair_chart(lower, use_container_width=True)
        else:
            st.altair_chart(top, use_container_width=True)
            if (not multi_team) and not (wp1_col and wp2_col) and team_for_drill == "PH":
                st.info("No WP1/WP2 UPLH columns found. Expected columns like 'WP1 UPLH' and 'WP2 UPLH'.")
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
            base_hc.mark_line(point=True).properties(height=260),
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
            base_ahu.mark_line(point=True).properties(height=260),
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
                    f"Pick a week for {team_name} per-person avg daily hours:",
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
                            wk_people["Avg Daily Hours"] >= 6.5, "≥ 6.5 (Over)", "< 6.5 (Under)"
                        )
                        wk_people["Delta"] = wk_people["Avg Daily Hours"] - 6.5
                        wk_people["DeltaLabel"] = wk_people["Delta"].map(lambda x: f"{x:+.2f}")
                        vmax = float(pd.to_numeric(wk_people["Avg Daily Hours"], errors="coerce").max())
                        pad  = max(0.3, (max(vmax, 6.5) * 0.12))  # a little headroom
                        y_lo = 0.0
                        y_hi = max(vmax, 6.5) + pad
                        y_scale = alt.Scale(domain=[y_lo, y_hi], nice=False, clamp=False)
                        order_people = (
                            wk_people.sort_values("Avg Daily Hours", ascending=False)["person"].tolist()
                        )
                        color_enc = alt.Color(
                            "OverUnder:N",
                            title="vs 6.5",
                            scale=alt.Scale(
                                domain=["≥ 6.5 (Over)", "< 6.5 (Under)"],
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
                                    alt.Tooltip("Delta:Q", title="Over/Under vs 6.5", format="+.2f"),
                                ],
                            )
                            .properties(height=240, title=f"{team_name} • Per-person Avg Daily Hours")
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
                                        domain=["≥ 6.5 (Over)", "< 6.5 (Under)"],
                                        range=["#22c55e", "#ef4444"]
                                    ),
                                ),
                            )
                        )
                        ref = alt.Chart(pd.DataFrame({"y": [6.5]})).mark_rule(strokeDash=[4, 3]).encode(y=alt.Y("y:Q", scale=y_scale))

                        st.altair_chart(bars + labels + ref, use_container_width=True)
        else:
            st.caption("Select exactly one team to drill into per-person daily hours.")
    else:
        st.info("No 'Actual HC used' data available in the selected range.")
with right2:
    st.subheader("Open Complaint Timeliness Trend")
    if "Open Complaint Timeliness" in f.columns and f["Open Complaint Timeliness"].notna().any():
        tml = f[["team", "period_date", "Open Complaint Timeliness"]].dropna().copy()
        max_val = tml["Open Complaint Timeliness"].max()
        divisor = 100.0 if pd.notna(max_val) and float(max_val) > 1.5 else 1.0
        tml["Timeliness %"] = tml["Open Complaint Timeliness"].astype(float) / divisor
        vmin = float(tml["Timeliness %"].min())
        vmax = float(tml["Timeliness %"].max())
        rng  = max(0.0, vmax - vmin)
        pad  = max(0.02, rng * 0.15)
        lo   = max(0.0, vmin - pad)
        hi   = min(1.0, vmax + pad)
        base_tml = alt.Chart(tml).encode(
            x=alt.X("period_date:T", title="Week"),
            y=alt.Y("Timeliness %:Q",
                    title="Timeliness",
                    axis=alt.Axis(format="%"),
                    scale=alt.Scale(domain=[lo, hi], clamp=True, nice=False)),
            color=alt.Color("team:N", title="Team") if len(teams_in_view) > 1 else alt.value("seagreen"),
            tooltip=["team:N", "period_date:T", alt.Tooltip("Timeliness %:Q", format=".0%")]
        )
        st.altair_chart(
            base_tml.mark_line(point=True).properties(height=260),
            use_container_width=True
        )
    else:
        st.info("No 'Open Complaint Timeliness' data available in the selected range.")
if len(teams_in_view) == 1:
    team_name = teams_in_view[0]
    st.subheader(f"{team_name} • Multi-Axis View")
    single = (
        f[f["team"] == team_name]
        .sort_values("period_date")
        .copy()
    )
    metric_options = [
        "HC in WIP",
        "Open Complaint Timeliness",
        "Actual UPLH",
        "Actual Output",
        "Actual Hours",
        "Actual HC used"
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
        st.caption("💡 Select **one** series to enable the 3-month projection.")
    if selected:
        display_to_col = {
            "HC in WIP": "HC in WIP",
            "Open Complaint Timeliness": "Open Complaint Timeliness",
            "Actual UPLH": "Actual UPLH",
            "Actual Output": "Actual Output",
            "Actual Hours": "Completed Hours",
            "Actual HC used": "Actual HC used",
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
            st.caption("🔮 Click **Show 3-month forecast** to project the selected series.")  # NEW
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
st.subheader("Efficiency vs Target (Actual / Target)")
eff = f.assign(Efficiency=lambda d: (d["Actual Output"] / d["Target Output"]))
eff = eff.replace([np.inf, -np.inf], np.nan).dropna(subset=["Efficiency"])
color_scale = alt.Scale(
    domain=[0, 1],
    range=["#d62728", "#2ca02c"]
)
eff_bar = (
    alt.Chart(eff)
    .mark_bar()
    .encode(
        x=alt.X("period_date:T", title="Week"),
        y=alt.Y("Efficiency:Q", title="x of Target"),
        color=alt.condition("datum.Efficiency >= 1", alt.value("#2ca02c"), alt.value("#d62728")),
        tooltip=[
            "team:N",
            "period_date:T",
            alt.Tooltip("Actual Output:Q", format=",.0f"),
            alt.Tooltip("Target Output:Q", format=",.0f"),
            alt.Tooltip("Efficiency:Q", format=".2f")
        ]
    )
)
ref_line = alt.Chart(pd.DataFrame({"y": [1.0]})).mark_rule(strokeDash=[4,3]).encode(y="y:Q")
st.altair_chart((eff_bar + ref_line).properties(height=260), use_container_width=True)
st.markdown("---")
st.subheader("Detailed Rows")
hide_cols = {"source_file", "fallback_used", "error", "Person Hours", "UPLH WP1", "UPLH WP2", "People in WIP", "Cell/Station Hours", "Outputs by Cell/Station", "Outputs by Person"}
drop_these = [c for c in f.columns if c in hide_cols or c.startswith("Unnamed:")]
f_table = (
    f.drop(columns=drop_these, errors="ignore")
    .sort_values(["team", "period_date"], ascending=[True, False])
)
st.dataframe(f_table, use_container_width=True)
