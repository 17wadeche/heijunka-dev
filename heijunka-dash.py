# Heijunka Dashboard — modernized
# A refreshed, modular, and playful take on your excellent Streamlit app.
# Key upgrades:
# - Sidebar filters + tabs for cleaner layout
# - URL permalinks for teams & dates (deep-link friendly)
# - KPI deltas vs previous period + celebratory balloons when beating target 🎈
# - Altair custom theme for a modern look (rounded, softer grid, better fonts)
# - Reusable helpers & tighter null handling
# - Export filtered data as CSV
# - Retains all of your drilldowns (PH people hours, WP1/WP2, etc.)

import os
from pathlib import Path
from typing import Iterable, List, Tuple
import pandas as pd
import numpy as np
import streamlit as st
import altair as alt
import json
import re

# ----------------------
# Config & Constants
# ----------------------
DEFAULT_DATA_PATH = Path(r"C:\heijunka-dev\metrics_aggregate_dev.xlsx")
DATA_URL = st.secrets.get("HEIJUNKA_DATA_URL", os.environ.get("HEIJUNKA_DATA_URL"))

st.set_page_config(page_title="Heijunka Metrics", layout="wide")

# Hide default Streamlit UI chrome for a cleaner vibe
st.markdown(
    """
    <style>
      [data-testid="stToolbar"] { display: none; }
      #MainMenu, header, footer { visibility: hidden; }
      .kpi-card { padding: 10px 14px; border-radius: 14px; background: rgba(0,0,0,0.04); }
      .metric-subtitle { opacity: 0.65; font-size: 0.85rem; margin-top: -6px; }
      .small-help { opacity: 0.6; font-size: 0.85rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

# Auto-refresh hourly if available (keeps wallboards fresh)
if hasattr(st, "autorefresh"):
    st.autorefresh(interval=60 * 60 * 1000, key="auto-refresh")

# ----------------------
# Altair Theme (modern)
# ----------------------

def _altair_theme():
    palette = [
        "#6C5CE7", "#00B894", "#0984E3", "#E17055", "#E84393",
        "#2D3436", "#55EFC4", "#FAB1A0", "#81ECEC", "#A29BFE",
    ]
    return {
        "config": {
            "view": {"strokeWidth": 0},
            "axis": {
                "domain": False,
                "labelColor": "#6b7280",
                "gridColor": "#e5e7eb",
                "titleColor": "#374151",
                "labelFontSize": 12,
                "titleFontSize": 12,
            },
            "legend": {"labelColor": "#374151", "titleColor": "#111827"},
            "range": {"category": palette},
            "point": {"filled": True, "size": 70},
            "line": {"strokeWidth": 2},
        }
    }

alt.themes.register("heijunka", _altair_theme)
alt.themes.enable("heijunka")

# ----------------------
# Data loading & cleaning
# ----------------------

@st.cache_data(show_spinner=False, ttl=15 * 60)
def load_data(data_path: str | None, data_url: str | None) -> pd.DataFrame:
    """Load CSV/JSON/Excel; apply postprocessing; return DataFrame (may be empty)."""
    if data_url:
        if data_url.lower().endswith(".json"):
            df = pd.read_json(data_url)
        else:
            df = pd.read_csv(data_url)
        return _postprocess(df)

    if not data_path:
        return pd.DataFrame()

    p = Path(data_path)
    if not p.exists():
        return pd.DataFrame()

    if p.suffix.lower() in (".xlsx", ".xlsm"):
        df = pd.read_excel(p, sheet_name="All Metrics")
    elif p.suffix.lower() == ".csv":
        df = pd.read_csv(p)
    elif p.suffix.lower() == ".json":
        df = pd.read_json(p)
    else:
        return pd.DataFrame()

    return _postprocess(df)


def _postprocess(df: pd.DataFrame) -> pd.DataFrame:
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
        s = (
            df["Open Complaint Timeliness"].astype(str).str.strip().replace({"": np.nan, "—": np.nan, "-": np.nan})
        )
        s = s.str.replace("%", "", regex=False).str.replace(",", "", regex=False)
        v = pd.to_numeric(s, errors="coerce")
        if pd.notna(v.max()) and float(v.max()) > 1.5:
            v = v / 100.0
        df["Open Complaint Timeliness"] = v

    for col in [
        "Total Available Hours",
        "Completed Hours",
        "Target Output",
        "Actual Output",
        "Target UPLH",
        "Actual UPLH",
        "HC in WIP",
        "Actual HC used",
    ]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    if {"Actual Output", "Target Output"}.issubset(df.columns):
        df["Efficiency vs Target"] = (df["Actual Output"] / df["Target Output"]).replace([np.inf, -np.inf], np.nan)
    if {"Completed Hours", "Total Available Hours"}.issubset(df.columns):
        df["Capacity Utilization"] = (df["Completed Hours"] / df["Total Available Hours"]).replace([np.inf, -np.inf], np.nan)

    return df

# ----------------------
# People helpers (PH drilldowns, WIP explode)
# ----------------------

def explode_people_in_wip(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "People in WIP" not in df.columns:
        return pd.DataFrame(columns=["team", "period_date", "person"])
    sub = df.loc[:, ["team", "period_date", "People in WIP"]].dropna(subset=["People in WIP"]).copy()
    rows: list[dict] = []

    def _as_names(x) -> list[str]:
        if isinstance(x, list):
            return [str(s).strip() for s in x if str(s).strip()]
        if isinstance(x, str):
            s = x.strip()
            try:
                obj = json.loads(s)
                if isinstance(obj, list):
                    return [str(v).strip() for v in obj if str(v).strip()]
                if isinstance(obj, dict):
                    return [str(k).strip() for k, v in obj.items() if str(k).strip()]
            except Exception:
                pass
            parts = [p.strip() for p in re.split(r"[,;\n\r]+", s) if p.strip()]
            return parts
        if isinstance(x, dict):
            return [str(k).strip() for k in x.keys() if str(k).strip()]
        return []

    for _, r in sub.iterrows():
        people = _as_names(r["People in WIP"])
        for person in people:
            rows.append(
                {
                    "team": r["team"],
                    "period_date": pd.to_datetime(r["period_date"], errors="coerce").normalize(),
                    "person": person,
                }
            )
    out = pd.DataFrame(rows)
    if not out.empty:
        out["period_date"] = pd.to_datetime(out["period_date"], errors="coerce").dt.normalize()
    return out


def explode_ph_person_hours(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "PH Person Hours" not in df.columns:
        return pd.DataFrame(columns=["team", "period_date", "person", "Actual Hours", "Available Hours", "Utilization"])
    sub = df.loc[df["team"] == "PH", ["team", "period_date", "PH Person Hours"]].dropna(subset=["PH Person Hours"]).copy()
    rows: list[dict] = []
    for _, r in sub.iterrows():
        payload = r["PH Person Hours"]
        try:
            obj = json.loads(payload) if isinstance(payload, str) else payload
            if not isinstance(obj, dict):
                continue
        except Exception:
            continue
        for person, vals in obj.items():
            a = pd.to_numeric((vals or {}).get("actual"), errors="coerce")
            t = pd.to_numeric((vals or {}).get("available"), errors="coerce")
            util = (a / t) if (pd.notna(a) and pd.notna(t) and t not in (0, 0.0)) else np.nan
            rows.append(
                {
                    "team": r["team"],
                    "period_date": pd.to_datetime(r["period_date"], errors="coerce"),
                    "person": str(person),
                    "Actual Hours": a,
                    "Available Hours": t,
                    "Utilization": util,
                }
            )
    out = pd.DataFrame(rows)
    if not out.empty:
        out["period_date"] = pd.to_datetime(out["period_date"], errors="coerce").dt.normalize()
    return out


def ahu_person_share_for_week(frame: pd.DataFrame, week, teams_in_view: List[str], ph_people_df: pd.DataFrame) -> pd.DataFrame:
    if frame.empty or "Actual HC used" not in frame.columns:
        return pd.DataFrame(columns=["team", "period_date", "person", "percent"])
    wk = pd.to_datetime(week, errors="coerce").normalize()
    if pd.isna(wk):
        return pd.DataFrame(columns=["team", "period_date", "person", "percent"])
    ppl = explode_people_in_wip(frame)
    out_rows: list[dict] = []
    for team in teams_in_view:
        team_ahu_series = frame.loc[(frame["team"] == team) & (frame["period_date"] == wk), "Actual HC used"].dropna()
        if team_ahu_series.empty:
            continue
        per_df = None
        if team == "PH" and ph_people_df is not None and not ph_people_df.empty:
            phw = ph_people_df.loc[(ph_people_df["team"] == "PH") & (ph_people_df["period_date"] == wk)]
            if not phw.empty and phw["Actual Hours"].notna().any():
                g = phw.groupby("person", as_index=False)["Actual Hours"].sum()
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

# ----------------------
# URL param helpers
# ----------------------

def _get_qp_list(key: str) -> List[str]:
    try:
        qp = st.query_params
        vals = qp.get_all(key) if hasattr(qp, "get_all") else qp.get(key, [])
    except Exception:
        qp = st.experimental_get_query_params()  # fallback for older Streamlit
        vals = qp.get(key, [])
    if vals is None:
        return []
    if isinstance(vals, str):
        return [vals]
    return [str(v) for v in vals]


def _set_qp_values(**kwargs) -> None:
    # Works on modern Streamlit; falls back gracefully
    try:
        for k, v in kwargs.items():
            st.query_params[k] = v
    except Exception:
        st.experimental_set_query_params(**kwargs)


# ----------------------
# Load data
# ----------------------

data_path = None if DATA_URL else str(DEFAULT_DATA_PATH)
_df = load_data(data_path, DATA_URL)

st.markdown("<h1 style='text-align:center;'>Heijunka Metrics Dashboard</h1>", unsafe_allow_html=True)

if _df.empty:
    st.warning("No data found yet. Ensure the source exists and has the 'All Metrics' sheet.")
    st.stop()

# ----------------------
# Sidebar Filters
# ----------------------
with st.sidebar:
    st.subheader("Filters")

    teams = sorted([t for t in _df["team"].dropna().unique()])
    default_teams = [teams[0]] if teams else []

    # Persist teams via URL params
    if "teams_sel" not in st.session_state:
        saved = [t for t in teams if t in _get_qp_list("teams")]
        st.session_state.teams_sel = saved or default_teams

    selected_teams = st.multiselect("Teams", teams, key="teams_sel")

    # Date range (also persisted)
    has_dates = _df["period_date"].notna().any()
    min_date = pd.to_datetime(_df["period_date"].min()).date() if has_dates else None
    max_date = pd.to_datetime(_df["period_date"].max()).date() if has_dates else None

    if min_date and max_date:
        colA, colB = st.columns(2)
        with colA:
            start = st.date_input("Start", value=min_date, min_value=min_date, max_value=max_date, key="start_date")
        with colB:
            end = st.date_input("End", value=max_date, min_value=min_date, max_value=max_date, key="end_date")
        if start > end:
            st.error("Start date cannot be after end date!")
            start, end = None, None
    else:
        start, end = None, None

    # Update URL params when filters change
    qp = {"teams": sorted(st.session_state.teams_sel)}
    if start and end:
        qp.update({"start": str(start), "end": str(end)})
    _set_qp_values(**qp)

    # Quick actions
    st.markdown("---")
    st.caption("Share this view")
    st.code(str(st.query_params), language="python")
    # (Download button moved below after filters are applied to ensure correct data & type)

# Apply filters
f = _df.copy()
if st.session_state.teams_sel:
    f = f[f["team"].isin(st.session_state.teams_sel)]
if start and end:
    f = f[(f["period_date"] >= pd.to_datetime(start)) & (f["period_date"] <= pd.to_datetime(end))]

if f.empty:
    st.info("No rows match your filters.")
    st.stop()
with st.sidebar:
    st.download_button(
        label="Download filtered CSV",
        data=f.to_csv(index=False).encode("utf-8"),
        file_name="heijunka_filtered.csv",
        mime="text/csv",
        use_container_width=True,
    )

ph_people = explode_ph_person_hours(f)
latest = (
    f.sort_values(["team", "period_date"]).groupby("team", as_index=False).tail(1)
)

# ----------------------
# KPI Row (with deltas vs previous period)
# ----------------------

kpi_cols = st.columns([2, 2, 2, 2, 2])

with kpi_cols[0]:
    st.markdown("<div class='kpi-card'><h4>Latest (Selected Teams)</h4><p class='metric-subtitle'>Aggregated across visible teams</p></div>", unsafe_allow_html=True)

# Helper to compute aggregated totals for latest and previous periods

def _aggregate_latest_and_prev(df: pd.DataFrame, cols: Iterable[str]) -> Tuple[pd.Series, pd.Series]:
    if df.empty:
        return pd.Series(dtype=float), pd.Series(dtype=float)
    # latest per team
    latest = df.sort_values(["team", "period_date"]).groupby("team", as_index=False).tail(1)
    # previous per team
    prev = (
        df.sort_values(["team", "period_date"]).groupby("team").nth(-2).reset_index()
    )
    return latest[cols].sum(numeric_only=True), prev[cols].sum(numeric_only=True)

agg_cols = [
    "Target Output",
    "Actual Output",
    "Total Available Hours",
    "Completed Hours",
    "HC in WIP",
    "Actual HC used",
]
latest_tot, prev_tot = _aggregate_latest_and_prev(f, agg_cols)

# Compute values
_tot_target = float(latest_tot.get("Target Output", np.nan))
_tot_actual = float(latest_tot.get("Actual Output", np.nan))
_tot_tahl = float(latest_tot.get("Total Available Hours", np.nan))
_tot_chl = float(latest_tot.get("Completed Hours", np.nan))
_tot_hc_wip = float(latest_tot.get("HC in WIP", np.nan))
_tot_hc_used = float(latest_tot.get("Actual HC used", np.nan))

_prev_target = float(prev_tot.get("Target Output", np.nan))
_prev_actual = float(prev_tot.get("Actual Output", np.nan))
_prev_tahl = float(prev_tot.get("Total Available Hours", np.nan))
_prev_chl = float(prev_tot.get("Completed Hours", np.nan))

_target_uplh = (_tot_target / _tot_tahl) if _tot_tahl else np.nan
_actual_uplh = (_tot_actual / _tot_chl) if _tot_chl else np.nan
_prev_target_uplh = (_prev_target / _prev_tahl) if _prev_tahl else np.nan
_prev_actual_uplh = (_prev_actual / _prev_chl) if _prev_chl else np.nan

# Celebrate if beating target overall
if _tot_target and _tot_actual and (_tot_actual >= _tot_target):
    st.balloons()


def _metric(col, label, value, prev_value=None, fmt="{:,.2f}"):
    if pd.isna(value):
        col.metric(label, "—")
        return
    try:
        value_str = fmt.format(value)
    except Exception:
        value_str = str(value)
    delta = None
    if prev_value is not None and not pd.isna(prev_value):
        try:
            delta_val = float(value) - float(prev_value)
            # Render as +x / -x with commas
            delta = ("+" if delta_val >= 0 else "") + f"{delta_val:,.2f} vs prev"
        except Exception:
            delta = None
    col.metric(label, value_str, delta=delta)


_metric(kpi_cols[1], "Target Output", _tot_target, _prev_target, "{:,.0f}")
_metric(kpi_cols[2], "Actual Output", _tot_actual, _prev_actual, "{:,.0f}")

# Actual vs Target (x)
if _tot_target:
    _metric(kpi_cols[3], "Actual vs Target", _tot_actual / _tot_target, None, "{:.2f}x")
else:
    _metric(kpi_cols[3], "Actual vs Target", np.nan)

# Utilization & UPLH row
kpi_cols2 = st.columns([2, 2, 2, 2, 2])

_metric(kpi_cols2[1], "Target UPLH", _target_uplh, _prev_target_uplh, "{:.2f}")
_metric(kpi_cols2[2], "Actual UPLH", _actual_uplh, _prev_actual_uplh, "{:.2f}")

if _tot_tahl:
    util_now = (_tot_chl / _tot_tahl) if _tot_tahl else np.nan
    util_prev = (_prev_chl / _prev_tahl) if _prev_tahl else np.nan
    _metric(kpi_cols2[3], "Capacity Utilization", util_now, util_prev, "{:.0%}")
else:
    _metric(kpi_cols2[3], "Capacity Utilization", np.nan)

# HC WIP / Actual HC used
kpi_cols3 = st.columns([2, 2, 2, 2, 2])
_metric(kpi_cols3[1], "HC in WIP", _tot_hc_wip, None, "{:,.0f}")
_metric(kpi_cols3[2], "Actual HC used", _tot_hc_used, None, "{:,.2f}")

# Timeliness KPI normalized
if "Open Complaint Timeliness" in latest.columns:
    raw_timeliness = latest["Open Complaint Timeliness"].dropna()
    timeliness_avg_raw = raw_timeliness.mean() if not raw_timeliness.empty else np.nan
    def _normalize_percent_value(v):
        if pd.isna(v):
            return np.nan
        try:
            v = float(v)
        except Exception:
            return np.nan
        return v if v <= 1.0 else v / 100.0
    timeliness_avg = _normalize_percent_value(timeliness_avg_raw)
    target_timeliness = 0.87
    delta = None
    if not pd.isna(timeliness_avg):
        delta_pct = timeliness_avg - target_timeliness
        delta = f"{delta_pct:+.0%} vs 87% target"
    kpi_cols3[3].metric("Open Complaint Timeliness", f"{timeliness_avg:.0%}" if pd.notna(timeliness_avg) else "—", delta=delta, delta_color="normal")
else:
    kpi_cols3[3].metric("Open Complaint Timeliness", "—")

st.markdown("---")

# ----------------------
# Tabs: Overview | Trends | People | Table
# ----------------------

oview, trends, people, table = st.tabs(["Overview", "Trends", "People", "Table"])

teams_in_view = sorted([t for t in f["team"].dropna().unique()])
multi_team = len(teams_in_view) > 1
team_sel = alt.selection_point(fields=["team"], bind="legend")

# -------- Overview: Efficiency vs Target bar + quick minis --------
with oview:
    st.subheader("Efficiency vs Target (Actual / Target)")
    eff = f.assign(Efficiency=lambda d: (d["Actual Output"] / d["Target Output"]))
    eff = eff.replace([np.inf, -np.inf], np.nan).dropna(subset=["Efficiency"]) 

    eff_bar = (
        alt.Chart(eff)
        .mark_bar()
        .encode(
            x=alt.X("period_date:T", title="Week"),
            y=alt.Y("Efficiency:Q", title="x of Target"),
            color=alt.condition("datum.Efficiency >= 1", alt.value("#16A34A"), alt.value("#DC2626")),
            tooltip=[
                "team:N",
                "period_date:T",
                alt.Tooltip("Actual Output:Q", format=",.0f"),
                alt.Tooltip("Target Output:Q", format=",.0f"),
                alt.Tooltip("Efficiency:Q", format=".2f"),
            ],
        )
    )
    ref_line = alt.Chart(pd.DataFrame({"y": [1.0]})).mark_rule(strokeDash=[4, 3]).encode(y="y:Q")
    st.altair_chart((eff_bar + ref_line).properties(height=260), use_container_width=True)

# -------- Trends: Hours | Output | UPLH + WP1/WP2 --------
with trends:
    left, mid, right = st.columns(3)

    with left:
        st.subheader("Hours Trend")
        have_hours = {"Total Available Hours", "Completed Hours"}.issubset(f.columns)
        if not have_hours:
            st.info("Hours columns not found (need 'Total Available Hours' and 'Completed Hours').")
        else:
            hrs_long = (
                f.melt(
                    id_vars=["team", "period_date"],
                    value_vars=["Total Available Hours", "Completed Hours"],
                    var_name="Metric",
                    value_name="Value",
                )
                .dropna(subset=["Value"]) 
                .assign(
                    Metric=lambda d: d["Metric"].replace({
                        "Total Available Hours": "Target Hours",
                        "Completed Hours": "Actual Hours",
                    })
                )
            )
            base = alt.Chart(hrs_long).encode(
                x=alt.X("period_date:T", title="Week"),
                y=alt.Y("Value:Q", title="Hours"),
                color=alt.Color("Metric:N", title="Series"),
                tooltip=["team:N", "period_date:T", "Metric:N", alt.Tooltip("Value:Q", format=",.0f")],
            )
            line = base.mark_line().encode(
                detail="team:N",
                opacity=alt.condition(team_sel, alt.value(1.0), alt.value(0.25)) if multi_team else alt.value(1.0),
            )
            pts = base.mark_point().encode(
                shape=alt.Shape("team:N", title="Team") if multi_team else alt.value("circle"),
                size=alt.value(45),
                opacity=alt.condition(team_sel, alt.value(1.0), alt.value(0.25)) if multi_team else alt.value(1.0),
            )
            st.altair_chart((line + pts).properties(height=280).add_params(team_sel), use_container_width=True)

    with mid:
        st.subheader("Output Trend")
        out_long = (
            f.melt(
                id_vars=["team", "period_date"],
                value_vars=["Target Output", "Actual Output"],
                var_name="Metric",
                value_name="Value",
            ).dropna(subset=["Value"])
        )
        base = alt.Chart(out_long).encode(
            x=alt.X("period_date:T", title="Week"),
            y=alt.Y("Value:Q", title="Output"),
            color=alt.Color("Metric:N", title="Series"),
            tooltip=["team:N", "period_date:T", "Metric:N", alt.Tooltip("Value:Q", format=",.0f")],
        )
        line = base.mark_line().encode(
            detail="team:N",
            opacity=alt.condition(team_sel, alt.value(1.0), alt.value(0.25)) if multi_team else alt.value(1.0),
        )
        pts = base.mark_point().encode(
            shape=alt.Shape("team:N", title="Team") if multi_team else alt.value("circle"),
            size=alt.value(45),
            opacity=alt.condition(team_sel, alt.value(1.0), alt.value(0.25)) if multi_team else alt.value(1.0),
        )
        st.altair_chart((line + pts).properties(height=280).add_params(team_sel), use_container_width=True)

    with right:
        st.subheader("UPLH Trend")
        have_target_uplh = "Target UPLH" in f.columns
        uplh_vars = ["Actual UPLH"] + (["Target UPLH"] if have_target_uplh else [])
        uplh_long = (
            f.melt(
                id_vars=["team", "period_date"],
                value_vars=uplh_vars,
                var_name="Metric",
                value_name="Value",
            ).dropna(subset=["Value"])
        )
        if not uplh_long.empty:
            vmin = float(pd.to_numeric(uplh_long["Value"], errors="coerce").min())
            vmax = float(pd.to_numeric(uplh_long["Value"], errors="coerce").max())
            rng = max(0.0, vmax - vmin)
            pad = max(0.2, rng * 0.15)
            lo = max(0.0, vmin - pad)
            hi = vmax + pad
            y_scale = alt.Scale(domain=[lo, hi], nice=False, clamp=False)
        else:
            y_scale = alt.Scale()

        sel_wk = alt.selection_point(name="wk_uplh", fields=["period_date"], on="click", clear="dblclick", empty="none")
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
        rule = alt.Chart(uplh_long).transform_filter(sel_wk).mark_rule(strokeDash=[4, 3]).encode(x="period_date:T")
        top = alt.layer(line, pts, rule).properties(height=280).add_params(team_sel, sel_wk)

        # Optional WP1/WP2
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
        if (not multi_team) and wp1_col and wp2_col:
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
            base_wp = alt.Chart(wp_long).transform_filter(sel_wk).transform_filter(team_sel)
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
        else:
            st.altair_chart(top, use_container_width=True)

# -------- People: HC/WIP trends + PH person shares --------
with people:
    left2, mid2, right2 = st.columns(3)

    with left2:
        st.subheader("HC in WIP Trend")
        if "HC in WIP" in f.columns and f["HC in WIP"].notna().any():
            hc = f[["team", "period_date", "HC in WIP"]].dropna()
            base_hc = alt.Chart(hc).encode(
                x=alt.X("period_date:T", title="Week"),
                y=alt.Y("HC in WIP:Q", title="HC in WIP"),
                color=alt.Color("team:N", title="Team") if len(teams_in_view) > 1 else alt.value("steelblue"),
                tooltip=["team:N", "period_date:T", alt.Tooltip("HC in WIP:Q", format=",.0f")],
            )
            st.altair_chart(base_hc.mark_line(point=True).properties(height=260), use_container_width=True)
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
                tooltip=["team:N", "period_date:T", alt.Tooltip("Actual HC used:Q", format=",.2f")],
            )
            st.altair_chart(base_ahu.mark_line(point=True).properties(height=260), use_container_width=True)

            if len(teams_in_view) == 1 and teams_in_view[0] == "PH":
                ahu_ph = ahu.loc[ahu["team"] == "PH"]
                all_weeks_ahu_ph = sorted(pd.to_datetime(ahu_ph["period_date"].dropna().unique()))
                if all_weeks_ahu_ph:
                    default_week = max(all_weeks_ahu_ph)
                    picked_ahu_week = st.selectbox(
                        "Pick a week to see each PH person's % of Actual HC used:",
                        options=all_weeks_ahu_ph,
                        index=all_weeks_ahu_ph.index(default_week) if default_week in all_weeks_ahu_ph else 0,
                        format_func=lambda d: pd.to_datetime(d).date().isoformat(),
                        key="ahu_week_select_ph",
                    )
                    comp = ahu_person_share_for_week(f, picked_ahu_week, ["PH"], ph_people)
                    if not comp.empty:
                        comp = comp.loc[(comp["percent"].astype(float) > 0) & comp["person"].notna()].copy()
                        comp = comp.sort_values(["percent", "person"], ascending=[False, True])
                    if comp.empty:
                        st.info("No PH people with a non-zero share for that week.")
                    else:
                        chart = (
                            alt.Chart(comp)
                            .mark_bar()
                            .encode(
                                x=alt.X("person:N", title="Person", sort="-y"),
                                y=alt.Y("percent:Q", title="% of Actual HC used", axis=alt.Axis(format="%")),
                                tooltip=[
                                    "person:N",
                                    alt.Tooltip("percent:Q", title="% of All HC used", format=".0%"),
                                    "period_date:T",
                                ],
                                color=alt.value("indianred"),
                            )
                            .properties(height=240)
                        )
                        st.altair_chart(chart, use_container_width=True)
                else:
                    st.info("No weeks available to drill down for PH.")
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
            rng = max(0.0, vmax - vmin)
            pad = max(0.02, rng * 0.15)
            lo = max(0.0, vmin - pad)
            hi = min(1.0, vmax + pad)
            base_tml = alt.Chart(tml).encode(
                x=alt.X("period_date:T", title="Week"),
                y=alt.Y("Timeliness %:Q", title="Timeliness", axis=alt.Axis(format="%"), scale=alt.Scale(domain=[lo, hi], clamp=True, nice=False)),
                color=alt.Color("team:N", title="Team") if len(teams_in_view) > 1 else alt.value("seagreen"),
                tooltip=["team:N", "period_date:T", alt.Tooltip("Timeliness %:Q", format=".0%")],
            )
            st.altair_chart(base_tml.mark_line(point=True).properties(height=260), use_container_width=True)
        else:
            st.info("No 'Open Complaint Timeliness' data available in the selected range.")

# -------- Table: detailed rows (with nicer defaults) --------
with table:
    st.subheader("Detailed Rows")
    hide_cols = {"source_file", "fallback_used", "error", "PH Person Hours", "UPLH WP1", "UPLH WP2", "People in WIP"}
    drop_these = [c for c in f.columns if c in hide_cols or str(c).startswith("Unnamed:")]
    f_table = f.drop(columns=drop_these, errors="ignore").sort_values(["team", "period_date"]) 
    st.dataframe(f_table, use_container_width=True)
