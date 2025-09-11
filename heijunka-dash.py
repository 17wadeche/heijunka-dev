# heijunka-dash.py
import os
from pathlib import Path
import pandas as pd
import numpy as np
import streamlit as st
import altair as alt
import json
DEFAULT_DATA_PATH = Path(r"C:\heijunka-dev\metrics_aggregate_dev.xlsx")
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
            df[col] = pd.to_numeric(df[col], errors="coerce")
    if {"Actual Output", "Target Output"}.issubset(df.columns):
        df["Efficiency vs Target"] = (df["Actual Output"] / df["Target Output"]).replace([np.inf, -np.inf], np.nan)
    if {"Completed Hours", "Total Available Hours"}.issubset(df.columns):
        df["Capacity Utilization"] = (df["Completed Hours"] / df["Total Available Hours"]).replace([np.inf, -np.inf], np.nan)
    return df
def explode_ph_person_hours(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "PH Person Hours" not in df.columns:
        return pd.DataFrame(columns=["team","period_date","person","Actual Hours","Available Hours","Utilization"])
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
            util = (a / t) if (pd.notna(a) and pd.notna(t) and t != 0) else np.nan
            rows.append({
                "team": r["team"],
                "period_date": pd.to_datetime(r["period_date"], errors="coerce"),
                "person": str(person),
                "Actual Hours": a,
                "Available Hours": t,
                "Utilization": util
            })
    out = pd.DataFrame(rows)
    if not out.empty:
        out["period_date"] = pd.to_datetime(out["period_date"], errors="coerce").dt.normalize()
    return out
data_path = None if DATA_URL else str(DEFAULT_DATA_PATH)
mtime_key = 0
if data_path:
    p = Path(data_path)
    mtime_key = p.stat().st_mtime if p.exists() else 0
df = load_data(data_path, DATA_URL)
st.markdown("<h1 style='text-align: center;'>Heijunka Metrics Dashboard</h1>", unsafe_allow_html=True)
if df.empty:
    st.warning("No data found yet. Make sure metrics_aggregate_dev.xlsx exists and has the 'All Metrics' sheet.")
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
ph_people = explode_ph_person_hours(f)
latest = (f.sort_values(["team", "period_date"])
            .groupby("team", as_index=False)
            .tail(1))
kpi_cols = st.columns(4)
def kpi(col, label, value, fmt="{:,.2f}"):
    if pd.isna(value):
        col.metric(label, "—")
    else:
        try:
            col.metric(label, fmt.format(value))
        except Exception:
            col.metric(label, str(value))
def kpi_vs_target(col, label, actual, target, fmt_val="{:,.2f}"):
    if pd.isna(actual) or pd.isna(target) or not target:
        col.metric(label, "—")
        return
    try:
        value_str = fmt_val.format(actual)
    except Exception:
        value_str = str(actual)
    diff = (float(actual) - float(target)) / float(target)
    delta_str = f"{diff:+.0%} vs target"
    col.metric(label, value_str, delta=delta_str, delta_color="normal")
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
kpi(kpi_cols2[1], "Target UPLH", (tot_target/tot_tahl if tot_tahl else np.nan), "{:.2f}")
kpi_vs_target(kpi_cols2[2], "Actual UPLH", actual_uplh, target_uplh, "{:.2f}")
kpi(kpi_cols2[3], "Capacity Utilization", (tot_chl/tot_tahl if tot_tahl else np.nan), "{:.0%}")
kpi_cols3 = st.columns(4)
kpi(kpi_cols3[1], "HC in WIP", tot_hc_wip, "{:,.0f}")
kpi(kpi_cols3[2], "Actual HC used", tot_hc_used, "{:,.2f}")
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
    ph_only = (len(teams_in_view) == 1 and teams_in_view[0] == "PH")
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
        if ph_only and "PH Person Hours" in f.columns:
            ph_people = explode_ph_person_hours(f)
        else:
            ph_people = pd.DataFrame()
        if ph_only and not ph_people.empty:
            sel_week = alt.selection_point(
                fields=["period_date"],
                on="click",
                nearest=True,
                clear="dblclick"
            )
            trend = (
                alt.Chart(hrs_long)
                .mark_line(point=True)
                .encode(
                    x=alt.X("period_date:T", title="Week"),
                    y=alt.Y("Value:Q", title="Hours"),
                    color=alt.Color("Metric:N", title="Series"),
                    tooltip=["team:N", "period_date:T", "Metric:N", alt.Tooltip("Value:Q", format=",.0f")],
                )
                .properties(height=280)
            )
            rule = (
                alt.Chart(hrs_long)
                .transform_filter(sel_week)
                .mark_rule(strokeDash=[4, 3])
                .encode(x="period_date:T")
            )
            bars = (
                alt.Chart(ph_people)
                .transform_filter(sel_week)
                .transform_fold(["Actual Hours", "Available Hours"], as_=["Metric", "Value"])
                .mark_bar()
                .encode(
                    x=alt.X("person:N", title="Person", sort=alt.Sort(field="person")),
                    y=alt.Y("Value:Q", title="Hours"),
                    color=alt.Color("Metric:N", title="Series"),
                    tooltip=[
                        "person:N", "Metric:N",
                        alt.Tooltip("Value:Q", format=",.1f"),
                        alt.Tooltip("period_date:T", title="Week"),
                    ],
                )
                .properties(height=230, title="Per-person (click a point above; double-click to clear)")
            )
            util = (
                alt.Chart(ph_people)
                .transform_filter(sel_week)
                .mark_point()
                .encode(
                    x=alt.X("person:N", title="Person", sort=alt.Sort(field="person")),
                    y=alt.Y("Utilization:Q", axis=alt.Axis(title="Utilization", format="%")),
                    tooltip=[
                        "person:N",
                        alt.Tooltip("Utilization:Q", format=".0%"),
                        alt.Tooltip("Actual Hours:Q", format=",.1f"),
                        alt.Tooltip("Available Hours:Q", format=",.1f"),
                        alt.Tooltip("period_date:T", title="Week"),
                    ],
                )
                .properties(height=230)
            )
            top = (trend + rule).add_params(sel_week)
            combined = alt.vconcat(
                top,
                (bars | util).resolve_scale(y="independent")
            )
            st.caption("Click a week to drill down (double-click to clear).")
            st.altair_chart(combined, use_container_width=True)
        else:
            team_sel = alt.selection_point(fields=["team"], bind="legend")
            base = alt.Chart(hrs_long).encode(
                x=alt.X("period_date:T", title="Week"),
                y=alt.Y("Value:Q", title="Hours"),
                color=alt.Color("Metric:N", title="Series"),
                tooltip=["team:N", "period_date:T", "Metric:N", alt.Tooltip("Value:Q", format=",.0f")],
            )
            line = base.mark_line(point=False).encode(
                detail="team:N",
                opacity=alt.condition(team_sel, alt.value(1.0), alt.value(0.25))
                if len(teams_in_view) > 1 else alt.value(1.0)
            )
            pts = base.mark_point().encode(
                shape=alt.Shape("team:N", title="Team") if len(teams_in_view) > 1 else alt.value("circle"),
                size=alt.value(45),
                opacity=alt.condition(team_sel, alt.value(1.0), alt.value(0.25))
                if len(teams_in_view) > 1 else alt.value(1.0)
            )
            st.altair_chart((line + pts).properties(height=280).add_params(team_sel), use_container_width=True)
st.write("Altair", alt.__version__, "PH rows:", len(ph_people), "hrs_long rows:", len(hrs_long), "PH-only:", ph_only)
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
with right:
    st.subheader("UPLH Trend")
    have_target_uplh = "Target UPLH" in f.columns
    uplh_vars = ["Actual UPLH"] + (["Target UPLH"] if have_target_uplh else [])
    uplh_long = (
        f.melt(
            id_vars=["team", "period_date"],
            value_vars=uplh_vars,
            var_name="Metric", value_name="Value"
        ).dropna(subset=["Value"])
    )
    base = alt.Chart(uplh_long).encode(
        x=alt.X("period_date:T", title="Week"),
        y=alt.Y("Value:Q", title="UPLH"),
        color=alt.Color("Metric:N", title="Series"),
        tooltip=["team:N", "period_date:T", "Metric:N", alt.Tooltip("Value:Q", format=",.2f")]
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
    selected = st.multiselect("Series", available, default=available, key="single_team_series")
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
hide_cols = {"source_file", "fallback_used", "error", "PH Person Hours"}
drop_these = [c for c in f.columns if c in hide_cols or c.startswith("Unnamed:")]
f_table = f.drop(columns=drop_these, errors="ignore").sort_values(["team", "period_date"])
st.dataframe(f_table, use_container_width=True)
