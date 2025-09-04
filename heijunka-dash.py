# heijunka-dash.py
import os
from pathlib import Path
import pandas as pd
import numpy as np
import streamlit as st
import altair as alt
DEFAULT_DATA_PATH = Path(r"C:\heijunka-dev\metrics_aggregate_dev.xlsx")
DATA_URL = st.secrets.get("HEIJUNKA_DATA_URL", os.environ.get("HEIJUNKA_DATA_URL"))
st.set_page_config(page_title="Heijunka Metrics", layout="wide")
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
                "Target UPLH", "Actual UPLH", "HC in WIP"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    if {"Actual Output", "Target Output"}.issubset(df.columns):
        df["Efficiency vs Target"] = (df["Actual Output"] / df["Target Output"]).replace([np.inf, -np.inf], np.nan)
    if {"Completed Hours", "Total Available Hours"}.issubset(df.columns):
        df["Capacity Utilization"] = (df["Completed Hours"] / df["Total Available Hours"]).replace([np.inf, -np.inf], np.nan)
    return df
data_path = None if DATA_URL else str(DEFAULT_DATA_PATH)
mtime_key = 0
if data_path:
    p = Path(data_path)
    mtime_key = p.stat().st_mtime if p.exists() else 0
df = load_data(data_path, DATA_URL)
st.title("Heijunka Metrics Dashboard")
if df.empty:
    st.warning("No data found yet. Make sure metrics_aggregate_dev.xlsx exists and has the 'All Metrics' sheet.")
    st.stop()
teams = sorted([t for t in df["team"].dropna().unique()])
default_teams = teams
col1, col2, col3 = st.columns([2,2,6], gap="large")
with col1:
    selected_teams = st.multiselect("Teams", teams, default=default_teams)
with col2:
    min_date = pd.to_datetime(df["period_date"].min()).date() if df["period_date"].notna().any() else None
    max_date = pd.to_datetime(df["period_date"].max()).date() if df["period_date"].notna().any() else None
    if min_date and max_date:
        start, end = st.date_input("Date range", (min_date, max_date))
    else:
        start, end = None, None
f = df.copy()
if selected_teams:
    f = f[f["team"].isin(selected_teams)]
if start and end:
    f = f[(f["period_date"] >= pd.to_datetime(start)) & (f["period_date"] <= pd.to_datetime(end))]

if f.empty:
    st.info("No rows match your filters.")
    st.stop()
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
tot_target = latest["Target Output"].sum(skipna=True)
tot_actual = latest["Actual Output"].sum(skipna=True)
tot_tahl  = latest["Total Available Hours"].sum(skipna=True)
tot_chl   = latest["Completed Hours"].sum(skipna=True)
tot_hc_wip = latest["HC in WIP"].sum(skipna=True) if "HC in WIP" in latest.columns else np.nan
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
kpi(kpi_cols[2], "Actual Output", tot_actual, "{:,.0f}")
kpi(kpi_cols[3], "Actual vs Target", (tot_actual/tot_target if tot_target else np.nan), "{:.2f}x")
kpi_cols2 = st.columns(4)
kpi(kpi_cols2[1], "Target UPLH", (tot_target/tot_tahl if tot_tahl else np.nan), "{:.2f}")
kpi(kpi_cols2[2], "Actual UPLH", (tot_actual/tot_chl if tot_chl else np.nan), "{:.2f}")
kpi(kpi_cols2[3], "Capacity Utilization", (tot_chl/tot_tahl if tot_tahl else np.nan), "{:.0%}")
kpi_cols3 = st.columns(4)
kpi(kpi_cols3[1], "HC in WIP", tot_hc_wip, "{:,.0f}")
kpi(kpi_cols3[2], "Open Complaint Timeliness (avg)", timeliness_avg, timeliness_fmt)
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
    if have_hours:
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
        base = alt.Chart(hrs_long).encode(
            x=alt.X("period_date:T", title="Week"),
            y=alt.Y("Value:Q", title="Hours"),
            color=alt.Color("Metric:N", title="Series"),
            tooltip=[
                "team:N", "period_date:T", "Metric:N",
                alt.Tooltip("Value:Q", format=",.0f")
            ]
        )
        line = base.mark_line(point=False).encode(
            detail="team:N",
            opacity=alt.condition(team_sel, alt.value(1.0), alt.value(0.25)) if multi_team else alt.value(1.0)
        )
        pts = base.mark_point().encode(
            shape=alt.Shape("team:N", title="Team") if multi_team else alt.value("circle"),
            size=alt.value(45),
            opacity=alt.condition(team_sel, alt.value(1.0), alt.value(0.25)) if multi_team else alt.value(1.0)
        )
        st.altair_chart((line + pts).properties(height=280).add_params(team_sel), use_container_width=True)
    else:
        st.info("Hours columns not found (need 'Total Available Hours' and 'Completed Hours').")
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
left2, right2 = st.columns(2)
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
        pad  = max(0.02, rng * 0.15)  # at least 2% padding
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
        base = alt.Chart(single).encode(
            x=alt.X("period_date:T", title="Week")
        )
        layers = []
        def side(i: int) -> str:
            return "left" if (i % 2 == 0) else "right"
        i = 0
        if "HC in WIP" in selected and "HC in WIP" in single.columns:
            layers.append(
                base.mark_line(point=True).encode(
                    y=alt.Y("HC in WIP:Q",
                            axis=alt.Axis(title="HC in WIP", orient=side(i))),
                    tooltip=["period_date:T", alt.Tooltip("HC in WIP:Q", format=",.0f")]
                )
            )
            i += 1
        if "Open Complaint Timeliness" in selected and "Open Complaint Timeliness" in single.columns:
            layers.append(
                base.mark_line(point=True).encode(
                    y=alt.Y("Open Complaint Timeliness:Q",
                            axis=alt.Axis(title="Timeliness", orient=side(i), format="%")),
                    tooltip=["period_date:T", alt.Tooltip("Open Complaint Timeliness:Q", format=".0%")]
                )
            )
            i += 1
        if "Actual UPLH" in selected and "Actual UPLH" in single.columns:
            layers.append(
                base.mark_line(point=True).encode(
                    y=alt.Y("Actual UPLH:Q",
                            axis=alt.Axis(title="Actual UPLH", orient=side(i))),
                    tooltip=["period_date:T", alt.Tooltip("Actual UPLH:Q", format=".2f")]
                )
            )
            i += 1
        if "Actual Output" in selected and "Actual Output" in single.columns:
            layers.append(
                base.mark_line(point=True).encode(
                    y=alt.Y("Actual Output:Q",
                            axis=alt.Axis(title="Actual Output", orient=side(i))),
                    tooltip=["period_date:T", alt.Tooltip("Actual Output:Q", format=",.0f")]
                )
            )
            i += 1
        if "Actual Hours" in selected and "Completed Hours" in single.columns:
            layers.append(
                base.mark_line(point=True).encode(
                    y=alt.Y("Completed Hours:Q",
                            axis=alt.Axis(title="Actual Hours", orient=side(i))),
                    tooltip=["period_date:T", alt.Tooltip("Completed Hours:Q", format=",.0f")]
                )
            )
            i += 1
        combo = alt.layer(*layers).resolve_scale(y="independent").properties(height=320)
        st.altair_chart(combo, use_container_width=True)
    else:
        st.info("Select at least one series to display.")
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
hide_cols = {"source_file", "fallback_used", "error"}
drop_these = [c for c in f.columns if c in hide_cols or c.startswith("Unnamed:")]
f_table = f.drop(columns=drop_these, errors="ignore").sort_values(["team", "period_date"])
st.dataframe(f_table, use_container_width=True)
