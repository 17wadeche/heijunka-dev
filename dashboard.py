from __future__ import annotations
from pathlib import Path
import pandas as pd
import streamlit as st
import altair as alt

NUMERIC_COLUMNS = [
  "available_hours",
  "completed_hours",
  "target_output",
  "actual_output",
  "target_uplh",
  "actual_uplh",
  "hc_in_wip",
  "actual_hc_used"
]

@st.cache_data(show_spinner="Loading team data...")
def load_data(path: Path) -> pd.DataFrame:
    data = pd.read_csv(path)
    required_columns = {"team", "period_date", *NUMERIC_COLUMNS}
    missing_columns = required_columns.difference(data.columns)
    if missing_columns:
        missing = ", ".join(sorted(missing_columns))
        raise ValueError(f"Dataset is missing required columns: {missing}")

    data["period_date"] = pd.to_datetime(data["period_date"], errors="coerce")
    for column in NUMERIC_COLUMNS:
        data[column] = pd.to_numeric(data[column], errors="coerce")
    return data.dropna(subset=["period_date", "team"])

def create_page(page_title: str, caption: str, file_path: str):
  st.set_page_config(page_title=page_title, layout="wide")
  st.title(page_title)
  st.caption(caption)
  try:
    data = load_data(file_path)
  except (OSError, ValueError, pd.errors.ParserError) as error:
    st.error(f"Could not load team data from `{file_path}`: {error}")
    st.stop()

  if data.empty:
    st.warning("The dataset has no valid team records to display.")
    st.stop()

  filtered = data.copy()
  team_data = data.copy()
  teams = sorted(data["team"].unique().tolist())
  selected_team = st.selectbox("Team", teams)

  if selected_team != "All":
    team_data = team_data[team_data["team"] == selected_team]

  week_start = sorted(team_data["period_date"].unique().tolist())
  selected_date = st.selectbox("Select Week", week_start)
  if selected_date != "All":
    filtered = team_data[team_data["period_date"] == selected_date]

  if filtered.empty:
    st.info("No records match the selected filters.")
    st.stop()

  #wip_pct = filtered["wip_pct"].mean()
  #nonwip_pct = filtered["nonwip_pct"].mean()

  total_completed = filtered["completed_hours"].sum()
  avg_completed = filtered["completed_hours"].mean()
  avg_available = filtered["available_hours"].mean()
  #total_nonwip = filtered["total_non_wip_hours"].sum()
  #total_ooo = filtered["ooo_hours"].sum()

  avg_hc = filtered["hc_in_wip"].mean()
  avg_in_wip = filtered["actual_hc_used"].mean()
  avg_out = filtered["actual_output"].mean()
  avg_uplh = filtered["actual_uplh"].mean()
  capacity = round(filtered["completed_hours"].sum()/filtered["available_hours"].sum(),2)
  #avg_closures = filtered["closures"].mean()

  st.caption(
    f"Weekly metrics for {selected_team}"
  )
  a1, a2, a3, a4, a5 = st.columns([0.8,0.8,0.8,0.8,0.8])
  #_safe_metric(_metric_from_df(metric_df, "wip_pct"), pct=True)
  a1.metric(
    "HC in WIP",
    round(avg_hc, 0),
  )
  a2.metric(
    "Actual HC Used",
    round(avg_in_wip, 1),
  )
  a3.metric(
    "Capacity Utilization",
    capacity
  )
  a4.metric(
    "Actual Output",
    round(avg_out, 1),
  )
  a5.metric(
    "Actual UPLH",
    round(avg_uplh, 1),
  )
  #a6.metric(
  #  "Closures*",
  #  round(avg_in_wip, 1),
  #)
  st.divider()

  st.subheader(fr"{selected_team} WIP Trends")
  # create charts
  out_chart = make_uplh_chart(team_data)
  st.altair_chart(out_chart, width="stretch")

  st.subheader("Team metrics")
  display_columns = [
    "team",
    "period_date",
    "available_hours",
    "completed_hours",
    "target_output",
    "actual_output",
    "target_uplh",
    "actual_uplh",
    "hc_in_wip",
    "actual_hc_used"
  ]
  st.dataframe(
    team_data[display_columns].sort_values(["team", "period_date"]),
    width="stretch",
    hide_index=True,
  )
  st.markdown(
    """
    <style>
    div[data-testid="stMetric"]{ text-align: center; }
    label[data-testid="stMetricLabel"]{ display: block; width: 100%; text-align: center; margin: 0; }
    label[data-testid="stMetricLabel"] p{ text-align: center !important; margin: 0 !important; }
    div[data-testid="stMetricValue"]{ text-align: center !important; width: 100%; }
    </style>
    """,
    unsafe_allow_html=True,
  )

def make_uplh_chart(data: pd.Dataframe):
  target_max = data["target_uplh"].max()
  max_height = data["actual_uplh"].max()
  if target_max > max_height:
    max_height = target_max
  base = alt.Chart(data).transform_fold(
    ['target_uplh', 'actual_uplh'],
    as_=['metric', 'uplh']
  ).encode(
    x=alt.X(
      fr"period_date:T",
      title="Week of",
      axis=alt.Axis(
        format="%Y-%m-%d",
        labelAngle=-35,
        tickCount="week",
      ),
      timeUnit="yearmonthdate",
    ),
    y=alt.Y(
      fr"uplh:Q",
      title="UPLH",
      axis=alt.Axis(),
      scale=alt.Scale(
        domain=[
          0,
          max(
            1.0,
            float(max_height) * 1.1,
          ),
        ]
      ),
    ),
    color='metric:N',
    tooltip=[
      alt.Tooltip(fr"period_date:N", title="Week of"),
      alt.Tooltip(fr"uplh:Q", title="UPLH"),
    ]
  )
  line = base.mark_line(strokeWidth=3).encode()
  points = base.mark_circle(size=80).encode()
  latest_point = alt.Chart(data).mark_circle(
    size=180,
    filled=True,
  ).encode(
    x=fr"period_date:T",
    y=fr"uplh:Q",
    tooltip=[
      alt.Tooltip(fr"period_date:T", title="Week of"),
      alt.Tooltip(fr"uplh:Q", title="UPLH"),
    ]
  )
  data_labels = base.mark_text(
    align="center",
    dy=-12,
    fontSize=11,
  )
  rule = alt.Chart(
    pd.DataFrame({"y": [0.80]})
  ).mark_rule(
    strokeDash=[6, 4]
  ).encode(
    y="y:Q"
  )
  trend_chart = (
    alt.layer(rule, line, points, latest_point)
    .properties(height=360)
    .interactive()
  )
  return trend_chart
