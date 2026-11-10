from __future__ import annotations
from pathlib import Path
import pandas as pd
import streamlit as st
import altair as alt

DATA_PATH = Path(__file__).resolve().parent.parent / "team_data" / "combined_portfolio_enterprise.csv"
NUMERIC_COLUMNS = [
    "available_hours",
    "completed_hours",
    "actual_uplh",
    "hc_in_wip",
    "actual_hc_used",
    "people_count",
    "total_non_wip_hours",
    "ooo_hours",
    "wip_pct",
    "nonwip_pct"
]


@st.cache_data(show_spinner="Loading team data...")
def load_data(path: Path) -> pd.DataFrame:
    data = pd.read_csv(path)
    required_columns = {"portfolio", "period_date", *NUMERIC_COLUMNS}
    missing_columns = required_columns.difference(data.columns)
    if missing_columns:
        missing = ", ".join(sorted(missing_columns))
        raise ValueError(f"Dataset is missing required columns: {missing}")

    data["period_date"] = pd.to_datetime(data["period_date"], errors="coerce")
    for column in NUMERIC_COLUMNS:
        data[column] = pd.to_numeric(data[column], errors="coerce")
    return data.dropna(subset=["period_date", "portfolio"])


st.set_page_config(page_title="Enterprise Dashboard", layout="wide")
st.title("Enterprise Dashboard")
st.caption("Metrics and trends from the combined team dataset.")

try:
    data = load_data(DATA_PATH)
except (OSError, ValueError, pd.errors.ParserError) as error:
    st.error(f"Could not load team data from `{DATA_PATH}`: {error}")
    st.stop()

if data.empty:
    st.warning("The dataset has no valid team records to display.")
    st.stop()

team_data = data.copy()

# Create Enterprise, Portfolio, OU selector
options = ["Enterprise", "Portfolio", "OU"]
selected_level = st.pills(" ", options, selection_mode="single", default="Enterprise")

if selected_level == "Portfolio":
    team_data = team_data[team_data["team"].isna()]
    team_data = team_data[team_data["level"] == "Portfolio"]
    portfolios = sorted(team_data["portfolio"].unique().tolist())
    selected_portfolio = st.selectbox("Portfolio", portfolios)

    if selected_portfolio != "All":
        team_data = team_data[team_data["portfolio"] == selected_portfolio]
elif selected_level == "OU":
    team_data = team_data[team_data["team"].notna()]
    team_data = team_data[team_data["level"] == "OU"]

    portfolios = sorted(team_data["portfolio"].unique().tolist())
    selected_portfolio = st.selectbox("Portfolio", portfolios)

    team_data = team_data[team_data["portfolio"] == selected_portfolio]
    ous = sorted(team_data["team"].unique().tolist())
    selected_team = st.selectbox("Team", ous)

    if selected_team != "All":
        team_data = team_data[team_data["team"] == selected_team]
else:
    team_data = data[data["level"] == "Enterprise"]
    selected_portfolio = "Enterprise"


week_start = sorted(team_data["period_date"].unique().tolist())
week_start.reverse()
selected_date = st.selectbox("Select Week", week_start)
if selected_date != "All":
    filtered = team_data[team_data["period_date"] == selected_date]
else:
    filtered = team_data.copy()

if filtered.empty:
    st.info("No records match the selected filters.")
    st.stop()

wip_pct = filtered["wip_pct"].mean()
nonwip_pct = filtered["nonwip_pct"].mean()
# overall metrics
total_completed = filtered["completed_hours"].sum()
total_nonwip = filtered["total_non_wip_hours"].sum()
total_ooo = filtered["ooo_hours"].sum()
# averages
unaccounted_hrs = team_data["unaccounted_hrs"].sum()
total_hours = team_data["hc_in_wip"].sum() * 40
avg_unaccounted = unaccounted_hrs/total_hours

total_hc = filtered["hc_in_wip"].sum()

st.caption(
    f"Metric cards summarize {selected_portfolio}"
)
_, c1, c2, _ = st.columns([1.2, 1.2, 1.2, 1.2])
#_safe_metric(_metric_from_df(metric_df, "wip_pct"), pct=True)
c1.metric(
    "Avg Per Person **WIP** Daily Hours",
    round(total_completed/total_hc/5, 2),
)
c2.metric(
    "Avg Per Person **Non-WIP** Daily Hours",
    round(total_nonwip/total_hc/5, 2),
)
_, p1, p2, _ = st.columns([1.2, 1.2, 1.2, 1.2])
p1.metric(
    "**WIP** Ratio",
    f"{wip_pct*100:,.1f}%"
)
p2.metric(
    "**Non-WIP** Ratio",
    f"{nonwip_pct*100:,.1f}%"
)
st.divider()
_, _, c3, c4, c5, _, _ = st.columns([.8, .8, 1.2, 1.2, 1.2, 1.0, 0.5])
c3.metric(
    "Avg **Other Team WIP** Hours",
    0
)
c4.metric(
    "Avg **OOO** Weekly Hours",
    round(team_data["ooo_hours"].mean(), 1),
)
unacc_pct = filtered
c5.metric(
    "Avg **Unaccounted** Weekly Hours",
    round(team_data["unaccounted_hrs"].mean(), 1),
)
_, _, p3, p4, p5, _, _ = st.columns([.8, .8, 1.2, 1.2, 1.2, 1.0, 0.5])
p3.metric(
    "**Other Team WIP** % of week",
    0,
)
p4.metric(
    "**OOO** % of week",
    fr"{round(team_data["ooo_pct"].mean(),1)*100}%",
)
p5.metric(
    "**Unaccounted** % remaining",
    fr"{round(avg_unaccounted,1)*100}%",
)
st.divider()

st.subheader(fr"{selected_portfolio} WIP Pct Trend")
base = alt.Chart(team_data).encode(
    x=alt.X(
        "period_date:T",
        title="Week of",
        axis=alt.Axis(
            format="%Y-%m-%d",
            labelAngle=-35,
            tickCount="week",
        ),
        timeUnit="yearmonthdate",
    ),
    y=alt.Y(
        "wip_pct:Q",
        title="WIP %",
        axis=alt.Axis(format=".0%"),
        scale=alt.Scale(
            domain=[
                0,
                max(
                    1.0,
                    float(team_data["wip_pct"].max()) * 1.1,
                ),
            ]
        ),
    ),
    tooltip=[
        alt.Tooltip("period_date:N", title="Week"),
        alt.Tooltip("wip_pct:Q", title="WIP %", format=".1%"),
    ]
)
line = base.mark_line(strokeWidth=3).encode()
points = base.mark_circle(size=80).encode()
latest_point = alt.Chart(filtered).mark_circle(
    size=180,
    filled=True,
).encode(
    x="period_date:T",
    y="wip_pct:Q",
    tooltip=[
        alt.Tooltip("period_date:T", title="Week"),
        alt.Tooltip("wip_pct:Q", title="WIP %", format=".1%"),
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
st.altair_chart(trend_chart, width="stretch")

st.subheader("Team metrics")
display_columns = [
    "team",
    "portfolio",
    "period_date",
    "completed_hours",
    "available_hours",
    "actual_uplh",
    "hc_in_wip",
    "people_count",
    "total_non_wip_hours",
    "ooo_hours",
    "wip_pct",
    "nonwip_pct"
]
st.dataframe(
    filtered[display_columns].sort_values(["portfolio", "period_date"]),
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
