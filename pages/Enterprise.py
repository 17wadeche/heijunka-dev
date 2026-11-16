from __future__ import annotations
from pathlib import Path
import pandas as pd
import streamlit as st
import altair as alt
import json
from itertools import chain

# load relevant configurations
CONFIG_PATH = Path(__file__).resolve().parent.parent / "config" / "enterprise_org.json"
with open(CONFIG_PATH, "r") as f:
    org_config = json.load(f)

FISCAL_PATH = Path(__file__).resolve().parent.parent / "config" / "fiscal_month.json"
with open(FISCAL_PATH, "r") as f:
    fiscal_data = json.load(f)

# Create iterable of FY27 Fiscal Months
fy27 = fiscal_data["FY27"]
fy27_fiscal_mos = set()
fy27_week_start = list()
fy27_month_weeks = dict()
for week, month in fy27.items():
    fy27_fiscal_mos.add(month)
    fy27_week_start.append(week)
    if month in fy27_month_weeks:
        fy27_month_weeks[month].append(week)
    else:
        fy27_month_weeks[month] = [week]

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
def get_teams_from_org_config(level: str, value: str):
  # get teams list from json
  teams = org_config["teams"]
  if len(value) > 0:
    # filter list on level and value provided
    return [t for t in teams if t[level] == value]
  # otherwise return full list of teams
  return teams


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
options = ["Enterprise", "Portfolio", "OU", "Team"]
portfolios = ["Neuroscience", "CRM", "MS", "IV"]

portfolio, dates = st.columns(2)
selection = ""

with portfolio:
    selected_level = st.pills(" ", options, selection_mode="single", default="Enterprise")
    if selected_level == "Enterprise":
        selected_portfolio = "Enterprise"
        selection = selected_portfolio
        team_data = data[data["level"] == "Enterprise"]
    else:
        selected_portfolio = st.selectbox("Portfolio", portfolios)
        if selected_level == "Portfolio":
            team_data = team_data[team_data["team"].isna()]
            team_data = team_data[team_data["portfolio"] == selected_portfolio]
            selection = selected_portfolio
        else:
            filtered_ous = get_teams_from_org_config("portfolio", selected_portfolio)
            ous_set = set([ou["ou"] for ou in filtered_ous])   # get unique list
            selected_ou =  st.selectbox("OUs", ous_set)
            if selected_level == "OU":
                team_data = team_data[team_data["ou"] == selected_ou]
                selection = selected_ou
            else:
                team_set = set([t["team"] for t in filtered_ous])
                selected_team = st.selectbox("Teams", team_set)
                team_data = team_data[team_data["team"] == selected_team]
                selection = selected_team

with dates:
    week_start = sorted(team_data["period_date"].unique().tolist())
    week_start.reverse()
    date_options = ["Week", "Fiscal Month", "Fiscal Quarter", "Fiscal Year"]
    date_level = st.pills(" ", date_options, selection_mode="single", default="Week")
    if date_level == "Fiscal Year":
        selected_dates = week_start
    elif date_level == "Fiscal Quarter":
        fq_choices = ["FY27 Q1", "FY27 Q2", "FY27 Q3", "FY27 Q4"]
        fq = st.pills("Choose Quarter", fq_choices, selection_mode="single", default="FY27 Q3")
        # Create layers to get dates
        if fq:
            st.caption(f"Auto selected months in {fq}")
            if fq == "FY27 Q1":
                fm_choices = ["FY27 May", "FY27 June", "FY27 July"]
            elif fq == "FY27 Q2":
                fm_choices = ["FY27 August", "FY27 September", "FY27 October"]
            elif fq == "FY27 Q3": 
                fm_choices = ["FY27 November", "FY27 December", "FY27 January"]
            elif fq == "FY27 Q4": 
                fm_choices = ["FY27 February", "FY27 March", "FY27 April"]
            if len(fm_choices) > 0:
                weeks_autoselected = list(chain.from_iterable(fy27_month_weeks[fm] for fm in fm_choices))
                selected_dates = [pd.Timestamp(x).normalize() for x in weeks_autoselected]
    elif date_level == "Fiscal Month":
        fm_choices = fy27_month_weeks.keys()
        selected_mos = st.multiselect("Select Fiscal Month(s)", fm_choices)
        # if month choices made
        if len(fm_choices) > 0:
            st.caption(f"Auto selected weeks in {", ".join(selected_mos)}")
            weeks_autoselected = list(chain.from_iterable(fy27_month_weeks[fm] for fm in selected_mos))
            selected_dates = [pd.Timestamp(x).normalize() for x in weeks_autoselected]
    else:
        selected_dates = st.multiselect("Select Week(s)", week_start)

if len(selected_dates) > 0:
    filtered = team_data[team_data["period_date"].isin(selected_dates)]
else:
    filtered = team_data.copy()

if filtered.empty:
    st.info("No records match the selected filters.")
    st.stop()

def sans_nones(dfColumn: pd.DataFrame):
    return dfColumn[dfColumn.notna()]
wip_pct = sans_nones(filtered["wip_pct"]).mean()
nonwip_pct = sans_nones(filtered["nonwip_pct"]).mean()
# overall metrics
total_completed = sans_nones(filtered["completed_hours"]).sum()
total_nonwip = sans_nones(filtered["total_non_wip_hours"]).sum()
total_ooo = sans_nones(filtered["ooo_hours"]).sum()
# averages
unaccounted_hrs = sans_nones(team_data["unaccounted_hrs"]).sum()
total_hours = sans_nones(team_data["hc_in_wip"]).sum() * 40
avg_unaccounted = unaccounted_hrs/total_hours

total_hc = filtered["hc_in_wip"].sum()

st.caption(
    f"Metric cards summarize {selection}"
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

st.subheader(fr"{selection} WIP Pct Trend")
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
