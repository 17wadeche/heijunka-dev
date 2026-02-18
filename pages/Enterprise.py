# pages/Enterprise.py
import hmac
import json
from pathlib import Path
import streamlit as st
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from utils.styles import apply_global_styles
apply_global_styles()
ENTERPRISE_PASSCODE = str(st.secrets.get("enterprise_passcode", "")).strip()
if not ENTERPRISE_PASSCODE:
    st.error("Enterprise passcode is not configured.")
    st.stop()
UNLOCK_KEY = "enterprise_unlocked"
if UNLOCK_KEY not in st.session_state:
    st.session_state[UNLOCK_KEY] = False
st.title("Enterprise Dashboard")
if not st.session_state[UNLOCK_KEY]:
    st.warning("This page is restricted.")
    code = st.text_input("Enter Enterprise passcode", type="password")
    if st.button("Unlock", type="primary"):
        if hmac.compare_digest(code.strip(), ENTERPRISE_PASSCODE):
            st.session_state[UNLOCK_KEY] = True
            st.rerun()
        else:
            st.error("Invalid passcode.")
    st.stop()
BASE_DIR = Path(__file__).resolve().parents[1]  # project root (one level above /pages)
CONFIG_PATH = BASE_DIR / "config" / "enterprise_org.json"
@st.cache_data(show_spinner=False)
def load_org_config(path: Path) -> list[dict]:
    if not path.exists():
        return []
    data = json.loads(path.read_text(encoding="utf-8"))
    teams = data.get("teams", [])
    if not isinstance(teams, list):
        return []
    normalized = []
    for row in teams:
        if not isinstance(row, dict):
            continue
        portfolio = str(row.get("portfolio", "")).strip()
        team = str(row.get("team", "")).strip()
        ou = str(row.get("ou", "")).strip()
        if portfolio and team:
            normalized.append(
                {
                    "Portfolio": portfolio,
                    "Team": team,
                    "OU": ou if ou else "Unassigned",
                }
            )
    return normalized
org_rows = load_org_config(CONFIG_PATH)
if not org_rows:
    st.error(
        "No org config found.\n\n"
        f"Expected file at: {CONFIG_PATH}\n"
        "Create config/enterprise_org.json and add teams."
    )
    st.stop()
st.subheader("Filters")
portfolios = sorted({r["Portfolio"] for r in org_rows})
ous = sorted({r["OU"] for r in org_rows})
teams = sorted({r["Team"] for r in org_rows})
c1, c2, c3 = st.columns(3)
with c1:
    sel_portfolios = st.multiselect("Portfolio", portfolios, default=portfolios)
with c2:
    sel_ous = st.multiselect("OU", ous, default=ous)
with c3:
    sel_teams = st.multiselect("Team", teams, default=teams)
filtered = [
    r for r in org_rows
    if r["Portfolio"] in sel_portfolios
    and r["OU"] in sel_ous
    and r["Team"] in sel_teams
]
st.subheader(f"Org View ({len(filtered)})")
grouped = {}
for r in filtered:
    grouped.setdefault(r["Portfolio"], {}).setdefault(r["OU"], []).append(r["Team"])
for p in grouped:
    for ou in grouped[p]:
        grouped[p][ou] = sorted(set(grouped[p][ou]))
for portfolio in sorted(grouped.keys()):
    with st.expander(f"📁 {portfolio}", expanded=True):
        ou_map = grouped[portfolio]
        for ou in sorted(ou_map.keys()):
            st.markdown(f"**OU:** {ou}")
            st.write(", ".join(ou_map[ou]) if ou_map[ou] else "—")
st.divider()
st.subheader("Table")
st.dataframe(filtered, use_container_width=True)
cA, cB = st.columns([1, 1])
with cA:
    if st.button("Reload config"):
        st.cache_data.clear()
        st.rerun()
with cB:
    if st.button("Lock page"):
        st.session_state[UNLOCK_KEY] = False
        st.rerun()
