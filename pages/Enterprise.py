import hmac
import streamlit as st
st.set_page_config(
    page_title="Enterprise",
    layout="wide",
    initial_sidebar_state="expanded" if st.session_state.sidebar_open else "collapsed",
)
st.markdown("""
<style>
/* Optional: hide hamburger + footer */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

/* ---- Robust toolbar hide rules ---- */

/* Hide entire top-right toolbar actions cluster (keeps sidebar/nav intact) */
div[data-testid="stToolbarActions"] {
    display: none !important;
}

/* Fallbacks for older/newer builds */
div[data-testid="stToolbar"] button,
div[data-testid="stToolbar"] a {
    display: none !important;
}

/* Extra fallback selectors */
header [data-testid="baseButton-header"],
header button[kind="header"],
header a[href*="github.com"],
header button[aria-label*="GitHub"],
header button[title*="GitHub"],
header button[aria-label*="Edit"],
header button[title*="Edit"],
header button[aria-label*="Source"],
header button[title*="Source"] {
    display: none !important;
}
</style>
""", unsafe_allow_html=True)
if "sidebar_open" not in st.session_state:
    st.session_state.sidebar_open = True
if st.button("Toggle sidebar"):
    st.session_state.sidebar_open = not st.session_state.sidebar_open
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
st.success("Welcome to Enterprise.")
st.info("Enterprise dashboard content goes here.")
if st.button("Lock page"):
    st.session_state[UNLOCK_KEY] = False
    st.rerun()
