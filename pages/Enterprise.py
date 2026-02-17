import hmac
import streamlit as st
st.set_page_config(page_title="Enterprise", layout="wide", initial_sidebar_state="expanded")
st.markdown("""
<style>
/* Hide hamburger main menu + footer */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

/* Keep header visible so sidebar toggle can exist */
header[data-testid="stHeader"] {
  visibility: visible !important;
  background: transparent !important;
}

/* Hide only the top-right toolbar items */
[data-testid="stToolbar"] {
  display: none !important;
}

/* Ensure sidebar collapsed/expand button stays visible & clickable */
[data-testid="collapsedControl"] {
  display: flex !important;
  visibility: visible !important;
  opacity: 1 !important;
  position: fixed !important;
  top: 0.75rem !important;
  left: 0.75rem !important;
  z-index: 999999 !important;
}
</style>
""", unsafe_allow_html=True)

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
