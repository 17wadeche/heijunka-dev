import hmac
import streamlit as st
st.set_page_config(page_title="Enterprise", layout="wide", initial_sidebar_state="expanded")
st.markdown("""
<style>
/* keep these hidden if you want */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

/* Hide only the pencil (Edit) button */
header button[title*="Edit"],
header button[aria-label*="Edit"] {
  display: none !important;
}

/* Hide only the GitHub icon/link */
header a[href*="github.com"],
header button[title*="GitHub"],
header button[aria-label*="GitHub"] {
  display: none !important;
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
