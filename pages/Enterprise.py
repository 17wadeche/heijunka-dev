import hmac
import streamlit as st
st.set_page_config(page_title="Enterprise", layout="wide")
ENTERPRISE_PASSCODE = str(st.secrets.get("enterprise_passcode", "")).strip()
if not ENTERPRISE_PASSCODE:
    st.error(
        "Enterprise passcode is not configured. "
        "Add `enterprise_passcode` in Streamlit Cloud Secrets."
    )
    st.stop()
if "enterprise_unlocked" not in st.session_state:
    st.session_state.enterprise_unlocked = False
def hide_cloud_chrome():
    st.markdown("""
    <style>
    /* Existing cleanup */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    [data-testid="stToolbar"] {display: none;}
    [data-testid="stDecoration"] {display: none;}
    header {visibility: hidden;}

    /* Streamlit Cloud bottom-right Manage app launcher */
    [data-testid="manage-app-button"] {display: none !important;}
    button[aria-label*="Manage app" i] {display: none !important;}
    div[aria-label*="Manage app" i] {display: none !important;}
    </style>
    """, unsafe_allow_html=True)
hide_cloud_chrome()
st.title("Enterprise Dashboard")
if not st.session_state.enterprise_unlocked:
    st.warning("This page is restricted.")
    code = st.text_input("Enter Enterprise passcode", type="password")
    if st.button("Unlock", type="primary"):
        if hmac.compare_digest(code.strip(), ENTERPRISE_PASSCODE):
            st.session_state.enterprise_unlocked = True
            st.rerun()
        else:
            st.error("Invalid passcode.")
    st.stop()
st.success("Welcome to Enterprise.")
st.info("Enterprise dashboard content goes here.")
if st.button("Lock page"):
    st.session_state.enterprise_unlocked = False
    st.rerun()
