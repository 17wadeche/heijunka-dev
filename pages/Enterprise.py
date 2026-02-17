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
