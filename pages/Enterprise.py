import json
import hmac
from pathlib import Path
from datetime import datetime, timezone
import streamlit as st

st.set_page_config(page_title="Enterprise", layout="wide")

# ---------------------------
# Config
# ---------------------------
ALLOWED_EMAILS = {
    str(e).strip().lower()
    for e in st.secrets.get("allowed_emails", [])
}
ALLOWED_DOMAINS = {
    str(d).strip().lower()
    for d in st.secrets.get("allowed_domains", [])
}

REQUEST_SINK = str(st.secrets.get("request_sink", "file")).strip().lower()
REQUESTS_FILE = Path(str(st.secrets.get("request_file", "access_requests.jsonl")))

# Native auth is NOT available in your current deployment
AUTH_ENABLED = bool(st.secrets.get("auth_enabled", False)) and hasattr(st, "login")

# Fallback shared gate secret (required when AUTH_ENABLED = false)
ENTERPRISE_PASSCODE = str(st.secrets.get("enterprise_passcode", "")).strip()

# ---------------------------
# Helpers
# ---------------------------
def is_allowed_email(email: str) -> bool:
    email = (email or "").strip().lower()
    if not email:
        return False
    if email in ALLOWED_EMAILS:
        return True
    if "@" in email:
        domain = email.rsplit("@", 1)[-1]
        return domain in ALLOWED_DOMAINS
    return False

def save_request(email: str, name: str, reason: str):
    payload = {
        "ts_utc": datetime.now(timezone.utc).isoformat(),
        "email": (email or "").strip().lower(),
        "name": (name or "").strip(),
        "reason": reason.strip(),
        "app": "enterprise_dashboard",
    }
    if REQUEST_SINK == "file":
        REQUESTS_FILE.parent.mkdir(parents=True, exist_ok=True)
        with REQUESTS_FILE.open("a", encoding="utf-8") as f:
            f.write(json.dumps(payload) + "\n")

def passcode_ok(user_input: str) -> bool:
    # constant-time compare
    return bool(ENTERPRISE_PASSCODE) and hmac.compare_digest(
        user_input.strip(), ENTERPRISE_PASSCODE
    )

# ---------------------------
# UI
# ---------------------------
st.title("Enterprise Dashboard")

if AUTH_ENABLED:
    # If you later get native auth, you can re-enable this branch.
    st.info("Native auth mode is enabled for this deployment.")
    # ... your st.login/st.user flow here ...
    st.stop()

# ---- Fallback: passcode-gated access ----
st.warning("This page is restricted.")

if "enterprise_unlocked" not in st.session_state:
    st.session_state.enterprise_unlocked = False

if not st.session_state.enterprise_unlocked:
    st.write("Enter your Enterprise access code to continue.")
    code = st.text_input("Enterprise access code", type="password")
    col1, col2 = st.columns([1, 3])

    with col1:
        if st.button("Unlock", type="primary"):
            if passcode_ok(code):
                st.session_state.enterprise_unlocked = True
                st.rerun()
            else:
                st.error("Invalid access code.")

    with col2:
        if st.button("Request access"):
            st.session_state.show_request = True

    if st.session_state.get("show_request", False):
        with st.expander("Request Enterprise access", expanded=True):
            req_email = st.text_input("Work email")
            req_name = st.text_input("Name")
            reason = st.text_area(
                "Why do you need access?",
                placeholder="Team, business need, timeframe…",
                height=120,
            )
            if st.button("Submit request"):
                if not req_email.strip() or not reason.strip():
                    st.warning("Please provide at least email and reason.")
                else:
                    try:
                        save_request(req_email, req_name, reason)
                        st.success("Request submitted. You’ll be contacted after review.")
                    except Exception as e:
                        st.error(f"Could not submit request: {e}")
    st.stop()

# ---- Authorized content ----
st.success("Welcome to Enterprise.")
st.info("Enterprise dashboard content goes here.")

if st.button("Lock page"):
    st.session_state.enterprise_unlocked = False
    st.rerun()
