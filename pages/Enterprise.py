import json
from pathlib import Path
from datetime import datetime, timezone
import streamlit as st

st.set_page_config(page_title="Enterprise", layout="wide")

# ---------------------------
# Config (prefer secrets)
# ---------------------------
ALLOWED_EMAILS = {
    str(e).strip().lower()
    for e in st.secrets.get("allowed_emails", ["you@company.com", "leader@company.com"])
}
ALLOWED_DOMAINS = {
    str(d).strip().lower()
    for d in st.secrets.get("allowed_domains", ["company.com"])
}

REQUEST_SINK = str(st.secrets.get("request_sink", "file")).strip().lower()
REQUESTS_FILE = Path(str(st.secrets.get("request_file", "access_requests.jsonl")))

# Optional explicit toggle so UI knows whether to show Sign in button.
# Set in secrets.toml: auth_enabled = true
AUTH_ENABLED = bool(st.secrets.get("auth_enabled", False))

# ---------------------------
# Compatibility-safe user helpers
# ---------------------------
def _user_attr(key: str, default=""):
    try:
        return getattr(st.user, key, default)
    except Exception:
        return default

def current_email() -> str:
    return str(_user_attr("email", "") or "").strip().lower()

def current_name() -> str:
    return str(_user_attr("name", "") or "").strip()

def is_logged_in() -> bool:
    flag = _user_attr("is_logged_in", None)
    if isinstance(flag, bool):
        return flag
    return bool(current_email())

def is_allowed(email: str) -> bool:
    if not email:
        return False
    if email in ALLOWED_EMAILS:
        return True
    if "@" in email:
        domain = email.rsplit("@", 1)[-1].lower()
        if domain in ALLOWED_DOMAINS:
            return True
    return False

def save_request(email: str, name: str, reason: str):
    payload = {
        "ts_utc": datetime.now(timezone.utc).isoformat(),
        "email": email,
        "name": name,
        "reason": reason.strip(),
        "app": "enterprise_dashboard",
    }
    if REQUEST_SINK == "file":
        REQUESTS_FILE.parent.mkdir(parents=True, exist_ok=True)
        with REQUESTS_FILE.open("a", encoding="utf-8") as f:
            f.write(json.dumps(payload) + "\n")

def try_login():
    try:
        st.login()
    except Exception:
        st.error(
            "Sign-in is not configured for this deployed app yet. "
            "Enable authentication in your Streamlit Cloud app settings, then try again."
        )

def try_logout():
    try:
        st.logout()
    except Exception:
        st.warning("Sign-out is currently unavailable in this deployment.")

# ---------------------------
# UI
# ---------------------------
st.title("Enterprise Dashboard")

# 1) Not logged in
if not is_logged_in():
    st.warning("This page is restricted.")
    st.write(
        "Please sign in to continue. If you don't have access yet, "
        "you can request it after signing in."
    )

    if AUTH_ENABLED:
        if st.button("Sign in", type="primary"):
            try_login()
    else:
        st.info(
            "Authentication is currently disabled for this deployment. "
            "After you enable auth in Streamlit Cloud, set `auth_enabled = true` in secrets."
        )
    st.stop()

email = current_email()
name = current_name()

# 2) Logged in but not authorized
if not is_allowed(email):
    st.error("You are signed in, but you don’t currently have Enterprise access.")

    with st.expander("Request access", expanded=True):
        st.write(f"Signed in as: **{email or 'Unknown account'}**")
        reason = st.text_area(
            "Why do you need access?",
            placeholder="Team, business need, timeframe…",
            height=120,
        )

        if st.button("Submit access request"):
            if not reason.strip():
                st.warning("Please add a short reason.")
            else:
                try:
                    save_request(email=email, name=name, reason=reason)
                    st.success("Request submitted. You’ll be contacted after review.")
                except Exception:
                    st.error("Could not submit request. Please try again.")

    if AUTH_ENABLED and st.button("Sign out"):
        try_logout()
    st.stop()

# 3) Authorized
st.success(f"Welcome, {name or email}!")
st.info("Enterprise dashboard content goes here.")

if AUTH_ENABLED and st.button("Sign out"):
    try_logout()
