import os
import json
from pathlib import Path
from datetime import datetime, timezone
import streamlit as st

st.set_page_config(page_title="Enterprise", layout="wide")

# ---- access policy ----
ALLOWED_EMAILS = {
    "you@company.com",
    "leader@company.com",
}
ALLOWED_DOMAINS = {"company.com"}  # optional

# Where to store requests (for local/dev). On Streamlit Cloud, use external sink (see below).
REQUESTS_FILE = Path("access_requests.jsonl")

def current_email() -> str:
    return str(getattr(st.user, "email", "") or "").strip().lower()

def is_allowed(email: str) -> bool:
    if not email:
        return False
    if email in ALLOWED_EMAILS:
        return True
    if "@" in email and email.split("@")[-1] in ALLOWED_DOMAINS:
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
    REQUESTS_FILE.parent.mkdir(parents=True, exist_ok=True)
    with REQUESTS_FILE.open("a", encoding="utf-8") as f:
        f.write(json.dumps(payload) + "\n")

st.title("Enterprise Dashboard")

# 1) Not logged in
if not st.user.is_logged_in:
    st.warning("This page is restricted.")
    st.write("Please sign in to continue. If you don't have access yet, you can request it after signing in.")
    if st.button("Sign in"):
        st.login()
    st.stop()

email = current_email()
name = str(getattr(st.user, "name", "") or "")

# 2) Logged in but not authorized
if not is_allowed(email):
    st.error("You are signed in, but you don’t currently have Enterprise access.")

    with st.expander("Request access", expanded=True):
        st.write(f"Signed in as: **{email}**")
        reason = st.text_area(
            "Why do you need access?",
            placeholder="Team, business need, timeframe…",
            height=120
        )
        submitted = st.button("Submit access request")
        if submitted:
            if not reason.strip():
                st.warning("Please add a short reason.")
            else:
                save_request(email=email, name=name, reason=reason)
                st.success("Request submitted. You’ll be contacted after review.")

    if st.button("Sign out"):
        st.logout()
    st.stop()

# 3) Authorized
st.success(f"Welcome, {name or email}!")
st.info("Enterprise dashboard content goes here.")
if st.button("Sign out"):
    st.logout()
