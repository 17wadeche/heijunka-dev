import streamlit as st

st.set_page_config(page_title="Enterprise", layout="wide")

def is_enterprise_allowed() -> bool:
    if not st.user.is_logged_in:
        return False

    allowed_emails = {
        "you@company.com",
        "leader@company.com",
    }
    allowed_domains = {"company.com"}

    email = str(getattr(st.user, "email", "") or "").lower().strip()
    if email in allowed_emails:
        return True
    if "@" in email and email.split("@")[-1] in allowed_domains:
        return True
    return False

st.title("Enterprise Dashboard")

if not st.user.is_logged_in:
    st.warning("This page is restricted. Please sign in.")
    if st.button("Sign in"):
        st.login()   # redirects to your configured OIDC provider
    st.stop()

if not is_enterprise_allowed():
    st.error("You are signed in but do not have access to this page.")
    if st.button("Sign out"):
        st.logout()
    st.stop()

# --- restricted content below ---
st.success(f"Welcome, {getattr(st.user, 'name', 'authorized user')}!")
st.info("Enterprise dashboard content here.")
if st.button("Sign out"):
    st.logout()
