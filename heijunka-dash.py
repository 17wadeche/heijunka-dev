import streamlit as st

st.set_page_config(page_title="Heijunka", layout="wide")

# ---------- auth helpers ----------
def is_enterprise_allowed() -> bool:
    # Not logged in => no access
    if not st.user.is_logged_in:
        return False

    # Adjust these to your policy:
    allowed_emails = {
        "you@company.com",
        "leader@company.com",
    }
    allowed_domains = {"company.com"}  # optional domain-wide allow

    email = str(getattr(st.user, "email", "") or "").lower().strip()
    if email in allowed_emails:
        return True
    if "@" in email and email.split("@")[-1] in allowed_domains:
        return True
    return False

# ---------- page declarations ----------
portfolio_page = st.Page(
    "pages/Interventional_Vascular.py",
    title="Portfolio Dashboard",   # label shown in side nav
    icon="📊",
)

# Enterprise page can still appear in nav; actual access enforced in page file
enterprise_page = st.Page(
    "pages/enterprise.py",
    title="Enterprise",            # label shown in side nav
    icon="🔒",
)

pg = st.navigation([enterprise_page, portfolio_page])
pg.run()
