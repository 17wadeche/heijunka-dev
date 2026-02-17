import hmac
import streamlit as st
import streamlit.components.v1 as components
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
import streamlit as st
import streamlit.components.v1 as components

def hide_cloud_chrome():
    # 1) CSS attempt (works when element is in same DOM context)
    st.markdown("""
    <style>
    #MainMenu, footer, header {visibility: hidden !important;}
    [data-testid="stToolbar"], [data-testid="stDecoration"] {display:none !important;}

    /* Known/likely Manage app launchers */
    [data-testid="manage-app-button"],
    [data-testid*="manage" i][data-testid*="app" i],
    button[aria-label*="Manage app" i],
    div[aria-label*="Manage app" i],
    [title*="Manage app" i] {
      display: none !important;
      visibility: hidden !important;
      opacity: 0 !important;
      pointer-events: none !important;
    }
    </style>
    """, unsafe_allow_html=True)

    # 2) JS fallback: continuously hide if Cloud re-injects it
    components.html("""
    <script>
    (function () {
      const sels = [
        '[data-testid="manage-app-button"]',
        '[data-testid*="manage" i][data-testid*="app" i]',
        'button[aria-label*="Manage app" i]',
        'div[aria-label*="Manage app" i]',
        '[title*="Manage app" i]',
        // very defensive fallback for floating bottom-right launcher containers
        'div[style*="position: fixed"][style*="bottom"][style*="right"]'
      ];

      function hideOne(el) {
        if (!el) return;
        // Avoid nuking important dialogs; target only likely launcher-sized elements
        const txt = (el.innerText || '').trim().toLowerCase();
        const aria = (el.getAttribute('aria-label') || '').toLowerCase();
        const title = (el.getAttribute('title') || '').toLowerCase();
        const isManage =
          txt.includes('manage app') || aria.includes('manage app') || title.includes('manage app');

        if (isManage || el.matches('[data-testid*="manage" i], [data-testid*="app" i]')) {
          el.style.setProperty('display', 'none', 'important');
          el.style.setProperty('visibility', 'hidden', 'important');
          el.style.setProperty('opacity', '0', 'important');
          el.style.setProperty('pointer-events', 'none', 'important');
        }
      }

      function sweep() {
        sels.forEach(sel => {
          document.querySelectorAll(sel).forEach(hideOne);
        });

        // Also scan for exact text nodes rendered as buttons/chips
        document.querySelectorAll('button, div, span, a').forEach(el => {
          const t = (el.textContent || '').trim().toLowerCase();
          if (t === 'manage app' || t.includes('manage app')) hideOne(el);
        });
      }

      sweep();
      setInterval(sweep, 800); // keep hiding after rerenders

      const mo = new MutationObserver(sweep);
      mo.observe(document.documentElement, { childList: true, subtree: true, attributes: true });
    })();
    </script>
    """, height=0, width=0)
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
