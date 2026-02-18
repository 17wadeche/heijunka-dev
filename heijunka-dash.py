import streamlit as st
from utils.styles import apply_global_styles
st.set_page_config(
    page_title="Enterprise Dashboard",
    layout="wide",
    initial_sidebar_state="expanded",
)
apply_global_styles()
pg = st.navigation([
    st.Page("pages/Enterprise.py", title="Enterprise"),
    st.Page("pages/Interventional_Vascular.py", title="Interventional Vascular"),
])
pg.run()