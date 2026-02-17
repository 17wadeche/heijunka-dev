import streamlit as st
st.set_page_config(
    page_title="Enterprise Dashboard",
    layout="wide",
    initial_sidebar_state="expanded",
)
pg = st.navigation(
    [
        st.Page("heijunka-dash.py", title="Enterprise"),
        st.Page("pages/Interventional_Vascular.py", title="Interventional Vascular"),
    ]
)
pg.run()