import streamlit as st
def apply_global_styles():
    st.markdown("""
        <style>
        [data-testid="stToolbarActions"] {
            display: none !important;
        }
        </style>
    """, unsafe_allow_html=True)