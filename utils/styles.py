import streamlit as st
def apply_global_styles():
    st.markdown("""
        <style>
        [data-testid="stToolbar"] {
            display: none !important;
        }
        [data-testid="collapsedControl"] {
            display: flex !important;
            visibility: visible !important;
            opacity: 1 !important;
            pointer-events: auto !important;
            position: fixed !important;
            top: 0.5rem !important;
            left: 0.5rem !important;
            z-index: 999999 !important;
        }
        </style>
    """, unsafe_allow_html=True)