import streamlit as st
def apply_global_styles():
    st.markdown("""
        <style>
        /* Hide the top-right toolbar */
        [data-testid="stToolbar"] {
            display: none !important;
        }
        /* Ensure sidebar collapse/expand button stays visible */
        [data-testid="collapsedControl"] {
            display: flex !important;
            visibility: visible !important;
        }
        </style>
    """, unsafe_allow_html=True)