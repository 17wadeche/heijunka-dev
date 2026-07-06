# heijunka-dash.py
import importlib
import sys
import time
from pathlib import Path
import streamlit as st
def _load_apply_global_styles(max_attempts: int = 5):
    repo_root = str(Path(__file__).resolve().parent)
    if repo_root not in sys.path:
        sys.path.insert(0, repo_root)
    last_error: BaseException | None = None
    for attempt in range(max_attempts):
        try:
            module = importlib.import_module("utils.styles")
            return module.apply_global_styles
        except (ImportError, KeyError) as exc:
            last_error = exc
            importlib.invalidate_caches()
            sys.modules.pop("utils.styles", None)
            sys.modules.pop("utils", None)
            if attempt < max_attempts - 1:
                time.sleep(0.2 * (attempt + 1))
    raise RuntimeError("Unable to import utils.styles after retrying") from last_error
apply_global_styles = _load_apply_global_styles()
st.set_page_config(
    page_title="Enterprise Dashboard",
    layout="wide",
    initial_sidebar_state="expanded",
)
apply_global_styles()
pg = st.navigation([
    st.Page("pages/Enterprise.py", title="Enterprise"),
    st.Page("pages/Interventional_Vascular.py", title="Interventional Vascular"),
    st.Page("pages/Neuroscience.py", title="Neuroscience"),
    st.Page("pages/Cardiac_Rhythm_Management.py", title="Cardiac Rhythm Management"),
    st.Page("pages/Medical_Surgical.py", title="Medical Surgical"),
])
pg.run()