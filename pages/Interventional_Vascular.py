import importlib
from pathlib import Path

DATA_PATH = Path(__file__).resolve().parent.parent / "team_data" / "iv_wip_teamdata.csv"
dashboard = importlib.import_module("dashboard")
# Initialize page
dashboard.create_page(
    "IV Heijunka Metrics Dashboard",
    "Metrics and trends from the individual teams.",
    DATA_PATH
)
