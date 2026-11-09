import importlib
from pathlib import Path

DATA_PATH = Path(__file__).resolve().parent.parent / "team_data" / "ms_wip_teamdata.csv"
dashboard = importlib.import_module("dashboard")
# Initialize page
dashboard.create_page(
    "MS Heijunka Metrics Dashboard",
    "Metrics and trends from the individual teams.",
    DATA_PATH
)
