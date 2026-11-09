import importlib
from pathlib import Path

DATA_PATH = Path(__file__).resolve().parent.parent / "team_data" / "neuroscience_wip_teamdata.csv"
dashboard = importlib.import_module("dashboard")
# Initialize page
dashboard.create_page(
    "NS Heijunka Metrics Dashboard",
    "Metrics and trends from the individual teams.",
    DATA_PATH
)
