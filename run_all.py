#!/usr/bin/env python
import subprocess
import sys
PYTHON_BIN = sys.executable
commands = [
    [PYTHON_BIN, "get_closures.py"],
    [PYTHON_BIN, "get_timeliness.py"],
    [PYTHON_BIN, "collect_metrics_dev.py", "--team", "ECT"],
    [PYTHON_BIN, "heijunka_new_layout.py", "--all"],
    [PYTHON_BIN, "collect_non_wip_new.py --config teams.json --metrics metrics.csv --all --out non_wip.csv"],
    [PYTHON_BIN, "push_selected_dates.py --date 2025-10-27"],
    [PYTHON_BIN, "push_selected_dates.py --date 2025-11-03"],
    [PYTHON_BIN, "push_selected_dates.py --date 2025-11-10"],
    [PYTHON_BIN, "push_selected_dates.py --date 2025-11-17"],
    [PYTHON_BIN, "push_selected_dates.py --date 2025-11-24"],
    [PYTHON_BIN, "push_selected_dates.py --date 2025-12-01"],
    [PYTHON_BIN, "push_selected_dates.py --date 2025-12-08"],
    [PYTHON_BIN, "push_selected_dates.py --date 2025-12-15"]
]
for cmd in commands:
    print(f"Running: {' '.join(cmd)}")
    result = subprocess.run(cmd)
    if result.returncode != 0:
        print(f"❌ Command failed: {' '.join(cmd)}")
        sys.exit(result.returncode)
print("✅ All scripts completed successfully.")