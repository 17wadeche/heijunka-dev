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
]
for cmd in commands:
    print(f"Running: {' '.join(cmd)}")
    result = subprocess.run(cmd)
    if result.returncode != 0:
        print(f"❌ Command failed: {' '.join(cmd)}")
        sys.exit(result.returncode)
print("✅ All scripts completed successfully.")