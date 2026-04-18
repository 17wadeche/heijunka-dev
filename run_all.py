#!/usr/bin/env python
import os
import subprocess
import sys
from datetime import date
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(SCRIPT_DIR)
PYTHON_BIN = sys.executable
commands = [
    [PYTHON_BIN, "scrape_wip_iv.py", "--all"],
    [PYTHON_BIN, "build_iv_non_wip_activities.py", "--config", "teams.json", "--metrics", "IV_DATA\\metrics.csv", "--all", "--out", "IV_DATA\\non_wip.csv"],
    [PYTHON_BIN, "scrape_wip_ect.py"],
    [PYTHON_BIN, "build_ect_non_wip_activities.py"],
    [PYTHON_BIN, "push_selected_dates.py", "--date", "2025-10-27"],
    [PYTHON_BIN, "push_selected_dates.py", "--date", "2025-11-03"],
    [PYTHON_BIN, "push_selected_dates.py", "--date", "2025-11-10"],
    [PYTHON_BIN, "push_selected_dates.py", "--date", "2025-11-17"],
    [PYTHON_BIN, "push_selected_dates.py", "--date", "2025-11-24"],
    [PYTHON_BIN, "push_selected_dates.py", "--date", "2025-12-01"],
    [PYTHON_BIN, "push_selected_dates.py", "--date", "2025-12-08"],
    [PYTHON_BIN, "push_selected_dates.py", "--date", "2025-12-15"],
    [PYTHON_BIN, "push_selected_dates.py", "--date", "2025-12-22"],
    [PYTHON_BIN, "push_selected_dates.py", "--date", "2025-12-29"],
    [PYTHON_BIN, "push_selected_dates.py", "--date", "2026-01-05"],
    [PYTHON_BIN, "push_selected_dates.py", "--date", "2026-01-12"],
    [PYTHON_BIN, "push_selected_dates.py", "--date", "2026-01-19"],
    [PYTHON_BIN, "push_selected_dates.py", "--date", "2026-01-26"],
    [PYTHON_BIN, "push_selected_dates.py", "--date", "2026-02-02"],
    [PYTHON_BIN, "push_selected_dates.py", "--date", "2026-02-09"],
    [PYTHON_BIN, "push_selected_dates.py", "--date", "2026-02-16"],
    [PYTHON_BIN, "push_selected_dates.py", "--date", "2026-02-23"],
    [PYTHON_BIN, "push_selected_dates.py", "--date", "2026-03-02"],
    [PYTHON_BIN, "push_selected_dates.py", "--date", "2026-03-09"],
    [PYTHON_BIN, "push_selected_dates.py", "--date", "2026-03-16"],
    [PYTHON_BIN, "push_selected_dates.py", "--date", "2026-03-23"],
    [PYTHON_BIN, "push_selected_dates.py", "--date", "2026-03-30"],
    [PYTHON_BIN, "push_selected_dates.py", "--date", "2026-04-06"],
    [PYTHON_BIN, "push_selected_dates.py", "--date", "2026-04-13"],
    [PYTHON_BIN, "scrape_wip_crm.py"],
    [PYTHON_BIN, "build_crm_non_wip_activities.py"],
    [PYTHON_BIN, "scrape_wip_ms.py"],
    [PYTHON_BIN, "build_ms_non_wip_activities.py"],
    [PYTHON_BIN, "Scrape_wip_ns.py"],
    [PYTHON_BIN, "build_ns_non_wip_activities.py"],
]
def run(cmd, *, cwd=None):
    subprocess.run(cmd, cwd=cwd, check=True)
def has_git_changes(*, cwd=None) -> bool:
    r = subprocess.run(["git", "diff", "--quiet"], cwd=cwd)
    if r.returncode == 1:
        return True
    r = subprocess.run(["git", "diff", "--quiet", "--staged"], cwd=cwd)
    return r.returncode == 1
def main():
    repo_root = SCRIPT_DIR
    try:
        for cmd in commands:
            run(cmd, cwd=repo_root)
        run(["git", "add", "-A"], cwd=repo_root)
        if has_git_changes(cwd=repo_root):
            msg = f"Automated update ({date.today().isoformat()})"
            run(["git", "commit", "-m", msg], cwd=repo_root)
            run(["git", "push"], cwd=repo_root)        
    except subprocess.CalledProcessError as e:
        sys.exit(e.returncode)
if __name__ == "__main__":
    main()