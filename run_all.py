#!/usr/bin/env python
import subprocess
import sys
from datetime import date
PYTHON_BIN = sys.executable
commands = [
    [PYTHON_BIN, "get_timeliness.py"],
    [PYTHON_BIN, "heijunka_new_layout.py", "--all"],
    [PYTHON_BIN, "collect_non_wip_new.py", "--config", "teams.json", "--metrics", "metrics.csv", "--all", "--out", "non_wip.csv"],
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
]
def run(cmd, *, cwd=None):
    print(f"Running: {' '.join(cmd)}")
    subprocess.run(cmd, cwd=cwd, check=True)
def has_git_changes(*, cwd=None) -> bool:
    r = subprocess.run(["git", "diff", "--quiet"], cwd=cwd)
    if r.returncode == 1:
        return True
    r = subprocess.run(["git", "diff", "--quiet", "--staged"], cwd=cwd)
    return r.returncode == 1
def main():
    repo_root = None
    try:
        for cmd in commands:
            run(cmd, cwd=repo_root)
        run(["git", "add", "-A"], cwd=repo_root)
        if has_git_changes(cwd=repo_root):
            msg = f"Automated update ({date.today().isoformat()})"
            run(["git", "commit", "-m", msg], cwd=repo_root)
            run(["git", "push"], cwd=repo_root)
            print("✅ Committed and pushed.")
        else:
            print("ℹ️ No git changes to commit. Skipping commit/push.")
        print("✅ All scripts completed successfully.")
    except subprocess.CalledProcessError as e:
        print(f"❌ Command failed (exit {e.returncode}): {e.cmd}")
        sys.exit(e.returncode)
if __name__ == "__main__":
    main()