import re
import sys
import argparse
import time
import glob
from pathlib import Path
from datetime import datetime as _dt, date as _date, timedelta
from dateutil import parser as dateparser
import pandas as pd
import csv
import numpy as np
import shutil
import subprocess
from openpyxl import load_workbook
import tempfile, uuid
import json
REPO_DIR = Path(r"C:\heijunka-dev")
REPO_CSV = REPO_DIR / "metrics_aggregate_dev.csv"
GIT_BRANCH = "main"
TIMELINESS_CSV = REPO_DIR / "timeliness.csv"
EXCLUDED_SOURCE_FILES = {
    r"C:\Users\wadec8\Medtronic PLC\SVT PXM Team - Archived Heijunka\SVT Future Heijunka.xlsm",
    r"c:\Users\wadec8\Medtronic PLC\CQXM - IV Resource Site - COS Supportive Materials\Heijunka Population & SW.xlsx",
    r"c:\Users\wadec8\Medtronic PLC\CQXM - IV Resource Site - COS Supportive Materials\PVH Future Heijunka (UPDATE) Template.xlsm",
    r"c:\Users\wadec8\Medtronic PLC\CQXM - IV Resource Site - COS Supportive Materials\Versatility Matrix_June 2024.xlsx",
    r"c:\Users\wadec8\Medtronic PLC\CQXM - IV Resource Site - Heijunka\CRDN Heijunka Template.xlsm",
    r"c:\Users\wadec8\Medtronic PLC\CAS Virtual VMB - PA Board\Scheduling Assistant.xlsx",
    r"c:\Users\wadec8\Medtronic PLC\CAS Virtual VMB - PA Board\Aging FACs 26 09.xlsx",
    r"c:\Users\wadec8\Medtronic PLC\CAS Virtual VMB - PA Board\Tier 1 Escalations and Recognition.pptx",
    r"c:\Users\wadec8\Medtronic PLC\CAS Virtual VMB - PA Board\Time Studies.docx",
    r"c:\Users\wadec8\Medtronic PLC\CAS Virtual VMB - PA Board\Updated Cryo Electronic data sheet.xlsx",
    r"c:\Users\wadec8\Medtronic PLC\CQXM - Aortic - Heijunka\Saved Heijunkas\Aortic Heijunka Template 2.0.xlsm",
    r"c:\Users\wadec8\Medtronic PLC\CQXM - Aortic - Heijunka\Saved Heijunkas\Aortic Heijunka Template 3.0.xlsm",
    r"c:\Users\wadec8\Medtronic PLC\Doran, Elaine - Heijunka Production Analysis\Archived Heijunka\ECT Future Heijunka 04 August 2025.xlsm",
    r"c:\Users\wadec8\Medtronic PLC\CQXM - IV Resource Site - COS Supportive Materials\Archive_Production Analysis\PVEV Heijunka_2020_2022_Archived.xlsm"
}
EXCLUDED_DIRS = {
    r"C:\Users\wadec8\Medtronic PLC\TCT CQXM - 1 WIP and Schedule\Weekly Heijunka Archived",
    r"C:\Users\wadec8\Medtronic PLC\TCT CQXM - 1 WIP and Schedule\Clinical",
    r"C:\Users\wadec8\Medtronic PLC\TCT CQXM - 1 WIP and Schedule\Commercial",
    r"C:\Users\wadec8\Medtronic PLC\TCT CQXM - 1 WIP and Schedule\Remediation",
    r"C:\Users\wadec8\Medtronic PLC\TCT CQXM - 1 WIP and Schedule\WIP Blitz Power Hour",
    r"C:\Users\wadec8\Medtronic PLC\SVT PXM Team - Archived Heijunka",
    r"c:\Users\wadec8\Medtronic PLC\Doran, Elaine - Heijunka Production Analysis\Heijunka Template",
    r"c:\Users\wadec8\Medtronic PLC\CQXM - IV Resource Site - COS Supportive Materials\Archive",
    r"c:\Users\wadec8\Medtronic PLC\CQXM - IV Resource Site - COS Supportive Materials\Finding Work Tool",
    r"c:\Users\wadec8\Medtronic PLC\CQXM - IV Resource Site - COS Supportive Materials\Archive_Production Analysis\2023",
    r"c:\Users\wadec8\Medtronic PLC\CQXM - IV Resource Site - COS Supportive Materials\Archive_Production Analysis\2022",
    r"c:\Users\wadec8\Medtronic PLC\CQXM - IV Resource Site - COS Supportive Materials\Archive_Production Analysis\2021",
    r"c:\Users\wadec8\Medtronic PLC\CQXM - IV Resource Site - COS Supportive Materials\Archive_Production Analysis\2020",
    r"c:\Users\wadec8\Medtronic PLC\CQXM - IV Resource Site - COS Supportive Materials\Archive_Production Analysis\2019",
    r"c:\Users\wadec8\Medtronic PLC\CQXM - IV Resource Site - COS Supportive Materials\Archive_Production Analysis\Prod Analysis Drafts for upcoming weeks"
    r"c:\Users\wadec8\Medtronic PLC\CQXM - IV Resource Site - COS Supportive Materials\PVH Smartsheet Gameboard",
    r"c:\Users\wadec8\Medtronic PLC\CQXM - IV Resource Site - COS Supportive Materials\Standard Works",
    r"c:\Users\wadec8\Medtronic PLC\CQXM - IV Resource Site - COS Supportive Materials\Upcoming Weeks Heijunka Drafts",
    r"c:\Users\wadec8\Medtronic PLC\CQXM - Aortic - Heijunka\Saved Heijunkas\Templates",
    r"c:\Users\wadec8\Medtronic PLC\CQXM - IV Resource Site - Heijunka\Archived\Archived Heijunka 2024"
}
EXCLUDED_DIRS = {s.lower().rstrip("\\").replace("/", "\\") for s in EXCLUDED_DIRS}
EXCLUDED_SOURCE_FILES = {s.lower().replace("/", "\\") for s in EXCLUDED_SOURCE_FILES}
def _is_excluded_path(p: Path) -> bool:
    try:
        sp = str(p).lower().replace("/", "\\")
        if sp in EXCLUDED_SOURCE_FILES:
            return True
        for d in EXCLUDED_DIRS:
            if sp == d or sp.startswith(d + "\\"):
                return True
        if p.name.lower() == "svt future heijunka.xlsm":
            return True
    except Exception:
        pass
    return False
TEAM_CONFIG = [
    {
        "name": "Aortic",
        "root": r"C:\Users\wadec8\Medtronic PLC\CQXM - Aortic - Heijunka",
        "pattern": "*.xls*",
        "period": {"sheet": "#12 Production Analysis", "cell": "C4"},
        "cells": {
            "Individual (WIP-Non WIP)": {
                "Total Available Hours": "I39",
            }
        },
        "sum_columns": {
            "#12 Production Analysis": {
                "Target Output": {"col": "E", "row_start": 7, "row_end": 199},
                "Actual Output": {"col": "I", "row_start": 7, "row_end": 199},
            }
        },
        "unique_key": ["team", "period_date"],
    },
    {
        "name": "ECT",
        "root": r"C:\Users\wadec8\Medtronic PLC\Doran, Elaine - Heijunka Production Analysis",
        "pattern": "*.xls*",
        "period": {"sheet": "#12 Production Analysis", "cell": "C4"},
        "cells": {
            "Individual (WIP-Non WIP)": {
                "Total Available Hours": "I39",
                "Completed Hours":       "I40",
            }
        },
        "sum_columns": {
            "#12 Production Analysis": {
                "Target Output": {"col": "F", "row_start": 7, "row_end": 199},
                "Actual Output": {"col": "I", "row_start": 7, "row_end": 199},
                "Completed Hours Detail": {"col": "G", "row_start": 7, "row_end": 199, "divide": 60, "skip_hidden": True},
            }
        },
        "unique_key": ["team", "period_date"],
    },
    {
        "name": "ECT",
        "root": r"C:\Users\wadec8\Medtronic PLC\Doran, Elaine - Heijunka Production Analysis\Archived Heijunka",
        "pattern": "*.xls*",
        "period": {"sheet": "#12 Production Analysis", "cell": "C4"},
        "cells": {
            "Individual (WIP-Non WIP)": {
                "Total Available Hours": "I39",
                "Completed Hours":       "I40",
            }
        },
        "sum_columns": {
            "#12 Production Analysis": {
                "Target Output": {"col": "F", "row_start": 7, "row_end": 199},
                "Actual Output": {"col": "I", "row_start": 7, "row_end": 199},
                "Completed Hours Detail": {"col": "G", "row_start": 7, "row_end": 199, "divide": 60, "skip_hidden": True},
            }
        },
        "unique_key": ["team", "period_date"],
    },
    {
        "name": "CRDN",
        "root": r"C:\Users\wadec8\Medtronic PLC\CQXM - IV Resource Site - Heijunka",
        "pattern": "*.xls*",
        "period": {"sheet": "#12 Production Analysis", "cell": "C4"},
        "cells": {
            "Individual (WIP-Non WIP)": {
                "Total Available Hours": "I39",
                "Completed Hours":       "I40",
            }
        },
        "sum_columns": {
            "#12 Production Analysis": {
                "Target Output": {"col": "F", "row_start": 7, "row_end": 199},
                "Actual Output": {"col": "I", "row_start": 7, "row_end": 199},
                "Completed Hours Detail": {"col": "G", "row_start": 7, "row_end": 199, "divide": 60, "skip_hidden": True},
            }
        },
        "unique_key": ["team", "period_date"],
    },
    {
        "name": "CRDN",
        "root": r"C:\Users\wadec8\Medtronic PLC\CQXM - IV Resource Site - Heijunka\Archived\Archived Heijunka 2025",
        "pattern": "*.xls*",
        "period": {"sheet": "#12 Production Analysis", "cell": "C4"},
        "cells": {
            "Individual (WIP-Non WIP)": {
                "Total Available Hours": "I39",
                "Completed Hours":       "I40",
            }
        },
        "sum_columns": {
            "#12 Production Analysis": {
                "Target Output": {"col": "F", "row_start": 7, "row_end": 199},
                "Actual Output": {"col": "I", "row_start": 7, "row_end": 199},
                "Completed Hours Detail": {"col": "G", "row_start": 7, "row_end": 199, "divide": 60, "skip_hidden": True},
            }
        },
        "unique_key": ["team", "period_date"],
    },
    {
        "name": "PVH",
        "root": r"C:\Users\wadec8\Medtronic PLC\CQXM - IV Resource Site - COS Supportive Materials",
        "pattern": "*.xls*",
        "period": {"sheet": "#12 Production Analysis", "cell": "C4"},
        "cells": {
            "Individual (WIP-Non WIP)": {
                "Total Available Hours": "I39",
                "Completed Hours":       "I40",
            }
        },
        "sum_columns": {
            "#12 Production Analysis": {
                "Target Output": {"col": "F", "row_start": 7, "row_end": 199},
                "Actual Output": {"col": "I", "row_start": 7, "row_end": 199},
                "Completed Hours Detail": {"col": "G", "row_start": 7, "row_end": 199, "divide": 60, "skip_hidden": True},
            }
        },
        "unique_key": ["team", "period_date"],
    },
    {
        "name": "PVH",
        "root": r"C:\Users\wadec8\Medtronic PLC\CQXM - IV Resource Site - COS Supportive Materials\Archive_Production Analysis\2025",
        "pattern": "*.xls*",
        "period": {"sheet": "#12 Production Analysis", "cell": "C4"},
        "cells": {
            "Individual (WIP-Non WIP)": {
                "Total Available Hours": "I39",
                "Completed Hours":       "I40",
            }
        },
        "sum_columns": {
            "#12 Production Analysis": {
                "Target Output": {"col": "F", "row_start": 7, "row_end": 199},
                "Actual Output": {"col": "I", "row_start": 7, "row_end": 199},
                "Completed Hours Detail": {"col": "G", "row_start": 7, "row_end": 199, "divide": 60, "skip_hidden": True},
            }
        },
        "unique_key": ["team", "period_date"],
    },
    {
        "name": "TCT Commercial",
        "root": r"C:\Users\wadec8\Medtronic PLC\TCT CQXM - Weekly Heijunka Archived",
        "pattern": "*.xlsb",
        "period": {"sheet": "#10 WIP Analysis", "cell": "D3"},
        "cells": {
            "Individual(WIP-Non WIP)": {
                "Total Available Hours": "I69",
                "Completed Hours": "I70",
            }
        },
        "sum_columns": {
            "Commercial #12 Prod Analysis": {
                "Target Output": "G",
                "Actual Output": "J",
            }
        },
    },
    {
        "name": "TCT Clinical",
        "root": r"C:\Users\wadec8\Medtronic PLC\TCT CQXM - Weekly Heijunka Archived",
        "pattern": "*.xlsb",
        "period": {"sheet": "#10 WIP Analysis", "cell": "D3"},
        "cells": {
            "Individual(WIP-Non WIP)": {
                "Total Available Hours": "AG69",
                "Completed Hours": "AG70",
            }
        },
        "sum_columns": {
            "Clinical #12 Prod Analysis": {
                "Target Output": "G",
                "Actual Output": "J",
            }
        },
    },
    {
        "name": "TCT Commercial",
        "root": r"C:\Users\wadec8\Medtronic PLC\TCT CQXM - 1 WIP and Schedule",
        "pattern": "*.xlsb",
        "period": {"sheet": "#10 WIP Analysis", "cell": "D3"},
        "cells": {
            "Individual(WIP-Non WIP)": {
                "Total Available Hours": "I69",
                "Completed Hours": "I70",
            }
        },
        "sum_columns": {
            "Commercial #12 Prod Analysis": {
                "Target Output": "G",
                "Actual Output": "J",
            }
        },
    },{
        "name": "TCT Clinical",
        "root": r"C:\Users\wadec8\Medtronic PLC\TCT CQXM - 1 WIP and Schedule",
        "pattern": "*.xlsb",
        "period": {"sheet": "#10 WIP Analysis", "cell": "D3"},
        "cells": {
            "Individual(WIP-Non WIP)": {
                "Total Available Hours": "AG69",
                "Completed Hours": "AG70",
            }
        },
        "sum_columns": {
            "Clinical #12 Prod Analysis": {
                "Target Output": "G",
                "Actual Output": "J",
            }
        },
    },
    {
        "name": "SVT",
        "pss_mode": True,
        "file_glob": r"C:\Users\wadec8\Medtronic PLC\SVT PXM Team - Heijunka_Schedule_Finding Work\SVT Heijunka*.xls*",
        "period": {"sheet": "#12 Production Analysis", "cell": "C4"},
        "cells_by_sheet": {
            "Individual": {
                "Total Available Hours": ["I24", "I27", "I30", "I33", "I36"],
            }
        },
        "sum_columns": {
            "#12 Production Analysis": {
                "Target Output": "F",
                "Actual Output": "I",
                "Completed Hours": {
                    "col": "G", "row_start": 1, "row_end": 200, "divide": 60, "skip_hidden": True
                },
            }
        },
        "unique_key": ["team", "period_date"],
    },
    {
        "name": "PSS",
        "pss_mode": True,
        "file": r"C:\Users\wadec8\Medtronic PLC\PSS Sharepoint - Documents\PSS_Heijunka.xlsm",
        "dropdown_iter": {
            "sheet": "Previous Weeks",
            "cell":  "A2",
            "source_hint": None
        },
        "cells": {
            "Total Available Hours": "R64",
            "Completed Hours":      "R54",
        },
        "sum_pairs": {
            "Target Output": ["X10", "Z10"],
            "Actual Output": ["X5",  "Z5"],
        },
        "unique_key": ["team", "period_date"],
    },
    {
        "name": "PH",
        "ph_mode": True,
        "file": r"C:\Users\wadec8\Medtronic PLC\Customer Quality Pelvic Health - Daily Tracker\PH Cell Heijunka.xlsx",
        "unique_key": ["team", "period_date"],
    },
]
SKIP_PATTERNS = [r"~\$", r"\.tmp$"]
DATE_REGEXES = [
    r"\b\d{4}[-_]\d{2}[-_]\d{2}\b",
    r"\b\d{8}\b",
    r"\b\d{1,2}[-_]\d{1,2}[-_]\d{2,4}\b",
    r"\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*[-_ ]\d{1,2}[-_, ]\d{2,4}\b",
]
USE_FILE_MTIME_IF_NO_DATE = True
OUT_XLSX = Path.cwd() / "metrics_aggregate_dev.xlsx"
OUT_CSV  = Path.cwd() / "metrics_aggregate_dev.csv"
def looks_like_temp(name: str) -> bool:
    return any(re.search(pat, name, flags=re.IGNORECASE) for pat in SKIP_PATTERNS)
def parse_date_from_text(text: str):
    for rx in DATE_REGEXES:
        m = re.search(rx, text, flags=re.IGNORECASE)
        if m:
            try:
                return dateparser.parse(m.group(0), dayfirst=False, yearfirst=False).date()
            except Exception:
                pass
    return None
def _filter_future_periods(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "period_date" not in df.columns:
        return df
    today = pd.Timestamp.today().normalize()
    d = pd.to_datetime(df["period_date"], errors="coerce").dt.normalize()
    keep = d.isna() | (d <= today)      # keep blanks and dates up to today
    return df.loc[keep].copy()
def _filter_ph_zero_hours(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "team" not in df.columns or "Total Available Hours" not in df.columns:
        return df
    tah = pd.to_numeric(df["Total Available Hours"], errors="coerce")
    keep = ~((df["team"] == "PH") & (tah.isna() | (tah <= 0)))
    return df.loc[keep].copy()
def _filter_ect_min_year(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "team" not in df.columns or "period_date" not in df.columns:
        return df
    cutoff = pd.Timestamp("2024-08-19")
    d = pd.to_datetime(df["period_date"], errors="coerce")
    keep = ~((df["team"] == "ECT") & (d < cutoff))
    return df.loc[keep].copy()
def _excel_serial_to_date(n) -> _date | None:
    try:
        return (_dt(1899, 12, 30) + timedelta(days=float(n))).date()
    except Exception:
        return None
YEAR_RX = re.compile(r"\b(?:19|20)\d{2}\b")
def _coerce_to_date_for_filter2(v, require_explicit_year: bool = False) -> _date | None:
    if isinstance(v, _dt):
        return v.date()
    if isinstance(v, _date):
        return v
    if isinstance(v, (int, float)):
        d = _excel_serial_to_date(v)
        if d and _dt(1900, 1, 1).date() <= d <= _dt(2100, 1, 1).date():
            return d
        return None
    s = str(v)
    if require_explicit_year and not YEAR_RX.search(s):
        return None
    try:
        return dateparser.parse(s, dayfirst=False, yearfirst=False).date()
    except Exception:
        return None
def _coerce_to_date_for_filter(v) -> _date | None:
    if isinstance(v, _dt):
        return v.date()
    if isinstance(v, _date):
        return v
    if isinstance(v, (int, float)):
        d = _excel_serial_to_date(v)
        if d and _dt(1900, 1, 1).date() <= d <= _dt(2100, 1, 1).date():
            return d
        return None
    try:
        d = dateparser.parse(str(v)).date()
        return d
    except Exception:
        return None
def _git(args: list[str], cwd: Path) -> tuple[int, str, str]:
    p = subprocess.Popen(args, cwd=str(cwd), stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    out, err = p.communicate()
    cmd = " ".join(args)
    if p.returncode != 0:
        print(f"[git] {cmd} -> code={p.returncode}\nSTDERR: {err}\nSTDOUT: {out}")
    else:
        if err.strip():
            print(f"[git] {cmd} (ok) stderr: {err.strip()}")
    return p.returncode, out.strip(), err.strip()
def _ensure_git_identity(repo_dir: Path):
    code, out, _ = _git(["git", "config", "--get", "user.email"], repo_dir)
    if code != 0 or not out:
        _git(["git", "config", "user.email", "heijunka-bot@example.com"], repo_dir)
    code, out, _ = _git(["git", "config", "--get", "user.name"], repo_dir)
    if code != 0 or not out:
        _git(["git", "config", "user.name", "Heijunka Bot"], repo_dir)
def git_autocommit_and_push(repo_dir: Path, file_path: Path, branch: str = "main"):
    if not (repo_dir / ".git").exists():
        print(f"[WARN] {repo_dir} is not a git repo; skipping auto-commit.")
        return
    if not file_path.exists():
        print(f"[WARN] {file_path} not found to commit.")
        return
    _git(["git", "config", "--global", "--add", "safe.directory", str(repo_dir)], repo_dir)
    _ensure_git_identity(repo_dir)
    code, _, _ = _git(["git", "--version"], repo_dir)
    if code != 0:
        print("[WARN] Git not available on PATH. Install Git or adjust PATH.")
        return
    code, remotes, _ = _git(["git", "remote"], repo_dir)
    if code != 0 or not remotes.strip():
        print("[WARN] No git remote configured (e.g., 'origin'); cannot push.")
        return
    remote = "origin" if "origin" in remotes.split() else remotes.split()[0]
    code, heads, _ = _git(["git", "branch", "--list", branch], repo_dir)
    if heads.strip():
        _git(["git", "checkout", branch], repo_dir)
    else:
        _git(["git", "checkout", "-B", branch], repo_dir)
    _, dirty, _ = _git(["git", "status", "--porcelain"], repo_dir)
    if dirty.strip():
        print("[git] Working tree has local changes; using --autostash during pull.")
    _git(["git", "fetch", remote], repo_dir)
    code, _, _ = _git(["git", "rev-parse", "--abbrev-ref", "--symbolic-full-name", "@{u}"], repo_dir)
    if code != 0:
        _git(["git", "branch", "--set-upstream-to", f"{remote}/{branch}", branch], repo_dir)
    code, _, _ = _git(["git", "pull", "--rebase", "--autostash", remote, branch], repo_dir)
    if code != 0:
        print("[WARN] 'git pull --rebase --autostash' failed; skipping commit to avoid conflicts. Resolve repo state manually.")
        return
    try:
        rel = file_path.relative_to(repo_dir).as_posix()
    except Exception:
        rel = str(file_path)
    _git(["git", "add", "--", rel], repo_dir)
    code, status, _ = _git(["git", "status", "--porcelain", "--", rel], repo_dir)
    if code != 0:
        print("[WARN] Could not check git status; aborting commit.")
        return
    if not status.strip():
        print("[git] No changes in CSV; skipping push.")
        return
    msg = f"Auto-update {rel} at {_dt.now().isoformat(timespec='seconds')}"
    code, _, _ = _git(["git", "commit", "-m", msg], repo_dir)
    if code != 0:
        print("[WARN] git commit failed; see logs above.")
        return
    code, _, _ = _git(["git", "push", "-u", remote, branch], repo_dir)
    if code != 0:
        print("[WARN] git push failed; see logs above. Check credentials / PAT / VPN.")
    else:
        print("[git] Pushed CSV update to", f"{remote}/{branch}")
def _ph_values_by_person(ws, col_end: str) -> tuple[dict, float | None, float | None]:
    def _to_float(v):
        if v is None: 
            return None
        try:
            return float(str(v).replace(",", "").strip())
        except Exception:
            return None
    rng_names = ws.Range(f"B30:{col_end}30").Value
    rng_actual = ws.Range(f"B50:{col_end}50").Value
    rng_avail  = ws.Range(f"B59:{col_end}59").Value
    if isinstance(rng_names, (tuple, list)) and isinstance(rng_names[0], (tuple, list)):
        names = list(rng_names[0])
    else:
        names = list(rng_names if isinstance(rng_names, (tuple, list)) else [rng_names])
    if isinstance(rng_actual, (tuple, list)) and isinstance(rng_actual[0], (tuple, list)):
        actuals = list(rng_actual[0])
    else:
        actuals = list(rng_actual if isinstance(rng_actual, (tuple, list)) else [rng_actual])
    if isinstance(rng_avail, (tuple, list)) and isinstance(rng_avail[0], (tuple, list)):
        avails = list(rng_avail[0])
    else:
        avails = list(rng_avail if isinstance(rng_avail, (tuple, list)) else [rng_avail])
    m = max(len(names), len(actuals), len(avails))
    names  += [None] * (m - len(names))
    actuals += [None] * (m - len(actuals))
    avails  += [None] * (m - len(avails))
    per = {}
    tot_actual = 0.0
    tot_avail  = 0.0
    any_actual = False
    any_avail  = False
    for nm, a, t in zip(names, actuals, avails):
        nm_str = (str(nm).strip() if nm is not None else "")
        if not nm_str:
            continue
        a_f = _to_float(a)
        t_f = _to_float(t)
        if a_f is not None:
            any_actual = True
            tot_actual += a_f
        if t_f is not None:
            any_avail = True
            tot_avail += t_f
        per[nm_str] = {"actual": (a_f if a_f is not None else 0.0),
                       "available": (t_f if t_f is not None else 0.0)}
    return per, (tot_actual if any_actual else None), (tot_avail if any_avail else None)
def _to_excel_com_value(v):
    if isinstance(v, (int, float, str)):
        return v
    if isinstance(v, _dt):
        return v
    if isinstance(v, _date):
        return _dt(v.year, v.month, v.day)
    return str(v)
def collect_ph_team(cfg: dict) -> list[dict]:
    file_path = Path(cfg.get("file", ""))
    team_name = cfg.get("name", "PH")
    src_display = str(file_path) if file_path else (cfg.get("file") or "")
    if not file_path or not file_path.exists():
        return [{"team": team_name, "source_file": src_display, "error": "PH file not found"}]
    try:
        import win32com.client as win32
        import pywintypes
    except Exception:
        return [{"team": team_name, "source_file": src_display,
                 "error": "pywin32 not installed; run 'pip install pywin32' to enable PH mode"}]
    import tempfile, uuid
    def _to_float(v):
        if v is None:
            return None
        try:
            return float(str(v).replace(",", "").strip())
        except Exception:
            return None
    rows: list[dict] = []
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    try:
        excel.DisplayAlerts = False
        try:
            excel.ScreenUpdating = False
        except Exception:
            pass
        try:
            excel.EnableEvents = False
        except Exception:
            pass
        wb = None
        tmp_copy = None
        open_path = str(file_path)
        try:
            tmp_copy = Path(tempfile.gettempdir()) / f"ph_heijunka_{uuid.uuid4().hex}{file_path.suffix}"
            shutil.copy2(file_path, tmp_copy)  # also forces hydration if available locally
            open_path = str(tmp_copy)
        except Exception:
            open_path = str(file_path)
        last_exc = None
        for _attempt in range(3):
            try:
                wb = excel.Workbooks.Open(
                    Filename=open_path,
                    ReadOnly=True,
                    UpdateLinks=0,                  # 0 = don't update
                    IgnoreReadOnlyRecommended=True,
                    Local=True                      # helps with localized paths
                )
                break
            except Exception as e:
                last_exc = e
                time.sleep(1.0)
        if wb is None:
            rows.append({"team": team_name, "source_file": src_display,
                        "error": f"PH mode failed after retries: {last_exc}"})
            return rows
        try:
            sheet_count = int(wb.Worksheets.Count)
            today_d = _dt.today().date()
            for idx in range(1, sheet_count + 1):
                ws = wb.Worksheets(idx)
                try:
                    _vis = int(getattr(ws, "Visible", -1))
                except Exception:
                    _vis = -1
                if _vis in (0, 2):
                    del ws
                    continue
                name = str(ws.Name).strip()
                if not re.search(r"\b(?:19|20)\d{4}\b", name) and not re.search(r"\b(?:19|20)\d{2}\b", name):
                    del ws
                    continue
                period_date = None
                try:
                    period_date = _coerce_to_date_for_filter2(name, require_explicit_year=True)
                except Exception:
                    period_date = None
                if not period_date or period_date > today_d:
                    del ws
                    continue
                try:
                    new_layout = bool(period_date and period_date > _dt(2025, 8, 30).date())
                    if new_layout:
                        ao  = (_to_float(ws.Range("Z2").Value)  or 0.0) + (_to_float(ws.Range("AB2").Value) or 0.0)
                        ch  = (_to_float(ws.Range("Z4").Value)  or 0.0) + (_to_float(ws.Range("AB4").Value) or 0.0)
                        to_ = (_to_float(ws.Range("Z7").Value)  or 0.0) + (_to_float(ws.Range("AB7").Value) or 0.0)
                        tah =  _to_float(ws.Range("T59").Value)
                        hc_end = "R"
                        uplh_wp1 = _to_float(ws.Range("Z5").Value)
                        uplh_wp2 = _to_float(ws.Range("AB5").Value)
                    else:
                        ao  = (_to_float(ws.Range("Y2").Value)  or 0.0) + (_to_float(ws.Range("AA2").Value) or 0.0)
                        ch  = (_to_float(ws.Range("Y4").Value)  or 0.0) + (_to_float(ws.Range("AA4").Value) or 0.0)
                        to_ = (_to_float(ws.Range("Y7").Value)  or 0.0) + (_to_float(ws.Range("AA7").Value) or 0.0)
                        tah =  _to_float(ws.Range("S59").Value)
                        hc_end = "Q"
                        uplh_wp1 = _to_float(ws.Range("Y5").Value)
                        uplh_wp2 = _to_float(ws.Range("AA5").Value)
                    per_person, sum_actual_row50, sum_avail_row59 = _ph_values_by_person(ws, hc_end)
                    if sum_actual_row50 is not None:
                        ch = sum_actual_row50
                    if sum_avail_row59 is not None:
                        tah = sum_avail_row59
                    if tah is None or float(tah) <= 0.0:
                        del ws
                        continue
                    try:
                        ppl_in_wip_list = [
                            nm for nm, vals in (per_person or {}).items()
                            if vals is not None and float(vals.get("actual") or 0) > 0
                        ]
                    except Exception:
                        ppl_in_wip_list = []
                    ppl_in_wip = ", ".join(ppl_in_wip_list) if ppl_in_wip_list else ""
                    try:
                        hc = _count_ph_hc_in_wip_com(ws, col_end=hc_end)
                    except Exception:
                        hc = None
                    rows.append({
                        "team": team_name,
                        "source_file": src_display,
                        "period_date": period_date,
                        "Total Available Hours": tah,
                        "Completed Hours": ch,
                        "Target Output": to_,
                        "Actual Output": ao,
                        "HC in WIP": hc,
                        "UPLH WP1": uplh_wp1,
                        "UPLH WP2": uplh_wp2,
                        "People in WIP": ppl_in_wip,
                        "Person Hours": json.dumps(per_person, ensure_ascii=False)
                    })
                finally:
                    del ws
        except Exception as e:
            rows.append({"team": team_name, "source_file": src_display, "error": f"PH mode failed: {e}"})
        finally:
            try:
                if wb is not None:
                    try:
                        wb.Close(SaveChanges=0)  # xlDoNotSaveChanges
                    except pywintypes.com_error:
                        pass
                    except Exception:
                        pass
            finally:
                try:
                    excel.Quit()
                except pywintypes.com_error:
                    pass
                except Exception:
                    pass
                if tmp_copy and Path(tmp_copy).exists():
                    try:
                        Path(tmp_copy).unlink()
                    except Exception:
                        pass
    except Exception as e:
        rows.append({"team": team_name, "source_file": src_display, "error": f"PH mode init failed: {e}"})
    DEBUG_PH = True
    if DEBUG_PH:
        picked = next((r for r in rows if isinstance(r, dict) and r.get("period_date") is not None), None)
        if picked:
            print("[PH] picked:", picked["period_date"], "from", src_display)
        else:
            msg = rows[0].get("error") if rows and isinstance(rows[0], dict) else "no valid rows"
            print(f"[PH] no valid week found ({msg})")
    return rows
def _get_validation_values_via_com(excel, ws, a1_cell: str, source_hint: str | None = None):
    vals = []
    formula = source_hint
    try:
        v = ws.Range(a1_cell).Validation
        if not formula:
            if int(v.Type) == 3 and v.Formula1:
                formula = v.Formula1
    except Exception:
        pass
    if not formula:
        return vals
    if isinstance(formula, str) and not formula.startswith("="):
        parts = [p.strip() for p in formula.split(",") if p.strip()]
        if parts:
            return parts
        formula = "=" + formula
    try:
        evaluated = excel.Evaluate(formula)
        try:
            _ = evaluated.Address
            rng_vals = evaluated.Value
            if isinstance(rng_vals, (tuple, list)):
                for row in (rng_vals if isinstance(rng_vals[0], (tuple, list)) else [rng_vals]):
                    for v in (row if isinstance(row, (tuple, list)) else [row]):
                        vals.append(v)
            else:
                vals.append(rng_vals)
        except Exception:
            if isinstance(evaluated, (tuple, list)):
                for row in (evaluated if isinstance(evaluated[0], (tuple, list)) else [evaluated]):
                    for v in (row if isinstance(row, (tuple, list)) else [row]):
                        vals.append(v)
            else:
                vals.append(evaluated)
    except Exception:
        if formula.startswith("="):
            raw = formula[1:]
            if "," in raw:
                vals = [p.strip() for p in raw.split(",") if p.strip()]
    out = []
    for v in vals:
        if v is None:
            continue
        if isinstance(v, str) and not v.strip():
            continue
        out.append(v)
    return out
def read_one_cell_openpyxl(ws, a1: str):
    r, c = a1_to_rowcol(a1)
    vals = list(ws.iter_rows(min_row=r, max_row=r, min_col=c, max_col=c, values_only=True))
    return vals[0][0] if vals else None
def read_one_cell_xlsb(file_path: Path, sheet: str, a1: str):
    from pandas import read_excel
    df = read_excel(file_path, sheet_name=sheet, engine="pyxlsb", header=None)
    m = re.fullmatch(r"([A-Za-z]+)(\d+)", a1.strip())
    if not m:
        return None
    col_letters, row_str = m.groups()
    r = int(row_str) - 1
    c = col_letter_to_index(col_letters) - 1
    try:
        return df.iat[r, c]
    except Exception:
        return None
def _read_cells_from_excel_com(ws, addr_map: dict[str, str]) -> dict:
    out = {}
    for key, a1 in (addr_map or {}).items():
        out[key] = ws.Range(a1).Value
    return out
def _sum_pairs_from_excel_com(ws, sum_pairs: dict[str, list[str]]) -> dict:
    out = {}
    for key, addrs in (sum_pairs or {}).items():
        total, any_vals = 0.0, False
        for a1 in addrs:
            v = ws.Range(a1).Value
            if v is None: 
                continue
            try:
                total += float(str(v).replace(",", "").strip()); any_vals = True
            except Exception:
                pass
        out[key] = total if any_vals else None
    return out
def _count_pss_hc_in_wip_com(ws, row_indices=None, col_start="B", col_end="O") -> int:
    if row_indices is None:
        row_indices = [34, 35, 38, 39, 42, 43, 46, 47, 50, 51]
    rmin, rmax = min(row_indices), max(row_indices)
    rng = ws.Range(f"{col_start}{rmin}:{col_end}{rmax}").Value  
    if not isinstance(rng, (tuple, list)) or not isinstance(rng[0], (tuple, list)):
        rng = (rng,)
    wanted_offsets = {ri - rmin for ri in row_indices}
    def _is_zero_or_blank(v):
        if v is None:
            return True
        s = str(v).strip()
        if not s:
            return True
        try:
            return float(s) == 0.0
        except Exception:
            return False
    cols = len(rng[0])
    count = 0
    for c in range(cols):
        any_nonzero = False
        for r_off in wanted_offsets:
            try:
                v = rng[r_off][c]
            except Exception:
                v = None
            if not _is_zero_or_blank(v):
                any_nonzero = True
                break
        if any_nonzero:
            count += 1
    return count
def _count_ph_hc_in_wip_com(ws, row_indices=None, col_start="B", col_end="Q") -> int:
    if row_indices is None:
        row_indices = [31, 32, 35, 36, 39, 40, 43, 44, 47, 48]
    rmin, rmax = min(row_indices), max(row_indices)
    rng = ws.Range(f"{col_start}{rmin}:{col_end}{rmax}").Value  
    if not isinstance(rng, (tuple, list)) or not isinstance(rng[0], (tuple, list)):
        rng = (rng,)
    wanted_offsets = {ri - rmin for ri in row_indices}
    def _is_zero_or_blank(v):
        if v is None:
            return True
        s = str(v).strip()
        if not s:
            return True
        try:
            return float(s) == 0.0
        except Exception:
            return False
    cols = len(rng[0])
    count = 0
    for c in range(cols):
        any_nonzero = False
        for r_off in wanted_offsets:
            try:
                v = rng[r_off][c]
            except Exception:
                v = None
            if not _is_zero_or_blank(v):
                any_nonzero = True
                break
        if any_nonzero:
            count += 1
    return count
def sum_cells_openpyxl(ws, addrs: list[str]):
    total, any_vals = 0.0, False
    for a in addrs:
        v = read_one_cell_openpyxl(ws, a)
        if v is None: 
            continue
        try:
            v = float(str(v).replace(",", "").strip())
            total += v; any_vals = True
        except Exception:
            pass
    return total if any_vals else None
def sum_cells_xlsb(file_path: Path, sheet: str, addrs: list[str]):
    from pandas import read_excel
    df = read_excel(file_path, sheet_name=sheet, engine="pyxlsb", header=None)
    total, any_vals = 0.0, False
    for a in addrs:
        m = re.fullmatch(r"([A-Za-z]+)(\d+)", a.strip())
        if not m:
            continue
        col_letters, row_str = m.groups()
        r = int(row_str) - 1
        c = col_letter_to_index(col_letters) - 1
        try:
            v = df.iat[r, c]
            if v is None: 
                continue
            v = float(str(v).replace(",", "").strip())
            total += v; any_vals = True
        except Exception:
            pass
    return total if any_vals else None
def infer_period_date(path: Path):
    d = parse_date_from_text(str(path))
    if d:
        return d
    if USE_FILE_MTIME_IF_NO_DATE:
        try:
            return _dt.fromtimestamp(path.stat().st_mtime).date()
        except Exception:
            return None
    return None
def safe_div(n, d):
    try:
        n = float(str(n).replace(",", "").strip())
        d = float(str(d).replace(",", "").strip())
        if d == 0:
            return None
        return n / d
    except Exception:
        return None
def col_letter_to_index(letter: str) -> int:
    letter = letter.strip().upper()
    num = 0
    for ch in letter:
        if not ('A' <= ch <= 'Z'):
            raise ValueError(f"Invalid column letter: {letter}")
        num = num * 26 + (ord(ch) - ord('A') + 1)
    return num
def a1_to_rowcol(a1: str) -> tuple[int, int]:
    m = re.fullmatch(r"([A-Za-z]+)(\d+)", a1.strip())
    if not m:
        raise ValueError(f"Bad cell address: {a1}")
    letters, row = m.groups()
    col = 0
    for ch in letters.upper():
        col = col * 26 + (ord(ch) - ord('A') + 1)
    return int(row), col
import re
def _passes_filters(row_dict: dict, include_contains: dict | None,
                    exclude_regex: dict | None) -> bool:
    if include_contains:
        for col, needle in include_contains.items():
            v = row_dict.get(col)
            if not isinstance(v, str) or needle.lower() not in v.lower():
                return False
    if exclude_regex:
        for col, rx in exclude_regex.items():
            v = row_dict.get(col)
            if isinstance(v, str) and re.search(rx, v, flags=re.IGNORECASE):
                return False
    return True
def sum_column_openpyxl_filtered(ws, target_col: str,
                                 include_contains: dict | None = None,
                                 exclude_regex: dict | None = None,
                                 row_start: int | None = None,
                                 row_end: int | None = None,
                                 skip_hidden: bool = False) -> float | None:
    t_idx = col_letter_to_index(target_col)
    need_cols = {t_idx}
    if include_contains:
        need_cols |= {col_letter_to_index(c) for c in include_contains.keys()}
    if exclude_regex:
        need_cols |= {col_letter_to_index(c) for c in exclude_regex.keys()}
    min_c, max_c = min(need_cols), max(need_cols)
    row_start = row_start or 1
    row_end = row_end or ws.max_row
    total, any_vals = 0.0, False
    for r_idx, row_vals in enumerate(ws.iter_rows(min_row=row_start, max_row=row_end,
                                                  min_col=min_c, max_col=max_c,
                                                  values_only=True), start=row_start):
        if skip_hidden and hasattr(ws, "row_dimensions"):
            rd = ws.row_dimensions.get(r_idx) if hasattr(ws.row_dimensions, "get") else None
            if rd is not None and getattr(rd, "hidden", False):
                continue
        row_map = {}
        for c_idx_off, val in enumerate(row_vals, start=min_c):
            row_map[chr(ord('A') + c_idx_off - 1)] = val  # 'A', 'B', ...
        if not _passes_filters(row_map, include_contains, exclude_regex):
            continue
        val = row_map.get(chr(ord('A') + t_idx - 1))
        try:
            if val is None:
                continue
            if isinstance(val, (int, float)):
                total += float(val); any_vals = True
            else:
                total += float(str(val).replace(",", "").strip()); any_vals = True
        except Exception:
            continue
    return total if any_vals else None
def _svt_hc_in_wip_openpyxl(ws,
                            key_col_letter: str = "C",
                            cond_col_letter: str = "G",
                            row_start: int = 7,
                            row_end: int = 200) -> int:
    kc = col_letter_to_index(key_col_letter)
    cc = col_letter_to_index(cond_col_letter)
    unique_keys = set()
    for r, row in enumerate(ws.iter_rows(min_row=row_start, max_row=row_end,
                                         min_col=min(kc, cc), max_col=max(kc, cc),
                                         values_only=True), start=row_start):
        vals = {col_letter_to_index(chr(ord('A') + i)): v
                for i, v in enumerate(range(min(kc, cc), max(kc, cc) + 1))}
        key_val = row[kc - min(kc, cc)]
        cond_val = row[cc - min(kc, cc)]
        try:
            if cond_val is None:
                continue
            cond_num = float(str(cond_val).replace(",", "").strip())
            if cond_num <= 0:
                continue
        except Exception:
            continue
        if key_val is None:
            continue
        key_str = str(key_val).strip()
        if key_str:
            unique_keys.add(key_str)
    return len(unique_keys)
def _svt_people_available_openpyxl(ws_individual) -> list[tuple[str, float | None]]:
    row_starts = [24, 27, 30, 33, 36]
    out = []
    for r in row_starts:
        nm = ws_individual[f"A{r}"].value
        nm = (str(nm).strip() if nm is not None else "")
        if not nm:
            continue
        try:
            v = ws_individual[f"I{r}"].value
            avail = safe_numeric(v)
        except Exception:
            avail = None
        out.append((nm, avail))
    return out
def _people_available_openpyxl_generic(ws,
                                       name_col: str = "A",
                                       avail_col: str = "I",
                                       start_row: int = 6,
                                       end_row: int = 32,
                                       step: int = 3) -> list[tuple[str, float | None]]:
    out: list[tuple[str, float | None]] = []
    BAD_NAMES = {"", "#REF!", "-", "–", "—", "0"}
    for r in range(start_row, end_row + 1, step):
        nm = ws[f"{name_col}{r}"].value
        nm = (str(nm).strip() if nm is not None else "")
        if nm in BAD_NAMES:
            continue
        avail = safe_numeric(ws[f"{avail_col}{r}"].value)
        out.append((nm, avail))
    return out
def _svt_completed_hours_by_person_openpyxl(ws_pa,
                                            people: list[str],
                                            name_col: str = "C",
                                            minutes_col: str = "G",
                                            row_start: int = 1,
                                            row_end:   int = 200,
                                            skip_hidden: bool = True) -> dict[str, float]:
    kc = col_letter_to_index(name_col)
    mc = col_letter_to_index(minutes_col)
    want = {p.strip().casefold(): p.strip() for p in people if str(p).strip()}
    totals_min = {canonical: 0.0 for canonical in want.keys()}
    min_c = min(kc, mc)
    max_c = max(kc, mc)
    for r_idx, row_vals in enumerate(
        ws_pa.iter_rows(min_row=row_start, max_row=row_end,
                        min_col=min_c, max_col=max_c, values_only=True),
        start=row_start
    ):
        if skip_hidden and hasattr(ws_pa, "row_dimensions"):
            rd = ws_pa.row_dimensions.get(r_idx) if hasattr(ws_pa.row_dimensions, "get") else None
            if rd is not None and getattr(rd, "hidden", False):
                continue
        name_val = row_vals[kc - min_c]
        mins_val = row_vals[mc - min_c]
        if name_val is None or mins_val is None:
            continue
        name_key = str(name_val).strip().casefold()
        if name_key not in totals_min:
            continue
        try:
            mins = float(str(mins_val).replace(",", "").strip())
            if mins > 0:
                totals_min[name_key] += mins
        except Exception:
            continue
    out = {}
    for key, total_m in totals_min.items():
        out[want[key]] = round(total_m / 60.0, 2)
    return out
def _tct_people_available_openpyxl(ws, name_col: str, avail_col: str,
                                   start_row: int = 24, end_row: int = 68, step: int = 3) -> list[tuple[str, float | None]]:
    rows = list(range(start_row, end_row + 1, step))
    out: list[tuple[str, float | None]] = []
    for r in rows:
        nm = ws[f"{name_col}{r}"].value
        nm = (str(nm).strip() if nm is not None else "")
        if not nm:
            continue
        avail = safe_numeric(ws[f"{avail_col}{r}"].value)
        out.append((nm, avail))
    return out
def _tct_people_available_xlsb(file_path: Path, sheet: str, name_col: str, avail_col: str,
                               start_row: int = 24, end_row: int = 68, step: int = 3) -> list[tuple[str, float | None]]:
    out: list[tuple[str, float | None]] = []
    for r in range(start_row, end_row + 1, step):
        nm = read_one_cell_xlsb(file_path, sheet, f"{name_col}{r}")
        nm = (str(nm).strip() if nm is not None else "")
        if not nm:
            continue
        av = read_one_cell_xlsb(file_path, sheet, f"{avail_col}{r}")
        out.append((nm, safe_numeric(av)))
    return out
def _completed_hours_by_person_pyxlsb(file_path: Path, sheet: str, people: list[str],
                                      name_col: str = "C", minutes_col: str = "H",
                                      row_start: int = 1, row_end: int = 200) -> dict[str, float]:
    from pandas import read_excel
    df = read_excel(file_path, sheet_name=sheet, engine="pyxlsb", header=None)
    rn = slice(row_start - 1, row_end)
    c_name = col_letter_to_index(name_col) - 1
    c_min  = col_letter_to_index(minutes_col) - 1
    sub = df.iloc[rn, [c_name, c_min]].copy()
    sub.columns = ["name", "mins"]
    sub["name_key"] = sub["name"].astype(str).str.strip().str.casefold()
    sub["mins"] = pd.to_numeric(sub["mins"], errors="coerce")
    want = {p.strip().casefold(): p.strip() for p in people if str(p).strip()}
    sub = sub[sub["name_key"].isin(want.keys())]
    grp = sub.dropna(subset=["mins"]).groupby("name_key")["mins"].sum()
    out = {}
    for k, mins in grp.items():
        out[want[k]] = round(float(mins) / 60.0, 2)
    for k, disp in want.items():
        out.setdefault(disp, 0.0)
    return out
def _hc_in_wip_from_file(file_path: Path, sheet_name: str,
                         key_col_letter: str = "C", cond_col_letter: str = "G",
                         row_start: int = 7, row_end: int = 200) -> int | None:
    ext = file_path.suffix.lower()
    try:
        if ext in (".xlsx", ".xlsm"):
            wb = load_workbook(file_path, data_only=True, read_only=True)
            if sheet_name not in wb.sheetnames:
                return None
            ws = wb[sheet_name]
            return _svt_hc_in_wip_openpyxl(ws, key_col_letter, cond_col_letter, row_start, row_end)
        elif ext == ".xlsb":
            from pandas import read_excel
            df = read_excel(file_path, sheet_name=sheet_name, engine="pyxlsb", header=None)
            rs = row_start - 1
            re_ = row_end     # pandas slice end is exclusive
            k = col_letter_to_index(key_col_letter) - 1
            c = col_letter_to_index(cond_col_letter) - 1
            cond_series = pd.to_numeric(df.iloc[rs:re_, c], errors="coerce")
            mask = cond_series > 0
            keys_series = df.iloc[rs:re_, k].where(mask)
            uniq = set()
            for v in keys_series.dropna().tolist():
                s = str(v).strip()
                if s and s.lower() != "nan":
                    uniq.add(s)
            return len(uniq)
        else:
            return None
    except Exception:
        return None
def _filter_pss_date_window(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    if "team" not in df.columns or "period_date" not in df.columns:
        return df
    min_d = pd.Timestamp("2024-06-03")
    max_d = pd.Timestamp.today().normalize()
    mask = ~(
        (df["team"] == "PSS") &
        (
            df["period_date"].isna() |
            (df["period_date"] < min_d) |
            (df["period_date"] > max_d)
        )
    )
    return df.loc[mask].copy()
def collect_pss_team(cfg: dict) -> list[dict]:
    file_path = None
    if "file" in cfg:
        file_path = Path(cfg["file"])
    elif "file_glob" in cfg:
        candidates = [Path(p) for p in glob.glob(cfg["file_glob"])]
        candidates = [p for p in candidates if p.exists()]
        if candidates:
            candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
            file_path = candidates[0]
    team_name = cfg["name"]
    src_display = str(file_path) if file_path else (cfg.get("file") or cfg.get("file_glob") or "")
    rows: list[dict] = []
    if file_path is None or not file_path.exists():
        return [{"team": team_name, "source_file": src_display, "error": "PSS-mode file not found"}]
    def _to_float(v):
        if v is None:
            return None
        try:
            return float(str(v).replace(",", "").strip())
        except Exception:
            return None
    dd = cfg.get("dropdown_iter")
    if dd:
        try:
            import win32com.client as win32
        except Exception:
            return [{"team": team_name, "source_file": src_display,
                     "error": "pywin32 not installed; run 'pip install pywin32' to enable PSS dropdown mode"}]
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = None
        try:
            wb = excel.Workbooks.Open(str(file_path))
            sheet_name = dd.get("sheet", "Previous Weeks")
            a1 = dd.get("cell", "A2")
            ws = wb.Worksheets(sheet_name)
            date_values = _get_validation_values_via_com(excel, ws, a1, dd.get("source_hint"))
            if not date_values:
                cur_val = ws.Range(a1).Value
                date_values = [cur_val] if cur_val else []
            seen = set(); uniq = []
            for v in date_values:
                k = str(v)
                if k not in seen:
                    seen.add(k); uniq.append(v)
            date_values = uniq
            min_d = _dt(2024, 6, 3).date()
            max_d = _dt.today().date() - timedelta(days=7)
            filtered = []
            for dv in date_values:
                d = _coerce_to_date_for_filter(dv)
                if d is None or d < min_d or d > max_d:
                    continue
                filtered.append((dv, d))
            for dv, period_date in filtered:
                ws.Range(a1).Value = _to_excel_com_value(dv)
                excel.CalculateFullRebuild()
                row = {"team": team_name, "source_file": src_display, "period_date": period_date}
                row.update(_read_cells_from_excel_com(ws, cfg.get("cells") or {}))
                row.update(_sum_pairs_from_excel_com(ws, cfg.get("sum_pairs") or {}))
                row["HC in WIP"] = _count_pss_hc_in_wip_com(ws)
                rows.append(row)
            try:
                ws_cw = wb.Worksheets("Current Week")
                tah = _to_float(ws_cw.Range("R61").Value)
                ch  = (_to_float(ws_cw.Range("X4").Value) or 0.0) + (_to_float(ws_cw.Range("Z4").Value) or 0.0)
                to_ = (_to_float(ws_cw.Range("X7").Value) or 0.0) + (_to_float(ws_cw.Range("Z7").Value) or 0.0)
                ao  = (_to_float(ws_cw.Range("X2").Value) or 0.0) + (_to_float(ws_cw.Range("Z2").Value) or 0.0)
                today = _dt.today().date()
                week_start = today - timedelta(days=today.weekday())  # Monday
                rows.append({
                    "team": team_name,
                    "source_file": src_display,
                    "period_date": week_start,
                    "Total Available Hours": tah,
                    "Completed Hours": ch,
                    "Target Output": to_,
                    "Actual Output": ao,
                    "HC in WIP": _count_pss_hc_in_wip_com(ws_cw)
                })
            except Exception:
                pass
        except Exception as e:
            rows.append({"team": team_name, "source_file": src_display, "error": f"PSS dropdown mode failed: {e}"})
        finally:
            if wb:
                wb.Close(SaveChanges=False)
            excel.Quit()
        return rows
    file_ext = file_path.suffix.lower()
    if file_ext in (".xlsx", ".xlsm"):
        need_visible_rows = any(
            isinstance(spec, dict) and spec.get("skip_hidden", False)
            for mapping in (cfg.get("sum_columns") or {}).values()
            for spec in mapping.values()
        )
        wb = load_workbook(file_path, data_only=True, read_only=not need_visible_rows)
        period_date = None
        if "period" in cfg:
            ps = cfg["period"]
            ps_sheet = ps.get("sheet")
            ps_cell  = ps.get("cell")
            if ps_sheet and ps_cell and ps_sheet in wb.sheetnames:
                ws_pd = wb[ps_sheet]
                raw = read_one_cell_openpyxl(ws_pd, ps_cell)
                if isinstance(raw, (_dt, _date)):
                    period_date = raw.date() if isinstance(raw, _dt) else raw
                elif isinstance(raw, (int, float)):
                    period_date = _excel_serial_to_date(raw)
                elif raw is not None:
                    try:
                        period_date = dateparser.parse(str(raw)).date()
                    except Exception:
                        period_date = None
        row = {"team": team_name, "source_file": src_display, "period_date": period_date}
        cbs = cfg.get("cells_by_sheet") or {}
        for sheet_name, mapping in cbs.items():
            if sheet_name not in wb.sheetnames:
                continue
            ws = wb[sheet_name]
            for out_name, addr in mapping.items():
                if isinstance(addr, list):
                    tot, anyv = 0.0, False
                    for a1 in addr:
                        v = read_one_cell_openpyxl(ws, a1)
                        fv = _to_float(v)
                        if fv is not None:
                            tot += fv; anyv = True
                    row[out_name] = (tot if anyv else None)
                else:
                    row[out_name] = read_one_cell_openpyxl(ws, addr)
        sc = cfg.get("sum_columns") or {}
        for sheet_name, mapping in sc.items():
            if sheet_name not in wb.sheetnames:
                continue
            ws = wb[sheet_name]
            for out_name, spec in mapping.items():
                try:
                    if isinstance(spec, str):
                        row[out_name] = sum_column(ws, spec)
                    elif isinstance(spec, dict):
                        col_letter   = spec.get("col")
                        row_start    = spec.get("row_start")
                        row_end      = spec.get("row_end")
                        include_filt = spec.get("include_contains")
                        exclude_rx   = spec.get("exclude_regex")
                        skip_hidden  = spec.get("skip_hidden", False)
                        divide       = spec.get("divide")
                        val = sum_column_openpyxl_filtered(
                            ws, col_letter,
                            include_contains=include_filt,
                            exclude_regex=exclude_rx,
                            row_start=row_start,
                            row_end=row_end,
                            skip_hidden=skip_hidden
                        )
                        if val is not None and divide:
                            try:
                                val = float(val) / float(divide)
                            except Exception:
                                pass
                        row[out_name] = val
                    else:
                        row[out_name] = None
                except Exception:
                    row[out_name] = None
            if team_name == "SVT":
                try:
                    if "#12 Production Analysis" in wb.sheetnames:
                        ws_hc = wb["#12 Production Analysis"]
                        row["HC in WIP"] = _svt_hc_in_wip_openpyxl(ws_hc)
                except Exception:
                    row["HC in WIP"] = None  
                try:
                    if "Individual" in wb.sheetnames and "#12 Production Analysis" in wb.sheetnames:
                        ws_ind = wb["Individual"]
                        ws_pa  = wb["#12 Production Analysis"]
                        people_avail = _svt_people_available_openpyxl(ws_individual=ws_ind)  
                        names = [n for n, _ in people_avail]
                        completed_hours = _svt_completed_hours_by_person_openpyxl(
                            ws_pa=ws_pa,
                            people=names,
                            name_col="C",
                            minutes_col="G",
                            row_start=1,
                            row_end=200,
                            skip_hidden=True
                        )
                        per_person = {}
                        for name, avail in people_avail:
                            a = completed_hours.get(name, 0.0)
                            try:
                                avail_f = float(avail) if avail is not None else 0.0
                            except Exception:
                                avail_f = 0.0
                            per_person[str(name).strip()] = {
                                "actual":   round(float(a or 0.0), 2),
                                "available": round(float(avail_f), 2),
                            }
                        row["Person Hours"] = json.dumps(per_person, ensure_ascii=False)
                except Exception:
                    pass
                if team_name == "SVT":
                    try:
                        by_person, by_cell = _outputs_person_and_cell_for_team(file_path, team_name)
                        if by_person:
                            row["Outputs by Person"] = json.dumps(by_person, ensure_ascii=False)
                        if by_cell:
                            row["Outputs by Cell/Station"] = json.dumps(by_cell, ensure_ascii=False)
                    except Exception:
                        pass
                try:
                    cs_hours = _cell_station_hours_for_team(file_path, team_name)
                    if cs_hours:
                        row["Cell/Station Hours"] = json.dumps(cs_hours, ensure_ascii=False)
                except Exception:
                    pass
            if team_name in ("TCT Commercial", "TCT Clinical"):
                try:
                    ind_sheet = "Individual(WIP-Non WIP)"
                    pa_sheet  = ("Commercial #12 Prod Analysis" if team_name == "TCT Commercial"
                                 else "Clinical #12 Prod Analysis")
                    if ind_sheet in wb.sheetnames:
                        ws_ind = wb[ind_sheet]
                        if team_name == "TCT Commercial":
                            people_avail = _tct_people_available_openpyxl(ws_ind, name_col="A", avail_col="I",
                                                                          start_row=24, end_row=68, step=3)
                        else:
                            people_avail = _tct_people_available_openpyxl(ws_ind, name_col="Z", avail_col="AG",
                                                                          start_row=24, end_row=68, step=3)
                    else:
                        people_avail = []
                    names = [n for n, _ in people_avail]
                    completed_hours = {}
                    if pa_sheet in wb.sheetnames and names:
                        ws_pa = wb[pa_sheet]
                        completed_hours = _svt_completed_hours_by_person_openpyxl(
                            ws_pa=ws_pa, people=names, name_col="C", minutes_col="H",
                            row_start=1, row_end=500, skip_hidden=True
                        )
                    if people_avail:
                        per_person = {}
                        for name, avail in people_avail:
                            a = completed_hours.get(name, 0.0)
                            per_person[str(name).strip()] = {
                                "actual":   round(float(a or 0.0), 2),
                                "available": round(float(avail or 0.0), 2),
                            }
                        row["Person Hours"] = json.dumps(per_person, ensure_ascii=False)
                except Exception:
                    pass
        rows.append(row)
        if "Current Week" in wb.sheetnames:
            ws_cw = wb["Current Week"]
            def _read(a1): return read_one_cell_openpyxl(ws_cw, a1)
            tah = _to_float(_read("R61"))
            ch  = (_to_float(_read("X4")) or 0.0) + (_to_float(_read("Z4")) or 0.0)
            to_ = (_to_float(_read("X7")) or 0.0) + (_to_float(_read("Z7")) or 0.0)
            ao  = (_to_float(_read("X2")) or 0.0) + (_to_float(_read("Z2")) or 0.0)
            today = _dt.today().date()
            week_start = today - timedelta(days=today.weekday())
            rows.append({
                "team": team_name,
                "source_file": src_display,
                "period_date": week_start,
                "Total Available Hours": tah,
                "Completed Hours": ch,
                "Target Output": to_,
                "Actual Output": ao,
                **({"HC in WIP": _svt_hc_in_wip_openpyxl(wb["#12 Production Analysis"])}
                    if (team_name == "SVT" and "#12 Production Analysis" in wb.sheetnames) else {})
            })
        return rows
    elif file_ext == ".xlsb":
        return [{"team": team_name, "source_file": src_display, "error": "XLSB PSS not supported in dropdown mode"}]
    else:
        return [{"team": team_name, "source_file": src_display, "error": f"Unsupported file type: {file_ext}"}]
def sum_column_pyxlsb_filtered(file_path: Path, sheet: str, target_col: str,
                               include_contains: dict | None = None,
                               exclude_regex: dict | None = None,
                               row_start: int | None = None,
                               row_end: int | None = None) -> float | None:
    from pandas import read_excel
    try:
        df = read_excel(file_path, sheet_name=sheet, engine="pyxlsb", header=None)
        t = col_letter_to_index(target_col) - 1
        s = pd.to_numeric(df.iloc[:, t], errors="coerce")
        mask = pd.Series(True, index=df.index)
        if include_contains:
            for col_letter, needle in include_contains.items():
                c = col_letter_to_index(col_letter) - 1
                mask &= df.iloc[:, c].astype(str).str.contains(needle, case=False, na=False)
        if exclude_regex:
            for col_letter, rx in exclude_regex.items():
                c = col_letter_to_index(col_letter) - 1
                mask &= ~df.iloc[:, c].astype(str).str.contains(rx, case=False, na=False, regex=True)
        if row_start or row_end:
            rs = (row_start - 1) if row_start else 0
            re_ = (row_end - 1) if row_end else len(df) - 1
            idx = df.index[rs:re_ + 1]
            mask = mask.loc[idx]
        s = s.where(mask)
        tot = s.dropna().sum()
        return float(tot) if pd.notna(tot) else None
    except Exception:
        return None
def sum_column_from_file(file_path: Path, sheet: str, col_letter: str,
                         include_contains: dict | None = None,
                         exclude_regex: dict | None = None,
                         row_start: int | None = None,
                         row_end: int | None = None,
                         skip_hidden: bool = False) -> float | None:
    ext = file_path.suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        wb = load_workbook(file_path, data_only=True, read_only=not skip_hidden)
        if sheet not in wb.sheetnames:
            return None
        ws = wb[sheet]
        return sum_column_openpyxl_filtered(ws, col_letter,
                                            include_contains, exclude_regex,
                                            row_start, row_end, skip_hidden)
    elif ext == ".xlsb":
        return sum_column_pyxlsb_filtered(file_path, sheet, col_letter,
                                          include_contains, exclude_regex,
                                          row_start, row_end)
    else:
        return None
def sum_column(ws, col_letter):
    col = 0
    for ch in col_letter.strip().upper():
        if not ('A' <= ch <= 'Z'):
            raise ValueError(f"Invalid column letter: {col_letter}")
        col = col * 26 + (ord(ch) - ord('A') + 1)
    total = 0.0
    any_vals = False
    for (val,) in ws.iter_rows(min_col=col, max_col=col, values_only=True):
        try:
            if val is None: 
                continue
            if isinstance(val, (int, float)):
                total += float(val); any_vals = True
            else:
                v = float(str(val).replace(",", "").strip())
                total += v; any_vals = True
        except Exception:
            continue
    return total if any_vals else None
def safe_numeric(x):
    if x is None:
        return None
    if isinstance(x, (int, float)):
        return float(x)
    try:
        s = str(x).strip().replace(",", "")
        return float(s)
    except Exception:
        return None
def normalize_period_date(df: pd.DataFrame) -> pd.DataFrame:
    if "period_date" in df.columns:
        df["period_date"] = pd.to_datetime(df["period_date"], errors="coerce").dt.normalize()
    return df
def read_with_openpyxl(file_path: Path, cells_cfg: dict, sumcols_cfg: dict) -> dict:
    need_visible_rows = any(
        isinstance(spec, dict) and spec.get("skip_hidden", False)
        for mapping in (sumcols_cfg or {}).values()
        for spec in mapping.values()
    )
    wb = load_workbook(file_path, data_only=True, read_only=not need_visible_rows)
    out = {}
    for sheet_name, mapping in (cells_cfg or {}).items():
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        ws = wb[sheet_name]
        for out_name, addr in mapping.items():
            try:
                r, c = a1_to_rowcol(addr)
                vals = list(ws.iter_rows(min_row=r, max_row=r, min_col=c, max_col=c, values_only=True))
                out[out_name] = vals[0][0] if vals else None
            except Exception:
                out[out_name] = None
    for sheet_name, mapping in (sumcols_cfg or {}).items():
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found")
        ws = wb[sheet_name]
        for out_name, spec in mapping.items():
            try:
                if isinstance(spec, str):
                    out[out_name] = sum_column(ws, spec)
                elif isinstance(spec, dict):
                    col_letter   = spec.get("col")
                    row_start    = spec.get("row_start")
                    row_end      = spec.get("row_end")
                    include_filt = spec.get("include_contains")
                    exclude_rx   = spec.get("exclude_regex")
                    skip_hidden  = spec.get("skip_hidden", False)
                    divide       = spec.get("divide")
                    val = sum_column_openpyxl_filtered(
                        ws,
                        col_letter,
                        include_contains=include_filt,
                        exclude_regex=exclude_rx,
                        row_start=row_start,
                        row_end=row_end,
                        skip_hidden=skip_hidden
                    )
                    if val is not None and divide:
                        try:
                            val = float(val) / float(divide)
                        except Exception:
                            pass
                    out[out_name] = val
                else:
                    out[out_name] = None
            except Exception:
                out[out_name] = None
    return out
def read_with_pyxlsb(file_path: Path, cells_cfg: dict, sumcols_cfg: dict) -> dict:
    from pandas import read_excel
    out = {}
    for sheet_name, mapping in (cells_cfg or {}).items():
        df = read_excel(file_path, sheet_name=sheet_name, engine="pyxlsb", header=None)
        for out_name, addr in mapping.items():
            m = re.fullmatch(r"([A-Za-z]+)(\d+)", addr.strip())
            if not m:
                out[out_name] = None
                continue
            col_letters, row_str = m.groups()
            r = int(row_str) - 1
            c = col_letter_to_index(col_letters) - 1
            try:
                val = df.iat[r, c]
            except Exception:
                val = None
            out[out_name] = val
    for sheet_name, mapping in (sumcols_cfg or {}).items():
        df = read_excel(file_path, sheet_name=sheet_name, engine="pyxlsb", header=None)
        for out_name, spec in mapping.items():
            try:
                if isinstance(spec, str):
                    c = col_letter_to_index(spec) - 1
                    series = pd.to_numeric(df.iloc[:, c], errors="coerce")
                    total = float(series.dropna().sum())
                    out[out_name] = total
                elif isinstance(spec, dict):
                    col_letter   = spec.get("col")
                    row_start    = spec.get("row_start")
                    row_end      = spec.get("row_end")
                    include_filt = spec.get("include_contains")
                    exclude_rx   = spec.get("exclude_regex")
                    divide       = spec.get("divide")
                    val = sum_column_pyxlsb_filtered(
                        file_path,
                        sheet_name,
                        col_letter,
                        include_contains=include_filt,
                        exclude_regex=exclude_rx,
                        row_start=row_start,
                        row_end=row_end
                    )
                    if val is not None and divide:
                        try:
                            val = float(val) / float(divide)
                        except Exception:
                            pass
                    out[out_name] = val
                else:
                    out[out_name] = None
            except Exception:
                out[out_name] = None
    return out
def _read_sheet_as_df(file_path: Path, sheet_name: str):
    from pandas import read_excel
    ext = file_path.suffix.lower()
    engine = "pyxlsb" if ext == ".xlsb" else None
    try:
        return read_excel(file_path, sheet_name=sheet_name, engine=engine, header=None)
    except Exception:
        return None
def _sum_output_target_by(df: pd.DataFrame, key_col_idx: int, out_col_idx: int, tgt_col_idx: int) -> dict:
    if df is None or df.empty:
        return {}
    cols = [key_col_idx, out_col_idx, tgt_col_idx]
    n = df.shape[1]
    if any(i >= n for i in cols):
        return {}
    sub = df.iloc[:, cols].copy()
    sub.columns = ["key", "output", "target"]
    sub["key"] = sub["key"].astype(str).str.strip()
    bad_keys = {"", "-", "–", "—", "nan"}  
    sub = sub[~sub["key"].isin(bad_keys)]
    sub = sub[sub["key"].str.len() > 0] 
    sub["output"] = pd.to_numeric(sub["output"], errors="coerce")
    sub["target"] = pd.to_numeric(sub["target"], errors="coerce")
    sub = sub[sub[["output", "target"]].notna().any(axis=1)]
    if sub.empty:
        return {}
    grp = sub.groupby("key", dropna=False, sort=True).agg({"output": "sum", "target": "sum"})
    out = {}
    for k, row in grp.iterrows():
        out[k] = {
            "output": float(round(row.get("output", 0.0) if pd.notna(row.get("output")) else 0.0, 2)),
            "target": float(round(row.get("target", 0.0) if pd.notna(row.get("target")) else 0.0, 2)),
        }
    return out
def _outputs_person_and_cell_for_team(file_path: Path, team_name: str) -> tuple[dict, dict]:
    team = (team_name or "").strip().casefold()
    if team in ("svt", "ect", "pvh", "crdn", "aortic"):
        sheet = "#12 Production Analysis"
        person_idx = 2   # C
        cell_idx   = 3   # D
        target_idx = 4   # F
        output_idx = 8   # I
        if team not in ("aortic",):
            target_idx = 5
    elif team == "tct clinical":
        sheet = "Clinical #12 Prod Analysis"
        person_idx = 2; cell_idx = 3; target_idx = 6; output_idx = 9
    elif team == "tct commercial":
        sheet = "Commercial #12 Prod Analysis"
        person_idx = 2; cell_idx = 3; target_idx = 6; output_idx = 9
    else:
        return {}, {}
    df = _read_sheet_as_df(file_path, sheet)
    if df is None:
        return {}, {}
    by_person = _sum_output_target_by(df, key_col_idx=person_idx, out_col_idx=output_idx, tgt_col_idx=target_idx)
    by_cell   = _sum_output_target_by(df, key_col_idx=cell_idx,   out_col_idx=output_idx, tgt_col_idx=target_idx)
    return by_person, by_cell
def _cell_station_hours_for_team(file_path: Path, team_name: str) -> dict:
    team = (team_name or "").strip().casefold()
    if team in ("svt", "ect", "pvh", "crdn", "aortic"):
        sheet = "#12 Production Analysis"
        key_idx = 3
        mins_idx = 6
    elif team == "tct clinical":
        sheet = "Clinical #12 Prod Analysis"; key_idx = 3; mins_idx = 7
    elif team == "tct commercial":
        sheet = "Commercial #12 Prod Analysis"; key_idx = 3; mins_idx = 7
    else:
        return {}
    df = _read_sheet_as_df(file_path, sheet)
    if df is None or df.empty: return {}
    n = df.shape[1]
    if key_idx >= n or mins_idx >= n: return {}
    sub = df.iloc[:, [key_idx, mins_idx]].copy()
    sub.columns = ["cell_station", "mins"]
    sub["cell_station"] = sub["cell_station"].astype(str).str.strip()
    bad = {"", "-", "–", "—", "nan"}
    sub = sub[~sub["cell_station"].isin(bad)]
    sub["mins"] = pd.to_numeric(sub["mins"], errors="coerce")
    sub = sub.dropna(subset=["mins"])
    if sub.empty: return {}
    agg = sub.groupby("cell_station", dropna=False)["mins"].sum()
    return {k: round(float(v) / 60.0, 2) for k, v in agg.items() if pd.notna(v) and float(v) > 0}
def read_metrics_from_file(file_path: Path, cells_cfg: dict, sumcols_cfg: dict) -> dict:
    ext = file_path.suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        return read_with_openpyxl(file_path, cells_cfg, sumcols_cfg)
    elif ext == ".xlsb":
        return read_with_pyxlsb(file_path, cells_cfg, sumcols_cfg)
    else:
        raise ValueError(f"Unsupported file type: {ext}")
def _aortic_hc_in_wip_from_file(file_path: Path,
                                sheet: str = "#12 Production Analysis",
                                col_c: str = "C",
                                row_start: int = 7,
                                row_end: int = 199,
                                skip_hidden: bool = True) -> int | None:
    ext = file_path.suffix.lower()
    def _unique_from_openpyxl(ws) -> int:
        c_idx = col_letter_to_index(col_c)
        names = set()
        for r_idx, row_vals in enumerate(
            ws.iter_rows(min_row=row_start, max_row=row_end,
                         min_col=c_idx, max_col=c_idx, values_only=True),
            start=row_start
        ):
            if skip_hidden and hasattr(ws, "row_dimensions"):
                rd = ws.row_dimensions.get(r_idx) if hasattr(ws.row_dimensions, "get") else None
                if rd is not None and getattr(rd, "hidden", False):
                    continue
            v = row_vals[0]
            s = str(v).strip() if v is not None else ""
            if s and s != "#REF!" and s != "0":
                names.add(s)
        return len(names)

    if ext in (".xlsx", ".xlsm"):
        try:
            wb = load_workbook(file_path, data_only=True, read_only=skip_hidden is False)
            if sheet not in wb.sheetnames:
                return None
            return _unique_from_openpyxl(wb[sheet])
        except Exception:
            return None
    elif ext == ".xlsb":
        from pandas import read_excel
        try:
            df = read_excel(file_path, sheet_name=sheet, engine="pyxlsb", header=None)
            c = col_letter_to_index(col_c) - 1
            series = df.iloc[row_start-1:row_end, c].astype(str)
            names = set()
            for s in series:
                s = s.strip()
                if s and s.lower() != "nan" and s != "#REF!" and s != "0":
                    names.add(s)
            return len(names)
        except Exception:
            return None
    else:
        return None
def _aortic_completed_hours_from_file(file_path: Path,
                                      sheet: str = "#12 Production Analysis",
                                      col_c: str = "C",
                                      col_d: str = "D",
                                      row_start: int = 7,
                                      row_end: int = 199,
                                      skip_hidden: bool = True) -> float | None:
    ext = file_path.suffix.lower()
    def _count_pairs_openpyxl(ws):
        c_idx = col_letter_to_index(col_c)
        d_idx = col_letter_to_index(col_d)
        min_c, max_c = min(c_idx, d_idx), max(c_idx, d_idx)
        hits = 0
        for r_idx, row_vals in enumerate(
            ws.iter_rows(min_row=row_start, max_row=row_end,
                         min_col=min_c, max_col=max_c, values_only=True),
            start=row_start
        ):
            if skip_hidden and hasattr(ws, "row_dimensions"):
                rd = ws.row_dimensions.get(r_idx) if hasattr(ws.row_dimensions, "get") else None
                if rd is not None and getattr(rd, "hidden", False):
                    continue
            c_val = row_vals[c_idx - min_c]
            d_val = row_vals[d_idx - min_c]
            if (str(c_val).strip() if c_val is not None else "") and (str(d_val).strip() if d_val is not None else ""):
                hits += 1
        return float(hits) * 2.0 if hits > 0 else 0.0
    if ext in (".xlsx", ".xlsm"):
        try:
            wb = load_workbook(file_path, data_only=True, read_only=skip_hidden is False)
            if sheet not in wb.sheetnames:
                return None
            ws = wb[sheet]
            return _count_pairs_openpyxl(ws)
        except Exception:
            return None
    elif ext == ".xlsb":
        from pandas import read_excel
        try:
            df = read_excel(file_path, sheet_name=sheet, engine="pyxlsb", header=None)
            c_i = col_letter_to_index(col_c) - 1
            d_i = col_letter_to_index(col_d) - 1
            sub = df.iloc[row_start-1:row_end, [c_i, d_i]].copy()
            sub = sub.dropna(how="all")
            def _nonempty(x): 
                return isinstance(x, (int,float)) or (isinstance(x, str) and x.strip() != "")
            mask = sub.iloc[:,0].apply(_nonempty) & sub.iloc[:,1].apply(_nonempty)
            hits = int(mask.sum())
            return float(hits) * 2.0 if hits > 0 else 0.0
        except Exception:
            return None
    else:
        return None
def _aortic_hours_by_col(file_path: Path,
                         sheet: str = "#12 Production Analysis",
                         col_letter: str = "C",
                         row_start: int = 7,
                         row_end: int = 199,
                         skip_hidden: bool = True) -> dict[str, float]:
    def _accumulate_openpyxl(ws) -> dict[str, float]:
        c_idx = col_letter_to_index(col_letter)
        out: dict[str, float] = {}
        for r_idx, row_vals in enumerate(
            ws.iter_rows(min_row=row_start, max_row=row_end,
                         min_col=c_idx, max_col=c_idx, values_only=True),
            start=row_start
        ):
            if skip_hidden and hasattr(ws, "row_dimensions"):
                rd = ws.row_dimensions.get(r_idx) if hasattr(ws.row_dimensions, "get") else None
                if rd is not None and getattr(rd, "hidden", False):
                    continue
            v = row_vals[0]
            s = str(v).strip() if v is not None else ""
            bad = {"", "#REF!", "nan"}
            if col_letter.upper() == "C":
                bad |= {"0", "-", "–", "—"}
            if col_letter.upper() == "D":
                bad |= {"-", "–", "—"}
            if s not in bad:
                out[s] = round(out.get(s, 0.0) + 2.0, 2)  # Rule 3/5: +2 hours per occurrence
        return out
    ext = file_path.suffix.lower()
    try:
        if ext in (".xlsx", ".xlsm"):
            wb = load_workbook(file_path, data_only=True, read_only=skip_hidden is False)
            if sheet not in wb.sheetnames:
                return {}
            return _accumulate_openpyxl(wb[sheet])
        elif ext == ".xlsb":
            from pandas import read_excel
            df = read_excel(file_path, sheet_name=sheet, engine="pyxlsb", header=None)
            c = col_letter_to_index(col_letter) - 1
            series = df.iloc[row_start-1:row_end, c].astype(str)
            out: dict[str, float] = {}
            for s in series:
                s = s.strip()
                bad = {"", "#REF!", "nan"}
                if col_letter.upper() == "C":
                    bad |= {"0", "-", "–", "—"}
                if col_letter.upper() == "D":
                    bad |= {"-", "–", "—"}
                if s not in bad:
                    out[s] = round(out.get(s, 0.0) + 2.0, 2)  # Rule 3/5
            return out
        else:
            return {}
    except Exception:
        return {}
def collect_for_team(team_cfg: dict) -> list[dict]:
    root = Path(team_cfg["root"])
    pattern = team_cfg.get("pattern", "*.xlsx")
    team_name = team_cfg["name"]
    cells_cfg = team_cfg.get("cells", {})
    sumcols_cfg = team_cfg.get("sum_columns", {})
    if not root.exists():
        print(f"[WARN] Root not found for {team_name}: {root}", file=sys.stderr)
        return []
    rows = []
    for p in root.rglob(pattern):
        if p.is_dir() or looks_like_temp(p.name) or _is_excluded_path(p):
            continue
        try:
            period = None
            per_cfg = team_cfg.get("period")
            if per_cfg and isinstance(per_cfg, dict):
                sheet = per_cfg.get("sheet")
                cell  = per_cfg.get("cell")
                if sheet and cell:
                    ext = p.suffix.lower()
                    try:
                        if ext == ".xlsb":
                            val = read_one_cell_xlsb(p, sheet, cell)
                        else:
                            wb = load_workbook(p, data_only=True, read_only=True)
                            if sheet in wb.sheetnames:
                                ws = wb[sheet]
                                val = read_one_cell_openpyxl(ws, cell)
                            else:
                                val = None
                        if isinstance(val, (_dt, _date)):
                            period = val.date() if isinstance(val, _dt) else val
                        elif isinstance(val, (int, float)):
                            period = _excel_serial_to_date(val)
                        elif val is not None:
                            period = dateparser.parse(str(val)).date()
                    except Exception:
                        period = None
            if period is None:
                period = infer_period_date(p)
            if team_name == "ECT":
                cutoff = _dt(2024, 8, 19).date()
                if isinstance(period, _dt):
                    period = period.date()
                if isinstance(period, _date) and period < cutoff:
                    continue
            if team_name.lower().startswith("tct"):
                today = _dt.today().date()
                if isinstance(period, _dt):
                    period = period.date()
                if isinstance(period, _date) and period > today:
                    print(f"[skip] TCT future period {period} -> {p}")
                    continue
            values = read_metrics_from_file(p, cells_cfg, sumcols_cfg)
            if team_name in ("ECT", "PVH", "CRDN", "Aortic"):
                try:
                    if team_name == "Aortic":
                        values["HC in WIP"] = _aortic_hc_in_wip_from_file(p, "#12 Production Analysis", col_c="C",
                                                                        row_start=7, row_end=199, skip_hidden=True)
                    else:
                        values["HC in WIP"] = _hc_in_wip_from_file(p, "#12 Production Analysis")
                except Exception:
                    values["HC in WIP"] = None
                try:
                    by_person, by_cell = _outputs_person_and_cell_for_team(p, team_name)
                    if by_person:
                        values["Outputs by Person"] = json.dumps(by_person, ensure_ascii=False)
                    if by_cell:
                        values["Outputs by Cell/Station"] = json.dumps(by_cell, ensure_ascii=False)
                except Exception:
                    pass
                try:
                    cs_hours = _cell_station_hours_for_team(p, team_name)
                    if cs_hours:
                        values["Cell/Station Hours"] = json.dumps(cs_hours, ensure_ascii=False)
                except Exception:
                    pass
                try:
                    wb_tmp = load_workbook(p, data_only=True, read_only=True)
                    per_person = {}
                    if "Individual (WIP-Non WIP)" in wb_tmp.sheetnames:
                        ws_ind = wb_tmp["Individual (WIP-Non WIP)"]
                        people_avail = _people_available_openpyxl_generic(
                            ws_ind, name_col="A", avail_col="I", start_row=6, end_row=32, step=3
                        )
                        names = [n for n, _ in people_avail]
                        if "#12 Production Analysis" in wb_tmp.sheetnames and names:
                            ws_pa = wb_tmp["#12 Production Analysis"]
                            completed_hours = _svt_completed_hours_by_person_openpyxl(
                                ws_pa=ws_pa, people=names, name_col="C", minutes_col="G",
                                row_start=7, row_end=199, skip_hidden=True
                            )
                            for name, avail in people_avail:
                                a = completed_hours.get(name, 0.0)
                                per_person[name] = {"actual": round(float(a or 0.0), 2),
                                                    "available": round(float(avail or 0.0), 2)}
                    if per_person:
                        values["Person Hours"] = json.dumps(per_person, ensure_ascii=False)
                except Exception:
                    pass
            if team_name == "Aortic":
                try:
                    per_actual = _aortic_hours_by_col(p, sheet="#12 Production Analysis", col_letter="C")
                    if per_actual:
                        existing_pp = {}
                        if "Person Hours" in values and values["Person Hours"]:
                            try:
                                existing_pp = json.loads(values["Person Hours"])
                            except Exception:
                                existing_pp = {}
                        BAD_NAMES = {"", "#REF!", "nan", "0", "-", "–", "—"}
                        existing_pp = {k: v for k, v in existing_pp.items() if str(k).strip() not in BAD_NAMES}
                        merged = {}
                        for name, hours in per_actual.items():
                            prev = existing_pp.get(name, {"actual": 0.0, "available": 0.0})
                            prev["actual"] = float(hours)
                            merged[name] = prev
                        for name, prev in existing_pp.items():
                            merged.setdefault(name, prev)
                        values["Person Hours"] = json.dumps(merged, ensure_ascii=False)
                except Exception:
                    pass
                try:
                    cs_hours = _aortic_hours_by_col(p, sheet="#12 Production Analysis", col_letter="D")
                    if cs_hours:
                        values["Cell/Station Hours"] = json.dumps(cs_hours, ensure_ascii=False)
                except Exception:
                    pass
                try:
                    ch = _aortic_completed_hours_from_file(
                        p, sheet="#12 Production Analysis", col_c="C", col_d="D",
                        row_start=7, row_end=199, skip_hidden=True
                    )
                    if ch is not None:
                        values["Completed Hours"] = ch
                except Exception:
                    pass
            if team_name in ("SVT", "TCT Clinical", "TCT Commercial"):
                try:
                    by_person, by_cell = _outputs_person_and_cell_for_team(p, team_name)
                    if by_person:
                        values["Outputs by Person"] = json.dumps(by_person, ensure_ascii=False)
                    if by_cell:
                        values["Outputs by Cell/Station"] = json.dumps(by_cell, ensure_ascii=False)
                except Exception:
                    pass
                try:
                    cs_hours = _cell_station_hours_for_team(p, team_name)
                    if cs_hours:
                        values["Cell/Station Hours"] = json.dumps(cs_hours, ensure_ascii=False)
                except Exception:
                    pass
            sheet_for_hc = None
            if team_name == "TCT Commercial":
                sheet_for_hc = "Commercial #12 Prod Analysis"
            elif team_name == "TCT Clinical":
                sheet_for_hc = "Clinical #12 Prod Analysis"
            if sheet_for_hc:
                values["HC in WIP"] = _hc_in_wip_from_file(p, sheet_for_hc)
            if team_name == "SVT":
                try:
                    wb_tmp = load_workbook(p, data_only=True, read_only=True)
                    if "#12 Production Analysis" in wb_tmp.sheetnames:
                        ws_hc = wb_tmp["#12 Production Analysis"]
                        values["HC in WIP"] = _svt_hc_in_wip_openpyxl(ws_hc)
                except Exception:
                    values["HC in WIP"] = None
            if p.suffix.lower() == ".xlsb" and team_name in ("TCT Commercial", "TCT Clinical"):
                try:
                    ind_sheet = "Individual(WIP-Non WIP)"
                    pa_sheet  = ("Commercial #12 Prod Analysis" if team_name == "TCT Commercial"
                                else "Clinical #12 Prod Analysis")
                    if team_name == "TCT Commercial":
                        people_avail = _tct_people_available_xlsb(
                            p, ind_sheet, name_col="A", avail_col="I", start_row=24, end_row=68, step=3
                        )
                    else:
                        people_avail = _tct_people_available_xlsb(
                            p, ind_sheet, name_col="Z", avail_col="AG", start_row=24, end_row=68, step=3
                        )
                    names = [n for n, _ in people_avail]
                    completed_hours = {}
                    if names:
                        completed_hours = _completed_hours_by_person_pyxlsb(
                            p, pa_sheet, names, name_col="C", minutes_col="H", row_start=1, row_end=200
                        )
                    if people_avail:
                        per_person = {}
                        for name, avail in people_avail:
                            a = completed_hours.get(name, 0.0)
                            per_person[str(name).strip()] = {
                                "actual":    round(float(a or 0.0), 2),
                                "available": round(float(avail or 0.0), 2),
                            }
                        values["Person Hours"] = json.dumps(per_person, ensure_ascii=False)
                except Exception:
                    pass
            rows.append({
                "team": team_name,
                "period_date": period,
                "source_file": str(p),
                **values
            })
        except Exception as e:
            error_cols = {}
            for mapping in cells_cfg.values():
                for out_name in mapping.keys():
                    error_cols.setdefault(out_name, None)
            for mapping in sumcols_cfg.values():
                for out_name in mapping.keys():
                    error_cols.setdefault(out_name, None)
            rows.append({
                "team": team_name,
                "period_date": infer_period_date(p),
                "source_file": str(p),
                "error": str(e),
                **error_cols
            })
    rows = apply_fallbacks_for_team(rows, team_cfg)
    return rows
def apply_fallbacks_for_team(rows: list[dict], team_cfg: dict) -> list[dict]:
    fb = team_cfg.get("fallback_total_available_hours")
    if not fb:
        return rows
    by_date = {}
    dated_rows = []
    for i, r in enumerate(rows):
        d = r.get("period_date")
        if d:
            dated_rows.append((d, i))
            by_date.setdefault(d, []).append(i)
    dated_rows.sort(key=lambda t: t[0])
    def find_nearest_earlier(target_date):
        lo, hi = 0, len(dated_rows) - 1
        best_idx = None
        while lo <= hi:
            mid = (lo + hi) // 2
            d_mid, idx_mid = dated_rows[mid]
            if d_mid < target_date:
                best_idx = idx_mid
                lo = mid + 1
            else:
                hi = mid - 1
        return best_idx
    for cur_date, row_idx in dated_rows:
        row = rows[row_idx]
        tah_val = row.get("Total Available Hours")
        try:
            is_zero = float(str(tah_val).replace(",", "").strip()) == 0.0
        except Exception:
            is_zero = False
        if not is_zero:
            continue
        target_prev = cur_date - timedelta(days=7)
        prev_indices = by_date.get(target_prev)
        candidate_idx = None
        if prev_indices:
            candidate_idx = prev_indices[0]
        else:
            candidate_idx = find_nearest_earlier(cur_date)
        if candidate_idx is None:
            continue
        prev_file = Path(rows[candidate_idx].get("source_file", ""))
        if not prev_file.exists():
            continue
        repl = sum_column_from_file(
            prev_file,
            fb["sheet"],
            fb["column"],
            include_contains=fb.get("include_contains"),
            exclude_regex=fb.get("exclude_regex"),
            row_start=fb.get("row_start"),
            row_end=fb.get("row_end"),
            skip_hidden=fb.get("skip_hidden", False),
        )
        if repl is not None:
            row["Total Available Hours"] = repl
            row["fallback_used"] = f"{fb['sheet']}!{fb['column']} from {prev_file.name}"
    return rows
def build_master(rows: list[dict]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows)
    df = normalize_period_date(df)
    base_cols = ["team", "period_date", "source_file"]
    metric_cols = [c for c in df.columns if c not in base_cols + ["error"]]
    cols = base_cols + metric_cols + (["error"] if "error" in df.columns else [])
    df = df.reindex(columns=cols)
    if "period_date" in df.columns:
        df = df.sort_values(["team", "period_date", "source_file"], ascending=[True, True, True])
    return df
def save_outputs(df: pd.DataFrame):
    if df.empty:
        print("No rows collected. Check paths/sheets/cells.")
        return
    df = _dedupe_by_team_unique_key(df)
    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as xlw:
        df.to_excel(xlw, index=False, sheet_name="All Metrics")
    preferred_cols = [
        "team", "period_date", "source_file",
        "Total Available Hours", "Completed Hours",
        "Target Output", "Actual Output",
        "Target UPLH", "Actual UPLH",
        "UPLH WP1", "UPLH WP2",
        "HC in WIP", "Actual HC Used",
        "People in WIP", "Person Hours",
        "Outputs by Person", "Outputs by Cell/Station",
        "Cell/Station Hours",
        "Open Complaint Timeliness",
        "fallback_used", "error",
    ]
    cols = [c for c in preferred_cols if c in df.columns]
    out = df.loc[:, cols].copy()
    if "period_date" in out.columns:
        out["period_date"] = pd.to_datetime(out["period_date"], errors="coerce").dt.strftime("%Y-%m-%d")
    numeric_cols = {"Total Available Hours", "Completed Hours", "Target Output", "Actual Output",
                    "Target UPLH", "Actual UPLH", "Actual HC Used", "UPLH WP1", "UPLH WP2"} & set(out.columns)
    for c in numeric_cols:
        out[c] = pd.to_numeric(out[c], errors="coerce")
    out = out.replace({np.nan: ""})
    out = out.drop_duplicates(keep="last")
    out.to_csv(
        OUT_CSV,
        index=False,
        sep=",",
        encoding="utf-8",
        lineterminator="\n",
        quoting=csv.QUOTE_MINIMAL,
        date_format="%Y-%m-%d",
    )
    print(f"Saved Excel: {OUT_XLSX.resolve()}")
    print(f"Saved CSV:   {OUT_CSV.resolve()}")
    REPO_DIR.mkdir(parents=True, exist_ok=True)
    def _samefile(a: Path, b: Path) -> bool:
        try:
            return a.resolve().samefile(b.resolve())
        except Exception:
            return str(a.resolve()).lower() == str(b.resolve()).lower()
    try:
        if not _samefile(OUT_CSV, REPO_CSV):
            shutil.copyfile(OUT_CSV, REPO_CSV)
            print(f"Copied CSV to: {REPO_CSV.resolve()}")
        else:
            print("[info] OUT_CSV and REPO_CSV are the same file; skipping copy.")
    except Exception as e:
        print(f"[WARN] Copy step failed: {e}", file=sys.stderr)
    target_for_git = REPO_CSV if REPO_CSV.exists() else OUT_CSV
    git_autocommit_and_push(REPO_DIR, target_for_git, branch=GIT_BRANCH)
def merge_with_existing(new_df: pd.DataFrame) -> pd.DataFrame:
    new_df = normalize_period_date(new_df)
    old_frames = []
    if OUT_XLSX.exists():
        try:
            old_xlsx = pd.read_excel(OUT_XLSX, sheet_name="All Metrics")
            old_frames.append(old_xlsx)
        except Exception:
            pass
    if REPO_CSV.exists():
        try:
            old_csv = pd.read_csv(REPO_CSV, dtype=str, keep_default_na=False)
            for col in old_csv.columns:
                if col.lower() in {"total available hours","completed hours","target output","actual output",
                                   "target uplh","actual uplh","actual hc used","hc in wip",
                                   "open complaint timeliness"}:
                    old_csv[col] = pd.to_numeric(old_csv[col], errors="coerce")
            old_frames.append(old_csv)
        except Exception:
            pass
    if old_frames:
        old = pd.concat(old_frames, ignore_index=True)
    else:
        return new_df
    old = normalize_period_date(old)
    if "team" in old.columns:
        old = old.loc[old["team"] != "PH"].copy()
    team_keys = {}
    for cfg in TEAM_CONFIG:
        key = cfg.get("unique_key", ["team", "period_date", "source_file"])
        team_keys[cfg["name"]] = key
    def make_key(df):
        keys = []
        for _, r in df.iterrows():
            kcols = team_keys.get(r.get("team"), ["team", "period_date", "source_file"])
            parts = []
            for c in kcols:
                if c == "period_date":
                    ts = pd.to_datetime(r.get(c), errors="coerce")
                    parts.append(ts.date().isoformat() if pd.notna(ts) else None)
                else:
                    parts.append(r.get(c))
            keys.append(tuple(parts))
        return pd.Series(keys, index=df.index)
    old = old.copy();      old["_origin"] = "old";      old["_key"] = make_key(old)
    new_df = new_df.copy(); new_df["_origin"] = "new";  new_df["_key"] = make_key(new_df)
    combined = pd.concat([old, new_df], ignore_index=True)
    combined = normalize_period_date(combined)
    if "source_file" in combined.columns:
        norm = combined["source_file"].astype(str).str.lower().str.replace("/", "\\")
        is_old = combined["_origin"] == "old"
        keep = pd.Series(True, index=combined.index)
        if EXCLUDED_SOURCE_FILES:
            keep &= is_old | (~norm.isin(EXCLUDED_SOURCE_FILES))
        if EXCLUDED_DIRS:
            keep &= is_old | (~norm.str.startswith(tuple(EXCLUDED_DIRS)))
        combined = combined.loc[keep].copy()
    latest_by_team = combined.groupby("team", dropna=False)["period_date"].transform("max")
    is_latest = (combined["period_date"] == latest_by_team) & combined["period_date"].notna()
    origin_rank = combined["_origin"].map({"old": 0, "new": 1}).fillna(1)
    past = (combined.loc[~is_latest]
            .assign(_origin_rank=origin_rank[~is_latest])
            .sort_values(["team", "period_date", "_origin_rank", "source_file"])
            .drop_duplicates(subset=["_key"], keep="first"))   # old wins for past
    curr = (combined.loc[is_latest]
            .assign(_origin_rank=origin_rank[is_latest])
            .sort_values(["team", "period_date", "_origin_rank", "source_file"])
            .drop_duplicates(subset=["_key"], keep="last"))    # new wins for latest
    combined = pd.concat([past, curr], ignore_index=True)
    combined = normalize_period_date(combined)
    combined = combined.drop(columns=[c for c in ["_key", "_origin", "_origin_rank"] if c in combined.columns])
    base_cols = ["team", "period_date", "source_file"]
    metric_cols = [c for c in combined.columns if c not in base_cols + ["error"]]
    cols = base_cols + metric_cols + (["error"] if "error" in combined.columns else [])
    combined = combined.reindex(columns=cols)
    if "period_date" in combined.columns:
        combined = combined.sort_values(["team", "period_date", "source_file"], ascending=[True, True, True])
    combined = _filter_future_periods(combined)
    combined = _filter_ph_zero_hours(combined)
    combined = _filter_pss_date_window(combined)
    combined = _filter_ect_min_year(combined)
    return combined
def _dedupe_by_team_unique_key(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    team_keys = {}
    for cfg in TEAM_CONFIG:
        key = cfg.get("unique_key", ["team", "period_date", "source_file"])
        team_keys[cfg["name"]] = key
    def _make_key_row(r) -> tuple:
        kcols = team_keys.get(r.get("team"), ["team", "period_date", "source_file"])
        parts = []
        for c in kcols:
            if c == "period_date":
                ts = pd.to_datetime(r.get(c), errors="coerce")
                parts.append(ts.normalize().date().isoformat() if pd.notna(ts) else None)
            else:
                parts.append(r.get(c))
        return tuple(parts)
    df = df.copy()
    df["_dedupe_key"] = df.apply(_make_key_row, axis=1)
    df = (df
          .sort_values(["team", "period_date", "source_file"], na_position="last")
          .drop_duplicates(subset=["_dedupe_key"], keep="last")
          .drop(columns=["_dedupe_key"]))
    return df
def _read_timeliness_csv_standardized() -> pd.DataFrame:
    p = TIMELINESS_CSV
    if not p.exists():
        return pd.DataFrame(columns=["team", "period_date", "Open Complaint Timeliness"])
    try:
        t = pd.read_csv(p, dtype=str, keep_default_na=False)
    except Exception as e:
        print(f"[timeliness] Failed to read {p}: {e}")
        return pd.DataFrame(columns=["team", "period_date", "Open Complaint Timeliness"])
    if t.shape[1] < 3:
        t = t.iloc[:, :3].copy()
        t.columns = ["team", "period_date", "Open Complaint Timeliness"]
    else:
        lower_cols = [str(c).strip().lower() for c in t.columns]
        def _first_match(names):
            for want in names:
                if want in lower_cols:
                    return t.columns[lower_cols.index(want)]
            return None
        team_col = _first_match(["team"])
        date_col = _first_match(["period_date", "period", "date"])
        val_col  = _first_match(["open complaint timeliness","timeliness","value","metric"])
        if not (team_col and date_col and val_col):
            t = t.iloc[:, :3].copy()
            t.columns = ["team", "period_date", "Open Complaint Timeliness"]
        else:
            t = t.rename(columns={
                team_col: "team", date_col: "period_date", val_col: "Open Complaint Timeliness"
            })
    t["team"] = t["team"].astype(str).str.strip()
    t["period_date"] = pd.to_datetime(t["period_date"], errors="coerce").dt.normalize()
    s = t["Open Complaint Timeliness"].astype(str).str.strip().str.replace("%","", regex=False)
    t["Open Complaint Timeliness"] = pd.to_numeric(s.str.extract(r'([+-]?\d+(?:\.\d+)?)', expand=False), errors="coerce")
    t = t.drop_duplicates(subset=["team","period_date"], keep="last")
    return t
def ensure_timeliness_placeholders(metrics_df: pd.DataFrame):
    if metrics_df.empty or "team" not in metrics_df.columns or "period_date" not in metrics_df.columns:
        return
    pairs = (
        metrics_df[["team","period_date"]]
        .dropna()
        .copy()
    )
    pairs["team"] = pairs["team"].astype(str).str.strip()
    pairs["period_date"] = pd.to_datetime(pairs["period_date"], errors="coerce").dt.normalize()
    pairs = pairs.dropna().drop_duplicates()
    t = _read_timeliness_csv_standardized()
    pairs["_team_key"] = pairs["team"].str.casefold()
    t["_team_key"] = t["team"].str.casefold()
    missing = (
        pairs.merge(t[["_team_key","period_date"]], on=["_team_key","period_date"], how="left", indicator=True)
             .loc[lambda d: d["_merge"] == "left_only", ["team","period_date"]]
             .drop_duplicates()
    )
    if missing.empty:
        print("[timeliness] No missing timeliness rows to add.")
        return
    placeholders = missing.copy()
    placeholders["Open Complaint Timeliness"] = np.nan
    out = pd.concat([t[["team","period_date","Open Complaint Timeliness"]], placeholders], ignore_index=True)
    out = out.drop_duplicates(subset=["team","period_date"], keep="last")
    out = out.sort_values(["team","period_date"]).reset_index(drop=True)
    TIMELINESS_CSV.parent.mkdir(parents=True, exist_ok=True)
    out_to_csv = out.copy()
    out_to_csv["period_date"] = pd.to_datetime(out_to_csv["period_date"], errors="coerce").dt.strftime("%Y-%m-%d")
    out_to_csv.to_csv(TIMELINESS_CSV, index=False)
    print(f"[timeliness] Added {len(placeholders)} placeholder row(s) to {TIMELINESS_CSV}.")
    git_autocommit_and_push(REPO_DIR, TIMELINESS_CSV, branch=GIT_BRANCH)
def add_open_complaint_timeliness(df: pd.DataFrame) -> pd.DataFrame:
    try:
        p = TIMELINESS_CSV
    except NameError:
        p = Path(r"C:\heijunka-dev") / "timeliness.csv"
    if not p.exists():
        print(f"[timeliness] {p} not found; skipping join.")
        return df
    try:
        t = pd.read_csv(p, dtype=str, keep_default_na=False)
    except Exception as e:
        print(f"[timeliness] Failed to read {p}: {e}")
        return df
    if t.shape[1] < 3:
        print("[timeliness] Expected at least 3 columns (team, date/period_date, value). Skipping join.")
        return df
    lower_cols = [str(c).strip().lower() for c in t.columns]
    def _first_match(names):
        for want in names:
            if want in lower_cols:
                return t.columns[lower_cols.index(want)]
        return None
    team_col = _first_match(["team"])
    date_col = _first_match(["period_date", "period", "date"])
    val_col  = _first_match(["open complaint timeliness", "timeliness", "value", "metric"])
    if not (team_col and date_col and val_col):
        t = t.iloc[:, :3].copy()
        t.columns = ["team", "period_date", "Open Complaint Timeliness"]
    else:
        t = t.rename(columns={
            team_col: "team",
            date_col: "period_date",
            val_col:  "Open Complaint Timeliness"
        })
    def _team_key(s: pd.Series) -> pd.Series:
        return s.astype(str).str.strip().str.casefold()
    def _to_num(series: pd.Series) -> pd.Series:
        s = series.astype(str).str.strip()
        s = s.str.replace("%", "", regex=False)
        extracted = s.str.extract(r'([+-]?\d+(?:\.\d+)?)', expand=False)
        return pd.to_numeric(extracted, errors="coerce")
    t["team"] = t["team"].astype(str).str.strip()
    t["_team_key"] = _team_key(t["team"])
    t["period_date"] = pd.to_datetime(t["period_date"], errors="coerce").dt.normalize()
    t["Open Complaint Timeliness"] = _to_num(t["Open Complaint Timeliness"])
    t = t.dropna(subset=["_team_key", "period_date"]).drop_duplicates(subset=["_team_key", "period_date"], keep="last")
    out = df.copy()
    if "team" not in out.columns or "period_date" not in out.columns:
        print("[timeliness] 'team'/'period_date' not found in metrics df; skipping join.")
        return df
    out["team"] = out["team"].astype(str).str.strip()
    out["_team_key"] = _team_key(out["team"])
    out["period_date"] = pd.to_datetime(out["period_date"], errors="coerce").dt.normalize()
    exact = out.merge(
        t[["_team_key", "period_date", "Open Complaint Timeliness"]],
        on=["_team_key", "period_date"],
        how="left",
        suffixes=("", "_t")
    )
    if "Open Complaint Timeliness" in out.columns:
        left_vals  = pd.to_numeric(exact["Open Complaint Timeliness"], errors="coerce")
        right_vals = pd.to_numeric(exact.get("Open Complaint Timeliness_t"), errors="coerce")
        exact["Open Complaint Timeliness"] = right_vals.combine_first(left_vals)
        if "Open Complaint Timeliness_t" in exact.columns:
            exact = exact.drop(columns=["Open Complaint Timeliness_t"])
    else:
        if "Open Complaint Timeliness_t" in exact.columns and "Open Complaint Timeliness" not in exact.columns:
            exact = exact.rename(columns={"Open Complaint Timeliness_t": "Open Complaint Timeliness"})
    mask_missing = exact["Open Complaint Timeliness"].isna() if "Open Complaint Timeliness" in exact.columns else pd.Series(False, index=exact.index)
    def _week_monday(s: pd.Series) -> pd.Series:
        d = pd.to_datetime(s, errors="coerce").dt.normalize()
        return d - pd.to_timedelta(d.dt.weekday, unit="D")
    if mask_missing.any():
        tmp = exact.loc[mask_missing].copy()
        tmp["_week_key"] = _week_monday(tmp["period_date"])
        t2 = t.copy()
        t2["_week_key"] = _week_monday(t2["period_date"])
        t2_week = (
            t2[["_team_key", "_week_key", "period_date", "Open Complaint Timeliness"]]
            .sort_values(["_team_key", "_week_key", "period_date"])
            .drop_duplicates(["_team_key", "_week_key"], keep="last")
        )
        wk = tmp.merge(
            t2_week,
            on=["_team_key", "_week_key"],
            how="left",
            suffixes=("", "_wk")
        )
        fill = pd.to_numeric(wk["Open Complaint Timeliness"], errors="coerce")
        exact.loc[mask_missing, "Open Complaint Timeliness"] = fill.values
        mask_missing = exact["Open Complaint Timeliness"].isna()
    if mask_missing.any():
        tol = pd.Timedelta(days=6)
        left_near = exact.loc[mask_missing, ["_team_key", "period_date"]].copy()
        left_near = left_near.sort_values(["_team_key", "period_date"])
        right_near = t[["_team_key", "period_date", "Open Complaint Timeliness"]].copy()
        right_near = right_near.sort_values(["_team_key", "period_date"])
        filled_vals = []
        for team_key, left_g in left_near.groupby("_team_key"):
            r_g = right_near[right_near["_team_key"] == team_key]
            if r_g.empty:
                filled_vals.extend([np.nan] * len(left_g))
                continue
            merged_g = pd.merge_asof(
                left_g.sort_values("period_date"),
                r_g.sort_values("period_date"),
                on="period_date",
                direction="nearest",
                tolerance=tol
            )
            filled_vals.extend(merged_g["Open Complaint Timeliness"].tolist())
        exact.loc[mask_missing, "Open Complaint Timeliness"] = pd.to_numeric(filled_vals, errors="coerce")
    exact = exact.drop(columns=[c for c in ["_team_key"] if c in exact.columns])
    return exact
def run_once():
    all_rows = []
    for cfg in TEAM_CONFIG:
        if cfg.get("pss_mode"):
            all_rows.extend(collect_pss_team(cfg))
        elif cfg.get("ph_mode"):
            all_rows.extend(collect_ph_team(cfg))
        else:
            all_rows.extend(collect_for_team(cfg))
    df = build_master(all_rows)
    df = _filter_ph_zero_hours(df) 
    df = _filter_future_periods(df)
    df = _filter_pss_date_window(df)
    df = _filter_ect_min_year(df)
    df["Target UPLH"] = df.apply(lambda r: safe_div(r.get("Target Output"), r.get("Total Available Hours")), axis=1)
    df["Actual UPLH"] = df.apply(lambda r: safe_div(r.get("Actual Output"), r.get("Completed Hours")), axis=1)
    df["Target UPLH"] = df["Target UPLH"].round(2)
    df["Actual UPLH"] = df["Actual UPLH"].round(2)
    if "UPLH WP1" in df.columns: df["UPLH WP1"] = pd.to_numeric(df["UPLH WP1"], errors="coerce").round(2)
    if "UPLH WP2" in df.columns: df["UPLH WP2"] = pd.to_numeric(df["UPLH WP2"], errors="coerce").round(2)
    df["Actual HC Used"] = pd.to_numeric(df.get("Completed Hours"), errors="coerce") / 32.5
    df["Actual HC Used"] = df["Actual HC Used"].round(2)
    df = merge_with_existing(df)
    try:
        if REPO_CSV.exists():
            repo = pd.read_csv(REPO_CSV, dtype=str, keep_default_na=False)
            repo_svt = repo[repo["team"] == "SVT"]
            curr_svt = df[df["team"] == "SVT"]
            r_n = pd.to_datetime(repo_svt["period_date"], errors="coerce").dt.normalize().nunique()
            c_n = pd.to_datetime(curr_svt["period_date"], errors="coerce").dt.normalize().nunique()
            if c_n + 3 < r_n:
                print("[safety] SVT periods dropped unexpectedly; restoring repo history union.")
                df = merge_with_existing(df)
    except Exception:
        pass
    ensure_timeliness_placeholders(df)
    df = add_open_complaint_timeliness(df)
    df = _dedupe_by_team_unique_key(df)
    save_outputs(df)
    if not df.empty:
        with pd.option_context("display.max_columns", None, "display.width", 180):
            print("\nPreview:")
            print(df.head(12).to_string(index=False))
def watch_mode():
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler
    class NewFileHandler(FileSystemEventHandler):
        def on_created(self, event):
            if event.is_directory:
                return
            name = event.src_path.lower()
            if (name.endswith(".xlsx") or name.endswith(".xlsm") or name.endswith(".xlsb")) and not looks_like_temp(name):
                print(f"[watch] New file: {event.src_path}")
                time.sleep(2)
                run_once()
    roots = []
    for cfg in TEAM_CONFIG:
        r = cfg.get("root")
        if r:
            roots.append(r)
    run_once()
    obs = Observer()
    for r in roots:
        rp = Path(r)
        if rp.exists():
            obs.schedule(NewFileHandler(), str(rp), recursive=True)
            print(f"[watch] Watching: {rp}")
        else:
            print(f"[watch][WARN] Missing root: {rp}")
    obs.start()
    print("[watch] Running. Ctrl+C to stop.")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        obs.stop()
    obs.join()
def main():
    parser = argparse.ArgumentParser(description="Aggregate metrics from synced SharePoint Excel files (.xlsx/.xlsm/.xlsb)")
    parser.add_argument("--watch", action="store_true", help="Watch folders and refresh on new files")
    args = parser.parse_args()
    if args.watch:
        watch_mode()
    else:
        run_once()
if __name__ == "__main__":
    main()