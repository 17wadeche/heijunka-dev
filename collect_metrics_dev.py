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
REPO_DIR = Path(r"C:\heijunka-dev")
REPO_CSV = REPO_DIR / "metrics_aggregate_dev.csv"
GIT_BRANCH = "main"
TIMELINESS_CSV = REPO_DIR / "timeliness.csv"
EXCLUDED_SOURCE_FILES = {
    r"C:\Users\wadec8\Medtronic PLC\SVT PXM Team - Archived Heijunka\SVT Future Heijunka.xlsm"
}
EXCLUDED_DIRS = {
    r"C:\Users\wadec8\Medtronic PLC\TCT CQXM - 1 WIP and Schedule\Weekly Heijunka Archived",
    r"C:\Users\wadec8\Medtronic PLC\TCT CQXM - 1 WIP and Schedule\Clinical",
    r"C:\Users\wadec8\Medtronic PLC\TCT CQXM - 1 WIP and Schedule\Commercial",
    r"C:\Users\wadec8\Medtronic PLC\TCT CQXM - 1 WIP and Schedule\Remediation",
    r"C:\Users\wadec8\Medtronic PLC\TCT CQXM - 1 WIP and Schedule\WIP Blitz Power Hour",
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
        "name": "SVT",
        "root": r"C:\Users\wadec8\Medtronic PLC\SVT PXM Team - Archived Heijunka",
        "pattern": "*.xls*",
        "cells": {
            "Individual": {
                "Total Available Hours": "I39",
            }
        },
        "sum_columns": {
            "#12 Production Analysis": {
                "Target Output": "F",
                "Actual Output": "I",
                "Completed Hours": {
                    "col": "G",
                    "row_start": 1,
                    "row_end": 200,
                    "divide": 60,
                    "skip_hidden": True,
                },
            }
        },
        "fallback_total_available_hours": {
            "sheet": "Next Weeks Hours",
            "column": "I",
            "include_contains": {"B": "Available WIP Hours"},
            "exclude_regex": {"A": r"^\s*(Team member|Total)\b"},
            "skip_hidden": True,
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
    },{
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
            "Commercial #12 Prod Analysis": {
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
def _excel_serial_to_date(n) -> _date | None:
    try:
        return (_dt(1899, 12, 30) + timedelta(days=float(n))).date()
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
    code, _, err = _git(["git", "--version"], repo_dir)
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
    _git(["git", "fetch", remote], repo_dir)
    code, up, _ = _git(["git", "rev-parse", "--abbrev-ref", "--symbolic-full-name", "@{u}"], repo_dir)
    if code != 0:
        _git(["git", "branch", "--set-upstream-to", f"{remote}/{branch}", branch], repo_dir)
    code, _, _ = _git(["git", "pull", "--rebase", remote, branch], repo_dir)
    if code != 0:
        print("[WARN] 'git pull --rebase' failed; skipping commit to avoid conflicts. Resolve repo state manually.")
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
def _to_excel_com_value(v):
    if isinstance(v, (int, float, str)):
        return v
    if isinstance(v, _dt):
        return v
    if isinstance(v, _date):
        return _dt(v.year, v.month, v.day)
    return str(v)
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
def read_metrics_from_file(file_path: Path, cells_cfg: dict, sumcols_cfg: dict) -> dict:
    ext = file_path.suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        return read_with_openpyxl(file_path, cells_cfg, sumcols_cfg)
    elif ext == ".xlsb":
        return read_with_pyxlsb(file_path, cells_cfg, sumcols_cfg)
    else:
        raise ValueError(f"Unsupported file type: {ext}")
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
            if team_name.lower().startswith("tct"):
                today = _dt.today().date()
                if isinstance(period, _dt):
                    period = period.date()
                if isinstance(period, _date) and period > today:
                    print(f"[skip] TCT future period {period} -> {p}")
                    continue
            values = read_metrics_from_file(p, cells_cfg, sumcols_cfg)
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
    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as xlw:
        df.to_excel(xlw, index=False, sheet_name="All Metrics")
    preferred_cols = [
        "team", "period_date", "source_file",
        "Total Available Hours", "Completed Hours",
        "Target Output", "Actual Output",
        "Target UPLH", "Actual UPLH",
        "HC in WIP", "Actual HC Used",
        "Open Complaint Timeliness",
        "fallback_used", "error",
    ]
    cols = [c for c in preferred_cols if c in df.columns]
    out = df.loc[:, cols].copy()
    if "period_date" in out.columns:
        out["period_date"] = pd.to_datetime(out["period_date"], errors="coerce").dt.strftime("%Y-%m-%d")
    numeric_cols = {"Total Available Hours", "Completed Hours", "Target Output", "Actual Output",
                    "Target UPLH", "Actual UPLH", "Actual HC Used"} & set(out.columns)
    for c in numeric_cols:
        out[c] = pd.to_numeric(out[c], errors="coerce")
    out = out.replace({np.nan: ""})
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
    if not OUT_XLSX.exists():
        return new_df
    try:
        old = pd.read_excel(OUT_XLSX, sheet_name="All Metrics")
    except Exception:
        return new_df
    old = normalize_period_date(old)
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
    old["_key"] = make_key(old)
    new_df["_key"] = make_key(new_df)
    combined = pd.concat([old, new_df], ignore_index=True)
    combined = combined.drop_duplicates(subset=["_key"], keep="last").drop(columns=["_key"])
    combined = normalize_period_date(combined)
    base_cols = ["team", "period_date", "source_file"]
    metric_cols = [c for c in combined.columns if c not in base_cols + ["error"]]
    cols = base_cols + metric_cols + (["error"] if "error" in combined.columns else [])
    combined = combined.reindex(columns=cols)
    if "period_date" in combined.columns:
        combined = combined.sort_values(["team", "period_date", "source_file"], ascending=[True, True, True])
    combined = _filter_pss_date_window(combined)
    if "source_file" in combined.columns:
        norm = combined["source_file"].astype(str).str.lower().str.replace("/", "\\")
        combined = combined[~norm.isin(EXCLUDED_SOURCE_FILES)]
        if EXCLUDED_DIRS:
            combined = combined[~norm.str.startswith(tuple(EXCLUDED_DIRS))]
    return combined
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
        print("[timeliness] Expected at least 3 columns (A, B, value). Skipping join.")
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
    t["team"] = t["team"].astype(str).str.strip()
    t["period_date"] = pd.to_datetime(t["period_date"], errors="coerce").dt.normalize()
    t = t.dropna(subset=["team", "period_date"]).drop_duplicates(subset=["team", "period_date"], keep="last")
    out = df.copy()
    if "team" not in out.columns or "period_date" not in out.columns:
        print("[timeliness] 'team'/'period_date' not found in metrics df; skipping join.")
        return df
    out["team"] = out["team"].astype(str).str.strip()
    out["period_date"] = pd.to_datetime(out["period_date"], errors="coerce").dt.normalize()
    merged = out.merge(
        t[["team", "period_date", "Open Complaint Timeliness"]],
        on=["team", "period_date"],
        how="left",
        suffixes=("_left", "_right")
    )
    left_name  = "Open Complaint Timeliness_left"
    right_name = "Open Complaint Timeliness_right"
    left  = pd.to_numeric(merged[left_name], errors="coerce") if left_name in merged.columns else None
    right = pd.to_numeric(merged[right_name], errors="coerce") if right_name in merged.columns else None
    if left is not None and right is not None:
        merged["Open Complaint Timeliness"] = left.combine_first(right)
        merged = merged.drop(columns=[left_name, right_name])
    elif right is not None:
        merged = merged.rename(columns={right_name: "Open Complaint Timeliness"})
    elif left is not None:
        merged = merged.rename(columns={left_name: "Open Complaint Timeliness"})
    return merged
def run_once():
    all_rows = []
    for cfg in TEAM_CONFIG:
        if cfg.get("pss_mode"):
            all_rows.extend(collect_pss_team(cfg))
        else:
            all_rows.extend(collect_for_team(cfg))
    df = build_master(all_rows)
    df = _filter_pss_date_window(df)
    if not df.empty and "source_file" in df.columns:
        norm = df["source_file"].astype(str).str.lower().str.replace("/", "\\")
        df = df[~norm.isin(EXCLUDED_SOURCE_FILES)]
        if EXCLUDED_DIRS:
            df = df[~norm.str.startswith(tuple(EXCLUDED_DIRS))]
    def safe_div(n, d):
        try:
            n = float(str(n).replace(",", "").strip())
            d = float(str(d).replace(",", "").strip())
            return None if d == 0 else n / d
        except Exception:
            return None
    df["Target UPLH"] = df.apply(lambda r: safe_div(r.get("Target Output"), r.get("Total Available Hours")), axis=1)
    df["Actual UPLH"] = df.apply(lambda r: safe_div(r.get("Actual Output"), r.get("Completed Hours")), axis=1)
    df["Target UPLH"] = df["Target UPLH"].round(2)
    df["Actual UPLH"] = df["Actual UPLH"].round(2)
    df["Actual HC Used"] = pd.to_numeric(df.get("Completed Hours"), errors="coerce") / 32.5
    df["Actual HC Used"] = df["Actual HC Used"].round(2)
    df = merge_with_existing(df)
    df = add_open_complaint_timeliness(df)
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