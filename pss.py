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

# -----------------------------
# Paths & constants
# -----------------------------
REPO_DIR = Path(r"C:\heijunka-dev")
REPO_CSV = REPO_DIR / "metrics_aggregate_dev.csv"
GIT_BRANCH = "main"

# Primary timeliness path + fallbacks
TIMELINESS_CSV = REPO_DIR / "timeliness.csv"
TIMELINESS_FALLBACKS = [
    Path.cwd() / "timeliness.csv",
    Path.cwd() / "data" / "timeliness.csv",
]

# Output filenames for PSS
OUT_XLSX = Path.cwd() / "pss.xlsx"
OUT_CSV  = Path.cwd() / "pss.csv"

# -----------------------------
# Team config (PSS only)
# -----------------------------
TEAM_CONFIG = [
    {
        "name": "PSS",
        "pss_mode": True,
        "file": r"C:\\Users\\wadec8\\Medtronic PLC\\PSS Sharepoint - Documents\\PSS_Heijunka.xlsm",
        "dropdown_iter": {
            "sheet": "Previous Weeks",
            "cell":  "A2",
            "source_hint": None,
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
    }
]

# -----------------------------
# Utilities
# -----------------------------

def _git(args: list[str], cwd: Path) -> tuple[int, str, str]:
    p = subprocess.Popen(args, cwd=str(cwd), stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    out, err = p.communicate()
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

# -----------------------------
# Excel helpers
# -----------------------------

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
    return [v for v in vals if v not in (None, "")]


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

# -----------------------------
# PSS collector
# -----------------------------

def _coerce_to_date_for_filter(v) -> _date | None:
    if isinstance(v, _dt):
        return v.date()
    if isinstance(v, _date):
        return v
    if isinstance(v, (int, float)):
        try:
            d = (_dt(1899, 12, 30) + timedelta(days=float(v))).date()
            if _dt(1900, 1, 1).date() <= d <= _dt(2100, 1, 1).date():
                return d
        except Exception:
            return None
        return None
    try:
        return dateparser.parse(str(v)).date()
    except Exception:
        return None


def collect_pss_team(cfg: dict) -> list[dict]:
    file_path = Path(cfg.get("file", "")) if cfg.get("file") else None
    team_name = cfg.get("name", "PSS")
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
                # Read key cells
                row["Total Available Hours"] = _to_float(ws.Range(cfg["cells"]["Total Available Hours"]).Value)
                row["Completed Hours"]      = _to_float(ws.Range(cfg["cells"]["Completed Hours"]).Value)
                # Sum pairs
                to_ = 0.0; any_to = False
                for a in cfg["sum_pairs"]["Target Output"]:
                    v = _to_float(ws.Range(a).Value)
                    if v is not None:
                        to_ += v; any_to = True
                ao_ = 0.0; any_ao = False
                for a in cfg["sum_pairs"]["Actual Output"]:
                    v = _to_float(ws.Range(a).Value)
                    if v is not None:
                        ao_ += v; any_ao = True
                row["Target Output"] = to_ if any_to else None
                row["Actual Output"] = ao_ if any_ao else None
                # HC in WIP
                row["HC in WIP"] = _count_pss_hc_in_wip_com(ws)
                rows.append(row)
            # Current Week snapshot
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
                    "HC in WIP": _count_pss_hc_in_wip_com(ws_cw),
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

    # If not dropdown mode, bail out with an explicit error (PSS expects dropdown mode)
    return [{"team": team_name, "source_file": src_display, "error": "PSS file found but dropdown mode not enabled"}]

# -----------------------------
# Frame transforms & calcs
# -----------------------------

def normalize_period_date(df: pd.DataFrame) -> pd.DataFrame:
    if "period_date" in df.columns:
        df["period_date"] = pd.to_datetime(df["period_date"], errors="coerce").dt.normalize()
    return df


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

# --- Timeliness join (robust like in tct.py) ---

def _parse_percentish(x):
    if x is None:
        return None
    s = str(x).strip()
    if not s:
        return None
    try:
        if s.endswith('%'):
            return float(s[:-1])
        return float(s)
    except Exception:
        return None


def _week_start(s):
    ts = pd.to_datetime(s, errors="coerce")
    return ts.dt.to_period("W-MON").apply(lambda p: p.start_time.normalize())


def _find_timeliness_csv() -> Path | None:
    if TIMELINESS_CSV.exists():
        return TIMELINESS_CSV
    for p in TIMELINESS_FALLBACKS:
        if p.exists():
            print(f"[timeliness] Using fallback: {p}")
            return p
    print(f"[timeliness] Not found in primary or fallbacks: {TIMELINESS_CSV}")
    return None


def add_open_complaint_timeliness(df: pd.DataFrame) -> pd.DataFrame:
    p = _find_timeliness_csv()
    if not p:
        return df
    try:
        t = pd.read_csv(p, dtype=str, keep_default_na=False)
    except Exception as e:
        print(f"[timeliness] Failed to read {p}: {e}")
        return df
    if t.empty:
        print("[timeliness] File empty; skipping join.")
        return df

    lower = [str(c).strip().lower() for c in t.columns]

    def _first_match(names):
        for want in names:
            if want in lower:
                return t.columns[lower.index(want)]
        return None

    team_col = _first_match(["team", "group", "area"])  # may be absent
    date_col = _first_match(["period_date", "period", "week", "date", "week_start", "week ending", "week_end"]) 
    val_col  = _first_match(["open complaint timeliness", "timeliness", "% timeliness", "value", "metric", "pct"]) 

    if date_col is None or val_col is None:
        cols = list(t.columns)
        if len(cols) >= 2:
            date_col = cols[0]
            val_col = cols[1]
            team_col = team_col or (cols[2] if len(cols) >= 3 else None)
        else:
            print("[timeliness] Could not identify columns; skipping join.")
            return df

    rename_map = {date_col: "period_date", val_col: "Open Complaint Timeliness"}
    if team_col:
        rename_map[team_col] = "team"
    t = t.rename(columns=rename_map)

    if "team" not in t.columns:
        t["team"] = "PSS"

    t["team"] = t["team"].astype(str).str.strip().replace({"": "PSS"})
    t["period_date"] = pd.to_datetime(t["period_date"], errors="coerce").dt.normalize()
    t["Open Complaint Timeliness"] = t["Open Complaint Timeliness"].apply(_parse_percentish)

    # Keep only PSS rows or blanks (which we treat as PSS broadcast)
    t = t[(t["team"].str.upper() == "PSS") | (t["team"].isna()) | (t["team"].eq("") )].copy()
    t["team"] = "PSS"

    t = t.dropna(subset=["period_date", "Open Complaint Timeliness"]).drop_duplicates(subset=["team", "period_date"], keep="last")
    t["week_start"] = _week_start(t["period_date"])  # Monday

    out = df.copy()
    if out.empty:
        return df
    if "team" not in out.columns or "period_date" not in out.columns:
        print("[timeliness] 'team'/'period_date' not found in metrics df; skipping join.")
        return df

    out["team"] = out["team"].astype(str).str.strip()
    out["period_date"] = pd.to_datetime(out["period_date"], errors="coerce").dt.normalize()
    out["week_start"] = _week_start(out["period_date"])  # Monday

    merged = out.merge(
        t[["team", "week_start", "Open Complaint Timeliness"]],
        on=["team", "week_start"],
        how="left",
    )

    if "week_start" in merged.columns:
        merged = merged.drop(columns=["week_start"]) 
    return merged

# -----------------------------
# Output
# -----------------------------

def save_outputs(df: pd.DataFrame):
    if df.empty:
        print("No rows collected. Check paths/sheets/cells.")
        return

    # Ensure fallback_used and error columns exist (even if empty)
    for c in ("fallback_used", "error"):
        if c not in df.columns:
            df[c] = ""

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

    # Optional repo copy + push (kept for parity with tct.py)
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

# -----------------------------
# Pipeline
# -----------------------------

def run_once():
    all_rows = []
    for cfg in TEAM_CONFIG:
        all_rows.extend(collect_pss_team(cfg))

    df = build_master(all_rows)

    # Derived metrics
    def _safe_div(n, d):
        try:
            n = float(str(n).replace(",", "").strip())
            d = float(str(d).replace(",", "").strip())
            return None if d == 0 else n / d
        except Exception:
            return None

    if not df.empty:
        df["Target UPLH"] = df.apply(lambda r: _safe_div(r.get("Target Output"), r.get("Total Available Hours")), axis=1)
        df["Actual UPLH"] = df.apply(lambda r: _safe_div(r.get("Actual Output"), r.get("Completed Hours")), axis=1)
        df["Target UPLH"] = df["Target UPLH"].round(2)
        df["Actual UPLH"] = df["Actual UPLH"].round(2)
        df["Actual HC Used"] = pd.to_numeric(df.get("Completed Hours"), errors="coerce") / 32.5
        df["Actual HC Used"] = df["Actual HC Used"].round(2)

    # Timeliness join
    df = add_open_complaint_timeliness(df)

    save_outputs(df)

    if not df.empty:
        with pd.option_context("display.max_columns", None, "display.width", 180):
            print("\nPreview:")
            print(df.head(12).to_string(index=False))


def main():
    parser = argparse.ArgumentParser(description="PSS-only metrics extractor")
    parser.add_argument("--watch", action="store_true", help="(ignored in PSS-only script)")
    args = parser.parse_args()
    run_once()


if __name__ == "__main__":
    main()
