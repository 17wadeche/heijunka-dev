# apply_closures.py
from pathlib import Path
import pandas as pd
import numpy as np
BASE = Path(r"C:\heijunka-dev")
METRICS_CSV  = BASE / "metrics_aggregate_dev.csv"
METRICS_XLSX = BASE / "metrics_aggregate_dev.xlsx"
CLOSURES_CSV = BASE / "closures.csv"
SHEET_NAME = "All Metrics"   # your dashboard reads this sheet
def _norm_date(s):
    dt = pd.to_datetime(s, errors="coerce")
    return pd.to_datetime(dt.dt.date)
def main():
    m = pd.read_csv(METRICS_CSV, dtype=str, encoding="utf-8-sig")
    if "period_date" not in m.columns or "team" not in m.columns:
        raise ValueError("metrics_aggregate_dev.csv must have 'team' and 'period_date' columns")
    c = pd.read_csv(CLOSURES_CSV, dtype=str, encoding="utf-8-sig")
    need = {"team", "period_date", "Closures"}
    if not need.issubset(c.columns):
        missing = need - set(c.columns)
        raise ValueError(f"closures.csv is missing columns: {sorted(missing)}")
    has_opened = "Opened" in c.columns  # NEW
    m["_team"] = m["team"].astype(str).str.strip()
    m["_date"] = _norm_date(m["period_date"])
    c["_team"] = c["team"].astype(str).str.strip()
    c["_date"] = _norm_date(c["period_date"])
    c = (
        c.sort_values(["_team", "_date"])
         .drop_duplicates(["_team", "_date"], keep="last")
         .copy()
    )
    c["Closures"] = pd.to_numeric(c["Closures"], errors="coerce")
    if has_opened:  # NEW
        c["Opened"] = pd.to_numeric(c["Opened"], errors="coerce")
    have_before_closures = "Closures" in m.columns
    have_before_opened   = "Opened" in m.columns  # NEW
    merge_cols = ["_team", "_date", "Closures"] + (["Opened"] if has_opened else [])  # NEW
    out = m.merge(
        c[merge_cols],
        on=["_team", "_date"],
        how="left",
        suffixes=("", "_from_closures"),
    )
    merged_new_name = "Closures_from_closures" if "Closures_from_closures" in out.columns else "Closures"
    if have_before_closures:
        existing = pd.to_numeric(out.get("Closures"), errors="coerce")
        newvals  = pd.to_numeric(out.get(merged_new_name), errors="coerce")
        out["Closures"] = np.where(newvals.notna(), newvals, existing)
        if merged_new_name != "Closures":
            out = out.drop(columns=[merged_new_name], errors="ignore")
    else:
        if merged_new_name != "Closures":
            out["Closures"] = out[merged_new_name]
            out = out.drop(columns=[merged_new_name], errors="ignore")
    if has_opened:
        merged_opened_name = "Opened_from_closures" if "Opened_from_closures" in out.columns else "Opened"
        if have_before_opened:
            existing_o = pd.to_numeric(out.get("Opened"), errors="coerce")
            new_o      = pd.to_numeric(out.get(merged_opened_name), errors="coerce")
            out["Opened"] = np.where(new_o.notna(), new_o, existing_o)
            if merged_opened_name != "Opened":
                out = out.drop(columns=[merged_opened_name], errors="ignore")
        else:
            if merged_opened_name != "Opened":
                out["Opened"] = out[merged_opened_name]
                out = out.drop(columns=[merged_opened_name], errors="ignore")
    out = out.drop(columns=["_team", "_date"], errors="ignore")
    out.to_csv(METRICS_CSV, index=False, encoding="utf-8-sig")
    if METRICS_XLSX.exists():
        with pd.ExcelWriter(METRICS_XLSX, engine="openpyxl", mode="a", if_sheet_exists="replace") as xw:
            out.to_excel(xw, sheet_name=SHEET_NAME, index=False)
    else:
        with pd.ExcelWriter(METRICS_XLSX, engine="openpyxl") as xw:
            out.to_excel(xw, sheet_name=SHEET_NAME, index=False)
    matched_closures = out["Closures"].notna().sum()
    print(f"Updated Closures for {matched_closures}/{len(out)} metric rows.")
    if has_opened:
        matched_opened = out["Opened"].notna().sum()
        print(f"Updated Opened for {matched_opened}/{len(out)} metric rows.")
    print("Wrote:")
    print(f" - {METRICS_CSV}")
    print(f" - {METRICS_XLSX} (sheet '{SHEET_NAME}')")
if __name__ == "__main__":
    main()