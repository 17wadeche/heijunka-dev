import argparse
import json
import re
from pathlib import Path
import numpy as np
import pandas as pd
BAD_NAMES = {
    "", "-", "–", "—", "nan", "NaN", "NAN",
    "n/a", "N/A", "na", "NA", "null", "NULL",
    "none", "None", "tm", "TM", "Totals", "TOTALS"
}
def norm_name(x: str) -> str:
    return " ".join(str(x or "").strip().split())
def is_real_person(name: str) -> bool:
    n = norm_name(name)
    if not n:
        return False
    if n in {"TM", "Totals"}:
        return False
    if n.strip().lower() in {b.lower() for b in BAD_NAMES}:
        return False
    return True
def parse_week_from_sheet(sheet_name: str) -> pd.Timestamp | None:
    s = str(sheet_name).strip()
    dt = pd.to_datetime(s, errors="coerce")
    if pd.notna(dt):
        return dt.normalize()
    s2 = re.sub(r"^\s*week\s+of\s+", "", s, flags=re.IGNORECASE).strip()
    dt = pd.to_datetime(s2, errors="coerce")
    if pd.notna(dt):
        return dt.normalize()
    return None
def safe_float(x) -> float:
    if x is None:
        return np.nan
    try:
        if pd.isna(x):
            return np.nan
    except Exception:
        pass
    if isinstance(x, (int, float, np.integer, np.floating)):
        return float(x)
    s = str(x).strip()
    if not s:
        return np.nan
    if s in {"-", "–", "—"}:
        return np.nan
    s = s.replace(",", "").replace("\u00a0", " ")
    m = re.search(r"[-+]?\d*\.?\d+", s)
    if not m:
        return np.nan
    try:
        return float(m.group(0))
    except Exception:
        return np.nan
def read_people_block(ws: pd.DataFrame) -> list[dict]:
    start_i = 3
    col_name = 0
    rows = []
    for i in range(start_i, len(ws)):
        name = norm_name(ws.iat[i, col_name] if col_name < ws.shape[1] else "")
        if not name:
            break  # stop on first blank name row
        if not is_real_person(name):
            continue
        b = safe_float(ws.iat[i, 1] if 1 < ws.shape[1] else np.nan)  # col B
        c = safe_float(ws.iat[i, 2] if 2 < ws.shape[1] else np.nan)  # col C
        if pd.isna(b): b = 0.0
        if pd.isna(c): c = 0.0
        rows.append({"row_i": i, "name": name, "B": b, "C": c})
    return rows
def extract_totals(ws: pd.DataFrame) -> tuple[float, float]:
    r = 17
    b18 = safe_float(ws.iat[r, 1] if 1 < ws.shape[1] else np.nan)
    c18 = safe_float(ws.iat[r, 2] if 2 < ws.shape[1] else np.nan)
    total_nonwip = (b18 - c18) if pd.notna(b18) and pd.notna(c18) else np.nan
    ooo = c18
    return total_nonwip, ooo
def build_nonwip_by_person(people_rows: list[dict]) -> dict[str, float]:
    out = {}
    for r in people_rows:
        b = r.get("B", np.nan)
        c = r.get("C", np.nan)
        v = (b - c) if pd.notna(b) and pd.notna(c) else np.nan
        if pd.isna(v):
            v = 0.0
        out[r["name"]] = float(round(v, 2))
    return out
def build_nonwip_activities(ws: pd.DataFrame, people_rows: list[dict]) -> list[dict]:
    header_i = 2
    start_col = 3   # D
    end_col = 35    # AJ
    activities = []
    max_col = ws.shape[1] - 1
    end_col = min(end_col, max_col)
    for pr in people_rows:
        i = pr["row_i"]
        name = pr["name"]
        for c in range(start_col, end_col + 1):
            activity = norm_name(ws.iat[header_i, c] if c < ws.shape[1] else "")
            if not activity:
                continue
            hours = safe_float(ws.iat[i, c])
            if pd.isna(hours) or hours <= 0:
                continue
            activities.append({
                "name": name,
                "activity": activity,
                "hours": float(round(hours, 2))
            })
    return activities
def load_completed_hours(csv_path: Path) -> pd.DataFrame:
    df = pd.read_csv(csv_path, dtype=str, keep_default_na=False)
    df.columns = [" ".join(str(c).split()) for c in df.columns]
    if "period_date" in df.columns:
        df["period_date"] = pd.to_datetime(df["period_date"], errors="coerce").dt.normalize()
    if "Completed Hours" in df.columns:
        df["Completed Hours"] = pd.to_numeric(df["Completed Hours"], errors="coerce")
    return df
def parse_person_hours_json(person_hours_cell: str) -> dict:
    if person_hours_cell is None or (isinstance(person_hours_cell, float) and pd.isna(person_hours_cell)):
        return {}
    if isinstance(person_hours_cell, dict):
        return person_hours_cell
    s = str(person_hours_cell).strip()
    if not s:
        return {}
    try:
        obj = json.loads(s)
        return obj if isinstance(obj, dict) else {}
    except Exception:
        return {}
def extract_wip_workers(ns_wip_row: pd.Series) -> list[str]:
    blob = parse_person_hours_json(ns_wip_row.get("Person Hours"))
    workers = []
    for k, v in blob.items():
        name = norm_name(k)
        if not is_real_person(name):
            continue
        if name == "0.0":
            continue
        actual = np.nan
        if isinstance(v, dict):
            actual = safe_float(v.get("actual"))
        else:
            actual = safe_float(v)
        if pd.notna(actual) and actual > 0:
            workers.append(name)
    workers = sorted(set(workers))
    return workers
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", required=True, help=r'Path to DBS NON WIP.xlsx')
    ap.add_argument("--ns_wip", required=True, help=r'Path to NS_WIP.csv')
    ap.add_argument("--ns_metrics", required=True, help=r'Path to NS_metrics.csv (must include Completed Hours)')
    ap.add_argument("--team", default="DBS")
    ap.add_argument("--out", default="ns_non_wip_activities.csv")
    args = ap.parse_args()
    xlsx_path = Path(args.xlsx)
    ns_wip_path = Path(args.ns_wip)
    ns_metrics_path = Path(args.ns_metrics)
    out_path = Path(args.out)
    team = args.team
    wip_df = pd.read_csv(ns_wip_path, dtype=str, keep_default_na=False)
    wip_df.columns = [" ".join(str(c).split()) for c in wip_df.columns]
    wip_df["period_date"] = pd.to_datetime(wip_df.get("period_date"), errors="coerce").dt.normalize()
    metrics_df = load_completed_hours(ns_metrics_path)
    sheets = pd.read_excel(xlsx_path, sheet_name=None, header=None)
    out_rows = []
    for sheet_name, ws in sheets.items():
        week = parse_week_from_sheet(sheet_name)
        if week is None or pd.isna(week):
            continue
        if ws.shape[0] < 18 or ws.shape[1] < 3:
            continue
        people_rows = read_people_block(ws)
        people = [r["name"] for r in people_rows]
        people_count = len(set(people))
        total_nonwip_hours, ooo_hours = extract_totals(ws)
        nonwip_by_person = build_nonwip_by_person(people_rows)
        nonwip_activities = build_nonwip_activities(ws, people_rows)
        wip_match = wip_df[(wip_df.get("team") == team) & (wip_df["period_date"] == week)]
        metrics_match = metrics_df[(metrics_df.get("team") == team) & (metrics_df["period_date"] == week)]
        wip_completed = np.nan
        if not wip_match.empty and "Completed Hours" in wip_match.columns:
            wip_completed = pd.to_numeric(wip_match.iloc[0]["Completed Hours"], errors="coerce")
        metrics_completed = np.nan
        if not metrics_match.empty and "Completed Hours" in metrics_match.columns:
            metrics_completed = pd.to_numeric(metrics_match.iloc[0]["Completed Hours"], errors="coerce")
        pct_in_wip = np.nan
        if pd.notna(wip_completed) and pd.notna(metrics_completed) and pd.notna(total_nonwip_hours):
            denom = float(metrics_completed) + float(total_nonwip_hours)
            pct_in_wip = float(wip_completed) / denom if denom != 0 else np.nan
        wip_workers = []
        if not wip_match.empty:
            wip_workers = extract_wip_workers(wip_match.iloc[0])
        wip_workers_count = len(wip_workers)
        ooo_map = {r["name"]: safe_float(r.get("C")) for r in people_rows}
        wip_workers_ooo_hours = float(
            round(sum(ooo_map.get(n, 0.0) for n in wip_workers if pd.notna(ooo_map.get(n, 0.0))), 2)
        )
        out_rows.append({
            "team": team,
            "period_date": week.date().isoformat(),
            "source_file": str(xlsx_path),
            "people_count": people_count,
            "total_non_wip_hours": float(round(total_nonwip_hours, 2)) if pd.notna(total_nonwip_hours) else np.nan,
            "OOO Hours": float(round(ooo_hours, 2)) if pd.notna(ooo_hours) else np.nan,
            "% in WIP": float(round(pct_in_wip, 6)) if pd.notna(pct_in_wip) else np.nan,
            "non_wip_by_person": json.dumps(nonwip_by_person, ensure_ascii=False),
            "non_wip_activities": json.dumps(nonwip_activities, ensure_ascii=False),
            "wip_workers": json.dumps(wip_workers, ensure_ascii=False),
            "wip_workers_count": wip_workers_count,
            "wip_workers_ooo_hours": wip_workers_ooo_hours,
        })
    out_df = pd.DataFrame(out_rows)
    if "period_date" in out_df.columns:
        out_df["period_date"] = pd.to_datetime(out_df["period_date"], errors="coerce").dt.normalize()
        out_df = out_df.sort_values(["team", "period_date"]).reset_index(drop=True)
    out_df.to_csv(out_path, index=False, encoding="utf-8-sig")
    print(f"Wrote {len(out_df)} rows -> {out_path}")
if __name__ == "__main__":
    main()