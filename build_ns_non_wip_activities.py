import json
import re
from pathlib import Path
import numpy as np
import pandas as pd
DEFAULTS = {
    "ns_wip": Path(r"C:\heijunka-dev\NS_WIP.csv"),
    "out": Path(r"C:\heijunka-dev\ns_non_wip_activities.csv"),
}
TEAM_SOURCES = {
    "DBS": {
        "xlsx": Path(r"C:\Users\wadec8\Medtronic PLC\DBS CQ Team - Documents\DBS NON WIP.xlsx"),
        "layout": {
            "people_start_row": 3,      # A4
            "totals_row": 17,           # row 18
            "activity_header_row": 2,   # row 3
            "activity_start_col": 3,    # D
            "activity_end_col": 35,     # AJ
            "min_rows": 18,
            "min_cols": 3,
        },
    },
    "SCS": {
        "xlsx": Path(r"C:\Users\wadec8\Medtronic PLC\Customer Quality SCS - Cell 17\SCS Non-D2D WIP Tracker 2026.xlsx"),
        "layout": {
            "people_start_row": 2,      # A3
            "totals_row": 25,           # row 26
            "activity_header_row": 1,   # row 2
            "activity_start_col": 3,    # D
            "activity_end_col": 34,     # AI
            "min_rows": 26,
            "min_cols": 3,
        },
    },
    "PH": {
        "xlsx": Path(r"C:\Users\wadec8\Medtronic PLC\Customer Quality Pelvic Health - Daily Tracker\Non-D2D WIP Tracker.xlsx"),
        "layout": {
            "people_start_row": 2,      # A3
            "totals_row": 16,           # row 26
            "activity_header_row": 1,   # row 2
            "activity_start_col": 3,    # D
            "activity_end_col": 34,     # AI
            "min_rows": 17,
            "min_cols": 3,
        },
    },
}
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
    if not s or s in {"-", "–", "—"}:
        return np.nan
    s = s.replace(",", "").replace("\u00a0", " ")
    m = re.search(r"[-+]?\d*\.?\d+", s)
    if not m:
        return np.nan
    try:
        return float(m.group(0))
    except Exception:
        return np.nan
def load_completed_hours(csv_path: Path) -> pd.DataFrame:
    df = pd.read_csv(csv_path, dtype=str, keep_default_na=False)
    df.columns = [" ".join(str(c).split()) for c in df.columns]
    if "period_date" in df.columns:
        df["period_date"] = pd.to_datetime(df["period_date"], errors="coerce").dt.normalize()
    if "Completed Hours" in df.columns:
        df["Completed Hours"] = pd.to_numeric(df["Completed Hours"], errors="coerce")
    return df
def parse_person_hours_json(person_hours_cell) -> dict:
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
        if not is_real_person(name) or name == "0.0":
            continue
        actual = safe_float(v.get("actual")) if isinstance(v, dict) else safe_float(v)
        if pd.notna(actual) and actual > 0:
            workers.append(name)
    return sorted(set(workers))
def read_people_block(ws: pd.DataFrame, start_i: int) -> list[dict]:
    rows = []
    for i in range(start_i, len(ws)):
        name = norm_name(ws.iat[i, 0] if 0 < ws.shape[1] else "")
        if not name:
            break
        if not is_real_person(name):
            continue
        b = safe_float(ws.iat[i, 1] if 1 < ws.shape[1] else np.nan)  # B
        c = safe_float(ws.iat[i, 2] if 2 < ws.shape[1] else np.nan)  # C
        if pd.isna(b): b = 0.0
        if pd.isna(c): c = 0.0
        rows.append({"row_i": i, "name": name, "B": b, "C": c})
    return rows
def extract_totals(ws: pd.DataFrame, totals_row_i: int) -> tuple[float, float]:
    b = safe_float(ws.iat[totals_row_i, 1] if 1 < ws.shape[1] else np.nan)
    c = safe_float(ws.iat[totals_row_i, 2] if 2 < ws.shape[1] else np.nan)
    total_nonwip = (b - c) if pd.notna(b) and pd.notna(c) else np.nan
    ooo = c if pd.notna(c) else np.nan
    return total_nonwip, ooo
def build_nonwip_by_person(people_rows: list[dict]) -> dict[str, float]:
    out = {}
    for r in people_rows:
        v = float(round(float(r.get("B", 0.0)) - float(r.get("C", 0.0)), 2))
        if v == 0.0:
            continue
        out[r["name"]] = v
    return out
def build_nonwip_activities(ws: pd.DataFrame, people_rows: list[dict], header_i: int, start_col: int, end_col: int) -> list[dict]:
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
            hours = safe_float(ws.iat[i, c] if c < ws.shape[1] else np.nan)
            if pd.isna(hours) or hours <= 0:
                continue
            activities.append({"name": name, "activity": activity, "hours": float(round(hours, 2))})
    return activities
def build_for_team(team: str, xlsx_path: Path, cfg: dict, wip_df: pd.DataFrame, metrics_df: pd.DataFrame) -> pd.DataFrame:
    if not xlsx_path.exists():
        print(f"[WARN] Missing XLSX for {team}: {xlsx_path}")
        return pd.DataFrame()
    sheets = pd.read_excel(xlsx_path, sheet_name=None, header=None)
    out_rows = []
    for sheet_name, ws in sheets.items():
        week = parse_week_from_sheet(sheet_name)
        if week is None or pd.isna(week):
            continue
        if ws.shape[0] < cfg["min_rows"] or ws.shape[1] < cfg["min_cols"]:
            continue
        people_rows = read_people_block(ws, start_i=cfg["people_start_row"])
        people_count = len(set(r["name"] for r in people_rows))
        total_nonwip_hours, ooo_hours = extract_totals(ws, totals_row_i=cfg["totals_row"])
        nonwip_by_person = build_nonwip_by_person(people_rows)
        nonwip_activities = build_nonwip_activities(
            ws,
            people_rows,
            header_i=cfg["activity_header_row"],
            start_col=cfg["activity_start_col"],
            end_col=cfg["activity_end_col"],
        )
        wip_match = wip_df[(wip_df["team"] == team) & (wip_df["period_date"] == week)]
        metrics_match = metrics_df[(metrics_df["team"] == team) & (metrics_df["period_date"] == week)]
        wip_completed = pd.to_numeric(wip_match.iloc[0].get("Completed Hours"), errors="coerce") if not wip_match.empty else np.nan
        metrics_completed = pd.to_numeric(metrics_match.iloc[0].get("Completed Hours"), errors="coerce") if not metrics_match.empty else np.nan
        pct_in_wip = np.nan
        if pd.notna(wip_completed) and pd.notna(metrics_completed) and pd.notna(total_nonwip_hours):
            denom = float(metrics_completed) + float(total_nonwip_hours)
            pct_in_wip = float(wip_completed) / denom if denom != 0 else np.nan
        wip_workers = extract_wip_workers(wip_match.iloc[0]) if not wip_match.empty else []
        wip_workers_count = len(wip_workers)
        ooo_map = {r["name"]: safe_float(r.get("C")) for r in people_rows}
        wip_workers_ooo_hours = float(round(sum(ooo_map.get(n, 0.0) for n in wip_workers), 2))
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
    df = pd.DataFrame(out_rows)
    if not df.empty:
        df["period_date"] = pd.to_datetime(df["period_date"], errors="coerce").dt.normalize()
        df = df.sort_values(["team", "period_date"]).reset_index(drop=True)
    return df
def main():
    wip_df = pd.read_csv(DEFAULTS["ns_wip"], dtype=str, keep_default_na=False)
    wip_df.columns = [" ".join(str(c).split()) for c in wip_df.columns]
    wip_df["period_date"] = pd.to_datetime(wip_df.get("period_date"), errors="coerce").dt.normalize()
    metrics_df = load_completed_hours(DEFAULTS["ns_wip"])
    built = []
    for team, info in TEAM_SOURCES.items():
        df_team = build_for_team(team, info["xlsx"], info["layout"], wip_df, metrics_df)
        if not df_team.empty:
            built.append(df_team)
    new_df = pd.concat(built, ignore_index=True) if built else pd.DataFrame()
    out_path = DEFAULTS["out"]
    if out_path.exists():
        old_df = pd.read_csv(out_path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
        old_df.columns = [" ".join(str(c).split()) for c in old_df.columns]
        if "period_date" in old_df.columns:
            old_df["period_date"] = pd.to_datetime(old_df["period_date"], errors="coerce").dt.normalize()
        combined = pd.concat([old_df, new_df], ignore_index=True)
        combined = combined.drop_duplicates(subset=["team", "period_date"], keep="last")
        combined = combined.sort_values(["team", "period_date"]).reset_index(drop=True)
    else:
        combined = new_df
    combined.to_csv(out_path, index=False, encoding="utf-8-sig")
    print(f"Wrote {len(combined)} rows -> {out_path}")
if __name__ == "__main__":
    main()