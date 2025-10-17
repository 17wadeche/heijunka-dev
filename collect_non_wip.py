# collect_non_wip.py
import csv
import json
from pathlib import Path
from datetime import datetime as _dt, date as _date, timedelta
import pandas as pd
from openpyxl import load_workbook
from dateutil import parser as dateparser
REPO_DIR = Path(r"C:\heijunka-dev")
REPO_CSV = REPO_DIR / "metrics_aggregate_dev.csv"
OUT_CSV  = REPO_DIR / "non_wip_activities.csv"
WEEKLY_HOURS_DEFAULT = 40.0
AORTIC_SHEET = "Individual (WIP-Non WIP)"
def _excel_serial_to_date(n):
    try:
        return (_dt(1899, 12, 30) + timedelta(days=float(n))).date()
    except Exception:
        return None
def _coerce_to_date(v):
    if isinstance(v, _dt): return v.date()
    if isinstance(v, _date): return v
    if isinstance(v, (int, float)): return _excel_serial_to_date(v)
    s = str(v).strip()
    if not s: return None
    try:
        return dateparser.parse(s, dayfirst=False, yearfirst=False).date()
    except Exception:
        return None
_BAD_EXACT = {
    "team member",
    "total available hours",
    "total pitches",
    "weeks production output",
    "workflow",
}
_BAD_PREFIX = (
    "release date",  # e.g., "Release Date: 04 Aug 2024"
    "revision",      # e.g., "REVISION: 001"
    "week starting", # e.g., "Week Starting: 08/05/2024"
)
def _looks_like_bad_header(raw: str) -> bool:
    if raw is None:
        return True
    s = str(raw).strip()
    if not s:
        return True
    s_nf = s.rstrip(":").strip()
    k = s_nf.casefold()
    if k in _BAD_EXACT:
        return True
    for pref in _BAD_PREFIX:
        if k.startswith(pref):
            return True
    if k in {"#ref!", "nan", "0", "-", "–", "—"}:
        return True
    return False
def _clean_name(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip().strip('"').strip("'")
    return "" if _looks_like_bad_header(s) else s
def _read_names_from_individual_sheet_xlsx(path: Path, sheet_name=AORTIC_SHEET) -> list[str]:
    try:
        wb = load_workbook(path, data_only=True, read_only=True)
    except Exception:
        return []
    if sheet_name not in wb.sheetnames:
        return []
    ws = wb[sheet_name]
    names = []
    for r in ws.iter_rows(min_row=1, max_row=200, min_col=1, max_col=1, values_only=True):
        nm = _clean_name(r[0])
        if nm:
            names.append(nm)
    seen, uniq = set(), []
    for n in names:
        k = n.casefold()
        if k not in seen:
            seen.add(k); uniq.append(n)
    return uniq
def _source_file_only(s: str) -> str:
    return s.split(" :: ", 1)[0].strip()
def _parse_person_hours_cell(s: str | None) -> dict[str, float]:
    if not s:
        return {}
    try:
        obj = json.loads(s)
        out = {}
        for name, vals in (obj or {}).items():
            try:
                out[str(name).strip()] = float((vals or {}).get("actual") or 0.0)
            except Exception:
                out[str(name).strip()] = 0.0
        return out
    except Exception:
        return {}
def _lookup_actual_hours(ph_by_name: dict[str, float], person: str) -> float:
    if person in ph_by_name:
        return float(ph_by_name[person])
    want = person.casefold()
    for k, v in ph_by_name.items():
        if k.casefold() == want:
            return float(v)
    return 0.0
def main():
    if not REPO_CSV.exists():
        raise FileNotFoundError(f"metrics CSV not found: {REPO_CSV}")
    df = pd.read_csv(REPO_CSV, dtype=str, keep_default_na=False)
    df = df[(df.get("team","") == "Aortic") & (df.get("source_file","") != "")]
    if df.empty:
        print("[non-wip] No Aortic rows found in metrics_aggregate_dev.csv")
        return
    df["period_date"] = df["period_date"].apply(_coerce_to_date)
    df = df.dropna(subset=["period_date"])
    df["period_date"] = pd.to_datetime(df["period_date"]).dt.date
    df["source_file_only"] = df["source_file"].apply(_source_file_only)
    ph_index: dict[tuple, dict[str, float]] = {}
    for _, r in df.iterrows():
        key = (r["period_date"], r["source_file_only"])
        ph = _parse_person_hours_cell(r.get("Person Hours"))
        if key not in ph_index:
            ph_index[key] = ph
        else:
            ph_index[key].update(ph)
    unique_refs = df[["team","period_date","source_file_only"]].drop_duplicates()
    out_rows = []
    for _, row in unique_refs.iterrows():
        team = row["team"]
        period_date = row["period_date"]
        src = row["source_file_only"]
        p = Path(src)
        if not p.exists() or p.suffix.lower() not in (".xlsx", ".xlsm"):
            continue
        people = _read_names_from_individual_sheet_xlsx(p)
        if not people:
            continue
        ph_by_name = ph_index.get((period_date, src), {})  # may be empty
        per_person_non_wip = {}
        total_non_wip = 0.0
        total_wip_capped = 0.0
        for person in people:
            wip_actual = _lookup_actual_hours(ph_by_name, person)
            non_wip = max(0.0, WEEKLY_HOURS_DEFAULT - float(wip_actual))
            per_person_non_wip[person] = round(non_wip, 2)
            total_non_wip += non_wip
            total_wip_capped += min(float(wip_actual), WEEKLY_HOURS_DEFAULT)
        people_count = len(people)
        weekly_total_available = WEEKLY_HOURS_DEFAULT * people_count if people_count > 0 else 0.0
        pct_in_wip = (total_wip_capped / weekly_total_available * 100.0) if weekly_total_available > 0 else 0.0
        pct_in_wip = round(pct_in_wip, 2)  # percent value, e.g., 73.5
        out_rows.append({
            "team": team,
            "period_date": period_date.isoformat(),
            "source_file": src,
            "people_count": people_count,
            "total_non_wip_hours": round(total_non_wip, 2),
            "% in WIP": pct_in_wip,
            "non_wip_by_person": json.dumps(per_person_non_wip, ensure_ascii=False),
        })
    if not out_rows:
        print("[non-wip] No weekly rows produced for Aortic.")
        return
    OUT_CSV.parent.mkdir(parents=True, exist_ok=True)
    cols = [
        "team",
        "period_date",
        "source_file",
        "people_count",
        "total_non_wip_hours",
        "% in WIP",
        "non_wip_by_person",
    ]
    with OUT_CSV.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(out_rows)
    print(f"[non-wip] Wrote {len(out_rows)} rows to {OUT_CSV}")
if __name__ == "__main__":
    main()
