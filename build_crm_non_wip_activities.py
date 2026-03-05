from __future__ import annotations
import argparse
import csv
import datetime as _dt
import json
import os
import re
from typing import Any, Dict, Iterable, List, Optional, Tuple
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
TEAM_BY_SOURCE: Dict[str, str] = {
    r"C:\Users\wadec8\Medtronic PLC\MCS COS Transformation - VMB Scheduling\Heijunka Current.xlsm": "MCS",
}
TEAM_BY_BASENAME: Dict[str, str] = {
    "Heijunka Current.xlsm": "MCS",
}
_AVAIL_PAT = re.compile(r"\bavailability\b", re.IGNORECASE)
_PROD_PAT = re.compile(r"\b(production|product)\s+analysis\b", re.IGNORECASE)
_MONTH_MAP = {
    "jan": 1, "january": 1,
    "feb": 2, "february": 2,
    "mar": 3, "march": 3,
    "apr": 4, "april": 4,
    "may": 5,
    "jun": 6, "june": 6,
    "jul": 7, "july": 7,
    "aug": 8, "august": 8,
    "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10,
    "nov": 11, "november": 11,
    "dec": 12, "december": 12,
}
CSV_COLUMNS = [
    "team",
    "period_date",
    "people_count",
    "total_non_wip_hours",
    "OOO Hours",
    "% in WIP",
    "non_wip_by_person",
    "non_wip_activities",
    "wip_workers",
    "wip_workers_count",
    "wip_workers_ooo_hours",
]
def _norm_path(p: str) -> str:
    return os.path.normpath(p)
def team_for_source(path: str) -> str:
    np = _norm_path(path)
    if np in TEAM_BY_SOURCE:
        return TEAM_BY_SOURCE[np]
    base = os.path.basename(np)
    return TEAM_BY_BASENAME.get(base, "")
def _norm_text(x: str) -> str:
    return re.sub(r"\s+", " ", (x or "").strip()).lower()
def is_excluded_person(name: str) -> bool:
    n = _norm_text(name)
    return n in {"", "do not use", "team member(s)"}
def parse_period_date_from_sheetname(sheet_name: str, *, default_year: Optional[int] = None) -> Optional[_dt.date]:
    if default_year is None:
        default_year = _dt.date.today().year
    s = sheet_name.strip()
    m = re.search(r"(\d{1,2})\s*[-/ ]\s*([A-Za-z]{3,9})(?:\s*[-/ ]\s*(\d{2,4}))?\s*$", s)
    if not m:
        return None
    day = int(m.group(1))
    mon_raw = m.group(2).strip().lower()
    year_raw = m.group(3)
    if mon_raw not in _MONTH_MAP:
        return None
    month = _MONTH_MAP[mon_raw]
    if year_raw:
        year = int(year_raw)
        if year < 100:
            year += 2000
    else:
        year = default_year
    try:
        return _dt.date(year, month, day)
    except ValueError:
        return None
def iso_date(d: Optional[_dt.date]) -> str:
    return d.isoformat() if isinstance(d, _dt.date) else ""
def _cell_number(v: Any) -> Optional[float]:
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        vs = v.strip()
        if vs == "":
            return None
        try:
            return float(vs)
        except ValueError:
            return None
    return None
def sum_range(ws: Worksheet, cell1: str, cell2: str) -> float:
    total = 0.0
    for row in ws[cell1:cell2]:
        for c in row:
            n = _cell_number(c.value)
            if n is not None:
                total += n
    return total
def read_merged_value(ws: Worksheet, top_left_cell: str) -> str:
    v = ws[top_left_cell].value
    return str(v).strip() if v is not None else ""
def find_sheets_by_period(wb, *, kind: str) -> Dict[_dt.date, str]:
    out: Dict[_dt.date, str] = {}
    for name in wb.sheetnames:
        if kind == "availability":
            if not _AVAIL_PAT.search(name):
                continue
        elif kind == "production":
            if not _PROD_PAT.search(name):
                continue
        else:
            raise ValueError("kind must be availability or production")
        d = parse_period_date_from_sheetname(name)
        if d:
            out[d] = name
    return out
def iter_prod_rows(ws_prod: Worksheet, start_row: int = 7, end_row: int = 406) -> Iterable[Tuple[int, str, str, Optional[float], Optional[float]]]:
    maxr = min(ws_prod.max_row, end_row)
    for r in range(start_row, maxr + 1):
        person = ws_prod[f"D{r}"].value
        station = ws_prod[f"E{r}"].value
        target = _cell_number(ws_prod[f"F{r}"].value)
        output = _cell_number(ws_prod[f"G{r}"].value)
        p = str(person).strip() if person is not None else ""
        s = str(station).strip() if station is not None else ""
        if p == "" and s == "" and target is None and output is None:
            continue
        yield (r, p, s, target, output)
def wip_workers_from_prod(ws_prod: Worksheet) -> List[str]:
    seen = set()
    for r, person, station, target, output in iter_prod_rows(ws_prod, start_row=7):
        if output is None:
            continue
        if is_excluded_person(person):
            continue
        seen.add(person)
    return sorted(seen)
NON_WIP_SPECS = [
    ("A13", ("B18", "F21"), "A", ("B17", "F17")),  # top-left block (often Do Not Use)
    ("A23", ("B28", "F31"), "A", ("B27", "F27")),  # left block lower
    ("H3",  ("I8", "M11"), "H", ("I7", "M7")),     # middle top
    ("H23", ("I28", "M31"), "H", ("I27", "M27")),  # middle
    ("H33", ("I38", "M41"), "H", ("I37", "M37")),  # middle lower
    ("O3",  ("P8", "T11"), "O", ("P7", "T7")),     # right top
    ("O13", ("P18", "T21"), "O", ("P17", "T17")),  # right
    ("O23", ("P28", "T31"), "O", ("P27", "T27")),  # right
    ("O43", ("P48", "T51"), "O", ("P47", "T47")),  # right lower
    ("O53", ("P58", "T61"), "O", ("P57", "T57")),  # right lower
]
def compute_total_non_wip_hours(ws_av: Worksheet) -> float:
    total = 0.0
    for _, (c1, c2), _, _ in NON_WIP_SPECS:
        total += sum_range(ws_av, c1, c2)
    return total
def compute_total_ooo_hours(ws_av: Worksheet) -> float:
    total = 0.0
    for _, _, _, (c1, c2) in NON_WIP_SPECS:
        total += sum_range(ws_av, c1, c2)
    return total
def compute_non_wip_by_person(ws_av: Worksheet) -> Dict[str, float]:
    out: Dict[str, float] = {}
    for name_cell, (c1, c2), _, _ in NON_WIP_SPECS:
        name = read_merged_value(ws_av, name_cell)
        if not name or is_excluded_person(name):
            continue
        out[name] = out.get(name, 0.0) + sum_range(ws_av, c1, c2)
    return out
def compute_non_wip_activities(ws_av: Worksheet) -> List[Dict[str, Any]]:
    agg: Dict[Tuple[str, str], float] = {}
    for name_cell, (c1, c2), label_col, _ in NON_WIP_SPECS:
        name = read_merged_value(ws_av, name_cell)
        if not name or is_excluded_person(name):
            continue
        min_row = ws_av[c1].row
        max_row = ws_av[c2].row
        min_col = ws_av[c1].column
        max_col = ws_av[c2].column
        for r in range(min_row, max_row + 1):
            activity_label = ws_av[f"{label_col}{r}"].value
            activity = str(activity_label).strip() if activity_label is not None else ""
            if activity == "":
                continue
            if _norm_text(activity) in {"weekday", "hours available", "available for cos1", "wip block?"}:
                continue
            hours_sum = 0.0
            for c in range(min_col, max_col + 1):
                val = _cell_number(ws_av.cell(row=r, column=c).value)
                if val is None or val == 0:
                    continue
                hours_sum += float(val)
            if hours_sum:
                key = (name, activity)
                agg[key] = agg.get(key, 0.0) + hours_sum
    return [
        {"name": n, "activity": a, "hours": float(h)}
        for (n, a), h in sorted(agg.items(), key=lambda x: (x[0][0].lower(), x[0][1].lower()))
    ]
def compute_ooo_by_person(ws_av: Worksheet) -> Dict[str, float]:
    out: Dict[str, float] = {}
    for name_cell, _, _, (c1, c2) in NON_WIP_SPECS:
        name = read_merged_value(ws_av, name_cell)
        if not name or is_excluded_person(name):
            continue
        out[name] = out.get(name, 0.0) + sum_range(ws_av, c1, c2)
    return out
def load_completed_hours_from_crm_wip(crm_wip_csv: str) -> Dict[Tuple[str, str], float]:
    out: Dict[Tuple[str, str], float] = {}
    if not os.path.exists(crm_wip_csv):
        return out
    with open(crm_wip_csv, "r", encoding="utf-8-sig", newline="") as fp:
        r = csv.DictReader(fp)
        for row in r:
            team = (row.get("team") or "").strip()
            period = (row.get("period_date") or "").strip()
            ch = row.get("Completed Hours")
            try:
                val = float(ch) if ch not in (None, "") else None
            except ValueError:
                val = None
            if team and period and val is not None:
                out[(team, period)] = val
    return out
def safe_div(n: float, d: float) -> Optional[float]:
    if d == 0:
        return None
    return n / d
def scrape_one_workbook(path: str, completed_hours_lookup: Dict[Tuple[str, str], float]) -> List[Dict[str, Any]]:
    team = team_for_source(path)
    wb = load_workbook(path, data_only=True)
    avail_sheets = find_sheets_by_period(wb, kind="availability")
    prod_sheets = find_sheets_by_period(wb, kind="production")
    periods = sorted(set(avail_sheets.keys()) | set(prod_sheets.keys()))
    rows: List[Dict[str, Any]] = []
    for period in periods:
        ws_av = wb[avail_sheets[period]] if period in avail_sheets else None
        ws_prod = wb[prod_sheets[period]] if period in prod_sheets else None
        if ws_av is None:
            continue
        period_iso = iso_date(period)
        total_non_wip_hours = compute_total_non_wip_hours(ws_av)
        ooo_hours = compute_total_ooo_hours(ws_av)
        non_wip_by_person = compute_non_wip_by_person(ws_av)
        non_wip_activities = compute_non_wip_activities(ws_av)
        wip_workers: List[str] = wip_workers_from_prod(ws_prod) if ws_prod is not None else []
        wip_workers_count = len(wip_workers)
        ooo_by_person = compute_ooo_by_person(ws_av)
        wip_workers_ooo_hours = float(sum(float(ooo_by_person.get(p, 0.0)) for p in wip_workers))
        completed_hours = completed_hours_lookup.get((team, period_iso), 0.0)
        pct_in_wip = safe_div(float(completed_hours), float(completed_hours) + float(total_non_wip_hours))
        row = {
            "team": team,
            "period_date": period_iso,
            "people_count": 10,
            "total_non_wip_hours": float(total_non_wip_hours),
            "OOO Hours": float(ooo_hours),
            "% in WIP": float(pct_in_wip) if pct_in_wip is not None else "",
            "non_wip_by_person": json.dumps(non_wip_by_person, ensure_ascii=False),
            "non_wip_activities": json.dumps(non_wip_activities, ensure_ascii=False),
            "wip_workers": json.dumps(wip_workers, ensure_ascii=False),
            "wip_workers_count": int(wip_workers_count),
            "wip_workers_ooo_hours": wip_workers_ooo_hours,
        }
        rows.append(row)
    return rows
def main() -> int:
    default_path = r"C:\Users\wadec8\Medtronic PLC\MCS COS Transformation - VMB Scheduling\Heijunka Current.xlsm"
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="*", help="Excel workbook(s) to scrape (.xlsx/.xlsm).")
    ap.add_argument("--crm_wip", default="CRM_WIP.csv", help="Path to CRM_WIP.csv (default: CRM_WIP.csv).")
    ap.add_argument("--out", default="crm_non_wip_activities.csv", help="Output CSV path.")
    args = ap.parse_args()
    files = args.files or [default_path]
    completed_hours_lookup = load_completed_hours_from_crm_wip(args.crm_wip)
    all_rows: List[Dict[str, Any]] = []
    for f in files:
        if not os.path.exists(f):
            continue
        all_rows.extend(scrape_one_workbook(f, completed_hours_lookup))
    with open(args.out, "w", newline="", encoding="utf-8") as fp:
        w = csv.DictWriter(fp, fieldnames=CSV_COLUMNS)
        w.writeheader()
        for r in all_rows:
            w.writerow({k: r.get(k, "") for k in CSV_COLUMNS})
    print(f"Wrote {len(all_rows)} row(s) to {args.out}")
    return 0
if __name__ == "__main__":
    raise SystemExit(main())