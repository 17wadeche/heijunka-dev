from __future__ import annotations
import argparse
import csv
import datetime as _dt
import json
import os
from typing import Any, Dict, Iterable, List, Optional, Tuple
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
TEAM_BY_SOURCE: Dict[str, str] = {
    r"C:\Users\wadec8\Medtronic PLC\CQXM RI-Heijunka live spreadsheet shared - Documents\WIP+Non-WIP Heijunka Template CQXM  VSS 2026 03 .xlsm": "VSS",
}
TEAM_BY_BASENAME: Dict[str, str] = {
    "WIP+Non-WIP Heijunka Template CQXM  VSS 2026 03 .xlsm": "VSS",
}
DEFAULT_FILES: List[str] = [
    r"C:\Users\wadec8\Medtronic PLC\CQXM RI-Heijunka live spreadsheet shared - Documents\WIP+Non-WIP Heijunka Template CQXM  VSS 2026 03 .xlsm",
]
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
AVAILABILITY_SHEET = "Available WIP+Non-WIP Hours"
def _norm_path(p: str) -> str:
    return os.path.normpath(p)
def team_for_source(path: str) -> str:
    np = _norm_path(path)
    if np in TEAM_BY_SOURCE:
        return TEAM_BY_SOURCE[np]
    return TEAM_BY_BASENAME.get(os.path.basename(np), "")
def _as_text(v: Any) -> str:
    return str(v).strip() if v is not None else ""
def _cell_number(v: Any) -> Optional[float]:
    if v is None:
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.strip().replace(",", "")
        if not s:
            return None
        try:
            return float(s)
        except ValueError:
            return None
    return None
def _as_date(v: Any) -> Optional[_dt.date]:
    if v is None:
        return None
    if isinstance(v, _dt.datetime):
        return v.date()
    if isinstance(v, _dt.date):
        return v
    if isinstance(v, str):
        s = v.strip()
        if not s:
            return None
        for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%d-%b-%y", "%d-%b-%Y"):
            try:
                return _dt.datetime.strptime(s, fmt).date()
            except ValueError:
                pass
    return None
def iso_date(d: Optional[_dt.date]) -> str:
    return d.isoformat() if isinstance(d, _dt.date) else ""
def dumps_json(obj: Any) -> str:
    return json.dumps(obj, ensure_ascii=False)
def safe_div(n: float, d: float) -> Optional[float]:
    if d in (0, 0.0) or d is None:
        return None
    return n / d
def iter_available_blocks(ws: Worksheet) -> Iterable[Tuple[int, _dt.date]]:
    for row in range(1, ws.max_row + 1):
        label = _as_text(ws.cell(row=row, column=1).value)
        if label.lower() == "week starting:":
            week_val = ws.cell(row=row, column=2).value
            week_date = _as_date(week_val)
            if week_date is not None:
                yield row, week_date
def find_next_week_row(starts: List[Tuple[int, _dt.date]], idx: int, max_row: int) -> int:
    if idx + 1 < len(starts):
        return starts[idx + 1][0] - 1
    return max_row
def find_actual_non_wip_header_row(ws: Worksheet, start_row: int, end_row: int) -> Optional[int]:
    for row in range(start_row, min(end_row, start_row + 25) + 1):
        row_vals = [_as_text(ws.cell(row=row, column=col).value).lower() for col in range(12, 23)]
        if "audit" in row_vals and "ooo" in row_vals:
            return row
    return None
def find_column_by_header(ws: Worksheet, header_row: int, header_name: str, start_col: int = 12, end_col: int = 22) -> Optional[int]:
    target = header_name.strip().lower()
    for col in range(start_col, end_col + 1):
        val = _as_text(ws.cell(row=header_row, column=col).value).strip().lower()
        if val == target:
            return col
    return None
def parse_available_sheet(ws: Worksheet) -> Dict[_dt.date, Dict[str, Any]]:
    results: Dict[_dt.date, Dict[str, Any]] = {}
    starts = list(iter_available_blocks(ws))
    for idx, (start_row, week_date) in enumerate(starts):
        end_row = find_next_week_row(starts, idx, ws.max_row)
        header_row = find_actual_non_wip_header_row(ws, start_row, end_row)
        people_count_names: List[str] = []
        non_wip_by_person: Dict[str, float] = {}
        non_wip_activities: List[Dict[str, Any]] = []
        ooo_by_person: Dict[str, float] = {}
        total_non_wip_hours = 0.0
        total_ooo_hours = 0.0
        if header_row is None:
            results[week_date] = {
                "people_count": 0,
                "total_non_wip_hours": 0.0,
                "ooo_hours": 0.0,
                "non_wip_by_person": {},
                "non_wip_activities": [],
                "ooo_by_person": {},
            }
            continue
        audit_col = find_column_by_header(ws, header_row, "Audit")
        last_col = find_column_by_header(ws, header_row, "Non-D2D")
        ooo_col = find_column_by_header(ws, header_row, "OOO")
        if audit_col is None or last_col is None or ooo_col is None:
            results[week_date] = {
                "people_count": 0,
                "total_non_wip_hours": 0.0,
                "ooo_hours": 0.0,
                "non_wip_by_person": {},
                "non_wip_activities": [],
                "ooo_by_person": {},
            }
            continue
        for row in range(header_row + 1, end_row + 1):
            name = _as_text(ws.cell(row=row, column=12).value)
            if not name:
                continue
            lower_name = name.strip().lower()
            if lower_name == "total":
                total_non_wip_hours = (_cell_number(ws.cell(row=row, column=last_col).value) or 0.0) / 60.0
                total_ooo_hours = (_cell_number(ws.cell(row=row, column=ooo_col).value) or 0.0) / 60.0
                continue
            if lower_name == "x":
                continue
            people_count_names.append(name)
            person_total_non_wip = (_cell_number(ws.cell(row=row, column=last_col).value) or 0.0) / 60.0
            person_ooo = (_cell_number(ws.cell(row=row, column=ooo_col).value) or 0.0) / 60.0
            non_wip_by_person[name] = person_total_non_wip
            ooo_by_person[name] = person_ooo
            for col in range(audit_col, last_col + 1):
                header = _as_text(ws.cell(row=header_row, column=col).value)
                mins = _cell_number(ws.cell(row=row, column=col).value) or 0.0
                hours = mins / 60.0
                if hours > 0:
                    non_wip_activities.append({
                        "name": name,
                        "activity": header,
                        "hours": hours,
                    })
        results[week_date] = {
            "people_count": len(set(people_count_names)),
            "total_non_wip_hours": total_non_wip_hours,
            "ooo_hours": total_ooo_hours,
            "non_wip_by_person": non_wip_by_person,
            "non_wip_activities": non_wip_activities,
            "ooo_by_person": ooo_by_person,
        }
    return results
def _norm_team(s: str) -> str:
    return (s or "").strip().upper()
def _norm_period_date(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return ""
    try:
        return _dt.date.fromisoformat(s).isoformat()
    except ValueError:
        return s
def script_dir() -> str:
    return os.path.dirname(os.path.abspath(__file__))
def parse_people_in_wip(raw: str) -> List[str]:
    s = (raw or "").strip()
    if not s:
        return []
    try:
        val = json.loads(s)
        if isinstance(val, list):
            return [str(x).strip() for x in val if str(x).strip()]
    except Exception:
        pass
    return []
def load_wip_lookup(path: str) -> Dict[Tuple[str, str], Dict[str, Any]]:
    lut: Dict[Tuple[str, str], Dict[str, Any]] = {}
    if not os.path.exists(path):
        return lut
    with open(path, "r", newline="", encoding="utf-8-sig") as fp:
        reader = csv.DictReader(fp)
        for row in reader:
            team = _norm_team(row.get("team", ""))
            period = _norm_period_date(row.get("period_date", ""))
            completed_hours = _cell_number(row.get("Completed Hours"))
            people_in_wip = parse_people_in_wip(row.get("People in WIP", ""))
            if team and period:
                lut[(team, period)] = {
                    "Completed Hours": float(completed_hours or 0.0),
                    "People in WIP": people_in_wip,
                }
    return lut
def find_wip_csv() -> Optional[str]:
    base = script_dir()
    candidates = [
        os.path.join(base, "MS_WIP.csv")
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return None
def blank_row_for_missing_file(path: str) -> Dict[str, Any]:
    return {
        "team": team_for_source(path),
        "period_date": "",
        "people_count": "",
        "total_non_wip_hours": "",
        "OOO Hours": "",
        "% in WIP": "",
        "non_wip_by_person": "",
        "non_wip_activities": "",
        "wip_workers": "",
        "wip_workers_count": "",
        "wip_workers_ooo_hours": "",
    }
def scrape_one_workbook(path: str, wip_lut: Dict[Tuple[str, str], Dict[str, Any]]) -> List[Dict[str, Any]]:
    team = team_for_source(path)
    wb = load_workbook(path, data_only=True)
    if AVAILABILITY_SHEET not in wb.sheetnames:
        return [{
            "team": team,
            "period_date": "",
            "people_count": "",
            "total_non_wip_hours": "",
            "OOO Hours": "",
            "% in WIP": "",
            "non_wip_by_person": "",
            "non_wip_activities": "",
            "wip_workers": "",
            "wip_workers_count": "",
            "wip_workers_ooo_hours": "",
        }]
    available_by_week = parse_available_sheet(wb[AVAILABILITY_SHEET])
    rows: List[Dict[str, Any]] = []
    for period in sorted(available_by_week.keys()):
        av = available_by_week[period]
        key = (_norm_team(team), iso_date(period))
        wip_info = wip_lut.get(key, {})
        completed_hours = float(wip_info.get("Completed Hours", 0.0) or 0.0)
        wip_workers = list(wip_info.get("People in WIP", []) or [])
        wip_workers_set = {x.strip() for x in wip_workers if str(x).strip()}
        total_non_wip_hours = float(av.get("total_non_wip_hours", 0.0) or 0.0)
        pct_in_wip = safe_div(completed_hours, completed_hours + total_non_wip_hours)
        ooo_by_person = av.get("ooo_by_person", {}) or {}
        wip_workers_ooo_hours = 0.0
        for worker in wip_workers_set:
            wip_workers_ooo_hours += float(ooo_by_person.get(worker, 0.0) or 0.0)
        row = {
            "team": team,
            "period_date": iso_date(period),
            "people_count": av.get("people_count", 0),
            "total_non_wip_hours": total_non_wip_hours,
            "OOO Hours": float(av.get("ooo_hours", 0.0) or 0.0),
            "% in WIP": float(pct_in_wip) if pct_in_wip is not None else "",
            "non_wip_by_person": dumps_json(av.get("non_wip_by_person", {})),
            "non_wip_activities": dumps_json(av.get("non_wip_activities", [])),
            "wip_workers": dumps_json(wip_workers),
            "wip_workers_count": len(wip_workers),
            "wip_workers_ooo_hours": wip_workers_ooo_hours,
        }
        rows.append(row)
    return rows
def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="*", help="Optional workbook(s) to scrape. If omitted, uses DEFAULT_FILES in the script.")
    ap.add_argument("--out", default="ms_non_wip_activities.csv", help="Output CSV path.")
    args = ap.parse_args()
    wip_csv = find_wip_csv()
    wip_lut: Dict[Tuple[str, str], Dict[str, Any]] = {}
    if wip_csv:
        wip_lut = load_wip_lookup(wip_csv)
    files = args.files or DEFAULT_FILES
    all_rows: List[Dict[str, Any]] = []
    for path in files:
        if not os.path.exists(path):
            all_rows.append(blank_row_for_missing_file(path))
            continue
        all_rows.extend(scrape_one_workbook(path, wip_lut))
    with open(args.out, "w", newline="", encoding="utf-8") as fp:
        writer = csv.DictWriter(fp, fieldnames=CSV_COLUMNS)
        writer.writeheader()
        for row in all_rows:
            writer.writerow({k: row.get(k, "") for k in CSV_COLUMNS})
    print(f"Wrote {len(all_rows)} row(s) to {args.out}")
    return 0
if __name__ == "__main__":
    raise SystemExit(main())