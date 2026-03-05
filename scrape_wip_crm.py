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
def _norm_path(p: str) -> str:
    return os.path.normpath(p)
def team_for_source(path: str) -> str:
    np = _norm_path(path)
    if np in TEAM_BY_SOURCE:
        return TEAM_BY_SOURCE[np]
    base = os.path.basename(np)
    return TEAM_BY_BASENAME.get(base, "")
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
def compute_total_available_hours(ws_av: Worksheet) -> float:
    blocks = [
        ("B", "F"),
        ("I", "M"),
        ("P", "T"),
    ]
    rows = [5, 15, 25, 35, 45, 55]
    total = 0.0
    for r in rows:
        for c1, c2 in blocks:
            total += sum_range(ws_av, f"{c1}{r}", f"{c2}{r}")
    return total
def iter_prod_rows(ws_prod: Worksheet, start_row: int = 7) -> Iterable[Tuple[int, str, str, Optional[float], Optional[float]]]:
    maxr = min(ws_prod.max_row, 406)
    for r in range(start_row, maxr + 1):
        person = ws_prod[f"D{r}"].value
        cell_station = ws_prod[f"E{r}"].value
        target = _cell_number(ws_prod[f"F{r}"].value)
        output = _cell_number(ws_prod[f"G{r}"].value)
        p = str(person).strip() if person is not None else ""
        cs = str(cell_station).strip() if cell_station is not None else ""
        if p == "" and cs == "" and target is None and output is None:
            continue
        yield (r, p, cs, target, output)
def compute_completed_hours(ws_prod: Worksheet) -> Tuple[float, Dict[str, float]]:
    total = 0.0
    by_person: Dict[str, float] = {}
    for r, person, cell_station, target, output in iter_prod_rows(ws_prod, start_row=7):
        if output is None:
            continue
        is_promoted = (cell_station == "Promoted PE - Initial MDR")
        h = 4.0 if is_promoted else 1.0
        total += h
        if person:
            by_person[person] = by_person.get(person, 0.0) + h
    return total, by_person
def compute_target_actual_output(ws_prod: Worksheet) -> Tuple[float, float]:
    targ = 0.0
    act = 0.0
    for r, person, cell_station, target, output in iter_prod_rows(ws_prod, start_row=7):
        if output is None:
            continue
        act += output
        if target is not None:
            targ += target
    return targ, act
def unique_people_in_wip(ws_prod: Worksheet) -> List[str]:
    seen = set()
    for r, person, cell_station, target, output in iter_prod_rows(ws_prod, start_row=7):
        if output is None:
            continue
        if person:
            seen.add(person)
    return sorted(seen)
def compute_outputs_by_person(ws_prod: Worksheet) -> Dict[str, Dict[str, float]]:
    out: Dict[str, Dict[str, float]] = {}
    for r, person, cell_station, target, output in iter_prod_rows(ws_prod, start_row=7):
        if output is None or not person:
            continue
        if person not in out:
            out[person] = {"output": 0.0, "target": 0.0}
        out[person]["output"] += output
        if target is not None:
            out[person]["target"] += target
    return out
def compute_outputs_by_station(ws_prod: Worksheet) -> Dict[str, Dict[str, float]]:
    out: Dict[str, Dict[str, float]] = {}
    for r, person, cell_station, target, output in iter_prod_rows(ws_prod, start_row=7):
        if output is None or not cell_station:
            continue
        if cell_station not in out:
            out[cell_station] = {"output": 0.0, "target": 0.0}
        out[cell_station]["output"] += output
        if target is not None:
            out[cell_station]["target"] += target
    return out
def compute_station_hours(ws_prod: Worksheet) -> Tuple[Dict[str, float], Dict[str, Dict[str, float]]]:
    station_hours: Dict[str, float] = {}
    station_hours_by_person: Dict[str, Dict[str, float]] = {}
    for r, person, cell_station, target, output in iter_prod_rows(ws_prod, start_row=7):
        if not person:
            continue
        if not cell_station:
            continue
        if person.strip().lower() == "do not use":
            continue
        if person.strip().lower() == "team member(s)":
            continue
        if cell_station.strip().lower() == "cell/station":
            continue
        h = 2.0 if cell_station == "Promoted PE - Initial MDR" else 1.0
        station_hours[cell_station] = station_hours.get(cell_station, 0.0) + h
        if cell_station not in station_hours_by_person:
            station_hours_by_person[cell_station] = {}
        station_hours_by_person[cell_station][person] = station_hours_by_person[cell_station].get(person, 0.0) + h
    return station_hours, station_hours_by_person
def compute_output_by_station_by_person(ws_prod: Worksheet) -> Dict[str, Dict[str, float]]:
    out: Dict[str, Dict[str, float]] = {}
    for r, person, cell_station, target, output in iter_prod_rows(ws_prod, start_row=7):
        if output is None or not person or not cell_station:
            continue
        out.setdefault(cell_station, {})
        out[cell_station][person] = out[cell_station].get(person, 0.0) + output
    return out
def compute_uplh_by_station_by_person(
    output_by_station_by_person: Dict[str, Dict[str, float]],
    hours_by_station_by_person: Dict[str, Dict[str, float]],
) -> Dict[str, Dict[str, float]]:
    out: Dict[str, Dict[str, float]] = {}
    for station, people_outputs in output_by_station_by_person.items():
        for person, out_val in people_outputs.items():
            hrs = hours_by_station_by_person.get(station, {}).get(person, 0.0)
            if hrs and hrs != 0.0:
                out.setdefault(station, {})
                out[station][person] = out_val / hrs
    return out
def compute_person_available_hours(ws_av: Worksheet) -> Dict[str, float]:
    pairs = [
        ("A13", ("B15", "F15")),
        ("A23", ("B25", "F25")),
        ("H3",  ("I5", "M5")),
        ("H23", ("I25", "M25")),
        ("H33", ("I35", "M35")),
        ("O3",  ("O5", "Q5")),
        ("O13", ("O15", "Q15")),
        ("O23", ("O25", "Q25")),
        ("O43", ("O45", "Q45")),
        ("O53", ("O55", "Q55")),
    ]
    out: Dict[str, float] = {}
    for name_cell, (c1, c2) in pairs:
        name = read_merged_value(ws_av, name_cell)
        if not name:
            continue
        avail = sum_range(ws_av, c1, c2)
        out[name] = out.get(name, 0.0) + avail
    return out
def build_person_hours_json(available_by_person: Dict[str, float], actual_by_person: Dict[str, float]) -> str:
    all_people = sorted(set(available_by_person.keys()) | set(actual_by_person.keys()))
    payload = {}
    for p in all_people:
        payload[p] = {
            "actual": float(actual_by_person.get(p, 0.0)),
            "available": float(available_by_person.get(p, 0.0)),
        }
    return json.dumps(payload, ensure_ascii=False)
def dumps_json(obj: Any) -> str:
    return json.dumps(obj, ensure_ascii=False)
def safe_div(n: float, d: float) -> Optional[float]:
    if d is None or d == 0:
        return None
    return n / d
def scrape_one_workbook(path: str) -> List[Dict[str, Any]]:
    team = team_for_source(path)
    wb = load_workbook(path, data_only=True)
    avail_sheets = find_sheets_by_period(wb, kind="availability")
    prod_sheets = find_sheets_by_period(wb, kind="production")
    periods = sorted(set(avail_sheets.keys()) | set(prod_sheets.keys()))
    rows: List[Dict[str, Any]] = []
    for period in periods:
        err_msgs: List[str] = []
        ws_av = wb[avail_sheets[period]] if period in avail_sheets else None
        ws_prod = wb[prod_sheets[period]] if period in prod_sheets else None
        total_available = None
        person_avail = {}
        if ws_av is not None:
            try:
                total_available = compute_total_available_hours(ws_av)
                person_avail = compute_person_available_hours(ws_av)
            except Exception as e:
                err_msgs.append(f"availability_parse_error: {e!r}")
        else:
            err_msgs.append("missing_availability_sheet")
        completed_hours = None
        actual_hours_by_person = {}
        target_output = None
        actual_output = None
        people = []
        outputs_by_person = {}
        outputs_by_station = {}
        station_hours = {}
        station_hours_by_person = {}
        output_by_station_by_person = {}
        uplh_by_station_by_person = {}
        if ws_prod is not None:
            try:
                completed_hours, actual_hours_by_person = compute_completed_hours(ws_prod)
                target_output, actual_output = compute_target_actual_output(ws_prod)
                people = unique_people_in_wip(ws_prod)
                outputs_by_person = compute_outputs_by_person(ws_prod)
                outputs_by_station = compute_outputs_by_station(ws_prod)
                station_hours, station_hours_by_person = compute_station_hours(ws_prod)
                output_by_station_by_person = compute_output_by_station_by_person(ws_prod)
                uplh_by_station_by_person = compute_uplh_by_station_by_person(
                    output_by_station_by_person, station_hours_by_person
                )
            except Exception as e:
                err_msgs.append(f"production_parse_error: {e!r}")
        else:
            err_msgs.append("missing_production_analysis_sheet")
        target_uplh = safe_div(float(target_output or 0.0), float(completed_hours or 0.0))
        actual_uplh = safe_div(float(actual_output or 0.0), float(completed_hours or 0.0))
        hc_in_wip = len(people) if people else 0
        actual_hc_used = safe_div(float(completed_hours or 0.0), 32.0)
        row: Dict[str, Any] = {
            "team": team,
            "period_date": iso_date(period),
            "source_file": path,
            "Total Available Hours": float(total_available) if total_available is not None else "",
            "Completed Hours": float(completed_hours) if completed_hours is not None else "",
            "Target Output": float(target_output) if target_output is not None else "",
            "Actual Output": float(actual_output) if actual_output is not None else "",
            "Target UPLH": float(target_uplh) if target_uplh is not None else "",
            "Actual UPLH": float(actual_uplh) if actual_uplh is not None else "",
            "UPLH WP1": "",
            "UPLH WP2": "",
            "HC in WIP": hc_in_wip if completed_hours is not None else "",
            "Actual HC Used": float(actual_hc_used) if actual_hc_used is not None else "",
            "People in WIP": dumps_json(people) if people else dumps_json([]) if ws_prod is not None else "",
            "Person Hours": build_person_hours_json(person_avail, actual_hours_by_person) if (ws_av is not None or ws_prod is not None) else "",
            "Outputs by Person": dumps_json(outputs_by_person) if ws_prod is not None else "",
            "Outputs by Cell/Station": dumps_json(outputs_by_station) if ws_prod is not None else "",
            "Cell/Station Hours": dumps_json(station_hours) if ws_prod is not None else "",
            "Hours by Cell/Station - by person": dumps_json(station_hours_by_person) if ws_prod is not None else "",
            "Output by Cell/Station - by person": dumps_json(output_by_station_by_person) if ws_prod is not None else "",
            "UPLH by Cell/Station - by person": dumps_json(uplh_by_station_by_person) if ws_prod is not None else "",
            "Open Complaint Timeliness": "",
            "error": "; ".join(err_msgs) if err_msgs else "",
            "Closures": "",
            "Opened": "",
        }
        rows.append(row)
    return rows
CSV_COLUMNS = [
    "team",
    "period_date",
    "source_file",
    "Total Available Hours",
    "Completed Hours",
    "Target Output",
    "Actual Output",
    "Target UPLH",
    "Actual UPLH",
    "UPLH WP1",
    "UPLH WP2",
    "HC in WIP",
    "Actual HC Used",
    "People in WIP",
    "Person Hours",
    "Outputs by Person",
    "Outputs by Cell/Station",
    "Cell/Station Hours",
    "Hours by Cell/Station - by person",
    "Output by Cell/Station - by person",
    "UPLH by Cell/Station - by person",
    "Open Complaint Timeliness",
    "error",
    "Closures",
    "Opened",
]
def main() -> int:
    default_path = r"C:\Users\wadec8\Medtronic PLC\MCS COS Transformation - VMB Scheduling\Heijunka Current.xlsm"
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="*", help="Excel workbook(s) to scrape (.xlsx/.xlsm).")
    ap.add_argument("--out", default="CRM_WIP.csv", help="Output CSV path (default: CRM_WIP.csv).")
    args = ap.parse_args()
    files = args.files or [default_path]
    all_rows: List[Dict[str, Any]] = []
    for f in files:
        if not os.path.exists(f):
            all_rows.append({
                "team": team_for_source(f),
                "period_date": "",
                "source_file": f,
                "Total Available Hours": "",
                "Completed Hours": "",
                "Target Output": "",
                "Actual Output": "",
                "Target UPLH": "",
                "Actual UPLH": "",
                "UPLH WP1": "",
                "UPLH WP2": "",
                "HC in WIP": "",
                "Actual HC Used": "",
                "People in WIP": "",
                "Person Hours": "",
                "Outputs by Person": "",
                "Outputs by Cell/Station": "",
                "Cell/Station Hours": "",
                "Hours by Cell/Station - by person": "",
                "Output by Cell/Station - by person": "",
                "UPLH by Cell/Station - by person": "",
                "Open Complaint Timeliness": "",
                "error": f"file_not_found: {f}",
                "Closures": "",
                "Opened": "",
            })
            continue
        all_rows.extend(scrape_one_workbook(f))
    with open(args.out, "w", newline="", encoding="utf-8") as fp:
        w = csv.DictWriter(fp, fieldnames=CSV_COLUMNS)
        w.writeheader()
        for r in all_rows:
            w.writerow({k: r.get(k, "") for k in CSV_COLUMNS})
    print(f"Wrote {len(all_rows)} row(s) to {args.out}")
    return 0
if __name__ == "__main__":
    raise SystemExit(main())