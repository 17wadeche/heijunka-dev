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
from zipfile import BadZipFile

# Existing MCS mappings plus new CDS folder mapping.
TEAM_BY_SOURCE: Dict[str, str] = {
    r"C:\Users\wadec8\Medtronic PLC\MCS COS Transformation - VMB Scheduling\Heijunka Current.xlsm": "MCS",
    r"C:\Users\wadec8\Medtronic PLC\Diagnostics MDR - Heijunka and Production Analysis\Archived Heijunka\2026": "CDS",
}
TEAM_BY_BASENAME: Dict[str, str] = {
    "Heijunka Current.xlsm": "MCS",
}
MCS_ROOT_HINT = _norm_mcs = os.path.normpath(r"C:\Users\wadec8\Medtronic PLC\MCS COS Transformation - VMB Scheduling")
CDS_ROOT_HINT = _norm_cds = os.path.normpath(r"C:\Users\wadec8\Medtronic PLC\Diagnostics MDR - Heijunka and Production Analysis\Archived Heijunka\2026")
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
def _norm_path(p: str) -> str:
    return os.path.normpath(p)
def team_for_source(path: str) -> str:
    np = _norm_path(path)
    if np in TEAM_BY_SOURCE:
        return TEAM_BY_SOURCE[np]
    if np.startswith(CDS_ROOT_HINT):
        return "CDS"
    if np.startswith(MCS_ROOT_HINT):
        return "MCS"
    base = os.path.basename(np)
    return TEAM_BY_BASENAME.get(base, "")
def parse_period_date_from_text(text: str, *, default_year: Optional[int] = None) -> Optional[_dt.date]:
    if default_year is None:
        default_year = _dt.date.today().year
    s = (text or "").strip()
    if not s:
        return None
    m = re.search(r"(\d{1,2})\s*[-/ ]\s*([A-Za-z]{3,9})(?:\s*[-/ ]\s*(\d{2,4}))?", s)
    if not m:
        return None
    day = int(m.group(1))
    mon_raw = m.group(2).strip().lower()
    year_raw = m.group(3)
    if mon_raw not in _MONTH_MAP:
        return None
    month = _MONTH_MAP[mon_raw]
    year = default_year
    if year_raw:
        year = int(year_raw)
        if year < 100:
            year += 2000
    try:
        return _dt.date(year, month, day)
    except ValueError:
        return None
def parse_period_date_from_sheetname(sheet_name: str, *, default_year: Optional[int] = None) -> Optional[_dt.date]:
    return parse_period_date_from_text(sheet_name, default_year=default_year)
def parse_period_date_from_filename(path: str, *, default_year: Optional[int] = None) -> Optional[_dt.date]:
    name = os.path.splitext(os.path.basename(path))[0]
    return parse_period_date_from_text(name, default_year=default_year)
def iso_date(d: Optional[_dt.date]) -> str:
    return d.isoformat() if isinstance(d, _dt.date) else ""
def _cell_number(v: Any) -> Optional[float]:
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        vs = v.strip().replace(",", "")
        if vs == "":
            return None
        try:
            return float(vs)
        except ValueError:
            return None
    return None
def is_excluded_person(person: str) -> bool:
    return person.strip().lower() == "do not use"
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
def _norm_team(s: str) -> str:
    return (s or "").strip().upper()
def _norm_period_date(s: str) -> str:
    ss = (s or "").strip()
    if not ss:
        return ""
    try:
        return _dt.date.fromisoformat(ss).isoformat()
    except ValueError:
        return ss
def load_timeliness_lookup(path: str) -> Dict[Tuple[str, str], str]:
    lut: Dict[Tuple[str, str], str] = {}
    if not path or not os.path.exists(path):
        return lut
    with open(path, "r", newline="", encoding="utf-8-sig") as fp:
        r = csv.DictReader(fp)
        for row in r:
            team = _norm_team(row.get("team", ""))
            period = _norm_period_date(row.get("period_date", ""))
            val = (row.get("Open Complaint Timeliness", "") or "").strip()
            if team and period:
                lut[(team, period)] = val
    return lut
def load_closures_lookup(path: str) -> Dict[Tuple[str, str], Tuple[str, str]]:
    lut: Dict[Tuple[str, str], Tuple[str, str]] = {}
    if not path or not os.path.exists(path):
        return lut
    with open(path, "r", newline="", encoding="utf-8-sig") as fp:
        r = csv.DictReader(fp)
        for row in r:
            team = _norm_team(row.get("team", ""))
            period = _norm_period_date(row.get("period_date", ""))
            closures = (row.get("Closures", "") or "").strip()
            opened = (row.get("Opened", "") or "").strip()
            if team and period:
                lut[(team, period)] = (closures, opened)
    return lut
def enrich_rows_with_metrics(
    rows: List[Dict[str, Any]],
    timeliness_lut: Dict[Tuple[str, str], str],
    closures_lut: Dict[Tuple[str, str], Tuple[str, str]],
) -> None:
    for r in rows:
        key = (_norm_team(r.get("team", "")), _norm_period_date(r.get("period_date", "")))
        if key in timeliness_lut:
            r["Open Complaint Timeliness"] = timeliness_lut[key]
        if key in closures_lut:
            c, o = closures_lut[key]
            r["Closures"] = c
            r["Opened"] = o
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
    blocks = [("B", "F"), ("I", "M"), ("P", "T")]
    rows = [5, 15, 25, 35, 45, 55]
    total = 0.0
    for r in rows:
        for c1, c2 in blocks:
            total += sum_range(ws_av, f"{c1}{r}", f"{c2}{r}")
    return total
def iter_prod_rows_mcs(ws_prod: Worksheet, start_row: int = 7) -> Iterable[Tuple[int, str, str, Optional[float], Optional[float]]]:
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
def compute_completed_hours_mcs(ws_prod: Worksheet) -> Tuple[float, Dict[str, float]]:
    total = 0.0
    by_person: Dict[str, float] = {}
    for r, person, cell_station, target, output in iter_prod_rows_mcs(ws_prod, start_row=7):
        if output is None:
            continue
        h = 4.0 if (cell_station == "Promoted PE - Initial MDR") else 1.0
        total += h
        if person and not is_excluded_person(person):
            by_person[person] = by_person.get(person, 0.0) + h
    return total, by_person
def compute_target_actual_output_mcs(ws_prod: Worksheet) -> Tuple[float, float]:
    targ = 0.0
    act = 0.0
    for r, person, cell_station, target, output in iter_prod_rows_mcs(ws_prod, start_row=7):
        if output is None:
            continue
        act += output
        if target is not None:
            targ += target
    return targ, act
def unique_people_in_wip_mcs(ws_prod: Worksheet) -> List[str]:
    seen = set()
    for r, person, cell_station, target, output in iter_prod_rows_mcs(ws_prod, start_row=7):
        if output is None or not person or is_excluded_person(person):
            continue
        seen.add(person)
    return sorted(seen)
def compute_outputs_by_person_mcs(ws_prod: Worksheet) -> Dict[str, Dict[str, float]]:
    out: Dict[str, Dict[str, float]] = {}
    for r, person, cell_station, target, output in iter_prod_rows_mcs(ws_prod, start_row=7):
        if output is None or not person or is_excluded_person(person):
            continue
        out.setdefault(person, {"output": 0.0, "target": 0.0})
        out[person]["output"] += output
        if target is not None:
            out[person]["target"] += target
    return out
def compute_outputs_by_station_mcs(ws_prod: Worksheet) -> Dict[str, Dict[str, float]]:
    out: Dict[str, Dict[str, float]] = {}
    for r, person, cell_station, target, output in iter_prod_rows_mcs(ws_prod, start_row=7):
        if output is None or not cell_station:
            continue
        out.setdefault(cell_station, {"output": 0.0, "target": 0.0})
        out[cell_station]["output"] += output
        if target is not None:
            out[cell_station]["target"] += target
    return out
def compute_station_hours_mcs(ws_prod: Worksheet) -> Tuple[Dict[str, float], Dict[str, Dict[str, float]]]:
    station_hours: Dict[str, float] = {}
    station_hours_by_person: Dict[str, Dict[str, float]] = {}
    for r, person, cell_station, target, output in iter_prod_rows_mcs(ws_prod, start_row=7):
        if not person or not cell_station:
            continue
        if person.strip().lower() in {"do not use", "team member(s)"}:
            continue
        if cell_station.strip().lower() == "cell/station":
            continue
        h = 2.0 if cell_station == "Promoted PE - Initial MDR" else 1.0
        station_hours[cell_station] = station_hours.get(cell_station, 0.0) + h
        station_hours_by_person.setdefault(cell_station, {})
        station_hours_by_person[cell_station][person] = station_hours_by_person[cell_station].get(person, 0.0) + h
    return station_hours, station_hours_by_person
def compute_output_by_station_by_person_mcs(ws_prod: Worksheet) -> Dict[str, Dict[str, float]]:
    out: Dict[str, Dict[str, float]] = {}
    for r, person, cell_station, target, output in iter_prod_rows_mcs(ws_prod, start_row=7):
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
        if not name or is_excluded_person(name):
            continue
        avail = sum_range(ws_av, c1, c2)
        out[name] = out.get(name, 0.0) + avail
    return out
def find_cds_production_sheet(wb) -> Optional[str]:
    for name in wb.sheetnames:
        if name.strip().lower() == "#12 production analysis":
            return name
    for name in wb.sheetnames:
        if _PROD_PAT.search(name):
            return name
    return None
def iter_prod_rows_cds(ws_prod: Worksheet, start_row: int = 9, end_row: int = 137) -> Iterable[Tuple[int, str, str, Optional[float], Optional[float], Optional[float]]]:
    maxr = min(ws_prod.max_row, end_row)
    for r in range(start_row, maxr + 1):
        person = ws_prod[f"C{r}"].value
        cell_station = ws_prod[f"D{r}"].value
        target = _cell_number(ws_prod[f"F{r}"].value)
        minutes = _cell_number(ws_prod[f"G{r}"].value)
        output = _cell_number(ws_prod[f"I{r}"].value)
        p = str(person).strip() if person is not None else ""
        cs = str(cell_station).strip() if cell_station is not None else ""
        if p == "" and cs == "" and target is None and minutes is None and output is None:
            continue
        yield (r, p, cs, target, minutes, output)
def compute_completed_hours_cds(ws_prod: Worksheet) -> Tuple[float, Dict[str, float]]:
    total_minutes = 0.0
    by_person_minutes: Dict[str, float] = {}
    for _, person, _, _, minutes, _ in iter_prod_rows_cds(ws_prod):
        if minutes is None:
            continue
        total_minutes += minutes
        if person and not is_excluded_person(person):
            by_person_minutes[person] = by_person_minutes.get(person, 0.0) + minutes
    return total_minutes / 60.0, {k: v / 60.0 for k, v in by_person_minutes.items()}
def compute_target_actual_output_cds(ws_prod: Worksheet) -> Tuple[float, float]:
    targ = 0.0
    act = 0.0
    for _, _, _, target, _, output in iter_prod_rows_cds(ws_prod):
        if target is not None:
            targ += target
        if output is not None:
            act += output
    return targ, act
def unique_people_in_wip_cds(ws_prod: Worksheet) -> List[str]:
    seen = set()
    for _, person, _, _, _, _ in iter_prod_rows_cds(ws_prod):
        if not person or is_excluded_person(person):
            continue
        seen.add(person)
    return sorted(seen)
def compute_outputs_by_person_cds(ws_prod: Worksheet) -> Dict[str, Dict[str, float]]:
    out: Dict[str, Dict[str, float]] = {}
    for _, person, _, target, _, output in iter_prod_rows_cds(ws_prod):
        if not person or is_excluded_person(person):
            continue
        out.setdefault(person, {"output": 0.0, "target": 0.0})
        if output is not None:
            out[person]["output"] += output
        if target is not None:
            out[person]["target"] += target
    return out
def compute_outputs_by_station_cds(ws_prod: Worksheet) -> Dict[str, Dict[str, float]]:
    out: Dict[str, Dict[str, float]] = {}
    for _, _, cell_station, target, _, output in iter_prod_rows_cds(ws_prod):
        if not cell_station:
            continue
        out.setdefault(cell_station, {"output": 0.0, "target": 0.0})
        if output is not None:
            out[cell_station]["output"] += output
        if target is not None:
            out[cell_station]["target"] += target
    return out
def compute_station_hours_cds(ws_prod: Worksheet) -> Tuple[Dict[str, float], Dict[str, Dict[str, float]]]:
    station_hours: Dict[str, float] = {}
    station_hours_by_person: Dict[str, Dict[str, float]] = {}
    for _, person, cell_station, _, minutes, _ in iter_prod_rows_cds(ws_prod):
        if not person or not cell_station or minutes is None or is_excluded_person(person):
            continue
        hours = minutes / 60.0
        station_hours[cell_station] = station_hours.get(cell_station, 0.0) + hours
        station_hours_by_person.setdefault(cell_station, {})
        station_hours_by_person[cell_station][person] = station_hours_by_person[cell_station].get(person, 0.0) + hours
    return station_hours, station_hours_by_person
def compute_output_by_station_by_person_cds(ws_prod: Worksheet) -> Dict[str, Dict[str, float]]:
    out: Dict[str, Dict[str, float]] = {}
    for _, person, cell_station, _, _, output in iter_prod_rows_cds(ws_prod):
        if not person or not cell_station or output is None:
            continue
        out.setdefault(cell_station, {})
        out[cell_station][person] = out[cell_station].get(person, 0.0) + output
    return out
def build_person_hours_json(available_by_person: Dict[str, float], actual_by_person: Dict[str, float]) -> str:
    all_people = sorted(set(available_by_person.keys()) | set(actual_by_person.keys()))
    payload = {}
    for p in all_people:
        if is_excluded_person(p):
            continue
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
def scrape_one_workbook_mcs(path: str) -> List[Dict[str, Any]]:
    team = team_for_source(path)
    wb = load_workbook(path, data_only=True)
    avail_sheets = find_sheets_by_period(wb, kind="availability")
    prod_sheets = find_sheets_by_period(wb, kind="production")
    periods = sorted(set(avail_sheets.keys()) | set(prod_sheets.keys()), reverse=True)
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
                completed_hours, actual_hours_by_person = compute_completed_hours_mcs(ws_prod)
                target_output, actual_output = compute_target_actual_output_mcs(ws_prod)
                people = unique_people_in_wip_mcs(ws_prod)
                outputs_by_person = compute_outputs_by_person_mcs(ws_prod)
                outputs_by_station = compute_outputs_by_station_mcs(ws_prod)
                station_hours, station_hours_by_person = compute_station_hours_mcs(ws_prod)
                output_by_station_by_person = compute_output_by_station_by_person_mcs(ws_prod)
                uplh_by_station_by_person = compute_uplh_by_station_by_person(output_by_station_by_person, station_hours_by_person)
            except Exception as e:
                err_msgs.append(f"production_parse_error: {e!r}")
        else:
            err_msgs.append("missing_production_analysis_sheet")
        target_uplh = safe_div(float(target_output or 0.0), float(completed_hours or 0.0))
        actual_uplh = safe_div(float(actual_output or 0.0), float(completed_hours or 0.0))
        hc_in_wip = len(people) if people else 0
        actual_hc_used = safe_div(float(completed_hours or 0.0), 32.0)
        rows.append({
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
        })
    return rows
def scrape_one_workbook_cds(path: str) -> List[Dict[str, Any]]:
    team = team_for_source(path)
    wb = load_workbook(path, data_only=True)
    err_msgs: List[str] = []
    prod_sheet_name = find_cds_production_sheet(wb)
    ws_prod = wb[prod_sheet_name] if prod_sheet_name else None
    if ws_prod is None:
        err_msgs.append("missing_#12_production_analysis_sheet")
    period = None
    if prod_sheet_name:
        period = parse_period_date_from_sheetname(prod_sheet_name, default_year=2026)
    if period is None:
        period = parse_period_date_from_filename(path, default_year=2026)
    if period is None:
        err_msgs.append("missing_period_date_from_sheet_or_filename")
    completed_hours = None
    actual_hours_by_person: Dict[str, float] = {}
    target_output = None
    actual_output = None
    people: List[str] = []
    outputs_by_person: Dict[str, Dict[str, float]] = {}
    outputs_by_station: Dict[str, Dict[str, float]] = {}
    station_hours: Dict[str, float] = {}
    station_hours_by_person: Dict[str, Dict[str, float]] = {}
    output_by_station_by_person: Dict[str, Dict[str, float]] = {}
    uplh_by_station_by_person: Dict[str, Dict[str, float]] = {}
    if ws_prod is not None:
        try:
            completed_hours, actual_hours_by_person = compute_completed_hours_cds(ws_prod)
            target_output, actual_output = compute_target_actual_output_cds(ws_prod)
            people = unique_people_in_wip_cds(ws_prod)
            outputs_by_person = compute_outputs_by_person_cds(ws_prod)
            outputs_by_station = compute_outputs_by_station_cds(ws_prod)
            station_hours, station_hours_by_person = compute_station_hours_cds(ws_prod)
            output_by_station_by_person = compute_output_by_station_by_person_cds(ws_prod)
            uplh_by_station_by_person = compute_uplh_by_station_by_person(output_by_station_by_person, station_hours_by_person)
        except Exception as e:
            err_msgs.append(f"production_parse_error: {e!r}")
    target_uplh = safe_div(float(target_output or 0.0), float(completed_hours or 0.0))
    actual_uplh = safe_div(float(actual_output or 0.0), float(completed_hours or 0.0))
    hc_in_wip = len(people) if people else 0
    actual_hc_used = safe_div(float(completed_hours or 0.0), 32.5)
    person_avail = {p: 32.0 for p in people}
    return [{
        "team": team,
        "period_date": iso_date(period),
        "source_file": path,
        "Total Available Hours": 160.0,
        "Completed Hours": float(completed_hours) if completed_hours is not None else "",
        "Target Output": float(target_output) if target_output is not None else "",
        "Actual Output": float(actual_output) if actual_output is not None else "",
        "Target UPLH": float(target_uplh) if target_uplh is not None else "",
        "Actual UPLH": float(actual_uplh) if actual_uplh is not None else "",
        "UPLH WP1": "",
        "UPLH WP2": "",
        "HC in WIP": hc_in_wip if ws_prod is not None else "",
        "Actual HC Used": float(actual_hc_used) if actual_hc_used is not None else "",
        "People in WIP": dumps_json(people) if ws_prod is not None else "",
        "Person Hours": build_person_hours_json(person_avail, actual_hours_by_person) if ws_prod is not None else "",
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
    }]
def scrape_one_workbook(path: str) -> List[Dict[str, Any]]:
    team = team_for_source(path)
    try:
        if team == "CDS":
            return scrape_one_workbook_cds(path)
        return scrape_one_workbook_mcs(path)
    except (BadZipFile, OSError, ValueError) as e:
        row = blank_row_for_missing_file(path)
        row["team"] = team
        row["error"] = f"workbook_open_error: {type(e).__name__}: {e}"
        return [row]
CSV_COLUMNS = [
    "team", "period_date", "source_file", "Total Available Hours", "Completed Hours",
    "Target Output", "Actual Output", "Target UPLH", "Actual UPLH", "UPLH WP1", "UPLH WP2",
    "HC in WIP", "Actual HC Used", "People in WIP", "Person Hours", "Outputs by Person",
    "Outputs by Cell/Station", "Cell/Station Hours", "Hours by Cell/Station - by person",
    "Output by Cell/Station - by person", "UPLH by Cell/Station - by person",
    "Open Complaint Timeliness", "error", "Closures", "Opened",
]
def iter_input_files(paths: List[str]) -> Iterable[str]:
    for p in paths:
        if os.path.isdir(p):
            for name in sorted(os.listdir(p)):
                lower = name.lower()
                if name.startswith("~$"):
                    continue
                if lower.endswith((".xlsx", ".xlsm", ".xls")):
                    yield os.path.join(p, name)
        else:
            yield p
def blank_row_for_missing_file(f: str) -> Dict[str, Any]:
    return {
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
    }
def main() -> int:
    default_paths = [
        r"C:\Users\wadec8\Medtronic PLC\MCS COS Transformation - VMB Scheduling\Heijunka Current.xlsm",
        r"C:\Users\wadec8\Medtronic PLC\Diagnostics MDR - Heijunka and Production Analysis\Archived Heijunka\2026",
    ]
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs="*", help="Excel workbook(s) and/or folders to scrape (.xlsx/.xlsm).")
    ap.add_argument("--out", default="CRM_WIP.csv", help="Output CSV path (default: CRM_WIP.csv).")
    ap.add_argument("--timeliness-csv", default="timeliness.csv", help="Path to timeliness.csv")
    ap.add_argument("--closures-csv", default="closures.csv", help="Path to closures.csv")
    args = ap.parse_args()
    inputs = args.files or default_paths
    all_rows: List[Dict[str, Any]] = []
    for f in iter_input_files(inputs):
        if not os.path.exists(f):
            all_rows.append(blank_row_for_missing_file(f))
            continue
        all_rows.extend(scrape_one_workbook(f))
    timeliness_lut = load_timeliness_lookup(args.timeliness_csv)
    closures_lut = load_closures_lookup(args.closures_csv)
    enrich_rows_with_metrics(all_rows, timeliness_lut, closures_lut)
    with open(args.out, "w", newline="", encoding="utf-8") as fp:
        w = csv.DictWriter(fp, fieldnames=CSV_COLUMNS)
        w.writeheader()
        for r in all_rows:
            w.writerow({k: r.get(k, "") for k in CSV_COLUMNS})
    print(f"Wrote {len(all_rows)} row(s) to {args.out}")
    return 0
if __name__ == "__main__":
    raise SystemExit(main())