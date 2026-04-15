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
MCS_DEFAULT_PATH = r"C:\Users\wadec8\Medtronic PLC\MCS COS Transformation - VMB Scheduling\Heijunka Current.xlsm"
DS_DEFAULT_DIR = r"C:\Users\wadec8\Medtronic PLC\Defibrillation Solutions - Schedule and PAB"
CPT_DEFAULT_DIR = r"C:\Users\wadec8\Medtronic PLC\Cardiac Pacing Therapies CQXM - Heijunka & PAB"
CDS_DEFAULT_DIR = r"C:\Users\wadec8\Medtronic PLC\Diagnostics MDR - Heijunka and Production Analysis"
CDS_ARCHIVE_PAB_DIR = r"C:\Users\wadec8\Medtronic PLC\Diagnostics MDR - Heijunka and Production Analysis\Archived PAB"
NI_DEFAULT_DIR = r"C:\Users\wadec8\Medtronic PLC\Tier1 PXM - Non Implantables - Heijunka"
NI_ARCHIVE_APRIL_2026_DIR = r"C:\Users\wadec8\Medtronic PLC\Tier1 PXM - Non Implantables - Heijunka\Archive\April 2026 - PAB"
TEAM_BY_SOURCE: Dict[str, str] = {
    os.path.normpath(MCS_DEFAULT_PATH): "MCS",
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
DS_PAB_SHEET = "#2 PAB"
DS_WIP_PLAN_SHEET = "# 1 WIP plan"
DS_PERF_WIP_SHEET = "#5 Performance WIP Time"
DS_NON_WIP_TYPES = {"essential non-wip", "non-wip"}
DS_PEOPLE_COUNT = 41
CPT_PAB_SHEET = "#3 PAB"
CPT_WIP_PLAN_SHEET = "# 1 WIP plan"
CPT_PERF_WIP_SHEET = "#6 Performance WIP Time"
CPT_NON_WIP_TYPES = {"essential non-wip", "non-wip"}
CPT_PEOPLE_COUNT = 43
CDS_PAB_SHEET = "#2 PAB"
CDS_WIP_PLAN_SHEET = "# 1 WIP plan"
CDS_PERF_METRICS_SHEET = "#4 Performance Metrics"
CDS_PERF_WIP_SHEET = "#5 Performance WIP Time"
CDS_NON_WIP_TYPES = {"essential non-wip", "non-wip"}
CDS_PEOPLE_COUNT = 6
NI_PAB_SHEET = "#2 PAB"
NI_WIP_PLAN_SHEET = "# 1 WIP plan"
NI_PERF_METRICS_SHEET = "#4 Performance Metrics"
NI_PERF_WIP_SHEET = "#5 Performance WIP Time"
NI_NON_WIP_TYPES = {"essential non-wip", "non-wip"}
NI_PEOPLE_COUNT = 8
def _norm_path(p: str) -> str:
    return os.path.normpath(p)
def team_for_source(path: str) -> str:
    np = _norm_path(path)
    if np in TEAM_BY_SOURCE:
        return TEAM_BY_SOURCE[np]
    ds_root = _norm_path(DS_DEFAULT_DIR)
    if np.startswith(ds_root + os.sep) or np == ds_root:
        return "DS"
    cpt_root = _norm_path(CPT_DEFAULT_DIR)
    if np.startswith(cpt_root + os.sep) or np == cpt_root:
        return "CPT"
    cds_root = _norm_path(CDS_DEFAULT_DIR)
    if np.startswith(cds_root + os.sep) or np == cds_root:
        return "CDS"
    ni_root = _norm_path(NI_DEFAULT_DIR)
    if np.startswith(ni_root + os.sep) or np == ni_root:
        return "NI"
    base = os.path.basename(np)
    if base in TEAM_BY_BASENAME:
        return TEAM_BY_BASENAME[base]
    np_lower = np.lower()
    if "defibrillation solutions" in np_lower:
        return "DS"
    if "cardiac pacing therapies" in np_lower:
        return "CPT"
    if "diagnostics mdr" in np_lower:
        return "CDS"
    if "non implantables" in np_lower:
        return "NI"
    return ""
def _norm_text(x: str) -> str:
    return re.sub(r"\s+", " ", (x or "").strip()).lower()
def _collapse_ws(x: str) -> str:
    return re.sub(r"\s+", " ", (x or "").strip())
def normalize_person_name(name: str) -> str:
    s = _collapse_ws(name)
    if not s:
        return ""
    return s.title()
def normalize_person_key(name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", _norm_text(name))
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
def parse_period_date_from_filename(filename: str) -> Optional[_dt.date]:
    base = os.path.basename(filename)
    stem, _ = os.path.splitext(base)
    patterns = [
        r"(?<!\d)(\d{1,2})\s*[-_ ]*\s*([A-Za-z]{3,9})(?:\s*[-_ ]*\s*(\d{2,4}))?(?!\d)",
        r"(?<!\d)(\d{4})[\s\-_]+(\d{1,2})[\s\-_]+(\d{1,2})(?!\d)",
        r"(?<!\d)(\d{1,2})[\s\-_\/]+(\d{1,2})[\s\-_\/]+(\d{4})(?!\d)",
    ]
    for i, pat in enumerate(patterns):
        m = re.search(pat, stem)
        if not m:
            continue
        try:
            if i == 0:
                day = int(m.group(1))
                mon_raw = m.group(2).lower()
                year_raw = m.group(3)
                year = int(year_raw) if year_raw else 2026
                if year < 100:
                    year += 2000
                if mon_raw not in _MONTH_MAP:
                    continue
                return _dt.date(year, _MONTH_MAP[mon_raw], day)
            if i == 1:
                year = int(m.group(1))
                month = int(m.group(2))
                day = int(m.group(3))
                return _dt.date(year, month, day)
            if i == 2:
                month = int(m.group(1))
                day = int(m.group(2))
                year = int(m.group(3))
                return _dt.date(year, month, day)
        except ValueError:
            continue
    return None
def parse_period_date_from_workbook_sheetnames(wb, *, default_year: Optional[int] = None) -> Optional[_dt.date]:
    for name in wb.sheetnames:
        d = parse_period_date_from_sheetname(name, default_year=default_year)
        if d is not None:
            return d
    return None
def parse_period_date_from_perf_metrics_cell(ws: Worksheet, cell_ref: str = "B3") -> Optional[_dt.date]:
    value = ws[cell_ref].value
    if value is None or value == "":
        return None
    if isinstance(value, _dt.datetime):
        return value.date()
    if isinstance(value, _dt.date):
        return value
    if isinstance(value, (int, float)):
        try:
            return _dt.datetime.fromordinal(_dt.datetime(1899, 12, 30).toordinal() + int(value)).date()
        except Exception:
            return None
    s = str(value).strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%d-%b-%Y", "%d-%b-%y", "%d %b %Y", "%d %B %Y"):
        try:
            return _dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None
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
        if vs.endswith("%"):
            try:
                return float(vs[:-1]) / 100.0
            except ValueError:
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
    for _, person, station, target, output in iter_prod_rows(ws_prod, start_row=7):
        if output is None:
            continue
        if is_excluded_person(person):
            continue
        norm_person = normalize_person_name(person)
        if not norm_person:
            continue
        seen.add(norm_person)
    return sorted(seen)
NON_WIP_SPECS = [
    ("A13", ("B18", "F21"), "A", ("B17", "F17")),
    ("A23", ("B28", "F31"), "A", ("B27", "F27")),
    ("H3",  ("I8", "M11"), "H", ("I7", "M7")),
    ("H23", ("I28", "M31"), "H", ("I27", "M27")),
    ("H33", ("I38", "M41"), "H", ("I37", "M37")),
    ("O3",  ("P8", "T11"), "O", ("P7", "T7")),
    ("O13", ("P18", "T21"), "O", ("P17", "T17")),
    ("O23", ("P28", "T31"), "O", ("P27", "T27")),
    ("O43", ("P48", "T51"), "O", ("P47", "T47")),
    ("O53", ("P58", "T61"), "O", ("P57", "T57")),
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
        raw_name = read_merged_value(ws_av, name_cell)
        if not raw_name or is_excluded_person(raw_name):
            continue
        name = normalize_person_name(raw_name)
        if not name:
            continue
        out[name] = out.get(name, 0.0) + sum_range(ws_av, c1, c2)
    return out
def compute_non_wip_activities(ws_av: Worksheet) -> List[Dict[str, Any]]:
    agg: Dict[Tuple[str, str], float] = {}
    for name_cell, (c1, c2), label_col, (ooo_c1, ooo_c2) in NON_WIP_SPECS:
        raw_name = read_merged_value(ws_av, name_cell)
        if not raw_name or is_excluded_person(raw_name):
            continue
        name = normalize_person_name(raw_name)
        if not name:
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
        ooo_hours = sum_range(ws_av, ooo_c1, ooo_c2)
        if ooo_hours:
            key = (name, "OOO")
            agg[key] = agg.get(key, 0.0) + float(ooo_hours)
    return [
        {"name": n, "activity": a, "hours": float(h)}
        for (n, a), h in sorted(agg.items(), key=lambda x: (x[0][0].lower(), x[0][1].lower()))
    ]
def compute_ooo_by_person(ws_av: Worksheet) -> Dict[str, float]:
    out: Dict[str, float] = {}
    for name_cell, _, _, (c1, c2) in NON_WIP_SPECS:
        raw_name = read_merged_value(ws_av, name_cell)
        if not raw_name or is_excluded_person(raw_name):
            continue
        name = normalize_person_name(raw_name)
        if not name:
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
def parse_people_in_wip_value(value: Any) -> List[str]:
    if value is None:
        return []
    if isinstance(value, list):
        raw_items = value
    else:
        s = str(value).strip()
        if not s:
            return []
        try:
            parsed = json.loads(s)
            raw_items = parsed if isinstance(parsed, list) else [parsed]
        except Exception:
            raw_items = re.split(r"[;,|\n]+", s)
    out: List[str] = []
    seen = set()
    for item in raw_items:
        name = normalize_person_name(str(item))
        if not name:
            continue
        key = normalize_person_key(name)
        if key in seen:
            continue
        seen.add(key)
        out.append(name)
    return out
def compute_ds_ooo_by_person(ws_wip_plan: Worksheet) -> Dict[str, float]:
    out: Dict[str, float] = {}
    for r in range(3, ws_wip_plan.max_row + 1):
        person = normalize_person_name(str(ws_wip_plan[f"DL{r}"].value or ""))
        if not person or is_excluded_person(person):
            continue
        hours = _cell_number(ws_wip_plan[f"EF{r}"].value)
        if hours is None or hours == 0:
            continue
        out[person] = out.get(person, 0.0) + float(hours)
    return out
def compute_ds_non_wip_activities(ws_pab: Worksheet, ws_wip_plan: Worksheet) -> List[Dict[str, Any]]:
    agg: Dict[Tuple[str, str], float] = {}
    for person, activity, hours in iter_ds_non_wip_rows(ws_pab):
        key = (person, activity)
        agg[key] = agg.get(key, 0.0) + float(hours)
    for person, hours in compute_ds_ooo_by_person(ws_wip_plan).items():
        key = (person, "OOO")
        agg[key] = agg.get(key, 0.0) + float(hours)
    return [
        {"name": person, "activity": activity, "hours": float(hours)}
        for (person, activity), hours in sorted(agg.items(), key=lambda x: (x[0][0].lower(), x[0][1].lower()))
        if hours != 0
    ]
def compute_cpt_ooo_by_person(ws_wip_plan: Worksheet) -> Dict[str, float]:
    out: Dict[str, float] = {}
    for r in range(3, ws_wip_plan.max_row + 1):
        person = normalize_person_name(str(ws_wip_plan[f"DB{r}"].value or ""))
        if not person or is_excluded_person(person):
            continue
        hours = _cell_number(ws_wip_plan[f"DT{r}"].value)
        if hours is None or hours == 0:
            continue
        out[person] = out.get(person, 0.0) + float(hours)
    return out
def compute_cpt_non_wip_activities(ws_pab: Worksheet, ws_wip_plan: Worksheet) -> List[Dict[str, Any]]:
    agg: Dict[Tuple[str, str], float] = {}
    for person, activity, hours in iter_cpt_non_wip_rows(ws_pab):
        key = (person, activity)
        agg[key] = agg.get(key, 0.0) + float(hours)
    for person, hours in compute_cpt_ooo_by_person(ws_wip_plan).items():
        key = (person, "OOO")
        agg[key] = agg.get(key, 0.0) + float(hours)
    return [
        {"name": person, "activity": activity, "hours": float(hours)}
        for (person, activity), hours in sorted(agg.items(), key=lambda x: (x[0][0].lower(), x[0][1].lower()))
        if hours != 0
    ]
def load_people_in_wip_from_crm_wip(crm_wip_csv: str) -> Dict[Tuple[str, str], List[str]]:
    out: Dict[Tuple[str, str], List[str]] = {}
    if not os.path.exists(crm_wip_csv):
        return out
    with open(crm_wip_csv, "r", encoding="utf-8-sig", newline="") as fp:
        r = csv.DictReader(fp)
        for row in r:
            team = (row.get("team") or "").strip()
            period = (row.get("period_date") or "").strip()
            if not team or not period:
                continue
            people = parse_people_in_wip_value(row.get("People in WIP"))
            if not people:
                continue
            key = (team, period)
            existing = out.get(key, [])
            seen = {normalize_person_key(x) for x in existing}
            for person in people:
                pkey = normalize_person_key(person)
                if pkey not in seen:
                    existing.append(person)
                    seen.add(pkey)
            out[key] = existing
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
def _get_required_sheet(wb, sheet_name: str) -> Worksheet:
    if sheet_name not in wb.sheetnames:
        raise KeyError(f"Missing required sheet: {sheet_name}")
    return wb[sheet_name]
def iter_ds_non_wip_rows(ws_pab: Worksheet, start_row: int = 2) -> Iterable[Tuple[str, str, float]]:
    for r in range(start_row, ws_pab.max_row + 1):
        category = _norm_text(str(ws_pab[f"D{r}"].value or ""))
        if category not in DS_NON_WIP_TYPES:
            continue
        person = normalize_person_name(str(ws_pab[f"C{r}"].value or ""))
        activity = _collapse_ws(str(ws_pab[f"E{r}"].value or ""))
        hours = _cell_number(ws_pab[f"I{r}"].value)
        if not person or hours is None:
            continue
        yield person, activity, float(hours)
def compute_ds_total_non_wip_hours(ws_pab: Worksheet) -> float:
    return float(sum(hours for _, _, hours in iter_ds_non_wip_rows(ws_pab)))
def compute_ds_non_wip_by_person(ws_pab: Worksheet) -> Dict[str, float]:
    out: Dict[str, float] = {}
    for person, _, hours in iter_ds_non_wip_rows(ws_pab):
        out[person] = out.get(person, 0.0) + float(hours)
    return {person: float(total) for person, total in out.items() if total != 0}
def compute_ds_non_wip_activities(ws_pab: Worksheet, ws_wip_plan: Worksheet) -> List[Dict[str, Any]]:
    agg: Dict[Tuple[str, str], float] = {}
    for person, activity, hours in iter_ds_non_wip_rows(ws_pab):
        key = (person, activity)
        agg[key] = agg.get(key, 0.0) + float(hours)
    for person, hours in compute_ds_ooo_by_person(ws_wip_plan).items():
        key = (person, "OOO")
        agg[key] = agg.get(key, 0.0) + float(hours)
    return [
        {"name": person, "activity": activity, "hours": float(hours)}
        for (person, activity), hours in sorted(
            agg.items(), key=lambda x: (x[0][0].lower(), x[0][1].lower())
        )
        if hours != 0
    ]
def compute_ds_wip_workers_ooo_hours(ws_wip_plan: Worksheet, wip_workers: List[str]) -> float:
    worker_keys = {normalize_person_key(x) for x in wip_workers if x}
    if not worker_keys:
        return 0.0
    total = 0.0
    for r in range(1, ws_wip_plan.max_row + 1):
        name = str(ws_wip_plan[f"DL{r}"].value or "").strip()
        if not name:
            continue
        if normalize_person_key(name) not in worker_keys:
            continue
        val = _cell_number(ws_wip_plan[f"EF{r}"].value)
        if val is not None:
            total += float(val)
    return float(total)
def iter_cpt_non_wip_rows(ws_pab: Worksheet, start_row: int = 2) -> Iterable[Tuple[str, str, float]]:
    for r in range(start_row, ws_pab.max_row + 1):
        category = _norm_text(str(ws_pab[f"D{r}"].value or ""))
        if category not in CPT_NON_WIP_TYPES:
            continue
        person = normalize_person_name(str(ws_pab[f"C{r}"].value or ""))
        activity = _collapse_ws(str(ws_pab[f"E{r}"].value or ""))
        hours = _cell_number(ws_pab[f"I{r}"].value)
        if not person or hours is None:
            continue
        yield person, activity, float(hours)
def compute_cpt_total_non_wip_hours(ws_pab: Worksheet) -> float:
    return float(sum(hours for _, _, hours in iter_cpt_non_wip_rows(ws_pab)))
def compute_cpt_non_wip_by_person(ws_pab: Worksheet) -> Dict[str, float]:
    out: Dict[str, float] = {}
    for person, _, hours in iter_cpt_non_wip_rows(ws_pab):
        out[person] = out.get(person, 0.0) + float(hours)
    return {person: float(total) for person, total in out.items() if total != 0}
def compute_cpt_non_wip_activities(ws_pab: Worksheet, ws_wip_plan: Worksheet) -> List[Dict[str, Any]]:
    agg: Dict[Tuple[str, str], float] = {}
    for person, activity, hours in iter_cpt_non_wip_rows(ws_pab):
        key = (person, activity)
        agg[key] = agg.get(key, 0.0) + float(hours)
    for person, hours in compute_cpt_ooo_by_person(ws_wip_plan).items():
        key = (person, "OOO")
        agg[key] = agg.get(key, 0.0) + float(hours)
    return [
        {"name": person, "activity": activity, "hours": float(hours)}
        for (person, activity), hours in sorted(
            agg.items(), key=lambda x: (x[0][0].lower(), x[0][1].lower())
        )
        if hours != 0
    ]
def compute_cpt_wip_workers_ooo_hours(ws_wip_plan: Worksheet, wip_workers: List[str]) -> float:
    worker_keys = {normalize_person_key(x) for x in wip_workers if x}
    if not worker_keys:
        return 0.0
    total = 0.0
    for r in range(1, ws_wip_plan.max_row + 1):
        name = str(ws_wip_plan[f"DB{r}"].value or "").strip()
        if not name:
            continue
        if normalize_person_key(name) not in worker_keys:
            continue
        val = _cell_number(ws_wip_plan[f"DT{r}"].value)
        if val is not None:
            total += float(val)
    return float(total)
def iter_team_non_wip_rows(ws_pab: Worksheet, non_wip_types: set[str], start_row: int = 2) -> Iterable[Tuple[str, str, float]]:
    for r in range(start_row, ws_pab.max_row + 1):
        category = _norm_text(str(ws_pab[f"D{r}"].value or ""))
        if category not in non_wip_types:
            continue
        person = normalize_person_name(str(ws_pab[f"C{r}"].value or ""))
        activity = _collapse_ws(str(ws_pab[f"E{r}"].value or ""))
        hours = _cell_number(ws_pab[f"I{r}"].value)
        if not person or hours is None:
            continue
        yield person, activity, float(hours)
def compute_ooo_by_person_from_wip_plan(ws_wip_plan: Worksheet, name_col: str, hours_col: str, start_row: int = 3) -> Dict[str, float]:
    out: Dict[str, float] = {}
    for r in range(start_row, ws_wip_plan.max_row + 1):
        person = normalize_person_name(str(ws_wip_plan[f"{name_col}{r}"].value or ""))
        if not person or is_excluded_person(person):
            continue
        hours = _cell_number(ws_wip_plan[f"{hours_col}{r}"].value)
        if hours is None or hours == 0:
            continue
        out[person] = out.get(person, 0.0) + float(hours)
    return out
def compute_team_total_non_wip_hours(ws_pab: Worksheet, non_wip_types: set[str]) -> float:
    return float(sum(hours for _, _, hours in iter_team_non_wip_rows(ws_pab, non_wip_types)))
def compute_team_non_wip_by_person(ws_pab: Worksheet, non_wip_types: set[str]) -> Dict[str, float]:
    out: Dict[str, float] = {}
    for person, _, hours in iter_team_non_wip_rows(ws_pab, non_wip_types):
        out[person] = out.get(person, 0.0) + float(hours)
    return {person: float(total) for person, total in out.items() if total != 0}
def compute_team_non_wip_activities(ws_pab: Worksheet, ws_wip_plan: Worksheet, non_wip_types: set[str], ooo_name_col: str, ooo_hours_col: str) -> List[Dict[str, Any]]:
    agg: Dict[Tuple[str, str], float] = {}
    for person, activity, hours in iter_team_non_wip_rows(ws_pab, non_wip_types):
        key = (person, activity)
        agg[key] = agg.get(key, 0.0) + float(hours)
    for person, hours in compute_ooo_by_person_from_wip_plan(ws_wip_plan, ooo_name_col, ooo_hours_col).items():
        key = (person, "OOO")
        agg[key] = agg.get(key, 0.0) + float(hours)
    return [
        {"name": person, "activity": activity, "hours": float(hours)}
        for (person, activity), hours in sorted(
            agg.items(), key=lambda x: (x[0][0].lower(), x[0][1].lower())
        )
        if hours != 0
    ]
def compute_team_wip_workers_ooo_hours(ws_wip_plan: Worksheet, wip_workers: List[str], name_col: str, hours_col: str) -> float:
    worker_keys = {normalize_person_key(x) for x in wip_workers if x}
    if not worker_keys:
        return 0.0
    total = 0.0
    for r in range(1, ws_wip_plan.max_row + 1):
        name = str(ws_wip_plan[f"{name_col}{r}"].value or "").strip()
        if not name:
            continue
        if normalize_person_key(name) not in worker_keys:
            continue
        val = _cell_number(ws_wip_plan[f"{hours_col}{r}"].value)
        if val is not None:
            total += float(val)
    return float(total)
def scrape_one_mapped_workbook(
    path: str,
    people_in_wip_lookup: Dict[Tuple[str, str], List[str]],
    *,
    team: str,
    pab_sheet: str,
    wip_plan_sheet: str,
    perf_metrics_sheet: Optional[str],
    perf_metrics_date_cell: Optional[str],
    perf_wip_sheet: str,
    pct_in_wip_cell: str,
    ooo_total_cell: str,
    ooo_name_col: str,
    ooo_hours_col: str,
    people_count: int,
    non_wip_types: set[str],
    period_date_offset_days: int = 0,
) -> List[Dict[str, Any]]:
    wb = load_workbook(path, data_only=True)
    period: Optional[_dt.date] = None
    if perf_metrics_sheet and perf_metrics_date_cell:
        ws_perf_metrics = _get_required_sheet(wb, perf_metrics_sheet)
        period = parse_period_date_from_perf_metrics_cell(ws_perf_metrics, perf_metrics_date_cell)
        if period is not None and period_date_offset_days:
            period = period + _dt.timedelta(days=period_date_offset_days)
    if period is None:
        period = parse_period_date_from_filename(path)
    if period is None:
        period = parse_period_date_from_workbook_sheetnames(wb)
    if period is None:
        return []
    period_iso = iso_date(period)
    ws_pab = _get_required_sheet(wb, pab_sheet)
    ws_wip_plan = _get_required_sheet(wb, wip_plan_sheet)
    ws_perf = _get_required_sheet(wb, perf_wip_sheet)
    total_non_wip_hours = compute_team_total_non_wip_hours(ws_pab, non_wip_types)
    ooo_hours = float(_cell_number(ws_wip_plan[ooo_total_cell].value) or 0.0)
    pct_in_wip = _cell_number(ws_perf[pct_in_wip_cell].value)
    non_wip_by_person = compute_team_non_wip_by_person(ws_pab, non_wip_types)
    non_wip_activities = compute_team_non_wip_activities(ws_pab, ws_wip_plan, non_wip_types, ooo_name_col, ooo_hours_col)
    wip_workers = people_in_wip_lookup.get((team, period_iso), [])
    wip_workers_count = len({normalize_person_key(x) for x in wip_workers if x})
    wip_workers_ooo_hours = compute_team_wip_workers_ooo_hours(ws_wip_plan, wip_workers, ooo_name_col, ooo_hours_col)
    row = {
        "team": team,
        "period_date": period_iso,
        "people_count": people_count,
        "total_non_wip_hours": float(total_non_wip_hours),
        "OOO Hours": float(ooo_hours),
        "% in WIP": float(pct_in_wip) if pct_in_wip is not None else "",
        "non_wip_by_person": json.dumps(non_wip_by_person, ensure_ascii=False),
        "non_wip_activities": json.dumps(non_wip_activities, ensure_ascii=False),
        "wip_workers": json.dumps(wip_workers, ensure_ascii=False),
        "wip_workers_count": int(wip_workers_count),
        "wip_workers_ooo_hours": float(wip_workers_ooo_hours),
    }
    return [row]
def scrape_one_ds_workbook(path: str, people_in_wip_lookup: Dict[Tuple[str, str], List[str]]) -> List[Dict[str, Any]]:
    return scrape_one_mapped_workbook(
        path,
        people_in_wip_lookup,
        team="DS",
        pab_sheet=DS_PAB_SHEET,
        wip_plan_sheet=DS_WIP_PLAN_SHEET,
        perf_metrics_sheet=None,
        perf_metrics_date_cell=None,
        perf_wip_sheet=DS_PERF_WIP_SHEET,
        pct_in_wip_cell="J2",
        ooo_total_cell="EF2",
        ooo_name_col="DL",
        ooo_hours_col="EF",
        people_count=DS_PEOPLE_COUNT,
        non_wip_types=DS_NON_WIP_TYPES,
    )
def scrape_one_cpt_workbook(path: str, people_in_wip_lookup: Dict[Tuple[str, str], List[str]]) -> List[Dict[str, Any]]:
    return scrape_one_mapped_workbook(
        path,
        people_in_wip_lookup,
        team="CPT",
        pab_sheet=CPT_PAB_SHEET,
        wip_plan_sheet=CPT_WIP_PLAN_SHEET,
        perf_metrics_sheet=None,
        perf_metrics_date_cell=None,
        perf_wip_sheet=CPT_PERF_WIP_SHEET,
        pct_in_wip_cell="J2",
        ooo_total_cell="DT2",
        ooo_name_col="DB",
        ooo_hours_col="DT",
        people_count=CPT_PEOPLE_COUNT,
        non_wip_types=CPT_NON_WIP_TYPES,
    )
def scrape_one_cds_workbook(path: str, people_in_wip_lookup: Dict[Tuple[str, str], List[str]]) -> List[Dict[str, Any]]:
    return scrape_one_mapped_workbook(
        path,
        people_in_wip_lookup,
        team="CDS",
        pab_sheet=CDS_PAB_SHEET,
        wip_plan_sheet=CDS_WIP_PLAN_SHEET,
        perf_metrics_sheet=CDS_PERF_METRICS_SHEET,
        perf_metrics_date_cell="B3",
        perf_wip_sheet=CDS_PERF_WIP_SHEET,
        pct_in_wip_cell="J2",
        ooo_total_cell="CP2",
        ooo_name_col="CC",
        ooo_hours_col="CP",
        people_count=CDS_PEOPLE_COUNT,
        non_wip_types=CDS_NON_WIP_TYPES,
        period_date_offset_days=-4,
    )
def scrape_one_ni_workbook(path: str, people_in_wip_lookup: Dict[Tuple[str, str], List[str]]) -> List[Dict[str, Any]]:
    return scrape_one_mapped_workbook(
        path,
        people_in_wip_lookup,
        team="NI",
        pab_sheet=NI_PAB_SHEET,
        wip_plan_sheet=NI_WIP_PLAN_SHEET,
        perf_metrics_sheet=NI_PERF_METRICS_SHEET,
        perf_metrics_date_cell="B3",
        perf_wip_sheet=NI_PERF_WIP_SHEET,
        pct_in_wip_cell="J2",
        ooo_total_cell="BR2",
        ooo_name_col="BI",
        ooo_hours_col="BR",
        people_count=NI_PEOPLE_COUNT,
        non_wip_types=NI_NON_WIP_TYPES,
        period_date_offset_days=-4,
    )
def expand_input_paths(paths: List[str]) -> List[str]:
    out: List[str] = []
    seen = set()
    def add_file(fp: str) -> None:
        np = _norm_path(fp)
        if np in seen:
            return
        if not os.path.isfile(np):
            return
        ext = os.path.splitext(np)[1].lower()
        if ext not in {".xlsx", ".xlsm"}:
            return
        team = team_for_source(np)
        base = os.path.basename(np)
        if team in {"CDS", "NI"} and "pab" not in base.lower():
            return
        seen.add(np)
        out.append(np)
    for p in paths:
        np = _norm_path(p)
        if os.path.isdir(np):
            for name in sorted(os.listdir(np)):
                add_file(os.path.join(np, name))
        else:
            add_file(np)
    return out
def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument(
        "files",
        nargs="*",
        help="Excel workbook(s) or directory/directories to scrape (.xlsx/.xlsm).",
    )
    ap.add_argument("--crm_wip", default="CRM_DATA\\CRM_WIP.csv", help="Path to CRM_WIP.csv (default: CRM_WIP.csv).")
    ap.add_argument("--out", default="CRM_DATA\\crm_non_wip_activities.csv", help="Output CSV path.")
    args = ap.parse_args()
    inputs = args.files or [
        MCS_DEFAULT_PATH,
        DS_DEFAULT_DIR,
        CPT_DEFAULT_DIR,
        CDS_DEFAULT_DIR,
        CDS_ARCHIVE_PAB_DIR,
        NI_DEFAULT_DIR,
        NI_ARCHIVE_APRIL_2026_DIR,
    ]
    files = expand_input_paths(inputs)
    completed_hours_lookup = load_completed_hours_from_crm_wip(args.crm_wip)
    people_in_wip_lookup = load_people_in_wip_from_crm_wip(args.crm_wip)
    all_rows: List[Dict[str, Any]] = []
    for f in files:
        team = team_for_source(f)
        try:
            if team == "DS":
                all_rows.extend(scrape_one_ds_workbook(f, people_in_wip_lookup))
            elif team == "CPT":
                all_rows.extend(scrape_one_cpt_workbook(f, people_in_wip_lookup))
            elif team == "CDS":
                all_rows.extend(scrape_one_cds_workbook(f, people_in_wip_lookup))
            elif team == "NI":
                all_rows.extend(scrape_one_ni_workbook(f, people_in_wip_lookup))
            else:
                all_rows.extend(scrape_one_workbook(f, completed_hours_lookup))
        except Exception as exc:
            print(f"Skipping {f}: {exc}")
    all_rows.sort(key=lambda r: (str(r.get("team", "")), str(r.get("period_date", ""))))
    with open(args.out, "w", newline="", encoding="utf-8") as fp:
        w = csv.DictWriter(fp, fieldnames=CSV_COLUMNS)
        w.writeheader()
        for r in all_rows:
            w.writerow({k: r.get(k, "") for k in CSV_COLUMNS})
    print(f"Wrote {len(all_rows)} row(s) to {args.out}")
    return 0
if __name__ == "__main__":
    raise SystemExit(main())
