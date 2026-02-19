import csv
import json
import os
from datetime import datetime
from typing import Any, Dict, Optional, Tuple
from openpyxl import load_workbook
HEADERS = [
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
def safe_float(v: Any) -> float:
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.strip()
        if s == "":
            return 0.0
        try:
            return float(s)
        except ValueError:
            return 0.0
    return 0.0
def safe_str(v: Any) -> str:
    return "" if v is None else str(v).strip()
def safe_div(n: float, d: float) -> Optional[float]:
    return None if d == 0 else (n / d)
def parse_sheet_date(sheet_name: str) -> str:
    name = sheet_name.strip()
    fmts = [
        "%b %d %Y",   # Jan 19 2026
        "%b %d, %Y",  # Jan 19, 2026
        "%B %d %Y",   # January 19 2026
        "%B %d, %Y",  # January 19, 2026
        "%Y-%m-%d",   # 2026-01-19
        "%m/%d/%Y",   # 01/19/2026
        "%m/%d/%y",   # 01/19/26
    ]
    for fmt in fmts:
        try:
            dt = datetime.strptime(name, fmt).date()
            return dt.isoformat()
        except ValueError:
            pass
    try:
        from dateutil import parser  # type: ignore
        dt = parser.parse(name, fuzzy=True).date()
        return dt.isoformat()
    except Exception:
        return name
def col_range_B_to_R() -> range:
    return range(2, 19)
def sum_range(ws, row_start: int, row_end: int, col: int) -> float:
    total = 0.0
    for r in range(row_start, row_end + 1):
        total += safe_float(ws.cell(row=r, column=col).value)
    return total
def read_lookup_csv(path: str) -> Tuple[Dict[Tuple[str, str], Dict[str, Any]], str]:
    lookup: Dict[Tuple[str, str], Dict[str, Any]] = {}
    if not os.path.exists(path):
        return lookup, f"Missing file: {os.path.basename(path)}"

    try:
        with open(path, "r", newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for row in reader:
                team = safe_str(row.get("team"))
                period_date = safe_str(row.get("period_date"))
                if team and period_date:
                    lookup[(team, period_date)] = row
        return lookup, ""
    except Exception as e:
        return lookup, f"Failed reading {os.path.basename(path)}: {e}"
def build_person_hours(ws) -> Dict[str, Dict[str, float]]:
    out: Dict[str, Dict[str, float]] = {}
    for c in col_range_B_to_R():
        name = safe_str(ws.cell(row=53, column=c).value)
        if not name:
            continue
        actual = safe_float(ws.cell(row=50, column=c).value)
        available = safe_float(ws.cell(row=59, column=c).value)
        out[name] = {"actual": actual, "available": available}
    return out
def build_outputs_by_person(ws) -> Dict[str, Dict[str, float]]:
    out: Dict[str, Dict[str, float]] = {}
    for c in col_range_B_to_R():
        name = safe_str(ws.cell(row=53, column=c).value)
        if not name:
            continue
        output = sum_range(ws, 11, 24, c)
        target = safe_float(ws.cell(row=25, column=c).value)
        if output != 0.0 or target != 0.0:
            out[name] = {"output": output, "target": target}
    return out
def build_outputs_by_cell(ws) -> Dict[str, Dict[str, float]]:
    return {
        "WP1": {"output": safe_float(ws["Z2"].value), "target": safe_float(ws["Z7"].value)},
        "WP2": {"output": safe_float(ws["AB2"].value), "target": safe_float(ws["AB7"].value)},
    }
def build_cell_station_hours(ws) -> Dict[str, float]:
    return {"WP1": safe_float(ws["Z4"].value), "WP2": safe_float(ws["AB4"].value)}
def build_hours_by_cell_by_person(ws) -> Dict[str, Dict[str, float]]:
    wp1_rows = [31, 35, 39, 43, 47]
    wp2_rows = [32, 36, 40, 44, 48]
    wp1: Dict[str, float] = {}
    wp2: Dict[str, float] = {}
    for c in col_range_B_to_R():
        name = safe_str(ws.cell(row=30, column=c).value)
        if not name:
            continue
        wp1_hours = sum(safe_float(ws.cell(row=r, column=c).value) for r in wp1_rows)
        wp2_hours = sum(safe_float(ws.cell(row=r, column=c).value) for r in wp2_rows)
        if wp1_hours != 0.0:
            wp1[name] = wp1_hours
        if wp2_hours != 0.0:
            wp2[name] = wp2_hours
    return {"WP1": wp1, "WP2": wp2}
def build_output_by_cell_by_person(ws) -> Dict[str, Dict[str, float]]:
    wp1_rows = [11, 14, 17, 20, 23]
    wp2_rows = [12, 15, 18, 21, 24]
    wp1: Dict[str, float] = {}
    wp2: Dict[str, float] = {}
    for c in col_range_B_to_R():
        name = safe_str(ws.cell(row=10, column=c).value)
        if not name:
            continue
        wp1_out = sum(safe_float(ws.cell(row=r, column=c).value) for r in wp1_rows)
        wp2_out = sum(safe_float(ws.cell(row=r, column=c).value) for r in wp2_rows)
        if wp1_out != 0.0:
            wp1[name] = wp1_out
        if wp2_out != 0.0:
            wp2[name] = wp2_out
    return {"WP1": wp1, "WP2": wp2}
def build_uplh_by_cell_by_person(
    output_by_cell_by_person: Dict[str, Dict[str, float]],
    hours_by_cell_by_person: Dict[str, Dict[str, float]],
) -> Dict[str, Dict[str, Optional[float]]]:
    out: Dict[str, Dict[str, Optional[float]]] = {"WP1": {}, "WP2": {}}
    for wp in ("WP1", "WP2"):
        for person, out_val in output_by_cell_by_person.get(wp, {}).items():
            hrs = hours_by_cell_by_person.get(wp, {}).get(person, 0.0)
            out[wp][person] = safe_div(out_val, hrs)
    return out
def count_hc_in_wip(ws) -> int:
    count = 0
    for c in col_range_B_to_R():
        if safe_float(ws.cell(row=50, column=c).value) != 0.0:
            count += 1
    return count
def scrape_workbook(source_file: str) -> list:
    excel_dir = os.path.dirname(os.path.abspath(source_file))
    timeliness_path = os.path.join(excel_dir, "timeliness.csv")
    closures_path = os.path.join(excel_dir, "closures.csv")
    timeliness_lu, timeliness_err = read_lookup_csv(timeliness_path)
    closures_lu, closures_err = read_lookup_csv(closures_path)
    wb = load_workbook(source_file, data_only=True)
    rows = []
    for ws in wb.worksheets:
        period_date = parse_sheet_date(ws.title)
        total_available_hours = safe_float(ws["T59"].value)
        completed_hours = safe_float(ws["T50"].value)
        target_output = safe_float(ws["Z7"].value) + safe_float(ws["AB7"].value)
        actual_output = safe_float(ws["Z2"].value) + safe_float(ws["AB2"].value)
        target_uplh = safe_div(target_output, completed_hours)
        actual_uplh = safe_div(actual_output, completed_hours)
        uplh_wp1 = safe_float(ws["Z5"].value)
        uplh_wp2 = safe_float(ws["AB5"].value)
        hc_in_wip = count_hc_in_wip(ws)
        actual_hc_used = safe_div(completed_hours, 32.5)
        person_hours = build_person_hours(ws)
        outputs_by_person = build_outputs_by_person(ws)
        outputs_by_cell = build_outputs_by_cell(ws)
        cell_station_hours = build_cell_station_hours(ws)
        hours_by_cell_by_person = build_hours_by_cell_by_person(ws)
        output_by_cell_by_person = build_output_by_cell_by_person(ws)
        uplh_by_cell_by_person = build_uplh_by_cell_by_person(
            output_by_cell_by_person, hours_by_cell_by_person
        )
        team = "PH"
        key = (team, period_date)
        open_complaint_timeliness = ""
        closures = ""
        opened = ""
        trow = timeliness_lu.get(key)
        if trow is not None:
            open_complaint_timeliness = safe_str(trow.get("Open Complaint Timeliness"))
        crow = closures_lu.get(key)
        if crow is not None:
            closures = safe_str(crow.get("Closures"))
            opened = safe_str(crow.get("Opened"))
        errs = []
        if timeliness_err:
            errs.append(timeliness_err)
        if closures_err:
            errs.append(closures_err)
        if not trow and not timeliness_err:
            errs.append(f"No timeliness match for {team} {period_date}")
        if not crow and not closures_err:
            errs.append(f"No closures match for {team} {period_date}")
        row = {
            "team": team,
            "period_date": period_date,
            "source_file": source_file,
            "Total Available Hours": total_available_hours,
            "Completed Hours": completed_hours,
            "Target Output": target_output,
            "Actual Output": actual_output,
            "Target UPLH": target_uplh,
            "Actual UPLH": actual_uplh,
            "UPLH WP1": uplh_wp1,
            "UPLH WP2": uplh_wp2,
            "HC in WIP": hc_in_wip,
            "Actual HC Used": actual_hc_used,
            "People in WIP": "",
            "Person Hours": json.dumps(person_hours, ensure_ascii=False),
            "Outputs by Person": json.dumps(outputs_by_person, ensure_ascii=False),
            "Outputs by Cell/Station": json.dumps(outputs_by_cell, ensure_ascii=False),
            "Cell/Station Hours": json.dumps(cell_station_hours, ensure_ascii=False),
            "Hours by Cell/Station - by person": json.dumps(hours_by_cell_by_person, ensure_ascii=False),
            "Output by Cell/Station - by person": json.dumps(output_by_cell_by_person, ensure_ascii=False),
            "UPLH by Cell/Station - by person": json.dumps(uplh_by_cell_by_person, ensure_ascii=False),
            "Open Complaint Timeliness": open_complaint_timeliness,
            "error": " | ".join(errs) if errs else "",
            "Closures": closures,
            "Opened": opened,
        }
        rows.append(row)
    return rows
def write_csv(rows: list, out_path: str) -> None:
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=HEADERS)
        w.writeheader()
        for r in rows:
            w.writerow({h: r.get(h, "") for h in HEADERS})
def main():
    source_file = r"C:\Users\wadec8\Medtronic PLC\Customer Quality Pelvic Health - Daily Tracker\PH Cell Heijunka.xlsx"
    out_file = "NS_metrics.csv"
    if not os.path.exists(source_file):
        raise FileNotFoundError(
            f"Input file not found: {source_file}\n"
            "Tip: edit `source_file` in this script or pass a valid path."
        )
    EPS = 1e-9
    rows = scrape_workbook(source_file)
    rows = [r for r in rows if abs(safe_float(r.get("Total Available Hours"))) > EPS]
    rows = [r for r in rows if safe_str(r.get("period_date")) != "2023-11-06"]
    write_csv(rows, out_file)
    print(f"Wrote {len(rows)} rows to {out_file}")
if __name__ == "__main__":
    main()