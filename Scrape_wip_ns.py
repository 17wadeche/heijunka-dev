import csv
import json
import os
from datetime import datetime, date
from typing import Any, Dict, Optional, Tuple
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
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
        "%b %d %Y",
        "%b %d, %Y",
        "%B %d %Y",
        "%B %d, %Y",
        "%Y-%m-%d",
        "%m/%d/%Y",
        "%m/%d/%y",
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
import re
from datetime import datetime, date
def parse_sheet_date_scs_missing_year(sheet_name: str) -> str:
    raw = (sheet_name or "").strip()
    low = raw.lower()
    if any(k in low for k in ["template", "agenda", "work instruction", "instructions"]):
        return ""
    m = re.search(r"\b([A-Za-z]{3,9})\s+(\d{1,2})\b", raw)
    if not m:
        return ""
    mon_txt = m.group(1)
    day_txt = m.group(2)
    mm = dd = None
    for fmt in ("%b %d", "%B %d"):
        try:
            dt = datetime.strptime(f"{mon_txt} {day_txt}", fmt)
            mm, dd = dt.month, dt.day
            break
        except ValueError:
            pass
    if mm is None or dd is None:
        return ""
    for y in range(date.today().year, 1999, -1):
        try:
            d = date(y, mm, dd)
        except ValueError:
            continue
        if d.weekday() == 0:  # Monday
            return d.isoformat()
    return ""
def col_range(start_col_letter: str, end_col_letter: str) -> range:
    start = column_index_from_string(start_col_letter.upper())
    end = column_index_from_string(end_col_letter.upper())
    return range(start, end + 1)
def sum_rows(ws, rows: list[int], col: int) -> float:
    return sum(safe_float(ws.cell(row=r, column=col).value) for r in rows)
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
def scrape_workbook_with_config(source_file: str, cfg: Dict[str, Any]) -> list[dict]:
    excel_dir = os.path.dirname(os.path.abspath(source_file))
    timeliness_path = os.path.join(excel_dir, "timeliness.csv")
    closures_path = os.path.join(excel_dir, "closures.csv")
    timeliness_lu, timeliness_err = read_lookup_csv(timeliness_path)
    closures_lu, closures_err = read_lookup_csv(closures_path)
    wb = load_workbook(source_file, data_only=True)
    rows_out: list[dict] = []
    cols = col_range(cfg["person_cols"][0], cfg["person_cols"][1])
    for ws in wb.worksheets:
        date_parser = cfg.get("date_parser", parse_sheet_date)
        period_date = date_parser(ws.title)
        if not period_date:
            continue
        taa_spec = cfg["cells"]["total_available_hours"]
        if isinstance(taa_spec, str):
            total_available_hours = safe_float(ws[taa_spec].value)
        else:
            if taa_spec.get("type") == "sum_range":
                rng = taa_spec["range"]
                total_available_hours = sum(safe_float(cell.value) for row in ws[rng] for cell in row)
            else:
                total_available_hours = 0.0
        completed_spec = cfg["cells"]["completed_hours"]
        if isinstance(completed_spec, str):
            completed_hours = safe_float(ws[completed_spec].value)
        else:
            if completed_spec.get("type") == "sum_range":
                rng = completed_spec["range"]  # e.g. "B60:V60"
                completed_hours = sum(safe_float(cell.value) for row in ws[rng] for cell in row)
            else:
                completed_hours = 0.0
        wp1_out = safe_float(ws[cfg["cells"]["wp1_output"]].value)
        wp2_out = safe_float(ws[cfg["cells"]["wp2_output"]].value)
        wp1_tgt = safe_float(ws[cfg["cells"]["wp1_target"]].value)
        wp2_tgt = safe_float(ws[cfg["cells"]["wp2_target"]].value)
        target_output = wp1_tgt + wp2_tgt
        actual_output = wp1_out + wp2_out
        target_uplh = safe_div(target_output, completed_hours)
        actual_uplh = safe_div(actual_output, completed_hours)
        uplh_wp1 = safe_float(ws[cfg["cells"]["uplh_wp1"]].value)
        uplh_wp2 = safe_float(ws[cfg["cells"]["uplh_wp2"]].value)
        hc_row = cfg["rows"]["hc_row"]
        hc_in_wip = 0
        for c in cols:
            if safe_float(ws.cell(row=hc_row, column=c).value) != 0.0:
                hc_in_wip += 1
        actual_hc_used = safe_div(completed_hours, 32.5)
        person_hours: Dict[str, Dict[str, float]] = {}
        name_row_ph = cfg["rows"]["person_name_row_for_person_hours"]
        actual_row_ph = cfg["rows"]["person_actual_row_for_person_hours"]
        avail_row_ph = cfg["rows"]["person_available_row_for_person_hours"]
        for c in cols:
            name = safe_str(ws.cell(row=name_row_ph, column=c).value)
            if not name:
                continue
            actual = safe_float(ws.cell(row=actual_row_ph, column=c).value)
            available = safe_float(ws.cell(row=avail_row_ph, column=c).value)
            person_hours[name] = {"actual": actual, "available": available}
        outputs_by_person: Dict[str, Dict[str, float]] = {}
        name_row_op = cfg["rows"]["person_name_row_for_outputs_by_person"]
        target_row_op = cfg["rows"]["person_target_row_for_outputs_by_person"]
        output_spec = cfg["outputs_by_person_output"]  
        for c in cols:
            name = safe_str(ws.cell(row=name_row_op, column=c).value)
            if not name:
                continue
            if output_spec["type"] == "row":
                output_val = safe_float(ws.cell(row=output_spec["row"], column=c).value)
            elif output_spec["type"] == "sum_rows":
                output_val = sum_rows(ws, output_spec["rows"], c)
            else:
                output_val = 0.0
            target_val = safe_float(ws.cell(row=target_row_op, column=c).value)
            if output_val != 0.0 or target_val != 0.0:
                outputs_by_person[name] = {"output": output_val, "target": target_val}
        outputs_by_cell = {
            "WP1": {"output": wp1_out, "target": wp1_tgt},
            "WP2": {"output": wp2_out, "target": wp2_tgt},
        }
        cell_station_hours = {
            "WP1": safe_float(ws[cfg["cells"]["wp1_hours"]].value),
            "WP2": safe_float(ws[cfg["cells"]["wp2_hours"]].value),
        }
        hours_by_cell_by_person = {"WP1": {}, "WP2": {}}
        name_row_hc = cfg["rows"]["person_name_row_for_hours_by_cell_by_person"]
        wp1_hour_rows = cfg["rows"]["wp1_hour_rows"]
        wp2_hour_rows = cfg["rows"]["wp2_hour_rows"]
        for c in cols:
            name = safe_str(ws.cell(row=name_row_hc, column=c).value)
            if not name:
                continue
            wp1_hrs = sum_rows(ws, wp1_hour_rows, c)
            wp2_hrs = sum_rows(ws, wp2_hour_rows, c)
            if wp1_hrs != 0.0:
                hours_by_cell_by_person["WP1"][name] = wp1_hrs
            if wp2_hrs != 0.0:
                hours_by_cell_by_person["WP2"][name] = wp2_hrs
        output_by_cell_by_person = {"WP1": {}, "WP2": {}}
        name_row_oc = cfg["rows"]["person_name_row_for_output_by_cell_by_person"]
        wp1_out_rows = cfg["rows"]["wp1_output_rows_by_person"]
        wp2_out_rows = cfg["rows"]["wp2_output_rows_by_person"]
        for c in cols:
            name = safe_str(ws.cell(row=name_row_oc, column=c).value)
            if not name:
                continue
            wp1_o = sum_rows(ws, wp1_out_rows, c)
            wp2_o = sum_rows(ws, wp2_out_rows, c)
            if wp1_o != 0.0:
                output_by_cell_by_person["WP1"][name] = wp1_o
            if wp2_o != 0.0:
                output_by_cell_by_person["WP2"][name] = wp2_o
        uplh_by_cell_by_person: Dict[str, Dict[str, Optional[float]]] = {"WP1": {}, "WP2": {}}
        for wp in ("WP1", "WP2"):
            for person, out_val in output_by_cell_by_person[wp].items():
                hrs = safe_float(hours_by_cell_by_person[wp].get(person, 0.0))
                uplh_by_cell_by_person[wp][person] = safe_div(out_val, hrs)
        team = cfg["team"]
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
        rows_out.append(
            {
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
        )
    return rows_out
def write_csv(rows: list, out_path: str) -> None:
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=HEADERS)
        w.writeheader()
        for r in rows:
            w.writerow({h: r.get(h, "") for h in HEADERS})
def main():
    ph_source_file = r"C:\Users\wadec8\Medtronic PLC\Customer Quality Pelvic Health - Daily Tracker\PH Cell Heijunka.xlsx"
    meic_source_file = r"C:\Users\wadec8\Medtronic PLC\Customer Quality Pelvic Health - Daily Tracker\MEIC\New MEIC PH Heijunka.xlsx"
    scs_source_file = r"C:\Users\wadec8\Medtronic PLC\Customer Quality SCS - Cell 17\Cell 1 - Heijunka.xlsx"
    scs_super_source_file = r"C:\Users\wadec8\Medtronic PLC\Customer Quality SCS - SCS Super Cell\Super Cell Heijunka.xlsx"
    out_file = "NS_metrics.csv"
    if not os.path.exists(ph_source_file):
        raise FileNotFoundError(f"Input file not found: {ph_source_file}")
    if not os.path.exists(meic_source_file):
        raise FileNotFoundError(f"Input file not found: {meic_source_file}")
    if not os.path.exists(scs_source_file):
        raise FileNotFoundError(f"Input file not found: {scs_source_file}")
    PH_CFG = {
        "team": "PH",
        "person_cols": ("B", "R"),
        "cells": {
            "total_available_hours": "T59",
            "completed_hours": "T50",
            "wp1_output": "Z2",
            "wp1_target": "Z7",
            "wp2_output": "AB2",
            "wp2_target": "AB7",
            "uplh_wp1": "Z5",
            "uplh_wp2": "AB5",
            "wp1_hours": "Z4",
            "wp2_hours": "AB4",
        },
        "rows": {
            "hc_row": 50,
            "person_name_row_for_person_hours": 53,
            "person_actual_row_for_person_hours": 50,
            "person_available_row_for_person_hours": 59,
            "person_name_row_for_outputs_by_person": 53,
            "person_target_row_for_outputs_by_person": 25,
            "person_name_row_for_hours_by_cell_by_person": 30,
            "wp1_hour_rows": [31, 35, 39, 43, 47],
            "wp2_hour_rows": [32, 36, 40, 44, 48],
            "person_name_row_for_output_by_cell_by_person": 10,
            "wp1_output_rows_by_person": [11, 14, 17, 20, 23],
            "wp2_output_rows_by_person": [12, 15, 18, 21, 24],
        },
        "outputs_by_person_output": {"type": "sum_rows", "rows": list(range(11, 25))},
    }
    MEIC_PH_CFG = {
        "team": "MEIC PH",
        "person_cols": ("B", "Q"),
        "cells": {
            "total_available_hours": "S111",
            "completed_hours": "S50",
            "wp1_output": "Y2",
            "wp1_target": "Y7",
            "wp2_output": "AA2",
            "wp2_target": "AA7",
            "uplh_wp1": "Y5",
            "uplh_wp2": "AA5",
            "wp1_hours": "Y4",
            "wp2_hours": "AA4",
        },
        "rows": {
            "hc_row": 50,
            "person_name_row_for_person_hours": 30,
            "person_actual_row_for_person_hours": 50,
            "person_available_row_for_person_hours": 111,
            "person_name_row_for_outputs_by_person": 30,
            "person_target_row_for_outputs_by_person": 25,
            "person_name_row_for_hours_by_cell_by_person": 30,
            "wp1_hour_rows": [31, 35, 39, 43, 47],
            "wp2_hour_rows": [32, 36, 40, 44, 48],
            "person_name_row_for_output_by_cell_by_person": 53,
            "wp1_output_rows_by_person": [54, 58, 62, 66, 70],
            "wp2_output_rows_by_person": [55, 57, 63, 67, 71],
        },
        "outputs_by_person_output": {"type": "row", "row": 73},
    }
    SCS_CELL1_CFG = {
        "team": "SCS Cell 1",
        "person_cols": ("B", "R"),
        "date_parser": parse_sheet_date_scs_missing_year,  # <--- key part
        "cells": {
            "total_available_hours": "S111",
            "completed_hours": "S50",
            "wp1_output": "T2",
            "wp1_target": "T7",
            "wp2_output": "V2",
            "wp2_target": "V7",
            "uplh_wp1": "T5",
            "uplh_wp2": "V5",
            "wp1_hours": "T4",
            "wp2_hours": "V4",
        },
        "rows": {
            "hc_row": 25,  # count non-zero in row 25, B..R
            "person_name_row_for_person_hours": 30,
            "person_actual_row_for_person_hours": 50,
            "person_available_row_for_person_hours": 111,
            "person_name_row_for_outputs_by_person": 53,
            "person_target_row_for_outputs_by_person": 25,
            "person_name_row_for_hours_by_cell_by_person": 30,
            "wp1_hour_rows": [31, 35, 39, 43, 47],
            "wp2_hour_rows": [32, 36, 40, 44, 48],
            "person_name_row_for_output_by_cell_by_person": 10,
            "wp1_output_rows_by_person": [54, 58, 62, 66, 70],
            "wp2_output_rows_by_person": [55, 59, 63, 67, 71],
        },
        "outputs_by_person_output": {"type": "row", "row": 73},
    }
    SCS_SUPER_CFG = {
        "team": "SCS Super Cell",
        "person_cols": ("B", "V"),
        "date_parser": parse_sheet_date_scs_missing_year,  # missing-year Monday logic
        "cells": {
            "total_available_hours": {"type": "sum_range", "range": "B60:V60"},
            "completed_hours": {"type": "sum_range", "range": "B60:V60"},
            "wp1_output": "AE2",
            "wp1_target": "AE7",
            "wp2_output": "AG2",
            "wp2_target": "AG7",
            "uplh_wp1": "AE5",
            "uplh_wp2": "AG5",
            "wp1_hours": "AE4",
            "wp2_hours": "AG4",
        },
        "rows": {
            "hc_row": 25,
            "person_name_row_for_person_hours": 30,
            "person_actual_row_for_person_hours": 50,
            "person_available_row_for_person_hours": 60,  # available from row 60
            "person_name_row_for_outputs_by_person": 10,
            "person_target_row_for_outputs_by_person": 25,
            "person_name_row_for_hours_by_cell_by_person": 30,
            "wp1_hour_rows": [31, 35, 39, 43, 47],
            "wp2_hour_rows": [32, 36, 40, 44, 48],
            "person_name_row_for_output_by_cell_by_person": 10,
            "wp1_output_rows_by_person": [11, 14, 17, 20, 23],
            "wp2_output_rows_by_person": [12, 15, 18, 21, 24],
        },
        "outputs_by_person_output": {"type": "sum_rows", "rows": list(range(11, 25))},
    }
    rows = []
    rows.extend(scrape_workbook_with_config(ph_source_file, PH_CFG))
    rows.extend(scrape_workbook_with_config(scs_source_file, SCS_CELL1_CFG))
    meic_rows = scrape_workbook_with_config(meic_source_file, MEIC_PH_CFG)
    scs_super_rows = scrape_workbook_with_config(scs_super_source_file, SCS_SUPER_CFG)
    print("SCS Super Cell rows scraped:", len(scs_super_rows))
    cutoff_super = date.fromisoformat("2025-06-30")
    scs_super_rows = [
        r for r in scs_super_rows
        if safe_str(r.get("period_date")) >= cutoff_super.isoformat()
    ]
    rows.extend(scs_super_rows)
    cutoff = date.fromisoformat("2025-09-01")
    meic_rows = [
        r for r in meic_rows
        if (
            safe_str(r.get("period_date")) >= "2025-09-01"
        )
    ]
    rows.extend(meic_rows)
    rows = [
        r for r in rows
        if (r.get("team") == "SCS Super Cell") or (safe_float(r.get("Total Available Hours")) != 0.0)
    ]
    rows = [r for r in rows if safe_str(r.get("period_date")) != "2023-11-06"]
    rows = [r for r in rows if safe_str(r.get("period_date")) != "2026-09-07"]
    def sort_key(r: dict) -> tuple:
        team = safe_str(r.get("team")).lower()
        d = safe_str(r.get("period_date"))
        if len(d) == 10 and d[4] == "-" and d[7] == "-":
            date_key = d
        else:
            date_key = "9999-12-31"
        return (team, date_key)
    rows.sort(key=sort_key)
    write_csv(rows, out_file)
    print(f"Wrote {len(rows)} rows to {out_file}")
if __name__ == "__main__":
    main()