import csv
import json
import os
import re
import shutil
import tempfile
import uuid
from datetime import datetime, date, timedelta
from typing import Any, Dict, Optional, Tuple
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
import win32com.client
import time
import pythoncom
import pywintypes
import argparse
import logging
import sys
import traceback
from contextlib import contextmanager
from threading import Thread, Event
import math
WIP_HEADERS = [
    "team",
    "period_date",
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
def json_load_safe(s: Any) -> dict:
    if s is None:
        return {}
    if isinstance(s, dict):
        return s
    txt = safe_str(s)
    if not txt:
        return {}
    try:
        return json.loads(txt)
    except Exception:
        return {}
def _sum_nested_person_map(dst: dict, src: dict, keys=("actual", "available")) -> None:
    for name, rec in (src or {}).items():
        if not isinstance(rec, dict):
            continue
        drec = dst.setdefault(name, {k: 0.0 for k in keys})
        for k in keys:
            drec[k] = safe_float(drec.get(k)) + safe_float(rec.get(k))
def _sum_nested_output_target_map(dst: dict, src: dict) -> None:
    for name, rec in (src or {}).items():
        if not isinstance(rec, dict):
            continue
        drec = dst.setdefault(name, {"output": 0.0, "target": 0.0})
        drec["output"] = safe_float(drec.get("output")) + safe_float(rec.get("output"))
        drec["target"] = safe_float(drec.get("target")) + safe_float(rec.get("target"))
def _sum_cell_map(dst: dict, src: dict) -> None:
    for cell, rec in (src or {}).items():
        if not isinstance(rec, dict):
            continue
        drec = dst.setdefault(cell, {"output": 0.0, "target": 0.0})
        drec["output"] = safe_float(drec.get("output")) + safe_float(rec.get("output"))
        drec["target"] = safe_float(drec.get("target")) + safe_float(rec.get("target"))
def _sum_simple_map(dst: dict, src: dict) -> None:
    for k, v in (src or {}).items():
        dst[k] = safe_float(dst.get(k)) + safe_float(v)
def _sum_cell_person_map(dst: dict, src: dict) -> None:
    for cell, people in (src or {}).items():
        if not isinstance(people, dict):
            continue
        dcell = dst.setdefault(cell, {})
        for person, val in people.items():
            dcell[person] = safe_float(dcell.get(person)) + safe_float(val)
def _recalc_uplh_by_cell_person(hours_by_cell_person: dict, out_by_cell_person: dict) -> dict:
    uplh = {}
    for cell in ("WP1", "WP2"):
        uplh[cell] = {}
        hmap = (hours_by_cell_person or {}).get(cell, {}) or {}
        omap = (out_by_cell_person or {}).get(cell, {}) or {}
        for person in set(list(hmap.keys()) + list(omap.keys())):
            hrs = safe_float(hmap.get(person))
            outv = safe_float(omap.get(person))
            uplh[cell][person] = safe_div(outv, hrs)
    return uplh
def build_ns_wip_rows(all_rows: list[dict]) -> list[dict]:
    et_combine_teams = {"O-Arm MEIC", "Nav", "Mazor", "AE MEIC", "CSF"}
    et_label = "Enabling Technologies"
    rollups = [
        ({"DBS C13", "DBS C14"}, "DBS"),
        ({"MEIC PH", "PH", "PH Cell 17"}, "PH"),
        ({"SCS Cell 1", "SCS Super Cell"}, "SCS"),
        (et_combine_teams, et_label),
    ]
    rename_map = {"TDD COS 1": "TDD"}
    buckets_by_label: Dict[str, Dict[str, list[dict]]] = {label: {} for _, label in rollups}
    passthrough: list[dict] = []
    for r in all_rows:
        team = safe_str(r.get("team"))
        if team in rename_map:
            r = dict(r)
            r["team"] = rename_map[team]
            team = r["team"]
        placed = False
        pd = safe_str(r.get("period_date"))
        for team_set, label in rollups:
            if team in team_set and pd:
                buckets_by_label[label].setdefault(pd, []).append(r)
                placed = True
                break
        if not placed:
            passthrough.append(r)
    out_rows: list[dict] = []
    out_rows.extend(passthrough)
    def _emit_rollup(label: str, period_date: str, rows: list[dict]) -> None:
        taa = 0.0
        ch = 0.0
        tgt_out = 0.0
        act_out = 0.0
        hc_wip = 0.0
        wp1_out_total = 0.0
        wp2_out_total = 0.0
        wp1_uplh_weighted_sum = 0.0
        wp2_uplh_weighted_sum = 0.0
        person_hours = {}
        outputs_by_person = {}
        outputs_by_cell = {}
        cell_station_hours = {}
        hours_by_cell_person = {}
        out_by_cell_person = {}
        open_timeliness = ""
        closures = ""
        opened = ""
        errs = []
        for r in rows:
            taa += safe_float(r.get("Total Available Hours"))
            ch += safe_float(r.get("Completed Hours"))
            tgt_out += safe_float(r.get("Target Output"))
            act_out += safe_float(r.get("Actual Output"))
            hc_wip += int(safe_float(r.get("HC in WIP")))
            cell_json = json_load_safe(r.get("Outputs by Cell/Station"))
            wp1_out = safe_float(((cell_json.get("WP1") or {}).get("output")))
            wp2_out = safe_float(((cell_json.get("WP2") or {}).get("output")))
            wp1_u = r.get("UPLH WP1")
            wp2_u = r.get("UPLH WP2")
            if wp1_out > 0:
                wp1_out_total += wp1_out
                wp1_uplh_weighted_sum += safe_float(wp1_u) * wp1_out
            if wp2_out > 0:
                wp2_out_total += wp2_out
                wp2_uplh_weighted_sum += safe_float(wp2_u) * wp2_out
            _sum_nested_person_map(person_hours, json_load_safe(r.get("Person Hours")), keys=("actual", "available"))
            _sum_nested_output_target_map(outputs_by_person, json_load_safe(r.get("Outputs by Person")))
            _sum_cell_map(outputs_by_cell, cell_json)
            _sum_simple_map(cell_station_hours, json_load_safe(r.get("Cell/Station Hours")))
            _sum_cell_person_map(hours_by_cell_person, json_load_safe(r.get("Hours by Cell/Station - by person")))
            _sum_cell_person_map(out_by_cell_person, json_load_safe(r.get("Output by Cell/Station - by person")))
            er = safe_str(r.get("error"))
            if er:
                errs.append(er)
        target_uplh = safe_div(tgt_out, ch)
        actual_uplh = safe_div(act_out, ch)
        uplh_wp1 = safe_div(wp1_uplh_weighted_sum, wp1_out_total)  # weighted avg
        uplh_wp2 = safe_div(wp2_uplh_weighted_sum, wp2_out_total)  # weighted avg
        actual_hc_used = safe_div(ch, 32.5)
        uplh_by_cell_person = _recalc_uplh_by_cell_person(hours_by_cell_person, out_by_cell_person)
        out_rows.append({
            "team": label,
            "period_date": period_date,
            "Total Available Hours": taa,
            "Completed Hours": ch,
            "Target Output": tgt_out,
            "Actual Output": act_out,
            "Target UPLH": target_uplh,
            "Actual UPLH": actual_uplh,
            "UPLH WP1": uplh_wp1,
            "UPLH WP2": uplh_wp2,
            "HC in WIP": hc_wip,
            "Actual HC Used": actual_hc_used,
            "People in WIP": "",
            "Person Hours": json.dumps(person_hours, ensure_ascii=False),
            "Outputs by Person": json.dumps(outputs_by_person, ensure_ascii=False),
            "Outputs by Cell/Station": json.dumps(outputs_by_cell, ensure_ascii=False),
            "Cell/Station Hours": json.dumps(cell_station_hours, ensure_ascii=False),
            "Hours by Cell/Station - by person": json.dumps(hours_by_cell_person, ensure_ascii=False),
            "Output by Cell/Station - by person": json.dumps(out_by_cell_person, ensure_ascii=False),
            "UPLH by Cell/Station - by person": json.dumps(uplh_by_cell_person, ensure_ascii=False),
            "Open Complaint Timeliness": open_timeliness,
            "error": " | ".join(errs) if errs else "",
            "Closures": closures,
            "Opened": opened,
        })
    for _, label in rollups:
        for period_date in sorted(buckets_by_label[label].keys()):
            _emit_rollup(label, period_date, buckets_by_label[label][period_date])
    def sort_key_wip(r: dict) -> tuple:
        team = safe_str(r.get("team")).lower()
        d = safe_str(r.get("period_date"))
        date_key = d if (len(d) == 10 and d[4] == "-" and d[7] == "-") else "9999-12-31"
        return (team, date_key)
    out_rows.sort(key=sort_key_wip)
    return out_rows
def write_csv_wip(rows: list[dict], out_path: str) -> None:
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=WIP_HEADERS)
        w.writeheader()
        for r in rows:
            w.writerow({h: r.get(h, "") for h in WIP_HEADERS})
def setup_logging(log_path: str = "NS_metrics.log") -> logging.Logger:
    logger = logging.getLogger("ns_metrics")
    logger.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setFormatter(fmt)
    fh.setLevel(logging.INFO)
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(fmt)
    ch.setLevel(logging.INFO)
    if not logger.handlers:
        logger.addHandler(fh)
        logger.addHandler(ch)
    return logger
@contextmanager
def heartbeat(logger: logging.Logger, label: str, every_seconds: int = 120):
    stop = Event()
    def _run():
        while not stop.wait(every_seconds):
            logger.info(f"[{label}] still running...")
    t = Thread(target=_run, daemon=True)
    t.start()
    try:
        yield
    finally:
        stop.set()
        t.join(timeout=1)
def run_team(logger: logging.Logger, team_name: str, fn):
    start = datetime.now()
    logger.info(f"[{team_name}] START")
    try:
        with heartbeat(logger, team_name, every_seconds=180):  # adjust heartbeat
            rows = fn()
        elapsed = datetime.now() - start
        logger.info(f"[{team_name}] DONE | rows={len(rows)} | elapsed={elapsed}")
        return rows
    except Exception as e:
        elapsed = datetime.now() - start
        logger.error(f"[{team_name}] FAIL | elapsed={elapsed} | error={e}")
        logger.error(traceback.format_exc())
        return []
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
def _excel_col_range(start_col: str, end_col: str) -> list[int]:
    s = column_index_from_string(start_col)
    e = column_index_from_string(end_col)
    return list(range(s, e + 1))
def _as_iso_date(v: Any) -> str:
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    s = safe_str(v)
    if not s:
        return ""
    return parse_sheet_date(s)
def _get_dropdown_values_from_validation(cell) -> list[Any]:
    try:
        formula1 = cell.Validation.Formula1
    except Exception:
        return []
    if not formula1:
        return []
    f = str(formula1).strip()
    if not f.startswith("="):
        vals = [x.strip() for x in f.split(",")]
        return [v for v in vals if v]
    try:
        rng = cell.Parent.Range(f.lstrip("="))
        vals = []
        v = rng.Value
        if isinstance(v, tuple):
            for row in v:
                if isinstance(row, tuple):
                    for item in row:
                        if safe_str(item):
                            vals.append(item)
                else:
                    if safe_str(row):
                        vals.append(row)
        else:
            if safe_str(v):
                vals.append(v)
        return vals
    except Exception:
        try:
            app = cell.Application
            rng = app.Range(f.lstrip("="))
            v = rng.Value
            vals = []
            if isinstance(v, tuple):
                for row in v:
                    if isinstance(row, tuple):
                        for item in row:
                            if safe_str(item):
                                vals.append(item)
                    else:
                        if safe_str(row):
                            vals.append(row)
            else:
                if safe_str(v):
                    vals.append(v)
            return vals
        except Exception:
            return []
def _com_call(fn, tries: int = 30, sleep_s: float = 0.25):
    for i in range(tries):
        try:
            return fn()
        except pywintypes.com_error as e:
            if e.args and e.args[0] == -2147418111:
                time.sleep(sleep_s)
                continue
            raise
    return fn()
def _open_workbook_via_temp_copy(excel, source_file: str):
    src = os.path.abspath(os.path.expandvars(source_file))
    if not os.path.exists(src):
        raise FileNotFoundError(f"File not found: {src}")
    tmp_dir = tempfile.gettempdir()
    base = os.path.splitext(os.path.basename(src))[0]
    ext = os.path.splitext(src)[1]
    tmp_path = os.path.join(tmp_dir, f"{base}__{uuid.uuid4().hex}{ext}")
    shutil.copy2(src, tmp_path)
    wb = excel.Workbooks.Open(
        tmp_path,
        UpdateLinks=0,
        ReadOnly=True,
        IgnoreReadOnlyRecommended=True,
        Notify=False,
        AddToMru=False,
        CorruptLoad=0,
    )
    return wb, tmp_path
def scrape_dbs_previous_weeks_xlsm(source_file: str, team: str, dropdown_override: Optional[list[Any]] = None) -> list[dict]:
    import shutil
    import tempfile
    import uuid
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")  # new isolated Excel instance
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.EnableEvents = False
    excel.AutomationSecurity = 3  
    wb = None
    tmp_path = None
    rows_out: list[dict] = []
    def _open_via_temp_copy(src_path: str):
        nonlocal tmp_path
        src = os.path.abspath(os.path.expandvars(src_path))
        if not os.path.exists(src):
            raise FileNotFoundError(f"File not found on disk: {src}")
        base = os.path.splitext(os.path.basename(src))[0]
        ext = os.path.splitext(src)[1]
        tmp_path = os.path.join(tempfile.gettempdir(), f"{base}__{uuid.uuid4().hex}{ext}")
        shutil.copy2(src, tmp_path)
        return _com_call(lambda: excel.Workbooks.Open(
            tmp_path,
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
            Notify=False,
            AddToMru=False,
            CorruptLoad=0,  # xlNormalLoad
        ))
    try:
        wb = _open_via_temp_copy(source_file)
        ws = _com_call(lambda: wb.Worksheets("Previous Weeks"))
        dd = _com_call(lambda: ws.Range("A2"))
        dropdown_values = dropdown_override if dropdown_override is not None else _get_dropdown_values_from_validation(dd)
        seen = set()
        dropdown_values = [v for v in dropdown_values if not (safe_str(v) in seen or seen.add(safe_str(v)))]
        cols = _excel_col_range("B", "M")
        excel_dir = os.path.dirname(os.path.abspath(os.path.expandvars(source_file)))
        timeliness_path = os.path.join(excel_dir, "timeliness.csv")
        closures_path = os.path.join(excel_dir, "closures.csv")
        timeliness_lu, timeliness_err = read_lookup_csv(timeliness_path)
        closures_lu, closures_err = read_lookup_csv(closures_path)
        for choice in dropdown_values:
            _com_call(lambda: setattr(dd, "Value", choice))
            _com_call(lambda: excel.Calculate())
            period_date = _as_iso_date(_com_call(lambda: dd.Value))
            if not period_date:
                continue
            total_available_hours = safe_float(_com_call(lambda: ws.Range("O69").Value))
            completed_hours = safe_float(_com_call(lambda: ws.Range("O59").Value))
            wp1_tgt = safe_float(_com_call(lambda: ws.Range("T10").Value))
            wp2_tgt = safe_float(_com_call(lambda: ws.Range("V10").Value))
            wp1_out = safe_float(_com_call(lambda: ws.Range("T5").Value))
            wp2_out = safe_float(_com_call(lambda: ws.Range("V5").Value))
            target_output = wp1_tgt + wp2_tgt
            actual_output = wp1_out + wp2_out
            target_uplh = safe_div(target_output, completed_hours)
            actual_uplh = safe_div(actual_output, completed_hours)
            uplh_wp1 = safe_float(_com_call(lambda: ws.Range("T8").Value))
            uplh_wp2 = safe_float(_com_call(lambda: ws.Range("V8").Value))
            hc_in_wip = 0
            for c in cols:
                if safe_float(_com_call(lambda c=c: ws.Cells(59, c).Value)) != 0.0:
                    hc_in_wip += 1
            actual_hc_used = safe_div(completed_hours, 32.5)
            person_hours: Dict[str, Dict[str, float]] = {}
            for c in cols:
                name = safe_str(_com_call(lambda c=c: ws.Cells(33, c).Value))
                if not name:
                    continue
                actual = safe_float(_com_call(lambda c=c: ws.Cells(59, c).Value))
                available = safe_float(_com_call(lambda c=c: ws.Cells(69, c).Value))
                person_hours[name] = {"actual": actual, "available": available}
            outputs_by_person: Dict[str, Dict[str, float]] = {}
            for c in cols:
                name = safe_str(_com_call(lambda c=c: ws.Cells(13, c).Value))
                if not name:
                    continue
                out_val = sum(
                    safe_float(_com_call(lambda r=r, c=c: ws.Cells(r, c).Value))
                    for r in range(14, 28)
                )
                tgt_val = safe_float(_com_call(lambda c=c: ws.Cells(28, c).Value))
                if out_val != 0.0 or tgt_val != 0.0:
                    outputs_by_person[name] = {"output": out_val, "target": tgt_val}
            outputs_by_cell = {
                "WP1": {"output": wp1_out, "target": wp1_tgt},
                "WP2": {"output": wp2_out, "target": wp2_tgt},
            }
            cell_station_hours = {
                "WP1": safe_float(_com_call(lambda: ws.Range("T7").Value)),
                "WP2": safe_float(_com_call(lambda: ws.Range("V7").Value)),
            }
            hours_by_cell_by_person = {"WP1": {}, "WP2": {}}
            wp1_hour_rows = [34, 39, 44, 49, 54]
            wp2_hour_rows = [35, 40, 45, 50, 55]
            for c in cols:
                name = safe_str(_com_call(lambda c=c: ws.Cells(33, c).Value))
                if not name:
                    continue
                wp1_hrs = sum(safe_float(_com_call(lambda r=r, c=c: ws.Cells(r, c).Value)) for r in wp1_hour_rows)
                wp2_hrs = sum(safe_float(_com_call(lambda r=r, c=c: ws.Cells(r, c).Value)) for r in wp2_hour_rows)
                if wp1_hrs != 0.0:
                    hours_by_cell_by_person["WP1"][name] = wp1_hrs
                if wp2_hrs != 0.0:
                    hours_by_cell_by_person["WP2"][name] = wp2_hrs
            output_by_cell_by_person = {"WP1": {}, "WP2": {}}
            wp1_out_rows = [14, 17, 20, 23, 26]
            wp2_out_rows = [15, 18, 21, 24, 27]
            for c in cols:
                name = safe_str(_com_call(lambda c=c: ws.Cells(13, c).Value))
                if not name:
                    continue
                wp1_o = sum(safe_float(_com_call(lambda r=r, c=c: ws.Cells(r, c).Value)) for r in wp1_out_rows)
                wp2_o = sum(safe_float(_com_call(lambda r=r, c=c: ws.Cells(r, c).Value)) for r in wp2_out_rows)
                if wp1_o != 0.0:
                    output_by_cell_by_person["WP1"][name] = wp1_o
                if wp2_o != 0.0:
                    output_by_cell_by_person["WP2"][name] = wp2_o
            uplh_by_cell_by_person: Dict[str, Dict[str, Optional[float]]] = {"WP1": {}, "WP2": {}}
            for wp in ("WP1", "WP2"):
                for person, out_val in output_by_cell_by_person[wp].items():
                    hrs = safe_float(hours_by_cell_by_person[wp].get(person, 0.0))
                    uplh_by_cell_by_person[wp][person] = safe_div(out_val, hrs)
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
            rows_out.append({
                "team": team,
                "period_date": period_date,
                "source_file": os.path.abspath(os.path.expandvars(source_file)),
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
            })
        return rows_out
    finally:
        try:
            if wb is not None:
                _com_call(lambda: wb.Close(SaveChanges=False), tries=10, sleep_s=0.3)
        except Exception:
            pass
        try:
            _com_call(lambda: excel.Quit(), tries=10, sleep_s=0.3)
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        try:
            if tmp_path and os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception:
            pass
def scrape_nav_previous_weeks_xlsm(source_file: str, team: str = "Nav", dropdown_override: Optional[list[Any]] = None) -> list[dict]:
    import shutil
    import tempfile
    import uuid
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")  # isolated Excel instance
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.EnableEvents = False
    excel.AutomationSecurity = 3  # force-disable macros
    wb = None
    tmp_path = None
    rows_out: list[dict] = []
    def _open_via_temp_copy(src_path: str):
        nonlocal tmp_path
        src = os.path.abspath(os.path.expandvars(src_path))
        if not os.path.exists(src):
            raise FileNotFoundError(f"File not found on disk: {src}")
        base = os.path.splitext(os.path.basename(src))[0]
        ext = os.path.splitext(src)[1]
        tmp_path = os.path.join(tempfile.gettempdir(), f"{base}__{uuid.uuid4().hex}{ext}")
        shutil.copy2(src, tmp_path)
        return _com_call(
            lambda: excel.Workbooks.Open(
                tmp_path,
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
                Notify=False,
                AddToMru=False,
                CorruptLoad=0,
            )
        )
    try:
        wb = _open_via_temp_copy(source_file)
        ws = _com_call(lambda: wb.Worksheets("Previous Weeks"))
        dd = _com_call(lambda: ws.Range("A2"))
        dropdown_values = dropdown_override if dropdown_override is not None else _get_dropdown_values_from_validation(dd)
        seen = set()
        dropdown_values = [v for v in dropdown_values if not (safe_str(v) in seen or seen.add(safe_str(v)))]
        cols = _excel_col_range("B", "V")
        excel_dir = os.path.dirname(os.path.abspath(os.path.expandvars(source_file)))
        timeliness_path = os.path.join(excel_dir, "timeliness.csv")
        closures_path = os.path.join(excel_dir, "closures.csv")
        timeliness_lu, timeliness_err = read_lookup_csv(timeliness_path)
        closures_lu, closures_err = read_lookup_csv(closures_path)
        today_iso = date.today().isoformat() 
        for choice in dropdown_values:
            _com_call(lambda: setattr(dd, "Value", choice))
            _com_call(lambda: excel.Calculate())
            period_date = _as_iso_date(_com_call(lambda: dd.Value))
            if not period_date:
                continue
            if period_date < "2025-06-02":
                continue
            if period_date > today_iso:
                continue
            total_available_hours = safe_float(_com_call(lambda: ws.Range("X64").Value))
            completed_hours = safe_float(_com_call(lambda: ws.Range("X54").Value))
            wp1_tgt = safe_float(_com_call(lambda: ws.Range("AD10").Value))
            wp2_tgt = safe_float(_com_call(lambda: ws.Range("AF10").Value))
            wp1_out = safe_float(_com_call(lambda: ws.Range("AD5").Value))
            wp2_out = safe_float(_com_call(lambda: ws.Range("AF5").Value))
            target_output = wp1_tgt + wp2_tgt
            actual_output = wp1_out + wp2_out
            if target_output < 0:
                continue
            target_uplh = safe_div(target_output, completed_hours)
            actual_uplh = safe_div(actual_output, completed_hours)
            uplh_wp1 = safe_float(_com_call(lambda: ws.Range("AD8").Value))
            uplh_wp2 = safe_float(_com_call(lambda: ws.Range("AF8").Value))
            hc_in_wip = 0
            for c in cols:
                if safe_float(_com_call(lambda c=c: ws.Cells(28, c).Value)) != 0.0:
                    hc_in_wip += 1
            actual_hc_used = safe_div(completed_hours, 32.5)
            person_hours: Dict[str, Dict[str, float]] = {}
            for c in cols:
                name = safe_str(_com_call(lambda c=c: ws.Cells(33, c).Value))
                if not name:
                    continue
                actual = safe_float(_com_call(lambda c=c: ws.Cells(54, c).Value))
                available = safe_float(_com_call(lambda c=c: ws.Cells(64, c).Value))
                person_hours[name] = {"actual": actual, "available": available}
            outputs_by_person: Dict[str, Dict[str, float]] = {}
            for c in cols:
                name = safe_str(_com_call(lambda c=c: ws.Cells(13, c).Value))
                if not name:
                    continue
                out_val = sum(
                    safe_float(_com_call(lambda r=r, c=c: ws.Cells(r, c).Value))
                    for r in range(14, 28)  # 14..27 inclusive
                )
                tgt_val = safe_float(_com_call(lambda c=c: ws.Cells(28, c).Value))
                if out_val != 0.0 or tgt_val != 0.0:
                    outputs_by_person[name] = {"output": out_val, "target": tgt_val}
            outputs_by_cell = {
                "WP1": {"output": wp1_out, "target": wp1_tgt},
                "WP2": {"output": wp2_out, "target": wp2_tgt},
            }
            cell_station_hours = {
                "WP1": safe_float(_com_call(lambda: ws.Range("AD7").Value)),
                "WP2": safe_float(_com_call(lambda: ws.Range("AF7").Value)),
            }
            hours_by_cell_by_person = {"WP1": {}, "WP2": {}}
            wp1_hour_rows = [34, 38, 42, 46, 50]
            wp2_hour_rows = [35, 39, 43, 47, 51]
            for c in cols:
                name = safe_str(_com_call(lambda c=c: ws.Cells(33, c).Value))
                if not name:
                    continue
                wp1_hrs = sum(safe_float(_com_call(lambda r=r, c=c: ws.Cells(r, c).Value)) for r in wp1_hour_rows)
                wp2_hrs = sum(safe_float(_com_call(lambda r=r, c=c: ws.Cells(r, c).Value)) for r in wp2_hour_rows)
                if wp1_hrs != 0.0:
                    hours_by_cell_by_person["WP1"][name] = wp1_hrs
                if wp2_hrs != 0.0:
                    hours_by_cell_by_person["WP2"][name] = wp2_hrs
            output_by_cell_by_person = {"WP1": {}, "WP2": {}}
            wp1_out_rows = [14, 17, 20, 23, 26]
            wp2_out_rows = [15, 18, 21, 24, 27]
            for c in cols:
                name = safe_str(_com_call(lambda c=c: ws.Cells(13, c).Value))
                if not name:
                    continue
                wp1_o = sum(safe_float(_com_call(lambda r=r, c=c: ws.Cells(r, c).Value)) for r in wp1_out_rows)
                wp2_o = sum(safe_float(_com_call(lambda r=r, c=c: ws.Cells(r, c).Value)) for r in wp2_out_rows)
                if wp1_o != 0.0:
                    output_by_cell_by_person["WP1"][name] = wp1_o
                if wp2_o != 0.0:
                    output_by_cell_by_person["WP2"][name] = wp2_o
            uplh_by_cell_by_person: Dict[str, Dict[str, Optional[float]]] = {"WP1": {}, "WP2": {}}
            for wp in ("WP1", "WP2"):
                for person, out_val in output_by_cell_by_person[wp].items():
                    hrs = safe_float(hours_by_cell_by_person[wp].get(person, 0.0))
                    uplh_by_cell_by_person[wp][person] = safe_div(out_val, hrs)
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
                    "source_file": os.path.abspath(os.path.expandvars(source_file)),
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
    finally:
        try:
            if wb is not None:
                _com_call(lambda: wb.Close(SaveChanges=False), tries=10, sleep_s=0.3)
        except Exception:
            pass
        try:
            _com_call(lambda: excel.Quit(), tries=10, sleep_s=0.3)
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        try:
            if tmp_path and os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception:
            pass
def scrape_meic_ae_oarm_previous_weeks_xlsm(source_file: str, team: str, dropdown_override: Optional[list[Any]] = None) -> list[dict]:
    import shutil
    import tempfile
    import uuid
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.EnableEvents = False
    excel.AutomationSecurity = 3  # disable macros
    wb = None
    tmp_path = None
    rows_out: list[dict] = []
    def _open_via_temp_copy(src_path: str):
        nonlocal tmp_path
        src = os.path.abspath(os.path.expandvars(src_path))
        if not os.path.exists(src):
            raise FileNotFoundError(f"File not found on disk: {src}")
        base = os.path.splitext(os.path.basename(src))[0]
        ext = os.path.splitext(src)[1]
        tmp_path = os.path.join(tempfile.gettempdir(), f"{base}__{uuid.uuid4().hex}{ext}")
        shutil.copy2(src, tmp_path)
        return _com_call(
            lambda: excel.Workbooks.Open(
                tmp_path,
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
                Notify=False,
                AddToMru=False,
                CorruptLoad=0,
            )
        )
    try:
        wb = _open_via_temp_copy(source_file)
        ws = _com_call(lambda: wb.Worksheets("Previous Weeks"))
        dd = _com_call(lambda: ws.Range("A2"))
        dropdown_values = dropdown_override if dropdown_override is not None else _get_dropdown_values_from_validation(dd)
        seen = set()
        dropdown_values = [v for v in dropdown_values if not (safe_str(v) in seen or seen.add(safe_str(v)))]
        cols = _excel_col_range("B", "P") 
        excel_dir = os.path.dirname(os.path.abspath(os.path.expandvars(source_file)))
        timeliness_path = os.path.join(excel_dir, "timeliness.csv")
        closures_path = os.path.join(excel_dir, "closures.csv")
        timeliness_lu, timeliness_err = read_lookup_csv(timeliness_path)
        closures_lu, closures_err = read_lookup_csv(closures_path)
        today_iso = date.today().isoformat()
        for choice in dropdown_values:
            _com_call(lambda: setattr(dd, "Value", choice))
            _com_call(lambda: excel.Calculate())
            period_date = _as_iso_date(_com_call(lambda: dd.Value))
            if not period_date:
                continue
            if period_date < "2025-06-02":
                continue
            if period_date > today_iso:
                continue
            total_available_hours = safe_float(_com_call(lambda: ws.Range("R64").Value))
            completed_hours = safe_float(_com_call(lambda: ws.Range("R54").Value))
            wp1_tgt = safe_float(_com_call(lambda: ws.Range("X10").Value))
            wp2_tgt = safe_float(_com_call(lambda: ws.Range("Z10").Value))
            wp1_out = safe_float(_com_call(lambda: ws.Range("X5").Value))
            wp2_out = safe_float(_com_call(lambda: ws.Range("Z5").Value))
            target_output = wp1_tgt + wp2_tgt
            actual_output = wp1_out + wp2_out
            if target_output < 0:
                continue
            target_uplh = safe_div(target_output, completed_hours)
            actual_uplh = safe_div(actual_output, completed_hours)
            uplh_wp1 = safe_float(_com_call(lambda: ws.Range("X8").Value))
            uplh_wp2 = safe_float(_com_call(lambda: ws.Range("Z8").Value))
            hc_in_wip = 0
            for c in cols:
                if safe_float(_com_call(lambda c=c: ws.Cells(28, c).Value)) != 0.0:
                    hc_in_wip += 1
            actual_hc_used = safe_div(completed_hours, 32.5)
            person_hours: Dict[str, Dict[str, float]] = {}
            for c in cols:
                name = safe_str(_com_call(lambda c=c: ws.Cells(33, c).Value))
                if not name:
                    continue
                actual = safe_float(_com_call(lambda c=c: ws.Cells(54, c).Value))
                available = safe_float(_com_call(lambda c=c: ws.Cells(64, c).Value))
                person_hours[name] = {"actual": actual, "available": available}
            outputs_by_person: Dict[str, Dict[str, float]] = {}
            for c in cols:
                name = safe_str(_com_call(lambda c=c: ws.Cells(13, c).Value))
                if not name:
                    continue
                out_val = sum(
                    safe_float(_com_call(lambda r=r, c=c: ws.Cells(r, c).Value))
                    for r in range(14, 28)  # 14..27 inclusive
                )
                tgt_val = safe_float(_com_call(lambda c=c: ws.Cells(28, c).Value))
                if out_val != 0.0 or tgt_val != 0.0:
                    outputs_by_person[name] = {"output": out_val, "target": tgt_val}
            outputs_by_cell = {
                "WP1": {"output": wp1_out, "target": wp1_tgt},
                "WP2": {"output": wp2_out, "target": wp2_tgt},
            }
            cell_station_hours = {
                "WP1": safe_float(_com_call(lambda: ws.Range("X7").Value)),
                "WP2": safe_float(_com_call(lambda: ws.Range("Z7").Value)),
            }
            hours_by_cell_by_person = {"WP1": {}, "WP2": {}}
            wp1_hour_rows = [34, 38, 42, 46, 50]
            wp2_hour_rows = [35, 39, 43, 47, 51]
            for c in cols:
                name = safe_str(_com_call(lambda c=c: ws.Cells(33, c).Value))
                if not name:
                    continue
                wp1_hrs = sum(safe_float(_com_call(lambda r=r, c=c: ws.Cells(r, c).Value)) for r in wp1_hour_rows)
                wp2_hrs = sum(safe_float(_com_call(lambda r=r, c=c: ws.Cells(r, c).Value)) for r in wp2_hour_rows)
                if wp1_hrs != 0.0:
                    hours_by_cell_by_person["WP1"][name] = wp1_hrs
                if wp2_hrs != 0.0:
                    hours_by_cell_by_person["WP2"][name] = wp2_hrs
            output_by_cell_by_person = {"WP1": {}, "WP2": {}}
            wp1_out_rows = [14, 17, 20, 23, 26]
            wp2_out_rows = [15, 18, 21, 24, 27]
            for c in cols:
                name = safe_str(_com_call(lambda c=c: ws.Cells(13, c).Value))
                if not name:
                    continue
                wp1_o = sum(safe_float(_com_call(lambda r=r, c=c: ws.Cells(r, c).Value)) for r in wp1_out_rows)
                wp2_o = sum(safe_float(_com_call(lambda r=r, c=c: ws.Cells(r, c).Value)) for r in wp2_out_rows)
                if wp1_o != 0.0:
                    output_by_cell_by_person["WP1"][name] = wp1_o
                if wp2_o != 0.0:
                    output_by_cell_by_person["WP2"][name] = wp2_o
            uplh_by_cell_by_person: Dict[str, Dict[str, Optional[float]]] = {"WP1": {}, "WP2": {}}
            for wp in ("WP1", "WP2"):
                for person, out_val in output_by_cell_by_person[wp].items():
                    hrs = safe_float(hours_by_cell_by_person[wp].get(person, 0.0))
                    uplh_by_cell_by_person[wp][person] = safe_div(out_val, hrs)
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
                    "source_file": os.path.abspath(os.path.expandvars(source_file)),
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
    finally:
        try:
            if wb is not None:
                _com_call(lambda: wb.Close(SaveChanges=False), tries=10, sleep_s=0.3)
        except Exception:
            pass
        try:
            _com_call(lambda: excel.Quit(), tries=10, sleep_s=0.3)
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        try:
            if tmp_path and os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception:
            pass
def scrape_previous_weeks_xlsm_with_filters(source_file: str, team: str, cfg: Dict[str, Any], dropdown_override: Optional[list[Any]] = None) -> list[dict]:
    import shutil
    import tempfile
    import uuid
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.EnableEvents = False
    excel.AutomationSecurity = 3  # disable macros
    wb = None
    tmp_path = None
    rows_out: list[dict] = []
    def _open_via_temp_copy(src_path: str):
        nonlocal tmp_path
        src = os.path.abspath(os.path.expandvars(src_path))
        if not os.path.exists(src):
            raise FileNotFoundError(f"File not found on disk: {src}")
        base = os.path.splitext(os.path.basename(src))[0]
        ext = os.path.splitext(src)[1]
        tmp_path = os.path.join(tempfile.gettempdir(), f"{base}__{uuid.uuid4().hex}{ext}")
        shutil.copy2(src, tmp_path)
        return _com_call(
            lambda: excel.Workbooks.Open(
                tmp_path,
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
                Notify=False,
                AddToMru=False,
                CorruptLoad=0,
            )
        )
    try:
        wb = _open_via_temp_copy(source_file)
        ws = _com_call(lambda: wb.Worksheets("Previous Weeks"))
        dd = _com_call(lambda: ws.Range("A2"))
        dropdown_values = dropdown_override if dropdown_override is not None else _get_dropdown_values_from_validation(dd)
        seen = set()
        dropdown_values = [v for v in dropdown_values if not (safe_str(v) in seen or seen.add(safe_str(v)))]
        cols = _excel_col_range(cfg["person_cols"][0], cfg["person_cols"][1])
        excel_dir = os.path.dirname(os.path.abspath(os.path.expandvars(source_file)))
        timeliness_path = os.path.join(excel_dir, "timeliness.csv")
        closures_path = os.path.join(excel_dir, "closures.csv")
        timeliness_lu, timeliness_err = read_lookup_csv(timeliness_path)
        closures_lu, closures_err = read_lookup_csv(closures_path)
        today_iso = date.today().isoformat()
        for choice in dropdown_values:
            _com_call(lambda: setattr(dd, "Value", choice))
            _com_call(lambda: excel.Calculate())
            period_date = _as_iso_date(_com_call(lambda: dd.Value))
            if not period_date:
                continue
            if period_date < "2025-06-02":
                continue
            if period_date > today_iso:
                continue
            total_available_hours = safe_float(_com_call(lambda: ws.Range(cfg["cells"]["total_available_hours"]).Value))
            completed_hours = safe_float(_com_call(lambda: ws.Range(cfg["cells"]["completed_hours"]).Value))
            wp1_tgt = safe_float(_com_call(lambda: ws.Range(cfg["cells"]["wp1_target"]).Value))
            wp2_tgt = safe_float(_com_call(lambda: ws.Range(cfg["cells"]["wp2_target"]).Value))
            wp1_out = safe_float(_com_call(lambda: ws.Range(cfg["cells"]["wp1_output"]).Value))
            wp2_out = safe_float(_com_call(lambda: ws.Range(cfg["cells"]["wp2_output"]).Value))
            target_output = wp1_tgt + wp2_tgt
            actual_output = wp1_out + wp2_out
            if target_output < 0:
                continue
            target_uplh = safe_div(target_output, completed_hours)
            actual_uplh = safe_div(actual_output, completed_hours)
            uplh_wp1 = safe_float(_com_call(lambda: ws.Range(cfg["cells"]["uplh_wp1"]).Value))
            uplh_wp2 = safe_float(_com_call(lambda: ws.Range(cfg["cells"]["uplh_wp2"]).Value))
            hc_in_wip = 0
            for c in cols:
                if safe_float(_com_call(lambda c=c: ws.Cells(28, c).Value)) != 0.0:
                    hc_in_wip += 1
            actual_hc_used = safe_div(completed_hours, 32.5)
            person_hours: Dict[str, Dict[str, float]] = {}
            name_row_ph = cfg["rows"]["person_name_row_for_person_hours"]
            actual_row_ph = cfg["rows"]["person_actual_row_for_person_hours"]
            avail_row_ph = cfg["rows"]["person_available_row_for_person_hours"]
            for c in cols:
                name = safe_str(_com_call(lambda c=c: ws.Cells(name_row_ph, c).Value))
                if not name:
                    continue
                actual = safe_float(_com_call(lambda c=c: ws.Cells(actual_row_ph, c).Value))
                available = safe_float(_com_call(lambda c=c: ws.Cells(avail_row_ph, c).Value))
                person_hours[name] = {"actual": actual, "available": available}
            outputs_by_person: Dict[str, Dict[str, float]] = {}
            name_row_op = cfg["rows"]["person_name_row_for_outputs_by_person"]
            target_row_op = cfg["rows"]["person_target_row_for_outputs_by_person"]
            for c in cols:
                name = safe_str(_com_call(lambda c=c: ws.Cells(name_row_op, c).Value))
                if not name:
                    continue
                out_val = sum(
                    safe_float(_com_call(lambda r=r, c=c: ws.Cells(r, c).Value))
                    for r in range(14, 28)  # 14..27
                )
                tgt_val = safe_float(_com_call(lambda c=c: ws.Cells(target_row_op, c).Value))
                if out_val != 0.0 or tgt_val != 0.0:
                    outputs_by_person[name] = {"output": out_val, "target": tgt_val}
            outputs_by_cell = {
                "WP1": {"output": wp1_out, "target": wp1_tgt},
                "WP2": {"output": wp2_out, "target": wp2_tgt},
            }
            cell_station_hours = {
                "WP1": safe_float(_com_call(lambda: ws.Range(cfg["cells"]["wp1_hours"]).Value)),
                "WP2": safe_float(_com_call(lambda: ws.Range(cfg["cells"]["wp2_hours"]).Value)),
            }
            hours_by_cell_by_person = {"WP1": {}, "WP2": {}}
            name_row_hc = cfg["rows"]["person_name_row_for_hours_by_cell_by_person"]
            wp1_hour_rows = cfg["rows"]["wp1_hour_rows"]
            wp2_hour_rows = cfg["rows"]["wp2_hour_rows"]
            for c in cols:
                name = safe_str(_com_call(lambda c=c: ws.Cells(name_row_hc, c).Value))
                if not name:
                    continue
                wp1_hrs = sum(safe_float(_com_call(lambda r=r, c=c: ws.Cells(r, c).Value)) for r in wp1_hour_rows)
                wp2_hrs = sum(safe_float(_com_call(lambda r=r, c=c: ws.Cells(r, c).Value)) for r in wp2_hour_rows)
                if wp1_hrs != 0.0:
                    hours_by_cell_by_person["WP1"][name] = wp1_hrs
                if wp2_hrs != 0.0:
                    hours_by_cell_by_person["WP2"][name] = wp2_hrs
            output_by_cell_by_person = {"WP1": {}, "WP2": {}}
            name_row_oc = cfg["rows"]["person_name_row_for_output_by_cell_by_person"]
            wp1_out_rows = cfg["rows"]["wp1_output_rows_by_person"]
            wp2_out_rows = cfg["rows"]["wp2_output_rows_by_person"]
            for c in cols:
                name = safe_str(_com_call(lambda c=c: ws.Cells(name_row_oc, c).Value))
                if not name:
                    continue
                wp1_o = sum(safe_float(_com_call(lambda r=r, c=c: ws.Cells(r, c).Value)) for r in wp1_out_rows)
                wp2_o = sum(safe_float(_com_call(lambda r=r, c=c: ws.Cells(r, c).Value)) for r in wp2_out_rows)
                if wp1_o != 0.0:
                    output_by_cell_by_person["WP1"][name] = wp1_o
                if wp2_o != 0.0:
                    output_by_cell_by_person["WP2"][name] = wp2_o
            uplh_by_cell_by_person: Dict[str, Dict[str, Optional[float]]] = {"WP1": {}, "WP2": {}}
            for wp in ("WP1", "WP2"):
                for person, out_val in output_by_cell_by_person[wp].items():
                    hrs = safe_float(hours_by_cell_by_person[wp].get(person, 0.0))
                    uplh_by_cell_by_person[wp][person] = safe_div(out_val, hrs)
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
                    "source_file": os.path.abspath(os.path.expandvars(source_file)),
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
    finally:
        try:
            if wb is not None:
                _com_call(lambda: wb.Close(SaveChanges=False), tries=10, sleep_s=0.3)
        except Exception:
            pass
        try:
            _com_call(lambda: excel.Quit(), tries=10, sleep_s=0.3)
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        try:
            if tmp_path and os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception:
            pass
def scrape_previous_weeks_xlsm_with_filters(source_file: str, team: str, cfg: Dict[str, Any], dropdown_override: Optional[list[Any]] = None) -> list[dict]:
    import shutil
    import tempfile
    import uuid
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.EnableEvents = False
    excel.AutomationSecurity = 3  # disable macros
    wb = None
    tmp_path = None
    rows_out: list[dict] = []
    def _open_via_temp_copy(src_path: str):
        nonlocal tmp_path
        src = os.path.abspath(os.path.expandvars(src_path))
        if not os.path.exists(src):
            raise FileNotFoundError(f"File not found on disk: {src}")
        base = os.path.splitext(os.path.basename(src))[0]
        ext = os.path.splitext(src)[1]
        tmp_path = os.path.join(tempfile.gettempdir(), f"{base}__{uuid.uuid4().hex}{ext}")
        shutil.copy2(src, tmp_path)
        return _com_call(
            lambda: excel.Workbooks.Open(
                tmp_path,
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
                Notify=False,
                AddToMru=False,
                CorruptLoad=0,
            )
        )
    try:
        wb = _open_via_temp_copy(source_file)
        ws = _com_call(lambda: wb.Worksheets("Previous Weeks"))
        dd = _com_call(lambda: ws.Range("A2"))
        dropdown_values = dropdown_override if dropdown_override is not None else _get_dropdown_values_from_validation(dd)
        seen = set()
        dropdown_values = [v for v in dropdown_values if not (safe_str(v) in seen or seen.add(safe_str(v)))]
        cols = _excel_col_range(cfg["person_cols"][0], cfg["person_cols"][1])
        excel_dir = os.path.dirname(os.path.abspath(os.path.expandvars(source_file)))
        timeliness_path = os.path.join(excel_dir, "timeliness.csv")
        closures_path = os.path.join(excel_dir, "closures.csv")
        timeliness_lu, timeliness_err = read_lookup_csv(timeliness_path)
        closures_lu, closures_err = read_lookup_csv(closures_path)
        today_iso = date.today().isoformat()
        for choice in dropdown_values:
            _com_call(lambda: setattr(dd, "Value", choice))
            _com_call(lambda: excel.Calculate())
            period_date = _as_iso_date(_com_call(lambda: dd.Value))
            if not period_date:
                continue
            if period_date < "2025-06-02":
                continue
            if period_date > today_iso:
                continue
            total_available_hours = safe_float(_com_call(lambda: ws.Range(cfg["cells"]["total_available_hours"]).Value))
            completed_hours = safe_float(_com_call(lambda: ws.Range(cfg["cells"]["completed_hours"]).Value))
            wp1_tgt = safe_float(_com_call(lambda: ws.Range(cfg["cells"]["wp1_target"]).Value))
            wp2_tgt = safe_float(_com_call(lambda: ws.Range(cfg["cells"]["wp2_target"]).Value))
            wp1_out = safe_float(_com_call(lambda: ws.Range(cfg["cells"]["wp1_output"]).Value))
            wp2_out = safe_float(_com_call(lambda: ws.Range(cfg["cells"]["wp2_output"]).Value))
            target_output = wp1_tgt + wp2_tgt
            actual_output = wp1_out + wp2_out
            if target_output < 0:
                continue
            target_uplh = safe_div(target_output, completed_hours)
            actual_uplh = safe_div(actual_output, completed_hours)
            uplh_wp1 = safe_float(_com_call(lambda: ws.Range(cfg["cells"]["uplh_wp1"]).Value))
            uplh_wp2 = safe_float(_com_call(lambda: ws.Range(cfg["cells"]["uplh_wp2"]).Value))
            hc_in_wip = 0
            for c in cols:
                if safe_float(_com_call(lambda c=c: ws.Cells(28, c).Value)) != 0.0:
                    hc_in_wip += 1
            actual_hc_used = safe_div(completed_hours, 32.5)
            person_hours: Dict[str, Dict[str, float]] = {}
            name_row_ph = cfg["rows"]["person_name_row_for_person_hours"]
            actual_row_ph = cfg["rows"]["person_actual_row_for_person_hours"]
            avail_row_ph = cfg["rows"]["person_available_row_for_person_hours"]
            for c in cols:
                name = safe_str(_com_call(lambda c=c: ws.Cells(name_row_ph, c).Value))
                if not name:
                    continue
                actual = safe_float(_com_call(lambda c=c: ws.Cells(actual_row_ph, c).Value))
                available = safe_float(_com_call(lambda c=c: ws.Cells(avail_row_ph, c).Value))
                person_hours[name] = {"actual": actual, "available": available}
            outputs_by_person: Dict[str, Dict[str, float]] = {}
            name_row_op = cfg["rows"]["person_name_row_for_outputs_by_person"]
            target_row_op = cfg["rows"]["person_target_row_for_outputs_by_person"]
            for c in cols:
                name = safe_str(_com_call(lambda c=c: ws.Cells(name_row_op, c).Value))
                if not name:
                    continue
                out_val = sum(
                    safe_float(_com_call(lambda r=r, c=c: ws.Cells(r, c).Value))
                    for r in range(14, 28)  # 14..27
                )
                tgt_val = safe_float(_com_call(lambda c=c: ws.Cells(target_row_op, c).Value))
                if out_val != 0.0 or tgt_val != 0.0:
                    outputs_by_person[name] = {"output": out_val, "target": tgt_val}
            outputs_by_cell = {
                "WP1": {"output": wp1_out, "target": wp1_tgt},
                "WP2": {"output": wp2_out, "target": wp2_tgt},
            }
            cell_station_hours = {
                "WP1": safe_float(_com_call(lambda: ws.Range(cfg["cells"]["wp1_hours"]).Value)),
                "WP2": safe_float(_com_call(lambda: ws.Range(cfg["cells"]["wp2_hours"]).Value)),
            }
            hours_by_cell_by_person = {"WP1": {}, "WP2": {}}
            name_row_hc = cfg["rows"]["person_name_row_for_hours_by_cell_by_person"]
            wp1_hour_rows = cfg["rows"]["wp1_hour_rows"]
            wp2_hour_rows = cfg["rows"]["wp2_hour_rows"]
            for c in cols:
                name = safe_str(_com_call(lambda c=c: ws.Cells(name_row_hc, c).Value))
                if not name:
                    continue
                wp1_hrs = sum(safe_float(_com_call(lambda r=r, c=c: ws.Cells(r, c).Value)) for r in wp1_hour_rows)
                wp2_hrs = sum(safe_float(_com_call(lambda r=r, c=c: ws.Cells(r, c).Value)) for r in wp2_hour_rows)
                if wp1_hrs != 0.0:
                    hours_by_cell_by_person["WP1"][name] = wp1_hrs
                if wp2_hrs != 0.0:
                    hours_by_cell_by_person["WP2"][name] = wp2_hrs
            output_by_cell_by_person = {"WP1": {}, "WP2": {}}
            name_row_oc = cfg["rows"]["person_name_row_for_output_by_cell_by_person"]
            wp1_out_rows = cfg["rows"]["wp1_output_rows_by_person"]
            wp2_out_rows = cfg["rows"]["wp2_output_rows_by_person"]
            for c in cols:
                name = safe_str(_com_call(lambda c=c: ws.Cells(name_row_oc, c).Value))
                if not name:
                    continue
                wp1_o = sum(safe_float(_com_call(lambda r=r, c=c: ws.Cells(r, c).Value)) for r in wp1_out_rows)
                wp2_o = sum(safe_float(_com_call(lambda r=r, c=c: ws.Cells(r, c).Value)) for r in wp2_out_rows)
                if wp1_o != 0.0:
                    output_by_cell_by_person["WP1"][name] = wp1_o
                if wp2_o != 0.0:
                    output_by_cell_by_person["WP2"][name] = wp2_o
            uplh_by_cell_by_person: Dict[str, Dict[str, Optional[float]]] = {"WP1": {}, "WP2": {}}
            for wp in ("WP1", "WP2"):
                for person, out_val in output_by_cell_by_person[wp].items():
                    hrs = safe_float(hours_by_cell_by_person[wp].get(person, 0.0))
                    uplh_by_cell_by_person[wp][person] = safe_div(out_val, hrs)
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
                    "source_file": os.path.abspath(os.path.expandvars(source_file)),
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
    finally:
        try:
            if wb is not None:
                _com_call(lambda: wb.Close(SaveChanges=False), tries=10, sleep_s=0.3)
        except Exception:
            pass
        try:
            _com_call(lambda: excel.Quit(), tries=10, sleep_s=0.3)
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        try:
            if tmp_path and os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception:
            pass
def week_monday_iso(d: date) -> str:
    monday = d - timedelta(days=d.weekday())
    return monday.isoformat()
def _parse_any_date_to_date(v: Any) -> Optional[date]:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = safe_str(v)
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    try:
        iso = parse_sheet_date(s)
        return date.fromisoformat(iso)
    except Exception:
        return None
def read_ent_team_tenure_mapping(
    xlsx_path: str,
    sheet_name: str = "Next Week Forecast",
    start_row: int = 2,
    end_row: int = 30,
    start_col_letter: str = "D",
    end_col_letter: str = "J",
    name_col_letter: str = "A",
) -> Tuple[float, Dict[str, float]]:
    wb = load_workbook(xlsx_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        ws = wb.active
    else:
        ws = wb[sheet_name]
    sc = column_index_from_string(start_col_letter)
    ec = column_index_from_string(end_col_letter)
    name_c = column_index_from_string(name_col_letter)
    total = 0.0
    per_person: Dict[str, float] = {}
    for r in range(start_row, end_row + 1):
        row_sum = 0.0
        for c in range(sc, ec + 1):
            row_sum += safe_float(ws.cell(row=r, column=c).value)
        total += row_sum
        nm = safe_str(ws.cell(row=r, column=name_c).value)
        if nm:
            per_person[nm] = row_sum
    return total, per_person
def _ent_cache_path(base_dir: str) -> str:
    return os.path.join(base_dir, "ent_total_available_cache.json")
def get_ent_total_available_for_week(
    mapping_xlsx: str,
    week_monday: str,
    today: Optional[date] = None,
) -> Tuple[float, Dict[str, float], str]:
    if today is None:
        today = date.today()
    base_dir = os.path.dirname(os.path.abspath(os.path.expandvars(mapping_xlsx)))
    cache_file = _ent_cache_path(base_dir)
    cache: Dict[str, Any] = {}
    if os.path.exists(cache_file):
        try:
            with open(cache_file, "r", encoding="utf-8") as f:
                cache = json.load(f) or {}
        except Exception:
            cache = {}
    is_monday = (today.weekday() == 0)
    if is_monday:
        total, per_person = read_ent_team_tenure_mapping(mapping_xlsx)
        cache[week_monday] = {"total": total, "per_person": per_person, "refreshed_on": today.isoformat()}
        try:
            with open(cache_file, "w", encoding="utf-8") as f:
                json.dump(cache, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
        return total, per_person, f"ENT mapping refreshed (Monday) and cached in {os.path.basename(cache_file)}"
    wk = cache.get(week_monday)
    if isinstance(wk, dict) and ("total" in wk) and ("per_person" in wk):
        return safe_float(wk.get("total")), (wk.get("per_person") or {}), f"ENT mapping loaded from cache {os.path.basename(cache_file)}"
    total, per_person = read_ent_team_tenure_mapping(mapping_xlsx)
    cache[week_monday] = {"total": total, "per_person": per_person, "refreshed_on": today.isoformat(), "note": "cache-miss fallback"}
    try:
        with open(cache_file, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    return total, per_person, f"ENT mapping cache miss; read mapping and cached in {os.path.basename(cache_file)}"
def scrape_ent_from_csv(
    ent_csv_path: str,
    mapping_xlsx_path: str,
    team: str = "ENT",
) -> list[dict]:
    if not os.path.exists(ent_csv_path):
        return [{
            "team": team,
            "period_date": "",
            "source_file": ent_csv_path,
        }]
    weekly: Dict[str, Dict[str, Any]] = {}
    with open(ent_csv_path, "r", newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        headers = next(reader, None)  # skip header row
        for row in reader:
            if not row or len(row) < 6:
                continue
            name = safe_str(row[0])
            d_raw = row[1]
            d_parsed = _parse_any_date_to_date(d_raw)
            if not d_parsed:
                continue
            wk = week_monday_iso(d_parsed)
            wp1_out = safe_float(row[2])
            wp2_out = safe_float(row[3])
            wp1_hrs = safe_float(row[4])
            wp2_hrs = safe_float(row[5])
            rec = weekly.setdefault(wk, {
                "wp1_out": 0.0, "wp2_out": 0.0,
                "wp1_hrs": 0.0, "wp2_hrs": 0.0,
                "by_person": {},  # name -> accumulators
            })
            rec["wp1_out"] += wp1_out
            rec["wp2_out"] += wp2_out
            rec["wp1_hrs"] += wp1_hrs
            rec["wp2_hrs"] += wp2_hrs
            if name:
                p = rec["by_person"].setdefault(name, {"wp1_out": 0.0, "wp2_out": 0.0, "wp1_hrs": 0.0, "wp2_hrs": 0.0})
                p["wp1_out"] += wp1_out
                p["wp2_out"] += wp2_out
                p["wp1_hrs"] += wp1_hrs
                p["wp2_hrs"] += wp2_hrs
    rows_out: list[dict] = []
    for period_date in sorted(weekly.keys()):
        agg = weekly[period_date]
        completed_hours = agg["wp1_hrs"] + agg["wp2_hrs"]
        actual_output = agg["wp1_out"] + agg["wp2_out"]
        taa, per_person_avail, taa_note = get_ent_total_available_for_week(mapping_xlsx_path, period_date)
        actual_uplh = safe_div(actual_output, completed_hours)
        uplh_wp1 = safe_div(agg["wp1_out"], agg["wp1_hrs"])
        uplh_wp2 = safe_div(agg["wp2_out"], agg["wp2_hrs"])
        hc_in_wip = 0
        for nm, pdata in (agg["by_person"] or {}).items():
            if (pdata["wp1_out"] + pdata["wp2_out"]) > 0:
                hc_in_wip += 1
        actual_hc_used = safe_div(completed_hours, 32.5)
        person_hours: Dict[str, Dict[str, float]] = {}
        for nm, pdata in (agg["by_person"] or {}).items():
            actual_person = pdata["wp1_hrs"] + pdata["wp2_hrs"]
            available_person = safe_float(per_person_avail.get(nm, 0.0))
            person_hours[nm] = {"actual": actual_person, "available": available_person}
        outputs_by_person: Dict[str, Dict[str, float]] = {}
        for nm, pdata in (agg["by_person"] or {}).items():
            out_person = pdata["wp1_out"] + pdata["wp2_out"]
            outputs_by_person[nm] = {"output": out_person, "target": 0.0}
        outputs_by_cell = {
            "WP1": {"output": agg["wp1_out"], "target": 0.0},
            "WP2": {"output": agg["wp2_out"], "target": 0.0},
        }
        cell_station_hours = {
            "WP1": agg["wp1_hrs"],
            "WP2": agg["wp2_hrs"],
        }
        hours_by_cell_by_person = {"WP1": {}, "WP2": {}}
        for nm, pdata in (agg["by_person"] or {}).items():
            hours_by_cell_by_person["WP1"][nm] = pdata["wp1_hrs"]
            hours_by_cell_by_person["WP2"][nm] = pdata["wp2_hrs"]
        output_by_cell_by_person = {"WP1": {}, "WP2": {}}
        for nm, pdata in (agg["by_person"] or {}).items():
            output_by_cell_by_person["WP1"][nm] = pdata["wp1_out"]
            output_by_cell_by_person["WP2"][nm] = pdata["wp2_out"]
        uplh_by_cell_by_person: Dict[str, Dict[str, Optional[float]]] = {"WP1": {}, "WP2": {}}
        for nm in (agg["by_person"] or {}).keys():
            uplh_by_cell_by_person["WP1"][nm] = safe_div(output_by_cell_by_person["WP1"][nm], hours_by_cell_by_person["WP1"][nm])
            uplh_by_cell_by_person["WP2"][nm] = safe_div(output_by_cell_by_person["WP2"][nm], hours_by_cell_by_person["WP2"][nm])
        errs = []
        if taa_note:
            errs.append(taa_note)
        rows_out.append({
            "team": team,
            "period_date": period_date,  # ALWAYS Monday
            "source_file": f"{os.path.abspath(os.path.expandvars(ent_csv_path))} | {os.path.abspath(os.path.expandvars(mapping_xlsx_path))}",
            "Total Available Hours": taa,
            "Completed Hours": completed_hours,
            "Target Output": "",  # blank per spec
            "Actual Output": actual_output,
            "Target UPLH": "",    # blank per spec
            "Actual UPLH": actual_uplh,
            "UPLH WP1": uplh_wp1,
            "UPLH WP2": uplh_wp2,
            "HC in WIP": hc_in_wip,
            "Actual HC Used": actual_hc_used,
            "People in WIP": "",  # blank per spec
            "Person Hours": json.dumps(person_hours, ensure_ascii=False),
            "Outputs by Person": json.dumps(outputs_by_person, ensure_ascii=False),
            "Outputs by Cell/Station": json.dumps(outputs_by_cell, ensure_ascii=False),
            "Cell/Station Hours": json.dumps(cell_station_hours, ensure_ascii=False),
            "Hours by Cell/Station - by person": json.dumps(hours_by_cell_by_person, ensure_ascii=False),
            "Output by Cell/Station - by person": json.dumps(output_by_cell_by_person, ensure_ascii=False),
            "UPLH by Cell/Station - by person": json.dumps(uplh_by_cell_by_person, ensure_ascii=False),
            "Open Complaint Timeliness": "",
            "error": " | ".join(errs) if errs else "",
            "Closures": "",
            "Opened": "",
        })
    return rows_out
def filter_rows_on_or_after(rows: list[dict], cutoff_iso: str) -> list[dict]:
    return [r for r in rows if safe_str(r.get("period_date")) >= cutoff_iso]
def _looks_like_iso_date(s: str) -> bool:
    s = safe_str(s)
    return (len(s) == 10 and s[4] == "-" and s[7] == "-")
def _is_monday_iso(s: str) -> bool:
    s = safe_str(s)
    if not _looks_like_iso_date(s):
        return False
    try:
        d = date.fromisoformat(s)
        return d.weekday() == 0
    except Exception:
        return False
def read_csv_rows(path: str) -> list[dict]:
    if not os.path.exists(path):
        return []
    with open(path, "r", newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))
def append_missing_placeholders_from_wip(
    wip_csv_path: str,
    closures_csv_path: str,
    timeliness_csv_path: str,
    logger: Optional[logging.Logger] = None,
) -> None:
    wip_rows = read_csv_rows(wip_csv_path)
    wip_keys: set[tuple[str, str]] = set()
    for r in wip_rows:
        team = safe_str(r.get("team"))
        pd = safe_str(r.get("period_date"))
        if not team or not pd:
            continue
        if _is_monday_iso(pd):
            wip_keys.add((team, pd))
    closures_rows = read_csv_rows(closures_csv_path)
    timeliness_rows = read_csv_rows(timeliness_csv_path)
    closures_keys = {(safe_str(r.get("team")), safe_str(r.get("period_date"))) for r in closures_rows if safe_str(r.get("team")) and safe_str(r.get("period_date"))}
    timeliness_keys = {(safe_str(r.get("team")), safe_str(r.get("period_date"))) for r in timeliness_rows if safe_str(r.get("team")) and safe_str(r.get("period_date"))}
    missing_closures = sorted([k for k in wip_keys if k not in closures_keys])
    missing_timeliness = sorted([k for k in wip_keys if k not in timeliness_keys])
    if missing_closures:
        file_exists = os.path.exists(closures_csv_path)
        with open(closures_csv_path, "a", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=["team", "period_date", "Closures", "Opened"])
            if (not file_exists) or (os.path.getsize(closures_csv_path) == 0):
                w.writeheader()
            for team, pd in missing_closures:
                w.writerow({"team": team, "period_date": pd, "Closures": "", "Opened": ""})
    if missing_timeliness:
        file_exists = os.path.exists(timeliness_csv_path)
        with open(timeliness_csv_path, "a", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=["team", "period_date", "Open Complaint Timeliness"])
            if (not file_exists) or (os.path.getsize(timeliness_csv_path) == 0):
                w.writeheader()
            for team, pd in missing_timeliness:
                w.writerow({"team": team, "period_date": pd, "Open Complaint Timeliness": ""})
    if logger:
        logger.info(f"[POST] WIP weekly keys={len(wip_keys)}")
        logger.info(f"[POST] closures placeholders appended={len(missing_closures)} -> {os.path.basename(closures_csv_path)}")
        logger.info(f"[POST] timeliness placeholders appended={len(missing_timeliness)} -> {os.path.basename(timeliness_csv_path)}")
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--team", default="all", help="Team to run (or 'all'). Example: --team PH")
    parser.add_argument("--log", default="NS_metrics.log", help="Log file path")
    args = parser.parse_args()
    logger = setup_logging(args.log)
    logger.info("=== NS Metrics Run START ===")
    logger.info(f"Selected team: {args.team}")
    ph_source_file = r"C:\Users\wadec8\Medtronic PLC\Customer Quality Pelvic Health - Daily Tracker\PH Cell Heijunka.xlsx"
    meic_source_file = r"C:\Users\wadec8\Medtronic PLC\Customer Quality Pelvic Health - Daily Tracker\MEIC\New MEIC PH Heijunka.xlsx"
    scs_source_file = r"C:\Users\wadec8\Medtronic PLC\Customer Quality SCS - Cell 17\Cell 1 - Heijunka.xlsx"
    scs_super_source_file = r"C:\Users\wadec8\Medtronic PLC\Customer Quality SCS - SCS Super Cell\Super Cell Heijunka.xlsx"
    cos_source_file = r"C:\Users\wadec8\Medtronic PLC\COS Cell - Documents\Heijunka v1.xlsx"
    nv_source_file = r"C:\Users\wadec8\Medtronic PLC\RTG Customer Quality Neurovascular - Documents\Cell\NV_Heijunka.xlsm"
    dbs_c13_source_file = r"C:\Users\wadec8\Medtronic PLC\DBS CQ Team - Documents\Heijunka_C13.xlsm"
    dbs_c14_source_file = r"C:\Users\wadec8\Medtronic PLC\DBS CQ Team - Documents\Heijunka_C14.xlsm"
    nav_source_file = r"C:\Users\wadec8\Medtronic PLC\MNAV Sharepoint - Navigation Work Reports\Heijunka_MNAV_Ranges_May2025.xlsm"
    ae_meic_source_file = r"C:\Users\wadec8\Medtronic PLC\MNAV Sharepoint - MEIC AE + OARM\AE_MEIC_Heijunka.xlsm"
    oarm_meic_source_file = r"C:\Users\wadec8\Medtronic PLC\MNAV Sharepoint - MEIC AE + OARM\OARM_MEIC_Heijunka.xlsm"
    mazor_source_file = r"C:\Users\wadec8\Medtronic PLC\MNAV Sharepoint - Caesarea Team\CAE - Heijunka_v2.xlsm"
    csf_source_file   = r"C:\Users\wadec8\Medtronic PLC\CQ CSF Management - Documents\CSF_Heijunka.xlsm"
    pss_source_file   = r"C:\Users\wadec8\Medtronic PLC\PSS Sharepoint - Documents\PSS_Heijunka.xlsm"
    ent_mapping_xlsx = r"C:\Users\wadec8\Medtronic PLC\ENT GEMBA Board - Heijunka 2.0 Files\Team & Tenure Mapping.xlsx"
    ent_data_csv     = r"C:\Users\wadec8\OneDrive - Medtronic PLC\ENT\ENT_Data.csv"
    ph_cell17_source_file = r"C:\Users\wadec8\Medtronic PLC\Customer Quality Pelvic Health - Cell 17\Cell 17 Heijunka.xlsx"
    out_file = "NS_metrics.csv"
    if not os.path.exists(ph_source_file):
        raise FileNotFoundError(f"Input file not found: {ph_source_file}")
    if not os.path.exists(meic_source_file):
        raise FileNotFoundError(f"Input file not found: {meic_source_file}")
    if not os.path.exists(scs_source_file):
        raise FileNotFoundError(f"Input file not found: {scs_source_file}")
    MAZOR_CFG = {
        "person_cols": ("B", "J"),
        "cells": {
            "total_available_hours": "R64",
            "completed_hours": "R54",
            "wp1_target": "X10",
            "wp2_target": "Z10",
            "wp1_output": "X5",
            "wp2_output": "Z5",
            "uplh_wp1": "X8",
            "uplh_wp2": "Z8",
            "wp1_hours": "X7",
            "wp2_hours": "Z7",
        },
        "rows": {
            "person_name_row_for_person_hours": 33,
            "person_actual_row_for_person_hours": 54,
            "person_available_row_for_person_hours": 64,
            "person_name_row_for_outputs_by_person": 13,
            "person_target_row_for_outputs_by_person": 28,
            "person_name_row_for_hours_by_cell_by_person": 33,
            "wp1_hour_rows": [34, 38, 42, 46, 50],
            "wp2_hour_rows": [35, 39, 43, 47, 51],
            "person_name_row_for_output_by_cell_by_person": 13,
            "wp1_output_rows_by_person": [14, 17, 20, 23, 26],
            "wp2_output_rows_by_person": [15, 18, 21, 24, 27],
        },
    }
    CSF_CFG = {
        "person_cols": ("B", "G"),
        "cells": {
            "total_available_hours": "I69",
            "completed_hours": "I69",  # per your spec
            "wp1_target": "N10",
            "wp2_target": "P10",
            "wp1_output": "N5",
            "wp2_output": "P5",
            "uplh_wp1": "N8",
            "uplh_wp2": "P8",
            "wp1_hours": "N7",
            "wp2_hours": "P7",
        },
        "rows": {
            "person_name_row_for_person_hours": 33,
            "person_actual_row_for_person_hours": 59,
            "person_available_row_for_person_hours": 69,
            "person_name_row_for_outputs_by_person": 13,
            "person_target_row_for_outputs_by_person": 28,
            "person_name_row_for_hours_by_cell_by_person": 33,
            "wp1_hour_rows": [34, 39, 44, 49, 54],
            "wp2_hour_rows": [35, 40, 45, 50, 55],
            "person_name_row_for_output_by_cell_by_person": 13,
            "wp1_output_rows_by_person": [14, 17, 20, 23, 26],
            "wp2_output_rows_by_person": [15, 18, 21, 24, 27],
        },
    }
    PSS_CFG = {
        "person_cols": ("B", "T"),
        "cells": {
            "total_available_hours": "W64",
            "completed_hours": "W54",
            "wp1_target": "AD10",
            "wp2_target": "AF10",
            "wp1_output": "AD5",
            "wp2_output": "AF5",
            "uplh_wp1": "AD8",
            "uplh_wp2": "AF8",
            "wp1_hours": "AD7",
            "wp2_hours": "AF7",
        },
        "rows": {
            "person_name_row_for_person_hours": 33,
            "person_actual_row_for_person_hours": 54,
            "person_available_row_for_person_hours": 64,
            "person_name_row_for_outputs_by_person": 13,
            "person_target_row_for_outputs_by_person": 28,
            "person_name_row_for_hours_by_cell_by_person": 33,
            "wp1_hour_rows": [34, 38, 42, 46, 50],
            "wp2_hour_rows": [35, 39, 43, 47, 51],
            "person_name_row_for_output_by_cell_by_person": 13,
            "wp1_output_rows_by_person": [14, 17, 20, 23, 26],
            "wp2_output_rows_by_person": [15, 18, 21, 24, 27],
        },
    }
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
    PH_CELL17_CFG = {
        "team": "PH Cell 17",
        "person_cols": ("B", "J"),
        "cells": {
            # Totals
            "total_available_hours": "L59",  # Total Available Hours
            "completed_hours": "L50",        # Completed Hours
            "wp1_output": "R2",
            "wp1_target": "R7",
            "wp2_output": "T2",
            "wp2_target": "T7",
            "uplh_wp1": "R5",
            "uplh_wp2": "T5",
            "wp1_hours": "R4",
            "wp2_hours": "T4",
        },
        "rows": {
            "hc_row": 25,
            "person_name_row_for_person_hours": 30,
            "person_actual_row_for_person_hours": 50,
            "person_available_row_for_person_hours": 59,
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
    TDD_COS1_CFG = {
        "team": "TDD COS 1",
        "person_cols": ("B", "P"),
        "date_parser": parse_sheet_date_scs_missing_year,  # missing-year Monday logic (handles hidden tabs too)
        "cells": {
            "total_available_hours": "R59",
            "completed_hours": "Q50",
            "wp1_output": "X2",
            "wp1_target": "X7",
            "wp2_output": "Z2",
            "wp2_target": "Z7",
            "uplh_wp1": "X5",
            "uplh_wp2": "Z5",
            "wp1_hours": "X4",
            "wp2_hours": "Z4",
        },
        "rows": {
            "hc_row": 50,  # count non-zero in row 50, B..P
            "person_name_row_for_person_hours": 30,
            "person_actual_row_for_person_hours": 50,
            "person_available_row_for_person_hours": 59,
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
    rows: list[dict] = []
    def extend_team(team_name: str, fn):
        out = run_team(logger, team_name, fn)   # logs START/DONE/FAIL + rows + elapsed
        rows.extend(out)
        return out
    def mondays_since(start_iso: str, end_d: date) -> list[str]:
        start = date.fromisoformat(start_iso)
        start = start - timedelta(days=start.weekday())  # ensure Monday
        out: list[str] = []
        d = start
        while d <= end_d:
            out.append(d.isoformat())
            d += timedelta(days=7)
        return out
    ALL_MONDAYS_SINCE_2025_06_02 = mondays_since("2025-06-02", date.today())
    extend_team("PH", lambda: scrape_workbook_with_config(ph_source_file, PH_CFG))
    extend_team("PH Cell 17", lambda: scrape_workbook_with_config(ph_cell17_source_file, PH_CELL17_CFG))
    extend_team("SCS Cell 1", lambda: scrape_workbook_with_config(scs_source_file, SCS_CELL1_CFG))
    meic_rows = run_team(logger, "MEIC PH", lambda: scrape_workbook_with_config(meic_source_file, MEIC_PH_CFG))
    cutoff_dbs = "2025-07-07"
    dbs_c13_rows = run_team(
        logger,
        "DBS C13",
        lambda: scrape_dbs_previous_weeks_xlsm(dbs_c13_source_file, "DBS C13", ALL_MONDAYS_SINCE_2025_06_02),
    )
    before = len(dbs_c13_rows)
    dbs_c13_rows = filter_rows_on_or_after(dbs_c13_rows, cutoff_dbs)
    logger.info(f"[DBS C13] filter >= {cutoff_dbs}: {before} -> {len(dbs_c13_rows)}")
    rows.extend(dbs_c13_rows)

    dbs_c14_rows = run_team(
        logger,
        "DBS C14",
        lambda: scrape_dbs_previous_weeks_xlsm(dbs_c14_source_file, "DBS C14", ALL_MONDAYS_SINCE_2025_06_02),
    )
    before = len(dbs_c14_rows)
    dbs_c14_rows = filter_rows_on_or_after(dbs_c14_rows, cutoff_dbs)
    logger.info(f"[DBS C14] filter >= {cutoff_dbs}: {before} -> {len(dbs_c14_rows)}")
    rows.extend(dbs_c14_rows)
    nv_rows = run_team(
        logger,
        "NV",
        lambda: scrape_dbs_previous_weeks_xlsm(nv_source_file, "NV", ALL_MONDAYS_SINCE_2025_06_02),
    )
    before = len(nv_rows)
    nv_rows = filter_rows_on_or_after(nv_rows, cutoff_dbs)
    logger.info(f"[NV] filter >= {cutoff_dbs}: {before} -> {len(nv_rows)}")
    rows.extend(nv_rows)
    extend_team("Nav", lambda: scrape_nav_previous_weeks_xlsm(nav_source_file, "Nav", ALL_MONDAYS_SINCE_2025_06_02))
    extend_team("AE MEIC", lambda: scrape_meic_ae_oarm_previous_weeks_xlsm(ae_meic_source_file, "AE MEIC", ALL_MONDAYS_SINCE_2025_06_02))
    extend_team("O-Arm MEIC", lambda: scrape_meic_ae_oarm_previous_weeks_xlsm(oarm_meic_source_file, "O-Arm MEIC", ALL_MONDAYS_SINCE_2025_06_02))
    extend_team("Mazor", lambda: scrape_previous_weeks_xlsm_with_filters(mazor_source_file, "Mazor", MAZOR_CFG, ALL_MONDAYS_SINCE_2025_06_02))
    extend_team("CSF",   lambda: scrape_previous_weeks_xlsm_with_filters(csf_source_file,   "CSF",   CSF_CFG,   ALL_MONDAYS_SINCE_2025_06_02))
    extend_team("PSS",   lambda: scrape_previous_weeks_xlsm_with_filters(pss_source_file,   "PSS",   PSS_CFG,   ALL_MONDAYS_SINCE_2025_06_02))
    extend_team("ENT",   lambda: scrape_ent_from_csv(ent_data_csv, ent_mapping_xlsx, team="ENT"))
    cos_rows = run_team(logger, "TDD COS 1", lambda: scrape_workbook_with_config(cos_source_file, TDD_COS1_CFG))
    cutoff_cos = date.fromisoformat("2025-06-02")
    before = len(cos_rows)
    cos_rows = [r for r in cos_rows if safe_str(r.get("period_date")) >= cutoff_cos.isoformat()]
    logger.info(f"[TDD COS 1] filter >= {cutoff_cos.isoformat()}: {before} -> {len(cos_rows)}")
    rows.extend(cos_rows)
    scs_super_rows = run_team(logger, "SCS Super Cell", lambda: scrape_workbook_with_config(scs_super_source_file, SCS_SUPER_CFG))
    logger.info(f"[SCS Super Cell] rows scraped (pre-filter): {len(scs_super_rows)}")
    cutoff_super = date.fromisoformat("2025-06-30")
    before = len(scs_super_rows)
    scs_super_rows = [r for r in scs_super_rows if safe_str(r.get("period_date")) >= cutoff_super.isoformat()]
    logger.info(f"[SCS Super Cell] filter >= {cutoff_super.isoformat()}: {before} -> {len(scs_super_rows)}")
    rows.extend(scs_super_rows)
    before = len(meic_rows)
    meic_rows = [r for r in meic_rows if safe_str(r.get("period_date")) >= "2025-09-01"]
    logger.info(f"[MEIC PH] filter >= 2025-09-01: {before} -> {len(meic_rows)}")
    rows.extend(meic_rows)
    before = len(rows)
    rows = [
        r for r in rows
        if (r.get("team") == "SCS Super Cell") or (safe_float(r.get("Total Available Hours")) != 0.0)
    ]
    logger.info(f"[ALL] filter TAA!=0 (except SCS Super Cell): {before} -> {len(rows)}")
    for bad in ("2023-11-06", "2026-09-07"):
        before = len(rows)
        rows = [r for r in rows if safe_str(r.get("period_date")) != bad]
        logger.info(f"[ALL] drop period_date == {bad}: {before} -> {len(rows)}")
    def sort_key(r: dict) -> tuple:
        team = safe_str(r.get("team")).lower()
        d = safe_str(r.get("period_date"))
        date_key = d if (len(d) == 10 and d[4] == "-" and d[7] == "-") else "9999-12-31"
        return (team, date_key)
    rows.sort(key=sort_key)
    write_csv(rows, out_file)
    logger.info(f"Wrote {len(rows)} rows to {out_file}")
    wip_rows = build_ns_wip_rows(rows)
    wip_out_file = "NS_WIP.csv"
    write_csv_wip(wip_rows, wip_out_file)
    logger.info(f"Wrote {len(wip_rows)} rows to {wip_out_file}")
    excel_dir = os.path.dirname(os.path.abspath(wip_out_file))
    timeliness_path = os.path.join(excel_dir, "timeliness.csv")
    closures_path = os.path.join(excel_dir, "closures.csv")
    append_missing_placeholders_from_wip(
        wip_csv_path=wip_out_file,
        closures_csv_path=closures_path,
        timeliness_csv_path=timeliness_path,
        logger=logger,
    )
    def apply_closures_timeliness_to_wip(
        wip_csv_path: str,
        closures_csv_path: str,
        timeliness_csv_path: str,
        logger: Optional[logging.Logger] = None,
    ) -> None:
        closures_rows = read_csv_rows(closures_csv_path)
        timeliness_rows = read_csv_rows(timeliness_csv_path)
        closures_lu: Dict[Tuple[str, str], Dict[str, Any]] = {}
        for r in closures_rows:
            team = safe_str(r.get("team"))
            pd = safe_str(r.get("period_date"))
            if team and pd:
                closures_lu[(team, pd)] = r
        timeliness_lu: Dict[Tuple[str, str], Dict[str, Any]] = {}
        for r in timeliness_rows:
            team = safe_str(r.get("team"))
            pd = safe_str(r.get("period_date"))
            if team and pd:
                timeliness_lu[(team, pd)] = r
        wip_rows = read_csv_rows(wip_csv_path)
        updated = 0
        for r in wip_rows:
            team = safe_str(r.get("team"))
            pd = safe_str(r.get("period_date"))
            if not team or not pd:
                continue
            c = closures_lu.get((team, pd))
            t = timeliness_lu.get((team, pd))
            if c is not None:
                r["Closures"] = safe_str(c.get("Closures"))
                r["Opened"] = safe_str(c.get("Opened"))
            if t is not None:
                r["Open Complaint Timeliness"] = safe_str(t.get("Open Complaint Timeliness"))
            if (safe_str(r.get("Closures")) or safe_str(r.get("Opened")) or safe_str(r.get("Open Complaint Timeliness"))):
                updated += 1
        with open(wip_csv_path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=WIP_HEADERS)
            w.writeheader()
            for r in wip_rows:
                w.writerow({h: r.get(h, "") for h in WIP_HEADERS})
        if logger:
            logger.info(f"[POST] NS_WIP.csv updated with closures/timeliness for {updated} rows")
    apply_closures_timeliness_to_wip(
        wip_csv_path=wip_out_file,
        closures_csv_path=closures_path,
        timeliness_csv_path=timeliness_path,
        logger=logger,
    )
    logger.info("=== NS Metrics Run END ===")
if __name__ == "__main__":
    main()