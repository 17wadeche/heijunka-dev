import csv
import json
import os
import re
import shutil
import tempfile
import uuid
from datetime import datetime, date
from typing import Any, Dict, Optional, Tuple
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
import win32com.client
import time
import pythoncom
import pywintypes
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
def scrape_dbs_previous_weeks_xlsm(source_file: str, team: str) -> list[dict]:
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
        dropdown_values = _get_dropdown_values_from_validation(dd)
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
def filter_rows_on_or_after(rows: list[dict], cutoff_iso: str) -> list[dict]:
    return [r for r in rows if safe_str(r.get("period_date")) >= cutoff_iso]
def main():
    ph_source_file = r"C:\Users\wadec8\Medtronic PLC\Customer Quality Pelvic Health - Daily Tracker\PH Cell Heijunka.xlsx"
    meic_source_file = r"C:\Users\wadec8\Medtronic PLC\Customer Quality Pelvic Health - Daily Tracker\MEIC\New MEIC PH Heijunka.xlsx"
    scs_source_file = r"C:\Users\wadec8\Medtronic PLC\Customer Quality SCS - Cell 17\Cell 1 - Heijunka.xlsx"
    scs_super_source_file = r"C:\Users\wadec8\Medtronic PLC\Customer Quality SCS - SCS Super Cell\Super Cell Heijunka.xlsx"
    cos_source_file = r"C:\Users\wadec8\Medtronic PLC\COS Cell - Documents\Heijunka v1.xlsx"
    nv_source_file = r"C:\Users\wadec8\Medtronic PLC\RTG Customer Quality Neurovascular - Documents\Cell\NV_Heijunka.xlsm"
    dbs_c13_source_file = r"C:\Users\wadec8\Medtronic PLC\DBS CQ Team - Documents\Heijunka_C13.xlsm"
    dbs_c14_source_file = r"C:\Users\wadec8\Medtronic PLC\DBS CQ Team - Documents\Heijunka_C14.xlsm"
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
    rows = []
    rows.extend(scrape_workbook_with_config(ph_source_file, PH_CFG))
    rows.extend(scrape_workbook_with_config(scs_source_file, SCS_CELL1_CFG))
    meic_rows = scrape_workbook_with_config(meic_source_file, MEIC_PH_CFG)
    cutoff_dbs = "2025-07-07"
    dbs_c13_rows = scrape_dbs_previous_weeks_xlsm(dbs_c13_source_file, "DBS C13")
    dbs_c13_rows = filter_rows_on_or_after(dbs_c13_rows, cutoff_dbs)
    rows.extend(dbs_c13_rows)
    dbs_c14_rows = scrape_dbs_previous_weeks_xlsm(dbs_c14_source_file, "DBS C14")
    dbs_c14_rows = filter_rows_on_or_after(dbs_c14_rows, cutoff_dbs)
    rows.extend(dbs_c14_rows)
    nv_rows = scrape_dbs_previous_weeks_xlsm(nv_source_file, "NV")
    nv_rows = filter_rows_on_or_after(nv_rows, cutoff_dbs)
    rows.extend(nv_rows)
    cos_rows = scrape_workbook_with_config(cos_source_file, TDD_COS1_CFG)
    cutoff_cos = date.fromisoformat("2025-01-06")
    cos_rows = [
        r for r in cos_rows
        if safe_str(r.get("period_date")) >= cutoff_cos.isoformat()
    ]
    rows.extend(cos_rows)
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