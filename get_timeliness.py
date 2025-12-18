import os
import re
import time
from datetime import datetime, date, timedelta
import pandas as pd
import win32com.client as win32
WORKBOOK_PATH = r"C:\Users\wadec8\OneDrive - Medtronic PLC\DSA-MDT-RPT-W-Go Green Initiative Monitor.xlsx"
SHEET_NAME = "Open Complaint Timeliness"
METRIC_NAME = "Open Complaint Timeliness"
TEAM_COLUMN_NAME = "Product Group"
TEAM_MAP = {
    "Coronary":"CRDN"
}
MONTHREL_ALLOWED = {0, -1}
OUT_DIR = r"C:\Users\wadec8\OneDrive - Medtronic PLC"
OUT_BASENAME = "open_complaint_timeliness_long"
INCLUDE_TOTALS = False
REFRESH_TIMEOUT_SECONDS = 1200
ROW_HIERARCHY_FIELD_NAME = "Operating Groups"
COL_HIERARCHY_FIELD_NAME = "Calendar"
TIMELINESS_CSV_PATH = r"C:\heijunka-dev\timeliness.csv"
AVERAGE_MERGE_GROUPS = {
    "PVH": {"EndoVenous", "Peripheral"}
}
MERGE_LOOKUP = {member: target
                for target, members in AVERAGE_MERGE_GROUPS.items()
                for member in members}
EXCLUDE_TEAMS = {
    "AAS",
    "AE",
    "Affera",
    "Airways",
    "Boulder-AST",
    "Boulder-GST",
    "Boxborough",
    "Brady",
    "CSF",
    "Caesarea",
    "DBS",
    "DISPOSABLES",
    "Diabetes",
    "Diagnostics",
    "DigitalTec",
    "ENT",
    "EPG",
    "Enterra",
    "GIS",
    "HF-MCS",
    "Heart Failure HP",
    "Heart Failure LP",
    "ILS-AST",
    "INPEN",
    "Implantable Tibial",
    "Infusion",
    "InterStim",
    "Intersect",
    "Kanghui",
    "Kyphon",
    "LSV",
    "MITG - Default",
    "MITG - Unmapped",
    "MOBILEAPP",
    "Medicrea",
    "NGP",
    "NV",
    "North Haven-AST",
    "North Haven-GST",
    "PARADIGM",
    "PCS-GWY",
    "PCS-MEIC",
    "PM-Apps",
    "PM-CTS",
    "PSS",
    "PTNM",
    "Pain Stim",
    "Patient Management",
    "Patient Monitoring",
    "RESERVOIR",
    "RF Ablation",
    "Renal Care Solutions",
    "SENSOR",
    "SOFTWARE",
    "Spine",
    "Surgical Robotics",
    "TRANSMITTER",
    "TYRX",
    "Tachy",
    "Ventilation",
    "Vitatron",
    "â€‹Undetermined",
    "Undetermined"
}
def excel_serial_to_date(n: float) -> date:
    return (datetime(1899, 12, 30) + timedelta(days=float(n))).date()
def week_monday(d: date) -> date:
    return d - timedelta(days=d.weekday())
def team_is_excluded(team_raw: str) -> bool:
    t = team_raw.strip()
    if t in EXCLUDE_TEAMS:
        return True
    return False
def parse_period_header(x):
    if x is None:
        return None
    if isinstance(x, datetime):
        return x.date()
    if isinstance(x, (int, float)) and x > 20000:
        try:
            return excel_serial_to_date(x)
        except Exception:
            return None
    s = str(x).strip()
    if not s:
        return None
    if re.fullmatch(r"\d{2}/\d{2}/\d{2}", s):
        try:
            return datetime.strptime(s, "%y/%m/%d").date()
        except Exception:
            return None
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            pass
    return None
def range_to_df(rng):
    vals = rng.Value
    if vals is None:
        return pd.DataFrame()
    if not isinstance(vals, (tuple, list)):
        return pd.DataFrame([[vals]])
    if vals and not isinstance(vals[0], (tuple, list)):
        return pd.DataFrame([list(vals)])
    return pd.DataFrame([list(r) for r in vals])
def is_total_like(s: str) -> bool:
    t = s.strip().lower()
    if t == "total":
        return True
    if " total" in t or t.endswith(" total"):
        return True
    if t in ("enterprise", "no portfolio"):
        return True
    return False
def connection_refreshing(conn) -> bool:
    try:
        if hasattr(conn, "OLEDBConnection") and conn.OLEDBConnection is not None:
            try:
                if conn.OLEDBConnection.Refreshing:
                    return True
            except Exception:
                pass
        if hasattr(conn, "ODBCConnection") and conn.ODBCConnection is not None:
            try:
                if conn.ODBCConnection.Refreshing:
                    return True
            except Exception:
                pass
    except Exception:
        pass
    return False
def wait_for_refresh(excel, wb, timeout_seconds=900, poll=2):
    wb.RefreshAll()
    try:
        excel.CalculateUntilAsyncQueriesDone()
    except Exception:
        pass
    t0 = time.time()
    while True:
        try:
            calc_busy = (excel.CalculationState != 0)
        except Exception:
            calc_busy = False
        conns_busy = False
        try:
            for c in wb.Connections:
                if connection_refreshing(c):
                    conns_busy = True
                    break
        except Exception:
            pass
        qt_busy = False
        try:
            for ws in wb.Worksheets:
                for qt in ws.QueryTables:
                    try:
                        if qt.Refreshing:
                            qt_busy = True
                            break
                    except Exception:
                        pass
                if qt_busy:
                    break
        except Exception:
            pass
        if not (calc_busy or conns_busy or qt_busy):
            return
        if time.time() - t0 > timeout_seconds:
            raise TimeoutError("Timed out waiting for Excel refresh to finish.")
        time.sleep(poll)
def find_best_pivot(ws, metric_name: str):
    if ws.PivotTables().Count == 1:
        return ws.PivotTables(1)
    for i in range(1, ws.PivotTables().Count + 1):
        pt = ws.PivotTables(i)
        try:
            df = range_to_df(pt.TableRange2).head(25).astype(str)
            if df.apply(lambda c: c.str.contains(metric_name, case=False, na=False)).any().any():
                return pt
        except Exception:
            continue
    return ws.PivotTables(1)
def clear_all_pivot_filters(pt):
    try:
        pt.ClearAllFilters()
    except Exception:
        pass
    try:
        for pf in pt.PivotFields():
            try:
                pf.ClearAllFilters()
            except Exception:
                pass
    except Exception:
        pass
def set_fmonthrel_strict(pt, allowed_values):
    field_mdx = "[Calendar].[fMonthRel].[fMonthRel]"
    pf = pt.PivotFields(field_mdx)
    try:
        pf.EnableMultiplePageItems = True
    except Exception:
        pass
    try:
        pf.ClearAllFilters()
    except Exception:
        pass
    allowed_values = sorted(list(allowed_values))
    candidates = [
        [f"[Calendar].[fMonthRel].&[{v}]" for v in allowed_values],
        [f"[Calendar].[fMonthRel].[fMonthRel].&[{v}]" for v in allowed_values],
    ]
    last_err = None
    for vil in candidates:
        try:
            pf.VisibleItemsList = vil
            return
        except Exception as e:
            last_err = e
    raise RuntimeError(f"Failed to set fMonthRel VisibleItemsList; last error: {last_err}")
def _set_tabular_and_repeat_labels(pt):
    try:
        pt.RowAxisLayout(1)
    except Exception:
        pass
    try:
        pt.RepeatAllLabels(2)
    except Exception:
        pass
    try:
        pt.ShowDrillIndicators = True
    except Exception:
        pass
def _drill_field_all_items(pf):
    try:
        pf.ShowAllItems = True
    except Exception:
        pass
    try:
        for pi in pf.PivotItems():
            try:
                pi.ShowDetail = True
            except Exception:
                try:
                    pi.DrilledDown = True
                except Exception:
                    pass
    except Exception:
        pass
def _pivot_has_weekly_columns(pt) -> bool:
    try:
        df = range_to_df(pt.TableRange2).head(20)
    except Exception:
        return False
    hits = 0
    for v in df.values.flatten().tolist():
        if parse_period_header(v) is not None:
            hits += 1
            if hits >= 2:
                return True
    return False
def force_expand_to_week_view_and_leaf(pt, max_passes=8):
    try:
        pt.ManualUpdate = True
    except Exception:
        pass
    _set_tabular_and_repeat_labels(pt)
    row_pf = None
    col_pf = None
    try:
        row_pf = pt.PivotFields(ROW_HIERARCHY_FIELD_NAME)
    except Exception:
        pass
    try:
        col_pf = pt.PivotFields(COL_HIERARCHY_FIELD_NAME)
    except Exception:
        pass
    for p in range(1, max_passes + 1):
        if _pivot_has_weekly_columns(pt):
            break
        if row_pf is not None:
            _drill_field_all_items(row_pf)
        if col_pf is not None:
            _drill_field_all_items(col_pf)
        try:
            for pf in pt.PivotFields():
                try:
                    orient = int(pf.Orientation)
                except Exception:
                    continue
                if orient in (1, 2):
                    _drill_field_all_items(pf)
        except Exception:
            pass
        try:
            pt.RefreshTable()
        except Exception:
            pass
        time.sleep(0.35)
        try:
            pt.RefreshTable()
        except Exception:
            pass
    try:
        pt.ManualUpdate = False
    except Exception:
        pass
def _normalize_period_date_series(s: pd.Series) -> pd.Series:
    s = s.astype(str).str.strip().str.lstrip("'")
    dt = pd.to_datetime(s, errors="coerce", infer_datetime_format=True)
    return dt.dt.strftime("%Y-%m-%d").fillna(s)
def update_timeliness_csv(path: str, updates_df: pd.DataFrame) -> tuple[int, int]:
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    existing = pd.read_csv(path, dtype={"team": str, "period_date": str})
    if "team" not in existing.columns or "period_date" not in existing.columns:
        raise RuntimeError(f"{path} must have columns: team, period_date, {METRIC_NAME}")
    if METRIC_NAME not in existing.columns:
        existing[METRIC_NAME] = pd.NA
    existing["team"] = existing["team"].astype(str).str.strip()
    existing["period_date"] = _normalize_period_date_series(existing["period_date"])
    upd = updates_df.copy()
    upd["team"] = upd["team"].astype(str).str.strip()
    upd["period_date"] = _normalize_period_date_series(upd["period_date"])
    existing_dt = pd.to_datetime(existing["period_date"], errors="coerce")
    upd_dt = pd.to_datetime(upd["period_date"], errors="coerce")
    max_existing = existing_dt.max()
    new_weeks = sorted([d for d in upd_dt.dropna().unique()
                        if pd.isna(max_existing) or d > max_existing])
    all_teams = sorted(set(existing["team"].dropna().tolist()) | set(upd["team"].dropna().tolist()))
    existing_keys = set(zip(existing["team"], existing["period_date"]))
    rows_to_add = []
    for week_dt in new_weeks:
        week_str = pd.Timestamp(week_dt).strftime("%Y-%m-%d")
        for team in all_teams:
            key = (team, week_str)
            if key not in existing_keys:
                rows_to_add.append({"team": team, "period_date": week_str, METRIC_NAME: pd.NA})
                existing_keys.add(key)
    added_count = len(rows_to_add)
    if rows_to_add:
        existing = pd.concat([existing, pd.DataFrame(rows_to_add)], ignore_index=True)
    upd_map = upd.set_index(["team", "period_date"])[METRIC_NAME]
    ex_idx = existing.set_index(["team", "period_date"])
    common = ex_idx.index.intersection(upd_map.index)
    ex_idx.loc[common, METRIC_NAME] = upd_map.loc[common].values
    updated_count = len(common)
    ex_idx.reset_index().to_csv(path, index=False)
    return updated_count, added_count
def main():
    if not os.path.exists(WORKBOOK_PATH):
        raise FileNotFoundError(WORKBOOK_PATH)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = os.path.join(OUT_DIR, f"{OUT_BASENAME}_{ts}.csv")
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(
        WORKBOOK_PATH,
        UpdateLinks=3,
        ReadOnly=True,
        IgnoreReadOnlyRecommended=True
    )
    try:
        wait_for_refresh(excel, wb, timeout_seconds=REFRESH_TIMEOUT_SECONDS, poll=2)
        ws = wb.Worksheets(SHEET_NAME)
        pt = find_best_pivot(ws, METRIC_NAME)
        clear_all_pivot_filters(pt)
        set_fmonthrel_strict(pt, MONTHREL_ALLOWED)
        try:
            pt.RefreshTable()
        except Exception:
            pass
        time.sleep(0.5)
        force_expand_to_week_view_and_leaf(pt, max_passes=8)
        wide = range_to_df(pt.TableRange2)
        header_row_idx = None
        for i in range(len(wide)):
            row_vals = [str(x).strip().lower() for x in wide.iloc[i, :].tolist()]
            if "portfolio" in row_vals and "operating unit" in row_vals:
                header_row_idx = i
                break
        if header_row_idx is None:
            header_row_idx = 0
        headers = list(wide.iloc[header_row_idx].values)
        data = wide.iloc[header_row_idx + 1 :].copy()
        data.columns = headers
        date_cols = []
        date_map = {}
        for col in data.columns:
            d = parse_period_header(col)
            if d is not None:
                date_cols.append(col)
                date_map[col] = week_monday(d)
        if not date_cols:
            raise RuntimeError(
                "No weekly date columns found. "
                "This usually means the Calendar hierarchy didn't drill to week level."
            )
        if TEAM_COLUMN_NAME not in data.columns:
            raise RuntimeError(
                f"'{TEAM_COLUMN_NAME}' column not found in pivot output. "
                f"First columns seen: {list(data.columns)[:30]}"
            )
        records = []
        for _, row in data.iterrows():
            team_raw = row.get(TEAM_COLUMN_NAME, None)
            if team_raw is None or not str(team_raw).strip():
                continue
            team_raw = str(team_raw).strip()
            if not INCLUDE_TOTALS and is_total_like(team_raw):
                continue
            if team_is_excluded(team_raw):
                continue
            team = TEAM_MAP.get(team_raw, team_raw)
            for c in date_cols:
                v = row.get(c, None)
                if v is None or (isinstance(v, str) and not v.strip()):
                    continue
                try:
                    val = float(v)
                except Exception:
                    continue
                if 0 <= val <= 1.5:
                    val *= 100.0
                iso = date_map[c].strftime("%Y-%m-%d")
                records.append({
                    "team": team,
                    "period_date": iso,
                    METRIC_NAME: round(val, 1),
                })
        out_df = pd.DataFrame(records)
        if out_df.empty:
            raise RuntimeError("Export produced 0 rows (pivot likely still collapsed or filtered unexpectedly).")
        out_df["team"] = out_df["team"].map(lambda t: MERGE_LOOKUP.get(t, t))
        out_df = (
            out_df
            .groupby(["team", "period_date"], as_index=False)[METRIC_NAME]
            .mean()
        )
        out_df[METRIC_NAME] = out_df[METRIC_NAME].round(1)
        out_df = out_df.sort_values(["team", "period_date"])
        updated, added = update_timeliness_csv(TIMELINESS_CSV_PATH, out_df)
        print(f"Updated {updated} row(s), added {added} row(s) in {TIMELINESS_CSV_PATH}")
        out_df.to_csv(out_path, index=False)
        print(f"Wrote: {out_path} ({len(out_df)} rows)")
    finally:
        wb.Close(SaveChanges=False)
        excel.Quit()
if __name__ == "__main__":
    main()