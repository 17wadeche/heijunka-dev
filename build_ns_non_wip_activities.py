import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple
import numpy as np
import pandas as pd
import win32com.client as win32
from tempfile import mkdtemp
import shutil
import warnings
from pathlib import Path
print(f"RUNNING FILE: {Path(__file__).resolve()}", flush=True)
DBS_C13_SOURCE_FILE = Path(r"C:\Users\wadec8\Medtronic PLC\DBS CQ Team - Documents\Cell 13 Heijunka V2.xlsx")
DBS_C14_SOURCE_FILE = Path(r"C:\Users\wadec8\Medtronic PLC\DBS CQ Team - Documents\Cell 14 Heijunka V2.xlsx")
warnings.filterwarnings(
    "ignore",
    message="Data Validation extension is not supported and will be removed"
)
MEIC_TRACKER_PATH = Path(
    r"C:\Users\wadec8\Medtronic PLC\MEIC_NMPH - Documents\NPH Tracker.xlsx"
)
def excel_cell(row_i_zero_based: int, col_i_zero_based: int) -> str:
    n = col_i_zero_based + 1
    letters = ""
    while n:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return f"{letters}{row_i_zero_based + 1}"
DBS_MEIC_NAMES = {"Divya", "Reshmita", "Shankar"}
PH_MEIC_NAMES = {"Sathya", "Arun", "Kavya"}
TEAM_TRACKER_SHEET = "Team Tracker"
NS_WIP_PATH = Path(r"C:\heijunka-dev\NS_WIP.csv")
NS_METRICS_PATH = Path(r"C:\heijunka-dev\NS_metrics.csv")
OUT_PATH = Path(r"C:\heijunka-dev\ns_non_wip_activities.csv")
BAD_NAMES = {
    "", "-", "–", "—", "nan", "NaN", "NAN",
    "n/a", "N/A", "na", "NA", "null", "NULL",
    "none", "None", "tm", "TM", "Totals", "TOTALS",
    "Team Hours Available", "TEAM HOURS AVAILABLE",
    "Mazor Hours Available", "MAZOR HOURS AVAILABLE",
    "Team 1 Hours Available", "Team Member"
}
def get_dbs_people_count_from_heijunka_files(
    file_paths: tuple[Path, Path] = (DBS_C13_SOURCE_FILE, DBS_C14_SOURCE_FILE),
    name_row_zero_based: int = 29,   # Excel row 30
) -> int:
    bad = {"", "open", "total", "uplh"}
    unique_names: set[str] = set()
    names_by_file: dict[str, list[str]] = {}
    for fp in file_paths:
        if not fp.exists():
            print(f"[DEBUG][DBS] missing file: {fp}", flush=True)
            names_by_file[str(fp)] = []
            continue
        ws_df = pd.read_excel(fp, sheet_name=0, header=None)
        if ws_df.shape[0] <= name_row_zero_based:
            print(f"[DEBUG][DBS] {fp.name}: row 30 not available", flush=True)
            names_by_file[str(fp)] = []
            continue
        row_vals = ws_df.iloc[name_row_zero_based].tolist()
        file_names_found = []
        for raw in row_vals:
            name = norm_name(raw)
            if not name:
                continue
            if name.strip().lower() in bad:
                continue
            if not is_real_person(name):
                continue
            file_names_found.append(name)
            unique_names.add(name)
        names_by_file[str(fp)] = file_names_found
        print(f"[DEBUG][DBS] {fp.name} row 30 names: {file_names_found}", flush=True)
    if len(file_paths) >= 2:
        s1 = set(names_by_file.get(str(file_paths[0]), []))
        s2 = set(names_by_file.get(str(file_paths[1]), []))
        print(f"[DEBUG][DBS] overlap names: {sorted(s1 & s2)}", flush=True)
    print(f"[DEBUG][DBS] merged unique names counted: {sorted(unique_names)}", flush=True)
    print(f"[DEBUG][DBS] merged unique people_count: {len(unique_names)}", flush=True)
    return len(unique_names)
def norm_name(x) -> str:
    return " ".join(str(x or "").strip().split())
def is_real_person(name: str) -> bool:
    n = norm_name(name)
    if not n:
        return False
    if n.strip().lower() in {b.lower() for b in BAD_NAMES}:
        return False
    return True
def safe_float(x) -> float:
    if x is None:
        return np.nan
    try:
        if pd.isna(x):
            return np.nan
    except Exception:
        pass
    if isinstance(x, (int, float, np.integer, np.floating)):
        return float(x)
    s = str(x).strip()
    if not s or s in {"-", "–", "—"}:
        return np.nan
    s = s.replace(",", "").replace("\u00a0", " ")
    m = re.search(r"[-+]?\d*\.?\d+", s)
    if not m:
        return np.nan
    try:
        return float(m.group(0))
    except Exception:
        return np.nan
def safe_float0(x) -> float:
    v = safe_float(x)
    return 0.0 if pd.isna(v) else float(v)
from datetime import datetime, date
def _excel_date_to_timestamp(v) -> Optional[pd.Timestamp]:
    if v is None:
        return None
    if isinstance(v, pd.Timestamp):
        try:
            return pd.Timestamp(v).tz_localize(None).normalize()
        except Exception:
            try:
                return pd.Timestamp(v).normalize()
            except Exception:
                return None
    if isinstance(v, datetime):
        try:
            return pd.Timestamp(v.replace(tzinfo=None)).normalize()
        except Exception:
            return None
    if isinstance(v, date):
        try:
            return pd.Timestamp(v).normalize()
        except Exception:
            return None
    s = str(v).strip()
    if not s or s.lower() in {"nan", "nat", "none"}:
        return None
    if re.fullmatch(r"[-+]?\d+(\.0+)?", s):
        try:
            num = float(s)
            if -1000 <= num <= 1000:
                return None
        except Exception:
            pass
    try:
        dt = pd.to_datetime(s, errors="coerce")
    except Exception:
        return None
    if pd.isna(dt):
        return None
    try:
        year = int(dt.year)
        if year < 2024 or year > 2035:
            return None
    except Exception:
        return None
    return pd.Timestamp(dt).normalize()
def _resolve_validation_list_values(wb, ws, cell_addr: str = "B1") -> List[pd.Timestamp]:
    values: List[pd.Timestamp] = []
    def _add(v):
        dt = _excel_date_to_timestamp(v)
        if dt is not None:
            values.append(dt)
    formula = None
    try:
        cell = ws.Range(cell_addr)
        formula = cell.Validation.Formula1
    except Exception:
        formula = None
    if formula:
        formula = str(formula).strip()
        if formula.startswith("="):
            formula = formula[1:]
        if "," in formula and "!" not in formula and ":" not in formula:
            for part in formula.split(","):
                _add(part.strip())
            return _clean_candidate_dates(values)
        try:
            src_rng = wb.Application.Evaluate(formula)
            if hasattr(src_rng, "Rows") and hasattr(src_rng, "Columns"):
                max_rows = min(src_rng.Rows.Count, 200)
                max_cols = min(src_rng.Columns.Count, 10)
                for r in range(1, max_rows + 1):
                    for c in range(1, max_cols + 1):
                        _add(src_rng.Cells(r, c).Value)
        except Exception:
            pass
        if values:
            return _clean_candidate_dates(values)
    candidate_sheet_names = [
        "Instructions for Use", "Instructions", "Lists", "Lookup", "Lookups",
        "Config", "Settings", "Setup"
    ]
    for sheet_name in candidate_sheet_names:
        try:
            sh = wb.Worksheets(sheet_name)
        except Exception:
            continue
        try:
            used = sh.UsedRange.Value
            if used is None:
                continue
            if not isinstance(used, tuple):
                used = ((used,),)
            max_rows = min(len(used), 200)
            for r in range(max_rows):
                row = used[r]
                if not isinstance(row, tuple):
                    row = (row,)
                max_cols = min(len(row), 5)
                for c in range(max_cols):
                    _add(row[c])
        except Exception:
            pass
    if values:
        return _clean_candidate_dates(values)
    try:
        for nm in wb.Names:
            try:
                ref = nm.RefersTo
                if not ref:
                    continue
                if str(ref).startswith("="):
                    ref = str(ref)[1:]
                src_rng = wb.Application.Evaluate(ref)
                if src_rng is None:
                    continue
                if hasattr(src_rng, "Rows") and hasattr(src_rng, "Columns"):
                    for r in range(1, src_rng.Rows.Count + 1):
                        for c in range(1, src_rng.Columns.Count + 1):
                            _add(src_rng.Cells(r, c).Value)
                else:
                    _add(src_rng)
            except Exception:
                pass
    except Exception:
        pass
    if values:
        return _clean_candidate_dates(values)
    try:
        current_dt = _excel_date_to_timestamp(ws.Range(cell_addr).Value)
        if current_dt is not None:
            return _clean_candidate_dates([current_dt])
    except Exception:
        pass
    return []
def _clean_candidate_dates(values: List[pd.Timestamp]) -> List[pd.Timestamp]:
    out = []
    seen = set()
    for dt in values:
        if dt is None or pd.isna(dt):
            continue
        ts = pd.Timestamp(dt).normalize()
        if ts.year < 2024 or ts.year > 2035:
            continue
        if ts.dayofweek != 0:
            continue
        if ts not in seen:
            seen.add(ts)
            out.append(ts)
    return sorted(out)
def load_csv(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    df.columns = [" ".join(str(c).split()) for c in df.columns]
    if "period_date" in df.columns:
        df["period_date"] = pd.to_datetime(df["period_date"], errors="coerce").dt.normalize()
    return df
def load_metrics(ns_metrics_path: Path) -> pd.DataFrame:
    df = load_csv(ns_metrics_path)
    if "Completed Hours" in df.columns:
        df["Completed Hours"] = pd.to_numeric(df["Completed Hours"], errors="coerce")
    return df
def parse_person_hours_json(cell) -> dict:
    if cell is None or (isinstance(cell, float) and pd.isna(cell)):
        return {}
    if isinstance(cell, dict):
        return cell
    s = str(cell).strip()
    if not s:
        return {}
    try:
        obj = json.loads(s)
        return obj if isinstance(obj, dict) else {}
    except Exception:
        return {}
def extract_wip_workers_from_row(row: pd.Series, person_hours_col: str = "Person Hours") -> List[str]:
    blob = parse_person_hours_json(row.get(person_hours_col))
    workers: List[str] = []
    for k, v in blob.items():
        name = norm_name(k)
        if not is_real_person(name) or name == "0.0":
            continue
        actual = safe_float(v.get("actual")) if isinstance(v, dict) else safe_float(v)
        if pd.notna(actual) and actual > 0:
            workers.append(name)
    return sorted(set(workers))
def read_people_block(
    ws: pd.DataFrame,
    start_row_i: int,
    end_row_i: Optional[int] = None,
    *,
    team: Optional[str] = None,
    sheet_name: Optional[str] = None,
    week: Optional[pd.Timestamp] = None,
) -> List[dict]:
    rows: List[dict] = []
    last_i = len(ws) - 1 if end_row_i is None else min(end_row_i, len(ws) - 1)
    is_ph = (team == "PH")
    for i in range(start_row_i, last_i + 1):
        raw_name = ws.iat[i, 0] if ws.shape[1] > 0 else ""
        name = norm_name(raw_name)
        a_cell = excel_cell(i, 0)
        b_cell = excel_cell(i, 1)
        c_cell = excel_cell(i, 2)
        b_raw = ws.iat[i, 1] if ws.shape[1] > 1 else np.nan
        c_raw = ws.iat[i, 2] if ws.shape[1] > 2 else np.nan
        b = safe_float(b_raw)
        c = safe_float(c_raw)
        if pd.isna(b):
            b = 0.0
        if pd.isna(c):
            c = 0.0
        rows.append({"row_i": i, "name": name, "B": b, "C": c})
    return rows
def build_activities(ws: pd.DataFrame, people_rows: List[dict], header_row_i: int, start_col_i: int, end_col_i: int) -> List[dict]:
    activities: List[dict] = []
    end_col_i = min(end_col_i, ws.shape[1] - 1)
    for pr in people_rows:
        i = pr["row_i"]
        name = pr["name"]
        for c in range(start_col_i, end_col_i + 1):
            label = norm_name(ws.iat[header_row_i, c] if c < ws.shape[1] else "")
            if not label:
                continue
            hrs = safe_float(ws.iat[i, c] if c < ws.shape[1] else np.nan)
            if pd.isna(hrs) or hrs <= 0:
                continue
            activities.append({
                "name": name,
                "activity": label,
                "hours": float(round(hrs, 2)),
            })
        ooo = float(round(safe_float0(pr.get("C", 0.0)), 2))
        if ooo > 0:
            activities.append({
                "name": name,
                "activity": "OOO",
                "hours": ooo,
            })
    return activities
@dataclass(frozen=True)
class StandardLayout:
    people_start_row: int
    totals_row: int
    activity_header_row: int
    activity_start_col: int
    activity_end_col: int
    min_rows: int
    min_cols: int
@dataclass(frozen=True)
class TeamSource:
    team: str
    xlsx: Path
    layout: Optional[StandardLayout] = None
    week_from_sheet: Optional[Callable[[str, pd.DataFrame], Optional[pd.Timestamp]]] = None
    custom_builder: Optional[Callable[..., Dict]] = None
    wip_workers_from: str = "NS_WIP"
    completed_hours_from: str = "NS_WIP"
def week_from_sheetname_date(sheet_name: str, ws: pd.DataFrame) -> Optional[pd.Timestamp]:
    s = str(sheet_name).strip()
    dt = pd.to_datetime(s, errors="coerce")
    if pd.notna(dt):
        return dt.normalize()
    s2 = re.sub(r"^\s*week\s+of\s+", "", s, flags=re.IGNORECASE).strip()
    dt = pd.to_datetime(s2, errors="coerce")
    if pd.notna(dt):
        return dt.normalize()
    return None
DEFAULT_YEAR_IF_MISSING = 2026
def _is_real_year(dt: pd.Timestamp, min_year: int = 2000) -> bool:
    try:
        return pd.notna(dt) and int(dt.year) >= min_year
    except Exception:
        return False
def _get_matching_worksheet(wb, preferred_name: str):
    preferred = preferred_name.strip().lower()
    candidates = []
    for ws in wb.Worksheets:
        try:
            nm = str(ws.Name).strip()
        except Exception:
            continue
        nm_lower = nm.lower()
        if nm_lower == preferred:
            return ws
        if preferred in nm_lower:
            candidates.append(ws)
    if candidates:
        return candidates[0]
    available = []
    for ws in wb.Worksheets:
        try:
            available.append(str(ws.Name))
        except Exception:
            pass
    raise ValueError(
        f"Could not find worksheet matching '{preferred_name}'. "
        f"Available sheets: {available}"
    )
def build_selector_rows_from_capacity_workbook(
    team_src: TeamSource,
    wip_df: pd.DataFrame,
    metrics_df: pd.DataFrame,
    selector_cell: str = "A2",
    sheet_name: str = "Capacity mgmt",
) -> pd.DataFrame:
    xlsx_path = team_src.xlsx
    print(f"[DEBUG] ENTER build_selector_rows_from_capacity_workbook for team={team_src.team!r}", flush=True)
    print(f"[DEBUG] workbook path for {team_src.team!r}: {xlsx_path}", flush=True)
    if not xlsx_path.exists():
        print(f"[WARN] Missing XLSX for {team_src.team}: {xlsx_path}", flush=True)
        return pd.DataFrame()
    out_rows: List[dict] = []
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    temp_dir = None
    try:
        wb = excel.Workbooks.Open(
            str(xlsx_path),
            UpdateLinks=3,
            ReadOnly=True
        )
        ws_com = _get_matching_worksheet(wb, sheet_name)
        excel.CalculateFullRebuild()
        all_dates = _resolve_validation_list_values(wb, ws_com, selector_cell)
        if not all_dates:
            current_dt = _excel_date_to_timestamp(ws_com.Range(selector_cell).Value)
            if current_dt is not None:
                all_dates = [current_dt]
        for week in all_dates:
            try:
                selector_range = ws_com.Range(selector_cell)
                selector_range.Value = week.to_pydatetime()
                try:
                    selector_range.NumberFormat = "yyyy/mm/dd"
                except Exception:
                    pass
                try:
                    ws_com.Calculate()
                except Exception:
                    pass
                try:
                    wb.RefreshAll()
                except Exception:
                    pass
                try:
                    excel.CalculateUntilAsyncQueriesDone()
                except Exception:
                    pass
                excel.CalculateFullRebuild()
                print(f"[DEBUG] {team_src.team} -> refreshing/recalculating for {week.date()}")
                try:
                    wb.RefreshAll()
                except Exception:
                    pass
                try:
                    excel.CalculateUntilAsyncQueriesDone()
                except Exception:
                    pass
                excel.CalculateFullRebuild()
                used = ws_com.UsedRange.Value
                if used is None:
                    continue
                if not isinstance(used, tuple):
                    used = ((used,),)
                ws_df = pd.DataFrame(list(used))
                if team_src.custom_builder is None:
                    raise ValueError(f"No custom_builder configured for {team_src.team}")
                built = team_src.custom_builder(team_src.team, ws_df, week)
                print(
                    f"[DEBUG] {team_src.team} built summary for {week.date()} -> "
                    f"people_count={built['people_count']}, "
                    f"ooo_hours={built['ooo_hours']}, "
                    f"total_nonwip_hours={built['total_nonwip_hours']}, "
                    f"activities={len(built['nonwip_activities'])}"
                )
                people_count = built["people_count"]
                total_nonwip_hours = built["total_nonwip_hours"]
                ooo_hours = built["ooo_hours"]
                nonwip_by_person = built["nonwip_by_person"]
                nonwip_activities = built["nonwip_activities"]
                ooo_map = built["ooo_map"]
                completed_src_df = metrics_df if team_src.completed_hours_from == "NS_metrics" else wip_df
                completed_match = completed_src_df[
                    (completed_src_df.get("team") == team_src.team) &
                    (completed_src_df["period_date"] == week)
                ]
                completed_hours = (
                    pd.to_numeric(completed_match.iloc[0].get("Completed Hours"), errors="coerce")
                    if not completed_match.empty else np.nan
                )
                pct_in_wip = np.nan
                if pd.notna(completed_hours) and pd.notna(total_nonwip_hours):
                    denom = float(completed_hours) + float(total_nonwip_hours)
                    pct_in_wip = float(completed_hours) / denom if denom != 0 else np.nan
                wip_source_df = metrics_df if team_src.wip_workers_from == "NS_metrics" else wip_df
                wip_match = wip_source_df[
                    (wip_source_df.get("team") == team_src.team) &
                    (wip_source_df["period_date"] == week)
                ]
                wip_workers = extract_wip_workers_from_row(wip_match.iloc[0]) if not wip_match.empty else []
                wip_workers_count = len(wip_workers)
                wip_workers_ooo_hours = float(round(sum(safe_float0(ooo_map.get(n, 0.0)) for n in wip_workers), 2))
                print(f"[DEBUG] team_src.team = {team_src.team!r}")
                team_name = str(team_src.team).strip()
                if team_name == "ENT":
                    people_count_final = people_count
                else:
                    people_count_final = get_people_count_from_wip(
                        wip_df=wip_df,
                        team=team_src.team,
                        week=week,
                        fallback=people_count,
                    )
                out_rows.append({
                    "team": team_src.team,
                    "period_date": week.date().isoformat(),
                    "source_file": str(xlsx_path),
                    "people_count": int(people_count_final),
                    "total_non_wip_hours": float(round(total_nonwip_hours, 2)) if pd.notna(total_nonwip_hours) else np.nan,
                    "OOO Hours": float(round(ooo_hours, 2)) if pd.notna(ooo_hours) else np.nan,
                    "% in WIP": float(round(pct_in_wip, 6)) if pd.notna(pct_in_wip) else np.nan,
                    "non_wip_by_person": json.dumps(nonwip_by_person, ensure_ascii=False),
                    "non_wip_activities": json.dumps(nonwip_activities, ensure_ascii=False),
                    "wip_workers": json.dumps(wip_workers, ensure_ascii=False),
                    "wip_workers_count": int(wip_workers_count),
                    "wip_workers_ooo_hours": float(wip_workers_ooo_hours),
                })
            except Exception as e:
                print(f"[WARN] Failed {team_src.team} week {week}: {e}")
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        excel.Quit()
        if temp_dir and Path(temp_dir).exists():
            shutil.rmtree(temp_dir, ignore_errors=True)
    df = pd.DataFrame(out_rows)
    if not df.empty:
        df["period_date"] = pd.to_datetime(df["period_date"], errors="coerce").dt.normalize()
        df = df.drop_duplicates(subset=["team", "period_date"], keep="last")
        df = df.sort_values(["team", "period_date"]).reset_index(drop=True)
    return df
def week_from_pss_meic_tab(sheet_name: str, ws: pd.DataFrame) -> Optional[pd.Timestamp]:
    s = str(sheet_name).strip()
    s_lower = s.lower()
    if "capacity mgmt" not in s_lower:
        return None
    m = re.search(r"\((\d{1,2})[.\-_/](\d{1,2})\)", s)
    if not m:
        return None 
    mm = int(m.group(1))
    dd = int(m.group(2))
    for r in range(0, min(6, ws.shape[0])):
        for c in range(0, min(6, ws.shape[1])):
            try:
                v = ws.iat[r, c]
            except Exception:
                continue
            dt = pd.to_datetime(v, errors="coerce")
            if _is_real_year(dt):
                try:
                    return pd.Timestamp(year=int(dt.year), month=mm, day=dd).normalize()
                except Exception:
                    pass
    try:
        return pd.Timestamp(year=DEFAULT_YEAR_IF_MISSING, month=mm, day=dd).normalize()
    except Exception:
        return None
def build_pss_meic_dated_row(team: str, ws: pd.DataFrame, week: Optional[pd.Timestamp] = None) -> Dict:
    NAME_COL = _col_letter_to_idx("A")
    COL_EXPECTED_WIP = _col_letter_to_idx("B")
    ACT_START = _col_letter_to_idx("C")
    ACT_END = _col_letter_to_idx("W")
    COL_OOO = _col_letter_to_idx("X")
    HEADER_ROW = 0          # Excel row 1
    PEOPLE_START_ROW = 1    # Excel row 2
    def header_label_for_col(c: int) -> str:
        txt = norm_name(ws.iat[HEADER_ROW, c] if ws.shape[0] > HEADER_ROW and ws.shape[1] > c else "")
        return txt
    people_rows: List[dict] = []
    nonwip_by_person: Dict[str, float] = {}
    activities: List[dict] = []
    seen_people = False
    blank_run = 0
    for i in range(PEOPLE_START_ROW, ws.shape[0]):
        raw_name = ws.iat[i, NAME_COL] if ws.shape[1] > NAME_COL else ""
        name = norm_name(raw_name)
        if is_real_person(name):
            seen_people = True
            blank_run = 0
            expected_wip = safe_float0(ws.iat[i, COL_EXPECTED_WIP] if ws.shape[1] > COL_EXPECTED_WIP else 0.0)
            ooo = safe_float0(ws.iat[i, COL_OOO] if ws.shape[1] > COL_OOO else 0.0)
            people_rows.append({
                "row_i": i,
                "name": name,
                "B": float(expected_wip),
                "OOO": float(ooo),
            })
            person_nonwip_total = 0.0
            for c in range(ACT_START, min(ACT_END, ws.shape[1] - 1) + 1):
                label = header_label_for_col(c)
                if not label:
                    continue
                hrs = safe_float(ws.iat[i, c] if ws.shape[0] > i and ws.shape[1] > c else np.nan)
                if pd.isna(hrs) or hrs <= 0:
                    continue
                hrs = float(round(float(hrs), 2))
                activities.append({
                    "name": name,
                    "activity": label,
                    "hours": hrs,
                })
                person_nonwip_total += hrs
            if ooo > 0:
                activities.append({
                    "name": name,
                    "activity": "OOO",
                    "hours": float(round(ooo, 2)),
                })
            person_nonwip_total = float(round(person_nonwip_total, 2))
            if person_nonwip_total != 0.0:
                nonwip_by_person[name] = person_nonwip_total
        else:
            if seen_people:
                blank_run += 1
                if blank_run >= 3:
                    break
    people_count = len(set(r["name"] for r in people_rows))
    ooo_hours = float(round(sum(r["OOO"] for r in people_rows), 2))
    total_nonwip_hours = float(round(sum(a["hours"] for a in activities), 2))
    return {
        "people_rows": people_rows,
        "people_count": people_count,
        "ooo_hours": ooo_hours,
        "total_nonwip_hours": total_nonwip_hours,
        "nonwip_by_person": nonwip_by_person,
        "nonwip_activities": activities,
        "ooo_map": {r["name"]: float(r["OOO"]) for r in people_rows},
    }
def week_from_mnav_capacity_tab(sheet_name: str, ws: pd.DataFrame) -> Optional[pd.Timestamp]:
    s = str(sheet_name).strip()
    s_lower = s.lower()
    if "capacity mgmt" not in s_lower:
        return None
    candidate_cells = [
        (1, 0),  # A2
        (0, 1),  # B1
        (1, 1),  # B2
        (0, 0),  # A1
    ]
    for r, c in candidate_cells:
        try:
            v = ws.iat[r, c]
            dt = pd.to_datetime(v, errors="coerce")
            if _is_real_year(dt):
                return dt.normalize()
        except Exception:
            pass
    m = re.search(r"\((\d{1,2})\.(\d{1,2})\)", s)
    if not m:
        return None
    mm = int(m.group(1))
    dd = int(m.group(2))
    for r in range(0, min(6, ws.shape[0])):
        for c in range(0, min(6, ws.shape[1])):
            try:
                v = ws.iat[r, c]
            except Exception:
                continue
            dt = pd.to_datetime(v, errors="coerce")
            if _is_real_year(dt):
                return pd.Timestamp(year=int(dt.year), month=mm, day=dd).normalize()
    return pd.Timestamp(year=DEFAULT_YEAR_IF_MISSING, month=mm, day=dd).normalize()
def build_meic_teamtracker_block(ws: pd.DataFrame) -> Dict:
    PEOPLE_START_ROW = 2         
    HEADER_ROW = 1                
    NAME_COL = 0                  
    NONWIP_COL = 1                
    OOO_COL = 2                   
    ACTIVITY_START_COL = 3        
    ACTIVITY_END_COL = 23         
    people_rows: List[dict] = []
    for i in range(PEOPLE_START_ROW, len(ws)):
        name = norm_name(ws.iat[i, NAME_COL] if ws.shape[1] > NAME_COL else "")
        if not name:
            continue
        if name.strip().lower() == "total":
            break
        if not is_real_person(name):
            continue
        nonwip = safe_float0(ws.iat[i, NONWIP_COL] if ws.shape[1] > NONWIP_COL else 0.0)
        ooo = safe_float0(ws.iat[i, OOO_COL] if ws.shape[1] > OOO_COL else 0.0)
        people_rows.append({
            "row_i": i,
            "name": name,
            "NONWIP": float(nonwip),
            "OOO": float(ooo),
        })
    people_count = len(set(r["name"] for r in people_rows))
    ooo_hours = float(round(sum(r["OOO"] for r in people_rows), 2))
    total_nonwip_hours = float(round(sum(r["NONWIP"] for r in people_rows), 2))
    nonwip_by_person: Dict[str, float] = {}
    for r in people_rows:
        v = float(round(r["NONWIP"], 2))
        if v != 0.0:
            nonwip_by_person[r["name"]] = v
    activities: List[dict] = []
    for pr in people_rows:
        i = pr["row_i"]
        name = pr["name"]
        for c in range(ACTIVITY_START_COL, min(ACTIVITY_END_COL, ws.shape[1] - 1) + 1):
            label = norm_name(ws.iat[HEADER_ROW, c] if ws.shape[1] > c else "")
            if not label:
                continue
            hrs = safe_float(ws.iat[i, c] if ws.shape[1] > c else np.nan)
            if pd.isna(hrs) or hrs <= 0:
                continue
            activities.append({
                "name": name,
                "activity": label,
                "hours": float(round(float(hrs), 2)),
            })
        ooo = float(round(pr["OOO"], 2))
        if ooo > 0:
            activities.append({
                "name": name,
                "activity": "OOO",
                "hours": ooo,
            })
    return {
        "people_rows": people_rows,
        "people_count": people_count,
        "ooo_hours": ooo_hours,
        "total_nonwip_hours": total_nonwip_hours,
        "nonwip_by_person": nonwip_by_person,
        "nonwip_activities": activities,
        "ooo_map": {r["name"]: float(r["OOO"]) for r in people_rows},
    }
def split_meic_snapshot_into_teams(built: Dict) -> Dict[str, Dict]:
    people_rows = built["people_rows"]
    activities = built["nonwip_activities"]
    def team_for_person(name: str) -> str:
        if name in DBS_MEIC_NAMES:
            return "DBS MEIC"
        if name in PH_MEIC_NAMES:
            return "PH MEIC"
        return "SCS MEIC"
    out: Dict[str, Dict] = {
        "DBS MEIC": {"people_rows": [], "activities": []},
        "PH MEIC": {"people_rows": [], "activities": []},
        "SCS MEIC": {"people_rows": [], "activities": []},
    }
    for r in people_rows:
        out[team_for_person(r["name"])]["people_rows"].append(r)
    for a in activities:
        out[team_for_person(a["name"])]["activities"].append(a)
    final = {}
    for team_name, data in out.items():
        prs = data["people_rows"]
        acts = data["activities"]
        nonwip_by_person = {}
        for r in prs:
            v = float(round(r["NONWIP"], 2))
            if v != 0.0:
                nonwip_by_person[r["name"]] = v
        final[team_name] = {
            "people_rows": prs,
            "people_count": len(set(r["name"] for r in prs)),
            "ooo_hours": float(round(sum(r["OOO"] for r in prs), 2)),
            "total_nonwip_hours": float(round(sum(r["NONWIP"] for r in prs), 2)),
            "nonwip_by_person": nonwip_by_person,
            "nonwip_activities": acts,
            "ooo_map": {r["name"]: float(r["OOO"]) for r in prs},
        }
    return final
def get_people_count_from_wip(
    wip_df: pd.DataFrame,
    team: str,
    week: pd.Timestamp,
    fallback: Optional[int] = None,
    component_teams: Optional[set] = None,
) -> int:
    team_name = str(team).strip()
    if team_name == "DBS":
        try:
            return get_dbs_people_count_from_heijunka_files(
                file_paths=(DBS_C13_SOURCE_FILE, DBS_C14_SOURCE_FILE),
                name_row_zero_based=29,   # Excel row 30
            )
        except Exception as e:
            print(f"[WARN][DBS] failed special people count: {e}", flush=True)
            return int(fallback or 0)
    if wip_df is None or wip_df.empty:
        return int(fallback or 0)
    base = wip_df[wip_df["period_date"] == week].copy()
    if base.empty:
        return int(fallback or 0)
    direct = base[base.get("team") == team]
    if not direct.empty:
        for col in ["HC in WIP", "HC_in_WIP", "hc in wip", "hc_in_wip"]:
            if col in direct.columns:
                vals = pd.to_numeric(direct[col], errors="coerce").dropna()
                if not vals.empty:
                    return int(vals.iloc[0])
    if component_teams:
        subset = base[base.get("team").isin(component_teams)]
        if not subset.empty:
            for col in ["HC in WIP", "HC_in_WIP", "hc in wip", "hc_in_wip"]:
                if col in subset.columns:
                    vals = pd.to_numeric(subset[col], errors="coerce").fillna(0)
                    return int(vals.sum())
    return int(fallback or 0)
def build_meic_rows_from_team_tracker(
    xlsx_path: Path,
    wip_df: pd.DataFrame,
    metrics_df: pd.DataFrame,
) -> pd.DataFrame:
    if not xlsx_path.exists():
        print(f"[WARN] Missing XLSX for MEIC tracker: {xlsx_path}")
        return pd.DataFrame()
    out_rows: List[dict] = []
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = True
    try:
        excel.AutomationSecurity = 1
    except Exception:
        pass
    wb = None
    temp_dir = None
    try:
        temp_dir = mkdtemp(prefix="meic_tracker_")
        temp_xlsx = Path(temp_dir) / xlsx_path.name
        shutil.copy2(xlsx_path, temp_xlsx)
        wb = excel.Workbooks.Open(str(temp_xlsx))
        ws_com = wb.Worksheets(TEAM_TRACKER_SHEET)
        excel.CalculateFullRebuild()
        all_dates = _resolve_validation_list_values(wb, ws_com, "B1")
        if not all_dates:
            current_dt = _excel_date_to_timestamp(ws_com.Range("B1").Value)
            if current_dt is not None:
                all_dates = [current_dt]
        for week in all_dates:
            try:
                ws_com.Range("B1").Value = week.to_pydatetime()
                wb.RefreshAll()
                excel.CalculateUntilAsyncQueriesDone()
                excel.CalculateFullRebuild()
                used = ws_com.UsedRange.Value
                if used is None:
                    continue
                if not isinstance(used, tuple):
                    used = ((used,),)
                ws_df = pd.DataFrame(list(used))
                built = build_meic_teamtracker_block(ws_df)
                split = split_meic_snapshot_into_teams(built)
                for team_name, team_built in split.items():
                    completed_match = metrics_df[
                        (metrics_df.get("team") == team_name) &
                        (metrics_df["period_date"] == week)
                    ]
                    completed_hours = (
                        pd.to_numeric(completed_match.iloc[0].get("Completed Hours"), errors="coerce")
                        if not completed_match.empty else np.nan
                    )
                    pct_in_wip = np.nan
                    if pd.notna(completed_hours) and pd.notna(team_built["total_nonwip_hours"]):
                        denom = float(completed_hours) + float(team_built["total_nonwip_hours"])
                        pct_in_wip = float(completed_hours) / denom if denom != 0 else np.nan
                    wip_match = metrics_df[
                        (metrics_df.get("team") == team_name) &
                        (metrics_df["period_date"] == week)
                    ]
                    wip_workers = extract_wip_workers_from_row(wip_match.iloc[0]) if not wip_match.empty else []
                    wip_workers_count = len(wip_workers)
                    wip_workers_ooo_hours = float(round(
                        sum(safe_float0(team_built["ooo_map"].get(n, 0.0)) for n in wip_workers), 2
                    ))
                    people_count_final = get_people_count_from_wip(
                        wip_df=wip_df,
                        team=team_name,
                        week=week,
                        fallback=team_built["people_count"],
                    )
                    out_rows.append({
                        "team": team_name,
                        "period_date": week.date().isoformat(),
                        "source_file": str(xlsx_path),
                        "people_count": int(people_count_final),
                        "total_non_wip_hours": float(round(team_built["total_nonwip_hours"], 2)),
                        "OOO Hours": float(round(team_built["ooo_hours"], 2)),
                        "% in WIP": float(round(pct_in_wip, 6)) if pd.notna(pct_in_wip) else np.nan,
                        "non_wip_by_person": json.dumps(team_built["nonwip_by_person"], ensure_ascii=False),
                        "non_wip_activities": json.dumps(team_built["nonwip_activities"], ensure_ascii=False),
                        "wip_workers": json.dumps(wip_workers, ensure_ascii=False),
                        "wip_workers_count": int(wip_workers_count),
                        "wip_workers_ooo_hours": float(wip_workers_ooo_hours),
                    })
            except Exception as e:
                print(f"[WARN] Failed MEIC week {week}: {e}")
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        excel.Quit()
        if temp_dir and Path(temp_dir).exists():
            shutil.rmtree(temp_dir, ignore_errors=True)
    df = pd.DataFrame(out_rows)
    if not df.empty:
        df["period_date"] = pd.to_datetime(df["period_date"], errors="coerce").dt.normalize()
        df = df.drop_duplicates(subset=["team", "period_date"], keep="last")
        df = df.sort_values(["team", "period_date"]).reset_index(drop=True)
    return df
def week_from_nv_tab(sheet_name: str, ws: pd.DataFrame) -> Optional[pd.Timestamp]:
    s = str(sheet_name).strip()
    m = re.fullmatch(r"(\d{2})([A-Za-z]{3})(\d{4})", s)
    if m:
        dt = pd.to_datetime(s, format="%d%b%Y", errors="coerce")
        if pd.notna(dt):
            return dt.normalize()
    dt = pd.to_datetime(s, errors="coerce")
    if pd.notna(dt):
        return dt.normalize()
    return None
def build_nv_row(team: str, ws: pd.DataFrame, week: Optional[pd.Timestamp] = None) -> Dict:
    PEOPLE_START = 1   
    PEOPLE_END   = 12 
    COL_EXPECTED = _col_letter_to_idx("B")   
    COL_OOO      = _col_letter_to_idx("X")  
    COL_NONWIP   = _col_letter_to_idx("Y")   
    DEDUCT_CELL  = "B19"
    ACT_HEADER_ROW = 0  
    ACT_START_COL  = _col_letter_to_idx("C")
    ACT_END_COL    = _col_letter_to_idx("W")
    m = re.fullmatch(r"([A-Za-z]+)(\d+)", DEDUCT_CELL.strip())
    deduct_col = _col_letter_to_idx(m.group(1))
    deduct_row = int(m.group(2)) - 1
    people_rows: List[dict] = []
    for i in range(PEOPLE_START, PEOPLE_END + 1):
        name = norm_name(ws.iat[i, 0] if ws.shape[1] > 0 else "")
        if not name:
            continue
        if not is_real_person(name):
            continue
        expected = safe_float0(ws.iat[i, COL_EXPECTED] if ws.shape[1] > COL_EXPECTED else 0.0)
        ooo      = safe_float0(ws.iat[i, COL_OOO]      if ws.shape[1] > COL_OOO      else 0.0)
        nonwip   = safe_float0(ws.iat[i, COL_NONWIP]   if ws.shape[1] > COL_NONWIP   else 0.0)
        people_rows.append(
            {"row_i": i, "name": name, "B": float(expected), "OOO": float(ooo), "NONWIP": float(nonwip)}
        )
    people_count = len(set(r["name"] for r in people_rows))
    ooo_hours = float(round(sum(r["OOO"] for r in people_rows), 2))
    deduct_val = safe_float0(ws.iat[deduct_row, deduct_col] if ws.shape[0] > deduct_row and ws.shape[1] > deduct_col else 0.0)
    total_nonwip_hours = float(round((people_count * 40.0) - float(deduct_val) - float(ooo_hours), 2))
    nonwip_by_person: Dict[str, float] = {}
    for r in people_rows:
        v = float(round(r["NONWIP"], 2))
        if v != 0.0:
            nonwip_by_person[r["name"]] = v
    activities: List[dict] = []
    for pr in people_rows:
        i = pr["row_i"]
        name = pr["name"]
        for c in range(ACT_START_COL, min(ACT_END_COL, ws.shape[1] - 1) + 1):
            label = norm_name(ws.iat[ACT_HEADER_ROW, c] if ws.shape[0] > ACT_HEADER_ROW and ws.shape[1] > c else "")
            if not label:
                continue
            hrs = safe_float(ws.iat[i, c] if ws.shape[0] > i and ws.shape[1] > c else np.nan)
            if pd.isna(hrs) or hrs <= 0:
                continue
            activities.append({"name": name, "activity": label, "hours": float(round(float(hrs), 2))})
        ooo = float(round(pr["OOO"], 2))
        if ooo > 0:
            activities.append({
                "name": name,
                "activity": "OOO",
                "hours": ooo,
            })
    return {
        "people_rows": people_rows,
        "people_count": people_count,
        "ooo_hours": ooo_hours,
        "total_nonwip_hours": total_nonwip_hours,
        "nonwip_by_person": nonwip_by_person,
        "nonwip_activities": activities,
        "ooo_map": {r["name"]: float(r["OOO"]) for r in people_rows},
    }
def build_mnav_row(team: str, ws: pd.DataFrame, week: Optional[pd.Timestamp] = None) -> Dict:
    PEOPLE_START = 2
    PEOPLE_END = 18
    COL_B = 1
    COL_AA = 26
    COL_AF = 31
    ooo_col = COL_AA if (week is not None and week.month == 2 and week.day == 16) else COL_AF
    COL_C = 2
    COL_AE = 30
    HEADER_ROW = 1
    B21_ROW = 20
    people_rows: List[dict] = []
    for i in range(PEOPLE_START, PEOPLE_END + 1):
        name = norm_name(ws.iat[i, 0] if ws.shape[1] > 0 else "")
        if not name:
            continue
        if not is_real_person(name):
            continue
        b = safe_float(ws.iat[i, COL_B] if ws.shape[1] > COL_B else np.nan)
        ooo = safe_float(ws.iat[i, ooo_col] if ws.shape[1] > ooo_col else np.nan)
        if pd.isna(b): b = 0.0
        if pd.isna(ooo): ooo = 0.0
        people_rows.append({"row_i": i, "name": name, "B": b, "OOO": ooo})
    people_count = len(set(r["name"] for r in people_rows))
    ooo_hours = float(round(sum(r["OOO"] for r in people_rows), 2))
    b21 = safe_float(ws.iat[B21_ROW, COL_B] if ws.shape[0] > B21_ROW and ws.shape[1] > COL_B else np.nan)
    if pd.isna(b21): b21 = 0.0
    total_nonwip_hours = float(round((people_count * 40.0) - float(b21) - float(ooo_hours), 2))
    nonwip_by_person: Dict[str, float] = {}
    for r in people_rows:
        v = float(round(40.0 - float(r["B"]) - float(r["OOO"]), 2))
        if v == 0.0:
            continue
        nonwip_by_person[r["name"]] = v
    activities: List[dict] = []
    for pr in people_rows:
        i = pr["row_i"]
        name = pr["name"]
        for c in range(COL_C, COL_AE + 1):
            label = norm_name(ws.iat[HEADER_ROW, c] if ws.shape[1] > c else "")
            if not label:
                continue
            hrs = safe_float(ws.iat[i, c] if ws.shape[1] > c else np.nan)
            if pd.isna(hrs) or hrs <= 0:
                continue
            activities.append({"name": name, "activity": label, "hours": float(round(hrs, 2))})
        ooo = float(round(pr["OOO"], 2))
        if ooo > 0:
            activities.append({
                "name": name,
                "activity": "OOO",
                "hours": ooo,
            })
    return {
        "people_rows": people_rows,
        "people_count": people_count,
        "ooo_hours": ooo_hours,
        "total_nonwip_hours": total_nonwip_hours,
        "nonwip_by_person": nonwip_by_person,
        "nonwip_activities": activities,
        "ooo_map": {r["name"]: float(r["OOO"]) for r in people_rows},
    }
ENABLE_TEAMS = {"AE MEIC", "CSF", "Mazor", "O-Arm MEIC", "Nav"}
ENABLE_TEAM_NAME = "Enabling Technologies"
MEIC_PARENT_MAP = {
    "PH": {"PH", "PH MEIC"},
    "DBS": {"DBS", "DBS MEIC"},
    "SCS": {"SCS", "SCS MEIC"},
    "PSS": {"PSS Intern", "PSS US", "PSS MEIC"}
}
def combine_meic_parent_teams(df: pd.DataFrame, wip_df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "team" not in df.columns or "period_date" not in df.columns:
        return df
    df = df.copy()
    df["period_date"] = pd.to_datetime(df["period_date"], errors="coerce").dt.normalize()
    teams_to_merge = set().union(*MEIC_PARENT_MAP.values())
    subset = df[df["team"].isin(teams_to_merge)].copy()
    rest = df[~df["team"].isin(teams_to_merge)].copy()
    if subset.empty:
        if "source_file" in rest.columns:
            rest = rest.drop(columns=["source_file"])
        return rest
    out_rows = []
    for parent_team, member_teams in MEIC_PARENT_MAP.items():
        g_team = subset[subset["team"].isin(member_teams)].copy()
        if g_team.empty:
            continue
        for period_date, g in g_team.groupby("period_date", dropna=False):
            nonwip_by_person = _merge_person_hours_dicts(g.get("non_wip_by_person"))
            nonwip_activities = _merge_activities_lists(g.get("non_wip_activities"))
            wip_workers_union = _merge_workers_union(g.get("wip_workers"))
            fallback_people_count = int(pd.to_numeric(g.get("people_count"), errors="coerce").fillna(0).sum())
            if parent_team in {"DBS", "SCS", "PH"}:
                people_count_final = fallback_people_count
            else:
                people_count_final = get_people_count_from_wip(
                    wip_df=wip_df,
                    team=parent_team,
                    week=period_date,
                    fallback=fallback_people_count,
                    component_teams=member_teams,
                )
            out_rows.append({
                "team": parent_team,
                "period_date": period_date,
                "people_count": int(people_count_final),
                "total_non_wip_hours": float(pd.to_numeric(g.get("total_non_wip_hours"), errors="coerce").fillna(0).sum()),
                "OOO Hours": float(pd.to_numeric(g.get("OOO Hours"), errors="coerce").fillna(0).sum()),
                "% in WIP": float(pd.to_numeric(g.get("% in WIP"), errors="coerce").mean()),
                "non_wip_by_person": json.dumps(nonwip_by_person, ensure_ascii=False),
                "non_wip_activities": json.dumps(nonwip_activities, ensure_ascii=False),
                "wip_workers": json.dumps(wip_workers_union, ensure_ascii=False),
                "wip_workers_count": int(pd.to_numeric(g.get("wip_workers_count"), errors="coerce").fillna(0).sum()),
                "wip_workers_ooo_hours": float(pd.to_numeric(g.get("wip_workers_ooo_hours"), errors="coerce").fillna(0).sum()),
            })
    merged_df = pd.DataFrame(out_rows)
    for dfx in (rest, merged_df):
        if "source_file" in dfx.columns:
            dfx.drop(columns=["source_file"], inplace=True)
    combined = pd.concat([rest, merged_df], ignore_index=True)
    combined["period_date"] = pd.to_datetime(combined["period_date"], errors="coerce").dt.normalize()
    combined = combined.sort_values(["team", "period_date"]).reset_index(drop=True)
    return combined
def _parse_json_dict(cell) -> dict:
    if cell is None or (isinstance(cell, float) and pd.isna(cell)):
        return {}
    if isinstance(cell, dict):
        return cell
    s = str(cell).strip()
    if not s:
        return {}
    try:
        obj = json.loads(s)
        return obj if isinstance(obj, dict) else {}
    except Exception:
        return {}
def _parse_json_list(cell) -> list:
    if cell is None or (isinstance(cell, float) and pd.isna(cell)):
        return []
    if isinstance(cell, list):
        return cell
    s = str(cell).strip()
    if not s:
        return []
    try:
        obj = json.loads(s)
        return obj if isinstance(obj, list) else []
    except Exception:
        return []
def _parse_json_str_list(cell) -> list:
    lst = _parse_json_list(cell)
    out = []
    for x in lst:
        n = norm_name(x)
        if n:
            out.append(n)
    return out
def _merge_person_hours_dicts(dict_cells: List) -> dict:
    merged: Dict[str, float] = {}
    for cell in dict_cells:
        d = _parse_json_dict(cell)
        for k, v in d.items():
            name = norm_name(k)
            if not name:
                continue
            hrs = safe_float0(v)
            merged[name] = float(round(merged.get(name, 0.0) + hrs, 2))
    merged = {k: v for k, v in merged.items() if v != 0.0}
    return merged
def _merge_activities_lists(list_cells: List) -> list:
    merged = []
    for cell in list_cells:
        merged.extend(_parse_json_list(cell))
    return merged
def _merge_workers_union(list_cells: List) -> list:
    s = set()
    for cell in list_cells:
        for name in _parse_json_str_list(cell):
            if is_real_person(name):
                s.add(name)
    return sorted(s)
def combine_enabling_technologies(df: pd.DataFrame, wip_df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "team" not in df.columns or "period_date" not in df.columns:
        return df
    df = df.copy()
    df["period_date"] = pd.to_datetime(df["period_date"], errors="coerce").dt.normalize()
    subset = df[df["team"].isin(ENABLE_TEAMS)].copy()
    rest = df[~df["team"].isin(ENABLE_TEAMS)].copy()
    if subset.empty:
        if "source_file" in rest.columns:
            rest = rest.drop(columns=["source_file"])
        return rest
    out_rows = []
    for period_date, g in subset.groupby("period_date", dropna=False):
        nonwip_by_person = _merge_person_hours_dicts(g.get("non_wip_by_person"))
        nonwip_activities = _merge_activities_lists(g.get("non_wip_activities"))
        wip_workers_union = _merge_workers_union(g.get("wip_workers"))
        people_count_final = get_people_count_from_wip(
            wip_df=wip_df,
            team=ENABLE_TEAM_NAME,
            week=period_date,
            fallback=int(pd.to_numeric(g.get("people_count"), errors="coerce").fillna(0).sum()),
            component_teams=ENABLE_TEAMS,
        )
        out_rows.append({
            "team": ENABLE_TEAM_NAME,
            "period_date": period_date,
            "people_count": int(people_count_final),
            "total_non_wip_hours": float(pd.to_numeric(g.get("total_non_wip_hours"), errors="coerce").fillna(0).sum()),
            "OOO Hours": float(pd.to_numeric(g.get("OOO Hours"), errors="coerce").fillna(0).sum()),
            "% in WIP": float(pd.to_numeric(g.get("% in WIP"), errors="coerce").mean()),
            "non_wip_by_person": json.dumps(nonwip_by_person, ensure_ascii=False),
            "non_wip_activities": json.dumps(nonwip_activities, ensure_ascii=False),
            "wip_workers": json.dumps(wip_workers_union, ensure_ascii=False),
            "wip_workers_count": int(pd.to_numeric(g.get("wip_workers_count"), errors="coerce").fillna(0).sum()),
            "wip_workers_ooo_hours": float(pd.to_numeric(g.get("wip_workers_ooo_hours"), errors="coerce").fillna(0).sum()),
        })
    enabling_df = pd.DataFrame(out_rows)
    for dfx in (rest, enabling_df):
        if "source_file" in dfx.columns:
            dfx.drop(columns=["source_file"], inplace=True)
    combined = pd.concat([rest, enabling_df], ignore_index=True)
    combined["period_date"] = pd.to_datetime(combined["period_date"], errors="coerce").dt.normalize()
    combined = combined.sort_values(["team", "period_date"]).reset_index(drop=True)
    return combined
def _col_letter_to_idx(letter: str) -> int:
    s = str(letter).strip().upper()
    n = 0
    for ch in s:
        if not ("A" <= ch <= "Z"):
            continue
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1
def build_capacity_fixed_row(
    team: str,
    ws: pd.DataFrame,
    *,
    people_start_row: int,
    people_end_row: int,
    expected_col_letter: str,   
    ooo_col_letter: str,        
    deduct_cell: str,            
    ooo_sum_start_row: int,      
    ooo_sum_end_row: int,       
    total_ooo_end_row: int,     
    activity_header_row: int,    
    activity_start_col_letter: str,
    activity_end_col_letter: str,
) -> Dict:
    col_b = _col_letter_to_idx(expected_col_letter)
    col_ooo = _col_letter_to_idx(ooo_col_letter)
    act_start = _col_letter_to_idx(activity_start_col_letter)
    act_end   = _col_letter_to_idx(activity_end_col_letter)
    m = re.fullmatch(r"([A-Za-z]+)(\d+)", deduct_cell.strip())
    if not m:
        raise ValueError(f"Bad deduct_cell: {deduct_cell}")
    deduct_col = _col_letter_to_idx(m.group(1))
    deduct_row = int(m.group(2)) - 1
    people_rows: List[dict] = []
    for i in range(people_start_row, people_end_row + 1):
        name = norm_name(ws.iat[i, 0] if ws.shape[1] > 0 else "")
        if not name:
            continue
        if not is_real_person(name):
            continue
        b = safe_float(ws.iat[i, col_b] if ws.shape[1] > col_b else np.nan)
        ooo = safe_float(ws.iat[i, col_ooo] if ws.shape[1] > col_ooo else np.nan)
        if pd.isna(b): b = 0.0
        if pd.isna(ooo): ooo = 0.0
        people_rows.append({"row_i": i, "name": name, "B": float(b), "OOO": float(ooo)})
    people_count = len(set(r["name"] for r in people_rows))
    ooo_hours = 0.0
    for r in range(ooo_sum_start_row, ooo_sum_end_row + 1):
        val = ws.iat[r, col_ooo] if (ws.shape[0] > r and ws.shape[1] > col_ooo) else 0.0
        ooo_hours += safe_float0(val)
    ooo_hours = float(round(ooo_hours, 2))
    deduct_val = safe_float(ws.iat[deduct_row, deduct_col] if ws.shape[0] > deduct_row and ws.shape[1] > deduct_col else 0.0)
    if pd.isna(deduct_val):
        deduct_val = 0.0
    total_ooo = 0.0
    for r in range(ooo_sum_start_row, total_ooo_end_row + 1):
        val = ws.iat[r, col_ooo] if (ws.shape[0] > r and ws.shape[1] > col_ooo) else 0.0
        total_ooo += safe_float0(val)
    total_ooo = float(round(total_ooo, 2))
    total_nonwip_hours = float(round((people_count * 40.0) - float(deduct_val) - float(total_ooo), 2))
    nonwip_by_person: Dict[str, float] = {}
    for r in people_rows:
        v = float(round(40.0 - float(r["B"]) - float(r["OOO"]), 2))
        if v != 0.0:
            nonwip_by_person[r["name"]] = v
    activities: List[dict] = []
    for pr in people_rows:
        i = pr["row_i"]
        name = pr["name"]
        for c in range(act_start, min(act_end, ws.shape[1] - 1) + 1):
            label = norm_name(ws.iat[activity_header_row, c] if ws.shape[0] > activity_header_row and ws.shape[1] > c else "")
            if not label:
                continue
            hrs = safe_float(ws.iat[i, c] if ws.shape[0] > i and ws.shape[1] > c else np.nan)
            if pd.isna(hrs) or hrs <= 0:
                continue
            activities.append({"name": name, "activity": label, "hours": float(round(hrs, 2))})
        ooo = float(round(pr["OOO"], 2))
        if ooo > 0:
            activities.append({
                "name": name,
                "activity": "OOO",
                "hours": ooo,
            })
    return {
        "people_rows": people_rows,
        "people_count": people_count,
        "ooo_hours": ooo_hours,
        "total_nonwip_hours": total_nonwip_hours,
        "nonwip_by_person": nonwip_by_person,
        "nonwip_activities": activities,
        "ooo_map": {r["name"]: float(r["OOO"]) for r in people_rows},
    }
OOO_LABELS = {"OUT OF OFFICE", "HOLIDAY", "OOO"}
def build_ent_row(team: str, ws: pd.DataFrame, week: Optional[pd.Timestamp] = None) -> Dict:
    PEOPLE_START = 2    # Excel row 3
    PEOPLE_END   = 25   # Excel row 26
    TOTAL_ROW    = 26   # Excel row 27
    COL_B  = _col_letter_to_idx("B")
    COL_Z  = _col_letter_to_idx("Z")
    COL_AA = _col_letter_to_idx("AA")
    ACT_START  = _col_letter_to_idx("C")
    ACT_END    = _col_letter_to_idx("AG")
    HEADER_ROW = 1      # Excel row 2
    people_rows: List[dict] = []
    for i in range(PEOPLE_START, PEOPLE_END + 1):
        name = norm_name(ws.iat[i, 0] if ws.shape[1] > 0 else "")
        if not name or not is_real_person(name):
            continue
        b = safe_float0(ws.iat[i, COL_B] if ws.shape[1] > COL_B else 0.0)
        z = safe_float0(ws.iat[i, COL_Z] if ws.shape[1] > COL_Z else 0.0)
        aa = safe_float0(ws.iat[i, COL_AA] if ws.shape[1] > COL_AA else 0.0)
        zaa_ooo = float(round(z + aa, 2))

        people_rows.append({
            "row_i": i,
            "name": name,
            "B": b,
            "Z_OOO": z,
            "AA_OOO": aa,
            "ZAA_OOO": zaa_ooo,
        })
    people_count = len(set(r["name"] for r in people_rows))
    activities: List[dict] = []
    nonwip_by_person: Dict[str, float] = {}
    ooo_map: Dict[str, float] = {}
    for pr in people_rows:
        i = pr["row_i"]
        name = pr["name"]
        person_nonwip_total = 0.0
        activity_ooo_total = 0.0
        for c in range(ACT_START, min(ACT_END, ws.shape[1] - 1) + 1):
            label = norm_name(ws.iat[HEADER_ROW, c] if ws.shape[0] > HEADER_ROW and ws.shape[1] > c else "")
            if not label:
                continue
            hrs = safe_float(ws.iat[i, c] if ws.shape[0] > i and ws.shape[1] > c else np.nan)
            if pd.isna(hrs) or hrs <= 0:
                continue
            hrs = float(round(hrs, 2))
            label_upper = label.strip().upper()
            if label_upper in OOO_LABELS:
                activity_ooo_total += hrs
                continue
            activities.append({
                "name": name,
                "activity": label,
                "hours": hrs,
            })
            person_nonwip_total += hrs
        person_ooo = float(round(
            activity_ooo_total if activity_ooo_total > 0 else pr["ZAA_OOO"],
            2
        ))
        if person_ooo > 0:
            activities.append({
                "name": name,
                "activity": "OOO",
                "hours": person_ooo,
            })
        if person_nonwip_total > 0:
            nonwip_by_person[name] = float(round(person_nonwip_total, 2))
        ooo_map[name] = person_ooo
    ooo_hours = float(round(sum(ooo_map.values()), 2))
    row_27_total = float(round(sum(
        safe_float0(ws.iat[TOTAL_ROW, c] if ws.shape[0] > TOTAL_ROW and ws.shape[1] > c else 0.0)
        for c in range(ACT_START, min(ACT_END, ws.shape[1] - 1) + 1)
    ), 2))
    total_nonwip_hours = float(round(row_27_total - ooo_hours, 2))
    return {
        "people_rows": people_rows,
        "people_count": people_count,
        "ooo_hours": ooo_hours,
        "total_nonwip_hours": total_nonwip_hours,
        "nonwip_by_person": nonwip_by_person,
        "nonwip_activities": activities,
        "ooo_map": ooo_map,
    }
def build_ae_meic_row(team: str, ws: pd.DataFrame, week: Optional[pd.Timestamp] = None) -> Dict:
    return build_capacity_fixed_row(
        team, ws,
        people_start_row=1, people_end_row=5,
        expected_col_letter="B",
        ooo_col_letter="Q",
        deduct_cell="B8",
        ooo_sum_start_row=1, ooo_sum_end_row=5, 
        total_ooo_end_row=5,             
        activity_header_row=0,           
        activity_start_col_letter="C",
        activity_end_col_letter="P",
    )
def build_pss_us_row(team: str, ws: pd.DataFrame, week: Optional[pd.Timestamp] = None) -> Dict:
    NAME_COL   = _col_letter_to_idx("A")
    COL_B      = _col_letter_to_idx("B")   # Expected WIP hrs
    ACT_START  = _col_letter_to_idx("C")
    ACT_END    = _col_letter_to_idx("R")   # last activity before OOO
    COL_OOO    = _col_letter_to_idx("S")   # Out of Office
    HEADER_ROW = 0                         # Excel row 1
    PEOPLE_START_ROW = 1                   # Excel row 2
    people_rows: List[dict] = []
    seen_people = False
    blank_run = 0
    max_row = ws.shape[0] - 1
    for i in range(PEOPLE_START_ROW, max_row + 1):
        raw_name = ws.iat[i, NAME_COL] if ws.shape[1] > NAME_COL else ""
        name = norm_name(raw_name)
        if is_real_person(name):
            seen_people = True
            blank_run = 0
            b = safe_float0(ws.iat[i, COL_B] if ws.shape[1] > COL_B else 0.0)
            ooo = safe_float0(ws.iat[i, COL_OOO] if ws.shape[1] > COL_OOO else 0.0)
            people_rows.append({
                "row_i": i,
                "name": name,
                "B": float(b),
                "OOO": float(ooo),
            })
        else:
            if seen_people:
                blank_run += 1
                if blank_run >= 3:
                    break
    people_count = len(set(r["name"] for r in people_rows))
    ooo_hours = float(round(sum(r["OOO"] for r in people_rows), 2))
    activities: List[dict] = []
    nonwip_by_person: Dict[str, float] = {}
    for pr in people_rows:
        i = pr["row_i"]
        name = pr["name"]
        person_total = 0.0
        for c in range(ACT_START, min(ACT_END, ws.shape[1] - 1) + 1):
            label = norm_name(ws.iat[HEADER_ROW, c] if ws.shape[0] > HEADER_ROW and ws.shape[1] > c else "")
            if not label:
                continue
            hrs = safe_float(ws.iat[i, c] if ws.shape[0] > i and ws.shape[1] > c else np.nan)
            if pd.isna(hrs) or hrs <= 0:
                continue
            hrs = float(round(float(hrs), 2))
            activities.append({
                "name": name,
                "activity": label,
                "hours": hrs,
            })
            person_total += hrs
        ooo = float(round(pr["OOO"], 2))
        if ooo > 0:
            activities.append({
                "name": name,
                "activity": "OOO",
                "hours": ooo,
            })
        if person_total != 0.0:
            nonwip_by_person[name] = float(round(person_total, 2))
    total_nonwip_hours = float(round(sum(a["hours"] for a in activities), 2))
    return {
        "people_rows": people_rows,
        "people_count": people_count,
        "ooo_hours": ooo_hours,
        "total_nonwip_hours": total_nonwip_hours,
        "nonwip_by_person": nonwip_by_person,
        "nonwip_activities": activities,
        "ooo_map": {r["name"]: float(r["OOO"]) for r in people_rows},
    }
def build_oarm_meic_row(team: str, ws: pd.DataFrame, week: Optional[pd.Timestamp] = None) -> Dict:
    return build_capacity_fixed_row(
        team, ws,
        people_start_row=1, people_end_row=8,
        expected_col_letter="B",
        ooo_col_letter="Q",
        deduct_cell="B11",
        ooo_sum_start_row=1, ooo_sum_end_row=8, 
        total_ooo_end_row=8,            
        activity_header_row=0,
        activity_start_col_letter="C",
        activity_end_col_letter="P",
    )
def build_mazor_row(team: str, ws: pd.DataFrame, week: Optional[pd.Timestamp] = None) -> Dict:
    return build_capacity_fixed_row(
        team, ws,
        people_start_row=1, people_end_row=7,
        expected_col_letter="B",
        ooo_col_letter="Z",
        deduct_cell="B10",
        ooo_sum_start_row=1, ooo_sum_end_row=8,
        total_ooo_end_row=7,      
        activity_header_row=0,
        activity_start_col_letter="C",
        activity_end_col_letter="Y",
    )
def week_from_pss_us_tab(sheet_name: str, ws: pd.DataFrame) -> Optional[pd.Timestamp]:
    s = str(sheet_name).strip()
    s_lower = s.lower()
    if "capacity mgmt" not in s_lower:
        return None
    m = re.search(r"\(([A-Za-z]{3,9})\.(\d{1,2})\)", s)
    if m:
        mon_txt = m.group(1)
        day = int(m.group(2))
        dt = pd.to_datetime(f"{mon_txt} {day} {DEFAULT_YEAR_IF_MISSING}", errors="coerce")
        if pd.notna(dt):
            return dt.normalize()
    dt = pd.to_datetime(s, errors="coerce")
    if pd.notna(dt):
        return dt.normalize()
    return None
def week_from_spine_tab(sheet_name: str, ws: pd.DataFrame) -> Optional[pd.Timestamp]:
    s = str(sheet_name).strip()
    s_lower = s.lower()
    bad_tab_words = {
        "instruction", "instructions", "setup", "config", "list", "lists",
        "lookup", "lookups", "read me", "readme", "cover", "template"
    }
    if any(word in s_lower for word in bad_tab_words):
        return None
    if ws.shape[0] < 18 or ws.shape[1] < 10:
        return None
    try:
        b1 = ws.iat[0, 1]
        dt = pd.to_datetime(b1, errors="coerce")
        if _is_real_year(dt):
            return dt.normalize()
    except Exception:
        pass
    dt = pd.to_datetime(s, errors="coerce")
    if _is_real_year(dt):
        return dt.normalize()
    m = re.search(r"(\d{1,2})[.\-_/](\d{1,2})[.\-_/](\d{2,4})", s)
    if m:
        mm = int(m.group(1))
        dd = int(m.group(2))
        yy = int(m.group(3))
        if yy < 100:
            yy += 2000
        try:
            return pd.Timestamp(year=yy, month=mm, day=dd).normalize()
        except Exception:
            pass
    return None
def week_from_ent_tab(sheet_name: str, ws: pd.DataFrame) -> Optional[pd.Timestamp]:
    s = str(sheet_name).strip()
    m = re.search(r"\((\d{1,2})[_\-. ]([A-Za-z]{3,9})\)", s)
    if not m:
        return None
    day = int(m.group(1))
    mon_txt = m.group(2)
    dt = pd.to_datetime(f"{day} {mon_txt} {DEFAULT_YEAR_IF_MISSING}", errors="coerce")
    if pd.isna(dt):
        return None
    return dt.normalize()
def build_spine_row(team: str, ws: pd.DataFrame, week: Optional[pd.Timestamp] = None) -> Dict:
    PEOPLE_START = 2  
    PEOPLE_END   = 17 
    COL_B   = _col_letter_to_idx("B")
    COL_OOO = _col_letter_to_idx("Y")
    ACT_START = _col_letter_to_idx("C")
    ACT_END   = _col_letter_to_idx("AC")
    HEADER_ROW = 1
    TEAM_HOURS_CELL = "B20"
    min_rows_needed = PEOPLE_END + 1   
    min_cols_needed = ACT_END + 1        
    if ws.shape[0] < min_rows_needed or ws.shape[1] < 2:
        return {
            "people_rows": [],
            "people_count": 0,
            "ooo_hours": 0.0,
            "total_nonwip_hours": np.nan,
            "nonwip_by_person": {},
            "nonwip_activities": [],
            "ooo_map": {},
        }
    m = re.fullmatch(r"([A-Za-z]+)(\d+)", TEAM_HOURS_CELL.strip())
    team_hours_col = _col_letter_to_idx(m.group(1))
    team_hours_row = int(m.group(2)) - 1
    people_rows: List[dict] = []
    last_people_row = min(PEOPLE_END, ws.shape[0] - 1)
    for i in range(PEOPLE_START, last_people_row + 1):
        name = norm_name(ws.iat[i, 0] if ws.shape[1] > 0 else "")
        if not name or not is_real_person(name):
            continue
        expected = safe_float0(ws.iat[i, COL_B] if ws.shape[1] > COL_B else 0.0)
        ooo      = safe_float0(ws.iat[i, COL_OOO] if ws.shape[1] > COL_OOO else 0.0)
        people_rows.append({
            "row_i": i,
            "name": name,
            "B": float(expected),
            "OOO": float(ooo),
        })
    people_count = len(set(r["name"] for r in people_rows))
    ooo_hours = float(round(sum(r["OOO"] for r in people_rows), 2))
    team_hours_available = safe_float0(
        ws.iat[team_hours_row, team_hours_col]
        if ws.shape[0] > team_hours_row and ws.shape[1] > team_hours_col else 0.0
    )
    total_nonwip_hours = float(round((people_count * 40.0) - team_hours_available - ooo_hours, 2))
    nonwip_by_person: Dict[str, float] = {}
    for r in people_rows:
        v = float(round(40.0 - float(r["B"]) - float(r["OOO"]), 2))
        if v != 0.0:
            nonwip_by_person[r["name"]] = v
    activities: List[dict] = []
    max_act_col = min(ACT_END, ws.shape[1] - 1)
    for pr in people_rows:
        i = pr["row_i"]
        name = pr["name"]
        for c in range(ACT_START, max_act_col + 1):
            if c == COL_OOO:
                continue
            label = norm_name(ws.iat[HEADER_ROW, c] if ws.shape[0] > HEADER_ROW and ws.shape[1] > c else "")
            if not label:
                continue
            hrs = safe_float(ws.iat[i, c] if ws.shape[0] > i and ws.shape[1] > c else np.nan)
            if pd.isna(hrs) or hrs <= 0:
                continue
            activities.append({
                "name": name,
                "activity": label,
                "hours": float(round(float(hrs), 2)),
            })
        ooo = float(round(pr["OOO"], 2))
        if ooo > 0:
            activities.append({
                "name": name,
                "activity": "OOO",
                "hours": ooo,
            })
    return {
        "people_rows": people_rows,
        "people_count": people_count,
        "ooo_hours": ooo_hours,
        "total_nonwip_hours": total_nonwip_hours,
        "nonwip_by_person": nonwip_by_person,
        "nonwip_activities": activities,
        "ooo_map": {r["name"]: float(r["OOO"]) for r in people_rows},
    }
def build_csf_row(team: str, ws: pd.DataFrame, week: Optional[pd.Timestamp] = None) -> Dict:
    return build_capacity_fixed_row(
        team, ws,
        people_start_row=1, people_end_row=5,       
        expected_col_letter="B",
        ooo_col_letter="AC",
        deduct_cell="B7",
        ooo_sum_start_row=1, ooo_sum_end_row=5,     
        total_ooo_end_row=5,                        
        activity_header_row=1,                       
        activity_start_col_letter="C",
        activity_end_col_letter="AB",
    )
TEAM_SOURCES: Dict[str, TeamSource] = {
    "PSS MEIC": TeamSource(
        team="PSS MEIC",
        xlsx=Path(r"C:\Users\wadec8\Medtronic PLC\PSS Sharepoint - Documents\PSS MEIC_Heijunka.xlsm"),
        week_from_sheet=week_from_pss_meic_tab,
        custom_builder=build_pss_meic_dated_row,
        wip_workers_from="NS_metrics",
        completed_hours_from="NS_metrics",
    ),
    "PSS US": TeamSource(
        team="PSS US",
        xlsx=Path(r"C:\Users\wadec8\Medtronic PLC\PSS Sharepoint - Documents\PSS_US_Heijunka.xlsm"),
        week_from_sheet=week_from_pss_us_tab,
        custom_builder=build_pss_us_row,
        wip_workers_from="NS_metrics",
        completed_hours_from="NS_metrics",
    ),
    "PSS Intern": TeamSource(
        team="PSS Intern",
        xlsx=Path(r"C:\Users\wadec8\Medtronic PLC\PSS Sharepoint - Documents\PSS MEIC_Interns Heijunka.xlsm"),
        week_from_sheet=week_from_pss_meic_tab,
        custom_builder=build_pss_meic_dated_row,
        wip_workers_from="NS_metrics",
        completed_hours_from="NS_metrics",
    ),
    "Spine": TeamSource(
        team="Spine",
        xlsx=Path(r"C:\Users\wadec8\Medtronic PLC\MEIC - RTG - Documents\Spine_Heijunka.xlsm"),
        week_from_sheet=week_from_spine_tab,
        custom_builder=build_spine_row,
        wip_workers_from="NS_metrics",
        completed_hours_from="NS_metrics",
    ),
    "DBS": TeamSource(
        team="DBS",
        xlsx=Path(r"C:\Users\wadec8\Medtronic PLC\DBS CQ Team - Documents\DBS NON WIP.xlsx"),
        layout=StandardLayout(
            people_start_row=3, totals_row=17,
            activity_header_row=2, activity_start_col=3, activity_end_col=35,
            min_rows=18, min_cols=3,
        ),
        week_from_sheet=week_from_sheetname_date,
        wip_workers_from="NS_WIP",
        completed_hours_from="NS_WIP",
    ),
    "SCS": TeamSource(
        team="SCS",
        xlsx=Path(r"C:\Users\wadec8\Medtronic PLC\Customer Quality SCS - Cell 17\SCS Non-D2D WIP Tracker 2026.xlsx"),
        layout=StandardLayout(
            people_start_row=2, totals_row=27,
            activity_header_row=1, activity_start_col=3, activity_end_col=36,
            min_rows=26, min_cols=3,
        ),
        week_from_sheet=week_from_sheetname_date,
        wip_workers_from="NS_WIP",
        completed_hours_from="NS_WIP",
    ),
    "TDD": TeamSource(
        team="TDD",
        xlsx=Path(r"C:\Users\wadec8\Medtronic PLC\RTG Customer Quality - Infusion - Documents\Non-D2D WIP Tracker TDD.xlsx"),
        layout=StandardLayout(
            people_start_row=2,      
            totals_row=20,           
            activity_header_row=1,   
            activity_start_col=3,    
            activity_end_col=34,     
            min_rows=21,
            min_cols=35,
        ),
        week_from_sheet=week_from_sheetname_date,
        wip_workers_from="NS_WIP",
        completed_hours_from="NS_WIP",
    ),
    "PH": TeamSource(
        team="PH",
        xlsx=Path(r"C:\Users\wadec8\Medtronic PLC\Customer Quality Pelvic Health - Other\PH Non-D2D WIP.xlsx"),
        layout=StandardLayout(
            people_start_row=2, totals_row=17,
            activity_header_row=1, activity_start_col=3, activity_end_col=37,
            min_rows=17, min_cols=3,
        ),
        week_from_sheet=week_from_sheetname_date,
        wip_workers_from="NS_WIP",
        completed_hours_from="NS_WIP",
    ),
    "NV": TeamSource(
        team="NV",
        xlsx=Path(r"C:\Users\wadec8\Medtronic PLC\RTG Customer Quality Neurovascular - Documents\Cell\NV_Heijunka.xlsm"),
        week_from_sheet=week_from_nv_tab,
        custom_builder=build_nv_row,
        wip_workers_from="NS_metrics",
        completed_hours_from="NS_metrics",
    ),
    "Nav": TeamSource(
        team="Nav",
        xlsx=Path(r"C:\Users\wadec8\Medtronic PLC\MNAV Sharepoint - Navigation Work Reports\Heijunka_MNAV_Ranges_May2025.xlsm"),
        week_from_sheet=week_from_mnav_capacity_tab,
        custom_builder=build_mnav_row,
        wip_workers_from="NS_metrics",
        completed_hours_from="NS_metrics",
    ),
    "AE MEIC": TeamSource(
        team="AE MEIC",
        xlsx=Path(r"C:\Users\wadec8\Medtronic PLC\MNAV Sharepoint - MEIC AE + OARM\AE_MEIC_Heijunka.xlsm"),
        week_from_sheet=week_from_mnav_capacity_tab,
        custom_builder=build_ae_meic_row,
        wip_workers_from="NS_metrics",
        completed_hours_from="NS_metrics",
    ),
    "O-Arm MEIC": TeamSource(
        team="O-Arm MEIC",
        xlsx=Path(r"C:\Users\wadec8\Medtronic PLC\MNAV Sharepoint - MEIC AE + OARM\OARM_MEIC_Heijunka.xlsm"),
        week_from_sheet=week_from_mnav_capacity_tab,
        custom_builder=build_oarm_meic_row,
        wip_workers_from="NS_metrics",
        completed_hours_from="NS_metrics",
    ),
    "Mazor": TeamSource(
        team="Mazor",
        xlsx=Path(r"C:\Users\wadec8\Medtronic PLC\MNAV Sharepoint - Caesarea Team\CAE - Heijunka_v2.xlsm"),
        week_from_sheet=week_from_mnav_capacity_tab,
        custom_builder=build_mazor_row,
        wip_workers_from="NS_metrics",
        completed_hours_from="NS_metrics",
    ),
    "CSF": TeamSource(
        team="CSF",
        xlsx=Path(r"c:\Users\wadec8\Medtronic PLC\CQ CSF Management - Documents\CSF_Heijunka.xlsm"),
        week_from_sheet=week_from_mnav_capacity_tab,
        custom_builder=build_csf_row,
        wip_workers_from="NS_metrics",
        completed_hours_from="NS_metrics",
    ),
    "ENT": TeamSource(
        team="ENT",
        xlsx=Path(r"C:\Users\wadec8\Medtronic PLC\ENT GEMBA Board - Capacity Management\ENT_Capacity Management for Non WIP_March 9th.xlsm"),
        week_from_sheet=week_from_ent_tab,
        custom_builder=build_ent_row,
        wip_workers_from="NS_metrics",
        completed_hours_from="NS_metrics",
    ),
    "DBS MEIC": TeamSource(
        team="DBS MEIC",
        xlsx=MEIC_TRACKER_PATH,
        wip_workers_from="NS_metrics",
        completed_hours_from="NS_metrics",
    ),
    "SCS MEIC": TeamSource(
        team="SCS MEIC",
        xlsx=MEIC_TRACKER_PATH,
        wip_workers_from="NS_metrics",
        completed_hours_from="NS_metrics",
    ),
    "PH MEIC": TeamSource(
        team="PH MEIC",
        xlsx=MEIC_TRACKER_PATH,
        wip_workers_from="NS_metrics",
        completed_hours_from="NS_metrics",
    ),
}
def build_nonwip_by_person_b_minus_c(people_rows: List[dict]) -> Dict[str, float]:
    out: Dict[str, float] = {}
    for r in people_rows:
        v = float(round(float(r.get("B", 0.0)) - float(r.get("C", 0.0)), 2))
        if v == 0.0:
            continue
        out[r["name"]] = v
    return out
def build_team_rows(team_src: TeamSource, wip_df: pd.DataFrame, metrics_df: pd.DataFrame) -> pd.DataFrame:
    if team_src.team in {"DBS MEIC", "SCS MEIC", "PH MEIC"}:
        if team_src.team != "DBS MEIC":
            return pd.DataFrame()
        return build_meic_rows_from_team_tracker(team_src.xlsx, wip_df=wip_df, metrics_df=metrics_df)
    xlsx_path = team_src.xlsx
    if not xlsx_path.exists():
        print(f"[WARN] Missing XLSX for {team_src.team}: {xlsx_path}")
        return pd.DataFrame()
    sheets = pd.read_excel(xlsx_path, sheet_name=None, header=None, engine="openpyxl")
    out_rows: List[dict] = []
    for sheet_name, ws in sheets.items():
        if team_src.week_from_sheet is None:
            continue
        week = team_src.week_from_sheet(sheet_name, ws)
        if week is None or pd.isna(week):
            continue
        if team_src.team in {"PSS US", "PSS MEIC"}:
            pss_min_date = pd.Timestamp("2026-02-01").normalize()
            pss_max_date = pd.Timestamp.today().normalize()
            if not (pss_min_date <= week.normalize() <= pss_max_date):
                continue
        if team_src.custom_builder is not None:
            built = team_src.custom_builder(team_src.team, ws, week)
            people_count = built["people_count"]
            total_nonwip_hours = built["total_nonwip_hours"]
            ooo_hours = built["ooo_hours"]
            nonwip_by_person = built["nonwip_by_person"]
            nonwip_activities = built["nonwip_activities"]
            ooo_map = built["ooo_map"]
        else:
            cfg = team_src.layout
            if cfg is None:
                continue
            if ws.shape[0] < cfg.min_rows or ws.shape[1] < cfg.min_cols:
                continue
            people_rows = read_people_block(
                ws,
                start_row_i=cfg.people_start_row,
                end_row_i=cfg.totals_row - 1,
                team=team_src.team,
                sheet_name=sheet_name,
                week=week,
            )
            people_count = len(set(r["name"] for r in people_rows))
            b = safe_float(ws.iat[cfg.totals_row, 1] if ws.shape[1] > 1 else np.nan)
            c = safe_float(ws.iat[cfg.totals_row, 2] if ws.shape[1] > 2 else np.nan)
            total_nonwip_hours = (b - c) if pd.notna(b) and pd.notna(c) else np.nan
            ooo_hours = c if pd.notna(c) else np.nan
            nonwip_by_person = build_nonwip_by_person_b_minus_c(people_rows)
            nonwip_activities = build_activities(
                ws, people_rows,
                header_row_i=cfg.activity_header_row,
                start_col_i=cfg.activity_start_col,
                end_col_i=cfg.activity_end_col,
            )
            ooo_map = {r["name"]: float(r.get("C", 0.0)) for r in people_rows}
        if team_src.completed_hours_from == "NS_metrics":
            completed_src_df = metrics_df
        else:
            completed_src_df = wip_df
        completed_match = completed_src_df[
            (completed_src_df.get("team") == team_src.team) &
            (completed_src_df["period_date"] == week)
        ]
        completed_hours = (
            pd.to_numeric(completed_match.iloc[0].get("Completed Hours"), errors="coerce")
            if not completed_match.empty else np.nan
        )
        pct_in_wip = np.nan
        if pd.notna(completed_hours) and pd.notna(total_nonwip_hours):
            denom = float(completed_hours) + float(total_nonwip_hours)
            pct_in_wip = float(completed_hours) / denom if denom != 0 else np.nan
        if team_src.wip_workers_from == "NS_metrics":
            wip_source_df = metrics_df
        else:
            wip_source_df = wip_df
        wip_match = wip_source_df[(wip_source_df.get("team") == team_src.team) & (wip_source_df["period_date"] == week)]
        wip_workers = extract_wip_workers_from_row(wip_match.iloc[0]) if not wip_match.empty else []
        wip_workers_count = len(wip_workers)
        wip_workers_ooo_hours = float(round(sum(safe_float0(ooo_map.get(n, 0.0)) for n in wip_workers), 2))
        use_original_people_count_teams = {"SCS", "TDD", "NV", "PH", "ENT"}
        if team_src.team == "DBS":
            people_count_final = get_dbs_people_count_from_heijunka_files(
                file_paths=(DBS_C13_SOURCE_FILE, DBS_C14_SOURCE_FILE),
                name_row_zero_based=29,   # Excel row 30
            )
        elif team_src.team in use_original_people_count_teams:
            people_count_final = int(people_count)
        else:
            people_count_final = get_people_count_from_wip(
                wip_df=wip_df,
                team=team_src.team,
                week=week,
                fallback=people_count,
            )
        out_rows.append({
            "team": team_src.team,
            "period_date": week.date().isoformat(),
            "source_file": str(xlsx_path),
            "people_count": int(people_count_final),
            "total_non_wip_hours": float(round(total_nonwip_hours, 2)) if pd.notna(total_nonwip_hours) else np.nan,
            "OOO Hours": float(round(ooo_hours, 2)) if pd.notna(ooo_hours) else np.nan,
            "% in WIP": float(round(pct_in_wip, 6)) if pd.notna(pct_in_wip) else np.nan,
            "non_wip_by_person": json.dumps(nonwip_by_person, ensure_ascii=False),
            "non_wip_activities": json.dumps(nonwip_activities, ensure_ascii=False),
            "wip_workers": json.dumps(wip_workers, ensure_ascii=False),
            "wip_workers_count": int(wip_workers_count),
            "wip_workers_ooo_hours": float(wip_workers_ooo_hours),
        })
    df = pd.DataFrame(out_rows)
    if not df.empty:
        df["period_date"] = pd.to_datetime(df["period_date"], errors="coerce").dt.normalize()
        df = df.sort_values(["team", "period_date"]).reset_index(drop=True)
    return df
def main():
    if not NS_WIP_PATH.exists():
        raise FileNotFoundError(f"Missing NS_WIP.csv: {NS_WIP_PATH}")
    if not NS_METRICS_PATH.exists():
        raise FileNotFoundError(f"Missing NS_metrics.csv: {NS_METRICS_PATH}")
    wip_df = load_csv(NS_WIP_PATH)
    metrics_df = load_metrics(NS_METRICS_PATH)
    built: List[pd.DataFrame] = []
    for team, src in TEAM_SOURCES.items():
        df_team = build_team_rows(src, wip_df=wip_df, metrics_df=metrics_df)
        if not df_team.empty:
            built.append(df_team)
    new_df = pd.concat(built, ignore_index=True) if built else pd.DataFrame()
    if OUT_PATH.exists():
        old_df = load_csv(OUT_PATH)
        old_df = old_df[
            ~old_df["team"].isin({ENABLE_TEAM_NAME, "PH", "DBS", "SCS"})
        ].copy()
        combined = pd.concat([old_df, new_df], ignore_index=True)
        combined["period_date"] = pd.to_datetime(combined["period_date"], errors="coerce").dt.normalize()
        combined = combined.drop_duplicates(subset=["team", "period_date"], keep="last")
        combined = combined.sort_values(["team", "period_date"]).reset_index(drop=True)
    else:
        combined = new_df
    combined = combine_enabling_technologies(combined, wip_df=wip_df)
    combined = combine_meic_parent_teams(combined, wip_df=wip_df)
    combined.to_csv(OUT_PATH, index=False, encoding="utf-8-sig")
    print(f"Wrote {len(combined)} rows -> {OUT_PATH}")
if __name__ == "__main__":
    main()