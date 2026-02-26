import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple
import numpy as np
import pandas as pd
NS_WIP_PATH = Path(r"C:\heijunka-dev\NS_WIP.csv")
NS_METRICS_PATH = Path(r"C:\heijunka-dev\NS_metrics.csv")
OUT_PATH = Path(r"C:\heijunka-dev\ns_non_wip_activities.csv")
BAD_NAMES = {
    "", "-", "–", "—", "nan", "NaN", "NAN",
    "n/a", "N/A", "na", "NA", "null", "NULL",
    "none", "None", "tm", "TM", "Totals", "TOTALS",
    "Team Hours Available", "TEAM HOURS AVAILABLE",
    "Mazor Hours Available", "MAZOR HOURS AVAILABLE",
}
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
def read_people_block(ws: pd.DataFrame, start_row_i: int) -> List[dict]:
    rows: List[dict] = []
    for i in range(start_row_i, len(ws)):
        name = norm_name(ws.iat[i, 0] if ws.shape[1] > 0 else "")
        if not name:
            break
        if not is_real_person(name):
            continue
        b = safe_float(ws.iat[i, 1] if ws.shape[1] > 1 else np.nan)
        c = safe_float(ws.iat[i, 2] if ws.shape[1] > 2 else np.nan)
        if pd.isna(b): b = 0.0
        if pd.isna(c): c = 0.0
        rows.append({"row_i": i, "name": name, "B": b, "C": c})
    return rows
def build_nonwip_by_person_b_minus_c(people_rows: List[dict]) -> Dict[str, float]:
    out: Dict[str, float] = {}
    for r in people_rows:
        v = float(round(float(r.get("B", 0.0)) - float(r.get("C", 0.0)), 2))
        if v == 0.0:
            continue
        out[r["name"]] = v
    return out
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
            activities.append({"name": name, "activity": label, "hours": float(round(hrs, 2))})
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
    wip_workers_from: str = "NS_WIP"          # where Person Hours comes from
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
def week_from_mnav_capacity_tab(sheet_name: str, ws: pd.DataFrame) -> Optional[pd.Timestamp]:
    s = str(sheet_name)
    if not s.lower().startswith("capacity mgmt"):
        return None
    try:
        b1 = ws.iat[0, 1]  # row 1 col B (0-indexed)
        dt = pd.to_datetime(b1, errors="coerce")
        if _is_real_year(dt):
            return dt.normalize()
    except Exception:
        pass
    m = re.search(r"\((\d{1,2})\.(\d{1,2})\)", s)
    if not m:
        return None
    mm = int(m.group(1))
    dd = int(m.group(2))
    for r in range(0, 6):
        for c in range(0, 6):
            try:
                v = ws.iat[r, c]
            except Exception:
                continue
            dt = pd.to_datetime(v, errors="coerce")
            if _is_real_year(dt):
                return pd.Timestamp(year=int(dt.year), month=mm, day=dd).normalize()
    return pd.Timestamp(year=DEFAULT_YEAR_IF_MISSING, month=mm, day=dd).normalize()
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
    return {
        "people_rows": people_rows,
        "people_count": people_count,
        "ooo_hours": ooo_hours,
        "total_nonwip_hours": total_nonwip_hours,
        "nonwip_by_person": nonwip_by_person,
        "nonwip_activities": activities,
        "ooo_map": {r["name"]: float(r["OOO"]) for r in people_rows},
    }
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
    expected_col_letter: str,    # column with "Expected Number of WIP Hours Per Week" (B)
    ooo_col_letter: str,         # OOO column (Q or Z)
    deduct_cell: str,            # e.g. "B8", "B11", "B10"
    ooo_sum_start_row: int,      # inclusive, 0-indexed row
    ooo_sum_end_row: int,        # inclusive, 0-indexed row
    total_ooo_end_row: int,      # inclusive, used for Total Non-WIP formula (sometimes differs)
    activity_header_row: int,    # row 1 => index 0, row 2 => index 1, etc
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
    deduct_row = int(m.group(2)) - 1  # Excel -> 0-indexed
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
        ooo_hours += safe_float(ws.iat[r, col_ooo] if ws.shape[1] > col_ooo and ws.shape[0] > r else 0.0) or 0.0
    ooo_hours = float(round(ooo_hours, 2))
    deduct_val = safe_float(ws.iat[deduct_row, deduct_col] if ws.shape[0] > deduct_row and ws.shape[1] > deduct_col else 0.0)
    if pd.isna(deduct_val):
        deduct_val = 0.0
    total_ooo = 0.0
    for r in range(ooo_sum_start_row, total_ooo_end_row + 1):
        total_ooo += safe_float(ws.iat[r, col_ooo] if ws.shape[1] > col_ooo and ws.shape[0] > r else 0.0) or 0.0
    total_ooo = float(total_ooo)
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
    return {
        "people_rows": people_rows,
        "people_count": people_count,
        "ooo_hours": ooo_hours,
        "total_nonwip_hours": total_nonwip_hours,
        "nonwip_by_person": nonwip_by_person,
        "nonwip_activities": activities,
        "ooo_map": {r["name"]: float(r["OOO"]) for r in people_rows},
    }
def build_ae_meic_row(team: str, ws: pd.DataFrame, week: Optional[pd.Timestamp] = None) -> Dict:
    return build_capacity_fixed_row(
        team, ws,
        people_start_row=1, people_end_row=5,
        expected_col_letter="B",
        ooo_col_letter="Q",
        deduct_cell="B8",
        ooo_sum_start_row=1, ooo_sum_end_row=5,     # Q2:Q6
        total_ooo_end_row=5,                        # Total uses Q2:Q6
        activity_header_row=0,                      # row 1
        activity_start_col_letter="C",
        activity_end_col_letter="P",
    )
def build_oarm_meic_row(team: str, ws: pd.DataFrame, week: Optional[pd.Timestamp] = None) -> Dict:
    return build_capacity_fixed_row(
        team, ws,
        people_start_row=1, people_end_row=8,
        expected_col_letter="B",
        ooo_col_letter="Q",
        deduct_cell="B11",
        ooo_sum_start_row=1, ooo_sum_end_row=8,     # Q2:Q9
        total_ooo_end_row=8,                        # Total uses Q2:Q9
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
        ooo_sum_start_row=1, ooo_sum_end_row=8,     # Z2:Z9 (OOO)
        total_ooo_end_row=7,                        # Z2:Z8 (Total Non-WIP)
        activity_header_row=0,
        activity_start_col_letter="C",
        activity_end_col_letter="Y",
    )
TEAM_SOURCES: Dict[str, TeamSource] = {
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
            people_start_row=2, totals_row=25,
            activity_header_row=1, activity_start_col=3, activity_end_col=34,
            min_rows=26, min_cols=3,
        ),
        week_from_sheet=week_from_sheetname_date,
        wip_workers_from="NS_WIP",
        completed_hours_from="NS_WIP",
    ),
    "PH": TeamSource(
        team="PH",
        xlsx=Path(r"C:\Users\wadec8\Medtronic PLC\Customer Quality Pelvic Health - Daily Tracker\Non-D2D WIP Tracker.xlsx"),
        layout=StandardLayout(
            people_start_row=2, totals_row=18,
            activity_header_row=1, activity_start_col=3, activity_end_col=34,
            min_rows=17, min_cols=3,
        ),
        week_from_sheet=week_from_sheetname_date,
        wip_workers_from="NS_WIP",
        completed_hours_from="NS_WIP",
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
}
def build_team_rows(team_src: TeamSource, wip_df: pd.DataFrame, metrics_df: pd.DataFrame) -> pd.DataFrame:
    xlsx_path = team_src.xlsx
    if not xlsx_path.exists():
        print(f"[WARN] Missing XLSX for {team_src.team}: {xlsx_path}")
        return pd.DataFrame()
    sheets = pd.read_excel(xlsx_path, sheet_name=None, header=None)
    out_rows: List[dict] = []
    for sheet_name, ws in sheets.items():
        if team_src.week_from_sheet is None:
            continue
        week = team_src.week_from_sheet(sheet_name, ws)
        if week is None or pd.isna(week):
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
            people_rows = read_people_block(ws, start_row_i=cfg.people_start_row)
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
        wip_workers_ooo_hours = float(round(sum(float(ooo_map.get(n, 0.0) or 0.0) for n in wip_workers), 2))
        out_rows.append({
            "team": team_src.team,
            "period_date": week.date().isoformat(),
            "source_file": str(xlsx_path),
            "people_count": int(people_count),
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
        combined = pd.concat([old_df, new_df], ignore_index=True)
        combined["period_date"] = pd.to_datetime(combined["period_date"], errors="coerce").dt.normalize()
        combined = combined.drop_duplicates(subset=["team", "period_date"], keep="last")
        combined = combined.sort_values(["team", "period_date"]).reset_index(drop=True)
    else:
        combined = new_df
    combined.to_csv(OUT_PATH, index=False, encoding="utf-8-sig")
    print(f"Wrote {len(combined)} rows -> {OUT_PATH}")
if __name__ == "__main__":
    main()