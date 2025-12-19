import os
import re
import time
from datetime import datetime, date, timedelta
import pandas as pd
import win32com.client as win32
WORKBOOK_PATH = r"C:\Users\wadec8\OneDrive - Medtronic PLC\DSA-MDT-RPT-W-Go Green Initiative Monitor.xlsx"
SHEET_NAME = "WIP Tracking"
OUT_DIR = r"C:\Users\wadec8\OneDrive - Medtronic PLC"
OUT_BASENAME = "wip_tracking_long"
REFRESH_TIMEOUT_SECONDS = 1200
TEAM_LIST_TEAM_COL = "team"
CLOSURES_CSV_PATH = r"C:\heijunka-dev\closures.csv"
CTTYPE_FIELD_CANDIDATES = [
    "[WipIO].[tType].[tType]",
]
CTTYPE_VALUE = "PE"
FMONTHREL_FIELD_MDX = "[Calendar].[fMonthRel].[fMonthRel]"
FMONTHREL_ALLOWED = {0, -1}
PORTFOLIO_FIELD_CANDIDATES = [
    "[BusinessMap].[Operating Groups].[Portfolio]",
    "[BusinessMap].[Operating Groups].[Product Group]",
    "[BusinessMap].[Operating Groups].[Operating Unit]",
    "[BusinessMap].[Operating Groups].[Integrated Operating Unit]",
]
INCOMING_HEADER_HINTS = ["incoming", "pe"]
CLOSED_TOTAL_HEADER_HINTS = ["closed", "total"]
EXCLUDE_TEAMS = {
    "Structural Heart and Aortic",
    "Neuroscience",
    "Medical Surgical",
    "Mechanical Circulatory Support",
    "Cardiovascular",
    "Cardiac Surgery",
    "Cardiac Rhythm Management",   
}
TEAM_MAP = {
    "Cardiac Ablation Solutions": "CAS",
    "Coronary & Renal Denervation": "CRDN",
    "Peripheral Vascular Health": "PVH",
    "Structural Heart": "TCT",
}
TEAM_MAP_NORM = {k.strip().lower(): v for k, v in TEAM_MAP.items()}
def map_team_name(team: str) -> str:
    t = team.strip()
    return TEAM_MAP_NORM.get(t.lower(), t)
EXCLUDE_TEAMS_NORM = {t.strip().lower() for t in EXCLUDE_TEAMS}
def is_excluded_team(team: str) -> bool:
    return team.strip().lower() in EXCLUDE_TEAMS_NORM
def excel_serial_to_date(n: float) -> date:
    return (datetime(1899, 12, 30) + timedelta(days=float(n))).date()
def week_monday(d: date) -> date:
    return d - timedelta(days=d.weekday())
def parse_period_value(x):
    if x is None:
        return None
    if isinstance(x, datetime):
        return x.date()
    if isinstance(x, date) and not isinstance(x, datetime):
        return x
    if isinstance(x, (int, float)) and x > 20000:
        try:
            return excel_serial_to_date(x)
        except Exception:
            return None
    s = str(x).strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%m/%d/%y", "%d-%b-%Y", "%d-%b-%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            pass
    return None
def load_teams_from_csv(path: str) -> list[str]:
    df = pd.read_csv(path, dtype={TEAM_LIST_TEAM_COL: str})
    teams = (
        df[TEAM_LIST_TEAM_COL]
        .dropna()
        .astype(str)
        .str.strip()
        .loc[lambda s: s.ne("")]
        .unique()
        .tolist()
    )
    return sorted(teams)
def _escape_mdx_member(v: str) -> str:
    return v.replace("]", "]]")
def select_item_on_field_guess(pf, caption: str):
    cap = caption.strip()
    if not cap:
        return False
    cap_mdx = _escape_mdx_member(cap)
    field = str(pf.Name)
    candidates = [
        f"{field}.&[{cap_mdx}]",
        f"{field}.&[{cap_mdx.upper()}]",
        f"{field}.&[{cap_mdx.lower()}]",
        f"{field}.[{pf.Caption}].&[{cap_mdx}]",
    ]
    try:
        pf.EnableMultiplePageItems = True
    except Exception:
        pass
    last_err = None
    for u in candidates:
        try:
            pf.ClearAllFilters()
        except Exception:
            pass
        try:
            pf.VisibleItemsList = [u]
            return True
        except Exception as e:
            last_err = e
    return False
def range_to_df(rng):
    vals = rng.Value
    if vals is None:
        return pd.DataFrame()
    if not isinstance(vals, (tuple, list)):
        return pd.DataFrame([[vals]])
    if vals and not isinstance(vals[0], (tuple, list)):
        return pd.DataFrame([list(vals)])
    return pd.DataFrame([list(r) for r in vals])
def norm_text(x) -> str:
    return re.sub(r"\s+", " ", str(x).replace("\n", " ").strip().lower())
def safe_float(x):
    if x is None:
        return None
    if isinstance(x, str) and not x.strip():
        return None
    try:
        return float(x)
    except Exception:
        return None
def cube_field_members_from_pivot_field(pf):
    def _items_to_members(pis):
        out = []
        for i in range(1, pis.Count + 1):
            pi = pis(i)
            cap = str(getattr(pi, "Caption", getattr(pi, "Name", ""))).strip()
            uniq = str(getattr(pi, "Name", "")).strip()  # OLAP unique name is usually pi.Name
            if not cap:
                continue
            if cap.lower() in ("(all)", "all"):
                continue
            out.append((cap, uniq))
        return out
    try:
        pis = pf.PivotItems()
        if pis.Count > 0:
            return _items_to_members(pis)
    except Exception:
        pass
    try:
        cf = pf.CubeField
        try:
            pis = cf.PivotItems()      # some Excel builds
        except Exception:
            pis = cf.PivotItems       # some expose as property
        if pis.Count > 0:
            return _items_to_members(pis)
    except Exception as e:
        raise RuntimeError(f"Could not enumerate members for '{pf.Name}': {e}")
    return []
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
def find_best_pivot(ws, must_contain_text: str):
    if ws.PivotTables().Count == 1:
        return ws.PivotTables(1)
    for i in range(1, ws.PivotTables().Count + 1):
        pt = ws.PivotTables(i)
        try:
            df = range_to_df(pt.TableRange2).head(30).astype(str)
            if df.apply(lambda c: c.str.contains(must_contain_text, case=False, na=False)).any().any():
                return pt
        except Exception:
            continue
    return ws.PivotTables(1)
def list_pivot_fields(pt):
    names = []
    try:
        for pf in pt.PivotFields():
            try:
                names.append(str(pf.Name))
            except Exception:
                pass
    except Exception:
        pass
    return sorted(set(names))
def get_pivot_field(pt, candidates):
    for name in candidates:
        try:
            return pt.PivotFields(name)
        except Exception:
            continue
    avail = []
    try:
        avail = list_pivot_fields(pt)
    except Exception:
        pass
    cand_l = [c.lower() for c in candidates]
    for a in avail:
        al = a.lower()
        if al in cand_l:
            try:
                return pt.PivotFields(a)
            except Exception:
                pass
    return None
def set_fmonthrel_strict(pt, allowed_values):
    pf = None
    try:
        pf = pt.PivotFields(FMONTHREL_FIELD_MDX)
    except Exception:
        raise RuntimeError(f"Could not access fMonthRel field as '{FMONTHREL_FIELD_MDX}'")
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
def pick_best_field_by_item_count(pt, candidates):
    best_pf = None
    best_count = -1
    for name in candidates:
        try:
            pf = pt.PivotFields(name)
        except Exception:
            continue
        try:
            cnt = 0
            for pi in pf.PivotItems():
                cap = str(getattr(pi, "Caption", pi.Name)).strip()
                if not cap or cap.lower() in ("(all)", "all"):
                    continue
                cnt += 1
            if cnt > best_count:
                best_count = cnt
                best_pf = pf
        except Exception:
            continue
    return best_pf
def iter_pivot_items(pf):
    try:
        pis = pf.PivotItems()   # usual COM pattern
    except Exception:
        pis = pf.PivotItems    # fallback
    try:
        cnt = int(pis.Count)
    except Exception:
        cnt = 0
    for i in range(1, cnt + 1):
        try:
            pi = pis(i)        # 1-based
        except Exception:
            pi = pis.Item(i)
        cap = str(getattr(pi, "Caption", pi.Name)).strip()
        nm = str(pi.Name).strip()
        yield cap, nm
def portfolio_leaf_items(pf):
    items = []
    for cap, nm in iter_pivot_items(pf):
        if not cap or cap.lower() in ("(all)", "all"):
            continue
        if cap in EXCLUDE_TEAMS:
            continue
        items.append((cap, nm))
    seen = set()
    out = []
    for cap, nm in items:
        key = cap.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append((cap, nm))
    return out
def select_portfolio_item(pt, portfolio_field_candidates, unique_name: str):
    pf = get_pivot_field(pt, portfolio_field_candidates)
    if pf is None:
        raise RuntimeError(f"Could not find Portfolio pivot field from candidates: {portfolio_field_candidates}")
    try:
        pf.ClearAllFilters()
    except Exception:
        pass
    try:
        pf.VisibleItemsList = [unique_name]
        return
    except Exception:
        pass
    try:
        for pi in pf.PivotItems():
            try:
                if str(pi.Name) == unique_name:
                    pf.CurrentPage = str(getattr(pi, "Caption", pi.Name))
                    return
            except Exception:
                continue
    except Exception:
        pass
    raise RuntimeError(f"Failed to select portfolio item: {unique_name}")
def find_metric_columns(df: pd.DataFrame):
    top = df.head(20)
    incoming_col = None
    closed_col = None
    for c in range(df.shape[1]):
        col_hits = " ".join(norm_text(x) for x in top.iloc[:, c].tolist())
        if incoming_col is None and all(h in col_hits for h in INCOMING_HEADER_HINTS):
            incoming_col = c
        if closed_col is None and all(h in col_hits for h in CLOSED_TOTAL_HEADER_HINTS):
            closed_col = c
    if incoming_col is None or closed_col is None:
        raise RuntimeError(
            f"Couldn't find metric columns. Found incoming_col={incoming_col}, closed_col={closed_col}.\n"
            f"Try adjusting INCOMING_HEADER_HINTS / CLOSED_TOTAL_HEADER_HINTS."
        )
    return incoming_col, closed_col
def extract_week_rows(df: pd.DataFrame, incoming_col: int, closed_col: int, team_name: str):
    records = []
    for r in range(df.shape[0]):
        d = None
        for c in range(min(6, df.shape[1])):
            d = parse_period_value(df.iat[r, c])
            if d is not None:
                break
        if d is None:
            continue
        inc = safe_float(df.iat[r, incoming_col])
        clo = safe_float(df.iat[r, closed_col])
        if inc is None and clo is None:
            continue
        records.append({
            "team": team_name,
            "period_date": week_monday(d).strftime("%Y-%m-%d"),
            "Opened": inc,
            "Closures": clo,
        })
    return records
def get_page_field(pt, want: str):
    want = want.strip().lower()
    try:
        pfs = pt.PageFields
    except Exception:
        pfs = pt.PageFields()
    for i in range(1, pfs.Count + 1):
        pf = pfs(i)
        cap = str(getattr(pf, "Caption", pf.Name)).strip()
        name = str(pf.Name).strip()
        print(f"PageField[{i}]: Caption='{cap}'  Name='{name}'")
    for i in range(1, pfs.Count + 1):
        pf = pfs(i)
        cap = str(getattr(pf, "Caption", pf.Name)).strip().lower()
        name = str(pf.Name).strip().lower()
        if cap == want or want in cap or want in name:
            return pf
    return None
def clean_label(x) -> str:
    s = "" if x is None else str(x)
    s = s.replace("\u00A0", " ").strip()
    return s
def expand_all_row_levels(pt, max_passes: int = 10, sleep: float = 0.25):
    xlRowField = 1
    xlColumnField = 2
    try:
        pt.ManualUpdate = True
    except Exception:
        pass
    for _ in range(max_passes):
        changed = False
        for pf in pt.PivotFields():
            try:
                orient = int(pf.Orientation)
            except Exception:
                continue
            if orient not in (xlRowField, xlColumnField):
                continue
            try:
                for pi in pf.PivotItems():
                    try:
                        if hasattr(pi, "ShowDetail"):
                            if not bool(pi.ShowDetail):
                                pi.ShowDetail = True
                                changed = True
                        elif hasattr(pi, "DrilledDown"):
                            if not bool(pi.DrilledDown):
                                pi.DrilledDown = True
                                changed = True
                    except Exception:
                        continue
            except Exception:
                continue
        try:
            pt.RefreshTable()
        except Exception:
            pass
        time.sleep(sleep)
        if not changed:
            break
    try:
        pt.ManualUpdate = False
    except Exception:
        pass
def extract_from_expanded_pivot(pt, teams_set=None):
    wide = range_to_df(pt.TableRange2)
    if wide.empty:
        return []
    incoming_col, closed_col = find_metric_columns(wide)
    month_map = {
        "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4,
        "MAY": 5, "JUN": 6, "JUL": 7, "AUG": 8,
        "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12,
    }
    def looks_like_fy(x):
        s = clean_label(x)
        if re.fullmatch(r"\d{4}", s):
            y = int(s)
            if 2000 <= y <= 2100:
                return y
        return None
    def fiscal_to_calendar_year(fy: int, month_num: int) -> int:
        return fy if month_num <= 4 else (fy - 1)
    def parse_week_without_year(s: str, fy: int):
        m = re.fullmatch(r"\s*(\d{1,2})[-/ ]([A-Za-z]{3})\s*", s)
        if not m or fy is None:
            return None
        day = int(m.group(1))
        mon = m.group(2).upper()
        if mon not in month_map:
            return None
        month_num = month_map[mon]
        year = fiscal_to_calendar_year(fy, month_num)
        try:
            return date(year, month_num, day)
        except Exception:
            return None
    def first_label_in_row(r):
        for c in range(min(6, wide.shape[1])):
            v = wide.iat[r, c]
            s = clean_label(v)
            if s:
                return s
        return ""
    records = []
    current_week = None
    current_fy = None
    for r in range(wide.shape[0]):
        fy = looks_like_fy(wide.iat[r, 0])
        if fy is not None:
            current_fy = fy
        d = None
        for c in range(min(6, wide.shape[1])):
            d = parse_period_value(wide.iat[r, c])
            if d is not None:
                break
        if d is None:
            lbl = first_label_in_row(r)
            d = parse_week_without_year(lbl, current_fy)
        if d is not None:
            current_week = week_monday(d).strftime("%Y-%m-%d")
            continue
        if current_week is None:
            continue
        team = first_label_in_row(r)
        if not team:
            continue
        if is_numeric_team_label(team):
            continue
        if team.strip().lower() in {"fiscal time groups", "enterprise", "all"}:
            continue
        if is_excluded_team(team):
            continue
        team = map_team_name(team)
        if teams_set is not None and team not in teams_set:
            continue
        inc = safe_float(wide.iat[r, incoming_col])
        clo = safe_float(wide.iat[r, closed_col])
        if inc is None and clo is None:
            continue
        records.append({
            "team": team,
            "period_date": current_week,
            "Opened": inc,
            "Closures": clo,
        })
    return records
def select_item_on_field(pf, unique_name: str):
    try:
        pf.EnableMultiplePageItems = True
    except Exception:
        pass
    try:
        pf.ClearAllFilters()
    except Exception:
        pass
    try:
        pf.VisibleItemsList = [unique_name]
        return
    except Exception:
        pass
    try:
        for pi in pf.PivotItems():
            if str(pi.Name) == unique_name:
                pf.CurrentPage = str(getattr(pi, "Caption", pi.Name))
                return
    except Exception:
        pass
    raise RuntimeError(f"Failed to select item on field '{pf.Name}': {unique_name}")
def is_numeric_team_label(x) -> bool:
    if x is None:
        return False
    if isinstance(x, (int, float)) and not isinstance(x, bool):
        return True
    s = clean_label(x)
    if not s:
        return False
    s2 = s.replace(",", "")
    return re.fullmatch(r"-?\d+(\.\d+)?", s2) is not None
def upsert_closures_csv(path: str, new_df: pd.DataFrame):
    if new_df is None or new_df.empty:
        return 0, 0
    in_df = new_df.copy()
    col_l = {c.lower(): c for c in in_df.columns}
    def _need(colname):
        if colname not in col_l:
            raise RuntimeError(f"new_df missing required column '{colname}'. Have: {list(in_df.columns)}")
        return col_l[colname]
    team_col = _need("team")
    date_col = _need("period_date")
    closures_col = col_l.get("closures")
    opened_col = col_l.get("opened")
    if closures_col is None or opened_col is None:
        raise RuntimeError(f"new_df must have 'Opened' and 'Closures'. Have: {list(in_df.columns)}")
    in_df = in_df.rename(columns={
        team_col: "team",
        date_col: "period_date",
        opened_col: "Opened",
        closures_col: "Closures",
    })
    in_df["team"] = in_df["team"].astype(str).str.strip()
    in_df["period_date"] = pd.to_datetime(in_df["period_date"], errors="coerce").dt.strftime("%Y-%m-%d")
    in_df = in_df.dropna(subset=["team", "period_date"])
    in_df = in_df[in_df["team"].ne("")]
    in_df = in_df.drop_duplicates(subset=["team", "period_date"], keep="last")
    if os.path.exists(path):
        cur = pd.read_csv(path, dtype={"team": str, "period_date": str})
    else:
        cur = pd.DataFrame(columns=["team", "period_date", "closures", "Opened"])
    cur_cols_l = {c.lower(): c for c in cur.columns}
    cur_team = cur_cols_l.get("team", "team")
    cur_date = cur_cols_l.get("period_date", "period_date")
    cur_opened = cur_cols_l.get("opened", "Opened")
    cur_closures = cur_cols_l.get("closures", "closures")
    cur_norm = cur.rename(columns={
        cur_team: "team",
        cur_date: "period_date",
        cur_opened: "Opened",
        cur_closures: "Closures",
    }).copy()
    cur_norm["team"] = cur_norm["team"].astype(str).str.strip()
    cur_norm["period_date"] = pd.to_datetime(cur_norm["period_date"], errors="coerce").dt.strftime("%Y-%m-%d")
    cur_norm = cur_norm.set_index(["team", "period_date"])
    in_norm = in_df.set_index(["team", "period_date"])
    common_idx = cur_norm.index.intersection(in_norm.index)
    before = cur_norm.loc[common_idx, ["Opened", "Closures"]].copy()
    cur_norm.update(in_norm[["Opened", "Closures"]])
    after = cur_norm.loc[common_idx, ["Opened", "Closures"]]
    updated_rows = int((before.ne(after)).any(axis=1).sum())
    missing_idx = in_norm.index.difference(cur_norm.index)
    added_rows = int(len(missing_idx))
    if added_rows:
        cur_norm = pd.concat([cur_norm, in_norm.loc[missing_idx, ["Opened", "Closures"]]], axis=0)
    out = cur_norm.reset_index()
    out = out.rename(columns={
        "Closures": cur_closures,  # likely "closures"
        "Opened": cur_opened,      # likely "Opened"
    })
    desired = [cur_team, cur_date, cur_closures, cur_opened]
    existing_order = [c for c in desired if c in out.columns] + [c for c in out.columns if c not in desired]
    out = out[existing_order]
    out = out.sort_values([cur_team, cur_date])
    out.to_csv(path, index=False)
    return updated_rows, added_rows
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
        IgnoreReadOnlyRecommended=True,
    )
    try:
        wait_for_refresh(excel, wb, timeout_seconds=REFRESH_TIMEOUT_SECONDS, poll=2)
        ws = wb.Worksheets(SHEET_NAME)
        pt = find_best_pivot(ws, must_contain_text="Incoming")
        try:
            pt.RefreshTable()
        except Exception:
            pass
        time.sleep(0.5)
        expand_all_row_levels(pt, max_passes=10, sleep=0.25)
        time.sleep(0.5)
        all_records = extract_from_expanded_pivot(pt, teams_set=None)
        out_df = pd.DataFrame(all_records)
        if out_df.empty:
            raise RuntimeError("Export produced 0 rows. (Is the pivot expanded + showing weeks/teams?)")
        out_df = out_df.sort_values(["team", "period_date"])
        updated, added = upsert_closures_csv(CLOSURES_CSV_PATH, out_df)
        print(f"Updated {updated} row(s), added {added} row(s) in {CLOSURES_CSV_PATH}")
        out_df.to_csv(out_path, index=False)
        print(f"Wrote: {out_path} ({len(out_df)} rows)")
    finally:
        wb.Close(SaveChanges=False)
        excel.Quit()
if __name__ == "__main__":
    main()
