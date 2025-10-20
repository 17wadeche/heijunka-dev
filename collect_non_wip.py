# collect_non_wip.py
import csv
import json
from pathlib import Path
from datetime import datetime as _dt, date as _date, timedelta
import re
import pandas as pd
from openpyxl import load_workbook
from dateutil import parser as dateparser
from openpyxl import load_workbook
_DAY_RANGES = {
    "Monday":    (7, 40),
    "Tuesday":   (42, 77),
    "Wednesday": (79, 118),
    "Thursday":  (120, 161),
    "Friday":    (163, 200),
}
TEAM_OOO_CFG = {
    "aortic":          {"sheet": "#12 Production Analysis",           "flag_col": "K"},
    "svt":             {"sheet": "#12 Production Analysis",           "flag_col": "K"},
    "crdn":            {"sheet": "#12 Production Analysis",           "flag_col": "K"},
    "ect":             {"sheet": "#12 Production Analysis",           "flag_col": "K"},
    "pvh":             {"sheet": "#12 Production Analysis",           "flag_col": "K"},
    "tct clinical":    {"sheet": "Clinical #12 Prod Analysis",        "flag_col": "L"},
    "tct commercial":  {"sheet": "Commercial #12 Prod Analysis",      "flag_col": "L"},
}
REPO_DIR = Path(r"C:\heijunka-dev")
REPO_CSV = REPO_DIR / "metrics_aggregate_dev.csv"
OUT_CSV  = REPO_DIR / "non_wip_activities.csv"
WEEKLY_HOURS_DEFAULT = 40.0
TEAM_CFG = {
    "aortic":         {"sheet_patterns": ["individual (wip non wip)"], "col": "A", "start": 1},
    "crdn":           {"sheet_patterns": ["individual (wip non wip)"], "col": "A", "start": 1},
    "ect":            {"sheet_patterns": ["individual (wip non wip)"], "col": "A", "start": 1},
    "pvh":            {"sheet_patterns": ["individual (wip non wip)"], "col": "A", "start": 1},
    "svt":            {"sheet_patterns": ["individual"],                "col": "A", "start": 1},
    "tct commercial": {"sheet_patterns": ["individual (wip non wip)"], "col": "A", "start": 1},
    "tct clinical":   {"sheet_patterns": ["individual (wip non wip)"], "col": "Z", "start": 1},
    "ph":             {"people_from": "person_hours"},
    "cas":            {"people_from": "person_hours"},
}
try:
    from pyxlsb import open_workbook as open_xlsb
except Exception:
    open_xlsb = None
def _excel_serial_to_date(n):
    try:
        return (_dt(1899, 12, 30) + timedelta(days=float(n))).date()
    except Exception:
        return None
def _coerce_to_date(v):
    if isinstance(v, _dt): return v.date()
    if isinstance(v, _date): return v
    if isinstance(v, (int, float)): return _excel_serial_to_date(v)
    s = str(v).strip()
    if not s: return None
    try:
        return dateparser.parse(s, dayfirst=False, yearfirst=False).date()
    except Exception:
        return None
_BAD_EXACT = {
    "team member",
    "tuesday",
    "4",
    "5",
    "6",
    "7",
    "commercial weeks production output",
    "clinical weeks production output",
    "0.0",
    "open",
    "0",
    "2025-10-06 00:00:00",
    "team member 1",
    "team member 2",
    "team member 3",
    "team member 4",
    "total available hours",
    "total pitches",
    "weeks production output",
    "workflow",
}
_BAD_PREFIX = ("release date", "revision", "week starting")
def _looks_like_bad_header(raw: str) -> bool:
    if raw is None:
        return True
    s = str(raw).strip()
    if not s:
        return True
    s_nf = s.rstrip(":").strip()
    k = s_nf.casefold()
    if k in _BAD_EXACT:
        return True
    for pref in _BAD_PREFIX:
        if k.startswith(pref):
            return True
    if k in {"#ref!", "nan", "0", "-", "–", "—"}:
        return True
    return False
def _clean_name(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip().strip('"').strip("'")
    return "" if _looks_like_bad_header(s) else s
def _col_letter_to_index(letter: str) -> int:
    letter = str(letter).strip().upper()
    if not letter:
        return 1
    acc = 0
    for ch in letter:
        if not ("A" <= ch <= "Z"):
            continue
        acc = acc * 26 + (ord(ch) - ord("A") + 1)
    return max(1, acc)
def _norm_sheet_name(s: str) -> str:
    s = (s or "").lower()
    s = s.replace("–", "-").replace("—", "-")        # normalize dashes
    s = s.replace("(", " ").replace(")", " ")
    s = s.replace("_", " ").replace("-", " ")
    s = re.sub(r"\s+", " ", s).strip()
    s = s.replace("wip non wip", "wip non wip")
    return s
def _norm_title(s: str) -> str:
    s = (s or "").strip().lower()
    s = s.replace("–", "-").replace("—", "-")
    return " ".join(s.split())
def _resolve_sheet_exact_or_fuzzy_openpyxl(wb, desired_title: str) -> str | None:
    want = _norm_title(desired_title)
    m = {_norm_title(n): n for n in wb.sheetnames}
    if want in m:
        return m[want]
    for actual in wb.sheetnames:
        ns = _norm_title(actual)
        if (want in ns) or ns.startswith(want) or ns.endswith(want):
            return actual
    return None
def _resolve_sheet_exact_or_fuzzy_xlsb(wb, desired_title: str) -> str | None:
    want = _norm_title(desired_title)
    names = list(getattr(wb, "sheets", []) or [])
    m = {_norm_title(n): n for n in names}
    if want in m:
        return m[want]
    for actual in names:
        ns = _norm_title(actual)
        if (want in ns) or ns.startswith(want) or ns.endswith(want):
            return actual
    return None
def extract_ooo_per_day(xlsx_path: Path, sheet_title: str, flag_col_letter: str) -> list[dict]:
    NAME_COL_IDX = _col_letter_to_index("C")
    FLAG_COL_IDX = _col_letter_to_index(flag_col_letter)
    ext = xlsx_path.suffix.lower()
    out = []
    if ext in (".xlsx", ".xlsm"):
        try:
            wb = load_workbook(xlsx_path, data_only=True, read_only=True)
        except Exception:
            return []
        sh_name = _resolve_sheet_exact_or_fuzzy_openpyxl(wb, sheet_title)
        if not sh_name:
            return []
        ws = wb[sh_name]
        for day, (rmin, rmax) in _DAY_RANGES.items():
            seen = set()
            for row in ws.iter_rows(
                min_row=rmin, max_row=rmax,
                min_col=min(NAME_COL_IDX, FLAG_COL_IDX),
                max_col=max(NAME_COL_IDX, FLAG_COL_IDX),
                values_only=True
            ):
                name = str(row[NAME_COL_IDX - min(NAME_COL_IDX, FLAG_COL_IDX)] or "").strip()
                flag = str(row[FLAG_COL_IDX - min(NAME_COL_IDX, FLAG_COL_IDX)] or "").strip().lower()
                if name and flag == "ooo":
                    key = name.casefold()
                    if key not in seen:
                        seen.add(key)
                        out.append({"day": day, "name": name, "activity": "OOO"})
        return out
    elif ext == ".xlsb" and open_xlsb is not None:
        try:
            with open_xlsb(xlsx_path) as wb:
                sh_name = _resolve_sheet_exact_or_fuzzy_xlsb(wb, sheet_title)
                if not sh_name:
                    return []
                ws = wb.get_sheet(sh_name)
                rows_by_index = {}
                for ridx, row in enumerate(ws.rows(), start=1):
                    rows_by_index[ridx] = [c.v for c in row]
                for day, (rmin, rmax) in _DAY_RANGES.items():
                    seen = set()
                    for ridx in range(rmin, rmax + 1):
                        r = rows_by_index.get(ridx) or []
                        name = str((r[NAME_COL_IDX - 1] if len(r) >= NAME_COL_IDX else "") or "").strip()
                        flag = str((r[FLAG_COL_IDX - 1] if len(r) >= FLAG_COL_IDX else "") or "").strip().lower()
                        if name and flag == "ooo":
                            key = name.casefold()
                            if key not in seen:
                                seen.add(key)
                                out.append({"day": day, "name": name, "activity": "OOO"})
            return out
        except Exception:
            return []
    return []
def _resolve_sheet_name(wb, desired_patterns: list[str]) -> str | None:
    if not desired_patterns:
        return None
    desired = {_norm_sheet_name(p) for p in desired_patterns}
    norm_to_actual = {_norm_sheet_name(n): n for n in wb.sheetnames}
    for want in desired:
        if want in norm_to_actual:
            return norm_to_actual[want]
    for actual in wb.sheetnames:
        ns = _norm_sheet_name(actual)
        for want in desired:
            if (want in ns) or ns.startswith(want) or ns.endswith(want):
                return actual
    return None
def _read_names_from_matching_sheets_row_xlsx(path: Path, sheet_patterns: list[str],
                                              row_number: int, max_cols: int = 400) -> list[str]:
    ext = path.suffix.lower()
    names = []
    want = {_norm_sheet_name(p) for p in (sheet_patterns or [])}
    if ext in (".xlsx", ".xlsm"):
        try:
            wb = load_workbook(path, data_only=True, read_only=True)
        except Exception:
            print(f"[non-wip] Could not open workbook for PH: {path}")
            return []
        for sh in wb.sheetnames:
            nsh = _norm_sheet_name(sh)
            if any(w in nsh for w in want):
                ws = wb[sh]
                for r in ws.iter_rows(min_row=row_number, max_row=row_number, min_col=1, max_col=max_cols, values_only=True):
                    for val in r:
                        nm = _clean_name(val)
                        if nm: names.append(nm)
    elif ext == ".xlsb":
        if open_xlsb is None:
            print("[non-wip] '.xlsb' requires 'pyxlsb'. Try: pip install pyxlsb")
            return []
        try:
            with open_xlsb(path) as wb:
                for sh in wb.sheets:
                    nsh = _norm_sheet_name(sh)
                    if any(w in nsh for w in want):
                        ws = wb.get_sheet(sh)
                        for ridx, row in enumerate(ws.rows(), start=1):
                            if ridx < row_number: 
                                continue
                            if ridx > row_number: 
                                break
                            for cidx, cell in enumerate(row, start=1):
                                if cidx > max_cols: 
                                    break
                                nm = _clean_name(cell.v)
                                if nm: names.append(nm)
        except Exception as e:
            print(f"[non-wip] Failed reading PH .xlsb {path.name}: {e}")
            return []
    seen, out = set(), []
    for n in names:
        k = n.casefold()
        if k not in seen:
            seen.add(k); out.append(n)
    return out
def _read_names_from_sheet_col_xlsx(path: Path, sheet_patterns: list[str], col_letter: str = "A",
                                    start_row: int = 1, max_rows: int = 400) -> list[str]:
    ext = path.suffix.lower()
    col_idx = _col_letter_to_index(col_letter)
    start_row = max(1, int(start_row))
    end_row = max(start_row, start_row + max_rows - 1)
    names = []
    if ext in (".xlsx", ".xlsm"):
        try:
            wb = load_workbook(path, data_only=True, read_only=True)
        except Exception:
            print(f"[non-wip] Could not open workbook: {path}")
            return []
        sheet_name = _resolve_sheet_name(wb, sheet_patterns)
        if not sheet_name:
            print(f"[non-wip] No sheet matched {sheet_patterns} in {path.name}")
            return []
        ws = wb[sheet_name]
        for r in ws.iter_rows(min_row=start_row, max_row=end_row, min_col=col_idx, max_col=col_idx, values_only=True):
            nm = _clean_name(r[0])
            if nm: names.append(nm)
    elif ext == ".xlsb":
        if open_xlsb is None:
            print("[non-wip] '.xlsb' requires 'pyxlsb'. Try: pip install pyxlsb")
            return []
        try:
            with open_xlsb(path) as wb:
                sheet_name = None
                norm_to_actual = {_norm_sheet_name(n): n for n in wb.sheets}
                desired = {_norm_sheet_name(p) for p in (sheet_patterns or [])}
                for want in desired:
                    if want in norm_to_actual:
                        sheet_name = norm_to_actual[want]; break
                if sheet_name is None:
                    for actual in wb.sheets:
                        ns = _norm_sheet_name(actual)
                        if any((want in ns) or ns.startswith(want) or ns.endswith(want) for want in desired):
                            sheet_name = actual; break
                if sheet_name is None:
                    print(f"[non-wip] No sheet matched {sheet_patterns} in {path.name}")
                    return []
                sh = wb.get_sheet(sheet_name)
                for ridx, row in enumerate(sh.rows(), start=1):
                    if ridx < start_row: 
                        continue
                    if ridx > end_row: 
                        break
                    val = None
                    for cidx, cell in enumerate(row, start=1):
                        if cidx == col_idx:
                            val = cell.v; break
                    nm = _clean_name(val)
                    if nm: names.append(nm)
        except Exception as e:
            print(f"[non-wip] Failed reading .xlsb {path.name}: {e}")
            return []
    seen, out = set(), []
    for n in names:
        k = n.casefold()
        if k not in seen:
            seen.add(k); out.append(n)
    return out
def _read_names_from_all_sheets_row_xlsx(path: Path, row_number: int, max_cols: int = 400) -> list[str]:
    try:
        wb = load_workbook(path, data_only=True, read_only=True)
    except Exception:
        print(f"[non-wip] Could not open workbook for PH: {path}")
        return []
    all_names = []
    for sh in wb.sheetnames:
        ws = wb[sh]
        for r in ws.iter_rows(min_row=row_number, max_row=row_number, min_col=1, max_col=max_cols, values_only=True):
            for val in r:
                nm = _clean_name(val)
                if nm:
                    all_names.append(nm)
    seen, out = set(), []
    for n in all_names:
        k = n.casefold()
        if k not in seen:
            seen.add(k); out.append(n)
    return out
def _source_file_only(s: str) -> str:
    return s.split(" :: ", 1)[0].strip()
def _parse_person_hours_cell(s: str | None) -> dict[str, float]:
    if not s:
        return {}
    try:
        obj = json.loads(s)
        out = {}
        for name, vals in (obj or {}).items():
            try:
                out[str(name).strip()] = float((vals or {}).get("actual") or 0.0)
            except Exception:
                out[str(name).strip()] = 0.0
        return out
    except Exception:
        return {}
def _lookup_actual_hours(ph_by_name: dict[str, float], person: str) -> float:
    if person in ph_by_name:
        return float(ph_by_name[person])
    want = person.casefold()
    for k, v in ph_by_name.items():
        if k.casefold() == want:
            return float(v)
    return 0.0
def _get_team_cfg(team_name: str):
    return TEAM_CFG.get(str(team_name).casefold())
def _read_people_from_file_for_team(xlsx_path: Path, team_name: str) -> list[str]:
    cfg = _get_team_cfg(team_name)
    if not cfg:
        return []
    if "row" in cfg and "sheet_patterns" in cfg:
        return _read_names_from_matching_sheets_row_xlsx(
            xlsx_path,
            sheet_patterns=cfg["sheet_patterns"],
            row_number=cfg["row"],
            max_cols=400,
        )
    patterns = cfg.get("sheet_patterns", [])
    col      = cfg.get("col", "A")
    start    = cfg.get("start", 1)
    return _read_names_from_sheet_col_xlsx(xlsx_path, sheet_patterns=patterns, col_letter=col, start_row=start)
def main():
    if not REPO_CSV.exists():
        raise FileNotFoundError(f"metrics CSV not found: {REPO_CSV}")
    df = pd.read_csv(REPO_CSV, dtype=str, keep_default_na=False)
    df["team_norm"] = df.get("team", "").astype(str).str.casefold()
    df = df[(df.get("source_file", "") != "")]
    if df.empty:
        print("[non-wip] No rows found in metrics_aggregate_dev.csv with a source file")
        return
    df["period_date"] = df["period_date"].apply(_coerce_to_date)
    df = df.dropna(subset=["period_date"])
    df["period_date"] = pd.to_datetime(df["period_date"]).dt.date
    df["source_file_only"] = df["source_file"].apply(_source_file_only)
    ph_index: dict[tuple, dict[str, float]] = {}
    for _, r in df.iterrows():
        key = (r["team_norm"], r["period_date"], r["source_file_only"])
        ph = _parse_person_hours_cell(r.get("Person Hours"))
        if key not in ph_index:
            ph_index[key] = ph
        else:
            ph_index[key].update(ph)
    unique_refs = df[["team", "team_norm", "period_date", "source_file_only"]].drop_duplicates()
    out_rows = []
    for _, row in unique_refs.iterrows():
        team        = row["team"]
        team_norm   = row["team_norm"]
        period_date = row["period_date"]
        src         = row["source_file_only"]
        ph_by_name = ph_index.get((team_norm, period_date, src), {})
        cfg = _get_team_cfg(team_norm)
        use_person_hours = bool(cfg and cfg.get("people_from") == "person_hours")
        people: list[str] = []
        if use_person_hours:
            people = [n for n in ph_by_name.keys() if _clean_name(n)]
            if not people:
                p = Path(src)
                if not p.exists():
                    print(f"[non-wip] No Person Hours names and file missing for team '{team}': {src}")
                    continue
                ext = p.suffix.lower()
                if ext not in (".xlsx", ".xlsm", ".xlsb"):
                    print(f"[non-wip] No Person Hours names and unsupported file type ({ext}) for team '{team}': {src}")
                    continue
                if ext == ".xlsb" and open_xlsb is None:
                    print("[non-wip] '.xlsb' requires 'pyxlsb'. Try: pip install pyxlsb")
                    continue
                people = _read_people_from_file_for_team(p, team_norm)
        else:
            p = Path(src)
            if not p.exists():
                print(f"[non-wip] Skip missing file: {src}")
                continue
            ext = p.suffix.lower()
            if ext not in (".xlsx", ".xlsm", ".xlsb"):
                print(f"[non-wip] Skip unsupported file type ({ext}): {src}")
                continue
            if ext == ".xlsb" and open_xlsb is None:
                print("[non-wip] '.xlsb' requires 'pyxlsb'. Try: pip install pyxlsb")
                continue
            people = _read_people_from_file_for_team(p, team_norm)
        if not people:
            source_hint = "Person Hours" if use_person_hours else "workbook"
            print(f"[non-wip] No names found for team '{team}' on {period_date} from {source_hint}")
            continue
        per_person_non_wip = {}
        total_non_wip = 0.0
        total_wip_capped = 0.0
        for person in people:
            wip_actual = _lookup_actual_hours(ph_by_name, person)
            non_wip = max(0.0, WEEKLY_HOURS_DEFAULT - float(wip_actual))
            per_person_non_wip[person] = round(non_wip, 2)
            total_non_wip += non_wip
            total_wip_capped += min(float(wip_actual), WEEKLY_HOURS_DEFAULT)
        people_count = len(people)
        weekly_total_available = WEEKLY_HOURS_DEFAULT * people_count if people_count > 0 else 0.0
        pct_in_wip = (total_wip_capped / weekly_total_available * 100.0) if weekly_total_available > 0 else 0.0
        pct_in_wip = round(pct_in_wip, 2)
        row_obj = {
            "team": team,
            "period_date": period_date.isoformat(),
            "source_file": src,
            "people_count": people_count,
            "total_non_wip_hours": round(total_non_wip, 2),
            "% in WIP": pct_in_wip,
            "non_wip_by_person": json.dumps(per_person_non_wip, ensure_ascii=False),
        }
        ooo_cfg = TEAM_OOO_CFG.get(team_norm)
        if ooo_cfg:
            p = Path(src)
            if p.exists():
                try:
                    details = extract_ooo_per_day(p, ooo_cfg["sheet"], ooo_cfg["flag_col"])
                except Exception:
                    details = []
            else:
                details = []
            row_obj["non_wip_activities"] = json.dumps(details, ensure_ascii=False)
        else:
            row_obj["non_wip_activities"] = "[]"
        out_rows.append(row_obj)
        out_rows.append({
            "team": team,
            "period_date": period_date.isoformat(),
            "source_file": src,
            "people_count": people_count,
            "total_non_wip_hours": round(total_non_wip, 2),
            "% in WIP": pct_in_wip,
            "non_wip_by_person": json.dumps(per_person_non_wip, ensure_ascii=False),
        })
    if not out_rows:
        print("[non-wip] No weekly rows produced for mapped teams.")
        return
    OUT_CSV.parent.mkdir(parents=True, exist_ok=True)
    cols = [
        "team",
        "period_date",
        "source_file",
        "people_count",
        "total_non_wip_hours",
        "% in WIP",
        "non_wip_by_person",
        "non_wip_activities",
    ]
    with OUT_CSV.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(out_rows)
    print(f"[non-wip] Wrote {len(out_rows)} rows to {OUT_CSV}")
if __name__ == "__main__":
    main()