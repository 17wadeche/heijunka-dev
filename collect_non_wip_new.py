#!/usr/bin/env python3
from __future__ import annotations
import argparse, csv, json, math, os, sys
from collections import defaultdict
from datetime import date, datetime
from typing import Any, Dict, Iterable, List, Optional, Tuple
def _to_date(v) -> Optional[date]:
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    if not s:
        return None
    try:
        n = float(s)
        return (datetime(1899, 12, 30) + timedelta(days=n)).date()
    except Exception:
        pass
    for fmt in ("%Y-%m-%d", "%m/%d/%Y"):
        try:
            if fmt == "%Y-%m-%d":
                return datetime.fromisoformat(s).date()
            return datetime.strptime(s, fmt).date()
        except Exception:
            pass
    return None
def _to_float(x) -> Optional[float]:
    if x is None:
        return None
    if isinstance(x, float):
        return x
    try:
        return float(str(x).replace(",", "").strip())
    except Exception:
        return None
def _clean(s: Any) -> str:
    if s is None:
        return ""
    if isinstance(s, float) and math.isnan(s):
        return ""
    return str(s).strip()
def _sheetnames_xlsb(path: str) -> List[str]:
    import pandas as pd
    with pd.ExcelFile(path, engine="pyxlsb") as xf:
        return list(xf.sheet_names)
def _rows_from_xlsb(path: str, sheet_name: str) -> Iterable[Tuple[Any, ...]]:
    import pandas as pd
    df = pd.read_excel(path, sheet_name=sheet_name, engine="pyxlsb", header=None)
    for row in df.itertuples(index=False, name=None):
        yield tuple(row)
def _sheetnames_xlsx_like(path: str) -> List[str]:
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True, read_only=True)
    return list(wb.sheetnames)
def _rows_from_xlsx_like(path: str, sheet_name: str) -> Iterable[Tuple[Any, ...]]:
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet_name]
    for r in ws.iter_rows(values_only=True):
        yield tuple(r)
def _get_io(path: str):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xlsb":
        return _sheetnames_xlsb, _rows_from_xlsb
    elif ext in (".xlsx", ".xlsm"):
        return _sheetnames_xlsx_like, _rows_from_xlsx_like
    else:
        raise ValueError(f"Unsupported workbook extension '{ext}'. Use .xlsb/.xlsx/.xlsm")
def _find_sheet_by_hint(sheet_names: List[str], hint: str) -> str:
    if not hint:
        raise ValueError("Empty sheet hint")
    exact = [nm for nm in sheet_names if nm.strip().lower() == hint.strip().lower()]
    if exact:
        return exact[0]
    contains = [nm for nm in sheet_names if hint.lower() in nm.lower()]
    if contains:
        return contains[0]
    raise ValueError(f"Sheet '{hint}' not found. Available: {sheet_names}")
def parse_available_people_and_nonwip(rows: Iterable[Tuple[Any, ...]]):
    nonwip_cols_idx = list(range(4, 9))
    people_by_week: Dict[date, set] = defaultdict(set)
    nonwip_by_week: Dict[date, Dict[str, float]] = defaultdict(lambda: defaultdict(float))
    current_week: Optional[date] = None
    current_person: Optional[str] = None
    for r in rows:
        r = r or tuple()
        week_raw = r[0] if len(r) >= 1 else None
        week = _to_date(week_raw) or current_week
        person_or_flag = _clean(r[2] if len(r) >= 3 else "")  # Column C
        if _to_date(week_raw):
            current_week = week
        if person_or_flag and person_or_flag.lower() not in {"available wip", "non-wip"}:
            current_person = person_or_flag
        if not (current_week and current_person):
            continue
        if person_or_flag.lower() in {"available wip", "non-wip"}:
            people_by_week[current_week].add(current_person)
        if person_or_flag.lower() == "non-wip":
            s = 0.0
            for c in nonwip_cols_idx:
                v = r[c] if len(r) > c else None
                fv = _to_float(v)
                if fv is not None:
                    s += fv
            if s:
                nonwip_by_week[current_week][current_person] += s
    return people_by_week, nonwip_by_week
def parse_prod_analysis(rows: Iterable[Tuple[Any, ...]]) -> Dict[date, Dict[str, Any]]:
    COL_DATE, COL_NAME, COL_FLAG, COL_MINUTES, COL_ACTIVITY = 0, 3, 4, 7, 11
    buckets: Dict[date, Dict[str, Any]] = defaultdict(lambda: {
        "ooo_hours": 0.0,
        "non_wip_activities": [],  # list of dicts {name, activity, hours}
    })
    for r in rows:
        r = r or tuple()
        wk = _to_date(r[COL_DATE] if len(r) > COL_DATE else None)
        if not wk:
            continue
        name = _clean(r[COL_NAME] if len(r) > COL_NAME else "")
        flag = _clean(r[COL_FLAG] if len(r) > COL_FLAG else "")
        mins = _to_float(r[COL_MINUTES] if len(r) > COL_MINUTES else None) or 0.0
        act  = _clean(r[COL_ACTIVITY] if len(r) > COL_ACTIVITY else "")
        if not (flag or mins or name or act):
            continue
        b = buckets[wk]
        if flag.lower() == "ooo" and mins > 0:
            b["ooo_hours"] += mins / 60.0
        if flag.lower() == "non wip" and mins > 0:
            b["non_wip_activities"].append({
                "name": name,
                "activity": act,
                "hours": round(mins / 60.0, 2),
            })
    for wk, b in buckets.items():
        b["ooo_hours"] = round(b["ooo_hours"], 2)
    return buckets
def load_completed_hours(metrics_csv: str) -> Dict[Tuple[str, str], float]:
    out: Dict[Tuple[str, str], float] = {}
    with open(metrics_csv, "r", encoding="utf-8") as f:
        r = csv.DictReader(f)
        for row in r:
            team = row.get("Team") or row.get("team") or ""
            wk = row.get("Week") or row.get("period_date") or ""
            ch = row.get("Completed Hours") or row.get("completed_hours") or "0"
            team = _clean(team)
            wk = _clean(wk)
            val = _to_float(ch) or 0.0
            if team and wk:
                out[(team, wk)] = out.get((team, wk), 0.0) + float(val)
    return out
def build_non_wip_rows(config_path: str,
                       chosen_teams: Optional[List[str]],
                       all_teams: bool,
                       metrics_csv: str) -> List[Dict[str, Any]]:
    with open(config_path, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    teams_to_run = list(cfg.keys()) if all_teams else (chosen_teams or [])
    if not teams_to_run:
        raise SystemExit("No teams specified. Use --all or --team <NAME> (repeatable).")
    completed_index = load_completed_hours(metrics_csv)
    out_rows: List[Dict[str, Any]] = []
    for team in teams_to_run:
        entry = cfg.get(team) or {}
        path = entry.get("workbook")
        if not path or not os.path.exists(path):
            raise SystemExit(f"[{team}] Workbook not found: {path}")
        prod_cfg = entry.get("prod_sheets") or entry.get("prod_sheet") or []
        prod_hints = prod_cfg if isinstance(prod_cfg, list) else [prod_cfg]
        avail_hint = entry.get("avail_sheet")
        if not avail_hint:
            raise SystemExit(f"[{team}] Missing 'avail_sheet' in config.")
        get_names, get_rows = _get_io(path)
        sheet_names = get_names(path)
        prod_rows_all: List[Tuple[Any, ...]] = []
        for hint in prod_hints:
            if not hint:
                continue
            try:
                nm = _find_sheet_by_hint(sheet_names, hint)
                prod_rows_all.extend(list(get_rows(path, nm)))
            except Exception as e:
                raise SystemExit(f"[{team}] Prod sheet '{hint}' error: {e}")
        try:
            avail_name = _find_sheet_by_hint(sheet_names, avail_hint)
            avail_rows = list(get_rows(path, avail_name))
        except Exception as e:
            raise SystemExit(f"[{team}] Available sheet '{avail_hint}' error: {e}")
        people_by_week, nonwip_by_week = parse_available_people_and_nonwip(avail_rows)
        prod_buckets = parse_prod_analysis(prod_rows_all)
        all_weeks = sorted({*people_by_week.keys(), *nonwip_by_week.keys(), *prod_buckets.keys()})
        for wk in all_weeks:
            iso = wk.isoformat()
            people_count = len(people_by_week.get(wk, set()))
            nonwip_by_person = {k: round(float(v), 2) for k, v in nonwip_by_week.get(wk, {}).items()}
            total_non_wip_hours = round(sum(nonwip_by_person.values()), 2)
            ooo_hours = float(prod_buckets.get(wk, {}).get("ooo_hours", 0.0) or 0.0)
            activities = prod_buckets.get(wk, {}).get("non_wip_activities", [])
            completed = float(completed_index.get((team, iso), 0.0))
            denom = (people_count * 40.0) - ooo_hours
            pct_in_wip = round((completed / denom * 100.0), 2) if denom > 0 else None
            out_rows.append({
                "Team": team,
                "Week": iso,
                "People Count": people_count,
                "Total Non-WIP Hours": total_non_wip_hours,
                "OOO Hours": round(ooo_hours, 2),
                "% in WIP": pct_in_wip,
                "Non-WIP by Person": json.dumps(nonwip_by_person, ensure_ascii=False),
                "Non-WIP Activities": json.dumps(activities, ensure_ascii=False),
            })
        print(f"[{team}] OK: {len(all_weeks)} week(s)")
    return out_rows
def main():
    ap = argparse.ArgumentParser(description="Collect Non-WIP/OOO metrics into a new CSV.")
    ap.add_argument("--config", required=True, help="Path to teams.json")
    ap.add_argument("--metrics", required=True, help="Path to metrics.csv produced by heijunka_new_layout.py")
    ap.add_argument("--team", action="append", help="Team name from teams.json (repeatable)")
    ap.add_argument("--all", action="store_true", help="Process all teams in config")
    ap.add_argument("--out", default="non_wip.csv", help="Output CSV path (default: non_wip.csv)")
    args = ap.parse_args()
    try:
        rows = build_non_wip_rows(
            config_path=args.config,
            chosen_teams=args.team,
            all_teams=args.all,
            metrics_csv=args.metrics,
        )
    except SystemExit as e:
        print(str(e), file=sys.stderr); sys.exit(2)
    except Exception as e:
        print(f"Failed: {e}", file=sys.stderr); sys.exit(1)
    if not rows:
        print("No data produced.", file=sys.stderr); sys.exit(1)
    cols = [
        "Team", "Week",
        "People Count",
        "Total Non-WIP Hours",
        "OOO Hours",
        "% in WIP",
        "Non-WIP by Person",
        "Non-WIP Activities",
    ]
    with open(args.out, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(rows)
    print(f"Wrote {len(rows)} rows -> {args.out}")
if __name__ == "__main__":
    from datetime import timedelta  # needed for Excel serial conversion
    main()