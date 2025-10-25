#!/usr/bin/env python3
from __future__ import annotations
import argparse, csv, os
from datetime import datetime, date
from typing import Dict, List, Tuple, Any, Optional, Iterable
def _clean(s: Any) -> str:
    return "" if s is None else str(s).strip()
def _to_date_iso(v: str) -> Optional[str]:
    s = _clean(v)
    if not s:
        return None
    try:
        return datetime.fromisoformat(s).date().isoformat()
    except Exception:
        pass
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).date().isoformat()
        except Exception:
            pass
    return None
def _read_csv(path: str) -> Tuple[List[Dict[str, str]], List[str]]:
    if not os.path.exists(path):
        return [], []
    with open(path, "r", encoding="utf-8", newline="") as f:
        r = csv.DictReader(f)
        rows = [dict(row) for row in r]
        headers = list(r.fieldnames or [])
    return rows, headers
def _write_csv(path: str, rows: List[Dict[str, Any]], headers: List[str]) -> None:
    with open(path, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers)
        w.writeheader()
        for row in rows:
            w.writerow({h: row.get(h, "") for h in headers})
def _merge_on_key(existing: List[Dict[str, Any]],
                  new_rows: List[Dict[str, Any]],
                  key_fn) -> List[Dict[str, Any]]:
    by_key: Dict[Any, Dict[str, Any]] = {key_fn(r): r for r in existing}
    for r in new_rows:
        by_key[key_fn(r)] = r  
    existing_keys = [key_fn(r) for r in existing]
    new_keys = [k for k in by_key.keys() if k not in existing_keys]
    merged = [by_key[k] for k in existing_keys if k in by_key] + [by_key[k] for k in new_keys]
    return merged
def _parse_dates(args) -> List[str]:
    dates: List[str] = []
    for d in (args.date or []):
        iso = _to_date_iso(d)
        if iso: dates.append(iso)
    if args.dates:
        for d in args.dates.split(","):
            iso = _to_date_iso(d)
            if iso: dates.append(iso)
    if args.date_from or args.date_to:
        if not args.date_from or not args.date_to:
            raise SystemExit("If using --date-from/--date-to, provide both.")
        start = datetime.fromisoformat(_to_date_iso(args.date_from)).date()
        end   = datetime.fromisoformat(_to_date_iso(args.date_to)).date()
        if end < start:
            raise SystemExit("--date-to must be >= --date-from")
        cur = start
        while cur <= end:
            dates.append(cur.isoformat())
            cur = cur.fromordinal(cur.toordinal() + 7)  
    seen = set(); out = []
    for d in dates:
        if d not in seen:
            seen.add(d); out.append(d)
    if not out:
        raise SystemExit("No valid dates provided. Use --date YYYY-MM-DD (repeatable), "
                         "--dates 2025-10-27,2025-11-03 or --date-from/--date-to.")
    return out
METRICS_SRC_HEADERS = [
    "Team","Week","Total Available Hours","Completed Hours","Target Output","Actual Output",
    "Target UPLH","Actual UPLH","HC in WIP","Actual HC Used","Person Hours","Outputs by Person",
    "Outputs by Cell/Station","Cell/Station Hours","Hours by Cell/Station - by person",
    "Output by Cell/Station - by person","UPLH by Cell/Station - by person","Closures",
    "Open Complaint Timeliness"
]
METRICS_AGG_HEADERS = [
    "team","period_date","source_file","Total Available Hours","Completed Hours","Target Output",
    "Actual Output","Target UPLH","Actual UPLH","UPLH WP1","UPLH WP2","HC in WIP",
    "Actual HC Used","People in WIP","Person Hours","Outputs by Person","Outputs by Cell/Station",
    "Cell/Station Hours","Hours by Cell/Station - by person","Output by Cell/Station - by person",
    "UPLH by Cell/Station - by person","Open Complaint Timeliness","error","Closures"
]
def project_metrics_row(src: Dict[str, str], source_file: str) -> Dict[str, Any]:
    return {
        "team": _clean(src.get("Team")),
        "period_date": _to_date_iso(src.get("Week")) or _clean(src.get("Week")),
        "source_file": source_file,
        "Total Available Hours": _clean(src.get("Total Available Hours")),
        "Completed Hours": _clean(src.get("Completed Hours")),
        "Target Output": _clean(src.get("Target Output")),
        "Actual Output": _clean(src.get("Actual Output")),
        "Target UPLH": _clean(src.get("Target UPLH")),
        "Actual UPLH": _clean(src.get("Actual UPLH")),
        "UPLH WP1": "",  # not present in metrics.csv
        "UPLH WP2": "",  # not present in metrics.csv
        "HC in WIP": _clean(src.get("HC in WIP")),
        "Actual HC Used": _clean(src.get("Actual HC Used")),
        "People in WIP": "",  # not present in metrics.csv
        "Person Hours": _clean(src.get("Person Hours")),
        "Outputs by Person": _clean(src.get("Outputs by Person")),
        "Outputs by Cell/Station": _clean(src.get("Outputs by Cell/Station")),
        "Cell/Station Hours": _clean(src.get("Cell/Station Hours")),
        "Hours by Cell/Station - by person": _clean(src.get("Hours by Cell/Station - by person")),
        "Output by Cell/Station - by person": _clean(src.get("Output by Cell/Station - by person")),
        "UPLH by Cell/Station - by person": _clean(src.get("UPLH by Cell/Station - by person")),
        "Open Complaint Timeliness": _clean(src.get("Open Complaint Timeliness")),
        "error": "",
        "Closures": _clean(src.get("Closures")),
    }
NONWIP_SRC_HEADERS = [
    "Team","Week","People Count","Total Non-WIP Hours","OOO Hours","% in WIP",
    "Non-WIP by Person","Non-WIP Activities"
]
NONWIP_OUT_HEADERS = [
    "team","period_date","source_file","people_count","total_non_wip_hours",
    "% in WIP","non_wip_by_person","non_wip_activities","OOO Hours"
]
def project_nonwip_row(src: Dict[str, str], source_file: str) -> Dict[str, Any]:
    return {
        "team": _clean(src.get("Team")),
        "period_date": _to_date_iso(src.get("Week")) or _clean(src.get("Week")),
        "source_file": source_file,
        "people_count": _clean(src.get("People Count")),
        "total_non_wip_hours": _clean(src.get("Total Non-WIP Hours")),
        "% in WIP": _clean(src.get("% in WIP")),
        "non_wip_by_person": _clean(src.get("Non-WIP by Person")),
        "non_wip_activities": _clean(src.get("Non-WIP Activities")),
        "OOO Hours": _clean(src.get("OOO Hours")),
    }
def push_metrics(dates_iso: List[str], src_path: str, out_path: str, source_file_value: str):
    src_rows, src_headers = _read_csv(src_path)
    if not src_rows:
        raise SystemExit(f"No rows in {src_path}")
    want = []
    for r in src_rows:
        wk = _to_date_iso(r.get("Week")) or _clean(r.get("Week"))
        if wk in dates_iso:
            want.append(project_metrics_row(r, source_file_value))
    if not want:
        print(f"[metrics] No matching rows for dates: {', '.join(dates_iso)}")
        return
    existing, _ = _read_csv(out_path)
    def key_fn(x): return (_clean(x.get("team")), _clean(x.get("period_date")), _clean(x.get("source_file")))
    merged = _merge_on_key(existing, want, key_fn)
    _write_csv(out_path, merged, METRICS_AGG_HEADERS)
    print(f"[metrics] Wrote {len(merged)} total rows -> {out_path} (added/updated {len(want)})")
def push_nonwip(dates_iso: List[str], src_path: str, out_path: str, source_file_value: str):
    src_rows, _ = _read_csv(src_path)
    if not src_rows:
        raise SystemExit(f"No rows in {src_path}")
    want = []
    for r in src_rows:
        wk = _to_date_iso(r.get("Week")) or _clean(r.get("Week"))
        if wk in dates_iso:
            want.append(project_nonwip_row(r, source_file_value))
    if not want:
        print(f"[non_wip] No matching rows for dates: {', '.join(dates_iso)}")
        return
    existing, _ = _read_csv(out_path)
    def key_fn(x): return (_clean(x.get("team")), _clean(x.get("period_date")), _clean(x.get("source_file")))
    merged = _merge_on_key(existing, want, key_fn)
    _write_csv(out_path, merged, NONWIP_OUT_HEADERS)
    print(f"[non_wip] Wrote {len(merged)} total rows -> {out_path} (added/updated {len(want)})")
def main():
    ap = argparse.ArgumentParser(description="Push selected dates from metrics.csv & non_wip.csv to their aggregate CSVs.")
    ap.add_argument("--dates", help="Comma-separated list of dates (e.g. 2025-10-27,2025-11-03)")
    ap.add_argument("--date", action="append", help="Specific date (repeatable). Example: --date 2025-10-27")
    ap.add_argument("--date-from", help="Inclusive start date (YYYY-MM-DD)")
    ap.add_argument("--date-to", help="Inclusive end date (YYYY-MM-DD)")
    ap.add_argument("--metrics", default="metrics.csv", help="Path to metrics.csv")
    ap.add_argument("--nonwip", default="non_wip.csv", help="Path to non_wip.csv")
    ap.add_argument("--out-metrics", default="metrics_aggregate_dev.csv", help="Output for metrics aggregate")
    ap.add_argument("--out-nonwip", default="non_wip_activities.csv", help="Output for non-wip aggregate")
    ap.add_argument("--source-file-metrics", default="metrics.csv", help="Value to place in 'source_file' for metrics rows")
    ap.add_argument("--source-file-nonwip", default="non_wip.csv", help="Value to place in 'source_file' for non-wip rows")
    ap.add_argument("--skip-metrics", action="store_true", help="Skip pushing metrics")
    ap.add_argument("--skip-nonwip", action="store_true", help="Skip pushing non_wip")
    args = ap.parse_args()
    dates_iso = _parse_dates(args)
    if not args.skip_metrics:
        if not os.path.exists(args.out_metrics):
            _write_csv(args.out_metrics, [], METRICS_AGG_HEADERS)
        push_metrics(dates_iso, args.metrics, args.out_metrics, args.source_file_metrics)
    if not args.skip_nonwip:
        if not os.path.exists(args.out_nonwip):
            _write_csv(args.out_nonwip, [], NONWIP_OUT_HEADERS)
        push_nonwip(dates_iso, args.nonwip, args.out_nonwip, args.source_file_nonwip)
if __name__ == "__main__":
    main()
