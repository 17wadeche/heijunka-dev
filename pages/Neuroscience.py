# pages/Neuroscience.py
import os, sys
from pathlib import Path
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
import pandas as pd
import numpy as np
import streamlit as st
import altair as alt
import json
import re
import unicodedata
from utils.nonwip_kpi_lookup import enterprise_nonwip_kpi_lookup
from utils.activity_map import ACTIVITY_MAP
from utils.styles import apply_global_styles
apply_global_styles()
NON_WIP_DEFAULT_PATH = Path(r"C:\heijunka-dev\NS_DATA\ns_non_wip_activities.csv")
def _safe_secret(name: str, default=None):
    import os
    try:
        return st.secrets.get(name, os.environ.get(name, default))
    except Exception:
        return os.environ.get(name, default)
NON_WIP_DATA_URL = _safe_secret("NS_NON_WIP_DATA_URL")
DATA_URL = _safe_secret("NS_HEIJUNKA_DATA_URL")
def _fmt_hours_minutes(x) -> str:
    try:
        total_mins = int(round(float(x) * 60))
    except Exception:
        return "0m"
    h, m = divmod(total_mins, 60)
    if h and m:
        return f"{h}h {m:02d}m"
    if h and not m:
        return f"{h}h"
    return f"{m}m"
from pathlib import Path
TEAMS_CONFIG_PATH = Path(__file__).resolve().parents[1] / "teams.json"
@st.cache_data(show_spinner=False, ttl=15 * 60)
def load_team_config(config_path: str | None = None) -> dict:
    p = Path(config_path) if config_path else TEAMS_CONFIG_PATH
    try:
        with open(p, "r", encoding="utf-8") as f:
            obj = json.load(f)
        return obj if isinstance(obj, dict) else {}
    except Exception:
        return {}
def irl_people_for_team(team: str, config: dict) -> set[str]:
    if not isinstance(config, dict):
        return set()
    team_cfg = config.get(str(team).strip(), {})
    if not isinstance(team_cfg, dict):
        return set()
    raw = team_cfg.get("irl_people", [])
    if not isinstance(raw, list):
        return set()
    return {str(x).strip() for x in raw if str(x).strip()}
def canonical_activity_key(label: str) -> str:
    raw = str(label or "").strip()
    if not raw:
        return ""
    mapped = ACTIVITY_MAP.get(raw.lower().strip(), raw)
    return re.sub(r"[^A-Z0-9]", "", str(mapped).upper())
@st.cache_data(show_spinner=False, ttl=15 * 60)
def load_non_wip(
    nw_path: str | None = None,
    nw_url: str | None = None,
    cache_tag: str = "NS", 
) -> pd.DataFrame:
    if nw_url is None:
        nw_url = NON_WIP_DATA_URL
    if nw_url:
        try:
            df = pd.read_csv(nw_url, dtype=str, keep_default_na=False, encoding="utf-8-sig")
        except Exception:
            import io, requests
            r = requests.get(nw_url, timeout=20)
            r.raise_for_status()
            df = pd.read_csv(
                io.StringIO(r.content.decode("utf-8-sig", errors="replace")),
                dtype=str, keep_default_na=False
            )
    else:
        p = Path(nw_path or NON_WIP_DEFAULT_PATH)
        if not p.exists():
            return pd.DataFrame(columns=[
                "team","period_date","source_file","people_count",
                "total_non_wip_hours","% in WIP","non_wip_by_person"
            ])
        df = pd.read_csv(p, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    if "period_date" in df.columns:
        df["period_date"] = pd.to_datetime(df["period_date"], errors="coerce").dt.normalize()
    for c in ["people_count", "total_non_wip_hours", "% in WIP", "OOO Hours"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    if "% in WIP" in df.columns and "% Non-WIP" not in df.columns:
        s = pd.to_numeric(df["% in WIP"], errors="coerce")
        if pd.notna(s.max()):
            if float(s.max()) <= 1.5:
                pct_wip_0_100 = s * 100.0
            else:
                pct_wip_0_100 = s
            df["% Non-WIP"] = 100.0 - pct_wip_0_100
    return df
@st.cache_data(show_spinner=False, ttl=15 * 60)
def explode_non_wip_by_person(nw: pd.DataFrame) -> pd.DataFrame:
    cols = ["team","period_date","person","Non-WIP Hours"]
    if nw.empty or "non_wip_by_person" not in nw.columns:
        return pd.DataFrame(columns=cols)
    rows = []
    sub = nw[["team","period_date","non_wip_by_person"]].dropna(subset=["non_wip_by_person"])
    for _, r in sub.iterrows():
        payload = r["non_wip_by_person"]
        try:
            obj = json.loads(payload) if isinstance(payload, str) else payload
            if not isinstance(obj, dict):
                continue
        except Exception:
            continue
        for person, hrs in obj.items():
            try:
                v = float(hrs)
            except Exception:
                v = np.nan
            rows.append({
                "team": r["team"],
                "period_date": pd.to_datetime(r["period_date"], errors="coerce").normalize(),
                "person": normalize_person_name(str(person).strip()),
                "Non-WIP Hours": v
            })
    out = pd.DataFrame(rows, columns=cols)
    if not out.empty:
        out["period_date"] = pd.to_datetime(out["period_date"], errors="coerce").dt.normalize()
    out = _filter_excluded_people_frame(out)
    return out
DEFAULT_DATA_PATH = Path(r"C:\heijunka-dev\NS_WIP.csv")
if hasattr(st, "autorefresh"):
    st.autorefresh(interval=60 * 60 * 1000, key="auto-refresh")
@st.cache_data(show_spinner=False, ttl=15 * 60)
def load_data(data_path: str | None, data_url: str | None):
    if data_url:
        try:
            lower = data_url.lower()
            if lower.endswith((".xlsx", ".xlsm", ".xls")):
                df = pd.read_excel(data_url, sheet_name="All Metrics")
            elif lower.endswith(".json"):
                df = pd.read_json(data_url)
            else:
                df = pd.read_csv(
                    data_url,
                    engine="python",      # enables sep=None sniffing
                    sep=None,             # auto-detect delimiter (comma, tab, semicolon…)
                    encoding="utf-8-sig", # handles BOM
                    on_bad_lines="skip",  # don't die on ragged rows
                    dtype=str,            # keep raw text; you coerce later in _postprocess
                )
        except pd.errors.ParserError:
            try:
                df = pd.read_csv(
                    data_url,
                    engine="python",
                    sep=";",
                    encoding="utf-8-sig",
                    on_bad_lines="skip",
                    dtype=str,
                )
            except Exception as e:
                st.error(f"Couldn't parse HEIJUNKA_DATA_URL as CSV: {e}")
                return pd.DataFrame()
        except Exception:
            import io, requests
            try:
                r = requests.get(data_url, timeout=20)
                r.raise_for_status()
                b = r.content
                head = b[:2048].lstrip()
                if head.startswith((b"{", b"[")):
                    df = pd.read_json(io.BytesIO(b))
                elif b[:2] == b"PK":
                    df = pd.read_excel(io.BytesIO(b), sheet_name="All Metrics")
                else:
                    df = pd.read_csv(
                        io.StringIO(b.decode("utf-8-sig", errors="replace")),
                        engine="python",
                        sep=None,
                        on_bad_lines="skip",
                        dtype=str,
                    )
            except Exception as e:
                st.error(f"Failed to fetch/parse HEIJUNKA_DATA_URL: {e}")
                return pd.DataFrame()
        return _postprocess(df)
    if not data_path:
        return pd.DataFrame()
    p = Path(data_path)
    if not p.exists():
        return pd.DataFrame()
    if p.suffix.lower() in (".xlsx", ".xlsm"):
        df = pd.read_excel(p, sheet_name="All Metrics")
    elif p.suffix.lower() == ".csv":
        df = pd.read_csv(p, engine="python", sep=None, encoding="utf-8-sig", on_bad_lines="skip", dtype=str)
    elif p.suffix.lower() == ".json":
        df = pd.read_json(p)
    else:
        return pd.DataFrame()
    return _postprocess(df)
def normalize_person_name(name: str) -> str:
    def _norm(x: str) -> str:
        x = str(x or "")
        x = unicodedata.normalize("NFKC", x)
        x = x.replace("\u00A0", " ")
        x = " ".join(x.split()).strip()
        x = x.lower()
        x = re.sub(r"[^\w\s]", "", x) 
        x = " ".join(x.split())
        return x
    raw = str(name or "")
    clean = " ".join(unicodedata.normalize("NFKC", raw).replace("\u00A0", " ").split()).strip()
    key = _norm(clean)
    aliases = {
        _norm("mirlay morin"): "Mirlay",
        _norm("nikita schazenbach"): "Nikita",
        _norm( "bandaru, phaneedra"): "Phaneedra Bandaru",
        _norm( "kasi m"): "Kasi Rajan",
        _norm("phaneedra"): "Phaneedra Bandaru",
        _norm("jacob"): "Jacob Woolley",
        _norm("pang lee"): "Pang",
        _norm("jacob g"): "Jacob Geraghty",
        _norm("jake"): "Jacob Geraghty",
        _norm("madison moeller"): "Madison",
        _norm("pavani uppari"): "Uppari Pavani",
        _norm("s, prabhu"): "Prabhu S",
        _norm("chandra"): "Nitheesh",
        _norm("damahe, jagruti"): "Jagruti Damahe",
        _norm("kallagunta, malleshwari"): "Malleshwari Kallagunta",
        _norm("gopikalyani ijigiri"): "Gopikalyani Iligiri",
        _norm("dey, pranjal"): "Pranjal Dey",
        _norm("shanmugasundaram, naveen"): "Naveen Shanmugasundaram",
        _norm("shanmugasundaram, naveenkumar"): "Naveen Shanmugasundaram",
        _norm("naveenkumar shanmugasundaram"): "Naveen Shanmugasundaram",
        _norm("badugu, aravind kumar"): "Aravind Kumar Badugu",
        _norm("aravind badugu"): "Aravind Kumar Badugu",
        _norm("aravind kumar badugu"): "Aravind Kumar Badugu",
        _norm("s, giridhar"): "Giridhar S",
        _norm("vemulapalli, reshmita"): "Reshmita",
        _norm("rick kennedy"): "Rick",
        _norm("surekha raju anantarapu"): "Surekha Raju",
        _norm("anwar, mohd faiz"): "Mohd Faiz Anwar",
        _norm("mohd anwar"): "Mohd Faiz Anwar",
        _norm("dominick olaes"): "Dominick",
        _norm("tabitha robertson"): "Tabitha",
        _norm("mariyadas, abhish"): "Abhish Mariyadas",
        _norm("abhish m"): "Abhish Mariyadas",
        _norm("m, kasi"): "Kasi M",
        _norm("kasi"): "Kasi M",
        _norm("divya, netti"): "Divya",
        _norm("megan r"): "Megan",
        _norm("patil, shankar"): "Shankar",
        _norm("yadav, trilok"): "Trilok",
        _norm("sengupta, trisha"): "Trisha",
        _norm("rauta, brajendra"): "Brajendra",
        _norm("m g"): "MG",
        _norm("sikkander, nadeem"): "Nadeem",
        _norm("Naidu, Priyadarshini"): "Priyadarshini",
        _norm("nath, koushik"): "Koushik Nath",
        _norm("raviteja, gade"): "Raviteja",
        _norm("iligiri, gopikalyani"): "Gopikalyani Iligiri",
        _norm("gundlapally, sinduja"): "Sinduja",
        _norm("s, sharavanan"): "Sharavanan",
        _norm(" tuniki, nitheesh chandra"): "Chandra",
        _norm("embari, chaitanya"): "Chaitanya",
        _norm("gowda, manjunath"): "Manjunath Gowda",
        _norm("andrew o"): "Andrew",
        _norm("kumar, shailesh"): "Shailesh Kumar",
        _norm("michael"): "Michael F",
        _norm("anu nandyala"): "Anu",
        _norm("kuche"): "Ku Che",
        _norm("goutham kumar, p"): "P Goutham Kumar",
    }
    return aliases.get(key, clean)
ENT_EXCLUSION_START = pd.Timestamp("2026-04-27")
ENT_EXCLUDED_PEOPLE = {
    normalize_person_name(x)
    for x in [
        "Aravind Kumar Badugu",
        "Jagruti Damahe",
        "Naveen Shanmugasundaram",
        "Prabhu S",
    ]
}
def is_excluded_from_team_after_date(team: str, week, person: str) -> bool:
    if str(team or "").strip().upper() != "ENT":
        return False
    wk = pd.to_datetime(week, errors="coerce")
    if pd.isna(wk):
        return False
    wk = wk.normalize()
    if wk < ENT_EXCLUSION_START:
        return False
    return normalize_person_name(str(person or "").strip()) in ENT_EXCLUDED_PEOPLE
def _filter_excluded_people_frame(
    df_in: pd.DataFrame,
    team_col: str = "team",
    week_col: str = "period_date",
    person_col: str = "person",
) -> pd.DataFrame:
    if df_in is None or df_in.empty:
        return df_in
    needed = {team_col, week_col, person_col}
    if not needed.issubset(df_in.columns):
        return df_in
    out = df_in.copy()
    exclude_mask = out.apply(
        lambda r: is_excluded_from_team_after_date(
            r.get(team_col), r.get(week_col), r.get(person_col)
        ),
        axis=1,
    )
    return out.loc[~exclude_mask].copy()
PSS_GROUPS = {
    "US": {"Abby", "Claire", "Nick", "Paige", "Gianna"},
    "MEIC": set(), 
}
def filter_people_df_by_group(df_in: pd.DataFrame, team: str, group_name: str | None) -> pd.DataFrame:
    if df_in is None or df_in.empty or team != "PSS" or not group_name:
        return df_in
    out = df_in.copy()
    if "person" not in out.columns:
        return out
    us_people = {normalize_person_name(x) for x in PSS_GROUPS["US"]}
    out["person"] = out["person"].astype(str).map(normalize_person_name)
    if group_name == "US":
        return out[out["person"].isin(us_people)].copy()
    if group_name == "MEIC":
        return out[~out["person"].isin(us_people)].copy()
    return out
def metric_row_filtered_to_group(row, team: str, group_name: str | None):
    if team != "PSS" or not group_name:
        return row
    us_people = {normalize_person_name(x) for x in PSS_GROUPS["US"]}
    def keep_person(name: str) -> bool:
        n = normalize_person_name(name)
        if group_name == "US":
            return n in us_people
        if group_name == "MEIC":
            return n not in us_people
        return True
    row = row.copy()
    json_person_cols = [
        "Person Hours",
        "Outputs by Person",
        "People in WIP",
        "non_wip_by_person",
    ]
    for col in json_person_cols:
        if col not in row.index or pd.isna(row[col]):
            continue
        try:
            payload = json.loads(row[col]) if isinstance(row[col], str) else row[col]
        except Exception:
            continue
        if isinstance(payload, dict):
            payload = {k: v for k, v in payload.items() if keep_person(str(k))}
        elif isinstance(payload, list):
            payload = [x for x in payload if keep_person(str(x))]
        row[col] = json.dumps(payload)
    if "non_wip_activities" in row.index and pd.notna(row["non_wip_activities"]):
        try:
            acts = json.loads(row["non_wip_activities"]) if isinstance(row["non_wip_activities"], str) else row["non_wip_activities"]
            if isinstance(acts, list):
                acts = [a for a in acts if keep_person(str(a.get("name", "")))]
                row["non_wip_activities"] = json.dumps(acts)
        except Exception:
            pass
    return row
PERSON_WEEKLY_HOURS = {
    "Chelsey": 16.0,
    "MG": 36.0,
    "Lindsey": 32.0,
}
def weekly_hours_for_person(name: str, default: float = 40.0) -> float:
    person = normalize_person_name(name)
    return float(PERSON_WEEKLY_HOURS.get(person, default))
def _postprocess(df: pd.DataFrame) -> pd.DataFrame:
    _NA_STRINGS = {
        "": np.nan, "-": np.nan, "–": np.nan, "—": np.nan,
        "nan": np.nan, "NaN": np.nan, "NAN": np.nan,
        "n/a": np.nan, "N/A": np.nan, "na": np.nan, "NA": np.nan, "null": np.nan, "NULL": np.nan
    }
    if df.empty:
        return df
    def _norm_name(x: str) -> str:
        s = str(x).strip()
        s = " ".join(s.split())
        return s
    df = df.rename(columns={c: _norm_name(c) for c in df.columns})
    canon_map = {}
    for c in list(df.columns):
        lc = c.lower()
        if lc == "hc in wip":
            canon_map[c] = "HC in WIP"
        elif lc in ("actual hc used", "actual_hc_used", "actual-hc-used"):
            canon_map[c] = "Actual HC used"
        elif lc in ("people in wip", "people_wip", "people-in-wip", "people_wip_list"):
            canon_map[c] = "People in WIP"
    if canon_map:
        df = df.rename(columns=canon_map)
    if "period_date" in df.columns:
        df["period_date"] = pd.to_datetime(df["period_date"], errors="coerce").dt.normalize()
    for col in ["Completed Hours", "Actual HC used"]:
        if col in df.columns:
            s = (
                df[col]
                .astype(str)
                .str.strip()
                .replace(_NA_STRINGS)
            )
            df[col] = pd.to_numeric(s, errors="coerce")  
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return df
def accounted_nonwip_by_person_from_row(row) -> tuple[dict[str, float], dict[str, float]]:
    payload = row.get("non_wip_activities", "[]")
    try:
        activities = json.loads(payload) if isinstance(payload, str) else payload
    except Exception:
        activities = []
    if not isinstance(activities, list) or not activities:
        return {}, {}
    team = row.get("team", "")
    wk = row.get("period_date")
    import re
    other_team_key = "OTHERTEAMWIP"
    accounted_other: dict[str, float] = {}
    accounted_nonother: dict[str, float] = {}
    for d in activities:
        name = normalize_person_name(str(d.get("name", "")).strip())
        if not name:
            continue
        if is_excluded_from_team_after_date(team, wk, name):
            continue
        raw_act = str(d.get("activity", "")).strip().upper()
        act_key = canonical_activity_key(d.get("activity", ""))
        if act_key == "OOO" or raw_act in {"OOO", "OUT OF OFFICE", "HOLIDAY"}:
            continue
        try:
            hrs = float(d.get("hours", 0) or 0)
        except Exception:
            hrs = 0.0
        if hrs <= 0:
            continue
        if act_key == other_team_key:
            accounted_other[name] = accounted_other.get(name, 0.0) + hrs
        else:
            accounted_nonother[name] = accounted_nonother.get(name, 0.0) + hrs
    accounted_other = {k: round(v, 2) for k, v in accounted_other.items()}
    accounted_nonother = {k: round(v, 2) for k, v in accounted_nonother.items()}
    return accounted_other, accounted_nonother
def build_ooo_table_from_row(row) -> pd.DataFrame:
    payload = row.get("non_wip_activities", "[]")
    try:
        obj = json.loads(payload) if isinstance(payload, str) else payload
    except Exception:
        obj = []
    if not isinstance(obj, list) or not obj:
        return pd.DataFrame(columns=["Activity", "Name", "Time"])
    team = row.get("team", "")
    wk = row.get("period_date")
    df = pd.DataFrame(obj)
    for c in ["activity", "name", "hours"]:
        if c not in df.columns:
            df[c] = None
    df["hours"] = pd.to_numeric(df["hours"], errors="coerce")
    df["activity"] = (
        df["activity"]
        .astype(str)
        .str.strip()
        .replace({
            "Out of Office": "OOO",
            "Holiday": "OOO",
        })
    )
    out = (
        df.groupby(["activity", "name"], as_index=False)["hours"]
          .sum()
          .rename(columns={"activity": "Activity", "name": "Name", "hours": "HoursRaw"})
          .assign(
              Activity=lambda d: d["Activity"].astype(str).str.strip(),
              Name=lambda d: d["Name"].astype(str).map(normalize_person_name),
          )
    )
    if not out.empty:
        out = out.loc[
            ~out.apply(
                lambda r: is_excluded_from_team_after_date(team, wk, r["Name"]),
                axis=1,
            )
        ].copy()
    out["Time"] = out["HoursRaw"].fillna(0).map(_fmt_hours_minutes)
    out = (
        out[["Activity", "Name", "Time", "HoursRaw"]]
        .sort_values(["Activity", "Name"])
        .reset_index(drop=True)
    )
    return out
def split_nonwip_activity_minutes(cat: pd.DataFrame) -> pd.DataFrame:
    import re
    if cat.empty:
        return cat
    def _canon_activity(label: str) -> str:
        s_orig = str(label or "").strip()
        if not s_orig:
            return s_orig
        s = re.sub(r"\s+", " ", s_orig).strip()
        s = re.sub(r"^[\.\,\;\:\-\–\—\s]+", "", s).strip()
        s = re.sub(r"[:\-\–\—]\s*\d+\s*$", "", s).strip()
        if not s:
            return s
        lower = s.lower()
        compact = re.sub(r"[^a-z0-9]", "", lower)
        if re.fullmatch(r"email(s)?(&|and|/)?im", compact):
            return "Email & IM"
        key = lower
        explicit_map = ACTIVITY_MAP
        if key in explicit_map:
            return explicit_map[key]
        acronym_tokens = {
            "im", "wip", "ooo", "sla", "qa", "hc", "pe", "wfh", "pto",
            "ri", "capa",
        }
        words = lower.split(" ")
        if len(words) == 1:
            w = words[0]
            if w.endswith("s") and not w.endswith("ss") and len(w) > 3:
                w = w[:-1]  # emails -> email
            if w in acronym_tokens:
                return w.upper()
            return w.capitalize()
        last = words[-1]
        if last.endswith("s") and not last.endswith("ss") and len(last) > 3:
            words[-1] = last[:-1]
        pretty = []
        for w in words:
            if not w:
                continue
            if w in acronym_tokens:
                pretty.append(w.upper())
            else:
                pretty.append(w.capitalize())
        return " ".join(pretty)
    rows: list[dict] = []
    for _, r in cat.iterrows():
        activity_text = str(r["Activity"])
        total_hours_raw = pd.to_numeric(r["Hours"], errors="coerce")
        total_hours = float(total_hours_raw) if pd.notna(total_hours_raw) else 0.0
        s = activity_text.replace(";", " ").replace(",", " ").replace(":", " ")
        s = re.sub(r"\s+", " ", s).strip()
        if not s:
            rows.append({"Activity": _canon_activity(activity_text), "Hours": total_hours})
            continue
        pattern = re.compile(
            r"(?P<num>\d+)\s*(?P<unit>h|hr|hrs|hour|hours|m|min|mins|minute|minutes)?\b",
            re.IGNORECASE,
        )
        sub_acts: list[tuple[str, int]] = []
        prev_end = 0
        for m in pattern.finditer(s):
            num = int(m.group("num"))
            unit = (m.group("unit") or "").lower()
            mins = num * 60 if unit in ("h", "hr", "hrs", "hour", "hours") else num
            label = s[prev_end:m.start()]
            prev_end = m.end()
            label = label.strip()
            if not label:
                continue
            label = re.sub(r"\([^)]*$", "", label)   
            label = re.sub(r"\(.*?\)", "", label)    
            label = re.sub(r"[:\-–—]+$", "", label)  
            label = label.strip(" ,;:()[]-–—")
            label = re.sub(r"\s+", " ", label).strip()
            label = _canon_activity(label)
            if label and mins > 0:
                sub_acts.append((label, mins))
        if sub_acts:
            has_delims = bool(re.search(r"[;,]", activity_text))
            if len(sub_acts) == 1 and not has_delims:
                rows.append({"Activity": _canon_activity(activity_text), "Hours": total_hours})
                continue
            total_minutes = sum(m for _, m in sub_acts)
            if total_hours <= 0 and total_minutes > 0:
                total_hours = total_minutes / 60.0
            if total_hours > 0 and total_minutes > 0:
                for label, mins in sub_acts:
                    h_sub = total_hours * (mins / total_minutes)
                    rows.append({"Activity": label, "Hours": h_sub})
            else:
                rows.append({"Activity": _canon_activity(activity_text), "Hours": total_hours})
        else:
            rows.append({"Activity": _canon_activity(activity_text), "Hours": total_hours})
    out = pd.DataFrame(rows)
    if out.empty:
        return cat
    out["Activity"] = out["Activity"].map(_canon_activity)
    return out.groupby("Activity", as_index=False)["Hours"].sum()
def ahu_person_share_for_week(frame: pd.DataFrame, week, teams_in_view: list[str], people_df: pd.DataFrame) -> pd.DataFrame:
    if frame.empty or "Actual HC used" not in frame.columns:
        return pd.DataFrame(columns=["team", "period_date", "person", "percent"])
    wk = pd.to_datetime(week, errors="coerce").normalize()
    if pd.isna(wk):
        return pd.DataFrame(columns=["team", "period_date", "person", "percent"])
    ppl = explode_people_in_wip(frame)
    out_rows: list[dict] = []
    for team in teams_in_view:
        team_ahu_series = (
            frame.loc[(frame["team"] == team) & (frame["period_date"] == wk), "Actual HC used"]
            .dropna()
        )
        if team_ahu_series.empty:
            continue
        per_df = None
        if people_df is not None and not people_df.empty:
            teamw = people_df.loc[
                (people_df["team"] == team) & (people_df["period_date"] == wk)
            ]
            if not teamw.empty and teamw["Actual Hours"].notna().any():
                g = teamw.groupby("person", as_index=False)["Actual Hours"].sum()
                tot = float(g["Actual Hours"].sum())
                if tot > 0:
                    per_df = g.assign(percent=lambda d: d["Actual Hours"] / tot)[["person", "percent"]]
        if per_df is None:
            sub = ppl.loc[(ppl["team"] == team) & (ppl["period_date"] == wk)]
            if not sub.empty:
                unique_people = sub["person"].dropna().drop_duplicates().tolist()
                n = len(unique_people)
                if n > 0:
                    per_df = pd.DataFrame({"person": unique_people, "percent": [1.0 / n] * n})
        if per_df is None or per_df.empty:
            continue
        per_df = per_df.assign(team=team, period_date=wk)
        out_rows.append(per_df[["team", "period_date", "person", "percent"]])
    if not out_rows:
        return pd.DataFrame(columns=["team", "period_date", "person", "percent"])
    return pd.concat(out_rows, ignore_index=True)
@st.cache_data(show_spinner=False, ttl=15 * 60)
def explode_people_in_wip(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "People in WIP" not in df.columns:
        return pd.DataFrame(columns=["team", "period_date", "person"])
    sub = df.loc[:, ["team", "period_date", "People in WIP"]].dropna(subset=["People in WIP"]).copy()
    BAD_NAMES = {"", "-", "–", "—", "nan", "NaN", "NAN", "n/a", "N/A", "na", "NA", "null", "NULL", "none", "None"}
    def _is_good_name(s: str) -> bool:
        return s.strip() and s.strip() not in BAD_NAMES
    rows: list[dict] = []
    def _as_names(x) -> list[str]:
        if isinstance(x, list):
            return [str(s).strip() for s in x if _is_good_name(str(s))]
        if isinstance(x, str):
            s = x.strip()
            try:
                obj = json.loads(s)
                if isinstance(obj, list):
                    return [str(v).strip() for v in obj if _is_good_name(str(v))]
                if isinstance(obj, dict):
                    return [str(k).strip() for k, v in obj.items() if _is_good_name(str(k))]
            except Exception:
                pass
            import re
            parts = [p.strip() for p in re.split(r"[,;\n\r]+", s) if _is_good_name(p)]
            return parts
        if isinstance(x, dict):
            return [str(k).strip() for k in x.keys() if _is_good_name(str(k))]
        return []
    for _, r in sub.iterrows():
        people = _as_names(r["People in WIP"])
        for person in people:
            rows.append({
                "team": r["team"],
                "period_date": pd.to_datetime(r["period_date"], errors="coerce").normalize(),
                "person": normalize_person_name(person)
            })
    out = pd.DataFrame(rows, columns=["team", "period_date", "person"])
    if not out.empty:
        out["period_date"] = pd.to_datetime(out["period_date"], errors="coerce").dt.normalize()
    out = _filter_excluded_people_frame(out)
    return out
@st.cache_data(show_spinner=False, ttl=15 * 60)
def explode_person_hours(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "Person Hours" not in df.columns:
        return pd.DataFrame(columns=[
            "team","period_date","person","Actual Hours","Available Hours","Utilization"
        ])
    BAD_NAMES = {"", "-", "–", "—", "nan", "NaN", "NAN", "n/a", "N/A", "na", "NA",
                 "null", "NULL", "none", "None"}
    def _is_good_name(s: str) -> bool:
        t = str(s).strip()
        return t and t not in BAD_NAMES and t.lower() not in {b.lower() for b in BAD_NAMES}
    rows: list[dict] = []
    sub = df.loc[:, ["team", "period_date", "Person Hours"]].dropna(subset=["Person Hours"]).copy()
    for _, r in sub.iterrows():
        payload = r["Person Hours"]
        try:
            obj = json.loads(payload) if isinstance(payload, str) else payload
            if not isinstance(obj, dict):
                continue
        except Exception:
            continue
        for person, vals in obj.items():
            if not _is_good_name(person):
                continue
            a = pd.to_numeric((vals or {}).get("actual"), errors="coerce")
            t = pd.to_numeric((vals or {}).get("available"), errors="coerce")
            a = float(a) if pd.notna(a) else 0.0
            t = float(t) if pd.notna(t) else 0.0
            if (a == 0.0) and (t == 0.0):
                continue
            util = (a / t) if t not in (0, 0.0) else np.nan
            rows.append({
                "team": r["team"],
                "period_date": pd.to_datetime(r["period_date"], errors="coerce").normalize(),
                "person": normalize_person_name(str(person).strip()),
                "Actual Hours": a,
                "Available Hours": t,
                "Utilization": util
            })
    out = pd.DataFrame(
        rows,
        columns=["team","period_date","person","Actual Hours","Available Hours","Utilization"]
    )
    if not out.empty:
        out["period_date"] = pd.to_datetime(out["period_date"], errors="coerce").dt.normalize()
    out = _filter_excluded_people_frame(out)
    return out
def build_person_weekly_accounting(
    team: str,
    week,
    nw_row,
    metrics_frame: pd.DataFrame,
    nw_frame: pd.DataFrame,
    week_hours: float = 40.0,
    irl_people: set[str] | None = None,
) -> pd.DataFrame:
    wk = pd.to_datetime(week, errors="coerce").normalize()
    long_nw = explode_non_wip_by_person(nw_frame)
    nw_people = long_nw.loc[
        (long_nw["team"] == team) & (long_nw["period_date"] == wk),
        ["person", "Non-WIP Hours"]
    ].copy()
    if nw_people.empty:
        nw_people = pd.DataFrame(columns=["person", "Non-WIP Hours"])
    nw_people["person"] = nw_people["person"].astype(str).str.strip()
    nw_people["Non-WIP Hours"] = pd.to_numeric(nw_people["Non-WIP Hours"], errors="coerce").fillna(0.0)
    person_hours = explode_person_hours(metrics_frame)
    wip_people = person_hours.loc[
        (person_hours["team"] == team) & (person_hours["period_date"] == wk),
        ["person", "Actual Hours"]
    ].copy()
    if wip_people.empty:
        wip_people = pd.DataFrame(columns=["person", "Actual Hours"])
    wip_people["person"] = wip_people["person"].astype(str).str.strip()
    wip_people["Completed Hours"] = pd.to_numeric(wip_people["Actual Hours"], errors="coerce").fillna(0.0)
    wip_people = wip_people.drop(columns=["Actual Hours"], errors="ignore")
    acct_other_map, acct_nonother_map = accounted_nonwip_by_person_from_row(nw_row)
    other_df = pd.DataFrame(
        [{"person": str(k).strip(), "Other Team WIP": float(v)} for k, v in acct_other_map.items()]
    )
    if other_df.empty:
        other_df = pd.DataFrame(columns=["person", "Other Team WIP"])
    acct_df = pd.DataFrame(
        [{"person": str(k).strip(), "Accounted Non-WIP": float(v)} for k, v in acct_nonother_map.items()]
    )
    if acct_df.empty:
        acct_df = pd.DataFrame(columns=["person", "Accounted Non-WIP"])
    payload = nw_row.get("non_wip_activities", "[]")
    try:
        activities = json.loads(payload) if isinstance(payload, str) else payload
    except Exception:
        activities = []
    ooo_by_person: dict[str, float] = {}
    if isinstance(activities, list):
        for item in activities:
            if not isinstance(item, dict):
                continue
            person = normalize_person_name(str(item.get("name", "")).strip())
            if is_excluded_from_team_after_date(team, wk, person):
                continue
            activity_key = canonical_activity_key(item.get("activity", ""))
            try:
                hrs = float(item.get("hours", 0) or 0)
            except Exception:
                hrs = 0.0
            if not person or hrs <= 0:
                continue
            if activity_key == "OOO":
                ooo_by_person[person] = ooo_by_person.get(person, 0.0) + hrs
    ooo_df = pd.DataFrame(
        [{"person": k, "OOO Hours": round(v, 2)} for k, v in ooo_by_person.items()]
    )
    if ooo_df.empty:
        ooo_df = pd.DataFrame(columns=["person", "OOO Hours"])
    def _clean_person_col(df_in: pd.DataFrame, value_col: str) -> pd.DataFrame:
        if df_in.empty:
            return pd.DataFrame(columns=["person", value_col])
        out = df_in.copy()
        out["person"] = (
            out["person"]
            .astype("string")
            .fillna("")
            .map(lambda x: normalize_person_name(str(x).strip()))
        )
        out["person"] = out["person"].replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
        out = out.dropna(subset=["person"]).copy()
        out[value_col] = pd.to_numeric(out[value_col], errors="coerce").fillna(0.0)
        return out[["person", value_col]]
    nw_people = _clean_person_col(nw_people, "Non-WIP Hours")
    wip_people = _clean_person_col(wip_people, "Completed Hours")
    other_df = _clean_person_col(other_df, "Other Team WIP")
    acct_df = _clean_person_col(acct_df, "Accounted Non-WIP")
    ooo_df = _clean_person_col(ooo_df, "OOO Hours")
    nw_people = nw_people.groupby("person", as_index=False)["Non-WIP Hours"].sum()
    wip_people = wip_people.groupby("person", as_index=False)["Completed Hours"].sum()
    other_df = other_df.groupby("person", as_index=False)["Other Team WIP"].sum()
    acct_df = acct_df.groupby("person", as_index=False)["Accounted Non-WIP"].sum()
    ooo_df = ooo_df.groupby("person", as_index=False)["OOO Hours"].sum()
    all_people = sorted(
        set(nw_people["person"].tolist())
        | set(wip_people["person"].tolist())
        | set(other_df["person"].tolist())
        | set(acct_df["person"].tolist())
        | set(ooo_df["person"].tolist())
    )
    people = pd.DataFrame({"person": pd.Series(all_people, dtype="string")})
    out = (
        people
        .merge(nw_people.astype({"person": "string"}), on="person", how="left")
        .merge(wip_people.astype({"person": "string"}), on="person", how="left")
        .merge(other_df.astype({"person": "string"}), on="person", how="left")
        .merge(acct_df.astype({"person": "string"}), on="person", how="left")
        .merge(ooo_df.astype({"person": "string"}), on="person", how="left")
    )
    numeric_cols = [
        "Non-WIP Hours",
        "Completed Hours",
        "Other Team WIP",
        "Accounted Non-WIP",
        "OOO Hours",
    ]
    for col in numeric_cols:
        if col not in out.columns:
            out[col] = 0.0
        out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0.0)
    out["person"] = out["person"].astype("string")
    out["person_key"] = out["person"].astype(str).str.strip().str.lower()
    irl_people_norm = {str(x).strip().lower() for x in (irl_people or set())}
    PERSON_WEEKLY_HOURS = {
        "chelsey": 16.0,
        "mg": 36.0,
        "lindsey": 32.0,
    }
    def get_expected_hours(person_key: str, default_hours: float, irl_people_norm: set[str]) -> float:
        if person_key in PERSON_WEEKLY_HOURS:
            return PERSON_WEEKLY_HOURS[person_key]
        if person_key in irl_people_norm:
            return 39.0
        return float(default_hours)
    out["Expected Hours"] = out["person_key"].apply(
        lambda p: get_expected_hours(p, week_hours, irl_people_norm)
    )
    calc_cols = [
        "Expected Hours",
        "OOO Hours",
        "Non-WIP Hours",
        "Completed Hours",
        "Other Team WIP",
        "Accounted Non-WIP",
    ]
    for col in calc_cols:
        out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0.0).astype("float64")
    non_ooo_total = out["Non-WIP Hours"].clip(lower=0.0)
    out["Other Team WIP"] = np.minimum(out["Other Team WIP"], non_ooo_total).astype("float64")
    remaining_nonwip = (non_ooo_total - out["Other Team WIP"]).clip(lower=0.0)
    out["Accounted Non-WIP"] = np.minimum(out["Accounted Non-WIP"], remaining_nonwip).astype("float64")
    out["Unaccounted"] = (
        out["Expected Hours"]
        - out["Completed Hours"]
        - out["OOO Hours"]
        - out["Other Team WIP"]
        - out["Accounted Non-WIP"]
    ).clip(lower=0.0).astype("float64")
    out["Total Used"] = (
        out["Completed Hours"]
        + out["OOO Hours"]
        + out["Other Team WIP"]
        + out["Accounted Non-WIP"]
    ).astype("float64")
    out["period_date"] = wk
    out["team"] = team
    return out.sort_values(["person"]).reset_index(drop=True)
data_path = None if DATA_URL else str(DEFAULT_DATA_PATH)
mtime_key = 0
if data_path:
    p = Path(data_path)
    mtime_key = p.stat().st_mtime if p.exists() else 0
df = load_data(data_path, DATA_URL)
def _first_valid_team(value, options):
    if value in options:
        return value
    return options[0] if options else None
_all_wip_teams = (
    sorted([t for t in df["team"].dropna().unique()])
    if not df.empty and "team" in df.columns
    else []
)
if "selected_team" not in st.session_state:
    existing_teams = st.session_state.get("teams_sel", [])
    if existing_teams:
        st.session_state.selected_team = existing_teams[0]
    elif _all_wip_teams:
        st.session_state.selected_team = _all_wip_teams[0]
PSS_GROUP_OPTIONS = ["All", "US", "MEIC"]
def _first_valid_pss_group(value):
    return value if value in PSS_GROUP_OPTIONS else "All"
if "selected_pss_group" not in st.session_state:
    st.session_state.selected_pss_group = _first_valid_pss_group(
        st.session_state.get("pss_group", "All")
    )
def kpi_card(
    container,
    label: str,
    value,
    fmt: str | None = None,
    color: str | None = None,
    help: str | None = None,
    subtext: str | None = None,
):
    if pd.isna(value):
        val_html = "—"
    else:
        try:
            val_html = (fmt or "{}").format(value)
        except Exception:
            val_html = str(value)
    help_icon = f"""<span title="{help}" style="cursor:help;margin-left:6px;color:#9ca3af;">ⓘ</span>""" if help else ""
    value_color = color or "#111827"
    subtext_html = f"""<div style="font-size:12px;color:#6b7280;margin-top:4px;">{subtext}</div>""" if subtext else ""
    container.markdown(
        f"""
        <div style="padding:12px 16px;border-radius:10px;border:1px solid #eee;">
          <div style="font-size:12px;color:#6b7280;display:flex;align-items:center;gap:4px;">
            <span>{label}</span>{help_icon}
          </div>
          <div style="font-size:28px;font-weight:700;color:{value_color};">{val_html}</div>
          {subtext_html}
        </div>
        """,
        unsafe_allow_html=True,
    )
def _capacity_subtext(hours_val, capacity_val) -> str | None:
    if pd.isna(hours_val) or pd.isna(capacity_val) or float(capacity_val) <= 0:
        return None
    pct = float(hours_val) / float(capacity_val)
    hrs_per_day = pct * 8.0
    return f"{pct:.1%} of capacity • {hrs_per_day:.1f}h/day"
def merged_people_count_for_week(team: str, week, metrics_frame: pd.DataFrame, nw_frame: pd.DataFrame) -> int:
    wk = pd.to_datetime(week, errors="coerce").normalize()
    if nw_frame is not None and not nw_frame.empty:
        raw_nw = nw_frame.copy()
        raw_nw["period_date"] = pd.to_datetime(raw_nw["period_date"], errors="coerce").dt.normalize()
        if "people_count" in raw_nw.columns:
            team_match = raw_nw.loc[
                (raw_nw["team"] == team) & (raw_nw["period_date"] == wk),
                "people_count",
            ]
            team_match = pd.to_numeric(team_match, errors="coerce").dropna()
            if not team_match.empty:
                return int(team_match.iloc[0])
    a = explode_non_wip_by_person(nw_frame)
    b = explode_person_hours(metrics_frame)
    c = explode_people_in_wip(metrics_frame)
    names = set()
    for df_, person_col in [(a, "person"), (b, "person"), (c, "person")]:
        sub = df_.loc[
            (df_["team"] == team) & (df_["period_date"] == wk),
            [person_col]
        ].copy()
        if not sub.empty:
            vals = (
                sub[person_col]
                .astype(str)
                .map(normalize_person_name)
                .str.strip()
            )
            names.update(x for x in vals if x)
    return len(names)
def percent_color(v: float | None, threshold: float, invert: bool = False) -> str:
    if v is None or pd.isna(v):
        return "#111827"
    good = (v >= threshold) if not invert else (v <= threshold)
    return "#22c55e" if good else "#ef4444"
st.markdown("<h1 style='text-align: center;'>NS Heijunka Metrics Dashboard</h1>", unsafe_allow_html=True)
def ent_capacity_hours_for_week(
    team: str,
    week,
    nw_frame: pd.DataFrame,
    irl_people: set[str] | None = None,
) -> float:
    wk = pd.to_datetime(week, errors="coerce").normalize()
    irl_people_norm = {str(x).strip().lower() for x in (irl_people or set())}
    if nw_frame is None or nw_frame.empty:
        return 0.0
    raw_nw = nw_frame.copy()
    raw_nw["period_date"] = pd.to_datetime(
        raw_nw["period_date"], errors="coerce"
    ).dt.normalize()
    row = raw_nw.loc[
        (raw_nw["team"] == team) & (raw_nw["period_date"] == wk)
    ]
    if row.empty:
        return 0.0
    people_count_series = pd.to_numeric(
        row["people_count"], errors="coerce"
    ).dropna()
    people_count = int(people_count_series.iloc[0]) if not people_count_series.empty else 0
    irl_count = 0
    if "non_wip_by_person" in row.columns:
        payload = row.iloc[0].get("non_wip_by_person")
        try:
            obj = json.loads(payload) if isinstance(payload, str) else payload
        except Exception:
            obj = {}
        if isinstance(obj, dict):
            names = {
                normalize_person_name(str(k).strip()).strip().lower()
                for k in obj.keys()
                if str(k).strip()
            }
            irl_count = sum(1 for n in names if n in irl_people_norm)
    irl_count = min(irl_count, people_count)
    non_irl_count = max(people_count - irl_count, 0)
    return float((irl_count * 39.0) + (non_irl_count * 40.0))
label = "Show WIP view" if st.session_state.get("nonwip_mode", False) else "Show Non-WIP view"
nonwip_mode = st.toggle(
    label,
    value=st.session_state.get("nonwip_mode", False),
    key="nonwip_mode",
    help="Switch between WIP and Non-WIP metrics"
)
if nonwip_mode:
    nw = load_non_wip()
    if nw.empty:
        st.info("No Non-WIP data found yet. Make sure non_wip_activities.csv exists.")
        st.stop()
    st.markdown("### Non-WIP Overview")
    teams_nw = sorted([t for t in nw["team"].dropna().unique()])
    c_team, c_week = st.columns(2)
    preferred_team = _first_valid_team(
        st.session_state.get("selected_team"),
        teams_nw,
    )
    if preferred_team is not None:
        st.session_state["nw_team"] = preferred_team
    if "selected_pss_group" not in st.session_state:
        st.session_state.selected_pss_group = "All"
    def _sync_from_nonwip_team():
        team = st.session_state.get("nw_team")
        if team:
            st.session_state.selected_team = team
            st.session_state.teams_sel = [team]
    def _sync_from_pss_group():
        st.session_state.selected_pss_group = _first_valid_pss_group(
            st.session_state.get("pss_group", "All")
        )
    with c_team:
        team_nw = st.selectbox(
            "Team",
            options=teams_nw,
            key="nw_team",
            on_change=_sync_from_nonwip_team,
        )
        pss_group = None
        if team_nw == "PSS":
            st.session_state["pss_group"] = _first_valid_pss_group(
                st.session_state.get("selected_pss_group", "All")
            )
            pss_group = st.selectbox(
                "Group",
                options=PSS_GROUP_OPTIONS,
                key="pss_group",
                on_change=_sync_from_pss_group,
            )
        else:
            pss_group = None
    st.session_state.selected_team = team_nw
    st.session_state.teams_sel = [team_nw]
    if team_nw == "PSS":
        st.session_state.selected_pss_group = _first_valid_pss_group(pss_group)
    today_nw = pd.Timestamp.today().normalize()
    weeks_nw = sorted(
        [
            d for d in pd.to_datetime(
                nw.loc[nw["team"] == team_nw, "period_date"].dropna().unique()
            )
            if pd.notna(d) and pd.to_datetime(d).normalize() <= today_nw
        ],
        reverse=True
    )
    if not weeks_nw:
        st.info("No weeks available for this team up to today.")
        st.stop()
    with c_week:
        week_nw = st.selectbox(
            "Week",
            options=weeks_nw,
            index=0,
            format_func=lambda d: pd.to_datetime(d).date().isoformat(),
            key="nw_week",
        )
    week_nw = pd.to_datetime(week_nw).normalize()
    sel = nw[(nw["team"] == team_nw) & (nw["period_date"] == week_nw)]
    if sel.empty:
        st.info("No Non-WIP row for that team/week.")
        st.stop()
    row = sel.iloc[0]
    row = metric_row_filtered_to_group(row, team_nw, pss_group if team_nw == "PSS" else None)
    if "% Non-WIP" in row.index and pd.notna(row["% Non-WIP"]):
        pct_non_wip = float(row["% Non-WIP"])
    else:
        pct_in_wip = float(row.get("% in WIP", np.nan))
        pct_non_wip = (100.0 - pct_in_wip) if pd.notna(pct_in_wip) else np.nan
    include_ooo_in_kpi_pct = st.toggle(
        "Include OOO Hours in KPI % of capacity",
        value=False,
        key="include_ooo_in_kpi_pct",
        help="When off, OOO Hours shows 0.0% of capacity and other KPI percentages are calculated against capacity excluding OOO hours.",
    )
    def colored_percent_metric(container, label: str, value: float | None, threshold=80.0):
        if pd.isna(value):
            container.metric(label, "—")
            return
        color = "#ef4444" if float(value) < threshold else "#22c55e"
        container.markdown(
            f"""
            <div style="padding:12px 16px;border-radius:10px;border:1px solid #eee;">
            <div style="font-size:12px;color:#6b7280;">{label}</div>
            <div style="font-size:28px;font-weight:700;color:{color};">{value:.2f}%</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    c1, c2, c3, c4, c5 = st.columns(5)
    teams_cfg = load_team_config()
    team_irl_people = irl_people_for_team(team_nw, teams_cfg)
    wk_people_kpi = build_person_weekly_accounting(
        team=team_nw,
        week=week_nw,
        nw_row=row,            # already filtered for PSS group
        metrics_frame=df,
        nw_frame=nw,
        week_hours=40.0,
        irl_people=team_irl_people,
    )
    if team_nw == "PSS":
        wk_people_kpi = filter_people_df_by_group(wk_people_kpi, team_nw, pss_group)
    wk_people_kpi = wk_people_kpi.copy()
    num_cols = [
        "Completed Hours",
        "Non-WIP Hours",
        "OOO Hours",
        "Other Team WIP",
        "Accounted Non-WIP",
        "Unaccounted",
        "Expected Hours",
    ]
    for col in num_cols:
        if col in wk_people_kpi.columns:
            wk_people_kpi[col] = pd.to_numeric(wk_people_kpi[col], errors="coerce").fillna(0.0)
    people_count_val = (
        float(
            wk_people_kpi["person"]
            .astype(str)
            .str.strip()
            .replace("", pd.NA)
            .dropna()
            .nunique()
        )
        if not wk_people_kpi.empty and "person" in wk_people_kpi.columns
        else np.nan
    )
    wip_hours_val = (
        float(wk_people_kpi["Completed Hours"].sum())
        if not wk_people_kpi.empty and "Completed Hours" in wk_people_kpi.columns
        else np.nan
    )
    other_team_wip_hours_val = (
        float(wk_people_kpi["Other Team WIP"].sum())
        if not wk_people_kpi.empty and "Other Team WIP" in wk_people_kpi.columns
        else 0.0
    )
    nonwip_hours_val = (
        float(wk_people_kpi["Accounted Non-WIP"].sum())
        if not wk_people_kpi.empty and "Accounted Non-WIP" in wk_people_kpi.columns
        else np.nan
    )
    ooo_hours_val = (
        float(wk_people_kpi["OOO Hours"].sum())
        if not wk_people_kpi.empty and "OOO Hours" in wk_people_kpi.columns
        else 0.0
    )
    unaccounted_hours_val = (
        float(wk_people_kpi["Unaccounted"].sum())
        if not wk_people_kpi.empty and "Unaccounted" in wk_people_kpi.columns
        else np.nan
    )
    capacity_val = (
        float(wk_people_kpi["Expected Hours"].sum())
        if not wk_people_kpi.empty and "Expected Hours" in wk_people_kpi.columns
        else np.nan
    )
    capacity_pct_basis = capacity_val
    if not include_ooo_in_kpi_pct and pd.notna(capacity_val):
        capacity_pct_basis = max(float(capacity_val) - float(ooo_hours_val), 0.0)
    wip_pct = (
        wip_hours_val / capacity_pct_basis
        if pd.notna(wip_hours_val) and pd.notna(capacity_pct_basis) and capacity_pct_basis > 0
        else np.nan
    )
    other_team_wip_pct = (
        other_team_wip_hours_val / capacity_pct_basis
        if pd.notna(other_team_wip_hours_val) and pd.notna(capacity_pct_basis) and capacity_pct_basis > 0
        else np.nan
    )
    nonwip_pct = (
        nonwip_hours_val / capacity_pct_basis
        if pd.notna(nonwip_hours_val) and pd.notna(capacity_pct_basis) and capacity_pct_basis > 0
        else np.nan
    )
    ooo_pct = (
        (ooo_hours_val / capacity_pct_basis)
        if include_ooo_in_kpi_pct and pd.notna(ooo_hours_val) and pd.notna(capacity_pct_basis) and capacity_pct_basis > 0
        else 0.0
    )
    unaccounted_pct = (
        unaccounted_hours_val / capacity_pct_basis
        if pd.notna(unaccounted_hours_val) and pd.notna(capacity_pct_basis) and capacity_pct_basis > 0
        else np.nan
    )
    _enterprise_kpi = enterprise_nonwip_kpi_lookup(
        team=team_nw,
        week=week_nw,
        nw_row=row,
        wk_people=wk_people_kpi,
        people_count=locals().get("people_count_merged", locals().get("people_count_val", np.nan)),
        completed_hours=wip_hours_val,
        total_non_wip_hours=locals().get("total_nonwip_hours_val", locals().get("nonwip_hours_val", np.nan)),
        factor_out_ooo=not include_ooo_in_kpi_pct,
        person_hours=locals().get("_ppl_hours_kpi"),
        people_in_wip=locals().get("_ppl_in_wip_kpi"),
        nw_frame=nw,
        metrics_frame=df,
        ent_capacity_callback=globals().get("ent_capacity_hours_for_week"),
        ent_capacity_kwargs={
            "team": team_nw,
            "week": week_nw,
            "nw_frame": nw,
            "irl_people": locals().get("team_irl_people", set()),
        },
    )
    capacity_val = _enterprise_kpi["capacity_hours"]
    capacity_pct_basis = _enterprise_kpi["pct_denom"]
    wip_hours_val = _enterprise_kpi["completed_hours"]
    other_team_wip_hours_val = _enterprise_kpi["other_team_wip_hours"]
    nonwip_hours_val = _enterprise_kpi["non_wip_hours"]
    ooo_hours_val = _enterprise_kpi["ooo_hours"]
    unaccounted_hours_val = _enterprise_kpi["unaccounted_hours"]
    wip_pct = _enterprise_kpi["wip_pct"]
    other_team_wip_pct = _enterprise_kpi["other_team_wip_pct"]
    nonwip_pct = _enterprise_kpi["non_wip_pct"]
    ooo_pct = _enterprise_kpi["ooo_pct"]
    unaccounted_pct = _enterprise_kpi["unaccounted_pct"]
    kpi_card(
        c1,
        "WIP Hours",
        wip_hours_val,
        fmt="{:,.1f}",
        color=percent_color(wip_pct, threshold=0.80, invert=False),
        subtext=_capacity_subtext(wip_hours_val, capacity_pct_basis),
    )
    kpi_card(
        c2,
        "Other Team WIP",
        other_team_wip_hours_val,
        fmt="{:,.1f}",
        subtext=_capacity_subtext(other_team_wip_hours_val, capacity_pct_basis),
    )
    kpi_card(
        c3,
        "Non-WIP Hours",
        nonwip_hours_val,
        fmt="{:,.1f}",
        color=percent_color(nonwip_pct, threshold=0.20, invert=True),
        subtext=_capacity_subtext(nonwip_hours_val, capacity_pct_basis),
    )
    kpi_card(
        c4,
        "OOO Hours",
        ooo_hours_val,
        fmt="{:,.1f}",
        subtext=_capacity_subtext(
            0.0 if not include_ooo_in_kpi_pct else ooo_hours_val,
            capacity_pct_basis,
        ),
    )
    kpi_card(
        c5,
        "Unaccounted Hours",
        unaccounted_hours_val,
        fmt="{:,.1f}",
        subtext=_capacity_subtext(unaccounted_hours_val, capacity_pct_basis),
    )
    st.markdown("---")
    st.markdown("#### Non-WIP Activities")
    if "non_wip_activities" not in row.index or row.get("non_wip_activities", "") in ("", "[]", None):
        st.info("No Non-WIP activities recorded for this selection.")
    else:
        act_tbl = build_ooo_table_from_row(row)
        if act_tbl.empty:
            st.info("No Non-WIP activities recorded for this selection.")
        else:
            display_tbl = act_tbl.drop(columns=["HoursRaw"], errors="ignore")
            st.dataframe(display_tbl, width="stretch", hide_index=True)
    teams_cfg = load_team_config()
    team_irl_people = irl_people_for_team(team_nw, teams_cfg)
    wk_people = build_person_weekly_accounting(
        team=team_nw,
        week=week_nw,
        nw_row=row,
        metrics_frame=df,
        nw_frame=nw,
        week_hours=40.0,
        irl_people=team_irl_people,
    )
    if team_nw == "PSS":
        wk_people = filter_people_df_by_group(wk_people, team_nw, pss_group)
    if wk_people.empty:
        st.info("No per-person weekly breakdown for this selection.")
    else:
        wk_people = wk_people.rename(columns={
            "Other Team WIP": "Accounted_Other",
            "Accounted Non-WIP": "Accounted_NonOther",
        })
        stack = (
            wk_people.melt(
                id_vars=["person", "period_date", "Non-WIP Hours", "Completed Hours"],
                value_vars=["OOO Hours","Accounted_Other", "Accounted_NonOther", "Unaccounted"],
                var_name="Category",
                value_name="Hours"
            )
            .dropna(subset=["Hours"])
        )
        stack = stack.merge(
            wk_people[[
                "person",
                "Completed Hours",
                "Non-WIP Hours",
                "OOO Hours",
                "Accounted_Other",
                "Accounted_NonOther",
                "Unaccounted"
            ]],
            on="person",
            how="left",
        )
        label_map = {
            "OOO Hours": "OOO",
            "Accounted_Other": "Other Team WIP",
            "Accounted_NonOther": "Accounted Non-WIP",
            "Unaccounted": "Unaccounted",
        }
        stack["CategoryLabel"] = stack["Category"].map(label_map)
        wk_people["StackTotal"] = (
            wk_people["OOO Hours"].fillna(0)
            + wk_people["Accounted_Other"].fillna(0)
            + wk_people["Accounted_NonOther"].fillna(0)
            + wk_people["Unaccounted"].fillna(0)
        )
        order_people = wk_people.sort_values("StackTotal", ascending=False)["person"].tolist()
        vmax = float(pd.to_numeric(wk_people["StackTotal"], errors="coerce").max())
        headroom = max(1.0, vmax * 0.18) if pd.notna(vmax) else 1.0
        y_scale = alt.Scale(domain=[0, vmax + headroom], nice=False, clamp=False)
        totals = (
            wk_people[["person", "period_date", "StackTotal"]]
            .rename(columns={"StackTotal": "Total"})
            .assign(Status=lambda d: np.where(d["Total"] <= 7.5, "Good (≤7.5)", "Over (>7.5)"))
        )
        outline = (
            alt.Chart(totals)
            .mark_bar(fillOpacity=0, strokeWidth=2)
            .encode(
                x=alt.X("person:N", sort=order_people),
                y=alt.Y("Total:Q", scale=y_scale),
                stroke=alt.Color(
                    "Status:N",
                    title="Total vs 7.5",
                    scale=alt.Scale(
                        domain=["Good (≤7.5)", "Over (>7.5)"],
                        range=["#22c55e", "#ef4444"],
                    ),
                ),
            )
        )
        bars = (
            alt.Chart(stack)
            .mark_bar(clip=False)
            .encode(
                x=alt.X(
                    "person:N",
                    title="Person",
                    sort=order_people,
                    axis=alt.Axis(
                        labelAngle=-35,
                        labelLimit=0,
                        labelOverlap=False,
                    ),
                ),
                y=alt.Y(
                    "Hours:Q",
                    title="Non-WIP Hours (week)",
                    stack="zero",
                    scale=y_scale,
                ),
                color=alt.Color(
                    "CategoryLabel:N",
                    title="Legend",
                    scale=alt.Scale(
                        domain=["OOO", "Other Team WIP", "Accounted Non-WIP", "Unaccounted"],
                        range=["#a855f7", "#2563eb", "#22c55e", "#9ca3af"],
                    ),
                ),
                tooltip=[
                    alt.Tooltip("person:N", title="Person"),
                    alt.Tooltip("Accounted_Other:Q", title="Other Team WIP Hours", format=",.2f"),
                    alt.Tooltip("Accounted_NonOther:Q", title="Accounted Non-WIP Hours", format=",.2f"),
                    alt.Tooltip("Unaccounted:Q", title="Unaccounted Hours", format=",.2f"),
                    alt.Tooltip("OOO Hours:Q", title="OOO Hours", format=",.2f"),
                    alt.Tooltip("period_date:T", title="Week"),
                ],
            )
        )
        ref = (
            alt.Chart(pd.DataFrame({"y": [7.5]}))
            .mark_rule(strokeDash=[4, 3], color="#6b7280")
            .encode(y=alt.Y("y:Q", scale=y_scale))
        )
        chart = (outline + ref + bars) \
            .properties(
                height=340,
                padding={"left": 8, "right": 12, "top": 36, "bottom": 64},
            ) \
            .configure_axis(labelOverlap=True) \
            .configure_view(stroke=None)
        st.altair_chart(chart, width="stretch")
        st.markdown("#### Non-WIP Activities")
        if "non_wip_activities" in row.index and row.get("non_wip_activities", "") not in ("", "[]", None):
            act_tbl2 = build_ooo_table_from_row(row)
            if not act_tbl2.empty and "HoursRaw" in act_tbl2.columns:
                cat = (
                    act_tbl2.groupby("Activity", as_index=False)["HoursRaw"]
                            .sum()
                            .rename(columns={"HoursRaw": "Hours"})
                )
                cat = cat[cat["Activity"].astype(str).str.strip().str.upper() != "OOO"].copy()
                cat = split_nonwip_activity_minutes(cat)
                if not cat.empty:
                    cat = cat.sort_values("Hours", ascending=False)
                    order_acts = cat["Activity"].tolist()
                    act_chart = (
                        alt.Chart(cat)
                        .mark_bar()
                        .encode(
                            x=alt.X(
                                "Activity:N",
                                title="Activity",
                                sort=order_acts,
                                axis=alt.Axis(
                                    labelAngle=-45,
                                    labelLimit=200,
                                    labelOverlap=False,
                                ),
                            ),
                            y=alt.Y("Hours:Q", title="Total Non-WIP Hours"),
                            tooltip=[
                                alt.Tooltip("Activity:N", title="Activity"),
                                alt.Tooltip("Hours:Q", title="Hours", format=",.2f"),
                            ],
                        )
                        .properties(
                            height=320,
                            padding={"left": 8, "right": 12, "top": 16, "bottom": 80},
                        )
                    )
                    st.altair_chart(act_chart, width="stretch")
    st.stop()
if df.empty:
    st.warning("No data found yet. Make sure metrics_aggregate_dev.csv exists and has the 'All Metrics' sheet.")
    st.stop()
def _get_qp_teams() -> list[str]:
    try:
        qp = st.query_params
        vals = qp.get_all("teams") if hasattr(qp, "get_all") else qp.get("teams", [])
    except Exception:
        qp = st.experimental_get_query_params()
        vals = qp.get("teams", [])
    if vals is None:
        return []
    if isinstance(vals, str):
        return [vals]
    return [str(v) for v in vals]
def _set_qp_teams(values: list[str]) -> None:
    try:
        st.query_params["teams"] = values
    except Exception:
        st.experimental_set_query_params(teams=values)
def _sets_equal(a, b) -> bool:
    return set(a) == set(b)
teams = sorted([t for t in df["team"].dropna().unique()])
default_team = _first_valid_team(
    st.session_state.get("selected_team"),
    teams,
)
default_teams = [default_team] if default_team else []
if "teams_sel" not in st.session_state:
    saved = [t for t in teams if t in _get_qp_teams()]
    st.session_state.teams_sel = saved or default_teams
else:
    st.session_state.teams_sel = [
        t for t in st.session_state.teams_sel
        if t in teams
    ] or default_teams
has_dates = df["period_date"].notna().any()
min_date = pd.to_datetime(df["period_date"].min()).date() if has_dates else None
max_date = pd.to_datetime(df["period_date"].max()).date() if has_dates else None
if has_dates and min_date and max_date:
    if "start_date" not in st.session_state:
        st.session_state["start_date"] = min_date
    if "end_date" not in st.session_state:
        st.session_state["end_date"] = max_date
    start = st.session_state["start_date"]
    end = st.session_state["end_date"]
    if start > end:
        st.error("Start date cannot be after end date!")
        start, end = min_date, max_date
        st.session_state["start_date"] = start
        st.session_state["end_date"] = end
else:
    start, end = None, None
col1, col2 = st.columns([6, 6], gap="large")
with col1:
    selected_teams = st.multiselect("Teams", teams, key="teams_sel")
if selected_teams:
    st.session_state.selected_team = selected_teams[0]
current_qp = _get_qp_teams()
if not _sets_equal(st.session_state.teams_sel, current_qp):
    _set_qp_teams(sorted(st.session_state.teams_sel))
f = df.copy()
if st.session_state.teams_sel:
    f = f[f["team"].isin(st.session_state.teams_sel)]
if start and end:
    f = f[(f["period_date"] >= pd.to_datetime(start)) & (f["period_date"] <= pd.to_datetime(end))]
if f.empty:
    st.info("No rows match your filters.")
    st.stop()
ppl_hours = explode_person_hours(f)
latest = (f.sort_values(["team", "period_date"])
            .groupby("team", as_index=False)
            .tail(1)
            .copy()
)
tot_hc_used = latest["Actual HC used"].sum(skipna=True) if "Actual HC used" in latest.columns else np.nan
nw_all = load_non_wip()
if not nw_all.empty:
    if "total_non_wip_hours" in nw_all.columns:
        nw_all = nw_all[["team", "period_date", "total_non_wip_hours"]].copy()
        nw_all["total_non_wip_hours"] = pd.to_numeric(nw_all["total_non_wip_hours"], errors="coerce")
    else:
        nw_all = nw_all.assign(total_non_wip_hours=np.nan)[["team", "period_date", "total_non_wip_hours"]]
else:
    nw_all = pd.DataFrame(columns=["team", "period_date", "total_non_wip_hours"])
latest_nw = latest.merge(nw_all, on=["team", "period_date"], how="left")
tot_nonwip = latest_nw["total_non_wip_hours"].sum(skipna=True) if "total_non_wip_hours" in latest_nw.columns else 0.0
left, right = st.columns(2)
base = alt.Chart(f).transform_calculate(
    week="toDate(datum.period_date)"
).encode(
    x=alt.X("period_date:T", title="Week")
)
teams_in_view = sorted([t for t in f["team"].dropna().unique()])
multi_team = len(teams_in_view) > 1
team_sel = alt.selection_point(fields=["team"], bind="legend")
with left:
    st.subheader("Actual WIP HC used Trend")
    if "Actual HC used" in f.columns and f["Actual HC used"].notna().any():
        ahu = f[["team", "period_date", "Actual HC used"]].dropna()
        base_ahu = alt.Chart(ahu).encode(
            x=alt.X("period_date:T", title="Week"),
            y=alt.Y("Actual HC used:Q", title="Actual HC used"),
            color=alt.Color("team:N", title="Team") if len(teams_in_view) > 1 else alt.value("indianred"),
            tooltip=["team:N", "period_date:T", alt.Tooltip("Actual HC used:Q", format=",.2f")]
        )
        st.altair_chart(
            base_ahu.mark_line(point=True).properties(height=280),
            width="stretch"
        )
        if len(teams_in_view) == 1:
            team_name = teams_in_view[0]
            if 'ppl_hours' not in locals():
                ppl_hours = explode_person_hours(f)
            team_people = ppl_hours.loc[ppl_hours["team"] == team_name].copy()
            if team_people.empty:
                st.info(f"No per-person data available for {team_name}.")
            else:
                all_weeks = sorted(
                    pd.to_datetime(team_people["period_date"].dropna().unique()),
                    reverse=True
                )
                picked_week = st.selectbox(
                    f"Week:",
                    options=all_weeks,
                    index=0,
                    format_func=lambda d: pd.to_datetime(d).date().isoformat(),
                    key="ahu_week_select_anyteam",
                )
                picked_week = pd.to_datetime(picked_week).normalize()
                wk_people = team_people.loc[team_people["period_date"] == picked_week].copy()
                if wk_people.empty:
                    st.info("No per-person data for the selected week.")
                else:
                    wk_people["Actual"] = pd.to_numeric(wk_people["Actual Hours"], errors="coerce")
                    wk_people = wk_people.loc[wk_people["Actual"].fillna(0) > 0].copy()
                    if wk_people.empty:
                        st.info("Nobody to show after filtering zero-hour entries.")
                    else:
                        wk_people["Avg Daily Hours"] = (wk_people["Actual"] / 5.0)
                        wk_people["OverUnder"] = np.where(
                            wk_people["Avg Daily Hours"] >= 6, "≥ 6 (Over)", "< 6 (Under)"
                        )
                        wk_people["Delta"] = wk_people["Avg Daily Hours"] - 6
                        wk_people["DeltaLabel"] = wk_people["Delta"].map(lambda x: f"{x:+.2f}")
                        vmax = float(pd.to_numeric(wk_people["Avg Daily Hours"], errors="coerce").max())
                        pad  = max(0.3, (max(vmax, 6) * 0.12))  # a little headroom
                        y_lo = 0.0
                        y_hi = max(vmax, 6) + pad
                        y_scale = alt.Scale(domain=[y_lo, y_hi], nice=False, clamp=False)
                        order_people = (
                            wk_people.sort_values("Avg Daily Hours", ascending=False)["person"].tolist()
                        )
                        color_enc = alt.Color(
                            "OverUnder:N",
                            title="vs 6",
                            scale=alt.Scale(
                                domain=["≥ 6 (Over)", "< 6 (Under)"],
                                range=["#22c55e", "#ef4444"]  # green / red
                            )
                        )
                        bars = (
                            alt.Chart(wk_people)
                            .mark_bar()
                            .encode(
                                x=alt.X("person:N", title="Person", sort=order_people),
                                y=alt.Y("Avg Daily Hours:Q", title="Avg Daily Hours (Actual/5)", scale=y_scale),
                                color=color_enc,
                                tooltip=[
                                    "period_date:T",
                                    "person:N",
                                    alt.Tooltip("Actual:Q", title="Actual Hours (week)", format=",.2f"),
                                    alt.Tooltip("Avg Daily Hours:Q", title="Avg Daily Hours", format=",.2f"),
                                    alt.Tooltip("Delta:Q", title="Over/Under vs 6", format="+.2f"),
                                ],
                            )
                            .properties(height=280)
                        )
                        label_pad = max(0.08, (y_hi - y_lo) * 0.03)
                        labels = (
                            alt.Chart(wk_people.assign(LabelY=lambda d: d["Avg Daily Hours"] + label_pad))
                            .mark_text(dy=-4)
                            .encode(
                                x="person:N",
                                y=alt.Y("LabelY:Q", scale=y_scale),
                                text="DeltaLabel:N",
                                color=alt.Color(
                                    "OverUnder:N",
                                    legend=None,
                                    scale=alt.Scale(
                                        domain=["≥ 6 (Over)", "< 6 (Under)"],
                                        range=["#22c55e", "#ef4444"]
                                    ),
                                ),
                            )
                        )
                        ref = alt.Chart(pd.DataFrame({"y": [6]})).mark_rule(strokeDash=[4, 3]).encode(y=alt.Y("y:Q", scale=y_scale))
                        st.altair_chart(bars + labels + ref, width="stretch")
        else:
            st.caption("Select exactly one team to drill into per-person daily hours.")
    else:
        st.info("No 'Actual HC used' data available in the selected range.")
with right:
    st.subheader("Hours Trend")
    _nw = load_non_wip()
    teams_cfg = load_team_config()
    irl_lookup = {t: irl_people_for_team(t, teams_cfg) for t in teams_in_view}
    mix_rows = []
    nw_sub = _nw[_nw["team"].isin(teams_in_view)].copy()
    if not nw_sub.empty:
        for _, nw_row in nw_sub.iterrows():
            team = str(nw_row.get("team", "")).strip()
            wk = pd.to_datetime(nw_row.get("period_date"), errors="coerce")
            if not team or pd.isna(wk):
                continue
            wk = wk.normalize()
            wk_people = build_person_weekly_accounting(
                team=team,
                week=wk,
                nw_row=nw_row,
                metrics_frame=f,
                nw_frame=_nw,
                week_hours=40.0,
                irl_people=irl_lookup.get(team, set()),
            )
            if wk_people.empty:
                continue
            wk_people = wk_people.copy()
            wk_people["WIP"] = pd.to_numeric(wk_people["Completed Hours"], errors="coerce").fillna(0.0)
            wk_people["Other Team WIP"] = pd.to_numeric(wk_people["Other Team WIP"], errors="coerce").fillna(0.0)
            wk_people["Non-WIP"] = pd.to_numeric(wk_people["Accounted Non-WIP"], errors="coerce").fillna(0.0)
            wk_people["OOO"] = pd.to_numeric(wk_people["OOO Hours"], errors="coerce").fillna(0.0)
            wk_people["Unaccounted"] = pd.to_numeric(wk_people["Unaccounted"], errors="coerce").fillna(0.0)
            wk_people["Denom"] = (
                wk_people["WIP"]
                + wk_people["Other Team WIP"]
                + wk_people["Non-WIP"]
                + wk_people["OOO"]
                + wk_people["Unaccounted"]
            )
            long_df = wk_people.melt(
                id_vars=["team", "period_date", "person", "Denom"],
                value_vars=[
                    "WIP",
                    "Other Team WIP",
                    "Non-WIP",
                    "OOO",
                    "Unaccounted",
                ],
                var_name="Category",
                value_name="Hours",
            )
            long_df["Pct"] = np.where(
                long_df["Denom"] > 0,
                long_df["Hours"] / long_df["Denom"],
                np.nan,
            )
            mix_rows.append(
                long_df[["team", "period_date", "person", "Category", "Hours", "Pct"]]
            )
    person_mix = (
        pd.concat(mix_rows, ignore_index=True)
        if mix_rows
        else pd.DataFrame(columns=["team", "period_date", "person", "Category", "Hours", "Pct"])
    )
    person_mix = person_mix.dropna(subset=["period_date", "person", "Pct"]).copy()
    if person_mix.empty:
        st.info("No WIP vs Other Team WIP vs Non-WIP vs OOO vs Unaccounted data available.")
    else:
        mix_weeks = sorted(person_mix["period_date"].dropna().unique(), reverse=True)
        picked_mix_week = st.selectbox(
            "Week",
            options=mix_weeks,
            index=0,
            format_func=lambda d: pd.to_datetime(d).date().isoformat(),
            key="time_mix_week_right2",
        )
        week_mix = person_mix[
            person_mix["period_date"] == pd.to_datetime(picked_mix_week).normalize()
        ].copy()
        chosen_mix_teams = []
        if multi_team:
            teams_present = sorted(week_mix["team"].dropna().unique().tolist())
            chosen_mix_teams = st.multiselect(
                "Team(s)",
                options=teams_present,
                default=teams_present,
                key="time_mix_teams_right2",
            )
            if chosen_mix_teams:
                week_mix = week_mix[week_mix["team"].isin(chosen_mix_teams)].copy()
        if week_mix.empty:
            st.info("No person mix data for that selection.")
        else:
            category_domain = [
                "WIP",
                "Other Team WIP",
                "Non-WIP",
                "OOO",
                "Unaccounted",
            ]
            category_colors = [
                "#2563eb",  # WIP
                "#8b5cf6",  # Other Team WIP
                "#22c55e",  # Non-WIP
                "#f59e0b",  # OOO
                "#9ca3af",  # Unaccounted
            ]
            category_order_map = {
                "WIP": 0,
                "Other Team WIP": 1,
                "Non-WIP": 2,
                "OOO": 3,
                "Unaccounted": 4,
            }
            week_mix["Category"] = week_mix["Category"].astype(str).str.strip()
            week_mix = week_mix[week_mix["Category"].isin(category_domain)].copy()
            week_mix["CategoryOrder"] = week_mix["Category"].map(category_order_map)
            week_mix["Pct"] = pd.to_numeric(week_mix["Pct"], errors="coerce").fillna(0.0)
            week_mix["Hours"] = pd.to_numeric(week_mix["Hours"], errors="coerce").fillna(0.0)
            top_controls_left, top_controls_right = st.columns([1, 1])
            with top_controls_left:
                factor_out_ooo_top = st.toggle(
                    "Factor out OOO (top chart)",
                    value=True,
                    key="time_mix_factor_out_ooo_top_right2",
                )
            top_mix = week_mix.copy()
            if factor_out_ooo_top and not top_mix.empty:
                weekly_person_totals = (
                    top_mix.groupby(["period_date", "person"], as_index=False)["Hours"]
                    .sum()
                    .rename(columns={"Hours": "TotalHours"})
                )
                weekly_ooo = (
                    top_mix[top_mix["Category"] == "OOO"]
                    .groupby(["period_date", "person"], as_index=False)["Hours"]
                    .sum()
                    .rename(columns={"Hours": "OOOHours"})
                )
                weekly_base = weekly_person_totals.merge(
                    weekly_ooo,
                    on=["period_date", "person"],
                    how="left",
                )
                weekly_base["OOOHours"] = weekly_base["OOOHours"].fillna(0.0)
                weekly_base["AdjDenom"] = (
                    weekly_base["TotalHours"] - weekly_base["OOOHours"]
                ).clip(lower=0.0)
                top_mix = top_mix[top_mix["Category"] != "OOO"].copy()
                top_mix = top_mix.merge(
                    weekly_base[["period_date", "person", "AdjDenom"]],
                    on=["period_date", "person"],
                    how="left",
                )
                top_mix["Pct"] = np.where(
                    top_mix["AdjDenom"] > 0,
                    top_mix["Hours"] / top_mix["AdjDenom"],
                    np.nan,
                )
                top_mix = top_mix.dropna(subset=["Pct"]).copy()
            top_categories = [c for c in category_domain if (not factor_out_ooo_top or c != "OOO")]
            top_colors = [category_colors[category_domain.index(c)] for c in top_categories]
            person_order = sorted(top_mix["person"].dropna().unique().tolist())
            label_src = top_mix.sort_values(["person", "CategoryOrder"]).copy()
            label_src["cum_pct"] = label_src.groupby("person")["Pct"].cumsum()
            label_src["y_mid"] = label_src["cum_pct"] - (label_src["Pct"] / 2.0)
            label_src = label_src[label_src["Pct"] >= 0.05].copy()
            bars = alt.Chart(top_mix).mark_bar().encode(
                x=alt.X(
                    "person:N",
                    title="Person",
                    sort=person_order,
                    axis=alt.Axis(
                        labelAngle=-90,
                        labelLimit=200,
                        labelOverlap=False,   # force all names to render
                    ),
                ),
                y=alt.Y(
                    "Pct:Q",
                    title="% of Time" if not factor_out_ooo_top else "% of Non-OOO Time",
                    stack="normalize",
                    axis=alt.Axis(format=".0%"),
                    scale=alt.Scale(domain=[0, 1]),
                ),
                color=alt.Color(
                    "Category:N",
                    title="Legend",
                    scale=alt.Scale(
                        domain=top_categories,
                        range=top_colors,
                    ),
                    sort=top_categories,
                    legend=alt.Legend(
                        orient="top",
                        direction="horizontal",
                        title=None,
                        labelLimit=200,
                    ),
                ),
                order=alt.Order("CategoryOrder:Q", sort="ascending"),
                tooltip=[
                    alt.Tooltip("team:N", title="Team"),
                    alt.Tooltip("person:N", title="Person"),
                    alt.Tooltip("Category:N", title="Category"),
                    alt.Tooltip("Hours:Q", title="Hours", format=",.2f"),
                    alt.Tooltip("Pct:Q", title="% of Time", format=".1%"),
                    alt.Tooltip("period_date:T", title="Week"),
                ],
            )
            labels = alt.Chart(label_src).mark_text(
                color="white",
                fontSize=11,
                fontWeight="bold",
                align="center",
                baseline="middle",
            ).encode(
                x=alt.X("person:N", sort=person_order),
                y=alt.Y(
                    "y_mid:Q",
                    scale=alt.Scale(domain=[0, 1]),
                    axis=None,
                ),
                detail="Category:N",
                text=alt.Text("Pct:Q", format=".0%"),
            )
            person_totals = (
                week_mix.groupby("person", as_index=False)["Hours"]
                .sum()
                .rename(columns={"Hours": "TotalHours"})
            )
            if "wk_people_kpi" in dir() and not wk_people_kpi.empty and "person" in wk_people_kpi.columns and "Expected Hours" in wk_people_kpi.columns:
                expected_hrs = wk_people_kpi[["person", "Expected Hours"]].copy()
                expected_hrs["person"] = expected_hrs["person"].astype(str).str.strip()
            else:
                expected_hrs = pd.DataFrame({"person": person_totals["person"], "Expected Hours": 40.0})
            person_totals["person"] = person_totals["person"].astype(str).str.strip()
            person_totals = person_totals.merge(expected_hrs, on="person", how="left")
            person_totals["Expected Hours"] = person_totals["Expected Hours"].fillna(40.0)
            overflow_df = person_totals[person_totals["TotalHours"] > person_totals["Expected Hours"]].copy()
            overflow_df["y_pos"] = 1.02  # just above the top of the 100% bar
            overflow_df["label"] = "⚠"
            overflow_layer = (
                alt.Chart(overflow_df)
                .mark_text(
                    fontSize=14,
                    fontWeight="bold",
                    color="#ef4444",
                    baseline="bottom",
                )
                .encode(
                    x=alt.X("person:N", sort=person_order),
                    y=alt.Y("y_pos:Q", scale=alt.Scale(domain=[0, 1.12]), axis=None),
                    text=alt.Text("label:N"),
                    tooltip=[
                        alt.Tooltip("person:N", title="Person"),
                        alt.Tooltip("TotalHours:Q", title="Total Hours", format=",.1f"),
                        alt.Tooltip("Expected Hours:Q", title="Expected Hours", format=",.1f"),
                    ],
                )
            )
            top_chart = (bars + labels + overflow_layer).properties(height=400)
            st.altair_chart(top_chart, width="stretch")
            st.markdown("##### Drill-down over time")
            people_for_drill = sorted(top_mix["person"].dropna().unique().tolist())
            picked_person_mix = st.selectbox(
                "Person",
                options=people_for_drill,
                key="time_mix_person_right2",
            )
            drill_controls_left, drill_controls_right = st.columns(2)
            with drill_controls_left:
                drill_window = st.segmented_control(
                    "Weeks",
                    options=[8, 12, 16],
                    default=16,
                    key="time_mix_window_right2",
                )
            with drill_controls_right:
                factor_out_ooo = st.toggle(
                    "Factor out OOO",
                    value=False,
                    key="time_mix_factor_out_ooo_right2",
                )
            drill_df = person_mix[person_mix["person"] == picked_person_mix].copy()
            if multi_team and chosen_mix_teams:
                drill_df = drill_df[drill_df["team"].isin(chosen_mix_teams)].copy()
            if drill_df.empty:
                st.info("No over-time data for that person.")
            else:
                drill_df["Category"] = drill_df["Category"].astype(str).str.strip()
                drill_df = drill_df[drill_df["Category"].isin(category_domain)].copy()
                drill_df["CategoryOrder"] = drill_df["Category"].map(category_order_map)
                drill_df["Pct"] = pd.to_numeric(drill_df["Pct"], errors="coerce").fillna(0.0)
                drill_df["Hours"] = pd.to_numeric(drill_df["Hours"], errors="coerce").fillna(0.0)
                drill_df["period_date"] = pd.to_datetime(drill_df["period_date"], errors="coerce")
                latest_weeks = (
                    pd.Series(drill_df["period_date"].dropna().sort_values().unique()).tolist()
                )[-int(drill_window):]
                drill_df = drill_df[drill_df["period_date"].isin(latest_weeks)].copy()
                if factor_out_ooo and not drill_df.empty:
                    base_df = drill_df.copy()
                    weekly_person_totals = (
                        base_df.groupby(["period_date", "person"], as_index=False)["Hours"]
                        .sum()
                        .rename(columns={"Hours": "TotalHours"})
                    )
                    weekly_ooo = (
                        base_df[base_df["Category"] == "OOO"]
                        .groupby(["period_date", "person"], as_index=False)["Hours"]
                        .sum()
                        .rename(columns={"Hours": "OOOHours"})
                    )
                    weekly_base = weekly_person_totals.merge(
                        weekly_ooo,
                        on=["period_date", "person"],
                        how="left",
                    )
                    weekly_base["OOOHours"] = weekly_base["OOOHours"].fillna(0.0)
                    weekly_base["AdjDenom"] = (
                        weekly_base["TotalHours"] - weekly_base["OOOHours"]
                    ).clip(lower=0.0)
                    drill_df = drill_df[drill_df["Category"] != "OOO"].copy()
                    drill_df = drill_df.merge(
                        weekly_base[["period_date", "person", "AdjDenom"]],
                        on=["period_date", "person"],
                        how="left",
                    )
                    drill_df["Pct"] = np.where(
                        drill_df["AdjDenom"] > 0,
                        drill_df["Hours"] / drill_df["AdjDenom"],
                        np.nan,
                    )
                    drill_df = drill_df.dropna(subset=["Pct"]).copy()
                if drill_df.empty:
                    st.info("No over-time data for that person after applying filters.")
                else:
                    drill_categories = [c for c in category_domain if (not factor_out_ooo or c != "OOO")]
                    drill_colors = [
                        category_colors[category_domain.index(c)]
                        for c in drill_categories
                    ]
                    drill_df = drill_df.sort_values(["period_date", "CategoryOrder"]).copy()
                    drill_label_src = drill_df.copy()
                    drill_label_src["cum_pct"] = drill_label_src.groupby("period_date")["Pct"].cumsum()
                    drill_label_src["y_mid"] = drill_label_src["cum_pct"] - (drill_label_src["Pct"] / 2.0)
                    drill_label_src = drill_label_src[drill_label_src["Pct"] >= 0.05].copy()
                    week_count = max(len(latest_weeks), 1)
                    drill_width = max(380, min(900, week_count * 52))
                    drill_bars = (
                        alt.Chart(drill_df)
                        .mark_bar(size=28)
                        .encode(
                            x=alt.X(
                                "period_date:T",
                                title="Week",
                                axis=alt.Axis(format="%m/%d", labelAngle=0),
                            ),
                            y=alt.Y(
                                "Pct:Q",
                                title="% of Time" if not factor_out_ooo else "% of Non-OOO Time",
                                stack="normalize",
                                axis=alt.Axis(format=".0%"),
                                scale=alt.Scale(domain=[0, 1]),
                            ),
                            color=alt.Color(
                                "Category:N",
                                title="Legend",
                                scale=alt.Scale(
                                    domain=drill_categories,
                                    range=drill_colors,
                                ),
                                sort=drill_categories,
                                legend=alt.Legend(
                                    orient="top",
                                    direction="horizontal",
                                    title=None,
                                    labelLimit=200,
                                ),
                            ),
                            order=alt.Order("CategoryOrder:Q", sort="ascending"),
                            tooltip=[
                                alt.Tooltip("team:N", title="Team"),
                                alt.Tooltip("period_date:T", title="Week"),
                                alt.Tooltip("Category:N", title="Category"),
                                alt.Tooltip("Hours:Q", title="Hours", format=",.2f"),
                                alt.Tooltip("Pct:Q", title="% of Time", format=".1%"),
                            ],
                        )
                    )
                    drill_labels = alt.Chart(drill_label_src).mark_text(
                        color="white",
                        fontSize=10,
                        fontWeight="bold",
                        align="center",
                        baseline="middle",
                    ).encode(
                        x=alt.X("period_date:T"),
                        y=alt.Y(
                            "y_mid:Q",
                            scale=alt.Scale(domain=[0, 1]),
                            axis=None,
                        ),
                        detail="Category:N",
                        text=alt.Text("Pct:Q", format=".0%"),
                    )
                    drill_totals = (
                        drill_df.groupby("period_date", as_index=False)["Hours"]
                        .sum()
                        .rename(columns={"Hours": "TotalHours"})
                    )
                    person_name = normalize_person_name(str(picked_person_mix).strip())
                    if (
                        "wk_people_kpi" in dir()
                        and not wk_people_kpi.empty
                        and "person" in wk_people_kpi.columns
                        and "Expected Hours" in wk_people_kpi.columns
                    ):
                        person_expected_match = wk_people_kpi.loc[
                            wk_people_kpi["person"].astype(str).map(normalize_person_name) == person_name,
                            "Expected Hours",
                        ]
                        if not person_expected_match.empty:
                            drill_expected_hrs = float(person_expected_match.iloc[0])
                        else:
                            drill_expected_hrs = weekly_hours_for_person(person_name, 40.0)
                    else:
                        drill_expected_hrs = weekly_hours_for_person(person_name, 40.0)
                    drill_totals["ExpectedHours"] = drill_expected_hrs
                    drill_overflow_df = drill_totals[drill_totals["TotalHours"] > drill_totals["ExpectedHours"]].copy()
                    drill_overflow_df["y_pos"] = 1.02
                    drill_overflow_df["label"] = "⚠"
                    drill_overflow_layer = (
                        alt.Chart(drill_overflow_df)
                        .mark_text(
                            fontSize=14,
                            fontWeight="bold",
                            color="#ef4444",
                            baseline="bottom",
                        )
                        .encode(
                            x=alt.X("period_date:T"),
                            y=alt.Y("y_pos:Q", scale=alt.Scale(domain=[0, 1.12]), axis=None),
                            text=alt.Text("label:N"),
                            tooltip=[
                                alt.Tooltip("period_date:T", title="Week"),
                                alt.Tooltip("TotalHours:Q", title="Total Hours", format=",.1f"),
                                alt.Tooltip("ExpectedHours:Q", title="Expected Hours", format=",.1f"),
                            ],
                        )
                    )
                    drill = (drill_bars + drill_labels + drill_overflow_layer).properties(
                        height=280,
                        width=drill_width,
                    )
                    st.altair_chart(drill, width="stretch")