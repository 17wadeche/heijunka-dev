# pages/Enterprise.py
from __future__ import annotations
import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import pandas as pd
import streamlit as st
def _candidate_repo_roots(start: Path) -> List[Path]:
    roots: List[Path] = []
    p = start.resolve()
    for parent in [p, *p.parents]:
        roots.append(parent)
    roots.extend(
        [
            Path("/mount/src/heijunka-dev"),
            Path("/mount/src/HEIJUNKA-DEV"),
            Path("/mount/src/heijunka-dev/..").resolve(),
            Path.cwd(),
        ]
    )
    seen = set()
    out: List[Path] = []
    for r in roots:
        try:
            rr = r.resolve()
        except Exception:
            rr = r
        if rr not in seen:
            seen.add(rr)
            out.append(rr)
    return out
def find_org_config_path() -> Tuple[Optional[Path], List[Path]]:
    attempted: List[Path] = []
    start = Path(__file__).resolve()
    preferred_root = start.parents[1] if len(start.parents) >= 2 else start.parent
    preferred = preferred_root / "config" / "enterprise_org.json"
    attempted.append(preferred)
    if preferred.exists():
        return preferred, attempted
    for root in _candidate_repo_roots(start):
        cand = root / "config" / "enterprise_org.json"
        attempted.append(cand)
        if cand.exists():
            return cand, attempted
    for root in _candidate_repo_roots(start):
        config_dir = root / "config"
        if config_dir.is_dir():
            try:
                for f in config_dir.iterdir():
                    attempted.append(f)
                    if f.is_file() and f.name.lower() == "enterprise_org.json":
                        return f, attempted
            except Exception:
                pass
    return None, attempted
@dataclass(frozen=True)
class TeamConfig:
    name: str
    enabled: bool = True
    meta: Dict[str, Any] = None  # extra fields like portfolio/ou/etc
@dataclass(frozen=True)
class OrgConfig:
    org_name: str
    teams: List[TeamConfig]
    raw: Dict[str, Any]
def _coerce_bool(v: Any, default: bool = True) -> bool:
    if v is None:
        return default
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return bool(v)
    if isinstance(v, str):
        return v.strip().lower() in {"1", "true", "yes", "y", "enabled", "on"}
    return default
def parse_org_config(data: Dict[str, Any]) -> OrgConfig:
    org_name = (
        data.get("org_name")
        or (data.get("org") or {}).get("name")
        or data.get("name")
        or "Enterprise"
    )
    teams_raw = data.get("teams") or data.get("Teams") or []
    teams: List[TeamConfig] = []
    if isinstance(teams_raw, list):
        for t in teams_raw:
            if isinstance(t, str):
                teams.append(TeamConfig(name=t, enabled=True, meta={}))
            elif isinstance(t, dict):
                name = t.get("name") or t.get("team") or t.get("Team")
                if not name:
                    continue
                enabled = _coerce_bool(t.get("enabled"), True)
                meta = {
                    k: v
                    for k, v in t.items()
                    if k not in {"name", "team", "Team", "enabled"}
                }
                teams.append(TeamConfig(name=str(name), enabled=enabled, meta=meta))
    elif isinstance(teams_raw, dict):
        for name, tmeta in teams_raw.items():
            if isinstance(tmeta, dict):
                enabled = _coerce_bool(tmeta.get("enabled"), True)
                meta = {k: v for k, v in tmeta.items() if k != "enabled"}
            else:
                enabled, meta = True, {}
            teams.append(TeamConfig(name=str(name), enabled=enabled, meta=meta))
    return OrgConfig(org_name=str(org_name), teams=teams, raw=data)
def load_org_config() -> Tuple[Optional[OrgConfig], Optional[str], List[str], Optional[str]]:
    cfg_path, attempted = find_org_config_path()
    attempted_str = [str(p) for p in attempted]
    if cfg_path is None:
        return None, None, attempted_str, None
    try:
        text = cfg_path.read_text(encoding="utf-8")
        data = json.loads(text)
        org = parse_org_config(data)
        if not org.teams:
            return (
                None,
                f"Found config at:\n{cfg_path}\n\n…but it has no teams. "
                "Add a `teams` list (strings or objects with `name`).",
                attempted_str,
                str(cfg_path),
            )
        return org, None, attempted_str, str(cfg_path)
    except Exception as e:
        return None, f"Failed to read/parse config:\n{cfg_path}\n\n{e}", attempted_str, str(cfg_path)
def _repo_root_from_cfg_path_str(cfg_path_str: Optional[str]) -> Path:
    if cfg_path_str:
        p = Path(cfg_path_str)
        if p.exists():
            if p.parent.name.lower() == "config":
                return p.parent.parent
            return p.parent
    return Path(__file__).resolve().parents[1]
def _try_read_csv(path: Path) -> Optional[pd.DataFrame]:
    try:
        if path.exists() and path.is_file():
            return pd.read_csv(path)
    except Exception:
        return None
    return None
@st.cache_data(show_spinner=False)
def load_common_data(repo_root_str: str) -> Dict[str, pd.DataFrame]:
    repo_root = Path(repo_root_str)
    candidates = {
        "metrics": repo_root / "metrics.csv",
        "metrics_aggregate_dev": repo_root / "metrics_aggregate_dev.csv",
        "non_wip": repo_root / "non_wip.csv",
        "non_wip_activities": repo_root / "non_wip_activities.csv",
        "closures": repo_root / "closures.csv",
        "timeliness": repo_root / "timeliness.csv",
        "Timeliness": repo_root / "Timeliness.csv",
        "NS_WIP": repo_root / "NS_WIP.csv",
        "ns_non_wip_activities": repo_root / "ns_non_wip_activities.csv",
        "CRM_WIP": repo_root / "CRM_WIP.csv",
        "crm_non_wip_activities": repo_root / "crm_non_wip_activities.csv",
    }
    out: Dict[str, pd.DataFrame] = {}
    for key, p in candidates.items():
        df = _try_read_csv(p)
        if df is not None and not df.empty:
            out[key] = df
    return out
def _maybe_apply_styles():
    try:
        from utils.styles import apply_global_styles  # type: ignore
        apply_global_styles()
    except Exception:
        pass
def _norm(s: str) -> str:
    return str(s).strip().lower().replace(" ", "_")
def _first_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    cols = {_norm(c): c for c in df.columns}
    for cand in candidates:
        if cand in cols:
            return cols[cand]
    return None
def _get_team_col(df: pd.DataFrame) -> Optional[str]:
    return _first_col(df, ["team", "team_name", "org_team", "squad"])
def _get_date_col(df: pd.DataFrame) -> Optional[str]:
    return _first_col(df, ["week", "period_date", "date", "day", "as_of", "timestamp"])
def _safe_to_datetime(df: pd.DataFrame, col: str) -> pd.Series:
    return pd.to_datetime(df[col], errors="coerce")
def _weekly_start(s: pd.Series) -> pd.Series:
    return s.dt.to_period("W-MON").dt.start_time
def _to_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce")
def _loads_json_maybe(v: Any) -> Any:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (dict, list)):
        return v
    if isinstance(v, str):
        t = v.strip()
        if not t:
            return None
        try:
            return json.loads(t)
        except Exception:
            return None
    return None
def _workdays_per_week_assumption() -> int:
    return 5
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
    explicit_map = {
        "email etc.":"Email & IM",
        "emails etc.":"Email & IM",
        "emails misc": "Email& IM",
        "capa": "CAPA",
        "em/etc": "Em Etc",
        "capa meeting": "CAPA",
        "scrum/checkin": "Scrum",
        "capa call": "CAPA",
        "capa working session": "CAPA",
        "capa update call": "CAPA",
        "pmpa weekly meeting": "Pmpa Meeting",
        "finish scheduling": "Scheduling",
        "rpa minutes and action": "RPA Meeting/Action",
        "aortil call": "Meeting",
        "audit checkin" : "Audit",
        "heijunka review/update" : "Heijunka",
        "scrumb": "Scrum",
        "brian meeting, scrum, collaboration": "Meeting",
        "training/letter shadowing, locating trainee work, email": "Training/Mentoring",
        "shadowing": "Training/Mentoring",
        "response to emails for product analysis, studies and literature processing question": "Email & IM",
        "reviewing letters, meeting" : "Meeting",
        "responding to emails and team collaboration": "Email & IM",
        "mtg" : "Meeting",
        "ng, coding, email/admin":"Next Gen",
        "ng, risk management, meetings, lates/event review, collab, gb, email/admin": "Meeting",
        "lates/event review, collab, ng, gb, email/admin": "Email & IM",
        "meetings, collaboration, coding, gb, emails, collab, event review":"Meeting",
        "scheduling/heijunka update": "Scheduling",
        "ng, collaboration, coding, meetings, event reviews, emails/admin":"Next Gen",
        "ng, collaboration, coding, meetings, lates/event reviews, emails/admin": "Meeting",
        "ng, gb, coding, meetings, event reviews, emails/admin": "Email & IM",
        "ng, gb, collaboration, meetings, event reviews, emails/admin": "Collaboration",
        "collab":"Collaboration",
        "team collab":"Collaboration",
        "20 minutes mir call":"Meeting",
        "ng, collaboration, event reviews, emails/admin":"Green Belt",
        "gb, collaboration, coding, meetings, event reviews, emails/admin": "Event Reviews",
        "late" : "Late Review",
        "brian meeting": "Meeting",
        "comm, task , rr practice": "Training/Mentoring",
        "training review/questions": "Training/Mentoring",
        "e-mail": "Email & IM",
        "scrum/checking": "Scrum",
        "heijunka population": "Heijunka",
        "clinical task training" : "Clinical Task",
        "training meeting": "Training/Mentoring",
        "problem solving meeting":"Problem Solving",
        "training (shadowing, scheduling, meeting, etc.)": "Training/Mentoring",
        "scrim": "Scrum",
        "review feedback": "Feedback",
        "rpa lab meeting": "RPA Meeting/Action",
        "rpa meeing": "RPA Meeting/Action",
        "film review meeting": "Meeting",
        "other querie": "Question",
        "meet": "Meeting",
        "weekly meeting": "Meeting",
        "email catch up": "Email & IM",
        "im": "Email & IM",
        "internal audit": "Audit",
        "ri response":"RI",
        "pmq cr pre-meeting q's review": "PMQ pre-meeting review Q's",
        "other queurie": "Question",
        "cqxm querie": "Question",
        "jumped to another meeting) global quality meeting": "Global Quality Meeting",
        "it support/restart": "It Support",
        "capa remediation review": "CAPA",
        "ect letter training": "Training/Mentoring",
        "ri": "RI",
        "ri aortic meeting": "RI",
        "scrum&action": "Scrum & Action",
        "rpa lab": "RPA Meeting/Action",
        "scrum & action": "Scrum & Action",
        "scrum and action": "Scrum & Action",
        "scrum& action": "Scrum & Action",
        "aged file review": "Aged WIP Review",
        "scrum &action": "Scrum & Action",
        "call": "Meeting",
        "metrics/schedule/scrum, code gov, malf matrix update": "Scrum & Action",
        "investigation meeting": "Meeting",
        "training":"Training/Mentoring",
        "meeting": "Meeting",
        "meeeting": "Meeting",
        ". meeting": "Meeting",  # extra safety; usually cleaned earlier
        "qa review": "QA Review",
        "qa review/correction": "QA Review",
        "qa review/update": "QA Review",
        "risk management knowledge sharing call": "Risk Management Knowledge Sharing Call",
        "risk mangement knowledge sharing call": "Risk Management Knowledge Sharing Call",
        "risk mgmt kniwledge session": "Risk Management Knowledge Sharing Call",
        "film meeting": "Meeting",
        "coding/risk mgmt meeting": "Meeting",
        "coding meeting": "Meeting",
        "film meeting": "Meeting",
        "literature meeting": "Meeting",
        "literature/readcube meeting": "Meeting",
        "aging wip":"Aged WIP Review",
        "aging file review":"Aged WIP Review",
        "report request/call": "Meeting",
        "town hall": "Meeting",
        "aem":"Meeting",
        "aem cqs meting":"Meeting",
        "scrum &call": "Scrum & Action",
        "it ticket/call": "IT Issue/Update",
        "loa catchup, it issues, meetings, email/admin, ng": "IT Issue/Update",
        "email & iml" : "Email & IM",
        "crdn call":"Meeting",
        "audit meeting": "Audit",
        "scrum": "Scrum & Action",
        "biohazrd kit approval": "Biohazard Kits Approval",
        "cas report": "Team Report",
        "knowledge sharing meeting/prep": "Meeting",
        "aged file/gemba review": "Gemba",
        "it issue": "IT Issue/Update",
        "pmq pre-meeting review q's":"PMQ Meeting",
        "pmq cr meeting":"PMQ Meeting",
        "pmq cr post-meeting action":"PMQ Meeting",
        "pmq cr fu email":"Email & IM",
        "pmq meeting":"PMQ Meeting",
        "infolding response meeting": "Meeting",
        "ti meeting": "Meeting",
        "tl meeting": "Meeting",
        "email/admin": "Email & IM",
        "emails admin": "Email & IM",
        "on":"Meeting",
        "emails/admin": "Email & IM",
        "enails/admin": "Email & IM",
        "ms meeting": "Meeting",
        "aem + townhall":"Meeting",
        
        "team meeting":"Meeting",
        "townhall":"Meeting",
        "meeting prep":"Meeting",
        "training/meeting": "Training/Mentoring",
        "training/mentoring": "Training/Mentoring",
        "reg inquirie": "RI",
        "training with natalie": "Training/Mentoring",
        "training w/ natalie": "Training/Mentoring",
        "independent training": "Training/Mentoring",
        "idenpendent training": "Training/Mentoring",
        "cross functional meeting": "Meeting",
        "weekly ttvr/tmvr meeting": "Meeting",
        "staff meeting": "Meeting",
        "aems + townhall": "Meeting",
        "aem +town hall": "Meeting",
        "aems +town hall": "Meeting",
        "email; training; computer repair":"Email & IM",
        "team meetin": "Meeting",
        "proformas x":"proforma",
        "aems +townhall": "Meeting",
        "ccrum/meeting": "Meeting",
        "crossfunctional meeting": "Meeting",
        "cross functional meeting": "Meeting",
        "cross funtional meeting":"Meeting",
        "scrumber":"Scrum & Action",
        "aem+townhall": "Meeting",
        "aems+townhall":"Meeting",
        "training qa": "Training/Mentoring",
        "ooo/appt": "OOO",
        "pmpa requests/update": "PMPA",
        "pmpa questions/request": "PMPA",
        "pmp questions/update": "PMPA",
        "pmpa question": "PMPA",
        "scrume" : "Scrum & Action",
        "gfe fax": "GFE",
        "gfe e-mail":"GFE",
        "questions/update": "Questions",
        "scurm": "Scrum & Action",
        "independednt cos": "COS",
        "meeting other": "Meeting",
        "pmq query":"PMQ Meeting",
        "reliant training":"Training/Mentoring",
        "lab meetig":"Meeting",
        "sh&a meeting":"Meeting",
        "tier 2 pvh":"Meeting",
        "tier 2 meeting crdn":"Meeting",
        "sha Meeting":"Meeting",
        "rpa lab monthly meeting": "RPA Meeting/Action",
        "nellcor meeting with julio - decision?":"Meeting",
        "monthly lab meeting":"Meeting",
        "tier 2 meeting;":"Meeting",
        "tier 2 meeting":"Meeting",
        "grad project meeting":"Meeting",
        "meeting majella aurelie":"Meeting",
        "quality aem":"Meeting",
        "emails ft": "Email & IM",
        "email To cell, setup": "Email & IM",
        "cornerstone training":"Training/Mentoring",
        "pmpa/r&d meeting":"PMPA",
        "scrum 30": "Scrum & Action",
        "scrum/meeting": "Scrum & Action",
        "training (on-boarding, product training, etc)":"Training/Mentoring",
        "blockers (it issues, gch issues, computer updates)": "IT Issue/Update",
        "scum": "Scrum & Action",
        "it update": "IT Issue/Update",
        "documentation reading": "Documentation",
        "sha meeting":"Meeting",
        "email to cell": "Email & IM",
        "email to cell, setup": "Email & IM",
        "gch issue": "IT Issue/Update",
        "escalated call":"Meeting",
        "set-up":"Set Up",
        "setup":"Set Up",
        "clinical safety plan review":"Clinical Safety Plan",
        "tl call on complex file":"Meeting",
        "out of office":"OOO",
        "2 paid 15 min":"2 paid 15 minute break",
        "classroom training (d2d, emea, etc)":"Training/Mentoring",
        "trainer/ mentor (development and delivery)":"Training/Mentoring",
        "collaboration (i.e. file discussion)":"Collaboration",
        "flex team support": "Other Team WIP",
        "qs":"Question",
        "extra ftq meeting":"Meeting",
        "practice letters":"Training/Mentoring",
        "letter training":"Training/Mentoring",
        "training, collaboration":"Training/Mentoring",
        "training related activities":"Training/Mentoring",
        "training related activities (meet, review, questions, updates, finish inbox)":"Training/Mentoring",
        "scrum/metrics/schedule": "Metrics & Schedule",
        "schedule": "Metrics & Schedule",
        "metric": "Metrics & Schedule",
        "wip schedule, setting up metrics table": "Metrics & Schedule",
        "metrics/schedule/scrum": "Metrics & Schedule",
        "rpa emails, ftq meeting": "RPA Meeting/Action",
        "practice letters, collaboration": "Collaboration",
        "rpa email": "RPA Meeting/Action",
        "rpa call": "RPA Meeting/Action",
        "collaboration/question": "Collaboration",
        "team collaboration": "Collaboration",
        "extra ftq meeting, rpa request":"Meeting",
        "1 hour training meeting":"Training/Mentoring",
        "affera training, rpa email request":"Training/Mentoring",
        "affera training, rpa email":"Training/Mentoring",
        "90 minutes training call":"Training/Mentoring",
        "affera training for sean & golden":"Training/Mentoring",
        "training related activity":"Training/Mentoring",
        "training questions. comm/task/rr practice review":"Training/Mentoring",
        "affera training":"Training/Mentoring",
        "training + scrum":"Training/Mentoring",
        "questions, grading rr/comm/task practice":"Training/Mentoring",
        "1 hour training call":"Training/Mentoring",
        "1 hour training":"Training/Mentoring",
        "training prep":"Training/Mentoring",
        "35 mins reverse-shadowing":"Training/Mentoring",
        "cornerstone, scrum":"Cornerstone",
        "cornerstone, email":"Cornerstone",
        "training and staff meeting":"Training/Mentoring",
        "60 minute training meeting":"Training/Mentoring",
        "training and meeting":"Training/Mentoring",
        "1to": "Meeting",
        "affera training, rpa collaboration":"Training/Mentoring",
        "trainings, practice rr":"Training/Mentoring",
        "audit report":"Audit",
        "aging review": "Aged WIP Review",
        "open file review": "WIP Review",
        "querie":"Question",
        "started compiling details to prepare email for clinical team" :"Email & IM",
        "training and question":"Training/Mentoring",
        "training reviews, explanation":"Training/Mentoring",
        "training, keytext for letters, collaboration":"Training/Mentoring",
        "training meeting and staff meeting":"Training/Mentoring",
        "affera training prep":"Training/Mentoring",
        "training/question":"Training/Mentoring",
        "training and collaboration":"Training/Mentoring",
        "1 hr 15 mins training":"Training/Mentoring",
        "scrum, email": "Email & IM",
        "oem integer meeting":"Meeting",
        "affera meeting":"Meeting",
        "affera ftq meeting":"Meeting",
        "affera mpxr meeting":"Meeting",
        "admin/email": "Email & IM",
        "emails, collaboration with lab": "Email & IM",
        "practice":"Training/Mentoring",
        "comm, task practce":"Training/Mentoring",
        "prism 2 training":"Training/Mentoring",
        "extra ftq meeting":"Meeting",
        "ris": "RI",
        "meetings; email": "Meeting",
        "training independent": "Training/Mentoring",
        "ad hoc ri data request from pmpa": "PMPA",
        "emails/adming":"Email & IM",
        "meetings other": "Meetings",
        "tm trainer":"Training/Mentoring",
        "emails qa scrumb":"Email & IM",
        "pmpa question/update": "PMPA",
        "cos1 training":"Training/Mentoring",
        "ooo for appt": "OOO",
        "emails/question": "Question",
        "questions/discussion": "Question",
        "training product training team meeting cross functional meeting aems + townhall":"Training/Mentoring",
        "lit scrum metric":"Scrum & Action",
        "pvh aged file review":"Aged WIP Review",
        "pmq cr post-meeting q's review":"PMQ Meeting",
        "it issues over im and phone": "IT Issue/Update",
        "imdrf code call": "Meeting",
        "calls+": "Meeting",
        "cornerstone training / gap report / peb correction":"Cornerstone",
        "capa (owner & support)":"CAPA",
        "aged file":"Aged WIP Review",
        "intake meeting": "Meeting",
        "scrum & after scrum meeting":"Scrum & Action",
        "interruptions/question": "Question",
        "emails other": "Email & IM",
        "im's/email": "Email & IM",
        "it/admin": "IT Issue/Update",
        "mentoring":"Training/Mentoring",
        "mentoring/interruption":"Training/Mentoring",
        "traing":"Training/Mentoring",
        "complex/consult training":"Training/Mentoring",
        "meeting with tm":"Meeting",
        "sme improvement projects/ problem solving":"Problem Solving",
        "training, scrum":"Training/Mentoring",
        "team communication":"Meeting",
        "aem (justin for all brian, chris, geoff for ftes)": "Team Meetings Aems Team Building",
        "respond to engineer email": "Email & IM",
        "training q":"Training/Mentoring",
        "training, practice, cornerstone":"Training/Mentoring",
        "ri meeting": "RI",
        "ri work": "RI",
        "questions, death event rr/notification":"Question",
        "training letter burdown": "Letter Burndown",
        "team collaboration repsonded to": "Collaboration",
        "issues with laptop": "IT Issue/Update",
        "meeting on file":"Meeting",
        "tm meeting":"Meeting",
        "gch crashing": "IT Issue/Update",
        "software update": "IT Issue/Update",
        "gch crashe": "IT Issue/Update",
        "pc restart": "IT Issue/Update",
        "gch slow and crashing": "IT Issue/Update",
        "call with tl on file":"Meeting",
        "tl call on file":"Meeting",
        "it issues/restart": "IT Issue/Update",
        "pc restart/update": "IT Issue/Update",
        "training louise": "Training/Mentoring",
        "cornerstone scrum":"Cornerstone",
        "louise training": "Training/Mentoring",
        "audit prep": "Audit",
        "audit support fda ri review": "Audit",
        "teams meeting":"Meeting",
        "it/computer issue": "IT Issue/Update",
        "lab meeting":"Meeting",
        "grad project work":"Project Work",
        "lab shadowing": "Training/Mentoring",
        "shadowing aortic pa": "Training/Mentoring",
        "fire marshall training": "Training/Mentoring",
        "it issue/update": "IT Issue/Update",
        "workday career development":"career development",
        "rfai call": "Meeting",
        "morning admin": "Admin",
        "lab monthly meeting project work":"Project Work",
        "meetings voyager transfer":"Meeting",
        "reading previous investigation write ups/consulting documentation": "Documentation",
        "admin and catheter tracking for nellcor": "Admin",
        "lab monthly meeting" :"Meeting",
        "aortic meeting" :"Meeting",
        "it": "IT Issue/Update",
        "aortic report": "Team Report",
        "sh report": "Team Report",
        "pvh report": "Team Report",
        "pmpa":"PMPA",
        "pmq querie": "PMQ Meeting",
        "next gen meeting":"Next Gen",
        "new ri":"RI",
        "it support": "IT Issue/Update",
        "rrtt report update": "rrtt",
        "pmpa questions/updat":"PMPA",
        "pvh call": "Meeting",
        "aging file": "Aged WIP Review",
        "is tool": "IS Tool Review",
        "srcum": "Scrum & Action",
        "readcube meeting":"Meeting",
        "readcube training": "Training/Mentoring",
        "vig training call": "Training/Mentoring",
        "pmq follow-up": "PMQ Meeting",
        "global quality meeting":"Meeting",
        "meetings & action":"Meeting",
        "proforma meeting":"Meeting",
        "email/meeting action": "Email & IM",
        "dhr meeting":"Meeting",
        "sh proforma training":"Training/Mentoring",
        "bsi audit report": "Audit",
        "aging": "Aged WIP Review",
        "pmq cr meeting & action":"PMQ Meeting",
        "pmq coding query":"PMQ Meeting",
        "aortic call":"Meeting",
        "gemba&prep": "Gemba",
        "one to one": "Meeting",
        "inv summ review": "Investigation Summary",
        "scrum, cornerstone":"Scrum & Action",
        "actions from scrum/minute":"Scrum & Action",
        "scrum, emails, meeting":"Scrum & Action",
        "knowledge sharing call": "Meeting",
        "pmq call":"PMQ Meeting",
        "pmq follow up":"PMQ Meeting",
        "ri review":"RI",
        "sh call": "Meeting",
        "sh vig meeting": "Meeting",
        "eu mdr call": "Meeting",
        "other emails & im":"Email & IM",
        "pmq mfg assessment call": "PMQ Meeting",
        "sh vig review/querie": "Question",
        "aged wip review": "Aged WIP Review",
        "file review": "Aged WIP Review",
        "rpa meeting": "RPA Meeting/Action",
        "rpa meeting and action": "RPA Meeting/Action",
        "fer capa report/review": "CAPA",
        "rpa meeting/action": "RPA Meeting/Action",
        "rpa meeting & action": "RPA Meeting/Action",
        "vig training": "Training/Mentoring",
        "laptop update": "IT Issue/Update",
        "laptop setup": "IT Issue/Update",
        "ivanti connection issue": "IT Issue/Update",
        "call with it":"IT Issue/Update",
        "it call":"IT Issue/Update",
        "gemba prep":"Gemba",
        "pmq query review":"PMQ Meeting",
        "ri update/submit":"RI",
        "pmpa questions/update":"PMPA",
        "precedent event/rd conflict": "RD Conflict",
        "Rd Conflict": "RD Conflict",
        "pmpa request":"PMPA",
        "pmpa update":"PMPA",
        "pmpa/questions/update":"PMPA",
        "email; meeting": "Email & IM",
        "email; training": "Email & IM",
        "emails/amin": "Email & IM",
        "email; article reivew for svt": "Email & IM",
        "email admin": "Email & IM",
        "rpa action": "RPA Meeting/Action",
        "rpa minutes write-up":"RPA Meeting/Action",
        "complex events consult": "Complex Event Consult",
        "email": "Email & IM",
        "codes governance call":"Meeting",
        "navion call":"Meeting",
        "sh meeting/event review":"Meeting",
        "pvh audit report":"Audit",
        "email/im & call":"Email & IM",
        "sh meeting":"Meeting",
        "lsh bridge issue":"LSH Bridge",
        "audit support":"Audit",
        "aortic meeting & post meeting call":"Meeting",
        "adn career development workshop":"Career Development",
        "ng workshop, lates, wmail/admin": "Next Gen",
        "a whole lot of collaboration, meet and greets, meetings, troubleshooting, and training": "Collaboration",
        "scrum and call":"Scrum & Action",
        "ri follow-up":"RI",
        "ri completion":"RI",
        "ri call": "RI",
        "processing question":"Question",
        "meetings; pe review for pmpa":"PMPA",
        "emails/discussion":"Email & IM",
        "mtgs 2 hrs": "Meeting",
        "sales rep training": "Training/Mentoring",
        "mir training": "Training/Mentoring",
        "catching up on email from being ooo for":"Email & IM",
        "rs pxm sh inbox":"RS Inbox",
        "rs sh inbox":"RS Inbox",
        "cornerstone/it ticket": "IT Issue/Update",
        "it ticket/tech issue": "IT Issue/Update",
        "ad hoc request from pmpa":"PMPA",
        "cornerstone + it ticket": "IT Issue/Update",
        "emails/other":"Email & IM",
        "catch up on email":"Email & IM",
        "weekly tmvr clinical event investigation meeting": "Meeting",
        "on clinical. it tickets and cornerstone": "IT Issue/Update",
        "restore/mpxr issue/discussion": "IT Issue/Update",
        "weekly ttvr clinical event investigation meeting": "Meeting",
        "prep for audit":"Audit",
        "sh staff meet": "Meeting",
        "submitted nurse consult":"Nurse Consult",
        "mns cornerstone":"Cornerstone",
        "qa observation":"QA",
        "pmpa discussion/canada discussion":"PMPA",
        "discussions/im": "Email & IM",
        "skills lab + consult w/corie":"Nurse Consult",
        "computer updates/rebooting/comp problem": "IT Issue/Update",
        "weekly cqxm/pmpa meeting":"PMPA",
        "meetings; pe review for pmpa.":"PMPA",
        "ri discussion":"RI",
        "pmpa requests & qa":"PMPA",
        "pacemaker complaint meeting":"Meeting",
        "mn mentoring": "Training/Mentoring",
        "steam pop review, cas coding":"Coding",
        "audit, building a community/potluck & costume party":"Audit",
        "internal audit backroom support":"Audit",
        "internal audit backroom":"Audit",
        "in-person discussions/independent research for south africa fvr process":"Collaboration",
        "impromptu discussion":"Collaboration",
        "skills lab vlt":"Skills Lab",
        "tvtr discussion":"Collaboration",
        "ris/internal audit questions about ris":"Audit",
        "em/technology etc":"Em Etc",
        "other/em etc":"Em Etc",
        "impromptu in-person discussion":"Collaboration",
        "emails/inbox": "Email & IM",
        "emails/misc": "Email & IM",
        "leading indicators and meeting":"Meeting",
        "wip mgmt/pmpa":"PMPA",
        "call with rep":"Meeting",
        "teams chat": "Email & IM",
        "30: pmpa updates/discussion misc":"PMPA",
        "misc update": "IT Issue/Update",
        "leading indicators/meeting/action":"Meeting",
        "it issue/update": "IT Issue/Update",
        "it issues/tix": "IT Issue/Update",
        "emails/misc/technology": "IT Issue/Update",
        "teams question": "Email & IM",
        "ms teams Chats/interruption": "Email & IM",
        "emails/in-person discussion": "Email & IM",
        "cq problem solving monthly meeting":"Problem Solving",
        "leading indicators meeting":"Meeting",
        "jar lid particulate discussion":"Jar Lid Discussion",
        "emails/action": "Email & IM",
        "mir training meeting":"Training/Mentoring",
        "redline doc":"Redline",
        "extra time on inbox":"RS Inbox",
        "tm suppor":"Tm Support",
        "ia backroom support scrum":"Audit",
        "ia backroom support scrum audit meeting":"Audit",
        "weekly team meeting":"Meeting",
        "laptop issue": "IT Issue/Update",
        "agile co":"Agile Change Order",
        "slow gch": "IT Issue/Update",
        "ia backroom audit meeting":"Audit",
        "scrum: 30.":"Scrum & Action",
        "qms training":"Training/Mentoring",
        "srum,parklife":"Scrum & Action",
        "meetings: 60 (includes answering questions)":"Meeting",
        "scrum 20":"Scrum & Action",
        "ste up":"Set Up",
        "training review":"Training/Mentoring",
        "escalated call&prep":"Meeting",
        "call with monica":"Meeting",
        "audit, meeting":"Audit",
        "new desk set up":"Desk",
        "emails etc": "Email & IM",
        "misc/desk":"Desk",
        "misc/email": "Email & IM",
        "ris/south africa final rr request":"RI",
        "questions/misc":"Questions",
        "cornertsone":"Cornerstone",
        "pmpa disucssion":"PMPA",
        "crdn strategy meeting":"Meeting",
        "providing kirsty vig training":"Training/Mentoring",
        "escalations training":"Training/Mentoring",
        "train":"Training/Mentoring",
        "questions/email": "Email & IM",
        "emails/inbox/it issue": "Email & IM",
        "intrepid question":"Question",
        "email misc.": "Email & IM",
        "career development/inclusion plan":"Career Development",
        "email misc": "Email & IM",
        "training/mentor review":"Training/Mentoring",
        "clinical training":"Training/Mentoring",
        "harmony meeting":"Meeting",
        "weekly tmvr meeting":"Meeting",
        "emails/it issue": "Email & IM",
        "emails misc.": "Email & IM",
        "questions/emails/inbox": "Email & IM",
        "emails/etc": "Email & IM",
        "email wip etc.": "Email & IM",
        "emails/teams message": "Email & IM",
        "it/gch slowness": "IT Issue/Update",
        "questions/pmpa":"PMPA",
        "meeting/email": "Email & IM",
        "questions/pmpa inquirie":"PMPA",
        "leading indicators action/meeting":"Meeting",
        "unplanned nextgen":"Next Gen",
        "reg inquiries next gen support ooo metrics/schedule/wip mgmt":"RI",
        "collaboration audit support problem solving late mdrs qa clinical pmpa questions/update":"Audit",
        "emails admin etc": "Email & IM",
        "weekly tvr meeting":"Meeting",
        "collaboration sh/crm":"Collaboration",
        "pvh next gen meeting":"Next Gen",
        "gemba & prep":"Gemba",
        "proformas x2":"Proforma",
        "1s":"Meeting",
        "yellow belt cornerstone":"Yellow Belt",
        "yellow belt cornerstone, question":"Yellow Belt",
        "training qs":"Question",
        "ran batchman for fers/fac":"Batchman",
        "scheduling/wip mgmt":"Schedule",
        "inbox/email": "Email & IM",
        "coordinating with pmpa":"PMPA",
        "emails wip etc": "Email & IM",
        "stragety meeting":"Meeting",
        "escalated call excel":"Meeting",
        "meeting with rep":"Meeting",
        "eu vig training":"Training/Mentoring",
        "training kirsty":"Training/Mentoring",
        "meeting with lab":"Meeting",
        "training + ri completion + chey wade email":"Training/Mentoring",
        "sr email": "Email & IM",
        "meeting on clinical study":"Meeting",
        "restart": "IT Issue/Update",
        "escalated call prep":"Meeting",
        "safety meeting":"Meeting",
        "computer reboot": "IT Issue/Update",
        "bsi audit meeting":"Audit",
        "it/tech issue": "IT Issue/Update",
        "pmpa discussion":"PMPA",
        "wip mgmt.":"WIP Mgmt",
        "weekly revalve meeting":"Meeting",
        "it ticket": "IT Issue/Update",
        "inbox":"RS Inbox",
        "technology": "IT Issue/Update",
        "late rr discussion":"Collaboration",
        "ng workshop, emails/admin, ms review": "Next Gen",
        "unplanned meetings, training/mentoring, unplanned questions/interruptions/people showing up at my desk with questions requiring research and collaboration": "Training/Mentoring",
        "ri repsonse":"RI",
        "rrt":"Rrtt",
        "cornerstone and admin":"Admin",
        "dmaic":"Problem Solving",
        "ng workshop, email/admin/collab": "Next Gen",
        "ng workshop, lates, email/admin": "Next Gen",
        "rpa":"RPA Meeting/Action",
        "on a call":"Meeting",
        "tier 2 pvh meeting":"Meeting",
        "(fw 15) (emails 15) (looking at aged files 30) (cf files additional testing 30)": "Aged WIP Review",
        "meetings, ng/coding": "Next Gen",
        "crdn weekly call":"Meeting",
        "(meetings 60) looking at aged files to see if any of them are workable (60)":"Meeting",
        "parkmore townhall":"Meeting",
        "prep for meeting": "Admin",
        "on call":"Meeting",
        "(fw - 15) (emails -15)":"Email & IM",
        "aortic cross functional call":"Meeting",
        "shiperp go/no go call":"Meeting",
        "shiperp pro-forma vs commercial invoice call":"Meeting",
        "rdn emails for generator return":"Email & IM",
        "call on file":"Meeting",
        "collaboration, problem statement for yellow belt training":"Yellow Belt",
        "career development plan":"Career Development",
        "ng, meetings, event review": "Next Gen",
        "review affera materials with ruth for ftq/mpxr":"Meeting",
        "affera emails/prism 2 hypercare meeting":"Meeting",
        "ftq review meeting (extra time) 60":"Meeting",
        "emails/rpa requests/collaboration":"Collaboration",
        "team colaboration":"Collaboration",
        "training: questions & reviewing event": "Training/Mentoring",
        "affera ftq project":"Project Work",
        "ftq project":"Project Work",
        "metrics, scrum, training":"Scrum & Action",
        "training, collaboration, ide snack & chat": "Training/Mentoring",
        "affera rpa - emails/phone call": "RPA Meeting/Action",
        "metrics/scrum/schedule":"Scrum & Action",
        "affera rpa email": "RPA Meeting/Action",
        "closed several checked) mostly training activtie": "Training/Mentoring",
        "affera email":"Email & IM",
        "meetings, training, admin/emails, ms requests, rpa request":"Admin",
        "meetings, training/education, admin/emails, ms/rpa, collaboration":"Admin",
        "burndown meeting, metrics, schedule":"Admin",
        "gemba and prep, metrics, schedule":"Gemba",
        "computer issues, collaboration, training plan update": "IT Issue/Update",
        "meetings, metrics/schedule":"Meeting",
        "meetings, collaboration, admin/email, cornerstone":"Meeting",
        "audit support - monitor backroom pull pesrs/reg reports =":"Audit",
        "emails, collaboration, rpa request":"Email & IM",
        "remediation activities 10 events reprocessed":"Remediation",
        "ng, regional, collaboration/event review, coding, email/admin, filling out pa board":"Next Gen",
        "internal audit pre-requests, audit review, affera email":"Audit",
        "training, ps event reviews, collaboration, meetings, emails/admin, filling out pa board":"Training/Mentoring",
        "rpa /trending/hypercare emails, phone calls, questions/collaboration,":"RPA Meeting/Action",
        "meetings, event reviews, collaboration, and a lot of time on the phone with my it besties working through loss of microsoft functionalitie":"Meeting",
        "ms, collaboration/event reviews, veteran's day program, emails/filling out pa board":"Email & IM",
        "ca rd audit":"Audit",
        "supp rr review practice events, emails - rpa, trending, audit":"Audit",
        "machine learning rd project request":"Project Work",
        "meetings, trending requests, ms and additional rds on friday report completed-this is wip but no proper category for it, filling out pa board":"Meeting",
        "figuring out how to undo-benefit selection and reapply when":"Admin",
        "training/review practice event":"Training/Mentoring",
        "scrum metrics, schedule for friday":"Scrum & Action",
        "scrum metrics, schedule for monday":"Scrum & Action",
        "delivery) practice1 letter":"Training/Mentoring",
        "benefit enrollment":"Admin",
        "rpa request":"RPA Meeting/Action",
        "project":"Project Work",
        "shadow":"Training/Mentoring",
        "trending team request":"Project Work",
        "training, collaboration, emails/admin":"Training/Mentoring",
        "meetings, collaboration, troubleshooting, training, email/admin":"Meeting",
        "meetings, late, traning, regional, emails/admin, collaboration":"Meeting",
        "precedent event review, training, meeting":"Training/Mentoring",
        "many meetings for collaboration, training time (delivery and receipt)":"Meeting",
        "meetings,collaboration, event review for potential capa, schedule metric":"Meeting",
        "meetings, lates and submission, collaboration, ms request, emails/admin":"Meeting",
        "emails - affera, rpa, hypercare":"Email & IM",
        "training, collaboration, pull work, scrum":"Training/Mentoring",
        "emails - rpa/audit, collaboration":"Email & IM",
        "emails, collaboration, metrics,":"Email & IM",
        "emails/teams - affera, rpa, audit, hypercare":"Email & IM",
        "collaboration (including rpa lab,)":"Collaboration",
        "training side work":"Training/Mentoring",
        "nextgen": "Next Gen",
        "flexcathcross breach meeting":"Meeting",
        "meetings, collaboration, troublehooting, training, email/admin":"Meeting",
        "monitoring audit backroom":"Audit",
        "email/teams - affera,rpa,mappers; collaboration":"Email & IM",
        "meetings, training, emails/admin":"Admin",
        "meetings, training, collaboration, emails/admin":"Admin",
        "collaboration, meetings, ps coding remediation, email/admin":"Meeting",
        "meetings, emails, schedule":"Email & IM",
        "emails/teams - rpa, collaboration":"Email & IM",
        "meetings, collaboration, ps coding remediation":"Meeting",
        "questions from team, collaboration":"Question",
        "scrum/ metrics/schedule":"Scrum & Action",
        "christmas burndown meeting 60 minute":"Meeting",
        "emails/rpa request/teams request":"Email & IM",
        "metrics, meetings, schedule":"Admin",
        "rpa requests, cornerstone":"RPA Meeting/Action",
        "burndown meeting":"Meeting",
        "emails, collaboration, rpa":"Email & IM",
        "practice letter": "Training/Mentoring",
        "audit support - monitor backroom pull pesrs/reg reports":"Audit",
        "audit support - pe search, pull pesrs, monitor backroom":"Audit",
        "completed pe (ftq) for audit":"Audit",
        "kick-off meeting for capa 605686 remediation activitie":"CAPA",
        "audit support - pesrs reg reports - total of":"Audit",
        "audit support - pull pesrs, monitor backroom":"Audit",
        "personal development plan":"Career Development",
        "career development review":"Career Development",
        "career developemnt plan":"Career Development",
        "training - event create, review, meet, question": "Training/Mentoring",
        "metrics, tl meeting, event request, meetings, ri":"Meeting",
        "consulting/email":"Email & IM",
        "getting training plans established, meetings scheduled, etc.":"Admin",
        "training took whole day, meetings went over, pulling training events took a while since training environment is having tech issues, many technical difficultie": "Training/Mentoring",
        "metrics, schedule, ri help":"Admin",
        "training, ml rd proj meeting, collaboration":"Training/Mentoring",
        "metrics, schedule, lab call, collaboration, training event review, data pull":"Admin",
        "1.5 hours training/shadow":"Training/Mentoring",
        "metrics, burndown meeting, schedule":"Admin",
        "computer issue": "IT Issue/Update",
        "email/teams requests rpa, audit":"Email & IM",
        "emails/teams questions/collaboration":"Email & IM",
        "us sr/rd shadowing":"Training/Mentoring",
        "metrics, meetings, schedule, email":"Admin",
        "meeting with jerlie on burndown":"Meeting",
        "email/teams rpa request": "RPA Meeting/Action",
        "rpa requests/collaboration": "RPA Meeting/Action",
        "team collabration":"Collaboration",
        "pulled metrics for scrum":"Scrum & Action",
        "training, collaboration, emails/admin":"Admin",
        "ng, collaboration, ps requests, meeting":"Next Gen",
        "collaboration, meetings, lates update":"Collaboration",
        "collaboration, holiday presentation":"Collaboration",
        "holiday program activitie":"Meeting",
        "training, late mdr, metrics, cross collab qs, metrics/schedule, cornerstone":"Training/Mentoring",
        "holiday program training - no meetings but lots of troubleshooting collaboration":"Collaboration",
        "precedent event review, metrics/scrum/schedule, late tracker, meeting":"Scrum & Action",
        "schedule, metrics, collaboration":"Admin",
        "metrics, schedule, stacy qs, code gov, teammate call":"Admin",
        "emails/teams question":"Email & IM",
        "email/teams rpa & regulatory":"Email & IM",
        "reviwed affera training material for anna wright/stacy r.":"Training/Mentoring",
        "reverse shadow":"Training/Mentoring",
        "team collaboration, review event with jyn & jerlie":"Collaboration",
        "tem collaboration":"Collaboration",
        "2 meetings, rpa questions/request":"Meeting",
        "training, admin, collaboration":"Admin",
        "two 30 minute meeting":"Meeting",
        "collaboration, meeting":"Meeting",
        "rpa/reliability email/phone call": "RPA Meeting/Action",
        "pulling metrics and reviewing what was required to be done, collobartion":"Admin",
        "metrics, d2d, planning, schedule":"Admin",
        "time to enter in cap mgmt tracker":"Fill Out Capacity Management Tools (heijunka, Non-wip Chart)",
        "affera ftq work":"Project Work",
        "meetings - affera ftq & affera data file":"Meeting",
        "training, scrum, collab":"Scrum & Action",
        "metrics/schedule/scrum":"Scrum & Action",
        "schedule, metrics, scrum, ageing event investigation, email":"Scrum & Action",
        "metrics, reliability meeting, schedule, scrum":"Scrum & Action",
        "metrics/schedule/scrum, lates meeting":"Scrum & Action",
        "ng, collaboration, email/admin":"Next Gen",
        "rpa requests, email": "RPA Meeting/Action",
        "rpa emails, affera meeting/team": "RPA Meeting/Action",
        "rpa requests - email, team": "RPA Meeting/Action",
        "meeting w/ reliability":"Meeting",
        "training, meetings, ng, collab":"Training/Mentoring",
        "metrics, scrum, scheudle, ri, batchman, townhall":"Scrum & Action",
        "ri, metrics/scrum/schedule":"Scrum & Action",
        "team collaboration & email":"Collaboration",
        "goal setting/mid yr development planning / town hall":"Career Development",
        "metrics, meetings, ri":"RI",
        "metrics, gemba prep, code gov, ri":"Admin",
        "team collaboration, met stacy poe":"Collaboration",
        "ng, ms, event review, emails, colloboration, filling out pa board": "Next Gen",
        "ng, coding, emails, filling out pa board":"Email & IM",
        "training, collaboration, meetings, ms/product, filling out pa, email": "Training/Mentoring",
        "ng, coding, event reviews, emails/messages, colloboration, filling out pa board": "Event Reviews",
        "/request for affera monitoring":"Data Request",
        "on training prep": "Training/Mentoring",
        "affera question":"Question",
        "reviewed the training from liam for the new mir": "Training/Mentoring",
        "nh training": "Training/Mentoring",
        "training nhs": "Training/Mentoring",
        "email/im":"Email & IM",
        "ri temp handoff mtg":"RI",
        "career dev review":"Career Development",
        "next gen call recording": "Next Gen",
        "acrum":"Scrum & Action",
        "sh letter querie":"Question",
        "call after scrum":"Scrum & Action",
        "global town hall":"Meeting",
        "restore lsh bridge review":"LSH Bridge",
        "rs inbox":"Inbox Maintenance",
        "next gen call":"Next Gen",
        "meeting & action": "Meeting",
        "cornerston":"Cornerstone",
        "team lead meeting & action": "Meeting",
        "pmq cr pre-meeting q's review & pre-meeting call":"PMQ Meeting",        
    }
    if lower in explicit_map:
        return explicit_map[lower]
    acronym_tokens = {
        "im", "wip", "ooo", "sla", "qa", "hc", "pe", "wfh", "pto",
        "ri", "capa",
    }
    words = lower.split(" ")
    if len(words) == 1:
        w = words[0]
        if w.endswith("s") and not w.endswith("ss") and len(w) > 3:
            w = w[:-1]
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
def split_nonwip_activity_minutes(cat: pd.DataFrame) -> pd.DataFrame:
    import numpy as np
    if cat.empty:
        return cat
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
    import numpy as np
    out = pd.DataFrame(rows)
    if out.empty:
        return cat
    out["Activity"] = out["Activity"].map(_canon_activity)
    return out.groupby("Activity", as_index=False)["Hours"].sum()
st.set_page_config(page_title="Enterprise Dashboard", layout="wide")
_maybe_apply_styles()
st.title("Enterprise Dashboard")
org, org_err, attempted_paths, cfg_path_str = load_org_config()
if org is None:
    st.error("No org config found.")
    expected_example = "/mount/src/heijunka-dev/config/enterprise_org.json"
    st.caption(
        "Expected file at: "
        f"`{expected_example}`\n\n"
        "Create `config/enterprise_org.json` and add teams.\n\n"
        "Tip: On Linux, folder names are case-sensitive. If your repo folder is `HEIJUNKA-DEV`, "
        "hard-coded paths to `heijunka-dev` will fail."
    )
    if org_err:
        st.code(org_err)
    with st.expander("Paths checked (debug)", expanded=False):
        st.write("\n".join(attempted_paths[:300]))
    st.stop()
repo_root = _repo_root_from_cfg_path_str(cfg_path_str)
data = load_common_data(str(repo_root))
enabled_teams = [t for t in org.teams if t.enabled]
all_team_names = [t.name for t in org.teams]
enabled_team_names = [t.name for t in enabled_teams] or all_team_names
with st.sidebar:
    st.subheader(org.org_name)
    all_portfolios = sorted(
        {
            str((t.meta or {}).get("portfolio")).strip()
            for t in org.teams
            if (t.meta or {}).get("portfolio") is not None
        }
    )
    portfolio_filter = st.multiselect(
        "Portfolio",
        options=all_portfolios,
        default=all_portfolios,
    )
    teams_after_portfolio = (
        [
            t
            for t in org.teams
            if str((t.meta or {}).get("portfolio")).strip() in set(portfolio_filter)
        ]
        if portfolio_filter
        else []
    )
    all_ous = sorted(
        {
            str((t.meta or {}).get("ou")).strip()
            for t in teams_after_portfolio
            if (t.meta or {}).get("ou") is not None
        }
    )
    ou_filter = st.multiselect(
        "OU",
        options=all_ous,
        default=all_ous,
    )
    teams_after_ou = (
        [
            t
            for t in teams_after_portfolio
            if str((t.meta or {}).get("ou")).strip() in set(ou_filter)
        ]
        if ou_filter
        else []
    )
    team_options = [t.name for t in teams_after_ou]
    default_teams = [t for t in enabled_team_names if t in team_options]
    if not default_teams and team_options:
        default_teams = team_options
    team_filter = st.multiselect(
        "Teams",
        options=team_options,
        default=default_teams,
    )
def filter_by_team(df: pd.DataFrame) -> pd.DataFrame:
    if not team_filter:
        return df.iloc[0:0]
    team_cols = [
        c
        for c in df.columns
        if c.strip().lower() in {"team", "team_name", "squad", "org_team"}
    ]
    if not team_cols:
        tc = _get_team_col(df)
        if tc:
            team_cols = [tc]
    if not team_cols:
        return df
    col = team_cols[0]
    return df[df[col].astype(str).isin(set(team_filter))]
def _date_bounds_for_df(df: pd.DataFrame) -> tuple[Optional[pd.Timestamp], Optional[pd.Timestamp], Optional[str]]:
    dc = _get_date_col(df)
    if not dc:
        return None, None, None
    ser = pd.to_datetime(df[dc], errors="coerce").dropna()
    if ser.empty:
        return None, None, dc
    return ser.min(), ser.max(), dc
def section_date_range(
    label: str,
    df: Optional[pd.DataFrame],
    key: str,
) -> tuple[Optional[pd.Timestamp], Optional[pd.Timestamp]]:
    if df is None or df.empty:
        return None, None
    mn, mx, _ = _date_bounds_for_df(df)
    if mn is None or mx is None:
        st.info("No date column detected for this section.")
        return None, None
    import datetime
    min_d = mn.date()
    max_d = mx.date()
    today_d = datetime.date.today()
    max_selectable = min(max_d, today_d)
    if max_selectable < min_d:
        max_selectable = min_d
        st.warning(
            f"{label}: all available data starts after today ({min_d}). "
            "Date range has been clamped to the first available date."
        )
    anchor_end = max_selectable
    presets = [
        "Custom",
        "Past week",
        "Past month",
        "Past 3 months",
        "Past 6 months",
        "Past year",
        "Past 2 years",
    ]
    days_map = {
        "Past week": 7,
        "Past month": 30,
        "Past 3 months": 90,
        "Past 6 months": 180,
        "Past year": 365,
        "Past 2 years": 730,
    }
    preset_key = f"{key}_preset"
    dates_key = f"{key}_dates"
    last_preset_key = f"{key}_last_preset"
    preset = st.selectbox(
        f"{label} — quick range",
        options=presets,
        index=0,
        key=preset_key,
        help="Choose a preset, or pick Custom to select exact dates.",
    )
    if preset in days_map:
        start_default = max(min_d, (pd.to_datetime(anchor_end) - pd.Timedelta(days=days_map[preset])).date())
        end_default = anchor_end
    else:
        start_default = min_d
        end_default = anchor_end
    prev = st.session_state.get(last_preset_key)
    if prev != preset:
        st.session_state[dates_key] = (start_default, end_default)
        st.session_state[last_preset_key] = preset
        st.rerun()
    if dates_key in st.session_state:
        v = st.session_state[dates_key]
        if isinstance(v, tuple) and len(v) == 2:
            s, e = v
            if hasattr(s, "date"):
                s = s.date()
            if hasattr(e, "date"):
                e = e.date()
            s = min(max(s, min_d), max_selectable)
            e = min(max(e, min_d), max_selectable)
            if e < s:
                e = s
            st.session_state[dates_key] = (s, e)
        else:
            d = v.date() if hasattr(v, "date") else v
            d = min(max(d, min_d), max_selectable)
            st.session_state[dates_key] = d
    dr = st.date_input(
        label,
        min_value=min_d,
        max_value=max_selectable,
        key=dates_key,
        help="Filters only this section.",
    )
    if isinstance(dr, tuple) and len(dr) == 2:
        start_d, end_d = dr
    else:
        start_d, end_d = start_default, end_default
    start_ts = pd.to_datetime(start_d)
    end_ts = pd.to_datetime(end_d) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    return start_ts, end_ts
def filter_by_date_range(df: pd.DataFrame, start_ts: Optional[pd.Timestamp], end_ts: Optional[pd.Timestamp]) -> pd.DataFrame:
    if start_ts is None or end_ts is None:
        return df
    dc = _get_date_col(df)
    if not dc:
        return df
    tmp = df.copy()
    tmp[dc] = pd.to_datetime(tmp[dc], errors="coerce")
    tmp = tmp.dropna(subset=[dc])
    return tmp[(tmp[dc] >= start_ts) & (tmp[dc] <= end_ts)]
def filter_df(df: pd.DataFrame, start_ts: Optional[pd.Timestamp], end_ts: Optional[pd.Timestamp]) -> pd.DataFrame:
    return filter_by_date_range(filter_by_team(df), start_ts, end_ts)
st.markdown(f"**Selected teams:** {len(team_filter)}")
if not team_filter:
    st.warning("No teams selected.")
    st.stop()
tabs = st.tabs(["Overview", "Non-WIP"])
def _get_metrics_df() -> Optional[pd.DataFrame]:
    if "metrics" in data:
        d = filter_by_team(data["metrics"])
        if not d.empty:
            return d
    if "metrics_aggregate_dev" in data:
        d = filter_by_team(data["metrics_aggregate_dev"])
        if not d.empty:
            return d
    if "NS_WIP" in data:
        d = filter_by_team(data["NS_WIP"])
        if not d.empty:
            return d
    if "CRM_WIP" in data:
        d = filter_by_team(data["CRM_WIP"])
        if not d.empty:
            return d
    return None
def _get_nonwip_df() -> Optional[pd.DataFrame]:
    if "ns_non_wip_activities" in data:
        d = filter_by_team(data["ns_non_wip_activities"])
        if not d.empty:
            return d
    if "crm_non_wip_activities" in data:
        d = filter_by_team(data["crm_non_wip_activities"])
        if not d.empty:
            return d
    if "non_wip" in data:
        d = filter_by_team(data["non_wip"])
        if not d.empty:
            return d
    if "non_wip_activities" in data:
        d = filter_by_team(data["non_wip_activities"])
        if not d.empty:
            return d
    return None
def _metrics_cols(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    return {
        "date": _get_date_col(df),
        "wip_hours": _first_col(df, ["completed_hours", "wip_hours", "completedhours"]),
        "avail_hours": _first_col(df, ["total_available_hours", "person_hours", "total_available_hours."]),
        "hc_used": _first_col(df, ["actual_hc_used", "actual_hc_used."]),
        "hc_in_wip": _first_col(df, ["hc_in_wip"]),
        "people_in_wip": _first_col(df, ["people_in_wip"]),
        "person_hours_json": _first_col(df, ["person_hours"]),
    }
def _nonwip_cols(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    return {
        "date": _get_date_col(df),
        "total_nonwip": _first_col(df, ["total_non-wip_hours", "total_non_wip_hours", "total_non_wip_hours."]),
        "people_count": _first_col(df, ["people_count", "people_count."]),
        "pct_in_wip": _first_col(df, ["%_in_wip", "%_in_wip."]),
        "ooohours": _first_col(df, ["ooo_hours", "ooo_hours."]),
        "nonwip_by_person_json": _first_col(df, ["non-wip_by_person", "non_wip_by_person", "non_wip_by_person."]),
        "wip_workers_json": _first_col(df, ["wip_workers", "wip_workers."]),
        "wip_workers_count": _first_col(df, ["wip_workers_count", "wip_workers_count."]),
        "wip_workers_ooo": _first_col(df, ["wip_workers_ooo_hours", "wip_workers_ooo_hours."]),
        "activities_json": _first_col(df, ["non-wip_activities", "non_wip_activities"]),
    }
def _people_lookup_by_week(dfnw: pd.DataFrame) -> Dict[pd.Timestamp, float]:
    out: Dict[pd.Timestamp, float] = {}
    nwc = _nonwip_cols(dfnw)
    if not (nwc["date"] and nwc["people_count"]):
        return out
    tmp = dfnw.copy()
    tmp[nwc["date"]] = pd.to_datetime(tmp[nwc["date"]], errors="coerce")
    tmp = tmp.dropna(subset=[nwc["date"]])
    tmp["people_count"] = _to_num(tmp[nwc["people_count"]]).fillna(0.0)
    grp = tmp.groupby(nwc["date"], as_index=False)["people_count"].sum()
    for _, r in grp.iterrows():
        out[pd.to_datetime(r[nwc["date"]])] = float(r["people_count"])
    return out
def _weekly_rollup_summary(
    dfm: Optional[pd.DataFrame],
    dfnw: Optional[pd.DataFrame],
    dfnw_act: Optional[pd.DataFrame],
    denom_mode: str,
    wd: int,
) -> Tuple[
    Optional[float], Optional[float], Optional[float], Optional[float],
    Optional[float], Optional[float], Optional[float], Optional[float],
]:
    if dfm is None or dfm.empty or dfnw is None or dfnw.empty:
        return (None, None, None, None, None, None, None, None)
    mc = _metrics_cols(dfm)
    nwc = _nonwip_cols(dfnw)
    if not (mc["date"] and mc["wip_hours"] and nwc["date"]):
        return (None, None, None, None, None, None, None, None)
    m = dfm.copy()
    m[mc["date"]] = _safe_to_datetime(m, mc["date"])
    m = m.dropna(subset=[mc["date"]])
    m["week_start"] = _weekly_start(m[mc["date"]])
    m["completed"] = _to_num(m[mc["wip_hours"]]).fillna(0.0)
    m_wk = m.groupby("week_start", as_index=False).agg(completed=("completed", "sum"))
    n = dfnw.copy()
    n[nwc["date"]] = _safe_to_datetime(n, nwc["date"])
    n = n.dropna(subset=[nwc["date"]])
    n["week_start"] = _weekly_start(n[nwc["date"]])
    if nwc["people_count"] and nwc["people_count"] in n.columns:
        n["people_count"] = _to_num(n[nwc["people_count"]]).fillna(0.0)
    else:
        n["people_count"] = 0.0
    if nwc["total_nonwip"] and nwc["total_nonwip"] in n.columns:
        n["nonwip_total"] = _to_num(n[nwc["total_nonwip"]]).fillna(0.0)
    else:
        n["nonwip_total"] = 0.0
    if nwc["ooohours"] and nwc["ooohours"] in n.columns:
        n["ooo_total"] = _to_num(n[nwc["ooohours"]]).fillna(0.0)
    else:
        n["ooo_total"] = 0.0
    def _col_num(colkey: str) -> pd.Series:
        c = nwc.get(colkey)
        if c and c in n.columns:
            return _to_num(n[c]).fillna(0.0)
        return pd.Series([0.0] * len(n), index=n.index)
    n["wip_workers_count"] = _col_num("wip_workers_count")
    n["wip_workers_ooo"] = _col_num("wip_workers_ooo")
    wip_workers_by_week: Dict[pd.Timestamp, List[str]] = {}
    if nwc.get("wip_workers_json") and nwc["wip_workers_json"] in n.columns:
        for _, r in n.iterrows():
            wk = r["week_start"]
            payload = _loads_json_maybe(r[nwc["wip_workers_json"]])
            if isinstance(payload, list):
                wip_workers_by_week[pd.to_datetime(wk)] = [str(x).strip() for x in payload if str(x).strip()]
    n_wk = n.groupby("week_start", as_index=False).agg(
        people_count=("people_count", "sum"),
        nonwip_total=("nonwip_total", "sum"),
        ooo_total=("ooo_total", "sum"),
        wip_workers_count=("wip_workers_count", "sum"),
        wip_workers_ooo=("wip_workers_ooo", "sum"),
    )
    nonwip_wipworkers_by_week: Dict[pd.Timestamp, float] = {}
    if denom_mode == "HC that worked in WIP" and dfnw_act is not None and not dfnw_act.empty:
        act = dfnw_act.copy()
        adc = _get_date_col(act)
        if adc is not None:
            act[adc] = _safe_to_datetime(act, adc)
            act = act.dropna(subset=[adc])
            act["week_start"] = _weekly_start(act[adc])
            act_cols = {_norm(c): c for c in act.columns}
            by_person_col = act_cols.get("non_wip_by_person") or act_cols.get("non-wip_by_person") or act_cols.get("nonwip_by_person")
            if by_person_col:
                for _, r in act.iterrows():
                    wk = pd.to_datetime(r["week_start"])
                    wips = set(wip_workers_by_week.get(wk, []))
                    if not wips:
                        continue
                    dct = _loads_json_maybe(r[by_person_col])
                    if not isinstance(dct, dict):
                        continue
                    s = 0.0
                    for name, hrs in dct.items():
                        if str(name).strip() in wips:
                            try:
                                s += float(hrs)
                            except Exception:
                                pass
                    nonwip_wipworkers_by_week[wk] = nonwip_wipworkers_by_week.get(wk, 0.0) + s
    wk = m_wk.merge(n_wk, on="week_start", how="inner").sort_values("week_start")
    if wk.empty:
        return (None, None, None, None, None, None, None, None)
    if denom_mode == "Total HC on team":
        wk["denom_hc"] = wk["people_count"]
        wk["ooo_use"] = wk["ooo_total"]
        wk["nonwip_use"] = wk["nonwip_total"]
    else:
        wk["denom_hc"] = wk["wip_workers_count"]
        wk["ooo_use"] = wk["wip_workers_ooo"]
        wk["nonwip_use"] = wk["week_start"].map(lambda x: nonwip_wipworkers_by_week.get(pd.to_datetime(x), float("nan")))
        wk["nonwip_use"] = wk["nonwip_use"].where(wk["nonwip_use"].notna(), wk["nonwip_total"])
    wk.loc[wk["denom_hc"] <= 0, "denom_hc"] = pd.NA
    wk["wip_daily_pp"] = wk["completed"] / (wk["denom_hc"] * float(wd))
    wk["nonwip_daily_pp"] = wk["nonwip_use"] / (wk["denom_hc"] * float(wd))
    wk["capacity_weekly"] = wk["denom_hc"] * float(wd) * 8.0
    wk["unacct_weekly"] = wk["capacity_weekly"] - wk["ooo_use"] - wk["completed"] - wk["nonwip_use"]
    wk["pct_wip"] = 100.0 * (wk["wip_daily_pp"] / 8.0)
    wk["pct_nonwip"] = 100.0 * (wk["nonwip_daily_pp"] / 8.0)
    wk["pct_ooo"] = 100.0 * (wk["ooo_use"] / wk["capacity_weekly"])
    wk["pct_unacct"] = 100.0 - (wk["pct_wip"] + wk["pct_nonwip"] + wk["pct_ooo"])
    avg_wip_daily_pp = float(wk["wip_daily_pp"].dropna().mean()) if wk["wip_daily_pp"].notna().any() else None
    avg_nonwip_daily_pp = float(wk["nonwip_daily_pp"].dropna().mean()) if wk["nonwip_daily_pp"].notna().any() else None
    avg_ooo_weekly = float(wk["ooo_use"].dropna().mean()) if wk["ooo_use"].notna().any() else None
    avg_unacct_weekly = float(wk["unacct_weekly"].dropna().mean()) if wk["unacct_weekly"].notna().any() else None
    pct_wip = float(wk["pct_wip"].dropna().mean()) if wk["pct_wip"].notna().any() else None
    pct_nonwip = float(wk["pct_nonwip"].dropna().mean()) if wk["pct_nonwip"].notna().any() else None
    pct_ooo = float(wk["pct_ooo"].dropna().mean()) if wk["pct_ooo"].notna().any() else None
    pct_unacct = float(wk["pct_unacct"].dropna().mean()) if wk["pct_unacct"].notna().any() else None
    return (
        avg_wip_daily_pp, avg_nonwip_daily_pp, avg_ooo_weekly, avg_unacct_weekly,
        pct_wip, pct_nonwip, pct_ooo, pct_unacct
    )
with tabs[0]:
    dfm_raw = _get_metrics_df()
    dfnw_raw = _get_nonwip_df()
    bounds_df = dfm_raw if (dfm_raw is not None and not dfm_raw.empty) else dfnw_raw
    ov_start, ov_end = section_date_range("Overview date range", bounds_df, key="dr_overview")
    dfm = filter_by_date_range(dfm_raw, ov_start, ov_end) if dfm_raw is not None else None
    dfnw = filter_by_date_range(dfnw_raw, ov_start, ov_end) if dfnw_raw is not None else None
    people_by_week: Dict[pd.Timestamp, float] = {}
    if dfnw is not None and not dfnw.empty:
        people_by_week = _people_lookup_by_week(dfnw)
    wd = _workdays_per_week_assumption()
    st.subheader("Summary")
    denom_mode = st.radio(
        "Per-person denominator for WIP daily hours",
        options=["Total HC on team", "HC that worked in WIP"],
        index=0,
        horizontal=True,
        help="Team-total daily WIP = Completed Hours/5. Per-person daily WIP divides by headcount too.",
    )
    dfnw_act_raw = filter_by_team(data["non_wip_activities"]) if "non_wip_activities" in data else None
    dfnw_act = filter_by_date_range(dfnw_act_raw, ov_start, ov_end) if dfnw_act_raw is not None else None
    (
        avg_daily_wip_per_person,
        avg_daily_nonwip_per_person,
        avg_weekly_ooo_hours,
        avg_weekly_unacct_hours,
        pct_wip,
        pct_nonwip,
        pct_ooo,
        pct_unacct,
    ) = _weekly_rollup_summary(dfm, dfnw, dfnw_act, denom_mode=denom_mode, wd=wd)
    st.markdown("""
    <style>
    div[data-testid="stMetric"]{ text-align: center; }
    label[data-testid="stMetricLabel"]{ display: block; width: 100%; text-align: center; margin: 0; }
    label[data-testid="stMetricLabel"] p{ text-align: center !important; margin: 0 !important; }
    div[data-testid="stMetricValue"]{ text-align: center !important; width: 100%; }
    </style>
    """, unsafe_allow_html=True)
    _, c1, c2, _ = st.columns([1.2, 1.2, 1.2, 1.2])
    c1.metric("Avg Per Person **WIP** Daily Hours", f"{avg_daily_wip_per_person:.2f}" if avg_daily_wip_per_person is not None else "—")
    c2.metric("Avg Per Person **Non-WIP** Daily Hours", f"{avg_daily_nonwip_per_person:.2f}" if avg_daily_nonwip_per_person is not None else "—")
    _, p1, p2, _ = st.columns([1.2, 1.2, 1.2, 1.2])
    p1.metric("**WIP** Ratio", f"{pct_wip:.1f}%" if pct_wip is not None else "—")
    p2.metric("**Non-WIP** Ratio", f"{pct_nonwip:.1f}%" if pct_nonwip is not None else "—")
    st.divider()
    _, _, c3, c4,_, _, _ = st.columns([1.35, 1.2, 1.2, 1.2, 1.2, 1.0, 0.5])
    c3.metric("Avg **OOO** Weekly Hours", f"{avg_weekly_ooo_hours:.2f}" if avg_weekly_ooo_hours is not None else "—")
    c4.metric("Avg **Unaccounted** Weekly Hours", f"{avg_weekly_unacct_hours:.2f}" if avg_weekly_unacct_hours is not None else "—")
    _, _, p3, p4, _, _, _ = st.columns([1.35, 1.2, 1.2, 1.2, 1.2, 1.0, 0.5])
    p3.metric("**OOO** % of week", f"{pct_ooo:.1f}%" if pct_ooo is not None else "—")
    p4.metric("**Unaccounted** % remaining", f"{pct_unacct:.1f}%" if pct_unacct is not None else "—")
    st.divider()
    st.subheader("Trend: avg daily WIP hours (week over week)")
    if dfm is not None:
        mc = _metrics_cols(dfm)
        if mc["date"] and mc["wip_hours"]:
            temp = dfm.copy()
            temp[mc["date"]] = _safe_to_datetime(temp, mc["date"])
            temp = temp.dropna(subset=[mc["date"]]).sort_values(mc["date"])
            temp["wip_hours"] = _to_num(temp[mc["wip_hours"]]).fillna(0.0)
            temp["week_start"] = _weekly_start(temp[mc["date"]])
            today = pd.Timestamp.now()
            current_week_start = today.to_period("W-MON").start_time
            temp = temp[temp["week_start"] <= current_week_start]
            wk = (
                temp.groupby("week_start", as_index=False)
                .agg(wip_hours=("wip_hours", "sum"))
                .sort_values("week_start")
            )
            wk["team_total"] = wk["wip_hours"] / float(wd)
            wk["per_person"] = pd.NA
            if denom_mode == "Total HC on team" and people_by_week:
                people_by_week_start = {
                    pd.to_datetime(k).to_period("W-MON").start_time: v
                    for k, v in people_by_week.items()
                }
                wk["people_count"] = wk["week_start"].map(people_by_week_start)
                wk["per_person"] = wk["wip_hours"] / (float(wd) * wk["people_count"])
                wk.loc[wk["people_count"].fillna(0) <= 0, "per_person"] = pd.NA
            elif mc["hc_in_wip"] and mc["hc_in_wip"] in temp.columns:
                temp["hc_in_wip"] = _to_num(temp[mc["hc_in_wip"]]).fillna(0.0)
                wk_hc = (
                    temp.groupby("week_start", as_index=False)
                    .agg(hc_in_wip=("hc_in_wip", "sum"))
                    .sort_values("week_start")
                )
                wk = wk.merge(wk_hc, on="week_start", how="left")
                wk["per_person"] = wk["wip_hours"] / (float(wd) * wk["hc_in_wip"])
                wk.loc[wk["hc_in_wip"].fillna(0) <= 0, "per_person"] = pd.NA
            wk["team_total_up_only"] = wk["team_total"].cummax()
            if wk["per_person"].notna().any():
                wk["per_person_up_only"] = wk["per_person"].cummax()
            view = st.selectbox("Trend view", ["Group total", "Per person"], index=1)
            if view == "Group total":
                st.line_chart(wk.set_index("week_start")["team_total_up_only"])
            else:
                if "per_person_up_only" not in wk.columns or wk["per_person_up_only"].notna().sum() == 0:
                    st.info("Per-person trend not available (no People Count / HC in WIP found for selected range).")
                else:
                    st.line_chart(wk.set_index("week_start")["per_person_up_only"])
        else:
            st.info("Need Week/period_date + Completed Hours to show trend.")
    else:
        st.info("No metrics data loaded for selected teams.")
EXCLUDED_NON_WIP = {"ooo", "non-wip", "non_wip", "other", "other team wip"}
def _norm_activity_name(val: Any) -> str:
    return str(val).strip().lower().replace("_", "-")
with tabs[1]:
    if (
        "non_wip" not in data
        and "non_wip_activities" not in data
        and "ns_non_wip_activities" not in data
    ):
        st.info("No non-WIP CSVs found (expected `non_wip.csv` and/or `non_wip_activities.csv`).")
        st.stop()
    st.markdown("### Non-WIP activities")
    source_raw = None
    if "ns_non_wip_activities" in data:
        cand = filter_by_team(data["ns_non_wip_activities"])
        if not cand.empty:
            source_raw = cand
    if source_raw is None and "non_wip" in data:
        cand = filter_by_team(data["non_wip"])
        if not cand.empty:
            source_raw = cand
    if source_raw is None and "non_wip_activities" in data:
        cand = filter_by_team(data["non_wip_activities"])
        if not cand.empty:
            source_raw = cand
    if source_raw is None or source_raw.empty:
        st.info("No Non-WIP activity data available after team filtering.")
        st.stop()
    nw_start, nw_end = section_date_range("Non-WIP date range", source_raw, key="dr_nonwip")
    source_df = filter_by_date_range(source_raw, nw_start, nw_end)
    if source_df.empty:
        st.info("No Non-WIP activity data available in this date range.")
        st.stop()
    dc = _get_date_col(source_df)
    json_col = None
    for c in source_df.columns:
        if _norm(c) in {"non-wip_activities", "non_wip_activities"}:
            json_col = c
            break
    if not (dc and json_col):
        st.info("Need `Week/period_date` and `Non-WIP Activities` (JSON list) to roll up activities.")
        st.stop()
    tmp = source_df.copy()
    tmp[dc] = _safe_to_datetime(tmp, dc)
    tmp = tmp.dropna(subset=[dc]).sort_values(dc)
    rows: List[Dict[str, Any]] = []
    for _, r in tmp.iterrows():
        wk = r[dc]
        payload = _loads_json_maybe(r[json_col])
        if not payload:
            continue
        if isinstance(payload, list):
            for item in payload:
                if not isinstance(item, dict):
                    continue
                act = item.get("activity") or item.get("Activity") or item.get("type")
                hrs = item.get("hours") or item.get("Hours")
                if act is None or hrs is None:
                    continue
                rows.append(
                    {
                        "week": wk,
                        "activity": str(act).strip(),
                        "hours": float(hrs) if str(hrs) != "" else 0.0,
                    }
                )
    if not rows:
        st.info("No parsable activity rows found in the JSON column.")
        st.stop()
    act_df = pd.DataFrame(rows)
    act_df["week"] = pd.to_datetime(act_df["week"], errors="coerce")
    act_df = act_df.dropna(subset=["week"])
    act_df["week_start"] = _weekly_start(act_df["week"])
    weekly_raw = (
        act_df.groupby(["week_start", "activity"], as_index=False)
        .agg(hours=("hours", "sum"))
    )
    normalised_chunks: List[pd.DataFrame] = []
    for wk_val, grp in weekly_raw.groupby("week_start"):
        cat = grp[["activity", "hours"]].rename(
            columns={"activity": "Activity", "hours": "Hours"}
        )
        cat_norm = split_nonwip_activity_minutes(cat)
        cat_norm["week_start"] = wk_val
        normalised_chunks.append(cat_norm)
    if normalised_chunks:
        weekly_by_activity = (
            pd.concat(normalised_chunks, ignore_index=True)
            .rename(columns={"Activity": "activity", "Hours": "hours"})
            .sort_values(["week_start", "hours"], ascending=[True, False])
        )
    else:
        weekly_by_activity = weekly_raw.copy()
    weekly_by_activity = weekly_by_activity[
        ~weekly_by_activity["activity"].map(_norm_activity_name).isin(EXCLUDED_NON_WIP)
    ]
    if weekly_by_activity.empty:
        st.info("No Non-WIP activity data available after exclusions.")
        st.stop()
    total_hours = (
        weekly_by_activity.groupby("activity", as_index=False)
        .agg(total_hours=("hours", "sum"))
        .sort_values("total_hours", ascending=False)
        .head(15)
        .reset_index(drop=True)
    )
    if total_hours.empty:
        st.info("No chartable Non-WIP activity data available after exclusions.")
        st.stop()
    import matplotlib.pyplot as plt
    def _short_label(s: Any, max_len: int = 22) -> str:
        s = str(s).strip()
        return s if len(s) <= max_len else s[: max_len - 3] + "..."
    chart_df = total_hours.copy()
    chart_df["label"] = chart_df["activity"].map(lambda x: _short_label(x, 22))
    fig, ax = plt.subplots(figsize=(14, 5.5))
    bars = ax.bar(chart_df["label"], chart_df["total_hours"])
    ax.set_ylabel("Total Hours")
    ax.set_xlabel("Non-WIP Activity")
    ax.set_title("Top 15 Non-WIP Activities by Total Hours")
    ax.tick_params(axis="x", rotation=45, labelsize=9)
    plt.setp(ax.get_xticklabels(), ha="right")
    for bar, val in zip(bars, chart_df["total_hours"]):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height(),
            f"{val:.1f}",
            ha="center",
            va="bottom",
            fontsize=9,
        )
    fig.tight_layout()
    st.pyplot(fig)
    st.caption("Top 15 activities by total hours for the selected period, sorted highest to lowest from left to right.")
    st.divider()
    st.markdown("#### Activity breakdown — pie chart")
    pie_start, pie_end = section_date_range("Pie chart date range", source_raw, key="dr_nonwip_pie")
    pie_source_df = filter_by_date_range(source_raw, pie_start, pie_end)
    if pie_source_df.empty:
        st.info("No Non-WIP activity data in the selected pie chart date range.")
        st.stop()
    pie_dc = _get_date_col(pie_source_df)
    pie_json_col = None
    for c in pie_source_df.columns:
        if _norm(c) in {"non-wip_activities", "non_wip_activities"}:
            pie_json_col = c
            break
    pie_rows: List[Dict[str, Any]] = []
    if pie_dc and pie_json_col:
        pie_tmp = pie_source_df.copy()
        pie_tmp[pie_dc] = _safe_to_datetime(pie_tmp, pie_dc)
        pie_tmp = pie_tmp.dropna(subset=[pie_dc]).sort_values(pie_dc)
        for _, r in pie_tmp.iterrows():
            payload = _loads_json_maybe(r[pie_json_col])
            if not payload:
                continue
            if isinstance(payload, list):
                for item in payload:
                    if not isinstance(item, dict):
                        continue
                    act = item.get("activity") or item.get("Activity") or item.get("type")
                    hrs = item.get("hours") or item.get("Hours")
                    if act is None or hrs is None:
                        continue
                    pie_rows.append(
                        {
                            "activity": str(act).strip(),
                            "hours": float(hrs) if str(hrs) != "" else 0.0,
                        }
                    )
    if pie_rows:
        pie_act_df = pd.DataFrame(pie_rows)
        pie_cat = pie_act_df.rename(columns={"activity": "Activity", "hours": "Hours"})
        pie_cat_norm = split_nonwip_activity_minutes(pie_cat)
        pie_rolled = pie_cat_norm.rename(columns={"Activity": "activity", "Hours": "hours"})
        pie_rolled = pie_rolled.groupby("activity", as_index=False).agg(hours=("hours", "sum"))
        pie_rolled = pie_rolled[
            ~pie_rolled["activity"].map(_norm_activity_name).isin(EXCLUDED_NON_WIP)
        ]
        pie_rolled = pie_rolled.sort_values("hours", ascending=False)
        if pie_rolled.empty:
            st.info('No pie chart data available after excluding "OOO" and "Non-WIP".')
            st.stop()
        if len(pie_rolled) > 14:
            top_pie = pie_rolled.head(14)
            other_pie = pd.DataFrame(
                [{
                    "activity": "Other",
                    "hours": float(pie_rolled["hours"].iloc[14:].sum()),
                }]
            )
            pie_df = pd.concat([top_pie, other_pie], ignore_index=True)
        else:
            pie_df = pie_rolled
        import matplotlib.pyplot as plt
        fig, ax = plt.subplots()
        ax.pie(
            pie_df["hours"],
            labels=pie_df["activity"],
            autopct="%1.0f%%",
            startangle=90,
        )
        ax.axis("equal")
        st.pyplot(fig)
    else:
        st.info("No parsable activity rows found for the selected pie chart date range.")