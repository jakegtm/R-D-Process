"""
R&D Monthly Activity Tracker
Run with: streamlit run app.py
"""

import streamlit as st
import json
from datetime import datetime, date
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import re

# ── Config ────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="R&D Activity Tracker",
    page_icon="🔬",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Constants ─────────────────────────────────────────────────────────────────

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)

EMPLOYEES = [
    "Bob Smith", "Sara", "Doug", "Trevor", "Doni",
    "Jonathan", "Steven", "Michael", "Joe", "Nicole Browne",
]

ENTITIES = ["107", "108", "109", "110"]   # base/default entities — custom ones persist separately

STATUS_LABELS = {
    "in-progress": "🟡 In Progress",
    "submitted":   "🔵 Ready for Review",
    "approved":    "🟢 Approved",
    "rejected":    "🔴 Rejected",
    "archived":    "📦 Archived",
    "not-started": "⚪ Not Started",
}

# Wizard steps — exactly the 9 columns in the Excel template
# (Month/Yr is derived from reporting_month, not asked separately)
WIZARD_STEPS = [
    {
        "field": "business_component",
        "label": "Business Component",
        "type": "text",
        "placeholder": "e.g. Software Development, Hardware Engineering...",
        "question": "What business component does this initiative belong to?",
        "hint": "Enter the area of the business this R&D work supports.",
        "required": True,
    },
    {
        "field": "initiative_name",
        "label": "Initiative Name",
        "type": "text",
        "placeholder": "e.g. Automated Quality Inspection System",
        "question": "What is the name of this R&D initiative?",
        "hint": "Use a short, consistent name — this carries over month-to-month for ongoing work.",
        "required": True,
    },
    {
        "field": "initiative_description",
        "label": "Initiative Description",
        "type": "textarea",
        "placeholder": "We are developing a system that will...",
        "question": "Describe what this initiative aims to achieve.",
        "hint": "Explain the goal and expected business outcome in 2–4 sentences.",
        "required": True,
    },
    {
        "field": "tech_uncertainty",
        "label": "Tech Uncertainty",
        "type": "textarea",
        "placeholder": "It is currently unknown whether... We are testing if...",
        "question": "What technical uncertainty are you working to resolve?",
        "hint": "What scientific or technical question are you trying to answer? What don't you know yet? This is key for R&D eligibility.",
        "required": True,
    },
    {
        "field": "start_date",
        "label": "Start Date",
        "type": "date",
        "question": "When did R&D work on this initiative begin?",
        "hint": "The actual date work first started — not the project plan date.",
        "required": True,
    },
    {
        "field": "expected_end_date",
        "label": "Expected End Date",
        "type": "date",
        "question": "When do you expect this initiative to conclude?",
        "hint": "Your best estimate. This can be updated in future months.",
        "required": True,
    },
    {
        "field": "activities",
        "label": "Activities to Eliminate Technical Uncertainty",
        "type": "textarea",
        "placeholder": "Prototyping new approach, running performance tests, analyzing results...",
        "question": "What activities are being conducted to eliminate the technical uncertainty?",
        "hint": "Describe the specific tasks, experiments, or development work happening this month.",
        "required": True,
    },
    {
        "field": "team_members",
        "label": "Team Members",
        "type": "multiselect",
        "question": "Which team members are working on this initiative?",
        "hint": "Select everyone contributing to this initiative this month.",
        "required": True,
    },
    {
        "field": "notes",
        "label": "Notes",
        "type": "textarea",
        "placeholder": "Optional — blockers, scope changes, anything useful for review...",
        "question": "Any additional notes or comments?",
        "hint": "Include anything helpful for the Oversight Lead's review. This field is optional.",
        "required": False,
    },
]

CARRYOVER_FIELDS = {
    "business_component", "initiative_name", "initiative_description",
    "tech_uncertainty", "start_date", "expected_end_date", "team_members",
}


# ── Date helpers ──────────────────────────────────────────────────────────────

def cur_month() -> str:
    return datetime.now().strftime("%Y-%m")

def prev_month_of(m: str) -> str:
    y, mo = map(int, m.split("-"))
    if mo == 1:
        return f"{y-1}-12"
    return f"{y}-{str(mo-1).zfill(2)}"

def next_month_of(m: str) -> str:
    y, mo = map(int, m.split("-"))
    if mo == 12:
        return f"{y+1}-01"
    return f"{y}-{str(mo+1).zfill(2)}"

def parse_month_label(s: str) -> str:
    """Parses a display label like 'May 2026' back into '2026-05'. Returns '' if unparseable."""
    if not s:
        return ""
    s = str(s).strip()
    for fmt in ("%B %Y", "%b %Y"):
        try:
            dt = datetime.strptime(s, fmt)
            return f"{dt.year}-{str(dt.month).zfill(2)}"
        except Exception:
            continue
    return ""

def fmt_month(m: str) -> str:
    if not m:
        return ""
    try:
        y, mo = m.split("-")
        return datetime(int(y), int(mo), 1).strftime("%B %Y")
    except Exception:
        return ""

def fmt_month_tab(m: str) -> str:
    """Short form for Excel sheet tab names, e.g. 'May 2026'. Never returns empty."""
    if not m:
        return "Report"
    try:
        y, mo = m.split("-")
        return datetime(int(y), int(mo), 1).strftime("%b %Y")
    except Exception:
        return m or "Report"

def available_months() -> list[str]:
    """6 months back + 1 forward, newest first."""
    today = date.today()
    result = []
    for delta in range(-6, 2):
        total = today.month - 1 + delta
        yr    = today.year + total // 12
        mo    = total % 12 + 1
        result.append(f"{yr}-{str(mo).zfill(2)}")
    result.sort(reverse=True)
    return result


# ── Filename helper ───────────────────────────────────────────────────────────

def export_filename(entity: str, reporting_month: str, tag: str = "") -> str:
    """
    107_Group_10_2025_Monthly_R_D_Tracking_Template.xlsx
    107_Group_10_2025_Consolidated_Submitted_Monthly_R_D_Tracking_Template.xlsx
    """
    try:
        dt = datetime.strptime(reporting_month, "%Y-%m")
        mo = str(dt.month)
        yr = str(dt.year)
    except Exception:
        mo, yr = "00", "0000"
    parts = [entity or "000", "Group", mo, yr]
    if tag:
        parts.append(tag)
    parts.append("Monthly_R_D_Tracking_Template")
    return "_".join(parts) + ".xlsx"


# ── Storage ───────────────────────────────────────────────────────────────────

def _safe_name(s: str) -> str:
    return s.replace(" ", "_").replace("/", "_")

def registry_path() -> Path:
    return DATA_DIR / "registry.json"

def load_registry() -> list:
    p = registry_path()
    return json.loads(p.read_text()) if p.exists() else []

def save_registry(users: list):
    registry_path().write_text(json.dumps(users, indent=2))


def custom_entities_path() -> Path:
    return DATA_DIR / "custom_entities.json"

def load_custom_entities() -> list[str]:
    p = custom_entities_path()
    return json.loads(p.read_text()) if p.exists() else []

def save_custom_entities(entities: list[str]):
    custom_entities_path().write_text(json.dumps(entities, indent=2))

def all_entities() -> list[str]:
    """Base entities (107-110) + any custom ones added by users, in order added."""
    custom = load_custom_entities()
    return ENTITIES + [e for e in custom if e not in ENTITIES]

def add_custom_entity(new_entity: str) -> bool:
    """
    Adds a new entity so it persists in the dropdown going forward.
    Returns True if it was newly added, False if it already existed.
    """
    new_entity = str(new_entity).strip()
    if not new_entity:
        return False
    existing = all_entities()
    if new_entity in existing:
        return False
    custom = load_custom_entities()
    custom.append(new_entity)
    save_custom_entities(custom)
    return True


def months_index_path(username: str) -> Path:
    return DATA_DIR / f"{_safe_name(username)}_months.json"

def load_user_months(username: str) -> list:
    p = months_index_path(username)
    return json.loads(p.read_text()) if p.exists() else []

def save_user_months(username: str, months: list):
    months_index_path(username).write_text(json.dumps(months, indent=2))

def submission_path(username: str, month: str) -> Path:
    return DATA_DIR / f"{_safe_name(username)}_{month}.json"

def load_submission(username: str, month: str) -> dict | None:
    p = submission_path(username, month)
    return json.loads(p.read_text()) if p.exists() else None

def save_draft(username: str, draft: dict):
    """
    Always keys by draft['reporting_month'] — this is the fix that ensures
    May submissions don't appear under June.
    """
    month = draft.get("reporting_month") or cur_month()
    submission_path(username, month).write_text(json.dumps(draft, indent=2))
    # Registry
    reg = load_registry()
    if username not in reg:
        reg.append(username)
        save_registry(reg)
    # Per-user month index — only track months that have real content
    # (avoids empty setup sessions polluting the admin view)
    has_content = bool(draft.get("initiatives")) or draft.get("status") not in ("in-progress", None)
    if has_content:
        months = load_user_months(username)
        if month not in months:
            months.append(month)
            save_user_months(username, months)

def load_all_submissions() -> dict:
    """
    Returns { username: { reporting_month: submission_dict } }
    Uses each user's month index so we find ALL months, not just current.
    """
    reg = load_registry()
    all_data: dict[str, dict] = {}
    for user in reg:
        all_data[user] = {}
        for month in load_user_months(user):
            sub = load_submission(user, month)
            if sub:
                all_data[user][month] = sub
    return all_data

def repair_months_index() -> int:
    """
    Scans the data/ folder for all submission JSON files and ensures every
    file with real initiatives is correctly registered in:
      - registry.json  (user list)
      - {username}_months.json  (per-user month index)

    Runs automatically on every admin page load.
    Returns the number of entries that were repaired/added.
    """
    repaired = 0
    pattern  = re.compile(r'^(.+)_(\d{4}-\d{2})\.json$')

    for filepath in DATA_DIR.glob("*.json"):
        # Skip index/registry files
        if filepath.stem in ("registry", ) or filepath.stem.endswith("_months"):
            continue

        m = pattern.match(filepath.name)
        if not m:
            continue

        raw_user = m.group(1).replace("_", " ")   # reverse _safe_name (best effort)
        month    = m.group(2)

        try:
            data = json.loads(filepath.read_text())
        except Exception:
            continue

        # Only index if it has real initiatives
        if not data.get("initiatives"):
            continue

        # Determine the actual username from registry (handles multi-word names)
        reg = load_registry()
        # Try to find matching user in registry
        username = None
        for u in reg:
            if _safe_name(u) == m.group(1):
                username = u
                break
        if username is None:
            # Not in registry — use the raw_user and add to registry
            # Check if the safe name of raw_user matches
            username = raw_user
            if username not in reg:
                reg.append(username)
                save_registry(reg)
                repaired += 1

        # Ensure month is in user's months index
        months = load_user_months(username)
        if month not in months:
            months.append(month)
            save_user_months(username, months)
            repaired += 1

    return repaired


def create_backup_excel(all_data: dict) -> bytes:
    """
    Creates a comprehensive Excel backup of ALL historical data.
    One sheet per entity (same format as consolidated export).
    Only includes initiatives that are currently in submissions
    (deleted initiatives are already removed from the data).
    Includes all statuses — nothing is filtered out.
    """
    combos = get_combos(all_data)
    if not combos:
        # Return a minimal workbook with a message
        wb  = Workbook()
        ws  = wb.active
        ws.title = "No Data"
        ws.cell(1, 1, "No submissions found.")
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()
    return build_excel_consolidated(all_data, combos, submitted_only=False)


def delete_all_history():
    """
    Deletes all submission JSON files and index files from the data/ directory.
    Preserves the directory itself. Irreversible — caller must confirm first.
    """
    for filepath in DATA_DIR.glob("*.json"):
        filepath.unlink()


# ── Excel Import / Restore ─────────────────────────────────────────────────────
# Lets an admin (or user) re-upload a previously downloaded backup/export file
# to reconstruct submissions after the server's storage has been wiped
# (e.g. app restart on ephemeral hosting). Reads the same layout this app
# writes — works with the consolidated/backup format (multi-sheet, "User"
# column) and with individual single-user exports (one sheet, no "User" col).

# Reverse lookup: "Approved" -> "approved", "📦 Archived" -> "archived", etc.
_STATUS_LABEL_TO_KEY: dict[str, str] = {}
for _k, _v in STATUS_LABELS.items():
    _STATUS_LABEL_TO_KEY[_v.strip().lower()] = _k
    _parts = _v.split(" ", 1)
    if len(_parts) > 1:
        _STATUS_LABEL_TO_KEY[_parts[1].strip().lower()] = _k

def _parse_status_label(s: str) -> str:
    if not s:
        return "in-progress"
    return _STATUS_LABEL_TO_KEY.get(str(s).strip().lower(), "in-progress")


def _cell_to_date_str(value) -> str:
    """Normalizes a cell value to a 'YYYY-MM-DD' string, however Excel stored it."""
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    return str(value).strip()


def _find_header_row(ws, max_scan: int = 8) -> tuple[int, dict[str, int]]:
    """
    Scans the first few rows of a sheet to find the header row — identified
    as the row containing 'Initiative Name'. Returns (row_index, {lower_header: col_idx}).
    Returns (0, {}) if no header row is found.
    """
    for r in range(1, max_scan + 1):
        row_vals = {}
        found_marker = False
        for c in range(1, ws.max_column + 1):
            val = ws.cell(r, c).value
            if val is None:
                continue
            sval = str(val).strip()
            row_vals[sval.lower()] = c
            if sval.lower() == "initiative name":
                found_marker = True
        if found_marker:
            return r, row_vals
    return 0, {}


def _extract_entity(ws) -> str:
    """Tries sheet title 'Entity 107', then row 2 subtitle 'Entity: 107'."""
    m = re.search(r"Entity\s+(\w+)", ws.title or "")
    if m:
        return m.group(1)
    subtitle = ws.cell(2, 1).value or ""
    # Subtitle text may live in different columns depending on User-col offset;
    # scan the first few cells of row 2 for the pattern.
    for c in range(1, min(ws.max_column, 10) + 1):
        v = ws.cell(2, c).value
        if v and "Entity:" in str(v):
            m2 = re.search(r"Entity:\s*(\w+)", str(v))
            if m2:
                return m2.group(1)
    return ""


def _extract_filing_and_username_from_subtitle(ws) -> tuple[str, str]:
    """
    For individual (single-user) exports without a 'User' column or
    'Filing Month' column: pulls 'Submitted by: X' and 'Filing: Month Year'
    out of the row-2 subtitle text.
    """
    username, filing = "", ""
    for c in range(1, min(ws.max_column, 10) + 1):
        v = ws.cell(2, c).value
        if not v:
            continue
        sv = str(v)
        m_user = re.search(r"Submitted by:\s*([^|]+?)\s*(?:\||$)", sv)
        if m_user:
            username = m_user.group(1).strip()
        m_fil = re.search(r"Filing:\s*([A-Za-z]+ \d{4})", sv)
        if m_fil:
            filing = parse_month_label(m_fil.group(1))
    return username, filing


def parse_import_workbook(file_bytes: bytes) -> tuple[dict, list[str]]:
    """
    Parses an uploaded .xlsx (consolidated/backup or individual export format)
    back into submission dicts.

    Returns (parsed, warnings) where:
      parsed   = { (username, entity, reporting_month): {
                      "entity": str, "reporting_month": str,
                      "activities_month": str, "status": str,
                      "initiatives": [ {...new_initiative()-shaped dicts...} ]
                  } }
      warnings = list of human-readable strings about rows/sheets that were
                 skipped or had to be guessed at.
    """
    warnings: list[str] = []
    parsed: dict = {}

    try:
        wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    except Exception as e:
        return {}, [f"Could not open file as an Excel workbook: {e}"]

    for ws in wb.worksheets:
        if ws.title in ("No Data",):
            continue

        hdr_row, hdrs = _find_header_row(ws)
        if not hdr_row:
            warnings.append(f"Sheet '{ws.title}': couldn't find a header row — skipped.")
            continue

        has_user_col = "user" in hdrs
        sheet_entity = _extract_entity(ws)
        fallback_username, fallback_filing = _extract_filing_and_username_from_subtitle(ws)

        if not sheet_entity:
            warnings.append(f"Sheet '{ws.title}': couldn't determine Entity — skipped.")
            continue

        col = lambda name: hdrs.get(name.lower())
        n_rows_in_sheet = 0

        for r in range(hdr_row + 1, ws.max_row + 1):
            name_col = col("initiative name")
            if not name_col:
                continue
            iname = ws.cell(r, name_col).value
            if not iname or not str(iname).strip():
                continue   # blank row

            def get(colname, default=""):
                ci = col(colname)
                if not ci:
                    return default
                v = ws.cell(r, ci).value
                return v if v is not None else default

            username = str(get("user", "")).strip() or fallback_username
            if not username:
                warnings.append(f"Sheet '{ws.title}' row {r}: no username found — skipped.")
                continue

            month_yr_label = str(get("month/yr", "")).strip()
            activities_month = parse_month_label(month_yr_label)

            filing_label = str(get("filing month", "")).strip()
            reporting_month = parse_month_label(filing_label)
            if not reporting_month:
                reporting_month = fallback_filing
            if not reporting_month and activities_month:
                # Last resort: assume the standard one-month convention
                reporting_month = next_month_of(activities_month)
                warnings.append(
                    f"Sheet '{ws.title}' row {r}: no Filing Month found — "
                    f"assumed {fmt_month(reporting_month)} (one month after activities)."
                )
            if not reporting_month:
                warnings.append(f"Sheet '{ws.title}' row {r}: couldn't determine a filing period — skipped.")
                continue

            status_key = _parse_status_label(str(get("status", "")))

            init = new_initiative()
            init["business_component"]     = str(get("business component", ""))
            init["initiative_name"]        = str(iname).strip()
            init["initiative_description"] = str(get("initiative description", ""))
            init["tech_uncertainty"]        = str(get("tech uncertainty", ""))
            init["start_date"]              = _cell_to_date_str(get("start date", None)) or None
            init["expected_end_date"]       = _cell_to_date_str(get("expected end date", None)) or None
            init["activities"]              = str(get("activities to eliminate technical uncertainty", ""))
            team_raw                        = str(get("team members", ""))
            init["team_members"]            = [t.strip() for t in team_raw.split(",") if t.strip()]
            init["notes"]                   = str(get("notes", ""))
            init["month_yr"]                = activities_month
            init["carry_over"]              = False

            if status_key in ("approved", "archived"):
                init["initiative_status"] = "approved"
            elif status_key == "rejected":
                init["initiative_status"] = "returned"
            else:
                init["initiative_status"] = "active"

            comp_raw = get("completion date", "")
            if comp_raw and status_key in ("approved", "archived"):
                for fmt_str in ("%Y-%m-%d %H:%M", "%Y-%m-%d"):
                    try:
                        dt = datetime.strptime(str(comp_raw).strip(), fmt_str)
                        init["approved_at"] = int(dt.timestamp() * 1000)
                        break
                    except Exception:
                        continue

            key = (username, sheet_entity, reporting_month)
            if key not in parsed:
                parsed[key] = {
                    "entity":           sheet_entity,
                    "reporting_month":  reporting_month,
                    "activities_month": activities_month or prev_month_of(reporting_month),
                    "status":           status_key,
                    "initiatives":      [],
                }
            parsed[key]["initiatives"].append(init)
            n_rows_in_sheet += 1

        if n_rows_in_sheet == 0:
            warnings.append(f"Sheet '{ws.title}': no usable initiative rows found.")

    return parsed, warnings


def apply_import(parsed: dict, overwrite: bool = False) -> dict:
    """
    Commits parsed import data to storage via save_draft.
    Skips (username, entity, period) combos that already have a real
    submission unless overwrite=True.
    Returns {"imported": [...], "skipped": [...]} — lists of
    "username — Entity X — Month Year" strings.
    """
    imported, skipped = [], []
    for (username, entity, rm), data in parsed.items():
        label = f"{username} — Entity {entity} — {fmt_month(rm)}"
        existing = load_submission(username, rm)
        if existing and existing.get("initiatives") and not overwrite:
            skipped.append(label)
            continue

        draft = {
            "initiatives":      data["initiatives"],
            "status":           data["status"] if data["status"] in STATUS_LABELS else "approved",
            "entity":           entity,
            "reporting_month":  rm,
            "activities_month": data["activities_month"],
        }
        save_draft(username, draft)
        imported.append(label)
    return {"imported": imported, "skipped": skipped}


def get_combos(all_data: dict) -> list[tuple[str, str]]:
    """
    Returns sorted list of unique (entity, reporting_month) tuples
    found across submissions that have at least one initiative.
    """
    combos: set[tuple[str, str]] = set()
    for user, months in all_data.items():
        for month, sub in months.items():
            if not sub.get("initiatives"):
                continue   # skip empty setup drafts
            e  = sub.get("entity", "")
            rm = sub.get("reporting_month", month)
            if e and rm:
                combos.add((e, rm))
    return sorted(combos, key=lambda x: (x[1], x[0]))


# ── Memory / history helpers ─────────────────────────────────────────────────

def get_past_ongoing_initiatives(username: str, current_reporting_month: str) -> list[dict]:
    """
    Returns all initiatives from ALL past submissions for this user where
    expected_end_date is still in the future (or not set).
    Deduplicates by initiative_name so the same initiative doesn't appear
    multiple times if it was carried over in earlier months.
    """
    months    = load_user_months(username)
    today     = date.today()
    seen      = set()
    result    = []

    # Walk months newest-first so we get the most recent version of each initiative
    for month in sorted(months, reverse=True):
        if month >= current_reporting_month:
            continue   # skip current or future periods
        sub = load_submission(username, month)
        if not sub:
            continue
        for init in (sub.get("initiatives") or []):
            name = init.get("initiative_name", "").strip().lower()
            if name in seen:
                continue   # already have a newer version
            seen.add(name)
            end = init.get("expected_end_date")
            if end:
                try:
                    if datetime.strptime(str(end), "%Y-%m-%d").date() < today:
                        continue   # already ended
                except Exception:
                    pass
            result.append(init)
    return result


def get_user_history(username: str) -> list[tuple[str, dict]]:
    """
    Returns [(month_str, submission_dict), ...] sorted newest-first.
    """
    months = load_user_months(username)
    history = []
    for month in sorted(months, reverse=True):
        sub = load_submission(username, month)
        if sub:
            history.append((month, sub))
    return history


# ── Entity rollover ───────────────────────────────────────────────────────────

def rollover_entity(
    all_data: dict,
    source_entity: str,
    source_month: str,
    target_month: str,
) -> list[str]:
    """
    For every user who has a submission under (source_entity, source_month),
    create a new in-progress draft under (source_entity, target_month)
    — unless they already have one there.

    Carries ALL initiatives regardless of end date — the admin is explicitly
    choosing what to roll over, so the app doesn't second-guess them.
    Returns list of usernames that were rolled over.
    """
    rolled: list[str] = []

    for username, months in all_data.items():
        src_sub = months.get(source_month)
        if not src_sub or src_sub.get("entity") != source_entity:
            continue

        # Don't overwrite an existing submission for that period
        existing = load_submission(username, target_month)
        # Only block if they have a real submission (with initiatives),
        # not just an empty setup draft saved when they picked a reporting month
        if existing and existing.get("entity") == source_entity and existing.get("initiatives"):
            continue

        initiatives = src_sub.get("initiatives") or []
        if not initiatives:
            continue

        # Deep-copy every initiative so all content is fully preserved.
        # Only reset the per-review-cycle fields (approval state) since
        # this is a new period with a fresh review cycle.
        import copy as _copy
        rolled_inits = []
        for i in initiatives:
            init = _copy.deepcopy(i)
            init["id"]                = f"{int(__import__('time').time()*1000)}{len(rolled_inits)}"
            init["initiative_status"] = "active"   # new period = new review
            init["approved_at"]       = None
            init["returned_at"]       = None
            init["carry_over"]        = True
            # Clear month_yr so the new period's activities_month is used in the export.
            # Historical month_yr is preserved in the source period's file.
            init["month_yr"]          = ""
            rolled_inits.append(init)

        new_sub: dict = {
            "initiatives":      rolled_inits,
            "status":           "in-progress",
            "entity":           source_entity,
            "reporting_month":  target_month,
            # Default activities_month to the month before the filing month.
            # User can change this in Report Setup if needed.
            "activities_month": prev_month_of(target_month),
        }
        save_draft(username, new_sub)
        rolled.append(username)

    return rolled


# ── Initiative helpers ────────────────────────────────────────────────────────

def new_initiative() -> dict:
    return {
        "id": f"{int(datetime.now().timestamp()*1000)}",
        "business_component":    "",
        "initiative_name":       "",
        "initiative_description":"",
        "tech_uncertainty":      "",
        "start_date":            None,
        "expected_end_date":     None,
        "activities":            "",
        "team_members":          [],
        "notes":                 "",
        "carry_over":            False,
        # Per-initiative status (set by admin actions)
        "initiative_status":     "active",   # active | approved | returned
        "approved_at":           None,       # ms timestamp when admin approved
        "returned_at":           None,       # ms timestamp when admin returned
        # month_yr: the reporting period this initiative row belongs to.
        # Set when first saved; preserved through admin rollovers so each row
        # always shows the correct original period in the export.
        "month_yr":              "",
    }

def carryover_initiative(src: dict) -> dict:
    init = new_initiative()
    for f in CARRYOVER_FIELDS:
        init[f] = src.get(f, init[f])
    init["carry_over"] = True
    return init

def empty_draft() -> dict:
    return {
        "initiatives":      [],
        "status":           "in-progress",
        "entity":           "",
        "reporting_month":  "",   # filing month  — storage key & Group# in filename
        "activities_month": "",   # activities month — what period the R&D work took place in
    }


# ── Excel export ──────────────────────────────────────────────────────────────
# Colors from the original template (extracted via openpyxl):
#   GREEN  = 9BBB59  (theme accent3 — header row)
#   BEIGE  = EEECE1  (theme lt2   — input cells)
#   YELLOW = DED900  (rgb         — Team Members / Selection cells)
#   Font   = Arial Narrow, bold white 12pt on headers

GREEN  = "9BBB59"
BEIGE  = "EEECE1"
YELLOW = "DED900"
WHITE  = "FFFFFF"
DARK   = "1F1F1F"
GRAY   = "666666"

def _thin_border():
    s = Side(style="thin", color="A0A0A0")
    return Border(left=s, right=s, top=s, bottom=s)

def _hdr_font(sz=12):
    return Font(name="Arial Narrow", bold=True, color=WHITE, size=sz)

def _data_font():
    return Font(name="Arial Narrow", color=DARK, size=11)

# Column definition for a single-combo sheet (User prepended for consolidated)
_COL_DEF = [
    # Widths match original template exactly (measured via openpyxl from source file)
    ("Month/Yr",                                               "_month",             14.15,  False),
    ("Filing Month",                                            "_filing",            13.0,   False),
    ("Business Component",                                     "business_component", 35.0,   False),
    ("Initiative Name",                                        "initiative_name",    19.26,  False),
    ("Initiative Description",                                 "initiative_description", 38.15, False),
    ("Tech Uncertainty",                                       "tech_uncertainty",   64.41,  False),
    ("Start Date",                                             "start_date",         15.68,  False),
    ("Expected End Date",                                      "expected_end_date",  15.68,  False),
    ("Activities to Eliminate Technical Uncertainty",          "activities",         60.26,  False),
    ("Team Members",                                           "team_members",       49.0,   True),
    ("Notes",                                                  "notes",              52.57,  False),
    ("Completion Date",                                        "_completion",        22.0,   False),
    ("Status",                                                 "_status",            22.0,   False),
]

def _write_sheet(ws, rows_data: list[dict], subtitle: str):
    """
    Matches the exact structure of the original template:
      Row 1 : Title (A1) + Legend header cells (C-E or D-F for consolidated)
      Row 2 : Legend swatches (Selection=yellow, subtitle text)
      Row 3 : Autofill label
      Row 4 : Empty spacer
      Row 5 : Column headers (green, Arial Narrow 12pt bold white)
      Row 6+: Data rows (auto height, beige fill / yellow for Team Members)
    """
    border      = _thin_border()
    fill_green  = PatternFill("solid", fgColor=GREEN)
    fill_beige  = PatternFill("solid", fgColor=BEIGE)
    fill_yellow = PatternFill("solid", fgColor=YELLOW)
    fill_red    = PatternFill("solid", fgColor="FF0000")

    has_user_col = any("user" in r for r in rows_data) if rows_data else False
    cols = ([("User", "user", 18, False)] if has_user_col else []) + list(_COL_DEF)

    # Legend columns offset right by 1 for consolidated sheet (extra User col)
    lc = 3 if not has_user_col else 4   # start column for legend

    # ── Row 1: Title + Legend header ─────────────────────────────────────────
    ws.row_dimensions[1].height = 18.3
    c = ws.cell(1, 1, "Monthly R&D Tracking Template")
    c.font = Font(name="Arial Narrow", bold=True, size=14, color=DARK)
    c = ws.cell(1, lc, "Legend:")
    c.font = Font(name="Arial Narrow", bold=True, size=12, color=DARK)
    c.alignment = Alignment(horizontal="right", wrap_text=True)
    c = ws.cell(1, lc + 1, "Input")
    c.font = Font(name="Arial Narrow", size=12, color=DARK)
    c.fill = fill_beige
    c.alignment = Alignment(wrap_text=True)
    c = ws.cell(1, lc + 2, "Missing Info")
    c.font = Font(name="Arial Narrow", size=11, color=DARK)
    c.fill = fill_red

    # ── Row 2: Selection swatch + subtitle ───────────────────────────────────
    ws.row_dimensions[2].height = 15.3
    c = ws.cell(2, lc + 1, "Selection")
    c.font = Font(name="Arial Narrow", size=12, color=DARK)
    c.fill = fill_yellow
    c.alignment = Alignment(wrap_text=True)
    c = ws.cell(2, lc + 2, subtitle)
    c.font = Font(name="Arial Narrow", size=10, color=GRAY, italic=True)

    # ── Row 3: Autofill label ─────────────────────────────────────────────────
    ws.row_dimensions[3].height = 15.3
    c = ws.cell(3, lc + 1, "Autofill")
    c.font = Font(name="Arial Narrow", bold=True, size=12, color=DARK)
    c.alignment = Alignment(wrap_text=True)

    # ── Row 4: Empty spacer ───────────────────────────────────────────────────
    ws.row_dimensions[4].height = 15

    # ── Row 5: Column headers ─────────────────────────────────────────────────
    ws.row_dimensions[5].height = 30
    for ci, (hdr, _, width, _) in enumerate(cols, 1):
        cell = ws.cell(5, ci, hdr)
        cell.fill      = fill_green
        cell.font      = _hdr_font(12)
        cell.border    = border
        cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
        ws.column_dimensions[get_column_letter(ci)].width = width

    # ── Row 6+: Data rows ─────────────────────────────────────────────────────
    # No explicit height — Excel auto-fits to content (matches original template behavior)
    for rn, row in enumerate(rows_data, 6):
        for ci, (_, key, _, is_team) in enumerate(cols, 1):
            val = row.get(key, "")
            cell = ws.cell(rn, ci, val)
            cell.font      = _data_font()
            cell.border    = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.fill      = fill_yellow if is_team else fill_beige

    # Freeze header rows 1-5 (same as original)
    ws.freeze_panes = "A6"


def _sub_to_rows(username: str, sub: dict, include_user: bool) -> list[dict]:
    """Convert a submission dict into a list of row dicts for _write_sheet."""
    rm     = sub.get("reporting_month", "")
    status = STATUS_LABELS.get(sub.get("status", ""), sub.get("status", ""))
    rows   = []
    for init in sub.get("initiatives") or []:
        row = {
            # Use initiative's own month_yr, then activities_month, then filing month
            "_month":  fmt_month(init.get("month_yr") or sub.get("activities_month") or rm),
            "_filing": fmt_month(rm),
            "_status": status,
            "_completion": (
                datetime.fromtimestamp(init["approved_at"] / 1000).strftime("%Y-%m-%d %H:%M")
                if init.get("approved_at") else
                datetime.fromtimestamp(sub["approved_at"] / 1000).strftime("%Y-%m-%d")
                if sub.get("approved_at") else ""
            ),
            "business_component":    init.get("business_component",    ""),
            "initiative_name":       init.get("initiative_name",        ""),
            "initiative_description":init.get("initiative_description", ""),
            "tech_uncertainty":      init.get("tech_uncertainty",       ""),
            "start_date":            str(init.get("start_date")        or ""),
            "expected_end_date":     str(init.get("expected_end_date") or ""),
            "activities":            init.get("activities",             ""),
            "team_members":          ", ".join(init.get("team_members") or []),
            "notes":                 init.get("notes",                  ""),
        }
        if include_user:
            row["user"] = username
        rows.append(row)
    return rows


def build_excel_individual(username: str, sub: dict) -> bytes:
    """Single-user, single-month export — matches original template layout."""
    wb = Workbook()
    ws = wb.active
    am     = sub.get("activities_month") or sub.get("reporting_month", "")
    fm     = sub.get("reporting_month", "")
    entity = sub.get("entity", "")
    ws.title = fmt_month_tab(am) or fmt_month_tab(fm) or "Report"
    subtitle = (
        f"Submitted by: {username}  |  Entity: {entity}  |  "
        f"Filing: {fmt_month(fm)}  |  Reporting Period: {fmt_month(am)}"
    )
    rows = _sub_to_rows(username, sub, include_user=False)
    _write_sheet(ws, rows, subtitle)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_excel_consolidated(
    all_data: dict,
    selected_combos: list[tuple[str, str]],
    submitted_only: bool = False,
    status_filter: list[str] | None = None,
) -> bytes:
    """
    One sheet per entity (tab = "Entity 107", "Entity 108", etc.).
    All periods for that entity appear as rows within the sheet,
    sorted by reporting_month then username.

    status_filter: if provided, only include submissions whose status is in
    this list (e.g. ["approved","archived"]). Takes precedence over submitted_only.
    If both are omitted/None/False, all statuses are included.
    """
    wb = Workbook()
    wb.remove(wb.active)

    # Collect unique entities from the selected combos, in sorted order
    entities = sorted({e for e, _ in selected_combos})

    if status_filter is not None:
        allowed_statuses = set(status_filter)
        status_desc = ", ".join(STATUS_LABELS.get(s, s).split(" ", 1)[-1] for s in status_filter) or "None selected"
    elif submitted_only:
        allowed_statuses = {"submitted", "approved", "archived"}
        status_desc = "Submitted, Approved & Archived"
    else:
        allowed_statuses = None
        status_desc = "All statuses"

    for entity in entities:
        ws = wb.create_sheet(title=f"Entity {entity}"[:31])
        subtitle = f"Entity: {entity}  |  {status_desc}"

        # Gather rows for this entity across all selected periods, sorted by period then user
        entity_combos = sorted(
            [(e, rm) for e, rm in selected_combos if e == entity],
            key=lambda x: x[1],   # sort by reporting_month (YYYY-MM)
        )
        rows = []
        for _, rm in entity_combos:
            for username in sorted(all_data.keys()):
                sub = all_data[username].get(rm)
                if not sub or sub.get("entity") != entity:
                    continue
                if allowed_statuses is not None and sub.get("status") not in allowed_statuses:
                    continue
                rows.extend(_sub_to_rows(username, sub, include_user=True))

        _write_sheet(ws, rows, subtitle)

    if not wb.sheetnames:
        ws = wb.create_sheet("No Data")
        ws.cell(1, 1, "No matching submissions found.")

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── CSS ───────────────────────────────────────────────────────────────────────

def inject_css():
    st.markdown("""
    <style>
    #MainMenu, footer, header { visibility: hidden; }
    .stApp { background: #f1f5f9; }
    .badge-approved    { background:#f0fdf4; color:#166534; border:1.5px solid #86efac; }
    .badge-submitted   { background:#eff6ff; color:#1e40af; border:1.5px solid #93c5fd; }
    .badge-in-progress { background:#fef9ec; color:#92600a; border:1.5px solid #fcd34d; }
    .badge-rejected    { background:#fff1f2; color:#9f1239; border:1.5px solid #fda4af; }
    .badge-not-started { background:#f8fafc; color:#64748b; border:1.5px solid #cbd5e1; }
    .rd-badge { display:inline-block; padding:3px 12px; border-radius:20px; font-size:12px; font-weight:700; }
    .progress-bar-outer { background:#e2e8f0; border-radius:4px; height:8px; overflow:hidden; margin-bottom:6px; }
    .progress-bar-inner { background:#c86a2a; height:100%; border-radius:4px; transition:width 0.3s; }
    .wizard-question { font-size:22px; font-weight:700; color:#1a3c5e; margin-bottom:6px; }
    .wizard-hint     { font-size:14px; color:#64748b; margin-bottom:20px; }
    .step-label      { font-size:11px; font-weight:700; color:#c86a2a; text-transform:uppercase; letter-spacing:.8px; }
    .carryover-banner { background:#fffbeb; border:1.5px solid #fcd34d; border-radius:10px; padding:10px 16px; margin-bottom:14px; font-size:13px; color:#92600a; }
    .prefilled-banner { background:#f0f9ff; border:1.5px solid #93c5fd; border-radius:10px; padding:10px 16px; margin-bottom:14px; font-size:13px; color:#1e40af; }
    .delete-confirm   { background:#fff1f2; border:1.5px solid #fda4af; border-radius:10px; padding:12px 16px; margin-top:8px; }
    </style>
    """, unsafe_allow_html=True)

def badge_html(status: str) -> str:
    cls_map = {
        "approved":    "badge-approved",
        "submitted":   "badge-submitted",
        "in-progress": "badge-in-progress",
        "rejected":    "badge-rejected",
    }
    cls   = cls_map.get(status, "badge-not-started")
    label = STATUS_LABELS.get(status, status)
    return f'<span class="rd-badge {cls}">{label}</span>'

def init_session():
    defaults = {
        "screen":      "login",
        "user":        None,
        "is_admin":    False,
        "draft":       empty_draft(),
        "wiz_init":    None,
        "wiz_step":    0,
        "wiz_mode":    "new",
        "confirm_del": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


# ── Login ─────────────────────────────────────────────────────────────────────

def screen_login():
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown("""
        <div style="text-align:center; margin-bottom:28px;">
            <div style="font-size:48px; margin-bottom:10px;">🔬</div>
            <h1 style="color:#1a3c5e; font-size:28px; margin:0 0 6px;">R&D Activity Tracker</h1>
            <p style="color:#64748b; margin:0;">Monthly Reporting Portal</p>
        </div>
        """, unsafe_allow_html=True)
        options = ["— Select your name —", "⚙ Admin (Oversight Lead)"] + EMPLOYEES
        sel = st.selectbox("Sign in as", options, label_visibility="collapsed")
        if st.button("Sign In →", use_container_width=True, type="primary"):
            if sel.startswith("— "):
                st.warning("Please select your name.")
            else:
                name     = sel.replace("⚙ ", "").replace(" (Oversight Lead)", "").strip()
                is_admin = "Admin" in sel
                st.session_state.user     = name
                st.session_state.is_admin = is_admin
                st.session_state.draft    = empty_draft()
                st.session_state.screen   = "admin" if is_admin else "dashboard"
                st.rerun()
        st.markdown(
            '<p style="text-align:center;font-size:12px;color:#94a3b8;margin-top:16px;">'
            'Your entries are saved automatically as you go.</p>',
            unsafe_allow_html=True,
        )


# ── Dashboard ─────────────────────────────────────────────────────────────────

# ── Reminder banners ──────────────────────────────────────────────────────────
# Streamlit has no background scheduler — these checks run only when someone
# opens the app, and simply surface a visible nudge based on today's date.

def render_user_reminder(user: str):
    """
    If it's the 5th or later and the user hasn't started this month's
    filing cycle yet, show a reminder banner on login.
    """
    today = date.today()
    if today.day < 5:
        return
    cur = cur_month()
    existing = load_submission(user, cur)
    if existing and existing.get("initiatives"):
        return  # already has a real submission for this period
    st.warning(
        f"📅 It's past the 5th and you haven't started your **{fmt_month(cur)}** filing yet. "
        "Set up your report below to get started."
    )


def render_admin_reminder(all_data: dict):
    """
    Surfaces a count of reports still pending review when the admin logs in.
    """
    pending = []
    for username, months in all_data.items():
        for month, sub in months.items():
            if sub.get("status") == "submitted" and sub.get("initiatives"):
                pending.append((username, month))
    if pending:
        names = ", ".join(f"{u} ({fmt_month(m)})" for u, m in pending[:5])
        extra = f" +{len(pending)-5} more" if len(pending) > 5 else ""
        st.warning(f"🔔 {len(pending)} report(s) pending your review: {names}{extra}")


def screen_dashboard():
    user      = st.session_state.user
    draft     = st.session_state.draft
    # "rejected" is treated the same as "in-progress" for editing/submitting
    # — user needs to be able to fix and resubmit after a rejection.
    # Locked states: submitted, approved, archived.
    # Rejected reports are also locked UNTIL the user acknowledges the rejection
    # (matching PA Flow 3 — user must explicitly confirm they've seen the feedback
    # before the edit/resubmit cycle opens up).
    ack_key  = f"ack_rejection_{draft.get('reporting_month')}"
    rejected_unacked = draft["status"] == "rejected" and not st.session_state.get(ack_key)
    submitted = draft["status"] in ("submitted", "approved", "archived") or rejected_unacked

    # Header
    c1, c2 = st.columns([5, 1])
    with c1:
        fm  = draft.get("reporting_month", "")
        am  = draft.get("activities_month", "")
        st.markdown("## 🔬 R&D Tracker")
        if fm and am:
            st.markdown(
                f"**Filing:** {fmt_month(fm)} &nbsp;|&nbsp; "
                f"**Reporting Period:** {fmt_month(am)}",
                unsafe_allow_html=True,
            )
        elif fm:
            st.markdown(f"**Filing Month:** {fmt_month(fm)}", unsafe_allow_html=True)
        else:
            st.markdown(
                "<span style='color:#c86a2a;'>⚠ Complete Report Setup below</span>",
                unsafe_allow_html=True,
            )
        st.caption(f"Signed in as **{user}**")
    with c2:
        st.write("")
        if st.button("Sign Out"):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.rerun()

    st.divider()
    render_user_reminder(user)

    # ── Report Setup ─────────────────────────────────────────────────────────
    setup_complete = bool(draft.get("entity") and draft.get("reporting_month"))
    filing_lbl     = fmt_month(draft.get("reporting_month",""))
    act_lbl        = fmt_month(draft.get("activities_month",""))
    expander_label = (
        f"⚙ Report Setup  —  Entity {draft['entity']}  ·  Filing: {filing_lbl}  ·  Reporting Period: {act_lbl}"
        if setup_complete else "⚙ Report Setup  ⚠ Please complete setup first"
    )
    with st.expander(expander_label, expanded=not setup_complete):
        st.caption(
            "**Filing Month** is the month you are submitting this report in "
            "— it determines the Group number in the filename (e.g. filing in June → Group 6). "
            "**Reporting Period** is the month the R&D activities actually took place "
"— this is the period that shows in the export (typically the previous month)."
        )
        su1, su2, su3 = st.columns(3)
        mlist = available_months()

        with su1:
            entity_options = all_entities() + ["+ Add new entity..."]
            eidx = entity_options.index(draft["entity"]) if draft.get("entity") in entity_options else 0
            entity_pick = st.selectbox("Entity", entity_options, index=eidx, key="su_entity")

            if entity_pick == "+ Add new entity...":
                new_e_col1, new_e_col2 = st.columns([3, 1])
                with new_e_col1:
                    new_entity_input = st.text_input(
                        "New entity number", key="su_new_entity",
                        placeholder="e.g. 111", label_visibility="collapsed",
                    )
                with new_e_col2:
                    if st.button("Add", key="su_add_entity"):
                        if new_entity_input.strip():
                            added = add_custom_entity(new_entity_input.strip())
                            st.session_state.su_entity = new_entity_input.strip()
                            if added:
                                st.success(f"Entity {new_entity_input.strip()} added — it will stay in the dropdown.")
                            st.rerun()
                # Keep previous entity selected until a new one is actually added
                chosen_entity = draft.get("entity") or (all_entities()[0] if all_entities() else "")
            else:
                chosen_entity = entity_pick

        with su2:
            # Filing Month — defaults to CURRENT month (you are filing now)
            def_filing = mlist[0]
            cur_filing  = draft.get("reporting_month") or def_filing
            filing_idx  = mlist.index(cur_filing) if cur_filing in mlist else 0
            chosen_filing = st.selectbox(
                "Filing Month",
                mlist,
                index=filing_idx,
                format_func=fmt_month,
                key="su_filing",
                help="The month you are submitting this report in. Sets the Group number in the filename.",
            )

        with su3:
            filing_changed_now = chosen_filing != draft.get("reporting_month", "")
            if filing_changed_now:
                # Auto-derive from new filing month
                cur_act = prev_month_of(chosen_filing)
            else:
                # Filing month unchanged — respect whatever the user has chosen
                cur_act = draft.get("activities_month") or prev_month_of(chosen_filing)
            act_idx = mlist.index(cur_act) if cur_act in mlist else (1 if len(mlist) > 1 else 0)
            # Key includes the filing month — when filing changes, Streamlit treats
            # this as a completely new widget and renders from index, not cached state.
            chosen_act = st.selectbox(
                "Reporting Period",
                mlist,
                index=act_idx,
                format_func=fmt_month,
                key=f"su_act_{chosen_filing}",
                help="The month R&D activities took place.",
            )

        entity_changed = chosen_entity != draft.get("entity")
        filing_changed = chosen_filing != draft.get("reporting_month")
        act_changed    = chosen_act    != draft.get("activities_month")

        if entity_changed or filing_changed or act_changed:
            if filing_changed:
                existing = load_submission(user, chosen_filing)
                draft    = existing if existing else empty_draft()
            draft["entity"]           = chosen_entity
            draft["reporting_month"]  = chosen_filing
            draft["activities_month"] = chosen_act
            save_draft(user, draft)
            st.session_state.draft = draft
            st.rerun()

        if setup_complete:
            fname = export_filename(draft["entity"], draft["reporting_month"])
            st.success(
                f"✓ **Filing Month:** {fmt_month(draft['reporting_month'])} "
                f"(Entity {draft['entity']}, filename: **{fname}**)  |  "
                f"**Reporting Period:** {fmt_month(draft.get('activities_month',''))}"
            )

    if not setup_complete:
        st.info("Complete the Report Setup above before adding initiatives.")
        return

    st.divider()

    # ── Carry-over from any past period ─────────────────────────────────────
    # Searches ALL past months for this user (not just previous month)
    # so nothing falls through the cracks.
    past_inits = get_past_ongoing_initiatives(user, draft["reporting_month"])

    # Filter out initiatives already in the current draft (by name)
    current_names = {i.get("initiative_name","").strip().lower() for i in draft["initiatives"]}
    past_inits    = [i for i in past_inits if i.get("initiative_name","").strip().lower() not in current_names]

    if not submitted and past_inits:
        st.markdown(f"""
        <div class="carryover-banner">
            <strong>📋 Ongoing initiatives from past periods</strong><br>
            {len(past_inits)} active initiative{"s" if len(past_inits)!=1 else ""} found across your history.
            Carry any forward — key details pre-fill; you only update activities and notes.
        </div>
        """, unsafe_allow_html=True)
        cols = st.columns(min(len(past_inits), 3))
        for idx, pi in enumerate(past_inits):
            with cols[idx % 3]:
                if st.button(f"↩ {pi.get('initiative_name','Unnamed')}", key=f"co_{pi['id']}"):
                    st.session_state.wiz_init = carryover_initiative(pi)
                    st.session_state.wiz_step = 0
                    st.session_state.wiz_mode = "carryover"
                    st.session_state.screen   = "wizard"
                    st.rerun()

    st.divider()

    # ── Status strip ─────────────────────────────────────────────────────────
    c1, c2, c3 = st.columns([4, 1, 1])
    with c1:
        st.markdown(f"**Report Status:** {badge_html(draft['status'])}", unsafe_allow_html=True)
        if draft.get("submitted_at"):
            ts = datetime.fromtimestamp(draft["submitted_at"]/1000).strftime("%b %d, %Y %H:%M")
            st.caption(f"Submitted {ts}")
        if draft["status"] == "rejected":
            rc = draft.get("rejection_comment", "")
            msg = "Your report was Rejected by the Oversight Lead."
            if rc:
                msg += f' Comment: "{rc}".'
            st.error(msg)
            # Show acknowledgement button matching PA Flow 3 — user must confirm
            # they've seen the rejection before the resubmit cycle can restart.
            if not st.session_state.get(f"ack_rejection_{draft.get('reporting_month')}"):
                st.warning(
                    "Please review the feedback above, update your initiatives if needed, "
                    "then click **Acknowledge & Prepare Resubmission** to confirm you've seen "
                    "the rejection and are ready to update."
                )
                if st.button("✓ Acknowledge & Prepare Resubmission", key="ack_rejection_btn"):
                    st.session_state[f"ack_rejection_{draft.get('reporting_month')}"] = True
                    st.rerun()
            else:
                st.info("✓ Rejection acknowledged. Update your initiatives below, then resubmit.")
        elif draft["status"] == "archived":
            arc_ts = draft.get("archived_at")
            arc_str = f" on {datetime.fromtimestamp(arc_ts/1000).strftime('%b %d, %Y')}" if arc_ts else ""
            st.info(f"📦 This period was archived{arc_str} and is now locked. Contact the Oversight Lead to reopen it.")
    with c2:
        st.metric("Initiatives", len(draft["initiatives"]))
    with c3:
        if st.session_state.is_admin and st.button("Admin View →"):
            st.session_state.screen = "admin"
            st.rerun()

    # ── Initiatives list ──────────────────────────────────────────────────────
    c1, c2 = st.columns([5, 1])
    with c1:
        act_m = draft.get("activities_month") or draft.get("reporting_month","")
        st.subheader(f"Initiatives — {fmt_month(act_m)}")
        st.caption(
            f"R&D activities that took place in **{fmt_month(act_m)}**. "
            "This is the period that will appear in the export."
        )
    with c2:
        if not submitted and st.button("＋ Add Initiative", type="primary"):
            st.session_state.wiz_init = new_initiative()
            st.session_state.wiz_step = 0
            st.session_state.wiz_mode = "new"
            st.session_state.screen   = "wizard"
            st.rerun()

    if not draft["initiatives"]:
        st.info("No initiatives added yet. Click **＋ Add Initiative** to begin this month's report.")
    else:
        # ── Quick-scan summary table ────────────────────────────────────────
        # Lets you see everything at a glance before opening any single card —
        # this is what matters once there are more than a couple initiatives.
        status_icon_map = {"approved": "✅", "returned": "🔴", "active": "🔵"}
        table_rows = []
        for init in draft["initiatives"]:
            istatus = init.get("initiative_status", "active")
            table_rows.append({
                " ":                  status_icon_map.get(istatus, "🔵"),
                "Initiative":         init.get("initiative_name", "Unnamed"),
                "Business Component": init.get("business_component", ""),
                "Team":               ", ".join(init.get("team_members") or []),
                "Start":              init.get("start_date", "—"),
                "End":                init.get("expected_end_date", "—"),
            })
        st.dataframe(
            table_rows,
            use_container_width=True,
            hide_index=True,
            column_config={" ": st.column_config.TextColumn(width="small")},
        )
        st.caption("✅ Approved by admin   ·   🔴 Rejected — needs revision   ·   🔵 Active / pending review")

        st.write("")

        for init in draft["initiatives"]:
            iid = init["id"]
            istatus = init.get("initiative_status", "active")
            status_icon = status_icon_map.get(istatus, "🔵")
            with st.expander(
                f"{status_icon} {'↩ ' if init.get('carry_over') else ''}"
                f"{init.get('initiative_name','Unnamed')} — {init.get('business_component','')}",
                expanded=False,
            ):
                st.markdown(f"**{init.get('initiative_description','—')}**")
                st.caption(
                    f"📅 {init.get('start_date','—')} → {init.get('expected_end_date','—')}  "
                    f"  👥 {', '.join(init.get('team_members') or ['—'])}"
                )
                if init.get("tech_uncertainty"):
                    st.markdown("**Technical Uncertainty:**")
                    st.markdown(
                        f'<div style="background:#f8fafc;padding:10px 14px;border-radius:6px;'
                        f'font-size:13px;color:#334155;">{init["tech_uncertainty"]}</div>',
                        unsafe_allow_html=True,
                    )
                if init.get("activities"):
                    st.markdown("**Activities to eliminate uncertainty:**")
                    st.caption(init["activities"])
                if init.get("notes"):
                    st.markdown(f"*Notes: {init['notes']}*")

                # Per-initiative status banner (set by admin)
                if istatus == "returned":
                    ret_ts = init.get("returned_at")
                    ret_str = f" on {datetime.fromtimestamp(ret_ts/1000).strftime('%b %d %H:%M')}" if ret_ts else ""
                    rc = init.get("rejection_comment", "")
                    rc_str = f' Comment: "{rc}".' if rc else ""
                    st.warning(f"⚠ This initiative was Rejected{ret_str}.{rc_str} Edit and resubmit.")
                elif istatus == "approved":
                    appr_ts = init.get("approved_at")
                    appr_str = f" on {datetime.fromtimestamp(appr_ts/1000).strftime('%b %d %H:%M')}" if appr_ts else ""
                    st.success(f"✓ Approved by admin{appr_str}.")

                is_archived = draft["status"] == "archived"
                st.write("")
                c1, c2, c3, c4 = st.columns([2.5, 0.8, 0.8, 0.8])
                with c2:
                    if st.button("✏ Edit", key=f"edit_{iid}", disabled=is_archived,
                                 help="This period is archived and locked." if is_archived else None):
                        st.session_state.wiz_init = dict(init)
                        st.session_state.wiz_step = 0
                        st.session_state.wiz_mode = "edit"
                        st.session_state.screen   = "wizard"
                        st.rerun()
                with c3:
                    if st.button("📤 Submit for Review", key=f"usub_{iid}", disabled=submitted,
                                 help="Marks your full report as Ready for Review"):
                        draft["status"]       = "submitted"
                        draft["submitted_at"] = int(datetime.now().timestamp()*1000)
                        save_draft(user, draft)
                        st.session_state.draft = draft
                        st.success("Report submitted!")
                        st.rerun()
                with c4:
                    if st.button("🗑 Delete", key=f"del_{iid}", disabled=is_archived,
                                 help="This period is archived and locked." if is_archived else None):
                        st.session_state.confirm_del = iid
                        st.rerun()

                if st.session_state.get("confirm_del") == iid:
                    st.markdown('<div class="delete-confirm">', unsafe_allow_html=True)
                    st.warning(
                        f"Delete **{init.get('initiative_name','this initiative')}**? "
                        "This cannot be undone."
                    )
                    ca, cb = st.columns(2)
                    with ca:
                        if st.button("Yes, delete it", key=f"conf_yes_{iid}", type="primary"):
                            draft["initiatives"] = [i for i in draft["initiatives"] if i["id"] != iid]
                            if submitted:
                                draft["status"] = "in-progress"
                                draft.pop("submitted_at", None)
                            save_draft(user, draft)
                            st.session_state.draft       = draft
                            st.session_state.confirm_del = None
                            st.success("Initiative deleted.")
                            st.rerun()
                    with cb:
                        if st.button("Cancel", key=f"conf_no_{iid}"):
                            st.session_state.confirm_del = None
                            st.rerun()
                    st.markdown('</div>', unsafe_allow_html=True)

    # ── Submit + Export ───────────────────────────────────────────────────────
    st.divider()
    c1, c2 = st.columns([3, 1])
    with c1:
        if not submitted and draft["initiatives"]:
            was_rejected = draft["status"] == "rejected"
            label = "Resubmit — Ready for Review ✓" if was_rejected else "Submit — Ready for Review ✓"
            caption = (
                "Resubmitting will send your updated report back to the Oversight Lead for review."
                if was_rejected else
                f"Sends your {len(draft['initiatives'])} initiative"
                f"{'s' if len(draft['initiatives'])!=1 else ''} to the Oversight Lead for review."
            )
            st.markdown("#### Ready to submit?" if not was_rejected else "#### Ready to resubmit?")
            st.caption(caption)
            if st.button(label, type="primary"):
                draft["status"]       = "submitted"
                draft["submitted_at"] = int(datetime.now().timestamp()*1000)
                draft.pop("rejection_comment", None)   # clear old rejection comment
                # Clear the acknowledgement flag since the cycle is now restarted
                st.session_state.pop(f"ack_rejection_{draft.get('reporting_month')}", None)
                save_draft(user, draft)
                st.session_state.draft = draft
                st.success("Report resubmitted!" if was_rejected else "Report submitted!")
                st.rerun()
    with c2:
        xlsx = build_excel_individual(user, draft)
        fname = export_filename(draft.get("entity",""), draft.get("reporting_month",""))
        st.download_button(
            "↓ Download My Report",
            data=xlsx,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # ── History ──────────────────────────────────────────────────────────────
    st.write("")
    render_history_section(user, draft.get("reporting_month",""))


# ── History (shown below submit on the dashboard) ───────────────────────────

def render_history_section(user: str, current_reporting_month: str):
    """Renders a collapsible history of all past submissions for this user."""
    history = get_user_history(user)
    # Exclude current period from history display
    history = [(m, s) for m, s in history if m != current_reporting_month]

    if not history:
        return

    with st.expander(f"📁 My Submission History ({len(history)} past period{'s' if len(history)!=1 else ''})", expanded=False):
        for month, sub in history:
            entity  = sub.get("entity", "—")
            status  = sub.get("status", "not-started")
            inits   = sub.get("initiatives") or []
            st.markdown(
                f"**{entity} — {fmt_month(month)}** &nbsp; {badge_html(status)} &nbsp; "
                f"*{len(inits)} initiative{'s' if len(inits)!=1 else ''}*",
                unsafe_allow_html=True,
            )
            if inits:
                for i in inits:
                    st.markdown(
                        f"&nbsp;&nbsp;&nbsp;• **{i.get('initiative_name','Unnamed')}** — "
                        f"{i.get('business_component','')}  "
                        f"📅 {i.get('start_date','—')} → {i.get('expected_end_date','—')}  "
                        f"👥 {', '.join(i.get('team_members') or ['—'])}",
                        unsafe_allow_html=True,
                    )
                    if i.get("tech_uncertainty"):
                        st.caption(f"   Uncertainty: {i['tech_uncertainty'][:120]}{'…' if len(i.get('tech_uncertainty',''))>120 else ''}")
            # Export button for this past period
            if sub.get("entity") and sub.get("reporting_month"):
                h_xlsx  = build_excel_individual(user, sub)
                h_fname = export_filename(sub["entity"], sub["reporting_month"])
                st.download_button(
                    f"↓ {entity} — {fmt_month(month)}",
                    data=h_xlsx,
                    file_name=h_fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"hist_dl_{month}",
                )
            st.write("")


# ── Wizard ────────────────────────────────────────────────────────────────────

def screen_wizard():
    user  = st.session_state.user
    init  = st.session_state.wiz_init
    step  = st.session_state.wiz_step
    mode  = st.session_state.wiz_mode
    s     = WIZARD_STEPS[step]
    total = len(WIZARD_STEPS)
    pct   = int((step+1)/total*100)

    is_carryover = mode == "carryover"
    is_prefilled = is_carryover and s["field"] in CARRYOVER_FIELDS

    c1, c2 = st.columns([5, 1])
    with c1:
        title = {"new":"New Initiative","edit":"Edit Initiative","carryover":"Carry Over Initiative"}[mode]
        st.markdown(f"### {title}")
    with c2:
        if st.button("✕ Cancel"):
            st.session_state.screen = "dashboard"
            st.rerun()

    st.markdown(f"""
    <div class="progress-bar-outer">
        <div class="progress-bar-inner" style="width:{pct}%"></div>
    </div>
    <p style="font-size:12px;color:#64748b;margin-top:2px;">
        Step {step+1} of {total} &nbsp;·&nbsp; {pct}% complete
    </p>
    """, unsafe_allow_html=True)

    if is_carryover and not is_prefilled:
        st.markdown(f"""<div class="carryover-banner">
            <strong>Updating for {fmt_month(st.session_state.draft.get('reporting_month',''))}:</strong>
            {init.get("initiative_name","this initiative")}
        </div>""", unsafe_allow_html=True)
    elif is_prefilled:
        st.markdown("""<div class="prefilled-banner">
            Pre-filled from last month — confirm or update before continuing.
        </div>""", unsafe_allow_html=True)

    st.markdown(f'<p class="step-label">Question {step+1} of {total}</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="wizard-question">{s["question"]}</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="wizard-hint">{s["hint"]}</p>', unsafe_allow_html=True)

    field   = s["field"]
    val     = init.get(field)
    new_val = val

    if s["type"] == "text":
        new_val = st.text_input(s["label"], value=val or "", placeholder=s.get("placeholder",""))
    elif s["type"] == "textarea":
        new_val = st.text_area(s["label"], value=val or "", placeholder=s.get("placeholder",""), height=140)
    elif s["type"] == "date":
        parsed = None
        if val:
            try:
                parsed = datetime.strptime(str(val), "%Y-%m-%d").date()
            except Exception:
                pass
        picked  = st.date_input(s["label"], value=parsed)
        new_val = picked.strftime("%Y-%m-%d") if picked else None
    elif s["type"] == "multiselect":
        new_val = st.multiselect(s["label"], EMPLOYEES, default=val or [])

    init[field] = new_val

    if s["required"]:
        is_valid = bool(new_val) if s["type"] == "multiselect" else bool(str(new_val or "").strip())
    else:
        is_valid = True

    c1, c2, c3 = st.columns([1, 4, 1])
    with c1:
        if st.button("← Back", disabled=(step == 0)):
            st.session_state.wiz_step -= 1
            st.rerun()
    with c3:
        is_last = step == total - 1
        if st.button("Save Initiative ✓" if is_last else "Next →", disabled=not is_valid, type="primary"):
            if is_last:
                draft = st.session_state.draft
                # Stamp month_yr on new/carryover initiatives so the export
                # always shows the correct period regardless of later rollovers.
                # Edit mode preserves whatever month_yr was already on the initiative.
                if mode in ("new", "carryover") and not init.get("month_yr"):
                    # Use activities_month (the R&D period) not filing month
                    init["month_yr"] = draft.get("activities_month") or draft.get("reporting_month", "")
                if mode == "edit":
                    draft["initiatives"] = [i if i["id"] != init["id"] else init for i in draft["initiatives"]]
                else:
                    draft["initiatives"].append(init)
                save_draft(user, draft)
                st.session_state.draft  = draft
                st.session_state.screen = "dashboard"
            else:
                st.session_state.wiz_init = init
                st.session_state.wiz_step += 1
            st.rerun()

    dots = "".join(
        f'<span style="display:inline-block;width:{"24px" if i==step else "8px"};height:8px;'
        f'border-radius:4px;margin:0 3px;'
        f'background:{"#c86a2a" if i<step else "#1a3c5e" if i==step else "#cbd5e1"}"></span>'
        for i in range(total)
    )
    st.markdown(f'<div style="text-align:center;margin-top:20px;">{dots}</div>', unsafe_allow_html=True)


# ── Admin ─────────────────────────────────────────────────────────────────────

def get_latest_period_per_entity(combos: list[tuple[str, str]]) -> dict:
    """
    Given all (entity, period) combos, returns the most recent period
    for each entity — used to suggest a one-click 'roll forward' target.
    """
    latest: dict[str, str] = {}
    for entity, rm in combos:
        if entity not in latest or rm > latest[entity]:
            latest[entity] = rm
    return latest


def screen_admin():
    st.markdown("## ⚙ Admin Dashboard")

    c1, c2 = st.columns([5, 1])
    with c2:
        if st.button("← Back"):
            st.session_state.screen = "login" if st.session_state.user == "Admin" else "dashboard"
            st.rerun()

    # Silently repair any missing index entries before loading data —
    # ensures files that exist on disk are always included in the view.
    repaired = repair_months_index()
    if repaired:
        st.info(f"🔧 Index repaired: {repaired} missing entr{'ies' if repaired!=1 else 'y'} restored from disk.")

    # Load everything, grouped by (entity, reporting_month)
    all_data = load_all_submissions()
    combos   = get_combos(all_data)   # [(entity, rm), ...]

    render_admin_reminder(all_data)

    # ── Summary stats (always visible, above the tabs) ───────────────────────
    all_subs  = [
        sub for ud in all_data.values() for sub in ud.values()
        if sub.get("initiatives")
    ]
    n_users   = sum(1 for ud in all_data.values() if any(s.get("initiatives") for s in ud.values()))
    n_sub     = sum(1 for s in all_subs if s.get("status") in ("submitted","approved"))
    n_appr    = sum(1 for s in all_subs if s.get("status") == "approved")
    n_arch    = sum(1 for s in all_subs if s.get("status") == "archived")
    n_inits   = sum(len(s.get("initiatives") or []) for s in all_subs)

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Team Members",     n_users)
    c2.metric("Periods/Entities", len(combos))
    c3.metric("Submitted",        n_sub)
    c4.metric("Approved",         n_appr)
    c5.metric("Archived",         n_arch)
    c6.metric("Total Initiatives",n_inits)

    st.write("")

    tab_submissions, tab_rollover, tab_export, tab_backup = st.tabs(
        ["📁 Submissions", "🔄 Rollover", "⬇ Export", "🗄 Backup & Data"]
    )

    # ═══════════════════════════════════════════════════════════════════════
    # TAB: Submissions
    # ═══════════════════════════════════════════════════════════════════════
    with tab_submissions:
        if not combos:
            st.info("No submissions found yet.")
        else:
            st.caption(
                "**Filing Month** = the month the report was submitted in (sets the Group # in the filename). "
                "Each report also has a separate **Reporting Period** — the month the R&D activities took place "
                "in — shown inside each report below."
            )
            # ── Filters ──────────────────────────────────────────────────────
            fc1, fc2, fc3 = st.columns([1.2, 1.5, 1.5])
            entities_present = sorted({e for e, _ in combos})
            periods_present  = sorted({rm for _, rm in combos}, reverse=True)

            with fc1:
                f_entity = st.selectbox(
                    "Entity", ["All"] + entities_present, key="f_entity"
                )
            with fc2:
                f_period = st.selectbox(
                    "Filing Month", ["All"] + [fmt_month(p) for p in periods_present], key="f_period"
                )
            with fc3:
                f_status = st.selectbox(
                    "Status",
                    ["All", "Ready for Review", "Approved", "Archived", "Rejected", "In Progress"],
                    key="f_status",
                )

            status_map = {
                "Ready for Review": "submitted", "Approved": "approved",
                "Archived": "archived", "Rejected": "rejected", "In Progress": "in-progress",
            }

            filtered_combos = [
                (e, rm) for e, rm in combos
                if (f_entity == "All" or e == f_entity)
                and (f_period == "All" or fmt_month(rm) == f_period)
            ]

            if not filtered_combos:
                st.caption("No periods match the selected filters.")

            for (entity, rm) in filtered_combos:
                combo_rows = []
                for username, months in all_data.items():
                    sub = months.get(rm)
                    if not sub or sub.get("entity") != entity:
                        continue
                    status = sub.get("status", "not-started")
                    inits  = sub.get("initiatives") or []
                    if not inits and status == "in-progress":
                        continue
                    if f_status != "All" and status != status_map.get(f_status):
                        continue
                    combo_rows.append((username, sub))

                if not combo_rows:
                    continue

                # Show both Filing Month and the typical Reporting Period for this group
                sample_am = next(
                    (s.get("activities_month") for _, s in combo_rows if s.get("activities_month")),
                    None,
                )
                am_suffix = f"  ·  Reporting Period: **{fmt_month(sample_am)}**" if sample_am else ""
                st.markdown(f"#### Entity {entity} — Filing Month: **{fmt_month(rm)}**{am_suffix}")

                for username, sub in combo_rows:
                    status = sub.get("status", "not-started")
                    inits  = sub.get("initiatives") or []
                    icon   = {"approved":"✅","submitted":"🔵","in-progress":"🟡",
                              "rejected":"🔴","archived":"📦","not-started":"⚪"}.get(status,"⚪")

                    with st.expander(
                        f"{icon}  {username}   —  {STATUS_LABELS.get(status,'—')}  "
                        f"({len(inits)} initiative{'s' if len(inits)!=1 else ''})",
                        expanded=(status == "submitted"),
                    ):
                        # ── Report-level info + export ────────────────────────
                        h1, h2 = st.columns([4, 1])
                        with h1:
                            rep_period = sub.get("activities_month", "")
                            st.markdown(
                                f"**Filing Month:** {fmt_month(rm)} &nbsp;|&nbsp; "
                                f"**Reporting Period:** {fmt_month(rep_period) or '—'}",
                                unsafe_allow_html=True,
                            )
                            if sub.get("submitted_at"):
                                ts = datetime.fromtimestamp(sub["submitted_at"]/1000).strftime("%b %d %H:%M")
                                st.caption(f"Submitted: {ts}")
                            if sub.get("approved_at"):
                                ts = datetime.fromtimestamp(sub["approved_at"]/1000).strftime("%b %d %H:%M")
                                st.caption(f"Report approved: {ts}")
                        with h2:
                            u_xlsx  = build_excel_individual(username, sub)
                            u_fname = export_filename(entity, rm).replace(".xlsx", f"_{_safe_name(username)}.xlsx")
                            st.download_button(
                                "↓ Export",
                                data=u_xlsx,
                                file_name=u_fname,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"dl_{entity}_{rm}_{username}",
                            )

                        # ── Report-level approve / return ─────────────────────
                        if status == "submitted":
                            ra1, ra2, ra3 = st.columns([3, 1, 1])
                            with ra2:
                                if st.button("✓ Approve All", key=f"appr_rpt_{entity}_{rm}_{username}", type="primary"):
                                    now = int(datetime.now().timestamp()*1000)
                                    sub["status"]      = "approved"
                                    sub["approved_at"] = now
                                    for i in sub.get("initiatives") or []:
                                        if i.get("initiative_status","active") == "active":
                                            i["initiative_status"] = "approved"
                                            i["approved_at"]       = now
                                    save_draft(username, sub)
                                    st.success(f"Report approved!")
                                    st.rerun()
                            with ra3:
                                if st.button("✕ Reject All", key=f"rej_rpt_{entity}_{rm}_{username}"):
                                    st.session_state[f"reject_report_{entity}_{rm}_{username}"] = True
                                    st.rerun()

                        # ── Archive (only available once approved) ────────────
                        if status == "approved":
                            arc1, arc2 = st.columns([4, 1])
                            with arc1:
                                st.caption("This report is approved and ready to be closed out for the period.")
                            with arc2:
                                if st.button("📦 Archive", key=f"archive_{entity}_{rm}_{username}"):
                                    now = int(datetime.now().timestamp()*1000)
                                    sub["status"]      = "archived"
                                    sub["archived_at"] = now
                                    save_draft(username, sub)
                                    st.success("Report archived.")
                                    st.rerun()
                        elif status == "archived":
                            arc_ts = sub.get("archived_at")
                            arc_str = datetime.fromtimestamp(arc_ts/1000).strftime("%b %d, %Y %H:%M") if arc_ts else ""
                            st.info(f"📦 Archived{f' on {arc_str}' if arc_str else ''}. Remains in all exports.")

                        if st.session_state.get(f"reject_report_{entity}_{rm}_{username}"):
                            reject_comment = st.text_area(
                                "Rejection comment (will be shown to the user):",
                                key=f"rc_rpt_{entity}_{rm}_{username}",
                                placeholder="Explain what needs to be updated...",
                                height=80,
                            )
                            rc1, rc2 = st.columns(2)
                            with rc1:
                                if st.button("Confirm Reject", key=f"rc_confirm_rpt_{entity}_{rm}_{username}", type="primary"):
                                    now = int(datetime.now().timestamp()*1000)
                                    sub["status"]            = "rejected"
                                    sub["rejected_at"]       = now
                                    sub["rejection_comment"] = reject_comment.strip()
                                    st.session_state.pop(f"reject_report_{entity}_{rm}_{username}", None)
                                    save_draft(username, sub)
                                    st.rerun()
                            with rc2:
                                if st.button("Cancel", key=f"rc_cancel_rpt_{entity}_{rm}_{username}"):
                                    st.session_state.pop(f"reject_report_{entity}_{rm}_{username}", None)
                                    st.rerun()

                        st.divider()

                        if not inits:
                            st.caption("No initiatives.")
                        elif status == "in-progress":
                            # Don't allow accepting/rejecting until the user has
                            # reviewed the rollover and submitted for review themselves.
                            st.info(
                                "🟡 This report is **In Progress** — the user hasn't submitted it yet. "
                                "They need to log in, review their initiatives, and click "
                                "**Submit for Review** before you can approve or reject."
                            )
                            for init in inits:
                                iid     = init["id"]
                                istatus = init.get("initiative_status","active")
                                iname   = init.get("initiative_name","Unnamed")
                                ico     = {"approved":"✅","returned":"🔴","active":"🔵"}.get(istatus,"🔵")
                                ic1, ic4 = st.columns([5, 0.8])
                                with ic1:
                                    co   = "↩ " if init.get("carry_over") else ""
                                    bc   = init.get("business_component","")
                                    team = ", ".join(init.get("team_members") or ["—"])
                                    sd   = init.get("start_date","—")
                                    ed   = init.get("expected_end_date","—")
                                    st.markdown(
                                        f"{ico} **{co}{iname}** — {bc}  \n"
                                        f"<small>👥 {team} &nbsp;|&nbsp; 📅 {sd} → {ed}</small>",
                                        unsafe_allow_html=True,
                                    )
                                with ic4:
                                    if st.button("🗑 Delete", key=f"del_i_{entity}_{rm}_{username}_{iid}"):
                                        sub["initiatives"] = [i for i in sub["initiatives"] if i["id"] != iid]
                                        if not sub["initiatives"]:
                                            sub["status"] = "in-progress"
                                        save_draft(username, sub)
                                        st.success(f"Deleted: {iname}")
                                        st.rerun()
                        else:
                            for init in inits:
                                iid     = init["id"]
                                istatus = init.get("initiative_status","active")
                                iname   = init.get("initiative_name","Unnamed")
                                ico     = {"approved":"✅","returned":"🔴","active":"🔵"}.get(istatus,"🔵")

                                ic1, ic2, ic3, ic4 = st.columns([3, 0.8, 0.8, 0.8])
                                with ic1:
                                    co   = "↩ " if init.get("carry_over") else ""
                                    bc   = init.get("business_component","")
                                    team = ", ".join(init.get("team_members") or ["—"])
                                    sd   = init.get("start_date","—")
                                    ed   = init.get("expected_end_date","—")
                                    st.markdown(
                                        f"{ico} **{co}{iname}** — {bc}  \n"
                                        f"<small>👥 {team} &nbsp;|&nbsp; 📅 {sd} → {ed}</small>",
                                        unsafe_allow_html=True,
                                    )
                                    if init.get("approved_at"):
                                        ts = datetime.fromtimestamp(init["approved_at"]/1000).strftime("%b %d %H:%M")
                                        st.caption(f"   Approved {ts}")
                                    if init.get("returned_at"):
                                        ts = datetime.fromtimestamp(init["returned_at"]/1000).strftime("%b %d %H:%M")
                                        rc = init.get("rejection_comment","")
                                        rc_str = f' — "{rc}"' if rc else ""
                                        st.caption(f"   Rejected {ts}{rc_str}")

                                with ic2:
                                    if istatus != "approved":
                                        if st.button("✓ Accept", key=f"appr_i_{entity}_{rm}_{username}_{iid}",
                                                     type="primary"):
                                            now = int(datetime.now().timestamp()*1000)
                                            init["initiative_status"] = "approved"
                                            init["approved_at"]       = now
                                            init.pop("returned_at", None)
                                            all_approved = all(
                                                i.get("initiative_status") == "approved"
                                                for i in sub.get("initiatives",[])
                                            )
                                            if all_approved:
                                                sub["status"]      = "approved"
                                                sub["approved_at"] = now
                                            save_draft(username, sub)
                                            st.rerun()

                                with ic3:
                                    if istatus != "returned":
                                        if st.button("✕ Reject", key=f"ret_i_{entity}_{rm}_{username}_{iid}"):
                                            st.session_state[f"reject_init_{iid}"] = True
                                            st.rerun()
                                if st.session_state.get(f"reject_init_{iid}"):
                                    reject_comment_i = st.text_area(
                                        f'Comment for rejecting "{iname}" (will be shown to the user):',
                                        key=f"rc_init_{iid}",
                                        placeholder="Explain what needs to be updated...",
                                        height=80,
                                    )
                                    ri1, ri2 = st.columns(2)
                                    with ri1:
                                        if st.button("Confirm Reject", key=f"rc_conf_init_{iid}", type="primary"):
                                            now = int(datetime.now().timestamp()*1000)
                                            init["initiative_status"] = "returned"
                                            init["returned_at"]       = now
                                            init["rejection_comment"] = reject_comment_i.strip()
                                            init.pop("approved_at", None)
                                            sub["status"] = "rejected"
                                            st.session_state.pop(f"reject_init_{iid}", None)
                                            save_draft(username, sub)
                                            st.rerun()
                                    with ri2:
                                        if st.button("Cancel", key=f"rc_cancel_init_{iid}"):
                                            st.session_state.pop(f"reject_init_{iid}", None)
                                            st.rerun()

                                with ic4:
                                    if st.button("🗑 Delete", key=f"del_i_{entity}_{rm}_{username}_{iid}"):
                                        sub["initiatives"] = [i for i in sub["initiatives"] if i["id"] != iid]
                                        if not sub["initiatives"]:
                                            sub["status"] = "in-progress"
                                        save_draft(username, sub)
                                        st.success(f"Deleted: {iname}")
                                        st.rerun()

                                st.write("")

                st.write("")

    # ═══════════════════════════════════════════════════════════════════════
    # TAB: Rollover
    # ═══════════════════════════════════════════════════════════════════════
    with tab_rollover:
        st.caption(
            "Pick one period to roll forward into a new Filing Month. "
            "The target list only shows months after the source — you can't "
            "roll into the same month or an earlier one. All initiatives are "
            "carried regardless of end date. Anyone who already has a real "
            "submission for the target period is skipped, never overwritten."
        )

        if not combos:
            st.info("No submissions to roll over yet.")
        else:
            combo_labels_ro = {
                f"Entity {e} — Filing: {fmt_month(rm)}": (e, rm) for e, rm in combos
            }
            sorted_labels = sorted(
                combo_labels_ro.keys(),
                key=lambda l: combo_labels_ro[l][1],
                reverse=True,
            )

            src_label = st.selectbox(
                "Roll FROM",
                sorted_labels,
                key="ro_src_single",
            )
            src_entity, src_month = combo_labels_ro[src_label]

            # Target list only ever shows months strictly AFTER the source —
            # rolling into the same month or an earlier one is not allowed.
            target_options = [m for m in available_months() if m > src_month]
            default_target = next_month_of(src_month)
            if default_target not in target_options:
                target_options = sorted(set(target_options) | {default_target})
            target_options = sorted(target_options)

            chosen_target = st.selectbox(
                "Roll TO (target Filing Month)",
                target_options,
                index=target_options.index(default_target) if default_target in target_options else 0,
                format_func=fmt_month,
                key=f"ro_target_single_{src_label}",
            )

            # Build preview for this single source → target
            preview = []  # (username, init_names)
            for username, months in all_data.items():
                sub = months.get(src_month)
                if not sub or sub.get("entity") != src_entity:
                    continue
                existing = load_submission(username, chosen_target)
                if existing and existing.get("entity") == src_entity and existing.get("initiatives"):
                    continue
                inits = sub.get("initiatives") or []
                if inits:
                    preview.append((username, [i.get("initiative_name", "Unnamed") for i in inits]))

            if not preview:
                st.info("No users have active initiatives to roll over into this target period.")
            else:
                st.markdown(
                    f"**Preview — Entity {src_entity}: Filing {fmt_month(src_month)} "
                    f"→ Filing {fmt_month(chosen_target)}**"
                )
                for u, names in preview:
                    st.markdown(
                        f"&nbsp;&nbsp;• **{u}** — " + ", ".join(f"*{n}*" for n in names),
                        unsafe_allow_html=True,
                    )

                total_subs = len(preview)
                if st.button(
                    f"🔄 Roll Over {total_subs} submission{'s' if total_subs!=1 else ''} "
                    f"→ Filing: {fmt_month(chosen_target)}",
                    type="primary",
                    key="do_rollover_single",
                ):
                    rolled = rollover_entity(all_data, src_entity, src_month, chosen_target)
                    if rolled:
                        st.success(
                            f"✓ Rolled over for: {', '.join(rolled)}. "
                            f"They'll find a pre-filled draft under Filing Month {fmt_month(chosen_target)} "
                            f"when they log in. They need to review their initiatives and click "
                            f"**Submit for Review** before you can approve."
                        )
                        st.rerun()
                    else:
                        st.info("Nothing to roll over.")

    # ═══════════════════════════════════════════════════════════════════════
    # TAB: Export
    # ═══════════════════════════════════════════════════════════════════════
    with tab_export:
        if not combos:
            st.info("No submissions found yet.")
        else:
            st.caption(
                "Build a custom export by choosing which periods/entities and which "
                "statuses to include. Excel gets one tab per entity."
            )

            # ── 1. Periods / Entities ────────────────────────────────────────
            st.markdown("**1. Periods & Entities**")
            combo_labels = {f"Entity {e} — Filing: {fmt_month(rm)}": (e, rm) for e, rm in combos}
            all_labels   = list(combo_labels.keys())

            sel_all_periods = st.checkbox("Include all periods/entities", value=True, key="sel_all_periods")
            if sel_all_periods:
                chosen_labels = all_labels
            else:
                chosen_labels = st.multiselect(
                    "Choose specific periods/entities:",
                    all_labels,
                    default=all_labels,
                    key="chosen_period_labels",
                )
            chosen_combos = [combo_labels[l] for l in chosen_labels]

            st.write("")

            # ── 2. Status filter (checkbox dropdown) ──────────────────────────
            st.markdown("**2. Status**")
            status_options = ["In Progress", "Ready for Review", "Approved", "Archived", "Rejected"]
            status_key_map = {
                "In Progress":     "in-progress",
                "Ready for Review":"submitted",
                "Approved":        "approved",
                "Archived":        "archived",
                "Rejected":        "rejected",
            }

            if "export_status_sel" not in st.session_state:
                st.session_state.export_status_sel = ["Ready for Review", "Approved", "Archived"]

            def _apply_status_preset(selected: list[str]):
                # Update the underlying selection AND clear each checkbox widget's
                # cached state — otherwise Streamlit keeps showing the checkbox's
                # old value instead of picking up the new preset (same fix used
                # elsewhere for selectbox widgets that don't refresh on their own).
                st.session_state.export_status_sel = selected
                for opt in status_options:
                    st.session_state.pop(f"status_chk_{opt}", None)
                st.rerun()

            # Quick preset buttons
            pc1, pc2, pc3, pc4 = st.columns(4)
            with pc1:
                if st.button("All statuses", key="preset_all"):
                    _apply_status_preset(status_options.copy())
            with pc2:
                if st.button("Closed out only", key="preset_closed",
                             help="Approved + Archived"):
                    _apply_status_preset(["Approved", "Archived"])
            with pc3:
                if st.button("Needs action", key="preset_action",
                             help="In Progress + Rejected"):
                    _apply_status_preset(["In Progress", "Rejected"])
            with pc4:
                if st.button("Reviewed only", key="preset_reviewed",
                             help="Ready for Review + Approved + Archived"):
                    _apply_status_preset(["Ready for Review", "Approved", "Archived"])

            # Checkbox dropdown — single clean button that opens a checklist panel,
            # instead of a multiselect whose pills can wrap and overflow.
            n_sel = len(st.session_state.export_status_sel)
            n_tot = len(status_options)
            button_label = (
                f"☑ {n_sel} of {n_tot} statuses selected ▾"
                if 0 < n_sel < n_tot else
                f"☑ All statuses selected ▾" if n_sel == n_tot else
                "☐ No statuses selected ▾"
            )
            with st.popover(button_label, use_container_width=True):
                for opt in status_options:
                    checked = st.checkbox(
                        opt,
                        value=opt in st.session_state.export_status_sel,
                        key=f"status_chk_{opt}",
                    )
                    if checked and opt not in st.session_state.export_status_sel:
                        st.session_state.export_status_sel.append(opt)
                        st.rerun()
                    elif not checked and opt in st.session_state.export_status_sel:
                        st.session_state.export_status_sel.remove(opt)
                        st.rerun()

            chosen_statuses_display = st.session_state.export_status_sel
            if chosen_statuses_display:
                st.caption("Selected: " + ", ".join(chosen_statuses_display))
            chosen_status_keys = [status_key_map[s] for s in chosen_statuses_display]

            st.write("")

            # ── 3. Live summary of what will be included ─────────────────────
            st.markdown("**3. Summary**")
            matched_rows  = 0
            matched_users = set()
            matched_inits = 0
            for entity, rm in chosen_combos:
                for username, months in all_data.items():
                    sub = months.get(rm)
                    if not sub or sub.get("entity") != entity:
                        continue
                    if sub.get("status") not in chosen_status_keys:
                        continue
                    inits = sub.get("initiatives") or []
                    if not inits:
                        continue
                    matched_rows += 1
                    matched_users.add(username)
                    matched_inits += len(inits)

            if not chosen_combos or not chosen_status_keys:
                st.warning("Select at least one period/entity and one status to export.")
            else:
                unique_entities = sorted({e for e,_ in chosen_combos})
                sc1, sc2, sc3, sc4 = st.columns(4)
                sc1.metric("Entities", len(unique_entities))
                sc2.metric("Reports", matched_rows)
                sc3.metric("Users", len(matched_users))
                sc4.metric("Initiatives", matched_inits)
                tab_preview = ", ".join(f'"Entity {e}"' for e in unique_entities)
                st.caption(f"Excel tabs: {tab_preview}")

            st.write("")

            # ── 4. Download ───────────────────────────────────────────────────
            st.markdown("**4. Download**")
            today = datetime.now().strftime("%m%d%y")
            status_tag = "_".join(s.replace(" ","") for s in chosen_statuses_display) if len(chosen_statuses_display) < len(status_options) else "AllStatuses"
            fname = f"Consolidated_Report_{status_tag}_{today}.xlsx"

            if chosen_combos and chosen_status_keys and matched_rows > 0:
                xlsx = build_excel_consolidated(all_data, chosen_combos, status_filter=chosen_status_keys)
                st.download_button(
                    f"↓ Download Consolidated Report ({matched_rows} report{'s' if matched_rows!=1 else ''})",
                    data=xlsx,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
            elif chosen_combos and chosen_status_keys:
                st.caption("No reports match this combination — nothing to download yet.")

    # ═══════════════════════════════════════════════════════════════════════
    # TAB: Backup & Data
    # ═══════════════════════════════════════════════════════════════════════
    with tab_backup:
        backup_date = datetime.now().strftime("%m%d%y")

        st.markdown("**Backup**")
        st.caption(
            "Downloads all historical data as an Excel file — one tab per entity, "
            "all periods, all users, all statuses. Only current initiatives are included; "
            "anything previously deleted is already gone. "
            "Save this somewhere safe (Google Drive, OneDrive) after each approval cycle."
        )
        backup_bytes = create_backup_excel(all_data)
        n_subs = sum(
            1 for f in DATA_DIR.glob("*.json")
            if not f.stem.endswith("_months") and f.stem != "registry"
            and bool(json.loads(f.read_text()).get("initiatives"))
        )
        col1, col2 = st.columns([2, 3])
        with col1:
            st.download_button(
                "↓ Download Full Data Backup (.xlsx)",
                data=backup_bytes,
                file_name=f"RD_Data_Backup_{backup_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with col2:
            st.caption(f"{n_subs} submission{'s' if n_subs != 1 else ''} with data · {len(backup_bytes)//1024} KB")

        st.divider()

        st.markdown("**⚠ Delete All History**")
        st.caption(
            "Permanently deletes all submission data from the server. "
            "Download the backup above first. This cannot be undone."
        )
        if not st.session_state.get("confirm_delete_all"):
            if st.button("🗑 Delete All History", key="del_all_btn"):
                st.session_state.confirm_delete_all = True
                st.rerun()
        else:
            st.error(
                "This will permanently delete ALL submissions for ALL users across ALL periods. "
                "Are you sure?"
            )
            da1, da2 = st.columns(2)
            with da1:
                if st.button("Yes, delete everything", key="del_all_confirm", type="primary"):
                    delete_all_history()
                    st.session_state.confirm_delete_all = False
                    st.success("All history deleted. The app will reload.")
                    st.rerun()
            with da2:
                if st.button("Cancel", key="del_all_cancel"):
                    st.session_state.confirm_delete_all = False
                    st.rerun()

        st.divider()

        # ── Restore from Backup ──────────────────────────────────────────────
        st.markdown("**📤 Restore from Backup**")
        st.caption(
            "If the server restarted and lost its data, upload a previously downloaded "
            "backup or export (.xlsx) here to rebuild submissions from it. "
            "Works with the full data backup, a consolidated report, or an individual user's export. "
            "All statuses are restored exactly as they were — Approved, Archived, In Progress, etc."
        )

        uploaded = st.file_uploader(
            "Upload a backup or export file", type=["xlsx"], key="restore_uploader"
        )

        if uploaded is not None:
            file_bytes = uploaded.getvalue()
            parsed, parse_warnings = parse_import_workbook(file_bytes)

            if not parsed:
                st.error("Couldn't find any usable data in this file.")
                if parse_warnings:
                    for w in parse_warnings:
                        st.caption(f"⚠ {w}")
            else:
                # Preview — split into "will import" vs "already has data"
                will_import, will_skip = [], []
                for (username, entity, rm), data in parsed.items():
                    existing = load_submission(username, rm)
                    already_has_data = bool(existing and existing.get("initiatives"))
                    row = {
                        "User": username,
                        "Entity": entity,
                        "Period": fmt_month(rm),
                        "Status": STATUS_LABELS.get(data["status"], data["status"]),
                        "Initiatives": len(data["initiatives"]),
                    }
                    (will_skip if already_has_data else will_import).append(row)

                pc1, pc2, pc3 = st.columns(3)
                pc1.metric("Periods found", len(parsed))
                pc2.metric("Will import", len(will_import))
                pc3.metric("Already have data", len(will_skip))

                if will_import:
                    st.markdown("**Will be imported:**")
                    for row in will_import:
                        st.markdown(
                            f"&nbsp;&nbsp;• **{row['User']}** — Entity {row['Entity']} — "
                            f"{row['Period']} — {row['Status']} "
                            f"({row['Initiatives']} initiative{'s' if row['Initiatives']!=1 else ''})",
                            unsafe_allow_html=True,
                        )

                if will_skip:
                    with st.expander(f"⚠ {len(will_skip)} period(s) already have data — won't be touched unless you choose to overwrite"):
                        for row in will_skip:
                            st.markdown(
                                f"&nbsp;&nbsp;• **{row['User']}** — Entity {row['Entity']} — "
                                f"{row['Period']} — {row['Status']}",
                                unsafe_allow_html=True,
                            )

                if parse_warnings:
                    with st.expander(f"ℹ {len(parse_warnings)} note(s) about this import"):
                        for w in parse_warnings:
                            st.caption(w)

                overwrite = False
                if will_skip:
                    overwrite = st.checkbox(
                        "Overwrite periods that already have data (replaces their current initiatives)",
                        value=False,
                        key="restore_overwrite",
                    )

                n_to_apply = len(parsed) if overwrite else len(will_import)
                if n_to_apply == 0:
                    st.info("Nothing to import — every period in this file already has data.")
                else:
                    if st.button(
                        f"📤 Restore {n_to_apply} period{'s' if n_to_apply!=1 else ''}",
                        type="primary",
                        key="confirm_restore",
                    ):
                        result = apply_import(parsed, overwrite=overwrite)
                        st.success(
                            f"✓ Restored {len(result['imported'])} period(s). "
                            + (f"Skipped {len(result['skipped'])} (already had data)." if result["skipped"] else "")
                        )
                        st.rerun()


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    inject_css()
    init_session()
    screen = st.session_state.screen
    if screen == "login":
        screen_login()
    elif screen == "dashboard":
        screen_dashboard()
    elif screen == "wizard":
        screen_wizard()
    elif screen == "admin":
        screen_admin()

if __name__ == "__main__":
    main()
