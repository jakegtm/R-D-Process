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

def ts_to_est(ms: int, fmt: str = "%b %d, %Y %I:%M %p") -> str:
    """Convert a millisecond UTC timestamp to a formatted EST string."""
    from datetime import timezone, timedelta
    est = timezone(timedelta(hours=-5))   # EST = UTC-5 (covers ET year-round simply)
    dt  = datetime.fromtimestamp(ms / 1000, tz=timezone.utc).astimezone(est)
    return dt.strftime(fmt) + " EST"

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


def permissions_path() -> Path:
    return DATA_DIR / "permissions.json"

def load_permissions() -> dict:
    """Returns {username: [entity_list]} — empty list means user sees all entities."""
    p = permissions_path()
    return json.loads(p.read_text()) if p.exists() else {}

def save_permissions(perms: dict):
    permissions_path().write_text(json.dumps(perms, indent=2))

def get_user_entities(username: str) -> list[str]:
    """Returns the entities this user is permitted to see.
    If no restriction is configured, returns all entities."""
    perms = load_permissions()
    if username not in perms or not perms[username]:
        return all_entities()
    return [e for e in perms[username] if e in all_entities()]


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
            # Skip resolved initiatives — user marked them as complete
            if init.get("resolved"):
                continue
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
    For every user who has an APPROVED or ARCHIVED submission under
    (source_entity, source_month), create a new in-progress draft under
    (source_entity, target_month) — unless they already have one there.

    Matching the Power Automate flow:
    - Only approved/archived submissions can be rolled (not in-progress/submitted).
    - The source submission is automatically archived after rollover.

    Returns list of usernames that were rolled over.
    """
    rolled: list[str] = []

    for username, months in all_data.items():
        src_sub = months.get(source_month)
        if not src_sub or src_sub.get("entity") != source_entity:
            continue

        # Only roll over approved or already-archived submissions
        if src_sub.get("status") not in ("approved", "archived"):
            continue

        # Don't overwrite an existing real submission in the target period
        existing = load_submission(username, target_month)
        if existing and existing.get("entity") == source_entity and existing.get("initiatives"):
            continue

        initiatives = src_sub.get("initiatives") or []
        if not initiatives:
            continue

        import copy as _copy
        rolled_inits = []
        for i in initiatives:
            init = _copy.deepcopy(i)
            init["id"]                = f"{int(__import__('time').time()*1000)}{len(rolled_inits)}"
            init["initiative_status"] = "active"
            init["approved_at"]       = None
            init["returned_at"]       = None
            init["carry_over"]        = True
            init["month_yr"]          = ""
            rolled_inits.append(init)

        new_sub: dict = {
            "initiatives":      rolled_inits,
            "status":           "in-progress",
            "entity":           source_entity,
            "reporting_month":  target_month,
            "activities_month": prev_month_of(target_month),
            # Tracks that this draft was created by admin rollover, not the user.
            # Used to show the user a notification on login.
            "rolled_over_from": source_month,
        }
        save_draft(username, new_sub)

        # Auto-archive the source submission (matching PA Flow 4 behaviour)
        if src_sub.get("status") != "archived":
            src_sub["status"]      = "archived"
            src_sub["archived_at"] = int(__import__("time").time() * 1000)
            save_draft(username, src_sub)

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
        # pathway: which flow the user chose this month for this carry-over initiative.
        # "continuing" | "resolved" | "new_direction" | "" (not yet chosen)
        "pathway":               "",
        # resolved: True once user marks initiative as complete via the Resolved pathway.
        # Resolved initiatives are excluded from future carry-over suggestions.
        "resolved":              False,
        # completion_date: user-set date when resolved (YYYY-MM-DD).
        # Populates the Completion Date column in the export.
        "completion_date":       None,
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
                init.get("completion_date") or
                (ts_to_est(init["approved_at"], "%Y-%m-%d %H:%M") if init.get("approved_at") else
                 datetime.fromtimestamp(sub["approved_at"] / 1000).strftime("%Y-%m-%d")
                 if sub.get("approved_at") else "")
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
    ws.title = f"Entity {entity}" if entity else (fmt_month_tab(am) or "Report")
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
    group_by: str = "entity",   # "entity" → one tab per entity | "user" → one tab per person
) -> bytes:
    """
    Builds a consolidated Excel workbook.

    group_by="entity" (default, for user self-exports):
        One sheet per entity tab named "Entity 107". All periods for that
        entity appear as rows sorted by period then username.

    group_by="user" (for admin consolidated reports):
        One sheet per team member tab named after the person. All their
        entities/periods appear as rows sorted by entity then period.

    status_filter: if provided, only include submissions whose status is in
    this list. Takes precedence over submitted_only.
    """
    wb = Workbook()
    wb.remove(wb.active)

    if status_filter is not None:
        allowed_statuses = set(status_filter)
        status_desc = ", ".join(STATUS_LABELS.get(s, s).split(" ", 1)[-1] for s in status_filter) or "None selected"
    elif submitted_only:
        allowed_statuses = {"submitted", "approved", "archived"}
        status_desc = "Submitted, Approved & Archived"
    else:
        allowed_statuses = None
        status_desc = "All statuses"

    if group_by == "user":
        # ── One tab per person ───────────────────────────────────────────────
        for username in sorted(all_data.keys()):
            rows = []
            # Sort by entity then reporting_month so a user's sheet reads chronologically
            combos_for_user = sorted(
                [(e, rm) for e, rm in selected_combos],
                key=lambda x: (x[0], x[1]),
            )
            for entity, rm in combos_for_user:
                sub = all_data[username].get(rm)
                if not sub or sub.get("entity") != entity:
                    continue
                if allowed_statuses is not None and sub.get("status") not in allowed_statuses:
                    continue
                rows.extend(_sub_to_rows(username, sub, include_user=False))

            if not rows:
                continue
            ws = wb.create_sheet(title=username[:31])
            subtitle = f"Team member: {username}  |  {status_desc}"
            _write_sheet(ws, rows, subtitle)

    else:
        # ── One tab per entity (original behaviour) ──────────────────────────
        entities = sorted({e for e, _ in selected_combos})
        for entity in entities:
            ws = wb.create_sheet(title=f"Entity {entity}"[:31])
            subtitle = f"Entity: {entity}  |  {status_desc}"
            entity_combos = sorted(
                [(e, rm) for e, rm in selected_combos if e == entity],
                key=lambda x: x[1],
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

def best_draft_for_user(username: str) -> dict:
    """
    Picks the most relevant draft to load on login:
    1. The earliest in-progress period that has real initiatives
       (likely a rollover waiting to be reviewed and submitted).
    2. The current calendar month's filing if it exists.
    3. A fresh empty draft for the current month.

    This ensures the user lands on something actionable — e.g. if
    they've already submitted June and July, and a rollover created
    an August in-progress draft, they open directly on August.
    """
    months_index = load_user_months(username)

    # Priority 1: any in-progress period with real initiatives
    # (sort ascending so earliest un-submitted period wins)
    in_progress = sorted([
        m for m in months_index
        if (sub := load_submission(username, m))
        and sub.get("initiatives")
        and sub.get("status") == "in-progress"
    ])
    if in_progress:
        return load_submission(username, in_progress[0])

    # Priority 2: current calendar month if it already has a submission
    cur = cur_month()
    cur_sub = load_submission(username, cur)
    if cur_sub and cur_sub.get("initiatives"):
        return cur_sub

    # Priority 3: most recent period with any real content
    for month in sorted(months_index, reverse=True):
        sub = load_submission(username, month)
        if sub and sub.get("initiatives"):
            return sub

    # Fallback: empty draft for current month
    draft = empty_draft()
    draft["reporting_month"]  = cur
    draft["activities_month"] = prev_month_of(cur)
    return draft


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
                # Auto-load the most actionable period (in-progress rollover first,
                # then current month, then most recent) so the user lands on the
                # right report without having to change Report Setup manually.
                if "Admin" not in sel:
                    best = best_draft_for_user(name)
                    st.session_state.draft = best
                    # Clear Report Setup widget keys so Streamlit renders
                    # the selectboxes from the new draft's values rather
                    # than whatever was cached from a prior session/period.
                    for k in ["su_entity", "su_filing"]:
                        st.session_state.pop(k, None)
                    # Clear the dynamic Reporting Period key too
                    rm = best.get("reporting_month", "")
                    st.session_state.pop(f"su_act_{rm}", None)
                else:
                    st.session_state.draft = empty_draft()
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

def _all_rejection_comments(sub: dict) -> list[str]:
    """Collect every rejection comment on a submission — report-level and per-initiative."""
    comments = []
    rc = sub.get("rejection_comment", "")
    if rc:
        comments.append(rc)
    for init in sub.get("initiatives") or []:
        irc = init.get("rejection_comment", "")
        if irc and init.get("initiative_status") == "returned":
            comments.append(f"{init.get('initiative_name','Initiative')}: {irc}")
    return comments


def render_user_reminder(user: str, current_filing_month: str = ""):
    """
    Three types of alerts, in priority order:
    1. Rejected periods — shown prominently regardless of what period is currently viewed.
    2. Unsubmitted in-progress periods from other filing months.
    3. Monthly filing reminder (past the 5th).
    """
    months_index = load_user_months(user)

    # ── Priority 1: Rejected periods — always show, takes precedence ─────────
    for month in sorted(months_index, reverse=True):
        sub = load_submission(user, month)
        if not sub or not sub.get("initiatives"):
            continue
        if sub.get("status") != "rejected":
            continue
        ack_key = f"ack_rejection_{month}"
        already_acked = st.session_state.get(ack_key, False)
        comments = _all_rejection_comments(sub)
        comment_str = ""
        if comments:
            comment_str = "  " + "  |  ".join(f'"{c}"' for c in comments)

        if month == current_filing_month:
            # Current period rejection is shown inline in the status strip — skip here
            continue

        # Show rejection banner for OTHER periods regardless of what's currently loaded
        st.error(
            f"🔴 Your **{fmt_month(month)}** report was Rejected by the Oversight Lead."
            + (f"  Comment: {comment_str}" if comment_str else "")
        )
        if not already_acked:
            st.warning(
                f"Switch to **Filing: {fmt_month(month)}** in Report Setup above, "
                "then click **Acknowledge & Prepare Resubmission** to begin updating."
            )

    # ── Priority 2: Unsubmitted in-progress drafts from other periods ────────
    pending_periods = []
    for month in months_index:
        if month == current_filing_month:
            continue
        sub = load_submission(user, month)
        if sub and sub.get("initiatives") and sub.get("status") == "in-progress":
            pending_periods.append((month, bool(sub.get("rolled_over_from"))))

    if pending_periods:
        parts = []
        for month, was_rolled in pending_periods:
            label = f"**{fmt_month(month)}**" + (" *(rolled over by admin)*" if was_rolled else "")
            parts.append(label)
        st.warning(
            f"⚠ You have an unsubmitted report for {', '.join(parts)}. "
            "Open **Report Setup** above, select that Filing Month, "
            "review your initiatives, and click **Submit for Review**."
        )

    # ── Priority 3: Monthly filing reminder (past the 5th) ───────────────────
    today = date.today()
    if today.day >= 5:
        cur = cur_month()
        existing = load_submission(user, cur)
        if not existing or not existing.get("initiatives"):
            st.info(
                f"📅 It's past the 5th and you haven't started your "
                f"**{fmt_month(cur)}** filing yet. "
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
    render_user_reminder(user, current_filing_month=draft.get("reporting_month", ""))

    # ── Rollover notification ─────────────────────────────────────────────────
    # Show a clear banner when the current draft was created by admin rollover
    # and hasn't been submitted yet — so the user knows exactly what happened.
    rolled_from = draft.get("rolled_over_from")
    if rolled_from and draft.get("status") == "in-progress":
        st.info(
            f"📋 **The Oversight Lead has rolled your {fmt_month(rolled_from)} report "
            f"forward to {fmt_month(draft.get('reporting_month',''))}.**  \n"
            "Your initiatives have been carried over and are ready for you to review. "
            "Update your activities for this month, then click **Submit for Review** when ready.",
            icon="📋",
        )

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
        # If the current draft is for a period outside the normal window
        # (e.g. a rollover created 2+ months ahead), include it so the
        # dropdown doesn't fall back to the wrong month.
        cur_draft_rm = draft.get("reporting_month")
        if cur_draft_rm and cur_draft_rm not in mlist:
            mlist = sorted(set(mlist) | {cur_draft_rm}, reverse=True)

        with su1:
            entity_options = get_user_entities(user) + ["+ Add new entity..."]
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
            st.session_state.show_entry_picker = False   # close picker when period changes
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
            ts = ts_to_est(draft["submitted_at"], "%b %d, %Y %I:%M %p")
            st.caption(f"Submitted {ts}")
        if draft["status"] == "rejected":
            comments = _all_rejection_comments(draft)
            msg = "Your report was Rejected by the Oversight Lead."
            if comments:
                msg += " " + "  |  ".join(f'"{c}"' for c in comments)
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
            arc_str = f" on {ts_to_est(arc_ts, '%b %d, %Y')}" if arc_ts else ""
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
        if not submitted:
            existing_inits = draft.get("initiatives", [])
            n_existing = len(existing_inits)

            # Show context if there are already initiatives in this report
            if n_existing > 0:
                n_approved = sum(1 for i in existing_inits if i.get("initiative_status") == "approved")
                n_returned = sum(1 for i in existing_inits if i.get("initiative_status") == "returned")
                parts = [f"{n_existing} initiative{'s' if n_existing!=1 else ''} already in this report"]
                if n_approved:
                    parts.append(f"{n_approved} accepted by admin")
                if n_returned:
                    parts.append(f"⚠ {n_returned} need{'s' if n_returned==1 else ''} revision")
                st.caption(" · ".join(parts) + " — adding more will include them in the same submission.")

            if st.button("＋ Add Initiative", type="primary", use_container_width=True):
                st.session_state.screen = "entry_picker"
                st.rerun()

    if not draft["initiatives"]:
        st.info("No initiatives added yet. Use **＋ Guided Entry** to add one at a time, or **📊 Bulk Entry** to fill a spreadsheet grid for multiple initiatives at once.")
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
                    ret_str = f" on {ts_to_est(ret_ts, '%b %d %I:%M %p')}" if ret_ts else ""
                    rc = init.get("rejection_comment", "")
                    rc_str = f' Comment: "{rc}".' if rc else ""
                    st.warning(f"⚠ This initiative was Rejected{ret_str}.{rc_str} Edit and resubmit.")
                elif istatus == "approved":
                    appr_ts = init.get("approved_at")
                    appr_str = f" on {ts_to_est(appr_ts, '%b %d %I:%M %p')}" if appr_ts else ""
                    st.success(f"✓ Approved by admin{appr_str}.")

                is_archived = draft["status"] == "archived"
                st.write("")
                c1, c2, c3, c4 = st.columns([2.5, 0.8, 0.8, 0.8])
                with c2:
                    if st.button("✏ Edit", key=f"edit_{iid}", disabled=is_archived,
                                 help="This period is archived and locked." if is_archived else None):
                        st.session_state.wiz_init = dict(init)
                        st.session_state.wiz_step = 0
                        # Carry-over initiatives go to the pathway selector first
                        if init.get("carry_over") and not is_archived:
                            st.session_state.screen = "pathway_select"
                        else:
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
        if draft.get("initiatives"):
            xlsx = build_excel_individual(user, draft)
            fname = export_filename(draft.get("entity",""), draft.get("reporting_month",""))
            st.download_button(
                "↓ Download This Filing",
                data=xlsx,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Downloads this period's report only. Use Export My Reports below to download across multiple periods.",
            )

    # ── Export ────────────────────────────────────────────────────────────────
    st.write("")
    render_user_export_section(user)

    # ── History ──────────────────────────────────────────────────────────────
    st.write("")
    render_history_section(user, draft.get("reporting_month",""))


# ── User Export ───────────────────────────────────────────────────────────────

def render_user_export_section(user: str):
    """
    Lets a signed-in user filter and download their own submissions as Excel.
    Filters: entity, filing month, reporting period, status.
    """
    all_months = load_user_months(user)
    if not all_months:
        return

    # Collect every submission for this user
    user_subs = []   # (reporting_month, sub)
    for month in all_months:
        sub = load_submission(user, month)
        if sub and sub.get("initiatives"):
            user_subs.append((month, sub))

    if not user_subs:
        return

    with st.expander("⬇ Export My Reports", expanded=False):
        st.caption("Filter your reports and download a custom Excel file.")
        st.write("")

        STATUS_LABELS_USER = {
            "in-progress": "🟡 In Progress",
            "submitted":   "🔵 Ready for Review",
            "approved":    "🟢 Approved",
            "rejected":    "🔴 Rejected",
            "archived":    "📦 Archived",
        }

        # ── Filters ──────────────────────────────────────────────────────────
        fc1, fc2, fc3 = st.columns(3)

        # Entity filter
        entities_present = sorted({s.get("entity","") for _, s in user_subs if s.get("entity")})
        with fc1:
            f_entity = st.selectbox(
                "Entity",
                ["All"] + entities_present,
                key="ue_entity",
            )

        # Filing Month filter
        filing_months = sorted({m for m, _ in user_subs}, reverse=True)
        with fc2:
            f_filing = st.selectbox(
                "Filing Month",
                ["All"] + [fmt_month(m) for m in filing_months],
                key="ue_filing",
            )

        # Reporting Period filter
        act_months = sorted(
            {s.get("activities_month","") for _, s in user_subs if s.get("activities_month")},
            reverse=True,
        )
        with fc3:
            f_period = st.selectbox(
                "Reporting Period",
                ["All"] + [fmt_month(m) for m in act_months],
                key="ue_period",
            )

        # Status checkboxes — flat, fully visible
        st.write("")
        st.caption("**Status**")
        statuses_present = sorted({s.get("status","") for _, s in user_subs})
        sel_statuses: list[str] = []
        status_cols = st.columns(len(statuses_present) if statuses_present else 1)
        for i, status_key in enumerate(statuses_present):
            with status_cols[i]:
                if st.checkbox(
                    STATUS_LABELS_USER.get(status_key, status_key),
                    value=True,
                    key=f"ue_status_{status_key}",
                ):
                    sel_statuses.append(status_key)

        # ── Apply filters ─────────────────────────────────────────────────────
        filtered = []
        for month, sub in user_subs:
            if f_entity != "All" and sub.get("entity") != f_entity:
                continue
            if f_filing != "All" and fmt_month(month) != f_filing:
                continue
            if f_period != "All" and fmt_month(sub.get("activities_month","")) != f_period:
                continue
            if sub.get("status") not in sel_statuses:
                continue
            filtered.append((month, sub))

        st.write("")

        # ── Preview & Download ────────────────────────────────────────────────
        if not filtered:
            st.caption("No reports match these filters.")
        else:
            # Summary line
            n_inits = sum(len(s.get("initiatives") or []) for _, s in filtered)
            st.caption(
                f"{len(filtered)} report{'s' if len(filtered)!=1 else ''} · "
                f"{n_inits} initiative{'s' if n_inits!=1 else ''}"
            )

            if len(filtered) == 1:
                # Single report → individual export
                month, sub = filtered[0]
                xlsx  = build_excel_individual(user, sub)
                fname = export_filename(sub.get("entity",""), month)
                st.caption("ℹ️ Only 1 period matches — downloading that single filing. Select more periods or change filters to get a consolidated export.")
                st.download_button(
                    f"↓ Download — {sub.get('entity','')} {fmt_month(month)} (1 period)",
                    data=xlsx,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    key="ue_dl_single",
                )
            else:
                # Multiple reports → consolidated export (user's data only)
                user_all_data = {user: {m: s for m, s in filtered}}
                chosen_combos = list({(s.get("entity",""), m) for m, s in filtered})
                chosen_status_keys = list(sel_statuses)
                xlsx = build_excel_consolidated(
                    user_all_data, chosen_combos,
                    status_filter=chosen_status_keys,
                    group_by="entity",
                )
                today = datetime.now().strftime("%m%d%y")
                fname = f"{user.replace(' ','_')}_RD_Report_{today}.xlsx"
                st.download_button(
                    f"↓ Download {len(filtered)} Reports",
                    data=xlsx,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    key="ue_dl_multi",
                )


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
            # Export button for this past period — only if there are initiatives
            if sub.get("entity") and sub.get("reporting_month") and inits:
                h_xlsx  = build_excel_individual(user, sub)
                h_fname = export_filename(sub["entity"], sub["reporting_month"])
                st.download_button(
                    f"↓ {entity} — {fmt_month(month)} (this period only)",
                    data=h_xlsx,
                    file_name=h_fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"hist_dl_{month}",
                    help="Downloads this single period's report. Use Export My Reports above for a custom multi-period download.",
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
        # Use a stable key so Streamlit preserves the selection across reruns.
        # On first render for this initiative (key not yet in session_state),
        # seed it from the saved value so edits pre-fill correctly.
        ms_key = f"wiz_team_{init.get('id','new')}_{step}"
        if ms_key not in st.session_state:
            st.session_state[ms_key] = val or []
        new_val = st.multiselect(
            s["label"], EMPLOYEES,
            key=ms_key,
        )

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

    # ── Navigation — radio stores the tab in session state so widget interactions
    # inside a section don't reset back to Submissions on every rerun. ─────────
    ADMIN_TABS = ["📁 Submissions", "🔄 Rollover", "⬇ Export", "🗄 Backup & Data", "⚙ Settings"]
    if "admin_tab" not in st.session_state:
        st.session_state.admin_tab = ADMIN_TABS[0]
    active_tab = st.radio(
        "##nav",
        ADMIN_TABS,
        index=ADMIN_TABS.index(st.session_state.admin_tab),
        horizontal=True,
        label_visibility="collapsed",
        key="admin_tab_radio",
    )
    st.session_state.admin_tab = active_tab
    st.write("")

    # ── Show only the active section ──────────────────────────────────────────
    # (Previously used st.tabs which resets on every widget rerun inside a tab)

    if active_tab == "📁 Submissions":
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
                                ts = ts_to_est(sub["submitted_at"], "%b %d %I:%M %p")
                                st.caption(f"Submitted: {ts}")
                            if sub.get("approved_at"):
                                ts = ts_to_est(sub["approved_at"], "%b %d %I:%M %p")
                                st.caption(f"Report approved: {ts}")
                        with h2:
                            if inits:
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
                            arc_str = ts_to_est(arc_ts, "%b %d, %Y %I:%M %p") if arc_ts else ""
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
                                        ts = ts_to_est(init["approved_at"], "%b %d %I:%M %p")
                                        st.caption(f"   Approved {ts}")
                                    if init.get("returned_at"):
                                        ts = ts_to_est(init["returned_at"], "%b %d %I:%M %p")
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


    if active_tab == "🔄 Rollover":
        st.caption(
            "Only **Approved** reports can be rolled over. The source period is automatically "
            "**Archived** after rollover. Anyone already in the target period is skipped."
        )

        if not combos:
            st.info("No submissions to roll over yet.")
        else:
            eligible_combos = [
                (e, rm) for e, rm in combos
                if any(
                    sub.get("status") in ("approved", "archived")
                    for ud in all_data.values()
                    if (sub := ud.get(rm)) and sub.get("entity") == e
                )
            ]

            if not eligible_combos:
                st.info("No approved reports to roll over yet. Approve a report in the Submissions tab first.")
            else:
                def _rollable_users_for(entity: str, src_month: str, target: str) -> list[str]:
                    out = []
                    for username, ud in all_data.items():
                        sub = ud.get(src_month)
                        if not sub or sub.get("entity") != entity:
                            continue
                        if sub.get("status") not in ("approved", "archived"):
                            continue
                        if not sub.get("initiatives"):
                            continue
                        existing = load_submission(username, target)
                        if existing and existing.get("entity") == entity and existing.get("initiatives"):
                            continue
                        out.append(username)
                    return out

                from collections import defaultdict
                grouped: dict[str, dict[str, list]] = defaultdict(lambda: defaultdict(list))
                for e, rm in sorted(eligible_combos, reverse=True):
                    grouped[e][rm[:4]].append(rm)

                st.markdown("**Select periods to roll — choose a target month for each:**")
                st.caption("Check a period to include it in the rollover. Each period can roll to a different target month.")
                st.write("")

                # Render: entity → year group → individual period rows
                # Each checked row shows an inline target month picker
                selections = []   # (entity, src_month, chosen_target, rollable_users)

                for entity in sorted(grouped.keys()):
                    with st.container():
                        st.markdown(f"**Entity {entity}**")

                        for year in sorted(grouped[entity].keys(), reverse=True):
                            months_in_year = grouped[entity][year]

                            # Year-level "select all" toggle
                            year_all_key = f"ro_all_{entity}_{year}"
                            # Default: check years that have any rollable period
                            any_rollable = any(
                                _rollable_users_for(entity, rm, next_month_of(rm))
                                for rm in months_in_year
                            )
                            year_all = st.checkbox(
                                f"  {year} — select all",
                                value=st.session_state.get(year_all_key, any_rollable),
                                key=year_all_key,
                            )

                            for rm in months_in_year:
                                chk_key = f"ro_chk_{entity}_{rm}"
                                tgt_key = f"ro_target_{entity}_{rm}"

                                # Default target = next month after source
                                default_tgt = next_month_of(rm)

                                # If year-all toggled on, force this period checked
                                if year_all:
                                    st.session_state[chk_key] = True

                                is_checked = st.session_state.get(chk_key, year_all)

                                # Build this period's row: [checkbox | label] [→ target picker if checked]
                                row_cols = st.columns([0.05, 2.2, 0.15, 1.6])

                                with row_cols[0]:
                                    checked = st.checkbox(
                                        "##",
                                        value=is_checked,
                                        key=chk_key,
                                        label_visibility="collapsed",
                                    )

                                with row_cols[1]:
                                    rollable_for_default = _rollable_users_for(entity, rm, default_tgt)
                                    user_label = (
                                        f"{len(rollable_for_default)} user{'s' if len(rollable_for_default)!=1 else ''} ready"
                                        if rollable_for_default else "nothing to roll"
                                    )
                                    color = "#2D6A2D" if rollable_for_default else "#999"
                                    st.markdown(
                                        f'<span style="font-size:15px;">{fmt_month(rm)}</span>'
                                        f'&nbsp;&nbsp;<span style="font-size:13px;color:{color};">{user_label}</span>',
                                        unsafe_allow_html=True,
                                    )

                                if checked:
                                    with row_cols[2]:
                                        st.markdown('<span style="font-size:18px;">→</span>', unsafe_allow_html=True)

                                    with row_cols[3]:
                                        # Target options: all months from source+1 up to 12 months forward
                                        # (not limited to available_months' normal window)
                                        import datetime as _dt
                                        _base = _dt.date.today()
                                        _ext  = [
                                            f"{(_base.year + ((_base.month + i - 1) // 12))}"
                                            f"-{str((_base.month - 1 + i) % 12 + 1).zfill(2)}"
                                            for i in range(1, 14)
                                        ]
                                        tgt_options = sorted(
                                            set(m for m in (_ext + available_months()) if m > rm)
                                            | {default_tgt}
                                        )
                                        cur_tgt = st.session_state.get(tgt_key, default_tgt)
                                        if cur_tgt not in tgt_options:
                                            cur_tgt = default_tgt
                                        chosen_tgt = st.selectbox(
                                            "##",
                                            tgt_options,
                                            index=tgt_options.index(cur_tgt),
                                            format_func=fmt_month,
                                            key=tgt_key,
                                            label_visibility="collapsed",
                                        )
                                        # Compute actual rollable users for the chosen target
                                        rollable = _rollable_users_for(entity, rm, chosen_tgt)
                                        if rollable:
                                            selections.append((entity, rm, chosen_tgt, rollable))
                                        elif rollable_for_default:
                                            st.caption("⚠ Already have data for this target")

                        st.write("")

                st.divider()

                if not selections:
                    st.caption("No periods selected with rollable users — tick checkboxes above to begin.")
                else:
                    total_users = sum(len(r) for _, _, _, r in selections)
                    st.markdown(
                        f"**Preview** — {len(selections)} rollover{'s' if len(selections)!=1 else ''}, "
                        f"{total_users} submission{'s' if total_users!=1 else ''} total"
                    )
                    for entity, src_month, tgt, rollable in selections:
                        st.markdown(
                            f"&nbsp;&nbsp;**Entity {entity}:** {fmt_month(src_month)} → **{fmt_month(tgt)}**",
                            unsafe_allow_html=True,
                        )
                        for u in rollable:
                            inits = [
                                i.get("initiative_name", "Unnamed")
                                for i in (all_data.get(u, {}).get(src_month, {}).get("initiatives") or [])
                            ]
                            st.markdown(
                                f"&nbsp;&nbsp;&nbsp;&nbsp;• **{u}** — " + ", ".join(f"*{n}*" for n in inits),
                                unsafe_allow_html=True,
                            )

                    st.write("")
                    if st.button(
                        f"🔄 Roll Forward {total_users} submission{'s' if total_users!=1 else ''}",
                        type="primary",
                        key="do_rollover_dynamic",
                    ):
                        all_rolled = []
                        for entity, src_month, tgt, rollable in selections:
                            subset = {u: all_data[u] for u in rollable if u in all_data}
                            rolled = rollover_entity(subset, entity, src_month, tgt)
                            all_rolled.extend(rolled)
                        if all_rolled:
                            targets = sorted({tgt for _, _, tgt, _ in selections})
                            st.success(
                                f"✓ Rolled over for: {', '.join(all_rolled)}. "
                                f"Source periods have been archived."
                            )
                            st.rerun()


    if active_tab == "⬇ Export":
        if not combos:
            st.info("No submissions found yet.")
        else:
            from collections import defaultdict as _dd
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

            ex_left, ex_right = st.columns([1, 1])

            # ── LEFT: Periods & Entities (checkboxes grouped by entity → year) ──
            with ex_left:
                st.markdown("**1. Periods & Entities**")

                # Group combos: entity → year → [months]
                ex_grouped = _dd(lambda: _dd(list))
                for e, rm in sorted(combos, reverse=True):
                    ex_grouped[e][rm[:4]].append(rm)

                chosen_combos = []
                for entity in sorted(ex_grouped.keys()):
                    st.markdown(f"**Entity {entity}**")
                    for year in sorted(ex_grouped[entity].keys(), reverse=True):
                        months_in_year = ex_grouped[entity][year]
                        year_all_key = f"ex_all_{entity}_{year}"

                        # "Select year" master checkbox — default on
                        year_all = st.checkbox(
                            f"  {year} — select all",
                            value=st.session_state.get(year_all_key, True),
                            key=year_all_key,
                        )
                        for rm in months_in_year:
                            chk_key = f"ex_chk_{entity}_{rm}"
                            if year_all:
                                st.session_state[chk_key] = True
                            chk = st.checkbox(
                                f"  {fmt_month(rm)}",
                                value=st.session_state.get(chk_key, True),
                                key=chk_key,
                            )
                            if chk:
                                chosen_combos.append((entity, rm))
                    st.write("")

            # ── RIGHT: Status (flat checkboxes, fully visible) ───────────────
            with ex_right:
                st.markdown("**2. Status**")

                st.write("")
                # Flat visible checkboxes — no popover, no hidden state
                new_sel = list(st.session_state.export_status_sel)
                for opt in status_options:
                    chk = st.checkbox(opt, value=opt in new_sel, key=f"st_chk_{opt}")
                    if chk and opt not in new_sel:
                        new_sel.append(opt)
                    elif not chk and opt in new_sel:
                        new_sel.remove(opt)
                if new_sel != st.session_state.export_status_sel:
                    st.session_state.export_status_sel = new_sel
                    st.rerun()
                chosen_statuses_display = st.session_state.export_status_sel
                chosen_status_keys = [status_key_map[s] for s in chosen_statuses_display]

            st.divider()

            # ── Summary ───────────────────────────────────────────────────────
            matched_rows = matched_inits = 0
            matched_users: set = set()
            for entity, rm in chosen_combos:
                for username, months in all_data.items():
                    sub = months.get(rm)
                    if not sub or sub.get("entity") != entity: continue
                    if sub.get("status") not in chosen_status_keys: continue
                    inits = sub.get("initiatives") or []
                    if not inits: continue
                    matched_rows += 1; matched_users.add(username); matched_inits += len(inits)

            if not chosen_combos or not chosen_status_keys:
                st.warning("Select at least one period and one status to export.")
            else:
                unique_entities = sorted({e for e,_ in chosen_combos})
                sc1, sc2, sc3, sc4 = st.columns(4)
                sc1.metric("Entities",    len(unique_entities))
                sc2.metric("Reports",     matched_rows)
                sc3.metric("Users",       len(matched_users))
                sc4.metric("Initiatives", matched_inits)
                st.caption("Excel tabs: " + ", ".join(f'"Entity {e}"' for e in unique_entities))

            st.write("")

            # ── Download ──────────────────────────────────────────────────────
            today = datetime.now().strftime("%m%d%y")
            status_tag = (
                "_".join(s.replace(" ","") for s in chosen_statuses_display)
                if len(chosen_statuses_display) < len(status_options) else "AllStatuses"
            )
            fname = f"Consolidated_Report_{status_tag}_{today}.xlsx"

            if chosen_combos and chosen_status_keys and matched_rows > 0:
                xlsx = build_excel_consolidated(all_data, chosen_combos, status_filter=chosen_status_keys, group_by="entity")
                st.download_button(
                    f"↓ Download Consolidated Report ({matched_rows} report{'s' if matched_rows!=1 else ''})",
                    data=xlsx, file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
            elif chosen_combos and chosen_status_keys:
                st.caption("No reports match this combination — nothing to download yet.")


    if active_tab == "🗄 Backup & Data":
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

    if active_tab == "⚙ Settings":
        st.markdown("**Entity Permissions**")
        st.caption(
            "By default, team members can see all entities. Set restrictions here "
            "to limit which entities each user can file reports for. "
            "Admins always see everything regardless of permissions."
        )

        all_ents = all_entities()
        perms = load_permissions()
        employees_list = [e for e in EMPLOYEES if e != "Admin"]

        changed = False
        for emp in employees_list:
            current = perms.get(emp, [])
            # Seed widget key on first render or after reset
            perm_key = f"perm_{emp}"
            if perm_key not in st.session_state:
                st.session_state[perm_key] = current if current else all_ents
            col1, col2 = st.columns([2, 4])
            with col1:
                st.markdown(f"**{emp}**")
            with col2:
                sel = st.multiselect(
                    f"##perm_{emp}",
                    all_ents,
                    key=perm_key,
                    label_visibility="collapsed",
                    placeholder="All entities (no restriction)",
                )
                # If they selected everything, treat as "no restriction" (empty list)
                new_perm = [] if set(sel) == set(all_ents) else list(sel)
                if new_perm != current:
                    perms[emp] = new_perm
                    changed = True

        st.write("")
        if changed:
            save_permissions(perms)
            st.success("✓ Permissions saved.")

        st.divider()
        st.markdown("**Current Restrictions** (users with fewer than all entities):")
        restricted = {u: v for u, v in perms.items() if v}
        if not restricted:
            st.caption("No restrictions set — all users can see all entities.")
        else:
            for u, ents in restricted.items():
                st.markdown(f"&nbsp;&nbsp;• **{u}**: {', '.join(ents)}", unsafe_allow_html=True)

        st.write("")
        if st.button("Reset all to default (all entities)", key="reset_perms"):
            save_permissions({})
            # Clear the multiselect widget cache so they re-render
            # with all entities selected instead of the old restricted values
            for emp in employees_list:
                st.session_state.pop(f"perm_{emp}", None)
            st.rerun()


# ── Entry Picker Screen ───────────────────────────────────────────────────────

def screen_entry_picker():
    """
    Shown after clicking + Add Initiative.
    Lets the user choose between Guided Entry (wizard) or Bulk Entry (grid).
    """
    draft  = st.session_state.draft
    entity = draft.get("entity", "")
    rm     = draft.get("reporting_month", "")
    am     = draft.get("activities_month", "") or rm

    st.markdown("## ＋ Add Initiative")
    st.caption(
        f"Entity **{entity}** · Filing: **{fmt_month(rm)}** · "
        f"Reporting Period: **{fmt_month(am)}**"
    )
    st.divider()
    st.markdown("### How would you like to enter your initiative(s)?")
    st.write("")

    c1, c2 = st.columns(2)

    with c1:
        st.markdown("#### 📝 Guided Entry")
        st.write(
            "Step-by-step wizard that walks you through each field one at a time. "
            "Best for adding one or two initiatives with full detail."
        )
        st.write("")
        if st.button("Start Guided Entry →", type="primary", use_container_width=True, key="pick_guided"):
            st.session_state.wiz_init = new_initiative()
            st.session_state.wiz_step = 0
            st.session_state.wiz_mode = "new"
            st.session_state.screen   = "wizard"
            st.rerun()

    with c2:
        st.markdown("#### 📊 Bulk Entry")
        st.write(
            "Spreadsheet-style grid where you fill in multiple rows at once. "
            "Best for entering several initiatives in one sitting."
        )
        st.write("")
        if st.button("Start Bulk Entry →", use_container_width=True, key="pick_bulk"):
            st.session_state.pop("bulk_df", None)
            st.session_state.screen = "bulk_entry"
            st.rerun()

    st.write("")
    st.divider()
    if st.button("← Back to Dashboard", key="pick_back"):
        st.session_state.screen = "dashboard"
        st.rerun()


# ── Bulk Entry Screen ──────────────────────────────────────────────────────────

def screen_bulk_entry():
    """
    Spreadsheet-style bulk initiative entry.
    Each row = one initiative. Adds to the existing draft on save.
    """
    import pandas as pd
    from datetime import date as _date

    user  = st.session_state.user
    draft = st.session_state.draft
    entity = draft.get("entity", "")
    rm     = draft.get("reporting_month", "")
    am     = draft.get("activities_month", "") or rm

    st.markdown("## 📊 Bulk Initiative Entry")
    st.caption(
        f"Entity **{entity}** · Filing: **{fmt_month(rm)}** · Reporting Period: **{fmt_month(am)}**  \n"
        "Fill in each row — one initiative per row. Required fields are marked \\*. "
        "Completely blank rows are skipped. Click **← Back** to return to the dashboard without saving."
    )
    st.divider()

    REQUIRED = ["Initiative Name", "Business Component", "Description",
                "Tech Uncertainty", "Start Date", "End Date", "Activities"]

    # ── Initialise the grid dataframe ─────────────────────────────────────────
    if "bulk_df" not in st.session_state:
        # Start with 5 blank rows; user can add more with the + button in the grid
        st.session_state.bulk_df = pd.DataFrame({
            "Initiative Name *":   [""] * 5,
            "Business Component *":[""] * 5,
            "Description *":       [""] * 5,
            "Tech Uncertainty *":  [""] * 5,
            "Start Date *":        [None] * 5,
            "End Date *":          [None] * 5,
            "Activities *":        [""] * 5,
            "Team Member":         [None] * 5,   # single-select dropdown
            "Notes":               [""] * 5,
        })

    # ── Render the grid ───────────────────────────────────────────────────────
    st.caption(
        "💡 **Tips:** Tab to move between cells · Click a Team Member cell for the dropdown "
        "· Click a date cell for the calendar picker · Use the ＋ row button at the bottom to add more rows"
    )

    edited = st.data_editor(
        st.session_state.bulk_df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=False,   # Row numbers visible so initiatives are identifiable
        column_config={
            "Initiative Name *": st.column_config.TextColumn(
                "Initiative Name *", width="medium",
                help="Short, unique name for this initiative",
            ),
            "Business Component *": st.column_config.TextColumn(
                "Business Component *", width="medium",
                help="The business component this R&D work belongs to",
            ),
            "Description *": st.column_config.TextColumn(
                "Initiative Description *", width="large",
                help="What this initiative is trying to accomplish",
            ),
            "Tech Uncertainty *": st.column_config.TextColumn(
                "Technical Uncertainty *", width="large",
                help="What technical question or uncertainty you are trying to resolve",
            ),
            "Start Date *": st.column_config.DateColumn(
                "Start Date *", format="YYYY-MM-DD",
                help="When work on this initiative began",
            ),
            "End Date *": st.column_config.DateColumn(
                "Expected End Date *", format="YYYY-MM-DD",
                help="Expected completion date",
            ),
            "Activities *": st.column_config.TextColumn(
                "Activities *", width="large",
                help="Activities carried out this month to eliminate the technical uncertainty",
            ),
            "Team Member": st.column_config.SelectboxColumn(
                "Team Member", width="medium",
                options=EMPLOYEES,
                help="Primary team member. Add more via the initiative card after saving.",
            ),
            "Notes": st.column_config.TextColumn(
                "Notes", width="medium",
                help="Optional — any additional context, blockers, or upcoming steps",
            ),
        },
        key="bulk_editor",
    )
    st.session_state.bulk_df = edited

    st.write("")

    # ── Validate & save ───────────────────────────────────────────────────────
    b1, b2 = st.columns([1, 1])
    with b1:
        if st.button("← Back to Dashboard", key="bulk_back"):
            st.session_state.screen = "dashboard"
            st.rerun()
    with b2:
        if st.button("Save Initiatives ✓", type="primary", key="bulk_save"):
            errors   = []
            to_add   = []

            for idx, row in edited.iterrows():
                row_num = idx + 1  # 1-based for user-facing messages

                name  = str(row.get("Initiative Name *",  "") or "").strip()
                bc    = str(row.get("Business Component *","") or "").strip()
                desc  = str(row.get("Description *",      "") or "").strip()
                unc   = str(row.get("Tech Uncertainty *", "") or "").strip()
                acts  = str(row.get("Activities *",       "") or "").strip()
                notes = str(row.get("Notes",              "") or "").strip()
                sd    = row.get("Start Date *")
                ed    = row.get("End Date *")
                team  = row.get("Team Member")

                # Completely blank row → skip silently
                text_vals = [name, bc, desc, unc, acts]
                if not any(text_vals) and sd is None and ed is None:
                    continue

                # Partially filled → collect errors
                row_errors = []
                if not name:  row_errors.append("Initiative Name")
                if not bc:    row_errors.append("Business Component")
                if not desc:  row_errors.append("Description")
                if not unc:   row_errors.append("Tech Uncertainty")
                if not acts:  row_errors.append("Activities")
                if sd is None:row_errors.append("Start Date")
                if ed is None:row_errors.append("End Date")

                if row_errors:
                    errors.append(f"Row {row_num} — missing: {', '.join(row_errors)}")
                    continue

                # Convert date objects to strings
                sd_str = sd.strftime("%Y-%m-%d") if hasattr(sd, "strftime") else str(sd)
                ed_str = ed.strftime("%Y-%m-%d") if hasattr(ed, "strftime") else str(ed)

                # Duplicate initiative name check within this batch + existing draft
                existing_names = [i.get("initiative_name","").strip().lower()
                                  for i in draft.get("initiatives", [])]
                batch_names    = [r.get("initiative_name","").strip().lower() for r in to_add]
                if name.lower() in existing_names or name.lower() in batch_names:
                    errors.append(f"Row {row_num} — \"{name}\" already exists in this report. "
                                  "Each initiative needs a unique name.")
                    continue

                init = new_initiative()
                init["initiative_name"]         = name
                init["business_component"]      = bc
                init["initiative_description"]  = desc
                init["tech_uncertainty"]        = unc
                init["start_date"]              = sd_str
                init["expected_end_date"]       = ed_str
                init["activities"]              = acts
                init["team_members"]            = [team] if team else []
                init["notes"]                   = notes
                init["month_yr"]               = am
                to_add.append(init)

            if errors:
                for e in errors:
                    st.error(e)
            elif not to_add:
                st.warning("No initiatives to save — fill in at least one row.")
            else:
                draft["initiatives"] = draft.get("initiatives", []) + to_add
                save_draft(user, draft)
                st.session_state.draft = draft
                st.session_state.pop("bulk_df", None)
                st.session_state.pop("bulk_editor", None)
                n = len(to_add)
                st.success(f"✓ Saved {n} initiative{'s' if n!=1 else ''}.")
                st.session_state.screen = "dashboard"
                st.rerun()


# ── Pathway Screens ────────────────────────────────────────────────────────────

def screen_pathway_select():
    """
    Shown before the wizard when a user edits a carry-over initiative.
    Asks: what happened with this initiative this month?
    Routes to the appropriate lightweight or full form.
    """
    user  = st.session_state.user
    init  = st.session_state.wiz_init
    draft = st.session_state.draft

    st.markdown("## What happened with this initiative this month?")
    st.markdown(
        f"**{init.get('initiative_name', 'Unnamed')}** — *{init.get('business_component', '')}*"
    )
    st.divider()

    existing_pathway = init.get("pathway", "")
    if existing_pathway:
        pathway_labels = {"continuing": "Still working on it",
                          "resolved": "Resolved this month",
                          "new_direction": "New direction"}
        st.caption(f"Previously chose: **{pathway_labels.get(existing_pathway, existing_pathway)}** — select again to change.")
        st.write("")

    c1, c2, c3 = st.columns(3)

    with c1:
        st.markdown("### 🔄 Still working on it")
        st.caption("Lightweight monthly update. Update activities and notes; all other details carry over as-is.")
        if st.button("Select →", key="pw_continuing", use_container_width=True, type="primary"):
            st.session_state.wiz_init["pathway"] = "continuing"
            st.session_state.screen = "continuing"
            st.rerun()

    with c2:
        st.markdown("### ✅ Resolved this month")
        st.caption("Capture the resolution and outcome. Sets a completion date; won't appear in future carry-overs.")
        if st.button("Select →", key="pw_resolved", use_container_width=True):
            st.session_state.wiz_init["pathway"] = "resolved"
            st.session_state.screen = "resolved"
            st.rerun()

    with c3:
        st.markdown("### 🔀 New direction")
        st.caption("The scope or approach has fundamentally changed. Full structured entry — same as adding a new initiative.")
        if st.button("Select →", key="pw_new_direction", use_container_width=True):
            st.session_state.wiz_init["pathway"] = "new_direction"
            st.session_state.wiz_mode = "edit"
            st.session_state.wiz_step = 0
            st.session_state.screen = "wizard"
            st.rerun()

    st.write("")
    if st.button("← Back to dashboard", key="pw_back"):
        st.session_state.screen = "dashboard"
        st.rerun()


def _save_initiative_update(user: str, init: dict):
    """Save updated initiative back into the draft and return to dashboard."""
    draft = st.session_state.draft
    if not init.get("month_yr"):
        init["month_yr"] = draft.get("activities_month") or draft.get("reporting_month", "")
    draft["initiatives"] = [
        i if i["id"] != init["id"] else init
        for i in draft["initiatives"]
    ]
    save_draft(user, draft)
    st.session_state.draft  = draft
    st.session_state.screen = "dashboard"
    st.rerun()


def screen_continuing():
    """Pathway 1 — Still Working On It. Prominent activities, dates + notes; other fields collapsible."""
    user  = st.session_state.user
    init  = dict(st.session_state.wiz_init)

    # Seed the team multiselect key on first render so it pre-fills correctly
    if "cont_team" not in st.session_state:
        st.session_state["cont_team"] = [
            m for m in (init.get("team_members") or []) if m in EMPLOYEES
        ]

    st.markdown(f"## 🔄 Monthly Update")
    st.markdown(f"**{init.get('initiative_name', 'Unnamed')}** — *{init.get('business_component', '')}*")
    st.caption("Update what changed this month. All fields are pre-filled from last month.")
    st.divider()

    activities = st.text_area(
        "What did you work on this month? *",
        value=init.get("activities", ""),
        height=160,
        placeholder="Describe the activities you carried out this month to advance this initiative...",
        key="cont_activities",
    )

    dc1, dc2 = st.columns(2)
    with dc1:
        sd = st.text_input(
            "Start Date (YYYY-MM-DD)",
            value=init.get("start_date") or "",
            key="cont_sd",
        )
    with dc2:
        ed = st.text_input(
            "Expected End Date (YYYY-MM-DD)",
            value=init.get("expected_end_date") or "",
            key="cont_ed",
        )

    notes = st.text_area(
        "Notes (optional)",
        value=init.get("notes", ""),
        height=80,
        placeholder="Any blockers, upcoming steps, or context worth noting...",
        key="cont_notes",
    )

    st.write("")

    with st.expander("✏ Update other details (optional)", expanded=False):
        st.caption("These fields carried over from last month. Update if anything has changed.")
        bc   = st.text_input("Business Component", value=init.get("business_component", ""), key="cont_bc")
        name = st.text_input("Initiative Name",    value=init.get("initiative_name", ""),    key="cont_name")
        desc = st.text_area("Initiative Description", value=init.get("initiative_description", ""), height=100, key="cont_desc")
        unc  = st.text_area("Technical Uncertainty",  value=init.get("tech_uncertainty", ""),        height=100, key="cont_unc")
        team = st.multiselect(
            "Team Members", EMPLOYEES,
            key="cont_team",
        )

    st.write("")
    b1, b2 = st.columns([1, 1])
    with b1:
        if st.button("← Back", key="cont_back"):
            st.session_state.screen = "pathway_select"
            st.rerun()
    with b2:
        if st.button("Save Update ✓", type="primary", key="cont_save"):
            if not activities.strip():
                st.error("Please describe what you worked on this month.")
            else:
                init["activities"]             = activities.strip()
                init["notes"]                  = notes.strip()
                init["start_date"]             = st.session_state.get("cont_sd") or None
                init["expected_end_date"]      = st.session_state.get("cont_ed") or None
                init["business_component"]     = st.session_state.get("cont_bc", init["business_component"])
                init["initiative_name"]        = st.session_state.get("cont_name", init["initiative_name"])
                init["initiative_description"] = st.session_state.get("cont_desc", init["initiative_description"])
                init["tech_uncertainty"]       = st.session_state.get("cont_unc", init["tech_uncertainty"])
                init["team_members"]           = st.session_state.get("cont_team", init.get("team_members", []))
                _save_initiative_update(user, init)


def screen_resolved():
    """Pathway 2 — Resolved This Month. Captures final activities, dates, completion date, and outcome."""
    user  = st.session_state.user
    init  = dict(st.session_state.wiz_init)

    st.markdown(f"## ✅ Resolution — {init.get('initiative_name', 'Unnamed')}")
    st.markdown(f"*{init.get('business_component', '')}*")
    st.caption("Capture what you did to reach resolution and the outcome. This initiative won't appear in next month's carry-over suggestions.")
    st.divider()

    activities = st.text_area(
        "What activities led to resolution this month? *",
        value=init.get("activities", ""),
        height=140,
        placeholder="Describe the final activities that eliminated the technical uncertainty...",
        key="res_activities",
    )

    rd1, rd2 = st.columns(2)
    with rd1:
        sd = st.text_input(
            "Start Date (YYYY-MM-DD)",
            value=init.get("start_date") or "",
            key="res_sd",
        )
    with rd2:
        ed = st.text_input(
            "Expected End Date (YYYY-MM-DD)",
            value=init.get("expected_end_date") or "",
            key="res_ed",
        )

    today_str = date.today().strftime("%Y-%m-%d")
    completion_date = st.text_input(
        "Completion Date (YYYY-MM-DD) *",
        value=init.get("completion_date") or today_str,
        key="res_completion",
    )
    outcome = st.text_area(
        "Outcome / Notes *",
        value=init.get("notes", ""),
        height=120,
        placeholder="Describe the outcome — what was resolved, what was learned, what the result was...",
        key="res_outcome",
    )

    st.write("")
    b1, b2 = st.columns([1, 1])
    with b1:
        if st.button("← Back", key="res_back"):
            st.session_state.screen = "pathway_select"
            st.rerun()
    with b2:
        if st.button("Save Resolution ✓", type="primary", key="res_save"):
            errors = []
            if not activities.strip():
                errors.append("Please describe the activities that led to resolution.")
            if not completion_date.strip():
                errors.append("Please enter a completion date.")
            else:
                try:
                    datetime.strptime(completion_date.strip(), "%Y-%m-%d")
                except ValueError:
                    errors.append("Completion date must be in YYYY-MM-DD format.")
            if not outcome.strip():
                errors.append("Please describe the outcome.")
            if errors:
                for e in errors:
                    st.error(e)
            else:
                init["activities"]      = activities.strip()
                init["notes"]           = outcome.strip()
                init["start_date"]      = st.session_state.get("res_sd") or None
                init["expected_end_date"] = st.session_state.get("res_ed") or None
                init["completion_date"] = completion_date.strip()
                init["resolved"]        = True
                _save_initiative_update(user, init)


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
    elif screen == "bulk_entry":
        screen_bulk_entry()
    elif screen == "entry_picker":
        screen_entry_picker()
    elif screen == "pathway_select":
        screen_pathway_select()
    elif screen == "continuing":
        screen_continuing()
    elif screen == "resolved":
        screen_resolved()
    elif screen == "admin":
        screen_admin()

if __name__ == "__main__":
    main()
