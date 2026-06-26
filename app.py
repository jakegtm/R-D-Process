"""
R&D Monthly Activity Tracker
Run with: streamlit run app.py
"""

import streamlit as st
import json
from datetime import datetime, date
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

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

ENTITIES = ["107", "108", "109", "110"]

STATUS_LABELS = {
    "in-progress": "🟡 In Progress",
    "submitted":   "🔵 Submitted for Review",
    "approved":    "🟢 Approved",
    "rejected":    "🔴 Returned — Needs Revision",
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

def fmt_month(m: str) -> str:
    if not m:
        return ""
    y, mo = m.split("-")
    return datetime(int(y), int(mo), 1).strftime("%B %Y")

def fmt_month_tab(m: str) -> str:
    """Short form for Excel sheet tab names, e.g. 'May 26'"""
    if not m:
        return ""
    y, mo = m.split("-")
    return datetime(int(y), int(mo), 1).strftime("%b %Y")

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
    # Per-user month index
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

def get_combos(all_data: dict) -> list[tuple[str, str]]:
    """
    Returns sorted list of unique (entity, reporting_month) tuples
    found across all submissions.
    """
    combos: set[tuple[str, str]] = set()
    for user, months in all_data.items():
        for month, sub in months.items():
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
        if existing and existing.get("entity") == source_entity:
            continue

        initiatives = src_sub.get("initiatives") or []
        if not initiatives:
            continue

        new_sub: dict = {
            "initiatives":     [carryover_initiative(i) for i in initiatives],
            "status":          "in-progress",
            "entity":          source_entity,
            "reporting_month": target_month,
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
    }

def carryover_initiative(src: dict) -> dict:
    init = new_initiative()
    for f in CARRYOVER_FIELDS:
        init[f] = src.get(f, init[f])
    init["carry_over"] = True
    return init

def empty_draft() -> dict:
    return {
        "initiatives":     [],
        "status":          "in-progress",
        "entity":          "",
        "reporting_month": "",
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
    # (header_label,                                            field_key,      width, is_team)
    ("Month/Yr",                                               "_month",        14,   False),
    ("Business Component",                                     "business_component", 35, False),
    ("Initiative Name",                                        "initiative_name",    19, False),
    ("Initiative Description",                                 "initiative_description", 38, False),
    ("Tech Uncertainty",                                       "tech_uncertainty",   64, False),
    ("Start Date",                                             "start_date",    16,   False),
    ("Expected End Date",                                      "expected_end_date", 16, False),
    ("Activities to Eliminate Technical Uncertainty",          "activities",    60,   False),
    ("Team Members",                                           "team_members",  49,   True),
    ("Notes",                                                  "notes",         53,   False),
    ("Completion Date",                                        "_completion",   22,   False),
    ("Status",                                                 "_status",       22,   False),
]

def _write_sheet(ws, rows_data: list[dict], subtitle: str):
    """
    Write title row, subtitle, spacer, header row, then data rows
    into worksheet ws.  rows_data is a list of dicts with keys from _COL_DEF
    plus 'user' for consolidated sheets.
    """
    border     = _thin_border()
    fill_green  = PatternFill("solid", fgColor=GREEN)
    fill_beige  = PatternFill("solid", fgColor=BEIGE)
    fill_yellow = PatternFill("solid", fgColor=YELLOW)

    has_user_col = any("user" in r for r in rows_data) if rows_data else False
    cols = ([("User", "user", 18, False)] if has_user_col else []) + list(_COL_DEF)

    # Row 1 — title
    ws.row_dimensions[1].height = 18.3
    c = ws.cell(1, 1, "Monthly R&D Tracking Template")
    c.font = Font(name="Arial Narrow", bold=True, size=14, color=DARK)

    # Row 2 — subtitle
    ws.row_dimensions[2].height = 15
    c = ws.cell(2, 1, subtitle)
    c.font = Font(name="Arial Narrow", size=10, color=GRAY, italic=True)

    # Row 3 — spacer
    ws.row_dimensions[3].height = 8

    # Row 4 — column headers
    ws.row_dimensions[4].height = 30
    for ci, (hdr, _, width, _is_team) in enumerate(cols, 1):
        cell = ws.cell(4, ci, hdr)
        cell.fill      = fill_green
        cell.font      = _hdr_font(12)
        cell.border    = border
        cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
        ws.column_dimensions[get_column_letter(ci)].width = width

    # Data rows
    for rn, row in enumerate(rows_data, 5):
        ws.row_dimensions[rn].height = 45
        for ci, (_, key, _, is_team) in enumerate(cols, 1):
            val = row.get(key, "")
            cell = ws.cell(rn, ci, val)
            cell.font      = _data_font()
            cell.border    = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.fill      = fill_yellow if is_team else fill_beige

    ws.freeze_panes = "A5"


def _sub_to_rows(username: str, sub: dict, include_user: bool) -> list[dict]:
    """Convert a submission dict into a list of row dicts for _write_sheet."""
    rm     = sub.get("reporting_month", "")
    status = STATUS_LABELS.get(sub.get("status", ""), sub.get("status", ""))
    rows   = []
    for init in sub.get("initiatives") or []:
        row = {
            "_month":  fmt_month(rm),
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
    ws.title = fmt_month_tab(sub.get("reporting_month", ""))
    entity = sub.get("entity", "")
    rm     = sub.get("reporting_month", "")
    subtitle = f"Submitted by: {username}  |  Entity: {entity}  |  Period: {fmt_month(rm)}"
    rows = _sub_to_rows(username, sub, include_user=False)
    _write_sheet(ws, rows, subtitle)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_excel_consolidated(
    all_data: dict,
    selected_combos: list[tuple[str, str]],
    submitted_only: bool = False,
) -> bytes:
    """
    Multi-sheet consolidated export.
    One sheet per (entity, reporting_month) period.
    Sheet name: "107 - May 2026"
    """
    wb = Workbook()
    wb.remove(wb.active)   # remove blank default sheet

    for (entity, rm) in selected_combos:
        sheet_name = f"{entity} - {fmt_month_tab(rm)}"[:31]
        ws = wb.create_sheet(title=sheet_name)
        subtitle = (
            f"Entity: {entity}  |  Period: {fmt_month(rm)}"
            + ("  |  Submitted & Approved only" if submitted_only else "")
        )
        rows = []
        for username, months in all_data.items():
            sub = months.get(rm)
            if not sub:
                continue
            if sub.get("entity") != entity:
                continue
            if submitted_only and sub.get("status") not in ("submitted", "approved"):
                continue
            rows.extend(_sub_to_rows(username, sub, include_user=True))

        _write_sheet(ws, rows, subtitle)

    # If nothing matched, add a placeholder sheet
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

def screen_dashboard():
    user      = st.session_state.user
    draft     = st.session_state.draft
    submitted = draft["status"] != "in-progress"

    # Header
    c1, c2 = st.columns([5, 1])
    with c1:
        period = fmt_month(draft.get("reporting_month")) or "— Set reporting month below"
        st.markdown(f"## 🔬 R&D Tracker — {period}")
        st.caption(f"Signed in as **{user}**")
    with c2:
        st.write("")
        if st.button("Sign Out"):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.rerun()

    st.divider()

    # ── Report Setup ─────────────────────────────────────────────────────────
    setup_complete = bool(draft.get("entity") and draft.get("reporting_month"))
    with st.expander(
        "⚙ Report Setup"
        + (f"  —  Entity {draft['entity']}  ·  {fmt_month(draft['reporting_month'])}"
           if setup_complete else "  ⚠ Please select entity and reporting month first"),
        expanded=not setup_complete,
    ):
        st.caption(
            "Export filename: "
            "`{entity}_Group_{month}_{year}_Monthly_R_D_Tracking_Template.xlsx`"
        )
        su1, su2 = st.columns(2)
        with su1:
            eidx         = ENTITIES.index(draft["entity"]) if draft.get("entity") in ENTITIES else 0
            chosen_entity = st.selectbox("Entity", ENTITIES, index=eidx, key="su_entity")
        with su2:
            mlist    = available_months()
            def_rm   = mlist[1] if len(mlist) > 1 else mlist[0]
            cur_rm   = draft.get("reporting_month") or def_rm
            rm_idx   = mlist.index(cur_rm) if cur_rm in mlist else 0
            chosen_rm = st.selectbox(
                "Reporting Month", mlist, index=rm_idx,
                format_func=fmt_month, key="su_rm",
            )

        entity_changed = chosen_entity != draft.get("entity")
        month_changed  = chosen_rm     != draft.get("reporting_month")

        if entity_changed or month_changed:
            if month_changed:
                # Load existing submission for that month if it exists
                existing = load_submission(user, chosen_rm)
                if existing:
                    draft = existing
                else:
                    draft = empty_draft()
            draft["entity"]          = chosen_entity
            draft["reporting_month"] = chosen_rm
            save_draft(user, draft)
            st.session_state.draft = draft
            st.rerun()

        if setup_complete:
            fname = export_filename(draft["entity"], draft["reporting_month"])
            st.success(f"Export will be named: **{fname}**")

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
            st.error("Your report was returned. Please revise and resubmit.")
    with c2:
        st.metric("Initiatives", len(draft["initiatives"]))
    with c3:
        if st.session_state.is_admin and st.button("Admin View →"):
            st.session_state.screen = "admin"
            st.rerun()

    # ── Initiatives list ──────────────────────────────────────────────────────
    c1, c2 = st.columns([5, 1])
    with c1:
        st.subheader(f"{fmt_month(draft['reporting_month'])} Initiatives")
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
        for init in draft["initiatives"]:
            iid = init["id"]
            with st.expander(
                f"{'↩ ' if init.get('carry_over') else ''}"
                f"{init.get('initiative_name','Unnamed')} — {init.get('business_component','')}",
                expanded=True,
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
                istatus = init.get("initiative_status", "active")
                if istatus == "returned":
                    ret_ts = init.get("returned_at")
                    ret_str = f" on {datetime.fromtimestamp(ret_ts/1000).strftime('%b %d %H:%M')}" if ret_ts else ""
                    st.warning(f"⚠ Admin returned this initiative for revision{ret_str}. Edit and resubmit.")
                elif istatus == "approved":
                    appr_ts = init.get("approved_at")
                    appr_str = f" on {datetime.fromtimestamp(appr_ts/1000).strftime('%b %d %H:%M')}" if appr_ts else ""
                    st.success(f"✓ Approved by admin{appr_str}.")

                st.write("")
                c1, c2, c3, c4 = st.columns([2.5, 0.8, 0.8, 0.8])
                with c2:
                    if st.button("✏ Edit", key=f"edit_{iid}"):
                        st.session_state.wiz_init = dict(init)
                        st.session_state.wiz_step = 0
                        st.session_state.wiz_mode = "edit"
                        st.session_state.screen   = "wizard"
                        st.rerun()
                with c3:
                    if st.button("📤 Submit", key=f"usub_{iid}", disabled=submitted,
                                 help="Submit your full report for review"):
                        draft["status"]       = "submitted"
                        draft["submitted_at"] = int(datetime.now().timestamp()*1000)
                        save_draft(user, draft)
                        st.session_state.draft = draft
                        st.success("Report submitted!")
                        st.rerun()
                with c4:
                    if st.button("🗑 Delete", key=f"del_{iid}"):
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
            st.markdown("#### Ready to submit?")
            st.caption(
                f"Sends your {len(draft['initiatives'])} initiative"
                f"{'s' if len(draft['initiatives'])!=1 else ''} to the Oversight Lead for review."
            )
            if st.button("Submit for Review ✓", type="primary"):
                draft["status"]       = "submitted"
                draft["submitted_at"] = int(datetime.now().timestamp()*1000)
                save_draft(user, draft)
                st.session_state.draft = draft
                st.success("Report submitted!")
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

    # Load everything, grouped by (entity, reporting_month)
    all_data = load_all_submissions()
    combos   = get_combos(all_data)   # [(entity, rm), ...]

    # ── Summary stats ────────────────────────────────────────────────────────
    all_subs  = [sub for ud in all_data.values() for sub in ud.values()]
    n_users   = len(all_data)
    n_sub     = sum(1 for s in all_subs if s.get("status") in ("submitted","approved"))
    n_appr    = sum(1 for s in all_subs if s.get("status") == "approved")
    n_inits   = sum(len(s.get("initiatives") or []) for s in all_subs)

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Team Members",     n_users)
    c2.metric("Periods/Entities", len(combos))
    c3.metric("Submitted",        n_sub)
    c4.metric("Approved",         n_appr)
    c5.metric("Total Initiatives",n_inits)

    st.divider()

    # ── Combo selector + consolidated export ─────────────────────────────────
    st.subheader("Consolidated Export")

    if not combos:
        st.info("No submissions found yet.")
    else:
        combo_labels  = {f"{e} — {fmt_month(rm)}": (e, rm) for e, rm in combos}
        all_labels    = list(combo_labels.keys())

        sel_all = st.checkbox("Select all periods/entities", value=True, key="sel_all_periods")
        if sel_all:
            chosen_labels = all_labels
        else:
            chosen_labels = st.multiselect(
                "Choose which periods/entities to include:",
                all_labels,
                default=all_labels,
            )

        chosen_combos = [combo_labels[l] for l in chosen_labels]

        submitted_only = st.checkbox("Submitted & Approved only (exclude In Progress / Returned)", value=True)

        if chosen_combos:
            tab_preview = ", ".join(f'"{e} - {fmt_month_tab(rm)}"' for e,rm in chosen_combos[:4])
            if len(chosen_combos) > 4:
                tab_preview += f" +{len(chosen_combos)-4} more"
            st.caption(f"Excel tabs: {tab_preview}")

        ex1, ex2 = st.columns(2)
        with ex1:
            if chosen_combos:
                xlsx = build_excel_consolidated(all_data, chosen_combos, submitted_only=False)
                # Use first combo for filename, or "Consolidated"
                fe, frm = chosen_combos[0]
                tag = "Consolidated_All" if len(chosen_combos) > 1 else "All"
                st.download_button(
                    f"↓ Download — All Statuses ({len(chosen_combos)} period{'s' if len(chosen_combos)!=1 else ''})",
                    data=xlsx,
                    file_name=export_filename(fe, frm, tag),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        with ex2:
            if chosen_combos:
                xlsx_sub = build_excel_consolidated(all_data, chosen_combos, submitted_only=True)
                tag2 = "Consolidated_Submitted" if len(chosen_combos) > 1 else "Submitted"
                st.download_button(
                    f"↓ Download — Submitted Only ({len(chosen_combos)} period{'s' if len(chosen_combos)!=1 else ''})",
                    data=xlsx_sub,
                    file_name=export_filename(fe, frm, tag2),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

    st.divider()

    # ── Entity Rollover ───────────────────────────────────────────────────────
    st.subheader("🔄 Entity Rollover")
    st.caption(
        "Copy all active initiatives from one entity/month period into a new period. "
        "Only initiatives whose expected end date hasn't passed will be carried. "
        "Users who already have a submission for the target period won't be overwritten."
    )

    if not combos:
        st.info("No submissions to roll over yet.")
    else:
        combo_labels_ro = {f"{e} — {fmt_month(rm)}": (e, rm) for e, rm in combos}
        ro1, ro2 = st.columns(2)

        with ro1:
            src_label = st.selectbox(
                "Source (roll FROM)",
                list(combo_labels_ro.keys()),
                key="ro_src",
            )
        src_entity, src_month = combo_labels_ro[src_label]

        with ro2:
            mlist_ro   = available_months()
            def_ro_idx = 0  # newest month as default target
            chosen_target = st.selectbox(
                "Target (roll TO)",
                mlist_ro,
                index=def_ro_idx,
                format_func=fmt_month,
                key="ro_target",
            )

        # Preview what would happen
        preview_users = []
        for username, months in all_data.items():
            sub = months.get(src_month)
            if not sub or sub.get("entity") != src_entity:
                continue
            existing = load_submission(username, chosen_target)
            if existing and existing.get("entity") == src_entity:
                continue
            inits = sub.get("initiatives") or []
            if inits:
                inits_to_carry = [i.get("initiative_name", "Unnamed") for i in inits]
                preview_users.append((username, inits_to_carry))

        if src_month == chosen_target:
            st.warning("Source and target are the same period — choose a different target month.")
        elif not preview_users:
            st.info("No users have active initiatives to roll over into this target period.")
        else:
            st.markdown(
                f"**Preview:** Rolling **{src_entity} — {fmt_month(src_month)}** → "
                f"**{fmt_month(chosen_target)}** will create in-progress drafts for:"
            )
            for uname, init_names in preview_users:
                st.markdown(
                    f"&nbsp;&nbsp;• **{uname}** — "
                    + ", ".join(f"*{n}*" for n in init_names),
                    unsafe_allow_html=True,
                )

            if st.button(
                f"🔄 Roll Over {src_entity} — {fmt_month(src_month)} → {fmt_month(chosen_target)}",
                type="primary",
                key="do_rollover",
            ):
                rolled = rollover_entity(all_data, src_entity, src_month, chosen_target)
                if rolled:
                    st.success(
                        f"✓ Rolled over for: {', '.join(rolled)}. "
                        f"They'll see a pre-filled draft when they sign in and select {fmt_month(chosen_target)}."
                    )
                    # Reload data so the new combos appear
                    all_data = load_all_submissions()
                    combos   = get_combos(all_data)
                    st.rerun()
                else:
                    st.info("Nothing to roll over (all users may already have a submission for that period).")

    st.divider()

    # ── Submissions by combo ──────────────────────────────────────────────────
    if not combos:
        return

    st.subheader("Team Submissions by Period")

    for (entity, rm) in combos:
        st.markdown(f"#### Entity {entity} — {fmt_month(rm)}")

        # Find all users who have a submission for this combo
        combo_rows = []
        for username, months in all_data.items():
            sub = months.get(rm)
            if not sub or sub.get("entity") != entity:
                continue
            combo_rows.append((username, sub))

        if not combo_rows:
            st.caption("No submissions.")
            continue

        for username, sub in combo_rows:
            status = sub.get("status", "not-started")
            inits  = sub.get("initiatives") or []
            icon   = {"approved":"✅","submitted":"🔵","in-progress":"🟡",
                      "rejected":"🔴","not-started":"⚪"}.get(status,"⚪")

            with st.expander(
                f"{icon}  {username}   —  {STATUS_LABELS.get(status,'—')}  "
                f"({len(inits)} initiative{'s' if len(inits)!=1 else ''})",
                expanded=(status == "submitted"),
            ):
                # ── Report-level info + export ────────────────────────────
                h1, h2 = st.columns([4, 1])
                with h1:
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

                # ── Report-level approve / return ─────────────────────────
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
                        if st.button("↩ Return All", key=f"rej_rpt_{entity}_{rm}_{username}"):
                            now = int(datetime.now().timestamp()*1000)
                            sub["status"]      = "rejected"
                            sub["rejected_at"] = now
                            save_draft(username, sub)
                            st.rerun()

                st.divider()

                # ── Per-initiative actions ────────────────────────────────
                if not inits:
                    st.caption("No initiatives.")
                else:
                    for init in inits:
                        iid     = init["id"]
                        istatus = init.get("initiative_status","active")
                        iname   = init.get("initiative_name","Unnamed")
                        ico     = {"approved":"✅","returned":"↩","active":"🔵"}.get(istatus,"🔵")

                        ic1, ic2, ic3, ic4 = st.columns([3, 0.8, 0.8, 0.8])
                        with ic1:
                            co = "↩ " if init.get("carry_over") else ""
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
                                st.caption(f"   Returned {ts}")

                        with ic2:
                            if istatus != "approved":
                                if st.button("✓ Accept", key=f"appr_i_{entity}_{rm}_{username}_{iid}",
                                             type="primary"):
                                    now = int(datetime.now().timestamp()*1000)
                                    init["initiative_status"] = "approved"
                                    init["approved_at"]       = now
                                    init.pop("returned_at", None)
                                    # If all initiatives approved, approve the report too
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
                                if st.button("↩ Return", key=f"ret_i_{entity}_{rm}_{username}_{iid}"):
                                    now = int(datetime.now().timestamp()*1000)
                                    init["initiative_status"] = "returned"
                                    init["returned_at"]       = now
                                    init.pop("approved_at", None)
                                    # Mark report as rejected so user is notified
                                    sub["status"] = "rejected"
                                    save_draft(username, sub)
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

        st.write("")  # spacing between combos


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
