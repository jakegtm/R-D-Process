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

STATUS_LABELS = {
    "in-progress": "🟡 In Progress",
    "submitted":   "🔵 Submitted for Review",
    "approved":    "🟢 Approved",
    "rejected":    "🔴 Returned — Needs Revision",
    "not-started": "⚪ Not Started",
}

# ── Wizard Steps — exactly the columns in the Excel template ─────────────────
#
#  Column order: Month/Yr | Business Component | Initiative Name |
#  Initiative Description | Tech Uncertainty | Start Date | Expected End Date |
#  Activities to Eliminate Technical Uncertainty | Team Members | Notes
#
#  Month/Yr is auto-filled from the current month — not asked as a question.
#
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

# Fields that pre-fill when carrying over from the previous month
CARRYOVER_FIELDS = {
    "business_component", "initiative_name", "initiative_description",
    "tech_uncertainty", "start_date", "expected_end_date", "team_members",
}
# Fields that reset each month (must be filled in fresh)
RESET_FIELDS = {"activities", "notes"}


# ── Data Helpers ──────────────────────────────────────────────────────────────

def cur_month() -> str:
    return datetime.now().strftime("%Y-%m")

def prev_month() -> str:
    n = datetime.now()
    if n.month == 1:
        return f"{n.year - 1}-12"
    return f"{n.year}-{str(n.month - 1).zfill(2)}"

def fmt_month(m: str) -> str:
    if not m:
        return ""
    y, mo = m.split("-")
    return datetime(int(y), int(mo), 1).strftime("%B %Y")

def data_path(username: str, month: str) -> Path:
    safe = username.replace(" ", "_").replace("/", "_")
    return DATA_DIR / f"{safe}_{month}.json"

def registry_path() -> Path:
    return DATA_DIR / "registry.json"

def load_registry() -> list:
    p = registry_path()
    return json.loads(p.read_text()) if p.exists() else []

def save_registry(users: list):
    registry_path().write_text(json.dumps(users, indent=2))

def load_submission(username: str, month: str) -> dict | None:
    p = data_path(username, month)
    return json.loads(p.read_text()) if p.exists() else None

def save_submission(username: str, month: str, data: dict):
    data_path(username, month).write_text(json.dumps(data, indent=2))
    reg = load_registry()
    if username not in reg:
        reg.append(username)
        save_registry(reg)

def new_initiative() -> dict:
    return {
        "id": f"{int(datetime.now().timestamp()*1000)}",
        # Matches Excel columns exactly
        "business_component":    "",
        "initiative_name":       "",
        "initiative_description":"",
        "tech_uncertainty":      "",
        "start_date":            None,
        "expected_end_date":     None,
        "activities":            "",
        "team_members":          [],
        "notes":                 "",
        # Meta
        "carry_over": False,
    }

def carryover_initiative(src: dict) -> dict:
    init = new_initiative()
    init["id"] = f"{int(datetime.now().timestamp()*1000)}"
    for f in CARRYOVER_FIELDS:
        init[f] = src.get(f, init[f])
    init["carry_over"] = True
    return init


# ── Excel Export ──────────────────────────────────────────────────────────────
#
# Colors extracted directly from the original template file:
#   Header fill  : #9BBB59  (theme accent3 — the green header row)
#   Header font  : white #FFFFFF, bold, Arial Narrow 12pt
#   Data fill    : #EEECE1  (theme lt2 — beige, used on most input cells)
#   Team Members : #DED900  (yellow — "Selection" cells in the original)
#   Title font   : bold, Arial Narrow 14pt
#   Borders      : thin, all sides

def build_excel(
    all_data: dict,
    only_user: str | None = None,
    submitted_only: bool = False,
    label: str = "",
) -> bytes:
    """
    Build an Excel workbook styled to match the original R&D tracking template.

    Parameters
    ----------
    all_data       : { username: { month_str: submission_dict } }
    only_user      : if set, only include rows for that user
    submitted_only : if True, skip initiatives from non-submitted reports
    label          : extra text shown in the subtitle row (e.g. "Consolidated — Submitted Only")
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "R&D Tracking"

    # ── Style constants ──────────────────────────────────────────────────────
    GREEN     = "9BBB59"   # header row fill  (template's accent3)
    BEIGE     = "EEECE1"   # data row fill    (template's lt2 / Input cells)
    YELLOW    = "DED900"   # Team Members col (template's Selection cells)
    WHITE     = "FFFFFF"
    DARK      = "1F1F1F"

    fill_green  = PatternFill("solid", fgColor=GREEN)
    fill_beige  = PatternFill("solid", fgColor=BEIGE)
    fill_yellow = PatternFill("solid", fgColor=YELLOW)

    thin = Side(style="thin", color="A0A0A0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def hdr_font(sz=12):
        return Font(name="Arial Narrow", bold=True, color=WHITE, size=sz)

    def data_font(bold=False):
        return Font(name="Arial Narrow", bold=bold, color=DARK, size=11)

    # ── Column layout ────────────────────────────────────────────────────────
    # Individual export: mirrors original template columns (A–J) + Status col
    # Consolidated:      prepends "User" column, then same layout
    is_consolidated = only_user is None

    if is_consolidated:
        headers = [
            "User",                                           # A  (extra for consolidated)
            "Month/Yr",                                       # B
            "Business Component",                             # C
            "Initiative Name",                                # D
            "Initiative Description",                         # E
            "Tech Uncertainty",                               # F
            "Start Date",                                     # G
            "Expected End Date",                              # H
            "Activities to Eliminate Technical Uncertainty",  # I
            "Team Members",                                   # J
            "Notes",                                          # K
            "Status",                                         # L
        ]
        col_widths  = [18, 14, 35, 19, 38, 64, 16, 16, 60, 49, 53, 22]
        team_col    = 10   # J — yellow
        status_col  = 12   # L
    else:
        headers = [
            "Month/Yr",                                       # A
            "Business Component",                             # B
            "Initiative Name",                                # C
            "Initiative Description",                         # D
            "Tech Uncertainty",                               # E
            "Start Date",                                     # F
            "Expected End Date",                              # G
            "Activities to Eliminate Technical Uncertainty",  # H
            "Team Members",                                   # I
            "Notes",                                          # J
            "Status",                                         # K
        ]
        col_widths  = [14, 35, 19, 38, 64, 16, 16, 60, 49, 53, 22]
        team_col    = 9    # I — yellow
        status_col  = 11   # K

    ncols = len(headers)

    # ── Row 1 — Title ────────────────────────────────────────────────────────
    ws.row_dimensions[1].height = 18.3
    ws.cell(row=1, column=1, value="Monthly R&D Tracking Template").font = Font(
        name="Arial Narrow", bold=True, size=14, color=DARK
    )

    # ── Row 2 — Subtitle / export info ──────────────────────────────────────
    ws.row_dimensions[2].height = 15
    subtitle = label or (f"Exported: {datetime.now().strftime('%B %Y')}")
    ws.cell(row=2, column=1, value=subtitle).font = Font(
        name="Arial Narrow", size=10, color="666666", italic=True
    )

    # ── Row 3 — Spacer ───────────────────────────────────────────────────────
    ws.row_dimensions[3].height = 8

    # ── Row 4 — Column headers ───────────────────────────────────────────────
    ws.row_dimensions[4].height = 30
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=4, column=ci, value=h)
        cell.fill      = fill_green
        cell.font      = hdr_font(12)
        cell.border    = border
        cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)

    # ── Column widths ────────────────────────────────────────────────────────
    for ci, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    # ── Data rows ─────────────────────────────────────────────────────────────
    row_num = 5
    for username, months in all_data.items():
        if only_user and username != only_user:
            continue
        for month, sub in months.items():
            if not sub or not sub.get("initiatives"):
                continue
            if submitted_only and sub.get("status") not in ("submitted", "approved"):
                continue
            for init in sub["initiatives"]:
                if is_consolidated:
                    vals = [
                        username,
                        fmt_month(month),
                        init.get("business_component",    ""),
                        init.get("initiative_name",        ""),
                        init.get("initiative_description", ""),
                        init.get("tech_uncertainty",       ""),
                        str(init.get("start_date")        or ""),
                        str(init.get("expected_end_date") or ""),
                        init.get("activities",             ""),
                        ", ".join(init.get("team_members") or []),
                        init.get("notes",                  ""),
                        STATUS_LABELS.get(sub.get("status",""), sub.get("status","")),
                    ]
                else:
                    vals = [
                        fmt_month(month),
                        init.get("business_component",    ""),
                        init.get("initiative_name",        ""),
                        init.get("initiative_description", ""),
                        init.get("tech_uncertainty",       ""),
                        str(init.get("start_date")        or ""),
                        str(init.get("expected_end_date") or ""),
                        init.get("activities",             ""),
                        ", ".join(init.get("team_members") or []),
                        init.get("notes",                  ""),
                        STATUS_LABELS.get(sub.get("status",""), sub.get("status","")),
                    ]

                ws.row_dimensions[row_num].height = 45
                for ci, val in enumerate(vals, 1):
                    cell = ws.cell(row=row_num, column=ci, value=val)
                    cell.font      = data_font()
                    cell.border    = border
                    cell.alignment = Alignment(vertical="top", wrap_text=True)
                    # Column-specific fill colors matching the template
                    if ci == team_col:
                        cell.fill = fill_yellow   # Team Members → yellow
                    else:
                        cell.fill = fill_beige    # Everything else → beige
                row_num += 1

    # ── Freeze header row ────────────────────────────────────────────────────
    ws.freeze_panes = f"A5"

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

    .carryover-banner {
        background:#fffbeb; border:1.5px solid #fcd34d;
        border-radius:10px; padding:10px 16px; margin-bottom:14px; font-size:13px; color:#92600a;
    }
    .prefilled-banner {
        background:#f0f9ff; border:1.5px solid #93c5fd;
        border-radius:10px; padding:10px 16px; margin-bottom:14px; font-size:13px; color:#1e40af;
    }
    .delete-confirm {
        background:#fff1f2; border:1.5px solid #fda4af;
        border-radius:10px; padding:12px 16px; margin-top:8px;
    }
    </style>
    """, unsafe_allow_html=True)


# ── Helpers ───────────────────────────────────────────────────────────────────

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
        "screen":       "login",
        "user":         None,
        "is_admin":     False,
        "draft":        {"initiatives": [], "status": "in-progress"},
        "wiz_init":     None,
        "wiz_step":     0,
        "wiz_mode":     "new",
        "confirm_del":  None,   # id of initiative pending delete confirmation
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
                if is_admin:
                    st.session_state.screen = "admin"
                else:
                    existing = load_submission(name, cur_month())
                    st.session_state.draft  = existing or {"initiatives": [], "status": "in-progress"}
                    st.session_state.screen = "dashboard"
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
    cm        = cur_month()
    pm        = prev_month()

    prev_sub   = load_submission(user, pm)
    today      = date.today()
    prev_inits = [
        i for i in (prev_sub or {}).get("initiatives", [])
        if not i.get("expected_end_date")
        or datetime.strptime(str(i["expected_end_date"]), "%Y-%m-%d").date() >= today
    ] if prev_sub else []

    # Header
    c1, c2 = st.columns([5, 1])
    with c1:
        st.markdown(f"## 🔬 R&D Tracker — {fmt_month(cm)}")
        st.caption(f"Signed in as **{user}**")
    with c2:
        st.write("")
        if st.button("Sign Out"):
            for k in ["screen","user","is_admin","draft","wiz_init","wiz_step","wiz_mode","confirm_del"]:
                st.session_state.pop(k, None)
            st.rerun()

    st.divider()

    # Status strip
    c1, c2, c3 = st.columns([4, 1, 1])
    with c1:
        st.markdown(f"**Report Status:** {badge_html(draft['status'])}", unsafe_allow_html=True)
        if draft.get("submitted_at"):
            ts = datetime.fromtimestamp(draft["submitted_at"] / 1000).strftime("%b %d, %Y %H:%M")
            st.caption(f"Submitted {ts}")
        if draft["status"] == "rejected":
            st.error("Your report was returned. Please revise and resubmit.")
    with c2:
        st.metric("Initiatives", len(draft["initiatives"]))
    with c3:
        if st.session_state.is_admin:
            if st.button("Admin View →"):
                st.session_state.screen = "admin"
                st.rerun()

    # Carry-over prompt
    if not submitted and prev_inits and not draft["initiatives"]:
        st.markdown(f"""
        <div class="carryover-banner">
            <strong>📋 Ongoing initiatives from {fmt_month(pm)}</strong><br>
            You had <strong>{len(prev_inits)}</strong> active initiative{"s" if len(prev_inits)!=1 else ""} last month.
            Carry any forward — key details will be pre-filled; you only update activities and notes.
        </div>
        """, unsafe_allow_html=True)
        cols = st.columns(min(len(prev_inits), 4))
        for idx, pi in enumerate(prev_inits):
            with cols[idx % 4]:
                if st.button(f"↩ {pi.get('initiative_name','Unnamed')}", key=f"co_{pi['id']}"):
                    st.session_state.wiz_init = carryover_initiative(pi)
                    st.session_state.wiz_step = 0
                    st.session_state.wiz_mode = "carryover"
                    st.session_state.screen   = "wizard"
                    st.rerun()

    st.divider()

    # Initiatives list header
    c1, c2 = st.columns([5, 1])
    with c1:
        st.subheader(f"{fmt_month(cm)} Initiatives")
    with c2:
        if not submitted:
            if st.button("＋ Add Initiative", type="primary"):
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
                # Info
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

                st.write("")

                # Action buttons — Edit and Delete always visible
                c1, c2, c3 = st.columns([3, 0.8, 0.8])
                with c2:
                    if not submitted:
                        if st.button("✏ Edit", key=f"edit_{iid}"):
                            st.session_state.wiz_init = dict(init)
                            st.session_state.wiz_step = 0
                            st.session_state.wiz_mode = "edit"
                            st.session_state.screen   = "wizard"
                            st.rerun()
                with c3:
                    # Delete works regardless of submitted status
                    if st.button("🗑 Delete", key=f"del_{iid}"):
                        st.session_state.confirm_del = iid
                        st.rerun()

                # Inline confirmation — only shown for this card
                if st.session_state.get("confirm_del") == iid:
                    st.markdown('<div class="delete-confirm">', unsafe_allow_html=True)
                    st.warning(
                        f"Delete **{init.get('initiative_name','this initiative')}**? "
                        "This cannot be undone."
                    )
                    ca, cb = st.columns(2)
                    with ca:
                        if st.button("Yes, delete it", key=f"conf_yes_{iid}", type="primary"):
                            draft["initiatives"] = [
                                i for i in draft["initiatives"] if i["id"] != iid
                            ]
                            # If they delete after submitting, reopen the report for editing
                            if submitted:
                                draft["status"] = "in-progress"
                                draft.pop("submitted_at", None)
                            save_submission(user, cm, draft)
                            st.session_state.draft      = draft
                            st.session_state.confirm_del = None
                            st.success("Initiative deleted.")
                            st.rerun()
                    with cb:
                        if st.button("Cancel", key=f"conf_no_{iid}"):
                            st.session_state.confirm_del = None
                            st.rerun()
                    st.markdown('</div>', unsafe_allow_html=True)

    # Submit + Export
    st.divider()
    c1, c2 = st.columns([3, 1])
    with c1:
        if not submitted and draft["initiatives"]:
            st.markdown("#### Ready to submit?")
            st.caption(
                f"This will send your {len(draft['initiatives'])} initiative"
                f"{'s' if len(draft['initiatives'])!=1 else ''} to the Oversight Lead for review."
            )
            if st.button("Submit for Review ✓", type="primary"):
                draft["status"]       = "submitted"
                draft["submitted_at"] = int(datetime.now().timestamp() * 1000)
                save_submission(user, cm, draft)
                st.session_state.draft = draft
                st.success("Report submitted!")
                st.rerun()
    with c2:
        all_data   = {user: {cm: draft}}
        xlsx_bytes = build_excel(
            all_data,
            only_user=user,
            label=f"Submitted by: {user}  |  Period: {fmt_month(cm)}",
        )
        st.download_button(
            "↓ Download My Report",
            data=xlsx_bytes,
            file_name=f"RD_{user.replace(' ','_')}_{cm}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ── Wizard ────────────────────────────────────────────────────────────────────

def screen_wizard():
    user  = st.session_state.user
    init  = st.session_state.wiz_init
    step  = st.session_state.wiz_step
    mode  = st.session_state.wiz_mode
    s     = WIZARD_STEPS[step]
    total = len(WIZARD_STEPS)
    pct   = int((step + 1) / total * 100)

    is_carryover = mode == "carryover"
    is_prefilled = is_carryover and s["field"] in CARRYOVER_FIELDS

    # Header
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

    # Context banners
    if is_carryover and not is_prefilled:
        st.markdown(f"""
        <div class="carryover-banner">
            <strong>Updating for {fmt_month(cur_month())}:</strong>
            {init.get("initiative_name","this initiative")}
        </div>
        """, unsafe_allow_html=True)
    elif is_prefilled:
        st.markdown("""
        <div class="prefilled-banner">
            Pre-filled from last month — confirm or update before continuing.
        </div>
        """, unsafe_allow_html=True)

    st.markdown(f'<p class="step-label">Question {step+1} of {total}</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="wizard-question">{s["question"]}</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="wizard-hint">{s["hint"]}</p>', unsafe_allow_html=True)

    # Input
    field   = s["field"]
    val     = init.get(field)
    new_val = val

    if s["type"] == "text":
        new_val = st.text_input(s["label"], value=val or "", placeholder=s.get("placeholder",""))

    elif s["type"] == "textarea":
        new_val = st.text_area(
            s["label"], value=val or "",
            placeholder=s.get("placeholder",""), height=140,
        )

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

    # Validation
    if s["required"]:
        if s["type"] == "multiselect":
            is_valid = bool(new_val)
        else:
            is_valid = bool(str(new_val or "").strip())
    else:
        is_valid = True

    # Navigation
    c1, c2, c3 = st.columns([1, 4, 1])
    with c1:
        if st.button("← Back", disabled=(step == 0)):
            st.session_state.wiz_step -= 1
            st.rerun()
    with c3:
        is_last = step == total - 1
        label   = "Save Initiative ✓" if is_last else "Next →"
        if st.button(label, disabled=not is_valid, type="primary"):
            if is_last:
                draft = st.session_state.draft
                if mode == "edit":
                    draft["initiatives"] = [
                        i if i["id"] != init["id"] else init
                        for i in draft["initiatives"]
                    ]
                else:
                    draft["initiatives"].append(init)
                save_submission(user, cur_month(), draft)
                st.session_state.draft  = draft
                st.session_state.screen = "dashboard"
            else:
                st.session_state.wiz_init = init
                st.session_state.wiz_step += 1
            st.rerun()

    # Step dots
    dots = "".join(
        f'<span style="display:inline-block;width:{"24px" if i==step else "8px"};height:8px;'
        f'border-radius:4px;margin:0 3px;'
        f'background:{"#c86a2a" if i<step else "#1a3c5e" if i==step else "#cbd5e1"}"></span>'
        for i in range(total)
    )
    st.markdown(f'<div style="text-align:center;margin-top:20px;">{dots}</div>', unsafe_allow_html=True)


# ── Admin ─────────────────────────────────────────────────────────────────────

def screen_admin():
    cm = cur_month()
    st.markdown(f"## ⚙ Admin Dashboard — {fmt_month(cm)}")

    c1, c2 = st.columns([5, 1])
    with c2:
        if st.button("← Back"):
            st.session_state.screen = (
                "login" if st.session_state.user == "Admin" else "dashboard"
            )
            st.rerun()

    reg      = load_registry()
    all_data = {}
    for u in reg:
        all_data[u] = {}
        for m in [cm, prev_month()]:
            sub = load_submission(u, m)
            if sub:
                all_data[u][m] = sub

    rows            = [(u, all_data.get(u, {}).get(cm)) for u in reg]
    submitted_count = sum(1 for _, s in rows if s and s.get("status") in ("submitted","approved"))
    approved_count  = sum(1 for _, s in rows if s and s.get("status") == "approved")
    total_inits     = sum(len((s or {}).get("initiatives",[]) or []) for _, s in rows)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Team Members",     len(reg))
    c2.metric("Submitted",        f"{submitted_count}/{len(reg)}")
    c3.metric("Approved",         approved_count)
    c4.metric("Total Initiatives",total_inits)

    # Exports — side by side
    ex1, ex2 = st.columns(2)
    with ex1:
        all_bytes = build_excel(
            all_data,
            label=f"Consolidated — All Users  |  Period: {fmt_month(cm)}",
        )
        st.download_button(
            "↓ Consolidated — All Users",
            data=all_bytes,
            file_name=f"RD_Consolidated_All_{cm}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with ex2:
        sub_bytes = build_excel(
            all_data,
            submitted_only=True,
            label=f"Consolidated — Submitted & Approved Only  |  Period: {fmt_month(cm)}",
        )
        st.download_button(
            "↓ Consolidated — Submitted Only",
            data=sub_bytes,
            file_name=f"RD_Consolidated_Submitted_{cm}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.divider()
    st.subheader(f"Team Submissions — {fmt_month(cm)}")

    if not reg:
        st.info("No submissions recorded yet.")
        return

    for u, sub in rows:
        status = (sub or {}).get("status", "not-started")
        inits  = (sub or {}).get("initiatives") or []
        icon   = {"approved":"✅","submitted":"🔵","in-progress":"🟡",
                  "rejected":"🔴","not-started":"⚪"}.get(status,"⚪")

        with st.expander(
            f"{icon}  {u}   —  {STATUS_LABELS.get(status,'—')}   "
            f"({len(inits)} initiative{'s' if len(inits)!=1 else ''})",
            expanded=False,
        ):
            if not sub:
                st.caption("No submission for this month.")
            else:
                c1, c2, c3 = st.columns([3, 1, 1])
                with c1:
                    if sub.get("submitted_at"):
                        ts = datetime.fromtimestamp(sub["submitted_at"]/1000).strftime("%b %d %H:%M")
                        st.caption(f"Submitted: {ts}")
                with c2:
                    if status == "submitted":
                        if st.button("✓ Approve", key=f"appr_{u}", type="primary"):
                            sub["status"]      = "approved"
                            sub["approved_at"] = int(datetime.now().timestamp()*1000)
                            save_submission(u, cm, sub)
                            st.success(f"{u}'s report approved!")
                            st.rerun()
                with c3:
                    if status == "submitted":
                        if st.button("↩ Return", key=f"rej_{u}"):
                            sub["status"]      = "rejected"
                            sub["rejected_at"] = int(datetime.now().timestamp()*1000)
                            save_submission(u, cm, sub)
                            st.warning(f"{u}'s report returned.")
                            st.rerun()

                u_bytes = build_excel(
                    {u: {cm: sub}},
                    only_user=u,
                    label=f"Submitted by: {u}  |  Period: {fmt_month(cm)}",
                )
                st.download_button(
                    f"↓ Export {u}'s Report",
                    data=u_bytes,
                    file_name=f"RD_{u.replace(' ','_')}_{cm}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{u}",
                )

                if inits:
                    st.markdown("**Initiatives:**")
                    for i in inits:
                        co = "↩ " if i.get("carry_over") else ""
                        st.markdown(
                            f"- **{co}{i.get('initiative_name','Unnamed')}** — "
                            f"{i.get('business_component','')} · "
                            f"👥 {', '.join(i.get('team_members') or ['—'])}"
                        )


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
