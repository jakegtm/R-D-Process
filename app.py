"""
R&D Monthly Activity Tracker
Run with: streamlit run app.py
"""

import streamlit as st
import json
import os
from datetime import datetime, date
from pathlib import Path
import calendar

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side
)
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

BIZ_COMPONENTS = [
    "Software Development", "Hardware Engineering", "Process Innovation",
    "Product Development", "Data & Analytics", "Infrastructure", "Other",
]

STATUS_LABELS = {
    "in-progress":  "🟡 In Progress",
    "submitted":    "🔵 Submitted for Review",
    "approved":     "🟢 Approved",
    "rejected":     "🔴 Returned — Needs Revision",
    "not-started":  "⚪ Not Started",
}

WIZARD_STEPS = [
    {
        "field": "business_component",
        "label": "Business Component",
        "type": "select",
        "opts": BIZ_COMPONENTS,
        "question": "What business component does this initiative belong to?",
        "hint": "Select the primary business area that this R&D work supports.",
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
        "label": "Technical Uncertainty",
        "type": "textarea",
        "placeholder": "It is currently unknown whether... We are testing if...",
        "question": "What technical uncertainty are you working to resolve?",
        "hint": "Core R&D eligibility question: what scientific or technical question are you trying to answer? What don't you know yet?",
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
        "label": "R&D Activities This Month",
        "type": "textarea",
        "placeholder": "Prototyping new approach, running performance tests, analyzing results...",
        "question": "What R&D activities are being conducted this month?",
        "hint": "Describe specific tasks, experiments, or development work happening this period.",
        "required": True,
    },
    {
        "field": "team_members",
        "label": "Team Members",
        "type": "multiselect",
        "question": "Which team members are contributing this month?",
        "hint": "Select everyone logging R&D hours to this initiative this period.",
        "required": True,
    },
    {
        "field": "expected_hours",
        "label": "Expected Hours",
        "type": "number",
        "question": "How many total R&D hours are expected this month?",
        "hint": "Combined estimate across all team members for this initiative.",
        "required": True,
    },
    {
        "field": "notes",
        "label": "Notes",
        "type": "textarea",
        "placeholder": "Optional — blockers, scope changes, anything useful for review...",
        "question": "Any additional notes or comments?",
        "hint": "Include anything helpful for the Oversight Lead's review.",
        "required": False,
    },

    # ── DONI'S QUESTIONS ──────────────────────────────────────────────────────
    # These are sample questions — edit the text, type, and opts as needed.
    # To remove a question: delete its block.
    # To add a question: copy a block and give it a unique "field" name.
    # Supported types: text, textarea, select, multiselect, number, date, yesno
    # ─────────────────────────────────────────────────────────────────────────

    {
        "field": "rd_activity_type",
        "label": "Type of R&D Activity",
        "type": "select",
        "opts": [
            "Basic Research — expanding general knowledge, no specific application yet",
            "Applied Research — directed toward a specific practical goal",
            "Experimental Development — using research to create new products or processes",
        ],
        "question": "What type of R&D activity best describes this initiative?",
        "hint": "This helps classify the work for reporting purposes. Choose the closest match.",
        "required": True,
        "doni": True,   # mark as Doni's question so it's easy to find
    },
    {
        "field": "rd_time_percentage",
        "label": "% of Time on R&D",
        "type": "number",
        "question": "What percentage of the team's time this month is R&D work on this initiative?",
        "hint": "Estimate the portion of hours that qualify as genuine R&D vs. routine work. Enter a number between 0 and 100.",
        "required": True,
        "doni": True,
    },
    {
        "field": "prior_month_progress",
        "label": "Progress vs. Last Month",
        "type": "select",
        "opts": [
            "Significant progress — major new findings or breakthroughs",
            "Moderate progress — advancing but no major breakthroughs yet",
            "Minimal progress — work ongoing, results inconclusive",
            "Blocked — work paused due to obstacles (explain in notes)",
            "N/A — this is a new initiative",
        ],
        "question": "How did work this month advance your understanding of the technical uncertainty compared to last month?",
        "hint": "Give an honest assessment of momentum. Use the Notes field to add detail.",
        "required": True,
        "doni": True,
    },
    {
        "field": "obstacles",
        "label": "Obstacles & Challenges",
        "type": "textarea",
        "placeholder": "e.g. Waiting on equipment, dependency on another team, unexpected test failures...",
        "question": "What obstacles or challenges did you encounter this month?",
        "hint": "Note anything that slowed progress or may affect future months. Leave blank if none.",
        "required": False,
        "doni": True,
    },
    {
        "field": "next_month_plan",
        "label": "Next Month Plan",
        "type": "textarea",
        "placeholder": "e.g. Complete prototype v2, run user testing, analyze data set...",
        "question": "What do you plan to work on for this initiative next month?",
        "hint": "A brief outline helps the Oversight Lead understand the trajectory of the work.",
        "required": True,
        "doni": True,
    },
    {
        "field": "uncertainty_status",
        "label": "Uncertainty Resolution Status",
        "type": "select",
        "opts": [
            "Still unresolved — uncertainty remains, work continues",
            "Partially resolved — some answers found, more work needed",
            "Fully resolved — technical uncertainty has been eliminated",
            "Abandoned — initiative discontinued (explain in notes)",
        ],
        "question": "What is the current status of the technical uncertainty for this initiative?",
        "hint": "Has your R&D work answered the core technical question yet?",
        "required": True,
        "doni": True,
    },
]

# Fields that carry over vs. fields that reset each month
CARRYOVER_FIELDS = {
    "business_component", "initiative_name", "initiative_description",
    "tech_uncertainty", "start_date", "expected_end_date", "team_members",
    "rd_activity_type",   # type of R&D doesn't usually change month to month
}
RESET_FIELDS = {
    "activities", "expected_hours", "notes",
    "rd_time_percentage", "prior_month_progress", "obstacles",
    "next_month_plan", "uncertainty_status",
}


# ── Data Helpers ──────────────────────────────────────────────────────────────

def cur_month() -> str:
    """Returns 'YYYY-MM'"""
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
    if p.exists():
        return json.loads(p.read_text())
    return []


def save_registry(users: list):
    registry_path().write_text(json.dumps(users, indent=2))


def load_submission(username: str, month: str) -> dict | None:
    p = data_path(username, month)
    if p.exists():
        return json.loads(p.read_text())
    return None


def save_submission(username: str, month: str, data: dict):
    data_path(username, month).write_text(json.dumps(data, indent=2))
    reg = load_registry()
    if username not in reg:
        reg.append(username)
        save_registry(reg)


def new_initiative() -> dict:
    return {
        "id": f"{int(datetime.now().timestamp()*1000)}",
        "business_component": "",
        "initiative_name": "",
        "initiative_description": "",
        "tech_uncertainty": "",
        "start_date": None,
        "expected_end_date": None,
        "activities": "",
        "team_members": [],
        "expected_hours": 0,
        "notes": "",
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

def build_excel(all_data: dict, only_user: str | None = None) -> bytes:
    """
    Builds and returns an Excel workbook as bytes.
    all_data = { username: { month: submission_dict } }
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "R&D Tracking"

    headers = [
        "User", "Month", "Business Component", "Initiative Name",
        "Description", "Technical Uncertainty", "Start Date",
        "Expected End Date", "Activities This Month", "Team Members",
        "Expected Hours", "Notes", "Status", "Submitted At",
    ]

    # Header row styling
    hdr_fill   = PatternFill("solid", fgColor="1a3c5e")
    hdr_font   = Font(color="FFFFFF", bold=True, size=11)
    hdr_align  = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

    ws.row_dimensions[1].height = 32
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.fill   = hdr_fill
        cell.font   = hdr_font
        cell.alignment = hdr_align
        cell.border = thin_border

    # Column widths
    col_widths = [16, 14, 22, 26, 42, 42, 12, 14, 42, 32, 14, 32, 22, 20]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Alternating row fill
    fill_even = PatternFill("solid", fgColor="EFF6FF")
    cell_font  = Font(size=10)
    cell_align = Alignment(vertical="top", wrap_text=True)

    row_num = 2
    for username, months in all_data.items():
        if only_user and username != only_user:
            continue
        for month, sub in months.items():
            if not sub or not sub.get("initiatives"):
                continue
            for init in sub["initiatives"]:
                fill = fill_even if row_num % 2 == 0 else None
                vals = [
                    username,
                    fmt_month(month),
                    init.get("business_component", ""),
                    init.get("initiative_name", ""),
                    init.get("initiative_description", ""),
                    init.get("tech_uncertainty", ""),
                    str(init.get("start_date") or ""),
                    str(init.get("expected_end_date") or ""),
                    init.get("activities", ""),
                    ", ".join(init.get("team_members") or []),
                    init.get("expected_hours", 0),
                    init.get("notes", ""),
                    STATUS_LABELS.get(sub.get("status", ""), sub.get("status", "")),
                    (
                        datetime.fromtimestamp(sub["submitted_at"] / 1000).strftime("%Y-%m-%d %H:%M")
                        if sub.get("submitted_at") else ""
                    ),
                ]
                ws.row_dimensions[row_num].height = 60
                for col_idx, val in enumerate(vals, 1):
                    cell = ws.cell(row=row_num, column=col_idx, value=val)
                    cell.font = cell_font
                    cell.alignment = cell_align
                    cell.border = thin_border
                    if fill:
                        cell.fill = fill
                row_num += 1

    # Freeze top row
    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── CSS ───────────────────────────────────────────────────────────────────────

def inject_css():
    st.markdown("""
    <style>
    /* Hide Streamlit branding */
    #MainMenu, footer, header { visibility: hidden; }

    /* App background */
    .stApp { background: #f1f5f9; }

    /* Card style */
    .rd-card {
        background: white;
        border-radius: 14px;
        padding: 24px;
        box-shadow: 0 1px 6px rgba(0,0,0,0.07);
        margin-bottom: 16px;
    }

    /* Badges */
    .badge-approved    { background:#f0fdf4; color:#166534; border:1.5px solid #86efac; }
    .badge-submitted   { background:#eff6ff; color:#1e40af; border:1.5px solid #93c5fd; }
    .badge-in-progress { background:#fef9ec; color:#92600a; border:1.5px solid #fcd34d; }
    .badge-rejected    { background:#fff1f2; color:#9f1239; border:1.5px solid #fda4af; }
    .badge-not-started { background:#f8fafc; color:#64748b; border:1.5px solid #cbd5e1; }
    .rd-badge {
        display:inline-block; padding:3px 12px; border-radius:20px;
        font-size:12px; font-weight:700;
    }

    /* Step wizard progress */
    .progress-bar-outer {
        background:#e2e8f0; border-radius:4px; height:8px;
        overflow:hidden; margin-bottom:6px;
    }
    .progress-bar-inner {
        background:#c86a2a; height:100%;
        border-radius:4px; transition:width 0.3s;
    }

    /* Wizard question */
    .wizard-question { font-size:22px; font-weight:700; color:#1a3c5e; margin-bottom:6px; }
    .wizard-hint     { font-size:14px; color:#64748b; margin-bottom:20px; }
    .step-label      { font-size:11px; font-weight:700; color:#c86a2a;
                       text-transform:uppercase; letter-spacing:.8px; }

    /* Carryover banner */
    .carryover-banner {
        background:#fffbeb; border:1.5px solid #fcd34d;
        border-radius:10px; padding:10px 16px; margin-bottom:14px;
        font-size:13px; color:#92600a;
    }
    .prefilled-banner {
        background:#f0f9ff; border:1.5px solid #93c5fd;
        border-radius:10px; padding:10px 16px; margin-bottom:14px;
        font-size:13px; color:#1e40af;
    }

    /* Page header */
    .rd-header {
        background:#1a3c5e; color:white;
        padding:14px 32px; border-radius:12px;
        margin-bottom:24px; display:flex;
        align-items:center; justify-content:space-between;
    }
    .rd-header h1 { margin:0; font-size:22px; color:white; }

    /* Stat box */
    .stat-box {
        background:white; border-radius:12px; text-align:center;
        padding:18px 12px; box-shadow:0 1px 4px rgba(0,0,0,0.07);
    }
    .stat-val  { font-size:32px; font-weight:800; color:#1a3c5e; margin:0; }
    .stat-lbl  { font-size:12px; color:#64748b; margin:0; }
    </style>
    """, unsafe_allow_html=True)


# ── Badge helper ──────────────────────────────────────────────────────────────

def badge_html(status: str) -> str:
    cls_map = {
        "approved":    "badge-approved",
        "submitted":   "badge-submitted",
        "in-progress": "badge-in-progress",
        "rejected":    "badge-rejected",
    }
    cls = cls_map.get(status, "badge-not-started")
    label = STATUS_LABELS.get(status, status)
    return f'<span class="rd-badge {cls}">{label}</span>'


# ── Session state init ────────────────────────────────────────────────────────

def init_session():
    defaults = {
        "screen":     "login",       # login | dashboard | wizard | admin
        "user":       None,
        "is_admin":   False,
        "draft":      {"initiatives": [], "status": "in-progress"},
        "wiz_init":   None,          # initiative dict being edited
        "wiz_step":   0,
        "wiz_mode":   "new",         # new | edit | carryover
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


# ── Screens ───────────────────────────────────────────────────────────────────

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

        with st.container():
            options = ["— Select your name —", "⚙ Admin (Oversight Lead)"] + EMPLOYEES
            sel = st.selectbox("Sign in as", options, label_visibility="collapsed")

            if st.button("Sign In →", use_container_width=True, type="primary"):
                if sel.startswith("— "):
                    st.warning("Please select your name.")
                else:
                    name = sel.replace("⚙ ", "").replace(" (Oversight Lead)", "").strip()
                    is_admin = "Admin" in sel

                    st.session_state.user     = name
                    st.session_state.is_admin = is_admin

                    if is_admin:
                        st.session_state.screen = "admin"
                    else:
                        existing = load_submission(name, cur_month())
                        st.session_state.draft = existing or {
                            "initiatives": [], "status": "in-progress"
                        }
                        st.session_state.screen = "dashboard"
                    st.rerun()

        st.markdown("""
        <p style="text-align:center; font-size:12px; color:#94a3b8; margin-top:16px;">
            Your entries are saved automatically as you go.
        </p>
        """, unsafe_allow_html=True)


def screen_dashboard():
    user  = st.session_state.user
    draft = st.session_state.draft
    submitted = draft["status"] != "in-progress"
    cm = cur_month()
    pm = prev_month()

    # Load previous month for carry-over
    prev_sub  = load_submission(user, pm)
    prev_inits = []
    if prev_sub:
        today = date.today()
        prev_inits = [
            i for i in (prev_sub.get("initiatives") or [])
            if not i.get("expected_end_date")
            or datetime.strptime(str(i["expected_end_date"]), "%Y-%m-%d").date() >= today
        ]

    # ── Header ──────────────────────────────────────────────────────────────
    c1, c2 = st.columns([5, 1])
    with c1:
        st.markdown(f"## 🔬 R&D Tracker — {fmt_month(cm)}")
        st.caption(f"Signed in as **{user}**")
    with c2:
        st.write("")
        if st.button("Sign Out"):
            for k in ["screen","user","is_admin","draft","wiz_init","wiz_step","wiz_mode"]:
                del st.session_state[k]
            st.rerun()

    st.divider()

    # ── Status strip ────────────────────────────────────────────────────────
    total_hours = sum(i.get("expected_hours", 0) or 0 for i in draft["initiatives"])
    c1, c2, c3, c4 = st.columns([3, 1, 1, 1])
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
        st.metric("Total Hours", total_hours)
    with c4:
        if st.session_state.is_admin:
            if st.button("Admin View →"):
                st.session_state.screen = "admin"
                st.rerun()

    # ── Carry-over prompt ────────────────────────────────────────────────────
    if not submitted and prev_inits and not draft["initiatives"]:
        st.markdown(f"""
        <div class="carryover-banner">
            <strong>📋 Ongoing initiatives from {fmt_month(pm)}</strong><br>
            You had <strong>{len(prev_inits)}</strong> active initiative{"s" if len(prev_inits)!=1 else ""} last month.
            Carry any forward — key details will be pre-filled; you only update activities and hours.
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

    # ── Initiatives list ─────────────────────────────────────────────────────
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
            with st.expander(
                f"{'↩ ' if init.get('carry_over') else ''}"
                f"{init.get('initiative_name','Unnamed')} — {init.get('business_component','')}",
                expanded=True,
            ):
                c1, c2, c3 = st.columns([3, 1, 1])
                with c1:
                    st.markdown(f"**{init.get('initiative_description','—')}**")
                    st.caption(
                        f"📅 {init.get('start_date','—')} → {init.get('expected_end_date','—')}  "
                        f"  👥 {', '.join(init.get('team_members') or ['—'])}  "
                        f"  ⏱ {init.get('expected_hours',0)} hrs"
                    )
                    if init.get("tech_uncertainty"):
                        st.markdown("**Technical Uncertainty:**")
                        st.markdown(
                            f'<div style="background:#f8fafc; padding:10px 14px; border-radius:6px; '
                            f'font-size:13px; color:#334155;">{init["tech_uncertainty"]}</div>',
                            unsafe_allow_html=True,
                        )
                    if init.get("activities"):
                        st.markdown("**Activities this month:**")
                        st.caption(init["activities"])
                    if init.get("notes"):
                        st.markdown(f"*Notes: {init['notes']}*")
                with c2:
                    if not submitted and st.button("✏ Edit", key=f"edit_{init['id']}"):
                        st.session_state.wiz_init = dict(init)
                        st.session_state.wiz_step = 0
                        st.session_state.wiz_mode = "edit"
                        st.session_state.screen   = "wizard"
                        st.rerun()
                with c3:
                    if not submitted and st.button("🗑 Remove", key=f"del_{init['id']}"):
                        draft["initiatives"] = [
                            i for i in draft["initiatives"] if i["id"] != init["id"]
                        ]
                        save_submission(user, cm, draft)
                        st.session_state.draft = draft
                        st.success("Initiative removed.")
                        st.rerun()

    # ── Submit + Export ──────────────────────────────────────────────────────
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
        # Individual export
        all_data = {user: {cm: draft}}
        xlsx_bytes = build_excel(all_data, only_user=user)
        st.download_button(
            "↓ Download My Report",
            data=xlsx_bytes,
            file_name=f"RD_{user.replace(' ','_')}_{cm}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


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

    # ── Header ──────────────────────────────────────────────────────────────
    c1, c2 = st.columns([5, 1])
    with c1:
        title = {"new":"New Initiative","edit":"Edit Initiative","carryover":"Carry Over Initiative"}[mode]
        st.markdown(f"### {title}")
    with c2:
        if st.button("✕ Cancel"):
            st.session_state.screen = "dashboard"
            st.rerun()

    # Progress bar
    st.markdown(f"""
    <div class="progress-bar-outer">
        <div class="progress-bar-inner" style="width:{pct}%"></div>
    </div>
    <p style="font-size:12px; color:#64748b; margin-top:2px;">
        Step {step+1} of {total} &nbsp;·&nbsp; {pct}% complete
    </p>
    """, unsafe_allow_html=True)

    # Context banners
    if is_carryover and not is_prefilled:
        st.markdown(f"""
        <div class="carryover-banner">
            <strong>Updating for {fmt_month(cur_month())}:</strong> {init.get("initiative_name","this initiative")}
        </div>
        """, unsafe_allow_html=True)
    elif is_prefilled:
        st.markdown("""
        <div class="prefilled-banner">
            Pre-filled from last month — confirm or update before continuing.
        </div>
        """, unsafe_allow_html=True)

    st.markdown(f'<p class="step-label">Question {step+1}</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="wizard-question">{s["question"]}</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="wizard-hint">{s["hint"]}</p>', unsafe_allow_html=True)

    # ── Input ────────────────────────────────────────────────────────────────
    field = s["field"]
    val   = init.get(field)
    new_val = val  # will be overwritten

    if s["type"] == "select":
        opts = s["opts"]
        idx  = opts.index(val) if val in opts else 0
        new_val = st.radio(s["label"], opts, index=idx, horizontal=True)

    elif s["type"] == "text":
        new_val = st.text_input(s["label"], value=val or "", placeholder=s.get("placeholder",""))

    elif s["type"] == "textarea":
        new_val = st.text_area(s["label"], value=val or "", placeholder=s.get("placeholder",""), height=130)

    elif s["type"] == "date":
        parsed = None
        if val:
            try:
                parsed = datetime.strptime(str(val), "%Y-%m-%d").date()
            except Exception:
                pass
        picked = st.date_input(s["label"], value=parsed)
        new_val = picked.strftime("%Y-%m-%d") if picked else None

    elif s["type"] == "number":
        new_val = st.number_input(s["label"], min_value=0, value=int(val or 0), step=4)

    elif s["type"] == "multiselect":
        new_val = st.multiselect(s["label"], EMPLOYEES, default=val or [])

    # Update in-place
    init[field] = new_val

    # ── Validation ───────────────────────────────────────────────────────────
    is_valid = True
    if s["required"]:
        if s["type"] == "multiselect":
            is_valid = bool(new_val)
        elif s["type"] == "number":
            is_valid = new_val is not None
        else:
            is_valid = bool(str(new_val or "").strip())

    # ── Navigation ───────────────────────────────────────────────────────────
    c1, c2, c3 = st.columns([1, 4, 1])
    with c1:
        if st.button("← Back", disabled=(step == 0)):
            st.session_state.wiz_step -= 1
            st.rerun()
    with c3:
        if step < total - 1:
            if st.button("Next →", disabled=not is_valid, type="primary"):
                st.session_state.wiz_init = init
                st.session_state.wiz_step += 1
                st.rerun()
        else:
            if st.button("Save Initiative ✓", disabled=not is_valid, type="primary"):
                # Merge back into draft
                draft = st.session_state.draft
                inits = draft["initiatives"]
                if mode == "edit":
                    draft["initiatives"] = [i if i["id"] != init["id"] else init for i in inits]
                else:
                    draft["initiatives"].append(init)
                save_submission(user, cur_month(), draft)
                st.session_state.draft  = draft
                st.session_state.screen = "dashboard"
                st.rerun()

    # Step dots
    dots = "".join(
        f'<span style="display:inline-block; width:{"24px" if i==step else "8px"}; height:8px; '
        f'border-radius:4px; margin:0 3px; background:{"#c86a2a" if i<step else "#1a3c5e" if i==step else "#cbd5e1"}"></span>'
        for i in range(total)
    )
    st.markdown(f'<div style="text-align:center; margin-top:20px;">{dots}</div>', unsafe_allow_html=True)


def screen_admin():
    cm = cur_month()
    st.markdown(f"## ⚙ Admin Dashboard — {fmt_month(cm)}")

    c1, c2 = st.columns([5, 1])
    with c2:
        if st.button("← Back"):
            st.session_state.screen = (
                "dashboard" if not st.session_state.is_admin
                or st.session_state.user != "Admin"
                else "login"
            )
            st.rerun()

    # Load all data
    reg = load_registry()
    all_data = {}
    for u in reg:
        all_data[u] = {}
        for m in [cm, prev_month()]:
            sub = load_submission(u, m)
            if sub:
                all_data[u][m] = sub

    # Stats
    rows = [(u, all_data.get(u, {}).get(cm)) for u in reg]
    submitted_count = sum(1 for _, s in rows if s and s.get("status") in ("submitted","approved"))
    approved_count  = sum(1 for _, s in rows if s and s.get("status") == "approved")
    total_inits     = sum(len((s or {}).get("initiatives",[]) or []) for _, s in rows)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Team Members",    len(reg))
    c2.metric("Submitted",       f"{submitted_count}/{len(reg)}")
    c3.metric("Approved",        approved_count)
    c4.metric("Total Initiatives", total_inits)

    # Consolidated export
    xlsx_bytes = build_excel(all_data)
    st.download_button(
        "↓ Download Consolidated Report (All Users)",
        data=xlsx_bytes,
        file_name=f"RD_Consolidated_{cm}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()
    st.subheader(f"Team Submissions — {fmt_month(cm)}")

    if not reg:
        st.info("No submissions recorded yet.")
        return

    for u, sub in rows:
        status  = (sub or {}).get("status", "not-started")
        inits   = (sub or {}).get("initiatives") or []
        icon    = {"approved":"✅","submitted":"🔵","in-progress":"🟡",
                   "rejected":"🔴","not-started":"⚪"}.get(status,"⚪")

        with st.expander(f"{icon}  {u}   —  {STATUS_LABELS.get(status,'—')}   ({len(inits)} initiative{'s' if len(inits)!=1 else ''})", expanded=False):
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

                # Individual export
                u_bytes = build_excel({u: {cm: sub}}, only_user=u)
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
                            f"👥 {', '.join(i.get('team_members') or ['—'])} · "
                            f"⏱ {i.get('expected_hours',0)} hrs"
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
