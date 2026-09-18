"""
Access roster — who can sign in, and which entities they may file for and see.

An admin downloads a template, fills it in, and uploads it. That file becomes
the authoritative list of users: anyone not on it has no access at all.

Why this exists at all: clicking through per-person entity permissions is fine
for ten people and impossible for a thousand. The roster is the same data as
the Settings checkboxes, in a form a client can maintain in Excel and hand back.

THIS IS NOT AUTHENTICATION. Sign-in is still a name picker with no password, so
the roster decides what a given name is *allowed* to do, not that the person
sitting at the keyboard is that person. It is a meaningful control once sign-in
is tied to SSO or a login, and until then it is an organising device — useful,
but do not describe it to a client as access control.

Storage goes through store.set_config so the roster lives with the rest of the
app config, but reads are cached on the config file's mtime: at 2,000 rows the
JSON is large enough that re-parsing it on every Streamlit rerun is noticeable.
"""

from __future__ import annotations

import io
import re
import time

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import store

CONFIG_KEY = "access_roster"

# Matches the export template's palette so the file looks like it belongs.
GREEN = "9BBB59"
BEIGE = "EEECE1"
WHITE = "FFFFFF"
DARK  = "1F1F1F"
GRAY  = "666666"

SHEET_PERMS = "Permissions"
SHEET_HELP  = "Instructions"

# "this person may file for every entity", including ones added later
ALL_ENTITIES = "*"

ROLE_PREPARER = "preparer"
ROLE_ADMIN    = "admin"
ROLES = (ROLE_PREPARER, ROLE_ADMIN)

# Column headers we recognise on upload. First match wins.
COL_HINTS = {
    "employee_id": ("employee id", "employee number", "emp id", "worker id", "id"),
    "name":        ("name", "preferred name", "full name", "employee name", "worker name"),
    "email":       ("email", "e-mail", "email address", "upn", "login"),
    "entities":    ("entities", "entity", "entities allowed", "entity access", "access"),
    "role":        ("role", "access level", "permission", "type"),
}

_cache: dict = {"stamp": None, "data": None}


# ── Storage ───────────────────────────────────────────────────────────────────

def _config_path():
    return store.CONFIG_DIR / f"{CONFIG_KEY}.json"


def load() -> dict:
    """{"meta": {...}, "users": [...], "_by_key": {...}} — empty if none loaded."""
    empty = {"meta": {}, "users": [], "_by_key": {}}
    path = _config_path()
    if not path.exists():
        _cache["stamp"], _cache["data"] = None, None
        return empty

    st_ = path.stat()
    stamp = (st_.st_mtime_ns, st_.st_size)
    if _cache["stamp"] == stamp and _cache["data"] is not None:
        return _cache["data"]

    data = store.get_config(CONFIG_KEY, None)
    if not isinstance(data, dict):
        return empty
    data.setdefault("meta", {})
    users = data.setdefault("users", [])
    data["_by_key"] = {u.get("display") or u.get("name", ""): u for u in users}

    _cache["stamp"], _cache["data"] = stamp, data
    return data


def save(users: list, meta: dict) -> None:
    store.set_config(CONFIG_KEY, {
        "meta": {**meta, "saved_at": int(time.time() * 1000), "n_users": len(users)},
        "users": users,
    })
    _cache["stamp"] = None


def clear() -> bool:
    if not is_loaded():
        return False
    store.set_config(CONFIG_KEY, {"meta": {}, "users": []})
    _cache["stamp"] = None
    return True


def is_loaded() -> bool:
    return bool(load()["users"])


def meta() -> dict:
    return load()["meta"]


def users() -> list:
    return list(load()["users"])


# ── Lookups ───────────────────────────────────────────────────────────────────

def sign_in_names() -> list:
    """Display names allowed to sign in, alphabetical."""
    return sorted((u.get("display") or u.get("name", "")) for u in load()["users"])


def get_user(display_name: str) -> dict | None:
    return load()["_by_key"].get(display_name)


def is_known(display_name: str) -> bool:
    return display_name in load()["_by_key"]


def entities_for(display_name: str, all_entities: list) -> list:
    """Entities this person may file for and see.

    Returns [] for someone not on the roster — deny by default. The caller must
    treat an empty list as "no access", never as "no restriction"; that
    inversion is the whole difference between this and the old behaviour.
    """
    u = get_user(display_name)
    if not u:
        return []
    ents = u.get("entities") or []
    if ALL_ENTITIES in ents:
        return list(all_entities)
    return [e for e in all_entities if e in ents]


def role_of(display_name: str) -> str:
    u = get_user(display_name)
    return (u.get("role") or ROLE_PREPARER) if u else ROLE_PREPARER


def is_admin(display_name: str) -> bool:
    return role_of(display_name) == ROLE_ADMIN


def admin_count() -> int:
    return sum(1 for u in load()["users"] if (u.get("role") or "") == ROLE_ADMIN)


def counts() -> tuple:
    """(users, admins)"""
    return len(load()["users"]), admin_count()


# ── Template ──────────────────────────────────────────────────────────────────

def _style_header(ws, headers: list, widths: list):
    thin = Side(style="thin", color="A0A0A0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for i, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(1, i, h)
        c.fill = PatternFill("solid", fgColor=GREEN)
        c.font = Font(name="Arial Narrow", bold=True, size=12, color=WHITE)
        c.border = border
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 26
    ws.freeze_panes = "A2"
    return border


def build_template(all_entities: list, existing: list | None = None) -> bytes:
    """The upload template. With `existing`, this is also the export of the
    current roster — same layout either way, so a round trip is lossless."""
    wb = Workbook()

    # ── Instructions ────────────────────────────────────────────────────────
    ws = wb.active
    ws.title = SHEET_HELP
    ws.column_dimensions["A"].width = 110
    lines = [
        ("How to fill in this file", True),
        ("", False),
        (f"Put one row per person on the '{SHEET_PERMS}' tab. Anyone not listed has NO access —", False),
        ("they cannot sign in, file a report, or see anything in the archive.", False),
        ("", False),
        ("Name  (required)", True),
        ("   Exactly how the person's name should appear when they sign in.", False),
        ("   The name IS the identifier, so every name must be different. If two", False),
        ("   people really are called the same thing, make the names different —", False),
        ("   e.g.  Amy Brooks (Clinical)  and  Amy Brooks (Regulatory).", False),
        ("   The file will not import while any name is duplicated: two identical", False),
        ("   names would share one account and write into each other's reports.", False),
        ("", False),
        ("   Careful when changing a name later. Reports are filed under the name,", False),
        ("   so renaming someone hides the reports they already submitted.", False),
        ("", False),
        ("Entities  (required)", True),
        ("   Which entities this person may FILE for and SEE. Separate several with", False),
        ("   commas, e.g.  107, 108", False),
        ("   Put  *  to mean every entity, including any added later.", False),
        (f"   Entities currently set up: {', '.join(all_entities) or '(none)'}", False),
        ("", False),
        ("Role  (optional)", True),
        ("   Leave blank for a normal preparer.", False),
        ("   Put  admin  to make someone an Oversight Lead, who reviews and accepts", False),
        ("   reports and can see every entity regardless of the Entities column.", False),
        ("", False),
        ("Email  (optional)", True),
        ("   Not used yet. Fill it in now and it is ready if sign-in moves to SSO.", False),
        ("", False),
        ("Uploading replaces the whole roster, so always start from an export of the", False),
        ("current list rather than a blank file if you only mean to change a few people.", False),
    ]
    for r, (text, bold) in enumerate(lines, 1):
        c = ws.cell(r, 1, text)
        c.font = Font(name="Arial Narrow", size=12, bold=bold,
                      color=DARK if bold or text else GRAY)
        c.alignment = Alignment(wrap_text=True, vertical="top")

    # ── Permissions ─────────────────────────────────────────────────────────
    ws = wb.create_sheet(SHEET_PERMS)
    headers = ["Name", "Email", "Entities", "Role"]
    border = _style_header(ws, headers, [32, 36, 26, 14])

    rows = existing or []
    for r, u in enumerate(rows, 2):
        ents = u.get("entities") or []
        ws.cell(r, 1, u.get("name", ""))
        ws.cell(r, 2, u.get("email", ""))
        ws.cell(r, 3, ALL_ENTITIES if ALL_ENTITIES in ents else ", ".join(ents))
        ws.cell(r, 4, u.get("role", "") if u.get("role") != ROLE_PREPARER else "")
        for cc in range(1, 5):
            cell = ws.cell(r, cc)
            cell.font = Font(name="Arial Narrow", size=11, color=DARK)
            cell.border = border
            cell.fill = PatternFill("solid", fgColor=BEIGE)
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    # Blank styled rows so an empty template doesn't look broken.
    if not rows:
        for r in range(2, 12):
            for cc in range(1, 5):
                cell = ws.cell(r, cc, "")
                cell.border = border
                cell.fill = PatternFill("solid", fgColor=BEIGE)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── Parsing an uploaded roster ────────────────────────────────────────────────

def _find_header(ws, max_scan: int = 10):
    """(row_index, {field: col_index}) for the header row, or (0, {})."""
    for r in range(1, max_scan + 1):
        found = {}
        for c in range(1, (ws.max_column or 0) + 1):
            v = ws.cell(r, c).value
            if v is None:
                continue
            label = str(v).strip().lower()
            for field, hints in COL_HINTS.items():
                if field in found:
                    continue
                if label in hints:
                    found[field] = c
        if "name" in found and "entities" in found:
            return r, found
    return 0, {}


def _split_entities(raw, all_entities: list):
    """Parse the Entities cell. Returns (entities, unknown)."""
    text = str(raw or "").strip()
    if not text:
        return [], []
    if text == ALL_ENTITIES or text.lower() in ("all", "any", "*"):
        return [ALL_ENTITIES], []
    parts = [p.strip() for p in re.split(r"[,;/|]+", text) if p.strip()]
    known   = [p for p in parts if p in all_entities]
    unknown = [p for p in parts if p not in all_entities]
    # Unknown codes are KEPT, not dropped: an entity that doesn't exist yet may
    # be added next week, and silently discarding it would leave the admin
    # believing access was granted. The caller warns about them instead.
    return known + unknown, unknown


def parse(file_bytes: bytes, all_entities: list):
    """Read an uploaded roster. Returns (users, warnings, errors).

    errors are fatal — nothing is imported. warnings are worth showing but the
    import can proceed.
    """
    try:
        wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    except Exception as e:
        return [], [], [f"Couldn't open that file as an Excel workbook: {e}"]

    ws = wb[SHEET_PERMS] if SHEET_PERMS in wb.sheetnames else None
    if ws is None:
        # Fall back to the first sheet that has a usable header.
        for cand in wb.worksheets:
            if cand.title == SHEET_HELP:
                continue
            if _find_header(cand)[0]:
                ws = cand
                break
    if ws is None:
        return [], [], [
            f"No '{SHEET_PERMS}' tab found, and no other sheet has both a Name "
            "and an Entities column."
        ]

    hdr_row, cols = _find_header(ws)
    if not hdr_row:
        return [], [], [
            "Couldn't find the header row. It needs a 'Name' column and an "
            "'Entities' column."
        ]

    def cell(r, field):
        ci = cols.get(field)
        if not ci:
            return ""
        v = ws.cell(r, ci).value
        if v is None:
            return ""
        if isinstance(v, float) and v.is_integer():
            v = int(v)
        return str(v).strip()

    users, warnings = [], []
    unknown_entities, no_entities, bad_roles = set(), [], []
    seen_names: dict = {}

    for r in range(hdr_row + 1, (ws.max_row or 0) + 1):
        name = cell(r, "name")
        if not name:
            continue

        ents, unknown = _split_entities(cell(r, "entities"), all_entities)
        unknown_entities.update(unknown)

        role = (cell(r, "role") or "").strip().lower()
        if role in ("", "preparer", "user", "employee"):
            role = ROLE_PREPARER
        elif role in ("admin", "administrator", "oversight lead", "reviewer"):
            role = ROLE_ADMIN
        else:
            bad_roles.append(f"row {r}: {role}")
            role = ROLE_PREPARER

        if not ents and role != ROLE_ADMIN:
            no_entities.append(name)

        users.append({
            "employee_id": cell(r, "employee_id"),
            "name":        name,
            "display":     name,          # disambiguated below
            "email":       cell(r, "email"),
            "entities":    ents,
            "role":        role,
        })
        seen_names[name] = seen_names.get(name, 0) + 1

    if not users:
        return [], [], ["No rows with a name were found on that sheet."]

    # The name is the identifier, so duplicates are fatal rather than a warning.
    # Two rows with the same name are one account: both people would sign in as
    # the same person and file into the same report, and the second row's entity
    # access would silently overwrite the first. There is no safe way to guess
    # which was meant, so the import stops and says which names to fix.
    #
    # An Employee ID column isn't in the template, because the client may not
    # have one to hand. If someone does add the column, it is used to separate
    # duplicates rather than rejecting the file.
    dupes = sorted(n for n, c in seen_names.items() if c > 1)
    for u in users:
        if seen_names[u["name"]] > 1 and u.get("employee_id"):
            u["display"] = f"{u['name']} ({u['employee_id']})"

    still = {}
    for u in users:
        still[u["display"]] = still.get(u["display"], 0) + 1
    collided = sorted(d for d, c in still.items() if c > 1)

    if collided:
        listed = ", ".join(collided[:5]) + ("…" if len(collided) > 5 else "")
        return [], [], [
            f"{len(collided)} duplicate name{'s' if len(collided) != 1 else ''} "
            f"in this file: {listed}. Names have to be unique — each one is a "
            "separate sign-in. Make them different (for example by adding the "
            "team in brackets) and upload again."
        ]
    if dupes:
        warnings.append(
            f"{len(dupes)} name{'s' if len(dupes) != 1 else ''} appeared twice "
            f"({', '.join(dupes[:3])}{'…' if len(dupes) > 3 else ''}) and were "
            "separated using the Employee ID column."
        )
    if unknown_entities:
        warnings.append(
            f"{len(unknown_entities)} entity code"
            f"{'s' if len(unknown_entities) != 1 else ''} "
            f"({', '.join(sorted(unknown_entities)[:5])}) "
            "aren't set up in the app. They're kept, but grant nothing until "
            "the entity is added in Settings."
        )
    if no_entities:
        warnings.append(
            f"{len(no_entities)} "
            f"{'person has' if len(no_entities) == 1 else 'people have'} no "
            f"entities ({', '.join(no_entities[:3])}"
            f"{'…' if len(no_entities) > 3 else ''}) — they can sign in but "
            "won't be able to file or see anything."
        )
    if bad_roles:
        warnings.append(
            f"{len(bad_roles)} unrecognised role value"
            f"{'s' if len(bad_roles) != 1 else ''} "
            f"({', '.join(bad_roles[:3])}) — treated as preparer."
        )

    return users, warnings, []


def diff_against_current(new_users: list) -> dict:
    """What an upload would change, for the confirmation screen."""
    current = {u.get("display"): u for u in load()["users"]}
    incoming = {u.get("display"): u for u in new_users}

    added   = [d for d in incoming if d not in current]
    removed = [d for d in current if d not in incoming]
    changed = []
    for d, u in incoming.items():
        old = current.get(d)
        if not old:
            continue
        if (sorted(old.get("entities") or []) != sorted(u.get("entities") or [])
                or (old.get("role") or "") != (u.get("role") or "")):
            changed.append(d)
    return {"added": sorted(added), "removed": sorted(removed),
            "changed": sorted(changed)}
