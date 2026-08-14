"""
Storage layer for the R&D Activity Tracker.

Everything the app knows about persistence goes through this module. The rest of
app.py never touches the filesystem, so swapping this implementation for a
SharePoint list (via Microsoft Graph) means rewriting only this file.

Two kinds of data, deliberately kept in separate directories:

    data/reports/   one JSON file per report
    data/config/    app configuration (entity list, permissions)

Keeping them apart is what stops "delete all reports" from taking out the
config with it, and stops a config file from being mistaken for a report.

A report is identified by the triple (user, entity, period):

    period  = filing month, "YYYY-MM"
    entity  = entity number, e.g. "107"
    user    = display name of the preparer

All three are part of the key. Earlier versions keyed only on (user, period),
which silently merged two entities filed in the same month into one record.
"""

from __future__ import annotations

import json
import re
import time
from pathlib import Path
from typing import Any

DATA_DIR    = Path("data")
REPORTS_DIR = DATA_DIR / "reports"
CONFIG_DIR  = DATA_DIR / "config"

for _d in (DATA_DIR, REPORTS_DIR, CONFIG_DIR):
    _d.mkdir(parents=True, exist_ok=True)


# ── Key handling ──────────────────────────────────────────────────────────────

_UNSAFE = re.compile(r"[^A-Za-z0-9]+")

def _slug(s: str) -> str:
    """Filesystem-safe fragment. Only used for the filename — the authoritative
    user/entity/period values are stored inside each file and always read from
    there, so slug collisions can never mislabel a report."""
    return _UNSAFE.sub("_", str(s or "").strip()).strip("_") or "none"


def report_key(user: str, entity: str, period: str) -> str:
    return f"{_slug(user)}__{_slug(entity)}__{_slug(period)}"


def _report_path(user: str, entity: str, period: str) -> Path:
    return REPORTS_DIR / f"{report_key(user, entity, period)}.json"


# ── Cache ─────────────────────────────────────────────────────────────────────
# Streamlit reruns the whole script on every interaction, so a naive
# implementation re-reads every file many times per click. Cache the parsed
# reports and invalidate whenever the directory's mtime changes.

_cache: dict[str, Any] = {"stamp": None, "reports": None}

def _dir_stamp() -> tuple:
    try:
        entries = sorted(
            (p.name, p.stat().st_mtime, p.stat().st_size)
            for p in REPORTS_DIR.glob("*.json")
        )
    except OSError:
        return ()
    return tuple(entries)

def invalidate_cache() -> None:
    _cache["stamp"] = None
    _cache["reports"] = None


# ── Reports ───────────────────────────────────────────────────────────────────

def _read_all() -> list[dict]:
    stamp = _dir_stamp()
    if _cache["stamp"] == stamp and _cache["reports"] is not None:
        return _cache["reports"]

    reports: list[dict] = []
    for path in REPORTS_DIR.glob("*.json"):
        try:
            data = json.loads(path.read_text())
        except (json.JSONDecodeError, OSError):
            continue
        if not isinstance(data, dict):
            continue
        # A report must carry its own identity. Anything that doesn't is skipped
        # rather than guessed at.
        if not (data.get("user") and data.get("entity") and data.get("reporting_month")):
            continue
        reports.append(data)

    _cache["stamp"] = stamp
    _cache["reports"] = reports
    return reports


def save_report(report: dict) -> None:
    """Write a report. Requires user, entity, and reporting_month to be set."""
    user   = report.get("user")
    entity = report.get("entity")
    period = report.get("reporting_month")
    if not (user and entity and period):
        raise ValueError(
            "save_report needs user, entity, and reporting_month. "
            f"Got user={user!r} entity={entity!r} reporting_month={period!r}"
        )
    report["updated_at"] = int(time.time() * 1000)
    _report_path(user, entity, period).write_text(json.dumps(report, indent=2))
    invalidate_cache()


def get_report(user: str, entity: str, period: str) -> dict | None:
    path = _report_path(user, entity, period)
    if not path.exists():
        return None
    try:
        data = json.loads(path.read_text())
    except (json.JSONDecodeError, OSError):
        return None
    return data if isinstance(data, dict) else None


def list_reports(
    user: str | None = None,
    entity: str | None = None,
    period: str | None = None,
    statuses: list[str] | None = None,
    non_empty: bool = True,
) -> list[dict]:
    """All reports matching the given filters, newest period first.

    non_empty=True (the default) hides reports with no initiatives — these are
    abandoned setup sessions and shouldn't appear in any list or count.
    """
    out = []
    for r in _read_all():
        if user is not None and r.get("user") != user:
            continue
        if entity is not None and r.get("entity") != entity:
            continue
        if period is not None and r.get("reporting_month") != period:
            continue
        if statuses is not None and r.get("status") not in statuses:
            continue
        if non_empty and not r.get("initiatives"):
            continue
        out.append(r)
    out.sort(
        key=lambda r: (r.get("reporting_month", ""), r.get("entity", ""), r.get("user", "")),
        reverse=True,
    )
    return out


def delete_report(user: str, entity: str, period: str) -> bool:
    path = _report_path(user, entity, period)
    if path.exists():
        path.unlink()
        invalidate_cache()
        return True
    return False


def delete_all_reports() -> int:
    """Delete every report. Config is in a separate directory and is untouched."""
    n = 0
    for path in REPORTS_DIR.glob("*.json"):
        path.unlink()
        n += 1
    invalidate_cache()
    return n


# ── Derived views ─────────────────────────────────────────────────────────────

def list_users() -> list[str]:
    return sorted({r["user"] for r in _read_all() if r.get("user")})


def list_periods(user: str | None = None) -> list[str]:
    """Filing months that have at least one non-empty report, newest first."""
    return sorted(
        {
            r["reporting_month"]
            for r in list_reports(user=user)
            if r.get("reporting_month")
        },
        reverse=True,
    )


def list_combos() -> list[tuple[str, str]]:
    """Unique (entity, period) pairs that have data, oldest period first."""
    combos = {
        (r["entity"], r["reporting_month"])
        for r in list_reports()
        if r.get("entity") and r.get("reporting_month")
    }
    return sorted(combos, key=lambda c: (c[1], c[0]))


def count_reports() -> int:
    return len(list_reports())


# ── Config ────────────────────────────────────────────────────────────────────

def get_config(name: str, default: Any = None) -> Any:
    path = CONFIG_DIR / f"{_slug(name)}.json"
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text())
    except (json.JSONDecodeError, OSError):
        return default


def set_config(name: str, value: Any) -> None:
    (CONFIG_DIR / f"{_slug(name)}.json").write_text(json.dumps(value, indent=2))
