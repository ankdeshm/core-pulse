"""CORE Pulse branding constants — single source of truth across UI, sidebar, and docs."""
from __future__ import annotations

from pathlib import Path

from app.config import settings

_CFG = settings()

APP_NAME = _CFG.get("app_name", "CORE Pulse")
APP_FULL_NAME = _CFG.get("app_full_name", "Operational Rollout & Adoption Intelligence")
TAGLINE = _CFG.get("tagline", "Measure adoption clearly. Scale rollout confidently.")
COMPANY = _CFG.get("company_long", "RL Inspection Company")
OPERATIONAL_PLATFORM = _CFG.get("operational_platform", "CORE")
AUTHOR_LINE = _CFG.get("author_line", "Developed and maintained by BI & Data Engineering")
VERSION = _CFG.get("version", "1.0.0")

ASSETS = Path(__file__).resolve().parent / "assets"
LOGO_PATH = ASSETS / "core_pulse_logo.svg"

TEAL = "#0F766E"
INDIGO = "#3730A3"
INK = "#0B1F3A"


def page_title(page: str | None = None) -> str:
    return f"{APP_NAME} · {page}" if page else f"{APP_NAME} · {APP_FULL_NAME}"


def footer_line() -> str:
    return f"{APP_NAME} v{VERSION} · {COMPANY} ({OPERATIONAL_PLATFORM}) · {AUTHOR_LINE}"
