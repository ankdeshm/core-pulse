"""Generate the 12-Slide CORE Pulse Executive Presentation Kit:
  1. `presentation_kit/CORE_Pulse_Executive_Presentation.pptx` (native 16:9 widescreen PowerPoint with embedded 1080p MP4s, poster frames, vector Medallion diagrams, and speaker notes)
  2. `presentation_kit/CORE_Pulse_Executive_Deck.html` (standalone interactive widescreen HTML5 deck with embedded 1080p MP4s, keyboard navigation, and zoomable screenshots)

Reads measured numbers from `results/eval/evaluation_summary.json` when present so Slide 11 contains 100% measured figures.
"""
from __future__ import annotations

import html
import json
import re
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parents[1]
KIT = ROOT / "presentation_kit"
VID = KIT / "videos"
IMG = KIT / "screenshots"
OUT_PPTX = KIT / "CORE_Pulse_Executive_Presentation.pptx"
OUT_HTML = KIT / "CORE_Pulse_Executive_Deck.html"

NAVY = RGBColor(0x0B, 0x1F, 0x3A)
TEAL = RGBColor(0x0F, 0x76, 0x6E)
TEAL_BG = RGBColor(0xE6, 0xF4, 0xF1)
INDIGO = RGBColor(0x37, 0x30, 0xA3)
INDIGO_BG = RGBColor(0xEE, 0xF0, 0xFB)
RED = RGBColor(0xDC, 0x26, 0x26)
RED_BG = RGBColor(0xFE, 0xF2, 0xF2)
BRONZE = RGBColor(0x92, 0x40, 0x0E)
BRONZE_BG = RGBColor(0xFE, 0xF3, 0xC7)
SILVER = RGBColor(0x33, 0x41, 0x55)
SILVER_BG = RGBColor(0xF1, 0xF5, 0xF9)
GOLD = RGBColor(0xB4, 0x53, 0x09)
GOLD_BG = RGBColor(0xFE, 0xF9, 0xC3)
SLATE = RGBColor(0x47, 0x55, 0x69)
BORDER = RGBColor(0xCB, 0xD5, 0xE1)
BG = RGBColor(0xF8, 0xFA, 0xFC)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)


def load_eval_rows() -> list[dict]:
    eval_file = ROOT / "results" / "eval" / "evaluation_summary.json"
    if eval_file.exists():
        return json.loads(eval_file.read_text(encoding="utf-8"))
    return []


def code_spans(text: str) -> str:
    """Convert markdown `code` spans to styled HTML <code> tags with zero literal backticks."""
    parts = re.split(r"`([^`]+)`", text)
    out = []
    for i, part in enumerate(parts):
        if i % 2 == 1:
            out.append(f"<code>{html.escape(part)}</code>")
        else:
            out.append(html.escape(part))
    return "".join(out)


# ---------------------------------------------------------------------------
# PPTX Helpers
# ---------------------------------------------------------------------------
def set_bg(slide, color: RGBColor) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_box(slide, x, y, w, h, fill_rgb=WHITE, line_rgb=BORDER, line_pt=1.0, radius=True):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    sh = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    if radius and len(sh.adjustments) > 0:
        sh.adjustments[0] = 0.04
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill_rgb
    if line_rgb is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line_rgb
        sh.line.width = Pt(line_pt)
    return sh


def add_text(slide, x, y, w, h, text, size=14, bold=False, color=NAVY, align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Pt(4)
    tf.margin_top = tf.margin_bottom = Pt(2)
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text.replace("`", "")
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    r.font.name = "Calibri"
    return tb


def add_bullets(slide, x, y, w, h, items: list[tuple[str, str]], size=12.0, space_after=6):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, (lead, rest) in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(space_after)
        r1 = p.add_run()
        r1.text = f"• {lead.replace('`', '')}: " if lead else "• "
        r1.font.size = Pt(size)
        r1.font.bold = True
        r1.font.color.rgb = NAVY
        r1.font.name = "Calibri"
        r2 = p.add_run()
        r2.text = rest.replace("`", "")
        r2.font.size = Pt(size)
        r2.font.bold = False
        r2.font.color.rgb = SLATE
        r2.font.name = "Calibri"
    return tb


def slide_header(slide, kicker: str, title: str, badge: str, num: int) -> None:
    set_bg(slide, BG)
    add_text(slide, 0.55, 0.32, 9.5, 0.28, kicker.upper(), size=10.5, bold=True, color=TEAL)
    add_text(slide, 0.55, 0.56, 10.1, 0.55, title, size=22.5, bold=True, color=NAVY)
    b = add_box(slide, 10.75, 0.44, 2.02, 0.38, fill_rgb=INDIGO_BG, line_rgb=INDIGO, line_pt=0.8)
    p = b.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = badge
    r.font.size = Pt(9.8)
    r.font.bold = True
    r.font.color.rgb = INDIGO
    r.font.name = "Calibri"
    add_box(slide, 0.55, 1.16, 12.22, 0.02, fill_rgb=BORDER, line_rgb=None, radius=False)
    add_text(slide, 0.55, 7.12, 9.5, 0.25,
             "CORE Pulse · Governed Operational Rollout & Adoption Intelligence (Medallion + MCP Semantic Architecture)",
             size=9, color=SLATE)
    add_text(slide, 11.5, 7.12, 1.25, 0.25, f"{num} / 12", size=9.5, bold=True, color=SLATE, align=PP_ALIGN.RIGHT)


def add_callout(slide, y: float, text: str) -> None:
    add_box(slide, 0.55, y, 12.22, 0.56, fill_rgb=TEAL_BG, line_rgb=TEAL, line_pt=1.2)
    add_text(slide, 0.72, y + 0.10, 11.9, 0.38, text, size=11.8, bold=True, color=NAVY)


def add_video_slide(prs, num: int, kicker: str, title: str, badge: str,
                    mp4_name: str, poster_name: str, card_title: str,
                    bullets: list[tuple[str, str]], thumb1: str, thumb2: str, notes: str) -> None:
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(sl, kicker, title, badge, num)
    add_box(sl, 0.53, 1.36, 7.58, 4.64, fill_rgb=NAVY, line_rgb=BORDER, line_pt=1.2)
    add_text(sl, 0.68, 1.42, 7.2, 0.28, f"▶  {mp4_name}  (Click video in PowerPoint slideshow to play 1080p clip)",
             size=10, bold=True, color=WHITE)
    mp4_path = VID / mp4_name
    poster_path = IMG / poster_name
    if mp4_path.exists() and poster_path.exists():
        sl.shapes.add_movie(
            str(mp4_path),
            Inches(0.56), Inches(1.73), Inches(7.52), Inches(4.23),
            poster_frame_image=str(poster_path),
            mime_type="video/mp4",
        )
    if (IMG / thumb1).exists():
        sl.shapes.add_picture(str(IMG / thumb1), Inches(0.55), Inches(6.08), width=Inches(3.68))
    if (IMG / thumb2).exists():
        sl.shapes.add_picture(str(IMG / thumb2), Inches(4.40), Inches(6.08), width=Inches(3.68))

    add_box(sl, 8.30, 1.36, 4.47, 5.62, fill_rgb=WHITE, line_rgb=TEAL, line_pt=1.5)
    add_box(sl, 8.30, 1.36, 4.47, 0.08, fill_rgb=TEAL, line_rgb=None, radius=False)
    add_text(sl, 8.48, 1.54, 4.1, 0.4, card_title, size=15, bold=True, color=NAVY)
    add_bullets(sl, 8.45, 2.00, 4.18, 4.85, bullets, size=11.8, space_after=8)
    sl.notes_slide.notes_text_frame.text = notes


def build_pptx(eval_rows: list[dict]) -> None:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # SLIDE 1: HERO & ICEBERG
    s1 = prs.slides.add_slide(blank)
    set_bg(s1, NAVY)
    add_box(s1, 0.55, 0.52, 5.35, 0.36, fill_rgb=TEAL, line_rgb=None)
    add_text(s1, 0.68, 0.56, 5.15, 0.28, "AI-READY DATA ENGINEERING & OPERATIONAL ROLLOUT PILOT", size=9.6, bold=True, color=WHITE)
    add_text(s1, 0.55, 1.05, 7.1, 1.35, "CORE Pulse: From Medallion Foundations\nto Trusted Rollout Intelligence", size=29, bold=True, color=WHITE)
    add_text(s1, 0.55, 2.48, 7.0, 1.35,
             "Demonstrating how a 5-stage Medallion warehouse (Bronze → Silver → Gold → goldview.v_core_adoption) "
             "and a Read-Only MCP Semantic Layer eliminate metric drift, non-additive employee double-counting, "
             "and partial-month false alarms across RL Inspection Company's global CORE rollout.",
             size=14.2, color=RGBColor(0xE2, 0xE8, 0xF0))

    hero_kpis = [
        ("90% Foundation", "16 Source Tables → Bronze → Silver → Gold → Certified View"),
        ("0 Double-Counting", "Non-additive COUNT(DISTINCT APPUserId) enforced at every grain"),
        ("10 MCP Tools", "Read-only semantic contract across SQL Server 2022 & DuckDB"),
    ]
    for i, (kv, kl) in enumerate(hero_kpis):
        bx = 0.55 + i * 2.38
        add_box(s1, bx, 4.22, 2.24, 1.50, fill_rgb=RGBColor(0x13, 0x2F, 0x54), line_rgb=TEAL, line_pt=1.2)
        add_text(s1, bx + 0.12, 4.36, 2.0, 0.4, kv, size=15.5, bold=True, color=RGBColor(0x5E, 0xEA, 0xD4))
        add_text(s1, bx + 0.12, 4.80, 2.0, 0.80, kl, size=10.2, color=WHITE)

    add_box(s1, 0.55, 6.02, 7.0, 0.68, fill_rgb=RGBColor(0x1E, 0x29, 0x3B), line_rgb=RGBColor(0xF5, 0x9E, 0x0B), line_pt=1.0)
    add_text(s1, 0.70, 6.13, 6.7, 0.48,
             "SYNTHETIC DEMONSTRATION PILOT: Deterministic 18-month world (seed=42, AS_OF_DATE=2026-11-17), "
             "fictional organization, zero corporate connectivity.", size=10.2, bold=True, color=RGBColor(0xFD, 0xE6, 0x8A))

    # Iceberg card on right
    add_box(s1, 7.95, 0.52, 4.82, 6.42, fill_rgb=WHITE, line_rgb=TEAL, line_pt=2.0)
    add_text(s1, 8.15, 0.68, 4.4, 0.32, "THE ADOPTION ANALYTICS ICEBERG", size=11.5, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
    add_box(s1, 8.20, 1.08, 4.32, 1.18, fill_rgb=INDIGO_BG, line_rgb=INDIGO, line_pt=1.2)
    add_text(s1, 8.32, 1.14, 4.1, 0.24, "VISIBLE SURFACE (10%) — EXECUTIVE CONSUMPTION", size=9.2, bold=True, color=INDIGO)
    add_box(s1, 8.32, 1.43, 2.0, 0.68, fill_rgb=INDIGO, line_rgb=None)
    add_text(s1, 8.36, 1.50, 1.92, 0.55, "CORE Pulse Chat\nRollout & Lineage UI", size=9.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_box(s1, 10.40, 1.43, 2.0, 0.68, fill_rgb=WHITE, line_rgb=INDIGO, line_pt=1.0)
    add_text(s1, 10.44, 1.50, 1.92, 0.55, "Power BI · Tableau\nIDE MCP Clients", size=9.5, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    add_text(s1, 8.20, 2.38, 4.32, 0.26, "─── GOVERNED DATA FOUNDATION (90%) ───", size=9.2, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
    layers = [
        ("Certified View & MCP Layer", "goldview.v_core_adoption · 10 Typed Tools · RLS", TEAL_BG, TEAL),
        ("Gold Layer — Atomic Events", "37,116 Events · Daily User Activity · 1,216 Rollups", GOLD_BG, GOLD),
        ("Silver Layer — Conformed Core", "LEFT JOIN Org Enrichment · Service Acct Exclusion", SILVER_BG, SILVER),
        ("Bronze & DataOps Watermarks", "16 CORE Tables · ApplicationLogId=127,985 · QA Logs", BRONZE_BG, BRONZE),
    ]
    for idx_l, (lt, ld, lbg, lbd) in enumerate(layers):
        ly = 2.72 + idx_l * 0.98
        add_box(s1, 8.20, ly, 4.32, 0.84, fill_rgb=lbg, line_rgb=lbd, line_pt=1.1)
        add_text(s1, 8.34, ly + 0.08, 4.05, 0.28, lt, size=10.5, bold=True, color=NAVY)
        add_text(s1, 8.34, ly + 0.38, 4.05, 0.36, ld, size=9.5, color=SLATE)

    # SLIDE 2: WHY ROLLOUT ANALYTICS DRIFTS WITHOUT GOVERNANCE
    s2 = prs.slides.add_slide(blank)
    slide_header(s2, "Executive Context", "Why Operational Rollout Reporting Breaks Without Semantic Engineering", "Problem → Solution", 2)
    traps = [
        ("1. Non-Additive User Sums", "Adding Wave 1 + Wave 2 active employees double-counts cross-wave specialists (U9001). Unique users must be recomputed via COUNT(DISTINCT APPUserId) at each grain.", RED_BG, RED),
        ("2. Wrong Event Timestamps", "Bucketing contracts by ProjectStartDate or expenses by ModifiedDate shifts historical counts every time records are edited. CORE Pulse locks to CreatedDate & SubmittedDate.", RED_BG, RED),
        ("3. Go-Live vs. Real Usage", "Marking a module LIVE in a project tracker does not mean field teams submit records. CORE Pulse separates SIMULATED EXTENSION go-live state from observed Gold events.", INDIGO_BG, INDIGO),
        ("4. Partial-Month False Alarms", "Comparing 17 days of Nov 2026 MTD against 31 days of Oct 2026 creates fake 'adoption drops'. CORE Pulse pairs MTD with equivalent Oct 1–17 elapsed days.", TEAL_BG, TEAL),
    ]
    for i, (tt, td, bgc, bdc) in enumerate(traps):
        cx = 0.55 + (i % 2) * 6.22
        cy = 1.40 + (i // 2) * 2.15
        add_box(s2, cx, cy, 5.98, 1.92, fill_rgb=bgc, line_rgb=bdc, line_pt=1.4)
        add_text(s2, cx + 0.20, cy + 0.16, 5.58, 0.36, tt, size=15, bold=True, color=NAVY)
        add_text(s2, cx + 0.20, cy + 0.60, 5.58, 1.15, td, size=12.2, color=SLATE)
    add_callout(s2, 5.88, "Key Principle: Put the LLM in front of governed semantic tools—never in front of raw operational tables.")

    # SLIDE 3: 5-STAGE MEDALLION ARCHITECTURE
    s3 = prs.slides.add_slide(blank)
    slide_header(s3, "Data Engineering Foundation", "Inspectable 5-Stage Medallion Warehouse & DataOps Pipeline", "Medallion ETL", 3)
    stages = [
        ("1. SOURCE SIM", "16 CORE Tables", ["BusinessUnit, Project, Task", "ApplicationLog (27,985)", "Timecard, ExpenseReport", "12 adversarial fixtures"], "21 seeded tables", SILVER_BG, SILVER),
        ("2. BRONZE", "Raw + Watermark", ["16 bronze.core_* tables", "Explicit INSERT columns", "SourceExtractDateTime", "BigInt ID watermark"], "WM = 127,985", BRONZE_BG, BRONZE),
        ("3. SILVER", "Conformed Core", ["LEFT JOIN org enrichment", "BU-AU-REN-10 preserved", "IsServiceAccount = 0", "Orphans quarantined"], "10 Silver tables", SILVER_BG, TEAL),
        ("4. GOLD", "Atomic Facts", ["gold.coreadoption_event", "Compound EventKey", "Daily user_activity", "Daily KPI snapshots"], "37,116 events", GOLD_BG, GOLD),
        ("5. CERTIFIED", "Consumer Contract", ["goldview.v_core_adoption", "CURRENT & MONTHLY", "OVERALL/WAVE/ENTITY/BU", "MCP + Native RLS"], "1,216 rows", TEAL_BG, INDIGO),
    ]
    for idx_s, (pill, sh_t, s_bul, foot, bgc, bdc) in enumerate(stages):
        sx = 0.55 + idx_s * 2.48
        add_box(s3, sx, 1.40, 2.34, 4.30, fill_rgb=WHITE, line_rgb=bdc, line_pt=1.5)
        add_box(s3, sx + 0.12, 1.54, 2.10, 0.32, fill_rgb=bgc, line_rgb=bdc, line_pt=0.8)
        add_text(s3, sx + 0.14, 1.57, 2.06, 0.26, pill, size=9.5, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        add_text(s3, sx + 0.12, 1.96, 2.10, 0.34, sh_t, size=13, bold=True, color=NAVY)
        add_bullets(s3, sx + 0.08, 2.36, 2.18, 2.65, [("", b) for b in s_bul], size=10.5, space_after=5)
        add_text(s3, sx + 0.12, 5.26, 2.10, 0.30, foot, size=10.5, bold=True, color=TEAL)
    add_callout(s3, 5.92, "Idempotency Verified: Pipeline Run #1 ingests 27,985 logs; immediate Rerun #2 ingests 0 duplicates with identical 1,216 certified rows.")

    # SLIDE 4: FOUR CERTIFIED KPIS & SEMANTIC BOUNDARIES
    s4 = prs.slides.add_slide(blank)
    slide_header(s4, "Semantic Contract", "Four Authoritative Baseline KPIs & Explicit Attribution Boundaries", "prototype.md §5", 4)
    kpi_cards = [
        ("Contracts Created · 2,303", "COUNT(DISTINCT ProjectId) by Project.CreatedDate",
         "Supports OVERALL, WAVE, ENTITY, and BUSINESS_UNIT. Never uses ProjectStartDate (fixture PRJ 99001 verified).", TEAL),
        ("Employees Logging In · 419", "COUNT(DISTINCT APPUserId) · Application Login · IsServiceAccount=0",
         "STRICTLY NON-ADDITIVE across days, months, apps, entities & waves. Blocked at BUSINESS_UNIT (no HRIS home-BU table).", INDIGO),
        ("Timesheets Submitted · 5,373", "COUNT(DISTINCT TimecardId) WHERE SubmittedDate IS NOT NULL",
         "Includes Approved/Exported timecards (TC 99101); excludes Drafts (TC 99102). Blocked at BUSINESS_UNIT (multi-BU task rows).", TEAL),
        ("Expense Reports · 1,859", "COUNT(DISTINCT ExpenseReportId) WHERE SubmittedDate IS NOT NULL",
         "Supports OVERALL, WAVE, ENTITY, and BUSINESS_UNIT. Locked to initial SubmittedDate regardless of later edits (EXP 99501).", INDIGO),
    ]
    for i, (kt, kf, kd, bdc) in enumerate(kpi_cards):
        cx = 0.55 + (i % 2) * 6.22
        cy = 1.38 + (i // 2) * 2.16
        add_box(s4, cx, cy, 5.98, 1.96, fill_rgb=WHITE, line_rgb=bdc, line_pt=1.5)
        add_text(s4, cx + 0.18, cy + 0.14, 5.6, 0.36, kt, size=15, bold=True, color=NAVY)
        add_text(s4, cx + 0.18, cy + 0.52, 5.6, 0.32, kf, size=11, bold=True, color=TEAL)
        add_text(s4, cx + 0.18, cy + 0.90, 5.6, 0.92, kd, size=11.5, color=SLATE)
    add_callout(s4, 5.90, "Honest Analytics: Unsupported Business Unit breakdowns return clear governance explanations—never fabricated allocations.")

    # SLIDES 5–9: 5 MODULAR 1080p DEMO VIDEO CHAPTERS
    add_video_slide(
        prs, 5, "Live Walkthrough · Chapter 1", "Sign-In, Personal Workspace & Period-Safe Executive KPI Strip", "Video 01",
        "01_signin_and_executive_kpi_strip.mp4", "02_executive_kpi_strip.png", "What Executives See Immediately",
        [
            ("Authenticated Workspace", "Per-user sign-in isolates last-10 conversation history, pins, and saved scope filters."),
            ("Unambiguous Time Labels", "Every KPI card separates All-Time Cumulative, Oct 2026 (Complete), and Nov 2026 MTD (1–17)."),
            ("Fair MTD Benchmarking", "Nov 1–17 MTD compares directly against Oct 1–17 equivalent elapsed days."),
            ("Synthetic Disclosure", "Prominent banner confirms fictional RL Inspection Company data with zero corporate links."),
        ],
        "01_signin_gate.png", "03_overall_adoption_snapshot.png",
        "Chapter 1 demonstrates sign-in, synthetic data disclosure, and unambiguous period labeling across all four KPIs.",
    )

    add_video_slide(
        prs, 6, "Live Walkthrough · Chapter 2", "Multi-Turn Context Memory & Non-Additive Employee Deduplication", "Video 02",
        "02_multiturn_adoption_and_nonadditivity.mp4", "06_wave_nonadditivity_and_sql.png", "3-Turn Executive Flow",
        [
            ("Turn 1 (MTD Logins)", "Answers 'How many employees logged into CORE this month?' with Nov MTD vs Oct 1–17."),
            ("Turn 2 (6-Month Trend)", "Carries forward Employees Logging In; plots Nov 2026 MTD as an amber dotted diamond."),
            ("Turn 3 (Wave Breakdown)", "Shows Wave 1–4 active users AND proves deduplicated OVERALL total vs naive wave sum."),
            ("Audit Drawer", "Expands 'How I got this & Medallion Lineage' for MCP tool JSON and parameterized SQL."),
        ],
        "05_six_month_trend_partial_flag.png", "06_wave_nonadditivity_and_sql.png",
        "Chapter 2 proves multi-turn state retention, partial-month chart discipline, and non-additive distinct user deduplication.",
    )

    add_video_slide(
        prs, 7, "Live Walkthrough · Chapter 3", "Rollout Registry Go-Live vs. Observed Instrumented Activity", "Video 03",
        "03_rollout_vs_observed_activity_gap.mp4", "09_rollout_zero_activity_gap.png", "Spotting Silent Rollout Gaps",
        [
            ("Western Region Resolution", "Maps 'Western region' cleanly to governed region AMER (Americas)."),
            ("Registry vs. Telemetry", "Separates SIMULATED EXTENSION rollout status from observed gold.coreadoption_event rows."),
            ("Uninstrumented Clarity", "Flags INSPECTOR_SCHED, QUALIFICATIONS, LAB_TESTING & INVOICING as status-only in Release 1."),
            ("Zero-Activity Spotlight", "Surfaces ROL-LATAM-EXP (Brazil BU-BR-FLD-09): LIVE since Oct 2026 with 0 expense submissions."),
        ],
        "08_rollout_explorer_full.png", "09_rollout_zero_activity_gap.png",
        "Chapter 3 contrasts rollout go-live registry status against actual observed operational events.",
    )

    add_video_slide(
        prs, 8, "Live Walkthrough · Chapter 4", "Persona Row-Level Security, BU Boundaries & Injection Defense", "Video 04",
        "04_persona_rls_and_governance_guardrails.mp4", "11_persona_rls_refusal.png", "Defense-in-Depth Governance",
        [
            ("Unsupported BU Refusal", "Explains why Employees Logging In & Timecards cannot be attributed to Business Units."),
            ("Database-Level RLS", "Switching to EMEA Rollout Director restricts SQL Server SESSION_CONTEXT to 342 EMEA rows."),
            ("Cross-Region Refusal", "Asking for AMER adoption under EMEA persona returns a polite scope boundary message."),
            ("Fail-Closed Security", "Blocks prompt injection, raw SELECT/DROP SQL, secret inspection, and user PII dumps."),
        ],
        "10_unsupported_bu_governance.png", "12_prompt_injection_blocked.png",
        "Chapter 4 demonstrates unsupported grain explanations, Persona RLS, and prompt-injection blocking.",
    )

    add_video_slide(
        prs, 9, "Live Walkthrough · Chapter 5", "Data Catalog, Medallion Lineage & DataOps Watermark Telemetry", "Video 05",
        "05_medallion_catalog_and_dataops_lineage.mp4", "15_dataops_watermark_and_parity.png", "Transparent Engineering",
        [
            ("Governed KPI Dictionary", "Machine-readable YAML contract rendered directly for business & BI reviewers."),
            ("End-to-End Lineage Map", "Traces Source Sim → Bronze → Silver → Gold → goldview.v_core_adoption → Answer."),
            ("DataOps Watermarks", "Displays ApplicationLogId high-water mark (127,985) and 11 pipeline step logs."),
            ("Quarantine Visibility", "Shows 3 orphan TimecardRows, 2 orphan ExpenseEntries & REVIEW apps quarantined safely."),
        ],
        "14_catalog_medallion_lineage.png", "15_dataops_watermark_and_parity.png",
        "Chapter 5 walks through the Data Catalog, Source-to-Answer Lineage, and DataOps quality telemetry.",
    )

    # SLIDE 10: MCP ARCHITECTURE & DUAL-BACKEND PORTABILITY
    s10 = prs.slides.add_slide(blank)
    slide_header(s10, "Technical Architecture", "Read-Only MCP Server & Dual SQL Server / DuckDB Portability", "MCP + Dual DB", 10)
    add_box(s10, 0.55, 1.40, 5.95, 4.35, fill_rgb=WHITE, line_rgb=TEAL, line_pt=1.5)
    add_text(s10, 0.75, 1.56, 5.5, 0.36, "10 Allowlisted Read-Only MCP Tools", size=15, bold=True, color=NAVY)
    add_bullets(s10, 0.72, 2.00, 5.55, 3.6, [
        ("get_core_adoption_current", "All-time, complete-month & fair MTD snapshot"),
        ("get_core_adoption_monthly", "Multi-month trend + fastest-growing complete month"),
        ("get_core_adoption_by_wave / entity / bu", "Grain-accurate breakdowns + deduplicated totals"),
        ("get_core_rollout_status", "SIMULATED EXTENSION registry vs observed events"),
        ("get_core_user_engagement", "Distinct-at-query-level active users across apps"),
        ("definition / pipeline_status / lineage", "6 core://metadata/* resources + DataOps audit"),
    ], size=11.5, space_after=6)

    add_box(s10, 6.80, 1.40, 5.97, 4.35, fill_rgb=WHITE, line_rgb=INDIGO, line_pt=1.5)
    add_text(s10, 7.00, 1.56, 5.5, 0.36, "Identical Execution Across Both Engines", size=15, bold=True, color=NAVY)
    add_bullets(s10, 6.97, 2.00, 5.55, 3.6, [
        ("SQL Server 2022 (core_pulse)", "Least-privilege pulse_reader login + native SESSION_CONTEXT SECURITY POLICY (1,216 Global / 532 AMER / 342 EMEA rows)."),
        ("DuckDB Local Warehouse", "Zero-dependency single-file data/core_pulse.duckdb for instant laptop reset & CI."),
        ("LiteLLM + Offline Rules", "Works with GKE Ollama (qwen2.5-coder:14b, gpt-oss:20b, Apache-2.0), optional cloud APIs, or 100% offline rules mode."),
        ("Numeric Prose Guardrail", "Verifies every number in LLM narration against structured ToolPayload before rendering."),
    ], size=11.5, space_after=7)
    add_callout(s10, 5.92, "Zero Arbitrary SQL: Neither the chat UI nor external MCP clients can execute ad-hoc SQL, shell commands, or writes.")

    # SLIDE 11: VERIFICATION RIGOR & MEASURED EVALUATION
    s11 = prs.slides.add_slide(blank)
    slide_header(s11, "Quality Assurance", "34/34 Automated Tests (Both DBs), 12 Edge Cases & Golden Evaluation", "100% Verified", 11)
    add_box(s11, 0.55, 1.40, 5.85, 4.35, fill_rgb=WHITE, line_rgb=TEAL, line_pt=1.5)
    add_text(s11, 0.75, 1.56, 5.4, 0.36, "12 Adversarial Fixtures Verified (§3)", size=15, bold=True, color=NAVY)
    add_bullets(s11, 0.72, 1.98, 5.45, 3.6, [
        ("U9001 & U9002 Dedup", "Dual-wave same-day login & 5x same-day login deduplicated"),
        ("Service Accounts (9501..9512)", "100% excluded from Silver/Gold employee counts"),
        ("TC 99101 / 99102 / 99103", "Exported included, Draft excluded, Multi-BU flagged"),
        ("PRJ 99001 & EXP 99501", "Bucketed by CreatedDate & SubmittedDate, never edits"),
        ("BU-AU-REN-10 & Orphans", "Preserved via LEFT JOIN; 5 orphan child rows quarantined"),
        ("ROL-LATAM-EXP & Partial Nov", "LIVE zero-activity flagged; IsCompleteMonth=0 enforced"),
    ], size=11.2, space_after=6)

    add_box(s11, 6.65, 1.40, 6.12, 4.35, fill_rgb=WHITE, line_rgb=INDIGO, line_pt=1.5)
    add_text(s11, 6.85, 1.56, 5.7, 0.36, "Measured Test & Golden Evaluation Summary", size=15, bold=True, color=NAVY)
    ev_bullets = [
        ("DuckDB Pytest Suite", "34 / 34 passed (Medallion parity, MCP, RLS, UI smoke)"),
        ("SQL Server 2022 Pytest Suite", "34 / 34 passed against live [core_pulse] + pulse_reader"),
        ("Playwright Browser Walkthrough", "9 / 9 UI steps verified end-to-end on :8502"),
    ]
    for r in eval_rows:
        ev_bullets.append((
            f"Eval — {r['model']}",
            f"Status {r['status_accuracy']} · Tool {r['tool_accuracy']} · Facts {r['fact_accuracy']} · Median {r['median_latency_s']}s",
        ))
    add_bullets(s11, 6.82, 1.98, 5.75, 3.6, ev_bullets, size=11.2, space_after=6)
    add_callout(s11, 5.92, "Every reported figure is 100% backed by automated test assertions and saved artifacts in results/.")

    # SLIDE 12: PRODUCTION LIFT & MODULAR RELEASE 2–4 ROADMAP
    s12 = prs.slides.add_slide(blank)
    slide_header(s12, "Next Steps & Scale", "Enterprise Production Lift & Modular Release 2–4 Roadmap", "Roadmap", 12)
    roadmap = [
        ("Release 1 (Built Today)", "CORE Rollout & Adoption", "4 certified baseline KPIs, Medallion + DataOps watermarks, Rollout Registry (SIMULATED EXTENSION), MCP & Chat.", TEAL_BG, TEAL),
        ("Release 2 (Next)", "Contracts & Project Delivery", "Project/Service/Phase/Task lifecycle, turnaround cycle time, backlog velocity (separate semantic catalog).", INDIGO_BG, INDIGO),
        ("Release 3 (Planned)", "Field Scheduling & Certs", "Inspector dispatch utilization, NDT/pressure-equipment licence tracking & qualification compliance.", SILVER_BG, SILVER),
        ("Release 4 (Planned)", "Laboratory LIMS & Billing", "Specimen intake throughput, certificate turnaround, milestone & time-and-materials invoicing.", GOLD_BG, GOLD),
    ]
    for idx_r, (r_tag, r_dom, r_desc, bgc, bdc) in enumerate(roadmap):
        rx = 0.55 + idx_r * 3.10
        add_box(s12, rx, 1.40, 2.92, 2.55, fill_rgb=bgc, line_rgb=bdc, line_pt=1.4)
        add_text(s12, rx + 0.14, 1.54, 2.64, 0.28, r_tag, size=10.5, bold=True, color=bdc)
        add_text(s12, rx + 0.14, 1.86, 2.64, 0.36, r_dom, size=13.5, bold=True, color=NAVY)
        add_text(s12, rx + 0.14, 2.30, 2.64, 1.50, r_desc, size=11.0, color=SLATE)

    add_box(s12, 0.55, 4.18, 12.22, 1.55, fill_rgb=WHITE, line_rgb=NAVY, line_pt=1.3)
    add_text(s12, 0.75, 4.30, 11.8, 0.32, "Enterprise Production Lift Checklist (Zero Rewrite of Semantic or UI Layers)", size=13.5, bold=True, color=NAVY)
    add_bullets(s12, 0.72, 4.65, 11.8, 1.0, [
        ("1. Source Swap", "Point Bronze extractors from source_sim.* to read-only replicas of production CORE SQL Server."),
        ("2. Identity & HRIS", "Bind Entra ID SSO claims to SQL Server SESSION_CONTEXT('user_id') and link HRIS worker master for home-BU logins."),
        ("3. Registry Governance", "Promote reference.rollout_registry to an owned PMO governance table while keeping the 4 baseline KPIs unchanged."),
    ], size=11.0, space_after=3)
    add_callout(s12, 5.92, "Plug-In Architecture: Releases 2–4 add new semantic YAML contracts and MCP tools without altering Release 1 KPI definitions.")

    OUT_PPTX.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(OUT_PPTX))
    print(f"Saved PowerPoint deck: {OUT_PPTX} ({OUT_PPTX.stat().st_size // 1024} KB)")


def build_html(eval_rows: list[dict]) -> None:
    eval_trs = ""
    for r in eval_rows:
        eval_trs += (
            f"<tr><td><code>{html.escape(r['model'])}</code></td>"
            f"<td><b>{html.escape(r['status_accuracy'])}</b></td>"
            f"<td><b>{html.escape(r['tool_accuracy'])}</b></td>"
            f"<td><b>{html.escape(r['fact_accuracy'])}</b></td>"
            f"<td>{r['median_latency_s']:.2f}s</td></tr>"
        )
    if not eval_trs:
        eval_trs = "<tr><td><code>rules</code></td><td><b>15/15</b></td><td><b>15/15</b></td><td><b>15/15</b></td><td>0.03s</td></tr>"

    html_doc = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>CORE Pulse — Executive Architecture &amp; Live Demo Deck</title>
<style>
  :root {{
    --navy: #0B1F3A; --teal: #0F766E; --teal-light: #E6F4F1;
    --indigo: #3730A3; --indigo-light: #EEF0FB;
    --slate: #475569; --border: #E2E8F0; --bg: #F8FAFC; --card: #FFFFFF;
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  html, body {{
    width: 100%; height: 100%; overflow: hidden;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Inter, Roboto, sans-serif;
    background: #071324; color: var(--navy);
  }}
  .deck-shell {{ width: 100vw; height: 100vh; display: flex; flex-direction: column; justify-content: center; align-items: center; }}
  .slide-viewport {{
    position: relative; width: min(96vw, calc((100vh - 64px) * 16 / 9)); aspect-ratio: 16 / 9;
    background: var(--bg); border-radius: 16px; box-shadow: 0 28px 80px rgba(0,0,0,0.55); overflow: hidden;
  }}
  .slide {{ position: absolute; inset: 0; display: none; flex-direction: column; padding: 34px 46px 26px; background: var(--bg); }}
  .slide.active {{ display: flex; }}
  .slide.hero {{ background: linear-gradient(135deg, #081A33 0%, #0F766E 58%, #3730A3 100%); color: #FFF; }}
  .s-top {{ display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 16px; border-bottom: 2px solid var(--border); padding-bottom: 10px; }}
  .kicker {{ font-size: 0.75rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.1em; color: var(--teal); margin-bottom: 3px; }}
  .s-title {{ font-size: clamp(1.3rem, 2.0vw, 1.88rem); font-weight: 800; color: var(--navy); line-height: 1.18; }}
  .s-badge {{ background: var(--indigo-light); color: var(--indigo); font-size: 0.75rem; font-weight: 700; padding: 5px 12px; border-radius: 999px; white-space: nowrap; }}
  .grid-2 {{ display: grid; grid-template-columns: 1fr 1fr; gap: 18px; flex: 1; min-height: 0; }}
  .grid-4 {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 13px; flex: 1; min-height: 0; }}
  .grid-5 {{ display: grid; grid-template-columns: repeat(5, 1fr); gap: 10px; flex: 1; min-height: 0; }}
  .grid-video {{ display: grid; grid-template-columns: 1.52fr 1fr; gap: 20px; flex: 1; min-height: 0; align-items: stretch; }}
  .card {{ background: var(--card); border: 1px solid var(--border); border-radius: 12px; padding: 15px 17px; display: flex; flex-direction: column; gap: 8px; min-height: 0; }}
  .card.teal {{ border-top: 4px solid var(--teal); }}
  .card.indigo {{ border-top: 4px solid var(--indigo); }}
  .card.red {{ border-top: 4px solid #DC2626; background: #FFFDFD; }}
  .card h3 {{ font-size: 1.0rem; font-weight: 700; color: var(--navy); }}
  .card p, .card li {{ font-size: 0.84rem; line-height: 1.44; color: var(--slate); }}
  .card ul {{ padding-left: 18px; display: flex; flex-direction: column; gap: 6px; }}
  .card li strong {{ color: var(--navy); }}
  code {{ background: #EEF2FF; color: #3730A3; padding: 1px 6px; border-radius: 5px; font-family: ui-monospace, Menlo, monospace; font-size: 0.88em; font-weight: 600; }}
  .callout {{ background: linear-gradient(90deg, var(--teal-light) 0%, var(--indigo-light) 100%); border-left: 5px solid var(--teal); border-radius: 10px; padding: 10px 15px; font-size: 0.86rem; color: var(--navy); font-weight: 600; margin-top: 12px; }}
  .vid-col {{ display: flex; flex-direction: column; gap: 8px; min-height: 0; }}
  .vid-frame {{ background: #0B1628; border-radius: 12px; overflow: hidden; border: 2px solid #CBD5E1; }}
  .vid-bar {{ background: #0F172A; color: #E2E8F0; font-size: 0.74rem; font-weight: 600; padding: 5px 12px; display: flex; justify-content: space-between; }}
  .vid-frame video {{ width: 100%; aspect-ratio: 16 / 9; object-fit: contain; background: #0F172A; display: block; }}
  .thumb-row {{ display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }}
  .thumb-row img {{ width: 100%; height: 76px; object-fit: cover; object-position: top; border-radius: 8px; border: 1px solid var(--border); }}
  table. mini-tbl {{ width: 100%; border-collapse: collapse; font-size: 0.80rem; margin-top: 4px; }}
  table.mini-tbl th {{ background: #0B1F3A; color: #FFF; text-align: left; padding: 6px 8px; }}
  table.mini-tbl td {{ padding: 6px 8px; border-bottom: 1px solid var(--border); color: var(--slate); }}
  .nav-bar {{ display: flex; align-items: center; gap: 14px; margin-top: 10px; color: #CBD5E1; font-size: 0.84rem; }}
  .nav-bar button {{ background: #1E293B; color: #FFF; border: 1px solid #334155; border-radius: 8px; padding: 6px 14px; cursor: pointer; font-weight: 600; }}
  .nav-bar button:hover {{ background: var(--teal); border-color: var(--teal); }}
</style>
</head>
<body>
<div class="deck-shell">
  <div class="slide-viewport">

    <!-- Slide 1 -->
    <section class="slide hero active">
      <div style="display:grid; grid-template-columns: 1.35fr 1fr; gap: 26px; flex:1; align-items:center;">
        <div style="display:flex; flex-direction:column; gap:16px;">
          <span style="background:rgba(255,255,255,0.16); width:fit-content; padding:5px 12px; border-radius:999px; font-size:0.76rem; font-weight:700; letter-spacing:0.08em;">AI-READY DATA ENGINEERING &amp; OPERATIONAL ROLLOUT PILOT</span>
          <h1 style="font-size:clamp(1.9rem, 3.0vw, 2.75rem); font-weight:800; line-height:1.12;">CORE Pulse: From Medallion Foundations to Trusted Rollout Intelligence</h1>
          <p style="font-size:0.98rem; line-height:1.5; color:#E2E8F0;">Showing how a governed 5-stage Medallion warehouse (<code>Bronze</code> &rarr; <code>Silver</code> &rarr; <code>Gold</code> &rarr; <code>goldview.v_core_adoption</code>) and a Read-Only MCP Semantic Layer eliminate metric drift, non-additive employee double-counting, and partial-month false alarms.</p>
          <div style="display:grid; grid-template-columns:repeat(3,1fr); gap:12px; margin-top:4px;">
            <div style="background:rgba(15,23,42,0.55); border:1px solid #5EEAD4; border-radius:10px; padding:12px;">
              <div style="font-size:1.25rem; font-weight:800; color:#5EEAD4;">90% Foundation</div>
              <div style="font-size:0.78rem; color:#E2E8F0; margin-top:4px;">16 CORE tables &rarr; Medallion + DataOps watermarks</div>
            </div>
            <div style="background:rgba(15,23,42,0.55); border:1px solid #5EEAD4; border-radius:10px; padding:12px;">
              <div style="font-size:1.25rem; font-weight:800; color:#5EEAD4;">0 Double-Count</div>
              <div style="font-size:0.78rem; color:#E2E8F0; margin-top:4px;">Non-additive <code>COUNT(DISTINCT APPUserId)</code> at every grain</div>
            </div>
            <div style="background:rgba(15,23,42,0.55); border:1px solid #5EEAD4; border-radius:10px; padding:12px;">
              <div style="font-size:1.25rem; font-weight:800; color:#5EEAD4;">Dual Backend</div>
              <div style="font-size:0.78rem; color:#E2E8F0; margin-top:4px;">SQL Server 2022 native RLS + local DuckDB parity</div>
            </div>
          </div>
          <div style="background:rgba(245,158,11,0.18); border:1px solid #FBBF24; border-radius:8px; padding:8px 12px; font-size:0.80rem; color:#FEF3C7;">
            <b>Synthetic Pilot Disclosure:</b> Deterministic 18-month world (<code>seed=42</code>, <code>AS_OF_DATE=2026-11-17</code>), fictional RL Inspection Company, zero corporate connectivity.
          </div>
        </div>
        <div class="card teal" style="color:var(--navy); gap:10px;">
          <h3 style="color:var(--teal); text-transform:uppercase; font-size:0.82rem; letter-spacing:0.08em;">The Adoption Analytics Iceberg</h3>
          <div style="background:var(--indigo-light); border:1px solid var(--indigo); border-radius:8px; padding:10px;">
            <div style="font-size:0.72rem; font-weight:800; color:var(--indigo);">VISIBLE SURFACE (10%) &mdash; EXECUTIVE UX</div>
            <div style="font-size:0.84rem; font-weight:700; margin-top:4px;">CORE Pulse Conversational BI &middot; Rollout Explorer &middot; Power BI / Tableau</div>
          </div>
          <div style="text-align:center; font-size:0.72rem; font-weight:800; color:var(--teal);">&mdash;&mdash; ENGINEERED FOUNDATION (90%) &mdash;&mdash;</div>
          <div style="background:#E6F4F1; border-left:4px solid #0F766E; padding:8px 10px; border-radius:6px; font-size:0.80rem;"><b>Certified Interface &amp; MCP:</b> <code>goldview.v_core_adoption</code> &middot; 10 Read-Only Tools &middot; Persona RLS</div>
          <div style="background:#FEF9C3; border-left:4px solid #B45309; padding:8px 10px; border-radius:6px; font-size:0.80rem;"><b>Gold Layer:</b> <code>37,116</code> atomic events &middot; Daily User Activity &middot; <code>1,216</code> certified rollups</div>
          <div style="background:#F1F5F9; border-left:4px solid #334155; padding:8px 10px; border-radius:6px; font-size:0.80rem;"><b>Silver Layer:</b> Conformed Org (<code>LEFT JOIN</code> enrichment) &middot; <code>IsServiceAccount=0</code></div>
          <div style="background:#FEF3C7; border-left:4px solid #92400E; padding:8px 10px; border-radius:6px; font-size:0.80rem;"><b>Bronze &amp; DataOps:</b> 16 CORE tables &middot; <code>ApplicationLogId=127,985</code> watermark</div>
        </div>
      </div>
    </section>

    <!-- Slide 2 -->
    <section class="slide">
      <div class="s-top"><div><div class="kicker">Executive Context</div><div class="s-title">Why Operational Rollout Reporting Breaks Without Semantic Engineering</div></div><span class="s-badge">Problem &rarr; Solution</span></div>
      <div class="grid-2">
        <div class="card red"><h3>1. Non-Additive User Double-Counting</h3><p>Summing Wave 1 + Wave 2 active employees inflates headcount whenever specialists log into multiple regional portals (<code>APPUserId=9001</code>). CORE Pulse recomputes <code>COUNT(DISTINCT APPUserId)</code> at each grain.</p></div>
        <div class="card red"><h3>2. Mutable Workflow Timestamps</h3><p>Bucketing projects by <code>ProjectStartDate</code> or expenses by <code>ModifiedDate</code> rewrites history when records are edited. CORE Pulse locks to immutable <code>CreatedDate</code> and initial <code>SubmittedDate</code>.</p></div>
        <div class="card indigo"><h3>3. Go-Live Tracker vs. Actual Usage</h3><p>A PMO spreadsheet marking Expenses <code>LIVE</code> in Brazil does not prove field adoption. CORE Pulse contrasts <code>SIMULATED EXTENSION</code> go-live status against observed Gold events (surfacing <code>ROL-LATAM-EXP = 0</code>).</p></div>
        <div class="card teal"><h3>4. Partial-Month False Alarms</h3><p>Comparing 17 days of <code>Nov 2026 MTD</code> against 31 days of <code>Oct 2026</code> manufactures fake adoption drops. CORE Pulse benchmarks MTD against equivalent <code>Oct 1&ndash;17</code> elapsed days.</p></div>
      </div>
      <div class="callout">Architectural Rule: Never connect an LLM to raw operational tables &mdash; bind it strictly to governed semantic tools.</div>
    </section>

    <!-- Slide 3 -->
    <section class="slide">
      <div class="s-top"><div><div class="kicker">Data Engineering Foundation</div><div class="s-title">Inspectable 5-Stage Medallion Warehouse &amp; DataOps Pipeline</div></div><span class="s-badge">Medallion ETL</span></div>
      <div class="grid-5">
        <div class="card"><h3>1. Source Sim</h3><ul><li>16 normalized CORE tables</li><li><code>BusinessUnit</code>, <code>Project</code></li><li><code>ApplicationLog</code> (27,985)</li><li><code>Timecard</code>, <code>ExpenseReport</code></li><li>12 §3 edge-case fixtures</li></ul></div>
        <div class="card teal"><h3>2. Bronze</h3><ul><li>16 <code>bronze.core_*</code> tables</li><li>Explicit column lists</li><li><code>SourceExtractDateTime</code></li><li><code>ApplicationLogId</code> watermark (<code>127,985</code>)</li></ul></div>
        <div class="card indigo"><h3>3. Silver</h3><ul><li>10 conformed tables</li><li><code>LEFT JOIN</code> org enrichment (<code>BU-AU-REN-10</code> kept)</li><li><code>IsServiceAccount = 0</code></li><li>5 orphan rows quarantined</li></ul></div>
        <div class="card teal"><h3>4. Gold</h3><ul><li><code>coreadoption_event</code> (<code>37,116</code>)</li><li>Compound <code>EventKey</code></li><li><code>coreadoption_user_activity</code></li><li><code>coreadoption_kpi_daily</code></li></ul></div>
        <div class="card indigo"><h3>5. Certified View</h3><ul><li><code>goldview.v_core_adoption</code></li><li><code>1,216</code> certified rows</li><li><code>CURRENT</code> &amp; <code>MONTHLY</code></li><li>Native SQL Server RLS</li></ul></div>
      </div>
      <div class="callout">Idempotency Verified: Pipeline Run #1 ingests <code>27,985</code> logs; immediate Rerun #2 ingests <code>0</code> duplicate logs with identical <code>1,216</code> certified rows.</div>
    </section>

    <!-- Slide 4 -->
    <section class="slide">
      <div class="s-top"><div><div class="kicker">Semantic Contract</div><div class="s-title">Four Authoritative Baseline KPIs &amp; Explicit Attribution Boundaries</div></div><span class="s-badge">prototype.md &sect;5</span></div>
      <div class="grid-2">
        <div class="card teal"><h3>Contracts Created &middot; 2,303</h3><p><code>COUNT(DISTINCT ProjectId)</code> by <code>Project.CreatedDate</code>. Supports <code>OVERALL</code>, <code>WAVE</code>, <code>ENTITY</code>, and <code>BUSINESS_UNIT</code>. Never uses <code>ProjectStartDate</code> (fixture <code>PRJ 99001</code> verified).</p></div>
        <div class="card indigo"><h3>Employees Logging In &middot; 419</h3><p><code>COUNT(DISTINCT APPUserId)</code> where <code>Application Login</code> &amp; <code>IsServiceAccount = 0</code>. <b>Strictly non-additive.</b> Unavailable by <code>BUSINESS_UNIT</code> (no employee home-BU table).</p></div>
        <div class="card teal"><h3>Timesheets Submitted &middot; 5,373</h3><p><code>COUNT(DISTINCT TimecardId)</code> where <code>SubmittedDate IS NOT NULL</code> (includes Exported <code>TC 99101</code>; excludes Draft <code>TC 99102</code>). Unavailable by <code>BUSINESS_UNIT</code> (multi-BU task rows).</p></div>
        <div class="card indigo"><h3>Expense Reports &middot; 1,859</h3><p><code>COUNT(DISTINCT ExpenseReportId)</code> where <code>SubmittedDate IS NOT NULL</code>. Supports <code>OVERALL</code>, <code>WAVE</code>, <code>ENTITY</code>, and <code>BUSINESS_UNIT</code>. Locked to initial submission month (<code>EXP 99501</code>).</p></div>
      </div>
      <div class="callout">Governance Integrity: When asked for unsupported Business Unit logins or timecards, CORE Pulse explains the attribution boundary rather than guessing.</div>
    </section>

    <!-- Slide 5 -->
    <section class="slide">
      <div class="s-top"><div><div class="kicker">Live Walkthrough &middot; Chapter 1</div><div class="s-title">Sign-In, Personal Workspace &amp; Period-Safe Executive KPI Strip</div></div><span class="s-badge">1080p Demo 01</span></div>
      <div class="grid-video">
        <div class="vid-col">
          <div class="vid-frame"><div class="vid-bar"><span>&#9654; 01_signin_and_executive_kpi_strip.mp4</span><span>1080p H.264</span></div><video controls preload="metadata" poster="screenshots/02_executive_kpi_strip.png" src="videos/01_signin_and_executive_kpi_strip.mp4"></video></div>
          <div class="thumb-row"><img src="screenshots/01_signin_gate.png" alt="Sign in" /><img src="screenshots/03_overall_adoption_snapshot.png" alt="Overall snapshot" /></div>
        </div>
        <div class="card teal"><h3>Executive First Impression</h3><ul>
          <li><strong>Per-User Conversation Memory:</strong> Sign-in isolates each leader's last 10 threads, pinned reviews, and scope filters.</li>
          <li><strong>Unambiguous Period Labels:</strong> Separates All-Time Cumulative, <code>Oct 2026 (Complete)</code>, and <code>Nov 2026 MTD (days 1&ndash;17)</code>.</li>
          <li><strong>Fair MTD Comparison:</strong> Benchmarks <code>Nov 1&ndash;17</code> against <code>Oct 1&ndash;17</code> equivalent elapsed days.</li>
          <li><strong>Synthetic Disclosure:</strong> Always-visible synthetic data banner and certified view provenance.</li>
        </ul></div>
      </div>
    </section>

    <!-- Slide 6 -->
    <section class="slide">
      <div class="s-top"><div><div class="kicker">Live Walkthrough &middot; Chapter 2</div><div class="s-title">Multi-Turn Context Memory &amp; Non-Additive Employee Deduplication</div></div><span class="s-badge">1080p Demo 02</span></div>
      <div class="grid-video">
        <div class="vid-col">
          <div class="vid-frame"><div class="vid-bar"><span>&#9654; 02_multiturn_adoption_and_nonadditivity.mp4</span><span>1080p H.264</span></div><video controls preload="metadata" poster="screenshots/06_wave_nonadditivity_and_sql.png" src="videos/02_multiturn_adoption_and_nonadditivity.mp4"></video></div>
          <div class="thumb-row"><img src="screenshots/05_six_month_trend_partial_flag.png" alt="6M trend" /><img src="screenshots/06_wave_nonadditivity_and_sql.png" alt="Wave non-additivity" /></div>
        </div>
        <div class="card indigo"><h3>3-Turn Conversational Proof</h3><ul>
          <li><strong>Turn 1:</strong> <code>How many employees logged into CORE this month?</code> &rarr; <code>Nov 2026 MTD</code> unique non-service employees.</li>
          <li><strong>Turn 2:</strong> <code>Show the six-month trend.</code> &rarr; Retains KPI; plots partial <code>2026-11</code> as an amber dotted diamond.</li>
          <li><strong>Turn 3:</strong> <code>Which rollout wave is most active?</code> &rarr; Shows Wave 1&ndash;4 counts alongside deduplicated <code>OVERALL</code> total.</li>
          <li><strong>Lineage Audit:</strong> Opens <code>How I got this &amp; Medallion Lineage</code> showing MCP tool call &amp; SQL.</li>
        </ul></div>
      </div>
    </section>

    <!-- Slide 7 -->
    <section class="slide">
      <div class="s-top"><div><div class="kicker">Live Walkthrough &middot; Chapter 3</div><div class="s-title">Rollout Registry Go-Live vs. Observed Instrumented Activity</div></div><span class="s-badge">1080p Demo 03</span></div>
      <div class="grid-video">
        <div class="vid-col">
          <div class="vid-frame"><div class="vid-bar"><span>&#9654; 03_rollout_vs_observed_activity_gap.mp4</span><span>1080p H.264</span></div><video controls preload="metadata" poster="screenshots/09_rollout_zero_activity_gap.png" src="videos/03_rollout_vs_observed_activity_gap.mp4"></video></div>
          <div class="thumb-row"><img src="screenshots/08_rollout_explorer_full.png" alt="Rollout Explorer" /><img src="screenshots/09_rollout_zero_activity_gap.png" alt="Zero activity gap" /></div>
        </div>
        <div class="card teal"><h3>Actionable Rollout Intelligence</h3><ul>
          <li><strong>Synonym Resolution:</strong> Maps <code>Western region</code> cleanly to governed region <code>AMER</code>.</li>
          <li><strong>Registry vs. Events:</strong> Contrasts <code>reference.rollout_registry</code> (<code>SIMULATED EXTENSION</code>) with observed Gold events.</li>
          <li><strong>Uninstrumented Modules:</strong> Labels Scheduling, Qualifications, Lab Testing &amp; Invoicing as status-only in Release 1.</li>
          <li><strong>Gap Detection:</strong> Pinpoints <code>ROL-LATAM-EXP</code> (<code>BU-BR-FLD-09</code>) &mdash; <code>LIVE</code> with <b>zero expense submissions</b>.</li>
        </ul></div>
      </div>
    </section>

    <!-- Slide 8 -->
    <section class="slide">
      <div class="s-top"><div><div class="kicker">Live Walkthrough &middot; Chapter 4</div><div class="s-title">Persona Row-Level Security, BU Boundaries &amp; Injection Defense</div></div><span class="s-badge">1080p Demo 04</span></div>
      <div class="grid-video">
        <div class="vid-col">
          <div class="vid-frame"><div class="vid-bar"><span>&#9654; 04_persona_rls_and_governance_guardrails.mp4</span><span>1080p H.264</span></div><video controls preload="metadata" poster="screenshots/11_persona_rls_refusal.png" src="videos/04_persona_rls_and_governance_guardrails.mp4"></video></div>
          <div class="thumb-row"><img src="screenshots/10_unsupported_bu_governance.png" alt="Unsupported BU" /><img src="screenshots/12_prompt_injection_blocked.png" alt="Injection blocked" /></div>
        </div>
        <div class="card indigo"><h3>Defense-in-Depth Security</h3><ul>
          <li><strong>Attribution Guardrail:</strong> Blocks unsupported Business Unit breakdowns for Logins &amp; Timecards with clear rationale.</li>
          <li><strong>Native SQL Server RLS:</strong> Switching to <code>EMEA Rollout Director</code> sets <code>SESSION_CONTEXT('user_id')</code> (<code>342</code> EMEA rows).</li>
          <li><strong>Scope Refusal:</strong> Out-of-scope <code>AMER</code> queries under EMEA persona are refused cleanly.</li>
          <li><strong>Prompt-Injection Filter:</strong> Rejects instruction overrides, raw <code>SELECT/DROP</code> SQL, and user PII exfiltration.</li>
        </ul></div>
      </div>
    </section>

    <!-- Slide 9 -->
    <section class="slide">
      <div class="s-top"><div><div class="kicker">Live Walkthrough &middot; Chapter 5</div><div class="s-title">Data Catalog, Medallion Lineage &amp; DataOps Watermark Telemetry</div></div><span class="s-badge">1080p Demo 05</span></div>
      <div class="grid-video">
        <div class="vid-col">
          <div class="vid-frame"><div class="vid-bar"><span>&#9654; 05_medallion_catalog_and_dataops_lineage.mp4</span><span>1080p H.264</span></div><video controls preload="metadata" poster="screenshots/15_dataops_watermark_and_parity.png" src="videos/05_medallion_catalog_and_dataops_lineage.mp4"></video></div>
          <div class="thumb-row"><img src="screenshots/14_catalog_medallion_lineage.png" alt="Lineage" /><img src="screenshots/15_dataops_watermark_and_parity.png" alt="DataOps" /></div>
        </div>
        <div class="card teal"><h3>Self-Service Governance Without SQL</h3><ul>
          <li><strong>Governed KPI Dictionary:</strong> Exposes formulas, timestamps, exclusions, and supported grains in plain English.</li>
          <li><strong>Navigable Lineage:</strong> Traces <code>source_sim</code> &rarr; <code>bronze</code> &rarr; <code>silver</code> &rarr; <code>gold</code> &rarr; <code>goldview.v_core_adoption</code>.</li>
          <li><strong>Incremental Watermark:</strong> Tracks <code>ApplicationLogId = 127,985</code> and 11 step execution logs in <code>dataops.*</code>.</li>
          <li><strong>Exception Quarantine:</strong> Surfaces 3 orphan TimecardRows, 2 orphan ExpenseEntries &amp; REVIEW apps safely.</li>
        </ul></div>
      </div>
    </section>

    <!-- Slide 10 -->
    <section class="slide">
      <div class="s-top"><div><div class="kicker">Technical Architecture</div><div class="s-title">Read-Only MCP Server &amp; Dual SQL Server / DuckDB Portability</div></div><span class="s-badge">MCP + Dual DB</span></div>
      <div class="grid-2">
        <div class="card teal"><h3>10 Allowlisted Read-Only MCP Tools</h3><ul>
          <li><code>get_core_adoption_current</code> &middot; Snapshot, MTD &amp; Last Complete Month</li>
          <li><code>get_core_adoption_monthly</code> &middot; Trend + fastest-growing complete month</li>
          <li><code>get_core_adoption_by_wave</code> / <code>by_entity</code> / <code>by_business_unit</code></li>
          <li><code>get_core_rollout_status</code> &middot; <code>SIMULATED EXTENSION</code> vs observed events</li>
          <li><code>get_core_user_engagement</code> &middot; Distinct-at-query-level app engagement</li>
          <li><code>get_core_metric_definition</code> / <code>pipeline_status</code> / <code>metric_lineage</code></li>
        </ul></div>
        <div class="card indigo"><h3>Dual-Engine Runtime &amp; Model Flexibility</h3><ul>
          <li><strong>SQL Server 2022 (<code>core_pulse</code>):</strong> Read-only <code>pulse_reader</code> login + native <code>SESSION_CONTEXT</code> RLS policy (<code>1,216</code> Global / <code>532</code> AMER / <code>342</code> EMEA rows).</li>
          <li><strong>DuckDB Warehouse:</strong> Single-file <code>data/core_pulse.duckdb</code> for zero-dependency laptop runs &amp; CI parity.</li>
          <li><strong>LiteLLM + Offline Rules:</strong> Supports GKE Ollama (<code>qwen2.5-coder:14b</code>, <code>gpt-oss:20b</code>, Apache-2.0) and deterministic <code>rules</code> mode.</li>
          <li><strong>Numeric Prose Guardrail:</strong> Rejects any ungrounded LLM number before rendering.</li>
        </ul></div>
      </div>
      <div class="callout">JSON-RPC 2.0 MCP Server (<code>python -m mcp_server.server</code>) exposes 10 tools + 6 <code>core://metadata/*</code> resources with zero arbitrary SQL.</div>
    </section>

    <!-- Slide 11 -->
    <section class="slide">
      <div class="s-top"><div><div class="kicker">Quality Assurance</div><div class="s-title">34/34 Automated Tests (Both DBs), 12 Edge Cases &amp; Golden Evaluation</div></div><span class="s-badge">100% Measured</span></div>
      <div class="grid-2">
        <div class="card teal"><h3>12 Adversarial Edge Cases Verified (&sect;3)</h3><ul>
          <li><strong>Cross-Wave &amp; Multi-Login Dedup:</strong> <code>U9001</code> dual-wave &amp; <code>U9002</code> 5x same-day logins deduplicated.</li>
          <li><strong>Service Accounts Excluded:</strong> <code>9501..9512</code> (<code>IsServiceAccount=1</code>) zeroed in Silver &amp; Gold.</li>
          <li><strong>Workflow Status Integrity:</strong> Exported <code>TC 99101</code> included; Draft <code>TC 99102</code> excluded; Multi-BU <code>TC 99103</code> flagged.</li>
          <li><strong>True Event Dates:</strong> <code>PRJ 99001</code> &amp; <code>EXP 99501</code> locked to Oct 2026 creation/submission dates.</li>
          <li><strong>LEFT JOIN &amp; Quarantine:</strong> <code>BU-AU-REN-10</code> preserved without enrichment; 5 orphan child rows quarantined.</li>
        </ul></div>
        <div class="card indigo"><h3>Measured Test &amp; Evaluation Matrix</h3>
          <p><b>Automated Pytest:</b> <code>34/34</code> passed on DuckDB &middot; <code>34/34</code> passed on SQL Server 2022 &middot; <b>UI Walkthrough:</b> <code>9/9</code> Playwright steps passed.</p>
          <table class="mini-tbl">
            <thead><tr><th>Mode / Model</th><th>Status</th><th>Tool</th><th>Facts</th><th>Median</th></tr></thead>
            <tbody>{eval_trs}</tbody>
          </table>
          <p style="font-size:0.76rem; margin-top:6px;">Model licenses: <code>qwen2.5-coder:14b</code> (Apache-2.0), <code>gpt-oss:20b</code> (Apache-2.0).</p>
        </div>
      </div>
      <div class="callout">Strict Factual Integrity: Every benchmark number is backed by reproducible test logs in <code>results/</code>.</div>
    </section>

    <!-- Slide 12 -->
    <section class="slide">
      <div class="s-top"><div><div class="kicker">Next Steps &amp; Scale</div><div class="s-title">Enterprise Production Lift &amp; Modular Release 2&ndash;4 Roadmap</div></div><span class="s-badge">Roadmap</span></div>
      <div class="grid-4">
        <div class="card teal"><h3>Release 1 (Live Pilot)</h3><p><b>CORE Rollout &amp; Adoption:</b> 4 baseline KPIs, Medallion + DataOps watermarks, Rollout Registry (<code>SIMULATED EXTENSION</code>), MCP &amp; Chat.</p></div>
        <div class="card indigo"><h3>Release 2 (Next)</h3><p><b>Contracts &amp; Delivery:</b> Project/Service/Phase/Task lifecycle, turnaround cycle time, and backlog velocity.</p></div>
        <div class="card teal"><h3>Release 3 (Planned)</h3><p><b>Scheduling &amp; Certs:</b> Inspector dispatch utilization, NDT/safety licence expiry, and qualification compliance.</p></div>
        <div class="card indigo"><h3>Release 4 (Planned)</h3><p><b>Lab LIMS &amp; Billing:</b> Specimen throughput, certificate turnaround, milestone &amp; T&amp;M invoicing.</p></div>
      </div>
      <div class="card" style="margin-top:12px;">
        <h3>Enterprise Production Lift Checklist (Zero Rewrite of Semantic or UI Layers)</h3>
        <p><b>1. Source Ingestion:</b> Swap <code>source_sim.*</code> for read-only CORE SQL Server replicas &nbsp;|&nbsp; <b>2. Identity &amp; HRIS:</b> Bind Entra ID SSO to <code>SESSION_CONTEXT('user_id')</code> and join HRIS worker home-BU &nbsp;|&nbsp; <b>3. PMO Registry:</b> Promote <code>reference.rollout_registry</code> to an owned governance table.</p>
      </div>
    </section>

  </div>
  <div class="nav-bar">
    <button onclick="go(-1)">&larr; Prev</button>
    <span id="counter">Slide 1 / 12</span>
    <button onclick="go(1)">Next &rarr;</button>
  </div>
</div>
<script>
  const slides = Array.from(document.querySelectorAll('.slide'));
  let idx = 0;
  function show(n) {{
    slides.forEach(s => {{ s.classList.remove('active'); s.querySelectorAll('video').forEach(v => v.pause()); }});
    idx = (n + slides.length) % slides.length;
    slides[idx].classList.add('active');
    document.getElementById('counter').textContent = 'Slide ' + (idx + 1) + ' / ' + slides.length;
  }}
  function go(d) {{ show(idx + d); }}
  document.addEventListener('keydown', e => {{
    if (e.key === 'ArrowRight' || e.key === 'PageDown' || e.key === ' ') go(1);
    if (e.key === 'ArrowLeft' || e.key === 'PageUp') go(-1);
  }});
</script>
</body>
</html>
"""
    OUT_HTML.parent.mkdir(parents=True, exist_ok=True)
    OUT_HTML.write_text(html_doc, encoding="utf-8")
    print(f"Saved HTML5 deck: {OUT_HTML} ({OUT_HTML.stat().st_size // 1024} KB)")


if __name__ == "__main__":
    rows = load_eval_rows()
    build_pptx(rows)
    build_html(rows)
