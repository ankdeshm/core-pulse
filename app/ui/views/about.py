"""About view: Architecture, Semantic Governance Contract, Production Lift, and Modular Release Roadmap."""
from __future__ import annotations

import streamlit as st

from app.config import backend
from app.ui import branding as B


def render() -> None:
    st.markdown(f"### ℹ️ About {B.APP_NAME} ({B.APP_FULL_NAME})")
    st.markdown(
        f"**{B.APP_NAME}** (`v{B.VERSION}`) is a governed operational rollout and adoption analytics pilot built for "
        f"**{B.COMPANY}**'s global **{B.OPERATIONAL_PLATFORM}** platform. It demonstrates end-to-end **Medallion Data Engineering** "
        f"(`source_sim` → `bronze` → `silver` → `gold` → `goldview.v_core_adoption`) paired with a **Read-Only MCP Semantic Layer** "
        f"and a **Tool-Grounded Conversational BI Experience**."
    )

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(
            "#### What Is Authoritative Baseline (`prototype.md`)\n"
            "- **Four Certified KPIs:** `Contracts Created`, `Employees Logging In`, `Timesheets Submitted`, and `Expense Reports`.\n"
            "- **True Event Timestamps:** `Project.CreatedDate`, `ApplicationLog.CreatedDate`, `Timecard.SubmittedDate`, and `ExpenseReport.SubmittedDate`.\n"
            "- **Non-Additive Distinct Users:** `COUNT(DISTINCT APPUserId)` (`IsServiceAccount = 0`) recomputed at each grain (`OVERALL`, `WAVE`, `ENTITY`), never summed from child rows.\n"
            "- **Explicit Attribution Limits:** Business Unit breakdowns are blocked for `Employees Logging In` (no employee home-BU table) and `Timesheets Submitted` (multi-BU task rows lack allocation rules).\n"
            "- **DataOps Governance:** `ApplicationLogId` incremental watermark, explicit `INSERT` column lists, orphan row quarantine, and parity checks."
        )
    with c2:
        st.markdown(
            "#### What Is Simulated vs. Enterprise Production Lift\n"
            "- **Synthetic World:** 100% deterministic seeded data (`seed=42`, `AS_OF_DATE=2026-11-17`, 18 months, 3 regions, 10 BUs) with zero corporate connectivity.\n"
            "- **Rollout Registry (`SIMULATED EXTENSION`):** `reference.rollout_registry` and `reference.module_catalog` model `LIVE` / `ONBOARDING` / `PLANNED` states and distinguish `INSTRUMENTED` vs `NOT_INSTRUMENTED` modules.\n"
            "- **Dual Database Runtime:** Runs identically on local **DuckDB** (`data/core_pulse.duckdb`) and **SQL Server 2022** (`core_pulse`, read-only login `pulse_reader`, native `SESSION_CONTEXT` Row-Level Security). Active backend: **`"
            + backend().upper() + "`**.\n"
            "- **Enterprise Lift Checklist:** Swap `source_sim.*` ingestion for read-only replicas of CORE SQL Server, map Entra ID SSO claims to `rls.user_region_access`, and connect HRIS worker master for home-BU login attribution."
        )

    st.markdown("#### Modular Platform Roadmap (Design Only — Preserves Baseline KPIs Unchanged)")
    st.markdown(
        "| Release | Analytics Domain | Planned Facts & Capabilities |\n"
        "|---|---|---|\n"
        "| **Release 1 (Live Pilot)** | **CORE Rollout & Adoption** | 4 baseline adoption KPIs, Medallion + DataOps watermarks, Rollout Registry (`SIMULATED EXTENSION`), MCP tools & conversational BI |\n"
        "| **Release 2** | **Contracts & Project Delivery** | Project/Service/Phase/Task lifecycle, turnaround cycle time, backlog velocity |\n"
        "| **Release 3** | **Field Scheduling & Qualifications** | Inspector dispatch utilization, NDT/safety certification expiry & compliance |\n"
        "| **Release 4** | **Laboratory LIMS & Billing** | Specimen throughput, certificate turnaround, milestone & T&M invoicing |"
    )
    st.caption(B.footer_line())
