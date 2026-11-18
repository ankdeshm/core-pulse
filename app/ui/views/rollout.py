"""Rollout Explorer view: Wave/Entity/Module deployment status vs observed instrumented activity."""
from __future__ import annotations

import streamlit as st

from app.semantic import service as Svc
from app.ui import common as C
from app.viz.charts import figure
from app.viz.tables import TABLE_CSS, render_html, to_csv_bytes


def render() -> None:
    cfg = st.session_state["cfg"]
    st.markdown("### 🧭 Rollout & Module Registry Explorer (`SIMULATED EXTENSION`)")
    st.info(
        "**Governance distinction:** Module deployment states (`LIVE`, `ONBOARDING`, `PLANNED`, `PAUSED`, `RETIRED`) "
        "come from `reference.rollout_registry` (`SIMULATED EXTENSION`). Application logins (`CORE_LOGIN`) are **never** "
        "used to claim usage of uninstrumented modules (`INSPECTOR_SCHED`, `QUALIFICATIONS`, `LAB_TESTING`, `INVOICING`)."
    )

    f1, f2, f3 = st.columns(3)
    st_pick = f1.selectbox("Filter by Rollout Status", ["All statuses", "LIVE", "ONBOARDING", "PLANNED", "PAUSED", "RETIRED"])
    mod_pick = f2.selectbox(
        "Filter by Module",
        ["All modules", "CORE_LOGIN", "PROJECTS", "TIMECARDS", "EXPENSES",
         "INSPECTOR_SCHED", "QUALIFICATIONS", "LAB_TESTING", "INVOICING"],
    )
    zero_only = f3.checkbox("Show only LIVE rollouts with ZERO observed events", value=False)

    res = Svc.get_core_rollout_status(
        module=None if mod_pick == "All modules" else mod_pick,
        rollout_status=None if st_pick == "All statuses" else st_pick,
        zero_activity_only=zero_only,
        user_id=cfg["user_id"],
        as_of=cfg["as_of"],
    )

    if res.status == "refusal":
        st.error(res.headline)
        return

    st.markdown(res.headline)
    c_cols = st.columns(4)
    for i, fact in enumerate(res.facts[:4]):
        parts = fact.split(":", 1)
        c_cols[i].metric(parts[0].strip(), parts[1].strip() if len(parts) > 1 else "")

    fig = figure(res.chart_spec)
    if fig is not None:
        st.plotly_chart(fig, width="stretch", key="rollout_chart")

    if not res.df.empty:
        st.markdown(TABLE_CSS + render_html(res.df), unsafe_allow_html=True)
        st.download_button(
            "⬇️ Download Rollout Matrix (CSV)",
            to_csv_bytes(res.df),
            file_name="core_pulse_rollout_registry.csv",
            mime="text/csv",
        )

    st.markdown("#### Investigate in Chat")
    b1, b2, b3 = st.columns(3)
    with b1:
        C.ask_button(
            "Ask: Live Expenses with Zero Activity",
            "Which entities are live on Expenses, and which live entities have zero expense submissions?",
            key="ro_ask_1",
        )
    with b2:
        C.ask_button(
            "Ask: Western Region Live Modules",
            "Which modules are live for the Western region?",
            key="ro_ask_2",
        )
    with b3:
        C.ask_button(
            "Ask: Are Live Modules Seeing Activity?",
            "Are those modules actually seeing activity?",
            key="ro_ask_3",
        )
