"""Interactive Plotly chart builders for CORE Pulse structured tool results."""
from __future__ import annotations

import plotly.graph_objects as go

TEAL = "#0F766E"
INDIGO = "#3730A3"
AMBER = "#D97706"
NAVY = "#0B1F3A"
PALETTE = [TEAL, INDIGO, "#0284C7", AMBER, "#64748B"]


def figure(spec: dict | None) -> go.Figure | None:
    if not spec or not spec.get("kind"):
        return None
    kind = spec["kind"]
    title = spec.get("title", "")

    if kind == "line":
        xs = spec.get("x", [])
        ys = spec.get("y", [])
        partials = spec.get("partial_flags", [False] * len(xs))
        fig = go.Figure()
        # Solid trace for complete months
        comp_idx = [i for i, p in enumerate(partials) if not p]
        if comp_idx:
            fig.add_trace(go.Scatter(
                x=[xs[i] for i in comp_idx],
                y=[ys[i] for i in comp_idx],
                mode="lines+markers",
                name="Complete month",
                line={"color": TEAL, "width": 3},
                marker={"size": 8, "color": TEAL},
            ))
        # Dashed connector + amber diamond for partial MTD month!
        part_idx = [i for i, p in enumerate(partials) if p]
        if part_idx:
            p_i = part_idx[0]
            seg_x = ([xs[p_i - 1]] if p_i > 0 else []) + [xs[p_i]]
            seg_y = ([ys[p_i - 1]] if p_i > 0 else []) + [ys[p_i]]
            fig.add_trace(go.Scatter(
                x=seg_x,
                y=seg_y,
                mode="lines+markers",
                name="Partial MTD (in progress)",
                line={"color": AMBER, "width": 2.5, "dash": "dot"},
                marker={"size": [0, 11] if len(seg_x) == 2 else [11], "symbol": "diamond", "color": AMBER},
            ))
        fig.update_layout(
            title={"text": title, "font": {"size": 15, "color": NAVY}},
            yaxis_title=spec.get("y_label", "Count"),
            xaxis_title="Calendar Month",
            template="plotly_white",
            height=370,
            margin={"l": 20, "r": 20, "t": 48, "b": 24},
            legend={"orientation": "h", "y": -0.2},
        )
        return fig

    if kind == "hbar":
        xs = list(reversed(spec.get("x", [])))
        ys = list(reversed(spec.get("y", [])))
        fig = go.Figure(go.Bar(
            x=xs, y=ys, orientation="h",
            marker_color=INDIGO,
            text=[f"{int(v):,}" for v in xs],
            textposition="auto",
        ))
        fig.update_layout(
            title={"text": title, "font": {"size": 15, "color": NAVY}},
            xaxis_title=spec.get("x_label", "Count"),
            template="plotly_white",
            height=max(280, 75 + 42 * len(ys)),
            margin={"l": 20, "r": 25, "t": 48, "b": 24},
        )
        return fig

    if kind == "grouped_bar":
        xs = spec.get("x", [])
        fig = go.Figure()
        for idx, s in enumerate(spec.get("series", [])):
            vals = s.get("values", [])
            fig.add_trace(go.Bar(
                x=xs, y=vals, name=s.get("name", f"Series {idx+1}"),
                marker_color=PALETTE[idx % len(PALETTE)],
                text=[f"{int(v):,}" for v in vals],
                textposition="outside",
            ))
        fig.update_layout(
            barmode="group",
            title={"text": title, "font": {"size": 15, "color": NAVY}},
            yaxis_title=spec.get("y_label", "Count"),
            template="plotly_white",
            height=380,
            margin={"l": 20, "r": 20, "t": 48, "b": 24},
            legend={"orientation": "h", "y": -0.2},
        )
        return fig

    if kind == "bar":
        xs = spec.get("x", [])
        ys = spec.get("y", [])
        fig = go.Figure(go.Bar(
            x=xs, y=ys,
            marker_color=TEAL,
            text=[f"{int(v):,}" for v in ys],
            textposition="auto",
        ))
        fig.update_layout(
            title={"text": title, "font": {"size": 15, "color": NAVY}},
            yaxis_title=spec.get("y_label", "Count"),
            template="plotly_white",
            height=350,
            margin={"l": 20, "r": 20, "t": 48, "b": 24},
        )
        return fig

    return None
