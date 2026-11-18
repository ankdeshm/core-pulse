"""User Guide view: renders docs/user_guide.md with inline live screenshots."""
from __future__ import annotations

import re

import streamlit as st

from app.config import ROOT

IMG_RE = re.compile(r"!\[([^\]]*)\]\(([^)]+)\)")


def render() -> None:
    guide_path = ROOT / "docs" / "user_guide.md"
    if not guide_path.exists():
        st.info("User Guide markdown (`docs/user_guide.md`) not yet generated.")
        return

    content = guide_path.read_text(encoding="utf-8")
    pos = 0
    for m in IMG_RE.finditer(content):
        before = content[pos:m.start()].strip()
        if before:
            st.markdown(before)
        caption, rel_path = m.group(1), m.group(2)
        img_file = (guide_path.parent / rel_path).resolve()
        if img_file.exists():
            st.image(str(img_file), caption=caption, width="stretch")
        else:
            st.caption(f"*(Screenshot `{rel_path}` — {caption})*")
        pos = m.end()
    tail = content[pos:].strip()
    if tail:
        st.markdown(tail)
