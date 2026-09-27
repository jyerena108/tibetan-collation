"""Wylie -> Uchen converter — apparatus intact.

See SPEC-wylie-to-unicode.md. MVP scope: a plain document (paragraphs and
real .docx footnotes) — the table case comes later.
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

import streamlit as st

from docio import fetch_google_doc
import wylie_uchen as wu


st.set_page_config(
    page_title="Wylie → Uchen",
    page_icon="ༀ",
    layout="centered",
)

st.title("ༀ Wylie → Uchen")
st.caption(
    "Converts Wylie to Uchen in a Word file or Google Doc, footnotes "
    "included, without touching the apparatus — sigla, page markers, and "
    "editorial terms like `om.`/`em.` stay exactly as written."
)

st.divider()

if "wu_raw" not in st.session_state:
    st.session_state.wu_raw = None
    st.session_state.wu_name = ""
    st.session_state.wu_notes = None
    st.session_state.wu_sigla = ""
    st.session_state.wu_terms = ""
    st.session_state.wu_flagged = []

# ── Step 1: get the file ────────────────────────────────────────────
st.subheader("1 · The document")

source_mode = st.radio(
    "Source",
    label_visibility="collapsed",
    options=["Upload a .docx file", "Google Doc link"],
    index=0,
    horizontal=True,
    help="A Google Doc must be shared “Anyone with the link → "
    "Viewer”: the app fetches it from a server and cannot use your "
    "Google account.",
)

raw, name = None, ""
if source_mode.startswith("Upload"):
    up = st.file_uploader("Word file (.docx)", type=["docx"])
    if up is not None:
        raw, name = up.read(), up.name.rsplit(".", 1)[0]
else:
    url = st.text_input(
        "Google Doc URL",
        placeholder="https://docs.google.com/document/d/…",
    )
    if url.strip():
        try:
            raw, name = fetch_google_doc(url, fmt="docx")
        except ValueError as exc:
            st.error(str(exc))

if raw is not None and st.button("Analyze", type="primary"):
    try:
        notes = wu.read_footnote_texts(raw)
    except Exception as exc:
        st.error(f"couldn't read that as a .docx with footnotes ({exc}).")
    else:
        sigla, terms, flagged = wu.detect_exceptions(notes)
        st.session_state.wu_raw = raw
        st.session_state.wu_name = name
        st.session_state.wu_notes = notes
        st.session_state.wu_sigla = ", ".join(sigla)
        st.session_state.wu_terms = ", ".join(terms)
        st.session_state.wu_flagged = flagged

# ── Step 2: confirm sigla and technical terms ───────────────────────
if st.session_state.wu_raw is not None:
    st.divider()
    st.subheader("2 · Sigla and technical terms")
    st.caption(
        f"Found {len(st.session_state.wu_notes)} footnotes. These are read "
        "off the note's own shape — a siglum is whatever labels a reading, "
        "a technical term is the one place a note carries a bare period "
        "(`om.`, `em.`, …). Add, remove, or just confirm."
    )

    sigla_text = st.text_input(
        "Sigla (comma-separated)",
        value=st.session_state.wu_sigla,
        help="Left exactly as written wherever they appear in a footnote.",
    )
    terms_text = st.text_input(
        "Technical terms (comma-separated)",
        value=st.session_state.wu_terms,
        help="Matched exactly, period included — `om.` and `om` are "
        "different tokens.",
    )

    flagged = st.session_state.wu_flagged
    if flagged:
        with st.expander(
            f"⚠️ {len(flagged)} footnote(s) left in Wylie — "
            "couldn't confirm the shape, worth a look",
            expanded=True,
        ):
            st.caption(
                "These don't reduce to `sigla reading] sigla reading` — "
                "usually because the editor's own note rides along in the "
                "footnote text. Left untouched rather than guessed at."
            )
            for fid in flagged:
                st.text(f"[{fid}] {st.session_state.wu_notes[fid].strip()}")

    st.divider()
    st.subheader("3 · Convert")

    if st.button("Convert to Uchen", type="primary"):
        sigla = [s.strip() for s in sigla_text.split(",") if s.strip()]
        terms = [t.strip() for t in terms_text.split(",") if t.strip()]
        try:
            out = wu.convert_docx_bytes(
                st.session_state.wu_raw, sigla, terms,
                flagged_ids=st.session_state.wu_flagged,
            )
        except Exception as exc:
            st.error(f"conversion failed: {exc}")
        else:
            st.success("Converted.")
            st.download_button(
                "Download converted .docx",
                data=out,
                file_name=f"{st.session_state.wu_name} (Uchen).docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
