"""The masthead and partner links, shared by every page of this app.

Both logos are white, so the band carries its own dark ground rather than
relying on the page's. That is what lets the branding look the same whether a
reader has Streamlit in light or dark mode — the alternative was pinning the
theme and taking that choice away from them.

The files live in ``assets/`` rather than being fetched from the partners'
sites: a page that hot-links someone else's logo breaks when they reorganise,
and asks their server for an image on every single load.
"""
import base64
import functools
from pathlib import Path

import streamlit as st

ITAS_URL = "https://itas-uni.eu/"
KARMAPAS_URL = "https://www.translating-karmapas.org/"

_ASSETS = Path(__file__).parent / "assets"
_BAND = "#16181A"          # the band's own ground, independent of the theme
_RULE = "rgba(255,255,255,.28)"   # echoes the hairline inside the ITAS logo


@functools.lru_cache(maxsize=8)
def _data_uri(name: str) -> str:
    """A logo as a data URI, read once and kept.

    Inlined rather than served, because Streamlit has no static route we can
    rely on and an <img src> pointing at the filesystem will not resolve.
    """
    return "data:image/png;base64," + base64.b64encode(
        (_ASSETS / name).read_bytes()
    ).decode("ascii")


def masthead() -> None:
    """The two partner logos on one dark band, above everything else.

    The rule between them is the same hairline the ITAS logo uses internally
    between its mark and its descriptor, so the pair reads as one lockup
    rather than as two logos parked side by side.
    """
    try:
        itas = _data_uri("itas-logo.png")
        karmapas = _data_uri("karmapas-logo.png")
    except OSError:
        return          # branding is decoration; never let it stop the tool

    st.markdown(
        f"""
        <div style="background:{_BAND};margin:-1rem -1rem 1.6rem;padding:15px 26px;
                    display:flex;align-items:center;gap:22px;flex-wrap:wrap;">
          <a href="{ITAS_URL}" target="_blank" rel="noopener"
             style="display:flex;align-items:center;">
            <img src="{itas}" alt="ITAS — International Institute for Tibetan
                 and Asian Studies" style="height:34px;display:block;">
          </a>
          <div style="width:1px;align-self:stretch;background:{_RULE};"></div>
          <a href="{KARMAPAS_URL}" target="_blank" rel="noopener"
             style="display:flex;align-items:center;">
            <img src="{karmapas}" alt="Translating the Karmapas' Works"
                 style="height:30px;display:block;">
          </a>
        </div>
        """,
        unsafe_allow_html=True,
    )


def partner_links() -> str:
    """One line naming both partners, for the foot of the page."""
    return (f"[International Institute for Tibetan and Asian Studies]({ITAS_URL})"
            f" · [Translating the Karmapas' Works]({KARMAPAS_URL})")
