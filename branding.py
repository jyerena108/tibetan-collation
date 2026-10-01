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

    Sized against the page title rather than in the abstract. The title sets
    at 44px, so a logo below that reads as a footnote to the page instead of
    as the masthead of it — which is what 34px, and even 48px, looked like.
    68px gives the ITAS mark clear precedence over the title and makes its
    descriptor line, "International Institute for Tibetan and Asian Studies",
    plainly readable.

    Neither file is being enlarged past what it holds: 68px is 83% of the
    ITAS artwork's own height, 60px is 67% of the Karmapa's. The Karmapa mark
    is set slightly shorter because it is a two-line wordmark against a
    single-line lockup; equal heights make it look the larger of the two.

    The heights are written inline as well as in the stylesheet. A stylesheet
    that gets stripped or fails to apply would silently leave both logos at
    their natural size with nothing to show anything was wrong, and the
    inline value is the one that renders in that case. The narrow-screen rule
    therefore needs !important to beat it.
    """
    try:
        itas = _data_uri("itas-logo.png")
        karmapas = _data_uri("karmapas-logo.png")
    except OSError:
        return          # branding is decoration; never let it stop the tool

    st.markdown(
        f"""
        <style>
          .tct-masthead {{
            background:{_BAND}; margin:-1rem -1rem 1.6rem; padding:20px 28px;
            display:flex; align-items:center; gap:28px; flex-wrap:wrap;
          }}
          .tct-masthead a {{ display:flex; align-items:center; }}
          .tct-masthead img {{ display:block; }}
          .tct-masthead .tct-itas {{ height:68px; }}
          .tct-masthead .tct-km   {{ height:60px; }}
          .tct-masthead .tct-rule {{
            width:1px; align-self:stretch; background:{_RULE};
          }}
          /* Narrow enough that the two marks stack: the rule would be left
             standing beside the first one with the second orphaned beneath
             it, so it goes, and the marks close up. */
          @media (max-width: 640px) {{
            .tct-masthead {{ gap:16px; padding:16px 20px; }}
            .tct-masthead .tct-rule {{ display:none; }}
            .tct-masthead .tct-itas {{ height:46px !important; }}
            .tct-masthead .tct-km   {{ height:40px !important; }}
          }}
        </style>
        <div class="tct-masthead">
          <a href="{ITAS_URL}" target="_blank" rel="noopener">
            <img class="tct-itas" src="{itas}" style="height:68px;display:block;"
                 alt="ITAS — International Institute for Tibetan and Asian Studies">
          </a>
          <div class="tct-rule"></div>
          <a href="{KARMAPAS_URL}" target="_blank" rel="noopener">
            <img class="tct-km" src="{karmapas}" style="height:60px;display:block;"
                 alt="Translating the Karmapas' Works">
          </a>
        </div>
        """,
        unsafe_allow_html=True,
    )


def partner_links() -> str:
    """One line naming both partners, for the foot of the page."""
    return (f"[International Institute for Tibetan and Asian Studies]({ITAS_URL})"
            f" · [Translating the Karmapas' Works]({KARMAPAS_URL})")
