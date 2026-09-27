"""Document I/O shared across the app's Streamlit pages.

Split out of collation_app_01.py so a second page can reuse it — e.g. the
Wylie -> Uchen converter, see SPEC-wylie-to-unicode.md — without importing
that file itself, which would re-run its top-level Streamlit UI.
"""

import re
import subprocess
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path


def deployed_version() -> str:
    """The last commit's date, as a display string.

    Streamlit Cloud rebuilds lag behind a push by a few minutes and give no
    visible sign of it — the page a user sees can be running older code with
    nothing to tell them so. Reading it straight from git, rather than
    hand-maintaining a version string that's one edit away from lying, means
    a stale deploy shows up as a stale date instead of as a silent mismatch.
    """
    try:
        out = subprocess.run(
            ["git", "log", "-1", "--format=%cd", "--date=format:%Y-%m-%d %H:%M"],
            cwd=Path(__file__).parent,
            capture_output=True, text=True, timeout=5, check=True,
        )
        return out.stdout.strip() or "unknown"
    except Exception:
        return "unknown"


# ── Google Docs ──────────────────────────────────────────────────────
# A Google Doc can be read without any API key through its plain-text export
# endpoint, but only when the document is shared "anyone with the link can
# view": Streamlit Cloud fetches from its own servers, so a restricted doc
# returns a sign-in page rather than the text. A document may hold several
# tabs, and ?tab= selects one — without it every tab is concatenated, each
# preceded by its title.
_GDOC_ID_RE = re.compile(r"/document/d/([A-Za-z0-9_-]+)")
_GDOC_TAB_RE = re.compile(r"[?&]tab=([A-Za-z0-9_.\-]+)")

# Google's text export puts footnote bodies at the end, after a rule of
# underscores. Left in, they collate as a trailing addition — one editorial
# note like "?=spungs" is enough to invent a variant. A run of underscores
# does not occur in Tibetan, in script or in Wylie, so the rule is safe.
_GDOC_FOOTNOTES_RE = re.compile(r"\n_{8,}\s*\n")


def strip_gdoc_footnotes(text: str) -> str:
    """Drop the footnote block Google appends after a rule of underscores."""
    m = _GDOC_FOOTNOTES_RE.search(text)
    return text[: m.start()] if m else text


# What each export format must come back as. A private document answers with
# a sign-in page carrying HTTP 200, so the content type is the only thing that
# tells a refusal from a document — and handing that HTML to the .docx parser
# would report a corrupt file instead of a sharing problem.
_GDOC_FORMATS = {
    "txt": ("text/plain", ".txt"),
    "docx": ("application/vnd.openxmlformats-officedocument", ".docx"),
}


def fetch_google_doc(url: str, timeout: int = 30, fmt: str = "txt"):
    """Fetch a Google Doc (or one of its tabs) as text or as a .docx.

    Returns ``(raw_bytes, display_name)``. Raises ValueError with a message
    meant for the user — a wrong link and a private document are the two
    things that actually go wrong, and they need different fixes.

    ``fmt`` is "txt" for a witness, whose text is all that is wanted, or
    "docx" for a collation report, which is a table: exported as text its
    columns collapse into one stream and there is no telling the witnesses
    apart again.
    """
    want_type, suffix = _GDOC_FORMATS[fmt]
    url = (url or "").strip()
    if not url:
        raise ValueError("no link given.")
    m = _GDOC_ID_RE.search(url)
    if not m:
        raise ValueError(
            "that does not look like a Google Doc link — it should contain "
            "`/document/d/…`. Copy the URL from the browser's address bar."
        )
    doc_id = m.group(1)
    tab = _GDOC_TAB_RE.search(url)
    export = f"https://docs.google.com/document/d/{doc_id}/export?format={fmt}"
    if tab:
        export += f"&tab={tab.group(1)}"

    req = urllib.request.Request(export, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            ctype = resp.headers.get("Content-Type", "")
            disposition = resp.headers.get("Content-Disposition", "")
            raw = resp.read()
    except urllib.error.HTTPError as exc:
        if exc.code in (401, 403):
            raise ValueError(
                "Google refused the request. Set the document's sharing to "
                "“Anyone with the link → Viewer”; the app fetches it from a "
                "server, so it cannot use your Google account."
            )
        if exc.code == 404:
            raise ValueError("no document found at that link.")
        raise ValueError(f"Google returned HTTP {exc.code}.")
    except Exception as exc:  # network trouble, DNS, timeout
        raise ValueError(f"could not reach Google Docs ({exc}).")

    if want_type not in ctype:
        raise ValueError(
            "Google returned a sign-in page instead of the document. Set the "
            "document's sharing to “Anyone with the link → Viewer”."
        )

    name = doc_id
    got = re.search(r"filename\*=UTF-8''([^;]+)", disposition)
    if got:
        name = urllib.parse.unquote(got.group(1))
    else:
        got = re.search(r'filename="([^"]+)"', disposition)
        if got:
            name = got.group(1)
    if name.lower().endswith(suffix):
        name = name[: -len(suffix)]
    name = re.sub(r"\s+", " ", name).strip()
    if tab:
        # every tab of one document reports the same title, so the tab id is
        # what tells two witnesses apart in the report header
        name = f"{name} · {tab.group(1)}"
    return raw, name
