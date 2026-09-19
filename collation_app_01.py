"""
Tibetan Collation Web App — Streamlit
======================================
Run with:
    pip install streamlit bayoo-docx
    streamlit run collation_app_01.py

This has to stay the very first statement in the file. Anywhere below the
sys.path line it stops being a docstring and becomes a bare expression, which
Streamlit's magic renders onto the page — the app used to open with a second
<h1> reading "Tibetan Collation Web App — Streamlit" above its real title.
"""

import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "Pydurma", "src"))

import io
import re
import tempfile
import urllib.error
import urllib.parse
import urllib.request
from collections import namedtuple
from pathlib import Path

import streamlit as st
from docx import Document
from docx.shared import Inches, Mm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from Pydurma.gen.normalizer_gen import GenericNormalizer
from Pydurma.gen.tokenizer_gen import GenericTokenizer
from Pydurma.encoder import Encoder
from Pydurma.aligners.fdmp import FDMPaligner
from Pydurma.utils.utils import column_matrix_to_row_matrix, token_row_to_text_row


# ─────────────────────────────────────────────
#  CONSTANTS
# ─────────────────────────────────────────────

COLOR_LIST = [
    "FFD37F", "9AD1FF", "A5FFB1", "FF9A9A", "C79CFF",
    "FFB27F", "7AFFD5", "FF7FBA", "B5FF7A", "7FB1FF",
    "FF7F7F", "FFE07F",
]

# Shad (Tibetan sentence/clause punctuation) — kept separate so it can be
# toggled on/off during collation. U+0F0D..U+0F11 and U+0F14. The ASCII "/"
# is the Wylie/EWTS transliteration of the shad, so transliterated texts get
# the same shad-ignore treatment (a single "/" and a double "//" both reduce
# to shad characters that are stripped when shad differences are ignored).
SHAD_CHARS_CORE = set(["།", "༎", "༏", "༐", "༑", "༔", "/"])

# Characters that are always ignored when detecting content differences
# (tsheg, yig-mgo, brackets, spaces, western punctuation).
PUNCT_TO_IGNORE_CORE = set([
    "་", "༄", "༅", "༈",
    "༼", "༽", "༌", "༗", "༘",
    " ", "\n", "\t", ",", ".", "?", ":", ";",
])

# Characters neutral for matching but still printed in a reading — see
# stack_mark_as_same in apply_preprocessing_options().
COMPARE_ONLY_IGNORE = set()

# Effective sets — (re)built by apply_preprocessing_options() below. The UI
# calls it with the user's preprocessing choices before running a collation;
# the module-level call further down installs the defaults for library use.
SHAD_CHARS = set(SHAD_CHARS_CORE)
PUNCT_TO_IGNORE_BASE = set(PUNCT_TO_IGNORE_CORE)
PUNCT_TO_IGNORE = PUNCT_TO_IGNORE_BASE | SHAD_CHARS


# ── Preprocessing (input normalization) ──────────────────────────────
# Folio/page tags like [354], [zhe 1], [kha zhe lnga] are reference markers,
# not text; left in place they collate as readings and pollute the apparatus.
_FOLIO_TAG_RE = re.compile(r"\[[^\[\]\n]{1,40}\]")


def strip_folio_tags(text: str) -> str:
    """Remove [..] folio/page tags, leaving a space so words don't fuse."""
    return _FOLIO_TAG_RE.sub(" ", text)


def extract_folio_tags(text: str):
    """Strip [..] tags like strip_folio_tags, but remember what and where.

    Returns ``(stripped_text, milestones)`` where milestones is a list of
    ``(offset, tag)`` pairs: ``offset`` is the character position in the
    *stripped* text (pointing at the space that replaced the tag) where the
    tag can be re-inserted later, e.g. into the golden document.
    """
    milestones = []
    out = []
    pos = 0  # length of stripped text built so far
    last = 0
    for m in _FOLIO_TAG_RE.finditer(text):
        out.append(text[last : m.start()])
        pos += m.start() - last
        milestones.append((pos, m.group(0)))
        out.append(" ")
        pos += 1
        last = m.end()
    out.append(text[last:])
    return "".join(out), milestones


def count_preprocessing_hits(text: str) -> dict:
    """Per-rule occurrence counts, for the preprocessing preview."""
    return {
        "folio/page tags [..]": len(_FOLIO_TAG_RE.findall(text)),
        "underscores _": text.count("_"),
        "pipes |": text.count("|"),
        "head marks @ # !": sum(text.count(c) for c in "@#!"),
    }


# ── Reading a collation report back in ─────────────────────────────
# The report carries everything needed to start again: a column per witness
# with its full text, the sigla in the header row, and the base marked
# "(golden, notes)". That makes it an editable carrier for all the witnesses
# at once — correct the OCR in the columns, feed the report back, re-collate.
#
# Page markers and note references are identified by shape, not by the italic
# and superscript formatting the report gives them. Formatting is the first
# thing to smear when a long table is edited by hand in Word, whereas
# "[BX1.2r.1]" (brackets around something containing a dot) and "[7]"
# (brackets around digits alone) survive any amount of retyping. A superscript
# run is only dropped when it really does look like a note reference, so text
# typed next to one cannot disappear silently.
_NOTE_REF_RE = re.compile(r"^\[\d+\]$")
_DOCX_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _docx_runs(node):
    """Yield (text, is_superscript) for every run under ``node``, in order."""
    W = _DOCX_NS
    for run in node.iter(W + "r"):
        parts = []
        for child in run:
            if child.tag == W + "t":
                parts.append(child.text or "")
            elif child.tag == W + "br":
                parts.append("\n")
        text = "".join(parts)
        if not text:
            continue
        rpr = run.find(W + "rPr")
        sup = rpr is not None and rpr.find(W + "vertAlign") is not None
        yield text, sup


def parse_collation_report(raw: bytes):
    """Recover the witnesses from a collation report produced by this tool.

    Returns ``(labels, texts)``, base first. Raises ValueError with a message
    for the user when the file is not a report this tool wrote.
    """
    import xml.etree.ElementTree as ET
    import zipfile

    W = _DOCX_NS
    try:
        with zipfile.ZipFile(io.BytesIO(raw)) as z:
            doc = ET.fromstring(z.read("word/document.xml"))
    except Exception:
        raise ValueError(
            "that is not a readable .docx file. Upload a collation report "
            "this tool produced."
        )

    table = doc.find(f".//{W}tbl")
    if table is None:
        raise ValueError(
            "no table found — this does not look like a collation report. "
            "The report is the side-by-side document, not the footnote one."
        )
    rows = table.findall(f"{W}tr")
    if len(rows) < 2:
        raise ValueError("the report's table has no witness row.")

    header = ["".join(t.text or "" for t in tc.iter(W + "t"))
              for tc in rows[0].findall(f"{W}tc")]
    labels = [h.replace("(golden, notes)", "").strip() for h in header]
    base_at = next((i for i, h in enumerate(header) if "golden" in h), 0)

    texts = []
    for tc in rows[1].findall(f"{W}tc"):
        parts = []
        for text, sup in _docx_runs(tc):
            if sup and _NOTE_REF_RE.match(text.strip()):
                continue  # the report's own [n] reference
            if not sup and _NOTE_REF_RE.match(text.strip()):
                continue  # same, with its formatting lost to editing
            parts.append(text)
        texts.append("".join(parts))

    if len(texts) != len(labels) or len(texts) < 2:
        raise ValueError("the report's header and witness columns do not match.")

    # the base has to come first for everything downstream
    order = [base_at] + [i for i in range(len(texts)) if i != base_at]
    return [labels[i] for i in order], [texts[i] for i in order]


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


# ── Page / folio markers ─────────────────────────────────────────────
# Witnesses carry their pagination in whatever form their source used —
# "p.292" from a PDF, "kha, 4r.7 (pdf 55)" from a pecha (volume, folio,
# recto/verso, line). Rather than force one format, the user pastes the first
# marker from each file and the shape is generalized from it: digit runs
# become \d+ and letter runs [A-Za-z]+, so a sample from volume ka, folio 1
# recto still matches later volumes and verso sides. A trailing parenthetical
# is made optional, so a sample carrying "(pdf 47)" does not exclude markers
# that omit it.


def pattern_from_example(example: str) -> str:
    """Generalize a sample page marker into a regex source string."""
    out, i = [], 0
    while i < len(example):
        c = example[i]
        if c.isdigit():
            j = i
            while j < len(example) and example[j].isdigit():
                j += 1
            out.append(r"\d+")
            i = j
        elif c.isalpha():
            j = i
            while j < len(example) and example[j].isalpha():
                j += 1
            out.append(r"[A-Za-z]+")
            i = j
        elif c.isspace():
            j = i
            while j < len(example) and example[j].isspace():
                j += 1
            out.append(r"\s+")
            i = j
        else:
            out.append(re.escape(c))
            i += 1
    pat = "".join(out)
    m = re.search(r"(?:\\s\+)?\\\(.*\\\)$", pat)
    if m:
        pat = pat[: m.start()] + "(?:" + pat[m.start():] + ")?"
    return pat


# House style is a bracketed tag containing a dot: [BX1.1v.1], [DX1.145r.1],
# or the shorter [AB1.272]. Requiring the dot is what separates a page marker
# from a plain folio tag like [354], which the folio-tag option handles.
DEFAULT_PAGE_MARKER_RE = r"\[[^\[\]\n]*\.[^\[\]\n]*\]"


def extract_page_markers(text: str, example: str):
    """Strip page markers from ``text``, remembering what and where.

    Returns ``(stripped_text, markers)`` with markers as ``(offset, marker)``
    pairs, the offset being the position in the *stripped* text from which that
    marker's page applies. Markers must be removed before collation or they
    align as readings and pollute the apparatus; their positions are what let
    the golden document show where each witness turned its page.

    With no ``example`` the default pattern is used: any bracketed tag
    containing a dot. That covers the agreed house style — ``[BX1.1v.1]``
    (siglum, folio, side, line) and the shorter ``[AB1.272]`` — while leaving
    a plain folio tag such as ``[354]`` or ``[zhe 1]`` to the folio-tag
    preprocessing option, which is a different thing.
    """
    example = (example or "").strip()
    try:
        rx = re.compile(pattern_from_example(example) if example
                        else DEFAULT_PAGE_MARKER_RE)
    except re.error:
        return text, []
    markers, out, pos, last = [], [], 0, 0
    for m in rx.finditer(text):
        out.append(text[last:m.start()])
        pos += m.start() - last
        markers.append((pos, m.group(0)))
        out.append(" ")
        pos += 1
        last = m.end()
    out.append(text[last:])
    return "".join(out), markers


# Tibetan carries no digits of its own, in script or in Wylie, so a token
# containing one is a strong signal of leftover pagination.
_DIGIT_TOKEN_RE = re.compile(r"\S*[0-9\u0f20-\u0f29]\S*")


# ── Upload decoding ──────────────────────────────────────────────────
# Uploads are not reliably UTF-8. OCR tools and PDF text extractors on macOS
# still emit Mac OS Roman, in which 0xCA is a non-breaking space — an invalid
# UTF-8 continuation byte. A bare .decode("utf-8") therefore killed the whole
# run with a raw UnicodeDecodeError, which tells a translator nothing.
#
# Mac Roman maps all 256 byte values and so never raises; it is the terminal
# fallback and anything listed after it would be unreachable. That means a
# Windows-1252 file is read as Mac Roman, which decodes its non-ASCII bytes
# differently — hence the UI reports which encoding was used, so a wrong guess
# shows up as visible mojibake rather than passing silently.
_UPLOAD_ENCODINGS = (
    ("utf-8-sig", "UTF-8"),        # also strips a byte-order mark, if present
    ("mac_roman", "Mac OS Roman"),
)


def decode_upload(raw: bytes):
    """Decode an uploaded file, tolerating non-UTF-8 input.

    Returns ``(text, encoding_label)``. Line endings are normalized to ``\n``:
    classic-Mac files use bare CR, which otherwise arrives as one enormous
    line and embeds stray carriage returns in the Word output. This is the
    only rewriting done here — the text itself is passed through untouched.
    """
    for codec, label in _UPLOAD_ENCODINGS:
        try:
            text = raw.decode(codec)
        except UnicodeDecodeError:
            continue
        return text.replace("\r\n", "\n").replace("\r", "\n"), label
    # Not reachable while Mac Roman is in the table above, but keep the app
    # alive rather than crashing if that table is ever narrowed.
    text = raw.decode("utf-8", errors="replace")
    return text.replace("\r\n", "\n").replace("\r", "\n"), "UTF-8 (damaged bytes replaced)"


def apply_preprocessing_options(
    underscore_as_space=True,
    pipe_as_shad=True,
    ignore_head_marks=True,
    stack_mark_as_same=True,
):
    """Install the effective character sets used by the collation.

    - underscore_as_space: EWTS writes an explicit space as ``_``; treat it
      as whitespace everywhere (ignore set, syllable splitting, a-chung
      reattachment).
    - pipe_as_shad: ``|`` is an alternate EWTS shad (common in OCR output).
    - ignore_head_marks: ``@``, ``#``, ``!`` transliterate yig-mgo ornaments
      (༄༅ …), which are structural, not textual.
    - stack_mark_as_same: EWTS ``+`` forces a stacked consonant, so
      ``seng+ge`` and ``seng ge`` are the same word written two ways — one
      following the Sanskrit original, one the naturalised Tibetan spelling.
      That is a graphic variant, not a textual one, and repeating it through
      an apparatus buries the readings that matter. Ignored here for matching
      only: ``+`` is never removed from a reading, so a genuine variant still
      prints as ``d+hi``. The orthographic profile counts what this hides.

    Rebinding the module-level sets keeps every existing function signature
    unchanged; the display/golden text itself is never rewritten by these.
    """
    global SHAD_CHARS, PUNCT_TO_IGNORE_BASE, PUNCT_TO_IGNORE
    global _SYLLABLE_SEP_RE, _SEP_CHARS, COMPARE_ONLY_IGNORE
    COMPARE_ONLY_IGNORE = {"+"} if stack_mark_as_same else set()
    SHAD_CHARS = set(SHAD_CHARS_CORE) | ({"|"} if pipe_as_shad else set())
    PUNCT_TO_IGNORE_BASE = (
        set(PUNCT_TO_IGNORE_CORE)
        | ({"_"} if underscore_as_space else set())
        | ({"@", "#", "!"} if ignore_head_marks else set())
    )
    PUNCT_TO_IGNORE = PUNCT_TO_IGNORE_BASE | SHAD_CHARS
    _SYLLABLE_SEP_RE = re.compile(r"[་༌\s_]+" if underscore_as_space else r"[་༌\s]+")
    _SEP_CHARS = set(["་", "༌", " ", "\t", "\n"]) | (
        {"_"} if underscore_as_space else set()
    )


# ─────────────────────────────────────────────
#  CORE LOGIC (unchanged from your script)
# ─────────────────────────────────────────────

# A3 landscape, in the units python-docx wants. The report is a table of up
# to six witnesses side by side and is also what gets edited by hand when
# correcting OCR, so width is the thing it needs most. Letter landscape left
# 8.5in of usable width once python-docx's 1.25in side margins were taken out
# — 1.42in a column at six witnesses. A3 with half-inch margins gives 15.54in,
# and is three inches taller as well, so more rows fit too. A3 is a standard
# size everything handles; a custom page wider than this prints unpredictably.
REPORT_PAGE_WIDTH = Mm(420)
REPORT_PAGE_HEIGHT = Mm(297)
REPORT_SIDE_MARGIN = Inches(0.5)
REPORT_TOP_MARGIN = Inches(0.6)


def set_report_page(document):
    """Put the report on a wide A3 landscape page with narrow side margins."""
    for section in document.sections:
        section.orientation = WD_ORIENT.LANDSCAPE
        # python-docx does not swap the dimensions when the orientation is
        # set, so they are given explicitly rather than exchanged.
        section.page_width = REPORT_PAGE_WIDTH
        section.page_height = REPORT_PAGE_HEIGHT
        section.left_margin = REPORT_SIDE_MARGIN
        section.right_margin = REPORT_SIDE_MARGIN
        section.top_margin = REPORT_TOP_MARGIN
        section.bottom_margin = REPORT_TOP_MARGIN


def ensure_footnote_reference_style(document):
    """Define the FootnoteReference character style with superscript.

    bayoo-docx emits each in-text reference mark as
    <w:rStyle w:val="FootnoteReference"/> but never defines that style, so
    Word renders the number at the baseline. We create it here so the mark
    renders raised/superscript like a real footnote number.
    """
    styles_el = document.styles.element
    target = None
    for st_el in styles_el.findall(qn("w:style")):
        if st_el.get(qn("w:styleId")) == "FootnoteReference":
            target = st_el
            break
    if target is None:
        target = OxmlElement("w:style")
        target.set(qn("w:type"), "character")
        target.set(qn("w:styleId"), "FootnoteReference")
        name = OxmlElement("w:name")
        name.set(qn("w:val"), "footnote reference")
        target.append(name)
        styles_el.append(target)
    rPr = target.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        target.append(rPr)
    va = rPr.find(qn("w:vertAlign"))
    if va is None:
        va = OxmlElement("w:vertAlign")
        rPr.append(va)
    va.set(qn("w:val"), "superscript")


def shrink_footnote_style(document, font_size=8, line_spacing_multiple=0.85):
    styles = document.styles
    for style_name in ("Footnote Text", "Footnote Reference"):
        try:
            s = styles[style_name]
            s.font.size = Pt(font_size)
            if style_name == "Footnote Reference":
                # ensure the in-text reference mark renders raised/superscript
                s.font.superscript = True
            if style_name == "Footnote Text":
                pf = s.paragraph_format
                pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
                pf.line_spacing = line_spacing_multiple
                pf.space_before = Pt(0)
                pf.space_after = Pt(0)
        except KeyError:
            pass


# Word processors and OCR routinely convert the Wylie a-chung apostrophe into a
# typographic quote. The characters look nearly identical but are distinct, so
# 'gyur and ’gyur compare as different words and every one produces a false
# variant — 9 of them in one witness of Jataka 35. Normalized for comparison
# only, exactly as "|" and "/" are treated as one shad; the text itself is
# never rewritten, so a witness still prints the character it actually uses.
_APOSTROPHE_VARIANTS = ("\u2018", "\u2019", "\u201b", "\u02bc")
_APOSTROPHE_RE = re.compile("[" + "".join(_APOSTROPHE_VARIANTS) + "]")


def normalize_apostrophes(s: str) -> str:
    """Fold typographic apostrophes onto the ASCII one used by Wylie."""
    return _APOSTROPHE_RE.sub("'", s)


def strip_ignorable(s: str, ignore_shad: bool = True) -> str:
    """Remove characters that don't count as content differences.

    When ignore_shad is True (default), shad punctuation is also stripped, so
    shad-only differences won't generate notes. When False, shad is preserved
    and therefore shad differences will surface as variant notes.

    Any whitespace is dropped, not just the ASCII space — OCR leaves
    non-breaking spaces (U+00A0) and the like, which the syllable splitter
    already treats as separators. Ignoring them here too keeps the comparison
    key consistent with the display, so an invisible space can't masquerade as
    content and produce an empty "om.] … om." note.
    """
    ignore_set = (PUNCT_TO_IGNORE if ignore_shad else PUNCT_TO_IGNORE_BASE)
    ignore_set = ignore_set | COMPARE_ONLY_IGNORE
    s = normalize_apostrophes(s)
    if not ignore_shad and "|" in SHAD_CHARS:
        # "|" and "/" are the same shad in different notation; when shad is
        # kept for comparison they must not read as a difference.
        s = s.replace("|", "/")
    return "".join(ch for ch in s if ch not in ignore_set and not ch.isspace())


_SEP_CHARS = set(["་", "༌", " ", "\t", "\n"])


def _reattach_stranded_achung(*rows):
    """Move a lone initial a-chung onto the syllable that follows it.

    The aligner sometimes strands an initial a-chung (Wylie ``'`` / Unicode
    ``འ``) at the end of one cell while the syllable it belongs to lands in the
    next cell — e.g. witness ``"su '"`` + ``"gyur"`` against base ``"su "`` +
    ``"gyur"``. Left alone this reads as a spurious ``su'`` variant. Moving the
    a-chung forward yields base ``"gyur"`` vs witness ``"'gyur"`` so the note
    correctly reads ``gyur] 'gyur``. Only a *lone* a-chung (preceded by a
    separator, i.e. an initial one) is moved; a final a-chung glued to letters
    such as ``dga'`` is left untouched. Per-row concatenation is preserved.
    """
    for row in rows:
        if row is None:
            continue
        n = len(row)
        for i in range(n):
            seg = row[i]
            if not seg or seg == "-":
                continue
            # Only when the a-chung is the very last character of the cell can
            # it move to the next cell's front without reordering anything in
            # between, so the base/golden text stays byte-for-byte intact.
            if seg[-1] not in A_CHUNG_CHARS:
                continue
            k = len(seg)
            while k > 0 and seg[k - 1] in A_CHUNG_CHARS:
                k -= 1
            # Lone/initial a-chung only: preceded by a separator, a shad, or
            # nothing. After a shad a new word begins, so a lone a-chung there
            # can only be the initial letter of the next word (e.g. "/ /'"
            # before "dir" is the 'a of "'dir").
            if k > 0 and seg[k - 1] not in _SEP_CHARS and seg[k - 1] not in SHAD_CHARS:
                continue
            run = seg[k:]
            j = i + 1
            while j < n and (not row[j] or row[j] == "-"):
                j += 1
            if j >= n:
                continue
            row[i] = seg[:k]
            row[j] = run + row[j]
    return rows


def _drop_empty_columns(rows):
    """Remove columns where every witness is blank.

    The aligner leaves these behind. They contribute nothing to any output,
    but they fall between a syllable and its suffix often enough to hide from
    the merge below that the two belong together.
    """
    if not rows:
        return rows
    width = max(len(r) for r in rows)
    for r in rows:
        r.extend([""] * (width - len(r)))
    for j in range(width - 1, -1, -1):
        if all(r[j] == "" for r in rows):
            for r in rows:
                del r[j]
    return rows


def _merge_final_achung(rows):
    """Rejoin a Tibetan syllable the aligner cut at an apostrophe.

    ``mda'`` and ``bka'`` come back as two cells, the word and then a cell
    holding nothing but the apostrophe, so a note reports ``mnga`` against
    ``mda`` — both readings a letter short of what the witnesses say.

    This is the mirror of _reattach_stranded_achung(), which handles the
    *initial* a-chung: ``su '`` + ``gyur`` becoming ``su `` + ``'gyur``, so an
    added a-chung reads ``gyur] 'gyur`` rather than a bare apostrophe. What
    separates the two is the character the previous cell ends on —

        mnga  +  '        ends in a letter  -> final, belongs backward
        su    +  ' gyur   ends in a space   -> initial, belongs forward

    so only the first is merged here, and the forward rule keeps the second.
    Cells are left alone when any witness has a gap in either of them, so an
    omission is never swallowed.
    """
    if not rows:
        return rows
    width = max(len(r) for r in rows)
    for r in rows:
        r.extend([""] * (width - len(r)))
    j = width - 1
    while j >= 1:
        here = [r[j] for r in rows]
        prev = [r[j - 1] for r in rows]
        gap = any(c == "-" for c in here) or any(p == "-" for p in prev)
        # The tokenizer breaks at every apostrophe, so pa'i arrives as pa +
        # 'i and mnga' as mnga + ' . Each is one Tibetan syllable — a stem
        # with its suffix inside a single tsheg unit — so a cell opening with
        # an apostrophe continues the one before it.
        opens_achung = any(c.strip() for c in here) and all(
            (not c.strip()) or c.lstrip()[:1] in A_CHUNG_CHARS
            for c in here
        )
        # …and the mirror: a cell ending in an apostrophe with the next
        # starting on a letter, which is how pa'ang is cut.
        closes_achung = any(
            p.rstrip()[-1:] in A_CHUNG_CHARS and p == p.rstrip()
            for p in prev if p.strip()
        ) and all(
            (not c.strip()) or c.lstrip()[:1] not in _SEP_CHARS
            for c in here
        )
        lone = opens_achung or closes_achung
        # The base arbitrates, because the marker is placed in the base's
        # text: if the base writes the two as one syllable the marker must not
        # split it, whatever the other witnesses do — GX1 writes "pa 'ang"
        # where the base has "pa'ang". Where the base has a gap there is
        # nothing to judge by, so every witness must agree instead.
        #
        # It matters that this reads the base rather than any witness. An
        # earlier version merged whenever *any* witness ran two cells
        # together, which joined "mi" and "'am" — a word and its particle,
        # not a stem and its suffix — and swallowed a transposition into one
        # unreadable note.
        def runs_on(p):
            return bool(p) and p[-1] not in _SEP_CHARS and p[-1] not in SHAD_CHARS

        if prev and prev[0].strip():
            backward = runs_on(prev[0])
        else:
            joined = [p for p in prev if p.strip()]
            backward = bool(joined) and all(runs_on(p) for p in joined)
        if lone and backward and not gap:
            for r in rows:
                r[j - 1] = r[j - 1] + r[j]
                del r[j]
        j -= 1
    return rows


def _merge_split_stacks(rows):
    """Rejoin a cell the aligner cut in the middle of a stacked word.

    Pydurma's tokenizer treats ``+`` as a token boundary, so ``seng+ge`` comes
    back as two tokens where ``seng ge`` comes back as one pair — the columns
    no longer line up and a note reports the fragment ``seng+`` against
    ``seng``. Merging the cell that ends in ``+`` with the one after it puts
    the word back together, so a reading is never half a word.

    Done whatever the stacking option is set to: this is about not reporting
    fragments, which is wrong either way. Cells are left alone when any
    witness has a gap in either of them, so an omission is never swallowed.
    """
    if not rows:
        return rows
    width = max(len(r) for r in rows)
    for r in rows:
        r.extend([""] * (width - len(r)))
    j = width - 2
    while j >= 0:
        cut = any(r[j].rstrip().endswith("+") for r in rows)
        gap = any(r[j] == "-" or r[j + 1] == "-" for r in rows)
        if cut and not gap:
            for r in rows:
                r[j] = r[j] + r[j + 1]
                del r[j + 1]
        j -= 1
    return rows


def align_witnesses(texts):
    """Align any number of witnesses against the first one (the base).

    Pydurma's FDMPaligner diffs every other witness against the base and sizes
    its matrix to the number of witnesses, so 2..N texts work with no change to
    the vendored library. It pops the base off the lists it is handed, so
    copies go in and the caller's data is left intact.
    """
    normalizer = GenericNormalizer()
    encoder = Encoder()
    tokenizer = GenericTokenizer(encoder, normalizer)
    aligner = FDMPaligner()

    pairs = [tokenizer.tokenize(t) for t in texts]
    token_lists = [p[0] for p in pairs]
    token_strings = [p[1] for p in pairs]

    matrix = aligner.get_alignment_matrix(list(token_strings), list(token_lists))
    row_matrix = column_matrix_to_row_matrix(matrix)

    aligned = [token_row_to_text_row(row_matrix[i], t) for i, t in enumerate(texts)]
    _reattach_stranded_achung(*aligned)
    _drop_empty_columns(aligned)
    _merge_final_achung(aligned)
    _merge_split_stacks(aligned)
    return aligned


def set_run_background_color(run, hex_color: str):
    hex_color = hex_color.lstrip("#").upper()
    rPr = run._element.get_or_add_rPr()
    shd = rPr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        rPr.append(shd)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)


OMITTED_MARK = "om."  # standard critical-apparatus mark for an omitted reading

TSHEG = "་"  # Tibetan intersyllabic tsheg, used to rejoin syllables

# a-chung: U+0F60 (འ) in Unicode Tibetan, apostrophe (') in Wylie/EWTS. When an
# alignment boundary strands a bare a-chung as its own token, it should re-join
# the adjacent syllable so an added/omitted a-chung reads as e.g. "su'" rather
# than surfacing as a meaningless standalone "'".
A_CHUNG_CHARS = set(["འ", "'"]) | set(_APOSTROPHE_VARIANTS)

# Syllable separators for display splitting: tsheg, no-break tsheg, whitespace.
_SYLLABLE_SEP_RE = re.compile(r"[་༌\s]+")


def _syllable_spans(seg: str, ignore_shad: bool):
    """Split a segment into display syllables, keeping where each one sits.

    Returns ``[(text, start, end)…]`` with start/end indexing ``seg`` itself,
    so a caller can both read a syllable and point at it. _syllables() is the
    text-only view of this, which keeps the note and the footnote marker
    working from one splitting — they used to disagree, and the marker ended
    up at the end of the cell rather than on the word the note is about.

    Tsheg and whitespace separate (and are dropped). Shad separates only when
    shad differences are ignored, so it stays visible otherwise. A bare
    a-chung is re-attached to its neighbour rather than kept as its own token.

    Every substitution below is length-preserving, so positions in the working
    copy are positions in ``seg``.
    """
    if not seg:
        return []
    # Fold typographic apostrophes so a note reports the Tibetan reading
    # ('das) rather than the typographic accident (‘das). One character for
    # one character, so offsets are unaffected.
    s = normalize_apostrophes(seg)
    if ignore_shad:
        # A space, not a deletion: the shad is itself a word separator, so
        # "ba//mtshan" has no other break between its two words and deleting
        # it rendered them fused as "bamtshan".
        for ch in SHAD_CHARS:
            s = s.replace(ch, " ")

    spans = []
    for m in re.finditer(r"[^\s་༌_]+" if "_" in _SEP_CHARS else r"[^\s་༌]+", s):
        # Drop ignorable non-content characters (head marks, western
        # punctuation) so readings never show e.g. "@#" from a source file.
        text = "".join(c for c in m.group(0) if c not in PUNCT_TO_IGNORE_BASE)
        if text:
            spans.append([text, m.start(), m.end()])

    merged = []
    pending = ""
    pending_start = None
    for text, start, end in spans:
        if all(c in A_CHUNG_CHARS for c in text):
            if merged:
                merged[-1][0] += text
                merged[-1][2] = end
            else:
                pending += text
                if pending_start is None:
                    pending_start = start
        else:
            merged.append([pending + text,
                           pending_start if pending_start is not None else start,
                           end])
            pending = ""
            pending_start = None
    if pending:
        merged.append([pending, pending_start, pending_start + len(pending)])
    return [tuple(m) for m in merged]


def _syllables(seg: str, ignore_shad: bool):
    """The display syllables of a segment — the text view of _syllable_spans."""
    return [text for text, _start, _end in _syllable_spans(seg, ignore_shad)]


def _trim_common_syllables(readings):
    """Trim syllables shared by *all* readings at the start and the end.

    ``readings`` is a list of syllable-lists. Returns ``(trimmed, pre)`` —
    the syllable-lists with the common leading and trailing syllables removed
    so only the differing part remains, and how many were taken off the front.
    That count is what lets the footnote marker be placed on the lemma rather
    than at the end of the cell. Because at least one reading always differs when a note exists,
    this never trims a reading down to nothing on every side simultaneously.
    """
    readings = [list(r) for r in readings]
    if len(readings) < 2:
        return readings, 0

    # Shared syllables are recognised by comparison key rather than by how
    # they print: "d+hi" and "dhi" are the same reading once the stacking mark
    # is neutral, and comparing the printed form would keep them, leaving a
    # note that cites shared syllables to report a one-word variant.
    def same(a, b):
        return strip_ignorable(a) == strip_ignorable(b)

    # common prefix
    n = min(len(r) for r in readings)
    pre = 0
    while pre < n and all(same(r[pre], readings[0][pre]) for r in readings):
        pre += 1
    readings = [r[pre:] for r in readings]
    # common suffix
    n = min(len(r) for r in readings)
    suf = 0
    while suf < n and all(same(r[-1 - suf], readings[0][-1 - suf]) for r in readings):
        suf += 1
    if suf:
        readings = [r[: len(r) - suf] for r in readings]
    return readings, pre


_TIBETAN_CHAR_RE = re.compile(r"[ༀ-࿿]")


def _reading_display(sylls) -> str:
    """Render a (trimmed) syllable-list for a note; empty = an omission.

    Syllables are rejoined with a tsheg for Tibetan-script input but with a
    plain space for Wylie/roman input — a tsheg between roman letters would
    mix scripts (e.g. "sgrog་la'ang" instead of "sgrog la'ang").
    """
    if not sylls:
        return OMITTED_MARK
    joiner = TSHEG if any(_TIBETAN_CHAR_RE.search(s) for s in sylls) else " "
    return joiner.join(sylls)


def build_note_text(segs, labels, positive=False, ignore_shad=True):
    """Build a single apparatus note in classic critical-edition style.

    ``segs`` and ``labels`` are base-first: index 0 is the base/golden witness
    and the rest are comparison witnesses, however many there are. Returns
    ``(note_text, lemma, lemma_end)``; note_text is "" when no witness
    differs, and lemma_end is the index, in the base's own syllables, just
    past the lemma — what the golden document uses to put the marker on the
    word the note is about rather than at the end of the cell.

    Format: ``<baseSigla> <lemma>] <sigla> <reading>; <sigla> <reading>``

    - The lemma is the base reading (or ``om.`` when the base omits it), and is
      itself labelled with the base siglum so it can be moved into the variant
      list unchanged if the base is later reassigned.
    - Sigla precede the reading they belong to on both sides of the bracket.
      Sigla sharing a reading are comma-separated (``V2, V4``); distinct
      readings are separated by ``; ``. No colon is used.
    - Multi-syllable segments are reduced to just the differing syllable(s):
      syllables shared by every witness are trimmed away, and the survivors
      keep their tsheg separators so they read correctly.
    - Negative apparatus (default): only witnesses that differ from the lemma
      are listed. A positive apparatus additionally credits the witnesses that
      agree by listing their sigla with the lemma (``V1, V3 la] V2 pa``)
      instead of repeating the reading.
    Pagination is not shown here. Each witness's page markers are embedded in
    the golden document's running text instead, at the point where that
    witness turns its page, which keeps the notes readable.
    """

    def siglum(i):
        return labels[i]
    # comparison keys (punctuation/tsheg-insensitive) decide agreement
    keys = [strip_ignorable(s, ignore_shad) for s in segs]
    # display syllables, with syllables shared by every witness trimmed away
    trimmed, pre = _trim_common_syllables(
        [_syllables(s, ignore_shad) for s in segs]
    )
    # Where the lemma ends, counted in the base's own syllables. For an
    # omission the lemma is empty and this marks the gap itself — the point
    # the missing words would occupy — which is where the marker belongs.
    lemma_end = pre + len(trimmed[0])

    lemma = _reading_display(trimmed[0])
    base_key = keys[0]

    lemma_labels = [siglum(0)]
    if positive:
        lemma_labels += [siglum(i) for i in range(1, len(segs)) if keys[i] == base_key]

    selected = [
        (siglum(i), trimmed[i]) for i in range(1, len(segs)) if keys[i] != base_key
    ]
    if not selected:
        return "", lemma, lemma_end

    # group witnesses that share the same displayed reading, preserving order
    groups = []  # list of [reading_display, [labels...]]
    for lab, t in selected:
        disp = _reading_display(t)
        for g in groups:
            if g[0] == disp:
                g[1].append(lab)
                break
        else:
            groups.append([disp, [lab]])

    parts = [f"{', '.join(labs)} {disp}" for disp, labs in groups]
    return (f"{', '.join(lemma_labels)} {lemma}] " + "; ".join(parts),
            lemma, lemma_end)


# One record per aligned cell, shared by both exporters.
def _base_index_for_mark(base_seg: str, witness_seg: str, within: int) -> int:
    """Where in the base's cell does a marker inside a witness's cell belong?

    A page marker usually falls exactly on a cell boundary, but when witnesses
    disagree around it the aligner produces one wide cell containing it — e.g.
    BX1 breaking a word as "me\ntog" against AB1's "me tog" puts AB1's page
    marker in the middle of the shared cell. Anchoring to the cell's start or
    end then lands the marker inside a word.

    The marker is placed after the same number of whitespace-separated words,
    then advanced over any whitespace so it sits on the following word rather
    than in the gap before it. Counting words rather than characters matters
    where the witnesses genuinely differ: DX1 reading "ni" against the base's
    "nyid" would put a character count in the middle of the base's word.
    """
    k = len(witness_seg[:within].split())
    i = seen = 0
    while i < len(base_seg) and seen < k:
        while i < len(base_seg) and base_seg[i].isspace():
            i += 1
        while i < len(base_seg) and not base_seg[i].isspace():
            i += 1
        seen += 1
    while i < len(base_seg) and base_seg[i].isspace():
        i += 1
    return i


# One witness's page marker, located twice: where it belongs in that
# witness's own column of the report, and where it belongs in the base's
# reading text in the golden document.
PageMark = namedtuple("PageMark", "witness own_at base_at text")


CollatedCell = namedtuple(
    "CollatedCell",
    "segs norms diffs base_missing has_diff note lemma lemma_end page_marks",
)


def collate_cells(aligned, labels, ignore_shad=True, positive=False, markers=None):
    """Walk the aligned rows once and decide everything about each cell.

    The report and the golden document used to run separate copies of this
    logic, each calling build_note_text() itself. The copies had to be kept
    identical by hand: if they ever disagreed about whether a cell yields a
    note, every footnote from that point on attached to the wrong word while
    still looking correctly numbered. Deciding once, here, removes that whole
    failure mode — both exporters consume this list.

    ``aligned`` is base-first, one row per witness; "-" marks an aligner gap.
    ``markers`` optionally gives each witness's page markers from
    extract_page_markers(). Each witness is tracked through its own text, so
    every cell records — in ``page_marks`` — the markers of any witnesses that
    begin a new page there. The golden document emits those inline, which is
    how one reading text can show where all the witnesses turned their pages.
    """
    n = len(aligned)
    width = max((len(r) for r in aligned), default=0)
    rows = [list(r) + [""] * (width - len(r)) for r in aligned]

    markers = markers or [[] for _ in range(n)]
    # per-witness running offset into its own (marker-stripped) text, and the
    # index of the next marker not yet reached
    offsets = [0] * n
    m_idx = [0] * n

    cells = []
    for j in range(width):
        raw = [rows[i][j] for i in range(n)]
        segs = ["" if c == "-" else c for c in raw]

        # A marker belongs to the cell that contains its offset, not the next
        # cell to start at or after it — testing against the cell's start
        # pushed any marker falling mid-cell onto the following word.
        page_marks = []
        for i in range(n):
            mk = markers[i] if i < len(markers) else []
            end = offsets[i] + len(segs[i])
            while m_idx[i] < len(mk) and mk[m_idx[i]][0] < max(end, offsets[i] + 1):
                pos, text = mk[m_idx[i]]
                within = max(0, pos - offsets[i])
                # own_at places it in this witness's own column in the report;
                # base_at places it in the base's reading text in the golden
                # document, where every witness's markers are woven together
                page_marks.append(
                    PageMark(
                        witness=i,
                        own_at=_base_index_for_mark(segs[i], segs[i], within),
                        base_at=_base_index_for_mark(segs[0], segs[i], within),
                        text=text,
                    )
                )
                m_idx[i] += 1
            offsets[i] = end
        page_marks.sort(key=lambda pm: pm.base_at)
        norms = [strip_ignorable(s, ignore_shad) for s in segs]
        base_norm = norms[0]

        # A witness omits the base's reading when its cell is a gap, or holds
        # only ignorable characters, while the base actually has content there.
        missing = [
            (raw[i] == "-" or (segs[i] and norms[i] == "")) and base_norm != ""
            for i in range(1, n)
        ]
        base_missing = base_norm == "" and any(norms[i] != "" for i in range(1, n))

        diffs = []
        for i in range(1, n):
            if missing[i - 1]:
                diffs.append(True)
            elif base_norm or norms[i]:
                diffs.append(base_norm != norms[i])
            else:
                diffs.append(False)

        has_diff = (base_norm != "" and any(diffs)) or base_missing

        note, lemma, lemma_end = "", "", 0
        if has_diff:
            note, lemma, lemma_end = build_note_text(
                segs, labels, positive=positive, ignore_shad=ignore_shad
            )

        cells.append(
            CollatedCell(segs, norms, diffs, base_missing, has_diff, note,
                         lemma, lemma_end, page_marks)
        )

    # Markers past the last cell (a page turning at the very end) still belong
    # in the output, so they ride on the final cell.
    trailing = []
    if cells:
        for i in range(n):
            mk = markers[i] if i < len(markers) else []
            trailing.extend(
                PageMark(witness=i,
                         own_at=len(cells[-1].segs[i]),
                         base_at=len(cells[-1].segs[0]),
                         text=m[1])
                for m in mk[m_idx[i]:]
            )
    if trailing and cells:
        cells[-1] = cells[-1]._replace(
            page_marks=sorted(cells[-1].page_marks + trailing,
                              key=lambda pm: pm.base_at)
        )
    return cells


def _write_cell(para, text, marks, shade):
    """Write one witness's cell into its column, splicing in its own markers.

    A witness's pagination belongs in that witness's column, at the point in
    its own text where the page turns — unlike the golden document, where
    every witness's markers are woven into the one reading text.
    """
    pos = 0
    for pm in sorted(marks, key=lambda m: m.own_at):
        at = max(pos, min(pm.own_at, len(text)))
        if at > pos:
            run = para.add_run(text[pos:at])
            if shade:
                set_run_background_color(run, shade)
            pos = at
        mark_run = para.add_run(pm.text)
        mark_run.italic = True
    if pos < len(text):
        run = para.add_run(text[pos:])
        if shade:
            set_run_background_color(run, shade)


# ── Verse structure ──────────────────────────────────────────────────
# Tibetan verse is isosyllabic: every pada of a passage carries the same
# syllable count, and a shad closes each one. That makes the boundaries
# findable — but only the metre can find them, because the shads are not
# reliable. A witness may write one inside a pada, where a pecha line ended,
# and may omit one between two padas entirely. Counting how often it omits
# them says something about the witness, which is why the profile reports it.
#
# Nothing here rewrites a text; it only measures.

VERSE_METRES = (7, 9, 11)     # a 13 or 15 admits prose clauses as "verse"
VERSE_MIN_RUN = 3             # fewer lines than this is not yet a passage
VERSE_MIN_RUN_UNUSUAL = 5     # a metre outside the classical set must show more
VERSE_TOL = 1                 # a pada may run a syllable over or short
VERSE_MAX_SPLIT = 2           # one missing shad, not a run of them — see below
_VERSE_SHADS = "\u0f0d\u0f0e\u0f0f\u0f10\u0f11\u0f14/|"
# A separator is a run of shads, even where they are written apart: GX1 writes
# the double shad as "/ /" and pecha-style Unicode as "\u0f0d    \u0f0d", and
# read as two singles every pada boundary disappears. Matching the whole run
# here keeps it one separator without deleting anything, so character
# positions still line up with the text the document is built from.
_VERSE_SPLIT_RE = re.compile(
    "([%s](?:[\\s_]*[%s])*)" % (re.escape(_VERSE_SHADS), re.escape(_VERSE_SHADS))
)
_VERSE_SYL_RE = re.compile("[\u0f0b\u0f0c\s_]+")
_VERSE_TAG_RE = re.compile(r"\[[^\[\]\n]*\]")

# One line of verse: the reading, the shad that closes it, its syllable count,
# and whether that shad had to be reconstructed because the witness omitted it.
VerseLine = namedtuple("VerseLine", "text shad syllables reconstructed end")
# One passage: its lines, its metre, and the syllable offset it begins at.
VerseBlock = namedtuple("VerseBlock", "lines metre at start")


def _verse_syllables(text):
    """Syllables, by tsheg in Tibetan script and by space in Wylie."""
    text = re.sub("[%s]" % re.escape(_VERSE_SHADS), " ", text)
    return [x for x in _VERSE_SYL_RE.split(text.strip()) if x]


def _verse_join(words, tibetan):
    """Rejoin syllables in their own script."""
    return (TSHEG.join(words) + TSHEG) if tibetan else " ".join(words)


def strip_page_tags(text):
    """Remove bracketed page/folio tags, remembering where each one stood.

    Returns ``(clean, marks)`` with marks as ``[(syllable_index, tag)…]``.
    A tag left in is read as a syllable and throws the metre off by one for
    every pada it touches — the same reason they are stripped before the
    collation, applied to counting rather than comparing. Remembering the
    positions is what lets an omission be cited by folio afterwards.
    """
    out, marks, seen, pos = [], [], 0, 0
    for m in _VERSE_TAG_RE.finditer(text):
        chunk = text[pos:m.start()]
        out.append(chunk)
        seen += len(_verse_syllables(chunk))
        marks.append((seen, m.group(0)))
        # blanked, not removed: every character position in the result still
        # matches the original, which is what lets a line break found here be
        # placed in the text the document is built from
        out.append(" " * (m.end() - m.start()))
        pos = m.end()
    out.append(text[pos:])
    return "".join(out), marks


def _verse_measurable(text):
    """The part of a text whose syllables can be counted.

    Page tags go, because a tag left in is read as a syllable and throws the
    metre off for every pada it touches. In a Tibetan-script text, runs of
    Latin script go too: a document may carry an English heading or a note,
    and "Part 2. Diamond Mind meditation" counted as five Tibetan syllables
    invents verse that is not there.
    """
    clean, marks = strip_page_tags(text)
    if _TIBETAN_CHAR_RE.search(clean):
        clean = re.sub(r"[A-Za-z][A-Za-z0-9.,'\-]*",
                       lambda m: " " * len(m.group(0)), clean)
    return clean, marks


def _verse_segments(text):
    """Split at shads, each segment keeping the shad that closes it.

    A run of shads counts as one separator even when written apart — see
    _VERSE_SPLIT_RE — so "/ /" closes a pada exactly as "//" does.
    """
    out, pos = [], 0
    for piece in _VERSE_SPLIT_RE.split(text):
        start = pos
        pos += len(piece)
        if piece and piece[0] in _VERSE_SHADS:
            if out:
                out[-1][1] = piece
                out[-1][3] = pos          # just past the closing shad
        elif piece.strip():
            spans = [(m.group(0), m.start() + start, m.end() + start)
                     for m in re.finditer(
                         r"[^\u0f0b\u0f0c\s_%s]+" % re.escape(_VERSE_SHADS), piece)]
            out.append([[w for w, _a, _b in spans], "", spans, pos])
    return [(w, sh, sp, e) for w, sh, sp, e in out if w]


def _verse_lines_at(segs, i, metre, tibetan):
    """Build lines of `metre` syllables from segs[i:], joining only.

    Never splitting is what lets a pada the witness broke across a line come
    back whole; it is also why a missing shad cannot be recovered here.
    """
    lines, k = [], i
    while k < len(segs):
        buf, n = [], 0
        j = k
        while j < len(segs) and n < metre - VERSE_TOL:
            buf.append(segs[j]); n += len(segs[j][0]); j += 1
        if not buf or not (metre - VERSE_TOL <= n <= metre + VERSE_TOL):
            break
        words = [w for seg in buf for w in seg[0]]
        lines.append(VerseLine(_verse_join(words, tibetan), buf[-1][1], n,
                               False, buf[-1][3]))
        k = j
    return lines, k


def _verse_metrical_before(segs, k, metre):
    """Is there a metrical line immediately before segs[k] — alone, or as two
    segments joined across a shad?"""
    if k <= 0:
        return False
    n = len(segs[k - 1][0])
    if metre - VERSE_TOL <= n <= metre + VERSE_TOL:
        return True
    if k > 1:
        n += len(segs[k - 2][0])
        return metre - VERSE_TOL <= n <= metre + VERSE_TOL
    return False


def _verse_extend_back(segs, start, metre, lines, tibetan, limit=0):
    """Walk back from a confirmed passage, recovering padas the punctuation
    hid, and return the index the passage really begins at.

    A segment may be split only where it is an EXACT multiple of the metre and
    a metrical line still stands behind it. Both conditions are needed: a
    21-syllable prose clause is also three sevens, and is given away by having
    nothing metrical before it. Without them this cuts prose into pieces, and
    through the middle of words.
    """
    i = start
    while i > limit:
        words, shad, spans, seg_end = segs[i - 1]
        n = len(words)
        parts = round(n / metre) if metre else 0
        # Only ever two. A segment of nine times the metre is not nine padas
        # with eight shads missing in a row; it is a prose clause whose length
        # divides by chance, which happens about one time in seven. DX1 has
        # one of 63 syllables, and without this cap it was cut into nine
        # "padas" straight through the middle of words.
        if 2 <= parts <= VERSE_MAX_SPLIT and n == parts * metre and \
                _verse_metrical_before(segs, i - 1, metre):
            made = []
            for p in range(parts):
                lo, hi = p * metre, (p + 1) * metre
                chunk = words[lo:hi]
                # an interior boundary ends at the last word of its chunk;
                # only the final piece reaches past the closing shad
                end = seg_end if p == parts - 1 else spans[hi - 1][2]
                made.append(VerseLine(_verse_join(chunk, tibetan), shad, metre,
                                      p < parts - 1, end))
            lines[:0] = made
            i -= 1
            continue
        if metre - VERSE_TOL <= n <= metre + VERSE_TOL:
            lines.insert(0, VerseLine(_verse_join(words, tibetan), shad, n,
                                      False, seg_end))
            i -= 1
            continue
        if i - 2 >= limit:
            prev = segs[i - 2][0]
            m = len(prev) + n
            if metre - VERSE_TOL <= m <= metre + VERSE_TOL:
                lines.insert(0, VerseLine(_verse_join(prev + words, tibetan),
                                          shad, m, False, seg_end))
                i -= 2
                continue
        break
    return i


def _verse_extend_on(segs, stop, metre, lines, tibetan):
    """Carry a confirmed passage forward past a boundary the witness omitted.

    The mirror of _verse_extend_back, and needed for the same reason: a run of
    intact padas ends the moment it meets a segment of twice the metre, and
    without this the omission sitting just after a passage is never reached.
    The same two conditions apply — an exact multiple, and a metrical line
    already standing before it, which here is the passage itself.
    """
    i = stop
    while i < len(segs):
        words, shad, spans, seg_end = segs[i]
        n = len(words)
        parts = round(n / metre) if metre else 0
        if 2 <= parts <= VERSE_MAX_SPLIT and n == parts * metre:
            for p in range(parts):
                lo, hi = p * metre, (p + 1) * metre
                chunk = words[lo:hi]
                end = seg_end if p == parts - 1 else spans[hi - 1][2]
                lines.append(VerseLine(_verse_join(chunk, tibetan), shad, metre,
                                       p < parts - 1, end))
            i += 1
            continue
        if metre - VERSE_TOL <= n <= metre + VERSE_TOL:
            lines.append(VerseLine(_verse_join(words, tibetan), shad, n,
                                   False, seg_end))
            i += 1
            continue
        break
    return i


def verse_blocks(text):
    """Every verse passage in a text, as VerseBlock records.

    The metres are read off the text itself rather than assumed, so a
    translation in a longer metre is found instead of walked past.

    A passage is a run of at least VERSE_MIN_RUN metrical lines most of which
    close with a double shad. That last test is what keeps prose out: verse
    lines end in a double 87-94% of the time across the Jataka witnesses,
    prose segments only 29-53%, and without it any run of prose clauses that
    happens to fall near a metre reads as verse.

    Page tags are stripped first; the text is measured, never rewritten.
    """
    tibetan = bool(_TIBETAN_CHAR_RE.search(text))
    clean, _marks = _verse_measurable(text)
    metres = derive_metres(text)
    segs = _verse_segments(clean)
    starts, seen = [], 0
    for words, _sh, _sp, _e in segs:
        starts.append(seen); seen += len(words)
    blocks, i, run = [], 0, 0
    while i < len(segs):
        found = None
        for m in metres:            # ascending: a 15 would eat two sevens
            lines, j = _verse_lines_at(segs, i, m, tibetan)
            # 7, 9 and 11 are the expected lengths and three lines settle
            # them. A length the text merely happens to use often has to show
            # more: a prose-heavy work puts a common clause length at the top
            # of the derived list, and four such clauses in a row are easy to
            # come by — BX1 has exactly one such run at fifteen.
            need = VERSE_MIN_RUN if m in VERSE_METRES else VERSE_MIN_RUN_UNUSUAL
            if len(lines) < need:
                continue
            doubled = sum(1 for l in lines if len(l.shad) >= 2)
            if doubled * 2 < len(lines):
                continue
            found = (lines, j, m)
            break
        if found:
            lines, j, m = found
            back = _verse_extend_back(segs, i, m, lines, tibetan, limit=run)
            j = _verse_extend_on(segs, j, m, lines, tibetan)
            blocks.append(VerseBlock(lines, m, starts[back],
                                     segs[back][2][0][1]))
            i = run = j
        else:
            i += 1
    return blocks


VERSE_STANZA = 4              # padas to a stanza; see golden_layout
VERSE_METRE_SHARE = 0.10      # a metre must hold a tenth of the closes


def derive_metres(text):
    """The metres a text actually uses, commonest first.

    Reading them off the text rather than assuming 7/9/11 is what lets a
    translation in a longer metre be found at all: a run of clean 15-syllable
    padas is invisible to a detector that only ever tries three lengths.
    Only runs closed by a double shad are counted, since a single shad ends a
    prose clause just as readily and would put every prose length in the list.
    """
    clean, _marks = _verse_measurable(text)
    lengths, total = {}, 0
    for words, shad, _spans, _end in _verse_segments(clean):
        if len(shad.strip()) >= 2:
            lengths[len(words)] = lengths.get(len(words), 0) + 1
            total += 1
    if not total:
        return VERSE_METRES
    # A real metre dominates; noise does not. In these witnesses 7 and 9 take
    # 26-76% of the double-shad segments each while every other length sits
    # under 8%, so a tenth of the total separates them cleanly — and a text
    # written mostly in a longer metre puts that metre at the top instead.
    floor = max(VERSE_MIN_RUN, total * VERSE_METRE_SHARE)
    good = [(n, c) for n, c in lengths.items() if 5 <= n <= 25 and c >= floor]
    good.sort(key=lambda x: -x[1])
    # Union with the classical lengths rather than replacing them. A metre
    # this text uses rarely — GX1 closes only 5 segments at eleven — falls
    # under the floor but is still worth trying, and the shortest-first order
    # means a longer metre is never reached by a shorter one anyway.
    return tuple(sorted(set(n for n, _c in good[:4]) | set(VERSE_METRES)))


def golden_layout(text, stanza=VERSE_STANZA):
    """Where to break the reading text, as ``(offset, kind)`` pairs.

    ``kind`` is "line" for a pada, "stanza" for the last pada of a stanza, and
    "para" for the end of a prose sentence. Offsets are into ``text`` itself,
    so the document can break exactly there.

    Prose is broken only at a sentence-final particle — ngo, to, do, so, go,
    'o — because a shad marks a pause of any weight and breaking at every one
    would chop the prose into clauses.
    """
    blocks = verse_blocks(text)
    clean, _marks = _verse_measurable(text)
    tibetan = bool(_TIBETAN_CHAR_RE.search(text))
    finals = _FINAL_T if tibetan else _FINAL_W
    spans = [(b.start, b.lines[-1].end) for b in blocks]

    breaks = []
    for b in blocks:
        if b.start > 0:
            # the passage opens a line of its own; without this its first
            # pada sits at the end of the prose paragraph above it
            breaks.append((b.start, "open"))
        for n, line in enumerate(b.lines):
            last = n + 1 == len(b.lines)
            at_stanza = (n + 1) % stanza == 0 and not last
            breaks.append((line.end, "stanza" if at_stanza else "line"))
    for words, shad, _sp, end in _verse_segments(clean):
        if any(lo <= end <= hi for lo, hi in spans):
            continue                       # inside verse; already broken
        if words and words[-1] in finals:
            breaks.append((end, "para"))
    # Carry each break past whatever blank follows the shad, so a line opens
    # on real text. Wylie writes an explicit space as "_", and a source that
    # ends a line with "pa//_" would otherwise start the next pada with a
    # stranded underscore.
    blank = re.compile(r"[\s_]+")
    moved = []
    for at, kind in breaks:
        m = blank.match(text, at)
        moved.append(((m.end() if m else at), kind))
    moved.sort()
    return moved, spans


def omitted_pada_shads(text):
    """How many pada boundaries this witness leaves unmarked.

    NOT_APPLICABLE when no verse is found at all — which is a real outcome,
    not a zero: a transcription that never distinguishes the double shad
    yields no passages, and reporting 0 there would claim the witness omits
    nothing when in fact nothing could be looked for.
    """
    blocks = verse_blocks(text)
    if not blocks:
        return NOT_APPLICABLE
    return sum(1 for b in blocks for l in b.lines if l.reconstructed)


def _junction_in(text, left_tail, right_head):
    """What one witness writes between two padas, as written.

    Returns ``(separator, folio)`` — the separator as written ("//", "/ /", a
    Unicode shad, or "" for nothing at all) and the page tag the junction
    falls under, or ``(None, "")`` when this witness does not have the
    passage. The tail and head are matched by syllable, so a variant reading
    elsewhere in the pada does not hide the junction.

    The folio comes from here rather than from the witness's own verse
    detection, so a witness identified as omitting only because the others
    mark the boundary still says where to look.
    """
    clean, marks = _verse_measurable(text)
    sylls, spans, pos = [], [], 0
    # Tokens must break exactly where _verse_syllables breaks — on the tsheg
    # as well as on whitespace — or a Tibetan-script text never matches its
    # own junctions and every witness reads as not having the passage.
    for m in re.finditer(r"[^\u0f0b\u0f0c\s_%s]+" % re.escape(_VERSE_SHADS), clean):
        sylls.append(m.group(0)); spans.append((m.start(), m.end()))
    want = list(left_tail) + list(right_head)
    n = len(left_tail)
    for i in range(len(sylls) - len(want) + 1):
        if sylls[i:i + len(want)] == want:
            gap = clean[spans[i + n - 1][1]:spans[i + n][0]]
            folio = ""
            for pos, mk in marks:
                if pos <= i + n:
                    folio = mk
            return re.sub(r"[\s_]+", " ", gap).strip(), folio
    return None, ""


def omitted_pada_sites(texts, labels, context=3):
    """Every pada boundary a witness leaves unmarked, with what the others
    write at the same place.

    A count says a witness runs padas together; it cannot say whether that is
    the witness or the OCR, and those call for opposite responses. The other
    witnesses settle it: where they all write a shad and one does not, the
    boundary is certain and the omission is that witness's. Where none of them
    writes it, the reading is shared and there is nothing to correct.
    """
    sites = []
    for label, text in zip(labels, texts):
        _clean, marks = _verse_measurable(text)
        for blk in verse_blocks(text):
            at = blk.at
            for n, line in enumerate(blk.lines):
                if line.reconstructed and n + 1 < len(blk.lines):
                    tag = ""
                    for pos, mk in marks:
                        if pos <= at:
                            tag = mk
                    left = _verse_syllables(line.text)
                    right = _verse_syllables(blk.lines[n + 1].text)
                    sites.append({
                        "by": label, "tag": tag,
                        "left": line.text, "right": blk.lines[n + 1].text,
                        "tail": left[-context:], "head": right[:context],
                    })
                at += line.syllables
    # one row per junction, however many witnesses omit it there
    rows, seen = [], {}
    for st_ in sites:
        key = (tuple(st_["tail"]), tuple(st_["head"]))
        if key in seen:
            seen[key]["omits"][st_["by"]] = st_["tag"]
            continue
        witnesses, folios = {}, {}
        for label, text in zip(labels, texts):
            sep, folio = _junction_in(text, st_["tail"], st_["head"])
            witnesses[label] = sep
            folios[label] = folio
        row = {"left": st_["left"], "right": st_["right"],
               "witnesses": witnesses, "folios": folios,
               "omits": {st_["by"]: st_["tag"]}}
        seen[key] = row
        rows.append(row)
    return rows


# What each row of the orthographic profile counts. These are features the
# collation normalises away, so without this table they leave no trace — yet
# they describe a witness: which scriptorium stacked its Sanskrit loans, which
# transcription used EWTS conventions. Stated once here rather than repeated
# through the apparatus, where a pattern of nine identical notes is invisible.
#
# Typographic apostrophes and non-breaking spaces are deliberately absent.
# They are artefacts of how a file was produced, not properties of the
# witness; both are still normalised, they are simply not evidence.
_STACK_WORD_RE = re.compile(r"[A-Za-z']*\+[A-Za-z']*")

# Unicode counterparts of the marks the Wylie counters look for. Subjoined
# consonants (U+0F90…) are deliberately absent: they are the Unicode spelling
# of every ordinary syllable, so counting them would answer "how much Tibetan
# is here", not "which loans did this scriptorium stack". EWTS + records a
# transcriber's choice; Unicode does not record that choice at all.
_UNI_HEAD_MARKS = "༄༅༆༇༈"
_UNI_SHAD = "།༎༏༐༑༔"
_NB_TSHEG = "༌"

_FINAL_W = {"ngo", "to", "do", "so", "go", "'o", "no", "bo", "mo", "ro", "lo"}
_FINAL_T = {"ངོ", "ཏོ", "དོ", "སོ", "གོ", "འོ", "ནོ", "བོ", "མོ", "རོ", "ལོ"}

_OMITTED_PADA_ROW = "Pāda-final shads omitted"

WYLIE, TIBETAN = "wylie", "tibetan"
BOTH_SCRIPTS = frozenset((WYLIE, TIBETAN))
NOT_APPLICABLE = "–"


def _script_of(text):
    """Which script a witness is written in, by the same test used for notes."""
    return TIBETAN if _TIBETAN_CHAR_RE.search(text) else WYLIE


# Each feature carries the marks it looks for in each script, so the row can
# be labelled in whichever script the witnesses are actually written in: a
# Wylie run should not grow Tibetan characters in its headings, nor the
# reverse. Order is the order the rows print in.
#
# ``scripts`` says where the feature can exist at all. A feature absent from
# every witness's script is not listed; one absent from a single witness's
# script shows NOT_APPLICABLE for that column, because a zero there would
# claim the text was searched and found wanting.
ORTHOGRAPHIC_FEATURES = (
    ("Stacked consonants", "+", None, frozenset((WYLIE,)),
     lambda t: len(re.findall(r"[A-Za-z']\+[A-Za-z']", t))),
    ("Head marks", "@ # !", "༄༅", BOTH_SCRIPTS,
     lambda t: sum(t.count(c) for c in "@#!")
     + sum(t.count(c) for c in _UNI_HEAD_MARKS)),
    ("Pipe written for shad", "|", "|", BOTH_SCRIPTS, lambda t: t.count("|")),
    ("Explicit space", "_", None, frozenset((WYLIE,)), lambda t: t.count("_")),
    ("Shad", "/", "།", BOTH_SCRIPTS,
     lambda t: t.count("/") + sum(t.count(c) for c in _UNI_SHAD)),
    ("Non-breaking tsheg", None, "༌", frozenset((TIBETAN,)),
     lambda t: t.count(_NB_TSHEG)),
    # The one row that counts something absent rather than present: how often
    # a witness runs two padas together with no shad between them. Invisible
    # in the apparatus, because the words are all there and in order — it is
    # the punctuation that is missing, and only the metre reveals it.
    (_OMITTED_PADA_ROW, None, None, BOTH_SCRIPTS, omitted_pada_shads),
)


def _feature_label(name, wylie_marks, tib_marks, present):
    """Name the row in the script(s) the witnesses are written in."""
    marks = []
    if TIBETAN in present and tib_marks:
        marks.append(tib_marks)
    if WYLIE in present and wylie_marks and wylie_marks not in marks:
        marks.append(wylie_marks)
    return f"{name}  {' '.join(marks)}" if marks else name


def orthographic_profile(texts, labels):
    """Per-witness counts of the features normalised before comparison.

    Returns ``(rows, stacked)`` where rows is ``[(feature, [counts…])…]`` and
    stacked maps each label to the distinct stacked forms it uses. Counts are
    integers, or NOT_APPLICABLE where the feature cannot occur in that
    witness's script.

    Rows and their labels follow the scripts actually present, so an all-Wylie
    run reads exactly as it always did and an all-Unicode one is not a column
    of zeros standing for features that were never findable.
    """
    scripts = [_script_of(t) for t in texts]
    present = frozenset(scripts)
    rows = []
    if len(present) > 1:
        # Only worth saying when it varies; it is what the dashes below mean.
        rows.append(("Script", ["Tibetan" if s is TIBETAN else "Wylie"
                                for s in scripts]))
    for name, wylie_marks, tib_marks, feat_scripts, fn in ORTHOGRAPHIC_FEATURES:
        if not (feat_scripts & present):
            continue
        counts = [fn(t) if s in feat_scripts else NOT_APPLICABLE
                  for t, s in zip(texts, scripts)]
        if name == _OMITTED_PADA_ROW and len(texts) > 1:
            # Count against the junctions the witnesses establish between
            # them, not against what each could find alone. A witness that
            # omits several boundaries may fall below the run needed to
            # detect any verse at all, and would report "not applicable"
            # while the table below plainly shows it omitting them.
            sites = omitted_pada_sites(texts, labels)
            if sites:
                # "" is the witness writing nothing between two padas;
                # None is the witness not having the passage at all, which is
                # not an omission and must not be counted as one.
                counts = [sum(1 for r in sites if r["witnesses"].get(lb) == "")
                          for lb in labels]
        rows.append((
            _feature_label(name, wylie_marks, tib_marks, present & feat_scripts),
            counts,
        ))
    stacked = {}
    for label, text in zip(labels, texts):
        forms = sorted({m.group(0).strip("/") for m in _STACK_WORD_RE.finditer(text)})
        stacked[label] = [f for f in forms if f]
    return rows, stacked


PROFILE_TABLE_STYLE = "Light Grid Accent 1"


def _style_table(table, numeric_from=1):
    """Give a table borders, a bold header, and right-aligned figures."""
    try:
        table.style = PROFILE_TABLE_STYLE
    except KeyError:  # a template without the built-in styles
        try:
            table.style = "Table Grid"
        except KeyError:
            pass
    table.autofit = True
    for cell in table.rows[0].cells:
        for para in cell.paragraphs:
            for run in para.runs:
                run.bold = True
    for row in table.rows[1:]:
        for cell in row.cells[numeric_from:]:
            text = cell.text.strip()
            if text and (text == NOT_APPLICABLE or all(ch.isdigit() for ch in text)):
                for para in cell.paragraphs:
                    para.alignment = WD_ALIGN_PARAGRAPH.RIGHT


def _trim_shared_for_display(readings):
    """Trim syllables shared by every witness, comparing what is printed.

    The apparatus trims by comparison key, where "seng ge'i" and "seng+ge'i"
    count as the same and would vanish. Here the stacking *is* the subject, so
    the printed form decides.
    """
    lists = [list(r) for r in readings]
    n = min(len(r) for r in lists)
    pre = 0
    while pre < n and all(r[pre] == lists[0][pre] for r in lists):
        pre += 1
    lists = [r[pre:] for r in lists]
    n = min(len(r) for r in lists)
    suf = 0
    while suf < n and all(r[-1 - suf] == lists[0][-1 - suf] for r in lists):
        suf += 1
    if suf:
        lists = [r[: len(r) - suf] for r in lists]
    return lists


def stacked_form_rows(cells, ignore_shad=True):
    """Every point where some witness stacks, and what each witness reads.

    Answers the question the bare list of forms could not: BX1 writes d+hi —
    what do the others have there? Readings are trimmed to the differing part,
    and identical patterns are counted rather than repeated.
    """
    counts = {}
    order = []
    for cell in cells:
        if not any("+" in seg for seg in cell.segs):
            continue
        trimmed = _trim_shared_for_display(
            [_syllables(seg, ignore_shad) for seg in cell.segs]
        )
        reading = tuple(" ".join(t) or OMITTED_MARK for t in trimmed)
        if reading not in counts:
            order.append(reading)
        counts[reading] = counts.get(reading, 0) + 1
    order.sort(key=lambda r: (-counts[r], r))
    return [(r, counts[r]) for r in order]


def add_orthographic_profile(doc, texts, labels, cells=None):
    """Append the profile, and the stacked-form comparison, to the report."""
    rows, _stacked = orthographic_profile(texts, labels)
    doc.add_paragraph()
    doc.add_heading("Orthographic profile", level=2)
    intro = (
        "Features normalised before comparison and therefore absent from the "
        "apparatus. They describe the witnesses rather than the text. The "
        "last row counts something absent instead: pādas the witness runs "
        "together with no shad between them, which only the metre reveals."
    )
    if any(_script_of(t) is TIBETAN for t in texts):
        # Otherwise a Tibetan-script run looks as though the stacking count
        # failed, when in fact there is no such count to take.
        intro += (
            " Explicit stacking is a feature of Wylie transcription: the "
            "EWTS + records a choice the transcriber made, which Tibetan "
            "script does not distinguish, so it is not counted there."
        )
    doc.add_paragraph(intro)
    table = doc.add_table(rows=1, cols=len(labels) + 1)
    hdr = table.rows[0].cells
    hdr[0].paragraphs[0].add_run("")
    for i, label in enumerate(labels):
        hdr[i + 1].paragraphs[0].add_run(label)
    for name, counts in rows:
        cs = table.add_row().cells
        cs[0].paragraphs[0].add_run(name)
        for i, n in enumerate(counts):
            cs[i + 1].paragraphs[0].add_run(str(n))
    _style_table(table)

    stacked_rows = stacked_form_rows(cells) if cells else []
    if not stacked_rows:
        return

    doc.add_paragraph()
    doc.add_heading("Stacked forms across the witnesses", level=2)
    doc.add_paragraph(
        "Each point where a witness writes a stacked consonant, and what the "
        "others read there. Identical patterns are counted, not repeated. "
        "A longer row means the stacked word sits beside another variant."
    )
    st_table = doc.add_table(rows=1, cols=len(labels) + 1)
    hdr = st_table.rows[0].cells
    for i, label in enumerate(labels):
        hdr[i].paragraphs[0].add_run(label)
    hdr[len(labels)].paragraphs[0].add_run("times")
    for reading, n in stacked_rows:
        cs = st_table.add_row().cells
        for i, text in enumerate(reading):
            cs[i].paragraphs[0].add_run(text)
        cs[len(labels)].paragraphs[0].add_run(str(n))
    _style_table(st_table, numeric_from=len(labels))

    sites = omitted_pada_sites(texts, labels)
    if not sites:
        return

    doc.add_paragraph()
    doc.add_heading("Pādas run together", level=2)
    doc.add_paragraph(
        "Where a witness writes two pādas with no shad between them, and what "
        "the others write at the same place. The count above says how often; "
        "this says where, and lets the witnesses answer for each other — a "
        "boundary the others all mark is certainly a boundary, so the omission "
        "is that witness's and can be corrected. One none of them marks is a "
        "shared reading, and there is nothing to correct."
    )
    st_cols = 2 + len(labels)
    site_table = doc.add_table(rows=1, cols=st_cols)
    hdr = site_table.rows[0].cells
    hdr[0].paragraphs[0].add_run("Pāda")
    hdr[1].paragraphs[0].add_run("Runs into")
    for i, label in enumerate(labels):
        hdr[2 + i].paragraphs[0].add_run(label)
    for row in sites:
        cs = site_table.add_row().cells
        cs[0].paragraphs[0].add_run(row["left"])
        cs[1].paragraphs[0].add_run(row["right"])
        for i, label in enumerate(labels):
            cell = cs[2 + i].paragraphs[0]
            sep = row["witnesses"].get(label)
            if sep is None:
                cell.add_run("—")          # this witness lacks the passage
            elif sep == "":
                run = cell.add_run(OMITTED_MARK)
                run.bold = True
                folio = row["folios"].get(label) or row["omits"].get(label, "")
                if folio:
                    cell.add_run("  " + folio)
            else:
                cell.add_run(sep)
    _style_table(site_table, numeric_from=st_cols)


def export_collation_report(cells, labels, names, profile_texts=None):
    """Side-by-side table of every witness, plus the numbered notes list.

    The page is landscape: with up to six witness columns a portrait page
    squeezes Tibetan script too narrow to read comfortably.
    """
    doc = Document()
    set_report_page(doc)
    doc.add_heading("Tibetan Collation Report", level=1)
    header = f"Base / golden: {names[0]}"
    for i, nm in enumerate(names[1:], start=1):
        header += f"  |  Comparison {i}: {nm}"
    doc.add_paragraph(header)

    n = len(labels)
    table = doc.add_table(rows=2, cols=n)
    hdr = table.rows[0].cells
    hdr[0].text = f"{labels[0]} (golden, notes)"
    for i in range(1, n):
        hdr[i].text = labels[i]

    paras = [table.rows[1].cells[i].paragraphs[0] for i in range(n)]

    notes = []
    note_active = False
    current_note_color = None
    color_idx = -1

    for cell in cells:
        # Every differing cell gets its own note. Adjacent differences are
        # separate variants (e.g. a particle change followed by a word
        # change), so grouping them would silently drop all but the first.
        note_start_here = False
        if cell.note:
            notes.append(cell.note)
            note_start_here = True
            note_active = True
            color_idx = (color_idx + 1) % len(COLOR_LIST)
            current_note_color = COLOR_LIST[color_idx]

        color = current_note_color if note_active else None
        note_number = len(notes)

        for i in range(n):
            seg_i = cell.segs[i]
            marks_i = [pm for pm in cell.page_marks if pm.witness == i]
            if i == 0:
                differs = any(cell.diffs)
                show_note = note_start_here
            else:
                differs = cell.diffs[i - 1] or cell.base_missing
                show_note = note_start_here and differs
            if not seg_i and not marks_i and not show_note:
                continue
            shade = color if (color and differs and cell.norms[i] != "") else None
            _write_cell(paras[i], seg_i, marks_i, shade)
            if show_note:
                m = paras[i].add_run(f"[{note_number}]")
                m.font.superscript = True

        if note_active and not cell.has_diff:
            note_active = False
            current_note_color = None

    if notes:
        doc.add_paragraph()
        doc.add_heading("Notes", level=2)
        for i, text in enumerate(notes, start=1):
            p = doc.add_paragraph()
            r = p.add_run(f"{i}")
            r.font.superscript = True
            r.bold = True
            p.add_run(" ")
            p.add_run(text)

    if profile_texts:
        add_orthographic_profile(doc, profile_texts, labels, cells=cells)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf, notes


def _is_shad_only(lemma: str) -> bool:
    """True when a note's lemma is nothing but shad punctuation.

    The lemma is what the footnote mark should sit on; when it is a shad the
    mark must stay there rather than move back onto the preceding word. The
    lemma now comes straight from build_note_text() instead of being re-parsed
    out of the finished note string, which also fixes the positive-apparatus
    case where the lemma carries several sigla ("V1, V3 /] …").
    """
    return bool(lemma) and all(c in SHAD_CHARS or c.isspace() for c in lemma)


def export_versions_document(texts, labels, patterns=None):
    """Every witness in sequence, one after another, for keeping.

    The report shows the witnesses side by side for comparison; this shows
    each one whole, which is what an archived text wants to be. Page markers
    are kept and italicised so the pagination survives; nothing is
    highlighted and no note references appear — the apparatus lives in the
    other two documents.
    """
    doc = Document()
    doc.add_heading("All source versions", level=1)
    doc.add_paragraph(
        "Each witness in full, in the order collated. "
        + ", ".join(labels)
    )

    # each witness may mark its pages its own way, so each gets its own
    compiled = []
    for i in range(len(texts)):
        pat = patterns[i] if patterns and i < len(patterns) else None
        try:
            compiled.append(re.compile(pat) if pat else None)
        except re.error:
            compiled.append(None)

    for idx, (label, text) in enumerate(zip(labels, texts)):
        rx = compiled[idx]
        doc.add_paragraph()
        doc.add_heading(label, level=2)
        para = doc.add_paragraph()
        if rx is None:
            para.add_run(text)
            continue
        pos = 0
        for m in rx.finditer(text):
            if m.start() > pos:
                para.add_run(text[pos:m.start()])
            mark = para.add_run(m.group(0))
            mark.italic = True
            pos = m.end()
        if pos < len(text):
            para.add_run(text[pos:])

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def _style_verse_line(p):
    """A pada sits on its own line, indented, with its stanza-mates close."""
    fmt = p.paragraph_format
    fmt.left_indent = Inches(0.35)
    fmt.space_after = Pt(0)
    fmt.space_before = Pt(0)
    return p


def export_golden_with_footnotes(cells, notes, labels, name1="base",
                                 milestones=None, ignore_shad=True,
                                 profile_texts=None, verse_layout=False):
    """Golden text with variant footnotes.

    ``cells`` is the very list the report was built from, so footnote numbering
    matches the report's notes by construction rather than by keeping a second
    copy of the diff logic in step with the first.

    Every witness's page markers are woven back into the running text at the
    cell where that witness turns its page, in italics, so one reading text
    shows where all the witnesses stood. The base's own markers are included;
    removing them all again gives back the base exactly.

    ``milestones`` is an optional list of ``(offset, tag)`` pairs from
    extract_folio_tags(): folio/page tags stripped before collation that are
    re-inserted here — as italic runs at their original character positions —
    without ever having been part of the alignment or the apparatus.
    """
    milestones = sorted(milestones) if milestones else []
    ms_index = 0  # next milestone still to be emitted
    char_pos = 0  # running offset into the (concatenated) base text

    # Where the reading text should break, if verse is to be set as verse.
    # Computed over the same concatenation the cells are emitted from, so an
    # offset here is the offset the document writes at.
    base_text = "".join(c.segs[0] for c in cells)
    breaks, verse_spans, verse_note = ([], [], "")
    if verse_layout:
        breaks, verse_spans = golden_layout(base_text)
        blocks = verse_blocks(base_text)
        if blocks:
            metres = sorted({b.metre for b in blocks})
            padas = sum(len(b.lines) for b in blocks)
            verse_note = (
                "Verse set one pāda per line, in stanzas of %d: %d passages, "
                "%d pādas, %s syllables. The file's own line breaks are "
                "discarded as the wrapping they are; not a character of the "
                "text is altered." % (
                    VERSE_STANZA, len(blocks), padas,
                    ", ".join(str(m) for m in metres[:-1]) + " and " + str(metres[-1])
                    if len(metres) > 1 else str(metres[0]))
            )
        else:
            # Silence here would read as "this text has no verse", which is
            # not what was found — nothing could be looked for.
            verse_note = (
                "No verse found, so the text is set as prose. Pāda boundaries "
                "are located by the double shad; a transcription that does not "
                "distinguish it from the single shad gives nothing to find."
            )

    def in_verse(at):
        return any(lo <= at < hi for lo, hi in verse_spans)

    def emit(p, text):
        """Write base text, splicing in any milestone tags it spans.

        With verse_layout on, the source's own line breaks are dropped: they
        are where the file happened to wrap, not where the text divides, and
        left in they cut across the padas being set.
        """
        nonlocal ms_index, char_pos
        if verse_layout:
            text = text.replace("\r", " ").replace("\n", " ")
        start = 0
        end_pos = char_pos + len(text)
        while ms_index < len(milestones) and milestones[ms_index][0] < end_pos:
            cut = milestones[ms_index][0] - char_pos
            if cut > start:
                p.add_run(text[start:cut])
            tag_run = p.add_run(milestones[ms_index][1])
            tag_run.italic = True
            start = cut
            ms_index += 1
        if start < len(text):
            p.add_run(text[start:])
        char_pos = end_pos

    doc = Document()
    doc.add_heading(f"{labels[0]} \u2014 golden text with footnotes", level=1)
    if len(labels) > 2:
        others = ", ".join(labels[1:-1]) + " and " + labels[-1]
    else:
        others = labels[1]
    doc.add_paragraph(f"Base: {name1}  |  Footnotes from comparison with {others}.")
    if verse_note:
        note_p = doc.add_paragraph()
        note_run = note_p.add_run(verse_note)
        note_run.italic = True
        note_run.font.size = Pt(9)

    if profile_texts:
        # Front matter: the profile describes the witnesses, so it belongs
        # before the text rather than after it, and on its own page so the
        # reading begins at the top of one.
        #
        # Portrait throughout, like the rest of the document — the
        # stacked-forms table wraps within its columns rather than the page
        # changing shape under the reader.
        add_orthographic_profile(doc, profile_texts, labels, cells=cells)
        doc.add_page_break()
    else:
        doc.add_paragraph()

    p_text = doc.add_paragraph()
    if verse_layout and verse_spans and in_verse(0):
        _style_verse_line(p_text)

    def start_line(kind, at):
        """Close the current paragraph and open the next one."""
        nonlocal p_text
        if kind == "stanza":
            gap = doc.add_paragraph()
            gap.paragraph_format.space_after = Pt(0)
        p_text = doc.add_paragraph()
        if in_verse(at):
            _style_verse_line(p_text)

    bi = 0
    note_index = 0
    for cell in cells:
        seg = cell.segs[0]
        note_start_here = bool(cell.note)
        if note_start_here:
            note_index += 1
        place_note = note_start_here and 1 <= note_index <= len(notes)

        # Everything that has to be spliced into this cell's base text, by
        # position: each witness's page markers, the footnote reference, and
        # any line break the layout puts inside this cell.
        inserts = [(pm.base_at, 0, pm.text) for pm in cell.page_marks]
        cell_start = char_pos
        while bi < len(breaks) and breaks[bi][0] <= cell_start + len(seg):
            at, kind = breaks[bi]
            # A break landing exactly on a cell boundary belongs to whichever
            # cell reaches it first, at offset 0 of the next one — dropping it
            # for falling outside both is how whole padas ran together.
            # kind 2 sorts after a footnote at the same spot, so the note
            # stays on its word and the break follows the shad.
            # -1 so a break precedes a page marker standing at the same
            # offset: the marker opens the new line rather than trailing the
            # old one. A footnote always sits earlier than its pada's break,
            # so it is never in contention here.
            inserts.append((max(0, at - cell_start), -1, kind))
            bi += 1

        if place_note and seg:
            # Put the reference mark right after the annotated word, before any
            # trailing space or shad, so it renders as "su² gyur" not
            # "su ²gyur" and "grag go²/" not "grag go/²".
            #
            # Exception: when the lemma *is* a shad (only possible with shad
            # differences reported), the note is about that shad, so the mark
            # must stay on it rather than jump back to the preceding word.
            lemma_is_shad = _is_shad_only(cell.lemma)
            cut = len(seg)
            if not lemma_is_shad:
                while cut > 0 and (seg[cut - 1].isspace() or seg[cut - 1] in SHAD_CHARS
                                   or seg[cut - 1] in PUNCT_TO_IGNORE_BASE):
                    cut -= 1
            else:
                while cut > 0 and seg[cut - 1].isspace():
                    cut -= 1
            if cut == 0:  # nothing to anchor to — keep the original placement
                cut = len(seg.rstrip())

            # One cell can hold several words while the note is about one of
            # them: "'di//'on te mi " carries a note on "'on te", and a marker
            # at the end of the cell attached it to "mi". Anchor to the end of
            # the lemma instead. An omission has an empty lemma and lemma_end
            # then points at the gap the missing words would occupy.
            spans = _syllable_spans(seg, ignore_shad)
            if 0 < cell.lemma_end <= len(spans):
                cut = spans[cell.lemma_end - 1][2]
            elif cell.lemma_end == 0 and spans:
                cut = spans[0][1]
            inserts.append((cut, 1, None))
        elif place_note:
            inserts.append((0, 1, None))

        # A page marker at the same spot as a footnote goes first, so the
        # marker opens the page and the note stays attached to its word.
        inserts.sort(key=lambda x: (x[0], x[1]))
        pos = 0
        for at, kind, payload in inserts:
            at = max(pos, min(at, len(seg)))
            if at > pos:
                emit(p_text, seg[pos:at])
                pos = at
            if kind == 0:
                mark_run = p_text.add_run(payload)
                mark_run.italic = True
            elif kind == 1:
                p_text.add_footnote(notes[note_index - 1])
            else:
                start_line(payload, cell_start + at)
        if pos < len(seg):
            emit(p_text, seg[pos:])

    # Any milestone past the last emitted character (e.g. a tag at the very
    # end of the base text) still needs to be written out.
    while ms_index < len(milestones):
        tag_run = p_text.add_run(milestones[ms_index][1])
        tag_run.italic = True
        ms_index += 1

    # Footnote styles are only present after add_footnote() has run, so apply
    # formatting now (before save) rather than on the empty document.
    ensure_footnote_reference_style(doc)
    shrink_footnote_style(doc, font_size=8, line_spacing_multiple=0.85)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# Install the default preprocessing sets (must run after every dependent
# constant above is defined; the UI re-applies with the user's choices).
apply_preprocessing_options()


# ─────────────────────────────────────────────
#  STREAMLIT UI
# ─────────────────────────────────────────────

st.set_page_config(
    page_title="Tibetan Collation Tool",
    page_icon="📜",
    layout="centered",
)

st.title("📜 Tibetan Collation Tool")
st.caption(
    "Collate two to six Tibetan witnesses into a critical apparatus. "
    "You get a collation report and a golden text with footnotes, in Word."
)

st.divider()

# ── Step 1: where the texts come from
# This reshapes the whole form, so it is asked before anything else and on its
# own — it used to sit inside a step headed "Upload texts", which presumed the
# answer.
st.subheader("1 · Where the texts come from")

source_mode = st.radio(
    "Source",
    label_visibility="collapsed",
    options=["Upload .txt files", "Google Doc links", "Upload a collation report"],
    index=0,
    horizontal=True,
    help="A Google Doc must be shared “Anyone with the link → Viewer”: the "
    "app fetches it from a server and cannot use your Google account. Paste "
    "the URL straight from the address bar — if the document has tabs, the "
    "?tab=… part selects the one you are looking at.",
)
use_links = source_mode.startswith("Google")
use_report = source_mode.startswith("Upload a collation")

def _page_marker_controls(slot, who):
    """The page-marker question for one witness, used by every source mode."""
    example = None
    if st.checkbox(
        "This text has page markers",
        value=False,
        key=f"pghas{slot}",
        help="Tick this if the text marks its pages, e.g. [V1.1v.1] or "
        "[V1.272]. They are removed before collation, shown in the golden "
        "document where each page begins, and kept in each column of the "
        "report.",
    ):
        example = st.text_input(
            "Optional: add example of first page marker",
            value="",
            key=f"pgx{slot}",
            placeholder="[V1.1v.1]",
            help="Leave blank for the usual bracketed style. Otherwise paste "
            "the first marker from this text and the pattern is worked out "
            "from it.",
        )
    return example


report_labels, report_texts = [], []
base_file = None
base_link = ""
comp_files, comp_links, comp_labels = [], [], []
page_examples = []

st.divider()
st.subheader("2 · The witnesses")

if use_report:
    # The report can arrive either way, because the correcting happens
    # wherever the correctors are: a .docx passed around, or one Google Doc
    # several people work in at once. It is fetched as .docx rather than as
    # text either way — the report is a table, and exported as text its
    # columns collapse into one stream with no telling the witnesses apart.
    report_via = st.radio(
        "Where the report is",
        options=["Upload the file", "Google Doc link"],
        horizontal=True,
        key="repvia",
        label_visibility="collapsed",
    )
    report_file = report_link = None
    if report_via.startswith("Upload"):
        report_file = st.file_uploader(
            "Collation report (.docx)",
            type=["docx"],
            key="repfile",
            help="A report this tool produced. Its columns are read back as "
            "the witnesses — correct the OCR in them and the collation is "
            "redone from your corrections.",
        )
    else:
        report_link = st.text_input(
            "Google Doc link to the collation report",
            key="replnk",
            placeholder="https://docs.google.com/document/d/…/edit",
            help="The report opened in Google Docs, shared “Anyone with the "
            "link → Viewer”. Its table is read back as the witnesses, so "
            "corrections made in the columns are what gets collated.",
        )

    _report_raw = None
    if report_file is not None:
        _report_raw = report_file.getvalue()
    elif report_link:
        try:
            _report_raw, _rep_name = fetch_google_doc(report_link, fmt="docx")
            st.success(f"Fetched from Google Docs: {_rep_name}")
        except ValueError as _exc:
            st.error(str(_exc))

    if _report_raw is not None:
        try:
            report_labels, report_texts = parse_collation_report(_report_raw)
        except ValueError as _exc:
            st.error(str(_exc))
            report_labels, report_texts = [], []

    if report_labels:
        label1 = report_labels[0]
        comp_labels = list(report_labels[1:])
        n_comp = len(comp_labels)
        st.caption(
            f"Read **{len(report_labels)}** witnesses from the report. "
            "The sigla and the base come from its header row."
        )
        _cols = st.columns(min(len(report_labels), 3))
        for _i, _lb in enumerate(report_labels):
            with _cols[_i % len(_cols)]:
                with st.container(border=True):
                    st.markdown(
                        f"**{_lb}**"
                        + ("  ·  base" if _i == 0 else "")
                    )
                    page_examples.append(_page_marker_controls(_i, _lb))
    else:
        label1 = "V1"
        n_comp = 0
else:
    # Base and comparisons are the same kind of thing — witnesses — so they
    # live in one step, each in its own bordered card. Loose stacked controls
    # became unreadable at five comparisons.
    with st.container(border=True):
        st.markdown("**Base**  ·  the apparatus is anchored here")
        if use_links:
            base_link = st.text_input(
                "Google Doc link",
                key="lnk0",
                placeholder="https://docs.google.com/document/d/…/edit?tab=t.…",
            )
        else:
            base_file = st.file_uploader(
                "Base text (.txt)",
                type=["txt"],
            )
        label1 = st.text_input("Siglum", value="V1", key="lab0")
        page_examples.append(_page_marker_controls(0, label1))

    n_comp = int(
        st.columns([1, 2])[0].number_input(
            "Comparison texts",
            min_value=1,
            max_value=5,
            value=2,
            step=1,
            help="Up to five witnesses can be collated against the base.",
        )
    )

    # Three to a row so five stay readable.
    _cols = st.columns(min(n_comp, 3))
    for _i in range(n_comp):
        with _cols[_i % len(_cols)]:
            with st.container(border=True):
                _default_siglum = f"V{_i + 2}"
                st.markdown(f"**Comparison {_i + 1}**")
                if use_links:
                    comp_files.append(None)
                    comp_links.append(
                        st.text_input(
                            "Google Doc link",
                            key=f"lnk{_i + 1}",
                            placeholder="https://docs.google.com/document/d/…",
                        )
                    )
                else:
                    comp_links.append("")
                    comp_files.append(
                        st.file_uploader(
                            "Text (.txt)", type=["txt"], key=f"c{_i + 1}"
                        )
                    )
                comp_labels.append(
                    st.text_input(
                        "Siglum", value=_default_siglum, key=f"lab{_i + 1}"
                    )
                )
                page_examples.append(
                    _page_marker_controls(_i + 1, _default_siglum)
                )

st.divider()
st.subheader("3 · How to collate")

# Ordered by what each option affects. The two editorial decisions — what the
# notes look like, and what counts as a difference — stay visible; how the
# files are read is plumbing and folds away, with a count so the state is
# still legible while closed.
apparatus_mode = st.radio(
    "Apparatus",
    options=["Negative (only variants)", "Positive (all witnesses)"],
    index=1,
    captions=[
        "Lists only the witnesses that differ from the base.",
        "Lists every witness at each variant, including those that agree.",
    ],
)
positive = apparatus_mode.startswith("Positive")

prep_stack = st.checkbox(
    "Treat + as a stacking mark — seng+ge matches seng ge",
    value=False,
    key="prep_stack",
    help="EWTS + forces a stacked consonant, so seng+ge and seng ge are the "
    "same word written two ways — one after the Sanskrit, one the naturalised "
    "Tibetan spelling. Off by default, so the difference is reported and you "
    "can see it; tick it to treat the two as one reading. Either way + is "
    "never removed from a reading, so a real variant still prints as d+hi, "
    "and the orthographic profile counts what ticking this hides.",
)

ignore_shad = st.checkbox(
    "Ignore shad (།) differences",
    value=True,
    key="ignoreshad",
    help="When checked, differences that consist only of shad punctuation "
    "(།, ༎, ༔ …) are not reported as variant notes. Uncheck to have shad "
    "differences show up in the apparatus.",
)

want_verse = st.checkbox(
    "Lay out verse as verse",
    value=False,
    key="wantverse",
    help="In the golden text, set each pāda of verse on its own line in "
    "stanzas of four, and break prose where a sentence actually ends. "
    "Pādas are found by metre — the syllable count the text itself uses — "
    "with the double shad marking where one closes, so a transcription that "
    "does not distinguish it from the single shad has nothing to find and "
    "the document says so. Nothing is rewritten: with this on the golden "
    "text differs from your file in whitespace only.",
)

want_profile = st.checkbox(
    "Add an orthographic profile",
    value=False,
    key="wantprofile",
    help="Two tables describing the witnesses rather than the text: they open "
    "the footnote document on a page of their own, and close the report. The "
    "first counts, per witness, the features normalised before comparison — "
    "stacked consonants, head marks, pipes written for shad, explicit spaces, "
    "shad — which otherwise leave no trace. The second takes every stacked "
    "form and shows what the other witnesses wrote at that point, so you can "
    "tell a spelling habit from a real reading.",
)

# The expander label reports how many cleanup options are on. Their values
# come from session_state because the checkboxes live inside the expander and
# so have not been drawn yet at this point; before the first interaction
# session_state is empty and the defaults (all on) stand.
_PREP_KEYS = ("prep_tags", "prep_keep", "prep_us", "prep_pipe", "prep_head")
_prep_on = sum(bool(st.session_state.get(k, True)) for k in _PREP_KEYS)

with st.expander(
    f"Reading the files — folio tags, _, |, head marks   ·   "
    f"{_prep_on} of {len(_PREP_KEYS)} on",
    expanded=False,
):
    st.caption(
        "Cleanup applied before collation. Each option only takes effect if "
        "its pattern actually appears in your texts — otherwise it changes "
        "nothing. Leaving them all on is safe."
    )
    prep_tags = st.checkbox(
        "Strip folio/page tags like [354], [zhe 1]",
        value=True,
        key="prep_tags",
        help="Bracketed reference markers are removed before alignment so "
        "they don't show up as spurious variants. Page markers such as "
        "[V1.1v.1] are handled separately and are not affected.",
    )
    prep_keep_tags = st.checkbox(
        "↳ …but keep the base text's tags in the golden output as milestones",
        value=True,
        key="prep_keep",
        disabled=not prep_tags,
        help="The stripped tags from the base/golden text are re-inserted "
        "into the downloaded golden document (in italics, at their original "
        "positions) as page references. They are never part of the "
        "collation itself. Tags from the comparison texts are not kept — "
        "the golden document reproduces only the base text.",
    )
    prep_underscore = st.checkbox(
        "Treat _ as a space (EWTS explicit space)",
        value=True,
        key="prep_us",
        help="In EWTS transliteration an underscore marks an explicit space. "
        "Without this, e.g. pa/_bdag and pa/ bdag read as different words.",
    )
    prep_pipe = st.checkbox(
        "Treat | as a shad (EWTS / OCR)",
        value=True,
        key="prep_pipe",
        help="Some OCR output writes the shad as a pipe. With this on, | "
        "behaves exactly like / — ignored or reported together with shad.",
    )
    prep_head = st.checkbox(
        "Ignore head marks @ # ! (yig-mgo ༄༅)",
        value=True,
        key="prep_head",
        help="These transliterate the ornamental head marks that open a "
        "section; they are structural, not textual, so they never count as "
        "variants.",
    )

st.divider()
st.subheader("4 · Run")

if use_report:
    ready = bool(report_texts)
elif use_links:
    ready = bool(base_link.strip()) and all(l.strip() for l in comp_links)
else:
    ready = base_file is not None and all(f is not None for f in comp_files)

if not ready:
    # A caption, not an alert: the disabled button already says it cannot run,
    # so this only needs to say what is missing.
    st.caption(
        "Add a collation report above to enable this."
        if use_report else
        "Paste a link for every witness above to enable this."
        if use_links else
        "Upload a file for every witness above to enable this."
    )

run_btn = st.button("▶ Run Collation", disabled=not ready, type="primary")

if run_btn and ready:
    uploads = [base_file] + comp_files
    links = [base_link] + comp_links
    labels = [label1] + comp_labels

    # .getvalue() (not .read()) — Streamlit keeps the uploaded file's read
    # cursor across reruns, so a second Run would read empty bytes and
    # silently reuse the previous results. Google Docs are fetched here, in
    # the Run branch, so editing a widget does not re-download anything.
    texts, encodings, names = [], [], []
    _fetched = []
    for _i in range(len(labels)):
        if use_report:
            # already read out of the report when it was uploaded
            _t, _e, _name = report_texts[_i], "UTF-8", labels[_i]
        elif use_links:
            with st.spinner(f"Fetching {labels[_i]} from Google Docs…"):
                try:
                    _raw, _name = fetch_google_doc(links[_i])
                except ValueError as _exc:
                    st.error(f"**{labels[_i]}** — {_exc}")
                    st.stop()
            _t, _e = decode_upload(_raw)
            # Google appends footnote bodies after a rule of underscores;
            # left in they collate as a trailing addition.
            _t = strip_gdoc_footnotes(_t)
            _fetched.append((labels[_i], _name, len(_t)))
        else:
            _raw = uploads[_i].getvalue()
            _name = uploads[_i].name
            _t, _e = decode_upload(_raw)
        names.append(_name)
        texts.append(_t)
        encodings.append(_e)

    if _fetched:
        st.success("Fetched from Google Docs:")
        for _lb, _nm, _n in _fetched:
            st.markdown(f"- **{_lb}** — {_nm} · {_n:,} characters")

    # A non-UTF-8 file is read on a best guess, so say so: a wrong guess
    # surfaces as odd characters in the notes rather than as an error.
    for _nm, _e in zip(names, encodings):
        if _e != "UTF-8":
            st.warning(
                f"**{_nm}** is not UTF-8 — read as **{_e}**. The "
                "collation will run; check the output for odd characters, "
                "and re-save the file as UTF-8 to remove any doubt."
            )

    # The archive belongs to the report round trip: having corrected the
    # witnesses inside a report, you need them back as standalone texts. When
    # they came from files or links the originals are already in hand, so
    # offering it there only invites the question of what it is for. Built
    # before any preprocessing, so it holds each witness exactly as it
    # stands, pagination included.
    versions_buf = None
    if use_report:
        versions_buf = export_versions_document(
            list(texts),
            labels,
            patterns=[
                (pattern_from_example(_ex) if _ex else DEFAULT_PAGE_MARKER_RE)
                if _ex is not None else None
                for _ex in (page_examples + [None] * len(labels))[: len(labels)]
            ],
        )

    # Apply the user's preprocessing choices
    apply_preprocessing_options(
        underscore_as_space=prep_underscore,
        pipe_as_shad=prep_pipe,
        ignore_head_marks=prep_head,
        stack_mark_as_same=prep_stack,
    )
    # counted from the texts as read, before anything is stripped out
    profile_texts = list(texts) if want_profile else None
    # Counted before cleanup is applied, so the preview reflects the input.
    prep_preview = {nm: count_preprocessing_hits(t) for nm, t in zip(names, texts)}

    # Page markers come out first. "Strip folio/page tags" removes anything
    # bracketed, so running it first would delete [V1.1v.1] before this ever
    # saw it. Their positions are kept so the golden document can show where
    # each witness turned its page.
    page_markers = []
    for _i, _t in enumerate(texts):
        _ex = page_examples[_i] if _i < len(page_examples) else None
        if _ex is None:
            page_markers.append([])
            continue
        texts[_i], _mk = extract_page_markers(_t, _ex)
        page_markers.append(_mk)
        if not _mk:
            _what = f"the example `{_ex.strip()}`" if _ex.strip() else "the default [V1.1v.1] style"
            st.warning(
                f"**{names[_i]}** — no page markers matched {_what}. That "
                "file will contribute no page references; check the example "
                "matches how its pages are actually marked."
            )

    golden_milestones = None
    if prep_tags:
        if prep_keep_tags:
            texts[0], golden_milestones = extract_folio_tags(texts[0])
        else:
            texts[0] = strip_folio_tags(texts[0])
        for _i in range(1, len(texts)):
            texts[_i] = strip_folio_tags(texts[_i])

    # A witness left unchecked keeps whatever is in its text, and anything
    # left there is collated as a reading. Tibetan — in script or in Wylie —
    # carries no digits of its own, so a digit surviving folio-tag stripping
    # is almost certainly pagination that is about to be read as a variant.
    # Cheap to detect and worth saying loudly, because the result is a note
    # like "V1 om.] V2 4r7 (pdf 55)" rather than an error.
    for _i, _t in enumerate(texts):
        if (page_examples[_i] if _i < len(page_examples) else None) is not None:
            continue
        _hits = _DIGIT_TOKEN_RE.findall(_t)
        if _hits:
            _eg = ", ".join(f"`{h}`" for h in dict.fromkeys(_hits[:3]))
            st.warning(
                f"**{names[_i]}** — “This text has page markers” is not "
                f"ticked for this file, but its text still contains numbers "
                f"({_eg}). Tibetan text has no digits of its own, so these "
                "are probably page markers — and they will be collated as "
                "readings, producing spurious variants. Tick that box for "
                "this file and give an example so they are removed."
            )

    with st.expander("Preprocessing preview", expanded=False):
        st.caption(
            "Occurrences of each preprocessing pattern found per file "
            "(counted before cleanup was applied)."
        )
        for fname, counts in prep_preview.items():
            hits = ", ".join(f"{k}: {v}" for k, v in counts.items() if v) or "nothing to clean"
            st.markdown(f"- **{fname}** — {hits}")
        if any(page_markers):
            st.caption("Page markers found (first and last shown):")
            for _nm, _mk in zip(names, page_markers):
                if _mk:
                    st.markdown(
                        f"- **{_nm}** — {len(_mk)} markers, "
                        f"`{_mk[0][1]}` … `{_mk[-1][1]}`"
                    )
                else:
                    st.markdown(f"- **{_nm}** — none")

    with st.spinner("Aligning texts… this may take a minute for long texts."):
        aligned = align_witnesses(texts)

    with st.spinner("Building collation report…"):
        cells = collate_cells(
            aligned, labels, ignore_shad=ignore_shad, positive=positive,
            markers=page_markers,
        )
        report_buf, notes = export_collation_report(
            cells, labels, names, profile_texts=profile_texts
        )

    footnote_buf = None
    try:
        with st.spinner("Building footnote document…"):
            footnote_buf = export_golden_with_footnotes(
                cells, notes, labels, name1=names[0],
                milestones=golden_milestones, ignore_shad=ignore_shad,
                profile_texts=profile_texts, verse_layout=want_verse,
            )
        if want_verse:
            _blocks = verse_blocks("".join(c.segs[0] for c in cells))
            if _blocks:
                st.caption(
                    "Verse: %d passages, %d pādas, %s syllables."
                    % (len(_blocks), sum(len(b.lines) for b in _blocks),
                       " and ".join(str(m) for m in sorted({b.metre for b in _blocks})))
                )
            else:
                st.warning(
                    "⚠️ No verse found, so the golden text is set as prose. "
                    "Pāda boundaries are located by the double shad — if this "
                    "text writes every shad as a single one, there is nothing "
                    "to find."
                )
    except AttributeError:
        st.warning(
            "⚠️ Footnote document skipped — `add_footnote` not available. "
            "Install **bayoo-docx** or **python-docx-2023** instead of python-docx."
        )

    # Store results in session_state so downloads persist after button clicks
    st.session_state["versions_buf"] = (
        versions_buf.getvalue() if versions_buf else None
    )
    st.session_state["report_buf"] = report_buf.getvalue()
    st.session_state["footnote_buf"] = footnote_buf.getvalue() if footnote_buf else None
    st.session_state["note_count"] = len(notes)

# Show download section whenever results are available in session_state
if "report_buf" in st.session_state:
    st.success(f"Done! Found **{st.session_state['note_count']}** difference note(s).")
    st.divider()
    st.subheader("5 · Download outputs")

    _DOCX_MIME = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    _has_versions = bool(st.session_state.get("versions_buf"))
    _dl = st.columns(3 if _has_versions else 2)
    dl1, dl2 = _dl[0], _dl[1]
    dl3 = _dl[2] if _has_versions else None
    with dl1:
        st.download_button(
            label="⬇ Collation report (.docx)",
            data=st.session_state["report_buf"],
            file_name="collation_report.docx",
            mime=_DOCX_MIME,
            key="dl_report",
        )
    with dl2:
        if st.session_state["footnote_buf"]:
            st.download_button(
                label="⬇ Golden text + footnotes (.docx)",
                data=st.session_state["footnote_buf"],
                file_name="golden_text_footnotes.docx",
                mime=_DOCX_MIME,
                key="dl_footnotes",
            )
        else:
            st.button("⬇ Golden text + footnotes (.docx)", disabled=True)
    if dl3 is not None:
        with dl3:
            st.download_button(
                label="⬇ All source versions (.docx)",
                data=st.session_state["versions_buf"],
                file_name="all_source_versions.docx",
                mime=_DOCX_MIME,
                key="dl_versions",
                help="Each witness in full, in sequence, with its page "
                "markers and nothing else — the corrected texts, ready to "
                "keep.",
            )

st.divider()
REPO_URL = "https://github.com/jyerena108/tibetan-collation"

with st.expander("ℹ️ About this tool"):
    st.markdown(f"""
**Tibetan Collation Tool** — aligns two or three versions of a Tibetan text
and generates a critical apparatus in Word format. Works with Unicode Tibetan
and Wylie/EWTS transliteration.

Notes read `V1 kyi] V2, V3 ni` — *where V1 reads `kyi`, V2 and V3 read
`ni`*. Omissions are marked `om.`

**How to cite**

> Yerena, J. *Tibetan Collation Tool* (2026). {REPO_URL}

If the alignment matters to your argument, please also cite Pydurma:

> Roux, E. & Kaldan, T. *Pydurma* (2023). https://github.com/openpecha/pydurma

**Run your own copy**

```bash
git clone {REPO_URL}.git
cd tibetan-collation
pip install -r requirements.txt
streamlit run collation_app_01.py
```

Full documentation is in the [README]({REPO_URL}#readme).
""")

st.caption(
    f"**Tibetan Collation Tool** by Yerena, J. — free and open source "
    f"([MIT]({REPO_URL}/blob/main/LICENSE)) · [source code]({REPO_URL})  \n"
    "Alignment by [Pydurma](https://github.com/openpecha/pydurma) "
    "(OpenPecha, MIT).  \n"
    "Provided as-is, without warranty. Please check collation results before "
    "relying on them in published work."
)
