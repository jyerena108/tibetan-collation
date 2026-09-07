import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "Pydurma", "src"))
"""
Tibetan Collation Web App — Streamlit
======================================
Run with:
    pip install streamlit bayoo-docx
    streamlit run collation_app.py
"""

import io
import re
import tempfile
from collections import namedtuple
from pathlib import Path

import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_LINE_SPACING
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
):
    """Install the effective character sets used by the collation.

    - underscore_as_space: EWTS writes an explicit space as ``_``; treat it
      as whitespace everywhere (ignore set, syllable splitting, a-chung
      reattachment).
    - pipe_as_shad: ``|`` is an alternate EWTS shad (common in OCR output).
    - ignore_head_marks: ``@``, ``#``, ``!`` transliterate yig-mgo ornaments
      (༄༅ …), which are structural, not textual.

    Rebinding the module-level sets keeps every existing function signature
    unchanged; the display/golden text itself is never rewritten by these.
    """
    global SHAD_CHARS, PUNCT_TO_IGNORE_BASE, PUNCT_TO_IGNORE
    global _SYLLABLE_SEP_RE, _SEP_CHARS
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

def set_landscape(document):
    """Turn the document landscape.

    python-docx does not swap the page dimensions when the orientation is
    changed, so width and height have to be exchanged by hand \u2014 otherwise Word
    still lays the page out portrait and the setting appears to do nothing.
    """
    for section in document.sections:
        w, h = section.page_width, section.page_height
        if w < h:
            section.orientation = WD_ORIENT.LANDSCAPE
            section.page_width, section.page_height = h, w


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
    ignore_set = PUNCT_TO_IGNORE if ignore_shad else PUNCT_TO_IGNORE_BASE
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
A_CHUNG_CHARS = set(["འ", "'"])

# Syllable separators for display splitting: tsheg, no-break tsheg, whitespace.
_SYLLABLE_SEP_RE = re.compile(r"[་༌\s]+")


def _syllables(seg: str, ignore_shad: bool):
    """Split an aligned segment into display syllables.

    Tsheg and whitespace act as separators (and are dropped). Shad is removed
    only when shad differences are ignored, so it stays visible otherwise.
    Letters (and any residual marks) are preserved inside each syllable. A bare
    a-chung is re-attached to its neighbour rather than kept as its own token.
    """
    if not seg:
        return []
    s = seg
    if ignore_shad:
        for ch in SHAD_CHARS:
            s = s.replace(ch, "")
    # Drop ignorable non-content characters (head marks, western punctuation)
    # from the display so readings never show e.g. "@#" from a source file.
    parts = []
    for p in _SYLLABLE_SEP_RE.split(s):
        p = "".join(c for c in p if c not in PUNCT_TO_IGNORE_BASE)
        if p:
            parts.append(p)
    merged = []
    pending = ""  # a leading bare a-chung waiting to attach to the next syllable
    for p in parts:
        if all(c in A_CHUNG_CHARS for c in p):
            if merged:
                merged[-1] = merged[-1] + p
            else:
                pending += p
        else:
            merged.append(pending + p)
            pending = ""
    if pending:
        merged.append(pending)
    return merged


def _trim_common_syllables(readings):
    """Trim syllables shared by *all* readings at the start and the end.

    ``readings`` is a list of syllable-lists. Returns new syllable-lists with
    the common leading/trailing syllables removed, so only the differing part
    remains. Because at least one reading always differs when a note exists,
    this never trims a reading down to nothing on every side simultaneously.
    """
    readings = [list(r) for r in readings]
    if len(readings) < 2:
        return readings
    # common prefix
    n = min(len(r) for r in readings)
    pre = 0
    while pre < n and all(r[pre] == readings[0][pre] for r in readings):
        pre += 1
    readings = [r[pre:] for r in readings]
    # common suffix
    n = min(len(r) for r in readings)
    suf = 0
    while suf < n and all(r[-1 - suf] == readings[0][-1 - suf] for r in readings):
        suf += 1
    if suf:
        readings = [r[: len(r) - suf] for r in readings]
    return readings


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
    ``(note_text, lemma)``; note_text is "" when no witness differs.

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
    """
    # comparison keys (punctuation/tsheg-insensitive) decide agreement
    keys = [strip_ignorable(s, ignore_shad) for s in segs]
    # display syllables, with syllables shared by every witness trimmed away
    trimmed = _trim_common_syllables([_syllables(s, ignore_shad) for s in segs])

    lemma = _reading_display(trimmed[0])
    base_key = keys[0]

    lemma_labels = [labels[0]]
    if positive:
        lemma_labels += [labels[i] for i in range(1, len(segs)) if keys[i] == base_key]

    selected = [
        (labels[i], trimmed[i]) for i in range(1, len(segs)) if keys[i] != base_key
    ]
    if not selected:
        return "", lemma

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
    return f"{', '.join(lemma_labels)} {lemma}] " + "; ".join(parts), lemma


# One record per aligned cell, shared by both exporters.
CollatedCell = namedtuple(
    "CollatedCell", "segs norms diffs base_missing has_diff note lemma"
)


def collate_cells(aligned, labels, ignore_shad=True, positive=False):
    """Walk the aligned rows once and decide everything about each cell.

    The report and the golden document used to run separate copies of this
    logic, each calling build_note_text() itself. The copies had to be kept
    identical by hand: if they ever disagreed about whether a cell yields a
    note, every footnote from that point on attached to the wrong word while
    still looking correctly numbered. Deciding once, here, removes that whole
    failure mode — both exporters consume this list.

    ``aligned`` is base-first, one row per witness; "-" marks an aligner gap.
    """
    n = len(aligned)
    width = max((len(r) for r in aligned), default=0)
    rows = [list(r) + [""] * (width - len(r)) for r in aligned]

    cells = []
    for j in range(width):
        raw = [rows[i][j] for i in range(n)]
        segs = ["" if c == "-" else c for c in raw]
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

        note, lemma = "", ""
        if has_diff:
            note, lemma = build_note_text(
                segs, labels, positive=positive, ignore_shad=ignore_shad
            )

        cells.append(
            CollatedCell(segs, norms, diffs, base_missing, has_diff, note, lemma)
        )
    return cells


def export_collation_report(cells, labels, names):
    """Side-by-side table of every witness, plus the numbered notes list.

    The page is landscape: with up to six witness columns a portrait page
    squeezes Tibetan script too narrow to read comfortably.
    """
    doc = Document()
    set_landscape(doc)
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

        seg = cell.segs[0]
        if seg:
            run = paras[0].add_run(seg)
            if color and any(cell.diffs) and cell.norms[0] != "":
                set_run_background_color(run, color)
            if note_start_here:
                m = paras[0].add_run(f"[{note_number}]")
                m.font.superscript = True
        elif note_start_here:
            m = paras[0].add_run(f"[{note_number}]")
            m.font.superscript = True

        for i in range(1, n):
            seg_i = cell.segs[i]
            if not seg_i:
                continue
            differs = cell.diffs[i - 1] or cell.base_missing
            run = paras[i].add_run(seg_i)
            if color and differs and cell.norms[i] != "":
                set_run_background_color(run, color)
            if note_start_here and differs:
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


def export_golden_with_footnotes(cells, notes, labels, name1="base", milestones=None):
    """Golden text with variant footnotes.

    ``cells`` is the very list the report was built from, so footnote numbering
    matches the report's notes by construction rather than by keeping a second
    copy of the diff logic in step with the first.

    ``milestones`` is an optional list of ``(offset, tag)`` pairs from
    extract_folio_tags(): folio/page tags stripped before collation that are
    re-inserted here — as italic runs at their original character positions —
    without ever having been part of the alignment or the apparatus.
    """
    milestones = sorted(milestones) if milestones else []
    ms_index = 0  # next milestone still to be emitted
    char_pos = 0  # running offset into the (concatenated) base text

    def emit(p, text):
        """Write base text, splicing in any milestone tags it spans."""
        nonlocal ms_index, char_pos
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
    doc.add_heading(f"{labels[0]} with Footnotes", level=1)
    if len(labels) > 2:
        others = ", ".join(labels[1:-1]) + " and " + labels[-1]
    else:
        others = labels[1]
    doc.add_paragraph(f"Base: {name1}  |  Footnotes from comparison with {others}.")
    doc.add_paragraph()

    p_text = doc.add_paragraph()

    note_index = 0
    for cell in cells:
        seg = cell.segs[0]
        note_start_here = bool(cell.note)
        if note_start_here:
            note_index += 1
        place_note = note_start_here and 1 <= note_index <= len(notes)

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
            content = seg[:cut]
            trailing = seg[cut:]
            if content:
                emit(p_text, content)
            p_text.add_footnote(notes[note_index - 1])
            if trailing:
                emit(p_text, trailing)
        else:
            if seg:
                emit(p_text, seg)
            if place_note:
                p_text.add_footnote(notes[note_index - 1])

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

# Streamlit markdown has no underline directive (bold/italics/strikethrough
# and colour only) and st.title() takes no unsafe_allow_html, so the title is
# rendered as an explicit <h1>. Only the wording is underlined, not the emoji.
st.markdown(
    "<h1>📜 <u>Tibetan Collation Tool</u></h1>",
    unsafe_allow_html=True,
)
st.caption("Upload your texts, run the collation, and download the Word outputs.")

st.divider()

# ── File uploads
st.subheader("1 · Upload texts")

col_a, col_b = st.columns([1, 1])
with col_a:
    base_file = st.file_uploader(
        "Base / golden text (.txt)",
        type=["txt"],
        help="This is the primary version — all notes are anchored here.",
    )
    label1 = st.text_input("Label for base text", value="V1")

with col_b:
    n_comp = int(
        st.number_input(
            "Number of comparison texts",
            min_value=1,
            max_value=5,
            value=2,
            step=1,
            help="Up to five witnesses can be collated against the base.",
        )
    )

st.divider()
st.subheader("2 · Comparison text(s)")

# Laid out three to a row so five uploaders stay readable.
comp_files = []
comp_labels = []
_cols = st.columns(min(n_comp, 3))
for _i in range(n_comp):
    with _cols[_i % len(_cols)]:
        comp_files.append(
            st.file_uploader(
                f"Comparison text {_i + 1} (.txt)", type=["txt"], key=f"c{_i + 1}"
            )
        )
        comp_labels.append(
            st.text_input(
                f"Label for comparison {_i + 1}",
                value=f"V{_i + 2}",
                key=f"lab{_i + 1}",
            )
        )

st.divider()
st.subheader("3 · Options")

with st.expander("Preprocessing", expanded=True):
    st.caption(
        "Cleanup applied before collation. Each option only takes effect if "
        "its pattern actually appears in your files — otherwise it changes "
        "nothing. Leaving them all checked is safe."
    )
    prep_tags = st.checkbox(
        "Strip folio/page tags like [354], [zhe 1]",
        value=True,
        help="Bracketed reference markers are removed before alignment so "
        "they don't show up as spurious variants.",
    )
    prep_keep_tags = st.checkbox(
        "↳ …but keep the base text's tags in the golden output as milestones",
        value=True,
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
        help="In EWTS transliteration an underscore marks an explicit space. "
        "Without this, e.g. pa/_bdag and pa/ bdag read as different words.",
    )
    prep_pipe = st.checkbox(
        "Treat | as a shad (EWTS / OCR)",
        value=True,
        help="Some OCR output writes the shad as a pipe. With this on, | "
        "behaves exactly like / — ignored or reported together with shad.",
    )
    prep_head = st.checkbox(
        "Ignore head marks @ # ! (yig-mgo ༄༅)",
        value=True,
        help="These transliterate the ornamental head marks that open a "
        "section; they are structural, not textual, so they never count as "
        "variants.",
    )

ignore_shad = st.checkbox(
    "Ignore shad (།) differences",
    value=True,
    help="When checked, differences that consist only of shad punctuation "
    "(།, ༎, ༔ …) are not reported as variant notes. Uncheck to have shad "
    "differences show up in the apparatus.",
)

apparatus_mode = st.radio(
    "Apparatus type",
    options=["Negative (only variants)", "Positive (all witnesses)"],
    index=0,
    help="Negative apparatus lists only the witnesses that differ from the "
    "base/golden reading. Positive apparatus lists every comparison witness "
    "at each variant point, including those that agree with the lemma.",
)
positive = apparatus_mode.startswith("Positive")

st.subheader("4 · Run")

ready = base_file is not None and all(f is not None for f in comp_files)

if not ready:
    st.info("Upload all required files above to enable the collation.")

run_btn = st.button("▶ Run Collation", disabled=not ready, type="primary")

if run_btn and ready:
    uploads = [base_file] + comp_files
    labels = [label1] + comp_labels
    names = [f.name for f in uploads]

    # .getvalue() (not .read()) — Streamlit keeps the uploaded file's read
    # cursor across reruns, so a second Run would read empty bytes and
    # silently reuse the previous results.
    texts, encodings = [], []
    for f in uploads:
        _t, _e = decode_upload(f.getvalue())
        texts.append(_t)
        encodings.append(_e)

    # A non-UTF-8 file is read on a best guess, so say so: a wrong guess
    # surfaces as odd characters in the notes rather than as an error.
    for _nm, _e in zip(names, encodings):
        if _e != "UTF-8":
            st.warning(
                f"**{_nm}** is not UTF-8 — read as **{_e}**. The "
                "collation will run; check the output for odd characters, "
                "and re-save the file as UTF-8 to remove any doubt."
            )

    # Apply the user's preprocessing choices
    apply_preprocessing_options(
        underscore_as_space=prep_underscore,
        pipe_as_shad=prep_pipe,
        ignore_head_marks=prep_head,
    )
    # Counted before cleanup is applied, so the preview reflects the input.
    prep_preview = {nm: count_preprocessing_hits(t) for nm, t in zip(names, texts)}

    golden_milestones = None
    if prep_tags:
        if prep_keep_tags:
            texts[0], golden_milestones = extract_folio_tags(texts[0])
        else:
            texts[0] = strip_folio_tags(texts[0])
        for _i in range(1, len(texts)):
            texts[_i] = strip_folio_tags(texts[_i])

    with st.expander("Preprocessing preview", expanded=False):
        st.caption(
            "Occurrences of each preprocessing pattern found per file "
            "(counted before cleanup was applied)."
        )
        for fname, counts in prep_preview.items():
            hits = ", ".join(f"{k}: {v}" for k, v in counts.items() if v) or "nothing to clean"
            st.markdown(f"- **{fname}** — {hits}")

    with st.spinner("Aligning texts… this may take a minute for long texts."):
        aligned = align_witnesses(texts)

    with st.spinner("Building collation report…"):
        cells = collate_cells(
            aligned, labels, ignore_shad=ignore_shad, positive=positive
        )
        report_buf, notes = export_collation_report(cells, labels, names)

    footnote_buf = None
    try:
        with st.spinner("Building footnote document…"):
            footnote_buf = export_golden_with_footnotes(
                cells, notes, labels, name1=names[0], milestones=golden_milestones
            )
    except AttributeError:
        st.warning(
            "⚠️ Footnote document skipped — `add_footnote` not available. "
            "Install **bayoo-docx** or **python-docx-2023** instead of python-docx."
        )

    # Store results in session_state so downloads persist after button clicks
    st.session_state["report_buf"] = report_buf.getvalue()
    st.session_state["footnote_buf"] = footnote_buf.getvalue() if footnote_buf else None
    st.session_state["note_count"] = len(notes)

# Show download section whenever results are available in session_state
if "report_buf" in st.session_state:
    st.success(f"Done! Found **{st.session_state['note_count']}** difference note(s).")
    st.divider()
    st.subheader("5 · Download outputs")

    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button(
            label="⬇ Collation report (.docx)",
            data=st.session_state["report_buf"],
            file_name="collation_report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="dl_report",
        )
    with dl2:
        if st.session_state["footnote_buf"]:
            st.download_button(
                label="⬇ Golden text + footnotes (.docx)",
                data=st.session_state["footnote_buf"],
                file_name="collation_footnotes.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="dl_footnotes",
            )
        else:
            st.button("⬇ Golden text + footnotes (.docx)", disabled=True)

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
