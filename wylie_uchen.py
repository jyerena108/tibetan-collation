"""Wylie -> Uchen conversion, apparatus intact.

Implements SPEC-wylie-to-unicode.md's MVP: convert a .docx's running Wylie
text and footnotes to Uchen, leaving page markers, footnote numbering,
sigla, and technical terms (``om.``, ``em.``, ...) untouched.

Deliberately does not reimplement transliteration or document I/O:
``pyewts`` does the Wylie -> Unicode conversion (see the spec's note on why),
and ``collation_app_01.decode_upload`` / ``fetch_google_doc`` are reused
as-is for getting bytes in.

The hard part is not the conversion — it's rewriting *part* of a run's text
while leaving the rest, without disturbing anything else in the .docx. This
module edits ``word/document.xml`` and ``word/footnotes.xml`` directly:
a run that needs no change is left completely untouched (byte-identical),
and a run that needs partial conversion is split into several runs, each
cloning the original's formatting, so nothing about the surrounding
document is ever rewritten.
"""

import io
import re
import zipfile
import xml.etree.ElementTree as ET

import pyewts

_converter = pyewts.pyewts()


def to_uchen(wylie: str) -> str:
    """Convert one piece of Wylie text to Uchen."""
    return _converter.toUnicode(wylie)


# ── docx XML plumbing ───────────────────────────────────────────────
_W_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W = "{%s}" % _W_URI
_WT, _WR, _WP, _WRPR = _W + "t", _W + "r", _W + "p", _W + "rPr"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# Root elements declare a long list of namespaces most of which nothing in
# the document actually uses. ElementTree only re-emits the ones it still
# sees referenced, which would silently drop the rest on write-back —
# harmless in practice, but registering every prefix the file declared keeps
# the round trip as close to identity as possible for anything we do not
# touch.
_XMLNS_RE = re.compile(r'xmlns:([A-Za-z0-9]+)="([^"]+)"')


def _register_namespaces(raw_xml: bytes) -> None:
    head = raw_xml[:4000].decode("utf-8", "ignore")
    for prefix, uri in _XMLNS_RE.findall(head):
        ET.register_namespace(prefix, uri)


def _clone_run_with_text(template_run, text):
    """A new <w:r>, formatted like ``template_run``, holding ``text``."""
    new_run = ET.Element(_WR)
    rpr = template_run.find(_WRPR)
    if rpr is not None:
        new_run.append(_deepcopy(rpr))
    t = ET.SubElement(new_run, _WT)
    t.text = text
    t.set(_XML_SPACE, "preserve")
    return new_run


def _deepcopy(elem):
    new = ET.Element(elem.tag, dict(elem.attrib))
    new.text, new.tail = elem.text, elem.tail
    for child in elem:
        new.append(_deepcopy(child))
    return new


def _replace_run(parent, old_run, new_runs):
    children = list(parent)
    idx = children.index(old_run)
    parent.remove(old_run)
    for offset, nr in enumerate(new_runs):
        parent.insert(idx + offset, nr)


def _group_tokens(tokens, is_exception):
    """Group ``(token, is_space)`` pairs into ``(text, convert)`` spans.

    A run of Wylie words is kept as *one* span even across the ordinary
    space(s) between them, rather than converting word by word: pyewts turns
    an inter-word space in its input into the tsheg itself, so feeding it
    one word at a time and rejoining with the original literal space would
    silently drop every tsheg. A space is only ever folded into a span this
    way when both its neighbours are convertible — a space touching an
    exception (a siglum, `]`, a page marker, ...) stays exactly as written,
    since nothing there should turn into a tsheg.
    """
    kinds = []  # "space" | "protect" | "convert", parallel to tokens
    for tok in tokens:
        if tok.isspace():
            kinds.append("space")
        elif is_exception(tok):
            kinds.append("protect")
        else:
            kinds.append("convert")

    spans, i, n = [], 0, len(tokens)
    while i < n:
        if kinds[i] != "convert":
            spans.append((tokens[i], False))
            i += 1
            continue
        j, buf = i, tokens[i]
        while j + 2 < n and kinds[j + 1] == "space" and kinds[j + 2] == "convert":
            buf += tokens[j + 1] + tokens[j + 2]
            j += 2
        spans.append((buf, True))
        i = j + 1
    return spans


def _apply_tokenizer(run, parent, tokenizer_re, is_exception):
    """Split ``run``'s text by ``tokenizer_re`` and convert non-exceptions.

    ``is_exception(token)`` decides, per non-space token, whether it is left
    alone. A run with nothing to convert is never touched, so untouched text
    stays byte-identical to the source. Returns the ``(node, convert)`` pairs
    now occupying ``run``'s old slot — a single ``(run, False)`` when nothing
    changed — for ``_bridge_tsheg`` to reason about afterwards.
    """
    t = run.find(_WT)
    if t is None or not t.text:
        return [(run, False)]
    tokens = [m.group(0) for m in tokenizer_re.finditer(t.text)]
    if not tokens:
        return [(run, False)]
    spans = [(to_uchen(text) if convert else text, convert)
             for text, convert in _group_tokens(tokens, is_exception)]
    if len(spans) == 1 and not spans[0][1]:
        return [(run, False)]  # nothing converted — left exactly as it was
    new_runs = [_clone_run_with_text(run, text) for text, _ in spans]
    _replace_run(parent, run, new_runs)
    return [(node, convert) for node, (_, convert) in zip(new_runs, spans)]


# A footnote reference is its own <w:r>, with no <w:t> at all, sitting right
# at the word boundary it marks — so the two Wylie words either side of it
# end up in different runs, each converted on its own by _apply_tokenizer
# above. That is fine for the words themselves, but the ordinary space
# between them, now stranded as the leading/trailing whitespace of one of
# those runs, no longer has a convertible neighbour *in its own run* to be
# folded into — so it survives conversion as a literal space instead of
# becoming the tsheg it would have if the reference weren't there.
#
# This second, paragraph-wide pass fixes exactly that: a run holding nothing
# but whitespace, with a converted run on each side once reference runs
# (which carry no text and so cannot end a word or start one) are looked
# through, is the tsheg the reference marker stood in the way of.
_TSHEG = "་"


def _paragraph_entries(p, tokenizer_re, is_exception):
    """``(node, kind)`` for every run in ``p``, after converting each.

    ``kind`` is ``"convert"``, ``"protect"``, or ``"milestone"`` (a run with
    no text — a footnote reference — transparent to the tsheg-bridging pass).
    """
    entries = []
    for run in list(p):
        if run.tag != _WR:
            continue
        t = run.find(_WT)
        if t is None or not t.text:
            entries.append((run, "milestone"))
            continue
        for node, convert in _apply_tokenizer(run, p, tokenizer_re, is_exception):
            entries.append((node, "convert" if convert else "protect"))
    return entries


def _nearest_content_kind(entries, start, step):
    i = start
    while 0 <= i < len(entries):
        kind = entries[i][1]
        if kind != "milestone":
            return kind
        i += step
    return None


def _bridge_tsheg(entries) -> None:
    """Turn a whitespace-only run between two converted runs into a tsheg."""
    for i, (node, kind) in enumerate(entries):
        if kind != "protect":
            continue
        t = node.find(_WT)
        if t is None or not t.text or not t.text.isspace():
            continue
        before = _nearest_content_kind(entries, i - 1, -1)
        after = _nearest_content_kind(entries, i + 1, 1)
        if before == "convert" and after == "convert":
            t.text = _TSHEG


# ── Main text: only page markers (brackets) are left alone ────────────
# Footnote *reference* numbers are a <w:footnoteReference/> element, not
# text, so there is nothing for a text-level tokenizer to protect there.
_MAIN_TOKEN_RE = re.compile(r"\s+|\[[^\[\]\n]*\]|[\[\]]|[^\s\[\]]+")


def _is_bracketed(tok: str) -> bool:
    return tok.startswith("[") or tok == "]"


# A golden-text-with-footnotes document (this tool's own export) often opens
# with a title and analysis sections ("Orthographic profile", "Stacked forms
# across the witnesses", ...) before the running text itself — present or
# not, unpredictably, per document. Those sections are English prose, not
# Wylie, and running the converter over them mangles them into nonsense.
#
# The running text itself always opens with a page marker giving every
# witness's starting folio, e.g. "[BX1, 1v.1; AB1, 272; DX1, 145r.1; GX1,
# 158v.1]" — so that marker is the boundary: everything from the first
# paragraph that contains a bracketed tag onward is converted, and whatever
# precedes it (if anything does) is left alone. A plain document with no such
# preamble has its first bracket, if any, in its very first paragraph, so
# this falls back to converting from the start.
_PAGE_MARKER_RE = re.compile(r"\[[^\[\]\n]*\]")


def _paragraph_text(p):
    return "".join(t.text or "" for t in p.iter(_WT))


def _find_golden_text_start(doc_root):
    body = doc_root.find(_W + "body")
    if body is None:
        return list(doc_root.iter(_WP))
    paras = list(body.findall(_WP))
    for i, p in enumerate(paras):
        if _PAGE_MARKER_RE.search(_paragraph_text(p)):
            return paras[i:]
    return paras


def _convert_body(doc_root) -> None:
    for p in _find_golden_text_start(doc_root):
        entries = _paragraph_entries(p, _MAIN_TOKEN_RE, _is_bracketed)
        _bridge_tsheg(entries)


# ── Footnotes: sigla, technical terms, and structural punctuation stay ──
_NOTE_TOKEN_RE = re.compile(r"\s+|[\]\,;]|[^\s\]\,;]+")


def _convert_footnotes(fn_root, sigla, terms, skip_ids) -> None:
    sigla, terms = set(sigla), set(terms)

    def is_exception(tok):
        return tok in ("]", ",", ";") or tok in sigla or tok in terms

    for note in fn_root.findall(_W + "footnote"):
        if note.get(_W + "id") in skip_ids:
            continue
        for p in note.iter(_WP):
            entries = _paragraph_entries(p, _NOTE_TOKEN_RE, is_exception)
            _bridge_tsheg(entries)


# ── First pass: find the sigla and technical terms ─────────────────────
# A note this tool writes is `<sigla> <reading>] <sigla> <reading>; ...`
# (build_note_text in collation_app_01.py) — a siglum is whatever labels a
# reading, and in every label this tool has produced or been given, that is
# an uppercase/digit token with no lowercase letter, because the reading
# beside it is lowercase Wylie. Technical terms (`om.`, `em.`, ...) are the
# one place apparatus text carries a bare period; Wylie has no use for one.
_SIGLUM_RE = re.compile(r"\b[A-Z][A-Z0-9]*\b")
_TERM_RE = re.compile(r"\b[a-z]+\.")

# Structural characters a well-formed note is built from. Anything outside
# this set — a question mark, curly quotes, an en dash — is the editor's own
# aside riding along in the footnote text, not apparatus grammar, and the
# note is set aside for a human rather than guessed at.
_NOTE_SAFE_RE = re.compile(
    r"^[A-Za-z0-9\s\[\]\(\),;/+'‘’.–-]*$"
)


def detect_exceptions(note_texts):
    """First pass over a batch of footnote texts.

    ``note_texts`` is ``{note_id: text}``. Returns ``(sigla, terms,
    flagged)``: ``sigla``/``terms`` are sorted lists of the tokens found in
    notes that look well-formed; ``flagged`` is the list of note ids that
    don't reduce to the sigla/reading shape and so are left untouched by
    conversion, for a human to look at.
    """
    sigla, terms, flagged = set(), set(), []
    for note_id, text in note_texts.items():
        if _NOTE_SAFE_RE.match(text):
            sigla.update(_SIGLUM_RE.findall(text))
            terms.update(_TERM_RE.findall(text))
        else:
            flagged.append(note_id)
    return sorted(sigla), sorted(terms), flagged


def read_footnote_texts(raw: bytes):
    """``{note_id: text}`` for every real footnote in a .docx's archive.

    Ids ``-1`` and ``0`` are Word's separator/continuation-separator
    placeholders, never real notes, and are skipped.
    """
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        if "word/footnotes.xml" not in z.namelist():
            return {}
        fn_root = ET.fromstring(z.read("word/footnotes.xml"))
    out = {}
    for note in fn_root.findall(_W + "footnote"):
        note_id = note.get(_W + "id")
        if note_id in ("-1", "0"):
            continue
        text = "".join(t.text or "" for t in note.iter(_WT))
        out[note_id] = text
    return out


# ── Putting it together ────────────────────────────────────────────────
def convert_docx_bytes(raw: bytes, sigla, terms, flagged_ids=()) -> bytes:
    """Convert a .docx's Wylie to Uchen, apparatus intact.

    ``sigla``/``terms`` are the user-confirmed exception lists. Notes whose
    id is in ``flagged_ids`` are left completely untouched. Every part of
    the .docx archive other than document.xml and footnotes.xml is copied
    over unchanged.
    """
    flagged_ids = set(flagged_ids)
    zin = zipfile.ZipFile(io.BytesIO(raw))

    doc_xml = zin.read("word/document.xml")
    _register_namespaces(doc_xml)
    doc_root = ET.fromstring(doc_xml)
    _convert_body(doc_root)

    fn_root = None
    if "word/footnotes.xml" in zin.namelist():
        fn_xml = zin.read("word/footnotes.xml")
        _register_namespaces(fn_xml)
        fn_root = ET.fromstring(fn_xml)
        _convert_footnotes(fn_root, sigla, terms, flagged_ids)

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == "word/document.xml":
                zout.writestr(item, ET.tostring(doc_root, encoding="UTF-8"))
            elif item.filename == "word/footnotes.xml" and fn_root is not None:
                zout.writestr(item, ET.tostring(fn_root, encoding="UTF-8"))
            else:
                zout.writestr(item, zin.read(item.filename))
    return out.getvalue()
