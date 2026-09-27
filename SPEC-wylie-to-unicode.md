# Wylie → Uchen converter, with the apparatus intact

A second tool for this app, not yet built. Written down from Dominik Dell's
specification of 2026-09 so that whoever builds it starts from the
requirement rather than from a summary of it.

## Why

Critical editions are mostly made in Wylie. Converting one to Uchen lets a
Tibetan reader see it — most do not read Wylie, and would rather not.
Converters exist on the web, but they take a plain text and know nothing of
footnotes. **The apparatus is the whole difficulty**: it carries signs that
must survive untouched — sigla, square brackets, `om.`, `em.` — sitting in
the middle of text that must be converted.

## Input and output

A Word file or Google Doc holding Wylie with footnotes, returned as the same
document in Uchen with its footnotes converted.

The existing tool already reads both — `decode_upload` and
`fetch_google_doc` — and already writes `.docx` with real footnotes. That
machinery is the reason this belongs in the same app.

## Conversion rules

Convert all the Wylie, in the main text **and** in the footnotes. The
exceptions define what is not Wylie.

**In the main text, leave alone:**

- everything inside square brackets (page markers)
- the footnote reference numbers

**In the footnotes, leave alone:**

- the footnote numbering
- `]`, commas, spaces
- the **sigla**
- the **technical terms** — `om.`, `em.`, and others

On the last two Dominik suggests a text box each: the user lists their sigla
and their terms. Worth weighing against reading them from the note's own
shape, since this tool writes notes as `BX1 kyi] AB1, GB1 ni` — the siglum
always stands at the head of a reading, and a parser for that already exists
in `parse_collation_report`. Decide this early: it shapes the whole
exceptions design.

That much is the MVP.

## Then: the table case

The Jātaka project keeps the Wylie in **column 1 of a table**, with stages of
translation in the columns beside it. The tool should reproduce that
structure.

Two options, user's choice:

1. replace the Wylie in column 1 with Uchen
2. keep column 1 and insert the Uchen as a new column 2

Procedure, as specified:

- copy the original file and work on the copy
- leave everything outside the table untouched
- leave every column but the first untouched
- read column 1 row by row; convert each row, footnotes included, by exactly
  the MVP logic
- replace column 1, or insert the new column 2, per the option chosen
- continue to the end of the table

## Later

Uchen → Wylie, the other direction.

---

## Notes for whoever builds it

**Do not write the transliteration.** `pyewts` is the standard EWTS ↔ Unicode
library, from the same OpenPecha work as the Pydurma vendored here. An EWTS
converter written from scratch is hundreds of orthographic rules and a long
tail of edge cases. Add the dependency, pin it as the others are pinned
(resolve against Python 3.11 — see `requirements.txt` for why), and spend the
effort on the exceptions, which is where the real problem is.

**It belongs as a second Streamlit page in this repo**, not a separate app:
one deploy, one pinned dependency set, one address to embed, and the `.docx`
and Google Doc code shared rather than copied and left to drift.

**Worth reusing as they stand:** `decode_upload`, `fetch_google_doc`,
`parse_collation_report` (for the note shape), and the run-walking in
`export_golden_with_footnotes`, which already splices content into a
paragraph without disturbing what is around it.

**The trap to expect.** Converting a run of text is easy; converting *part* of
a run while leaving the rest is where `.docx` work goes wrong. A footnote
like `BX1 kyi] AB1, GB1 ni` has to come out with `BX1`, `]`, `,` and `AB1`
untouched and only `kyi` and `ni` converted — inside one paragraph, possibly
inside one run. Build that case first and the rest follows.
