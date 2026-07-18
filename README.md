# Tibetan Collation Tool

A web app for collating Tibetan texts and producing a critical apparatus in
Word format. Upload two or three versions of a text, and the tool aligns them,
finds the differences, and generates the notes for you.

It works with both Unicode Tibetan (བོད་ཡིག) and Wylie/EWTS transliteration.

> **Live app:** _add your Streamlit URL here_

---

## What it produces

Two Word documents:

| File | Contents |
|---|---|
| **Collation report** | A numbered list of every variant found |
| **Golden text + footnotes** | Your base text with real Word footnotes at each variant |

Notes follow standard critical-edition style, with the siglum before the
reading on both sides of the bracket:

```
BX1 kyi] AB1, GB1 ni
```

Read as: *where BX1 reads `kyi`, witnesses AB1 and GB1 read `ni`.* Witnesses
sharing a reading are grouped with commas; different readings are separated by
`;`. An omission is marked `om.`. The lemma carries its own siglum, so it can
be moved into the variant list unchanged if you later reassign the base text.

Only the differing syllable is shown, not the whole phrase — if two witnesses
agree on the surrounding words, those are trimmed away.

---

## Using the app

### 1 · Upload texts

Upload a **base (golden) text** — this is the version the apparatus is
anchored to — and one or two **comparison texts**. Plain `.txt`, UTF-8.

Give each one a **siglum** (short label like `BX1`, `AB1`, `GB1`). These are
what appear in the notes.

### 2 · Options

**Apparatus type**

- **Negative** (default) — lists only the witnesses that *differ* from the base
- **Positive** — lists *every* witness at each variant point, including those
  that agree

**Ignore shad (།) differences**

On by default. Differences consisting only of shad punctuation (`།`, `༎`, `༔`,
and the Wylie `/`) are not reported. Turn it off to have shad variants appear
in the apparatus.

**Preprocessing**

Cleanup applied before collation. Each option only takes effect if its pattern
actually appears in your files, so leaving them all checked is safe.

| Option | What it does |
|---|---|
| Strip folio/page tags | Removes `[354]`, `[zhe 1]` and similar reference markers so they don't collate as text |
| ↳ keep as milestones | Puts the base text's tags back into the golden document, in italics, at their original positions — without them ever entering the collation |
| Treat `_` as a space | EWTS writes an explicit space as `_`; without this, `pa/_bdag` and `pa/ bdag` read as different words |
| Treat `\|` as a shad | Some OCR output writes the shad as a pipe |
| Ignore head marks `@ # !` | These transliterate the yig-mgo ornaments (༄༅) that open a section — structural, not textual |

After running, a **Preprocessing preview** shows how many of each pattern were
found per file, so you can see exactly what was cleaned.

### 3 · Run and download

Click **Run Collation**, then download either output. Changing any option and
running again re-collates with the new settings.

---

## Notes on Tibetan handling

- **Scripts** — syllables are rejoined with a tsheg (`་`) for Unicode Tibetan
  and with a space for Wylie, so notes never mix scripts.
- **A-chung** — a lone a-chung (`འ` / `'`) stranded at an alignment boundary is
  reattached to the syllable it belongs to, so an added a-chung reads
  `gyur] 'gyur` rather than surfacing as a bare `'`.
- **Footnote marks** are placed on the annotated word, before any trailing
  shad or space (`grag go¹²/`). When a note is *about* a shad, the mark stays
  on the shad instead.
- **Your text is never rewritten.** The golden output reproduces the base
  witness exactly as written, apart from folio tags you chose to strip.
  Preprocessing affects how texts are *compared*, not what they say.

---

## Running it locally

```bash
git clone https://github.com/jyerena108/tibetan-collation.git
cd tibetan-collation
pip install -r requirements.txt
streamlit run collation_app_01.py
```

A browser tab opens at `http://localhost:8501`.

Requires Python 3.11. The footnote output needs `bayoo-docx` (included in
`requirements.txt`), which adds the `add_footnote()` method that plain
`python-docx` lacks — without it the collation report still works, but the
footnote document is skipped.

### Deploying your own copy

Push to a GitHub repo and connect it at [share.streamlit.io](https://share.streamlit.io).
No server needed. Streamlit redeploys automatically on every push.

---

## Credits

Built by **Yerena, J.**

Alignment is powered by **[Pydurma](https://github.com/openpecha/pydurma)** by
Elie Roux and Tenzin Kaldan ([OpenPecha](https://openpecha.org)), MIT licensed
and vendored in `Pydurma/`.

Thanks to the translators who tested the tool and shaped its output format
through their feedback.

## License

MIT — see [LICENSE](LICENSE). Pydurma is separately MIT licensed; see
`Pydurma/LICENSE`.
