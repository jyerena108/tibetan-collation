# Tibetan Collation Tool

A web app for collating Tibetan texts and producing a critical apparatus in
Word format. Give it a base text and up to five comparison witnesses; it aligns
them, finds the differences, and writes the notes.

It works with both Unicode Tibetan (བོད་ཡིག) and Wylie/EWTS transliteration.

> **Live app:** <https://tibetan-collation-001.streamlit.app/>

---

## What it produces

| File | Contents |
|---|---|
| **Collation report** | Every witness side by side, plus the numbered notes. A3 landscape, so six columns stay readable |
| **Golden text + footnotes** | Your base text with real Word footnotes at each variant, opening with the orthographic profile when you ask for one |
| **All source versions** | Each witness in full, in sequence — offered only when the input was a collation report (see [the round trip](#the-round-trip)) |

Notes follow standard critical-edition style, with the siglum before the
reading on both sides of the bracket:

```
V1 kyi] V2, V3 ni
```

Read as: *where V1 reads `kyi`, witnesses V2 and V3 read `ni`.* Witnesses
sharing a reading are grouped with commas; different readings are separated by
`;`. An omission is marked `om.`. The lemma carries its own siglum, so it can
be moved into the variant list unchanged if you later reassign the base text.

Only the differing syllables are shown, not the whole phrase — anything every
witness agrees on is trimmed away.

---

## Using the app

### 1 · Where the texts come from

One choice, which reshapes the rest of the form.

| Source | Notes |
|---|---|
| **Upload .txt files** | Plain text, any common encoding |
| **Google Doc links** | Paste the URL from the address bar |
| **Upload a collation report** | A report this tool made, after you have corrected it — as a `.docx` or as a Google Doc link |

**Google Docs** are read through the plain-text export, so no API key is
needed — but the document must be shared **Anyone with the link → Viewer**,
because the app fetches it from a server and cannot use your Google account. A
private document returns a sign-in page, and the tool says so rather than
collating it. If the document has tabs, the `?tab=…` in the URL selects the one
you are looking at; without it you would get every tab concatenated. Footnote
bodies, which Google appends after a rule of underscores, are dropped.

**A collation report is fetched as a `.docx`, not as text**, because it is a
table: exported as text its columns collapse into a single stream and there is
no telling the witnesses apart again. Everything else is the same — the same
sharing requirement, the same errors.

**Encoding** is worked out for you. Files are read as UTF-8, and when that
fails — PDF extractors and older Mac tools often write Mac OS Roman — the text
is decoded on a best guess rather than the run dying, with a notice naming the
encoding used. Windows and classic-Mac line endings are normalised.

### 2 · The witnesses

The **base** comes first — the version the apparatus is anchored to — then one
to five **comparisons**, each in its own card.

Give each a **siglum**: a short label like `V1`, `V2`. These are what appear in
the notes. Defaults are `V1` for the base and `V2`…`V6` for the comparisons;
rename them to whatever your edition uses.

When the source is a collation report, this step reads the witnesses, their
sigla, and which one is the base straight out of the report instead.

**Page / folio markers.** Each witness has its own **This text has page
markers** box. Tick it and you can give an example:

- Leave the box **blank** for the house style — a bracketed tag containing a
  dot, like `[V1.1v.1]` (siglum, folio, side, line) or the shorter `[V1.272]`.
  A plain folio tag such as `[354]` is *not* a page marker; that belongs to the
  folio-tag option below.
- Or paste the first marker from that file — `Pdf.50`, `p.292`,
  `kha, 1r.1 (pdf 47)` — and the pattern is worked out from it. Later markers
  are found even when the case changes (`Pdf.294` then `pdf.295`), the volume
  changes (`ka` then `kha`), or the side changes (`4r` then `4v`). A trailing
  parenthetical is optional, so a sample carrying `(pdf 47)` still matches
  markers without one.

Markers are removed before collation — left in, they align as readings and
invent variants — and then **woven back into both output documents**:

- in the **golden text**, every witness's markers appear inline at the point
  where *that* witness turns its page, so one reading text shows where all of
  them stood:

  ```
  [BX1.1v.1][DX1.145r.1][AB1.272] skyes bu dam pa rnams ni/rmad du byung
  ba'i cha lugs … de nas de nyid [DX1.146r.1]gnyid sad do//…
  ```

- in the **report**, each column carries only its own witness's markers.

Every version keeps the format its own source uses; nothing is converted. The
notes themselves carry no page references — repeating one on every variant made
them longer without making them more useful.

If a file has markers but the box is left unticked, the tool warns: Tibetan
carries no digits of its own, so numbers surviving in the text are a reliable
sign of pagination about to be read as a variant.

### 3 · How to collate

**Apparatus**

- **Positive** (default) — lists every witness at each variant, including those
  that agree with the base
- **Negative** — lists only the witnesses that *differ*

**Treat `+` as a stacking mark — `seng+ge` matches `seng ge`** — off by
default, so the difference is reported and you can see it. Tick it to read the
two spellings as one word. See [below](#orthography-versus-text).

**Ignore shad (།) differences** — on by default. Differences consisting only of
shad punctuation (`།`, `༎`, `༔`, and the Wylie `/`) are not reported. Turn it
off to have them appear.

**Add an orthographic profile** — off by default. Two tables that describe the
witnesses rather than the text: they open the footnote document on a page of
their own, and close the report.

The first counts the features normalised before comparison, which otherwise
leave no trace:

```
                             BX1     AB1     DX1     GX1
Stacked consonants  +          3      16       2       4
Head marks  @ # !              2       1       0       0
Pipe written for shad  |       0       0       0       7
Explicit space  _              5       0       0       0
Shad  /                      317     292     286     279
Pāda-final shads omitted       2       1       0       1
```

Read as a profile that says something: AB1 stacks sixteen forms where the base
stacks three, and BX1 runs two pairs of verse lines together where DX1 runs
none.

**That last row counts something absent**, which is why nothing else surfaces
it. Tibetan verse is isosyllabic — every pāda of a passage carries the same
syllable count, classically 7, 9 or 11 — and a shad closes each one. When a
witness omits that shad, two pādas run together as one long line. The words are
all present and in order, so the apparatus sees nothing wrong; only the metre
reveals it.

**Every omission is also listed**, in a third table, with what each witness
writes at the same place — because a count cannot say whether it is a reading
or a slip of the OCR, and those call for opposite responses:

| Pāda | Runs into | BX1 | AB1 | DX1 | GX1 |
|---|---|---|---|---|---|
| …brtson 'grus brtan pa'i mthus | 'jig rten mchog gi spyod pa bstan… | **om.** `[BX1.3r.1]` | `/ /` | `// /` | `/ /` |
| dam pa'i spyod las gzhan du byas | dud 'gro'i gzugs can bdag nyid che | **om.** `[BX1.3v.1]` | `/ /` | `//` | **om.** `[GX1.160v.1]` |
| …ngo mtshar skye 'gyur | skyes bu dam pa de yi mthu yis byas | — | **om.** `[AB1.273]` | `//` | — |

**Read across, and the witnesses answer for each other.** In the first row
three of them write a shad, so the boundary is certainly a boundary and BX1 is
the one that dropped it — with a folio reference saying where to correct it in
[the round trip](#the-round-trip). In the second, two omit and two do not: a
shared reading rather than a slip. A `—` means that witness does not have the
passage, which is not an omission and is not counted as one.

Every omission carries the folio it falls on in **that** witness, so there is
somewhere to go and look. Each witness's shad is shown **as it writes it**, so
`/ /` against `//` also tells you how the transcriptions differ.

Page/folio tags are removed before any of this is measured — left in they are
read as a syllable and throw the metre off for every pāda they touch — and
remembering where they stood is what lets an omission be cited by folio. They
are untouched in the output documents.

The count is `–`, not `0`, when no verse could be found at all. That happens
when a transcription never distinguishes the double shad from the single: verse
lines end in a double 87–93% of the time across these witnesses and prose
segments only 27–52%, so without that distinction there is nothing to measure,
and a zero there would claim the witness omits nothing.

**The table follows the script your texts are in.** Wylie witnesses are
labelled in Wylie, exactly as above. Unicode Tibetan witnesses are labelled in
Tibetan — `Head marks  ༄༅`, `Shad  །` — and gain a row for the non-breaking
tsheg `༌`, which is normalised away silently otherwise. Rows that cannot exist
in a script are not listed for it rather than reported as zero, and a run
mixing the two marks those cells `–` and adds a **Script** row saying which
witness is which.

Explicit stacking is the one feature with no Unicode counterpart. EWTS `+`
records a choice the transcriber made; Tibetan script does not distinguish
stacked from unstacked, so there is nothing there to count. The subjoined
consonants are not a substitute — they spell every ordinary syllable, so
counting them would measure how much Tibetan is present, not how the witness
handled its Sanskrit. A Unicode run therefore has no stacking row and no
stacked-forms table.

The second takes every stacked form and shows what the other witnesses wrote at
that point, so you can tell a spelling habit from a real reading:

| BX1 | AB1 | DX1 | GX1 | × |
|---|---|---|---|---|
| seng ge'i | **seng+ge'i** | seng ge'i | seng ge'i | 7 |
| **d+hi** | de | de | de | 2 |
| seng ge | **seng+ge** | seng ge | seng ge | 2 |
| pad ma'i | **pad+ma'i** | **pad+ma'i** | **pad+ma'i** | 1 |

`seng+ge` against `seng ge` is AB1 spelling the same word another way. `d+hi`
against `de` is a different reading, and stays in the apparatus however you set
the stacking option.

**Reading the files** — collapsed, and safe to leave alone. Each option only
takes effect if its pattern actually appears in your texts.

| Option | What it does |
|---|---|
| Strip folio/page tags | Removes `[354]`, `[zhe 1]` and similar reference markers so they don't collate as text |
| ↳ keep as milestones | Puts the base text's tags back into the golden document, in italics, at their original positions — without them ever entering the collation |
| Treat `_` as a space | EWTS writes an explicit space as `_`; without this, `pa/_bdag` and `pa/ bdag` read as different words |
| Treat `\|` as a shad | Some OCR output writes the shad as a pipe |
| Ignore head marks `@ # !` | These transliterate the yig-mgo ornaments (༄༅) that open a section — structural, not textual |

After running, a **Preprocessing preview** reports how many of each pattern were
found per file.

### 4 · Run, 5 · Download

Click **Run Collation**, then take the documents you need. Change any option and
run again to re-collate.

---

## The round trip

The report carries every witness in full, which makes it an editable container
for all of them at once. That supports a workflow the tool was built around:

1. Produce your first OCR versions and revise them by hand
2. Collate → report and footnote file
3. **Correct the OCR inside the report's columns** — this gives you version 2
4. Feed the corrected report back in → final apparatus, plus an archive of the
   corrected texts

Step 3 can happen in Word or in Google Docs, and step 4 takes the report either
way — upload the file, or paste a link to it. The Google route is the one to
use when the correcting is shared: several people in one document, or one
person moving between machines.

Step 4 re-collates from scratch, because once the witness texts change the old
notes are stale by definition. It reads page markers and note references by
shape rather than by their italic and superscript formatting, since formatting
is the first thing to smear when a long table is edited by hand in Word.

---

## Orthography versus text

Some differences are graphic rather than textual, and repeating them through an
apparatus buries the readings that matter. The tool normalises these for
*matching* only — they are never removed from a reading, so a genuine variant
still prints what the witness actually wrote. Apostrophes are folded always;
stacking is the option described above.

- **Stacked consonants.** EWTS `+` forces a stack, so `seng+ge` and `seng ge`
  are the same word written two ways — one after the Sanskrit original, one the
  naturalised Tibetan spelling. **Reported by default**, so you see the
  difference; tick **Treat `+` as a stacking mark** in section 3 to read the two
  as one word. Either way `+` is never removed from a reading, so a real variant
  still reads `V1 d+hi] V2, V3, V4 de`.
- **Typographic apostrophes.** Word processors and OCR turn the Wylie a-chung
  into a curly quote (`’` or `‘`). These fold onto the ASCII apostrophe for
  comparison *and* in the notes, so `'gyur` and `’gyur` are not reported as a
  variant and a note reads `'das`, not `‘das`.

The orthographic profile counts both features, whether or not they are
being ignored on a given run.

---

## Notes on Tibetan handling

- **Scripts** — syllables are rejoined with a tsheg (`་`) for Unicode Tibetan
  and with a space for Wylie, so notes never mix scripts.
- **Syllables stay whole.** The aligner breaks at every apostrophe, so `pa'i`
  arrives as `pa` + `'i` and `mnga'` as `mnga` + `'`. Each is one syllable — a
  stem with its suffix inside a single tsheg unit — so they are rejoined before
  the notes are built. Without this the apparatus cited fragments: `spyad pa]
  sbyang ba` where the text reads `spyad pa'i`.
- **A-chung** — an initial a-chung stranded at an alignment boundary moves
  forward to the syllable it belongs to, so an added one reads `gyur] 'gyur`
  rather than surfacing as a bare `'`. A final one moves back, so `mda'` is
  never cited as `mda`.
- **Footnote marks sit on the lemma** — the last word the note is about, not
  the end of whatever the aligner happened to group together. Where the base
  omits something, the mark goes at the gap, between the words the missing
  ones would have sat between.
- **Your text is never rewritten.** The golden document reproduces the base
  witness exactly as written, apart from the page markers and folio tags it
  re-inserts, and line endings, which are normalised when the file is read.
  Everything else affects how texts are *compared*, not what they say.

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
