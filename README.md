# odf-layout-diff

Prototype. Compares where text lands in two PDF renderings of the same
document and reports the differences as a short, ranked list of symptoms.

The intended use is a style-correction loop for documents that LibreOffice
lays out differently from the application that produced them:

1. **Reference**: a PDF with the intended layout, for example the original
   `.docx` exported by Microsoft Word.
2. **Candidate**: the same document converted to ODF and rendered by
   LibreOffice.
3. This tool reports what moved and, where it can tell, why.
4. A language model (a small local one is the target) picks a fix from a
   constrained set of style operations, applied through
   [odf-rs](../odf-rs)'s `odf-tool protocol`.
5. Render again and repeat until the report is empty.

The point of step 3 is that the model should not have to infer
"this line wrapped because the text is slightly wider" from raw geometry or
from megabytes of style data. The tool does that mechanically and hands the
model a few kilobytes.

odf-rs itself does no layout, so this lives outside it.

## Requirements

- Python 3.11 or later, standard library only
- `pdftotext` from poppler-utils

## Usage

```bash
python3 layout_diff.py reference.pdf candidate.pdf \
    --projection project.json > report.json
```

`--projection` is optional and repeatable: an `odf-tool protocol` response
to a `project` request on the candidate `.odt` (one file per window). With
it, each symptom names the projection nodes (`address`, `styleName`) whose
text matches, which is what a `set_style_properties` operation needs.

`--threshold` (default 3.0 pt) is the smallest vertical shift reported;
`--limit` (default 10) caps the number of symptoms.

### Documents exported by Microsoft Word

`python3 experiments/prepare_word_odt.py word.odt prepared.odt` prepares an
`.odt` that Word exported before it is compared or fixed (needs `soffice`):

- It fills in the text grid's base height (the line pitch: text area height
  over lines per page) and ruby height (0), and turns a lines and
  characters grid without a character pitch into a lines-only grid, as
  LibreOffice's DOCX import does. Word leaves both heights out; LibreOffice
  then lays every line out at least 30 pt high where Word's pitch is
  typically about 18 pt, and the document gains pages.
- It drops a fixed pitch from font faces with a non-ASCII family name.
  With a UI in another language, LibreOffice finds such a name only
  through fontconfig, where a fixed pitch puts monospace fonts first.
- It saves the result again with LibreOffice: Word's export does not
  satisfy the ODF schema, and `odf-tool` refuses to edit it.

Run the tests with:

```bash
python3 -m unittest discover -s tests -t . -v
```

## Report

```json
{
  "reportVersion": 2,
  "pages": {"reference": 2, "candidate": 2},
  "matchedGlyphRatio": 0.912,
  "rowsOnAnotherPage": [...],
  "symptoms": [...],
  "omittedSymptoms": 0
}
```

- `matchedGlyphRatio`: the share of reference characters aligned with the
  candidate. A low value means the two PDFs do not hold the same text (or
  extract it in a very different order), and the rest of the report should
  not be trusted.
- `rowsOnAnotherPage`: reference rows that landed on a different candidate
  page, grouped by page pair.
- `symptoms`, ranked by the number of points they move the layout:
  - `text-block-taller` / `text-block-shorter`: a text block
    (`pdftotext`'s block) gained or lost lines. `lineChange` is the net
    number of added line breaks, `heightChangePt` the change in its
    vertical extent, `breakCount` the number of breaks that do not cancel
    out and `breaks` the first three of them.
  - `vertical-shift`: everything from `row` down moved by `shiftPt`
    relative to `rowAbove`, and no reported block change explains it.
    `breaksInRowAbove` lists line breaks the candidate added inside
    `rowAbove` across block boundaries (for example, text pushed to the
    right of a cell by spaces). When it finds none, `extraRows` lists the
    texts of candidate rows that appeared between the two rows without
    holding text of another reference row, and `hint` is `wrapped-row` if
    each of them is text of `rowAbove` (typically its last character wrapped
    on its own, too short to align) or `unknown` otherwise.
  - `shifted-line-start`: a line starts `shiftPt` points right (positive)
    or left (negative) of the reference while staying on the row its
    reference row landed on, so nothing wraps differently: an indent or
    margin moved it, for example out of its table cell. Reported from 6 pt
    on (about a cell's padding). Lines of one block moved by the same
    amount form one symptom (`lineCount`). `hint` is `shifted-start`, or
    `wider-text` / `narrower-text` when the end of the line (including the
    later segments of its paragraph on that row, known from `--projection`
    excerpts) stayed within 3 pt: right-aligned text whose width changed,
    which no indent fixes. A line start already given as a break's
    `startShiftPt`, a later segment of the same paragraph (see `wide-gap`),
    centred text whose two ends moved apart and two copies of the same text
    on one row that moved by opposite amounts (each aligned with the other
    copy because the cells were extracted in another order) are left out.
- With `--projection`, every symptom kind carries `nodes`: for block
  changes and shifts with breaks, matched from the block text and the text
  around the breaks; for other shifts, from the lines of `rowAbove`; for
  moved lines, from their block and line starts.
- Each break has `textBeforeBreak`/`textAfterBreak` (normalized: NFKC, no
  whitespace), `startShiftPt` (how far right the candidate starts the text
  before the break; `null` unless that text begins a line, or a segment of
  one, in both renderings), `gapPt` (the horizontal gap between the two sides where
  they share a line), `widthRatio` (candidate width over reference width of
  the text on both sides), and a `hint`, the first that applies:
  - `shifted-start`: the text before the break starts at least 3 pt away
    from the reference position. Indentation (including font-relative
    units such as LibreOffice's `loext:text-indent="18ic"`), alignment or
    cell padding are the usual causes.
  - `wider-text`: the candidate draws the same characters wider
    (`widthRatio` above 1.005). Letter spacing, character scaling, font
    size or font substitution are the usual causes.
  - `wide-gap`: the text is not wider, but the break falls at a gap of at
    least 20 pt, typically a run of spaces used for alignment. Also used
    when the text before the break continues, after a gap of at least
    20 pt, a paragraph whose earlier text sits left of it on the same row
    (known from `--projection` excerpts) and starts at least 3 pt further
    right: the spaces before it got wider, which is not an indent, so no
    `startShiftPt` is given. Without a projection such text is treated as
    a line start.
  - `unknown`: neither applies.

## How it works

1. `pdftotext -bbox-layout` gives pages, blocks, lines and word boxes.
   `pdftotext` splits one visual line where the font or baseline changes
   (for example before a full-width bracket); lines of one block on the
   same row that overlap or touch are joined back. Every word is split
   into normalized characters with an interpolated x position.
2. The two character streams are aligned with `difflib`, with a
   separator between text blocks that never matches, so a match cannot run
   from one block into the next. Matching blocks shorter than three
   characters are dropped. The two PDFs may extract text in a different
   order (a wrapped cell's text can move ahead of a neighbouring column), so
   a second pass pairs every unmatched run of at least four characters that
   occurs exactly once among the unmatched candidate characters.
3. Lines are grouped into visual rows by vertical centre.
4. Within each reference block, a line break is *added* where two
   consecutive characters share a reference row but the second starts a new,
   lower row further left in the candidate, and *removed* the other way
   round. A block is reported only when the added and removed breaks do not
   cancel out, so reflowed paragraphs and side-by-side cells that merely
   sit at different heights stay out of the report.
5. Each reference row's shift is the median offset of its characters on the
   candidate row of its leftmost run. A shift is reported when it exceeds
   the threshold, persists into the next row, and is not directly below a
   reported block. Rows repeated on several pages (running headers and
   footers) are ignored.
6. A line start (the first character of a `pdftotext` line in both
   renderings) whose candidate character sits on the row its reference row
   landed on is compared horizontally.

## Limitations

- Text only: ruling lines, cell borders and images are not compared. Text
  moved out of its cell is found only as a line start that moved relative
  to the reference, not against the border itself.
- Both PDFs must contain the same text. A filled-in form compared with a
  blank one will report the filled-in values as differences.
- The reference renderer matters. LibreOffice-based renderers (Collabora)
  share the candidate's layout engine and are not a useful reference.
- Line breaks before text too short to align (for example a lone "日"
  wrapped onto the next line) are not found as breaks; the resulting shift
  is reported with the added row in `extraRows` instead.
