# Task: static layout-property comparator for Word (.docx) vs ODF (.odt) tables and page setup

## Context
This is a local research task on document-format interoperability (no security, no network, no git repository).
Real government forms were authored in Microsoft Word and converted to ODF. When LibreOffice renders the
converted .odt, row heights, borders and line spacing drift from Word's layout. We want a mechanical,
per-table/per-row/per-cell comparison of the layout-relevant properties that Word stores (.docx) against
what the .odt stores, so a human can see which properties were lost or changed by conversion.

## Hard rules
- Work ONLY inside this directory: the current working directory (OUT). Write every file here.
- Inputs are read-only: a local directory of paired Word/ODF government-form samples (third-party,
  outside this repository; do not modify, move or copy them into any git repository). Do not open any
  other path except Python's standard library.
- Python 3.10+ standard library only (zipfile, xml.etree.ElementTree, json, re, argparse, unittest).
- Do not interpret OOXML defaults you are unsure about: always report the RAW attribute (or `null` when absent)
  next to any resolved value, and name the source level of every resolved value.
- Do not end your turn to wait for anything; everything here is synchronous.

## Deliverables (all in OUT)
1. `static_props.py` with subcommands:
   - `dump-docx FILE.docx` -> JSON on stdout
   - `dump-odt FILE.odt` -> JSON on stdout
   - `compare WORD.docx FILE.odt` -> JSON on stdout (diff records + category counts)
   - `compare-odt A.odt B.odt` -> JSON on stdout (same categories, ODT on both sides)
2. `test_static_props.py` (unittest): builds tiny synthetic .docx/.odt packages in memory/tempdir and checks
   unit conversions (twips, eighths of a point, cm/in/mm/pt), row-height rule mapping, cell margin and border
   resolution precedence, span/covered-cell column alignment. Run it and include the result.
3. `out/<short>.json` and `out/<short>.md` for these four pairs (short name in brackets):
   - [plan200] `1-200萬元以上設置計畫(範本)11307【檢核表加註】.docx` vs `.odt` (same stem)
   - [labor] `1-3勞、就、災、健保及勞退合一加保申報表_1150820(改後).docx` vs `1-3勞、就、災、健保及勞退合一加保申報表_1150810(改後).odt`
     (different date suffix, but text is 99% identical; a few cells differ in text)
   - [shejia] `社家署110年社福考核指標-公益彩券組.docx` vs `.odt`
   - [mohw] `衛福部-個案訪視交通補助費(按公里數核算)印領清冊.docx` vs `.odt`
4. `out/summary.md`: one table, rows = pairs, columns = category counts (see below), plus the 10 most frequent
   concrete mismatch patterns overall (e.g. "Word trHeight exact 20pt -> ODT min-row-height 20pt: 37 rows").
5. Print `ls -l`, `sha256sum` and `wc -l` of every deliverable to stdout at the end, then a final message that
   starts with `## SUMMARY` (at most 30 lines): what was built, test result, the summary table, and any
   property you could not resolve and why.

## What to extract

### Document level
- docx: every `w:sectPr` (body and paragraph-level): pgSz (w,h,orient), pgMar (top,bottom,left,right,header,footer,gutter),
  docGrid (type, linePitch, charSpace), titlePg, header/footer references; `word/settings.xml` full `w:compat`
  children and compatSetting name/val; docDefaults rPr/pPr; font usage counts of w:rFonts (ascii, hAnsi, eastAsia, cs)
  across styles and runs; theme font resolution if `asciiTheme`/`eastAsiaTheme` is used (report theme name and mapped face).
- odt: every `style:page-layout` (all attributes of page-layout-properties, header-style/footer-style
  header-footer-properties), master-page -> page-layout mapping, `style:font-face` (name, svg:font-family,
  style:font-pitch, font-family-generic), `style:default-style` per family, and all `config:config-item` names/values
  under `ooo:configuration-settings` in settings.xml.

### Tables (document order; nested tables with a path like `3/r2c1/0`)
Word side, per table: tblStyle id, tblW, tblLayout, tblInd, jc, tblCellSpacing, tblCellMar, tblBorders, tblLook,
gridCol widths; per row: trHeight (val, hRule RAW), cantSplit, tblHeader, tblPrEx overrides; per cell: grid column
start, gridSpan, vMerge (restart/continue), tcW, tcMar, tcBorders, vAlign, noWrap, textDirection, shading;
first and all paragraphs in the cell: resolved spacing (before, after, line, lineRule), snapToGrid, contextualSpacing,
jc, ind, font size (sz/szCs) and eastAsia font, each with the source level (direct / paragraph style chain /
table style / docDefaults). Resolve basedOn chains. Resolve table-style conditional formatting (tblStylePr) only
if simple; otherwise report it as unresolved.

Resolution precedence to implement and report (with source):
- cell margin: tcMar > tblPrEx/tblCellMar > tblPr/tblCellMar > table style tblCellMar > (unresolved default; report null)
- cell border side: tcBorders > tblPrEx/tblBorders > tblPr/tblBorders > table style tblBorders; for tblBorders
  use top/bottom/left/right (or start/end) on the outer edges of the table and insideH/insideV on inner edges,
  according to the cell's grid position and spans. Report val, sz (eighths of a point -> pt), space, color.

ODT side, per `table:table`: table style props (style:width, style:rel-width, table:align, fo:margin-*,
table:border-model), column widths (table-column styles, number-columns-repeated), per row: style:row-height,
style:min-row-height, style:use-optimal-row-height, fo:keep-together, number-rows-repeated; per cell:
number-columns-spanned/rows-spanned, covered-table-cell, fo:padding and fo:padding-* , fo:border and fo:border-*,
style:vertical-align, style:writing-mode, fo:background-color; paragraphs in the cell: resolved fo:line-height,
style:line-height-at-least, style:line-spacing, fo:margin-top/bottom, style:snap-to-layout-grid,
style:contextual-spacing, fo:text-align, indents (fo: and loext:), fo:font-size / style:font-size-asian,
style:font-name / style:font-name-asian, each with source level (automatic style / parent chain / default-style).
Note: in LibreOffice an automatic style's parent-style-name that points to another automatic style is not resolved;
report such a parent as `parentIsAutomatic: true` rather than silently resolving it.

### Alignment and comparison
- Tables: by order; if counts differ, match by normalized text similarity (NFKC, whitespace removed, first 300 chars).
- Rows: by index; if counts differ, align by row text; report unmatched rows.
- Cells: by grid column start (Word gridBefore/gridSpan; ODT spans and covered cells).
- Categories (count each mismatch once per row/cell/side):
  - `rowHeightRule`: Word (hRule RAW, val pt) vs ODT (row-height exact / min-row-height / none). Record the pair
    as a pattern string, e.g. `word:atLeast|absent 14.2pt -> odt:min 14.2pt`. Count value diffs > 0.5 pt separately.
  - `cellPadding`: per side, |diff| > 0.5 pt, or resolved on one side only.
  - `cellBorder`: per side: presence mismatch, width diff > 0.25 pt, style mismatch (single/double/dotted...).
  - `vAlign`, `columnWidth` (> 1 pt), `tableWidth` (> 1 pt), `paraLineSpacing` (rule or value),
    `paraSpacingBeforeAfter` (> 0.5 pt), `snapToGrid`, `fontSize` (> 0.25 pt), `eastAsiaFont`.
  - Document: page size/margins (> 0.5 pt), header/footer distances, grid (Word type/linePitch/charSpace vs ODT
    layout-grid-mode/base-height/ruby-height/base-width/lines).
- Units: twips/20 = pt; border sz/8 = pt; 1in = 72pt; 1cm = 72/2.54 pt.

Keep the JSON reasonably compact (no full paragraph text; at most 40 characters of cell text as a label).
