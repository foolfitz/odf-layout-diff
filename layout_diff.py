#!/usr/bin/env python3
"""Compare where text lands in a reference PDF and a candidate PDF.

The reference is a rendering trusted to have the intended layout (for
example the original form exported by Microsoft Word); the candidate is the
same document rendered by LibreOffice. The report lists mechanical symptoms
-- text blocks that gained or lost lines, vertical shifts no line change
explains, rows moved to another page -- ranked by how many points they move
the layout, so that a small language model does not have to infer them from
raw geometry.

Text positions come from poppler's ``pdftotext -bbox-layout``. Only the
standard library is used.
"""
from __future__ import annotations

import argparse
import difflib
import json
import statistics
import subprocess
import sys
import unicodedata
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from dataclasses import dataclass, field

REPORT_VERSION = 2
XHTML = "{http://www.w3.org/1999/xhtml}"

# Lines whose vertical centres are closer than this fraction of the smaller
# line height belong to the same visual row.
ROW_TOLERANCE = 0.4
# Matching blocks shorter than this are ignored: short runs of common
# characters ("年", "：") align spuriously.
MIN_MATCH_BLOCK = 3
# Private-use characters that never occur in extracted text; one per side,
# so a separator never matches.
SEPARATORS = ("\ue000", "\ue001")
# The second alignment pass only moves runs at least this long.
MIN_MOVED_RUN = 4
# A row holding fewer matched glyphs than this does not count as a line.
MIN_RUN = 2
# The next glyph starts this far left of the previous one: a new line began.
BACK_LEFT_PT = 1.0
# `hint` thresholds; see README.md.
SHIFTED_START_PT = 3.0
WIDER_TEXT_RATIO = 1.005
WIDE_GAP_PT = 20.0
# A line that does not wrap differently is reported as moved sideways from
# this distance on: about a table cell's default padding, so smaller shifts
# keep the text inside its cell.
LINE_START_SHIFT_PT = 6.0
# A `pdftotext` line in the same block and row that starts no further right
# than this past the end of the previous one continues it.
LINE_JOIN_PT = 1.0
TEXT_LIMIT = 60
NODE_KEY_LENGTH = 8
MAX_BREAKS = 3


@dataclass
class Word:
    x_min: float
    x_max: float
    text: str


@dataclass
class Line:
    page: int
    block: int
    x_min: float
    y_min: float
    x_max: float
    y_max: float
    words: list[Word]

    @property
    def center(self) -> float:
        return (self.y_min + self.y_max) / 2

    @property
    def height(self) -> float:
        return self.y_max - self.y_min

    @property
    def text(self) -> str:
        return " ".join(word.text for word in self.words)


@dataclass
class Glyph:
    char: str
    line: int
    x: float
    # Words on one extracted line split into segments at gaps of at least
    # WIDE_GAP_PT: `pdftotext` often joins side-by-side cells into one line.
    segment: int


@dataclass
class Row:
    page: int
    anchor: float
    height: float
    lines: list[int] = field(default_factory=list)


@dataclass
class Layout:
    page_heights: list[float]
    lines: list[Line]
    glyphs: list[Glyph]
    rows: list[Row] = field(default_factory=list)
    row_of_line: list[int] = field(default_factory=list)

    def page_offset(self, page: int) -> float:
        return sum(self.page_heights[: page - 1])

    def line_of(self, glyph: int) -> Line:
        return self.lines[self.glyphs[glyph].line]

    def row_of(self, glyph: int) -> int:
        return self.row_of_line[self.glyphs[glyph].line]

    def absolute_center(self, glyph: int) -> float:
        line = self.line_of(glyph)
        return self.page_offset(line.page) + line.center

    def row_text(self, row: int) -> str:
        lines = sorted(self.rows[row].lines, key=lambda index: self.lines[index].x_min)
        return shorten(" ".join(self.lines[index].text for index in lines))

    def is_new_line(self, before: int, after: int) -> bool:
        """`after` sits on a lower row than `before` and starts further left."""
        upper, lower = self.rows[self.row_of(before)], self.rows[self.row_of(after)]
        return (lower.page, lower.anchor) > (upper.page, upper.anchor) and (
            self.glyphs[after].x < self.glyphs[before].x - BACK_LEFT_PT
        )


def normalize(text: str) -> str:
    """NFKC without whitespace: the unit both renderings are aligned on."""
    return "".join(
        char for char in unicodedata.normalize("NFKC", text) if not char.isspace()
    )


def shorten(text: str) -> str:
    return text if len(text) <= TEXT_LIMIT else text[: TEXT_LIMIT - 1] + "…"


def parse_layout(xml_text: str) -> Layout:
    root = ET.fromstring(xml_text)
    page_heights: list[float] = []
    lines: list[Line] = []
    glyphs: list[Glyph] = []
    block_index = -1
    segment_index = -1
    for page_number, page in enumerate(root.iter(f"{XHTML}page"), start=1):
        page_heights.append(float(page.get("height", "0")))
        for block in page.iter(f"{XHTML}block"):
            block_index += 1
            block_lines: list[Line] = []
            for element in block.iter(f"{XHTML}line"):
                words = [
                    Word(float(word.get("xMin")), float(word.get("xMax")), word.text or "")
                    for word in element.iter(f"{XHTML}word")
                ]
                if not words:
                    continue
                line = Line(
                    page_number,
                    block_index,
                    float(element.get("xMin")),
                    float(element.get("yMin")),
                    float(element.get("xMax")),
                    float(element.get("yMax")),
                    words,
                )
                if block_lines and continues_line(block_lines[-1], line):
                    join_lines(block_lines[-1], line)
                else:
                    block_lines.append(line)
            for line in block_lines:
                line_index = len(lines)
                lines.append(line)
                words = line.words
                segment_index += 1
                for previous, word in zip([None, *words], words):
                    if previous is not None and word.x_min - previous.x_max >= WIDE_GAP_PT:
                        segment_index += 1
                    chars = normalize(word.text)
                    for position, char in enumerate(chars):
                        x = word.x_min + (word.x_max - word.x_min) * (position + 0.5) / len(chars)
                        glyphs.append(Glyph(char, line_index, x, segment_index))
    layout = Layout(page_heights, lines, glyphs)
    assign_rows(layout)
    return layout


def continues_line(previous: Line, line: Line) -> bool:
    """`pdftotext` splits one visual line where the font or baseline changes
    (for example before a full-width bracket); the parts overlap or touch."""
    return (
        abs(line.center - previous.center) <= ROW_TOLERANCE * min(line.height, previous.height)
        and previous.x_min <= line.x_min <= previous.x_max + LINE_JOIN_PT
    )


def join_lines(previous: Line, line: Line) -> None:
    previous.words = sorted(previous.words + line.words, key=lambda word: word.x_min)
    previous.x_max = max(previous.x_max, line.x_max)
    previous.y_min = min(previous.y_min, line.y_min)
    previous.y_max = max(previous.y_max, line.y_max)


def assign_rows(layout: Layout) -> None:
    order = sorted(
        range(len(layout.lines)),
        key=lambda index: (layout.lines[index].page, layout.lines[index].center),
    )
    layout.row_of_line = [0] * len(layout.lines)
    current: Row | None = None
    for index in order:
        line = layout.lines[index]
        if (
            current is None
            or current.page != line.page
            or line.center - current.anchor > ROW_TOLERANCE * min(line.height, current.height)
        ):
            current = Row(line.page, line.center, line.height)
            layout.rows.append(current)
        current.lines.append(index)
        layout.row_of_line[index] = len(layout.rows) - 1


def pdf_layout(path: str) -> Layout:
    completed = subprocess.run(
        ["pdftotext", "-bbox-layout", path, "-"],
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    return parse_layout(completed.stdout)


def glyph_stream(layout: Layout, separator: str) -> tuple[str, list[int]]:
    """The layout's characters with `separator` between text blocks, and the
    glyph index behind each position (-1 for a separator)."""
    chars: list[str] = []
    index: list[int] = []
    previous_block = None
    for glyph_index, glyph in enumerate(layout.glyphs):
        block = layout.lines[glyph.line].block
        if previous_block is not None and block != previous_block:
            chars.append(separator)
            index.append(-1)
        previous_block = block
        chars.append(glyph.char)
        index.append(glyph_index)
    return "".join(chars), index


def align(reference: Layout, candidate: Layout) -> list[tuple[int, int]]:
    """(reference glyph, candidate glyph) pairs holding the same character.

    Each side separates its text blocks with a character the other side never
    contains, so no match spans a block boundary: otherwise a shared tail such
    as "ame)" followed by the next block's text can outscore the block's own
    text. Runs the extraction order moved are paired in a second pass.
    """
    reference_text, reference_index = glyph_stream(reference, SEPARATORS[0])
    candidate_text, candidate_index = glyph_stream(candidate, SEPARATORS[1])
    matcher = difflib.SequenceMatcher(None, reference_text, candidate_text, autojunk=False)
    pairs = []
    for block in matcher.get_matching_blocks():
        if block.size >= MIN_MATCH_BLOCK:
            pairs.extend(
                (reference_index[block.a + offset], candidate_index[block.b + offset])
                for offset in range(block.size)
            )
    return sorted(pairs + moved_runs(reference, candidate, pairs))


def moved_runs(reference: Layout, candidate: Layout, pairs: list[tuple[int, int]]) -> list[tuple[int, int]]:
    """Pairs for unmatched reference runs found exactly once among the
    unmatched candidate glyphs, wherever they are."""
    matched_reference = {a for a, _ in pairs}
    matched_candidate = {b for _, b in pairs}
    candidate_text = "".join(
        SEPARATORS[1] if index in matched_candidate else glyph.char
        for index, glyph in enumerate(candidate.glyphs)
    )
    runs: list[list[int]] = []
    current: list[int] = []
    for index in range(len(reference.glyphs)):
        if index in matched_reference or (
            current and reference.line_of(current[-1]).block != reference.line_of(index).block
        ):
            if current:
                runs.append(current)
            current = [] if index in matched_reference else [index]
            continue
        current.append(index)
    if current:
        runs.append(current)
    found = []
    for run in sorted(runs, key=lambda run: (-len(run), run[0])):
        if len(run) < MIN_MOVED_RUN:
            continue
        text = "".join(reference.glyphs[index].char for index in run)
        start = candidate_text.find(text)
        if start < 0 or candidate_text.find(text, start + 1) >= 0:
            continue
        found.extend(zip(run, range(start, start + len(run))))
        candidate_text = candidate_text[:start] + SEPARATORS[1] * len(run) + candidate_text[start + len(run) :]
    return found


def counted_rows(layout: Layout, glyphs: list[int]) -> set[int]:
    counts = Counter(layout.row_of(glyph) for glyph in glyphs)
    return {row for row, count in counts.items() if count >= MIN_RUN}


def text_height(layout: Layout, glyphs: list[int], rows: set[int]) -> float:
    """Vertical extent of the lines holding `glyphs`, summed per page."""
    extents: dict[int, list[float]] = {}
    for glyph in glyphs:
        if layout.row_of(glyph) not in rows:
            continue
        line = layout.line_of(glyph)
        low, high = extents.get(line.page, [line.y_min, line.y_max])
        extents[line.page] = [min(low, line.y_min), max(high, line.y_max)]
    return sum(high - low for low, high in extents.values())


def span(layout: Layout, glyphs: list[int]) -> float:
    xs = [layout.glyphs[glyph].x for glyph in glyphs]
    return max(xs) - min(xs) if xs else 0.0


def run_around(
    reference: Layout, members: list[tuple[int, int]], index: int, row_of, step: int
) -> list[tuple[int, int]]:
    """The contiguous members sharing a row and a reference segment with
    `members[index]`, walking by `step`."""
    key = (row_of(members[index]), reference.glyphs[members[index][0]].segment)
    run = []
    while 0 <= index < len(members) and (
        row_of(members[index]),
        reference.glyphs[members[index][0]].segment,
    ) == key:
        run.append(members[index])
        index += step
    return run if step > 0 else run[::-1]


def starts_segment(layout: Layout, glyph: int) -> bool:
    return glyph == 0 or layout.glyphs[glyph - 1].segment != layout.glyphs[glyph].segment


def segment_text(layout: Layout, glyph: int, step: int, limit: int = NODE_KEY_LENGTH) -> str:
    """Up to `limit` characters of `glyph`'s segment, starting at it (`step`
    1) or ending at it (`step` -1)."""
    segment = layout.glyphs[glyph].segment
    chars = []
    while 0 <= glyph < len(layout.glyphs) and layout.glyphs[glyph].segment == segment and len(chars) < limit:
        chars.append(layout.glyphs[glyph].char)
        glyph += step
    return "".join(chars if step > 0 else reversed(chars))


def segment_end(layout: Layout, glyph: int) -> int:
    while glyph + 1 < len(layout.glyphs) and layout.glyphs[glyph + 1].segment == layout.glyphs[glyph].segment:
        glyph += 1
    return glyph


def paragraph_neighbour(layout: Layout, glyph: int, nodes: list[dict]) -> int | None:
    """The glyph just left of `glyph` on its row, when a projection excerpt
    shows both belong to one paragraph: `glyph` then starts a later segment
    of that paragraph, after a run of spaces, not a line."""
    if not nodes:
        return None
    row, x = layout.row_of(glyph), layout.glyphs[glyph].x
    left = [index for index, other in enumerate(layout.glyphs) if other.x < x and layout.row_of(index) == row]
    if not left:
        return None
    neighbour = max(left, key=lambda index: layout.glyphs[index].x)
    joined = segment_text(layout, neighbour, -1) + segment_text(layout, glyph, 1)
    return neighbour if any(joined in node["excerpt"] for node in nodes) else None


def paragraph_end(layout: Layout, glyph: int, nodes: list[dict]) -> int:
    """The last glyph of the line `glyph` starts, following later segments
    of the same paragraph on its row (see `paragraph_neighbour`)."""
    end = segment_end(layout, glyph)
    if not nodes:
        return end
    row = layout.row_of(glyph)
    while True:
        x = layout.glyphs[end].x
        right = [index for index, other in enumerate(layout.glyphs) if other.x > x and layout.row_of(index) == row]
        if not right:
            return end
        neighbour = min(right, key=lambda index: layout.glyphs[index].x)
        joined = segment_text(layout, end, -1) + segment_text(layout, neighbour, 1)
        if not starts_segment(layout, neighbour) or not any(joined in node["excerpt"] for node in nodes):
            return end
        end = segment_end(layout, neighbour)


def same_text_on_row(reference: Layout, candidate: Layout, end: int, row: int) -> int | None:
    """The candidate glyph that ends, on `row`, a segment ending with the
    text of the reference segment `end` ends, nearest its reference
    position: repeated text (a note after each of several labels) is often
    left unaligned."""
    key = segment_text(reference, end, -1)
    ends = [
        glyph
        for glyph in range(len(candidate.glyphs))
        if candidate.row_of(glyph) == row
        and segment_end(candidate, glyph) == glyph
        and segment_text(candidate, glyph, -1) == key
    ]
    return min(ends, key=lambda glyph: abs(candidate.glyphs[glyph].x - reference.glyphs[end].x), default=None)


def describe_break(
    kind: str,
    reference: Layout,
    candidate: Layout,
    members: list[tuple[int, int]],
    index: int,
    nodes: list[dict],
    reported_starts: set[int],
) -> dict:
    """The break between `members[index]` and `members[index + 1]`. A
    reference glyph given a `startShiftPt` is added to `reported_starts`."""
    if kind == "added":
        row_of = lambda pair: candidate.row_of(pair[1])  # noqa: E731
        shared, shared_index = reference, 0
    else:
        row_of = lambda pair: reference.row_of(pair[0])  # noqa: E731
        shared, shared_index = candidate, 1
    before = run_around(reference, members, index, row_of, -1)
    after = run_around(reference, members, index + 1, row_of, 1)
    reference_width = span(reference, [a for a, _ in before]) + span(reference, [a for a, _ in after])
    candidate_width = span(candidate, [b for _, b in before]) + span(candidate, [b for _, b in after])
    ratio = candidate_width / reference_width if reference_width > 0 else None
    # The gap is measured in the layout where both sides share one line.
    gap = shared.glyphs[after[0][shared_index]].x - shared.glyphs[before[-1][shared_index]].x
    # A start position only means something where the text before the break
    # begins a line in both layouts: mid-line it just reflects the reflow, and
    # after a run of spaces in its own paragraph it reflects that run's width.
    first, first_candidate = before[0]
    shift = candidate.glyphs[first_candidate].x - reference.glyphs[first].x
    start_shift = None
    after_wider_gap = False
    if starts_segment(reference, first) and starts_segment(candidate, first_candidate):
        neighbour = paragraph_neighbour(reference, first, nodes)
        if neighbour is None:
            start_shift = shift
            reported_starts.add(first)
        else:
            after_wider_gap = (
                shift >= SHIFTED_START_PT
                and reference.glyphs[first].x - reference.glyphs[neighbour].x >= WIDE_GAP_PT
            )
    if start_shift is not None and abs(start_shift) >= SHIFTED_START_PT:
        hint = "shifted-start"
    elif ratio is not None and ratio > WIDER_TEXT_RATIO:
        hint = "wider-text"
    elif gap >= WIDE_GAP_PT or after_wider_gap:
        hint = "wide-gap"
    else:
        hint = "unknown"
    return {
        "kind": kind,
        "textBeforeBreak": shorten("".join(reference.glyphs[a].char for a, _ in before)),
        "textAfterBreak": shorten("".join(reference.glyphs[a].char for a, _ in after)),
        "startShiftPt": round(start_shift, 1) if start_shift is not None else None,
        "gapPt": round(gap, 1),
        "widthRatio": round(ratio, 3) if ratio is not None else None,
        "hint": hint,
    }


def without_reflow(breaks: list[dict]) -> list[dict]:
    """Drops adjacent added/removed pairs that only move a break further
    along the same text: the line count does not change."""
    kept: list[dict] = []
    for item in breaks:
        if kept and kept[-1]["kind"] != item["kind"]:
            earlier = kept[-1]["textBeforeBreak"].rstrip("…")
            later = item["textBeforeBreak"].rstrip("…")
            if earlier.startswith(later) or later.startswith(earlier):
                kept.pop()
                continue
        kept.append(item)
    return kept


def row_breaks(
    reference: Layout,
    candidate: Layout,
    members: list[tuple[int, int]],
    nodes: list[dict],
    reported_starts: set[int],
) -> list[dict]:
    """Line breaks the candidate added inside one reference row, across
    whatever text blocks the row is made of."""
    members = sorted(members, key=lambda pair: (reference.glyphs[pair[0]].x, pair[0]))
    found = []
    for index in range(len(members) - 1):
        (_, b1), (_, b2) = members[index], members[index + 1]
        if candidate.row_of(b1) != candidate.row_of(b2) and candidate.is_new_line(b1, b2):
            found.append(describe_break("added", reference, candidate, members, index, nodes, reported_starts))
    return without_reflow(found)


def break_keys(breaks: list[dict]) -> list[str]:
    keys = []
    for item in breaks:
        keys.append(item["textAfterBreak"][:NODE_KEY_LENGTH])
        keys.append(item["textBeforeBreak"].rstrip("…")[-NODE_KEY_LENGTH:])
    return keys


def block_changes(
    reference: Layout,
    candidate: Layout,
    pairs: list[tuple[int, int]],
    nodes: list[dict],
    reported_starts: set[int],
) -> list[dict]:
    """Reference text blocks whose matched text spans a different number of
    lines in the candidate."""
    by_block: dict[int, list[tuple[int, int]]] = defaultdict(list)
    for pair in pairs:
        by_block[reference.line_of(pair[0]).block].append(pair)
    changes = []
    for block in sorted(by_block):
        members = sorted(by_block[block])
        reference_glyphs = [a for a, _ in members]
        candidate_glyphs = [b for _, b in members]
        reference_rows = counted_rows(reference, reference_glyphs)
        candidate_rows = counted_rows(candidate, candidate_glyphs)
        breaks = []
        for index in range(len(members) - 1):
            (a1, b1), (a2, b2) = members[index], members[index + 1]
            same_reference_row = reference.row_of(a1) == reference.row_of(a2)
            same_candidate_row = candidate.row_of(b1) == candidate.row_of(b2)
            if same_reference_row and not same_candidate_row and candidate.is_new_line(b1, b2):
                if {candidate.row_of(b1), candidate.row_of(b2)} <= candidate_rows:
                    breaks.append(
                        describe_break("added", reference, candidate, members, index, nodes, reported_starts)
                    )
            elif same_candidate_row and not same_reference_row and reference.is_new_line(a1, a2):
                if {reference.row_of(a1), reference.row_of(a2)} <= reference_rows:
                    breaks.append(
                        describe_break("removed", reference, candidate, members, index, nodes, reported_starts)
                    )
        # Counting breaks rather than rows keeps a block that holds several
        # side-by-side cells from reporting their mutual misalignment.
        line_change = sum(1 if item["kind"] == "added" else -1 for item in breaks)
        if line_change == 0:
            continue
        first_line = reference.line_of(members[0][0])
        block_lines = [line for line in reference.lines if line.block == block]
        breaks = without_reflow(breaks)
        change = {
            "kind": "text-block-taller" if line_change > 0 else "text-block-shorter",
            "page": first_line.page,
            "y": round(first_line.center, 1),
            "text": block_text(reference, block),
            "lineChange": line_change,
            "breakCount": len(breaks),
            "heightChangePt": round(
                text_height(candidate, candidate_glyphs, candidate_rows)
                - text_height(reference, reference_glyphs, reference_rows),
                1,
            ),
            "breaks": breaks[:MAX_BREAKS],
        }
        if nodes:
            keys = [normalize(" ".join(line.text for line in block_lines))[:NODE_KEY_LENGTH]]
            change["nodes"] = match_nodes(nodes, keys + break_keys(breaks))
        changes.append({"impact": abs(change["heightChangePt"]), "rows": reference_rows, "symptom": change})
    return changes


def block_text(layout: Layout, block: int) -> str:
    return shorten(" ".join(line.text for line in layout.lines if line.block == block))


def row_keys(layout: Layout, row: int) -> list[str]:
    """Node keys for a row: the start and end of each of its lines."""
    keys = []
    for index in layout.rows[row].lines:
        text = normalize(layout.lines[index].text)
        keys += [text[:NODE_KEY_LENGTH], text[-NODE_KEY_LENGTH:]]
    return keys


def extra_rows(
    reference: Layout,
    candidate: Layout,
    pairs_by_candidate_row: dict[int, list[tuple[int, int]]],
    upper: int,
    lower: int,
    target_of: dict[int, int],
) -> list[str]:
    """Texts of candidate rows between where reference rows `upper` and
    `lower` landed that hold no text of another reference row: lines the
    candidate added, such as the last character of `upper` wrapped on its
    own (too short to align, so no break is found)."""
    between = {
        normalize("".join(reference.lines[index].text for index in reference.rows[row].lines))
        for row in range(upper + 1, lower)
    }
    found = []
    for row in range(target_of[upper] + 1, target_of[lower]):
        if {reference.row_of(a) for a, _ in pairs_by_candidate_row[row]} - {upper}:
            continue
        text = normalize("".join(candidate.lines[index].text for index in candidate.rows[row].lines))
        if text not in between:
            found.append(candidate.row_text(row))
    return found


def line_start_shifts(
    reference: Layout,
    candidate: Layout,
    pairs: list[tuple[int, int]],
    nodes: list[dict],
    reported_starts: set[int],
    target_of: dict[int, int],
) -> list[dict]:
    """Lines that start at another x in the candidate while staying on the
    row their reference row landed on: an indent or margin change that moves
    text without wrapping it, for example out of its table cell. Consecutive
    lines of one block moved by the same amount form one symptom."""
    partner = dict(pairs)
    starts_by_text: dict[str, list[int]] = {}
    for glyph in range(len(reference.glyphs)):
        if starts_segment(reference, glyph):
            starts_by_text.setdefault(segment_text(reference, glyph, 1), []).append(glyph)

    def on_another_copy(a: int, b: int) -> bool:
        """The candidate text starts where another copy of the same text
        starts in the reference, on a line beside it: the two copies (for
        example one label in two cells, which may sit a few points apart
        vertically) were aligned with each other because the renderings
        extract them in another order."""
        line = reference.line_of(a)
        return any(
            other != a
            and reference.line_of(other).page == line.page
            and reference.line_of(other).y_min < line.y_max
            and line.y_min < reference.line_of(other).y_max
            and abs(reference.glyphs[other].x - candidate.glyphs[b].x) < LINE_START_SHIFT_PT
            for other in starts_by_text[segment_text(reference, a, 1)]
        )

    groups: list[dict] = []
    for a, b in pairs:
        if a in reported_starts or not (starts_segment(reference, a) and starts_segment(candidate, b)):
            continue
        if target_of.get(reference.row_of(a)) != candidate.row_of(b):
            continue
        shift = candidate.glyphs[b].x - reference.glyphs[a].x
        if abs(shift) < LINE_START_SHIFT_PT or paragraph_neighbour(reference, a, nodes) is not None:
            continue
        if on_another_copy(a, b):
            continue
        end = paragraph_end(reference, a, nodes)
        end_partner = partner.get(end)
        if end_partner is None:
            end_partner = same_text_on_row(reference, candidate, end, candidate.row_of(b))
        hint = "shifted-start"
        if end_partner is not None and candidate.row_of(end_partner) == candidate.row_of(b):
            end_shift = candidate.glyphs[end_partner].x - reference.glyphs[end].x
            # Centred text that got wider or narrower moves both ends apart.
            if shift * end_shift < 0 and abs(shift + end_shift) <= LINE_JOIN_PT:
                continue
            # Right-aligned text keeps its end: the line got wider or narrower.
            if end != a and abs(end_shift) < SHIFTED_START_PT:
                hint = "wider-text" if shift < 0 else "narrower-text"
        block = reference.line_of(a).block
        same = [
            group
            for group in groups
            if group["block"] == block and group["hint"] == hint and abs(shift - group["shift"]) <= LINE_JOIN_PT
        ]
        if same:
            same[0]["glyphs"].append(a)
        else:
            groups.append({"block": block, "shift": shift, "hint": hint, "glyphs": [a]})
    found = []
    for group in groups:
        first = group["glyphs"][0]
        line = reference.line_of(first)
        symptom = {
            "kind": "shifted-line-start",
            "page": line.page,
            "y": round(line.center, 1),
            "text": shorten(segment_text(reference, first, 1, TEXT_LIMIT)),
            "lineCount": len(group["glyphs"]),
            "shiftPt": round(group["shift"], 1),
            "hint": group["hint"],
        }
        if nodes:
            keys = [normalize(block_text(reference, group["block"]))[:NODE_KEY_LENGTH]]
            symptom["nodes"] = match_nodes(nodes, keys + [segment_text(reference, g, 1) for g in group["glyphs"]])
        found.append({"impact": abs(group["shift"]), "rows": set(), "symptom": symptom})
    return found


def row_positions(reference: Layout, candidate: Layout, pairs: list[tuple[int, int]]) -> list[dict]:
    """Where each reference row landed, judged by the candidate row of its
    leftmost run: text the candidate wrapped onto a later row does not move
    the row itself."""
    grouped: dict[int, list[tuple[int, int]]] = defaultdict(list)
    for pair in pairs:
        grouped[reference.row_of(pair[0])].append(pair)
    # Rows repeated on several pages (running headers and footers) stay put
    # while the body moves, so they would report the body's shift reversed.
    pages_by_text: dict[str, set[int]] = defaultdict(set)
    for row in reference.rows:
        pages_by_text[normalize("".join(reference.lines[index].text for index in row.lines))].add(row.page)
    positions = []
    for row, members in grouped.items():
        text = normalize("".join(reference.lines[index].text for index in reference.rows[row].lines))
        if len(pages_by_text[text]) > 1:
            continue
        members.sort(key=lambda pair: (reference.glyphs[pair[0]].x, pair[0]))
        target = None
        for pair in members:
            row_members = [other for other in members if candidate.row_of(other[1]) == candidate.row_of(pair[1])]
            if len(row_members) >= MIN_RUN:
                target = candidate.row_of(pair[1])
                break
        if target is None:
            continue
        source = reference.rows[row]
        positions.append(
            {
                "row": row,
                "page": source.page,
                "anchor": source.anchor,
                "candidateRow": target,
                "candidatePage": candidate.rows[target].page,
                "shift": statistics.median(
                    candidate.absolute_center(b) - reference.absolute_center(a)
                    for a, b in members
                    if candidate.row_of(b) == target
                ),
            }
        )
    positions.sort(key=lambda item: (item["page"], item["anchor"]))
    return positions


def compare(
    reference: Layout,
    candidate: Layout,
    nodes: list[dict] | None = None,
    threshold: float = 3.0,
    limit: int = 10,
) -> dict:
    nodes = nodes or []
    pairs = align(reference, candidate)
    pairs_by_row: dict[int, list[tuple[int, int]]] = defaultdict(list)
    pairs_by_candidate_row: dict[int, list[tuple[int, int]]] = defaultdict(list)
    for pair in pairs:
        pairs_by_row[reference.row_of(pair[0])].append(pair)
        pairs_by_candidate_row[candidate.row_of(pair[1])].append(pair)
    reported_starts: set[int] = set()
    found = block_changes(reference, candidate, pairs, nodes, reported_starts)
    explained_rows = set().union(*(item["rows"] for item in found)) if found else set()

    positions = row_positions(reference, candidate, pairs)
    target_of = {position["row"]: position["candidateRow"] for position in positions}
    moved: dict[tuple[int, int], list[int]] = defaultdict(list)
    kept = [position for position in positions if position["candidatePage"] == position["page"]]
    for position in positions:
        if position["candidatePage"] != position["page"]:
            moved[(position["page"], position["candidatePage"])].append(position["row"])
    for index in range(1, len(kept)):
        previous, current = kept[index - 1], kept[index]
        if previous["page"] != current["page"] or previous["row"] in explained_rows:
            continue
        delta = current["shift"] - previous["shift"]
        if abs(delta) < threshold:
            continue
        # A shift that the next row undoes is jitter between neighbouring
        # cells, not a displacement of the layout.
        following = kept[index + 1] if index + 1 < len(kept) else None
        if following is not None and following["page"] == current["page"]:
            if abs(following["shift"] - previous["shift"]) < threshold:
                continue
        causes = row_breaks(reference, candidate, pairs_by_row[previous["row"]], nodes, reported_starts)
        symptom = {
            "kind": "vertical-shift",
            "page": current["page"],
            "y": round(current["anchor"], 1),
            "rowAbove": reference.row_text(previous["row"]),
            "row": reference.row_text(current["row"]),
            "shiftPt": round(delta, 1),
            "breaksInRowAbove": causes[:MAX_BREAKS],
        }
        if causes:
            if nodes:
                symptom["nodes"] = match_nodes(nodes, break_keys(causes))
        else:
            added = extra_rows(reference, candidate, pairs_by_candidate_row, previous["row"], current["row"], target_of)
            above = normalize("".join(reference.lines[index].text for index in reference.rows[previous["row"]].lines))
            symptom["extraRows"] = added
            symptom["hint"] = "wrapped-row" if added and all(normalize(text) in above for text in added) else "unknown"
            if nodes:
                symptom["nodes"] = match_nodes(nodes, row_keys(reference, previous["row"]))
        found.append({"impact": abs(delta), "rows": set(), "symptom": symptom})
    found.extend(line_start_shifts(reference, candidate, pairs, nodes, reported_starts, target_of))

    found.sort(key=lambda item: (-item["impact"], item["symptom"]["page"], item["symptom"]["y"]))
    overflow = [
        {
            "page": page,
            "candidatePage": candidate_page,
            "rowCount": len(rows),
            "rows": [reference.row_text(row) for row in rows[:5]],
        }
        for (page, candidate_page), rows in sorted(moved.items())
    ]
    return {
        "reportVersion": REPORT_VERSION,
        "pages": {"reference": len(reference.page_heights), "candidate": len(candidate.page_heights)},
        "matchedGlyphRatio": round(len(pairs) / max(len(reference.glyphs), 1), 3),
        "rowsOnAnotherPage": overflow,
        "symptoms": [item["symptom"] for item in found[:limit]],
        "omittedSymptoms": max(len(found) - limit, 0),
    }


def load_projection_nodes(paths: list[str]) -> list[dict]:
    """Paragraph and heading nodes from `odf-tool protocol` `project` responses."""
    nodes = []
    for path in paths:
        with open(path, encoding="utf-8") as handle:
            data = json.load(handle)
        for node in data.get("result", data).get("nodes", []):
            if node.get("kind") in ("paragraph", "heading"):
                nodes.append(
                    {
                        "address": node["address"],
                        "styleName": node.get("styleName"),
                        "excerpt": normalize(node.get("excerpt", "")),
                    }
                )
    return nodes


def match_nodes(nodes: list[dict], keys: list[str], limit: int = 3) -> list[dict]:
    """The projection nodes whose excerpt contains the most keys."""
    keys = sorted({key for key in keys if len(key) >= MIN_MATCH_BLOCK})
    scored = []
    for node in nodes:
        score = sum(1 for key in keys if key in node["excerpt"])
        if score:
            address = tuple(int(part) for part in node["address"].split("/"))
            scored.append((score, address, node))
    if not scored:
        return []
    best = max(score for score, _, _ in scored)
    winners = sorted((item for item in scored if item[0] == best), key=lambda item: item[1])
    return [{"address": node["address"], "styleName": node["styleName"]} for _, _, node in winners[:limit]]


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("reference", help="PDF with the intended layout")
    parser.add_argument("candidate", help="PDF rendered by LibreOffice")
    parser.add_argument(
        "--projection",
        action="append",
        default=[],
        metavar="JSON",
        help="odf-tool project response for the candidate document (repeatable, one per window)",
    )
    parser.add_argument("--threshold", type=float, default=3.0, help="smallest vertical shift reported, in pt")
    parser.add_argument("--limit", type=int, default=10, help="most symptoms reported")
    arguments = parser.parse_args(argv)
    report = compare(
        pdf_layout(arguments.reference),
        pdf_layout(arguments.candidate),
        load_projection_nodes(arguments.projection),
        arguments.threshold,
        arguments.limit,
    )
    json.dump(report, sys.stdout, ensure_ascii=False, indent=2)
    sys.stdout.write("\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
