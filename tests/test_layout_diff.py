"""Unit tests on synthetic `pdftotext -bbox-layout` documents."""
import unittest

import layout_diff


def word(x_min: float, x_max: float, text: str) -> tuple:
    return (x_min, x_max, text)


def line(y_min: float, y_max: float, *words: tuple) -> tuple:
    return (y_min, y_max, words)


def document(*pages: list[list[tuple]]) -> str:
    """Each page is a list of blocks; each block is a list of lines."""
    parts = ['<html xmlns="http://www.w3.org/1999/xhtml"><body><doc>']
    for blocks in pages:
        parts.append('<page width="595" height="842"><flow>')
        for block in blocks:
            parts.append("<block>")
            for y_min, y_max, words in block:
                x_min = min(w[0] for w in words)
                x_max = max(w[1] for w in words)
                parts.append(f'<line xMin="{x_min}" yMin="{y_min}" xMax="{x_max}" yMax="{y_max}">')
                for w_min, w_max, text in words:
                    parts.append(
                        f'<word xMin="{w_min}" yMin="{y_min}" xMax="{w_max}" yMax="{y_max}">{text}</word>'
                    )
                parts.append("</line>")
            parts.append("</block>")
        parts.append("</flow></page>")
    parts.append("</doc></body></html>")
    return "".join(parts)


def compare(reference: str, candidate: str, nodes: list[dict] | None = None) -> dict:
    return layout_diff.compare(
        layout_diff.parse_layout(reference), layout_diff.parse_layout(candidate), nodes
    )


def kinds(report: dict) -> list[str]:
    return [symptom["kind"] for symptom in report["symptoms"]]


def all_breaks(report: dict) -> list[dict]:
    return [
        item
        for symptom in report["symptoms"]
        for item in symptom.get("breaks", []) + symptom.get("breaksInRowAbove", [])
    ]


class ParseTests(unittest.TestCase):
    def test_glyphs_are_placed_inside_their_word(self) -> None:
        layout = layout_diff.parse_layout(document([[line(100, 110, word(10, 30, "AB"))]]))
        self.assertEqual([glyph.char for glyph in layout.glyphs], ["A", "B"])
        self.assertEqual([glyph.x for glyph in layout.glyphs], [15.0, 25.0])
        self.assertEqual(len(layout.rows), 1)

    def test_normalization_drops_whitespace_and_folds_width(self) -> None:
        self.assertEqual(layout_diff.normalize("Ａ　B（c）"), "AB(c)")

    def test_a_line_split_where_the_font_changes_is_joined(self) -> None:
        # pdftotext splits a line before a full-width bracket set in another
        # font; the two parts overlap.
        layout = layout_diff.parse_layout(
            document([[line(100, 110, word(10, 60, "改或使用。")), line(100, 110, word(57, 80, "【Athin"))]])
        )
        self.assertEqual(len(layout.lines), 1)
        self.assertEqual("".join(glyph.char for glyph in layout.glyphs), "改或使用。【Athin")
        self.assertEqual(len({glyph.segment for glyph in layout.glyphs}), 1)


class SymptomTests(unittest.TestCase):
    def test_a_block_that_wraps_in_the_candidate_is_reported_as_taller(self) -> None:
        reference = document(
            [
                [line(100, 110, word(10, 50, "Country"), word(55, 65, "of"), word(70, 100, "Birth:"))],
                [line(130, 140, word(10, 30, "Next"))],
            ]
        )
        candidate = document(
            [
                [
                    line(100, 110, word(10, 56, "Country"), word(61, 73, "of")),
                    line(112, 122, word(10, 44, "Birth:")),
                ],
                [line(142, 152, word(10, 30, "Next"))],
            ]
        )
        report = compare(reference, candidate)
        # The shift of "Next" is explained by the wrap, so it is not repeated.
        self.assertEqual(kinds(report), ["text-block-taller"])
        symptom = report["symptoms"][0]
        self.assertEqual(symptom["lineChange"], 1)
        self.assertEqual(symptom["heightChangePt"], 12.0)
        self.assertEqual(len(symptom["breaks"]), 1)
        self.assertEqual(symptom["breaks"][0]["textBeforeBreak"], "Countryof")
        self.assertEqual(symptom["breaks"][0]["textAfterBreak"], "Birth:")
        self.assertEqual(symptom["breaks"][0]["hint"], "wider-text")

    def test_a_break_at_a_wide_gap_explains_the_shift_below_it(self) -> None:
        reference = document(
            [
                [line(100, 110, word(10, 60, "Signature:"))],
                [line(100, 110, word(300, 330, "(note)"))],
                [line(130, 140, word(10, 30, "Next"))],
                [line(160, 170, word(10, 30, "Last"))],
            ]
        )
        candidate = document(
            [
                [line(100, 110, word(10, 60, "Signature:"))],
                [line(124, 134, word(10, 40, "(note)"))],
                [line(154, 164, word(10, 30, "Next"))],
                [line(184, 194, word(10, 30, "Last"))],
            ]
        )
        nodes = [
            {"address": "2/1", "styleName": "P1", "excerpt": layout_diff.normalize("Signature: (note)")},
            {"address": "2/2", "styleName": "P2", "excerpt": "Next"},
        ]
        report = compare(reference, candidate, nodes)
        self.assertEqual(kinds(report), ["vertical-shift"])
        symptom = report["symptoms"][0]
        self.assertEqual(symptom["shiftPt"], 24.0)
        self.assertEqual(len(symptom["breaksInRowAbove"]), 1)
        cause = symptom["breaksInRowAbove"][0]
        self.assertEqual(cause["textAfterBreak"], "(note)")
        self.assertEqual(cause["gapPt"], 245.0)
        self.assertEqual(cause["widthRatio"], 1.0)
        self.assertEqual(cause["hint"], "wide-gap")
        self.assertEqual(symptom["nodes"], [{"address": "2/1", "styleName": "P1"}])

    def test_text_extracted_in_another_order_still_aligns_and_a_shifted_start_is_hinted(self) -> None:
        # The reference extracts "(Given Name)" after the next row's block;
        # the candidate, having wrapped it, extracts it right after
        # "(Surname)". "(Surname)" also starts 71 pt further right.
        reference = document(
            [
                [line(200, 210, word(315, 349, "(Surname)"))],
                [line(217, 227, word(52, 80, "Basic"))],
                [line(200, 210, word(501, 523, "(Given"), word(525, 547, "Name)"))],
                [line(240, 250, word(36, 66, "Below"))],
                [line(270, 280, word(36, 60, "Last"))],
            ]
        )
        candidate = document(
            [
                [line(200, 210, word(386, 420, "(Surname)"))],
                [line(212, 222, word(170, 192, "(Given"), word(194, 216, "Name)"))],
                [line(229, 239, word(52, 80, "Basic"))],
                [line(252, 262, word(36, 66, "Below"))],
                [line(282, 292, word(36, 60, "Last"))],
            ]
        )
        report = compare(reference, candidate)
        self.assertEqual(report["matchedGlyphRatio"], 1.0)
        self.assertEqual(kinds(report), ["vertical-shift"])
        symptom = report["symptoms"][0]
        self.assertEqual(symptom["shiftPt"], 12.0)
        cause = symptom["breaksInRowAbove"][0]
        self.assertEqual(cause["textBeforeBreak"], "(Surname)")
        self.assertEqual(cause["textAfterBreak"], "(GivenName)")
        self.assertEqual(cause["startShiftPt"], 71.0)
        self.assertEqual(cause["hint"], "shifted-start")

    def test_text_in_the_middle_of_a_candidate_line_has_no_start_shift(self) -> None:
        # The candidate breaks the first line earlier: "three" begins a
        # reference line but sits after "two" in the candidate.
        reference = document(
            [
                [
                    line(100, 110, word(10, 30, "one"), word(35, 55, "two")),
                    line(112, 122, word(10, 40, "three"), word(45, 70, "four")),
                ],
                [line(140, 150, word(10, 30, "Next"))],
            ]
        )
        candidate = document(
            [
                [
                    line(100, 110, word(10, 30, "one")),
                    line(112, 122, word(10, 30, "two"), word(35, 65, "three")),
                    line(124, 134, word(10, 35, "four")),
                ],
                [line(152, 162, word(10, 30, "Next"))],
            ]
        )
        found = all_breaks(compare(reference, candidate))
        self.assertIn("three", [item["textBeforeBreak"] for item in found])
        self.assertNotIn("shifted-start", [item["hint"] for item in found])

    def test_a_later_segment_of_the_same_paragraph_is_not_a_shifted_start(self) -> None:
        # "(note)" follows "Signature:" after a run of spaces in one
        # paragraph. The candidate's spaces are wider, so "(note)" starts
        # 20 pt further right and "more" wraps: the space run, not an indent.
        reference = document(
            [
                [line(100, 110, word(10, 60, "Signature:"))],
                [line(100, 110, word(300, 330, "(note)"), word(335, 360, "more"))],
                [line(130, 140, word(10, 30, "Next"))],
            ]
        )
        candidate = document(
            [
                [line(100, 110, word(10, 60, "Signature:"))],
                [line(100, 110, word(320, 350, "(note)")), line(124, 134, word(10, 35, "more"))],
                [line(154, 164, word(10, 30, "Next"))],
            ]
        )
        nodes = [{"address": "2/1", "styleName": "P1", "excerpt": layout_diff.normalize("Signature: (note) more")}]
        (without_nodes,) = [item for item in all_breaks(compare(reference, candidate)) if item["textBeforeBreak"] == "(note)"]
        self.assertEqual(without_nodes["startShiftPt"], 20.0)
        self.assertEqual(without_nodes["hint"], "shifted-start")
        (cause,) = [item for item in all_breaks(compare(reference, candidate, nodes)) if item["textBeforeBreak"] == "(note)"]
        self.assertIsNone(cause["startShiftPt"])
        self.assertEqual(cause["hint"], "wide-gap")

    def test_a_neighbour_cell_sitting_lower_is_not_a_line_break(self) -> None:
        reference = document(
            [
                [line(100, 110, word(10, 60, "Signature:"))],
                [line(100, 110, word(300, 330, "(note)"))],
                [line(130, 140, word(10, 30, "Next"))],
            ]
        )
        candidate = document(
            [
                [line(100, 110, word(10, 60, "Signature:"))],
                [line(106, 116, word(300, 330, "(note)"))],
                [line(130, 140, word(10, 30, "Next"))],
            ]
        )
        self.assertEqual(compare(reference, candidate)["symptoms"], [])

    def test_a_reflow_that_keeps_the_line_count_is_not_reported(self) -> None:
        reference = document(
            [
                [
                    line(100, 110, word(10, 30, "one"), word(35, 55, "two"), word(60, 90, "three")),
                    line(112, 122, word(10, 35, "four"), word(40, 60, "five")),
                ]
            ]
        )
        candidate = document(
            [
                [
                    line(100, 110, word(10, 30, "one"), word(35, 55, "two")),
                    line(112, 122, word(10, 40, "three"), word(45, 70, "four"), word(75, 95, "five")),
                ]
            ]
        )
        self.assertEqual(compare(reference, candidate)["symptoms"], [])

    def test_rows_pushed_onto_the_next_page_are_listed(self) -> None:
        reference = document(
            [[line(100, 110, word(10, 40, "Alpha"))], [line(700, 710, word(10, 40, "Omega"))]],
            [[line(100, 110, word(10, 40, "Tail"))]],
        )
        candidate = document(
            [[line(100, 110, word(10, 40, "Alpha"))]],
            [[line(50, 60, word(10, 40, "Omega"))], [line(150, 160, word(10, 40, "Tail"))]],
        )
        report = compare(reference, candidate)
        self.assertEqual(
            report["rowsOnAnotherPage"],
            [{"page": 1, "candidatePage": 2, "rowCount": 1, "rows": ["Omega"]}],
        )

    def test_a_lone_wrapped_character_is_reported_as_an_extra_row(self) -> None:
        # The last "日" of the row wraps on its own; one character is too
        # short to align, so no break is found, but the added row is.
        reference = document(
            [
                [line(100, 110, word(10, 60, "Dates:"))],
                [line(100, 110, word(300, 310, "日"))],
                [line(130, 140, word(10, 30, "Next"))],
                [line(160, 170, word(10, 30, "Last"))],
            ]
        )
        candidate = document(
            [
                [line(100, 110, word(10, 60, "Dates:"))],
                [line(112, 122, word(10, 20, "日"))],
                [line(142, 152, word(10, 30, "Next"))],
                [line(172, 182, word(10, 30, "Last"))],
            ]
        )
        nodes = [
            {"address": "2/1", "styleName": "P1", "excerpt": layout_diff.normalize("Dates: 日")},
            {"address": "2/2", "styleName": "P2", "excerpt": "Next"},
        ]
        report = compare(reference, candidate, nodes)
        self.assertEqual(kinds(report), ["vertical-shift"])
        symptom = report["symptoms"][0]
        self.assertEqual(symptom["breaksInRowAbove"], [])
        self.assertEqual(symptom["extraRows"], ["日"])
        self.assertEqual(symptom["hint"], "wrapped-row")
        self.assertEqual(symptom["nodes"], [{"address": "2/1", "styleName": "P1"}])

    def test_a_shift_without_an_added_row_has_no_hint_but_names_nodes(self) -> None:
        reference = document(
            [
                [line(100, 110, word(10, 60, "Title"))],
                [line(130, 140, word(10, 30, "Next"))],
                [line(160, 170, word(10, 30, "Last"))],
            ]
        )
        candidate = document(
            [
                [line(100, 110, word(10, 60, "Title"))],
                [line(136, 146, word(10, 30, "Next"))],
                [line(166, 176, word(10, 30, "Last"))],
            ]
        )
        nodes = [{"address": "1", "styleName": "P1", "excerpt": "Title"}]
        (symptom,) = compare(reference, candidate, nodes)["symptoms"]
        self.assertEqual(symptom["extraRows"], [])
        self.assertEqual(symptom["hint"], "unknown")
        self.assertEqual(symptom["nodes"], [{"address": "1", "styleName": "P1"}])

    def test_lines_moved_sideways_without_wrapping_are_one_symptom(self) -> None:
        # A negative margin moves both lines of the paragraph 20 pt left,
        # out of their cell; nothing wraps differently.
        reference = document(
            [
                [line(100, 110, word(40, 80, "first"), word(85, 120, "line")), line(112, 122, word(40, 90, "second"))],
                [line(100, 110, word(300, 340, "Other"))],
                [line(140, 150, word(10, 30, "Next"))],
            ]
        )
        candidate = document(
            [
                [line(100, 110, word(20, 60, "first"), word(65, 100, "line")), line(112, 122, word(20, 70, "second"))],
                [line(100, 110, word(300, 340, "Other"))],
                [line(140, 150, word(10, 30, "Next"))],
            ]
        )
        nodes = [{"address": "2/3", "styleName": "P3", "excerpt": layout_diff.normalize("first line second")}]
        (symptom,) = compare(reference, candidate, nodes)["symptoms"]
        self.assertEqual(symptom["kind"], "shifted-line-start")
        self.assertEqual(symptom["shiftPt"], -20.0)
        self.assertEqual(symptom["lineCount"], 2)
        self.assertEqual(symptom["hint"], "shifted-start")
        self.assertEqual(symptom["nodes"], [{"address": "2/3", "styleName": "P3"}])

    def test_centred_text_that_got_wider_is_not_a_moved_line(self) -> None:
        reference = document([[line(100, 110, word(200, 300, "Centred"))], [line(130, 140, word(10, 30, "Next"))]])
        candidate = document([[line(100, 110, word(190, 310, "Centred"))], [line(130, 140, word(10, 30, "Next"))]])
        self.assertEqual(compare(reference, candidate)["symptoms"], [])

    def test_a_small_sideways_move_is_not_reported(self) -> None:
        reference = document([[line(100, 110, word(40, 80, "Label"))], [line(130, 140, word(10, 30, "Next"))]])
        candidate = document([[line(100, 110, word(44, 84, "Label"))], [line(130, 140, word(10, 30, "Next"))]])
        self.assertEqual(compare(reference, candidate)["symptoms"], [])

    def test_two_copies_of_a_label_extracted_in_another_order_are_not_moved_lines(self) -> None:
        # Two cells of one row hold the same label; the candidate extracts the
        # right cell first, so each copy aligns with the other one.
        left = [line(100, 110, word(100, 150, "法定代理人"))]
        right = [line(100, 110, word(300, 350, "法定代理人"))]
        after = [line(140, 150, word(10, 30, "Next"))]
        reference = document([left, right, after])
        candidate = document([right, left, after])
        self.assertEqual(compare(reference, candidate)["symptoms"], [])

    def test_a_right_aligned_line_that_got_wider_is_hinted_as_wider_text(self) -> None:
        # One right-aligned paragraph with spaces between its two parts: the
        # end stays where it was, the start moves left.
        after = [line(140, 150, word(10, 30, "Next"))]
        reference = document(
            [[line(100, 110, word(218, 271, "申請人："))], [line(100, 110, word(426, 505, "（簽名蓋章）"))], after]
        )
        candidate = document(
            [[line(100, 110, word(205, 259, "申請人："))], [line(100, 110, word(421, 503, "（簽名蓋章）"))], after]
        )
        nodes = [{"address": "1/3", "styleName": "P32", "excerpt": layout_diff.normalize("申請人：　　（簽名蓋章）")}]
        (symptom,) = compare(reference, candidate, nodes)["symptoms"]
        self.assertEqual(symptom["kind"], "shifted-line-start")
        self.assertEqual(symptom["hint"], "wider-text")

    def test_an_unaligned_repeated_note_is_found_by_its_text_on_the_row(self) -> None:
        note = "（簽名蓋章）"
        reference = layout_diff.parse_layout(
            document([[line(100, 110, word(426, 505, note))], [line(120, 130, word(426, 505, note))]])
        )
        candidate = layout_diff.parse_layout(
            document([[line(120, 130, word(300, 380, note))], [line(100, 110, word(421, 503, note))]])
        )
        end = len(note) - 1
        found = layout_diff.same_text_on_row(reference, candidate, end, candidate.row_of(len(note)))
        self.assertEqual(found, 2 * len(note) - 1)

    def test_a_running_footer_is_not_reported_as_a_shift(self) -> None:
        reference = document(
            [[line(100, 110, word(10, 40, "Body"))], [line(800, 810, word(10, 50, "Footer1"))]],
            [[line(100, 110, word(10, 40, "Rest"))], [line(800, 810, word(10, 50, "Footer1"))]],
        )
        candidate = document(
            [[line(100, 110, word(10, 40, "Body"))], [line(800, 810, word(10, 50, "Footer1"))]],
            [[line(170, 180, word(10, 40, "Rest"))], [line(800, 810, word(10, 50, "Footer1"))]],
        )
        self.assertEqual(compare(reference, candidate)["symptoms"], [])


class MatchNodesTests(unittest.TestCase):
    def test_only_the_nodes_matching_the_most_keys_are_returned(self) -> None:
        nodes = [
            {"address": "2/10", "styleName": "P2", "excerpt": "Signature:"},
            {"address": "2/9", "styleName": "P1", "excerpt": "Signature:(note)"},
        ]
        self.assertEqual(
            layout_diff.match_nodes(nodes, ["(note)", "Signature:"]),
            [{"address": "2/9", "styleName": "P1"}],
        )


if __name__ == "__main__":
    unittest.main()
