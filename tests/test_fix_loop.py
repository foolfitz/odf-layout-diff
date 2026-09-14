"""Unit tests for the fix loop's decisions, without a model or a renderer."""
import unittest

from experiments import fix_loop


def report(symptoms: list[dict], rows_on_another_page: int = 0) -> dict:
    overflow = [{"page": 1, "candidatePage": 2, "rowCount": rows_on_another_page, "rows": []}]
    return {"rowsOnAnotherPage": overflow if rows_on_another_page else [], "symptoms": symptoms}


def block(text: str, height_change: float, hint: str = "wider-text") -> dict:
    return {
        "kind": "text-block-taller",
        "page": 1,
        "y": 100.0,
        "text": text,
        "heightChangePt": height_change,
        "breaks": [{"hint": hint}],
        "nodes": [{"address": "2/1", "styleName": "P1"}],
    }


def moved_line(text: str, shift: float, address: str = "2/2") -> dict:
    return {
        "kind": "shifted-line-start",
        "page": 1,
        "y": 200.0,
        "text": text,
        "shiftPt": shift,
        "hint": "shifted-start",
        "nodes": [{"address": address, "styleName": "P2"}],
    }


def shift_below(address: str, shift: float) -> dict:
    return {
        "kind": "vertical-shift",
        "page": 1,
        "y": 300.0,
        "row": "Next",
        "shiftPt": shift,
        "breaksInRowAbove": [{"hint": "wide-gap"}],
        "nodes": [{"address": address, "styleName": "P1"}],
    }


class HintTests(unittest.TestCase):
    def test_hints_allow_only_the_keys_that_explain_them(self) -> None:
        self.assertEqual(fix_loop.allowed_keys(block("a", 5, "wide-gap")), [fix_loop.LETTER_SPACING])
        self.assertEqual(fix_loop.allowed_keys(moved_line("b", -20)), [fix_loop.TEXT_INDENT, fix_loop.MARGIN_LEFT])

    def test_a_symptom_without_a_known_cause_or_node_is_not_offered(self) -> None:
        unknown = block("a", 5, "unknown")
        without_node = dict(block("b", 5), nodes=[])
        offered = block("c", 5)
        self.assertEqual(fix_loop.fixable(report([unknown, without_node, offered])), [offered])


class JudgeTests(unittest.TestCase):
    def test_a_smaller_symptom_is_kept(self) -> None:
        verdict = fix_loop.judge(report([block("a", 8.8)]), report([block("a", 1.9)]))
        self.assertEqual(verdict, {"kept": True, "reason": None, "newSymptoms": []})

    def test_overcorrecting_past_zero_is_no_improvement(self) -> None:
        verdict = fix_loop.judge(report([block("a", -1.3)]), report([block("a", -7.2)]))
        self.assertEqual(verdict["reason"], "NO_IMPROVEMENT")

    def test_text_of_the_same_paragraph_moving_sideways_is_a_new_symptom(self) -> None:
        # A negative margin removes the extra line but pushes the text out of its cell.
        verdict = fix_loop.judge(report([shift_below("2/1", 24.0)]), report([moved_line("sig", -21.3, "2/1")]))
        self.assertFalse(verdict["kept"])
        self.assertEqual(verdict["reason"], "NEW_SYMPTOM")
        self.assertEqual(verdict["newSymptoms"], [{"kind": "shifted-line-start", "page": 1, "text": "sig", "pt": 21.3}])

    def test_a_symptom_changing_kind_at_the_same_paragraph_is_not_new(self) -> None:
        # Tighter spacing moves the wrap: the shift below becomes a taller block.
        verdict = fix_loop.judge(report([shift_below("2/1", 24.0)]), report([block("(note)", 22.5)]))
        self.assertTrue(verdict["kept"])

    def test_a_symptom_at_another_paragraph_is_new(self) -> None:
        verdict = fix_loop.judge(report([shift_below("2/9", 24.0)]), report([block("(note)", 2.0)]))
        self.assertEqual(verdict["reason"], "NEW_SYMPTOM")

    def test_more_rows_on_another_page_reverts(self) -> None:
        verdict = fix_loop.judge(report([block("a", 20.0)], 1), report([], 2))
        self.assertEqual(verdict["reason"], "MORE_ROWS_ON_ANOTHER_PAGE")


class CheckChangeTests(unittest.TestCase):
    def change(self, **overrides: object) -> dict:
        change = {
            "symptomId": 1,
            "address": "2/1",
            "propertyKey": fix_loop.LETTER_SPACING,
            "propertyValue": "-0.02cm",
        }
        return {**change, **overrides}

    def test_a_change_must_match_the_symptom_it_names(self) -> None:
        shown = [block("a", 5, "wide-gap")]
        self.assertIsNone(fix_loop.check_change(self.change(), shown, set()))
        self.assertEqual(fix_loop.check_change(self.change(symptomId=2), shown, set()), "REJECTED_UNKNOWN_SYMPTOM")
        self.assertEqual(fix_loop.check_change(self.change(address="2/9"), shown, set()), "REJECTED_ADDRESS_NOT_OFFERED")
        self.assertEqual(
            fix_loop.check_change(self.change(propertyKey=fix_loop.MARGIN_LEFT), shown, set()),
            "REJECTED_KEY_NOT_ALLOWED",
        )

    def test_a_change_already_tried_is_rejected(self) -> None:
        tried = {("2/1", fix_loop.LETTER_SPACING, "-0.02cm")}
        self.assertEqual(fix_loop.check_change(self.change(), [block("a", 5)], tried), "REJECTED_ALREADY_TRIED")


if __name__ == "__main__":
    unittest.main()
