# SPDX-License-Identifier: MPL-2.0
# Copyright 2026 OSSII
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
"""Unit tests for restoring the distributed alignment Word's .odt export drops."""
import pathlib
import tempfile
import unittest
import zipfile

from experiments import restore_distributed_alignment as alignment
from experiments import static_props as sp

W = sp.NS["w"]
O = sp.NS["office"]
S = sp.NS["style"]
X = sp.NS["text"]
F = sp.NS["fo"]

START = '<style:paragraph-properties fo:text-align="start"/>'
JUSTIFIED = '<style:paragraph-properties fo:text-align="justify" fo:text-align-last="justify"/>'
# The label as Word writes it, spread by the distributed alignment, and as the
# .odt holds it: the join has to see one text.
LABEL, ODT_LABEL = "姓　名", "姓名"


def word_paragraph(text: str, alignment_value: str = "") -> str:
    properties = f'<w:jc w:val="{alignment_value}"/>' if alignment_value else ""
    return f"<w:p><w:pPr>{properties}</w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>"


def word_document(*paragraphs: str) -> str:
    """word/document.xml, with the encoding declaration Word writes."""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W}"><w:body>{"".join(paragraphs)}</w:body></w:document>'
    )


def distributed(*paragraphs: str) -> list[tuple[str, bool]]:
    return alignment.word_paragraphs(word_document(*paragraphs).encode("utf-8"))


def style(name: str, properties: str = START) -> str:
    return (
        f'<style:style style:name="{name}" style:family="paragraph" style:parent-style-name="Standard">'
        f"{properties}</style:style>"
    )


def paragraph(name: str, text: str) -> str:
    return f'<text:p text:style-name="{name}">{text}</text:p>'


def odt_content(styles: str, body: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<office:document-content xmlns:office="{O}" xmlns:style="{S}" xmlns:text="{X}" xmlns:fo="{F}">'
        f"<office:automatic-styles>{styles}</office:automatic-styles>"
        f"<office:body><office:text>{body}</office:text></office:body></office:document-content>"
    )


def patched(name: str) -> str:
    return f"paragraph style {name}: text-align justify, text-align-last justify"


class WordTests(unittest.TestCase):
    def test_only_a_distribute_on_the_paragraph_itself_is_read(self) -> None:
        found = distributed(
            word_paragraph(LABEL, "distribute"),
            word_paragraph("備　註", "both"),
            word_paragraph("說明"),
        )
        self.assertEqual(found, [(ODT_LABEL, True), ("備註", False), ("說明", False)])


class ParagraphTests(unittest.TestCase):
    def test_a_self_closing_paragraph_is_a_paragraph_of_its_own(self) -> None:
        content = odt_content("", paragraph("P1", ODT_LABEL) + '<text:p text:style-name="PE"/>' + paragraph("P2", "職業"))
        self.assertEqual(
            [(found.style, found.text) for found in alignment.paragraphs(content)],
            [("P1", ODT_LABEL), ("PE", ""), ("P2", "職業")],
        )

    def test_a_paragraph_is_told_from_an_element_whose_name_starts_the_same(self) -> None:
        content = odt_content("", '<text:p text:style-name="P1">頁 <text:page-number>1</text:page-number></text:p>')
        self.assertEqual([(found.style, found.text) for found in alignment.paragraphs(content)], [("P1", "頁1")])


class RestoreTests(unittest.TestCase):
    def test_a_distributed_paragraph_gets_both_alignment_attributes(self) -> None:
        content = odt_content(style("P1"), paragraph("P1", ODT_LABEL))
        edited, changes = alignment.with_distributed_alignment(content, distributed(word_paragraph(LABEL, "distribute")))
        self.assertEqual(edited, odt_content(style("P1", JUSTIFIED), paragraph("P1", ODT_LABEL)))
        self.assertEqual(changes, [patched("P1")])

    def test_a_paragraph_that_was_not_distributed_is_left_byte_identical(self) -> None:
        word = distributed(word_paragraph(LABEL, "distribute"), word_paragraph("備註"))
        content = odt_content(style("P1") + style("P2"), paragraph("P1", ODT_LABEL) + paragraph("P2", "備註"))
        edited, changes = alignment.with_distributed_alignment(content, word)
        self.assertEqual(edited, odt_content(style("P1", JUSTIFIED) + style("P2"), paragraph("P1", ODT_LABEL) + paragraph("P2", "備註")))
        self.assertEqual(changes, [patched("P1")])

    def test_an_empty_paragraph_between_two_targets_does_not_shift_the_styles(self) -> None:
        # The trap: a regex that runs from <text:p to the next </text:p> takes
        # the empty paragraph's tag with the next paragraph's text.
        word = distributed(word_paragraph(LABEL, "distribute"), word_paragraph("職　業", "distribute"))
        body = (
            paragraph("P1", ODT_LABEL)
            + '<text:p text:style-name="PE"/>'
            + paragraph("P2", "職業")
            + '<text:p text:style-name="PE"/>'
        )
        styles = style("P1") + style("PE") + style("P2")
        edited, changes = alignment.with_distributed_alignment(odt_content(styles, body), word)
        self.assertEqual(edited, odt_content(style("P1", JUSTIFIED) + style("PE") + style("P2", JUSTIFIED), body))
        self.assertEqual(changes, [patched("P1"), patched("P2")])

    def test_a_style_without_paragraph_properties_gains_them_first(self) -> None:
        for styles, expected in (
            (
                '<style:style style:name="P1" style:family="paragraph"/>',
                f'<style:style style:name="P1" style:family="paragraph">{JUSTIFIED}</style:style>',
            ),
            (
                style("P1", '<style:text-properties fo:font-size="12pt"/>'),
                style("P1", JUSTIFIED + '<style:text-properties fo:font-size="12pt"/>'),
            ),
        ):
            content = odt_content(styles, paragraph("P1", ODT_LABEL))
            edited, changes = alignment.with_distributed_alignment(content, distributed(word_paragraph(LABEL, "distribute")))
            self.assertEqual(edited, odt_content(expected, paragraph("P1", ODT_LABEL)))
            self.assertEqual(changes, [patched("P1")])

    def test_a_paragraph_that_already_has_a_last_line_alignment_is_left_alone(self) -> None:
        properties = '<style:paragraph-properties fo:text-align="center" fo:text-align-last="center"/>'
        content = odt_content(style("P1", properties), paragraph("P1", ODT_LABEL))
        self.assertEqual(
            alignment.with_distributed_alignment(content, distributed(word_paragraph(LABEL, "distribute"))),
            (content, []),
        )

    def test_a_style_shared_with_a_paragraph_that_was_not_distributed_is_left_alone(self) -> None:
        # Patching it would spread that paragraph's text out as well.
        word = distributed(word_paragraph(LABEL, "distribute"), word_paragraph("備註"))
        content = odt_content(style("P1"), paragraph("P1", ODT_LABEL) + paragraph("P1", "備註"))
        self.assertEqual(
            alignment.with_distributed_alignment(content, word),
            (content, ["paragraph style P1: used by 2 paragraphs, 1 distributed; not patched"]),
        )

    def test_a_style_that_content_xml_does_not_hold_is_reported(self) -> None:
        content = odt_content("", paragraph("Standard", ODT_LABEL))
        self.assertEqual(
            alignment.with_distributed_alignment(content, distributed(word_paragraph(LABEL, "distribute"))),
            (content, ["paragraph style Standard: not a paragraph style in content.xml; not patched"]),
        )


class AmbiguityTests(unittest.TestCase):
    """Duplicate label text is the rule: several cells read 姓　名."""

    def test_every_copy_is_patched_when_every_copy_was_distributed(self) -> None:
        word = distributed(word_paragraph(LABEL, "distribute"), word_paragraph(LABEL, "distribute"))
        content = odt_content(style("P1") + style("P2"), paragraph("P1", ODT_LABEL) + paragraph("P2", ODT_LABEL))
        edited, changes = alignment.with_distributed_alignment(content, word)
        self.assertEqual(edited, odt_content(style("P1", JUSTIFIED) + style("P2", JUSTIFIED), paragraph("P1", ODT_LABEL) + paragraph("P2", ODT_LABEL)))
        self.assertEqual(changes, [patched("P1"), patched("P2")])

    def test_a_text_distributed_in_only_some_of_its_copies_is_left_alone(self) -> None:
        # Nothing in either file says which .odt copy is the distributed one.
        word = distributed(word_paragraph(LABEL, "distribute"), word_paragraph(ODT_LABEL))
        content = odt_content(style("P1") + style("P2"), paragraph("P1", ODT_LABEL) + paragraph("P2", ODT_LABEL))
        self.assertEqual(
            alignment.with_distributed_alignment(content, word),
            (content, [f'text "{ODT_LABEL}": distributed in 1 of 2 .docx paragraphs; not patched']),
        )

    def test_a_count_that_differs_between_the_two_files_is_left_alone(self) -> None:
        word = distributed(word_paragraph(LABEL, "distribute"), word_paragraph(LABEL, "distribute"))
        content = odt_content(style("P1"), paragraph("P1", ODT_LABEL))
        self.assertEqual(
            alignment.with_distributed_alignment(content, word),
            (content, [f'text "{ODT_LABEL}": 2 in the .docx, 1 in the .odt; not patched']),
        )


class PackageTests(unittest.TestCase):
    def test_only_content_xml_changes_and_the_mimetype_stays_first_and_stored(self) -> None:
        styles = f'<office:document-styles xmlns:office="{O}"/>'
        content = odt_content(style("P1") + style("P2"), paragraph("P1", ODT_LABEL) + paragraph("P2", "備註"))
        with tempfile.TemporaryDirectory() as directory:
            word = pathlib.Path(directory) / "word.docx"
            source = pathlib.Path(directory) / "word.odt"
            target = pathlib.Path(directory) / "restored.odt"
            with zipfile.ZipFile(word, "w", zipfile.ZIP_DEFLATED) as package:
                package.writestr(
                    "word/document.xml",
                    word_document(word_paragraph(LABEL, "distribute"), word_paragraph("備註")),
                )
            with zipfile.ZipFile(source, "w", zipfile.ZIP_DEFLATED) as package:
                package.writestr("content.xml", content)
                package.writestr("mimetype", "application/vnd.oasis.opendocument.text")
                package.writestr("styles.xml", styles)
            changes = alignment.restore(word, source, target)
            with zipfile.ZipFile(target) as package:
                infos = package.infolist()
                edited = package.read("content.xml").decode("utf-8")
                self.assertEqual([info.filename for info in infos], ["mimetype", "content.xml", "styles.xml"])
                self.assertEqual(infos[0].compress_type, zipfile.ZIP_STORED)
                self.assertEqual(package.read("styles.xml").decode("utf-8"), styles)
        self.assertEqual(edited, content.replace(style("P1"), style("P1", JUSTIFIED)))
        self.assertEqual(changes, [patched("P1")])

    def test_a_docx_without_a_document_part_is_refused(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            word = pathlib.Path(directory) / "word.docx"
            source = pathlib.Path(directory) / "word.odt"
            with zipfile.ZipFile(word, "w") as package:
                package.writestr("[Content_Types].xml", "<Types/>")
            with zipfile.ZipFile(source, "w") as package:
                package.writestr("content.xml", odt_content("", ""))
            with self.assertRaises(RuntimeError):
                alignment.restore(word, source, pathlib.Path(directory) / "restored.odt")


if __name__ == "__main__":
    unittest.main()
