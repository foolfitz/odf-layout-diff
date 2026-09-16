# SPDX-License-Identifier: MPL-2.0
# Copyright 2026 OSSII
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
"""Unit tests for preparing an .odt from a LibreOffice-family converter."""
import pathlib
import tempfile
import unittest
import zipfile

from experiments import prepare_converter_odt
from tests.test_prepare_word_odt import ADJUST, WORD_GRID, item, office_settings, page_layout

# A converter writes the grid out in full: only the settings are short of an item.
GRID = page_layout(
    'fo:page-height="842pt" style:layout-grid-mode="line" style:layout-grid-lines="29" '
    'style:layout-grid-base-height="29.14pt" style:layout-grid-ruby-height="0pt"'
)
NO_GRID = GRID.replace('"line"', '"none"')
OTHER_ITEM = item("AddExternalLeading", "false")


def converted_package(path: pathlib.Path, styles: str, settings: str, content: str = "<office:document-content/>") -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as package:
        package.writestr("content.xml", content)
        package.writestr("mimetype", "application/vnd.oasis.opendocument.text")
        package.writestr("styles.xml", styles)
        package.writestr("settings.xml", settings)


def parts(path: pathlib.Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as package:
        return {name: package.read(name) for name in package.namelist()}


class GridTests(unittest.TestCase):
    def test_only_a_line_or_character_grid_counts(self) -> None:
        self.assertTrue(prepare_converter_odt.has_grid(GRID))
        self.assertTrue(prepare_converter_odt.has_grid(GRID.replace('"line"', '"both"')))
        self.assertFalse(prepare_converter_odt.has_grid(NO_GRID))
        self.assertFalse(prepare_converter_odt.has_grid(""))


class PrepareTests(unittest.TestCase):
    def prepare(self, styles: str, settings: str, content: str = "<office:document-content/>") -> tuple[list[str], dict[str, bytes]]:
        with tempfile.TemporaryDirectory() as directory:
            source = pathlib.Path(directory) / "converted.odt"
            target = pathlib.Path(directory) / "prepared.odt"
            converted_package(source, styles, settings, content)
            changes = prepare_converter_odt.prepare(source, target)
            return changes, parts(target)

    def test_a_grid_without_the_item_has_the_table_line_heights_detached(self) -> None:
        changes, prepared = self.prepare(GRID, office_settings(OTHER_ITEM))
        self.assertEqual(changes, [f"settings.xml: {ADJUST}=false"])
        settings = prepared["settings.xml"].decode()
        self.assertIn(item(ADJUST, "false"), settings)
        self.assertIn(OTHER_ITEM, settings)

    def test_a_document_without_a_grid_is_left_alone(self) -> None:
        # The option does nothing without a grid, so no part is edited. The
        # package is still rewritten on the way through the clamp, so this
        # compares the parts, not the bytes of the container.
        settings = office_settings(OTHER_ITEM)
        changes, prepared = self.prepare(NO_GRID, settings)
        self.assertEqual(changes, [])
        self.assertEqual(prepared, parts_of(NO_GRID, settings))

    def test_an_item_that_is_already_off_is_left_alone(self) -> None:
        settings = office_settings(item(ADJUST, "false"))
        changes, prepared = self.prepare(GRID, settings)
        self.assertEqual(changes, [])
        self.assertEqual(prepared["settings.xml"].decode(), settings)

    def test_the_grid_and_the_fonts_are_not_touched(self) -> None:
        # A Word-shaped defect in a converted document is still not this
        # script's business: it writes settings.xml and nothing else.
        font = '<style:font-face style:name="標楷體" svg:font-family="標楷體" style:font-pitch="fixed"/>'
        styles, content = font + page_layout(WORD_GRID), "<office:document-content>" + font + "</office:document-content>"
        changes, prepared = self.prepare(styles, office_settings(OTHER_ITEM), content)
        self.assertEqual(changes, [f"settings.xml: {ADJUST}=false"])
        self.assertEqual(prepared["styles.xml"].decode(), styles)
        self.assertEqual(prepared["content.xml"].decode(), content)

    def test_a_negative_padding_is_clamped_here_too(self) -> None:
        content = (
            "<office:document-content>"
            '<style:graphic-properties fo:padding-top="-0.004cm"/>'
            "</office:document-content>"
        )
        changes, prepared = self.prepare(GRID, office_settings(OTHER_ITEM), content)
        self.assertIn('fo:padding-top="0cm"', prepared["content.xml"].decode())
        self.assertIn("content.xml: 1 negative paddings clamped to 0", changes)

    def test_a_negative_padding_is_clamped_without_a_grid(self) -> None:
        """The clamp is not gated on the grid, and not on a re-save either.

        A converted document with no grid has nothing for the settings item to
        act on, but an invalid padding still stops anything that validates a
        document before editing it. Two documents in the corpus carry one as
        authored, with no re-save anywhere in their history -- so gating the
        clamp on either condition would miss exactly those.
        """
        content = (
            "<office:document-content>"
            '<style:graphic-properties fo:padding="-1pt"/>'
            "</office:document-content>"
        )
        changes, prepared = self.prepare(NO_GRID, office_settings(OTHER_ITEM), content)
        self.assertIn('fo:padding="0cm"', prepared["content.xml"].decode())
        self.assertEqual(changes, ["content.xml: 1 negative paddings clamped to 0"])


def parts_of(styles: str, settings: str) -> dict[str, bytes]:
    with tempfile.TemporaryDirectory() as directory:
        path = pathlib.Path(directory) / "converted.odt"
        converted_package(path, styles, settings)
        return parts(path)


if __name__ == "__main__":
    unittest.main()
