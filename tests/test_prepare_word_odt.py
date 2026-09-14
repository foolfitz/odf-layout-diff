"""Unit tests for preparing a Word-exported .odt, without LibreOffice."""
import pathlib
import tempfile
import unittest
import zipfile

from experiments import prepare_word_odt


def page_layout(properties: str, header_footer: str = "") -> str:
    return (
        '<office:automatic-styles><style:page-layout style:name="PL0">'
        f"<style:page-layout-properties {properties}/>{header_footer}"
        "</style:page-layout></office:automatic-styles>"
    )


WORD_GRID = (
    'fo:page-height="842pt" fo:margin-top="40pt" fo:margin-bottom="60pt" '
    'style:layout-grid-mode="line" style:layout-grid-lines="40" style:layout-grid-base-width="0in"'
)


class GridTests(unittest.TestCase):
    def test_a_grid_without_heights_gets_the_line_pitch_and_no_ruby(self) -> None:
        header = (
            '<style:header-style><style:header-footer-properties fo:min-height="22pt"/></style:header-style>'
            '<style:footer-style><style:header-footer-properties fo:min-height="-5pt"/></style:footer-style>'
        )
        styles, changes = prepare_word_odt.fill_grid(page_layout(WORD_GRID, header))
        # (842 - 40 - 60 - 22) / 40; a negative footer height takes no space.
        self.assertIn('style:layout-grid-base-height="18.00pt" style:layout-grid-ruby-height="0pt"/>', styles)
        self.assertEqual(changes, ["page layout PL0: grid line->line, base height 18.00pt, ruby height 0pt"])

    def test_the_default_page_layout_before_it_does_not_hide_a_page_layout(self) -> None:
        default = (
            "<office:styles><style:default-page-layout>"
            '<style:page-layout-properties style:layout-grid-standard-mode="true"/>'
            "</style:default-page-layout></office:styles>"
        )
        styles, changes = prepare_word_odt.fill_grid(default + page_layout(WORD_GRID))
        self.assertIn("layout-grid-base-height", styles)
        self.assertEqual(len(changes), 1)

    def test_a_lines_and_characters_grid_without_a_character_pitch_becomes_lines_only(self) -> None:
        both = WORD_GRID.replace('"line"', '"both"')
        styles, _ = prepare_word_odt.fill_grid(page_layout(both))
        self.assertIn('style:layout-grid-mode="line"', styles)
        with_pitch = both.replace('base-width="0in"', 'base-width="0.5cm"')
        styles, _ = prepare_word_odt.fill_grid(page_layout(with_pitch))
        self.assertIn('style:layout-grid-mode="both"', styles)

    def test_a_grid_with_a_height_or_no_grid_is_left_alone(self) -> None:
        for properties in (
            WORD_GRID + ' style:layout-grid-base-height="0.5cm"',
            WORD_GRID.replace('"line"', '"none"'),
        ):
            original = page_layout(properties)
            self.assertEqual(prepare_word_odt.fill_grid(original), (original, []))


class FontTests(unittest.TestCase):
    def test_a_fixed_pitch_is_dropped_only_for_a_non_ascii_family_name(self) -> None:
        xml = (
            '<style:font-face style:name="標楷體" svg:font-family="標楷體" style:font-pitch="fixed"/>'
            '<style:font-face style:name="Courier New" svg:font-family="&apos;Courier New&apos;" style:font-pitch="fixed"/>'
        )
        edited, changes = prepare_word_odt.drop_fixed_pitch(xml)
        self.assertEqual(
            edited,
            '<style:font-face style:name="標楷體" svg:font-family="標楷體"/>'
            '<style:font-face style:name="Courier New" svg:font-family="&apos;Courier New&apos;" style:font-pitch="fixed"/>',
        )
        self.assertEqual(changes, ["font face 標楷體: fixed pitch dropped"])


class PackageTests(unittest.TestCase):
    def test_the_package_keeps_its_parts_with_the_mimetype_first_and_stored(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source, target = pathlib.Path(directory) / "word.odt", pathlib.Path(directory) / "edited.odt"
            with zipfile.ZipFile(source, "w", zipfile.ZIP_DEFLATED) as package:
                package.writestr("content.xml", "<office:document-content/>")
                package.writestr("mimetype", "application/vnd.oasis.opendocument.text")
                package.writestr("styles.xml", page_layout(WORD_GRID))
            changes = prepare_word_odt.edit_package(source, target)
            with zipfile.ZipFile(target) as package:
                infos = package.infolist()
                self.assertEqual([info.filename for info in infos], ["mimetype", "content.xml", "styles.xml"])
                self.assertEqual(infos[0].compress_type, zipfile.ZIP_STORED)
                self.assertIn("layout-grid-base-height", package.read("styles.xml").decode())
            self.assertEqual(len(changes), 1)


if __name__ == "__main__":
    unittest.main()
