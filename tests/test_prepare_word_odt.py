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
ADJUST = "AdjustTableLineHeightsToGridHeight"
OOO = ' xmlns:ooo="http://openoffice.org/2004/office"'
# Word's settings.xml: a self-closing root, and no xmlns:ooo to resolve the
# QName the settings block is named with.
WORD_SETTINGS = (
    '<?xml version="1.0" encoding="UTF-8"?>'
    '<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"'
    ' xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" office:version="1.2"/>'
)


def item(name: str, value: str, kind: str = "boolean") -> str:
    return f'<config:config-item config:name="{name}" config:type="{kind}">{value}</config:config-item>'


def office_settings(items: str, namespace: str = OOO) -> str:
    """A settings.xml of the shape LibreOffice writes."""
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"'
        ' xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"'
        f'{namespace} office:version="1.3"><office:settings>'
        '<config:config-item-set config:name="ooo:view-settings">'
        f'{item("ViewAreaTop", "0", "int")}</config:config-item-set>'
        '<config:config-item-set config:name="ooo:configuration-settings">'
        f"{items}</config:config-item-set></office:settings></office:document-settings>"
    )


def word_package(path: pathlib.Path, settings: str = WORD_SETTINGS) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as package:
        package.writestr("content.xml", "<office:document-content/>")
        package.writestr("mimetype", "application/vnd.oasis.opendocument.text")
        package.writestr("styles.xml", page_layout(WORD_GRID))
        package.writestr("settings.xml", settings)


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


class SettingsTests(unittest.TestCase):
    def test_a_self_closing_root_gains_the_namespace_and_the_item(self) -> None:
        edited, changes = prepare_word_odt.with_compat_settings(WORD_SETTINGS)
        root = prepare_word_odt.SETTINGS_ROOT.search(edited)
        # Without the declaration on the root the whole file is ignored, so the
        # item alone would not be enough.
        self.assertIn('xmlns:ooo="http://openoffice.org/2004/office"', root.group(0))
        self.assertIn(item(ADJUST, "false"), edited)
        self.assertEqual(changes, ["settings.xml: xmlns:ooo declared", f"settings.xml: {ADJUST}=false"])

    def test_an_item_that_snaps_table_text_to_the_grid_is_turned_off(self) -> None:
        edited, changes = prepare_word_odt.with_compat_settings(office_settings(item(ADJUST, "true")))
        self.assertIn(item(ADJUST, "false"), edited)
        self.assertNotIn(item(ADJUST, "true"), edited)
        self.assertEqual(changes, [f"settings.xml: {ADJUST}=false"])

    def test_settings_that_already_carry_the_item_are_left_alone(self) -> None:
        original = office_settings(item(ADJUST, "false"))
        self.assertEqual(prepare_word_odt.with_compat_settings(original), (original, []))


class OrderTests(unittest.TestCase):
    def test_the_settings_go_into_the_package_the_resave_produced(self) -> None:
        seen: list[str] = []

        def resave(source: pathlib.Path, target: pathlib.Path) -> None:
            """LibreOffice, which rewrites settings.xml with the item back on."""
            with zipfile.ZipFile(source) as package:
                seen.append(package.read("styles.xml").decode())
                parts = {name: package.read(name) for name in package.namelist()}
            parts["settings.xml"] = office_settings(
                item(ADJUST, "true") + item("PrinterName", "resaved", "string")
            ).encode("utf-8")
            with zipfile.ZipFile(target, "w") as output:
                for name, data in parts.items():
                    output.writestr(name, data)

        with tempfile.TemporaryDirectory() as directory:
            source, target = pathlib.Path(directory) / "word.odt", pathlib.Path(directory) / "prepared.odt"
            word_package(source)
            changes = prepare_word_odt.prepare(source, target, resave)
            with zipfile.ZipFile(target) as package:
                settings = package.read("settings.xml").decode()
                styles = package.read("styles.xml").decode()
        # The settings were injected into what the re-save produced, not before it.
        self.assertIn('config:name="PrinterName"', settings)
        self.assertIn(item(ADJUST, "false"), settings)
        # And the grid was filled before the re-save, which would otherwise have
        # frozen LibreOffice's own default height.
        self.assertIn("layout-grid-base-height", seen[0])
        self.assertIn("layout-grid-base-height", styles)
        self.assertIn(f"settings.xml: {ADJUST}=false", changes)

    def test_without_a_resave_the_package_keeps_the_edit_and_the_settings(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            source, target = pathlib.Path(directory) / "word.odt", pathlib.Path(directory) / "prepared.odt"
            word_package(source)
            prepare_word_odt.prepare(source, target)
            with zipfile.ZipFile(target) as package:
                infos = package.infolist()
                settings = package.read("settings.xml").decode()
                self.assertEqual(infos[0].filename, "mimetype")
                self.assertEqual(infos[0].compress_type, zipfile.ZIP_STORED)
                self.assertIn("layout-grid-base-height", package.read("styles.xml").decode())
        self.assertIn(item(ADJUST, "false"), settings)
        self.assertIn('xmlns:ooo="http://openoffice.org/2004/office"', settings)


if __name__ == "__main__":
    unittest.main()
