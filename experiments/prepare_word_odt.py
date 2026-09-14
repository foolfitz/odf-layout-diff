#!/usr/bin/env python3
"""Experiment: prepare an .odt exported by Microsoft Word for the fix loop.

Word writes a text grid (`style:layout-grid-mode`, `style:layout-grid-lines`)
without `style:layout-grid-base-height` and `style:layout-grid-ruby-height`.
LibreOffice then uses its defaults, 20 pt and 10 pt, so every line takes at
least 30 pt where Word's line pitch is typically about 18 pt, and the
document gains pages. The grid is filled in the way LibreOffice's DOCX
import sets it: the base height is the line pitch (the text area height
divided by the number of lines), the ruby height is 0, and a lines and
characters grid without a character pitch (`style:layout-grid-base-width`
of 0) becomes a lines-only grid.

A fixed pitch is dropped from font faces whose family name is not ASCII.
LibreOffice registers a font under the family name of its UI language only,
looks any other name up through fontconfig, and a fixed pitch there puts
monospace fonts first: with an English UI, 標楷體 is replaced.

LibreOffice then saves the result again. Word's .odt export does not satisfy
the ODF schema (its manifest has no `manifest:version`, for example), and
`odf-tool` refuses to edit such a document; the resaved one renders the same.

    python3 experiments/prepare_word_odt.py word.odt prepared.odt
"""
import argparse
import json
import pathlib
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

POINTS_PER_UNIT = {"pt": 1.0, "pc": 12.0, "in": 72.0, "cm": 72.0 / 2.54, "mm": 72.0 / 25.4}
LENGTH = re.compile(r"^(-?[0-9]*\.?[0-9]+)(pt|pc|in|cm|mm)$")
# Element names are followed by whitespace, not a word boundary: "-" would
# let <style:page-layout match <style:page-layout-properties.
PAGE_LAYOUT = re.compile(r"<style:page-layout\s[^>]*>.*?</style:page-layout>", re.S)
PAGE_LAYOUT_PROPERTIES = re.compile(r"<style:page-layout-properties\s[^>]*>")
HEADER_FOOTER_PROPERTIES = re.compile(r"<style:header-footer-properties\s[^>]*>")
FONT_FACE = re.compile(r"<style:font-face\s[^>]*>")


def attribute(tag: str, name: str) -> str | None:
    found = re.search(rf'\s{re.escape(name)}="([^"]*)"', tag)
    return found.group(1) if found else None


def points(length: str | None) -> float | None:
    found = LENGTH.match(length.strip()) if length else None
    return float(found.group(1)) * POINTS_PER_UNIT[found.group(2)] if found else None


def add_attributes(tag: str, attributes: str) -> str:
    ending = "/>" if tag.endswith("/>") else ">"
    return tag[: -len(ending)].rstrip() + attributes + ending


def fill_grid(styles: str) -> tuple[str, list[str]]:
    """`styles.xml` with each Word text grid given a base and ruby height."""
    changes: list[str] = []

    def page_layout(match: re.Match) -> str:
        block = match.group(0)
        name = attribute(block[: block.index(">") + 1], "style:name")
        found = PAGE_LAYOUT_PROPERTIES.search(block)
        if not found:
            return block
        properties = found.group(0)
        mode = attribute(properties, "style:layout-grid-mode")
        lines = attribute(properties, "style:layout-grid-lines")
        if mode not in ("line", "both") or not lines or int(lines) <= 0:
            return block
        if attribute(properties, "style:layout-grid-base-height") or attribute(properties, "style:layout-grid-ruby-height"):
            return block
        height = points(attribute(properties, "fo:page-height"))
        if height is None:
            return block
        margins = sum(points(attribute(properties, key)) or 0.0 for key in ("fo:margin-top", "fo:margin-bottom"))
        # Word's top and bottom margins start below the header and above the
        # footer; its export moves that space into their minimum heights.
        header_footer = sum(
            max(0.0, points(attribute(tag, "fo:min-height")) or 0.0) for tag in HEADER_FOOTER_PROPERTIES.findall(block)
        )
        pitch = (height - margins - header_footer) / int(lines)
        if pitch <= 0:
            return block
        edited = add_attributes(
            properties, f' style:layout-grid-base-height="{pitch:.2f}pt" style:layout-grid-ruby-height="0pt"'
        )
        new_mode = mode
        if mode == "both" and points(attribute(properties, "style:layout-grid-base-width")) == 0:
            edited = edited.replace('style:layout-grid-mode="both"', 'style:layout-grid-mode="line"')
            new_mode = "line"
        changes.append(f"page layout {name}: grid {mode}->{new_mode}, base height {pitch:.2f}pt, ruby height 0pt")
        return block.replace(properties, edited, 1)

    return PAGE_LAYOUT.sub(page_layout, styles), changes


def drop_fixed_pitch(xml: str) -> tuple[str, list[str]]:
    """The part with `style:font-pitch="fixed"` removed from font faces whose
    family name is not ASCII."""
    changes: list[str] = []

    def font_face(match: re.Match) -> str:
        tag = match.group(0)
        family = attribute(tag, "svg:font-family") or ""
        if attribute(tag, "style:font-pitch") != "fixed" or family.isascii():
            return tag
        changes.append(f"font face {attribute(tag, 'style:name')}: fixed pitch dropped")
        return re.sub(r'\sstyle:font-pitch="fixed"', "", tag)

    return FONT_FACE.sub(font_face, xml), changes


def edit_package(source: pathlib.Path, target: pathlib.Path) -> list[str]:
    """Writes `source` with the grid filled in and fixed pitches dropped to `target`."""
    changes: list[str] = []
    with zipfile.ZipFile(source) as package, zipfile.ZipFile(target, "w") as output:
        names = package.namelist()
        # The mimetype entry comes first and is stored uncompressed.
        for name in sorted(names, key=lambda item: item != "mimetype"):
            data = package.read(name)
            if name in ("styles.xml", "content.xml"):
                text, found = drop_fixed_pitch(data.decode("utf-8"))
                changes.extend(found)
                if name == "styles.xml":
                    text, found = fill_grid(text)
                    changes.extend(found)
                data = text.encode("utf-8")
            compression = zipfile.ZIP_STORED if name == "mimetype" else zipfile.ZIP_DEFLATED
            output.writestr(zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0)), data, compress_type=compression)
    return changes


def resave(source: pathlib.Path, target: pathlib.Path, soffice: str) -> None:
    with tempfile.TemporaryDirectory() as directory:
        outdir = pathlib.Path(directory)
        subprocess.run(
            [soffice, "--headless", "--convert-to", "odt", "--outdir", str(outdir), str(source)],
            capture_output=True,
            timeout=300,
            check=True,
        )
        saved = outdir / source.name
        if not saved.exists():
            raise RuntimeError(f"LibreOffice did not save {source.name} (is another instance running?)")
        shutil.copyfile(saved, target)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("source", help=".odt exported by Microsoft Word")
    parser.add_argument("output", help="prepared .odt")
    parser.add_argument("--soffice", default="soffice")
    parser.add_argument("--no-resave", action="store_true", help="only edit the package")
    arguments = parser.parse_args()
    output = pathlib.Path(arguments.output)
    with tempfile.TemporaryDirectory() as directory:
        edited = pathlib.Path(directory) / "edited.odt"
        changes = edit_package(pathlib.Path(arguments.source), edited)
        if arguments.no_resave:
            shutil.copyfile(edited, output)
        else:
            resave(edited, output, arguments.soffice)
    print(json.dumps({"changes": changes, "resaved": not arguments.no_resave}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
