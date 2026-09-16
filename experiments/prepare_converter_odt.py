#!/usr/bin/env python3
# SPDX-License-Identifier: MPL-2.0
# Copyright 2026 OSSII
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
"""Experiment: prepare an .odt produced by a LibreOffice-family converter.

The conversion tools used by government offices, and LibreOffice itself, write
a complete text grid, so there is nothing to fill in. What their `settings.xml`
does not carry is `AdjustTableLineHeightsToGridHeight`: that compatibility
option was only split off in 2025-07 (tdf#167583) with a default of true, and
the ODT import has no rule for the missing item, so table text is snapped to
the grid where Word's is not. In the corpus every converted document that had a
grid was missing the item.

The option does nothing without a grid, so the grid decides whether the package
is touched at all. Unlike a Word export this does not fill the grid and does not
touch font pitch: those are defects of that export, and these documents carry a
base height and a usable font pitch already. There is no re-save either -- the
package these tools write is already one LibreOffice reads back unchanged.

What it does share with the Word path is the padding clamp: a negative
`fo:padding` is invalid whatever wrote it, and a handful of these documents
carry one as authored.

    python3 experiments/prepare_converter_odt.py converted.odt prepared.odt
"""
import argparse
import json
import pathlib
import sys
import tempfile
import zipfile

try:
    from experiments import prepare_word_odt
except ImportError:  # run as a script, with experiments/ on the path
    import prepare_word_odt

STYLES = "styles.xml"


def has_grid(styles: str) -> bool:
    """Whether a page layout in `styles.xml` turns the text grid on."""
    return any(
        prepare_word_odt.attribute(properties, "style:layout-grid-mode") in ("line", "both")
        for properties in prepare_word_odt.PAGE_LAYOUT_PROPERTIES.findall(styles)
    )


def prepare(source: pathlib.Path, target: pathlib.Path) -> list[str]:
    """Writes `source` to `target` with the negative paddings clamped, and with
    the compatibility settings when there is a grid for them to act on.

    The clamp is not conditional on the grid. An invalid padding stops anything
    that validates a document before editing it, whatever the grid says, and
    these documents never went through a re-save -- what they carry is what
    their converter wrote.
    """
    with zipfile.ZipFile(source) as package:
        styles = package.read(STYLES).decode("utf-8") if STYLES in package.namelist() else ""
    if not has_grid(styles):
        return prepare_word_odt.clamp_padding(source, target)
    with tempfile.TemporaryDirectory() as directory:
        clamped = pathlib.Path(directory) / "clamped.odt"
        changes = prepare_word_odt.clamp_padding(source, clamped)
        changes.extend(prepare_word_odt.inject_settings(clamped, target))
    return changes


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("source", help=".odt from a LibreOffice-family converter")
    parser.add_argument("output", help="prepared .odt")
    arguments = parser.parse_args()
    changes = prepare(pathlib.Path(arguments.source), pathlib.Path(arguments.output))
    print(json.dumps({"changes": changes}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
