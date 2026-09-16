#!/usr/bin/env python3
# SPDX-License-Identifier: MPL-2.0
# Copyright 2026 OSSII
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
"""Experiment: restore the distributed alignment Word's .odt export drops.

Word's `<w:jc w:val="distribute"/>` spreads a short label over the whole line,
which is how a form cell reads 姓　名 rather than 姓名. Its .odt export writes
that as `fo:text-align="start"` (sometimes `"justify"`) and never writes
`fo:text-align-last`, so LibreOffice collapses the label to the left of the
cell. Five Word-exported .odt files in the study carried no `fo:text-align-last`
at all, while LibreOffice's own DOCX and DOC import keeps the alignment
(`DomainMapper` turns `ST_Jc_distribute` into a last line adjustment of BLOCK):
the defect is in Word's export. Giving the affected paragraphs
`fo:text-align="justify"` with `fo:text-align-last="justify"` was measured to
restore the layout; `style:justify-single-word` alone does nothing and is not
written.

The exported .odt does not say which paragraphs were distributed -- `"start"`
is what any left-aligned paragraph carries -- so the Word original is the
source of truth, and only a `.docx` is read. A `.doc` is a binary format this
does not parse: there is no partial support for it.

The two files are joined on paragraph text alone, NFKC-normalized with all
whitespace removed, so that Word's 姓　名 and the .odt's `姓<text:s/>名` are one
label. Form labels repeat, though, and several cells reading 姓　名 cannot be
told apart by text. A text is therefore patched only when every `.docx`
paragraph carrying it was distributed and both files hold the same number of
paragraphs with it; then every `.odt` paragraph with that text is patched, and
which copy is which never has to be guessed. A text distributed in some copies
only, or one whose counts differ, is left alone and reported -- patching the
wrong paragraph is worse than patching none. Paragraphs with no text are not
matched, and a style shared with a paragraph that was not distributed is left
alone too, since patching it would align that paragraph as well. Only `<text:p>`
is matched and only paragraph styles in `content.xml` are patched; anything
else is reported, again rather than guessed.

This is independent of `prepare_word_odt.py` and can run before or after it:
the match is on text, so a re-save that renames the automatic styles does not
affect it.

    python3 experiments/restore_distributed_alignment.py word.docx word.odt restored.odt
"""
import argparse
import html
import json
import pathlib
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from dataclasses import dataclass

try:
    from experiments import prepare_word_odt, static_props
except ImportError:  # run as a script, with experiments/ on the path
    import prepare_word_odt
    import static_props

CONTENT = "content.xml"
DOCUMENT = "word/document.xml"
PARAGRAPH = "text:p"
STYLE = "style:style"
PARAGRAPH_PROPERTIES = "style:paragraph-properties"
ALIGNMENT = {"fo:text-align": "justify", "fo:text-align-last": "justify"}
DISTRIBUTE = "distribute"
WORD_PARAGRAPH = static_props.q("w", "p")
WORD_PROPERTIES = static_props.q("w", "pPr")
WORD_ALIGNMENT = static_props.q("w", "jc")
WORD_TEXT = static_props.q("w", "t")
WORD_VALUE = static_props.q("w", "val")
TAG = re.compile(r"<[^>]*>")


@dataclass(frozen=True)
class Element:
    """One element of an XML part: its opening tag, where that tag starts and
    ends, the markup it holds and where the element ends."""

    tag: str
    start: int
    open_end: int
    inner: str
    end: int


@dataclass(frozen=True)
class Paragraph:
    element: Element
    style: str | None
    text: str


def elements(xml: str, name: str) -> list[Element]:
    """Every `name` element of `xml` in document order, nested ones included.

    A self-closing tag is a whole element. Matching `<text:p` up to the next
    `>` and then on to the next `</text:p>` instead swallows `<text:p .../>`:
    the element found then carries the empty paragraph's tag and the *next*
    paragraph's text. That is how an earlier attempt at this rule patched the
    paragraph before the one it had matched, and measured no effect twice.
    """
    token = re.compile(rf"<{re.escape(name)}(?=[\s/>])[^>]*>|</{re.escape(name)}\s*>")
    open_tags: list[tuple[int, int, str]] = []
    found: list[Element] = []
    for match in token.finditer(xml):
        tag = match.group(0)
        if tag.startswith("</"):
            if not open_tags:
                raise RuntimeError(f"</{name}> without a <{name}>")
            start, open_end, opening = open_tags.pop()
            found.append(Element(opening, start, open_end, xml[open_end : match.start()], match.end()))
        elif tag.endswith("/>"):
            found.append(Element(tag, match.start(), match.end(), "", match.end()))
        else:
            open_tags.append((match.start(), match.end(), tag))
    if open_tags:
        raise RuntimeError(f"<{name}> without a </{name}>")
    return sorted(found, key=lambda element: element.start)


def text_of(xml: str) -> str:
    """The text of a fragment, normalized so that the two formats compare: NFKC
    and no whitespace, which is what makes 姓　名 and `姓<text:s/>名` one label."""
    return static_props.normalize_text(html.unescape(TAG.sub("", xml)))


def word_paragraphs(document: bytes) -> list[tuple[str, bool]]:
    """The text of every paragraph of `word/document.xml`, in document order,
    and whether its paragraph properties carry `<w:jc w:val="distribute"/>`.

    Only a `w:jc` of the paragraph itself counts. An alignment inherited from a
    paragraph style is not resolved, so such a paragraph is reported as missing
    from the .odt rather than patched.
    """
    root = ET.fromstring(document)
    found: list[tuple[str, bool]] = []
    for paragraph in root.iter(WORD_PARAGRAPH):
        properties = paragraph.find(WORD_PROPERTIES)
        alignment = properties.find(WORD_ALIGNMENT) if properties is not None else None
        distributed = alignment is not None and alignment.get(WORD_VALUE) == DISTRIBUTE
        text = "".join(node.text or "" for node in paragraph.iter(WORD_TEXT))
        found.append((static_props.normalize_text(text), distributed))
    return found


def paragraphs(content: str) -> list[Paragraph]:
    """Every `<text:p>` of `content.xml` with its style name and its text."""
    return [
        Paragraph(element, prepare_word_odt.attribute(element.tag, "text:style-name"), text_of(element.inner))
        for element in elements(content, PARAGRAPH)
    ]


def targets(found: list[Paragraph], word: list[tuple[str, bool]]) -> tuple[list[Paragraph], list[str]]:
    """The paragraphs of `content.xml` whose text was distributed throughout the
    .docx, and a note for every text that was not taken."""
    total = Counter(text for text, _ in word if text)
    distributed = Counter(text for text, is_distributed in word if text and is_distributed)
    copies: dict[str, list[Paragraph]] = {}
    for paragraph in found:
        if paragraph.text:
            copies.setdefault(paragraph.text, []).append(paragraph)
    matched: list[Paragraph] = []
    notes: list[str] = []
    for text in sorted(distributed):
        here = copies.get(text, [])
        label = static_props.text_label(text)
        if distributed[text] != total[text]:
            # Copies of one label that were not all distributed: only their
            # order could tell them apart, and nothing here checks that order.
            notes.append(f'text "{label}": distributed in {distributed[text]} of {total[text]} .docx paragraphs; not patched')
        elif len(here) != total[text]:
            notes.append(f'text "{label}": {total[text]} in the .docx, {len(here)} in the .odt; not patched')
        else:
            matched.extend(here)
    return matched, notes


def target_styles(found: list[Paragraph], matched: list[Paragraph]) -> tuple[list[str], list[str]]:
    """The style names to patch, and a note for every style left alone."""
    users = Counter(paragraph.style for paragraph in found)
    wanted = Counter(paragraph.style for paragraph in matched)
    names: list[str] = []
    notes: list[str] = []
    for name in sorted(wanted, key=lambda style: (style is None, style or "")):
        if name is None:
            notes.append(f"distributed paragraphs without a text:style-name: {wanted[None]}; not patched")
        elif users[name] != wanted[name]:
            notes.append(f"paragraph style {name}: used by {users[name]} paragraphs, {wanted[name]} distributed; not patched")
        else:
            names.append(name)
    return names, notes


def set_attribute(tag: str, name: str, value: str) -> str:
    """`tag` with `name` set to `value`. Word writes `fo:text-align="start"`
    where the paragraph was distributed, so an existing value is replaced."""
    if prepare_word_odt.attribute(tag, name) is None:
        return prepare_word_odt.add_attributes(tag, f' {name}="{value}"')
    return re.sub(rf'(\s{re.escape(name)}=")[^"]*"', lambda match: match.group(1) + value + '"', tag, count=1)


def restored(content: str, element: Element) -> str | None:
    """The `<style:style>` at `element` with both alignment attributes on its
    paragraph properties, or None when it already carries `fo:text-align-last`:
    a last line alignment that is already there was meant, and is left alone."""
    style = content[element.start : element.end]
    properties = elements(style, PARAGRAPH_PROPERTIES)
    if properties:
        first = properties[0]
        if prepare_word_odt.attribute(first.tag, "fo:text-align-last") is not None:
            return None
        tag = first.tag
        for name, value in ALIGNMENT.items():
            tag = set_attribute(tag, name, value)
        return style[: first.start] + tag + style[first.open_end :]
    added = "<{}{}/>".format(PARAGRAPH_PROPERTIES, "".join(f' {name}="{value}"' for name, value in ALIGNMENT.items()))
    if element.tag.endswith("/>"):
        # A style with no properties at all: it needs a body to hold them.
        return f"{element.tag[:-2].rstrip()}>{added}</{STYLE}>"
    # A style's paragraph properties come before its text properties.
    opening = element.open_end - element.start
    return style[:opening] + added + style[opening:]


def with_distributed_alignment(content: str, word: list[tuple[str, bool]]) -> tuple[str, list[str]]:
    """`content.xml` with the paragraph styles of the paragraphs that were
    distributed in the .docx given both alignment attributes.

    The notes of everything left alone follow the changes in the same list.
    """
    found = paragraphs(content)
    matched, notes = targets(found, word)
    names, refused = target_styles(found, matched)
    notes.extend(refused)
    wanted, seen = set(names), set()
    changes: list[str] = []
    patches: list[tuple[Element, str]] = []
    for element in elements(content, STYLE):
        name = prepare_word_odt.attribute(element.tag, "style:name")
        if name not in wanted or prepare_word_odt.attribute(element.tag, "style:family") != "paragraph":
            continue
        seen.add(name)
        replacement = restored(content, element)
        if replacement is None:
            continue
        patches.append((element, replacement))
        changes.append(f"paragraph style {name}: text-align justify, text-align-last justify")
    notes.extend(f"paragraph style {name}: not a paragraph style in {CONTENT}; not patched" for name in sorted(wanted - seen))
    for element, replacement in reversed(patches):
        content = content[: element.start] + replacement + content[element.end :]
    return content, changes + notes


def restore(word: pathlib.Path, source: pathlib.Path, target: pathlib.Path) -> list[str]:
    """Writes `source`, an .odt Word exported from `word`, with its distributed
    alignment restored to `target`."""
    with zipfile.ZipFile(word) as package:
        if DOCUMENT not in package.namelist():
            raise RuntimeError(f"{word.name} has no {DOCUMENT}")
        distributed = word_paragraphs(package.read(DOCUMENT))
    with zipfile.ZipFile(source) as package:
        if CONTENT not in package.namelist():
            raise RuntimeError(f"{source.name} has no {CONTENT}")

    def edit(name: str, data: bytes) -> tuple[bytes, list[str]]:
        if name != CONTENT:
            return data, []
        text, changes = with_distributed_alignment(data.decode("utf-8"), distributed)
        return text.encode("utf-8"), changes

    return prepare_word_odt.rewrite_package(source, target, edit)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("word", help=".docx the .odt was exported from")
    parser.add_argument("source", help=".odt exported by Microsoft Word")
    parser.add_argument("output", help="restored .odt")
    arguments = parser.parse_args()
    changes = restore(
        pathlib.Path(arguments.word), pathlib.Path(arguments.source), pathlib.Path(arguments.output)
    )
    print(json.dumps({"changes": changes}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
