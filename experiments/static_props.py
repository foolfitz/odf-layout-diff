#!/usr/bin/env python3
# SPDX-License-Identifier: MPL-2.0
# Copyright 2026 OSSII
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
"""Static OOXML/ODF table and page-layout property inspector.

This intentionally reads package XML only.  It does not render documents and it
does not guess format defaults: raw values accompany every converted value.
"""

from __future__ import annotations

import argparse
import difflib
import json
import re
import sys
import unicodedata
from collections import Counter
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "fo": "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
    "config": "urn:oasis:names:tc:opendocument:xmlns:config:1.0",
    "svg": "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
    "loext": "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0",
}
URI_PREFIX = {v: k for k, v in NS.items()}


def q(prefix: str, local: str) -> str:
    return "{%s}%s" % (NS[prefix], local)


def short_qname(name: str) -> str:
    if name.startswith("{"):
        uri, local = name[1:].split("}", 1)
        return f"{URI_PREFIX.get(uri, uri)}:{local}"
    return name


def raw_attrs(node: ET.Element | None) -> dict:
    return {} if node is None else {short_qname(k): v for k, v in node.attrib.items()}


def wattr(node: ET.Element | None, name: str) -> str | None:
    return None if node is None else node.get(q("w", name))


def oattr(node: ET.Element | None, prefix: str, name: str) -> str | None:
    return None if node is None else node.get(q(prefix, name))


def rounded(value: float | None) -> float | None:
    return None if value is None else round(value, 4)


_LENGTH_RE = re.compile(r"^\s*([+-]?(?:\d+(?:\.\d*)?|\.\d+))\s*(in|cm|mm|pt)\s*$", re.I)


def length_to_pt(raw: str | None) -> float | None:
    """Convert an explicit ODF length; percentages and unitless values stay unresolved."""
    if raw is None:
        return None
    match = _LENGTH_RE.match(raw)
    if not match:
        return None
    value, unit = float(match.group(1)), match.group(2).lower()
    factor = {"pt": 1.0, "in": 72.0, "cm": 72.0 / 2.54, "mm": 72.0 / 25.4}[unit]
    return rounded(value * factor)


def twips_to_pt(raw: str | None) -> float | None:
    try:
        return rounded(float(raw) / 20.0) if raw is not None else None
    except ValueError:
        return None


def eighths_to_pt(raw: str | None) -> float | None:
    try:
        return rounded(float(raw) / 8.0) if raw is not None else None
    except ValueError:
        return None


def explicit_length(raw: str | None, source: str | None = None) -> dict:
    out = {"raw": raw, "pt": length_to_pt(raw)}
    if source is not None:
        out["source"] = source
    if raw is not None and out["pt"] is None:
        out["unresolved"] = "not an explicit in/cm/mm/pt length"
    return out


def twip_measure(raw: str | None, source: str | None = None) -> dict:
    out = {"raw": raw, "pt": twips_to_pt(raw)}
    if source is not None:
        out["source"] = source
    if raw is not None and out["pt"] is None:
        out["unresolved"] = "not a numeric twip value"
    return out


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", value or ""))


def text_label(value: str) -> str:
    compact = re.sub(r"\s+", " ", unicodedata.normalize("NFKC", value or "")).strip()
    return compact[:40]


def element_text(node: ET.Element) -> str:
    return "".join(node.itertext())


def top_level_descendants(node: ET.Element, target_tag: str) -> list[ET.Element]:
    """Find target descendants in order without descending into a found target."""
    result = []
    def visit(parent):
        for child in list(parent):
            if child.tag == target_tag:
                result.append(child)
            else:
                visit(child)
    visit(node)
    return result


def bool_word(node: ET.Element | None) -> bool | None:
    if node is None:
        return None
    value = wattr(node, "val")
    return True if value is None else value.lower() not in {"0", "false", "off", "no"}


def word_bool_property(node: ET.Element | None, source: str | None = None) -> dict:
    return {"raw": wattr(node, "val"), "value": bool_word(node),
            "present": node is not None, "source": source if node is not None else None}


def clean_internal(value):
    if isinstance(value, dict):
        return {k: clean_internal(v) for k, v in value.items() if not k.startswith("_")}
    if isinstance(value, list):
        return [clean_internal(v) for v in value]
    return value


def xml_from_zip(package: ZipFile, name: str) -> ET.Element | None:
    try:
        return ET.fromstring(package.read(name))
    except KeyError:
        return None


class DocxStyles:
    def __init__(self, root: ET.Element | None):
        self.root = root
        self.styles: dict[tuple[str, str], ET.Element] = {}
        self.doc_ppr = None
        self.doc_rpr = None
        if root is None:
            return
        defaults = root.find("w:docDefaults", NS)
        if defaults is not None:
            self.doc_ppr = defaults.find("w:pPrDefault/w:pPr", NS)
            self.doc_rpr = defaults.find("w:rPrDefault/w:rPr", NS)
        for style in root.findall("w:style", NS):
            kind, ident = wattr(style, "type"), wattr(style, "styleId")
            if kind and ident:
                self.styles[(kind, ident)] = style

    def style_chain(self, kind: str, ident: str | None):
        seen = set()
        while ident and ident not in seen:
            seen.add(ident)
            node = self.styles.get((kind, ident))
            if node is None:
                return
            yield ident, node
            based = node.find("w:basedOn", NS)
            ident = wattr(based, "val")

    def paragraph_levels(self, ppr: ET.Element | None, table_style: str | None):
        if ppr is not None:
            yield "direct", ppr
        style_id = wattr(ppr.find("w:pStyle", NS) if ppr is not None else None, "val")
        for ident, style in self.style_chain("paragraph", style_id):
            level = style.find("w:pPr", NS)
            if level is not None:
                yield f"paragraph style {ident}", level
        yield from self.table_levels(table_style, "pPr")
        if self.doc_ppr is not None:
            yield "docDefaults", self.doc_ppr

    def run_levels(self, paragraph: ET.Element, ppr: ET.Element | None, table_style: str | None):
        for run in paragraph.findall(".//w:r", NS):
            rpr = run.find("w:rPr", NS)
            if rpr is not None:
                yield "direct run", rpr
                break
        if ppr is not None:
            p_rpr = ppr.find("w:rPr", NS)
            if p_rpr is not None:
                yield "direct paragraph run properties", p_rpr
        style_id = wattr(ppr.find("w:pStyle", NS) if ppr is not None else None, "val")
        for ident, style in self.style_chain("paragraph", style_id):
            level = style.find("w:rPr", NS)
            if level is not None:
                yield f"paragraph style {ident}", level
        yield from self.table_levels(table_style, "rPr")
        if self.doc_rpr is not None:
            yield "docDefaults", self.doc_rpr

    def table_levels(self, ident: str | None, property_name: str):
        for style_id, style in self.style_chain("table", ident):
            whole = [x for x in style.findall("w:tblStylePr", NS) if wattr(x, "type") == "wholeTable"]
            for cond in whole:
                level = cond.find(f"w:{property_name}", NS)
                if level is not None:
                    yield f"table style {style_id} wholeTable", level
            level = style.find(f"w:{property_name}", NS)
            if level is not None:
                yield f"table style {style_id}", level

    def table_conditional_status(self, ident: str | None) -> dict:
        unsupported = []
        applied = []
        for style_id, style in self.style_chain("table", ident):
            for cond in style.findall("w:tblStylePr", NS):
                kind = wattr(cond, "type")
                if kind == "wholeTable":
                    applied.append({"style": style_id, "type": kind})
                else:
                    unsupported.append({"style": style_id, "type": kind})
        return {
            "appliedSimple": applied,
            "unresolved": unsupported,
            "reason": ("conditional region/edge precedence was not resolved" if unsupported else None),
        }

    def defaults_dump(self) -> dict:
        return {"pPr": element_tree_dump(self.doc_ppr), "rPr": element_tree_dump(self.doc_rpr)}


def element_tree_dump(node: ET.Element | None):
    if node is None:
        return None
    return {
        "tag": short_qname(node.tag),
        "attributes": raw_attrs(node),
        "children": [element_tree_dump(child) for child in node],
    }


def first_property(levels, property_name: str, attr_name: str | None = None):
    for source, level in levels:
        node = level.find(f"w:{property_name}", NS)
        if node is None:
            continue
        if attr_name is None:
            return node, source
        value = wattr(node, attr_name)
        if value is not None:
            return value, source
    return (None, None)


def resolved_word_scalar(levels, prop: str, attr_name: str = "val") -> dict:
    value, source = first_property(list(levels), prop, attr_name)
    return {"raw": value, "value": value, "source": source}


def resolved_word_bool(levels, prop: str) -> dict:
    for source, level in levels:
        node = level.find(f"w:{prop}", NS)
        if node is not None:
            return {"raw": wattr(node, "val"), "value": bool_word(node), "source": source}
    return {"raw": None, "value": None, "source": None}


def docx_theme(root: ET.Element | None) -> dict:
    result = {"themeName": None, "fontSchemeName": None, "major": {}, "minor": {}, "supplemental": {}}
    if root is None:
        return result
    result["themeName"] = root.get("name")
    scheme = root.find(".//a:fontScheme", NS)
    if scheme is None:
        return result
    result["fontSchemeName"] = scheme.get("name")
    for key, tag in (("major", "majorFont"), ("minor", "minorFont")):
        node = scheme.find(f"a:{tag}", NS)
        if node is None:
            continue
        for slot in ("latin", "ea", "cs"):
            child = node.find(f"a:{slot}", NS)
            result[key][slot] = None if child is None else child.get("typeface") or None
        result["supplemental"][key] = [
            {"script": x.get("script"), "typeface": x.get("typeface")}
            for x in node.findall("a:font", NS)
        ]
    return result


def theme_font(theme: dict, theme_ref: str | None) -> dict | None:
    if not theme_ref:
        return None
    match = re.match(r"^(major|minor)(HAnsi|EastAsia|Bidi)$", theme_ref)
    if not match:
        return {"raw": theme_ref, "themeName": theme.get("themeName"), "mappedFace": None,
                "unresolved": "unrecognized OOXML theme font reference"}
    family, suffix = match.groups()
    slot = {"HAnsi": "latin", "EastAsia": "ea", "Bidi": "cs"}[suffix]
    face = theme.get(family, {}).get(slot)
    out = {"raw": theme_ref, "themeName": theme.get("themeName"),
           "fontSchemeName": theme.get("fontSchemeName"), "mappedFace": face}
    if not face:
        out["unresolved"] = f"theme {family}/{slot} typeface is empty; script-specific fallback was not guessed"
    return out


def word_font_from_levels(levels, theme: dict) -> dict:
    for source, level in levels:
        fonts = level.find("w:rFonts", NS)
        if fonts is None:
            continue
        explicit = wattr(fonts, "eastAsia")
        ref = wattr(fonts, "eastAsiaTheme")
        if explicit is not None:
            return {"raw": explicit, "value": explicit, "source": source, "theme": None}
        if ref is not None:
            mapping = theme_font(theme, ref)
            return {"raw": ref, "value": mapping.get("mappedFace"), "source": source, "theme": mapping}
    return {"raw": None, "value": None, "source": None, "theme": None}


def word_font_size(levels) -> dict:
    levels = list(levels)
    for prop in ("sz", "szCs"):
        value, source = first_property(levels, prop, "val")
        if value is not None:
            try:
                pt = rounded(float(value) / 2.0)
            except ValueError:
                pt = None
            return {"raw": value, "pt": pt, "property": prop, "source": source}
    return {"raw": None, "pt": None, "property": None, "source": None}


def word_spacing(levels) -> dict:
    levels = list(levels)
    output = {}
    for name in ("before", "after"):
        raw, source = first_property(levels, "spacing", name)
        output[name] = twip_measure(raw, source)
    raw, source = first_property(levels, "spacing", "line")
    rule, rule_source = first_property(levels, "spacing", "lineRule")
    resolved = None
    if raw is not None:
        if rule in {"exact", "atLeast"}:
            resolved = {"rule": "exact" if rule == "exact" else "min", "value": twips_to_pt(raw), "unit": "pt"}
        elif rule == "auto":
            try:
                resolved = {"rule": "proportional", "value": rounded(float(raw) / 240.0), "unit": "line"}
            except ValueError:
                resolved = None
    output["line"] = {"raw": raw, "lineRuleRaw": rule, "source": source,
                      "ruleSource": rule_source, "resolved": resolved}
    if raw is not None and resolved is None:
        output["line"]["unresolved"] = ("lineRule is absent; its default was not guessed" if rule is None
                                            else "unsupported Word line rule/value")
    return output


def docx_paragraph(paragraph: ET.Element, styles: DocxStyles, table_style: str | None, theme: dict) -> dict:
    ppr = paragraph.find("w:pPr", NS)
    levels = list(styles.paragraph_levels(ppr, table_style))
    run_levels = list(styles.run_levels(paragraph, ppr, table_style))
    ind = {}
    for name in ("left", "right", "firstLine", "hanging", "start", "end"):
        raw, source = first_property(levels, "ind", name)
        ind[name] = twip_measure(raw, source)
    return {
        "style": wattr(ppr.find("w:pStyle", NS) if ppr is not None else None, "val"),
        "spacing": word_spacing(levels),
        "snapToGrid": resolved_word_bool(iter(levels), "snapToGrid"),
        "contextualSpacing": resolved_word_bool(iter(levels), "contextualSpacing"),
        "jc": resolved_word_scalar(iter(levels), "jc"),
        "ind": ind,
        "fontSize": word_font_size(run_levels),
        "eastAsiaFont": word_font_from_levels(run_levels, theme),
    }


def word_table_measure(node: ET.Element | None, source: str) -> dict:
    raw, kind = wattr(node, "w"), wattr(node, "type")
    out = {"raw": raw, "typeRaw": kind, "pt": None, "percent": None, "source": source if node is not None else None}
    if kind == "dxa":
        out["pt"] = twips_to_pt(raw)
    elif kind == "pct" and raw is not None:
        try:
            out["percent"] = rounded(float(raw) / 50.0)
        except ValueError:
            pass
    elif raw is not None and kind not in {"auto", "nil"}:
        out["unresolved"] = ("width type is absent; its default was not guessed" if kind is None
                             else "unsupported OOXML table width type")
    return out


def word_border_node(node: ET.Element | None, source: str | None, edge_used: str | None = None) -> dict:
    raw_val = wattr(node, "val")
    present = None if node is None else raw_val not in {"nil", "none"}
    result = {
        "raw": raw_attrs(node), "present": present, "style": raw_val,
        "widthPt": eighths_to_pt(wattr(node, "sz")), "spacePt": None,
        "color": wattr(node, "color"), "source": source, "edgeUsed": edge_used,
    }
    space = wattr(node, "space")
    if space is not None:
        try:
            result["spacePt"] = rounded(float(space))
        except ValueError:
            result["spaceUnresolved"] = "non-numeric OOXML border space"
    return result


def word_side_node(container: ET.Element | None, side: str):
    if container is None:
        return None, None
    names = {"left": ("start", "left"), "right": ("end", "right")}.get(side, (side,))
    for name in names:
        child = container.find(f"w:{name}", NS)
        if child is not None:
            return child, name
    return None, None


def resolve_word_margin(side: str, tcpr: ET.Element | None, trprex: ET.Element | None,
                        tblpr: ET.Element | None, styles: DocxStyles, table_style: str | None) -> dict:
    levels = []
    if tcpr is not None:
        levels.append(("direct cell", tcpr.find("w:tcMar", NS)))
    if trprex is not None:
        levels.append(("row tblPrEx", trprex.find("w:tblCellMar", NS)))
    if tblpr is not None:
        levels.append(("direct table", tblpr.find("w:tblCellMar", NS)))
    for source, style_tblpr in styles.table_levels(table_style, "tblPr"):
        levels.append((source, style_tblpr.find("w:tblCellMar", NS)))
    for source, mar in levels:
        node, used = word_side_node(mar, side)
        if node is not None:
            kind = wattr(node, "type")
            out = {"raw": wattr(node, "w"), "pt": twips_to_pt(wattr(node, "w")) if kind == "dxa" else None,
                   "source": source, "typeRaw": kind, "sideUsed": used}
            if kind is None:
                out["unresolved"] = "margin width type is absent; its default was not guessed"
            elif kind != "dxa":
                out["unresolved"] = "non-dxa cell margin was not converted to points"
            return out
    return {"raw": None, "pt": None, "source": None, "unresolved": "no stored margin; format default not guessed"}


def table_edge_name(side: str, row_index: int, row_count: int, start: int, end: int, grid_count: int) -> str:
    if side == "top":
        return "top" if row_index == 0 else "insideH"
    if side == "bottom":
        return "bottom" if row_index == row_count - 1 else "insideH"
    if side == "left":
        return "left" if start == 0 else "insideV"
    return "right" if end >= grid_count else "insideV"


def resolve_word_border(side: str, tcpr: ET.Element | None, trprex: ET.Element | None,
                        tblpr: ET.Element | None, styles: DocxStyles, table_style: str | None,
                        row_index: int, row_count: int, start: int, end: int, grid_count: int) -> dict:
    tc_borders = tcpr.find("w:tcBorders", NS) if tcpr is not None else None
    node, used = word_side_node(tc_borders, side)
    if node is not None:
        return word_border_node(node, "direct cell", used)
    edge = table_edge_name(side, row_index, row_count, start, end, grid_count)
    levels = []
    if trprex is not None:
        levels.append(("row tblPrEx", trprex.find("w:tblBorders", NS)))
    if tblpr is not None:
        levels.append(("direct table", tblpr.find("w:tblBorders", NS)))
    for source, style_tblpr in styles.table_levels(table_style, "tblPr"):
        levels.append((source, style_tblpr.find("w:tblBorders", NS)))
    for source, borders in levels:
        lookup = edge
        if edge == "left":
            node, used = word_side_node(borders, "left")
        elif edge == "right":
            node, used = word_side_node(borders, "right")
        else:
            node = borders.find(f"w:{edge}", NS) if borders is not None else None
            used = edge if node is not None else None
        if node is not None:
            return word_border_node(node, source, used)
    result = word_border_node(None, None, edge)
    result["unresolved"] = "no stored border at any implemented precedence level"
    return result


def word_row_height(trpr: ET.Element | None) -> dict:
    node = trpr.find("w:trHeight", NS) if trpr is not None else None
    raw, rule = wattr(node, "val"), wattr(node, "hRule")
    if node is None:
        kind = "none"
    elif rule is None:
        kind = "atLeast|absent"
    else:
        kind = {"exact": "exact", "atLeast": "min", "auto": "auto"}.get(rule, rule)
    return {"raw": {"val": raw, "hRule": rule}, "pt": twips_to_pt(raw), "kind": kind,
            "source": "direct row" if node is not None else None}


def dump_word_tbl_props(tblpr: ET.Element | None, styles: DocxStyles, style_id: str | None) -> dict:
    def child(name):
        return tblpr.find(f"w:{name}", NS) if tblpr is not None else None
    return {
        "tblStyle": style_id,
        "tblW": word_table_measure(child("tblW"), "direct table"),
        "tblLayout": raw_attrs(child("tblLayout")),
        "tblInd": word_table_measure(child("tblInd"), "direct table"),
        "jc": raw_attrs(child("jc")),
        "tblCellSpacing": word_table_measure(child("tblCellSpacing"), "direct table"),
        "tblCellMar": element_tree_dump(child("tblCellMar")),
        "tblBorders": element_tree_dump(child("tblBorders")),
        "tblLook": raw_attrs(child("tblLook")),
        "conditionalFormatting": styles.table_conditional_status(style_id),
    }


def parse_docx_table(tbl: ET.Element, path: str, styles: DocxStyles, theme: dict, tables: list):
    tblpr = tbl.find("w:tblPr", NS)
    style_id = wattr(tblpr.find("w:tblStyle", NS) if tblpr is not None else None, "val")
    grid = [twip_measure(wattr(x, "w"), "table grid") for x in tbl.findall("w:tblGrid/w:gridCol", NS)]
    row_nodes = tbl.findall("w:tr", NS)
    table = {
        "path": path, "text": text_label(element_text(tbl)), "_textFull": element_text(tbl),
        "properties": dump_word_tbl_props(tblpr, styles, style_id), "gridColumns": grid, "rows": [],
    }
    tables.append(table)
    grid_count = len(grid)
    for ri, tr in enumerate(row_nodes):
        trpr = tr.find("w:trPr", NS)
        trprex = tr.find("w:tblPrEx", NS)
        before = wattr(trpr.find("w:gridBefore", NS) if trpr is not None else None, "val")
        try:
            column = int(before or 0)
        except ValueError:
            column = 0
        row = {
            "index": ri, "text": text_label(element_text(tr)), "_textFull": element_text(tr),
            "gridBeforeRaw": before, "height": word_row_height(trpr),
            "cantSplit": word_bool_property(trpr.find("w:cantSplit", NS) if trpr is not None else None, "direct row"),
            "tblHeader": word_bool_property(trpr.find("w:tblHeader", NS) if trpr is not None else None, "direct row"),
            "tblPrEx": element_tree_dump(trprex), "cells": [],
        }
        for ci, tc in enumerate(tr.findall("w:tc", NS)):
            tcpr = tc.find("w:tcPr", NS)
            span_raw = wattr(tcpr.find("w:gridSpan", NS) if tcpr is not None else None, "val")
            try:
                span = max(1, int(span_raw or 1))
            except ValueError:
                span = 1
            start, end = column, column + span
            vmerge_node = tcpr.find("w:vMerge", NS) if tcpr is not None else None
            vmerge = None
            if vmerge_node is not None:
                vmerge = wattr(vmerge_node, "val") or "continue"
            paragraphs = [docx_paragraph(p, styles, style_id, theme) for p in tc.findall("w:p", NS)]
            cell = {
                "index": ci, "gridColumnStart": start, "gridSpan": span, "gridSpanRaw": span_raw,
                "vMerge": vmerge, "vMergeRaw": wattr(vmerge_node, "val"),
                "tcW": word_table_measure(tcpr.find("w:tcW", NS) if tcpr is not None else None, "direct cell"),
                "margins": {}, "borders": {},
                "vAlign": raw_attrs(tcpr.find("w:vAlign", NS) if tcpr is not None else None),
                "noWrap": word_bool_property(tcpr.find("w:noWrap", NS) if tcpr is not None else None, "direct cell"),
                "textDirection": raw_attrs(tcpr.find("w:textDirection", NS) if tcpr is not None else None),
                "shading": raw_attrs(tcpr.find("w:shd", NS) if tcpr is not None else None),
                "text": text_label("".join(element_text(p) for p in tc.findall("w:p", NS))),
                "paragraphs": paragraphs,
            }
            for side in ("top", "bottom", "left", "right"):
                cell["margins"][side] = resolve_word_margin(side, tcpr, trprex, tblpr, styles, style_id)
                cell["borders"][side] = resolve_word_border(
                    side, tcpr, trprex, tblpr, styles, style_id, ri, len(row_nodes), start, end, grid_count)
            row["cells"].append(cell)
            for nested_index, child in enumerate(top_level_descendants(tc, q("w", "tbl"))):
                parse_docx_table(child, f"{path}/r{ri}c{start}/{nested_index}", styles, theme, tables)
            column = end
        table["rows"].append(row)


def docx_sections(document: ET.Element | None) -> list:
    result = []
    if document is None:
        return result
    for index, sect in enumerate(document.findall(".//w:sectPr", NS)):
        pgsz, pgmar, grid = sect.find("w:pgSz", NS), sect.find("w:pgMar", NS), sect.find("w:docGrid", NS)
        refs = []
        for name in ("headerReference", "footerReference"):
            for ref in sect.findall(f"w:{name}", NS):
                refs.append({"kind": name, "type": wattr(ref, "type"), "relationshipId": ref.get(q("r", "id"))})
        result.append({
            "index": index,
            "pageSize": {"raw": raw_attrs(pgsz), "width": twip_measure(wattr(pgsz, "w"), "sectPr"),
                         "height": twip_measure(wattr(pgsz, "h"), "sectPr"), "orientRaw": wattr(pgsz, "orient")},
            "pageMargins": {name: twip_measure(wattr(pgmar, name), "sectPr")
                            for name in ("top", "bottom", "left", "right", "header", "footer", "gutter")},
            "docGrid": {"raw": raw_attrs(grid), "typeRaw": wattr(grid, "type"),
                        "linePitch": twip_measure(wattr(grid, "linePitch"), "sectPr"),
                        "charSpaceRaw": wattr(grid, "charSpace")},
            "titlePg": word_bool_property(sect.find("w:titlePg", NS), "sectPr"), "headerFooterReferences": refs,
        })
    return result


def docx_compat(settings: ET.Element | None) -> dict:
    compat = settings.find("w:compat", NS) if settings is not None else None
    children = []
    if compat is not None:
        for child in compat:
            item = {"name": short_qname(child.tag), "attributes": raw_attrs(child),
                    "xml": element_tree_dump(child)}
            if child.tag == q("w", "compatSetting"):
                item["compatSetting"] = {"name": wattr(child, "name"), "uri": wattr(child, "uri"), "val": wattr(child, "val")}
            children.append(item)
    return {"children": children}


def font_usage(theme: dict, styles_root: ET.Element | None, document_root: ET.Element | None) -> dict:
    attributes = ("ascii", "hAnsi", "eastAsia", "cs", "asciiTheme", "hAnsiTheme", "eastAsiaTheme", "csTheme")
    def count_roots(*roots):
        counts = {name: Counter() for name in attributes}
        for root in roots:
            if root is None:
                continue
            for node in root.findall(".//w:rFonts", NS):
                for name in counts:
                    value = wattr(node, name)
                    if value is not None:
                        counts[name][value] += 1
        return counts
    by_location = {"styles": count_roots(styles_root), "documentRuns": count_roots(document_root)}
    counts = {name: Counter() for name in attributes}
    for location in by_location.values():
        for name, values in location.items():
            counts[name].update(values)
    raw = {name: dict(sorted(counter.items())) for name, counter in counts.items()}
    locations = {where: {name: dict(sorted(counter.items())) for name, counter in values.items()}
                 for where, values in by_location.items()}
    resolved = {}
    for attribute in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "csTheme"):
        resolved[attribute] = [
            {"reference": reference, "count": count, **(theme_font(theme, reference) or {})}
            for reference, count in sorted(counts[attribute].items())
        ]
    return {"counts": raw, "byLocation": locations, "themeResolution": resolved}


def dump_docx(path: str | Path, include_internal: bool = False) -> dict:
    with ZipFile(path) as package:
        document = xml_from_zip(package, "word/document.xml")
        styles_root = xml_from_zip(package, "word/styles.xml")
        settings = xml_from_zip(package, "word/settings.xml")
        theme_root = xml_from_zip(package, "word/theme/theme1.xml")
    styles = DocxStyles(styles_root)
    theme = docx_theme(theme_root)
    tables = []
    if document is not None:
        body = document.find("w:body", NS)
        top_index = 0
        if body is not None:
            for child in top_level_descendants(body, q("w", "tbl")):
                parse_docx_table(child, str(top_index), styles, theme, tables)
                top_index += 1
    result = {
        "format": "docx", "file": str(path),
        "document": {"sections": docx_sections(document), "compat": docx_compat(settings),
                     "docDefaults": styles.defaults_dump(), "fontUsage": font_usage(theme, styles_root, document), "themeFonts": theme},
        "tables": tables,
    }
    return result if include_internal else clean_internal(result)


class OdtStyles:
    def __init__(self, styles_root: ET.Element | None, content_root: ET.Element | None):
        self.automatic: dict[tuple[str, str], ET.Element] = {}
        self.named: dict[tuple[str, str], ET.Element] = {}
        self.defaults: dict[str, ET.Element] = {}
        for root in (styles_root, content_root):
            if root is None:
                continue
            for section in root.findall("office:automatic-styles", NS):
                for node in section.findall("style:style", NS):
                    self._add(self.automatic, node)
            for section in root.findall("office:styles", NS):
                for node in section.findall("style:style", NS):
                    self._add(self.named, node)
                for node in section.findall("style:default-style", NS):
                    family = oattr(node, "style", "family")
                    if family:
                        self.defaults[family] = node

    @staticmethod
    def _add(target, node):
        name, family = oattr(node, "style", "name"), oattr(node, "style", "family")
        if name and family:
            target[(family, name)] = node

    def node(self, family: str, name: str | None):
        if name is None:
            return None, None
        if (family, name) in self.automatic:
            return self.automatic[(family, name)], "automatic"
        if (family, name) in self.named:
            return self.named[(family, name)], "named"
        return None, None

    def levels(self, family: str, name: str | None):
        node, kind = self.node(family, name)
        seen = set()
        first = True
        while node is not None:
            ident = oattr(node, "style", "name")
            key = (family, ident)
            if key in seen:
                break
            seen.add(key)
            if first and kind == "automatic":
                source = f"automatic style {ident}"
            else:
                source = f"parent style {ident}" if not first else f"named style {ident}"
            yield source, node
            parent = oattr(node, "style", "parent-style-name")
            if not parent:
                break
            if kind == "automatic" and (family, parent) in self.automatic:
                return
            node = self.named.get((family, parent))
            kind = "named"
            first = False
        default = self.defaults.get(family)
        if default is not None:
            yield f"default-style {family}", default

    def parent_status(self, family: str, name: str | None) -> dict:
        node, kind = self.node(family, name)
        parent = oattr(node, "style", "parent-style-name")
        return {"parentStyleName": parent,
                "parentIsAutomatic": bool(kind == "automatic" and parent and (family, parent) in self.automatic)}

    def attr(self, family: str, name: str | None, prop_tag: str, attr_prefix: str, attr_name: str):
        for source, style in self.levels(family, name):
            props = style.find(f"style:{prop_tag}", NS)
            value = oattr(props, attr_prefix, attr_name)
            if value is not None:
                return value, source
        return None, None

    def all_attrs(self, family: str, name: str | None, prop_tag: str) -> dict:
        merged = {}
        sources = {}
        levels = list(self.levels(family, name))
        for source, style in reversed(levels):
            props = style.find(f"style:{prop_tag}", NS)
            if props is not None:
                for key, value in raw_attrs(props).items():
                    merged[key] = value
                    sources[key] = source
        return {"raw": merged, "sources": sources, **self.parent_status(family, name)}

    def defaults_dump(self) -> list:
        return [{"family": family, "properties": element_tree_dump(node)} for family, node in sorted(self.defaults.items())]


def odt_length_attr(styles: OdtStyles, family: str, name: str | None, prop_tag: str,
                    attr_prefix: str, attr_name: str) -> dict:
    raw, source = styles.attr(family, name, prop_tag, attr_prefix, attr_name)
    return explicit_length(raw, source)


def resolve_odt_box_side(styles: OdtStyles, family: str, name: str | None, prop_tag: str,
                         side: str, base_name: str) -> tuple[str | None, str | None, str | None]:
    for source, style in styles.levels(family, name):
        props = style.find(f"style:{prop_tag}", NS)
        if props is None:
            continue
        side_raw = oattr(props, "fo", f"{base_name}-{side}")
        if side_raw is not None:
            return side_raw, source, f"fo:{base_name}-{side}"
        raw = oattr(props, "fo", base_name)
        if raw is not None:
            return raw, source, f"fo:{base_name}"
    return None, None, None


def odt_border(raw: str | None, source: str | None, attribute_used: str | None) -> dict:
    result = {"raw": raw, "present": None, "style": None, "widthPt": None,
              "color": None, "source": source, "attributeUsed": attribute_used}
    if raw is None:
        result["unresolved"] = "no stored border at any implemented style level"
        return result
    if raw.strip().lower() == "none":
        result.update({"present": False, "style": "none"})
        return result
    parts = raw.split()
    result["present"] = True
    for part in parts:
        if length_to_pt(part) is not None:
            result["widthPt"] = length_to_pt(part)
        elif part.lower() in {"solid", "double", "dotted", "dashed", "groove", "ridge", "inset", "outset"}:
            result["style"] = part.lower()
        elif part.startswith("#") or part.lower() in {"black", "white", "red", "blue", "green", "transparent"}:
            result["color"] = part
    if result["style"] is None or result["widthPt"] is None:
        result["parseNote"] = "border shorthand was only partially parsed"
    return result


def odt_row_height(styles: OdtStyles, style_name: str | None) -> dict:
    exact, esource = styles.attr("table-row", style_name, "table-row-properties", "style", "row-height")
    minimum, msource = styles.attr("table-row", style_name, "table-row-properties", "style", "min-row-height")
    optimal, osource = styles.attr("table-row", style_name, "table-row-properties", "style", "use-optimal-row-height")
    if exact is not None:
        kind, raw, source = "exact", exact, esource
    elif minimum is not None:
        kind, raw, source = "min", minimum, msource
    elif optimal in {"true", "1"}:
        kind, raw, source = "optimal", optimal, osource
    else:
        kind, raw, source = "none", None, None
    return {"kind": kind, "raw": raw, "pt": length_to_pt(raw), "source": source,
            "rowHeightRaw": exact, "minRowHeightRaw": minimum, "useOptimalRaw": optimal}


def odt_line_spacing(styles: OdtStyles, style_name: str | None) -> dict:
    line, source = styles.attr("paragraph", style_name, "paragraph-properties", "fo", "line-height")
    at_least, at_source = styles.attr("paragraph", style_name, "paragraph-properties", "style", "line-height-at-least")
    spacing, sp_source = styles.attr("paragraph", style_name, "paragraph-properties", "style", "line-spacing")
    resolved = None
    raw, primary_source = line, source
    if line is not None:
        if line.endswith("%"):
            try:
                resolved = {"rule": "proportional", "value": rounded(float(line[:-1]) / 100.0), "unit": "line"}
            except ValueError:
                pass
        elif length_to_pt(line) is not None:
            resolved = {"rule": "exact", "value": length_to_pt(line), "unit": "pt"}
        elif line == "normal":
            resolved = {"rule": "normal", "value": None, "unit": None}
    elif at_least is not None:
        raw, primary_source = at_least, at_source
        if length_to_pt(at_least) is not None:
            resolved = {"rule": "min", "value": length_to_pt(at_least), "unit": "pt"}
    elif spacing is not None:
        raw, primary_source = spacing, sp_source
        if length_to_pt(spacing) is not None:
            resolved = {"rule": "leading", "value": length_to_pt(spacing), "unit": "pt"}
    result = {"raw": raw, "source": primary_source, "lineHeightRaw": line,
              "lineHeightAtLeastRaw": at_least, "lineSpacingRaw": spacing, "resolved": resolved}
    if raw is not None and resolved is None:
        result["unresolved"] = "unsupported ODF line-height form"
    return result


def odt_first_text_style(paragraph: ET.Element) -> str | None:
    for span in paragraph.findall(".//text:span", NS):
        if normalize_text(element_text(span)):
            return oattr(span, "text", "style-name")
    return None


def odt_text_attr(paragraph: ET.Element, styles: OdtStyles, paragraph_style: str | None,
                  prefix: str, name: str):
    text_style = odt_first_text_style(paragraph)
    if text_style:
        value, source = styles.attr("text", text_style, "text-properties", prefix, name)
        if value is not None:
            return value, source, text_style
    value, source = styles.attr("paragraph", paragraph_style, "text-properties", prefix, name)
    return value, source, text_style


def odt_paragraph(paragraph: ET.Element, styles: OdtStyles) -> dict:
    style_name = oattr(paragraph, "text", "style-name")
    ind = {}
    for prefix, name in (("fo", "text-indent"), ("fo", "margin-left"), ("fo", "margin-right"),
                         ("loext", "margin-left"), ("loext", "margin-right")):
        raw, source = styles.attr("paragraph", style_name, "paragraph-properties", prefix, name)
        ind[f"{prefix}:{name}"] = explicit_length(raw, source)
    font_size, fs_source, text_style = odt_text_attr(paragraph, styles, style_name, "fo", "font-size")
    asian_size, as_source, _ = odt_text_attr(paragraph, styles, style_name, "style", "font-size-asian")
    font_name, fn_source, _ = odt_text_attr(paragraph, styles, style_name, "style", "font-name")
    asian_name, an_source, _ = odt_text_attr(paragraph, styles, style_name, "style", "font-name-asian")
    before, bsource = styles.attr("paragraph", style_name, "paragraph-properties", "fo", "margin-top")
    after, asource = styles.attr("paragraph", style_name, "paragraph-properties", "fo", "margin-bottom")
    snap, snap_source = styles.attr("paragraph", style_name, "paragraph-properties", "style", "snap-to-layout-grid")
    contextual, csource = styles.attr("paragraph", style_name, "paragraph-properties", "style", "contextual-spacing")
    align, align_source = styles.attr("paragraph", style_name, "paragraph-properties", "fo", "text-align")
    return {
        "style": style_name, "firstTextStyle": text_style, **styles.parent_status("paragraph", style_name),
        "spacing": {"before": explicit_length(before, bsource), "after": explicit_length(after, asource),
                    "line": odt_line_spacing(styles, style_name)},
        "snapToGrid": {"raw": snap, "value": snap, "source": snap_source},
        "contextualSpacing": {"raw": contextual, "value": contextual, "source": csource},
        "textAlign": {"raw": align, "value": align, "source": align_source}, "ind": ind,
        "fontSize": {"raw": font_size, "pt": length_to_pt(font_size), "source": fs_source},
        "fontSizeAsian": {"raw": asian_size, "pt": length_to_pt(asian_size), "source": as_source},
        "fontName": {"raw": font_name, "value": font_name, "source": fn_source},
        "fontNameAsian": {"raw": asian_name, "value": asian_name, "source": an_source},
    }


def nested_paragraphs(cell: ET.Element):
    result = []
    def visit(node):
        for child in list(node):
            if child.tag == q("table", "table"):
                continue
            if child.tag in {q("text", "p"), q("text", "h")}:
                result.append(child)
            else:
                visit(child)
    visit(cell)
    return result


def odt_table_children(table: ET.Element, target: str) -> list:
    result = []
    grouping = {q("table", "table-header-rows"), q("table", "table-rows"), q("table", "table-row-group"),
                q("table", "table-header-columns"), q("table", "table-columns"), q("table", "table-column-group")}
    def visit(node):
        for child in list(node):
            if child.tag == target:
                result.append(child)
            elif child.tag in grouping:
                visit(child)
    visit(table)
    return result


def odt_table_props(styles: OdtStyles, name: str | None) -> dict:
    props = styles.all_attrs("table", name, "table-properties")
    def length(prefix, attr):
        raw, source = styles.attr("table", name, "table-properties", prefix, attr)
        return explicit_length(raw, source)
    rel, rel_source = styles.attr("table", name, "table-properties", "style", "rel-width")
    align, align_source = styles.attr("table", name, "table-properties", "table", "align")
    model, model_source = styles.attr("table", name, "table-properties", "table", "border-model")
    margins = {}
    for side in ("top", "bottom", "left", "right"):
        raw, source, used = resolve_odt_box_side(styles, "table", name, "table-properties", side, "margin")
        margins[side] = explicit_length(raw, source)
        margins[side]["attributeUsed"] = used
    return {"style": name, **styles.parent_status("table", name), "all": props,
            "width": length("style", "width"), "relWidth": {"raw": rel, "value": rel, "source": rel_source},
            "align": {"raw": align, "value": align, "source": align_source},
            "margins": margins,
            "borderModel": {"raw": model, "value": model, "source": model_source}}


def parse_odt_table(table_node: ET.Element, path: str, styles: OdtStyles, tables: list):
    style_name = oattr(table_node, "table", "style-name")
    cols = []
    for col_node in odt_table_children(table_node, q("table", "table-column")):
        col_style = oattr(col_node, "table", "style-name")
        repeat_raw = oattr(col_node, "table", "number-columns-repeated")
        try:
            repeat = max(1, int(repeat_raw or 1))
        except ValueError:
            repeat = 1
        width = odt_length_attr(styles, "table-column", col_style, "table-column-properties", "style", "column-width")
        rel, rel_source = styles.attr("table-column", col_style, "table-column-properties", "style", "rel-column-width")
        for repeat_index in range(repeat):
            cols.append({"style": col_style, "width": width, "relWidth": {"raw": rel, "value": rel, "source": rel_source},
                         "numberColumnsRepeatedRaw": repeat_raw, "repeatIndex": repeat_index})
    row_nodes = odt_table_children(table_node, q("table", "table-row"))
    item = {"path": path, "name": oattr(table_node, "table", "name"), "text": text_label(element_text(table_node)),
            "_textFull": element_text(table_node), "properties": odt_table_props(styles, style_name),
            "gridColumns": cols, "rows": []}
    tables.append(item)
    logical_row = 0
    for xml_row_index, row_node in enumerate(row_nodes):
        row_style = oattr(row_node, "table", "style-name")
        repeat_raw = oattr(row_node, "table", "number-rows-repeated")
        try:
            repeat = max(1, int(repeat_raw or 1))
        except ValueError:
            repeat = 1
        # Very large repeated blank spreadsheet ranges are represented, not exploded.
        expand = repeat if repeat <= 4096 else 1
        for repeat_index in range(expand):
            keep, keep_source = styles.attr("table-row", row_style, "table-row-properties", "fo", "keep-together")
            row = {"index": logical_row, "xmlRowIndex": xml_row_index, "repeatIndex": repeat_index,
                   "numberRowsRepeatedRaw": repeat_raw, "repeatExpansionTruncated": repeat > 4096,
                   "text": text_label(element_text(row_node)), "_textFull": element_text(row_node),
                   "height": odt_row_height(styles, row_style),
                   "keepTogether": {"raw": keep, "value": keep, "source": keep_source}, "cells": []}
            column = 0
            for xml_cell_index, cell_node in enumerate(list(row_node)):
                if cell_node.tag not in {q("table", "table-cell"), q("table", "covered-table-cell")}:
                    continue
                covered = cell_node.tag == q("table", "covered-table-cell")
                repeat_cell_raw = oattr(cell_node, "table", "number-columns-repeated")
                try:
                    repeat_cell = max(1, int(repeat_cell_raw or 1))
                except ValueError:
                    repeat_cell = 1
                span_raw = oattr(cell_node, "table", "number-columns-spanned")
                row_span_raw = oattr(cell_node, "table", "number-rows-spanned")
                try:
                    span = max(1, int(span_raw or 1))
                except ValueError:
                    span = 1
                cell_style = oattr(cell_node, "table", "style-name")
                paragraphs = [odt_paragraph(p, styles) for p in nested_paragraphs(cell_node)]
                for cell_repeat_index in range(repeat_cell if repeat_cell <= 4096 else 1):
                    cell = {"index": len(row["cells"]), "xmlCellIndex": xml_cell_index,
                            "gridColumnStart": column, "numberColumnsSpanned": span,
                            "numberColumnsSpannedRaw": span_raw, "numberRowsSpannedRaw": row_span_raw,
                            "numberColumnsRepeatedRaw": repeat_cell_raw, "repeatIndex": cell_repeat_index,
                            "covered": covered, "style": cell_style, **styles.parent_status("table-cell", cell_style),
                            "margins": {}, "borders": {},
                            "vAlign": {}, "writingMode": {}, "backgroundColor": {},
                            "text": text_label(element_text(cell_node)), "paragraphs": paragraphs}
                    for side in ("top", "bottom", "left", "right"):
                        raw, source, used = resolve_odt_box_side(styles, "table-cell", cell_style,
                                                                 "table-cell-properties", side, "padding")
                        cell["margins"][side] = explicit_length(raw, source)
                        cell["margins"][side]["attributeUsed"] = used
                        raw, source, used = resolve_odt_box_side(styles, "table-cell", cell_style,
                                                                 "table-cell-properties", side, "border")
                        cell["borders"][side] = odt_border(raw, source, used)
                    for key, prefix, attr_name in (("vAlign", "style", "vertical-align"),
                                                   ("writingMode", "style", "writing-mode"),
                                                   ("backgroundColor", "fo", "background-color")):
                        raw, source = styles.attr("table-cell", cell_style, "table-cell-properties", prefix, attr_name)
                        cell[key] = {"raw": raw, "value": raw, "source": source}
                    row["cells"].append(cell)
                    column += 1
            item["rows"].append(row)
            logical_row += 1
        if repeat > 4096:
            logical_row += repeat - 1
        # Nested table paths are attached to the logical row's anchor cell.
        for cell_node in [x for x in list(row_node) if x.tag == q("table", "table-cell")]:
            start = 0
            prior = list(row_node)[:list(row_node).index(cell_node)]
            for p in prior:
                if p.tag in {q("table", "table-cell"), q("table", "covered-table-cell")}:
                    try:
                        start += max(1, int(oattr(p, "table", "number-columns-repeated") or 1))
                    except ValueError:
                        start += 1
            for nested_index, child in enumerate(top_level_descendants(cell_node, q("table", "table"))):
                parse_odt_table(child, f"{path}/r{logical_row - expand}c{start}/{nested_index}", styles, tables)


def odt_document(styles_root: ET.Element | None, content_root: ET.Element | None,
                 settings_root: ET.Element | None, styles: OdtStyles) -> dict:
    page_layouts = []
    if styles_root is not None:
        for layout in styles_root.findall(".//style:page-layout", NS):
            props = layout.find("style:page-layout-properties", NS)
            header = layout.find("style:header-style/style:header-footer-properties", NS)
            footer = layout.find("style:footer-style/style:header-footer-properties", NS)
            page_layouts.append({"name": oattr(layout, "style", "name"),
                                 "pageLayoutProperties": raw_attrs(props),
                                 "headerFooterProperties": {"header": raw_attrs(header), "footer": raw_attrs(footer)}})
    masters = []
    if styles_root is not None:
        for master in styles_root.findall(".//style:master-page", NS):
            masters.append({"name": oattr(master, "style", "name"),
                            "pageLayoutName": oattr(master, "style", "page-layout-name"),
                            "nextStyleName": oattr(master, "style", "next-style-name")})
    faces = []
    seen_faces = set()
    for root in (styles_root, content_root):
        if root is None:
            continue
        for face in root.findall(".//style:font-face", NS):
            record = {"name": oattr(face, "style", "name"), "fontFamily": oattr(face, "svg", "font-family"),
                      "fontPitch": oattr(face, "style", "font-pitch"),
                      "fontFamilyGeneric": oattr(face, "style", "font-family-generic")}
            key = tuple(record.values())
            if key not in seen_faces:
                seen_faces.add(key)
                faces.append(record)
    config = []
    if settings_root is not None:
        for item_set in settings_root.findall(".//config:config-item-set", NS):
            if oattr(item_set, "config", "name") != "ooo:configuration-settings":
                continue
            for node in item_set.findall(".//config:config-item", NS):
                config.append({"name": oattr(node, "config", "name"), "type": oattr(node, "config", "type"),
                               "value": node.text or ""})
    return {"pageLayouts": page_layouts, "masterPages": masters, "fontFaces": faces,
            "defaultStyles": styles.defaults_dump(), "configurationSettings": config}


def dump_odt(path: str | Path, include_internal: bool = False) -> dict:
    with ZipFile(path) as package:
        content = xml_from_zip(package, "content.xml")
        styles_root = xml_from_zip(package, "styles.xml")
        settings = xml_from_zip(package, "settings.xml")
    styles = OdtStyles(styles_root, content)
    tables = []
    if content is not None:
        body = content.find("office:body", NS)
        root_tables = []
        if body is not None:
            root_tables = top_level_descendants(body, q("table", "table"))
        for index, table in enumerate(root_tables):
            # visit() does not descend into a table, so these are top-level tables.
            parse_odt_table(table, str(index), styles, tables)
    result = {"format": "odt", "file": str(path),
              "document": odt_document(styles_root, content, settings, styles), "tables": tables}
    return result if include_internal else clean_internal(result)


CATEGORIES = [
    "rowHeightRule", "rowHeightValue", "cellPadding", "cellBorder", "vAlign",
    "columnWidth", "tableWidth", "paraLineSpacing", "paraSpacingBeforeAfter",
    "snapToGrid", "fontSize", "eastAsiaFont", "pageSize", "pageMargins",
    "headerFooterDistance", "grid", "unmatchedTable", "unmatchedRow", "unmatchedCell",
]


def similarity(left: str, right: str) -> float:
    left, right = normalize_text(left)[:300], normalize_text(right)[:300]
    if not left and not right:
        return 1.0
    return difflib.SequenceMatcher(None, left, right).ratio()


def align_items(left: list, right: list, text_key="_textFull"):
    if len(left) == len(right):
        return [(a, b) for a, b in zip(left, right)], [], []
    candidates = []
    for li, a in enumerate(left):
        for ri, b in enumerate(right):
            candidates.append((similarity(a.get(text_key, ""), b.get(text_key, "")), -abs(li - ri), li, ri))
    used_l, used_r, pairs = set(), set(), []
    for score, _, li, ri in sorted(candidates, reverse=True):
        if li not in used_l and ri not in used_r and score >= 0.12:
            used_l.add(li); used_r.add(ri); pairs.append((left[li], right[ri]))
    pairs.sort(key=lambda p: left.index(p[0]))
    return pairs, [x for i, x in enumerate(left) if i not in used_l], [x for i, x in enumerate(right) if i not in used_r]


def fmt_pt(value):
    return "null" if value is None else f"{value:g}pt"


def scalar_bool(value):
    raw = value.get("value") if isinstance(value, dict) else value
    if raw is None:
        return None
    if isinstance(raw, bool):
        return raw
    return str(raw).lower() not in {"false", "0", "off", "no"}


def measure_mismatch(a: dict, b: dict, threshold: float, missing=True) -> bool:
    av, bv = a.get("pt"), b.get("pt")
    if av is None or bv is None:
        return missing and (av is None) != (bv is None)
    return abs(av - bv) > threshold


def normalize_border_style(value: str | None) -> str | None:
    return {"solid": "single", "dash": "dashed", "dot": "dotted", "nil": "none"}.get(value, value)


def add_diff(ctx: dict, category: str, path: str, left, right, pattern: str | None = None, detail: str | None = None):
    record = {"category": category, "path": path, ctx["leftLabel"]: clean_internal(left), ctx["rightLabel"]: clean_internal(right)}
    if pattern:
        record["pattern"] = pattern
        ctx["patterns"][pattern] += 1
    if detail:
        record["detail"] = detail
    ctx["diffs"].append(record)
    ctx["counts"][category] += 1


def row_height_pattern(label: str, height: dict) -> str:
    kind = height.get("kind")
    return f"{label}:{kind} {fmt_pt(height.get('pt'))}"


def compare_paragraphs(ctx, path, left_paras, right_paras):
    max_count = max(len(left_paras), len(right_paras))
    mismatch = {k: None for k in ("paraLineSpacing", "paraSpacingBeforeAfter", "snapToGrid", "fontSize", "eastAsiaFont")}
    for index in range(max_count):
        if index >= len(left_paras) or index >= len(right_paras):
            continue
        a, b = left_paras[index], right_paras[index]
        aline, bline = a["spacing"]["line"], b["spacing"]["line"]
        ar, br = aline.get("resolved"), bline.get("resolved")
        if ar is not None and br is not None:
            if ar.get("rule") != br.get("rule") or ar.get("unit") != br.get("unit") or (
                    ar.get("value") is not None and br.get("value") is not None and abs(ar["value"] - br["value"]) > 0.01):
                mismatch["paraLineSpacing"] = (a["spacing"]["line"], b["spacing"]["line"],
                    f"line {ar.get('rule')} {ar.get('value')} {ar.get('unit')} -> {br.get('rule')} {br.get('value')} {br.get('unit')}")
        elif (ar is None) != (br is None):
            mismatch["paraLineSpacing"] = (aline, bline, "line resolved on one side only")
        for which in ("before", "after"):
            if measure_mismatch(a["spacing"][which], b["spacing"][which], 0.5, missing=False):
                mismatch["paraSpacingBeforeAfter"] = (a["spacing"], b["spacing"],
                    f"paragraph {which} {fmt_pt(a['spacing'][which].get('pt'))} -> {fmt_pt(b['spacing'][which].get('pt'))}")
        if scalar_bool(a["snapToGrid"]) != scalar_bool(b["snapToGrid"]):
            mismatch["snapToGrid"] = (a["snapToGrid"], b["snapToGrid"],
                                       f"snapToGrid {scalar_bool(a['snapToGrid'])} -> {scalar_bool(b['snapToGrid'])}")
        afs = a.get("fontSizeAsian") if a.get("fontSizeAsian", {}).get("pt") is not None else a.get("fontSize", {})
        bfs = b.get("fontSizeAsian") if b.get("fontSizeAsian", {}).get("pt") is not None else b.get("fontSize", {})
        if measure_mismatch(afs, bfs, 0.25, missing=False):
            mismatch["fontSize"] = (afs, bfs, f"font size {fmt_pt(afs.get('pt'))} -> {fmt_pt(bfs.get('pt'))}")
        afont = a.get("fontNameAsian", a.get("eastAsiaFont", {})).get("value")
        bfont = b.get("fontNameAsian", b.get("eastAsiaFont", {})).get("value")
        if afont != bfont and (afont is not None or bfont is not None):
            mismatch["eastAsiaFont"] = (a.get("fontNameAsian", a.get("eastAsiaFont")),
                                         b.get("fontNameAsian", b.get("eastAsiaFont")),
                                         f"East Asia font {afont or 'null'} -> {bfont or 'null'}")
    for category, value in mismatch.items():
        if value is not None:
            add_diff(ctx, category, path, value[0], value[1], value[2])


def compare_cells(ctx, table_path, row_left, row_right):
    left = {c["gridColumnStart"]: c for c in row_left["cells"] if not c.get("covered")}
    right = {c["gridColumnStart"]: c for c in row_right["cells"] if not c.get("covered")}
    left_covered = {c["gridColumnStart"]: c for c in row_left["cells"] if c.get("covered")}
    right_covered = {c["gridColumnStart"]: c for c in row_right["cells"] if c.get("covered")}
    columns = set(left) | set(right)
    # A Word vMerge continuation is a real tc while ODF represents the same grid
    # position as covered-table-cell.  Horizontal covered cells have no opposite
    # anchor and are consequently not treated as extra cells.
    for col in list(columns):
        if col not in left and col in left_covered:
            left[col] = left_covered[col]
        if col not in right and col in right_covered:
            right[col] = right_covered[col]
    for col in sorted(columns):
        path = f"{table_path}/r{row_left['index']}c{col}"
        if col not in left or col not in right:
            add_diff(ctx, "unmatchedCell", path, left.get(col), right.get(col),
                     f"cell at grid column {col} exists on one side only")
            continue
        a, b = left[col], right[col]
        for side in ("top", "bottom", "left", "right"):
            am, bm = a["margins"][side], b["margins"][side]
            if measure_mismatch(am, bm, 0.5):
                add_diff(ctx, "cellPadding", path + f"/{side}", am, bm,
                         f"padding {side} {fmt_pt(am.get('pt'))} -> {fmt_pt(bm.get('pt'))}")
            ab, bb = a["borders"][side], b["borders"][side]
            style_a, style_b = normalize_border_style(ab.get("style")), normalize_border_style(bb.get("style"))
            pa, pb = ab.get("present"), bb.get("present")
            presence = pa != pb and True in {pa, pb}
            width = False
            if ab.get("widthPt") is not None and bb.get("widthPt") is not None:
                width = abs(ab["widthPt"] - bb["widthPt"]) > 0.25
            elif (ab.get("widthPt") is None) != (bb.get("widthPt") is None) and ab.get("present") and bb.get("present"):
                width = True
            style = pa is True and pb is True and style_a != style_b
            if presence or width or style:
                add_diff(ctx, "cellBorder", path + f"/{side}", ab, bb,
                         f"border {side} {style_a or 'null'} {fmt_pt(ab.get('widthPt'))} -> {style_b or 'null'} {fmt_pt(bb.get('widthPt'))}")
        aval = a.get("vAlign", {}).get("w:val", a.get("vAlign", {}).get("value"))
        bval = b.get("vAlign", {}).get("w:val", b.get("vAlign", {}).get("value"))
        aval = {"middle": "center"}.get(aval, aval); bval = {"middle": "center"}.get(bval, bval)
        if aval != bval and (aval is not None or bval is not None):
            add_diff(ctx, "vAlign", path, a.get("vAlign"), b.get("vAlign"), f"vAlign {aval or 'null'} -> {bval or 'null'}")
        compare_paragraphs(ctx, path, a["paragraphs"], b["paragraphs"])


def compare_rows(ctx, table_path, left_rows, right_rows):
    pairs, unmatched_l, unmatched_r = align_items(left_rows, right_rows)
    for row in unmatched_l:
        add_diff(ctx, "unmatchedRow", f"{table_path}/r{row['index']}", row, None, "unmatched row on left")
    for row in unmatched_r:
        add_diff(ctx, "unmatchedRow", f"{table_path}/r{row['index']}", None, row, "unmatched row on right")
    for a, b in pairs:
        path = f"{table_path}/r{a['index']}"
        ah, bh = a["height"], b["height"]
        if ah.get("kind") != bh.get("kind"):
            pattern = f"{row_height_pattern(ctx['leftShort'], ah)} -> {row_height_pattern(ctx['rightShort'], bh)}"
            add_diff(ctx, "rowHeightRule", path, ah, bh, pattern)
        if ah.get("pt") is not None and bh.get("pt") is not None and abs(ah["pt"] - bh["pt"]) > 0.5:
            add_diff(ctx, "rowHeightValue", path, ah, bh,
                     f"row height {fmt_pt(ah['pt'])} -> {fmt_pt(bh['pt'])}")
        compare_cells(ctx, table_path, a, b)


def width_signature(width: dict):
    if width.get("pt") is not None:
        return "pt", width["pt"]
    if width.get("percent") is not None:
        return "percent", width["percent"]
    raw = width.get("raw")
    if isinstance(raw, str) and raw.endswith("%"):
        try:
            return "percent", float(raw[:-1])
        except ValueError:
            pass
    return None, None


def compare_tables(ctx, left_tables, right_tables):
    pairs, unmatched_l, unmatched_r = align_items(left_tables, right_tables)
    for table in unmatched_l:
        add_diff(ctx, "unmatchedTable", table["path"], table, None, "unmatched table on left")
    for table in unmatched_r:
        add_diff(ctx, "unmatchedTable", table["path"], None, table, "unmatched table on right")
    for a, b in pairs:
        path = a["path"]
        aw, bw = a["properties"].get("tblW", a["properties"].get("width", {})), b["properties"].get("tblW", b["properties"].get("width", {}))
        au, av = width_signature(aw); bu, bv = width_signature(bw)
        if au == bu == "pt" and abs(av - bv) > 1.0 or (au and bu and au != bu):
            add_diff(ctx, "tableWidth", path, aw, bw, f"table width {au}:{av} -> {bu}:{bv}")
        for index in range(max(len(a["gridColumns"]), len(b["gridColumns"]))):
            ac = a["gridColumns"][index] if index < len(a["gridColumns"]) else None
            bc = b["gridColumns"][index] if index < len(b["gridColumns"]) else None
            am = ac if ac is None or "pt" in ac else ac.get("width", {})
            bm = bc if bc is None or "pt" in bc else bc.get("width", {})
            if ac is None or bc is None or measure_mismatch(am, bm, 1.0, missing=False):
                add_diff(ctx, "columnWidth", f"{path}/col{index}", ac, bc,
                         f"column width {fmt_pt(None if am is None else am.get('pt'))} -> {fmt_pt(None if bm is None else bm.get('pt'))}")
        compare_rows(ctx, path, a["rows"], b["rows"])


def odt_layout_order(document: dict) -> list:
    by_name = {x["name"]: x for x in document.get("pageLayouts", [])}
    ordered = [by_name[x["pageLayoutName"]] for x in document.get("masterPages", []) if x.get("pageLayoutName") in by_name]
    return ordered or document.get("pageLayouts", [])


def doc_page_views(dump: dict):
    if dump["format"] == "docx":
        return dump["document"].get("sections", [])
    result = []
    for layout in odt_layout_order(dump["document"]):
        attrs = layout["pageLayoutProperties"]
        size = {"width": explicit_length(attrs.get("fo:page-width"), "page-layout"),
                "height": explicit_length(attrs.get("fo:page-height"), "page-layout"),
                "orientRaw": attrs.get("style:print-orientation"), "raw": attrs}
        margins = {side: explicit_length(attrs.get(f"fo:margin-{side}", attrs.get("fo:margin")), "page-layout")
                   for side in ("top", "bottom", "left", "right")}
        margins.update({"header": {"raw": None, "pt": None, "source": None,
                                   "unresolved": "ODF header style does not store a directly equivalent edge distance"},
                        "footer": {"raw": None, "pt": None, "source": None,
                                   "unresolved": "ODF footer style does not store a directly equivalent edge distance"},
                        "gutter": {"raw": None, "pt": None, "source": None}})
        grid_keys = ("style:layout-grid-mode", "style:layout-grid-base-height", "style:layout-grid-ruby-height",
                     "style:layout-grid-base-width", "style:layout-grid-lines")
        result.append({"pageSize": size, "pageMargins": margins,
                       "headerFooterProperties": layout["headerFooterProperties"],
                       "docGrid": {"raw": {key: attrs.get(key) for key in grid_keys}}})
    return result


def compare_document(ctx, left_dump, right_dump):
    left, right = doc_page_views(left_dump), doc_page_views(right_dump)
    for index in range(max(len(left), len(right))):
        path = f"document/page{index}"
        if index >= len(left) or index >= len(right):
            add_diff(ctx, "pageSize", path, left[index] if index < len(left) else None,
                     right[index] if index < len(right) else None, "page-layout count differs")
            continue
        a, b = left[index], right[index]
        for dimension in ("width", "height"):
            if measure_mismatch(a["pageSize"][dimension], b["pageSize"][dimension], 0.5, missing=False):
                add_diff(ctx, "pageSize", path + "/" + dimension, a["pageSize"][dimension], b["pageSize"][dimension],
                         f"page {dimension} {fmt_pt(a['pageSize'][dimension].get('pt'))} -> {fmt_pt(b['pageSize'][dimension].get('pt'))}")
        for side in ("top", "bottom", "left", "right"):
            if measure_mismatch(a["pageMargins"][side], b["pageMargins"][side], 0.5, missing=False):
                add_diff(ctx, "pageMargins", path + "/" + side, a["pageMargins"][side], b["pageMargins"][side],
                         f"page margin {side} {fmt_pt(a['pageMargins'][side].get('pt'))} -> {fmt_pt(b['pageMargins'][side].get('pt'))}")
        for side in ("header", "footer"):
            # Only compare direct edge distances; ODF header/footer box properties remain raw.
            if a["pageMargins"][side].get("pt") is not None and b["pageMargins"][side].get("pt") is not None and \
                    abs(a["pageMargins"][side]["pt"] - b["pageMargins"][side]["pt"]) > 0.5:
                add_diff(ctx, "headerFooterDistance", path + "/" + side, a["pageMargins"][side], b["pageMargins"][side],
                         f"{side} distance {fmt_pt(a['pageMargins'][side]['pt'])} -> {fmt_pt(b['pageMargins'][side]['pt'])}")
        agrid = {key: value for key, value in a["docGrid"].get("raw", {}).items() if value is not None}
        bgrid = {key: value for key, value in b["docGrid"].get("raw", {}).items() if value is not None}
        if (agrid or bgrid) and agrid != bgrid:
            add_diff(ctx, "grid", path, a["docGrid"], b["docGrid"], "stored document grid properties differ")


def compare_dumps(left: dict, right: dict) -> dict:
    labels = ("word", "odt") if left["format"] == "docx" else ("odtA", "odtB")
    ctx = {"leftLabel": labels[0], "rightLabel": labels[1], "leftShort": labels[0], "rightShort": labels[1],
           "counts": Counter({x: 0 for x in CATEGORIES}), "patterns": Counter(), "diffs": []}
    compare_document(ctx, left, right)
    compare_tables(ctx, left["tables"], right["tables"])
    return {"schemaVersion": 1, "left": {"format": left["format"], "file": left["file"]},
            "right": {"format": right["format"], "file": right["file"]},
            "categoryCounts": dict(ctx["counts"]), "patternCounts": dict(ctx["patterns"].most_common()),
            "diffs": ctx["diffs"], "unresolvedNotes": [
                "Word defaults absent from XML are not guessed.",
                "Only wholeTable OOXML conditional table formatting is applied; region/edge conditions are reported unresolved.",
                "An ODF automatic style whose parent is another automatic style is deliberately not followed.",
                "ODF header/footer box metrics are retained raw; no direct OOXML edge-distance equivalent is inferred.",
            ]}


def compare_docx_odt(docx: str | Path, odt: str | Path) -> dict:
    return compare_dumps(dump_docx(docx, True), dump_odt(odt, True))


def compare_odt_odt(left: str | Path, right: str | Path) -> dict:
    return compare_dumps(dump_odt(left, True), dump_odt(right, True))


def comparison_markdown(result: dict, title: str | None = None, diff_limit: int = 40) -> str:
    title = title or "Static layout-property comparison"
    lines = [f"# {title}", "", f"- Left: `{result['left']['file']}`", f"- Right: `{result['right']['file']}`", "",
             "## Category counts", "", "| Category | Count |", "|---|---:|"]
    lines += [f"| {key} | {value} |" for key, value in result["categoryCounts"].items()]
    lines += ["", "## Most frequent concrete patterns", "", "| Pattern | Count |", "|---|---:|"]
    patterns = sorted(result.get("patternCounts", {}).items(), key=lambda x: (-x[1], x[0]))[:10]
    lines += [f"| {pattern.replace('|', '&#124;')} | {count} |" for pattern, count in patterns]
    if not patterns:
        lines.append("| _(none)_ | 0 |")
    lines += ["", f"## First {min(diff_limit, len(result['diffs']))} differences", ""]
    for diff in result["diffs"][:diff_limit]:
        lines.append(f"- `{diff['category']}` `{diff['path']}` — {diff.get('pattern', diff.get('detail', 'mismatch'))}")
    lines += ["", "## Resolution limits", ""]
    lines += [f"- {note}" for note in result.get("unresolvedNotes", [])]
    return "\n".join(lines) + "\n"


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    p = sub.add_parser("dump-docx"); p.add_argument("file")
    p = sub.add_parser("dump-odt"); p.add_argument("file")
    p = sub.add_parser("compare"); p.add_argument("word"); p.add_argument("odt")
    p = sub.add_parser("compare-odt"); p.add_argument("left"); p.add_argument("right")
    args = parser.parse_args(argv)
    if args.command == "dump-docx":
        result = dump_docx(args.file)
    elif args.command == "dump-odt":
        result = dump_odt(args.file)
    elif args.command == "compare":
        result = compare_docx_odt(args.word, args.odt)
    else:
        result = compare_odt_odt(args.left, args.right)
    json.dump(result, sys.stdout, ensure_ascii=False, separators=(",", ":"))
    sys.stdout.write("\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
