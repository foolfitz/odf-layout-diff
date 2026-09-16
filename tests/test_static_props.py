# SPDX-License-Identifier: MPL-2.0
# Copyright 2026 OSSII
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from experiments import static_props as sp


W = sp.NS["w"]
O = sp.NS["office"]
S = sp.NS["style"]
T = sp.NS["table"]
X = sp.NS["text"]
F = sp.NS["fo"]


def package(path, members):
    with ZipFile(path, "w") as archive:
        for name, value in members.items():
            archive.writestr(name, value)


def wtag(name, attrs="", body=""):
    return f"<w:{name}{attrs}>{body}</w:{name}>"


def docx_xml(tables):
    return f'''<w:document xmlns:w="{W}"><w:body>{tables}</w:body></w:document>'''


def word_table(tc_mar="", row_mar="", table_mar="", tc_border="", row_border="",
               table_border="", height="", span=""):
    tbl_pr = f'''<w:tblPr><w:tblStyle w:val="TS"/>{table_mar}{table_border}</w:tblPr>'''
    tr_pr = f"<w:trPr>{height}</w:trPr>"
    ex = f"<w:tblPrEx>{row_mar}{row_border}</w:tblPrEx>" if row_mar or row_border else ""
    tc_pr = f"<w:tcPr>{span}{tc_mar}{tc_border}</w:tcPr>"
    return f'''<w:tbl>{tbl_pr}<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>
      <w:tr>{tr_pr}{ex}<w:tc>{tc_pr}<w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'''


WORD_STYLES = f'''<w:styles xmlns:w="{W}">
  <w:style w:type="table" w:styleId="TS"><w:tblPr>
    <w:tblCellMar><w:left w:w="100" w:type="dxa"/></w:tblCellMar>
    <w:tblBorders><w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders>
  </w:tblPr></w:style>
</w:styles>'''


def margin(value):
    return f'<w:tblCellMar><w:left w:w="{value}" w:type="dxa"/></w:tblCellMar>'


def cell_margin(value):
    return f'<w:tcMar><w:left w:w="{value}" w:type="dxa"/></w:tcMar>'


def border(value):
    return f'<w:tblBorders><w:left w:val="single" w:sz="{value}"/></w:tblBorders>'


def cell_border(value):
    return f'<w:tcBorders><w:left w:val="double" w:sz="{value}"/></w:tcBorders>'


ODT_STYLES = f'''<office:document-styles xmlns:office="{O}" xmlns:style="{S}" xmlns:fo="{F}">
  <office:styles>
    <style:default-style style:family="table-cell"><style:table-cell-properties fo:padding="1pt" fo:border="0.5pt solid #000000" fo:background-color="#00ff00"/></style:default-style>
  </office:styles>
</office:document-styles>'''


def odt_content(rows, auto_styles):
    return f'''<office:document-content xmlns:office="{O}" xmlns:style="{S}" xmlns:table="{T}" xmlns:text="{X}" xmlns:fo="{F}">
      <office:automatic-styles>{auto_styles}</office:automatic-styles>
      <office:body><office:text><table:table table:name="T">{rows}</table:table></office:text></office:body>
    </office:document-content>'''


class StaticPropsTests(unittest.TestCase):
    def test_unit_conversions(self):
        self.assertEqual(sp.twips_to_pt("400"), 20.0)
        self.assertEqual(sp.eighths_to_pt("6"), 0.75)
        self.assertAlmostEqual(sp.length_to_pt("1in"), 72.0)
        self.assertAlmostEqual(sp.length_to_pt("2.54cm"), 72.0)
        self.assertAlmostEqual(sp.length_to_pt("25.4mm"), 72.0)
        self.assertAlmostEqual(sp.length_to_pt("9pt"), 9.0)
        self.assertIsNone(sp.length_to_pt("50%"))

    def test_word_row_height_margin_and_border_precedence(self):
        direct = word_table(cell_margin("400"), margin("300"), margin("200"),
                            cell_border("16"), border("12"), border("8"),
                            '<w:trHeight w:val="400" w:hRule="exact"/>')
        row = word_table("", margin("300"), margin("200"), "", border("12"), border("8"),
                         '<w:trHeight w:val="280" w:hRule="atLeast"/>')
        table = word_table("", "", margin("200"), "", "", border("8"))
        style = word_table()
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "x.docx"
            package(path, {"word/document.xml": docx_xml(direct + row + table + style),
                           "word/styles.xml": WORD_STYLES})
            result = sp.dump_docx(path)
        self.assertEqual(result["tables"][0]["rows"][0]["height"]["kind"], "exact")
        self.assertEqual(result["tables"][0]["rows"][0]["height"]["pt"], 20.0)
        self.assertEqual(result["tables"][1]["rows"][0]["height"]["kind"], "min")
        self.assertEqual(result["tables"][1]["rows"][0]["height"]["pt"], 14.0)
        self.assertEqual(result["tables"][2]["rows"][0]["height"]["kind"], "none")
        cells = [item["rows"][0]["cells"][0] for item in result["tables"]]
        self.assertEqual([c["margins"]["left"]["raw"] for c in cells], ["400", "300", "200", "100"])
        self.assertEqual([c["margins"]["left"]["source"] for c in cells],
                         ["direct cell", "row tblPrEx", "direct table", "table style TS"])
        self.assertEqual([c["borders"]["left"]["widthPt"] for c in cells], [2.0, 1.5, 1.0, 0.5])
        self.assertEqual(cells[0]["borders"]["left"]["style"], "double")

    def test_odt_row_height_padding_border_and_automatic_parent(self):
        auto = '''
          <style:style style:name="r1" style:family="table-row"><style:table-row-properties style:row-height="1in"/></style:style>
          <style:style style:name="r2" style:family="table-row"><style:table-row-properties style:min-row-height="10mm"/></style:style>
          <style:style style:name="baseAuto" style:family="table-cell"><style:table-cell-properties fo:padding="9pt"/></style:style>
          <style:style style:name="c1" style:family="table-cell" style:parent-style-name="baseAuto"><style:table-cell-properties fo:padding="2pt" fo:padding-left="3pt" fo:border="0.75pt dotted #ff0000"/></style:style>'''
        rows = '''
          <table:table-row table:style-name="r1"><table:table-cell table:style-name="c1"><text:p>x</text:p></table:table-cell></table:table-row>
          <table:table-row table:style-name="r2"><table:table-cell><text:p>y</text:p></table:table-cell></table:table-row>
          <table:table-row><table:table-cell><text:p>z</text:p></table:table-cell></table:table-row>'''
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "x.odt"
            package(path, {"content.xml": odt_content(rows, auto), "styles.xml": ODT_STYLES})
            result = sp.dump_odt(path)
        table = result["tables"][0]
        self.assertEqual(table["rows"][0]["height"]["kind"], "exact")
        self.assertEqual(table["rows"][0]["height"]["pt"], 72.0)
        self.assertEqual(table["rows"][1]["height"]["kind"], "min")
        self.assertAlmostEqual(table["rows"][1]["height"]["pt"], 72 / 2.54, places=3)
        self.assertEqual(table["rows"][2]["height"]["kind"], "none")
        cell = table["rows"][0]["cells"][0]
        self.assertEqual(cell["margins"]["left"]["pt"], 3.0)
        self.assertEqual(cell["margins"]["right"]["pt"], 2.0)
        self.assertEqual(cell["borders"]["top"]["style"], "dotted")
        self.assertEqual(cell["borders"]["top"]["widthPt"], 0.75)
        self.assertTrue(cell["parentIsAutomatic"])
        self.assertIsNone(cell["backgroundColor"]["value"])

    def test_span_and_covered_cell_column_alignment(self):
        word = word_table(span='<w:gridSpan w:val="2"/>')
        # Add a following Word cell after the spanning anchor.
        word = word.replace('</w:tr></w:tbl>', '<w:tc><w:tcPr/><w:p><w:r><w:t>next</w:t></w:r></w:p></w:tc></w:tr></w:tbl>')
        rows = '''<table:table-row>
          <table:table-cell table:number-columns-spanned="2"><text:p>x</text:p></table:table-cell>
          <table:covered-table-cell/>
          <table:table-cell><text:p>next</text:p></table:table-cell>
        </table:table-row>'''
        with tempfile.TemporaryDirectory() as directory:
            wpath, opath = Path(directory) / "x.docx", Path(directory) / "x.odt"
            package(wpath, {"word/document.xml": docx_xml(word), "word/styles.xml": WORD_STYLES})
            package(opath, {"content.xml": odt_content(rows, ""), "styles.xml": ODT_STYLES})
            wd, od = sp.dump_docx(wpath), sp.dump_odt(opath)
            compared = sp.compare_dumps(sp.dump_docx(wpath, True), sp.dump_odt(opath, True))
        wcells = wd["tables"][0]["rows"][0]["cells"]
        ocells = od["tables"][0]["rows"][0]["cells"]
        self.assertEqual([c["gridColumnStart"] for c in wcells], [0, 2])
        self.assertEqual([c["gridColumnStart"] for c in ocells], [0, 1, 2])
        self.assertTrue(ocells[1]["covered"])
        self.assertEqual([c["gridColumnStart"] for c in ocells if not c["covered"]], [0, 2])
        self.assertEqual(compared["categoryCounts"]["unmatchedCell"], 0)


if __name__ == "__main__":
    unittest.main()
