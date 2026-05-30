import unittest
import xml.etree.ElementTree as ET
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZIP_DEFLATED, ZipFile

from dotex.tex_to_docx import (
  CitationTarget,
    WORD_ATTR_PREFIX,
    DocumentLayoutHints,
    TableLayoutHint,
    TemplateDocxHints,
    ZoteroDocxContext,
    build_caption_placeholder,
    build_initial_conversion_diagnostics,
    cleanup_table_cell,
    make_cross_reference_anchor,
    postprocess_generated_docx,
)


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def make_hints() -> TemplateDocxHints:
    return TemplateDocxHints(
        caption_style_id="Caption",
        table_style_id="DemoTable",
        table_paragraph_style_id="TableBody",
        normal_style_id="Normal",
        title_style_id="Title",
        heading_1_style_id="Heading1",
        heading_2_style_id="Heading2",
        heading_3_style_id="Heading3",
        bibliography_style_id="Bibliography",
        zotero_item_uri_prefix=None,
    )


def make_context() -> ZoteroDocxContext:
    return ZoteroDocxContext(
        bibliography_entries=[],
        unmatched_notices=[],
        by_anchor={},
        by_normalized_url={},
        by_normalized_doi={},
    )


class PostprocessDocxTests(unittest.TestCase):
    def test_postprocess_leaves_styles_part_untouched(self) -> None:
        document_xml = b"""<?xml version='1.0' encoding='utf-8'?>
<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:body>
    <w:tbl>
      <w:tr><w:tc><w:p/></w:tc></w:tr>
      <w:tr><w:tc><w:p/></w:tc></w:tr>
    </w:tbl>
  </w:body>
</w:document>
"""
        styles_xml = b"""<?xml version='1.0' encoding='utf-8'?>
<w:styles xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:style w:type='paragraph' w:styleId='Normal' w:default='1'>
    <w:name w:val='Normal'/>
  </w:style>
</w:styles>
"""

        with TemporaryDirectory() as temp_dir:
            docx_path = Path(temp_dir) / "sample.docx"
            with ZipFile(docx_path, "w", compression=ZIP_DEFLATED) as archive:
                archive.writestr("word/document.xml", document_xml)
                archive.writestr("word/styles.xml", styles_xml)

            diagnostics = build_initial_conversion_diagnostics(
                DocumentLayoutHints(
                    length_context={},
                    tables=[TableLayoutHint(centered=True, total_width_ratio=None, column_width_ratios=[], column_alignments=[])],
                    figures=[],
                )
            )

            postprocess_generated_docx(
                docx_path,
              docx_path,
                make_hints(),
                make_context(),
                DocumentLayoutHints(
                    length_context={},
                    tables=[TableLayoutHint(centered=True, total_width_ratio=None, column_width_ratios=[], column_alignments=[])],
                    figures=[],
                ),
                diagnostics,
            )

            with ZipFile(docx_path) as archive:
                self.assertEqual(archive.read("word/styles.xml"), styles_xml)
                processed_document = ET.fromstring(archive.read("word/document.xml"))

        table_borders = processed_document.find(".//w:tblPr/w:tblBorders", NS)
        self.assertIsNotNone(table_borders)
        assert table_borders is not None
        top_border = table_borders.find("w:top", NS)
        bottom_border = table_borders.find("w:bottom", NS)
        self.assertIsNotNone(top_border)
        self.assertIsNotNone(bottom_border)
        assert top_border is not None and bottom_border is not None
        self.assertEqual(top_border.get(f"{WORD_ATTR_PREFIX}val"), "single")
        self.assertEqual(bottom_border.get(f"{WORD_ATTR_PREFIX}val"), "single")

        first_cell = processed_document.find(".//w:tr[1]/w:tc[1]", NS)
        last_cell = processed_document.find(".//w:tr[last()]/w:tc[1]", NS)
        self.assertIsNotNone(first_cell)
        self.assertIsNotNone(last_cell)
        assert first_cell is not None and last_cell is not None
        first_cell_top = first_cell.find("./w:tcPr/w:tcBorders/w:top", NS)
        last_cell_bottom = last_cell.find("./w:tcPr/w:tcBorders/w:bottom", NS)
        self.assertIsNotNone(first_cell_top)
        self.assertIsNotNone(last_cell_bottom)
        assert first_cell_top is not None and last_cell_bottom is not None
        self.assertEqual(first_cell_top.get(f"{WORD_ATTR_PREFIX}val"), "single")
        self.assertEqual(last_cell_bottom.get(f"{WORD_ATTR_PREFIX}val"), "single")

    def test_cleanup_table_cell_strips_layout_commands_and_keeps_inline_math(self) -> None:
        cell = r"\begin{minipage}[b]{\linewidth}\raggedright $\Delta SEI$ = $SEI(T1)$ - $SEI(T0)$ {[]}\protect\phantomsection\end{minipage}"

        cleaned = cleanup_table_cell(cell)

        self.assertEqual(cleaned, "Δ SEI = SEI(T1) - SEI(T0) []")

    def test_postprocess_zotero_mode_keeps_hidden_caption_bookmarks_and_rewrites_crossrefs(self) -> None:
        label = "tab-1"
        caption_placeholder = build_caption_placeholder("table", label, "示例表格", "1")
        crossref_anchor = make_cross_reference_anchor(label)
        document_xml = f"""<?xml version='1.0' encoding='utf-8'?>
<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:body>
    <w:p>
      <w:r><w:t>{caption_placeholder}</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t xml:space='preserve'>见 </w:t></w:r>
      <w:hyperlink w:anchor='{crossref_anchor}'>
        <w:r><w:t>表 1</w:t></w:r>
      </w:hyperlink>
      <w:r><w:t xml:space='preserve'>。</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
""".encode("utf-8")
        styles_xml = b"""<?xml version='1.0' encoding='utf-8'?>
<w:styles xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:style w:type='paragraph' w:styleId='Normal' w:default='1'><w:name w:val='Normal'/></w:style>
  <w:style w:type='paragraph' w:styleId='Caption'><w:name w:val='Caption'/></w:style>
</w:styles>
"""

        with TemporaryDirectory() as temp_dir:
            docx_path = Path(temp_dir) / "sample.docx"
            with ZipFile(docx_path, "w", compression=ZIP_DEFLATED) as archive:
                archive.writestr("word/document.xml", document_xml)
                archive.writestr("word/styles.xml", styles_xml)

            diagnostics = build_initial_conversion_diagnostics(
                DocumentLayoutHints(length_context={}, tables=[], figures=[])
            )
            postprocess_generated_docx(
                docx_path,
              docx_path,
                make_hints(),
                make_context(),
                DocumentLayoutHints(length_context={}, tables=[], figures=[]),
                diagnostics,
                enable_zotero=True,
            )

            with ZipFile(docx_path) as archive:
                processed_document = ET.fromstring(archive.read("word/document.xml"))

        bookmark_starts = processed_document.findall(".//w:bookmarkStart", NS)
        bookmark_ends = processed_document.findall(".//w:bookmarkEnd", NS)
        self.assertEqual(len(bookmark_starts), 1)
        self.assertEqual(len(bookmark_ends), 1)
        self.assertRegex(bookmark_starts[0].get(f"{WORD_ATTR_PREFIX}name") or "", r"^_Ref\d+$")
        self.assertEqual(len(processed_document.findall(".//w:hyperlink", NS)), 0)
        instr_texts = [instr.text or "" for instr in processed_document.findall(".//w:instrText", NS)]
        self.assertTrue(any("SEQ 表" in text for text in instr_texts))
        self.assertTrue(any(text.startswith(" REF _Ref") and "\\h" in text for text in instr_texts))

        paragraphs = processed_document.findall(".//w:body/w:p", NS)
        texts = ["".join(t.text or "" for t in paragraph.findall(".//w:t", NS)).strip() for paragraph in paragraphs]
        self.assertIn("表 1 示例表格", texts[0])
        self.assertEqual(texts[1], "见 表 1。")

    def test_postprocess_zotero_mode_keeps_citation_fields_in_paragraphs_with_crossrefs(self) -> None:
        label = "fig-1"
        caption_placeholder = build_caption_placeholder("figure", label, "概念模型", "1")
        crossref_anchor = make_cross_reference_anchor(label)
        document_xml = f"""<?xml version='1.0' encoding='utf-8'?>
<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:body>
    <w:p>
      <w:r><w:t>{caption_placeholder}</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t xml:space='preserve'>见 </w:t></w:r>
      <w:hyperlink w:anchor='Appleton_2006'>
        <w:r><w:t>(Appleton et al. 2006)</w:t></w:r>
      </w:hyperlink>
      <w:r><w:t xml:space='preserve'> 与 </w:t></w:r>
      <w:hyperlink w:anchor='{crossref_anchor}'>
        <w:r><w:t>图 1</w:t></w:r>
      </w:hyperlink>
      <w:r><w:t>。</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
""".encode("utf-8")
        styles_xml = b"""<?xml version='1.0' encoding='utf-8'?>
<w:styles xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:style w:type='paragraph' w:styleId='Normal' w:default='1'><w:name w:val='Normal'/></w:style>
  <w:style w:type='paragraph' w:styleId='Caption'><w:name w:val='Caption'/></w:style>
</w:styles>
"""
        context = make_context()
        context.by_anchor["Appleton_2006"] = CitationTarget(
            source_key="Appleton_2006",
            formatted_reference="Appleton et al. 2006",
            zotero_item_key="APPLETON2006",
            item_data={"id": 1567, "title": "Appleton et al. 2006"},
            uri="http://zotero.org/users/14586934/items/APPLETON2006",
            anchor_id="Appleton_2006",
        )

        with TemporaryDirectory() as temp_dir:
            docx_path = Path(temp_dir) / "sample.docx"
            with ZipFile(docx_path, "w", compression=ZIP_DEFLATED) as archive:
                archive.writestr("word/document.xml", document_xml)
                archive.writestr("word/styles.xml", styles_xml)

            diagnostics = build_initial_conversion_diagnostics(
                DocumentLayoutHints(length_context={}, tables=[], figures=[])
            )
            postprocess_generated_docx(
                docx_path,
              docx_path,
                make_hints(),
                context,
                DocumentLayoutHints(length_context={}, tables=[], figures=[]),
                diagnostics,
                enable_zotero=True,
            )

            with ZipFile(docx_path) as archive:
                processed_document = ET.fromstring(archive.read("word/document.xml"))

        instr_texts = [instr.text or "" for instr in processed_document.findall(".//w:instrText", NS)]
        self.assertTrue(any("ADDIN ZOTERO_ITEM CSL_CITATION" in text for text in instr_texts))
        self.assertTrue(any(text.startswith(" REF _Ref") and "\\h" in text for text in instr_texts))


if __name__ == "__main__":
    unittest.main()