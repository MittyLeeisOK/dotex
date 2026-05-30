import unittest
import xml.etree.ElementTree as ET
from pathlib import Path
from tempfile import TemporaryDirectory

from dotex.converter import (
    REL_ATTR_PREFIX,
    THREE_LINE_OUTER_BORDER_SIZE,
    WORD_ATTR_PREFIX,
    CitationFieldShell,
    CitationTarget,
    DocumentLayoutHints,
    ZoteroDocxContext,
    TABLE_LAYOUT_NOTICE,
    TableLayoutHint,
    TemplateDocxHints,
    apply_bibliography_hints,
    apply_table_hints,
    build_zotero_docx_context,
    build_initial_conversion_diagnostics,
    build_zotero_citation_field_elements,
    convert_citation_hyperlinks_to_zotero_fields,
    normalize_tex_for_pandoc,
    normalize_internal_anchor_bookmarks,
    normalize_tree_run_fonts,
    prune_unused_hyperlink_relationships_for_part,
    resolve_citation_hyperlink_target,
    strip_zotero_field_run_fonts,
    style_default_internal_hyperlinks,
    strip_all_bookmarks,
    strip_internal_hyperlink_styles,
)
from dotex.zotero_resolver import build_zotero_item_uri, parse_bibliography_entries


def make_context() -> ZoteroDocxContext:
    return ZoteroDocxContext(
        bibliography_entries=[],
        unmatched_notices=[],
        by_anchor={},
        by_normalized_url={},
        by_normalized_doi={},
    )


class ZoteroModeTests(unittest.TestCase):
    def test_normalize_tex_for_pandoc_uses_internal_anchor_links_for_citations(self) -> None:
        with TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            tex_path = root / "manuscript.tex"
            tex_path.write_text(
                "\\begin{document}\\parencite{reeve2009}\\section{参考文献}\\input{refs.bib}\\end{document}\n",
                encoding="utf-8",
            )
            (root / "refs_display.json").write_text('{"reeve2009": "Reeve 2009"}', encoding="utf-8")
            (root / "refs.bib").write_text(
                "@article{reeve2009,\n  title = {Demo title},\n  author = {Reeve, John},\n  year = {2009}\n}\n",
                encoding="utf-8",
            )

            normalized = normalize_tex_for_pandoc(tex_path, use_citation_hyperlinks=True)

        self.assertIn("[Reeve 2009](#", normalized)

    def test_build_zotero_docx_context_can_use_direct_companion_without_database(self) -> None:
        with TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            tex_path = root / "manuscript.tex"
            tex_path.write_text(
                "\\begin{document}\\parencite{reeve2009}\\section{参考文献}\\input{refs.bib}\\end{document}\n",
                encoding="utf-8",
            )
            (root / "refs_display.json").write_text('{"reeve2009": "Reeve 2009"}', encoding="utf-8")
            (root / "refs.bib").write_text(
                "@article{reeve2009,\n  title = {Demo title},\n  author = {Reeve, John},\n  year = {2009},\n  doi = {10.1000/demo}\n}\n",
                encoding="utf-8",
            )
            (root / "dotex_zotero_items.json").write_text(
                """
{
  "version": 1,
  "items": [
    {
      "key": "reeve2009",
      "source_key": "10.1000/demo",
      "formatted_reference": "Reeve 2009",
      "zotero_item_key": "ABCD1234",
      "uri": "http://zotero.org/users/local/items/ABCD1234",
      "item_data": {
        "id": 123,
        "type": "article-journal",
        "title": "Demo title"
      }
    }
  ]
}
""".strip(),
                encoding="utf-8",
            )

            context = build_zotero_docx_context(
                tex_path,
                TemplateDocxHints(
                    caption_style_id=None,
                    table_style_id=None,
                    table_paragraph_style_id=None,
                    normal_style_id=None,
                    title_style_id="Title",
                    heading_1_style_id="Heading1",
                    heading_2_style_id="Heading2",
                    heading_3_style_id="Heading3",
                    bibliography_style_id=None,
                    zotero_item_uri_prefix=None,
                ),
                source_text=tex_path.read_text(encoding="utf-8"),
                enable_zotero=True,
                zotero_database=root / "missing.sqlite",
            )

        self.assertEqual(len(context.bibliography_entries), 1)
        self.assertEqual(context.bibliography_entries[0].item_data["id"], 123)
        self.assertEqual(context.bibliography_entries[0].uri, "http://zotero.org/users/local/items/ABCD1234")

    def test_table_notice_added_when_tables_present(self) -> None:
        diagnostics = build_initial_conversion_diagnostics(
            DocumentLayoutHints(
                length_context={},
                tables=[TableLayoutHint(centered=True, total_width_ratio=None, column_width_ratios=[], column_alignments=[])],
                figures=[],
            )
        )

        self.assertEqual(diagnostics.notices, [TABLE_LAYOUT_NOTICE])

    def test_table_cells_are_left_aligned_while_table_is_centered(self) -> None:
        document = ET.Element(f"{WORD_ATTR_PREFIX}document")
        body = ET.SubElement(document, f"{WORD_ATTR_PREFIX}body")
        table = ET.SubElement(body, f"{WORD_ATTR_PREFIX}tbl")
        row = ET.SubElement(table, f"{WORD_ATTR_PREFIX}tr")
        cell = ET.SubElement(row, f"{WORD_ATTR_PREFIX}tc")
        paragraph = ET.SubElement(cell, f"{WORD_ATTR_PREFIX}p")
        paragraph_properties = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}pPr")
        paragraph_alignment = ET.SubElement(paragraph_properties, f"{WORD_ATTR_PREFIX}jc")
        paragraph_alignment.set(f"{WORD_ATTR_PREFIX}val", "center")

        changed = apply_table_hints(
            document,
            TemplateDocxHints(
                caption_style_id=None,
                table_style_id="DemoTable",
                table_paragraph_style_id="TableBody",
                normal_style_id=None,
                title_style_id="Title",
                heading_1_style_id="Heading1",
                heading_2_style_id="Heading2",
                heading_3_style_id="Heading3",
                bibliography_style_id=None,
                zotero_item_uri_prefix=None,
            ),
            DocumentLayoutHints(length_context={}, tables=[], figures=[]),
        )

        self.assertTrue(changed)
        table_style = table.find("./w:tblPr/w:tblStyle", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
        self.assertIsNotNone(table_style)
        assert table_style is not None
        self.assertEqual(table_style.get(f"{WORD_ATTR_PREFIX}val"), "DemoTable")
        table_look = table.find("./w:tblPr/w:tblLook", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
        self.assertIsNotNone(table_look)
        assert table_look is not None
        self.assertEqual(table_look.get(f"{WORD_ATTR_PREFIX}val"), "0020")
        self.assertEqual(table_look.get(f"{WORD_ATTR_PREFIX}firstRow"), "1")
        self.assertEqual(table_look.get(f"{WORD_ATTR_PREFIX}noHBand"), "0")
        self.assertEqual(table_look.get(f"{WORD_ATTR_PREFIX}noVBand"), "0")
        table_alignment = table.find("./w:tblPr/w:jc", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
        self.assertIsNotNone(table_alignment)
        assert table_alignment is not None
        self.assertEqual(table_alignment.get(f"{WORD_ATTR_PREFIX}val"), "center")
        updated_alignment = paragraph.find("./w:pPr/w:jc", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
        self.assertIsNotNone(updated_alignment)
        assert updated_alignment is not None
        self.assertEqual(updated_alignment.get(f"{WORD_ATTR_PREFIX}val"), "left")

    def test_table_hints_prefer_style_defined_outer_borders(self) -> None:
        document = ET.Element(f"{WORD_ATTR_PREFIX}document")
        body = ET.SubElement(document, f"{WORD_ATTR_PREFIX}body")
        table = ET.SubElement(body, f"{WORD_ATTR_PREFIX}tbl")
        header_row = ET.SubElement(table, f"{WORD_ATTR_PREFIX}tr")
        header_cell = ET.SubElement(header_row, f"{WORD_ATTR_PREFIX}tc")
        ET.SubElement(header_cell, f"{WORD_ATTR_PREFIX}p")
        body_row = ET.SubElement(table, f"{WORD_ATTR_PREFIX}tr")
        body_cell = ET.SubElement(body_row, f"{WORD_ATTR_PREFIX}tc")
        ET.SubElement(body_cell, f"{WORD_ATTR_PREFIX}p")

        changed = apply_table_hints(
            document,
            TemplateDocxHints(
                caption_style_id=None,
                table_style_id="DemoTable",
                table_paragraph_style_id="TableBody",
                normal_style_id=None,
                title_style_id="Title",
                heading_1_style_id="Heading1",
                heading_2_style_id="Heading2",
                heading_3_style_id="Heading3",
                bibliography_style_id=None,
                zotero_item_uri_prefix=None,
            ),
            DocumentLayoutHints(length_context={}, tables=[], figures=[]),
        )

        self.assertTrue(changed)
        table_borders = table.find("./w:tblPr/w:tblBorders", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
        self.assertIsNotNone(table_borders)
        assert table_borders is not None
        top_border = table_borders.find("./w:top", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
        bottom_border = table_borders.find("./w:bottom", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
        self.assertIsNotNone(top_border)
        self.assertIsNotNone(bottom_border)
        assert top_border is not None and bottom_border is not None
        self.assertEqual(top_border.get(f"{WORD_ATTR_PREFIX}val"), "single")
        self.assertEqual(bottom_border.get(f"{WORD_ATTR_PREFIX}val"), "single")
        self.assertEqual(top_border.get(f"{WORD_ATTR_PREFIX}sz"), THREE_LINE_OUTER_BORDER_SIZE)
        self.assertEqual(bottom_border.get(f"{WORD_ATTR_PREFIX}sz"), THREE_LINE_OUTER_BORDER_SIZE)

        header_bottom = header_cell.find("./w:tcPr/w:tcBorders/w:bottom", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
        body_bottom = body_cell.find("./w:tcPr/w:tcBorders/w:bottom", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
        self.assertIsNotNone(header_bottom)
        self.assertIsNotNone(body_bottom)
        assert header_bottom is not None and body_bottom is not None
        self.assertEqual(header_bottom.get(f"{WORD_ATTR_PREFIX}val"), "single")
        self.assertEqual(body_bottom.get(f"{WORD_ATTR_PREFIX}val"), "single")

        def test_unused_hyperlink_relationships_are_pruned_for_document_part(self) -> None:
                document_xml = f"""<?xml version='1.0' encoding='utf-8'?>
<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                        xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>
    <w:body>
        <w:p>
            <w:hyperlink r:id='rId1'><w:r><w:t>Used</w:t></w:r></w:hyperlink>
        </w:p>
    </w:body>
</w:document>
""".encode("utf-8")
                rels_xml = b"""<?xml version='1.0' encoding='utf-8'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
    <Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink' Target='https://example.com/used' TargetMode='External'/>
    <Relationship Id='rId2' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink' Target='https://example.com/unused' TargetMode='External'/>
    <Relationship Id='rId3' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles' Target='styles.xml'/>
</Relationships>
"""

                pruned_xml, changed = prune_unused_hyperlink_relationships_for_part(document_xml, rels_xml)

                self.assertTrue(changed)
                self.assertIsNotNone(pruned_xml)
                assert pruned_xml is not None
                root = ET.fromstring(pruned_xml)
                rel_ids = [rel.get("Id") for rel in root]
                self.assertEqual(rel_ids, ["rId1", "rId3"])

        def test_empty_hyperlink_relationship_part_is_removed(self) -> None:
                footnotes_xml = b"""<?xml version='1.0' encoding='utf-8'?>
<w:footnotes xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
    <w:footnote w:type='separator' w:id='-1'><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
</w:footnotes>
"""
                rels_xml = b"""<?xml version='1.0' encoding='utf-8'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
    <Relationship Id='rId99' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink' Target='https://example.com/unused' TargetMode='External'/>
</Relationships>
"""

                pruned_xml, changed = prune_unused_hyperlink_relationships_for_part(footnotes_xml, rels_xml)

                self.assertTrue(changed)
                self.assertIsNone(pruned_xml)

    def test_grouped_header_rows_are_merged_and_border_second_header_row(self) -> None:
        document = ET.Element(f"{WORD_ATTR_PREFIX}document")
        body = ET.SubElement(document, f"{WORD_ATTR_PREFIX}body")
        table = ET.SubElement(body, f"{WORD_ATTR_PREFIX}tbl")

        header_row_1 = ET.SubElement(table, f"{WORD_ATTR_PREFIX}tr")
        for value in ["组别", "SEI (T0)", "", "LCQ (T1)", ""]:
            cell = ET.SubElement(header_row_1, f"{WORD_ATTR_PREFIX}tc")
            paragraph = ET.SubElement(cell, f"{WORD_ATTR_PREFIX}p")
            if value:
                run = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}r")
                text = ET.SubElement(run, f"{WORD_ATTR_PREFIX}t")
                text.text = value

        header_row_2 = ET.SubElement(table, f"{WORD_ATTR_PREFIX}tr")
        for value in ["", "Mean", "SD", "Mean", "SD"]:
            cell = ET.SubElement(header_row_2, f"{WORD_ATTR_PREFIX}tc")
            paragraph = ET.SubElement(cell, f"{WORD_ATTR_PREFIX}p")
            if value:
                run = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}r")
                text = ET.SubElement(run, f"{WORD_ATTR_PREFIX}t")
                text.text = value

        body_row = ET.SubElement(table, f"{WORD_ATTR_PREFIX}tr")
        for value in ["对照组", "3.700", "0.535", "4.947", "1.108"]:
            cell = ET.SubElement(body_row, f"{WORD_ATTR_PREFIX}tc")
            paragraph = ET.SubElement(cell, f"{WORD_ATTR_PREFIX}p")
            run = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}r")
            text = ET.SubElement(run, f"{WORD_ATTR_PREFIX}t")
            text.text = value

        changed = apply_table_hints(
            document,
            TemplateDocxHints(
                caption_style_id=None,
                table_style_id="DemoTable",
                table_paragraph_style_id="TableBody",
                normal_style_id=None,
                title_style_id="Title",
                heading_1_style_id="Heading1",
                heading_2_style_id="Heading2",
                heading_3_style_id="Heading3",
                bibliography_style_id=None,
                zotero_item_uri_prefix=None,
            ),
            DocumentLayoutHints(length_context={}, tables=[], figures=[]),
        )

        self.assertTrue(changed)
        merged_first_row_cells = header_row_1.findall("w:tc", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
        self.assertEqual(len(merged_first_row_cells), 3)
        self.assertEqual(
            merged_first_row_cells[1].find("./w:tcPr/w:gridSpan", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}).get(f"{WORD_ATTR_PREFIX}val"),
            "2",
        )
        self.assertEqual(
            merged_first_row_cells[2].find("./w:tcPr/w:gridSpan", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}).get(f"{WORD_ATTR_PREFIX}val"),
            "2",
        )

        first_header_vmerge = merged_first_row_cells[0].find("./w:tcPr/w:vMerge", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
        second_row_first_cell = header_row_2.findall("w:tc", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})[0]
        second_header_vmerge = second_row_first_cell.find("./w:tcPr/w:vMerge", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
        self.assertEqual(first_header_vmerge.get(f"{WORD_ATTR_PREFIX}val"), "restart")
        self.assertEqual(second_header_vmerge.get(f"{WORD_ATTR_PREFIX}val"), "continue")

        first_row_group_bottom = merged_first_row_cells[1].find("./w:tcPr/w:tcBorders/w:bottom", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
        second_row_bottom = second_row_first_cell.find("./w:tcPr/w:tcBorders/w:bottom", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
        self.assertEqual(first_row_group_bottom.get(f"{WORD_ATTR_PREFIX}val"), "single")
        self.assertEqual(second_row_bottom.get(f"{WORD_ATTR_PREFIX}val"), "single")

    def test_fallback_target_without_anchor_or_relationship(self) -> None:
        paragraph = ET.Element(f"{WORD_ATTR_PREFIX}p")
        hyperlink = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}hyperlink")
        run = ET.SubElement(hyperlink, f"{WORD_ATTR_PREFIX}r")
        text = ET.SubElement(run, f"{WORD_ATTR_PREFIX}t")
        text.text = "Smith 2020"

        context = make_context()
        target = resolve_citation_hyperlink_target(hyperlink, {}, context)

        self.assertIsNotNone(target)
        assert target is not None
        self.assertTrue(target.source_key.startswith("inline-cite-Smith-2020"))
        self.assertEqual(target.formatted_reference, "Smith 2020")
        self.assertEqual(target.item_data.get("title"), "Smith 2020")
        self.assertIn(target.source_key, context.by_anchor)

    def test_zotero_fields_are_not_inserted_inside_tables(self) -> None:
        document = ET.Element(f"{WORD_ATTR_PREFIX}document")
        body = ET.SubElement(document, f"{WORD_ATTR_PREFIX}body")
        table = ET.SubElement(body, f"{WORD_ATTR_PREFIX}tbl")
        row = ET.SubElement(table, f"{WORD_ATTR_PREFIX}tr")
        cell = ET.SubElement(row, f"{WORD_ATTR_PREFIX}tc")
        paragraph = ET.SubElement(cell, f"{WORD_ATTR_PREFIX}p")
        hyperlink = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}hyperlink")
        hyperlink.set(f"{WORD_ATTR_PREFIX}anchor", "Reeve_2009")
        run = ET.SubElement(hyperlink, f"{WORD_ATTR_PREFIX}r")
        text = ET.SubElement(run, f"{WORD_ATTR_PREFIX}t")
        text.text = "Reeve 2009"

        context = make_context()
        target = CitationTarget(
            source_key="Reeve_2009",
            formatted_reference="Reeve 2009",
            zotero_item_key="REEVE2009",
            item_data={"id": 1, "title": "Reeve 2009"},
            uri="http://zotero.org/users/1/items/REEVE2009",
            anchor_id="Reeve_2009",
        )
        context.by_anchor["Reeve_2009"] = target

        changed = convert_citation_hyperlinks_to_zotero_fields(document, {}, context)

        self.assertFalse(changed)
        self.assertIsNotNone(paragraph.find("./w:hyperlink", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}))
        self.assertEqual(len(paragraph.findall(".//w:instrText", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})), 0)

    def test_zotero_fields_are_not_inserted_into_paragraphs_with_existing_fields(self) -> None:
        document = ET.Element(f"{WORD_ATTR_PREFIX}document")
        body = ET.SubElement(document, f"{WORD_ATTR_PREFIX}body")
        paragraph = ET.SubElement(body, f"{WORD_ATTR_PREFIX}p")

        ref_run = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}r")
        fld_char = ET.SubElement(ref_run, f"{WORD_ATTR_PREFIX}fldChar")
        fld_char.set(f"{WORD_ATTR_PREFIX}fldCharType", "begin")

        hyperlink = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}hyperlink")
        hyperlink.set(f"{WORD_ATTR_PREFIX}anchor", "TenHove_2024")
        run = ET.SubElement(hyperlink, f"{WORD_ATTR_PREFIX}r")
        text = ET.SubElement(run, f"{WORD_ATTR_PREFIX}t")
        text.text = "Ten Hove et al. 2024"

        context = make_context()
        target = CitationTarget(
            source_key="TenHove_2024",
            formatted_reference="Ten Hove et al. 2024",
            zotero_item_key="TENHOVE2024",
            item_data={"id": 1, "title": "Ten Hove et al. 2024"},
            uri="http://zotero.org/users/1/items/TENHOVE2024",
            anchor_id="TenHove_2024",
        )
        context.by_anchor["TenHove_2024"] = target

        changed = convert_citation_hyperlinks_to_zotero_fields(document, {}, context)

        self.assertFalse(changed)
        self.assertIsNotNone(paragraph.find("./w:hyperlink", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}))
        self.assertEqual(len(paragraph.findall(".//w:instrText", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})), 0)

    def test_convert_citation_hyperlinks_prefers_preserved_field_shell(self) -> None:
        document = ET.Element(f"{WORD_ATTR_PREFIX}document")
        body = ET.SubElement(document, f"{WORD_ATTR_PREFIX}body")
        paragraph = ET.SubElement(body, f"{WORD_ATTR_PREFIX}p")
        hyperlink = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}hyperlink")
        hyperlink.set(f"{WORD_ATTR_PREFIX}anchor", "Reeve_2009")
        run = ET.SubElement(hyperlink, f"{WORD_ATTR_PREFIX}r")
        text = ET.SubElement(run, f"{WORD_ATTR_PREFIX}t")
        text.text = "(Reeve 2009)"

        context = make_context()
        target = CitationTarget(
            source_key="Reeve_2009",
            formatted_reference="Reeve 2009",
            zotero_item_key="REEVE2009",
            item_data={"id": 1, "title": "Reeve 2009"},
            uri="http://zotero.org/users/1/items/REEVE2009",
            anchor_id="Reeve_2009",
        )
        context.by_anchor["Reeve_2009"] = target
        context.citation_field_shells[(("Reeve_2009",), "(Reeve 2009)")] = [
            CitationFieldShell(
                field_nodes_xml=[
                    '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:color w:val="003399"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>',
                    '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:color w:val="003399"/><w:lang w:eastAsia="zh-CN"/></w:rPr><w:instrText xml:space="preserve"> ADDIN ZOTERO_ITEM CSL_CITATION {"citationID":"cite-1","properties":{"formattedCitation":"(Reeve 2009)","plainCitation":"Reeve 2009"}} </w:instrText></w:r>',
                    '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:color w:val="003399"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>',
                    '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:color w:val="003399"/></w:rPr><w:t>(Reeve 2009)</w:t></w:r>',
                    '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:color w:val="003399"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>',
                ]
            )
        ]

        changed = convert_citation_hyperlinks_to_zotero_fields(document, {}, context)

        self.assertTrue(changed)
        self.assertEqual(len(paragraph.findall("./w:hyperlink", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})), 0)
        instr_run = paragraph.findall("./w:r", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})[1]
        lang = instr_run.find(f"{WORD_ATTR_PREFIX}rPr/{WORD_ATTR_PREFIX}lang")
        self.assertIsNotNone(lang)
        assert lang is not None
        self.assertEqual(lang.get(f"{WORD_ATTR_PREFIX}eastAsia"), "zh-CN")
        instr_text = "".join(
            (run.find(f"{WORD_ATTR_PREFIX}instrText").text or "")
            for run in paragraph.findall("./w:r", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
            if run.find(f"{WORD_ATTR_PREFIX}instrText") is not None
        )
        self.assertIn('"plainCitation":"(Reeve 2009)"', instr_text)
        self.assertIn('"noteIndex":0', instr_text)
        self.assertNotIn('"dontUpdate"', instr_text)
        self.assertRegex(instr_text, r'"citationID":"[A-Za-z0-9]{8}"')
        self.assertNotIn('"citationID":"cite-1"', instr_text)
        self.assertEqual(context.citation_field_shells, {})

    def test_bookmarks_are_stripped(self) -> None:
        document = ET.Element(f"{WORD_ATTR_PREFIX}document")
        body = ET.SubElement(document, f"{WORD_ATTR_PREFIX}body")
        paragraph = ET.SubElement(body, f"{WORD_ATTR_PREFIX}p")

        bookmark_start = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}bookmarkStart")
        bookmark_start.set(f"{WORD_ATTR_PREFIX}id", "0")
        bookmark_start.set(f"{WORD_ATTR_PREFIX}name", "_Ref1")
        bookmark_end = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}bookmarkEnd")
        bookmark_end.set(f"{WORD_ATTR_PREFIX}id", "0")

        changed = strip_all_bookmarks(document)

        self.assertTrue(changed)
        self.assertEqual(len(document.findall('.//w:bookmarkStart', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})), 0)
        self.assertEqual(len(document.findall('.//w:bookmarkEnd', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})), 0)

    def test_internal_anchor_hyperlinks_are_flattened(self) -> None:
        document = ET.Element(f"{WORD_ATTR_PREFIX}document")
        body = ET.SubElement(document, f"{WORD_ATTR_PREFIX}body")
        paragraph = ET.SubElement(body, f"{WORD_ATTR_PREFIX}p")

        hyperlink = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}hyperlink")
        hyperlink.set(f"{WORD_ATTR_PREFIX}anchor", "ref-1")
        run = ET.SubElement(hyperlink, f"{WORD_ATTR_PREFIX}r")
        text = ET.SubElement(run, f"{WORD_ATTR_PREFIX}t")
        text.text = "Smith 2020"

        changed = strip_internal_hyperlink_styles(document)

        self.assertTrue(changed)
        self.assertEqual(len(document.findall('.//w:hyperlink', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})), 0)
        self.assertEqual(len(document.findall('.//w:t', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})), 1)

    def test_default_mode_keeps_only_reference_bookmarks(self) -> None:
        document = ET.Element(f"{WORD_ATTR_PREFIX}document")
        body = ET.SubElement(document, f"{WORD_ATTR_PREFIX}body")
        body_paragraph = ET.SubElement(body, f"{WORD_ATTR_PREFIX}p")
        body_bookmark = ET.SubElement(body_paragraph, f"{WORD_ATTR_PREFIX}bookmarkStart")
        body_bookmark.set(f"{WORD_ATTR_PREFIX}id", "0")
        body_bookmark.set(f"{WORD_ATTR_PREFIX}name", "_fig-demo")
        body_bookmark_end = ET.SubElement(body_paragraph, f"{WORD_ATTR_PREFIX}bookmarkEnd")
        body_bookmark_end.set(f"{WORD_ATTR_PREFIX}id", "0")
        heading = ET.SubElement(body, f"{WORD_ATTR_PREFIX}p")
        heading_run = ET.SubElement(heading, f"{WORD_ATTR_PREFIX}r")
        heading_text = ET.SubElement(heading_run, f"{WORD_ATTR_PREFIX}t")
        heading_text.text = "参考文献"
        paragraph = ET.SubElement(body, f"{WORD_ATTR_PREFIX}p")
        bookmark_start = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}bookmarkStart")
        bookmark_start.set(f"{WORD_ATTR_PREFIX}id", "1")
        bookmark_start.set(f"{WORD_ATTR_PREFIX}name", "Walsh_2011")
        run = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}r")
        text = ET.SubElement(run, f"{WORD_ATTR_PREFIX}t")
        text.text = "Reference text"
        bookmark_end = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}bookmarkEnd")
        bookmark_end.set(f"{WORD_ATTR_PREFIX}id", "1")
        citation = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}hyperlink")
        citation.set(f"{WORD_ATTR_PREFIX}anchor", "Walsh_2011")

        changed = normalize_internal_anchor_bookmarks(document, {"Walsh_2011"})

        self.assertTrue(changed)
        self.assertEqual(len(document.findall('.//w:bookmarkStart', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})), 1)
        self.assertEqual(bookmark_start.get(f"{WORD_ATTR_PREFIX}name"), "_Walsh_2011")
        self.assertEqual(citation.get(f"{WORD_ATTR_PREFIX}anchor"), "_Walsh_2011")
        children = list(paragraph)
        self.assertEqual(children.index(bookmark_end), children.index(bookmark_start) + 1)

    def test_default_mode_moves_body_level_reference_bookmarks_into_paragraph(self) -> None:
        document = ET.Element(f"{WORD_ATTR_PREFIX}document")
        body = ET.SubElement(document, f"{WORD_ATTR_PREFIX}body")
        heading = ET.SubElement(body, f"{WORD_ATTR_PREFIX}p")
        heading_run = ET.SubElement(heading, f"{WORD_ATTR_PREFIX}r")
        heading_text = ET.SubElement(heading_run, f"{WORD_ATTR_PREFIX}t")
        heading_text.text = "参考文献"

        bookmark_start = ET.SubElement(body, f"{WORD_ATTR_PREFIX}bookmarkStart")
        bookmark_start.set(f"{WORD_ATTR_PREFIX}id", "1")
        bookmark_start.set(f"{WORD_ATTR_PREFIX}name", "Walsh_2011")

        paragraph = ET.SubElement(body, f"{WORD_ATTR_PREFIX}p")
        run = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}r")
        text = ET.SubElement(run, f"{WORD_ATTR_PREFIX}t")
        text.text = "Reference text"

        bookmark_end = ET.SubElement(body, f"{WORD_ATTR_PREFIX}bookmarkEnd")
        bookmark_end.set(f"{WORD_ATTR_PREFIX}id", "1")

        citation = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}hyperlink")
        citation.set(f"{WORD_ATTR_PREFIX}anchor", "Walsh_2011")

        changed = normalize_internal_anchor_bookmarks(document, {"Walsh_2011"})

        self.assertTrue(changed)
        self.assertEqual(
            len([child for child in list(body) if child.tag in {f"{WORD_ATTR_PREFIX}bookmarkStart", f"{WORD_ATTR_PREFIX}bookmarkEnd"}]),
            0,
        )
        self.assertEqual(bookmark_start.get(f"{WORD_ATTR_PREFIX}name"), "_Walsh_2011")
        self.assertEqual(citation.get(f"{WORD_ATTR_PREFIX}anchor"), "_Walsh_2011")
        children = list(paragraph)
        self.assertEqual(children.index(bookmark_end), children.index(bookmark_start) + 1)

    def test_default_internal_links_are_sky_blue_without_underline(self) -> None:
        document = ET.Element(f"{WORD_ATTR_PREFIX}document")
        body = ET.SubElement(document, f"{WORD_ATTR_PREFIX}body")
        paragraph = ET.SubElement(body, f"{WORD_ATTR_PREFIX}p")
        hyperlink = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}hyperlink")
        hyperlink.set(f"{WORD_ATTR_PREFIX}anchor", "_Walsh_2011")
        run = ET.SubElement(hyperlink, f"{WORD_ATTR_PREFIX}r")
        run_properties = ET.SubElement(run, f"{WORD_ATTR_PREFIX}rPr")
        run_style = ET.SubElement(run_properties, f"{WORD_ATTR_PREFIX}rStyle")
        run_style.set(f"{WORD_ATTR_PREFIX}val", "Hyperlink")
        text = ET.SubElement(run, f"{WORD_ATTR_PREFIX}t")
        text.text = "Walsh 2011"

        changed = style_default_internal_hyperlinks(document, {"Walsh_2011"})

        self.assertTrue(changed)
        self.assertEqual(len(run_properties.findall(f"{WORD_ATTR_PREFIX}rStyle")), 0)
        color = run_properties.find(f"{WORD_ATTR_PREFIX}color")
        underline = run_properties.find(f"{WORD_ATTR_PREFIX}u")
        self.assertIsNotNone(color)
        self.assertIsNotNone(underline)
        assert color is not None and underline is not None
        self.assertEqual(color.get(f"{WORD_ATTR_PREFIX}val"), "00B0F0")
        self.assertEqual(underline.get(f"{WORD_ATTR_PREFIX}val"), "none")

    def test_default_internal_links_absorb_surrounding_parentheses(self) -> None:
        document = ET.Element(f"{WORD_ATTR_PREFIX}document")
        body = ET.SubElement(document, f"{WORD_ATTR_PREFIX}body")
        paragraph = ET.SubElement(body, f"{WORD_ATTR_PREFIX}p")

        opening_run = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}r")
        opening_text = ET.SubElement(opening_run, f"{WORD_ATTR_PREFIX}t")
        opening_text.text = " ("

        hyperlink = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}hyperlink")
        hyperlink.set(f"{WORD_ATTR_PREFIX}anchor", "_Walsh_2011")
        run = ET.SubElement(hyperlink, f"{WORD_ATTR_PREFIX}r")
        text = ET.SubElement(run, f"{WORD_ATTR_PREFIX}t")
        text.text = "Walsh 2011"

        closing_run = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}r")
        closing_text = ET.SubElement(closing_run, f"{WORD_ATTR_PREFIX}t")
        closing_text.text = ")。"

        changed = style_default_internal_hyperlinks(document, {"Walsh_2011"})

        self.assertTrue(changed)
        hyperlink_text = "".join(node.text or "" for node in hyperlink.findall(".//w:t", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}))
        self.assertEqual(hyperlink_text, "(Walsh 2011)")
        remaining_text = "".join(node.text or "" for node in paragraph.findall("./w:r/w:t", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}))
        self.assertEqual(remaining_text, " 。")

    def test_relationship_internal_hyperlinks_are_flattened(self) -> None:
        document = ET.Element(f"{WORD_ATTR_PREFIX}document")
        body = ET.SubElement(document, f"{WORD_ATTR_PREFIX}body")
        paragraph = ET.SubElement(body, f"{WORD_ATTR_PREFIX}p")

        hyperlink = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}hyperlink")
        hyperlink.set(f"{REL_ATTR_PREFIX}id", "rId1")
        run = ET.SubElement(hyperlink, f"{WORD_ATTR_PREFIX}r")
        text = ET.SubElement(run, f"{WORD_ATTR_PREFIX}t")
        text.text = "Smith 2020"

        changed = strip_internal_hyperlink_styles(document, {"rId1": "#ref-1"})

        self.assertTrue(changed)
        self.assertEqual(len(document.findall('.//w:hyperlink', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})), 0)
        self.assertEqual(len(document.findall('.//w:t', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})), 1)

    def test_external_hyperlinks_are_preserved(self) -> None:
        document = ET.Element(f"{WORD_ATTR_PREFIX}document")
        body = ET.SubElement(document, f"{WORD_ATTR_PREFIX}body")
        paragraph = ET.SubElement(body, f"{WORD_ATTR_PREFIX}p")

        hyperlink = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}hyperlink")
        hyperlink.set(f"{REL_ATTR_PREFIX}id", "rId1")
        run = ET.SubElement(hyperlink, f"{WORD_ATTR_PREFIX}r")
        text = ET.SubElement(run, f"{WORD_ATTR_PREFIX}t")
        text.text = "ORCID"

        changed = strip_internal_hyperlink_styles(document, {"rId1": "https://orcid.org/0000-0000"})

        self.assertFalse(changed)
        self.assertEqual(len(document.findall('.//w:hyperlink', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})), 1)

    def test_external_fragment_hyperlinks_are_preserved(self) -> None:
        document = ET.Element(f"{WORD_ATTR_PREFIX}document")
        body = ET.SubElement(document, f"{WORD_ATTR_PREFIX}body")
        paragraph = ET.SubElement(body, f"{WORD_ATTR_PREFIX}p")

        hyperlink = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}hyperlink")
        hyperlink.set(f"{REL_ATTR_PREFIX}id", "rId1")
        run = ET.SubElement(hyperlink, f"{WORD_ATTR_PREFIX}r")
        text = ET.SubElement(run, f"{WORD_ATTR_PREFIX}t")
        text.text = "Website"

        changed = strip_internal_hyperlink_styles(document, {"rId1": "https://example.com/page#section"})

        self.assertFalse(changed)
        self.assertEqual(len(document.findall('.//w:hyperlink', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})), 1)

    def test_ref_anchor_index_maps_without_bookmarks(self) -> None:
        paragraph = ET.Element(f"{WORD_ATTR_PREFIX}p")
        hyperlink = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}hyperlink")
        hyperlink.set(f"{WORD_ATTR_PREFIX}anchor", "_Ref1")
        run = ET.SubElement(hyperlink, f"{WORD_ATTR_PREFIX}r")
        text = ET.SubElement(run, f"{WORD_ATTR_PREFIX}t")
        text.text = "Lang et al. 2025"

        target = CitationTarget(
            source_key="bib-1",
            formatted_reference="Lang et al. 2025",
            zotero_item_key="ABCD1234",
            item_data={"id": "ABCD1234", "title": "Demo"},
            uri="http://zotero.org/users/local/items/ABCD1234",
            anchor_id="bib-1",
        )
        context = ZoteroDocxContext(
            bibliography_entries=[target],
            unmatched_notices=[],
            by_anchor={},
            by_normalized_url={},
            by_normalized_doi={},
        )

        resolved = resolve_citation_hyperlink_target(hyperlink, {}, context)

        self.assertIsNotNone(resolved)
        assert resolved is not None
        self.assertEqual(resolved.zotero_item_key, "ABCD1234")
        self.assertEqual(context.by_anchor.get("_Ref1"), resolved)

    def test_display_text_resolves_to_bibliography_entry_before_fallback(self) -> None:
        paragraph = ET.Element(f"{WORD_ATTR_PREFIX}p")
        hyperlink = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}hyperlink")
        run = ET.SubElement(hyperlink, f"{WORD_ATTR_PREFIX}r")
        text = ET.SubElement(run, f"{WORD_ATTR_PREFIX}t")
        text.text = "Giannakos et al. 2025"

        target = CitationTarget(
            source_key="https://doi.org/10.1080/0144929X.2024.2394886",
            formatted_reference=(
                "Giannakos, Michail, et al. (2025) "
                "‘The Promise and Challenges of Generative AI in Education’"
            ),
            zotero_item_key="REALITEM1",
            item_data={"id": "REALITEM1", "title": "The Promise and Challenges of Generative AI in Education"},
            uri="http://zotero.org/users/local/items/REALITEM1",
            anchor_id="bib-giannakos",
        )
        context = ZoteroDocxContext(
            bibliography_entries=[target],
            unmatched_notices=[],
            by_anchor={},
            by_normalized_url={},
            by_normalized_doi={},
        )

        resolved = resolve_citation_hyperlink_target(hyperlink, {}, context)

        self.assertIs(resolved, target)
        assert resolved is not None
        self.assertEqual(resolved.item_data["title"], "The Promise and Challenges of Generative AI in Education")

    def test_display_text_uses_initial_to_disambiguate_bibliography_entry(self) -> None:
        paragraph = ET.Element(f"{WORD_ATTR_PREFIX}p")
        hyperlink = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}hyperlink")
        run = ET.SubElement(hyperlink, f"{WORD_ATTR_PREFIX}r")
        text = ET.SubElement(run, f"{WORD_ATTR_PREFIX}t")
        text.text = "H. Li et al. 2025"

        h_li = CitationTarget(
            source_key="doi-h-li",
            formatted_reference="Li, H., et al. (2025) ‘Target Article’",
            zotero_item_key="HLI2025",
            item_data={"id": "HLI2025", "title": "Target Article"},
            uri="http://zotero.org/users/local/items/HLI2025",
            anchor_id="bib-h-li",
        )
        y_li = CitationTarget(
            source_key="doi-y-li",
            formatted_reference="Li, Y., et al. (2025) ‘Other Article’",
            zotero_item_key="YLI2025",
            item_data={"id": "YLI2025", "title": "Other Article"},
            uri="http://zotero.org/users/local/items/YLI2025",
            anchor_id="bib-y-li",
        )
        context = ZoteroDocxContext(
            bibliography_entries=[y_li, h_li],
            unmatched_notices=[],
            by_anchor={},
            by_normalized_url={},
            by_normalized_doi={},
        )

        resolved = resolve_citation_hyperlink_target(hyperlink, {}, context)

        self.assertIs(resolved, h_li)

    def test_zotero_item_uri_uses_user_id(self) -> None:
        self.assertEqual(
            build_zotero_item_uri("ABCD1234", "14586934"),
            "http://zotero.org/users/14586934/items/ABCD1234",
        )

    def test_zotero_field_uses_numeric_item_id_and_plain_citation(self) -> None:
        paragraph = ET.Element(f"{WORD_ATTR_PREFIX}p")
        hyperlink = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}hyperlink")
        target = CitationTarget(
            source_key="doi-demo",
            formatted_reference="Walsh, S. (2011) Demo",
            zotero_item_key="A5IY9AKJ",
            item_data={"id": 1587, "type": "book", "title": "Demo"},
            uri="http://zotero.org/users/14586934/items/A5IY9AKJ",
            anchor_id="bib-demo",
        )

        field_runs = build_zotero_citation_field_elements("(Walsh 2011)", [target], hyperlink)
        instr_text = "".join(
            (run.find(f"{WORD_ATTR_PREFIX}instrText").text or "")
            for run in field_runs
            if run.find(f"{WORD_ATTR_PREFIX}instrText") is not None
        )
        instruction_runs = [run for run in field_runs if run.find(f"{WORD_ATTR_PREFIX}instrText") is not None]

        self.assertIn('"plainCitation":"(Walsh 2011)"', instr_text)
        self.assertIn('"noteIndex":0', instr_text)
        self.assertIn('"unsorted":true', instr_text)
        self.assertNotIn('"dontUpdate"', instr_text)
        self.assertRegex(instr_text, r'"citationID":"[A-Za-z0-9]{8}"')
        self.assertNotIn('"citationID":"cite-', instr_text)
        self.assertIn('"id":1587', instr_text)
        self.assertIn('"uris":["http://zotero.org/users/14586934/items/A5IY9AKJ"]', instr_text)
        self.assertEqual(len(instruction_runs), 1)
        for run in field_runs:
            fonts = run.find(f"{WORD_ATTR_PREFIX}rPr/{WORD_ATTR_PREFIX}rFonts")
            self.assertIsNone(fonts)
            color = run.find(f"{WORD_ATTR_PREFIX}rPr/{WORD_ATTR_PREFIX}color")
            self.assertIsNotNone(color)
            assert color is not None
            self.assertEqual(color.get(f"{WORD_ATTR_PREFIX}val"), "003399")

    def test_strip_zotero_field_run_fonts_keeps_body_fonts(self) -> None:
        root = ET.fromstring(
            """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:body>
                    <w:p>
                        <w:r>
                            <w:rPr><w:color w:val="003399"/></w:rPr>
                            <w:fldChar w:fldCharType="begin"/>
                        </w:r>
                        <w:r>
                            <w:rPr><w:color w:val="003399"/></w:rPr>
                            <w:instrText xml:space="preserve"> ADDIN ZOTERO_ITEM CSL_CITATION {"citationID":"cite-1"} </w:instrText>
                        </w:r>
                        <w:r>
                            <w:rPr><w:color w:val="003399"/></w:rPr>
                            <w:fldChar w:fldCharType="separate"/>
                        </w:r>
                        <w:r>
                            <w:rPr><w:color w:val="003399"/></w:rPr>
                            <w:t>(Demo 2025)</w:t>
                        </w:r>
                        <w:r>
                            <w:rPr><w:color w:val="003399"/></w:rPr>
                            <w:fldChar w:fldCharType="end"/>
                        </w:r>
                        <w:r>
                            <w:rPr><w:b/></w:rPr>
                            <w:t>Body text</w:t>
                        </w:r>
                    </w:p>
                </w:body>
            </w:document>
            """
        )

        self.assertTrue(normalize_tree_run_fonts(root))
        self.assertTrue(strip_zotero_field_run_fonts(root))

        namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        runs = root.findall(".//w:p/w:r", namespace)
        for run in runs[:5]:
            fonts = run.find(f"{WORD_ATTR_PREFIX}rPr/{WORD_ATTR_PREFIX}rFonts")
            self.assertIsNone(fonts)

        body_fonts = runs[5].find(f"{WORD_ATTR_PREFIX}rPr/{WORD_ATTR_PREFIX}rFonts")
        self.assertIsNotNone(body_fonts)

    def test_parse_bibliography_entries_supports_refs_bib_with_refs_display(self) -> None:
        with TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            (temp_path / "refs_display.json").write_text(
                '{"appleton2006": "Appleton et al.\u00a02006"}',
                encoding="utf-8",
            )
            (temp_path / "refs.bib").write_text(
                """@article{appleton2006,
  title = {Measuring Cognitive and Psychological Engagement},
  author = {Appleton, James J. and Christenson, Sandra L.},
  year = {2006},
  doi = {10.1016/j.jsp.2006.04.002},
  journal = {Journal of School Psychology}
}
""",
                encoding="utf-8",
            )

            entries = parse_bibliography_entries(temp_path / "refs.bib")

        self.assertEqual(len(entries), 1)
        self.assertEqual(entries[0].source_key, "10.1016/j.jsp.2006.04.002")
        self.assertEqual(entries[0].formatted_reference, "Appleton et al.\u00a02006")
        self.assertEqual(entries[0].parsed_title, "Measuring Cognitive and Psychological Engagement")

    def test_zotero_bibliography_field_preserves_ppr_first(self) -> None:
        document = ET.Element(f"{WORD_ATTR_PREFIX}document")
        body = ET.SubElement(document, f"{WORD_ATTR_PREFIX}body")

        heading = ET.SubElement(body, f"{WORD_ATTR_PREFIX}p")
        heading_run = ET.SubElement(heading, f"{WORD_ATTR_PREFIX}r")
        heading_text = ET.SubElement(heading_run, f"{WORD_ATTR_PREFIX}t")
        heading_text.text = "参考文献"

        bibliography_paragraph = ET.SubElement(body, f"{WORD_ATTR_PREFIX}p")
        paragraph_properties = ET.SubElement(bibliography_paragraph, f"{WORD_ATTR_PREFIX}pPr")
        paragraph_style = ET.SubElement(paragraph_properties, f"{WORD_ATTR_PREFIX}pStyle")
        paragraph_style.set(f"{WORD_ATTR_PREFIX}val", "41")
        bibliography_run = ET.SubElement(bibliography_paragraph, f"{WORD_ATTR_PREFIX}r")
        bibliography_text = ET.SubElement(bibliography_run, f"{WORD_ATTR_PREFIX}t")
        bibliography_text.text = "Walsh, S. (2011) Demo"

        changed = apply_bibliography_hints(
            document,
            TemplateDocxHints(
                caption_style_id=None,
                table_style_id=None,
                table_paragraph_style_id=None,
                normal_style_id=None,
                title_style_id="Title",
                heading_1_style_id="Heading1",
                heading_2_style_id="Heading2",
                heading_3_style_id="Heading3",
                bibliography_style_id="41",
                zotero_item_uri_prefix=None,
            ),
            make_context(),
            enable_zotero=True,
        )

        self.assertTrue(changed)
        children = list(bibliography_paragraph)
        self.assertEqual(children[0].tag, f"{WORD_ATTR_PREFIX}pPr")
        instr = bibliography_paragraph.find(".//w:instrText", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
        self.assertIsNotNone(instr)
        assert instr is not None
        self.assertIn("ZOTERO_BIBL", instr.text or "")
        bibliography_end_paragraph = list(body)[-1]
        self.assertIsNot(bibliography_end_paragraph, bibliography_paragraph)
        end_children = list(bibliography_end_paragraph)
        self.assertEqual(end_children[0].tag, f"{WORD_ATTR_PREFIX}pPr")
        self.assertIsNone(bibliography_paragraph.find("./w:r/w:fldChar[@w:fldCharType='end']", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}))
        bibliography_end = bibliography_end_paragraph.find("./w:r/w:fldChar", {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
        self.assertIsNotNone(bibliography_end)
        assert bibliography_end is not None
        self.assertEqual(bibliography_end.get(f"{WORD_ATTR_PREFIX}fldCharType"), "end")


if __name__ == "__main__":
    unittest.main()
