import unittest
import xml.etree.ElementTree as ET

import dotex.tex_to_docx as tex_to_docx_module

from dotex.tex_to_docx import (
    DEFAULT_ZOTERO_FIELD_COLOR,
    WORD_ATTR_PREFIX,
    ConversionDiagnostics,
    TemplateDocxHints,
    apply_native_cross_reference_fields,
    build_caption_placeholder,
    format_numbered_reference,
    infer_reference_prefixes,
    make_cross_reference_anchor,
)


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def make_template_hints() -> TemplateDocxHints:
    return TemplateDocxHints(
        caption_style_id="Caption",
        table_style_id=None,
        table_paragraph_style_id=None,
        normal_style_id=None,
        title_style_id="Title",
        heading_1_style_id="Heading1",
        heading_2_style_id="Heading2",
        heading_3_style_id="Heading3",
        bibliography_style_id=None,
        zotero_item_uri_prefix=None,
    )


class CrossReferenceTests(unittest.TestCase):
    def test_caption_placeholder_becomes_native_field_and_ref_field(self) -> None:
        document = ET.Element(f"{WORD_ATTR_PREFIX}document")
        body = ET.SubElement(document, f"{WORD_ATTR_PREFIX}body")

        caption_paragraph = ET.SubElement(body, f"{WORD_ATTR_PREFIX}p")
        caption_run = ET.SubElement(caption_paragraph, f"{WORD_ATTR_PREFIX}r")
        caption_text = ET.SubElement(caption_run, f"{WORD_ATTR_PREFIX}t")
        caption_text.text = build_caption_placeholder("figure", "fig:demo", "示例图题注", "1")

        body_paragraph = ET.SubElement(body, f"{WORD_ATTR_PREFIX}p")
        hyperlink = ET.SubElement(body_paragraph, f"{WORD_ATTR_PREFIX}hyperlink")
        hyperlink.set(f"{WORD_ATTR_PREFIX}anchor", make_cross_reference_anchor("fig:demo"))
        hyperlink_run = ET.SubElement(hyperlink, f"{WORD_ATTR_PREFIX}r")
        hyperlink_text = ET.SubElement(hyperlink_run, f"{WORD_ATTR_PREFIX}t")
        hyperlink_text.text = "图1"

        diagnostics = ConversionDiagnostics()
        changed = apply_native_cross_reference_fields(document, make_template_hints(), {}, diagnostics)

        self.assertTrue(changed)
        bookmark_start = caption_paragraph.find("./w:bookmarkStart", NS)
        self.assertIsNotNone(bookmark_start)
        assert bookmark_start is not None
        self.assertRegex(bookmark_start.get(f"{WORD_ATTR_PREFIX}name") or "", r"^_Ref\d+$")
        caption_instr = caption_paragraph.find(".//w:instrText", NS)
        self.assertIsNotNone(caption_instr)
        assert caption_instr is not None
        self.assertIn("SEQ 图", caption_instr.text or "")
        self.assertIsNone(body_paragraph.find("./w:hyperlink", NS))
        ref_instr = body_paragraph.find(".//w:instrText", NS)
        self.assertIsNotNone(ref_instr)
        assert ref_instr is not None
        self.assertRegex(ref_instr.text or "", r" REF _Ref\d+ \\h ")
        for run in body_paragraph.findall("./w:r", NS):
            color = run.find("./w:rPr/w:color", NS)
            self.assertIsNotNone(color)
            assert color is not None
            self.assertEqual(color.get(f"{WORD_ATTR_PREFIX}val"), DEFAULT_ZOTERO_FIELD_COLOR)
        self.assertEqual(diagnostics.warnings, [])
        self.assertEqual(len(diagnostics.cross_reference_targets), 1)
        self.assertTrue(diagnostics.cross_reference_targets[0].referenced)

    def test_missing_cross_reference_falls_back_to_blue_plain_text(self) -> None:
        document = ET.Element(f"{WORD_ATTR_PREFIX}document")
        body = ET.SubElement(document, f"{WORD_ATTR_PREFIX}body")

        body_paragraph = ET.SubElement(body, f"{WORD_ATTR_PREFIX}p")
        hyperlink = ET.SubElement(body_paragraph, f"{WORD_ATTR_PREFIX}hyperlink")
        hyperlink.set(f"{WORD_ATTR_PREFIX}anchor", make_cross_reference_anchor("fig:missing"))
        hyperlink_run = ET.SubElement(hyperlink, f"{WORD_ATTR_PREFIX}r")
        hyperlink_text = ET.SubElement(hyperlink_run, f"{WORD_ATTR_PREFIX}t")
        hyperlink_text.text = "图9"

        diagnostics = ConversionDiagnostics()
        changed = apply_native_cross_reference_fields(document, make_template_hints(), {}, diagnostics)

        self.assertTrue(changed)
        self.assertIsNone(body_paragraph.find("./w:hyperlink", NS))
        self.assertEqual("".join(node.text or "" for node in body_paragraph.findall(".//w:t", NS)), "图9")
        color = body_paragraph.find("./w:r/w:rPr/w:color", NS)
        self.assertIsNotNone(color)
        assert color is not None
        self.assertEqual(color.get(f"{WORD_ATTR_PREFIX}val"), DEFAULT_ZOTERO_FIELD_COLOR)
        self.assertTrue(any("交叉引用 图9 未能解析到对应题注" in warning for warning in diagnostics.warnings))

    def test_unreferenced_caption_emits_warning(self) -> None:
        document = ET.Element(f"{WORD_ATTR_PREFIX}document")
        body = ET.SubElement(document, f"{WORD_ATTR_PREFIX}body")
        caption_paragraph = ET.SubElement(body, f"{WORD_ATTR_PREFIX}p")
        caption_run = ET.SubElement(caption_paragraph, f"{WORD_ATTR_PREFIX}r")
        caption_text = ET.SubElement(caption_run, f"{WORD_ATTR_PREFIX}t")
        caption_text.text = build_caption_placeholder("table", "tab:demo", "示例表题注", "2")

        diagnostics = ConversionDiagnostics()
        changed = apply_native_cross_reference_fields(document, make_template_hints(), {}, diagnostics)

        self.assertTrue(changed)
        self.assertTrue(
            any("表2 示例表题注 在文内不含交叉引用" in warning for warning in diagnostics.warnings)
        )

    def test_infer_reference_prefixes_supports_english_refs(self) -> None:
        prefixes = infer_reference_prefixes(
            r"See Table~\ref{tab:demo}, Figure \ref{fig:demo}, and \hyperref[_Ref1]{Table 7}."
        )

        self.assertEqual(prefixes["tab:demo"], "Table")
        self.assertEqual(prefixes["fig:demo"], "Figure")
        self.assertEqual(prefixes["_Ref1"], "Table")
        self.assertEqual(format_numbered_reference("Table", "7"), "Table 7")
        self.assertEqual(format_numbered_reference("表", "7"), "表7")

    def test_caption_placeholder_uses_english_prefix_when_document_is_english(self) -> None:
        old_document_uses_cjk = tex_to_docx_module.CURRENT_DOCUMENT_USES_CJK
        old_reference_prefixes = dict(tex_to_docx_module.CURRENT_REFERENCE_PREFIXES)
        try:
            tex_to_docx_module.CURRENT_DOCUMENT_USES_CJK = False
            tex_to_docx_module.CURRENT_REFERENCE_PREFIXES = {"tab:demo": "Table"}

            document = ET.Element(f"{WORD_ATTR_PREFIX}document")
            body = ET.SubElement(document, f"{WORD_ATTR_PREFIX}body")

            caption_paragraph = ET.SubElement(body, f"{WORD_ATTR_PREFIX}p")
            caption_run = ET.SubElement(caption_paragraph, f"{WORD_ATTR_PREFIX}r")
            caption_text = ET.SubElement(caption_run, f"{WORD_ATTR_PREFIX}t")
            caption_text.text = build_caption_placeholder("table", "tab:demo", "Example table caption", "2")

            diagnostics = ConversionDiagnostics()
            changed = apply_native_cross_reference_fields(document, make_template_hints(), {}, diagnostics)

            self.assertTrue(changed)
            self.assertIn(
                "Table 2 Example table caption",
                "".join(node.text or "" for node in caption_paragraph.findall(".//w:t", NS)),
            )
            caption_instr = caption_paragraph.find(".//w:instrText", NS)
            self.assertIsNotNone(caption_instr)
            assert caption_instr is not None
            self.assertIn("SEQ Table", caption_instr.text or "")
        finally:
            tex_to_docx_module.CURRENT_DOCUMENT_USES_CJK = old_document_uses_cjk
            tex_to_docx_module.CURRENT_REFERENCE_PREFIXES = old_reference_prefixes


if __name__ == "__main__":
    unittest.main()