import unittest
import xml.etree.ElementTree as ET

from dotex.tex_to_docx import (
    WORD_ATTR_PREFIX,
    TemplateDocxHints,
    WESTERN_FONT_FAMILY,
    normalize_document_style_usage,
    sanitize_styles_part,
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


class StyleCleanupTests(unittest.TestCase):
    def test_sanitize_styles_prunes_unrelated_styles_and_sets_tnr(self) -> None:
        styles_xml = """<?xml version='1.0' encoding='utf-8'?>
<w:styles xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:asciiTheme='minorHAnsi' w:hAnsiTheme='minorHAnsi'/></w:rPr></w:rPrDefault>
  </w:docDefaults>
  <w:latentStyles w:count='2'><w:lsdException w:name='Heading 4'/></w:latentStyles>
  <w:style w:type='paragraph' w:styleId='Normal' w:default='1'><w:name w:val='Normal'/></w:style>
  <w:style w:type='paragraph' w:styleId='Title'><w:name w:val='Title'/></w:style>
  <w:style w:type='paragraph' w:styleId='Heading1'><w:name w:val='heading 1'/><w:basedOn w:val='Normal'/></w:style>
  <w:style w:type='numbering' w:styleId='NoList'><w:name w:val='No List'/></w:style>
    <w:style w:type='table' w:styleId='TableBase' w:default='1'><w:name w:val='Normal Table'/></w:style>
    <w:style w:type='table' w:styleId='DemoTable'><w:name w:val='Demo Table'/><w:basedOn w:val='TableBase'/></w:style>
  <w:style w:type='paragraph' w:styleId='WeirdStyle'><w:name w:val='Weird Style'/></w:style>
  <w:style w:type='character' w:styleId='DefaultParagraphFont' w:default='1'><w:name w:val='Default Paragraph Font'/></w:style>
</w:styles>
""".encode("utf-8")

        sanitized_xml, changed = sanitize_styles_part(styles_xml, make_hints())
        root = ET.fromstring(sanitized_xml)
        style_ids = [style.get(f"{WORD_ATTR_PREFIX}styleId") for style in root.findall("w:style", NS)]

        self.assertTrue(changed)
        self.assertIn("Normal", style_ids)
        self.assertIn("Title", style_ids)
        self.assertIn("Heading1", style_ids)
        self.assertIn("NoList", style_ids)
        self.assertIn("TableBase", style_ids)
        self.assertIn("DemoTable", style_ids)
        self.assertNotIn("WeirdStyle", style_ids)
        fonts = root.find("./w:docDefaults/w:rPrDefault/w:rPr/w:rFonts", NS)
        self.assertIsNotNone(fonts)
        assert fonts is not None
        self.assertEqual(fonts.get(f"{WORD_ATTR_PREFIX}ascii"), WESTERN_FONT_FAMILY)
        self.assertEqual(fonts.get(f"{WORD_ATTR_PREFIX}hAnsi"), WESTERN_FONT_FAMILY)
        self.assertEqual(fonts.get(f"{WORD_ATTR_PREFIX}cs"), WESTERN_FONT_FAMILY)
        latent_styles = root.find("w:latentStyles", NS)
        self.assertIsNotNone(latent_styles)
        assert latent_styles is not None
        self.assertEqual(latent_styles.get(f"{WORD_ATTR_PREFIX}count"), "2")
        self.assertEqual(len(list(latent_styles)), 1)

    def test_normalize_document_style_usage_collapses_weird_styles(self) -> None:
        document = ET.Element(f"{WORD_ATTR_PREFIX}document")
        body = ET.SubElement(document, f"{WORD_ATTR_PREFIX}body")
        paragraph = ET.SubElement(body, f"{WORD_ATTR_PREFIX}p")
        paragraph_properties = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}pPr")
        paragraph_style = ET.SubElement(paragraph_properties, f"{WORD_ATTR_PREFIX}pStyle")
        paragraph_style.set(f"{WORD_ATTR_PREFIX}val", "WeirdStyle")
        run = ET.SubElement(paragraph, f"{WORD_ATTR_PREFIX}r")
        run_properties = ET.SubElement(run, f"{WORD_ATTR_PREFIX}rPr")
        run_style = ET.SubElement(run_properties, f"{WORD_ATTR_PREFIX}rStyle")
        run_style.set(f"{WORD_ATTR_PREFIX}val", "WeirdRun")
        text = ET.SubElement(run, f"{WORD_ATTR_PREFIX}t")
        text.text = "Demo"

        changed = normalize_document_style_usage(document, make_hints())

        self.assertTrue(changed)
        updated_style = paragraph.find("./w:pPr/w:pStyle", NS)
        self.assertIsNotNone(updated_style)
        assert updated_style is not None
        self.assertEqual(updated_style.get(f"{WORD_ATTR_PREFIX}val"), "Normal")
        self.assertIsNone(paragraph.find(".//w:rStyle", NS))


if __name__ == "__main__":
    unittest.main()