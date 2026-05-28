import unittest
import xml.etree.ElementTree as ET

from dotex.converter import (
    WORD_ATTR_PREFIX,
    CitationTarget,
    ZoteroDocxContext,
    resolve_citation_hyperlink_target,
    strip_all_bookmarks,
    strip_internal_hyperlink_styles,
)


def make_context() -> ZoteroDocxContext:
    return ZoteroDocxContext(
        bibliography_entries=[],
        unmatched_notices=[],
        by_anchor={},
        by_normalized_url={},
        by_normalized_doi={},
    )


class ZoteroModeTests(unittest.TestCase):
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


if __name__ == "__main__":
    unittest.main()
