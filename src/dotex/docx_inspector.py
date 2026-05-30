from __future__ import annotations

import json
import zipfile
from collections import Counter
from dataclasses import asdict, dataclass
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}
PACKAGE_RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships"

W = NS["w"]


def qn(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def get_attr(el: ET.Element | None, ns: str, tag: str) -> str | None:
    if el is None:
        return None
    return el.get(qn(ns, tag))


def read_xml(zf: zipfile.ZipFile, name: str) -> ET.Element | None:
    if name not in zf.namelist():
        return None
    return ET.fromstring(zf.read(name))


def clean_text(value: str | None) -> str:
    if not value:
        return ""
    return " ".join(value.replace("\u00a0", " ").split())


@dataclass
class StyleRecord:
    style_id: str
    name: str
    style_type: str
    is_default: bool
    custom_style: bool
    based_on: str | None
    next_style: str | None
    linked_style: str | None
    ui_priority: str | None
    quick_format: bool


@dataclass
class TemplateManifest:
    source: dict
    styles: dict
    document: dict
    theme: dict
    doc_defaults: dict
    zotero: dict

    def to_json(self) -> str:
        return json.dumps(asdict(self), ensure_ascii=False, indent=2, sort_keys=True)


class WordTemplateInspector:
    def __init__(self, docx_path: Path):
        self.docx_path = docx_path
        self.zf = zipfile.ZipFile(docx_path)
        self.document = read_xml(self.zf, "word/document.xml")
        self.styles_root = read_xml(self.zf, "word/styles.xml")
        self.settings_root = read_xml(self.zf, "word/settings.xml")
        self.theme_root = read_xml(self.zf, "word/theme/theme1.xml")
        self.numbering_root = read_xml(self.zf, "word/numbering.xml")

    def inspect(self) -> TemplateManifest:
        styles = self._collect_styles()
        paragraph_usage = self._collect_paragraph_style_usage(styles)
        run_style_usage = self._collect_run_style_usage(styles)
        table_style_usage = self._collect_table_style_usage(styles)
        section_properties = self._collect_sections()
        captions = self._collect_caption_samples(styles)

        return TemplateManifest(
            source={
                "docx_path": str(self.docx_path),
                "package_part_count": len(self.zf.namelist()),
                "has_numbering": self.numbering_root is not None,
            },
            styles={
                "counts_by_type": dict(Counter(record.style_type for record in styles.values())),
                "defaults": self._collect_default_styles(),
                "table_styles": [
                    asdict(record)
                    for record in sorted(styles.values(), key=lambda item: item.name.lower())
                    if record.style_type == "table"
                ],
                "items": [
                    asdict(record)
                    for record in sorted(styles.values(), key=lambda item: (item.style_type, item.name.lower()))
                ],
            },
            document={
                "paragraph_style_usage": paragraph_usage,
                "run_style_usage": run_style_usage,
                "table_style_usage": table_style_usage,
                "caption_samples": captions,
                "sections": section_properties,
                "link_artifacts": self._collect_link_artifacts(),
            },
            theme=self._collect_theme_fonts(),
            doc_defaults=self._collect_doc_defaults(),
            zotero=self._collect_zotero_traces(),
        )

    def _collect_styles(self) -> dict[str, StyleRecord]:
        if self.styles_root is None:
            return {}

        records: dict[str, StyleRecord] = {}
        for style in self.styles_root.findall("w:style", NS):
            style_id = get_attr(style, W, "styleId") or ""
            name_el = style.find("w:name", NS)
            based_on_el = style.find("w:basedOn", NS)
            next_el = style.find("w:next", NS)
            linked_el = style.find("w:link", NS)
            ui_priority_el = style.find("w:uiPriority", NS)
            records[style_id] = StyleRecord(
                style_id=style_id,
                name=get_attr(name_el, W, "val") or style_id,
                style_type=get_attr(style, W, "type") or "unknown",
                is_default=(get_attr(style, W, "default") == "1"),
                custom_style=(get_attr(style, W, "customStyle") == "1"),
                based_on=get_attr(based_on_el, W, "val"),
                next_style=get_attr(next_el, W, "val"),
                linked_style=get_attr(linked_el, W, "val"),
                ui_priority=get_attr(ui_priority_el, W, "val"),
                quick_format=(style.find("w:qFormat", NS) is not None),
            )
        return records

    def _collect_default_styles(self) -> dict[str, str]:
        defaults: dict[str, str] = {}
        if self.styles_root is None:
            return defaults
        for style in self.styles_root.findall("w:style", NS):
            if get_attr(style, W, "default") != "1":
                continue
            style_type = get_attr(style, W, "type") or "unknown"
            defaults[style_type] = get_attr(style, W, "styleId") or ""
        return defaults

    def _style_name(self, style_id: str | None, styles: dict[str, StyleRecord]) -> str | None:
        if style_id is None:
            return None
        record = styles.get(style_id)
        return record.name if record is not None else style_id

    def _collect_paragraph_style_usage(self, styles: dict[str, StyleRecord]) -> list[dict]:
        if self.document is None:
            return []
        counter: Counter[str] = Counter()
        for paragraph in self.document.findall(".//w:p", NS):
            pstyle = paragraph.find("./w:pPr/w:pStyle", NS)
            style_id = get_attr(pstyle, W, "val") or "Normal"
            counter[style_id] += 1
        return [
            {
                "style_id": style_id,
                "style_name": self._style_name(style_id, styles),
                "count": count,
            }
            for style_id, count in counter.most_common()
        ]

    def _collect_run_style_usage(self, styles: dict[str, StyleRecord]) -> list[dict]:
        if self.document is None:
            return []
        counter: Counter[str] = Counter()
        for run in self.document.findall(".//w:r", NS):
            rstyle = run.find("./w:rPr/w:rStyle", NS)
            style_id = get_attr(rstyle, W, "val")
            if style_id:
                counter[style_id] += 1
        return [
            {
                "style_id": style_id,
                "style_name": self._style_name(style_id, styles),
                "count": count,
            }
            for style_id, count in counter.most_common()
        ]

    def _collect_table_style_usage(self, styles: dict[str, StyleRecord]) -> list[dict]:
        if self.document is None:
            return []
        counter: Counter[str] = Counter()
        for table in self.document.findall(".//w:tbl", NS):
            tbl_style = table.find("./w:tblPr/w:tblStyle", NS)
            style_id = get_attr(tbl_style, W, "val") or "<direct-formatting-or-none>"
            counter[style_id] += 1
        return [
            {
                "style_id": style_id,
                "style_name": self._style_name(style_id, styles),
                "count": count,
            }
            for style_id, count in counter.most_common()
        ]

    def _collect_caption_samples(self, styles: dict[str, StyleRecord]) -> list[dict]:
        if self.document is None:
            return []
        samples: list[dict] = []
        for paragraph in self.document.findall(".//w:p", NS):
            pstyle = paragraph.find("./w:pPr/w:pStyle", NS)
            style_id = get_attr(pstyle, W, "val")
            style_name = self._style_name(style_id, styles)
            if (style_name or "").strip().lower() != "caption":
                continue
            text = clean_text("".join(node.text or "" for node in paragraph.findall(".//w:t", NS)))
            if text:
                samples.append(
                    {
                        "style_id": style_id,
                        "style_name": style_name,
                        "text": text,
                    }
                )
            if len(samples) >= 10:
                break
        return samples

    def _collect_sections(self) -> list[dict]:
        if self.document is None:
            return []
        sections: list[dict] = []
        for index, sect in enumerate(self.document.findall(".//w:sectPr", NS), start=1):
            page_size = sect.find("./w:pgSz", NS)
            margins = sect.find("./w:pgMar", NS)
            cols = sect.find("./w:cols", NS)
            doc_grid = sect.find("./w:docGrid", NS)
            sections.append(
                {
                    "index": index,
                    "page_size_twips": {
                        "width": get_attr(page_size, W, "w"),
                        "height": get_attr(page_size, W, "h"),
                        "orientation": get_attr(page_size, W, "orient"),
                    },
                    "page_margins_twips": {
                        "top": get_attr(margins, W, "top"),
                        "right": get_attr(margins, W, "right"),
                        "bottom": get_attr(margins, W, "bottom"),
                        "left": get_attr(margins, W, "left"),
                        "header": get_attr(margins, W, "header"),
                        "footer": get_attr(margins, W, "footer"),
                        "gutter": get_attr(margins, W, "gutter"),
                    },
                    "columns": {
                        "count": get_attr(cols, W, "num"),
                        "space": get_attr(cols, W, "space"),
                    },
                    "doc_grid": {
                        "type": get_attr(doc_grid, W, "type"),
                        "line_pitch": get_attr(doc_grid, W, "linePitch"),
                        "char_space": get_attr(doc_grid, W, "charSpace"),
                    },
                }
            )
        return sections

    def _collect_doc_defaults(self) -> dict:
        if self.styles_root is None:
            return {}
        rpr_default = self.styles_root.find("./w:docDefaults/w:rPrDefault/w:rPr", NS)
        ppr_default = self.styles_root.find("./w:docDefaults/w:pPrDefault/w:pPr", NS)
        fonts = rpr_default.find("./w:rFonts", NS) if rpr_default is not None else None
        size = rpr_default.find("./w:sz", NS) if rpr_default is not None else None
        lang = rpr_default.find("./w:lang", NS) if rpr_default is not None else None
        spacing = ppr_default.find("./w:spacing", NS) if ppr_default is not None else None
        indent = ppr_default.find("./w:ind", NS) if ppr_default is not None else None
        jc = ppr_default.find("./w:jc", NS) if ppr_default is not None else None

        return {
            "run_fonts": {
                "ascii": get_attr(fonts, W, "ascii"),
                "h_ansi": get_attr(fonts, W, "hAnsi"),
                "east_asia": get_attr(fonts, W, "eastAsia"),
                "complex_script": get_attr(fonts, W, "cs"),
                "ascii_theme": get_attr(fonts, W, "asciiTheme"),
                "h_ansi_theme": get_attr(fonts, W, "hAnsiTheme"),
                "east_asia_theme": get_attr(fonts, W, "eastAsiaTheme"),
            },
            "font_size_half_points": get_attr(size, W, "val"),
            "language": {
                "val": get_attr(lang, W, "val"),
                "east_asia": get_attr(lang, W, "eastAsia"),
                "bidi": get_attr(lang, W, "bidi"),
            },
            "paragraph_spacing_twips": {
                "before": get_attr(spacing, W, "before"),
                "after": get_attr(spacing, W, "after"),
                "line": get_attr(spacing, W, "line"),
                "line_rule": get_attr(spacing, W, "lineRule"),
            },
            "paragraph_indent_twips": {
                "left": get_attr(indent, W, "left"),
                "right": get_attr(indent, W, "right"),
                "first_line": get_attr(indent, W, "firstLine"),
                "hanging": get_attr(indent, W, "hanging"),
            },
            "paragraph_alignment": get_attr(jc, W, "val"),
        }

    def _collect_theme_fonts(self) -> dict:
        if self.theme_root is None:
            return {}
        major = self.theme_root.find(".//a:themeElements/a:fontScheme/a:majorFont", NS)
        minor = self.theme_root.find(".//a:themeElements/a:fontScheme/a:minorFont", NS)
        return {
            "major": self._collect_font_scheme(major),
            "minor": self._collect_font_scheme(minor),
        }

    def _collect_font_scheme(self, root: ET.Element | None) -> dict:
        if root is None:
            return {}
        latin = root.find("./a:latin", NS)
        ea = root.find("./a:ea", NS)
        cs = root.find("./a:cs", NS)
        script_fonts = []
        for child in root:
            local_tag = child.tag.rsplit("}", 1)[-1]
            if local_tag not in {"font", "latin", "ea", "cs"}:
                continue
            if local_tag == "font":
                script_fonts.append(
                    {
                        "script": child.get("script"),
                        "typeface": child.get("typeface"),
                    }
                )
        return {
            "latin": latin.get("typeface") if latin is not None else None,
            "east_asia": ea.get("typeface") if ea is not None else None,
            "complex_script": cs.get("typeface") if cs is not None else None,
            "script_fonts": script_fonts,
        }

    def _collect_zotero_traces(self) -> dict:
        hits: list[dict] = []
        markers = ("ZOTERO_ITEM CSL_CITATION", "CSL_BIBLIOGRAPHY")
        citation_field_count = 0
        bibliography_field_count = 0
        for name in sorted(self.zf.namelist()):
            if not name.startswith("word/") or not name.endswith(".xml"):
                continue
            raw_text = self.zf.read(name).decode("utf-8", errors="ignore")
            citation_field_count += raw_text.count("ZOTERO_ITEM CSL_CITATION")
            bibliography_field_count += raw_text.count("CSL_BIBLIOGRAPHY")
            part_hits = [marker for marker in markers if marker in raw_text]
            if not part_hits:
                continue
            sample = ""
            for marker in markers:
                idx = raw_text.find(marker)
                if idx != -1:
                    sample = raw_text[idx : idx + 240]
                    break
            hits.append(
                {
                    "part": name,
                    "markers": part_hits,
                    "sample": sample,
                }
            )
        return {
            "detected": bool(hits),
            "parts_with_hits": hits,
            "field_code_hit_count": len(hits),
            "citation_field_count": citation_field_count,
            "bibliography_field_count": bibliography_field_count,
        }

    def _collect_link_artifacts(self) -> dict:
        bookmark_start_count = 0
        bookmark_end_count = 0
        non_hidden_bookmark_count = 0
        expanded_bookmark_count = 0
        hyperlink_count = 0
        internal_hyperlink_count = 0
        document_relationships = self._document_relationship_targets()

        for name in sorted(self.zf.namelist()):
            if not name.startswith("word/") or not name.endswith(".xml"):
                continue
            root = read_xml(self.zf, name)
            if root is None:
                continue
            bookmark_ends = {
                get_attr(bookmark, W, "id"): bookmark
                for bookmark in root.findall(".//w:bookmarkEnd", NS)
                if get_attr(bookmark, W, "id") is not None
            }
            bookmark_starts = root.findall(".//w:bookmarkStart", NS)
            bookmark_start_count += len(bookmark_starts)
            bookmark_end_count += len(bookmark_ends)
            for bookmark in bookmark_starts:
                name = get_attr(bookmark, W, "name") or ""
                if not name.startswith("_"):
                    non_hidden_bookmark_count += 1
                bookmark_end = bookmark_ends.get(get_attr(bookmark, W, "id") or "")
                if bookmark_end is not None and not is_collapsed_bookmark(root, bookmark, bookmark_end):
                    expanded_bookmark_count += 1
            for hyperlink in root.findall(".//w:hyperlink", NS):
                hyperlink_count += 1
                anchor = get_attr(hyperlink, W, "anchor")
                rel_id = get_attr(hyperlink, NS["r"], "id")
                relationship_target = document_relationships.get(rel_id or "")
                if anchor or is_internal_relationship_target(relationship_target):
                    internal_hyperlink_count += 1

        return {
            "bookmark_start_count": bookmark_start_count,
            "bookmark_end_count": bookmark_end_count,
            "non_hidden_bookmark_count": non_hidden_bookmark_count,
            "expanded_bookmark_count": expanded_bookmark_count,
            "hyperlink_count": hyperlink_count,
            "internal_hyperlink_count": internal_hyperlink_count,
        }

    def _document_relationship_targets(self) -> dict[str, str]:
        if "word/_rels/document.xml.rels" not in self.zf.namelist():
            return {}
        root = ET.fromstring(self.zf.read("word/_rels/document.xml.rels"))
        relationships: dict[str, str] = {}
        for relationship in root.findall(f"{{{PACKAGE_RELATIONSHIP_NAMESPACE}}}Relationship"):
            rel_id = relationship.get("Id")
            target = relationship.get("Target")
            if rel_id and target:
                relationships[rel_id] = target
        return relationships


def is_internal_relationship_target(target: str | None) -> bool:
    if not target:
        return False
    stripped = target.strip()
    if stripped.startswith("#"):
        return True
    if any(marker in stripped for marker in (":", "/", "\\")):
        return False
    return True


def is_collapsed_bookmark(root: ET.Element, start: ET.Element, end: ET.Element) -> bool:
    parent = find_direct_parent(root, start)
    if parent is None or parent is not find_direct_parent(root, end):
        return False
    children = list(parent)
    return children.index(end) == children.index(start) + 1


def find_direct_parent(root: ET.Element, target: ET.Element) -> ET.Element | None:
    for parent in root.iter():
        for child in list(parent):
            if child is target:
                return parent
    return None


def inspect_template(docx_path: Path) -> TemplateManifest:
    inspector = WordTemplateInspector(docx_path)
    return inspector.inspect()
