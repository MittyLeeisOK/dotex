from __future__ import annotations

import os
import re
import shutil
import tempfile
import xml.etree.ElementTree as ET
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import pypandoc


WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DRAWING_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main"
WORDPROCESSING_DRAWING_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
MATH_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/math"
PACKAGE_RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships"
XML_NAMESPACES = {
    "w": WORD_NAMESPACE,
    "r": RELATIONSHIP_NAMESPACE,
    "a": DRAWING_NAMESPACE,
    "wp": WORDPROCESSING_DRAWING_NAMESPACE,
    "m": MATH_NAMESPACE,
}
REL_ATTR_PREFIX = f"{{{RELATIONSHIP_NAMESPACE}}}"


@dataclass
class DocxToTexResult:
    source_docx: Path
    output_tex: Path
    media_dir: Path
    extracted_media_count: int
    table_count: int
    graphics_count: int
    math_count: int


@dataclass
class DocxFigure:
    target: str
    caption: str | None
    width_inches: float | None


def convert_docx_to_tex(
    docx_path: Path,
    output_tex: Path,
    media_dir: Path | None = None,
    standalone: bool = True,
) -> DocxToTexResult:
    source_docx = docx_path.resolve()
    output = output_tex.resolve()
    output.parent.mkdir(parents=True, exist_ok=True)

    media_root = (media_dir or output.with_name(f"{output.stem}_media")).resolve()
    if media_root.exists():
        shutil.rmtree(media_root)
    media_root.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="docx-to-tex-") as temp_dir_name:
        temp_dir = Path(temp_dir_name)
        prepared_docx = temp_dir / source_docx.name
        math_placeholders = prepare_docx_for_reverse_conversion(source_docx, prepared_docx)

        extract_media_arg = media_root.name if media_root.parent == output.parent else str(media_root)
        output_name = output.name if output.parent else str(output)
        extra_args = ["--wrap=none", f"--extract-media={extract_media_arg}"]
        if standalone:
            extra_args.insert(0, "--standalone")

        with working_directory(output.parent):
            pypandoc.convert_file(
                str(prepared_docx),
                "latex",
                outputfile=output_name,
                extra_args=extra_args,
            )

    flatten_pandoc_media_subdir(media_root)
    latex_text = output.read_text(encoding="utf-8")
    latex_text = normalize_converted_latex(latex_text, media_root, math_placeholders)
    latex_text = ensure_fallback_figures(source_docx, latex_text, media_root)
    output.write_text(latex_text, encoding="utf-8")

    return DocxToTexResult(
        source_docx=source_docx,
        output_tex=output,
        media_dir=media_root,
        extracted_media_count=count_media_files(media_root),
        table_count=count_table_environments(latex_text),
        graphics_count=latex_text.count("\\includegraphics"),
        math_count=count_math_markers(latex_text),
    )


def prepare_docx_for_reverse_conversion(source_docx: Path, prepared_docx: Path) -> dict[str, str]:
    math_placeholders: dict[str, str] = {}
    with ZipFile(source_docx) as source_zip:
        archive_entries = [(info, source_zip.read(info.filename)) for info in source_zip.infolist()]

    document_xml = next((data for info, data in archive_entries if info.filename == "word/document.xml"), None)
    if document_xml is None:
        shutil.copy2(source_docx, prepared_docx)
        return math_placeholders

    document_tree = ET.fromstring(document_xml)
    replace_omml_math_with_placeholders(document_tree, math_placeholders)
    updated_document = ET.tostring(document_tree, encoding="utf-8", xml_declaration=True)

    with ZipFile(prepared_docx, "w", compression=ZIP_DEFLATED) as target_zip:
        for info, data in archive_entries:
            if info.filename == "word/document.xml":
                target_zip.writestr(info, updated_document)
                continue
            target_zip.writestr(info, data)
    return math_placeholders


def replace_omml_math_with_placeholders(element: ET.Element, math_placeholders: dict[str, str]) -> None:
    children = list(element)
    for child in children:
        if child.tag == f"{{{MATH_NAMESPACE}}}oMath":
            latex_math = omml_to_latex(child).strip()
            placeholder = f"TEXDOCXMATH{len(math_placeholders)}TOKEN"
            math_placeholders[placeholder] = latex_math
            insert_math_placeholder_runs(element, child, placeholder)
            continue
        replace_omml_math_with_placeholders(child, math_placeholders)


def insert_math_placeholder_runs(parent: ET.Element, math_element: ET.Element, placeholder: str) -> None:
    replacement_run = ET.Element(f"{{{WORD_NAMESPACE}}}r")
    text_node = ET.Element(f"{{{WORD_NAMESPACE}}}t")
    text_node.text = placeholder
    replacement_run.append(text_node)
    children = list(parent)
    index = children.index(math_element)
    parent.remove(math_element)
    parent.insert(index, replacement_run)


def omml_to_latex(element: ET.Element) -> str:
    local_name = local_tag_name(element.tag)
    if local_name in {"oMath", "e", "sup", "sub", "num", "den", "deg", "radPr", "dPr", "naryPr"}:
        return "".join(omml_to_latex(child) for child in list(element))
    if local_name == "r":
        return "".join(omml_to_latex(child) for child in list(element))
    if local_name == "t":
        return element.text or ""
    if local_name == "sSup":
        base = omml_to_latex(find_math_child(element, "e"))
        superscript = omml_to_latex(find_math_child(element, "sup"))
        return f"{base}^{{{superscript}}}"
    if local_name == "sSub":
        base = omml_to_latex(find_math_child(element, "e"))
        subscript = omml_to_latex(find_math_child(element, "sub"))
        return f"{base}_{{{subscript}}}"
    if local_name == "sSubSup":
        base = omml_to_latex(find_math_child(element, "e"))
        subscript = omml_to_latex(find_math_child(element, "sub"))
        superscript = omml_to_latex(find_math_child(element, "sup"))
        return f"{base}_{{{subscript}}}^{{{superscript}}}"
    if local_name == "f":
        numerator = omml_to_latex(find_math_child(element, "num"))
        denominator = omml_to_latex(find_math_child(element, "den"))
        return f"\\frac{{{numerator}}}{{{denominator}}}"
    if local_name == "rad":
        degree = omml_to_latex(find_math_child(element, "deg"))
        radicand = omml_to_latex(find_math_child(element, "e"))
        if degree:
            return f"\\sqrt[{degree}]{{{radicand}}}"
        return f"\\sqrt{{{radicand}}}"
    if local_name == "d":
        return f"({omml_to_latex(find_math_child(element, 'e'))})"
    if local_name == "nary":
        body = omml_to_latex(find_math_child(element, "e"))
        return body
    return "".join(omml_to_latex(child) for child in list(element))


def find_math_child(element: ET.Element, local_name: str) -> ET.Element:
    for child in list(element):
        if local_tag_name(child.tag) == local_name:
            return child
    return ET.Element(f"{{{MATH_NAMESPACE}}}{local_name}")


def local_tag_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def flatten_pandoc_media_subdir(media_root: Path) -> None:
    nested_media = media_root / "media"
    if not nested_media.exists() or not nested_media.is_dir():
        return

    for child in list(nested_media.iterdir()):
        target = media_root / child.name
        if target.exists():
            if target.is_dir():
                shutil.rmtree(target)
            else:
                target.unlink()
        shutil.move(str(child), str(target))
    nested_media.rmdir()


def normalize_converted_latex(latex_text: str, media_root: Path, math_placeholders: dict[str, str]) -> str:
    normalized = latex_text.replace("\\includegraphics{", "\\includegraphics[]{}")
    normalized = normalized.replace("\\includegraphics[]{}", "\\includegraphics{")
    media_pattern = re.escape(f"{media_root.name}/media/")
    normalized = re.sub(media_pattern, f"{media_root.name}/", normalized)
    normalized = normalized.replace("\\tightlist", "")
    normalized = normalized.replace("\r\n", "\n")
    for placeholder, latex_math in math_placeholders.items():
        normalized = normalized.replace(placeholder, f"${latex_math}$")
    return normalized


def ensure_fallback_figures(source_docx: Path, latex_text: str, media_root: Path) -> str:
    if latex_text.count("\\includegraphics") > 0:
        return latex_text

    figures = extract_docx_figures(source_docx)
    if not figures:
        return latex_text

    extract_missing_figure_media(source_docx, figures, media_root)
    updated_text = latex_text
    for figure in figures:
        if not figure.caption:
            continue
        includegraphics = build_fallback_figure_block(figure, media_root)
        updated_text = updated_text.replace(figure.caption, includegraphics, 1)
    return updated_text


def extract_docx_figures(docx_path: Path) -> list[DocxFigure]:
    with ZipFile(docx_path) as archive:
        document_root = ET.fromstring(archive.read("word/document.xml"))
        relationships = parse_document_relationships(archive.read("word/_rels/document.xml.rels"))

    body = document_root.find("w:body", XML_NAMESPACES)
    if body is None:
        return []

    paragraphs = [child for child in list(body) if local_tag_name(child.tag) == "p"]
    figures: list[DocxFigure] = []
    for index, paragraph in enumerate(paragraphs):
        drawing = paragraph.find(".//w:drawing", XML_NAMESPACES)
        if drawing is None:
            continue
        embed_id = next((element.get(f"{REL_ATTR_PREFIX}embed") for element in drawing.iter() if element.get(f"{REL_ATTR_PREFIX}embed")), None)
        target = relationships.get(embed_id or "")
        if not target:
            continue
        caption = find_following_caption(paragraphs, index + 1)
        extent = drawing.find(".//wp:extent", XML_NAMESPACES)
        width_inches = None
        if extent is not None and extent.get("cx"):
            width_inches = int(extent.get("cx", "0")) / 914400
        figures.append(DocxFigure(target=target, caption=caption, width_inches=width_inches))
    return figures


def parse_document_relationships(rels_xml: bytes) -> dict[str, str]:
    relationships_root = ET.fromstring(rels_xml)
    relationships: dict[str, str] = {}
    for rel in relationships_root.findall(f"{{{PACKAGE_RELATIONSHIP_NAMESPACE}}}Relationship"):
        rel_id = rel.get("Id")
        target = rel.get("Target")
        if rel_id and target:
            relationships[rel_id] = target
    return relationships


def find_following_caption(paragraphs: list[ET.Element], start_index: int) -> str | None:
    for paragraph in paragraphs[start_index:]:
        text = paragraph_text(paragraph)
        if not text:
            continue
        if re.match(r"^(图|Figure)\s*[0-9]+", text):
            return text
        break
    return None


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", XML_NAMESPACES)).strip()


def extract_missing_figure_media(source_docx: Path, figures: list[DocxFigure], media_root: Path) -> None:
    with ZipFile(source_docx) as archive:
        for figure in figures:
            target_name = Path(figure.target).name
            destination = media_root / target_name
            if destination.exists():
                continue
            archive_path = f"word/{figure.target.lstrip('/')}"
            try:
                data = archive.read(archive_path)
            except KeyError:
                continue
            destination.write_bytes(data)


def build_fallback_figure_block(figure: DocxFigure, media_root: Path) -> str:
    width_clause = ""
    if figure.width_inches:
        width_clause = f"[width={figure.width_inches:.5f}in]"
    resource_path = f"{media_root.name}/{Path(figure.target).name}"
    caption = figure.caption or Path(figure.target).stem
    return (
        "\\begin{figure}\n"
        "\\centering\n"
        f"\\includegraphics{width_clause}{{{resource_path}}}\n"
        f"\\caption{{\\textbf{{{caption}}}}}\n"
        "\\end{figure}"
    )


def count_media_files(media_root: Path) -> int:
    if not media_root.exists():
        return 0
    return sum(1 for path in media_root.rglob("*") if path.is_file())


def count_table_environments(latex_text: str) -> int:
    return sum(
        latex_text.count(token)
        for token in ("\\begin{table}", "\\begin{table*}", "\\begin{longtable}", "\\begin{tabular}")
    )


def count_math_markers(latex_text: str) -> int:
    tokens = [
        "\\(",
        "\\[",
        "\\begin{equation}",
        "\\begin{align}",
        "\\begin{gather}",
        "\\begin{multline}",
        "$$",
    ]
    inline_dollar_pairs = len(re.findall(r"(?<!\\)\$(?!\$)", latex_text)) // 2
    return sum(latex_text.count(token) for token in tokens) + inline_dollar_pairs


@contextmanager
def working_directory(path: Path):
    previous = Path.cwd()
    os.chdir(path)
    try:
        yield
    finally:
        os.chdir(previous)