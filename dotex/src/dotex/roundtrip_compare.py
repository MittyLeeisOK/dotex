from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from zipfile import ZipFile


WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MATH_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/math"
XML_NAMESPACES = {"w": WORD_NAMESPACE, "m": MATH_NAMESPACE}


@dataclass
class DocxMetrics:
    path: str
    table_count: int
    drawing_count: int
    media_count: int
    omml_formula_count: int
    caption_like_paragraphs: int


@dataclass
class TexMetrics:
    path: str
    table_count: int
    longtable_count: int
    figure_count: int
    includegraphics_count: int
    math_count: int
    caption_count: int
    label_count: int


@dataclass
class RoundtripComparison:
    source_docx: DocxMetrics
    source_tex: TexMetrics
    generated_docx: DocxMetrics


def analyze_docx(docx_path: Path) -> DocxMetrics:
    with ZipFile(docx_path) as archive:
        document_root = ET.fromstring(archive.read("word/document.xml"))
        media_count = sum(1 for name in archive.namelist() if name.startswith("word/media/") and not name.endswith("/"))

    paragraphs = document_root.findall(".//w:p", XML_NAMESPACES)
    caption_like_paragraphs = 0
    for paragraph in paragraphs:
        text = "".join(node.text or "" for node in paragraph.findall(".//w:t", XML_NAMESPACES)).strip()
        if re.match(r"^(图|表格)\s*[0-9]+", text):
            caption_like_paragraphs += 1

    return DocxMetrics(
        path=str(docx_path),
        table_count=len(document_root.findall(".//w:tbl", XML_NAMESPACES)),
        drawing_count=len(document_root.findall(".//w:drawing", XML_NAMESPACES)),
        media_count=media_count,
        omml_formula_count=len(document_root.findall(".//m:oMath", XML_NAMESPACES))
        + len(document_root.findall(".//m:oMathPara", XML_NAMESPACES)),
        caption_like_paragraphs=caption_like_paragraphs,
    )


def analyze_tex(tex_path: Path) -> TexMetrics:
    text = tex_path.read_text(encoding="utf-8")
    return TexMetrics(
        path=str(tex_path),
        table_count=text.count("\\begin{table}") + text.count("\\begin{table*}") + text.count("\\begin{longtable}"),
        longtable_count=text.count("\\begin{longtable}"),
        figure_count=text.count("\\begin{figure}") + text.count("\\begin{figure*}"),
        includegraphics_count=text.count("\\includegraphics"),
        math_count=count_tex_math(text),
        caption_count=text.count("\\caption{"),
        label_count=text.count("\\label{"),
    )


def count_tex_math(text: str) -> int:
    tokens = [
        "\\(",
        "\\[",
        "\\begin{equation}",
        "\\begin{align}",
        "\\begin{gather}",
        "\\begin{multline}",
        "$$",
    ]
    return sum(text.count(token) for token in tokens)


def build_roundtrip_comparison(
    source_docx_path: Path,
    source_tex_path: Path,
    generated_docx_path: Path,
) -> RoundtripComparison:
    return RoundtripComparison(
        source_docx=analyze_docx(source_docx_path),
        source_tex=analyze_tex(source_tex_path),
        generated_docx=analyze_docx(generated_docx_path),
    )


def render_roundtrip_report(comparison: RoundtripComparison) -> str:
    source_docx = comparison.source_docx
    source_tex = comparison.source_tex
    generated_docx = comparison.generated_docx

    lines = [
        "# Roundtrip Comparison Report",
        "",
        "## Inputs",
        "",
        f"- Source DOCX: {source_docx.path}",
        f"- Source TeX: {source_tex.path}",
        f"- Generated DOCX: {generated_docx.path}",
        "",
        "## Tables",
        "",
        "| Surface | Count | Notes |",
        "| --- | ---: | --- |",
        f"| Source DOCX tables | {source_docx.table_count} | Word tables in original manuscript |",
        f"| Source TeX tables | {source_tex.table_count} | `table`, `table*`, and `longtable` environments |",
        f"| Source TeX longtables | {source_tex.longtable_count} | subset of source TeX tables |",
        f"| Generated DOCX tables | {generated_docx.table_count} | Word tables after TeX→DOCX conversion |",
        "",
        "## Figures And Media",
        "",
        "| Surface | Count | Notes |",
        "| --- | ---: | --- |",
        f"| Source DOCX drawings | {source_docx.drawing_count} | drawing elements in original Word file |",
        f"| Source DOCX media files | {source_docx.media_count} | files under `word/media/` |",
        f"| Source TeX figure environments | {source_tex.figure_count} | `figure` and `figure*` environments |",
        f"| Source TeX includegraphics | {source_tex.includegraphics_count} | explicit graphics inclusions |",
        f"| Generated DOCX drawings | {generated_docx.drawing_count} | drawing elements after TeX→DOCX conversion |",
        f"| Generated DOCX media files | {generated_docx.media_count} | files under `word/media/` |",
        "",
        "## Formulas",
        "",
        "| Surface | Count | Notes |",
        "| --- | ---: | --- |",
        f"| Source DOCX OMML formulas | {source_docx.omml_formula_count} | Office Math objects in the original DOCX |",
        f"| Source TeX math markers | {source_tex.math_count} | display and inline LaTeX math markers |",
        f"| Generated DOCX OMML formulas | {generated_docx.omml_formula_count} | Office Math objects after TeX→DOCX conversion |",
        "",
        "## Caption And Label Signals",
        "",
        f"- Source DOCX caption-like paragraphs: {source_docx.caption_like_paragraphs}",
        f"- Source TeX captions: {source_tex.caption_count}",
        f"- Source TeX labels: {source_tex.label_count}",
        f"- Generated DOCX caption-like paragraphs: {generated_docx.caption_like_paragraphs}",
        "",
        "## Observations",
        "",
    ]

    observations = build_observations(comparison)
    if observations:
        lines.extend(f"- {item}" for item in observations)
    else:
        lines.append("- No notable differences detected by the structural comparison.")

    lines.append("")
    return "\n".join(lines)


def build_observations(comparison: RoundtripComparison) -> list[str]:
    source_docx = comparison.source_docx
    source_tex = comparison.source_tex
    generated_docx = comparison.generated_docx
    observations: list[str] = []

    if source_tex.table_count != generated_docx.table_count:
        observations.append(
            f"Source TeX exposes {source_tex.table_count} table environments, while the generated DOCX currently contains {generated_docx.table_count} Word tables."
        )
    else:
        observations.append("Source TeX table count and generated DOCX table count are aligned at the manuscript level.")

    if source_tex.figure_count != generated_docx.drawing_count:
        observations.append(
            f"Source TeX has {source_tex.figure_count} figure environments, while the generated DOCX contains {generated_docx.drawing_count} drawing elements."
        )
    else:
        observations.append("Source TeX figure count and generated DOCX drawing count are aligned at the manuscript level.")

    if source_docx.omml_formula_count == 0 and generated_docx.omml_formula_count == 0:
        observations.append("Neither the original DOCX nor the generated DOCX stores formulas as OMML objects in this manuscript.")
    elif source_docx.omml_formula_count != generated_docx.omml_formula_count:
        observations.append(
            f"The original DOCX stores {source_docx.omml_formula_count} OMML formulas, while the generated DOCX stores {generated_docx.omml_formula_count}, so formula carriers are not symmetric across the roundtrip."
        )
    if source_tex.math_count == 0:
        observations.append("The source TeX does not expose standalone math environments in this manuscript, so formula roundtrip risk is currently low but not exercised by this sample.")
    if generated_docx.media_count != source_docx.media_count:
        observations.append(
            f"The generated DOCX contains {generated_docx.media_count} media files versus {source_docx.media_count} in the original DOCX, so media packaging is not a one-to-one proxy for visual figure count."
        )

    return observations