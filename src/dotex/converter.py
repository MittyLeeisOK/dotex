from __future__ import annotations

import hashlib
import json
import re
import xml.etree.ElementTree as ET
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

import pypandoc

from dotex.zotero_resolver import (
    copied_zotero_database,
    normalize_doi,
    normalize_url,
    parse_bibliography_entries,
    resolve_bibliography_against_zotero,
)


WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types"
DRAWINGML_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main"
MATH_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/math"
XML_NAMESPACES = {
    "w": WORD_NAMESPACE,
    "r": RELATIONSHIP_NAMESPACE,
}
WORD_ATTR_PREFIX = f"{{{WORD_NAMESPACE}}}"
REL_ATTR_PREFIX = f"{{{RELATIONSHIP_NAMESPACE}}}"

DEFAULT_ZOTERO_DATABASE = Path("~/Zotero/zotero.sqlite").expanduser()

ET.register_namespace("w", WORD_NAMESPACE)
ET.register_namespace("r", RELATIONSHIP_NAMESPACE)

DEFAULT_PAPER_WIDTH_TWIPS = int(round(8.27 * 1440))
CURRENT_LENGTH_CONTEXT: dict[str, float] = {}
CURRENT_BIBLIOGRAPHY_ANCHORS: dict[str, str] = {}
DEFAULT_ZOTERO_FIELD_COLOR = "003399"
DEFAULT_INTERNAL_LINK_COLOR = DEFAULT_ZOTERO_FIELD_COLOR
ZOTERO_CITATION_INSTRUCTION_PREFIX = " ADDIN ZOTERO_ITEM CSL_CITATION "
ZOTERO_CITATION_SCHEMA_URL = "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"
ZOTERO_CITATION_ID_ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
TABLE_LAYOUT_NOTICE = "程序无法识别表格排版，请注意手动处理表格样式"
CAPTION_PLACEHOLDER_PREFIX = "DOTEX_CAPTION "
CROSS_REFERENCE_ANCHOR_PREFIX = "dotex-xref-"
NATIVE_CROSS_REFERENCE_BOOKMARK_PREFIX = "_Ref"
DIRECT_ZOTERO_COMPANION_FILENAME = "dotex_zotero_items.json"
WESTERN_FONT_FAMILY = "Times New Roman"
FIELD_INSTRUCTION_CHUNK_SIZE = 240
THREE_LINE_OUTER_BORDER_SIZE = "12"
THREE_LINE_HEADER_BORDER_SIZE = "4"
OPENING_CITATION_BRACKETS = ("(", "（")
CLOSING_CITATION_BRACKETS = (")", "）")
TABLE_PROPERTY_CHILD_ORDER = (
    "tblStyle",
    "tblpPr",
    "tblOverlap",
    "bidiVisual",
    "tblStyleRowBandSize",
    "tblStyleColBandSize",
    "tblW",
    "jc",
    "tblCellSpacing",
    "tblInd",
    "tblBorders",
    "shd",
    "tblLayout",
    "tblCellMar",
    "tblLook",
    "tblCaption",
    "tblDescription",
)
TABLE_BORDER_CHILD_ORDER = ("top", "left", "bottom", "right", "insideH", "insideV")
CELL_PROPERTY_CHILD_ORDER = (
    "cnfStyle",
    "tcW",
    "gridSpan",
    "hMerge",
    "vMerge",
    "tcBorders",
    "shd",
    "noWrap",
    "tcMar",
    "textDirection",
    "tcFitText",
    "vAlign",
    "hideMark",
    "headers",
)
CELL_BORDER_CHILD_ORDER = ("top", "left", "bottom", "right")


@dataclass
class ConversionResult:
    source_tex: Path
    template_docx: Path
    normalized_source_path: Path
    output_docx: Path
    diagnostics: ConversionDiagnostics


@dataclass
class CrossReferenceTarget:
    kind: str
    label: str
    number: str
    caption_text: str
    bookmark_name: str | None = None
    referenced: bool = False


@dataclass
class ConversionDiagnostics:
    warnings: list[str] = field(default_factory=list)
    notices: list[str] = field(default_factory=list)
    cross_reference_targets: list[CrossReferenceTarget] = field(default_factory=list)
    fallback_citation_count: int = 0

    def add_warning(self, message: str) -> None:
        if message not in self.warnings:
            self.warnings.append(message)

    def add_notice(self, message: str) -> None:
        if message not in self.notices:
            self.notices.append(message)

    def missing_cross_reference_targets(self) -> list[CrossReferenceTarget]:
        return [target for target in self.cross_reference_targets if not target.referenced]


@dataclass
class TemplateDocxHints:
    caption_style_id: str | None
    table_style_id: str | None
    table_paragraph_style_id: str | None
    normal_style_id: str | None
    title_style_id: str | None
    heading_1_style_id: str | None
    heading_2_style_id: str | None
    heading_3_style_id: str | None
    bibliography_style_id: str | None
    zotero_item_uri_prefix: str | None


@dataclass
class TableLayoutHint:
    centered: bool
    total_width_ratio: float | None
    column_width_ratios: list[float | None]
    column_alignments: list[str | None]


@dataclass
class FigureLayoutHint:
    centered: bool
    width_ratio: float | None


@dataclass
class DocumentLayoutHints:
    length_context: dict[str, float]
    tables: list[TableLayoutHint]
    figures: list[FigureLayoutHint]


@dataclass
class CitationTarget:
    source_key: str
    formatted_reference: str
    zotero_item_key: str | None
    item_data: dict
    uri: str | None
    anchor_id: str


@dataclass
class CitationFieldShell:
    field_nodes_xml: list[str]


@dataclass
class UnmatchedZoteroNotice:
    source_key: str
    formatted_reference: str
    import_url: str | None


@dataclass
class ZoteroDocxContext:
    bibliography_entries: list[CitationTarget]
    unmatched_notices: list[UnmatchedZoteroNotice]
    by_anchor: dict[str, CitationTarget]
    by_normalized_url: dict[str, CitationTarget]
    by_normalized_doi: dict[str, CitationTarget]
    citation_field_shells: dict[tuple[tuple[str, ...], str], list[CitationFieldShell]] = field(default_factory=dict)

    def lookup(self, target: str | None = None, anchor: str | None = None) -> CitationTarget | None:
        if anchor and anchor in self.by_anchor:
            return self.by_anchor[anchor]
        if not target:
            return None
        normalized_url = normalize_url(target)
        if normalized_url and normalized_url in self.by_normalized_url:
            return self.by_normalized_url[normalized_url]
        normalized_doi = normalize_doi(target)
        if normalized_doi and normalized_doi in self.by_normalized_doi:
            return self.by_normalized_doi[normalized_doi]
        return None

    def lookup_display_text(self, display_text: str) -> CitationTarget | None:
        signature = citation_display_signature(display_text)
        if signature is None:
            return None
        candidates = [
            entry
            for entry in self.bibliography_entries
            if bibliography_entry_matches_signature(entry, signature)
        ]
        if len(candidates) == 1:
            return candidates[0]
        return None


def build_initial_conversion_diagnostics(document_layout_hints: DocumentLayoutHints) -> ConversionDiagnostics:
    diagnostics = ConversionDiagnostics()
    if document_layout_hints.tables:
        diagnostics.add_notice(TABLE_LAYOUT_NOTICE)
    return diagnostics


def convert_tex_to_docx(
    tex_path: Path,
    template_docx: Path,
    output_docx: Path,
    artifacts_dir: Path,
    bibliography_path: Path | None = None,
    bibliography_heading: str = "参考文献",
    enable_zotero: bool = False,
    use_native_bookmarks: bool = True,
    zotero_database: Path | None = None,
) -> ConversionResult:
    source_tex = tex_path.resolve()
    source_text = source_tex.read_text(encoding="utf-8")
    template = template_docx.resolve()
    output = output_docx.resolve()
    artifact_root = artifacts_dir.resolve()
    artifact_root.mkdir(parents=True, exist_ok=True)
    template_hints = infer_template_docx_hints(template)
    document_layout_hints = collect_document_layout_hints(source_text)
    diagnostics = build_initial_conversion_diagnostics(document_layout_hints)
    zotero_context = build_zotero_docx_context(
        source_tex,
        template_hints,
        bibliography_path=bibliography_path,
        source_text=source_text,
        enable_zotero=enable_zotero,
        zotero_database=zotero_database,
    )

    normalized_source = artifact_root / "normal_manuscript.normalized.md"
    normalized_source.write_text(
        normalize_tex_for_pandoc(
            source_tex,
            source_text=source_text,
            length_context=document_layout_hints.length_context,
            use_citation_hyperlinks=enable_zotero,
        ),
        encoding="utf-8",
    )

    output.parent.mkdir(parents=True, exist_ok=True)
    resource_path = str(source_tex.parent)
    pypandoc.convert_file(
        str(normalized_source),
        "docx",
        format=(
            "markdown+raw_html+tex_math_dollars+pipe_tables+fenced_divs"
            "+link_attributes+implicit_figures"
        ),
        outputfile=str(output),
        extra_args=[
            "--standalone",
            f"--reference-doc={template}",
            f"--resource-path={resource_path}",
            "--wrap=none",
        ],
    )
    postprocess_generated_docx(
        output,
        template,
        template_hints,
        zotero_context,
        document_layout_hints,
        diagnostics,
        bibliography_heading=bibliography_heading,
        enable_zotero=enable_zotero,
        use_native_bookmarks=use_native_bookmarks,
    )

    return ConversionResult(
        source_tex=source_tex,
        template_docx=template,
        normalized_source_path=normalized_source,
        output_docx=output,
        diagnostics=diagnostics,
    )


def inject_preceding_labels_into_environments(body: str, env_names: list[str]) -> str:
    """Move standalone \\label{id} that immediately precede \\begin{env} inside the environment block.

    This is needed when pandoc emits \\label{} before \\begin{figure} rather than inside the caption.
    After injection, extract_label() can find the label inside the block.
    """
    for env_name in env_names:
        body = re.sub(
            r"(\\label\{[^}]+\}\{?\}?)\s*(\\begin\{" + re.escape(env_name) + r"\}(?:\[[^\]]*\])?)",
            lambda m: m.group(2) + "\n" + m.group(1),
            body,
        )
    return body


def unwrap_formatting_wrapped_environments(body: str, env_names: list[str]) -> str:
    """Strip inline formatting wrappers like \textbf{...} around block environments.

    Pandoc sometimes wraps a whole figure environment in inline formatting commands.
    When that happens, labels emitted immediately before the wrapper no longer sit
    directly in front of \begin{figure}, so later label injection misses them.
    """
    wrappers = ("textbf", "emph")
    for env_name in env_names:
        for wrapper in wrappers:
            body = re.sub(
                r"\\" + wrapper + r"\{\s*(\\begin\{" + re.escape(env_name) + r"\}(?:\[[^\]]*\])?.*?\\end\{" + re.escape(env_name) + r"\})\s*\}",
                lambda m: m.group(1),
                body,
                flags=re.S,
            )
    return body


def convert_hyperref_commands(body: str) -> str:
    """Convert \\hyperref[label]{text} to cross-reference anchor links.

    These links are later rewritten by rewrite_cross_reference_hyperlinks into
    Word REF fields pointing to caption bookmarks.
    """
    def replace_hyperref(m: re.Match) -> str:
        label = m.group(1)
        text = m.group(2)
        anchor = make_cross_reference_anchor(label)
        return f"[{text}](#{anchor})"

    return re.sub(r"\\hyperref\[([^\]]+)\]\{([^}]+)\}", replace_hyperref, body)


def normalize_tex_for_pandoc(
    tex_path: Path,
    source_text: str | None = None,
    length_context: dict[str, float] | None = None,
    use_citation_hyperlinks: bool = False,
) -> str:
    if source_text is None:
        source_text = tex_path.read_text(encoding="utf-8")
    body = extract_document_body(source_text)
    metadata, body = extract_front_matter(body)
    # If \title{} was in the preamble (before \begin{document}), it won't be
    # found by extract_front_matter above. Fall back to searching the preamble.
    if not metadata.get("title"):
        begin_doc = source_text.find("\\begin{document}")
        preamble = source_text[:begin_doc] if begin_doc != -1 else ""
        preamble_meta, _ = extract_front_matter(preamble)
        if preamble_meta.get("title"):
            metadata["title"] = preamble_meta["title"]
        if not metadata.get("authors") and preamble_meta.get("authors"):
            metadata["authors"] = preamble_meta["authors"]
    labels = parse_label_numbers(tex_path.with_suffix(".aux"))
    global CURRENT_LABELS, CURRENT_LENGTH_CONTEXT, CURRENT_BIBLIOGRAPHY_ANCHORS
    CURRENT_LABELS = labels
    CURRENT_LENGTH_CONTEXT = length_context or extract_length_context(source_text)
    CURRENT_BIBLIOGRAPHY_ANCHORS = build_bibliography_anchor_map(tex_path, source_text)

    body = replace_command_two_args(body, "litref", render_litref_link)
    body = replace_command_one_arg(body, "tabref", lambda label: render_cross_reference(label, labels, "表"))
    body = replace_command_one_arg(body, "figref", lambda label: render_cross_reference(label, labels, "图"))
    body = replace_command_one_arg(body, "eqref", lambda label: render_cross_reference(label, labels, "公式", wrap_parentheses=True))
    body = replace_command_one_arg(body, "detokenize", lambda value: value)
    body = replace_generic_refs(body, labels)
    bib_display = _load_refs_display(tex_path.parent)
    if bib_display:
        bib_key_source = _load_bib_key_to_source_key(tex_path.parent) if use_citation_hyperlinks else {}

        def _render_cite_item(k: str, include_hyperlink: bool) -> str:
            display = bib_display.get(k, k).replace("\u00a0", " ")
            if not include_hyperlink:
                return display
            source_key = bib_key_source.get(k.strip(), k.strip())
            anchor = CURRENT_BIBLIOGRAPHY_ANCHORS.get(source_key, make_bibliography_anchor_id(source_key))
            return f"[{display}](#{anchor})"

        def _expand_parencite(keys: str) -> str:
            parts = [_render_cite_item(k.strip(), use_citation_hyperlinks) for k in keys.split(",")]
            return "(" + "; ".join(parts) + ")"

        def _expand_textcite(keys: str) -> str:
            return "; ".join(_render_cite_item(k.strip(), use_citation_hyperlinks) for k in keys.split(","))

        for cmd in ("parencite", "citep", "cite"):
            body = replace_command_one_arg(body, cmd, _expand_parencite)
        body = replace_command_one_arg(body, "textcite", _expand_textcite)
        body = replace_command_one_arg(body, "citet", _expand_textcite)
    body = inline_bibliography_inputs(body, tex_path.parent)
    body = strip_layout_only_commands(body)
    body = unwrap_formatting_wrapped_environments(body, ["figure"])
    body = inject_preceding_labels_into_environments(body, ["figure"])
    body = normalize_table_syntax(body)
    body = convert_figure_blocks(body)
    body = convert_longtable_blocks(body)
    body = convert_table_blocks(body)
    body = convert_equation_blocks(body)
    body = convert_hyperref_commands(body)
    body = convert_section_commands(body)
    body = convert_block_commands(body)
    body = convert_inline_tex_to_markdown(body)
    body = cleanup_markdown(body)
    return compose_markdown(metadata, body)


def extract_front_matter(body: str) -> tuple[dict, str]:
    title, body = pop_first_command(body, "title")
    author, body = pop_first_command(body, "author")
    date, body = pop_first_command(body, "date")
    body = body.replace("\\maketitle", "", 1)
    authors = []
    if author:
        authors = [
            convert_inline_tex_to_plain(part.strip())
            for part in author.split("\\and")
            if part.strip()
        ]
    metadata = {
        "title": convert_inline_tex_to_plain(title or ""),
        "authors": authors,
        "date": convert_inline_tex_to_plain(date or ""),
    }
    return metadata, body


def pop_first_command(text: str, command: str) -> tuple[str | None, str]:
    token = f"\\{command}"
    start = text.find(token)
    if start == -1:
        return None, text
    cursor = skip_whitespace(text, start + len(token))
    if cursor >= len(text) or text[cursor] != "{":
        return None, text
    value, end = read_braced(text, cursor)
    return value, text[:start] + text[end:]


def extract_document_body(text: str) -> str:
    match = re.search(r"\\begin\{document\}(.*)\\end\{document\}", text, flags=re.S)
    if match is None:
        return text
    return match.group(1)


def collect_document_layout_hints(source_text: str) -> DocumentLayoutHints:
    length_context = extract_length_context(source_text)
    body = extract_document_body(source_text)
    tables: list[TableLayoutHint] = []
    figures: list[FigureLayoutHint] = []

    index = 0
    while index < len(body):
        next_env: str | None = None
        next_start = -1
        for candidate in ("figure", "table", "longtable"):
            start = body.find(f"\\begin{{{candidate}}}", index)
            if start == -1:
                continue
            if next_start == -1 or start < next_start:
                next_start = start
                next_env = candidate
        if next_env is None or next_start == -1:
            break
        block, index = read_environment_block(body, next_env, next_start)
        if next_env == "figure":
            figures.append(build_figure_layout_hint(block, length_context))
        else:
            tables.append(build_table_layout_hint(block, next_env, length_context))

    return DocumentLayoutHints(length_context=length_context, tables=tables, figures=figures)


def read_environment_block(text: str, env_name: str, start: int) -> tuple[str, int]:
    end_token = f"\\end{{{env_name}}}"
    end = text.find(end_token, start)
    if end == -1:
        return text[start:], len(text)
    block_end = end + len(end_token)
    return text[start:block_end], block_end


def extract_length_context(source_text: str) -> dict[str, float]:
    page_width = infer_page_width_twips(source_text)
    margins = infer_geometry_margins_twips(source_text)
    text_width = max(page_width - margins["left"] - margins["right"], 1)

    lengths: dict[str, float] = {
        "paperwidth": float(page_width),
        "textwidth": float(text_width),
        "linewidth": float(text_width),
        "oddsidemargin": float(margins["left"] - unit_to_twips("1in")),
        "evensidemargin": float(margins["left"] - unit_to_twips("1in")),
    }

    for name in re.findall(r"\\newlength\{\\([^}]+)\}", source_text):
        lengths.setdefault(name, 0.0)

    for name, expression in parse_setlength_commands(source_text):
        try:
            lengths[name] = evaluate_tex_length(expression, lengths)
        except ValueError:
            continue

    return lengths


def infer_page_width_twips(source_text: str) -> int:
    lowered = source_text.casefold()
    if "letterpaper" in lowered:
        return int(round(8.5 * 1440))
    return DEFAULT_PAPER_WIDTH_TWIPS


def infer_geometry_margins_twips(source_text: str) -> dict[str, int]:
    default_margin = unit_to_twips("1in")
    margins = {"left": default_margin, "right": default_margin, "top": default_margin, "bottom": default_margin}
    match = re.search(r"\\usepackage\[([^\]]+)\]\{geometry\}", source_text)
    if match is None:
        return margins

    parsed: dict[str, int] = {}
    for option in (item.strip() for item in match.group(1).split(",") if item.strip()):
        if "=" not in option:
            continue
        key, value = [part.strip() for part in option.split("=", 1)]
        if re.fullmatch(r"[0-9]*\.?[0-9]+(?:cm|mm|in|pt)", value):
            parsed[key] = unit_to_twips(value)

    if "margin" in parsed:
        value = parsed["margin"]
        margins = {"left": value, "right": value, "top": value, "bottom": value}
    margins["left"] = parsed.get("left", margins["left"])
    margins["right"] = parsed.get("right", margins["right"])
    margins["top"] = parsed.get("top", margins["top"])
    margins["bottom"] = parsed.get("bottom", margins["bottom"])
    return margins


def parse_setlength_commands(source_text: str) -> list[tuple[str, str]]:
    commands: list[tuple[str, str]] = []
    token = "\\setlength"
    index = 0
    while True:
        start = source_text.find(token, index)
        if start == -1:
            break
        cursor = skip_whitespace(source_text, start + len(token))
        try:
            name, cursor = read_braced(source_text, cursor)
            expression, cursor = read_braced(source_text, skip_whitespace(source_text, cursor))
        except ValueError:
            index = start + len(token)
            continue
        commands.append((name.lstrip("\\"), expression.strip()))
        index = cursor
    return commands


def evaluate_tex_length(expression: str, lengths: dict[str, float]) -> float:
    sanitized = expression.replace("\\dimexpr", "").replace("\\relax", "").strip()
    if not sanitized:
        return 0.0
    sanitized = re.sub(r"(?<=\d)(\\[A-Za-z@]+)", r"*\1", sanitized)
    tokens: list[str] = []
    index = 0
    while index < len(sanitized):
        char = sanitized[index]
        if char.isspace():
            index += 1
            continue
        if char in "+-*/()":
            tokens.append(char)
            index += 1
            continue
        if char == "\\":
            match = re.match(r"\\([A-Za-z@]+)", sanitized[index:])
            if match is None:
                raise ValueError(f"Unsupported length token in: {expression}")
            tokens.append(str(lengths.get(match.group(1), 0.0)))
            index += len(match.group(0))
            continue
        match = re.match(r"[0-9]*\.?[0-9]+(?:cm|mm|in|pt)?", sanitized[index:])
        if match is None:
            raise ValueError(f"Unsupported length expression: {expression}")
        raw_value = match.group(0)
        if re.fullmatch(r"[0-9]*\.?[0-9]+(?:cm|mm|in|pt)", raw_value):
            tokens.append(str(unit_to_twips(raw_value)))
        else:
            tokens.append(raw_value)
        index += len(raw_value)

    return float(eval("".join(tokens), {"__builtins__": {}}, {}))


def unit_to_twips(raw_value: str) -> int:
    match = re.fullmatch(r"([0-9]*\.?[0-9]+)(cm|mm|in|pt)", raw_value)
    if match is None:
        raise ValueError(f"Unsupported length unit: {raw_value}")
    value = float(match.group(1))
    unit = match.group(2)
    if unit == "in":
        inches = value
    elif unit == "cm":
        inches = value / 2.54
    elif unit == "mm":
        inches = value / 25.4
    else:
        inches = value / 72.0
    return int(round(inches * 1440))


def build_figure_layout_hint(block: str, length_context: dict[str, float]) -> FigureLayoutHint:
    image_match = re.search(r"\\includegraphics(?:\[(?P<options>[^\]]*)\])?\{(?P<path>[^{}]+)\}", block)
    options = image_match.group("options") if image_match is not None else ""
    return FigureLayoutHint(
        centered=layout_block_is_centered(block),
        width_ratio=parse_includegraphics_width_ratio(options, length_context),
    )


def build_table_layout_hint(block: str, env_name: str, length_context: dict[str, float]) -> TableLayoutHint:
    inner_env_name, total_width_ratio, spec = extract_table_environment_layout(block, env_name, length_context)
    column_width_ratios: list[float | None] = []
    column_alignments: list[str | None] = []
    if spec:
        column_width_ratios, column_alignments, total_width_ratio = parse_column_layout_spec(
            spec,
            length_context,
            total_width_ratio,
        )
    centered = layout_block_is_centered(block)
    if inner_env_name == "longtable":
        centered = True
    return TableLayoutHint(
        centered=centered,
        total_width_ratio=total_width_ratio,
        column_width_ratios=column_width_ratios,
        column_alignments=column_alignments,
    )


def extract_table_environment_layout(
    block: str,
    env_name: str,
    length_context: dict[str, float],
) -> tuple[str, float | None, str | None]:
    if env_name == "longtable":
        arguments = parse_environment_arguments(block, "longtable")
        return "longtable", None, arguments[0] if arguments else None

    candidates: list[tuple[int, str]] = []
    for candidate in ("tabularx", "tabular", "longtable"):
        start = block.find(f"\\begin{{{candidate}}}")
        if start != -1:
            candidates.append((start, candidate))
    if not candidates:
        return env_name, None, None

    _, inner_env = min(candidates)
    inner_block = extract_first_environment_block(block, inner_env)
    if inner_block is None:
        return inner_env, None, None
    arguments = parse_environment_arguments(inner_block, inner_env)
    if inner_env == "tabularx":
        total_width_ratio = evaluate_width_ratio(arguments[0], length_context) if arguments else None
        return inner_env, total_width_ratio, arguments[1] if len(arguments) > 1 else None
    return inner_env, None, arguments[0] if arguments else None


def parse_environment_arguments(block: str, env_name: str) -> list[str]:
    token = f"\\begin{{{env_name}}}"
    start = block.find(token)
    if start == -1:
        return []
    cursor = start + len(token)
    cursor = skip_whitespace(block, cursor)
    if cursor < len(block) and block[cursor] == "[":
        _, cursor = read_bracketed(block, cursor)
    arguments: list[str] = []
    while True:
        cursor = skip_whitespace(block, cursor)
        if cursor >= len(block) or block[cursor] != "{":
            break
        value, cursor = read_braced(block, cursor)
        arguments.append(value)
    return arguments


def read_bracketed(text: str, start: int) -> tuple[str, int]:
    if start >= len(text) or text[start] != "[":
        raise ValueError("expected opening bracket")
    depth = 0
    parts: list[str] = []
    index = start
    while index < len(text):
        char = text[index]
        if char == "[" and not is_escaped(text, index):
            depth += 1
            if depth > 1:
                parts.append(char)
        elif char == "]" and not is_escaped(text, index):
            depth -= 1
            if depth == 0:
                return "".join(parts), index + 1
            parts.append(char)
        else:
            parts.append(char)
        index += 1
    raise ValueError("unclosed bracket")


def layout_block_is_centered(block: str) -> bool:
    return "\\centering" in block or "\\begin{center}" in block


def parse_column_layout_spec(
    spec: str,
    length_context: dict[str, float],
    total_width_ratio: float | None,
) -> tuple[list[float | None], list[str | None], float | None]:
    parsed_columns = parse_columns(spec, length_context)
    if not parsed_columns:
        return [], [], total_width_ratio

    fixed_sum = sum(column["width_ratio"] or 0.0 for column in parsed_columns if column["width_ratio"] is not None)
    unresolved_indices = [index for index, column in enumerate(parsed_columns) if column["width_ratio"] is None]
    if total_width_ratio is None and unresolved_indices:
        total_width_ratio = max(fixed_sum, 1.0)
    if unresolved_indices:
        target_width = max(total_width_ratio or fixed_sum, fixed_sum)
        remaining = max(target_width - fixed_sum, 0.0)
        share = remaining / len(unresolved_indices) if unresolved_indices else 0.0
        if share == 0.0:
            share = max(target_width, 1.0) / len(unresolved_indices)
        for index in unresolved_indices:
            parsed_columns[index]["width_ratio"] = share
    if total_width_ratio is None:
        total_width_ratio = sum(column["width_ratio"] or 0.0 for column in parsed_columns)

    return (
        [column["width_ratio"] for column in parsed_columns],
        [column["alignment"] for column in parsed_columns],
        total_width_ratio,
    )


def parse_columns(spec: str, length_context: dict[str, float]) -> list[dict[str, float | str | None]]:
    columns: list[dict[str, float | str | None]] = []
    decorators: list[str] = []
    index = 0
    while index < len(spec):
        char = spec[index]
        if char.isspace() or char in "|!":
            index += 1
            continue
        if char in {">", "<"}:
            index += 1
            index = skip_whitespace(spec, index)
            if index < len(spec) and spec[index] == "{":
                decorator, index = read_braced(spec, index)
                if char == ">":
                    decorators.append(decorator)
            continue
        if char == "@":
            index += 1
            index = skip_whitespace(spec, index)
            if index < len(spec) and spec[index] == "{":
                _, index = read_braced(spec, index)
            continue
        if char == "*":
            index += 1
            repeat_count_text, index = read_braced(spec, skip_whitespace(spec, index))
            repeated_spec, index = read_braced(spec, skip_whitespace(spec, index))
            try:
                repeat_count = max(int(repeat_count_text), 1)
            except ValueError:
                repeat_count = 1
            repeated_columns = parse_columns(repeated_spec, length_context)
            for _ in range(repeat_count):
                columns.extend(
                    {"width_ratio": item["width_ratio"], "alignment": item["alignment"]}
                    for item in repeated_columns
                )
            decorators = []
            continue
        if char in "lcrX":
            columns.append(
                {
                    "width_ratio": None,
                    "alignment": resolve_column_alignment(decorators, default_alignment=default_column_alignment(char)),
                }
            )
            decorators = []
            index += 1
            continue
        if char in "pmb":
            index += 1
            width_expression, index = read_braced(spec, skip_whitespace(spec, index))
            columns.append(
                {
                    "width_ratio": evaluate_width_ratio(width_expression, length_context),
                    "alignment": resolve_column_alignment(decorators, default_alignment="left"),
                }
            )
            decorators = []
            continue
        if char == "{":
            _, index = read_braced(spec, index)
            continue
        index += 1
    return columns


def resolve_column_alignment(decorators: list[str], default_alignment: str) -> str:
    decorator_text = " ".join(decorators)
    if "raggedleft" in decorator_text:
        return "right"
    if "centering" in decorator_text:
        return "center"
    if "raggedright" in decorator_text:
        return "left"
    return default_alignment


def default_column_alignment(column_type: str) -> str:
    if column_type == "c":
        return "center"
    if column_type == "r":
        return "right"
    return "left"


def evaluate_width_ratio(width_expression: str | None, length_context: dict[str, float]) -> float | None:
    if not width_expression:
        return None
    text_width = max(length_context.get("textwidth", 0.0), 1.0)
    try:
        width_value = evaluate_tex_length(width_expression, length_context)
    except ValueError:
        return None
    return width_value / text_width if width_value > 0 else None


def parse_label_numbers(aux_path: Path) -> dict[str, str]:
    if not aux_path.exists():
        return {}

    labels: dict[str, str] = {}
    pattern = re.compile(r"\\newlabel\{([^}]+)\}\{\{([^}]*)\}")
    for line in aux_path.read_text(encoding="utf-8").splitlines():
        match = pattern.search(line)
        if match is None:
            continue
        key, number = match.groups()
        if number:
            labels[key] = number
    return labels


def resolve_ref(label: str, labels: dict[str, str], prefix: str) -> str:
    number = labels.get(label)
    return f"{prefix} {number}" if number else label


def render_cross_reference(
    label: str,
    labels: dict[str, str],
    prefix: str,
    wrap_parentheses: bool = False,
) -> str:
    number = labels.get(label)
    if not number:
        return label
    display_text = format_numbered_reference(prefix, number)
    if wrap_parentheses:
        display_text = f"({display_text})"
    anchor = make_cross_reference_anchor(label)
    return f"[{display_text}](#{anchor})"


def replace_generic_refs(text: str, labels: dict[str, str]) -> str:
    def repl(match: re.Match[str]) -> str:
        label = match.group(1)
        if label.startswith("fig:"):
            return render_cross_reference(label, labels, "图")
        if label.startswith("tab:"):
            return render_cross_reference(label, labels, "表")
        if label.startswith("eq:"):
            return render_cross_reference(label, labels, "公式")
        return labels.get(label, label)

    return re.sub(r"\\ref\*?\{([^}]+)\}", repl, text)


def replace_command_one_arg(text: str, command: str, replacer) -> str:
    return replace_command(text, command, 1, lambda args: replacer(args[0]))


def replace_command_two_args(text: str, command: str, replacer) -> str:
    return replace_command(text, command, 2, lambda args: replacer(args[0], args[1]))


def replace_command(text: str, command: str, arg_count: int, replacer) -> str:
    token = f"\\{command}"
    chunks: list[str] = []
    index = 0
    while index < len(text):
        if text.startswith(token, index) and (index + len(token) == len(text) or not text[index + len(token)].isalpha()):
            cursor = index + len(token)
            args: list[str] = []
            try:
                for _ in range(arg_count):
                    cursor = skip_whitespace(text, cursor)
                    if cursor >= len(text) or text[cursor] != "{":
                        raise ValueError("missing argument")
                    arg, cursor = read_braced(text, cursor)
                    args.append(arg)
            except ValueError:
                chunks.append(text[index])
                index += 1
                continue
            chunks.append(replacer(args))
            index = cursor
            continue

        chunks.append(text[index])
        index += 1
    return "".join(chunks)


def skip_whitespace(text: str, index: int) -> int:
    while index < len(text) and text[index].isspace():
        index += 1
    return index


def read_braced(text: str, start: int) -> tuple[str, int]:
    if start >= len(text) or text[start] != "{":
        raise ValueError("expected opening brace")

    depth = 0
    parts: list[str] = []
    index = start
    while index < len(text):
        char = text[index]
        if char == "{" and not is_escaped(text, index):
            depth += 1
            if depth > 1:
                parts.append(char)
        elif char == "}" and not is_escaped(text, index):
            depth -= 1
            if depth == 0:
                return "".join(parts), index + 1
            parts.append(char)
        else:
            parts.append(char)
        index += 1

    raise ValueError("unclosed brace")


def is_escaped(text: str, index: int) -> bool:
    slash_count = 0
    probe = index - 1
    while probe >= 0 and text[probe] == "\\":
        slash_count += 1
        probe -= 1
    return slash_count % 2 == 1


def _load_refs_display(tex_dir: Path) -> dict[str, str]:
    """Load refs_display.json if it exists alongside the tex file."""
    refs_path = tex_dir / "refs_display.json"
    if refs_path.exists():
        return json.loads(refs_path.read_text(encoding="utf-8"))
    return {}


def _load_bib_key_to_doi(tex_dir: Path) -> dict[str, str]:
    """Parse refs.bib to build {bib_key: doi} mapping."""
    refs_bib = tex_dir / "refs.bib"
    if not refs_bib.exists():
        return {}
    text = refs_bib.read_text(encoding="utf-8")
    result: dict[str, str] = {}
    for entry_match in re.finditer(r"@\w+\{([^,\s{}\'\"]+),", text):
        key = entry_match.group(1).strip()
        entry_start = entry_match.end()
        next_entry = text.find("@", entry_start)
        entry_text = text[entry_start: next_entry if next_entry != -1 else len(text)]
        doi_match = re.search(r"\bdoi\s*=\s*\{([^}]+)\}", entry_text, re.IGNORECASE)
        if doi_match:
            result[key] = doi_match.group(1).strip()
    return result


def _load_bib_key_to_source_key(tex_dir: Path) -> dict[str, str]:
    refs_bib = tex_dir / "refs.bib"
    if not refs_bib.exists():
        return {}
    text = refs_bib.read_text(encoding="utf-8")
    result: dict[str, str] = {}
    for entry_match in re.finditer(r"@\w+\{([^,\s{}\'\"]+),", text):
        key = entry_match.group(1).strip()
        entry_start = entry_match.end()
        next_entry = text.find("@", entry_start)
        entry_text = text[entry_start: next_entry if next_entry != -1 else len(text)]
        doi_match = re.search(r"\bdoi\s*=\s*\{([^}]+)\}", entry_text, re.IGNORECASE)
        url_match = re.search(r"\burl\s*=\s*\{([^}]+)\}", entry_text, re.IGNORECASE)
        if doi_match:
            result[key] = doi_match.group(1).strip()
        elif url_match:
            result[key] = url_match.group(1).strip()
        else:
            result[key] = key
    return result


def inline_bibliography_inputs(body: str, base_dir: Path) -> str:
    pattern = re.compile(r"\\input\{([^}]+)\}")

    def repl(match: re.Match[str]) -> str:
        relative_path = match.group(1)
        bib_path = (base_dir / relative_path).resolve()
        if not bib_path.exists():
            return match.group(0)
        return render_bibliography_entries(bib_path)

    return pattern.sub(repl, body)


def render_plain_tex_bibliography(bib_path: Path) -> str:
    """Convert a plain-TeX formatted bibliography (no \\bibentry macros) to markdown."""
    text = bib_path.read_text(encoding="utf-8")
    # Strip leading comment lines
    text = re.sub(r"^%[^\n]*\n", "", text, flags=re.M)
    # Strip TeX layout preamble wrappers: {\small / {\footnotesize etc.
    text = re.sub(r"^\{\\(?:small|large|normalsize|scriptsize|footnotesize)\s*$", "", text, flags=re.M)
    # Strip layout-only lines
    text = re.sub(r"^\\(?:sloppy|raggedright|noindent|centering)\s*$", "", text, flags=re.M)
    text = re.sub(r"^\\setlength\{[^}]*\}\{[^}]*\}\s*$", "", text, flags=re.M)
    # Strip trailing closing brace from {\small group
    text = re.sub(r"^\}\s*$", "", text, flags=re.M)
    # Convert TeX inline formatting to markdown
    text = replace_command_one_arg(text, "emph", lambda v: f"*{v}*")
    text = replace_command_one_arg(text, "textbf", lambda v: f"**{v}**")
    # Strip URL wrapper macros
    text = replace_command_one_arg(text, "url", lambda v: v)
    text = replace_command_one_arg(text, "nolinkurl", lambda v: v)
    # Fix TeX escaped characters
    text = text.replace(r"\%", "%")
    text = text.replace(r"\_", "_")
    text = text.replace(r"\&", "&")
    text = text.replace(r"\#", "#")
    # Normalize excess blank lines
    text = re.sub(r"\n{3,}", "\n\n", text.strip())
    return text


def render_bibliography_entries(bib_path: Path) -> str:
    text = bib_path.read_text(encoding="utf-8")
    # If the file has no \bibentry commands, fall back to plain-TeX rendering
    if "\\bibentry" not in text:
        return render_plain_tex_bibliography(bib_path)
    text = replace_command_one_arg(text, "nolinkurl", lambda value: f"\\url{{{value}}}")
    entries: list[str] = []
    token = "\\bibentry"
    index = 0
    while True:
        start = text.find(token, index)
        if start == -1:
            break
        cursor = start + len(token)
        try:
            cursor = skip_whitespace(text, cursor)
            source_key, cursor = read_braced(text, cursor)
            cursor = skip_whitespace(text, cursor)
            entry_text, cursor = read_braced(text, cursor)
        except ValueError:
            index = start + len(token)
            continue
        anchor = CURRENT_BIBLIOGRAPHY_ANCHORS.get(source_key.strip(), make_bibliography_anchor_id(source_key))
        entries.append(render_anchor_div(anchor, entry_text.strip()))
        index = cursor
    return "\n\n".join(entries)


def render_litref_link(target: str, text: str) -> str:
    anchor = CURRENT_BIBLIOGRAPHY_ANCHORS.get(target.strip(), make_bibliography_anchor_id(target))
    return f"[{text}](#{anchor})"


def make_cross_reference_anchor(label: str) -> str:
    return f"{CROSS_REFERENCE_ANCHOR_PREFIX}{make_anchor_id(label)}"


def make_native_cross_reference_bookmark(label: str) -> str:
    digest = hashlib.sha1(label.encode("utf-8")).digest()
    numeric_suffix = int.from_bytes(digest[:5], "big") % 900_000_000 + 100_000_000
    return f"{NATIVE_CROSS_REFERENCE_BOOKMARK_PREFIX}{numeric_suffix}"


def caption_prefix_for_kind(kind: str) -> str:
    mapping = {
        "figure": "图",
        "table": "表",
        "equation": "公式",
    }
    return mapping.get(kind, kind)


def sequence_identifier_for_kind(kind: str) -> str:
    mapping = {
        "figure": "图",
        "table": "表",
        "equation": "公式",
    }
    return mapping.get(kind, "SEQ")


def format_numbered_reference(prefix: str, number: str) -> str:
    if not number:
        return prefix
    return f"{prefix}{number}"


def build_caption_placeholder(kind: str, label: str | None, caption_text: str, number: str | None) -> str:
    payload = {
        "kind": kind,
        "label": label or "",
        "caption": caption_text,
        "number": number or "",
    }
    return f"{CAPTION_PLACEHOLDER_PREFIX}{json.dumps(payload, ensure_ascii=False, separators=(',', ':'))}"


def parse_caption_placeholder(text: str) -> dict[str, str] | None:
    normalized = text.strip()
    if not normalized.startswith(CAPTION_PLACEHOLDER_PREFIX):
        return None
    payload = normalized[len(CAPTION_PLACEHOLDER_PREFIX) :].strip()
    if not payload:
        return None
    payload = (
        payload.replace("“", '"')
        .replace("”", '"')
        .replace("‘", "'")
        .replace("’", "'")
    )
    try:
        loaded = json.loads(payload)
    except json.JSONDecodeError:
        return None
    if not isinstance(loaded, dict):
        return None
    return {
        "kind": str(loaded.get("kind", "")),
        "label": str(loaded.get("label", "")),
        "caption": str(loaded.get("caption", "")),
        "number": str(loaded.get("number", "")),
    }


def render_native_caption_block(kind: str, label: str | None, caption: str) -> str:
    caption_text = convert_inline_tex_to_plain(caption)
    number = CURRENT_LABELS.get(label, "") if label else ""
    if not caption_text and not number and not label:
        return ""
    return render_custom_style_block(
        build_caption_placeholder(kind, label, caption_text, number),
        "caption",
    )


def build_bibliography_anchor_map(tex_path: Path, source_text: str) -> dict[str, str]:
    bibliography_path = infer_bibliography_path(tex_path, source_text, None)
    if bibliography_path is None or not bibliography_path.exists():
        return {}
    used: set[str] = set()
    anchors: dict[str, str] = {}
    for index, entry in enumerate(parse_bibliography_entries(bibliography_path), start=1):
        anchors[entry.source_key] = make_author_year_bookmark_name(entry.formatted_reference, index, used)
    return anchors


def make_author_year_bookmark_name(reference_text: str, index: int, used: set[str]) -> str:
    author = "Ref"
    initial = ""
    year = ""
    match = re.match(r"\s*([^,\(\.]+)\s*,\s*([^,\(\.]?)", reference_text.replace("\xa0", " "))
    if match:
        author = normalize_bookmark_token(match.group(1)) or "Ref"
        if match.group(2):
            initial = normalize_bookmark_token(match.group(2)[:1])
    year_match = re.search(r"(19|20)\d{2}", reference_text)
    if year_match:
        year = year_match.group(0)
    parts = [part for part in [author, initial, year] if part]
    if not parts:
        parts = ["Ref", f"{index:03d}"]
    base = "_".join(parts)
    candidate = base
    suffix = 2
    while candidate in used:
        candidate = f"{base}_{suffix}"
        suffix += 1
    used.add(candidate)
    return candidate


def normalize_bookmark_token(value: str) -> str:
    normalized = re.sub(r"[^0-9A-Za-z\u4e00-\u9fff]+", "_", value).strip("_")
    if not normalized:
        return ""
    if normalized[0].isdigit():
        normalized = f"Ref_{normalized}"
    return normalized


def convert_figure_blocks(body: str) -> str:
    return replace_environment_blocks(body, "figure", render_figure_block)


def render_figure_block(block: str) -> str:
    image_match = re.search(r"\\includegraphics(?:\[(?P<options>[^\]]*)\])?\{(?P<path>[^{}]+)\}", block)
    if image_match is None:
        return ""

    caption = extract_simple_command_arg(block, "caption") or ""
    label = extract_label(block)
    image_path = image_match.group("path").strip()
    options = image_match.group("options") or ""
    image_attributes = render_image_attributes(None, options)
    caption_block = render_native_caption_block("figure", label, caption)
    return f"\n\n![]({image_path}){image_attributes}\n\n{caption_block}\n\n"


def convert_longtable_blocks(body: str) -> str:
    return replace_environment_blocks(body, "longtable", render_longtable_block)


def render_longtable_block(block: str) -> str:
    caption = extract_simple_command_arg(block, "caption") or ""
    label = extract_label(block)
    rows = parse_table_rows(block, env_name="longtable")
    return render_markdown_table(rows, caption, label, kind="table")


def convert_table_blocks(body: str) -> str:
    return replace_environment_blocks(body, "table", render_table_block)


def render_table_block(block: str) -> str:
    caption = extract_simple_command_arg(block, "caption") or ""
    label = extract_label(block)
    tabular_block = extract_first_environment_block(block, "tabular")
    if tabular_block is None:
        caption_block = render_native_caption_block("table", label, caption)
        if caption_block:
            return f"\n\n{caption_block}\n\n"
        return "\n\n"
    rows = parse_table_rows(tabular_block, env_name="tabular")
    return render_markdown_table(rows, caption, label, kind="table")


def convert_equation_blocks(body: str) -> str:
    for env_name in (
        "equation",
        "equation*",
        "align",
        "align*",
        "gather",
        "gather*",
        "multline",
        "multline*",
        "displaymath",
    ):
        body = replace_environment_blocks(
            body,
            env_name,
            lambda block, env_name=env_name: render_equation_block(block, env_name),
        )
    return body


def render_equation_block(block: str, env_name: str) -> str:
    label = extract_label(block)
    math_body = strip_environment_wrapper(block, env_name)
    math_body = re.sub(r"\\label\{[^}]+\}", "", math_body)
    math_body = math_body.strip()
    if not math_body:
        return "\n\n"
    caption_block = render_native_caption_block("equation", label, "") if label else ""
    parts = [f"$$\n{math_body}\n$$"]
    if caption_block:
        parts.append(caption_block)
    content = "\n\n".join(parts)
    return f"\n\n{content}\n\n"


def replace_environment_blocks(text: str, env_name: str, replacer) -> str:
    begin_token = f"\\begin{{{env_name}}}"
    end_token = f"\\end{{{env_name}}}"
    parts: list[str] = []
    index = 0
    while True:
        start = text.find(begin_token, index)
        if start == -1:
            parts.append(text[index:])
            break
        end = text.find(end_token, start)
        if end == -1:
            parts.append(text[index:])
            break
        block_end = end + len(end_token)
        parts.append(text[index:start])
        parts.append(replacer(text[start:block_end]))
        index = block_end
    return "".join(parts)


def extract_first_environment_block(text: str, env_name: str) -> str | None:
    begin_token = f"\\begin{{{env_name}}}"
    end_token = f"\\end{{{env_name}}}"
    start = text.find(begin_token)
    if start == -1:
        return None
    end = text.find(end_token, start)
    if end == -1:
        return None
    return text[start : end + len(end_token)]


def extract_simple_command_arg(text: str, command: str) -> str | None:
    token = f"\\{command}"
    start = text.find(token)
    if start == -1:
        return None
    cursor = skip_whitespace(text, start + len(token))
    if cursor >= len(text) or text[cursor] != "{":
        return None
    value, _ = read_braced(text, cursor)
    return value


def extract_label(text: str) -> str | None:
    match = re.search(r"\\label\{([^}]+)\}", text)
    return match.group(1).strip() if match else None


def remove_block_command_from_text(text: str, command: str) -> str:
    """Remove all \\command{...} occurrences (with proper nested-brace matching)."""
    result: list[str] = []
    cursor = 0
    token = f"\\{command}"
    while cursor < len(text):
        idx = text.find(token, cursor)
        if idx == -1:
            result.append(text[cursor:])
            break
        result.append(text[cursor:idx])
        try:
            _, end = read_braced(text, idx + len(token))
            cursor = end
        except ValueError:
            result.append(text[idx : idx + len(token)])
            cursor = idx + len(token)
    return "".join(result)


def parse_table_rows(block: str, env_name: str) -> list[list[str]]:
    content = strip_environment_wrapper(block, env_name)
    content = remove_block_command_from_text(content, "caption")
    content = re.sub(r"\\label\{[^}]+\}", "", content)
    content = re.sub(r"\\endfirsthead.*?\\endhead", "", content, flags=re.S)
    content = re.sub(r"\\endfoot.*?\\endlastfoot", "", content, flags=re.S)
    for token in [
        "\\toprule", "\\midrule", "\\bottomrule", "\\hline", "\\addlinespace",
        "\\tabularnewline", "\\endlastfoot", "\\endhead", "\\endfirsthead", "\\endfoot",
    ]:
        content = content.replace(token, "\n")
    content = re.sub(r"\\noalign\{[^}]*\}", "", content)
    content = re.sub(r"\\cmidrule(?:\([^)]*\))?\{[^}]+\}", "\n", content)

    rows: list[list[str]] = []
    for raw_row in split_outside_braces(content, "\\\\"):
        row_text = raw_row.strip()
        if not row_text:
            continue
        cells = split_outside_braces(row_text, "&")
        expanded_cells: list[str] = []
        for cell in cells:
            expanded_cells.extend(expand_multicolumn_cell(cell.strip()))
        cleaned_cells = [cleanup_table_cell(cell) for cell in expanded_cells]
        while cleaned_cells and not cleaned_cells[-1]:
            cleaned_cells.pop()
        if any(cleaned_cells):
            rows.append(cleaned_cells)

    if len(rows) >= 2 and rows[0] == rows[1]:
        rows.pop(1)
    return rows


def strip_environment_wrapper(block: str, env_name: str) -> str:
    begin_token = f"\\begin{{{env_name}}}"
    cursor = block.find(begin_token)
    if cursor == -1:
        return block
    cursor += len(begin_token)
    # Skip optional [] argument
    cursor = skip_whitespace(block, cursor)
    if cursor < len(block) and block[cursor] == "[":
        depth = 1
        cursor += 1
        while cursor < len(block) and depth > 0:
            c = block[cursor]
            if c == "[":
                depth += 1
            elif c == "]":
                depth -= 1
            cursor += 1
    # Skip all required {} argument(s) using proper nested-brace matching
    cursor = skip_whitespace(block, cursor)
    while cursor < len(block) and block[cursor] == "{":
        try:
            _, cursor = read_braced(block, cursor)
        except ValueError:
            break
        cursor = skip_whitespace(block, cursor)
    content = block[cursor:]
    end_token = f"\\end{{{env_name}}}"
    if end_token in content:
        content = content.rsplit(end_token, 1)[0]
    return content


def split_outside_braces(text: str, separator: str) -> list[str]:
    parts: list[str] = []
    buffer: list[str] = []
    depth = 0
    index = 0
    while index < len(text):
        if text.startswith(separator, index) and depth == 0:
            parts.append("".join(buffer))
            buffer = []
            index += len(separator)
            continue
        char = text[index]
        if char == "{" and not is_escaped(text, index):
            depth += 1
        elif char == "}" and not is_escaped(text, index) and depth > 0:
            depth -= 1
        buffer.append(char)
        index += 1
    parts.append("".join(buffer))
    return parts


def expand_multicolumn_cell(cell: str) -> list[str]:
    stripped = cell.strip()
    if not stripped.startswith("\\multicolumn"):
        return [stripped]
    cursor = len("\\multicolumn")
    try:
        count_text, cursor = read_braced(stripped, skip_whitespace(stripped, cursor))
        _, cursor = read_braced(stripped, skip_whitespace(stripped, cursor))
        value, _ = read_braced(stripped, skip_whitespace(stripped, cursor))
    except ValueError:
        return [stripped]
    return [value.strip()]


def cleanup_table_cell(cell: str) -> str:
    # Strip minipage wrappers and alignment commands before inline TeX processing
    cell = re.sub(r"(?m)^\s*%\s*", "", cell)
    cell = re.sub(r"\\begin\{minipage\}(?:\[[^\]]*\])?\{[^}]*\}", "", cell)
    cell = re.sub(r"\\end\{minipage\}", "", cell)
    cell = re.sub(r"\\(raggedright|raggedleft|centering|arraybackslash)\b\s*", "", cell)
    cell = re.sub(r"\\noalign\{[^}]*\}", "", cell)
    text = convert_inline_tex_to_markdown(cell)
    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text).strip()
    text = text.replace("|", "\\|")
    return text


def render_markdown_table(rows: list[list[str]], caption: str, label: str | None, kind: str) -> str:
    if not rows:
        caption_block = render_native_caption_block(kind, label, caption)
        if not caption_block:
            return "\n\n"
        return f"\n\n{caption_block}\n\n"

    column_count = max(len(row) for row in rows)
    padded_rows = [row + [""] * (column_count - len(row)) for row in rows]
    header = padded_rows[0]
    body_rows = padded_rows[1:] or [[""] * column_count]

    lines = [format_markdown_row(header), format_markdown_row(["---"] * column_count)]
    lines.extend(format_markdown_row(row) for row in body_rows)

    table_text = "\n".join(lines)
    caption_block = render_native_caption_block(kind, label, caption)
    parts: list[str] = []
    if caption_block:
        parts.append(caption_block)
    parts.append(table_text)
    content = "\n\n".join(parts)
    return f"\n\n{content}\n\n"


def format_markdown_row(row: list[str]) -> str:
    return "| " + " | ".join(cell or " " for cell in row) + " |"


def format_caption_text(
    caption: str,
    label: str | None,
    labels: dict[str, str] | None,
    kind: str,
) -> str:
    caption_text = convert_inline_tex_to_plain(caption)
    if labels is None:
        labels = CURRENT_LABELS
    number = labels.get(label, "") if label else ""
    prefix = caption_prefix_for_kind(kind)
    if number and caption_text:
        return f"{format_numbered_reference(prefix, number)} {caption_text}".strip()
    if caption_text:
        return caption_text
    if number:
        return format_numbered_reference(prefix, number)
    return ""


def make_anchor_id(label: str) -> str:
    anchor = re.sub(r"[^0-9A-Za-z_-]+", "-", label).strip("-")
    return anchor or "anchor"


def make_bibliography_anchor_id(source_key: str) -> str:
    return make_anchor_id(f"bib-{source_key}")


def render_anchor_div(label: str | None, content: str) -> str:
    if not label:
        return content
    return f"::: {{#{make_anchor_id(label)}}}\n\n{content}\n\n:::"


def render_custom_style_block(text: str, style_name: str) -> str:
    if not text:
        return ""
    return f"::: {{custom-style=\"{style_name}\"}}\n{text}\n:::"


def render_image_attributes(anchor: str | None, options: str) -> str:
    attributes: list[str] = []
    if anchor:
        attributes.append(f"#{anchor}")
    width = parse_includegraphics_width(options)
    if width:
        attributes.append(f"width={width}")
    return f"{{{' '.join(attributes)}}}" if attributes else ""


def parse_includegraphics_width(options: str) -> str | None:
    if not options:
        return None
    match = re.search(r"width\s*=\s*([^,]+)", options)
    if match is None:
        return None
    raw_width = match.group(1).strip()
    ratio = parse_includegraphics_width_ratio(raw_width, CURRENT_LENGTH_CONTEXT)
    if ratio is not None:
        return format_width_percentage(ratio)
    return None


def parse_includegraphics_width_ratio(options_or_width: str, length_context: dict[str, float]) -> float | None:
    if not options_or_width:
        return None
    raw_width = options_or_width
    if "," in raw_width or "=" in raw_width:
        match = re.search(r"width\s*=\s*([^,]+)", raw_width)
        if match is None:
            return None
        raw_width = match.group(1).strip()
    return evaluate_width_ratio(raw_width, length_context)


def format_width_percentage(ratio: float) -> str:
    percentage = ratio * 100
    if abs(round(percentage) - percentage) < 0.05:
        return f"{round(percentage):.0f}%"
    return f"{percentage:.2f}%"


def convert_section_commands(body: str) -> str:
    patterns = [
        (r"\\subsubsection\*?\{([^{}]+)\}", r"\n\n### \1\n"),
        (r"\\subsection\*?\{([^{}]+)\}", r"\n\n## \1\n"),
        (r"\\section\*?\{([^{}]+)\}", r"\n\n# \1\n"),
    ]
    for pattern, replacement in patterns:
        body = re.sub(pattern, replacement, body)
    return body


def convert_block_commands(body: str) -> str:
    body = body.replace("\\begin{center}", "")
    body = body.replace("\\end{center}", "")
    body = re.sub(r"^\s*\\centering\s*$", "", body, flags=re.M)
    body = re.sub(r"^\s*\\label\{[^}]+\}\s*$", "", body, flags=re.M)
    body = body.replace("\\\\", "  \n")
    return body


def convert_inline_tex_to_markdown(text: str) -> str:
    previous = None
    current = normalize_inline_tex_artifacts(text)
    while current != previous:
        previous = current
        current = replace_command_two_args(
            current,
            "href",
            lambda url, label: f"[{convert_inline_tex_to_markdown(label)}]({url})",
        )
        current = replace_command_one_arg(current, "url", lambda value: f"<{value}>")
        current = replace_command_one_arg(current, "nolinkurl", lambda value: f"<{value}>")
        current = replace_command_one_arg(
            current,
            "emph",
            lambda value: f"*{convert_inline_tex_to_markdown(value)}*",
        )
        current = replace_command_one_arg(
            current,
            "textbf",
            lambda value: f"**{convert_inline_tex_to_markdown(value)}**",
        )
        current = replace_command_one_arg(
            current,
            "textsuperscript",
            lambda value: f"<sup>{convert_inline_tex_to_markdown(value)}</sup>",
        )
        current = replace_command_one_arg(
            current,
            "hl",
            lambda value: convert_inline_tex_to_markdown(value),
        )

    replacements = {
        "``": "\u201c",
        "''": "\u201d",
        "`": "\u2018",
        "\\%": "%",
        "\\_": "_",
        "\\&": "&",
        "\\#": "#",
        "\\linewidth": "",
        "\\and": "; ",
        "~": " ",
        "{}": "",
    }
    for old, new in replacements.items():
        current = current.replace(old, new)
    current = re.sub(r"\\label\{[^}]+\}", "", current)
    return current


def convert_inline_tex_to_plain(text: str) -> str:
    current = normalize_inline_tex_artifacts(text)
    previous = None
    while current != previous:
        previous = current
        current = replace_command_two_args(current, "href", lambda url, label: convert_inline_tex_to_plain(label))
        current = replace_command_one_arg(current, "url", lambda value: value)
        current = replace_command_one_arg(current, "nolinkurl", lambda value: value)
        current = replace_command_one_arg(current, "emph", lambda value: convert_inline_tex_to_plain(value))
        current = replace_command_one_arg(current, "textbf", lambda value: convert_inline_tex_to_plain(value))
        current = replace_command_one_arg(current, "textsuperscript", lambda value: convert_inline_tex_to_plain(value))
        current = replace_command_one_arg(current, "hl", lambda value: convert_inline_tex_to_plain(value))
    replacements = {
        "\\%": "%",
        "\\_": "_",
        "\\&": "&",
        "\\#": "#",
        "\\and": "; ",
        "~": " ",
        "{}": "",
    }
    for old, new in replacements.items():
        current = current.replace(old, new)
    current = re.sub(r"\\label\{[^}]+\}", "", current)
    current = re.sub(r"\s+", " ", current).strip()
    return current


def normalize_inline_tex_artifacts(text: str) -> str:
    current = text
    current = re.sub(r"\\(?:protect|phantomsection)\b", "", current)
    current = re.sub(r"\\label\{[^}]+\}\{\}", "", current)
    current = re.sub(r"\\tabularnewline\b", "", current)
    current = re.sub(r"\$([^$]+)\$", lambda match: simplify_tex_math_fragment(match.group(1)), current)
    replacements = {
        "{[]}": "[]",
        "{[}": "[",
        "{]}": "]",
        "{(}": "(",
        "{)}": ")",
    }
    for old, new in replacements.items():
        current = current.replace(old, new)
    return current


def simplify_tex_math_fragment(fragment: str) -> str:
    current = fragment
    command_map = {
        r"\Delta": "Δ",
        r"\alpha": "α",
        r"\beta": "β",
        r"\gamma": "γ",
        r"\mu": "μ",
        r"\sigma": "σ",
    }
    for old, new in command_map.items():
        current = current.replace(old, new)
    current = current.replace(r"\%", "%")
    current = current.replace(r"\_", "_")
    current = current.replace(r"\{", "{")
    current = current.replace(r"\}", "}")
    current = current.replace(r"\left", "")
    current = current.replace(r"\right", "")
    current = re.sub(r"\\text\{([^{}]+)\}", r"\1", current)
    current = re.sub(r"\\mathrm\{([^{}]+)\}", r"\1", current)
    current = re.sub(r"\\operatorname\{([^{}]+)\}", r"\1", current)
    current = current.replace("{", "")
    current = current.replace("}", "")
    current = re.sub(r"\\([A-Za-z]+)", r"\1", current)
    current = re.sub(r"\s+", " ", current).strip()
    return current


def cleanup_markdown(body: str) -> str:
    body = re.sub(r"^\s+$", "", body, flags=re.M)
    body = re.sub(r"\n{3,}", "\n\n", body)
    return body.strip() + "\n"


def compose_markdown(metadata: dict, body: str) -> str:
    front_matter = ["---"]
    title = metadata.get("title")
    if title:
        front_matter.append(f'title: "{escape_yaml(title)}"')
    authors = metadata.get("authors") or []
    if authors:
        front_matter.append("author:")
        for author in authors:
            front_matter.append(f'  - "{escape_yaml(author)}"')
    if metadata.get("date"):
        front_matter.append(f'date: "{escape_yaml(metadata["date"])}"')
    front_matter.append("---")
    return "\n".join(front_matter) + "\n\n" + body


def escape_yaml(value: str) -> str:
    return value.replace("\\", "\\\\").replace('"', '\\"')


def strip_layout_only_commands(body: str) -> str:
    replacements = {
        "\\begin{widetableblock}": "",
        "\\end{widetableblock}": "",
        "\\begin{landscape}": "",
        "\\end{landscape}": "",
        "\\begingroup": "",
        "\\endgroup": "",
        "\\protect": "",
        "\\phantomsection": "",
        "\\endfirsthead": "",
        "\\endhead": "",
        "\\endfoot": "",
        "\\endlastfoot": "",
        "\\widetablewidth": "\\linewidth",
    }
    for old, new in replacements.items():
        body = body.replace(old, new)

    cleaned_lines: list[str] = []
    for line in body.splitlines():
        stripped = line.strip()
        if not stripped:
            cleaned_lines.append("")
            continue
        if stripped in {"{\\small", "\\small", "{\\scriptsize", "\\scriptsize", "{\\footnotesize", "\\footnotesize", "}"}:
            continue
        if stripped.startswith("\\setlength{"):
            continue
        if stripped.startswith("\\hbadness") or stripped.startswith("\\hfuzz"):
            continue
        if stripped.startswith("\\Urlmuskip"):
            continue
        if stripped.startswith("\\captionsetup"):
            continue
        if stripped.startswith("\\renewcommand{"):
            continue
        cleaned_lines.append(line)

    return "\n".join(cleaned_lines)


def normalize_table_syntax(body: str) -> str:
    normalized_lines: list[str] = []
    inside_longtable = False
    for line in body.splitlines():
        stripped = line.strip()
        if stripped.startswith("\\begin{tabularx}"):
            normalized_lines.append(rewrite_begin_line(line, "tabularx"))
            continue
        if stripped.startswith("\\begin{tabular}{"):
            normalized_lines.append(rewrite_begin_line(line, "tabular"))
            continue
        if stripped.startswith("\\begin{longtable}"):
            inside_longtable = True
            # Column spec (possibly multi-line) is stripped by strip_environment_wrapper; pass through as-is
            normalized_lines.append(line)
            continue

        line = line.replace("\\end{tabularx}", "\\end{tabular}")
        line = re.sub(r"\\cmidrule(?:\([^)]*\))?\{[^}]+\}", r"\\hline", line)
        for old in ["\\toprule", "\\midrule", "\\bottomrule"]:
            line = line.replace(old, "\\hline")
        line = line.replace("\\addlinespace", "")
        if "\\caption{" in line:
            if inside_longtable:
                if not line.rstrip().endswith("\\"):
                    line = f"{line.rstrip()}\\\\"
            elif line.rstrip().endswith("\\"):
                line = re.sub(r"\\+\s*$", "", line)
        if stripped.startswith("\\end{longtable}"):
            inside_longtable = False
        normalized_lines.append(line)

    normalized = "\n".join(normalized_lines)
    normalized = re.sub(r"\n{3,}", "\n\n", normalized)
    return normalized


def rewrite_begin_line(line: str, env: str) -> str:
    indent = line[: len(line) - len(line.lstrip())]
    stripped = line.strip()
    if env == "tabularx":
        token = "\\begin{tabularx}"
        cursor = len(token)
        cursor = skip_whitespace(stripped, cursor)
        _, cursor = read_braced(stripped, cursor)
        cursor = skip_whitespace(stripped, cursor)
        spec, _ = read_braced(stripped, cursor)
        simple_spec = simplify_column_spec(spec)
        return f"{indent}\\begin{{tabular}}{{{simple_spec}}}"

    token = f"\\begin{{{env}}}"
    cursor = len(token)
    cursor = skip_whitespace(stripped, cursor)
    spec, _ = read_braced(stripped, cursor)
    simple_spec = simplify_column_spec(spec)
    return f"{indent}\\begin{{{env}}}{{{simple_spec}}}"


def simplify_column_spec(spec: str) -> str:
    columns: list[str] = []
    index = 0
    while index < len(spec):
        char = spec[index]
        if char.isspace() or char in "|!":
            index += 1
            continue
        if char in {">", "<"}:
            index += 1
            index = skip_whitespace(spec, index)
            if index < len(spec) and spec[index] == "{":
                _, index = read_braced(spec, index)
            continue
        if char == "@":
            index += 1
            index = skip_whitespace(spec, index)
            if index < len(spec) and spec[index] == "{":
                _, index = read_braced(spec, index)
            continue
        if char == "*":
            index += 1
            index = skip_whitespace(spec, index)
            count_text, index = read_braced(spec, index)
            index = skip_whitespace(spec, index)
            repeated_spec, index = read_braced(spec, index)
            repeated_columns = simplify_column_spec(repeated_spec)
            try:
                repeat_count = int(count_text)
            except ValueError:
                repeat_count = 1
            columns.extend(list(repeated_columns) * repeat_count)
            continue
        if char in "lcrX":
            columns.append("l" if char == "X" else char)
            index += 1
            continue
        if char in "pmb":
            index += 1
            index = skip_whitespace(spec, index)
            if index < len(spec) and spec[index] == "{":
                _, index = read_braced(spec, index)
            columns.append("l")
            continue
        if char == "{":
            _, index = read_braced(spec, index)
            continue
        index += 1

    return "".join(columns) or "l"


CURRENT_LABELS: dict[str, str] = {}


def infer_bibliography_path(source_tex: Path, source_text: str, override: Path | None = None) -> Path | None:
    if override is not None:
        return override.expanduser().resolve()

    for relative_path in re.findall(r"\\input\{([^}]+)\}", extract_document_body(source_text)):
        candidate = (source_tex.parent / relative_path).resolve()
        if not candidate.exists():
            continue
        try:
            if parse_bibliography_entries(candidate):
                return candidate
        except OSError:
            continue

    fallbacks = [
        (source_tex.parent / "bibliography_links.tex").resolve(),
        (source_tex.parent / "refs.bib").resolve(),
    ]
    first_existing: Path | None = None
    for fallback in fallbacks:
        if not fallback.exists():
            continue
        if first_existing is None:
            first_existing = fallback
        try:
            if parse_bibliography_entries(fallback):
                return fallback
        except OSError:
            continue
    return first_existing


def build_zotero_docx_context(
    source_tex: Path,
    template_hints: TemplateDocxHints,
    bibliography_path: Path | None = None,
    source_text: str | None = None,
    enable_zotero: bool = False,
    zotero_database: Path | None = None,
) -> ZoteroDocxContext:
    if source_text is None:
        source_text = source_tex.read_text(encoding="utf-8")
    bibliography_path = infer_bibliography_path(source_tex, source_text, bibliography_path)
    companion_payload = load_direct_zotero_companion_payload(source_tex)
    direct_targets = load_direct_zotero_targets(source_tex, source_text or "", companion_payload)
    citation_field_shells = load_direct_citation_field_shells(companion_payload)
    if bibliography_path is None or not bibliography_path.exists():
        if direct_targets is not None:
            return build_direct_zotero_context(direct_targets, citation_field_shells)
        return ZoteroDocxContext([], [], {}, {}, {})
    if direct_targets is not None:
        return build_direct_zotero_context(direct_targets, citation_field_shells)
    if enable_zotero and zotero_database is not None and not zotero_database.exists():
        raise FileNotFoundError(f"Zotero database not found: {zotero_database}")
    if enable_zotero and zotero_database is None and not DEFAULT_ZOTERO_DATABASE.exists():
        raise FileNotFoundError(f"Zotero mode requires a local Zotero database at {DEFAULT_ZOTERO_DATABASE}")
    if enable_zotero and zotero_database is None and DEFAULT_ZOTERO_DATABASE.exists():
        with copied_zotero_database(DEFAULT_ZOTERO_DATABASE) as database_snapshot:
            return build_zotero_docx_context(
                source_tex,
                template_hints,
                bibliography_path=bibliography_path,
                source_text=source_text,
                enable_zotero=enable_zotero,
                zotero_database=database_snapshot,
            )

    bibliography_entries = parse_bibliography_entries(bibliography_path)
    if not bibliography_entries:
        refs_bib = (source_tex.parent / "refs.bib").resolve()
        if refs_bib.exists() and refs_bib != bibliography_path:
            bibliography_entries = parse_bibliography_entries(refs_bib)
            if bibliography_entries:
                bibliography_path = refs_bib
    records_by_source: dict[str, object] = {}
    csl_by_key: dict[str, dict] = {}
    unmatched_notices: list[UnmatchedZoteroNotice] = []
    zotero_database = zotero_database or DEFAULT_ZOTERO_DATABASE
    if enable_zotero and zotero_database.exists():
        try:
            report, csl_items = resolve_bibliography_against_zotero(bibliography_path, zotero_database)
            records_by_source = {record.source_key: record for record in report.records}
            csl_by_key = {str(item.get("id")): item for item in csl_items}
            unmatched_notices = [
                UnmatchedZoteroNotice(
                    source_key=record.source_key,
                    formatted_reference=record.formatted_reference,
                    import_url=derive_import_url(record.source_key),
                )
                for record in report.records
                if not getattr(record, "matched", False)
            ]
        except Exception:
            records_by_source = {}
            csl_by_key = {}
            unmatched_notices = []

    resolved_entries: list[CitationTarget] = []
    by_anchor: dict[str, CitationTarget] = {}
    by_url: dict[str, CitationTarget] = {}
    by_doi: dict[str, CitationTarget] = {}
    anchor_map = build_bibliography_anchor_map(source_tex, source_text)
    for index, entry in enumerate(bibliography_entries, start=1):
        record = records_by_source.get(entry.source_key)
        zotero_item_key = getattr(record, "zotero_item_key", None)
        zotero_item_id = getattr(record, "zotero_item_id", None)
        item_data = synthesize_citation_item_data(entry, record, csl_by_key, index)
        if zotero_item_id is not None:
            item_data["id"] = zotero_item_id
        uri = getattr(record, "zotero_uri", None)
        if uri is None and zotero_item_key and template_hints.zotero_item_uri_prefix:
            uri = f"{template_hints.zotero_item_uri_prefix}{zotero_item_key}"
        anchor_id = anchor_map.get(entry.source_key, make_bibliography_anchor_id(entry.source_key))

        target = CitationTarget(
            source_key=entry.source_key,
            formatted_reference=entry.formatted_reference,
            zotero_item_key=zotero_item_key,
            item_data=item_data,
            uri=uri,
            anchor_id=anchor_id,
        )
        resolved_entries.append(target)
        by_anchor[anchor_id] = target

        normalized_source_url = normalize_url(entry.source_key)
        if normalized_source_url:
            by_url[normalized_source_url] = target
        normalized_source_doi = normalize_doi(entry.source_key)
        if normalized_source_doi:
            by_doi[normalized_source_doi] = target

    return ZoteroDocxContext(resolved_entries, unmatched_notices, by_anchor, by_url, by_doi)


def load_direct_zotero_companion_payload(source_tex: Path) -> dict | None:
    companion_path = source_tex.parent / DIRECT_ZOTERO_COMPANION_FILENAME
    if not companion_path.exists():
        return None
    try:
        payload = json.loads(companion_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    return payload if isinstance(payload, dict) else None


def load_direct_zotero_targets(
    source_tex: Path,
    source_text: str,
    companion_payload: dict | None = None,
) -> list[CitationTarget] | None:
    payload = companion_payload or load_direct_zotero_companion_payload(source_tex)
    if payload is None:
        return None
    raw_items = payload.get("items") if isinstance(payload, dict) else None
    if not isinstance(raw_items, list) or not raw_items:
        return None

    anchor_map = build_bibliography_anchor_map(source_tex, source_text)
    direct_targets: list[CitationTarget] = []
    for raw_item in raw_items:
        if not isinstance(raw_item, dict):
            continue
        source_key = str(raw_item.get("source_key") or raw_item.get("key") or "").strip()
        if not source_key:
            continue
        formatted_reference = str(raw_item.get("formatted_reference") or source_key)
        item_data = raw_item.get("item_data") if isinstance(raw_item.get("item_data"), dict) else {}
        zotero_item_key = raw_item.get("zotero_item_key")
        uri = raw_item.get("uri")
        anchor_id = anchor_map.get(source_key, make_bibliography_anchor_id(source_key))
        direct_targets.append(
            CitationTarget(
                source_key=source_key,
                formatted_reference=formatted_reference,
                zotero_item_key=str(zotero_item_key) if zotero_item_key else None,
                item_data=dict(item_data),
                uri=str(uri) if uri else None,
                anchor_id=anchor_id,
            )
        )
    return direct_targets or None


def load_direct_citation_field_shells(
    companion_payload: dict | None,
) -> dict[tuple[tuple[str, ...], str], list[CitationFieldShell]]:
    if not isinstance(companion_payload, dict):
        return {}
    raw_citations = companion_payload.get("citations")
    if not isinstance(raw_citations, list):
        return {}

    citation_field_shells: dict[tuple[tuple[str, ...], str], list[CitationFieldShell]] = {}
    for raw_citation in raw_citations:
        if not isinstance(raw_citation, dict):
            continue
        raw_keys = raw_citation.get("source_keys")
        raw_field_nodes = raw_citation.get("field_nodes_xml")
        if not isinstance(raw_keys, list) or not isinstance(raw_field_nodes, list):
            continue
        source_keys = [str(key).strip() for key in raw_keys if str(key).strip()]
        field_nodes_xml = [str(node) for node in raw_field_nodes if isinstance(node, str) and node.strip()]
        formatted_citation = normalize_citation_display_text(str(raw_citation.get("formatted_citation") or ""))
        if not source_keys or not field_nodes_xml or not formatted_citation:
            continue
        signature = make_citation_field_shell_signature(source_keys, formatted_citation)
        citation_field_shells.setdefault(signature, []).append(CitationFieldShell(field_nodes_xml=field_nodes_xml))
    return citation_field_shells


def build_direct_zotero_context(
    entries: list[CitationTarget],
    citation_field_shells: dict[tuple[tuple[str, ...], str], list[CitationFieldShell]] | None = None,
) -> ZoteroDocxContext:
    by_anchor = {entry.anchor_id: entry for entry in entries}
    by_url = {
        normalized_url: entry
        for entry in entries
        if (normalized_url := normalize_url(entry.source_key))
    }
    by_doi = {
        normalized_doi: entry
        for entry in entries
        if (normalized_doi := normalize_doi(entry.source_key))
    }
    return ZoteroDocxContext(entries, [], by_anchor, by_url, by_doi, citation_field_shells or {})


def derive_import_url(source_key: str) -> str | None:
    normalized_doi = normalize_doi(source_key)
    if normalized_doi:
        return f"https://doi.org/{normalized_doi}"
    return normalize_url(source_key)


def synthesize_citation_item_data(entry, report_record: object | None, csl_by_key: dict[str, dict], synthetic_index: int) -> dict:
    zotero_item_key = getattr(report_record, "zotero_item_key", None)
    if getattr(report_record, "matched", False) and zotero_item_key in csl_by_key:
        item_data = dict(csl_by_key[zotero_item_key])
        item_data["id"] = getattr(report_record, "zotero_item_id", None) or zotero_item_key
        if getattr(report_record, "zotero_doi", None) and "DOI" not in item_data:
            item_data["DOI"] = report_record.zotero_doi
        if getattr(report_record, "zotero_url", None) and "URL" not in item_data:
            item_data["URL"] = report_record.zotero_url
        return item_data

    item_data = {
        "id": f"generated-{synthetic_index}",
        "type": "article-journal" if normalize_doi(entry.source_key) else "webpage",
        "title": entry.parsed_title or entry.formatted_reference,
    }
    normalized_url = normalize_url(entry.source_key)
    normalized_doi = normalize_doi(entry.source_key)
    if normalized_url:
        item_data["URL"] = normalized_url
    if normalized_doi:
        item_data["DOI"] = normalized_doi
    year_match = re.search(r"(19|20)\d{2}", entry.formatted_reference)
    if year_match:
        item_data["issued"] = {"date-parts": [[int(year_match.group(0))]]}
    return item_data


def postprocess_generated_docx(
    output_docx: Path,
    template_docx: Path,
    template_hints: TemplateDocxHints,
    zotero_context: ZoteroDocxContext,
    document_layout_hints: DocumentLayoutHints,
    diagnostics: ConversionDiagnostics,
    bibliography_heading: str = "参考文献",
    enable_zotero: bool = False,
    use_native_bookmarks: bool = True,
) -> None:
    with ZipFile(output_docx) as source_zip:
        archive_entries = [(info, source_zip.read(info.filename)) for info in source_zip.infolist()]
    archive_part_map = {info.filename: data for info, data in archive_entries}

    document_xml = next((data for info, data in archive_entries if info.filename == "word/document.xml"), None)
    if document_xml is None:
        return

    rels_xml = next((data for info, data in archive_entries if info.filename == "word/_rels/document.xml.rels"), None)
    relationship_targets = parse_document_relationships(rels_xml)

    document_tree = ET.fromstring(document_xml)
    changed = False
    changed |= apply_table_hints(document_tree, template_hints, document_layout_hints)
    changed |= apply_figure_hints(document_tree, document_layout_hints)
    changed |= apply_body_paragraph_hints(document_tree, template_hints, bibliography_heading)
    changed |= apply_bibliography_hints(
        document_tree,
        template_hints,
        zotero_context,
        bibliography_heading=bibliography_heading,
        enable_zotero=enable_zotero,
    )
    if enable_zotero:
        changed |= convert_citation_hyperlinks_to_zotero_fields(document_tree, relationship_targets, zotero_context)
    changed |= apply_native_cross_reference_fields(
        document_tree,
        template_hints,
        relationship_targets,
        diagnostics,
        use_native_bookmarks=use_native_bookmarks,
    )
    changed |= normalize_document_style_usage(document_tree, template_hints, bibliography_heading)
    if enable_zotero:
        changed |= strip_internal_hyperlink_styles(document_tree, relationship_targets)
        preserved_prefixes = (NATIVE_CROSS_REFERENCE_BOOKMARK_PREFIX,) if use_native_bookmarks else ()
        changed |= strip_all_bookmarks(document_tree, preserved_prefixes=preserved_prefixes)
    else:
        changed |= apply_default_citation_anchor_policy(
            document_tree,
            relationship_targets,
            zotero_context,
            bibliography_heading,
        )

    updated_parts: dict[str, bytes] = {
        "word/document.xml": ET.tostring(document_tree, encoding="utf-8", xml_declaration=True)
    }
    removed_parts: set[str] = set()

    if enable_zotero:
        zotero_preferences_part = load_template_zotero_preferences_part(template_docx)
        if zotero_preferences_part is not None and archive_part_map.get("docProps/custom.xml") != zotero_preferences_part:
            updated_parts["docProps/custom.xml"] = zotero_preferences_part
            changed = True
        custom_xml_updates = load_template_custom_xml_updates(
            template_docx,
            archive_part_map,
            updated_parts,
        )
        if custom_xml_updates:
            updated_parts.update(custom_xml_updates)
            changed = True

    theme_xml = next((data for info, data in archive_entries if info.filename == "word/theme/theme1.xml"), None)
    if theme_xml is not None:
        sanitized_theme_xml, theme_changed = sanitize_theme_part(theme_xml)
        if theme_changed:
            updated_parts["word/theme/theme1.xml"] = sanitized_theme_xml
            changed = True

    relationship_updates, relationship_removals, relationships_changed = prune_unused_hyperlink_relationship_parts(
        archive_part_map,
        updated_parts,
    )
    if relationships_changed:
        updated_parts.update(relationship_updates)
        removed_parts.update(relationship_removals)
        changed = True

    for info, data in archive_entries:
        if info.filename in updated_parts:
            continue
        if not info.filename.startswith("word/") or not info.filename.endswith(".xml"):
            continue
        if info.filename == "word/styles.xml":
            continue

    if not changed:
        return

    temp_output = output_docx.with_suffix(".tmp.docx")
    try:
        with ZipFile(temp_output, "w", compression=ZIP_DEFLATED) as target_zip:
            written_parts: set[str] = set()
            for info, data in archive_entries:
                if info.filename in removed_parts:
                    continue
                if info.filename in updated_parts:
                    target_zip.writestr(info, updated_parts[info.filename])
                    written_parts.add(info.filename)
                    continue
                target_zip.writestr(info, data)
                written_parts.add(info.filename)
            for part_name, data in updated_parts.items():
                if part_name in written_parts or part_name in removed_parts:
                    continue
                target_zip.writestr(part_name, data)
        temp_output.replace(output_docx)
    finally:
        if temp_output.exists():
            temp_output.unlink()


def load_template_zotero_preferences_part(template_docx: Path) -> bytes | None:
    try:
        with ZipFile(template_docx) as template_zip:
            custom_part = template_zip.read("docProps/custom.xml")
    except (FileNotFoundError, KeyError, OSError):
        return None
    return custom_part if b"ZOTERO_PREF" in custom_part else None


def load_template_custom_xml_updates(
    template_docx: Path,
    archive_part_map: dict[str, bytes],
    updated_parts: dict[str, bytes],
) -> dict[str, bytes]:
    try:
        with ZipFile(template_docx) as template_zip:
            template_part_map = {name: template_zip.read(name) for name in template_zip.namelist()}
    except (FileNotFoundError, OSError):
        return {}

    custom_part_names = [name for name in template_part_map if name.startswith("customXml/")]
    if not custom_part_names:
        return {}

    updates: dict[str, bytes] = {}
    for part_name in custom_part_names:
        part_data = template_part_map[part_name]
        current_data = updated_parts.get(part_name, archive_part_map.get(part_name))
        if current_data != part_data:
            updates[part_name] = part_data

    merged_document_rels = merge_template_custom_xml_relationships(
        updated_parts.get("word/_rels/document.xml.rels", archive_part_map.get("word/_rels/document.xml.rels")),
        template_part_map.get("word/_rels/document.xml.rels"),
    )
    if merged_document_rels is not None:
        updates["word/_rels/document.xml.rels"] = merged_document_rels

    merged_content_types = merge_template_custom_xml_content_types(
        updated_parts.get("[Content_Types].xml", archive_part_map.get("[Content_Types].xml")),
        template_part_map.get("[Content_Types].xml"),
    )
    if merged_content_types is not None:
        updates["[Content_Types].xml"] = merged_content_types

    return updates


def merge_template_custom_xml_relationships(
    current_rels_xml: bytes | None,
    template_rels_xml: bytes | None,
) -> bytes | None:
    if current_rels_xml is None or template_rels_xml is None:
        return None

    current_root = ET.fromstring(current_rels_xml)
    template_root = ET.fromstring(template_rels_xml)
    changed = False
    existing_pairs = {
        (relationship.get("Type"), relationship.get("Target"))
        for relationship in current_root.findall(f"{{{PACKAGE_RELATIONSHIP_NAMESPACE}}}Relationship")
    }
    next_id = next_relationship_id(current_root)
    for relationship in template_root.findall(f"{{{PACKAGE_RELATIONSHIP_NAMESPACE}}}Relationship"):
        relationship_type = relationship.get("Type")
        target = relationship.get("Target")
        if relationship_type != "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml":
            continue
        pair = (relationship_type, target)
        if pair in existing_pairs:
            continue
        new_relationship = ET.Element(f"{{{PACKAGE_RELATIONSHIP_NAMESPACE}}}Relationship")
        new_relationship.set("Id", next_id)
        next_id = increment_relationship_id(next_id)
        if relationship_type is not None:
            new_relationship.set("Type", relationship_type)
        if target is not None:
            new_relationship.set("Target", target)
        current_root.append(new_relationship)
        existing_pairs.add(pair)
        changed = True
    if not changed:
        return None
    return ET.tostring(current_root, encoding="utf-8", xml_declaration=True)


def merge_template_custom_xml_content_types(
    current_content_types_xml: bytes | None,
    template_content_types_xml: bytes | None,
) -> bytes | None:
    if current_content_types_xml is None or template_content_types_xml is None:
        return None

    current_root = ET.fromstring(current_content_types_xml)
    template_root = ET.fromstring(template_content_types_xml)
    changed = False
    existing_part_names = {
        override.get("PartName")
        for override in current_root.findall(f"{{{CONTENT_TYPES_NAMESPACE}}}Override")
    }
    for override in template_root.findall(f"{{{CONTENT_TYPES_NAMESPACE}}}Override"):
        part_name = override.get("PartName")
        if not isinstance(part_name, str) or not part_name.startswith("/customXml/"):
            continue
        if part_name in existing_part_names:
            continue
        current_root.append(ET.fromstring(ET.tostring(override, encoding="utf-8")))
        existing_part_names.add(part_name)
        changed = True
    if not changed:
        return None
    return ET.tostring(current_root, encoding="utf-8", xml_declaration=True)


def next_relationship_id(root: ET.Element) -> str:
    numeric_ids = []
    for relationship in root.findall(f"{{{PACKAGE_RELATIONSHIP_NAMESPACE}}}Relationship"):
        relationship_id = relationship.get("Id") or ""
        match = re.fullmatch(r"rId(\d+)", relationship_id)
        if match is not None:
            numeric_ids.append(int(match.group(1)))
    return f"rId{max(numeric_ids, default=0) + 1}"


def increment_relationship_id(relationship_id: str) -> str:
    match = re.fullmatch(r"rId(\d+)", relationship_id)
    if match is None:
        return relationship_id
    return f"rId{int(match.group(1)) + 1}"


def apply_native_cross_reference_fields(
    document_tree: ET.Element,
    template_hints: TemplateDocxHints,
    relationship_targets: dict[str, str],
    diagnostics: ConversionDiagnostics,
    use_native_bookmarks: bool = True,
) -> bool:
    changed, caption_targets = replace_caption_placeholders(
        document_tree,
        template_hints,
        diagnostics,
        use_native_bookmarks=use_native_bookmarks,
    )
    changed |= rewrite_cross_reference_hyperlinks(
        document_tree,
        relationship_targets,
        caption_targets,
        diagnostics,
        use_field_refs=use_native_bookmarks,
    )
    for target in diagnostics.missing_cross_reference_targets():
        diagnostics.add_warning(
            f"{describe_cross_reference_target(target)} 在文内不含交叉引用，可能导致编辑审核不通过"
        )
    return changed


def allowed_paragraph_style_ids(hints: TemplateDocxHints) -> set[str]:
    return {
        style_id
        for style_id in (
            hints.normal_style_id,
            hints.title_style_id,
            hints.heading_1_style_id,
            hints.heading_2_style_id,
            hints.heading_3_style_id,
            hints.table_paragraph_style_id,
            hints.caption_style_id,
            hints.bibliography_style_id,
        )
        if style_id
    }


def normalize_document_style_usage(
    document_tree: ET.Element,
    hints: TemplateDocxHints,
    bibliography_heading: str = "参考文献",
) -> bool:
    changed = False
    allowed_styles = allowed_paragraph_style_ids(hints)
    bibliography_index = find_bibliography_body_index(document_tree, bibliography_heading)
    table_paragraph_ids = {
        id(paragraph)
        for table in document_tree.findall(".//w:tbl", XML_NAMESPACES)
        for paragraph in table.findall(".//w:p", XML_NAMESPACES)
    }

    for paragraph in document_tree.findall(".//w:p", XML_NAMESPACES):
        style_id = get_paragraph_style_id(paragraph)
        body_index = find_body_child_index_containing(document_tree, paragraph)
        target_style = None
        paragraph_text = get_paragraph_text(paragraph)

        if id(paragraph) in table_paragraph_ids:
            target_style = hints.table_paragraph_style_id or hints.normal_style_id or style_id
        elif bibliography_index is not None and body_index is not None and body_index > bibliography_index:
            target_style = hints.bibliography_style_id or hints.normal_style_id or style_id
        elif paragraph_text == bibliography_heading and hints.heading_1_style_id:
            target_style = hints.heading_1_style_id
        elif style_id not in allowed_styles:
            target_style = hints.normal_style_id or style_id

        if target_style and target_style != style_id:
            set_paragraph_style_id(paragraph, target_style)
            changed = True

    if strip_run_styles_in_tree(document_tree):
        changed = True
    if normalize_tree_run_fonts(document_tree):
        changed = True
    if strip_zotero_field_run_fonts(document_tree):
        changed = True
    return changed


def strip_run_styles_in_tree(root: ET.Element) -> bool:
    changed = False
    for run_properties in root.findall(".//w:rPr", XML_NAMESPACES):
        if remove_run_style(run_properties):
            changed = True
    return changed


def normalize_tree_run_fonts(root: ET.Element) -> bool:
    changed = False
    for run_properties in root.findall(".//w:rPr", XML_NAMESPACES):
        if ensure_times_new_roman_rfonts(run_properties):
            changed = True
    return changed


def strip_zotero_field_run_fonts(root: ET.Element) -> bool:
    changed = False
    for paragraph in root.findall(".//w:p", XML_NAMESPACES):
        in_field = False
        is_zotero_field = False
        field_children: list[ET.Element] = []
        for child in list(paragraph):
            if not in_field:
                if child.tag != f"{WORD_ATTR_PREFIX}r":
                    continue
                field_char = child.find("w:fldChar", XML_NAMESPACES)
                if field_char is None or field_char.get(f"{WORD_ATTR_PREFIX}fldCharType") != "begin":
                    continue
                in_field = True
                is_zotero_field = False
                field_children = [child]
                continue

            field_children.append(child)
            if child.tag != f"{WORD_ATTR_PREFIX}r":
                continue
            instruction_text = "".join(text.text or "" for text in child.findall("w:instrText", XML_NAMESPACES))
            if "ADDIN ZOTERO_ITEM CSL_CITATION" in instruction_text:
                is_zotero_field = True
            field_char = child.find("w:fldChar", XML_NAMESPACES)
            if field_char is None or field_char.get(f"{WORD_ATTR_PREFIX}fldCharType") != "end":
                continue
            if is_zotero_field:
                for field_child in field_children:
                    for run_properties in field_child.findall(".//w:rPr", XML_NAMESPACES):
                        for fonts in run_properties.findall("w:rFonts", XML_NAMESPACES):
                            run_properties.remove(fonts)
                            changed = True
            in_field = False
            is_zotero_field = False
            field_children = []
    return changed


def ensure_times_new_roman_rfonts(run_properties: ET.Element) -> bool:
    changed = False
    fonts = run_properties.find("w:rFonts", XML_NAMESPACES)
    if fonts is None:
        fonts = ET.Element(f"{WORD_ATTR_PREFIX}rFonts")
        run_properties.insert(0, fonts)
        changed = True
    for attr_name in ("ascii", "hAnsi", "cs"):
        qualified = f"{WORD_ATTR_PREFIX}{attr_name}"
        if fonts.get(qualified) != WESTERN_FONT_FAMILY:
            fonts.set(qualified, WESTERN_FONT_FAMILY)
            changed = True
    for theme_attr in ("asciiTheme", "hAnsiTheme", "cstheme", "csTheme"):
        qualified = f"{WORD_ATTR_PREFIX}{theme_attr}"
        if qualified in fonts.attrib:
            del fonts.attrib[qualified]
            changed = True
    return changed


def replace_caption_placeholders(
    document_tree: ET.Element,
    template_hints: TemplateDocxHints,
    diagnostics: ConversionDiagnostics,
    use_native_bookmarks: bool = True,
) -> tuple[bool, dict[str, CrossReferenceTarget]]:
    changed = False
    caption_targets: dict[str, CrossReferenceTarget] = {}
    next_bookmark_id = allocate_bookmark_id(document_tree) if use_native_bookmarks else 0
    fallback_sequence_numbers = {"figure": 0, "table": 0, "equation": 0}

    for paragraph in document_tree.findall(".//w:p", XML_NAMESPACES):
        placeholder = parse_caption_placeholder(get_paragraph_text(paragraph))
        if placeholder is None:
            continue

        kind = placeholder.get("kind", "") or "figure"
        label = placeholder.get("label", "")
        caption_text = placeholder.get("caption", "")
        number = placeholder.get("number", "")
        fallback_sequence_numbers[kind] = fallback_sequence_numbers.get(kind, 0) + 1
        display_number = number or str(fallback_sequence_numbers[kind])
        bookmark_name: str | None = None
        bookmark_id: str | None = None
        if use_native_bookmarks:
            bookmark_name = make_native_cross_reference_bookmark(label or f"{kind}-{next_bookmark_id}")
            bookmark_id = str(next_bookmark_id)
            next_bookmark_id += 1

        clear_paragraph_content(paragraph)
        if template_hints.caption_style_id:
            set_paragraph_style_id(paragraph, template_hints.caption_style_id)
        reference_rpr = first_run_properties(paragraph)
        new_children: list[ET.Element] = []
        if bookmark_id is not None and bookmark_name is not None:
            new_children.append(build_bookmark_boundary(bookmark_id, bookmark_name, is_start=True))
        new_children.append(build_field_run(text=caption_prefix_for_kind(kind) + " ", rpr_template=reference_rpr))
        new_children.extend(build_sequence_field_elements(kind, display_number, reference_rpr))
        if bookmark_id is not None:
            new_children.append(build_bookmark_boundary(bookmark_id, None, is_start=False))
        if caption_text:
            new_children.append(build_field_run(text=f" {caption_text}", rpr_template=reference_rpr))
        append_paragraph_children(paragraph, new_children)

        target = CrossReferenceTarget(
            kind=kind,
            label=label or bookmark_name,
            number=display_number,
            caption_text=caption_text,
            bookmark_name=bookmark_name,
        )
        diagnostics.cross_reference_targets.append(target)
        if label:
            caption_targets[make_cross_reference_anchor(label)] = target
        else:
            diagnostics.add_warning(
                f"{describe_cross_reference_target(target)} 未定义标签，无法生成稳定交叉引用，可能导致编辑审核不通过"
            )
        changed = True

    return changed, caption_targets


def rewrite_cross_reference_hyperlinks(
    document_tree: ET.Element,
    relationship_targets: dict[str, str],
    caption_targets: dict[str, CrossReferenceTarget],
    diagnostics: ConversionDiagnostics,
    use_field_refs: bool = True,
) -> bool:
    changed = False
    for paragraph in document_tree.findall(".//w:p", XML_NAMESPACES):
        for child in list(paragraph):
            if child.tag != f"{WORD_ATTR_PREFIX}hyperlink":
                continue
            target_anchor = extract_hyperlink_target_anchor(child, relationship_targets)
            if target_anchor is None or not target_anchor.startswith(CROSS_REFERENCE_ANCHOR_PREFIX):
                continue
            target = caption_targets.get(target_anchor)
            display_text = get_element_text(child)
            if target is None:
                diagnostics.add_warning(f"交叉引用 {display_text} 未能解析到对应题注，可能导致编辑审核不通过")
                flatten_hyperlink_child(paragraph, child, reference_style=True)
                changed = True
                continue
            if not use_field_refs or target.bookmark_name is None:
                flatten_hyperlink_child(paragraph, child, reference_style=True)
                target.referenced = True
                changed = True
                continue
            insert_at = list(paragraph).index(child)
            paragraph.remove(child)
            for offset, field_run in enumerate(
                build_reference_field_elements(display_text, target.bookmark_name, child)
            ):
                paragraph.insert(insert_at + offset, field_run)
            target.referenced = True
            changed = True
    return changed


def describe_cross_reference_target(target: CrossReferenceTarget) -> str:
    numbered = format_numbered_reference(caption_prefix_for_kind(target.kind), target.number)
    if target.caption_text:
        return f"{numbered} {target.caption_text}".strip()
    return numbered


def allocate_bookmark_id(document_tree: ET.Element) -> int:
    bookmark_ids: list[int] = []
    for bookmark in document_tree.findall(".//w:bookmarkStart", XML_NAMESPACES):
        raw_id = bookmark.get(f"{WORD_ATTR_PREFIX}id")
        if raw_id is None:
            continue
        try:
            bookmark_ids.append(int(raw_id))
        except ValueError:
            continue
    return max(bookmark_ids, default=-1) + 1


def clear_paragraph_content(paragraph: ET.Element) -> None:
    for child in list(paragraph):
        if child.tag == f"{WORD_ATTR_PREFIX}pPr":
            continue
        paragraph.remove(child)


def append_paragraph_children(paragraph: ET.Element, children: list[ET.Element]) -> None:
    insert_at = 1 if list(paragraph) and list(paragraph)[0].tag == f"{WORD_ATTR_PREFIX}pPr" else 0
    for offset, child in enumerate(children):
        paragraph.insert(insert_at + offset, child)


def build_bookmark_boundary(bookmark_id: str, bookmark_name: str | None, is_start: bool) -> ET.Element:
    tag_name = "bookmarkStart" if is_start else "bookmarkEnd"
    bookmark = ET.Element(f"{WORD_ATTR_PREFIX}{tag_name}")
    bookmark.set(f"{WORD_ATTR_PREFIX}id", bookmark_id)
    if is_start and bookmark_name is not None:
        bookmark.set(f"{WORD_ATTR_PREFIX}name", bookmark_name)
    return bookmark


def build_sequence_field_elements(kind: str, display_number: str, rpr_template: ET.Element | None) -> list[ET.Element]:
    instruction = f" SEQ {sequence_identifier_for_kind(kind)} \\* ARABIC "
    return [
        build_field_run(fld_char_type="begin", rpr_template=rpr_template),
        build_field_run(instr_text=instruction, rpr_template=rpr_template),
        build_field_run(fld_char_type="separate", rpr_template=rpr_template),
        build_field_run(text=display_number, rpr_template=rpr_template),
        build_field_run(fld_char_type="end", rpr_template=rpr_template),
    ]


def build_reference_field_elements(
    display_text: str,
    bookmark_name: str,
    hyperlink_element: ET.Element,
) -> list[ET.Element]:
    reference_rpr = ensure_reference_run_properties(first_run_properties(hyperlink_element))
    instruction = f" REF {bookmark_name} \\h "
    return [
        build_field_run(fld_char_type="begin", rpr_template=reference_rpr),
        build_field_run(instr_text=instruction, rpr_template=reference_rpr),
        build_field_run(fld_char_type="separate", rpr_template=reference_rpr),
        build_field_run(text=display_text, rpr_template=reference_rpr),
        build_field_run(fld_char_type="end", rpr_template=reference_rpr),
    ]


def extract_hyperlink_target_anchor(
    hyperlink: ET.Element,
    relationship_targets: dict[str, str],
) -> str | None:
    anchor = hyperlink.get(f"{WORD_ATTR_PREFIX}anchor")
    if anchor:
        return anchor
    rel_id = hyperlink.get(f"{REL_ATTR_PREFIX}id")
    if rel_id is None:
        return None
    return extract_anchor_from_relationship_target(relationship_targets.get(rel_id))


def infer_template_docx_hints(template_docx: Path) -> TemplateDocxHints:
    with ZipFile(template_docx) as template_zip:
        styles_tree = ET.fromstring(template_zip.read("word/styles.xml"))
        document_tree = ET.fromstring(template_zip.read("word/document.xml"))
        raw_document_xml = template_zip.read("word/document.xml").decode("utf-8", errors="ignore")

    caption_style_id: str | None = None
    normal_style_id: str | None = None
    title_style_id: str | None = None
    heading_1_style_id: str | None = None
    heading_2_style_id: str | None = None
    heading_3_style_id: str | None = None
    bibliography_style_id: str | None = None
    for paragraph in document_tree.findall(".//w:p", XML_NAMESPACES):
        if paragraph_contains_instruction(paragraph, "ZOTERO_BIBL"):
            bibliography_style_id = get_paragraph_style_id(paragraph)
            if bibliography_style_id:
                break

    for style in styles_tree.findall("w:style", XML_NAMESPACES):
        style_id = style.get(f"{WORD_ATTR_PREFIX}styleId")
        style_type = style.get(f"{WORD_ATTR_PREFIX}type")
        name_element = style.find("w:name", XML_NAMESPACES)
        style_name = name_element.get(f"{WORD_ATTR_PREFIX}val") if name_element is not None else ""
        normalized_style_name = re.sub(r"\s+", " ", style_name).strip().lower()
        if style_type == "paragraph" and style_name.lower() == "caption":
            caption_style_id = style_id
        if style_type == "paragraph" and style_name == "Normal":
            normal_style_id = style_id
        if style_type == "paragraph" and title_style_id is None and normalized_style_name in {"title", "大标题"}:
            title_style_id = style_id
        if style_type == "paragraph" and heading_1_style_id is None and normalized_style_name in {"heading 1", "heading1", "标题 1", "标题1"}:
            heading_1_style_id = style_id
        if style_type == "paragraph" and heading_2_style_id is None and normalized_style_name in {"heading 2", "heading2", "标题 2", "标题2"}:
            heading_2_style_id = style_id
        if style_type == "paragraph" and heading_3_style_id is None and normalized_style_name in {"heading 3", "heading3", "标题 3", "标题3"}:
            heading_3_style_id = style_id
        if bibliography_style_id is None and style_type == "paragraph" and re.search(
            r"(书目|bibliography|references?)",
            style_name,
            re.IGNORECASE,
        ):
            bibliography_style_id = style_id

    table_style_counter: Counter[str] = Counter()
    for table in document_tree.findall(".//w:tbl", XML_NAMESPACES):
        table_properties = table.find("w:tblPr", XML_NAMESPACES)
        if table_properties is None:
            continue
        table_style = table_properties.find("w:tblStyle", XML_NAMESPACES)
        if table_style is None:
            continue
        style_id = table_style.get(f"{WORD_ATTR_PREFIX}val")
        if style_id:
            table_style_counter[style_id] += 1

    table_paragraph_counter: Counter[str] = Counter()
    for paragraph in document_tree.findall(".//w:tbl//w:p", XML_NAMESPACES):
        paragraph_properties = paragraph.find("w:pPr", XML_NAMESPACES)
        if paragraph_properties is None:
            continue
        paragraph_style = paragraph_properties.find("w:pStyle", XML_NAMESPACES)
        if paragraph_style is None:
            continue
        style_id = paragraph_style.get(f"{WORD_ATTR_PREFIX}val")
        if style_id:
            table_paragraph_counter[style_id] += 1

    uri_prefix_match = re.search(r"http://zotero\.org/users/\d+/items/", raw_document_xml)

    return TemplateDocxHints(
        caption_style_id=caption_style_id,
        table_style_id=table_style_counter.most_common(1)[0][0] if table_style_counter else None,
        table_paragraph_style_id=(
            table_paragraph_counter.most_common(1)[0][0] if table_paragraph_counter else None
        ),
        normal_style_id=normal_style_id,
        title_style_id=title_style_id,
        heading_1_style_id=heading_1_style_id,
        heading_2_style_id=heading_2_style_id,
        heading_3_style_id=heading_3_style_id,
        bibliography_style_id=bibliography_style_id,
        zotero_item_uri_prefix=uri_prefix_match.group(0) if uri_prefix_match else None,
    )


def collect_referenced_style_ids(archive_entries: list[tuple[ZipInfo, bytes]]) -> set[str]:
    referenced_style_ids: set[str] = set()
    for info, data in archive_entries:
        if info.filename == "word/styles.xml":
            continue
        if not info.filename.startswith("word/") or not info.filename.endswith(".xml"):
            continue
        try:
            tree = ET.fromstring(data)
        except ET.ParseError:
            continue
        for xpath in (".//w:pStyle", ".//w:rStyle", ".//w:tblStyle"):
            for element in tree.findall(xpath, XML_NAMESPACES):
                style_id = element.get(f"{WORD_ATTR_PREFIX}val")
                if style_id:
                    referenced_style_ids.add(style_id)
    return referenced_style_ids


def sanitize_styles_part(
    styles_xml: bytes,
    hints: TemplateDocxHints,
    referenced_style_ids: set[str] | None = None,
) -> tuple[bytes, bool]:
    styles_tree = ET.fromstring(styles_xml)
    changed = False
    style_map = {
        style.get(f"{WORD_ATTR_PREFIX}styleId"): style
        for style in styles_tree.findall("w:style", XML_NAMESPACES)
        if style.get(f"{WORD_ATTR_PREFIX}styleId")
    }
    keep_ids = collect_allowed_style_ids(style_map, hints, referenced_style_ids or set())

    for style in list(styles_tree.findall("w:style", XML_NAMESPACES)):
        style_id = style.get(f"{WORD_ATTR_PREFIX}styleId")
        if style_id not in keep_ids:
            styles_tree.remove(style)
            changed = True
            continue
        if sanitize_style_element(style, keep_ids):
            changed = True

    if sanitize_doc_defaults(styles_tree):
        changed = True
    if prune_latent_styles(styles_tree):
        changed = True

    return ET.tostring(styles_tree, encoding="utf-8", xml_declaration=True), changed


def collect_allowed_style_ids(
    style_map: dict[str, ET.Element],
    hints: TemplateDocxHints,
    referenced_style_ids: set[str],
) -> set[str]:
    keep_ids = allowed_paragraph_style_ids(hints) | referenced_style_ids
    if hints.table_style_id:
        keep_ids.add(hints.table_style_id)

    for style_id, style in style_map.items():
        style_type = style.get(f"{WORD_ATTR_PREFIX}type")
        is_default = style.get(f"{WORD_ATTR_PREFIX}default") == "1"
        if style_type == "character" and (is_default or style_id == "DefaultParagraphFont"):
            keep_ids.add(style_id)
        elif style_type == "table" and is_default:
            keep_ids.add(style_id)
        elif style_type == "numbering":
            keep_ids.add(style_id)

    for style_id in list(keep_ids):
        add_style_dependency_chain(style_id, style_map, keep_ids)
    return keep_ids


def add_style_dependency_chain(style_id: str, style_map: dict[str, ET.Element], keep_ids: set[str]) -> None:
    style = style_map.get(style_id)
    if style is None:
        return
    based_on = style.find("w:basedOn", XML_NAMESPACES)
    parent_style_id = based_on.get(f"{WORD_ATTR_PREFIX}val") if based_on is not None else None
    if parent_style_id and parent_style_id not in keep_ids:
        keep_ids.add(parent_style_id)
        add_style_dependency_chain(parent_style_id, style_map, keep_ids)


def sanitize_style_element(style: ET.Element, keep_ids: set[str]) -> bool:
    changed = False
    for child_name in ("basedOn", "next", "link"):
        child = style.find(f"w:{child_name}", XML_NAMESPACES)
        if child is None:
            continue
        target = child.get(f"{WORD_ATTR_PREFIX}val")
        if target not in keep_ids:
            style.remove(child)
            changed = True

    for run_properties in style.findall(".//w:rPr", XML_NAMESPACES):
        if ensure_times_new_roman_rfonts(run_properties):
            changed = True
    return changed


def sanitize_doc_defaults(styles_tree: ET.Element) -> bool:
    changed = False
    doc_defaults = ensure_styles_child(styles_tree, "docDefaults")
    rpr_default = ensure_styles_child(doc_defaults, "rPrDefault")
    run_properties = ensure_styles_child(rpr_default, "rPr")
    if ensure_times_new_roman_rfonts(run_properties):
        changed = True
    return changed


def prune_latent_styles(styles_tree: ET.Element) -> bool:
    return False


def sanitize_theme_part(theme_xml: bytes) -> tuple[bytes, bool]:
    theme_tree = ET.fromstring(theme_xml)
    changed = False
    for path in (
        f".//{{{DRAWINGML_NAMESPACE}}}themeElements/{{{DRAWINGML_NAMESPACE}}}fontScheme/{{{DRAWINGML_NAMESPACE}}}majorFont/{{{DRAWINGML_NAMESPACE}}}latin",
        f".//{{{DRAWINGML_NAMESPACE}}}themeElements/{{{DRAWINGML_NAMESPACE}}}fontScheme/{{{DRAWINGML_NAMESPACE}}}minorFont/{{{DRAWINGML_NAMESPACE}}}latin",
    ):
        element = theme_tree.find(path)
        if element is None:
            continue
        if element.get("typeface") != WESTERN_FONT_FAMILY:
            element.set("typeface", WESTERN_FONT_FAMILY)
            changed = True
    return ET.tostring(theme_tree, encoding="utf-8", xml_declaration=True), changed


def ensure_styles_child(parent: ET.Element, child_name: str) -> ET.Element:
    child = parent.find(f"w:{child_name}", XML_NAMESPACES)
    if child is not None:
        return child
    child = ET.Element(f"{WORD_ATTR_PREFIX}{child_name}")
    parent.append(child)
    return child


def parse_document_relationships(rels_xml: bytes | None) -> dict[str, str]:
    if rels_xml is None:
        return {}
    relationships_root = ET.fromstring(rels_xml)
    relationships: dict[str, str] = {}
    for rel in relationships_root.findall(f"{{{PACKAGE_RELATIONSHIP_NAMESPACE}}}Relationship"):
        rel_id = rel.get("Id")
        target = rel.get("Target")
        if rel_id and target:
            relationships[rel_id] = target
    return relationships


def collect_used_relationship_ids(part_xml: bytes) -> set[str]:
    try:
        root = ET.fromstring(part_xml)
    except ET.ParseError:
        return set()

    used_ids: set[str] = set()
    for element in root.iter():
        rel_id = element.get(f"{REL_ATTR_PREFIX}id")
        if rel_id:
            used_ids.add(rel_id)
    return used_ids


def owning_part_for_relationship_part(part_name: str) -> str | None:
    if not part_name.startswith("word/_rels/") or not part_name.endswith(".rels"):
        return None
    relative_name = part_name[len("word/_rels/") :]
    if not relative_name.endswith(".rels"):
        return None
    return f"word/{relative_name[:-5]}"


def prune_unused_hyperlink_relationships_for_part(
    part_xml: bytes,
    rels_xml: bytes,
) -> tuple[bytes | None, bool]:
    try:
        relationships_root = ET.fromstring(rels_xml)
    except ET.ParseError:
        return rels_xml, False

    used_ids = collect_used_relationship_ids(part_xml)
    changed = False
    for relationship in list(relationships_root.findall(f"{{{PACKAGE_RELATIONSHIP_NAMESPACE}}}Relationship")):
        relationship_type = relationship.get("Type") or ""
        relationship_id = relationship.get("Id")
        if not relationship_type.endswith("/hyperlink"):
            continue
        if relationship_id and relationship_id in used_ids:
            continue
        relationships_root.remove(relationship)
        changed = True

    if not changed:
        return rels_xml, False
    if not list(relationships_root):
        return None, True
    return ET.tostring(relationships_root, encoding="utf-8", xml_declaration=True), True


def prune_unused_hyperlink_relationship_parts(
    archive_part_map: dict[str, bytes],
    updated_parts: dict[str, bytes],
) -> tuple[dict[str, bytes], set[str], bool]:
    current_parts = dict(archive_part_map)
    current_parts.update(updated_parts)
    relationship_updates: dict[str, bytes] = {}
    relationship_removals: set[str] = set()
    changed = False

    for part_name, rels_xml in current_parts.items():
        owner_part = owning_part_for_relationship_part(part_name)
        if owner_part is None:
            continue
        owner_xml = current_parts.get(owner_part)
        if owner_xml is None:
            relationship_removals.add(part_name)
            changed = True
            continue
        pruned_xml, part_changed = prune_unused_hyperlink_relationships_for_part(owner_xml, rels_xml)
        if not part_changed:
            continue
        changed = True
        if pruned_xml is None:
            relationship_removals.add(part_name)
            relationship_updates.pop(part_name, None)
            continue
        relationship_updates[part_name] = pruned_xml

    return relationship_updates, relationship_removals, changed


def apply_body_paragraph_hints(
    document_tree: ET.Element,
    hints: TemplateDocxHints,
    bibliography_heading: str = "参考文献",
) -> bool:
    if hints.normal_style_id is None:
        return False
    changed = False
    paragraphs = direct_body_paragraphs(document_tree)
    bibliography_heading_index = find_bibliography_heading_index(paragraphs, bibliography_heading)
    for index, paragraph in enumerate(paragraphs):
        if bibliography_heading_index is not None and index > bibliography_heading_index:
            continue
        style_id = get_paragraph_style_id(paragraph)
        if style_id not in {None, "BodyText", "FirstParagraph"}:
            continue
        set_paragraph_style_id(paragraph, hints.normal_style_id)
        changed = True

    if hints.title_style_id and not any(get_paragraph_style_id(paragraph) == hints.title_style_id for paragraph in paragraphs):
        heading_style_ids = {
            style_id
            for style_id in (hints.heading_1_style_id, hints.heading_2_style_id, hints.heading_3_style_id)
            if style_id
        }
        first_heading_index = next(
            (
                index
                for index, paragraph in enumerate(paragraphs)
                if (
                    get_paragraph_style_id(paragraph) in heading_style_ids
                    or get_paragraph_text(paragraph).strip() in {"摘要", "Abstract"}
                )
            ),
            None,
        )
        if first_heading_index is not None:
            for paragraph in reversed(paragraphs[:first_heading_index]):
                text = get_paragraph_text(paragraph).strip()
                if not text:
                    continue
                if len(text) < 8 or len(text) > 120:
                    continue
                if text in {"总编室意见", bibliography_heading}:
                    continue
                if re.match(r"^[0-9０-９]+[、.．)]", text):
                    continue
                if text.endswith(("：", ":", "。", "！", "!", "？", "?", "；", ";")):
                    continue
                if get_paragraph_style_id(paragraph) != hints.title_style_id:
                    set_paragraph_style_id(paragraph, hints.title_style_id)
                    changed = True
                if set_paragraph_alignment(paragraph, "center"):
                    changed = True
                break
    return changed


def strip_all_bookmarks(
    document_tree: ET.Element,
    preserved_prefixes: tuple[str, ...] = (),
) -> bool:
    changed = False
    preserved_ids: set[str] = set()
    for bookmark in document_tree.findall(".//w:bookmarkStart", XML_NAMESPACES):
        name = bookmark.get(f"{WORD_ATTR_PREFIX}name") or ""
        bookmark_id = bookmark.get(f"{WORD_ATTR_PREFIX}id") or ""
        if any(name.startswith(prefix) for prefix in preserved_prefixes):
            if bookmark_id:
                preserved_ids.add(bookmark_id)
            continue
        parent = find_direct_parent(document_tree, bookmark)
        if parent is None:
            continue
        parent.remove(bookmark)
        changed = True
    for bookmark in document_tree.findall(".//w:bookmarkEnd", XML_NAMESPACES):
        bookmark_id = bookmark.get(f"{WORD_ATTR_PREFIX}id") or ""
        if bookmark_id in preserved_ids:
            continue
        parent = find_direct_parent(document_tree, bookmark)
        if parent is None:
            continue
        parent.remove(bookmark)
        changed = True
    return changed


def apply_default_citation_anchor_policy(
    document_tree: ET.Element,
    relationship_targets: dict[str, str],
    zotero_context: ZoteroDocxContext,
    bibliography_heading: str = "参考文献",
) -> bool:
    allowed_anchors = {entry.anchor_id for entry in zotero_context.bibliography_entries}
    changed = remove_disallowed_internal_hyperlinks(document_tree, relationship_targets, allowed_anchors)
    changed |= normalize_internal_anchor_bookmarks(
        document_tree,
        allowed_anchors,
        bibliography_heading,
        preserved_prefixes=(NATIVE_CROSS_REFERENCE_BOOKMARK_PREFIX,),
    )
    changed |= style_default_internal_hyperlinks(document_tree, allowed_anchors)
    return changed


def normalize_internal_anchor_bookmarks(
    document_tree: ET.Element,
    allowed_anchors: set[str] | None = None,
    bibliography_heading: str = "参考文献",
    preserved_prefixes: tuple[str, ...] = (),
) -> bool:
    changed = False
    allowed_anchors = allowed_anchors or set()
    bibliography_body_index = find_bibliography_body_index(document_tree, bibliography_heading)
    body = document_tree.find("w:body", XML_NAMESPACES)
    bookmark_starts = document_tree.findall(".//w:bookmarkStart", XML_NAMESPACES)
    bookmark_body_indices = {
        id(bookmark): find_body_child_index_containing(document_tree, bookmark)
        for bookmark in bookmark_starts
    }
    bookmark_ends = {
        bookmark.get(f"{WORD_ATTR_PREFIX}id"): bookmark
        for bookmark in document_tree.findall(".//w:bookmarkEnd", XML_NAMESPACES)
        if bookmark.get(f"{WORD_ATTR_PREFIX}id") is not None
    }
    remove_bookmark_ids: set[str] = set()
    renamed_anchors: dict[str, str] = {}

    for bookmark in bookmark_starts:
        name = bookmark.get(f"{WORD_ATTR_PREFIX}name")
        bookmark_id = bookmark.get(f"{WORD_ATTR_PREFIX}id")
        if name and any(name.startswith(prefix) for prefix in preserved_prefixes):
            continue
        body_index = bookmark_body_indices.get(id(bookmark))
        keep = bool(
            name in allowed_anchors
            and bibliography_body_index is not None
            and body_index is not None
            and body_index > bibliography_body_index
        )
        if not keep:
            if bookmark_id is not None:
                remove_bookmark_ids.add(bookmark_id)
            parent = find_direct_parent(document_tree, bookmark)
            if parent is not None:
                parent.remove(bookmark)
                changed = True
            continue

        if name and not name.startswith("_"):
            hidden_name = f"_{name}"
            bookmark.set(f"{WORD_ATTR_PREFIX}name", hidden_name)
            renamed_anchors[name] = hidden_name
            changed = True

        bookmark_end = bookmark_ends.get(bookmark_id)
        if bookmark_end is None:
            continue
        start_parent = find_direct_parent(document_tree, bookmark)
        end_parent = find_direct_parent(document_tree, bookmark_end)
        if start_parent is None or end_parent is None:
            continue

        target_paragraph: ET.Element | None = None
        if start_parent.tag == f"{WORD_ATTR_PREFIX}p":
            target_paragraph = start_parent
        elif end_parent.tag == f"{WORD_ATTR_PREFIX}p":
            target_paragraph = end_parent
        elif body is not None and start_parent is body and end_parent is body:
            body_children = list(body)
            start_body_index = body_children.index(bookmark)
            end_body_index = body_children.index(bookmark_end)
            for child in body_children[start_body_index + 1 : end_body_index]:
                if child.tag == f"{WORD_ATTR_PREFIX}p":
                    target_paragraph = child
                    break
            if target_paragraph is None:
                for child in body_children[end_body_index + 1 :]:
                    if child.tag == f"{WORD_ATTR_PREFIX}p":
                        target_paragraph = child
                        break
            if target_paragraph is None:
                for child in reversed(body_children[:start_body_index]):
                    if child.tag == f"{WORD_ATTR_PREFIX}p":
                        target_paragraph = child
                        break

        if target_paragraph is not None and (start_parent is body or end_parent is body):
            insert_at = paragraph_content_insert_index(target_paragraph)
            if start_parent is not target_paragraph:
                start_parent.remove(bookmark)
            else:
                insert_at = list(target_paragraph).index(bookmark)
            if end_parent is not target_paragraph:
                end_parent.remove(bookmark_end)
            target_paragraph.insert(insert_at, bookmark)
            target_paragraph.insert(insert_at + 1, bookmark_end)
            changed = True
            continue

        start_index = list(start_parent).index(bookmark)
        if start_parent is end_parent:
            end_index = list(start_parent).index(bookmark_end)
            if end_index == start_index + 1:
                continue
        end_parent.remove(bookmark_end)
        start_parent.insert(start_index + 1, bookmark_end)
        changed = True

    for bookmark_id in remove_bookmark_ids:
        bookmark_end = bookmark_ends.get(bookmark_id)
        if bookmark_end is None:
            continue
        parent = find_direct_parent(document_tree, bookmark_end)
        if parent is not None:
            parent.remove(bookmark_end)
            changed = True
    for hyperlink in document_tree.findall(".//w:hyperlink", XML_NAMESPACES):
        anchor = hyperlink.get(f"{WORD_ATTR_PREFIX}anchor")
        if anchor in renamed_anchors:
            hyperlink.set(f"{WORD_ATTR_PREFIX}anchor", renamed_anchors[anchor])
            changed = True
    return changed


def remove_disallowed_internal_hyperlinks(
    document_tree: ET.Element,
    relationship_targets: dict[str, str],
    allowed_anchors: set[str],
) -> bool:
    changed = False
    for paragraph in document_tree.findall(".//w:p", XML_NAMESPACES):
        for child in list(paragraph):
            if child.tag != f"{WORD_ATTR_PREFIX}hyperlink":
                continue
            target_anchor = child.get(f"{WORD_ATTR_PREFIX}anchor")
            rel_id = child.get(f"{REL_ATTR_PREFIX}id")
            if target_anchor is None and rel_id:
                target_anchor = extract_anchor_from_relationship_target(relationship_targets.get(rel_id))
            if target_anchor is None or target_anchor in allowed_anchors:
                continue
            flatten_hyperlink_child(paragraph, child)
            changed = True
    return changed


def flatten_hyperlink_child(
    paragraph: ET.Element,
    hyperlink: ET.Element,
    reference_style: bool = False,
) -> None:
    insert_at = list(paragraph).index(hyperlink)
    replacement_runs: list[ET.Element] = []
    for run in hyperlink.findall("w:r", XML_NAMESPACES):
        new_run = ET.Element(f"{WORD_ATTR_PREFIX}r")
        run_properties = run.find("w:rPr", XML_NAMESPACES)
        cloned_properties = clone_element(run_properties) if run_properties is not None else None
        if reference_style:
            cloned_properties = ensure_reference_run_properties(cloned_properties)
        elif cloned_properties is not None:
            remove_run_style(cloned_properties)
        if cloned_properties is not None:
            new_run.append(cloned_properties)
        for node in list(run):
            if node.tag == f"{WORD_ATTR_PREFIX}rPr":
                continue
            new_run.append(clone_element(node))
        if len(new_run):
            replacement_runs.append(new_run)
    if not replacement_runs:
        replacement_runs.append(
            build_field_run(
                text=get_element_text(hyperlink),
                rpr_template=ensure_reference_run_properties(None) if reference_style else None,
            )
        )
    paragraph.remove(hyperlink)
    for offset, new_run in enumerate(replacement_runs):
        paragraph.insert(insert_at + offset, new_run)


def style_default_internal_hyperlinks(document_tree: ET.Element, allowed_anchors: set[str]) -> bool:
    changed = False
    hidden_allowed_anchors = {f"_{anchor}" for anchor in allowed_anchors}
    citation_anchors = allowed_anchors | hidden_allowed_anchors
    for paragraph in document_tree.findall(".//w:p", XML_NAMESPACES):
        if absorb_adjacent_citation_brackets(paragraph, citation_anchors):
            changed = True
    for hyperlink in document_tree.findall(".//w:hyperlink", XML_NAMESPACES):
        anchor = hyperlink.get(f"{WORD_ATTR_PREFIX}anchor")
        if anchor not in citation_anchors:
            continue
        for run in hyperlink.findall("w:r", XML_NAMESPACES):
            if apply_plain_sky_blue_link_style(run):
                changed = True
    return changed


def absorb_adjacent_citation_brackets(paragraph: ET.Element, allowed_anchors: set[str]) -> bool:
    changed = False
    for hyperlink in paragraph.findall("w:hyperlink", XML_NAMESPACES):
        anchor = hyperlink.get(f"{WORD_ATTR_PREFIX}anchor")
        if anchor not in allowed_anchors:
            continue
        children = list(paragraph)
        index = children.index(hyperlink)
        if index > 0:
            opening_bracket = pop_run_edge_bracket(paragraph, children[index - 1], trailing=True)
            if opening_bracket:
                hyperlink.insert(0, build_text_run(opening_bracket))
                changed = True
        children = list(paragraph)
        index = children.index(hyperlink)
        if index + 1 < len(children):
            closing_bracket = pop_run_edge_bracket(paragraph, children[index + 1], trailing=False)
            if closing_bracket:
                hyperlink.append(build_text_run(closing_bracket))
                changed = True
    return changed


def pop_run_edge_bracket(paragraph: ET.Element, node: ET.Element, trailing: bool) -> str | None:
    if node.tag != f"{WORD_ATTR_PREFIX}r":
        return None
    text_nodes = node.findall("w:t", XML_NAMESPACES)
    if not text_nodes:
        return None
    text_node = text_nodes[-1] if trailing else text_nodes[0]
    text = text_node.text or ""
    if not text:
        return None
    brackets = OPENING_CITATION_BRACKETS if trailing else CLOSING_CITATION_BRACKETS
    probe = text[-1] if trailing else text[0]
    if probe not in brackets:
        return None
    text_node.text = text[:-1] if trailing else text[1:]
    if get_element_text(node) == "":
        paragraph.remove(node)
    return probe


def build_text_run(text: str) -> ET.Element:
    run = ET.Element(f"{WORD_ATTR_PREFIX}r")
    text_node = ET.SubElement(run, f"{WORD_ATTR_PREFIX}t")
    if text[:1].isspace() or text[-1:].isspace():
        text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text_node.text = text
    return run


def apply_plain_sky_blue_link_style(run: ET.Element) -> bool:
    changed = False
    run_properties = run.find("w:rPr", XML_NAMESPACES)
    if run_properties is None:
        run_properties = ET.Element(f"{WORD_ATTR_PREFIX}rPr")
        run.insert(0, run_properties)
        changed = True
    ensure_reference_run_properties(run_properties)
    for existing in list(run_properties.findall("w:u", XML_NAMESPACES)):
        run_properties.remove(existing)
        changed = True
    underline = ET.SubElement(run_properties, f"{WORD_ATTR_PREFIX}u")
    underline.set(f"{WORD_ATTR_PREFIX}val", "none")
    return True


def find_bibliography_body_index(document_tree: ET.Element, bibliography_heading: str) -> int | None:
    body = document_tree.find("w:body", XML_NAMESPACES)
    if body is None:
        return None
    for index, child in enumerate(list(body)):
        if child.tag == f"{WORD_ATTR_PREFIX}p" and get_element_text(child).strip() == bibliography_heading:
            return index
    return None


def find_body_child_index_containing(document_tree: ET.Element, target: ET.Element) -> int | None:
    body = document_tree.find("w:body", XML_NAMESPACES)
    if body is None:
        return None
    for index, child in enumerate(list(body)):
        if child is target or any(descendant is target for descendant in child.iter()):
            return index
    return None


def populate_zotero_anchor_aliases_from_bibliography(
    document_tree: ET.Element,
    zotero_context: ZoteroDocxContext,
    bibliography_heading: str = "参考文献",
) -> None:
    if not zotero_context.bibliography_entries:
        return

    paragraphs = direct_body_paragraphs(document_tree)
    bibliography_heading_index = find_bibliography_heading_index(paragraphs, bibliography_heading)
    if bibliography_heading_index is None:
        return

    body = document_tree.find("w:body", XML_NAMESPACES)
    if body is None:
        return

    heading_paragraph = paragraphs[bibliography_heading_index]
    body_children = list(body)
    heading_body_index = body_children.index(heading_paragraph)

    used_source_keys: set[str] = set()
    sequential_index = 0
    pending_bookmark_names: list[str] = []
    for child in body_children[heading_body_index + 1 :]:
        if child.tag == f"{WORD_ATTR_PREFIX}sectPr":
            break
        if child.tag == f"{WORD_ATTR_PREFIX}bookmarkStart":
            name = child.get(f"{WORD_ATTR_PREFIX}name")
            if name and not name.startswith("_Toc"):
                pending_bookmark_names.append(name)
            continue
        if child.tag == f"{WORD_ATTR_PREFIX}bookmarkEnd":
            continue
        if child.tag != f"{WORD_ATTR_PREFIX}p":
            pending_bookmark_names = []
            continue
        bookmark_names = pending_bookmark_names + bibliography_bookmark_names(child)
        pending_bookmark_names = []
        if not bookmark_names:
            continue

        target = match_bibliography_paragraph_target(child, zotero_context.bibliography_entries, used_source_keys)
        if target is None:
            while sequential_index < len(zotero_context.bibliography_entries):
                candidate = zotero_context.bibliography_entries[sequential_index]
                sequential_index += 1
                if candidate.source_key not in used_source_keys:
                    target = candidate
                    break
        if target is None:
            continue

        used_source_keys.add(target.source_key)
        for bookmark_name in bookmark_names:
            zotero_context.by_anchor[bookmark_name] = target


def bibliography_bookmark_names(paragraph: ET.Element) -> list[str]:
    bookmark_names: list[str] = []
    for bookmark in paragraph.findall(".//w:bookmarkStart", XML_NAMESPACES):
        name = bookmark.get(f"{WORD_ATTR_PREFIX}name")
        if name and not name.startswith("_Toc"):
            bookmark_names.append(name)
    return bookmark_names


def strip_bibliography_bookmarks(document_tree: ET.Element, bibliography_heading_index: int) -> bool:
    body = document_tree.find("w:body", XML_NAMESPACES)
    if body is None:
        return False

    paragraphs = direct_body_paragraphs(document_tree)
    if bibliography_heading_index >= len(paragraphs):
        return False

    heading_paragraph = paragraphs[bibliography_heading_index]
    body_children = list(body)
    heading_body_index = body_children.index(heading_paragraph)
    removable_nodes: list[ET.Element] = []
    removed_in_paragraphs = False
    for child in body_children[heading_body_index + 1 :]:
        if child.tag == f"{WORD_ATTR_PREFIX}sectPr":
            break
        if child.tag in {f"{WORD_ATTR_PREFIX}bookmarkStart", f"{WORD_ATTR_PREFIX}bookmarkEnd"}:
            removable_nodes.append(child)
            continue
        if child.tag == f"{WORD_ATTR_PREFIX}p" and strip_bookmarks_in_paragraph(child):
            removed_in_paragraphs = True

    for removable in removable_nodes:
        body.remove(removable)
    return bool(removable_nodes) or removed_in_paragraphs


def match_bibliography_paragraph_target(
    paragraph: ET.Element,
    bibliography_entries: list[CitationTarget],
    used_source_keys: set[str],
) -> CitationTarget | None:
    paragraph_text = get_element_text(paragraph)
    normalized_paragraph_text = normalize_reference_text(paragraph_text)

    exact_matches = [
        entry
        for entry in bibliography_entries
        if entry.source_key not in used_source_keys
        and normalize_reference_text(entry.formatted_reference) == normalized_paragraph_text
    ]
    if len(exact_matches) == 1:
        return exact_matches[0]

    doi_matches = [
        entry
        for entry in bibliography_entries
        if entry.source_key not in used_source_keys and reference_text_mentions_source_key(paragraph_text, entry.source_key)
    ]
    if len(doi_matches) == 1:
        return doi_matches[0]
    return None


def normalize_reference_text(text: str) -> str:
    normalized = text.replace("\xa0", " ").strip().lower()
    normalized = re.sub(r"\s+", " ", normalized)
    return normalized


def reference_text_mentions_source_key(text: str, source_key: str) -> bool:
    normalized_text = normalize_reference_text(text)
    normalized_doi = normalize_doi(source_key)
    if normalized_doi and normalized_doi.lower() in normalized_text:
        return True
    normalized_url = normalize_url(source_key)
    if normalized_url and normalized_url.lower() in normalized_text:
        return True
    return False


def apply_bibliography_hints(
    document_tree: ET.Element,
    hints: TemplateDocxHints,
    zotero_context: ZoteroDocxContext,
    bibliography_heading: str = "参考文献",
    enable_zotero: bool = False,
) -> bool:
    paragraphs = direct_body_paragraphs(document_tree)
    bibliography_heading_index = find_bibliography_heading_index(paragraphs, bibliography_heading)
    if bibliography_heading_index is None:
        return False

    bibliography_paragraphs = paragraphs[bibliography_heading_index + 1 :]
    if not bibliography_paragraphs:
        return False

    changed = False
    for paragraph in bibliography_paragraphs:
        if hints.bibliography_style_id and get_paragraph_style_id(paragraph) != hints.bibliography_style_id:
            set_paragraph_style_id(paragraph, hints.bibliography_style_id)
            changed = True
        if flatten_hyperlinks_in_paragraph(paragraph):
            changed = True

    if enable_zotero and strip_bibliography_bookmarks(document_tree, bibliography_heading_index):
        changed = True

    if not enable_zotero:
        return changed

    first_paragraph = bibliography_paragraphs[0]
    if not paragraph_contains_instruction(first_paragraph, "ZOTERO_BIBL"):
        reference_rpr = first_run_properties(first_paragraph)
        insert_at = paragraph_content_insert_index(first_paragraph)
        instruction_runs = build_instruction_field_runs(
            ' ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY ',
            reference_rpr,
        )
        first_paragraph.insert(insert_at, build_field_run(fld_char_type="begin", rpr_template=reference_rpr))
        for offset, run in enumerate(instruction_runs):
            first_paragraph.insert(insert_at + 1 + offset, run)
        first_paragraph.insert(
            insert_at + 1 + len(instruction_runs),
            build_field_run(fld_char_type="separate", rpr_template=reference_rpr),
        )
        body = document_tree.find(f".//{{{WORD_NAMESPACE}}}body")
        if body is None:
            return changed

        last_paragraph = bibliography_paragraphs[-1]
        end_paragraph = ET.Element(f"{WORD_ATTR_PREFIX}p")
        last_paragraph_properties = last_paragraph.find(f"{{{WORD_NAMESPACE}}}pPr")
        if last_paragraph_properties is not None:
            end_paragraph.append(clone_element(last_paragraph_properties))
        end_paragraph.append(build_field_run(fld_char_type="end", rpr_template=reference_rpr))

        body_children = list(body)
        insert_index = body_children.index(last_paragraph) + 1
        body.insert(insert_index, end_paragraph)
        changed = True
    return changed


def convert_citation_hyperlinks_to_zotero_fields(
    document_tree: ET.Element,
    relationship_targets: dict[str, str],
    zotero_context: ZoteroDocxContext,
) -> bool:
    changed = False
    occurrence_counts: dict[tuple[tuple[str, ...], str], int] = {}
    table_paragraph_ids = {
        id(paragraph)
        for paragraph in document_tree.findall(".//w:tbl//w:p", XML_NAMESPACES)
    }
    for paragraph in document_tree.findall(".//w:p", XML_NAMESPACES):
        if id(paragraph) in table_paragraph_ids:
            continue
        if paragraph.find(f".//{{{WORD_NAMESPACE}}}fldChar") is not None:
            continue
        if paragraph.find(f".//{{{MATH_NAMESPACE}}}oMath") is not None:
            continue
        if paragraph.find(f".//{{{MATH_NAMESPACE}}}oMathPara") is not None:
            continue
        children = list(paragraph)
        index = 0
        while index < len(children):
            child = children[index]
            citation_target = resolve_citation_hyperlink_target(child, relationship_targets, zotero_context)
            if citation_target is None:
                index += 1
                continue

            display_text, citation_targets, removable_nodes = collect_citation_cluster(
                paragraph,
                children,
                index,
                relationship_targets,
                zotero_context,
            )
            if not citation_targets:
                index += 1
                continue

            insert_at = list(paragraph).index(citation_targets[0][1])
            reference_element = citation_targets[0][1]
            for _, removable_node in citation_targets:
                paragraph.remove(removable_node)
            for removable_node in removable_nodes:
                if removable_node in paragraph:
                    paragraph.remove(removable_node)

            field_targets = [target for target, _ in citation_targets]
            signature = make_citation_field_shell_signature(
                [target.source_key for target in field_targets],
                display_text,
            )
            occurrence_index = occurrence_counts.get(signature, 0) + 1
            occurrence_counts[signature] = occurrence_index
            preserved_field_runs = pop_matching_citation_field_shell(
                zotero_context,
                field_targets,
                display_text,
                occurrence_index,
            )
            for offset, field_run in enumerate(
                preserved_field_runs
                or build_zotero_citation_field_elements(
                    display_text,
                    field_targets,
                    reference_element,
                    occurrence_index=occurrence_index,
                )
            ):
                paragraph.insert(insert_at + offset, field_run)

            children = list(paragraph)
            index = insert_at + 1
            changed = True
    return changed


def cluster_next_index(
    children: list[ET.Element],
    start_index: int,
    citation_targets: list[tuple[CitationTarget, ET.Element]],
    removable_nodes: list[ET.Element],
) -> int:
    cluster_nodes = {node for _target, node in citation_targets}
    cluster_nodes.update(removable_nodes)
    index = start_index
    while index < len(children) and (children[index] in cluster_nodes or index == start_index):
        index += 1
    return index


def strip_bookmarks_in_paragraph(paragraph: ET.Element) -> bool:
    changed = False
    for bookmark in paragraph.findall(".//w:bookmarkStart", XML_NAMESPACES):
        parent = find_direct_parent(paragraph, bookmark)
        if parent is None:
            continue
        parent.remove(bookmark)
        changed = True
    for bookmark in paragraph.findall(".//w:bookmarkEnd", XML_NAMESPACES):
        parent = find_direct_parent(paragraph, bookmark)
        if parent is None:
            continue
        parent.remove(bookmark)
        changed = True
    return changed


def find_direct_parent(root: ET.Element, target: ET.Element) -> ET.Element | None:
    for parent in root.iter():
        for child in list(parent):
            if child is target:
                return parent
    return None


def infer_bibliography_target_from_anchor(anchor: str, zotero_context: ZoteroDocxContext) -> CitationTarget | None:
    match = re.fullmatch(r"(?:_+)?[Rr]ef(\d+)", anchor.strip())
    if not match:
        return None
    index = int(match.group(1)) - 1
    if index < 0 or index >= len(zotero_context.bibliography_entries):
        return None
    return zotero_context.bibliography_entries[index]


def citation_display_signature(display_text: str) -> tuple[str, str, str | None] | None:
    normalized = display_text.replace("\xa0", " ").strip()
    normalized = re.sub(r"^[\s\(\[\{（【]+|[\s\)\]\}）】.,;；，]+$", "", normalized)
    year_match = re.search(r"(19|20)\d{2}", normalized)
    if year_match is None:
        return None

    before_year = normalized[: year_match.start()].strip(" ,;；，")
    before_year = re.sub(r"\bet\s+al\.?$", "", before_year, flags=re.IGNORECASE).strip(" ,;；，")
    if not before_year:
        return None

    tokens = before_year.split()
    surname = normalize_citation_name_token(tokens[-1])
    if not surname:
        return None

    initial: str | None = None
    if len(tokens) > 1 and re.fullmatch(r"[A-Za-z]\.", tokens[-2]):
        initial = tokens[-2][0].lower()
    return surname, year_match.group(0), initial


def bibliography_entry_matches_signature(entry: CitationTarget, signature: tuple[str, str, str | None]) -> bool:
    surname, year, initial = signature
    reference = entry.formatted_reference.replace("\xa0", " ")
    if year not in reference:
        return False

    match = re.match(r"\s*([^,\(\.]+)\s*,\s*([^,\(\.])?", reference)
    if match is None:
        return False
    entry_surname = normalize_citation_name_token(match.group(1))
    if entry_surname != surname:
        return False
    if initial is None:
        return True
    entry_initial = (match.group(2) or "").lower()
    return entry_initial == initial


def normalize_citation_name_token(value: str) -> str:
    normalized = re.sub(r"[^A-Za-z\u4e00-\u9fff]+", "", value).lower()
    return normalized


def resolve_citation_hyperlink_target(
    element: ET.Element,
    relationship_targets: dict[str, str],
    zotero_context: ZoteroDocxContext,
) -> CitationTarget | None:
    if element.tag != f"{WORD_ATTR_PREFIX}hyperlink":
        return None
    display_text = get_element_text(element)
    anchor = element.get(f"{WORD_ATTR_PREFIX}anchor")
    if anchor:
        target = zotero_context.lookup(anchor=anchor)
        if target is None:
            target = infer_bibliography_target_from_anchor(anchor, zotero_context)
            if target is not None:
                zotero_context.by_anchor[anchor] = target
        if target is not None:
            return target
        display_target = zotero_context.lookup_display_text(display_text)
        if display_target is not None:
            zotero_context.by_anchor[anchor] = display_target
            return display_target
        if looks_like_citation_display_text(display_text):
            fallback_target = synthesize_inline_citation_target(display_text, anchor)
            zotero_context.by_anchor[anchor] = fallback_target
            return fallback_target
    rel_id = element.get(f"{REL_ATTR_PREFIX}id")
    if rel_id is not None:
        target = relationship_targets.get(rel_id)
        target_anchor = extract_anchor_from_relationship_target(target)
        if target_anchor:
            resolved_target = zotero_context.lookup(anchor=target_anchor)
            if resolved_target is None:
                resolved_target = infer_bibliography_target_from_anchor(target_anchor, zotero_context)
                if resolved_target is not None:
                    zotero_context.by_anchor[target_anchor] = resolved_target
            if resolved_target is not None:
                return resolved_target
            display_target = zotero_context.lookup_display_text(display_text)
            if display_target is not None:
                zotero_context.by_anchor[target_anchor] = display_target
                return display_target
            if looks_like_citation_display_text(display_text):
                fallback_target = synthesize_inline_citation_target(display_text, target_anchor)
                zotero_context.by_anchor[target_anchor] = fallback_target
                return fallback_target
        resolved_target = zotero_context.lookup(target)
        if resolved_target is not None:
            return resolved_target
        display_target = zotero_context.lookup_display_text(display_text)
        if display_target is not None:
            return display_target
    display_target = zotero_context.lookup_display_text(display_text)
    if display_target is not None:
        return display_target
    if looks_like_citation_display_text(display_text):
        synthetic_anchor = make_anchor_id(f"inline-cite-{display_text}")
        fallback_target = synthesize_inline_citation_target(display_text, synthetic_anchor)
        zotero_context.by_anchor[synthetic_anchor] = fallback_target
        return fallback_target
    return None


def extract_anchor_from_relationship_target(target: str | None) -> str | None:
    if not target:
        return None
    if target.startswith("#"):
        anchor = target[1:]
    elif "#" in target:
        anchor = target.split("#", 1)[1]
    elif "://" not in target and "/" not in target and "\\" not in target:
        anchor = target
    else:
        return None
    anchor = anchor.strip()
    return anchor or None


def looks_like_citation_display_text(text: str) -> bool:
    normalized = text.replace("\xa0", " ").strip()
    if not normalized or len(normalized) > 200:
        return False
    has_year = bool(re.search(r"(19|20)\d{2}", normalized))
    has_numeric_citation = bool(
        re.fullmatch(r"[\[(（【]?\s*\d+(?:\s*[-,，;；–]\s*\d+)*\s*[\])）】]?", normalized)
    )
    has_word_like_text = bool(re.search(r"[\w\u4e00-\u9fff]", normalized))
    return has_word_like_text and (has_year or has_numeric_citation)


def synthesize_inline_citation_target(display_text: str, anchor: str) -> CitationTarget:
    normalized = display_text.replace("\xa0", " ").strip()
    year_match = re.search(r"(19|20)\d{2}", normalized)
    item_data: dict[str, object] = {
        "id": anchor,
        "type": "article-journal",
        "title": normalized,
    }
    if year_match:
        item_data["issued"] = {"date-parts": [[int(year_match.group(0))]]}
    return CitationTarget(
        source_key=anchor,
        formatted_reference=normalized,
        zotero_item_key=None,
        item_data=item_data,
        uri=None,
        anchor_id=anchor,
    )


def collect_citation_cluster(
    paragraph: ET.Element,
    children: list[ET.Element],
    start_index: int,
    relationship_targets: dict[str, str],
    zotero_context: ZoteroDocxContext,
) -> tuple[str, list[tuple[CitationTarget, ET.Element]], list[ET.Element]]:
    citation_nodes: list[tuple[CitationTarget, ET.Element]] = []
    removable_nodes: list[ET.Element] = []
    display_parts: list[str] = []

    if start_index > 0:
        prefix = consume_trailing_citation_prefix(children[start_index - 1])
        if prefix:
            display_parts.append(prefix)
            remove_if_empty_run(paragraph, children[start_index - 1])

    index = start_index
    while index < len(children):
        child = children[index]
        citation_target = resolve_citation_hyperlink_target(child, relationship_targets, zotero_context)
        if citation_target is not None:
            citation_nodes.append((citation_target, child))
            display_parts.append(get_element_text(child))
            index += 1
            continue

        if is_standalone_citation_separator(child):
            display_parts.append(get_element_text(child))
            removable_nodes.append(child)
            index += 1
            continue

        suffix = consume_leading_citation_suffix(child)
        if suffix:
            display_parts.append(suffix)
            if get_element_text(child).strip():
                break
            removable_nodes.append(child)
            index += 1
        break

    return "".join(display_parts), citation_nodes, removable_nodes


def consume_trailing_citation_prefix(element: ET.Element) -> str:
    if element.tag != f"{WORD_ATTR_PREFIX}r":
        return ""
    text = get_element_text(element)
    match = re.search(r"([\(\[]\s*)$", text)
    if match is None:
        return ""
    trim_run_text_suffix(element, len(match.group(1)))
    return match.group(1)


def consume_leading_citation_suffix(element: ET.Element) -> str:
    if element.tag != f"{WORD_ATTR_PREFIX}r":
        return ""
    text = get_element_text(element)
    match = re.match(r"^([\)\]]\s*)", text)
    if match is None:
        return ""
    trim_run_text_prefix(element, len(match.group(1)))
    return match.group(1)


def is_standalone_citation_separator(element: ET.Element) -> bool:
    if element.tag != f"{WORD_ATTR_PREFIX}r":
        return False
    text = get_element_text(element)
    return bool(text) and bool(re.fullmatch(r"[\s;,]+", text))


def remove_if_empty_run(paragraph: ET.Element, element: ET.Element) -> None:
    if element.tag == f"{WORD_ATTR_PREFIX}r" and not get_element_text(element):
        paragraph.remove(element)


def trim_run_text_suffix(run: ET.Element, length: int) -> None:
    text_nodes = run.findall("w:t", XML_NAMESPACES)
    remaining = max(length, 0)
    for text_node in reversed(text_nodes):
        if remaining == 0:
            break
        text = text_node.text or ""
        if not text:
            continue
        if remaining >= len(text):
            text_node.text = ""
            remaining -= len(text)
            continue
        text_node.text = text[:-remaining]
        break


def trim_run_text_prefix(run: ET.Element, length: int) -> None:
    text_nodes = run.findall("w:t", XML_NAMESPACES)
    remaining = max(length, 0)
    for text_node in text_nodes:
        if remaining == 0:
            break
        text = text_node.text or ""
        if not text:
            continue
        if remaining >= len(text):
            text_node.text = ""
            remaining -= len(text)
            continue
        text_node.text = text[remaining:]
        break


def strip_internal_hyperlink_styles(
    document_tree: ET.Element,
    relationship_targets: dict[str, str] | None = None,
) -> bool:
    changed = False
    for paragraph in document_tree.findall(".//w:p", XML_NAMESPACES):
        if flatten_internal_hyperlinks_in_paragraph(paragraph, relationship_targets or {}):
            changed = True
    for hyperlink in document_tree.findall(".//w:hyperlink", XML_NAMESPACES):
        for run_properties in hyperlink.findall(".//w:rPr", XML_NAMESPACES):
            if remove_run_style(run_properties):
                changed = True
    return changed


def flatten_internal_hyperlinks_in_paragraph(
    paragraph: ET.Element,
    relationship_targets: dict[str, str],
) -> bool:
    changed = False
    for child in list(paragraph):
        if child.tag != f"{WORD_ATTR_PREFIX}hyperlink":
            continue
        rel_id = child.get(f"{REL_ATTR_PREFIX}id")
        if rel_id and not is_internal_hyperlink_target(relationship_targets.get(rel_id)):
            continue
        insert_at = list(paragraph).index(child)
        replacement_runs: list[ET.Element] = []
        for run in child.findall("w:r", XML_NAMESPACES):
            new_run = ET.Element(f"{WORD_ATTR_PREFIX}r")
            run_properties = run.find("w:rPr", XML_NAMESPACES)
            normalized_properties = ensure_reference_run_properties(
                clone_element(run_properties) if run_properties is not None else None
            )
            if normalized_properties is not None:
                new_run.append(normalized_properties)
            for node in list(run):
                if node.tag == f"{WORD_ATTR_PREFIX}rPr":
                    continue
                new_run.append(clone_element(node))
            if len(new_run):
                replacement_runs.append(new_run)
        if not replacement_runs:
            replacement_runs.append(
                build_field_run(
                    text=get_element_text(child),
                    rpr_template=ensure_reference_run_properties(None),
                )
            )
        paragraph.remove(child)
        for offset, new_run in enumerate(replacement_runs):
            paragraph.insert(insert_at + offset, new_run)
        changed = True
    return changed


def is_internal_hyperlink_target(target: str | None) -> bool:
    if not target:
        return False
    stripped = target.strip()
    if stripped.startswith("#"):
        return True
    if any(marker in stripped for marker in (":", "/", "\\")):
        return False
    return True


def normalize_citation_display_text(text: str) -> str:
    return text.replace("\xa0", " ").strip()


def make_citation_field_shell_signature(
    source_keys: list[str],
    display_text: str,
) -> tuple[tuple[str, ...], str]:
    return tuple(source_keys), normalize_citation_display_text(display_text)


def pop_matching_citation_field_shell(
    zotero_context: ZoteroDocxContext,
    citation_targets: list[CitationTarget],
    display_text: str,
    occurrence_index: int = 1,
) -> list[ET.Element] | None:
    signature = make_citation_field_shell_signature(
        [target.source_key for target in citation_targets],
        display_text,
    )
    matching_shells = zotero_context.citation_field_shells.get(signature)
    if not matching_shells:
        return None

    shell = matching_shells[0]
    try:
        field_runs = [ET.fromstring(node_xml) for node_xml in shell.field_nodes_xml]
    except ET.ParseError:
        return None
    normalize_preserved_zotero_field_runs(
        field_runs,
        display_text,
        [target.source_key for target in citation_targets],
        occurrence_index,
    )

    matching_shells.pop(0)
    if not matching_shells:
        del zotero_context.citation_field_shells[signature]
    return field_runs


def normalize_preserved_zotero_field_runs(
    field_runs: list[ET.Element],
    display_text: str,
    source_keys: list[str],
    occurrence_index: int,
) -> None:
    instruction_nodes = [
        instruction
        for field_run in field_runs
        for instruction in field_run.findall(f"{WORD_ATTR_PREFIX}instrText")
    ]
    if not instruction_nodes:
        return

    instruction_text = "".join(instruction.text or "" for instruction in instruction_nodes)
    if ZOTERO_CITATION_INSTRUCTION_PREFIX not in instruction_text:
        return

    prefix, raw_payload = instruction_text.split(ZOTERO_CITATION_INSTRUCTION_PREFIX, 1)
    payload_text = raw_payload.strip()
    try:
        payload = json.loads(payload_text)
    except json.JSONDecodeError:
        return

    raw_citation_items = payload.get("citationItems")
    if not isinstance(raw_citation_items, list):
        raw_citation_items = []

    canonical_payload = build_canonical_zotero_citation_payload(
        display_text,
        source_keys,
        occurrence_index,
        raw_citation_items,
        payload,
    )
    updated_instruction = (
        prefix
        + ZOTERO_CITATION_INSTRUCTION_PREFIX
        + json.dumps(canonical_payload, ensure_ascii=False, separators=(",", ":"))
        + " "
    )
    if updated_instruction != instruction_text:
        rewrite_instruction_node_texts(instruction_nodes, updated_instruction)

    for field_run in field_runs:
        if field_run.tag != f"{WORD_ATTR_PREFIX}r":
            continue
        run_properties = field_run.find("w:rPr", XML_NAMESPACES)
        normalized_properties = ensure_zotero_field_run_properties(run_properties)
        if run_properties is None:
            field_run.insert(0, normalized_properties)


def rewrite_instruction_node_texts(
    instruction_nodes: list[ET.Element],
    updated_instruction: str,
) -> None:
    remaining = updated_instruction
    segment_lengths = [len(node.text or "") for node in instruction_nodes]
    for index, node in enumerate(instruction_nodes):
        if index == len(instruction_nodes) - 1:
            node.text = remaining
            break
        segment_length = segment_lengths[index]
        node.text = remaining[:segment_length]
        remaining = remaining[segment_length:]


def build_canonical_zotero_citation_payload(
    display_text: str,
    source_keys: list[str],
    occurrence_index: int,
    citation_items: list[dict[str, object]],
    raw_payload: dict | None = None,
) -> dict[str, object]:
    payload = dict(raw_payload) if isinstance(raw_payload, dict) else {}
    raw_citation_id = payload.get("citationID")
    payload["citationID"] = (
        raw_citation_id
        if is_native_style_citation_id(raw_citation_id)
        else make_stable_zotero_citation_id(source_keys, display_text, occurrence_index)
    )
    payload["properties"] = build_canonical_zotero_citation_properties(display_text, payload.get("properties"))
    payload["citationItems"] = citation_items
    payload["schema"] = payload.get("schema") or ZOTERO_CITATION_SCHEMA_URL
    return payload


def build_canonical_zotero_citation_properties(
    display_text: str,
    raw_properties: dict | None = None,
) -> dict[str, object]:
    properties = raw_properties if isinstance(raw_properties, dict) else {}
    note_index = properties.get("noteIndex")
    if not isinstance(note_index, int):
        try:
            note_index = int(note_index)
        except (TypeError, ValueError):
            note_index = 0

    unsorted = properties.get("unsorted")
    if not isinstance(unsorted, bool):
        unsorted = True

    return {
        "unsorted": unsorted,
        "formattedCitation": display_text,
        "plainCitation": plain_citation_text(display_text),
        "noteIndex": note_index,
    }


def is_native_style_citation_id(raw_citation_id: object) -> bool:
    return isinstance(raw_citation_id, str) and bool(re.fullmatch(r"[A-Za-z0-9]{8}", raw_citation_id))


def make_stable_zotero_citation_id(
    source_keys: list[str],
    display_text: str,
    occurrence_index: int,
) -> str:
    seed = json.dumps(
        {
            "source_keys": source_keys,
            "display_text": normalize_citation_display_text(display_text),
            "occurrence_index": occurrence_index,
        },
        ensure_ascii=False,
        separators=(",", ":"),
    )
    value = int.from_bytes(hashlib.sha1(seed.encode("utf-8")).digest()[:6], "big")
    chars: list[str] = []
    alphabet = ZOTERO_CITATION_ID_ALPHABET
    base = len(alphabet)
    for _ in range(8):
        value, remainder = divmod(value, base)
        chars.append(alphabet[remainder])
    return "".join(chars)


def build_zotero_citation_field_elements(
    display_text: str,
    citation_targets: list[CitationTarget],
    hyperlink_element: ET.Element,
    occurrence_index: int = 1,
) -> list[ET.Element]:
    reference_rpr = ensure_zotero_field_run_properties(first_run_properties(hyperlink_element))
    citation_items: list[dict[str, object]] = []
    source_keys: list[str] = []
    for citation_target in citation_targets:
        item_data = dict(citation_target.item_data)
        item_id = item_data.get("id", citation_target.zotero_item_key or citation_target.source_key)
        item_data["id"] = item_id
        citation_items.append(
            {
                "id": item_id,
                "uris": [citation_target.uri] if citation_target.uri else [],
                "itemData": item_data,
            }
        )
        source_keys.append(citation_target.source_key)

    payload = build_canonical_zotero_citation_payload(
        display_text,
        source_keys,
        occurrence_index,
        citation_items,
    )
    instruction = " ADDIN ZOTERO_ITEM CSL_CITATION " + json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + " "
    return [
        build_field_run(fld_char_type="begin", rpr_template=reference_rpr),
        *build_instruction_field_runs(instruction, reference_rpr),
        build_field_run(fld_char_type="separate", rpr_template=reference_rpr),
        build_field_run(text=display_text, rpr_template=reference_rpr),
        build_field_run(fld_char_type="end", rpr_template=reference_rpr),
    ]


def plain_citation_text(display_text: str) -> str:
    return display_text.replace("\xa0", " ").strip()


def build_field_run(
    *,
    text: str | None = None,
    fld_char_type: str | None = None,
    instr_text: str | None = None,
    rpr_template: ET.Element | None = None,
) -> ET.Element:
    run = ET.Element(f"{WORD_ATTR_PREFIX}r")
    if rpr_template is not None:
        run.append(clone_element(rpr_template))
    if fld_char_type is not None:
        field_char = ET.Element(f"{WORD_ATTR_PREFIX}fldChar")
        field_char.set(f"{WORD_ATTR_PREFIX}fldCharType", fld_char_type)
        run.append(field_char)
        return run
    if instr_text is not None:
        instruction = ET.Element(f"{WORD_ATTR_PREFIX}instrText")
        instruction.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        instruction.text = instr_text
        run.append(instruction)
        return run
    text_node = ET.Element(f"{WORD_ATTR_PREFIX}t")
    if text is not None and (text.startswith(" ") or text.endswith(" ")):
        text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text_node.text = text or ""
    run.append(text_node)
    return run


def build_instruction_field_runs(
    instr_text: str,
    rpr_template: ET.Element | None = None,
    chunk_size: int = FIELD_INSTRUCTION_CHUNK_SIZE,
) -> list[ET.Element]:
    return [build_field_run(instr_text=instr_text, rpr_template=rpr_template)]


def flatten_hyperlinks_in_paragraph(paragraph: ET.Element) -> bool:
    changed = False
    for child in list(paragraph):
        if child.tag != f"{WORD_ATTR_PREFIX}hyperlink":
            continue
        insert_at = list(paragraph).index(child)
        replacement_runs: list[ET.Element] = []
        for run in child.findall("w:r", XML_NAMESPACES):
            new_run = ET.Element(f"{WORD_ATTR_PREFIX}r")
            run_properties = run.find("w:rPr", XML_NAMESPACES)
            if run_properties is not None:
                cloned_properties = clone_element(run_properties)
                remove_run_style(cloned_properties)
                new_run.append(cloned_properties)
            for node in list(run):
                if node.tag == f"{WORD_ATTR_PREFIX}rPr":
                    continue
                new_run.append(clone_element(node))
            if len(new_run):
                replacement_runs.append(new_run)
        if not replacement_runs:
            replacement_runs.append(build_field_run(text=get_element_text(child)))
        paragraph.remove(child)
        for offset, new_run in enumerate(replacement_runs):
            paragraph.insert(insert_at + offset, new_run)
        changed = True
    return changed


def paragraph_content_insert_index(paragraph: ET.Element) -> int:
    for index, child in enumerate(list(paragraph)):
        if child.tag == f"{WORD_ATTR_PREFIX}pPr":
            return index + 1
    return 0


def paragraph_contains_instruction(paragraph: ET.Element, token: str) -> bool:
    for instruction in paragraph.findall(".//w:instrText", XML_NAMESPACES):
        if token in (instruction.text or ""):
            return True
    return False


def first_run_properties(element: ET.Element) -> ET.Element | None:
    run_properties = element.find(".//w:rPr", XML_NAMESPACES)
    if run_properties is None:
        return None
    cloned_properties = clone_element(run_properties)
    remove_run_style(cloned_properties)
    return cloned_properties


def ensure_reference_run_properties(run_properties: ET.Element | None) -> ET.Element:
    if run_properties is None:
        run_properties = ET.Element(f"{WORD_ATTR_PREFIX}rPr")
    remove_run_style(run_properties)
    for fonts in run_properties.findall("w:rFonts", XML_NAMESPACES):
        run_properties.remove(fonts)
    for size_tag in ("sz", "szCs"):
        for size_node in run_properties.findall(f"w:{size_tag}", XML_NAMESPACES):
            run_properties.remove(size_node)
    color = run_properties.find("w:color", XML_NAMESPACES)
    if color is None:
        color = ET.SubElement(run_properties, f"{WORD_ATTR_PREFIX}color")
    color.set(f"{WORD_ATTR_PREFIX}val", DEFAULT_ZOTERO_FIELD_COLOR)
    return run_properties


def ensure_zotero_field_run_properties(run_properties: ET.Element | None) -> ET.Element:
    return ensure_reference_run_properties(run_properties)


def clone_element(element: ET.Element) -> ET.Element:
    return ET.fromstring(ET.tostring(element, encoding="utf-8"))


def remove_run_style(run_properties: ET.Element) -> bool:
    removed = False
    for run_style in run_properties.findall("w:rStyle", XML_NAMESPACES):
        run_properties.remove(run_style)
        removed = True
    return removed


def direct_body_paragraphs(document_tree: ET.Element) -> list[ET.Element]:
    body = document_tree.find("w:body", XML_NAMESPACES)
    if body is None:
        return []
    return [child for child in body if child.tag == f"{WORD_ATTR_PREFIX}p"]


def find_bibliography_heading_index(paragraphs: list[ET.Element], bibliography_heading: str = "参考文献") -> int | None:
    for index, paragraph in enumerate(paragraphs):
        if get_paragraph_text(paragraph) == bibliography_heading:
            return index
    return None


def get_paragraph_text(paragraph: ET.Element) -> str:
    return "".join(text.text or "" for text in paragraph.findall(".//w:t", XML_NAMESPACES)).strip()


def get_element_text(element: ET.Element) -> str:
    return "".join(text.text or "" for text in element.findall(".//w:t", XML_NAMESPACES))


def get_paragraph_style_id(paragraph: ET.Element) -> str | None:
    paragraph_style = paragraph.find("./w:pPr/w:pStyle", XML_NAMESPACES)
    if paragraph_style is None:
        return None
    return paragraph_style.get(f"{WORD_ATTR_PREFIX}val")


def set_paragraph_style_id(paragraph: ET.Element, style_id: str) -> None:
    paragraph_properties = ensure_child(paragraph, "pPr")
    paragraph_style = ensure_child(paragraph_properties, "pStyle")
    paragraph_style.set(f"{WORD_ATTR_PREFIX}val", style_id)


def apply_table_hints(
    document_tree: ET.Element,
    hints: TemplateDocxHints,
    document_layout_hints: DocumentLayoutHints,
) -> bool:
    changed = False
    for index, table in enumerate(document_tree.findall(".//w:tbl", XML_NAMESPACES)):
        table_properties = ensure_child(table, "tblPr")
        table_style = table_properties.find("w:tblStyle", XML_NAMESPACES)
        if hints.table_style_id:
            if table_style is None:
                table_style = ensure_ordered_word_child(table_properties, "tblStyle", TABLE_PROPERTY_CHILD_ORDER)
            if table_style.get(f"{WORD_ATTR_PREFIX}val") != hints.table_style_id:
                table_style.set(f"{WORD_ATTR_PREFIX}val", hints.table_style_id)
                changed = True
        elif table_style is not None:
            table_properties.remove(table_style)
            changed = True

        table_look = ensure_ordered_word_child(table_properties, "tblLook", TABLE_PROPERTY_CHILD_ORDER)
        changed |= set_word_attributes(
            table_look,
            {
                "val": "0020",
                "firstRow": "1",
                "lastRow": "0",
                "firstColumn": "0",
                "lastColumn": "0",
                "noHBand": "0",
                "noVBand": "0",
            },
        )

        header_rows = 1
        header_changed, header_rows = normalize_grouped_header_rows(table)
        changed |= header_changed
        changed |= apply_three_line_table_borders(table, header_rows=header_rows)

        table_alignment = ensure_ordered_word_child(table_properties, "jc", TABLE_PROPERTY_CHILD_ORDER)
        if table_alignment.get(f"{WORD_ATTR_PREFIX}val") != "center":
            table_alignment.set(f"{WORD_ATTR_PREFIX}val", "center")
            changed = True

        if hints.table_paragraph_style_id:
            for row in table.findall("w:tr", XML_NAMESPACES):
                for cell in row.findall("w:tc", XML_NAMESPACES):
                    for paragraph in cell.findall(".//w:p", XML_NAMESPACES):
                        changed |= set_paragraph_alignment(paragraph, "left")
                        paragraph_properties = ensure_child(paragraph, "pPr")
                        paragraph_style = ensure_child(paragraph_properties, "pStyle")
                        if paragraph_style.get(f"{WORD_ATTR_PREFIX}val") != hints.table_paragraph_style_id:
                            paragraph_style.set(f"{WORD_ATTR_PREFIX}val", hints.table_paragraph_style_id)
                            changed = True
                    _ = get_grid_span(cell)
    return changed


def normalize_grouped_header_rows(table: ET.Element) -> tuple[bool, int]:
    rows = table.findall("w:tr", XML_NAMESPACES)
    if len(rows) < 2:
        return False, 1

    first_cells = rows[0].findall("w:tc", XML_NAMESPACES)
    second_cells = rows[1].findall("w:tc", XML_NAMESPACES)
    if len(first_cells) != len(second_cells):
        return False, 1

    group_specs: list[tuple[int, int]] = []
    index = 0
    while index < len(first_cells):
        if not get_table_cell_text(first_cells[index]):
            index += 1
            continue
        span = 1
        probe = index + 1
        while probe < len(first_cells) and not get_table_cell_text(first_cells[probe]):
            span += 1
            probe += 1
        if span > 1 and all(get_table_cell_text(second_cells[column]) for column in range(index, index + span)):
            group_specs.append((index, span))
        index = probe if span > 1 else index + 1

    if not group_specs:
        return False, 1

    changed = False
    grouped_columns = {column for start, span in group_specs for column in range(start, start + span)}
    for column, first_cell in enumerate(first_cells):
        if column in grouped_columns:
            continue
        if get_table_cell_text(first_cell) and not get_table_cell_text(second_cells[column]):
            changed |= set_vertical_merge(first_cell, "restart")
            changed |= set_vertical_merge(second_cells[column], "continue")

    first_row = rows[0]
    for start, span in reversed(group_specs):
        current_cells = first_row.findall("w:tc", XML_NAMESPACES)
        lead_cell = current_cells[start]
        changed |= set_grid_span(lead_cell, span)
        for continuation_cell in current_cells[start + 1 : start + span]:
            first_row.remove(continuation_cell)
            changed = True

    return changed, 2


def apply_three_line_table_borders(table: ET.Element, header_rows: int = 1) -> bool:
    changed = False
    table_properties = ensure_child(table, "tblPr")
    table_borders = ensure_ordered_word_child(table_properties, "tblBorders", TABLE_PROPERTY_CHILD_ORDER)
    for border_name, attributes in (
        (
            "top",
            {
                "val": "single",
                "sz": THREE_LINE_OUTER_BORDER_SIZE,
                "space": "0",
                "color": "auto",
            },
        ),
        ("left", {"val": "nil"}),
        (
            "bottom",
            {
                "val": "single",
                "sz": THREE_LINE_OUTER_BORDER_SIZE,
                "space": "0",
                "color": "auto",
            },
        ),
        ("right", {"val": "nil"}),
        ("insideH", {"val": "nil"}),
        ("insideV", {"val": "nil"}),
    ):
        border = ensure_ordered_word_child(table_borders, border_name, TABLE_BORDER_CHILD_ORDER)
        changed |= set_word_attributes(border, attributes)

    rows = table.findall("w:tr", XML_NAMESPACES)
    header_separator_needed = len(rows) > header_rows
    for row_index, row in enumerate(rows):
        for cell in row.findall("w:tc", XML_NAMESPACES):
            cell_properties = ensure_ordered_word_child(cell, "tcPr", CELL_PROPERTY_CHILD_ORDER)
            cell_borders = ensure_ordered_word_child(cell_properties, "tcBorders", CELL_PROPERTY_CHILD_ORDER)
            top_attributes = {"val": "nil"}
            if row_index == 0:
                top_attributes = {
                    "val": "single",
                    "sz": THREE_LINE_OUTER_BORDER_SIZE,
                    "space": "0",
                    "color": "auto",
                }
            bottom_attributes = {"val": "nil"}
            if row_index == len(rows) - 1:
                bottom_attributes = {
                    "val": "single",
                    "sz": THREE_LINE_OUTER_BORDER_SIZE,
                    "space": "0",
                    "color": "auto",
                }
            if header_separator_needed:
                if header_rows > 1 and row_index == 0 and get_grid_span(cell) > 1:
                    bottom_attributes = {
                        "val": "single",
                        "sz": THREE_LINE_HEADER_BORDER_SIZE,
                        "space": "0",
                        "color": "auto",
                    }
                elif row_index == header_rows - 1:
                    bottom_attributes = {
                        "val": "single",
                        "sz": THREE_LINE_HEADER_BORDER_SIZE,
                        "space": "0",
                        "color": "auto",
                    }
            for border_name, attributes in (
                ("top", top_attributes),
                ("left", {"val": "nil"}),
                ("bottom", bottom_attributes),
                ("right", {"val": "nil"}),
            ):
                border = ensure_ordered_word_child(cell_borders, border_name, CELL_BORDER_CHILD_ORDER)
                changed |= set_word_attributes(border, attributes)
    return changed


def apply_figure_hints(document_tree: ET.Element, document_layout_hints: DocumentLayoutHints) -> bool:
    changed = False
    for paragraph in document_tree.findall(".//w:p", XML_NAMESPACES):
        if paragraph.find(".//w:drawing", XML_NAMESPACES) is None:
            continue
        changed |= set_paragraph_alignment(paragraph, "center")
    return changed


def get_grid_span(cell: ET.Element) -> int:
    grid_span = cell.find("./w:tcPr/w:gridSpan", XML_NAMESPACES)
    if grid_span is None:
        return 1
    try:
        return max(int(grid_span.get(f"{WORD_ATTR_PREFIX}val", "1")), 1)
    except ValueError:
        return 1


def set_grid_span(cell: ET.Element, span: int) -> bool:
    if span <= 1:
        return False
    cell_properties = ensure_ordered_word_child(cell, "tcPr", CELL_PROPERTY_CHILD_ORDER)
    grid_span = ensure_ordered_word_child(cell_properties, "gridSpan", CELL_PROPERTY_CHILD_ORDER)
    return set_word_attributes(grid_span, {"val": str(span)})


def set_vertical_merge(cell: ET.Element, merge_kind: str) -> bool:
    cell_properties = ensure_ordered_word_child(cell, "tcPr", CELL_PROPERTY_CHILD_ORDER)
    vertical_merge = ensure_ordered_word_child(cell_properties, "vMerge", CELL_PROPERTY_CHILD_ORDER)
    attributes = {"val": merge_kind} if merge_kind else {}
    return set_word_attributes(vertical_merge, attributes)


def get_table_cell_text(cell: ET.Element) -> str:
    texts: list[str] = []
    for paragraph in cell.findall(".//w:p", XML_NAMESPACES):
        text = get_paragraph_text(paragraph)
        if text:
            texts.append(text)
    return " ".join(texts).strip()


def set_paragraph_alignment(paragraph: ET.Element, alignment: str) -> bool:
    paragraph_properties = ensure_child(paragraph, "pPr")
    alignment_element = ensure_child(paragraph_properties, "jc")
    alignment_value = {"left": "left", "center": "center", "right": "right"}.get(alignment)
    if alignment_value is None:
        return False
    if alignment_element.get(f"{WORD_ATTR_PREFIX}val") == alignment_value:
        return False
    alignment_element.set(f"{WORD_ATTR_PREFIX}val", alignment_value)
    return True


def ensure_child(parent: ET.Element, child_name: str) -> ET.Element:
    child = parent.find(f"w:{child_name}", XML_NAMESPACES)
    if child is not None:
        return child
    child = ET.Element(f"{WORD_ATTR_PREFIX}{child_name}")
    parent.insert(0, child)
    return child


def ensure_ordered_word_child(parent: ET.Element, child_name: str, ordered_names: tuple[str, ...]) -> ET.Element:
    child = parent.find(f"w:{child_name}", XML_NAMESPACES)
    if child is not None:
        return child
    child = ET.Element(f"{WORD_ATTR_PREFIX}{child_name}")
    insert_at = len(parent)
    try:
        child_order = ordered_names.index(child_name)
    except ValueError:
        child_order = len(ordered_names)
    for index, existing in enumerate(list(parent)):
        existing_name = word_local_name(existing.tag)
        try:
            existing_order = ordered_names.index(existing_name)
        except ValueError:
            existing_order = len(ordered_names)
        if existing_order > child_order:
            insert_at = index
            break
    parent.insert(insert_at, child)
    return child


def set_word_attributes(element: ET.Element, attributes: dict[str, str]) -> bool:
    changed = False
    desired = {f"{WORD_ATTR_PREFIX}{name}": value for name, value in attributes.items()}
    for existing_name in list(element.attrib):
        if existing_name not in desired:
            del element.attrib[existing_name]
            changed = True
    for attr_name, attr_value in desired.items():
        if element.get(attr_name) != attr_value:
            element.set(attr_name, attr_value)
            changed = True
    return changed


def word_local_name(tag: str) -> str:
    return tag.split("}", 1)[1] if tag.startswith("{") else tag
