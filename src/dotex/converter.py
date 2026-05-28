from __future__ import annotations

import json
import re
import xml.etree.ElementTree as ET
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import pypandoc

from dotex.zotero_resolver import (
    normalize_doi,
    normalize_url,
    parse_bibliography_entries,
    resolve_bibliography_against_zotero,
)


WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships"
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


@dataclass
class ConversionResult:
    source_tex: Path
    template_docx: Path
    normalized_source_path: Path
    output_docx: Path


@dataclass
class TemplateDocxHints:
    caption_style_id: str | None
    table_style_id: str | None
    table_paragraph_style_id: str | None
    normal_style_id: str | None
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


def convert_tex_to_docx(
    tex_path: Path,
    template_docx: Path,
    output_docx: Path,
    artifacts_dir: Path,
    bibliography_path: Path | None = None,
    bibliography_heading: str = "参考文献",
    enable_zotero: bool = False,
) -> ConversionResult:
    source_tex = tex_path.resolve()
    source_text = source_tex.read_text(encoding="utf-8")
    template = template_docx.resolve()
    output = output_docx.resolve()
    artifact_root = artifacts_dir.resolve()
    artifact_root.mkdir(parents=True, exist_ok=True)
    template_hints = infer_template_docx_hints(template)
    document_layout_hints = collect_document_layout_hints(source_text)
    zotero_context = build_zotero_docx_context(
        source_tex,
        template_hints,
        bibliography_path=bibliography_path,
        source_text=source_text,
        enable_zotero=enable_zotero,
    )

    normalized_source = artifact_root / "normal_manuscript.normalized.md"
    normalized_source.write_text(
        normalize_tex_for_pandoc(
            source_tex,
            source_text=source_text,
            length_context=document_layout_hints.length_context,
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
        template_hints,
        zotero_context,
        document_layout_hints,
        bibliography_heading=bibliography_heading,
        enable_zotero=enable_zotero,
    )

    return ConversionResult(
        source_tex=source_tex,
        template_docx=template,
        normalized_source_path=normalized_source,
        output_docx=output,
    )


def normalize_tex_for_pandoc(
    tex_path: Path,
    source_text: str | None = None,
    length_context: dict[str, float] | None = None,
) -> str:
    if source_text is None:
        source_text = tex_path.read_text(encoding="utf-8")
    body = extract_document_body(source_text)
    metadata, body = extract_front_matter(body)
    labels = parse_label_numbers(tex_path.with_suffix(".aux"))
    global CURRENT_LABELS, CURRENT_LENGTH_CONTEXT
    CURRENT_LABELS = labels
    CURRENT_LENGTH_CONTEXT = length_context or extract_length_context(source_text)

    body = replace_command_two_args(body, "litref", render_litref_link)
    body = replace_command_one_arg(body, "tabref", lambda label: render_cross_reference(label, labels, "表格"))
    body = replace_command_one_arg(body, "figref", lambda label: render_cross_reference(label, labels, "图"))
    body = replace_command_one_arg(body, "detokenize", lambda value: value)
    body = replace_generic_refs(body, labels)
    body = inline_bibliography_inputs(body, tex_path.parent)
    body = strip_layout_only_commands(body)
    body = normalize_table_syntax(body)
    body = convert_figure_blocks(body)
    body = convert_longtable_blocks(body)
    body = convert_table_blocks(body)
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


def render_cross_reference(label: str, labels: dict[str, str], prefix: str) -> str:
    number = labels.get(label)
    if not number:
        return label
    anchor = make_anchor_id(label)
    return f"[{prefix} {number}](#{anchor})"


def replace_generic_refs(text: str, labels: dict[str, str]) -> str:
    def repl(match: re.Match[str]) -> str:
        label = match.group(1)
        if label.startswith("fig:"):
            return render_cross_reference(label, labels, "图")
        if label.startswith("tab:"):
            return render_cross_reference(label, labels, "表格")
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


def inline_bibliography_inputs(body: str, base_dir: Path) -> str:
    pattern = re.compile(r"\\input\{([^}]+)\}")

    def repl(match: re.Match[str]) -> str:
        relative_path = match.group(1)
        bib_path = (base_dir / relative_path).resolve()
        if not bib_path.exists():
            return match.group(0)
        return render_bibliography_entries(bib_path)

    return pattern.sub(repl, body)


def render_bibliography_entries(bib_path: Path) -> str:
    text = bib_path.read_text(encoding="utf-8")
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
        entries.append(render_anchor_div(make_bibliography_anchor_id(source_key), entry_text.strip()))
        index = cursor
    return "\n\n".join(entries)


def render_litref_link(target: str, text: str) -> str:
    return f"[{text}](#{make_bibliography_anchor_id(target)})"


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
    anchor = make_anchor_id(label) if label else None
    caption_text = format_caption_text(caption, label, labels=None, kind="figure")
    image_attributes = render_image_attributes(anchor, options)
    caption_block = render_custom_style_block(caption_text, "caption") if caption_text else ""
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
        caption_text = format_caption_text(caption, label, labels=None, kind="table")
        if caption_text:
            return f"\n\n{render_anchor_div(label, render_custom_style_block(caption_text, 'caption'))}\n\n"
        return "\n\n"
    rows = parse_table_rows(tabular_block, env_name="tabular")
    return render_markdown_table(rows, caption, label, kind="table")


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


def parse_table_rows(block: str, env_name: str) -> list[list[str]]:
    content = strip_environment_wrapper(block, env_name)
    content = re.sub(r"\\caption\{.*?\}\\label\{.*?\}\\*", "", content, flags=re.S)
    content = re.sub(r"\\caption\{.*?\}", "", content, flags=re.S)
    content = re.sub(r"\\label\{[^}]+\}", "", content)
    content = re.sub(r"\\endfirsthead.*?\\endhead", "", content, flags=re.S)
    content = re.sub(r"\\endfoot.*?\\endlastfoot", "", content, flags=re.S)
    for token in ["\\toprule", "\\midrule", "\\bottomrule", "\\hline", "\\addlinespace"]:
        content = content.replace(token, "\n")
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
        if any(cleaned_cells):
            rows.append(cleaned_cells)

    if len(rows) >= 2 and rows[0] == rows[1]:
        rows.pop(1)
    return rows


def strip_environment_wrapper(block: str, env_name: str) -> str:
    begin_pattern = re.compile(rf"\\begin\{{{env_name}\}}(?:\[[^\]]*\])?(?:\{{[^}}]*\}})*")
    content = begin_pattern.sub("", block, count=1)
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
    try:
        count = max(int(count_text), 1)
    except ValueError:
        count = 1
    return [value.strip()] + [""] * (count - 1)


def cleanup_table_cell(cell: str) -> str:
    text = convert_inline_tex_to_markdown(cell)
    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text).strip()
    text = text.replace("|", "\\|")
    return text


def render_markdown_table(rows: list[list[str]], caption: str, label: str | None, kind: str) -> str:
    if not rows:
        caption_text = format_caption_text(caption, label, labels=None, kind=kind)
        if not caption_text:
            return "\n\n"
        return f"\n\n{render_anchor_div(label, render_custom_style_block(caption_text, 'caption'))}\n\n"

    column_count = max(len(row) for row in rows)
    padded_rows = [row + [""] * (column_count - len(row)) for row in rows]
    header = padded_rows[0]
    body_rows = padded_rows[1:] or [[""] * column_count]

    lines = [format_markdown_row(header), format_markdown_row(["---"] * column_count)]
    lines.extend(format_markdown_row(row) for row in body_rows)

    table_text = "\n".join(lines)
    caption_text = format_caption_text(caption, label, labels=None, kind=kind)
    parts: list[str] = []
    if caption_text:
        parts.append(render_custom_style_block(caption_text, "caption"))
    parts.append(table_text)
    anchored_table = render_anchor_div(label, "\n\n".join(parts))
    return f"\n\n{anchored_table}\n\n"


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
    prefix = "图" if kind == "figure" else "表格"
    if number and caption_text:
        return f"{prefix} {number} {caption_text}".strip()
    if caption_text:
        return caption_text
    if number:
        return f"{prefix} {number}"
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
        (r"\\subsubsection\*\{([^{}]+)\}", r"\n\n### \1\n"),
        (r"\\subsection\*\{([^{}]+)\}", r"\n\n## \1\n"),
        (r"\\section\*\{([^{}]+)\}", r"\n\n# \1\n"),
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
    current = text
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

    replacements = {
        "\\%": "%",
        "\\_": "_",
        "\\&": "&",
        "\\#": "#",
        "\\linewidth": "",
        "\\and": "; ",
        "~": " ",
    }
    for old, new in replacements.items():
        current = current.replace(old, new)
    current = re.sub(r"\\label\{[^}]+\}", "", current)
    return current


def convert_inline_tex_to_plain(text: str) -> str:
    current = text
    previous = None
    while current != previous:
        previous = current
        current = replace_command_two_args(current, "href", lambda url, label: convert_inline_tex_to_plain(label))
        current = replace_command_one_arg(current, "url", lambda value: value)
        current = replace_command_one_arg(current, "nolinkurl", lambda value: value)
        current = replace_command_one_arg(current, "emph", lambda value: convert_inline_tex_to_plain(value))
        current = replace_command_one_arg(current, "textbf", lambda value: convert_inline_tex_to_plain(value))
        current = replace_command_one_arg(current, "textsuperscript", lambda value: convert_inline_tex_to_plain(value))
    replacements = {
        "\\%": "%",
        "\\_": "_",
        "\\&": "&",
        "\\#": "#",
        "\\and": "; ",
        "~": " ",
    }
    for old, new in replacements.items():
        current = current.replace(old, new)
    current = re.sub(r"\\label\{[^}]+\}", "", current)
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
        if stripped.startswith("\\begin{longtable}{"):
            inside_longtable = True
            normalized_lines.append(rewrite_begin_line(line, "longtable"))
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
            if "\\bibentry" in candidate.read_text(encoding="utf-8"):
                return candidate
        except OSError:
            continue

    fallback = (source_tex.parent / "bibliography_links.tex").resolve()
    return fallback if fallback.exists() else None


def build_zotero_docx_context(
    source_tex: Path,
    template_hints: TemplateDocxHints,
    bibliography_path: Path | None = None,
    source_text: str | None = None,
    enable_zotero: bool = False,
) -> ZoteroDocxContext:
    if source_text is None:
        source_text = source_tex.read_text(encoding="utf-8")
    bibliography_path = infer_bibliography_path(source_tex, source_text, bibliography_path)
    if bibliography_path is None or not bibliography_path.exists():
        return ZoteroDocxContext([], [], {}, {}, {})

    bibliography_entries = parse_bibliography_entries(bibliography_path)
    records_by_source: dict[str, object] = {}
    csl_by_key: dict[str, dict] = {}
    unmatched_notices: list[UnmatchedZoteroNotice] = []
    if enable_zotero and DEFAULT_ZOTERO_DATABASE.exists():
        try:
            report, csl_items = resolve_bibliography_against_zotero(bibliography_path, DEFAULT_ZOTERO_DATABASE)
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
    for index, entry in enumerate(bibliography_entries, start=1):
        record = records_by_source.get(entry.source_key)
        zotero_item_key = getattr(record, "zotero_item_key", None)
        item_data = synthesize_citation_item_data(entry, record, csl_by_key, index)
        uri = None
        if zotero_item_key and template_hints.zotero_item_uri_prefix:
            uri = f"{template_hints.zotero_item_uri_prefix}{zotero_item_key}"
        elif zotero_item_key:
            uri = zotero_item_key
        anchor_id = make_bibliography_anchor_id(entry.source_key)

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


def derive_import_url(source_key: str) -> str | None:
    normalized_doi = normalize_doi(source_key)
    if normalized_doi:
        return f"https://doi.org/{normalized_doi}"
    return normalize_url(source_key)


def synthesize_citation_item_data(entry, report_record: object | None, csl_by_key: dict[str, dict], synthetic_index: int) -> dict:
    zotero_item_key = getattr(report_record, "zotero_item_key", None)
    if getattr(report_record, "matched", False) and zotero_item_key in csl_by_key:
        item_data = dict(csl_by_key[zotero_item_key])
        item_data["id"] = zotero_item_key
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
    template_hints: TemplateDocxHints,
    zotero_context: ZoteroDocxContext,
    document_layout_hints: DocumentLayoutHints,
    bibliography_heading: str = "参考文献",
    enable_zotero: bool = False,
) -> None:
    with ZipFile(output_docx) as source_zip:
        archive_entries = [(info, source_zip.read(info.filename)) for info in source_zip.infolist()]

    document_xml = next((data for info, data in archive_entries if info.filename == "word/document.xml"), None)
    if document_xml is None:
        return

    rels_xml = next((data for info, data in archive_entries if info.filename == "word/_rels/document.xml.rels"), None)
    relationship_targets = parse_document_relationships(rels_xml)

    document_tree = ET.fromstring(document_xml)
    changed = False
    if enable_zotero:
        populate_zotero_anchor_aliases_from_bibliography(document_tree, zotero_context, bibliography_heading)
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
    changed |= strip_internal_hyperlink_styles(document_tree)
    if not changed:
        return

    updated_document = ET.tostring(document_tree, encoding="utf-8", xml_declaration=True)
    temp_output = output_docx.with_suffix(".tmp.docx")
    with ZipFile(temp_output, "w", compression=ZIP_DEFLATED) as target_zip:
        for info, data in archive_entries:
            if info.filename == "word/document.xml":
                target_zip.writestr(info, updated_document)
                continue
            target_zip.writestr(info, data)
    temp_output.replace(output_docx)


def infer_template_docx_hints(template_docx: Path) -> TemplateDocxHints:
    with ZipFile(template_docx) as template_zip:
        styles_tree = ET.fromstring(template_zip.read("word/styles.xml"))
        document_tree = ET.fromstring(template_zip.read("word/document.xml"))
        raw_document_xml = template_zip.read("word/document.xml").decode("utf-8", errors="ignore")

    caption_style_id: str | None = None
    normal_style_id: str | None = None
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
        if style_type == "paragraph" and style_name.lower() == "caption":
            caption_style_id = style_id
        if style_type == "paragraph" and style_name == "Normal":
            normal_style_id = style_id
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
        bibliography_style_id=bibliography_style_id,
        zotero_item_uri_prefix=uri_prefix_match.group(0) if uri_prefix_match else None,
    )


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
    return changed


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
        first_paragraph.insert(0, build_field_run(fld_char_type="begin", rpr_template=reference_rpr))
        first_paragraph.insert(
            1,
            build_field_run(
                instr_text=' ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY ',
                rpr_template=reference_rpr,
            ),
        )
        first_paragraph.insert(2, build_field_run(fld_char_type="separate", rpr_template=reference_rpr))
        bibliography_paragraphs[-1].append(build_field_run(fld_char_type="end", rpr_template=reference_rpr))
        changed = True
    return changed


def convert_citation_hyperlinks_to_zotero_fields(
    document_tree: ET.Element,
    relationship_targets: dict[str, str],
    zotero_context: ZoteroDocxContext,
) -> bool:
    changed = False
    for paragraph in document_tree.findall(".//w:p", XML_NAMESPACES):
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
            for offset, field_run in enumerate(
                build_zotero_citation_field_elements(display_text, field_targets, reference_element)
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


def resolve_citation_hyperlink_target(
    element: ET.Element,
    relationship_targets: dict[str, str],
    zotero_context: ZoteroDocxContext,
) -> CitationTarget | None:
    if element.tag != f"{WORD_ATTR_PREFIX}hyperlink":
        return None
    anchor = element.get(f"{WORD_ATTR_PREFIX}anchor")
    if anchor:
        target = zotero_context.lookup(anchor=anchor)
        if target is not None:
            return target
        display_text = get_element_text(element)
        if looks_like_citation_display_text(display_text):
            fallback_target = synthesize_inline_citation_target(display_text, anchor)
            zotero_context.by_anchor[anchor] = fallback_target
            return fallback_target
    rel_id = element.get(f"{REL_ATTR_PREFIX}id")
    if rel_id is None:
        return None
    return zotero_context.lookup(relationship_targets.get(rel_id))


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


def strip_internal_hyperlink_styles(document_tree: ET.Element) -> bool:
    changed = False
    for hyperlink in document_tree.findall(".//w:hyperlink", XML_NAMESPACES):
        for run_properties in hyperlink.findall(".//w:rPr", XML_NAMESPACES):
            if remove_run_style(run_properties):
                changed = True
    return changed


def build_zotero_citation_field_elements(
    display_text: str,
    citation_targets: list[CitationTarget],
    hyperlink_element: ET.Element,
) -> list[ET.Element]:
    reference_rpr = first_run_properties(hyperlink_element)
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

    payload = {
        "citationID": f"cite-{abs(hash((display_text, tuple(source_keys)))) % 1_000_000_000}",
        "properties": {
            "unsorted": False,
            "formattedCitation": display_text,
            "plainCitation": display_text,
            "noteIndex": 0,
        },
        "citationItems": citation_items,
    }
    instruction = " ADDIN ZOTERO_ITEM CSL_CITATION " + json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    return [
        build_field_run(fld_char_type="begin", rpr_template=reference_rpr),
        build_field_run(instr_text=instruction, rpr_template=reference_rpr),
        build_field_run(fld_char_type="separate", rpr_template=reference_rpr),
        build_field_run(text=display_text, rpr_template=reference_rpr),
        build_field_run(fld_char_type="end", rpr_template=reference_rpr),
    ]


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
        layout_hint = document_layout_hints.tables[index] if index < len(document_layout_hints.tables) else None
        table_properties = ensure_child(table, "tblPr")
        if hints.table_style_id:
            table_style = ensure_child(table_properties, "tblStyle")
            if table_style.get(f"{WORD_ATTR_PREFIX}val") != hints.table_style_id:
                table_style.set(f"{WORD_ATTR_PREFIX}val", hints.table_style_id)
                changed = True

        table_alignment = ensure_child(table_properties, "jc")
        if table_alignment.get(f"{WORD_ATTR_PREFIX}val") != "center":
            table_alignment.set(f"{WORD_ATTR_PREFIX}val", "center")
            changed = True

        if hints.table_paragraph_style_id:
            for row in table.findall("w:tr", XML_NAMESPACES):
                column_index = 0
                for cell in row.findall("w:tc", XML_NAMESPACES):
                    alignment = None
                    if layout_hint and layout_hint.column_alignments:
                        alignment = layout_hint.column_alignments[min(column_index, len(layout_hint.column_alignments) - 1)]
                    for paragraph in cell.findall(".//w:p", XML_NAMESPACES):
                        if alignment:
                            changed |= set_paragraph_alignment(paragraph, alignment)
                        paragraph_properties = ensure_child(paragraph, "pPr")
                        paragraph_style = ensure_child(paragraph_properties, "pStyle")
                        if paragraph_style.get(f"{WORD_ATTR_PREFIX}val") != hints.table_paragraph_style_id:
                            paragraph_style.set(f"{WORD_ATTR_PREFIX}val", hints.table_paragraph_style_id)
                            changed = True
                    column_index += get_grid_span(cell)
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