from __future__ import annotations

import json
import os
import re
import shutil
import tempfile
import xml.etree.ElementTree as ET
from contextlib import contextmanager
from copy import deepcopy
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
WORD_ATTR_PREFIX = f"{{{WORD_NAMESPACE}}}"
REL_ATTR_PREFIX = f"{{{RELATIONSHIP_NAMESPACE}}}"
BIBLIOGRAPHY_FILENAME = "bibliography_links.tex"
REFS_BIB_FILENAME = "refs.bib"
REFS_DISPLAY_FILENAME = "refs_display.json"
PARENCITE_DEFS_FILENAME = "parencite_defs.tex"
ZOTERO_ITEMS_FILENAME = "dotex_zotero_items.json"
BIBLIOGRAPHY_SECTION_TITLES = ("参考文献", "References", "Bibliography")
CAPTION_NUMBER_PREFIX_PATTERN = re.compile(
    r"^(?P<label>图|表|figure|table)\s*(?P<number>\d+)\s*[:：.]?\s*(?P<body>.+?)\s*$",
    re.IGNORECASE,
)
ZOTERO_CITATION_FIELD_TOKEN = "ADDIN ZOTERO_ITEM CSL_CITATION"
PAGE_LAYOUT_PACKAGE = "\\usepackage[a4paper,left=2.2cm,right=2.2cm,top=2.4cm,bottom=2.4cm]{geometry}"
GRAPHICS_SUPPORT_BLOCK = r"""\makeatletter
\newlength{\dotexgraphicmaxwidth}
\setlength{\dotexgraphicmaxwidth}{0.92\linewidth}
\def\maxwidth{\ifdim\Gin@nat@width>\dotexgraphicmaxwidth\dotexgraphicmaxwidth\else\Gin@nat@width\fi}
\def\maxheight{\ifdim\Gin@nat@height>\textheight\textheight\else\Gin@nat@height\fi}
\newcommand{\dotexcapwidth}[1]{\ifdim#1>\dotexgraphicmaxwidth\dotexgraphicmaxwidth\else#1\fi}
\newcommand{\dotexcapheight}[1]{\ifdim#1>\textheight\textheight\else#1\fi}
\makeatother
"""
TABLE_LAYOUT_SUPPORT_BLOCK = r"""\newlength{\dotextablewidthbonus}
\setlength{\dotextablewidthbonus}{1.2cm}
\newlength{\dotextablewidth}
\AtBeginDocument{\setlength{\dotextablewidth}{\dimexpr\linewidth+\dotextablewidthbonus\relax}}
"""
_LONGTABLE_FULL_WIDTH_THRESHOLD = 0.95
DEFAULT_PROJECT_LATEXMKRC = """$pdf_mode = 5;
$max_repeat = 5;

$xelatex = 'xelatex -synctex=1 -interaction=nonstopmode -halt-on-error -file-line-error %O %S';

$clean_ext = 'aux bbl bcf blg fdb_latexmk fls lof log lot out run.xml synctex.gz toc xdv';
"""
DEFAULT_PROJECT_MAKEFILE = """MAIN ?= {main_stem}
TEX_IMAGE ?= texlive/texlive:latest
LATEXMK ?= latexmk
TECTONIC ?= tectonic

.PHONY: pdf local-pdf tectonic-pdf docker-pdf clean distclean shell check

pdf:
	@if command -v $(LATEXMK) >/dev/null 2>&1 && command -v xelatex >/dev/null 2>&1; then \\
		$(MAKE) local-pdf; \\
	elif command -v $(TECTONIC) >/dev/null 2>&1; then \\
		$(MAKE) tectonic-pdf; \\
	elif command -v docker >/dev/null 2>&1; then \\
		$(MAKE) docker-pdf; \\
	else \\
		printf '%s\\n' \\
			'Error: no supported LaTeX build environment found.' \\
			'Install MacTeX (latexmk + xelatex), tectonic, or Docker Desktop, then rerun make.' >&2; \\
		exit 127; \\
	fi

local-pdf:
	@if ! command -v $(LATEXMK) >/dev/null 2>&1 || ! command -v xelatex >/dev/null 2>&1; then \\
		printf '%s\\n' \\
			'Error: local build requires latexmk and xelatex.' \\
			'Install MacTeX, then rerun make local-pdf.' >&2; \\
		exit 127; \\
	fi
	$(LATEXMK) -xelatex "$(MAIN).tex"

tectonic-pdf:
	@if ! command -v $(TECTONIC) >/dev/null 2>&1; then \\
		printf '%s\\n' \\
			'Error: tectonic is not installed or not on PATH.' \\
			'Install tectonic, then rerun make tectonic-pdf.' >&2; \\
		exit 127; \\
	fi
	$(TECTONIC) -X compile "$(MAIN).tex"

docker-pdf:
	@if ! command -v docker >/dev/null 2>&1; then \\
		printf '%s\\n' \\
			'Error: docker is not installed or not on PATH.' \\
			'Install Docker Desktop, then rerun make docker-pdf.' >&2; \\
		exit 127; \\
	fi
	docker run --rm \\
		-v "$(CURDIR)":/work \\
		-w /work \\
		$(TEX_IMAGE) \\
		latexmk -xelatex "$(MAIN).tex"

clean:
	@if command -v $(LATEXMK) >/dev/null 2>&1; then \\
		$(LATEXMK) -c "$(MAIN).tex"; \\
	else \\
		rm -f "$(MAIN)".aux "$(MAIN)".bbl "$(MAIN)".bcf "$(MAIN)".blg \\
			"$(MAIN)".fdb_latexmk "$(MAIN)".fls "$(MAIN)".lof "$(MAIN)".log \\
			"$(MAIN)".lot "$(MAIN)".out "$(MAIN)".run.xml "$(MAIN)".synctex.gz \\
			"$(MAIN)".toc "$(MAIN)".xdv; \\
	fi

distclean: clean
	rm -f "$(MAIN).pdf"

shell:
	@if ! command -v docker >/dev/null 2>&1; then \\
		printf '%s\\n' \\
			'Error: docker is not installed or not on PATH.' \\
			'Install Docker Desktop, then rerun make shell.' >&2; \\
		exit 127; \\
	fi
	docker run --rm -it \\
		-v "$(CURDIR)":/work \\
		-w /work \\
		$(TEX_IMAGE) \\
		bash

check:
	@printf 'latexmk: '; command -v $(LATEXMK) || printf 'not found\\n'
	@printf 'xelatex: '; command -v xelatex || printf 'not found\\n'
	@printf 'tectonic:'; command -v $(TECTONIC) || printf ' not found\\n'
	@printf 'docker:  '; command -v docker || printf 'not found\\n'
"""


@dataclass
class DocxToTexResult:
    source_docx: Path
    output_tex: Path
    project_dir: Path
    media_dir: Path
    bibliography_path: Path | None
    citation_support_paths: list[Path]
    extracted_media_count: int
    table_count: int
    graphics_count: int
    math_count: int


@dataclass
class DocxFigure:
    target: str
    caption: str | None
    width_inches: float | None


@dataclass
class RecoveredBibliographyItem:
    tex_key: str
    source_key: str
    formatted_reference: str
    zotero_item_key: str | None
    uri: str | None
    item_data: dict


@dataclass
class RecoveredCitationCommand:
    latex_command: str
    source_keys: list[str]


@dataclass
class RecoveredCitationShell:
    source_keys: list[str]
    formatted_citation: str
    field_nodes_xml: list[str]


@dataclass
class ReverseConversionPreparation:
    math_placeholders: dict[str, str]
    citation_placeholders: dict[str, str]
    bibliography_items: list[RecoveredBibliographyItem]
    citation_shells: list[RecoveredCitationShell]


def convert_docx_to_tex(
    docx_path: Path,
    output_tex: Path,
    media_dir: Path | None = None,
    standalone: bool = True,
    plain_citation: bool = False,
    plain_ref: bool = False,
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
        preparation = prepare_docx_for_reverse_conversion(
            source_docx,
            prepared_docx,
            recover_zotero_citations=not plain_citation,
        )

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
    latex_text = normalize_converted_latex(
        latex_text,
        media_root,
        preparation.math_placeholders,
        citation_placeholders=preparation.citation_placeholders,
        preserve_refs=not plain_ref,
    )
    latex_text = ensure_fallback_figures(source_docx, latex_text, media_root)
    latex_text = ensure_latex_build_support(latex_text)
    citation_support_paths = write_citation_support_files(
        output.parent,
        preparation.bibliography_items,
        preparation.citation_shells,
    )
    if preparation.bibliography_items:
        latex_text = ensure_parencite_support(latex_text)
    latex_text, bibliography_text = split_bibliography_section(latex_text)
    output.write_text(latex_text, encoding="utf-8")
    bibliography_path = write_bibliography_companion(output.parent, bibliography_text)
    write_project_scaffold(output.parent, output.stem)

    return DocxToTexResult(
        source_docx=source_docx,
        output_tex=output,
        project_dir=output.parent,
        media_dir=media_root,
        bibliography_path=bibliography_path,
        citation_support_paths=citation_support_paths,
        extracted_media_count=count_media_files(media_root),
        table_count=count_table_environments(latex_text),
        graphics_count=latex_text.count("\\includegraphics"),
        math_count=count_math_markers(latex_text),
    )


def prepare_docx_for_reverse_conversion(
    source_docx: Path,
    prepared_docx: Path,
    recover_zotero_citations: bool = True,
) -> ReverseConversionPreparation:
    preparation = ReverseConversionPreparation(
        math_placeholders={},
        citation_placeholders={},
        bibliography_items=[],
        citation_shells=[],
    )
    with ZipFile(source_docx) as source_zip:
        archive_entries = [(info, source_zip.read(info.filename)) for info in source_zip.infolist()]

    document_xml = next((data for info, data in archive_entries if info.filename == "word/document.xml"), None)
    if document_xml is None:
        shutil.copy2(source_docx, prepared_docx)
        return preparation

    document_tree = ET.fromstring(document_xml)
    replace_omml_math_with_placeholders(document_tree, preparation.math_placeholders)
    if recover_zotero_citations:
        replace_zotero_citation_fields_with_placeholders(
            document_tree,
            preparation.citation_placeholders,
            preparation.bibliography_items,
            preparation.citation_shells,
        )
    updated_document = ET.tostring(document_tree, encoding="utf-8", xml_declaration=True)

    with ZipFile(prepared_docx, "w", compression=ZIP_DEFLATED) as target_zip:
        for info, data in archive_entries:
            if info.filename == "word/document.xml":
                target_zip.writestr(info, updated_document)
                continue
            target_zip.writestr(info, data)
    return preparation


def replace_zotero_citation_fields_with_placeholders(
    document_tree: ET.Element,
    citation_placeholders: dict[str, str],
    bibliography_items: list[RecoveredBibliographyItem],
    citation_shells: list[RecoveredCitationShell],
) -> None:
    item_index: dict[str, RecoveredBibliographyItem] = {}
    used_keys: set[str] = set()
    for paragraph in document_tree.findall(".//w:p", XML_NAMESPACES):
        replace_zotero_citation_fields_in_paragraph(
            paragraph,
            citation_placeholders,
            item_index,
            used_keys,
            citation_shells,
        )
    bibliography_items.extend(item_index.values())


def replace_zotero_citation_fields_in_paragraph(
    paragraph: ET.Element,
    citation_placeholders: dict[str, str],
    item_index: dict[str, RecoveredBibliographyItem],
    used_keys: set[str],
    citation_shells: list[RecoveredCitationShell],
) -> None:
    children = list(paragraph)
    index = 0
    while index < len(children):
        field_range = collect_word_field_range(children, index)
        if field_range is None:
            index += 1
            continue
        next_index, field_nodes, instruction_text = field_range
        payload = parse_zotero_citation_payload(instruction_text)
        if payload is None:
            index = next_index
            continue
        recovered_citation = build_recovered_citation_command(payload, item_index, used_keys)
        if recovered_citation is None or not recovered_citation.latex_command:
            index = next_index
            continue
        placeholder = f"TEXDOCXCITE{len(citation_placeholders)}TOKEN"
        citation_placeholders[placeholder] = recovered_citation.latex_command
        citation_shells.append(
            RecoveredCitationShell(
                source_keys=list(recovered_citation.source_keys),
                formatted_citation=extract_formatted_citation(payload, field_nodes),
                field_nodes_xml=[ET.tostring(node, encoding="unicode") for node in field_nodes],
            )
        )
        replacement_run = clone_run_with_text(field_nodes[0], placeholder)
        for node in field_nodes:
            paragraph.remove(node)
        paragraph.insert(index, replacement_run)
        children = list(paragraph)
        index += 1


def collect_word_field_range(
    children: list[ET.Element],
    start_index: int,
) -> tuple[int, list[ET.Element], str] | None:
    child = children[start_index]
    if child.tag != f"{WORD_ATTR_PREFIX}r":
        return None
    field_char = child.find("w:fldChar", XML_NAMESPACES)
    if field_char is None or field_char.get(f"{WORD_ATTR_PREFIX}fldCharType") != "begin":
        return None

    instruction_parts: list[str] = []
    field_nodes: list[ET.Element] = []
    index = start_index
    while index < len(children):
        current = children[index]
        field_nodes.append(current)
        if current.tag == f"{WORD_ATTR_PREFIX}r":
            for instruction in current.findall("w:instrText", XML_NAMESPACES):
                instruction_parts.append(instruction.text or "")
            current_field_char = current.find("w:fldChar", XML_NAMESPACES)
            if current_field_char is not None and current_field_char.get(f"{WORD_ATTR_PREFIX}fldCharType") == "end":
                return index + 1, field_nodes, "".join(instruction_parts)
        index += 1
    return None


def parse_zotero_citation_payload(instruction_text: str) -> dict | None:
    if ZOTERO_CITATION_FIELD_TOKEN not in instruction_text:
        return None
    payload_text = instruction_text.split(ZOTERO_CITATION_FIELD_TOKEN, 1)[1].strip()
    if not payload_text:
        return None
    try:
        payload = json.loads(payload_text)
    except json.JSONDecodeError:
        return None
    return payload if isinstance(payload, dict) else None


def build_recovered_citation_command(
    payload: dict,
    item_index: dict[str, RecoveredBibliographyItem],
    used_keys: set[str],
) -> RecoveredCitationCommand | None:
    raw_items = payload.get("citationItems")
    if not isinstance(raw_items, list) or not raw_items:
        return None

    display_parts = split_plain_citation_parts(payload, len(raw_items))
    recovered_keys: list[str] = []
    recovered_source_keys: list[str] = []
    for index, raw_item in enumerate(raw_items):
        if not isinstance(raw_item, dict):
            continue
        recovered_item = register_recovered_bibliography_item(
            raw_item,
            display_parts[index] if index < len(display_parts) else None,
            item_index,
            used_keys,
        )
        if recovered_item is None:
            continue
        recovered_keys.append(recovered_item.tex_key)
        recovered_source_keys.append(recovered_item.source_key)

    if not recovered_keys:
        return None
    command_name = "parencite" if is_parenthetical_citation(payload) else "textcite"
    return RecoveredCitationCommand(
        latex_command=f"\\{command_name}{{{','.join(recovered_keys)}}}",
        source_keys=recovered_source_keys,
    )


def extract_formatted_citation(payload: dict, field_nodes: list[ET.Element]) -> str:
    properties = payload.get("properties") if isinstance(payload.get("properties"), dict) else {}
    formatted = str(properties.get("formattedCitation") or "").replace("\xa0", " ").strip()
    if formatted:
        return formatted
    return "".join(
        text.text or ""
        for node in field_nodes
        for text in node.findall(".//w:t", XML_NAMESPACES)
    ).replace("\xa0", " ").strip()


def split_plain_citation_parts(payload: dict, item_count: int) -> list[str | None]:
    properties = payload.get("properties") if isinstance(payload.get("properties"), dict) else {}
    plain_citation = str(properties.get("plainCitation") or "").replace("\xa0", " ").strip()
    if not plain_citation:
        return [None] * item_count
    if is_parenthetical_citation(payload):
        plain_citation = strip_wrapping_citation_brackets(plain_citation)
    if item_count == 1:
        return [plain_citation.strip()] if plain_citation.strip() else [None]
    parts = [part.strip() for part in re.split(r"\s*[;；]\s*", plain_citation) if part.strip()]
    if len(parts) == item_count:
        return parts
    return [None] * item_count


def strip_wrapping_citation_brackets(text: str) -> str:
    normalized = text.strip()
    bracket_pairs = [("(", ")"), ("[", "]"), ("（", "）"), ("【", "】")]
    for opener, closer in bracket_pairs:
        if normalized.startswith(opener) and normalized.endswith(closer):
            return normalized[1:-1].strip()
    return normalized


def is_parenthetical_citation(payload: dict) -> bool:
    properties = payload.get("properties") if isinstance(payload.get("properties"), dict) else {}
    formatted = str(properties.get("formattedCitation") or "").strip()
    return bool(formatted) and formatted[0] in "(（[【" and formatted[-1] in ")）]】"


def register_recovered_bibliography_item(
    raw_item: dict,
    display_text: str | None,
    item_index: dict[str, RecoveredBibliographyItem],
    used_keys: set[str],
) -> RecoveredBibliographyItem | None:
    item_data = raw_item.get("itemData") if isinstance(raw_item.get("itemData"), dict) else {}
    if not item_data:
        return None

    uri = extract_zotero_uri(raw_item)
    zotero_item_key = extract_zotero_item_key(raw_item, uri)
    identity = build_recovered_item_identity(item_data, uri, zotero_item_key)
    existing = item_index.get(identity)
    if existing is not None:
        return existing

    formatted_reference = (display_text or synthesize_display_text(item_data)).strip()
    source_key = infer_source_key(item_data, uri, zotero_item_key)
    tex_key = allocate_tex_citation_key(item_data, zotero_item_key, used_keys)
    recovered_item = RecoveredBibliographyItem(
        tex_key=tex_key,
        source_key=source_key,
        formatted_reference=formatted_reference or tex_key,
        zotero_item_key=zotero_item_key,
        uri=uri,
        item_data=dict(item_data),
    )
    item_index[identity] = recovered_item
    return recovered_item


def extract_zotero_uri(raw_item: dict) -> str | None:
    uris = raw_item.get("uris")
    if isinstance(uris, list) and uris:
        candidate = uris[0]
        if isinstance(candidate, str) and candidate.strip():
            return candidate.strip()
    uri = raw_item.get("uri")
    if isinstance(uri, str) and uri.strip():
        return uri.strip()
    return None


def extract_zotero_item_key(raw_item: dict, uri: str | None) -> str | None:
    item_data = raw_item.get("itemData") if isinstance(raw_item.get("itemData"), dict) else {}
    raw_key = raw_item.get("key") or item_data.get("id")
    if isinstance(raw_key, str) and raw_key and not raw_key.isdigit():
        return raw_key
    if uri and "/items/" in uri:
        return uri.rsplit("/items/", 1)[-1].strip() or None
    return None


def build_recovered_item_identity(item_data: dict, uri: str | None, zotero_item_key: str | None) -> str:
    item_id = item_data.get("id")
    if item_id not in {None, ""}:
        return f"id:{item_id}"
    doi = str(item_data.get("DOI") or "").strip().lower()
    if doi:
        return f"doi:{doi}"
    url = str(item_data.get("URL") or "").strip().lower()
    if url:
        return f"url:{url}"
    if uri:
        return f"uri:{uri}"
    title = normalize_tex_key_token(str(item_data.get("title") or ""))
    year = extract_item_year(item_data) or ""
    return f"fallback:{zotero_item_key or ''}:{title}:{year}"


def infer_source_key(item_data: dict, uri: str | None, zotero_item_key: str | None) -> str:
    for field_name in ("DOI", "URL"):
        value = item_data.get(field_name)
        if isinstance(value, str) and value.strip():
            return value.strip()
    if zotero_item_key:
        return zotero_item_key
    if uri:
        return uri
    title = str(item_data.get("title") or "").strip()
    return title or "generated-reference"


def allocate_tex_citation_key(item_data: dict, zotero_item_key: str | None, used_keys: set[str]) -> str:
    base = build_tex_citation_key_base(item_data, zotero_item_key)
    candidate = base
    suffix_index = 0
    while candidate in used_keys:
        suffix_index += 1
        candidate = f"{base}{chr(ord('a') + suffix_index)}"
    used_keys.add(candidate)
    return candidate


def build_tex_citation_key_base(item_data: dict, zotero_item_key: str | None) -> str:
    authors = item_data.get("author") if isinstance(item_data.get("author"), list) else []
    author_token = ""
    if authors:
        first_author = authors[0] if isinstance(authors[0], dict) else {}
        author_token = normalize_tex_key_token(
            str(first_author.get("family") or first_author.get("literal") or first_author.get("given") or "")
        )
    year = extract_item_year(item_data) or ""
    if author_token and year:
        return f"{author_token}{year}"
    if author_token:
        return author_token
    if zotero_item_key:
        normalized_key = normalize_tex_key_token(zotero_item_key)
        if normalized_key:
            return normalized_key
    title = normalize_tex_key_token(str(item_data.get("title") or "reference"))
    return title or "reference"


def normalize_tex_key_token(value: str) -> str:
    normalized = re.sub(r"[^A-Za-z0-9]+", "", value).lower()
    return normalized[:40] or "reference"


def extract_item_year(item_data: dict) -> str | None:
    issued = item_data.get("issued")
    if isinstance(issued, dict):
        date_parts = issued.get("date-parts")
        if isinstance(date_parts, list) and date_parts and isinstance(date_parts[0], list) and date_parts[0]:
            year = date_parts[0][0]
            if year not in {None, ""}:
                return str(year)
    raw_date = str(item_data.get("date") or "")
    match = re.search(r"(19|20)\d{2}", raw_date)
    if match is not None:
        return match.group(0)
    return None


def synthesize_display_text(item_data: dict) -> str:
    authors = item_data.get("author") if isinstance(item_data.get("author"), list) else []
    year = extract_item_year(item_data)
    if authors:
        first_author = authors[0] if isinstance(authors[0], dict) else {}
        author_text = str(first_author.get("family") or first_author.get("literal") or first_author.get("given") or "").strip()
        if len(authors) > 1 and author_text:
            author_text = f"{author_text} et al."
        if author_text and year:
            return f"{author_text} {year}"
        if author_text:
            return author_text
    title = str(item_data.get("title") or "").strip()
    if title and year:
        return f"{title} {year}"
    return title or year or "reference"


def clone_run_with_text(source_run: ET.Element, text: str) -> ET.Element:
    replacement_run = ET.Element(f"{WORD_ATTR_PREFIX}r")
    properties = source_run.find("w:rPr", XML_NAMESPACES)
    if properties is not None:
        replacement_run.append(deepcopy(properties))
    text_node = ET.SubElement(replacement_run, f"{WORD_ATTR_PREFIX}t")
    text_node.text = text
    return replacement_run


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


def normalize_converted_latex(
    latex_text: str,
    media_root: Path,
    math_placeholders: dict[str, str],
    citation_placeholders: dict[str, str] | None = None,
    preserve_refs: bool = True,
) -> str:
    normalized = latex_text.replace("\\includegraphics{", "\\includegraphics[]{}")
    normalized = normalized.replace("\\includegraphics[]{}", "\\includegraphics{")
    media_pattern = re.escape(f"{media_root.name}/media/")
    normalized = re.sub(media_pattern, f"{media_root.name}/", normalized)
    normalized = normalized.replace("\r\n", "\n")
    for placeholder, latex_math in math_placeholders.items():
        normalized = normalized.replace(placeholder, f"${latex_math}$")
    for placeholder, latex_citation in (citation_placeholders or {}).items():
        normalized = normalized.replace(placeholder, latex_citation)
    if not preserve_refs:
        normalized = flatten_cross_reference_markup(normalized)
    return normalized


def flatten_cross_reference_markup(latex_text: str) -> str:
    flattened = re.sub(r"\\hyperref\[[^\]]+\]\{([^{}]+)\}", r"\1", latex_text)
    flattened = re.sub(r"\\protect\\phantomsection\\label\{_Ref[^}]+\}\{\}", "", flattened)
    flattened = re.sub(r"\\label\{_Ref[^}]+\}(?:\{\})?", "", flattened)
    return flattened


def ensure_latex_build_support(latex_text: str) -> str:
    supported = latex_text
    supported = ensure_latex_package(supported, PAGE_LAYOUT_PACKAGE)
    if "\\includegraphics" in supported:
        supported = ensure_latex_package(supported, "\\usepackage{graphicx}")
        supported = ensure_graphics_support_block(supported)
        supported = normalize_includegraphics_commands(supported)
    if "\\begin{longtable}" in supported:
        supported = ensure_table_layout_support(supported)
        supported = normalize_longtable_widths(supported)
    if "\\begin{figure}" in supported and "\\begin{longtable}" in supported:
        supported = ensure_latex_package(supported, "\\usepackage{float}")
        supported = fix_figure_placement(supported)
    if contains_cjk_text(supported):
        supported = ensure_latex_package(supported, "\\usepackage[UTF8]{ctex}")
    supported = normalize_caption_number_prefixes(supported)
    return supported


def ensure_parencite_support(latex_text: str) -> str:
    support_input = f"\\input{{{PARENCITE_DEFS_FILENAME}}}"
    if support_input in latex_text:
        return latex_text
    begin_document = latex_text.find("\\begin{document}")
    if begin_document == -1:
        return f"{support_input}\n\n{latex_text}"
    return f"{latex_text[:begin_document]}{support_input}\n\n{latex_text[begin_document:]}"


def ensure_latex_package(latex_text: str, package_line: str) -> str:
    if package_line in latex_text:
        return latex_text

    documentclass_match = re.search(r"\\documentclass(?:\[[^\]]*\])?\{[^}]+\}\n", latex_text)
    if documentclass_match is None:
        return f"{package_line}\n{latex_text}"
    insert_at = documentclass_match.end()
    return f"{latex_text[:insert_at]}{package_line}\n\n{latex_text[insert_at:]}"


def ensure_graphics_support_block(latex_text: str) -> str:
    if "\\def\\maxwidth" in latex_text and "\\dotexcapwidth" in latex_text:
        return latex_text

    graphicx_match = re.search(r"\\usepackage(?:\[[^\]]*\])?\{graphicx\}\n", latex_text)
    if graphicx_match is not None:
        insert_at = graphicx_match.end()
        return f"{latex_text[:insert_at]}{GRAPHICS_SUPPORT_BLOCK}\n{latex_text[insert_at:]}"
    return f"{GRAPHICS_SUPPORT_BLOCK}\n{latex_text}"


def ensure_table_layout_support(latex_text: str) -> str:
    if "\\dotextablewidthbonus" in latex_text:
        return latex_text

    longtable_match = re.search(r"\\usepackage(?:\[[^\]]*\])?\{[^}]*longtable[^}]*\}\n", latex_text)
    if longtable_match is not None:
        insert_at = longtable_match.end()
        return f"{latex_text[:insert_at]}{TABLE_LAYOUT_SUPPORT_BLOCK}\n{latex_text[insert_at:]}"
    return f"{TABLE_LAYOUT_SUPPORT_BLOCK}\n{latex_text}"


def normalize_includegraphics_commands(latex_text: str) -> str:
    includegraphics_pattern = re.compile(r"\\includegraphics(?:\[(?P<options>[^\]]*)\])?\{")

    def repl(match: re.Match[str]) -> str:
        options = match.group("options")
        if options is None:
            return "\\includegraphics[width=\\maxwidth,height=\\maxheight,keepaspectratio]{"

        parts = [part.strip() for part in options.split(",") if part.strip()]
        width_value: str | None = None
        height_value: str | None = None
        keep_aspect = False
        remaining: list[str] = []
        for part in parts:
            lowered = part.lower()
            if lowered.startswith("width="):
                width_value = part.split("=", 1)[1].strip()
                continue
            if lowered.startswith("height="):
                height_value = part.split("=", 1)[1].strip()
                continue
            if lowered == "keepaspectratio":
                keep_aspect = True
                continue
            remaining.append(part)

        normalized_parts = [normalize_includegraphics_width(width_value), normalize_includegraphics_height(height_value)]
        if keep_aspect or True:
            normalized_parts.append("keepaspectratio")
        normalized_parts.extend(remaining)
        return f"\\includegraphics[{','.join(normalized_parts)}]{{"

    return includegraphics_pattern.sub(repl, latex_text)


def fix_figure_placement(latex_text: str) -> str:
    """Add [H] placement specifier to figure environments that have none,
    preventing them from floating into the middle of multi-page longtables."""
    return re.sub(r"\\begin\{figure\}(?!\[)", r"\\begin{figure}[H]", latex_text)


def normalize_longtable_widths(latex_text: str) -> str:
    """Replace \\linewidth with \\dotextablewidth in longtable column specs and
    insert per-table LTleft/LTright centering before each longtable.

    Full-width tables (column fraction sum >= threshold) get extended centering
    that pushes them slightly beyond the text margins symmetrically.  Narrow
    tables (fraction sum < threshold) use the default \\fill centering so they
    remain properly centred regardless of their actual width.
    """
    begin_token = "\\begin{longtable}"
    end_token = "\\end{longtable}"
    parts: list[str] = []
    index = 0
    while True:
        start = latex_text.find(begin_token, index)
        if start == -1:
            parts.append(latex_text[index:])
            break
        end = latex_text.find(end_token, start)
        if end == -1:
            parts.append(latex_text[index:])
            break
        block_end = end + len(end_token)
        block = latex_text[start:block_end]
        widened = block.replace("(\\linewidth - ", "(\\dotextablewidth - ")
        fraction_sum = sum(
            float(f) for f in re.findall(r"\\real\{([0-9.]+)\}", block)
        )
        if fraction_sum >= _LONGTABLE_FULL_WIDTH_THRESHOLD:
            centering = (
                "\\setlength{\\LTleft}{\\dimexpr-\\dotextablewidthbonus/2\\relax}\n"
                "\\setlength{\\LTright}{\\dimexpr-\\dotextablewidthbonus/2\\relax}\n"
            )
        else:
            centering = (
                "\\setlength{\\LTleft}{\\fill}\n"
                "\\setlength{\\LTright}{\\fill}\n"
            )
        parts.append(latex_text[index:start])
        parts.append(centering)
        parts.append(widened)
        index = block_end
    return "".join(parts)


def normalize_includegraphics_width(width_value: str | None) -> str:
    if not width_value:
        return "width=\\maxwidth"
    if width_value == "\\maxwidth" or width_value.startswith("\\dotexcapwidth{"):
        return f"width={width_value}"
    return f"width=\\dotexcapwidth{{{width_value}}}"


def normalize_includegraphics_height(height_value: str | None) -> str:
    if not height_value:
        return "height=\\maxheight"
    if height_value == "\\maxheight" or height_value.startswith("\\dotexcapheight{"):
        return f"height={height_value}"
    return f"height=\\dotexcapheight{{{height_value}}}"


def normalize_caption_number_prefixes(latex_text: str) -> str:
    lines = latex_text.splitlines(keepends=True)
    return "".join(normalize_caption_line(line) for line in lines)


def normalize_caption_line(line: str) -> str:
    newline = "\n" if line.endswith("\n") else ""
    core = line[:-1] if newline else line

    bold_match = re.match(r"(?P<prefix>\s*\\caption\{\\textbf\{)(?P<content>.*?)(?P<suffix>\}\}\s*)$", core)
    if bold_match is not None:
        cleaned = strip_caption_number_prefix(bold_match.group("content"))
        return f"{bold_match.group('prefix')}{cleaned}{bold_match.group('suffix')}{newline}"

    plain_match = re.match(
        r"(?P<prefix>\s*\\caption\{(?:\\protect\\phantomsection\\label\{[^}]+\}\{\})?)(?P<content>.*?)(?P<suffix>\}\s*(?:\\tabularnewline\s*)?)$",
        core,
    )
    if plain_match is not None:
        cleaned = strip_caption_number_prefix(plain_match.group("content"))
        return f"{plain_match.group('prefix')}{cleaned}{plain_match.group('suffix')}{newline}"

    return line


def strip_caption_number_prefix(text: str) -> str:
    match = CAPTION_NUMBER_PREFIX_PATTERN.match(text.strip())
    if match is None:
        return text
    body = match.group("body").strip()
    return body or text


def contains_cjk_text(latex_text: str) -> bool:
    return re.search(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]", latex_text) is not None


def split_bibliography_section(
    latex_text: str,
    bibliography_filename: str = BIBLIOGRAPHY_FILENAME,
) -> tuple[str, str | None]:
    heading_pattern = "|".join(re.escape(title) for title in BIBLIOGRAPHY_SECTION_TITLES)
    bibliography_pattern = re.compile(
        rf"(?ms)(^\\section\*?\{{(?:{heading_pattern})\}}(?:\\label\{{[^}}]+\}})?\s*\n)(?P<body>.*?)(?=^\\end\{{document\}}\s*$|\Z)"
    )
    matches = list(bibliography_pattern.finditer(latex_text))
    if not matches:
        return latex_text, None

    match = matches[-1]
    bibliography_body = match.group("body").strip()
    if not bibliography_body:
        return latex_text, None

    replacement = f"{match.group(1)}\\input{{{bibliography_filename}}}\n\n"
    updated_latex = f"{latex_text[:match.start()]}{replacement}{latex_text[match.end():]}"
    return updated_latex, build_bibliography_companion_text(bibliography_body)


def build_bibliography_companion_text(bibliography_body: str) -> str:
    trimmed = bibliography_body.strip()
    return (
        "% Generated by dotex convert-tex.\n"
        "{\\small\n"
        "\\sloppy\n"
        "\\raggedright\n"
        "\\setlength{\\emergencystretch}{3em}\n"
        "\\setlength{\\parindent}{0pt}\n"
        "\\setlength{\\parskip}{0.5em}\n"
        f"{trimmed}\n"
        "}\n"
    )


def write_bibliography_companion(project_dir: Path, bibliography_text: str | None) -> Path | None:
    bibliography_path = project_dir / BIBLIOGRAPHY_FILENAME
    if bibliography_text is None:
        if bibliography_path.exists():
            bibliography_path.unlink()
        return None

    bibliography_path.write_text(bibliography_text, encoding="utf-8")
    return bibliography_path


def write_citation_support_files(
    project_dir: Path,
    bibliography_items: list[RecoveredBibliographyItem],
    citation_shells: list[RecoveredCitationShell] | None = None,
) -> list[Path]:
    support_paths = [
        project_dir / REFS_DISPLAY_FILENAME,
        project_dir / REFS_BIB_FILENAME,
        project_dir / ZOTERO_ITEMS_FILENAME,
        project_dir / PARENCITE_DEFS_FILENAME,
    ]
    if not bibliography_items:
        for support_path in support_paths:
            if support_path.exists():
                support_path.unlink()
        return []

    refs_display_path = project_dir / REFS_DISPLAY_FILENAME
    refs_display_path.write_text(
        json.dumps({item.tex_key: item.formatted_reference for item in bibliography_items}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    refs_bib_path = project_dir / REFS_BIB_FILENAME
    refs_bib_path.write_text(build_refs_bib_text(bibliography_items), encoding="utf-8")

    zotero_items_path = project_dir / ZOTERO_ITEMS_FILENAME
    zotero_items_path.write_text(build_zotero_items_payload(bibliography_items, citation_shells or []), encoding="utf-8")

    parencite_defs_path = project_dir / PARENCITE_DEFS_FILENAME
    parencite_defs_path.write_text(build_parencite_defs_text(bibliography_items), encoding="utf-8")
    return [refs_display_path, refs_bib_path, zotero_items_path, parencite_defs_path]


def build_refs_bib_text(bibliography_items: list[RecoveredBibliographyItem]) -> str:
    entries = [build_bibtex_entry(item) for item in bibliography_items]
    return "\n\n".join(entries) + "\n"


def build_bibtex_entry(item: RecoveredBibliographyItem) -> str:
    item_data = item.item_data
    entry_type = map_csl_type_to_bibtex(item_data.get("type"))
    fields: list[tuple[str, str]] = []
    title = str(item_data.get("title") or "").strip()
    if title:
        fields.append(("title", escape_bibtex_value(title)))
    author_text = format_bibtex_authors(item_data.get("author"))
    if author_text:
        fields.append(("author", escape_bibtex_value(author_text)))
    year = extract_item_year(item_data)
    if year:
        fields.append(("year", year))
    container_title = str(item_data.get("container-title") or "").strip()
    if container_title:
        field_name = "journal" if entry_type == "article" else "booktitle"
        fields.append((field_name, escape_bibtex_value(container_title)))
    publisher = str(item_data.get("publisher") or "").strip()
    if publisher:
        fields.append(("publisher", escape_bibtex_value(publisher)))
    volume = str(item_data.get("volume") or "").strip()
    if volume:
        fields.append(("volume", volume))
    issue = str(item_data.get("issue") or "").strip()
    if issue:
        fields.append(("number", issue))
    page = str(item_data.get("page") or "").strip()
    if page:
        fields.append(("pages", escape_bibtex_value(page)))
    doi = str(item_data.get("DOI") or "").strip()
    if doi:
        fields.append(("doi", escape_bibtex_value(doi)))
    url = str(item_data.get("URL") or "").strip()
    if url:
        fields.append(("url", escape_bibtex_value(url)))

    rendered_fields = "\n".join(f"  {name} = {{{value}}}," for name, value in fields)
    return f"@{entry_type}{{{item.tex_key},\n{rendered_fields}\n}}"


def map_csl_type_to_bibtex(item_type: object) -> str:
    mapping = {
        "article-journal": "article",
        "book": "book",
        "chapter": "incollection",
        "paper-conference": "inproceedings",
        "report": "techreport",
        "thesis": "phdthesis",
        "webpage": "misc",
    }
    return mapping.get(str(item_type or ""), "misc")


def format_bibtex_authors(authors: object) -> str:
    if not isinstance(authors, list):
        return ""
    parts: list[str] = []
    for author in authors:
        if not isinstance(author, dict):
            continue
        literal = str(author.get("literal") or "").strip()
        family = str(author.get("family") or "").strip()
        given = str(author.get("given") or "").strip()
        if literal:
            parts.append(literal)
        elif family and given:
            parts.append(f"{family}, {given}")
        elif family:
            parts.append(family)
        elif given:
            parts.append(given)
    return " and ".join(parts)


def escape_bibtex_value(value: str) -> str:
    escaped = value.replace("\\", "\\\\")
    escaped = escaped.replace("{", "\\{").replace("}", "\\}")
    return escaped


def build_zotero_items_payload(
    bibliography_items: list[RecoveredBibliographyItem],
    citation_shells: list[RecoveredCitationShell] | None = None,
) -> str:
    payload = {
        "version": 1,
        "items": [
            {
                "key": item.tex_key,
                "source_key": item.source_key,
                "formatted_reference": item.formatted_reference,
                "zotero_item_key": item.zotero_item_key,
                "uri": item.uri,
                "item_data": item.item_data,
            }
            for item in bibliography_items
        ],
    }
    if citation_shells:
        payload["citations"] = [
            {
                "source_keys": citation_shell.source_keys,
                "formatted_citation": citation_shell.formatted_citation,
                "field_nodes_xml": citation_shell.field_nodes_xml,
            }
            for citation_shell in citation_shells
        ]
    return json.dumps(payload, ensure_ascii=False, indent=2)


def build_parencite_defs_text(bibliography_items: list[RecoveredBibliographyItem]) -> str:
    registrations = "\n".join(
        f"\\dotexregistercitation{{{item.tex_key}}}{{{escape_tex_text(item.formatted_reference)}}}"
        for item in bibliography_items
    )
    return (
        "% Generated by dotex convert-tex.\n"
        "\\ExplSyntaxOn\n"
        "\\prop_new:N \\g_dotex_citation_prop\n"
        "\\cs_new_protected:Npn \\dotexregistercitation #1 #2\n"
        "  { \\prop_gput:Nnn \\g_dotex_citation_prop {#1} {#2} }\n"
        "\\cs_new:Npn \\dotexrendercitation #1\n"
        "  {\n"
        "    \\clist_set:Nn \\l_tmpa_clist {#1}\n"
        "    \\seq_clear:N \\l_tmpa_seq\n"
        "    \\clist_map_inline:Nn \\l_tmpa_clist\n"
        "      {\n"
        "        \\prop_get:NnNTF \\g_dotex_citation_prop {##1} \\l_tmpa_tl\n"
        "          { \\seq_put_right:NV \\l_tmpa_seq \\l_tmpa_tl }\n"
        "          { \\seq_put_right:Nn \\l_tmpa_seq {##1} }\n"
        "      }\n"
        "    \\seq_use:Nn \\l_tmpa_seq {; }\n"
        "  }\n"
        "\\cs_set_protected:Npn \\parencite #1 { (\\dotexrendercitation{#1}) }\n"
        "\\cs_set_protected:Npn \\citep #1 { (\\dotexrendercitation{#1}) }\n"
        "\\cs_set_protected:Npn \\cite #1 { (\\dotexrendercitation{#1}) }\n"
        "\\cs_set_protected:Npn \\textcite #1 { \\dotexrendercitation{#1} }\n"
        "\\cs_set_protected:Npn \\citet #1 { \\dotexrendercitation{#1} }\n"
        f"{registrations}\n"
        "\\ExplSyntaxOff\n"
    )


def escape_tex_text(value: str) -> str:
    replacements = {
        "\\": "\\textbackslash{}",
        "{": "\\{",
        "}": "\\}",
        "%": "\\%",
        "#": "\\#",
        "&": "\\&",
        "$": "\\$",
        "_": "\\_",
    }
    return "".join(replacements.get(char, char) for char in value)


def build_project_makefile(main_stem: str) -> str:
    return DEFAULT_PROJECT_MAKEFILE.format(main_stem=main_stem)


def write_project_scaffold(project_dir: Path, main_stem: str) -> None:
    (project_dir / "Makefile").write_text(build_project_makefile(main_stem), encoding="utf-8")
    (project_dir / ".latexmkrc").write_text(DEFAULT_PROJECT_LATEXMKRC, encoding="utf-8")


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
    resource_path = f"{media_root.name}/{Path(figure.target).name}"
    caption = strip_caption_number_prefix(figure.caption or Path(figure.target).stem)
    return (
        "\\begin{figure}\n"
        "\\centering\n"
        f"\\includegraphics[width=\\maxwidth,height=\\maxheight,keepaspectratio]{{{resource_path}}}\n"
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