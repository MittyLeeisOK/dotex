from __future__ import annotations

import argparse
import re
import tempfile
from contextlib import ExitStack, nullcontext
from dataclasses import dataclass
from importlib.resources import as_file, files
from pathlib import Path
from xml.sax.saxutils import escape as xml_escape
from zipfile import ZIP_DEFLATED, ZipFile

from dotex.tex_to_docx import (
    DEFAULT_ZOTERO_DATABASE,
    convert_tex_to_docx,
    derive_import_url,
    infer_bibliography_path,
)
from dotex.docx_to_tex import convert_docx_to_tex
from dotex.inspect_docx import TemplateManifest, inspect_template
from dotex.compare_roundtrip import build_roundtrip_comparison, render_roundtrip_report
from dotex.resolve_zotero import copied_zotero_database, resolve_bibliography_against_zotero


DEFAULT_FORMAT_SCORE_THRESHOLD = 90.0
DEFAULT_TEMPLATE_RESOURCE = files("dotex").joinpath("templates/default_reference.docx")


@dataclass(frozen=True)
class ConvertDocxRequest:
    tex_path: Path
    output_path: Path
    template_override: Path | None
    bibliography_path: Path | None
    bibliography_heading: str
    plain_citation: bool
    plain_ref: bool

    @property
    def enable_zotero(self) -> bool:
        return not self.plain_citation


@dataclass(frozen=True)
class TemplateSelection:
    path: Path
    is_builtin: bool


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="dotex",
        description="Bidirectional high-fidelity DOCX and TeX tooling for manuscript workflows.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    inspect_parser = subparsers.add_parser(
        "inspect-template",
        help="Extract styles, page settings, table styles, and Zotero traces from a DOCX template.",
    )
    inspect_parser.add_argument("docx", type=Path, help="Path to the source DOCX template.")
    inspect_parser.add_argument(
        "--output",
        type=Path,
        help="Write the manifest JSON to this path instead of stdout.",
    )
    inspect_parser.set_defaults(func=run_inspect_template)

    convert_parser = subparsers.add_parser(
        "convert-docx",
        help="Convert the manuscript TeX into a DOCX using a Word template as reference-doc.",
    )
    convert_parser.add_argument("tex", type=Path, help="Path to the source TeX manuscript.")
    convert_parser.add_argument(
        "-t",
        "--template",
        type=Path,
        help="Optional DOCX template used as pandoc reference-doc. If omitted, the toolkit uses its built-in default reference template.",
    )
    convert_parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Output DOCX path. Defaults to the TeX path with a .docx suffix.",
    )
    convert_parser.add_argument(
        "--artifacts-dir",
        type=Path,
        help=argparse.SUPPRESS,
    )
    convert_parser.add_argument(
        "--debug",
        action="store_true",
        help=argparse.SUPPRESS,
    )
    convert_parser.add_argument(
        "--bibliography",
        type=Path,
        help="Optional bibliography file containing \\bibentry definitions. Defaults to auto-detection from the TeX source.",
    )
    convert_parser.add_argument(
        "--bibliography-heading",
        type=str,
        default="参考文献",
        help="Heading text used to locate the bibliography section in the generated DOCX.",
    )
    convert_parser.add_argument(
        "--plaincitation",
        dest="plain_citation",
        action="store_true",
        default=False,
        help="Flatten citations instead of emitting editable Zotero fields.",
    )
    convert_parser.add_argument(
        "--plainref",
        dest="plain_ref",
        action="store_true",
        default=False,
        help="Flatten caption and cross-reference fields instead of preserving native REF/SEQ fields.",
    )
    convert_parser.set_defaults(func=run_convert_docx)

    convert_tex_parser = subparsers.add_parser(
        "convert-tex",
        help="Convert a DOCX manuscript into LaTeX and extract its image resources.",
    )
    convert_tex_parser.add_argument("docx", type=Path, help="Path to the source DOCX manuscript.")
    convert_tex_parser.add_argument(
        "--output",
        type=Path,
        help="Output project directory or TeX file path. Defaults to a project directory named after the DOCX stem, containing STEM.tex.",
    )
    convert_tex_parser.add_argument(
        "--media-dir",
        type=Path,
        help="Directory for extracted image resources. Defaults to OUTPUT_TEX stem plus _media.",
    )
    convert_tex_parser.add_argument(
        "--standalone",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Emit a standalone LaTeX document preamble. Disable to emit body-focused LaTeX only.",
    )
    convert_tex_parser.add_argument(
        "--plaincitation",
        dest="plain_citation",
        action="store_true",
        default=False,
        help="Flatten Zotero citation fields to visible text instead of restoring TeX cite commands.",
    )
    convert_tex_parser.add_argument(
        "--plainref",
        dest="plain_ref",
        action="store_true",
        default=False,
        help="Flatten Word REF/SEQ cross references to plain text instead of preserving cross-reference markup.",
    )
    convert_tex_parser.set_defaults(func=run_convert_tex)

    normalize_stub = subparsers.add_parser(
        "normalize-tex",
        help="Normalize manuscript-specific TeX into an intermediate representation.",
    )
    normalize_stub.set_defaults(func=run_not_implemented)

    compare_parser = subparsers.add_parser(
        "compare-roundtrip",
        help="Diagnostic report tool for checking roundtrip fidelity across tables, figures, and formulas.",
    )
    compare_parser.add_argument("source_docx", type=Path, help="Path to the original DOCX manuscript.")
    compare_parser.add_argument("source_tex", type=Path, help="Path to the source TeX manuscript.")
    compare_parser.add_argument("generated_docx", type=Path, help="Path to the generated DOCX manuscript.")
    compare_parser.add_argument(
        "--output",
        type=Path,
        help="Write the comparison report here. Defaults to generated DOCX stem plus .roundtrip-comparison.md.",
    )
    compare_parser.set_defaults(func=run_compare_roundtrip)

    zotero_parser = subparsers.add_parser(
        "resolve-zotero",
        help="Resolve bibliography_links entries against a local Zotero sqlite database.",
    )
    zotero_parser.add_argument(
        "bibliography",
        type=Path,
        help="Path to bibliography_links.tex or another file containing \\bibentry definitions.",
    )
    zotero_parser.add_argument(
        "--database",
        type=Path,
        default=Path("~/Zotero/zotero.sqlite").expanduser(),
        help="Path to the local Zotero sqlite database.",
    )
    zotero_parser.add_argument(
        "--output",
        type=Path,
        default=Path("artifacts/zotero_resolution.json"),
        help="Write the resolution report JSON here.",
    )
    zotero_parser.add_argument(
        "--csl-json",
        type=Path,
        default=Path("artifacts/zotero_library_subset.json"),
        help="Write matched Zotero items as CSL JSON here.",
    )
    zotero_parser.set_defaults(func=run_resolve_zotero)

    return parser


def run_inspect_template(args: argparse.Namespace) -> int:
    manifest = inspect_template(args.docx)
    output_path: Path | None = args.output
    if output_path is None:
        print(manifest.to_json())
        return 0

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(manifest.to_json(), encoding="utf-8")
    print(f"Wrote template manifest to {output_path}")
    return 0


def build_convert_docx_request(args: argparse.Namespace) -> ConvertDocxRequest:
    tex_path = args.tex.resolve()
    source_text = tex_path.read_text(encoding="utf-8")
    return ConvertDocxRequest(
        tex_path=tex_path,
        output_path=(args.output or default_docx_output_path(args.tex)).resolve(),
        template_override=args.template.resolve() if args.template is not None else None,
        bibliography_path=infer_bibliography_path(tex_path, source_text, args.bibliography),
        bibliography_heading=args.bibliography_heading,
        plain_citation=args.plain_citation,
        plain_ref=args.plain_ref,
    )


def resolve_template_selection(
    stack: ExitStack,
    override: Path | None,
) -> TemplateSelection:
    if override is not None:
        return TemplateSelection(path=override.resolve(), is_builtin=False)
    template_path = Path(stack.enter_context(as_file(DEFAULT_TEMPLATE_RESOURCE))).resolve()
    return TemplateSelection(path=template_path, is_builtin=True)


def run_convert_docx_conversion(
    request: ConvertDocxRequest,
    template_path: Path,
    artifacts_dir: Path,
    zotero_database: Path | None,
):
    return convert_tex_to_docx(
        tex_path=request.tex_path,
        template_docx=template_path,
        output_docx=request.output_path,
        artifacts_dir=artifacts_dir,
        bibliography_path=request.bibliography_path,
        bibliography_heading=request.bibliography_heading,
        enable_zotero=request.enable_zotero,
        use_native_bookmarks=not request.plain_ref,
        zotero_database=zotero_database,
    )


def attach_zotero_resolution_summary(
    validation_report: dict,
    request: ConvertDocxRequest,
    result,
    zotero_database: Path | None,
) -> None:
    bibliography_path = request.bibliography_path
    if bibliography_path is None or not bibliography_path.exists():
        return
    if zotero_database is None or not zotero_database.exists():
        return

    zotero_report, _csl_items = resolve_bibliography_against_zotero(
        bibliography_path,
        zotero_database,
    )
    unmatched_checklist_path = result.output_docx.with_suffix(".zotero-import-checklist.xlsx")
    should_write_checklist = zotero_report.unmatched_entries > 0
    if should_write_checklist:
        write_unmatched_import_workbook(zotero_report, unmatched_checklist_path)
    elif unmatched_checklist_path.exists():
        unmatched_checklist_path.unlink()

    validation_report["zotero_resolution"] = {
        "matched_entries": zotero_report.matched_entries,
        "unmatched_entries": zotero_report.unmatched_entries,
        "total_entries": zotero_report.total_entries,
        "complete": zotero_report.unmatched_entries == 0,
        "report_path": None,
        "csl_json_path": None,
        "unmatched_checklist_path": str(unmatched_checklist_path) if should_write_checklist else None,
    }


def build_convert_docx_validation_report(
    request: ConvertDocxRequest,
    template: TemplateSelection,
    result,
    zotero_database: Path | None,
) -> dict:
    reference_manifest = inspect_template(template.path)
    generated_manifest = inspect_template(result.output_docx)
    if template.is_builtin:
        validation_report = build_builtin_template_validation_report(
            reference_manifest,
            generated_manifest,
            expect_zotero=request.enable_zotero,
        )
    else:
        validation_report = build_validation_report(
            reference_manifest,
            generated_manifest,
            expect_zotero=request.enable_zotero,
        )

    if request.enable_zotero:
        attach_zotero_resolution_summary(validation_report, request, result, zotero_database)
    return validation_report


def emit_convert_docx_summary(template: TemplateSelection, result, validation_report: dict) -> None:
    template_mode = "builtin-default" if template.is_builtin else "reference-docx"
    print(f"Reference DOCX ({template_mode}): {template.path}")
    print(f"Wrote DOCX to {result.output_docx}")

    zotero_resolution = validation_report.get("zotero_resolution", {})
    if zotero_resolution.get("unmatched_checklist_path"):
        print(
            "Wrote unmatched Zotero checklist to "
            f"{zotero_resolution['unmatched_checklist_path']}"
        )
    if zotero_resolution.get("unmatched_entries", 0) > 0:
        checklist_path = zotero_resolution.get("unmatched_checklist_path")
        print(
            "WARNING: Zotero 缺少 "
            f"{zotero_resolution['unmatched_entries']} 条文献匹配，"
            f"建议按 {checklist_path} 导入后重新转换"
        )

    for notice in result.diagnostics.notices:
        print(f"NOTICE: {notice}")
    for warning in result.diagnostics.warnings:
        print(f"WARNING: {warning}")

    print(
        "Self-check: "
        f"format score {validation_report['format_score']:.2f}; "
        f"citation fields {validation_report['zotero_fields']['citation_field_count']}; "
        f"bibliography fields {validation_report['zotero_fields']['bibliography_field_count']}; "
        f"bookmarks {validation_report['link_artifacts']['bookmark_start_count'] + validation_report['link_artifacts']['bookmark_end_count']}; "
        f"non-hidden bookmarks {validation_report['link_artifacts']['non_hidden_bookmark_count']}; "
        f"expanded bookmarks {validation_report['link_artifacts']['expanded_bookmark_count']}; "
        f"internal hyperlinks {validation_report['link_artifacts']['internal_hyperlink_count']}; "
        f"pass={all(validation_report['passes'].values())}"
    )


def run_convert_docx(args: argparse.Namespace) -> int:
    request = build_convert_docx_request(args)
    with ExitStack() as stack:
        template = resolve_template_selection(stack, request.template_override)
        artifacts_dir = Path(stack.enter_context(tempfile.TemporaryDirectory(prefix="dotex-artifacts-"))).resolve()
        artifacts_dir.mkdir(parents=True, exist_ok=True)

        zotero_database = None
        if request.enable_zotero and DEFAULT_ZOTERO_DATABASE.exists():
            zotero_database = stack.enter_context(copied_zotero_database(DEFAULT_ZOTERO_DATABASE))

        result = run_convert_docx_conversion(
            request,
            template.path,
            artifacts_dir,
            zotero_database,
        )
        validation_report = build_convert_docx_validation_report(
            request,
            template,
            result,
            zotero_database,
        )
        emit_convert_docx_summary(template, result, validation_report)
        return 0 if all(validation_report["passes"].values()) else 1


def default_docx_output_path(tex_path: Path) -> Path:
    return tex_path.with_suffix(".docx")


def default_tex_output_path(docx_path: Path) -> Path:
    project_dir = docx_path.with_suffix("")
    return project_dir / f"{docx_path.stem}.tex"


def resolve_convert_tex_output_path(docx_path: Path, output: Path | None) -> Path:
    if output is None:
        return default_tex_output_path(docx_path)
    if output.suffix.lower() == ".tex":
        return output
    return output / f"{docx_path.stem}.tex"


def run_convert_tex(args: argparse.Namespace) -> int:
    output_path = resolve_convert_tex_output_path(args.docx, args.output)
    result = convert_docx_to_tex(
        docx_path=args.docx,
        output_tex=output_path,
        media_dir=args.media_dir,
        standalone=args.standalone,
        plain_citation=args.plain_citation,
        plain_ref=args.plain_ref,
    )

    print(f"Wrote LaTeX project to {result.project_dir}")
    print(f"Wrote TeX to {result.output_tex}")
    if result.bibliography_path is not None:
        print(f"Wrote bibliography companion to {result.bibliography_path}")
    print(f"Wrote extracted media to {result.media_dir}")
    print(
        "Structure summary: "
        f"tables {result.table_count}; "
        f"graphics {result.graphics_count}; "
        f"math markers {result.math_count}; "
        f"media files {result.extracted_media_count}"
    )
    return 0


def run_compare_roundtrip(args: argparse.Namespace) -> int:
    comparison = build_roundtrip_comparison(
        source_docx_path=args.source_docx,
        source_tex_path=args.source_tex,
        generated_docx_path=args.generated_docx,
    )
    report_text = render_roundtrip_report(comparison)
    output_path = args.output or args.generated_docx.with_suffix(".roundtrip-comparison.md")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(report_text, encoding="utf-8")
    print(f"Wrote roundtrip comparison report to {output_path}")
    return 0


def run_resolve_zotero(args: argparse.Namespace) -> int:
    with copied_zotero_database(args.database) as zotero_database:
        report, csl_items = resolve_bibliography_against_zotero(args.bibliography, zotero_database)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(report.to_json(), encoding="utf-8")

    args.csl_json.parent.mkdir(parents=True, exist_ok=True)
    args.csl_json.write_text(json_dumps(csl_items), encoding="utf-8")

    print(
        "Resolved Zotero entries: "
        f"{report.matched_entries}/{report.total_entries} matched; "
        f"{report.unmatched_entries} unmatched"
    )
    print(f"Wrote resolution report to {args.output}")
    print(f"Wrote CSL JSON subset to {args.csl_json}")
    return 0


def run_not_implemented(args: argparse.Namespace) -> int:
    raise SystemExit(f"Command '{args.command}' is reserved but not implemented yet.")


def json_dumps(payload: object) -> str:
    import json

    return json.dumps(payload, ensure_ascii=False, indent=2)


def write_unmatched_import_workbook(report, output_path: Path) -> None:
    unmatched_records = [record for record in report.records if not record.matched]
    rows = [["Title", "Formatted reference", "Source key", "Import URL"]]
    if not unmatched_records:
        rows.append(["No unmatched entries.", "", "", ""])
    else:
        for record in unmatched_records:
            rows.append(
                [
                    record.parsed_title or record.formatted_reference,
                    record.formatted_reference,
                    record.source_key,
                    derive_import_url(record.source_key) or "",
                ]
            )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(output_path, "w", compression=ZIP_DEFLATED) as workbook:
        workbook.writestr("[Content_Types].xml", build_xlsx_content_types_xml())
        workbook.writestr("_rels/.rels", build_xlsx_root_relationships_xml())
        workbook.writestr("xl/workbook.xml", build_xlsx_workbook_xml())
        workbook.writestr("xl/_rels/workbook.xml.rels", build_xlsx_workbook_relationships_xml())
        workbook.writestr("xl/styles.xml", build_xlsx_styles_xml())
        workbook.writestr("xl/worksheets/sheet1.xml", build_xlsx_sheet_xml(rows))


def build_xlsx_content_types_xml() -> str:
    return """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">
    <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>
    <Default Extension=\"xml\" ContentType=\"application/xml\"/>
    <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>
    <Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>
    <Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>
</Types>
"""


def build_xlsx_root_relationships_xml() -> str:
    return """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
    <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>
</Relationships>
"""


def build_xlsx_workbook_xml() -> str:
    return """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">
    <sheets>
        <sheet name=\"Unmatched Zotero\" sheetId=\"1\" r:id=\"rId1\"/>
    </sheets>
</workbook>
"""


def build_xlsx_workbook_relationships_xml() -> str:
    return """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">
    <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>
    <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>
</Relationships>
"""


def build_xlsx_styles_xml() -> str:
    return """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
    <fonts count=\"1\">
        <font>
            <sz val=\"11\"/>
            <name val=\"Calibri\"/>
            <family val=\"2\"/>
        </font>
    </fonts>
    <fills count=\"2\">
        <fill><patternFill patternType=\"none\"/></fill>
        <fill><patternFill patternType=\"gray125\"/></fill>
    </fills>
    <borders count=\"1\">
        <border><left/><right/><top/><bottom/><diagonal/></border>
    </borders>
    <cellStyleXfs count=\"1\">
        <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>
    </cellStyleXfs>
    <cellXfs count=\"1\">
        <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>
    </cellXfs>
    <cellStyles count=\"1\">
        <cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>
    </cellStyles>
</styleSheet>
"""


def build_xlsx_sheet_xml(rows: list[list[str]]) -> str:
    last_column = excel_column_name(max((len(row) for row in rows), default=1))
    last_row = max(len(rows), 1)
    row_xml = []
    for row_index, row in enumerate(rows, start=1):
        cells = []
        for column_index, value in enumerate(row, start=1):
            if value == "":
                continue
            cell_ref = f"{excel_column_name(column_index)}{row_index}"
            cells.append(
                f'<c r="{cell_ref}" t="inlineStr">{build_xlsx_inline_string(value)}</c>'
            )
        row_xml.append(f'<row r="{row_index}">{"".join(cells)}</row>')

    return (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<dimension ref="A1:{last_column}{last_row}"/>'
        '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        '<cols>'
        '<col min="1" max="1" width="28" customWidth="1"/>'
        '<col min="2" max="2" width="80" customWidth="1"/>'
        '<col min="3" max="3" width="42" customWidth="1"/>'
        '<col min="4" max="4" width="42" customWidth="1"/>'
        '</cols>'
        '<sheetData>'
        f'{"".join(row_xml)}'
        '</sheetData>'
        f'<autoFilter ref="A1:{last_column}{last_row}"/>'
        '</worksheet>'
    )


def build_xlsx_inline_string(value: str) -> str:
    sanitized = sanitize_xlsx_text(value)
    preserve = ' xml:space="preserve"' if sanitized != sanitized.strip() or "\n" in sanitized else ""
    return f'<is><t{preserve}>{xml_escape(sanitized)}</t></is>'


def sanitize_xlsx_text(value: str) -> str:
    return "".join(character if character in "\t\n\r" or ord(character) >= 32 else " " for character in value)


def excel_column_name(index: int) -> str:
    letters: list[str] = []
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters.append(chr(ord("A") + remainder))
    return "".join(reversed(letters))


def build_validation_report(
    reference_manifest: TemplateManifest,
    generated_manifest: TemplateManifest,
    expect_zotero: bool = True,
) -> dict:
    reference_paragraphs = usage_map(reference_manifest.document["paragraph_style_usage"])
    generated_paragraphs = usage_map(generated_manifest.document["paragraph_style_usage"])

    section_similarity = 1.0 if reference_manifest.document["sections"] == generated_manifest.document["sections"] else 0.0
    table_style_similarity = (
        1.0
        if reference_manifest.document["table_style_usage"] == generated_manifest.document["table_style_usage"]
        or uses_direct_table_formatting(generated_manifest.document["table_style_usage"])
        else 0.0
    )

    reference_caption_count = count_style(reference_manifest.document["caption_samples"], "caption")
    generated_caption_count = count_style(generated_manifest.document["caption_samples"], "caption")
    caption_count_similarity = 1 - abs(reference_caption_count - generated_caption_count) / max(reference_caption_count, 1)

    body_style_similarity = relative_similarity(
        reference_paragraphs,
        generated_paragraphs,
        {"Normal", "heading 1", "heading 2", "Title", "caption", "表格正文"},
    )
    bibliography_styles = {
        style
        for style, count in reference_paragraphs.items()
        if count > 0 and re.search(r"(书目|bibliography|references?)", style, re.IGNORECASE)
    }
    bibliography_style_presence = style_presence(bibliography_styles, generated_paragraphs)
    if not bibliography_styles:
        bibliography_style_presence = 1.0
    table_paragraph_style_similarity = (
        1.0
        if generated_paragraphs.get("表格正文", 0) >= reference_paragraphs.get("表格正文", 0) * 0.9
        else 0.0
    )
    zotero_presence = 1.0 if generated_manifest.zotero.get("detected") else 0.0
    if not expect_zotero:
        zotero_presence = 1.0

    weights = {
        "sections": 0.20,
        "table_styles": 0.20,
        "caption_style_count": 0.10,
        "body_styles": 0.20,
        "bibliography_style_presence": 0.10,
        "table_paragraph_style": 0.10,
        "zotero_presence": 0.10,
    }
    format_score = 100 * (
        section_similarity * weights["sections"]
        + table_style_similarity * weights["table_styles"]
        + max(0.0, caption_count_similarity) * weights["caption_style_count"]
        + body_style_similarity * weights["body_styles"]
        + bibliography_style_presence * weights["bibliography_style_presence"]
        + table_paragraph_style_similarity * weights["table_paragraph_style"]
        + zotero_presence * weights["zotero_presence"]
    )

    citation_field_count = int(generated_manifest.zotero.get("citation_field_count", 0))
    bibliography_field_count = int(generated_manifest.zotero.get("bibliography_field_count", 0))
    link_artifacts = generated_manifest.document.get("link_artifacts", {})
    bookmark_count = int(link_artifacts.get("bookmark_start_count", 0)) + int(link_artifacts.get("bookmark_end_count", 0))
    internal_hyperlink_count = int(link_artifacts.get("internal_hyperlink_count", 0))
    non_hidden_bookmark_count = int(link_artifacts.get("non_hidden_bookmark_count", 0))
    expanded_bookmark_count = int(link_artifacts.get("expanded_bookmark_count", 0))
    bookmark_pass = non_hidden_bookmark_count == 0
    internal_hyperlink_pass = internal_hyperlink_count == 0 if expect_zotero else True

    return {
        "reference_docx": reference_manifest.source["docx_path"],
        "generated_docx": generated_manifest.source["docx_path"],
        "template_mode": "custom-reference-docx",
        "format_score": round(format_score, 2),
        "weights": weights,
        "components": {
            "sections_similarity": round(section_similarity, 4),
            "table_style_similarity": round(table_style_similarity, 4),
            "caption_count_similarity": round(max(0.0, caption_count_similarity), 4),
            "body_style_similarity": round(body_style_similarity, 4),
            "bibliography_style_presence": round(bibliography_style_presence, 4),
            "table_paragraph_style_similarity": round(table_paragraph_style_similarity, 4),
            "zotero_presence": round(zotero_presence, 4),
        },
        "counts": {
            "reference_caption_count": reference_caption_count,
            "generated_caption_count": generated_caption_count,
            "reference_paragraph_styles": reference_paragraphs,
            "generated_paragraph_styles": generated_paragraphs,
        },
        "zotero_fields": {
            "detected": generated_manifest.zotero.get("detected", False),
            "citation_field_count": citation_field_count,
            "bibliography_field_count": bibliography_field_count,
        },
        "link_artifacts": link_artifacts,
        "residual_run_styles": generated_manifest.document["run_style_usage"],
        "passes": {
            "format_score_threshold": format_score >= DEFAULT_FORMAT_SCORE_THRESHOLD,
            "zotero_detected": (not expect_zotero) or bool(generated_manifest.zotero.get("detected")),
            "citation_fields_present": (not expect_zotero) or citation_field_count > 0,
            "bibliography_field_present": (not expect_zotero) or bibliography_field_count > 0,
            "bookmark_policy": bookmark_pass,
            "internal_hyperlink_policy": internal_hyperlink_pass,
            "bibliography_style_present": bibliography_style_presence == 1.0,
        },
    }


def build_builtin_template_validation_report(
    reference_manifest: TemplateManifest,
    generated_manifest: TemplateManifest,
    expect_zotero: bool = True,
) -> dict:
    reference_paragraphs = usage_map(reference_manifest.document["paragraph_style_usage"])
    generated_paragraphs = usage_map(generated_manifest.document["paragraph_style_usage"])
    reference_table_styles = usage_map(reference_manifest.document["table_style_usage"])
    generated_table_styles = usage_map(generated_manifest.document["table_style_usage"])

    section_similarity = 1.0 if reference_manifest.document["sections"] == generated_manifest.document["sections"] else 0.0

    required_paragraph_styles = {style for style, count in reference_paragraphs.items() if count > 0}
    paragraph_style_presence = style_presence(required_paragraph_styles, generated_paragraphs)

    required_table_styles = {style for style, count in reference_table_styles.items() if count > 0}
    table_style_presence = (
        1.0
        if uses_direct_table_formatting(generated_manifest.document["table_style_usage"])
        else style_presence(required_table_styles, generated_table_styles)
    )

    bibliography_styles = {style for style in required_paragraph_styles if re.search(r"(书目|bibliography|references?)", style, re.IGNORECASE)}
    bibliography_style_presence = style_presence(bibliography_styles, generated_paragraphs)
    if not bibliography_styles:
        bibliography_style_presence = 1.0

    zotero_presence = 1.0 if generated_manifest.zotero.get("detected") else 0.0
    if not expect_zotero:
        zotero_presence = 1.0

    weights = {
        "sections": 0.40,
        "required_paragraph_styles": 0.25,
        "required_table_styles": 0.15,
        "bibliography_style_presence": 0.10,
        "zotero_presence": 0.10,
    }
    format_score = 100 * (
        section_similarity * weights["sections"]
        + paragraph_style_presence * weights["required_paragraph_styles"]
        + table_style_presence * weights["required_table_styles"]
        + bibliography_style_presence * weights["bibliography_style_presence"]
        + zotero_presence * weights["zotero_presence"]
    )

    citation_field_count = int(generated_manifest.zotero.get("citation_field_count", 0))
    bibliography_field_count = int(generated_manifest.zotero.get("bibliography_field_count", 0))
    link_artifacts = generated_manifest.document.get("link_artifacts", {})
    bookmark_count = int(link_artifacts.get("bookmark_start_count", 0)) + int(link_artifacts.get("bookmark_end_count", 0))
    internal_hyperlink_count = int(link_artifacts.get("internal_hyperlink_count", 0))
    non_hidden_bookmark_count = int(link_artifacts.get("non_hidden_bookmark_count", 0))
    expanded_bookmark_count = int(link_artifacts.get("expanded_bookmark_count", 0))
    bookmark_pass = non_hidden_bookmark_count == 0
    internal_hyperlink_pass = internal_hyperlink_count == 0 if expect_zotero else True

    return {
        "reference_docx": reference_manifest.source["docx_path"],
        "generated_docx": generated_manifest.source["docx_path"],
        "template_mode": "builtin-default",
        "format_score": round(format_score, 2),
        "weights": weights,
        "components": {
            "sections_similarity": round(section_similarity, 4),
            "required_paragraph_style_presence": round(paragraph_style_presence, 4),
            "required_table_style_presence": round(table_style_presence, 4),
            "bibliography_style_presence": round(bibliography_style_presence, 4),
            "zotero_presence": round(zotero_presence, 4),
        },
        "counts": {
            "reference_required_paragraph_styles": sorted(required_paragraph_styles),
            "generated_paragraph_styles": generated_paragraphs,
            "reference_required_table_styles": sorted(required_table_styles),
            "generated_table_styles": generated_table_styles,
        },
        "zotero_fields": {
            "detected": generated_manifest.zotero.get("detected", False),
            "citation_field_count": citation_field_count,
            "bibliography_field_count": bibliography_field_count,
        },
        "link_artifacts": link_artifacts,
        "residual_run_styles": generated_manifest.document["run_style_usage"],
        "passes": {
            "format_score_threshold": format_score >= DEFAULT_FORMAT_SCORE_THRESHOLD,
            "zotero_detected": (not expect_zotero) or bool(generated_manifest.zotero.get("detected")),
            "citation_fields_present": (not expect_zotero) or citation_field_count > 0,
            "bibliography_field_present": (not expect_zotero) or bibliography_field_count > 0,
            "bookmark_policy": bookmark_pass,
            "internal_hyperlink_policy": internal_hyperlink_pass,
            "bibliography_style_present": bibliography_style_presence == 1.0,
        },
    }


def usage_map(style_usage: list[dict]) -> dict[str, int]:
    return {str(item.get("style_name")): int(item.get("count", 0)) for item in style_usage}


def uses_direct_table_formatting(style_usage: list[dict]) -> bool:
    return bool(style_usage) and all(
        str(item.get("style_name")) == "<direct-formatting-or-none>" for item in style_usage
    )


def count_style(samples: list[dict], style_name: str) -> int:
    return sum(1 for item in samples if item.get("style_name") == style_name)


def style_presence(required_styles: set[str], generated_styles: dict[str, int]) -> float:
    if not required_styles:
        return 1.0
    present = sum(1 for style in required_styles if generated_styles.get(style, 0) > 0)
    return present / len(required_styles)


def relative_similarity(reference: dict[str, int], generated: dict[str, int], keys: set[str]) -> float:
    total = max(
        sum(reference.get(key, 0) for key in keys),
        sum(generated.get(key, 0) for key in keys),
        1,
    )
    difference = sum(abs(reference.get(key, 0) - generated.get(key, 0)) for key in keys)
    return max(0.0, 1 - difference / (2 * total))


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)
