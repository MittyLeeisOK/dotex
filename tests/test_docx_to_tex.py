import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZIP_DEFLATED, ZipFile

from dotex.docx_to_tex import (
    RecoveredBibliographyItem,
    RecoveredCitationShell,
    build_project_makefile,
    ensure_latex_build_support,
    normalize_converted_latex,
    prepare_docx_for_reverse_conversion,
    split_bibliography_section,
    write_citation_support_files,
)


class DocxToTexTests(unittest.TestCase):
    def test_normalize_converted_latex_preserves_tightlist_macro(self) -> None:
        latex_text = (
            "\\providecommand{\\tightlist}{%\n"
            "  \\setlength{\\itemsep}{0pt}\\setlength{\\parskip}{0pt}}\n"
            "\\tightlist\n"
        )

        normalized = normalize_converted_latex(
            latex_text,
            media_root=__import__("pathlib").Path("demo_media"),
            math_placeholders={},
        )

        self.assertIn("\\providecommand{\\tightlist}{%", normalized)
        self.assertIn("\\tightlist", normalized)
        self.assertNotIn("\\providecommand{}", normalized)

    def test_prepare_docx_for_reverse_conversion_recovers_zotero_citation_fields(self) -> None:
        with TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source_docx = root / "source.docx"
            prepared_docx = root / "prepared.docx"
            document_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:r><w:fldChar w:fldCharType="begin"/></w:r>
            <w:r><w:instrText xml:space="preserve"> ADDIN ZOTERO_ITEM CSL_CITATION {"citationItems":[{"uris":["http://zotero.org/users/local/items/ABCD1234"],"itemData":{"id":123,"type":"article-journal","title":"Demo title","author":[{"family":"Reeve","given":"John"}],"issued":{"date-parts":[[2009]]},"DOI":"10.1000/demo"}}],"properties":{"formattedCitation":"(Reeve 2009)","plainCitation":"(Reeve 2009)"}} </w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate"/></w:r>
            <w:r><w:t>(Reeve 2009)</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
        </w:p>
    </w:body>
</w:document>
"""
            with ZipFile(source_docx, "w", compression=ZIP_DEFLATED) as archive:
                archive.writestr("word/document.xml", document_xml)

            preparation = prepare_docx_for_reverse_conversion(source_docx, prepared_docx)

        self.assertEqual(len(preparation.citation_placeholders), 1)
        self.assertEqual(list(preparation.citation_placeholders.values()), ["\\parencite{reeve2009}"])
        self.assertEqual(len(preparation.bibliography_items), 1)
        self.assertEqual(preparation.bibliography_items[0].formatted_reference, "Reeve 2009")
        self.assertEqual(preparation.bibliography_items[0].source_key, "10.1000/demo")
        self.assertEqual(len(preparation.citation_shells), 1)
        self.assertEqual(preparation.citation_shells[0].formatted_citation, "(Reeve 2009)")
        self.assertEqual(preparation.citation_shells[0].source_keys, ["10.1000/demo"])
        self.assertTrue(any('fldCharType="begin"' in node for node in preparation.citation_shells[0].field_nodes_xml))

    def test_split_bibliography_section_writes_input_placeholder(self) -> None:
        latex_text = (
            "\\section{正文}\n"
            "前文内容。\n\n"
            "\\section{参考文献}\\label{refs}\n\n"
            "First reference.\n\n"
            "Second reference.\n\n"
            "\\end{document}\n"
        )

        updated_latex, bibliography_text = split_bibliography_section(latex_text)

        self.assertIn("\\input{bibliography_links.tex}", updated_latex)
        self.assertNotIn("First reference.", updated_latex)
        self.assertIn("First reference.", bibliography_text)
        self.assertIn("Second reference.", bibliography_text)
        self.assertTrue(updated_latex.rstrip().endswith("\\end{document}"))

    def test_build_project_makefile_uses_main_stem(self) -> None:
        makefile_text = build_project_makefile("demo_manuscript")

        self.assertIn("MAIN ?= demo_manuscript", makefile_text)
        self.assertIn('$(LATEXMK) -xelatex "$(MAIN).tex"', makefile_text)
        self.assertIn("\npdf:\n\t@if command -v $(LATEXMK)", makefile_text)
        self.assertIn("printf '%s\\n' \\", makefile_text)

    def test_normalize_converted_latex_restores_citations_and_flattens_refs(self) -> None:
        latex_text = (
            "TEXDOCXCITE0TOKEN\n"
            "\\hyperref[_Ref123]{图 1}\n"
            "\\caption{\\protect\\phantomsection\\label{_Ref123}{}图 1 示例}\n"
        )

        normalized = normalize_converted_latex(
            latex_text,
            media_root=Path("demo_media"),
            math_placeholders={},
            citation_placeholders={"TEXDOCXCITE0TOKEN": "\\parencite{reeve2009}"},
            preserve_refs=False,
        )

        self.assertIn("\\parencite{reeve2009}", normalized)
        self.assertIn("图 1", normalized)
        self.assertNotIn("\\hyperref[_Ref123]", normalized)
        self.assertNotIn("\\label{_Ref123}", normalized)

    def test_write_citation_support_files_writes_refs_bib_and_display_map(self) -> None:
        with TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            support_paths = write_citation_support_files(
                root,
                [
                    RecoveredBibliographyItem(
                        tex_key="reeve2009",
                        source_key="10.1000/demo",
                        formatted_reference="Reeve 2009",
                        zotero_item_key="ABCD1234",
                        uri="http://zotero.org/users/local/items/ABCD1234",
                        item_data={
                            "id": 123,
                            "type": "article-journal",
                            "title": "Demo title",
                            "author": [{"family": "Reeve", "given": "John"}],
                            "issued": {"date-parts": [[2009]]},
                            "DOI": "10.1000/demo",
                        },
                    )
                ],
                [
                    RecoveredCitationShell(
                        source_keys=["10.1000/demo"],
                        formatted_citation="(Reeve 2009)",
                        field_nodes_xml=["<w:r xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:fldChar w:fldCharType=\"begin\"/></w:r>"],
                    )
                ],
            )

            refs_display = (root / "refs_display.json").read_text(encoding="utf-8")
            refs_bib = (root / "refs.bib").read_text(encoding="utf-8")
            zotero_items = (root / "dotex_zotero_items.json").read_text(encoding="utf-8")
            parencite_defs = (root / "parencite_defs.tex").read_text(encoding="utf-8")

        self.assertEqual(len(support_paths), 4)
        self.assertIn('"reeve2009": "Reeve 2009"', refs_display)
        self.assertIn("@article{reeve2009,", refs_bib)
        self.assertIn("10.1000/demo", refs_bib)
        self.assertIn('"source_key": "10.1000/demo"', zotero_items)
        self.assertIn('"citations": [', zotero_items)
        self.assertIn('"formatted_citation": "(Reeve 2009)"', zotero_items)
        self.assertIn("\\dotexregistercitation{reeve2009}{Reeve 2009}", parencite_defs)

    def test_normalize_converted_latex_adds_graphicx_and_ctex_when_needed(self) -> None:
        latex_text = (
            "\\documentclass[]{article}\n"
            "中文内容。\n"
            "\\includegraphics{demo_media/figure.png}\n"
            "\\begin{document}\n"
            "\\end{document}\n"
        )

        normalized = normalize_converted_latex(
            latex_text,
            media_root=Path("demo_media"),
            math_placeholders={},
        )
        normalized = ensure_latex_build_support(normalized)

        self.assertIn("\\usepackage{graphicx}", normalized)
        self.assertIn("\\usepackage[a4paper,left=2.2cm,right=2.2cm,top=2.4cm,bottom=2.4cm]{geometry}", normalized)
        self.assertIn("\\usepackage[UTF8]{ctex}", normalized)
        self.assertEqual(normalized.count("\\usepackage{graphicx}"), 1)
        self.assertEqual(normalized.count("\\usepackage[UTF8]{ctex}"), 1)

    def test_ensure_latex_build_support_adds_graphicx_after_late_figure_insertion(self) -> None:
        latex_text = (
            "\\documentclass[]{article}\n"
            "\\begin{document}\n"
            "中文内容。\n"
            "\\includegraphics{demo_media/late-figure.png}\n"
            "\\end{document}\n"
        )

        supported = ensure_latex_build_support(latex_text)

        self.assertIn("\\usepackage{graphicx}", supported)
        self.assertIn("0.92\\linewidth", supported)
        self.assertIn("\\usepackage[UTF8]{ctex}", supported)

    def test_ensure_latex_build_support_normalizes_caption_prefixes(self) -> None:
        latex_text = (
            "\\documentclass[]{article}\n"
            "\\usepackage{graphicx}\n"
            "\\begin{document}\n"
            "\\caption{\\protect\\phantomsection\\label{_Ref1}{}表 1 组别分配}\\tabularnewline\n"
            "\\caption{\\textbf{图 3 基于SDT的AI教学话语策略与AI默认输出模式的比较}}\n"
            "\\end{document}\n"
        )

        supported = ensure_latex_build_support(latex_text)

        self.assertIn("\\caption{\\protect\\phantomsection\\label{_Ref1}{}组别分配}\\tabularnewline", supported)
        self.assertIn("\\caption{\\textbf{基于SDT的AI教学话语策略与AI默认输出模式的比较}}", supported)
        self.assertNotIn("表 1 组别分配", supported)
        self.assertNotIn("图 3 基于SDT的AI教学话语策略与AI默认输出模式的比较", supported)

    def test_ensure_latex_build_support_caps_includegraphics_width(self) -> None:
        latex_text = (
            "\\documentclass[]{article}\n"
            "\\usepackage{graphicx}\n"
            "\\begin{document}\n"
            "\\includegraphics[width=7.63830in]{demo_media/image3.png}\n"
            "\\end{document}\n"
        )

        supported = ensure_latex_build_support(latex_text)

        self.assertIn("\\def\\maxwidth", supported)
        self.assertIn("\\dotexgraphicmaxwidth", supported)
        self.assertIn("\\dotexcapwidth{7.63830in}", supported)
        self.assertIn("height=\\maxheight", supported)
        self.assertIn("keepaspectratio", supported)
        self.assertNotIn("\x0c", supported)
        self.assertNotIn("\t extheight", supported)

    def test_ensure_latex_build_support_widens_longtables(self) -> None:
        latex_text = (
            "\\documentclass[]{article}\n"
            "\\usepackage{longtable,booktabs,array}\n"
            "\\begin{document}\n"
            "\\begin{longtable}[]{@{}\n"
            "  >{\\raggedright\\arraybackslash}p{(\\linewidth - 4\\tabcolsep) * \\real{0.5}}\n"
            "  >{\\raggedright\\arraybackslash}p{(\\linewidth - 4\\tabcolsep) * \\real{0.5}}@{}}\n"
            "demo\\tabularnewline\n"
            "\\end{longtable}\n"
            "\\end{document}\n"
        )

        supported = ensure_latex_build_support(latex_text)

        self.assertIn("\\dotextablewidthbonus", supported)
        self.assertIn("\\setlength{\\LTleft}{\\dimexpr-\\dotextablewidthbonus/2\\relax}", supported)
        self.assertIn("(\\dotextablewidth - 4\\tabcolsep)", supported)


if __name__ == "__main__":
    unittest.main()