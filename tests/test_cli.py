import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from dotex.cli import (
    build_convert_docx_request,
    build_parser,
    default_docx_output_path,
    default_tex_output_path,
    resolve_template_selection,
    resolve_convert_tex_output_path,
)


class CliTests(unittest.TestCase):
    def test_default_docx_output_path_uses_plain_docx_suffix(self) -> None:
        tex_path = Path("/tmp/manuscript_v3.tex")

        self.assertEqual(default_docx_output_path(tex_path), Path("/tmp/manuscript_v3.docx"))

    def test_default_docx_output_path_replaces_non_tex_suffix(self) -> None:
        tex_path = Path("/tmp/subdir/manuscript.source.tex")

        self.assertEqual(default_docx_output_path(tex_path), Path("/tmp/subdir/manuscript.source.docx"))

    def test_default_tex_output_path_uses_project_directory(self) -> None:
        docx_path = Path("/tmp/manuscript_v3.docx")

        self.assertEqual(default_tex_output_path(docx_path), Path("/tmp/manuscript_v3/manuscript_v3.tex"))

    def test_resolve_convert_tex_output_path_accepts_directory_override(self) -> None:
        docx_path = Path("/tmp/manuscript_v3.docx")
        output_dir = Path("/tmp/exported_project")

        self.assertEqual(
            resolve_convert_tex_output_path(docx_path, output_dir),
            Path("/tmp/exported_project/manuscript_v3.tex"),
        )

    def test_resolve_convert_tex_output_path_keeps_explicit_tex_file(self) -> None:
        docx_path = Path("/tmp/manuscript_v3.docx")
        explicit_output = Path("/tmp/custom/main.tex")

        self.assertEqual(resolve_convert_tex_output_path(docx_path, explicit_output), explicit_output)

    def test_resolve_template_selection_defaults_to_builtin_template(self) -> None:
        with TemporaryDirectory() as temp_dir:
            from contextlib import ExitStack

            with ExitStack() as stack:
                selection = resolve_template_selection(stack, None)

            self.assertTrue(selection.is_builtin)
            self.assertTrue(selection.path.exists())

    def test_resolve_template_selection_uses_explicit_override(self) -> None:
        with TemporaryDirectory() as temp_dir:
            from contextlib import ExitStack

            root = Path(temp_dir)
            template_path = root / "explicit-template.docx"
            template_path.write_text("", encoding="utf-8")

            with ExitStack() as stack:
                selection = resolve_template_selection(stack, template_path)

            self.assertFalse(selection.is_builtin)
            self.assertEqual(selection.path, template_path.resolve())

    def test_convert_docx_defaults_to_zotero_and_native_refs(self) -> None:
        with TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            tex_path = root / "manuscript_v3.tex"
            tex_path.write_text("\\begin{document}demo\\end{document}\n", encoding="utf-8")
            parser = build_parser()

            request = build_convert_docx_request(parser.parse_args(["convert-docx", str(tex_path)]))

        self.assertTrue(request.enable_zotero)
        self.assertFalse(request.plain_citation)
        self.assertFalse(request.plain_ref)

    def test_convert_docx_plain_flags_disable_default_rich_modes(self) -> None:
        with TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            tex_path = root / "manuscript_v3.tex"
            tex_path.write_text("\\begin{document}demo\\end{document}\n", encoding="utf-8")
            parser = build_parser()

            request = build_convert_docx_request(
                parser.parse_args(["convert-docx", str(tex_path), "--plaincitation", "--plainref"])
            )

        self.assertFalse(request.enable_zotero)
        self.assertTrue(request.plain_citation)
        self.assertTrue(request.plain_ref)

    def test_convert_tex_plain_flags_are_opt_in(self) -> None:
        parser = build_parser()

        default_args = parser.parse_args(["convert-tex", "/tmp/manuscript_v3.docx"])
        plain_args = parser.parse_args(["convert-tex", "/tmp/manuscript_v3.docx", "--plaincitation", "--plainref"])

        self.assertFalse(default_args.plain_citation)
        self.assertFalse(default_args.plain_ref)
        self.assertTrue(plain_args.plain_citation)
        self.assertTrue(plain_args.plain_ref)


if __name__ == "__main__":
    unittest.main()