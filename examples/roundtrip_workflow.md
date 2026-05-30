# Roundtrip Workflow

This workflow shows a standard dotex roundtrip path.

## 1. Install The Toolkit

```bash
curl -fsSL https://raw.githubusercontent.com/MittyLeeisOK/dotex/main/install.sh | bash -s -- --install --yes
```

## 2. Convert DOCX To TeX

```bash
dotex convert-tex \
  /path/to/original.docx \
  --output /path/to/original_from_docx
```

Outputs:

- `/path/to/original_from_docx/original_from_docx.tex`
- `/path/to/original_from_docx/bibliography_links.tex` when the DOCX ends with a references section
- `/path/to/original_from_docx/original_from_docx_media/`
- `/path/to/original_from_docx/.latexmkrc`
- `/path/to/original_from_docx/Makefile`

## 3. Convert TeX Back To DOCX

```bash
dotex convert-docx \
  /path/to/source.tex \
  -o /path/to/source.docx
```

Use `-t /path/to/reference-template.docx` only when you want to override the bundled default template.

Possible companion output when Zotero mode is enabled:

Use `-z` or `--zotero` on `convert-docx` if you want this mode.

- `/path/to/source.zotero-import-checklist.xlsx`

## 4. Compare The Roundtrip

```bash
dotex compare-roundtrip \
  /path/to/original.docx \
  /path/to/source.tex \
  /path/to/source.docx \
  --output /path/to/source.roundtrip-comparison.md
```

The generated comparison report focuses on:

- tables
- figure and media surfaces
- formula carriers
- caption and label signals

## 5. Review The Structural Gaps

Use the roundtrip report to answer these questions quickly:

1. Did table counts stay aligned?
2. Did figure counts stay aligned?
3. Did the DOCX introduce OMML formulas that were not present in the source TeX?
4. Did the converted TeX recover image resources and inline math markers?

If the answer to any of those is no, inspect the generated TeX or DOCX before
doing further manuscript edits.