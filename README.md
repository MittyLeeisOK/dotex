# dotex

dotex 1.0 is a DOCX and TeX roundtrip tool for manuscript workflows.

It is built for one default idea:

- DOCX to TeX should keep citation and cross-reference structure editable.
- TeX to DOCX should restore editable Zotero fields and native Word caption references.
- Plain text output should be optional, not the default.

## What dotex does

dotex helps you move between Word and LaTeX without throwing away the structures that matter in real manuscripts.

By default it preserves or restores:

- editable Zotero citation fields
- an editable Zotero bibliography field
- native Word figure and table references
- native Word caption numbering with REF and SEQ fields
- hidden `_Ref...` bookmarks used by Word cross-references
- extracted images and manuscript support files for roundtrip work

dotex is not just a thin Pandoc wrapper. It converts first, then postprocesses the DOCX so the result stays useful in Word.

## Install

### Quick install

```bash
curl -fsSL https://raw.githubusercontent.com/MittyLeeisOK/dotex/main/install.sh | bash -s -- --install --yes
```

If `pandoc` is already installed and you do not want automatic dependency installation:

```bash
curl -fsSL https://raw.githubusercontent.com/MittyLeeisOK/dotex/main/install.sh | bash -s -- --install --yes --skip-deps
```

### Local development install

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -U pip
python -m pip install -e .
```

## Run dotex

Use either of these entrypoints:

```bash
dotex
```

or:

```bash
python -m dotex
```

Do not use `python -m dotex.cli` as the public entrypoint.

## Quick start

Most users only need these two commands.

### 1. Convert DOCX to a TeX project

```bash
dotex convert-tex /path/to/manuscript.docx
```

By default this creates a project directory named after the DOCX stem.

Example:

- `dotex convert-tex manuscript.docx`
- writes `manuscript/manuscript.tex`

The project directory contains the TeX file, extracted media, and roundtrip support files.

### 2. Convert TeX to DOCX

```bash
dotex convert-docx /path/to/manuscript/manuscript.tex
```

By default this writes a DOCX next to the input TeX.

Example:

- `dotex convert-docx manuscript/manuscript.tex`
- writes `manuscript/manuscript.docx`

dotex restores editable Zotero fields and native Word caption references by default.

### 3. Do a full roundtrip

```bash
dotex convert-tex manuscript.docx
dotex convert-docx manuscript/manuscript.tex
```

This is the normal roundtrip workflow.

You usually do not need `--template`.

If the TeX project came from `convert-tex`, dotex will try to reuse a nearby reference DOCX automatically. Use `--template` only when you want to override that detection manually.

## The two downgrade switches

dotex defaults to rich output.

Use these only when you explicitly want flattened output:

- `--plaincitation`: flatten citations instead of emitting Zotero fields
- `--plainref`: flatten caption and cross-reference fields instead of preserving Word REF/SEQ fields

Examples:

```bash
dotex convert-docx manuscript.tex -o manuscript.docx --plaincitation
dotex convert-docx manuscript.tex -o manuscript.docx --plaincitation --plainref
dotex convert-tex manuscript.docx --output output_project --plaincitation --plainref
```

## Main commands

### `convert-tex`

```bash
dotex convert-tex INPUT.docx --output OUTPUT_PROJECT
```

Converts a DOCX manuscript into a TeX project.

### `convert-docx`

```bash
dotex convert-docx INPUT.tex -o OUTPUT.docx
```

Converts a TeX manuscript into a DOCX file.

### `inspect-template`

```bash
dotex inspect-template TEMPLATE.docx --output manifest.json
```

Extracts template information from a DOCX reference file.

### `compare-roundtrip`

```bash
dotex compare-roundtrip original.docx source.tex generated.docx --output report.md
```

Builds a diagnostic report that compares an original DOCX, the source TeX, and the generated DOCX.

This is mainly a QA and debugging command, not part of the normal author workflow.

### `resolve-zotero`

```bash
dotex resolve-zotero bibliography_links.tex --output report.json
```

Resolves bibliography entries against a local Zotero database.

### `normalize-tex`

```bash
dotex normalize-tex
```

Advanced normalization entrypoint for manuscript-specific TeX processing.

## What `convert-tex` writes

The output project always includes:

- the main `.tex` file
- the media directory
- `Makefile`
- `.latexmkrc`

When a bibliography is found, dotex can also write:

- `bibliography_links.tex`
- `refs.bib`
- `refs_display.json`
- `parencite_defs.tex`
- `dotex_zotero_items.json`

These files make DOCX to TeX to DOCX roundtrip work without rebuilding everything from scratch.

## How templates are chosen

Most users do not need to pass `--template`.

For `convert-docx`, dotex looks for a reference DOCX in this order:

1. `--template`
2. a sibling DOCX with the same stem as the input TeX
3. a parent-level DOCX whose name matches the TeX folder name
4. the built-in default template

This lets a roundtrip project reuse the original Word styling automatically when possible.

If none of these exist, dotex falls back to the built-in default template.

## Validation output

`convert-docx` ends with a self-check summary. It reports items such as:

- format score
- citation field count
- bibliography field count
- bookmark count
- internal hyperlink count
- final pass value

If Zotero matching is incomplete, dotex can also write a `.zotero-import-checklist.xlsx` file.

## Requirements

- Python 3.9 or newer
- Pandoc
- `.docx` input files, not legacy `.doc`
- UTF-8 TeX files

For best results, TeX manuscripts should keep stable `\label{...}` values for figures, tables, and equations.

## Current limits

- very complex table layouts may still need manual review
- broken or malformed Word fields cannot always be reconstructed semantically
- figures or tables without stable targets cannot receive reliable roundtrip references
- Word can delay field refresh in some environments until reopen or manual refresh

## For maintainers

AI-oriented implementation notes and debugging memory live in:

- `AI_Dev_Memory/README.md`