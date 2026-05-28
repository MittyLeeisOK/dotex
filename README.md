# dotex

dotex is a standalone, reusable high-fidelity TeX to DOCX toolkit for
table-heavy manuscripts that need to stay close to an existing Word reference
document.

After installation, you can run either `dotex` or `python3 -m dotex`.

The toolkit supports both directions:

1. TeX to DOCX with template-aware postprocessing for tables, figures,
  bibliography styling, cross-references, and optional Zotero fields.
2. DOCX to TeX with extracted media, OMML inline math recovery, and fallback
  figure reconstruction when pandoc does not emit figure blocks.
3. Roundtrip comparison across original DOCX, source TeX, and generated DOCX.

## Publishing This Repository

Do not publish this directory as-is after running it on a real manuscript unless
you first clean local outputs.

- Commit-safe files are the toolkit source and docs: `src/`, `examples/`,
  `pyproject.toml`, `README.md`, and small placeholder files such as `.gitkeep`.
- `artifacts/` and `reports/` are local run outputs. They can contain manuscript
  text, extracted figures, generated DOCX files, Zotero resolution snapshots,
  and validation reports derived from your private document.
- This repository is configured so `artifacts/` and `reports/` stay out of git
  by default. Before pushing to GitHub, keep those directories empty except for
  placeholder files.

## Quick Install

For a normal end-user install from a local checkout:

```bash
bash install.sh --install
```

The installer now tries to install `pandoc` automatically when possible and prints stage-by-stage progress during installation.

If you want to skip the automatic dependency step:

```bash
bash install.sh --install --skip-deps
```

The installer defaults to a user-scope install under `~/.local/share/dotex`
with launchers written to `~/.local/bin`. Use `--system` to install under `/opt`
and `/usr/local/bin`.

On first install, the script prints:

- the main command entrypoints
- the most common conversion commands
- the runtime prerequisites such as Python 3.9+ and `pandoc`
- Zotero-related notes and output-file warnings

## Commands

Install in editable mode:

```bash
python3 -m pip install -e .
```

or

```bash
pip install -e .
dotex --help
```

Inspect a Word template:

```bash
dotex inspect-template \
  /path/to/reference-template.docx \
  --output artifacts/normal_manuscript.json
```

Convert TeX to DOCX with Zotero fields enabled:

```bash
dotex convert-docx \
  /path/to/manuscript.tex \
  -o /path/to/manuscript.docx \
  -z
```

Convert DOCX back to TeX and extract image resources next to the output:

```bash
dotex convert-tex \
  /path/to/manuscript.docx \
  --output /path/to/manuscript.tex
```

Generate a structural comparison report for an original DOCX, source TeX, and
generated DOCX:

```bash
dotex compare-roundtrip \
  /path/to/original.docx \
  /path/to/source.tex \
  /path/to/generated.docx \
  --output /path/to/generated.roundtrip-comparison.md
```

Convert TeX to DOCX while keeping citations as internal bibliography anchors:

```bash
dotex convert-docx \
  /path/to/manuscript.tex \
  -o /path/to/manuscript.docx
```

If `--template` is omitted, the tool uses the bundled default reference
template that matches the toolkit's built-in manuscript style.

If you want to override that default style, pass an explicit Word template:

```bash
dotex convert-docx \
  /path/to/manuscript.tex \
  -t /path/to/reference-template.docx \
  -o /path/to/manuscript.docx
```

Resolve a bibliography file against the local Zotero sqlite database:

```bash
dotex resolve-zotero \
  /path/to/bibliography_links.tex
```

Important `convert-docx` options:

- `-t` / `--template`: optional Word reference DOCX. If omitted, the tool uses the
  bundled default reference template. Pass this only when you want to override
  the toolkit's default page settings and styles.
- `-o` / `--output`: output DOCX path. Defaults to the source TeX path with a
  `.docx` suffix.
- `--bibliography`: optional path to a bibliography file containing
  `\bibentry{key}{entry}` definitions. If omitted, the tool auto-detects a
  `\input{...}` file in the TeX body that contains `\bibentry`, then falls back
  to `bibliography_links.tex` next to the source TeX.
- `--bibliography-heading`: heading text used to find the bibliography section
  in the generated DOCX. Default: `参考文献`.
- `-z` / `--zotero`: emit Zotero field reconstruction. If omitted, citations stay as internal anchors.

Important `convert-tex` options:

- `--output`: output TeX path. Defaults to the DOCX path with a `.tex` suffix.
- `--media-dir`: explicit directory for extracted media. Defaults to
  `OUTPUT_TEX_STEM_media`.
- `--standalone` / `--no-standalone`: emit a full document preamble or a body-only
  LaTeX fragment.

## Reverse Conversion Guarantees

`convert-tex` uses pandoc as the baseline converter and then applies targeted
fallbacks for the failure modes that matter most in manuscript workflows.

- Tables: pandoc-generated longtables and table environments are preserved.
- Image resources: extracted into a flat media directory next to the output TeX.
- Figures: if pandoc emits zero `\includegraphics` blocks for a DOCX that clearly
  contains drawings, the toolkit reconstructs figure blocks from DrawingML and
  the following caption paragraph.
- Formulas: inline OMML objects are replaced with stable placeholders before
  pandoc runs, then restored as LaTeX inline math after conversion.

This means the toolkit is intentionally stricter than raw pandoc for DOCX to
TeX conversion: it does not assume pandoc alone is sufficient for every DOCX
produced in a roundtrip workflow.

## Supported TeX Patterns

The converter is intentionally pattern-based. It preserves layout best when the
source follows these conventions.

### Tables

Supported well:

- `table`, `table*`, and `longtable`
- `tabular`, `tabularx`, and `longtable` column specs using `l`, `c`, `r`, `p{}`,
  `m{}`, `b{}`, and `X`
- `\centering` or `center` around the table block
- widths expressed through ratios such as `0.92\linewidth`
- widths derived from lengths defined with `\setlength`, especially when those
  lengths are built from `\linewidth`, `\textwidth`, or simple arithmetic
- longtable width controls such as `\LTleft`, `\LTright`, and `\LTcapwidth`

To preserve table width and column width, keep layout commands inside the table
environment. For example, put `\scriptsize`, `\tabcolsep`, widening lengths, and
alignment commands inside the same `table` or `longtable` block that owns the
actual tabular material.

Centered widened tables are emitted with table left indent forced to `0` and the
table object itself centered, which avoids Word shifting the table to the right.

### Figures

Supported well:

- `figure` and `figure*`
- `\includegraphics[width=...]`
- `\centering` or `center` around the figure
- captions written with `\caption{...}` and labels written with `\label{...}`

Figure centering and width restoration depend on width values being readable
from the TeX source. Widths tied to custom macros are supported when those
macros ultimately resolve from lengths set in the preamble or nearby code.

Centered figures are emitted with zero paragraph indentation to avoid rightward
drift after width expansion.

When converting from DOCX back to TeX, figure captions are used as the fallback
anchor for reconstructing figure environments if the DOCX drawing survives but
raw pandoc image emission fails.

### Cross-references

Supported well:

- `\tabref{label}`
- `\figref{label}`
- `\label{...}` on figures and tables

These are converted into Word-internal links.

### Bibliography and citations

Expected pattern:

- in-text citations use `\litref{key}{visible text}`
- bibliography entries are provided as `\bibentry{key}{formatted entry}`
- the main TeX file includes the bibliography file with `\input{...}`

The first argument of `\litref` and `\bibentry` must be the same logical key.
In the current workflow this key is typically a DOI URL or a stable URL-like
identifier. The converter derives a stable internal bibliography anchor from
that key and uses it for both the in-text citation target and the bibliography
entry anchor.

## Zotero Mode

When `-z` / `--zotero` is used:

- in-text citations are first normalized as internal bibliography links
- DOCX postprocessing converts matched bibliography links into Zotero `CSL_CITATION` fields
- citation groups such as `(A 2024; B 2025)` are reconstructed as a single field
- bibliography paragraphs receive a `CSL_BIBLIOGRAPHY` field wrapper
- entries missing from the local Zotero library are written to a companion Excel checklist next to the generated DOCX
- when Zotero mode is enabled, all recognized citation groups are emitted as Zotero fields even if some items are not yet matched in the local Zotero library
- unmatched items therefore remain present inside Zotero field payloads, but may still lack a resolved Zotero URI until those items are imported and the DOCX is regenerated
- an unmatched import checklist is written next to the generated DOCX as
  `OUTPUT_DOCX.zotero-import-checklist.xlsx`

When `-z` / `--zotero` is not used:

- citations remain normal Word internal hyperlinks
- clicking a citation jumps to the matching bibliography entry
- no Zotero fields are inserted, so self-check expects `0` citation fields and
  `0` bibliography fields

## How Zotero Matching Works

`resolve-zotero` and the Zotero-enabled conversion mode match bibliography
entries against the local Zotero library in this order:

1. DOI
2. URL
3. normalized title

Best results require the local Zotero database to contain the cited items. If a
bibliography entry is not found locally, the converter leaves its in-text
citation data inside the generated Zotero field and adds that entry to an
unmatched import checklist workbook for manual Zotero import. After those items
are imported into Zotero, rerun `convert-docx` to rebuild the field payloads
with resolved library URIs.

The default database path is `~/Zotero/zotero.sqlite`.

Directly writing into the Zotero sqlite database is intentionally unsupported in
this tool. The database is opened read-only, and the intended workflow is:

1. read the unmatched DOI/URL hints from the generated DOCX or checklist artifact
2. import those items into Zotero through Zotero itself
3. rerun `convert-docx`

## If The TeX Does Not Follow These Rules

The tool is designed to degrade predictably, not magically infer arbitrary TeX.
When unsupported patterns are used, the most common fallbacks are:

- table widths collapse toward pandoc defaults
- column widths become closer to averaged widths
- figure widths lose manuscript-specific widening
- citations stay as plain links instead of being upgraded into the intended
  structured form

The fastest fixes are usually:

1. rewrite custom citation macros to `\litref{key}{text}`
2. move bibliography entries into `\bibentry{key}{...}` form
3. keep table and figure layout commands inside their owning environment
4. replace opaque custom width macros with `\setlength`-based lengths that
   resolve from `\linewidth` or `\textwidth`
5. keep captions and labels close to the figure or table they belong to

If a manuscript uses different conventions, prefer adding a thin normalization
step in the TeX source or extending the parser for that convention rather than
hardcoding another manuscript-specific branch.

## Self-check Outputs

After each `convert-docx` run, the tool rewrites these artifacts under the
artifact directory:

- `original_docx_manifest.json`
- `generated_docx_manifest.json`
- `docx_validation_report.json`
- `zotero_resolution.json` and `zotero_library_subset.json` when Zotero mode is
  enabled and the local sqlite database is available

When Zotero mode is enabled and the local sqlite database is available, the tool
also writes `OUTPUT_DOCX.zotero-import-checklist.xlsx` next to the generated DOCX.

## Roundtrip Comparison

`compare-roundtrip` does a structural comparison focused on the surfaces that are
most fragile in manuscript conversion:

- Word tables versus TeX table environments
- Word drawings and media versus TeX figure and `\includegraphics` usage
- OMML formulas versus LaTeX math markers
- caption-like paragraphs, `\caption`, and `\label` signals

The report is intentionally structural rather than semantic. It is designed to
flag likely fidelity gaps quickly, not to prove textual identity.

## Example Workflow

See [examples/roundtrip_workflow.md](examples/roundtrip_workflow.md) for a
minimal end-to-end workflow covering DOCX to TeX, TeX to DOCX, and structural
roundtrip comparison.

The current self-check compares structural formatting features such as section
layout, table styles, caption counts, body paragraph styles, bibliography style
presence, table cell paragraph styles, and Zotero field presence when Zotero
mode is enabled.