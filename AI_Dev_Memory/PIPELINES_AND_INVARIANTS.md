# Pipelines And Invariants

## Mental model

dotex has two production-grade one-way chains:

1. DOCX -> TeX
2. TeX -> DOCX

The important thing is not that they are inverse in theory.
The important thing is that both chains preserve the same editing affordances in practice.

## Stable default model

Without downgrade flags:

- DOCX -> TeX restores editable citation and cross-reference structure.
- TeX -> DOCX emits editable Zotero fields and native Word REF/SEQ fields.
- Plain text output is not the default.

The only downgrade switches are:

- `--plaincitation`
- `--plainref`

If an implementation change makes rich behavior opt-in again, that is a regression.

## Control paths that matter

### DOCX -> TeX

Primary file:

- `src/dotex/docx_to_tex.py`

Critical responsibilities:

- detect `ADDIN ZOTERO_ITEM CSL_CITATION` fields
- restore `\parencite` / `\textcite`
- preserve occurrence-level citation shells in `dotex_zotero_items.json`
- preserve bibliography companion files:
  - `refs.bib`
  - `refs_display.json`
  - `parencite_defs.tex`
  - `dotex_zotero_items.json`

Key implementation idea:

- item-level metadata is not enough for stable roundtrip behavior
- occurrence-level field shells must be preserved when possible

### TeX -> DOCX

Primary file:

- `src/dotex/tex_to_docx.py`

Critical responsibilities:

- rebuild Zotero citation fields from companion data or live resolution
- rebuild Zotero bibliography field
- rebuild Word-native caption and cross-reference fields
- preserve package-level Zotero metadata from the reference DOCX
- canonicalize emitted citation payloads so direct and roundtrip outputs converge

## Invariants that must hold

### Citation invariants

- editable Zotero citation fields exist in output DOCX
- editable Zotero bibliography field exists in output DOCX
- parenthetical citations keep `plainCitation == formattedCitation == display text`
- emitted payloads use 8-character alphanumeric `citationID` values
- legacy `cite-...` ids are repaired when encountered in dirty source shells
- `dontUpdate` is omitted from canonical emitted payloads

### Caption and cross-reference invariants

- hidden `_Ref...` bookmarks must exist for native Word caption targets
- caption numbering uses Word-native `SEQ` fields, not custom sequence names
- caption cross-references use `REF _Ref... \h`
- internal hyperlink nodes may be removed, but the hidden `_Ref...` structure must survive

### Package invariants

- `docProps/custom.xml` must preserve Zotero custom properties when the reference/source DOCX has them
- `customXml/item1.xml` and related parts must be preserved when the reference/source DOCX has them
- `word/_rels/document.xml.rels` must preserve the customXml relationship
- `[Content_Types].xml` must preserve the customXml itemProps override

### Style invariants

- do not rewrite `word/styles.xml` in postprocess
- global font normalization must not damage Zotero field runs
- Zotero field runs should stay blue and should not keep accidental `w:rFonts` added by later normalization

## Stability target across both chains

The stable end state is not "copy the source shell exactly no matter what".

The stable end state is:

- preserve native-looking source shells when they are already clean
- repair obviously synthetic or unstable payload details into a canonical shape
- make DOCX -> TeX -> DOCX and direct TeX -> DOCX land in the same compatibility class

## Known failure mode: dirty source DOCX

A roundtrip is only as clean as the reference/source DOCX used as the template.

Real finding from the manuscript workflow:

- the active source DOCX contained many citation fields already carrying synthetic `cite-...` ids
- repeated citations could share the same `source_keys + formatted_citation` signature while still having different original shells

Implication:

- never assume the source DOCX is pristine just because it opens in Word
- inspect the package before treating it as a fidelity anchor

## Validation target

When evaluating a change, prefer this order:

1. real manuscript roundtrip behavior
2. narrow behavior tests around the touched slice
3. package-level parity checks for the touched feature
4. self-check summary only as a supporting signal
