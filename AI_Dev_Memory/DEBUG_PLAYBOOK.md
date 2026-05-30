# Debug Playbook

## First principles

- Start from the real failing artifact, not the theory.
- Compare package parts and field XML before changing code.
- Prefer one discriminating check over broad exploration.
- When a real manuscript exists, use it.

## Good anchors

Use one of these as the first anchor:

- a specific citation string
- a specific caption reference
- a specific DOCX package part
- a failing narrow test
- a specific converter function already known to control the behavior

## Minimal discriminating checks

### Zotero edit warning

Check in this order:

1. extract the exact field block from `word/document.xml`
2. compare source vs generated `formattedCitation`
3. compare source vs generated `plainCitation`
4. compare source vs generated `citationID` class
5. compare `citationItems[].id` and `uris`
6. compare package-level Zotero parts

If field-level parity holds, move up to package-level checks.

### Citation looks smaller or loses blue color

Treat this as a display-shell problem first, not a payload problem.

Check in this order:

1. extract one exact `ADDIN ZOTERO_ITEM CSL_CITATION` block from source and generated `word/document.xml`
2. compare the visible result runs, not just the JSON payload in `instrText`
3. inspect `w:color`, `w:sz`, `w:szCs`, `w:rFonts`, and `w:rStyle` on runs inside the field block
4. if the source already carries explicit small sizes or missing color, do not preserve those display properties blindly
5. confirm direct-builder fields and preserved-shell fields land in the same normalized display shape

### Caption / REF / SEQ issue

Check in this order:

1. does the output keep hidden `_Ref...` bookmarks?
2. does the body contain `REF _Ref... \h` fields?
3. do captions use `SEQ 图/表/公式`?
4. only then inspect style cleanup side effects

### Word repair / save-as problem

Check in this order:

1. whether `word/styles.xml` was touched
2. relationships and content types
3. malformed field cleanup or hyperlink cleanup
4. recently added package parts

## Real manuscript lessons worth reusing

- if the main title disappears into body text after roundtrip, inspect whether DOCX -> TeX downgraded the title into the first body paragraph; recover it with paragraph-style hints instead of editing `styles.xml`
- treat `\hl{...}` as a content-bearing wrapper; dropping it deletes user text rather than merely removing formatting
- if a complex longtable becomes chaotic in DOCX, inspect the TeX -> markdown downgrade before touching DOCX postprocess; `%` inside `\multicolumn{...}{%` is layout syntax, not cell content
- do not naively expand every `\multicolumn` into filler columns when targeting markdown tables; that creates phantom blank columns and shifts real cells

## Commands worth reusing

Real validation commands should stay narrow and reproducible.

Common command patterns:

- generate roundtrip project from a real DOCX
- generate DOCX directly from TeX using a real reference DOCX
- extract and diff a single field block from `word/document.xml`
- hash or compare key package parts:
  - `docProps/custom.xml`
  - `customXml/item1.xml`
  - `word/_rels/document.xml.rels`
  - `[Content_Types].xml`

## Anti-patterns

- do not assume source DOCX is pristine
- do not use broad repo exploration after the controlling path is already known
- do not patch unrelated tests while a local hypothesis is still under validation
- do not trust self-check alone when Word/Zotero interactive behavior is the complaint
- do not chase exact citationID values as the compatibility target; chase the correct citationID class and payload shape

## What counts as a real fix

A real fix should satisfy all of these when applicable:

- narrow test for the touched slice passes
- real manuscript output is regenerated
- output package keeps required Zotero metadata parts
- direct TeX -> DOCX and roundtrip DOCX -> TeX -> DOCX land in the same compatibility class

## Release mindset for future AI developers

If a change helps one chain but not the other, it is incomplete.

If a change copies source XML more faithfully but increases dependence on a dirty source DOCX, it is incomplete.

If a change only looks right in plain XML diff but has not been checked against Word/Zotero behavior on a real manuscript, confidence should stay low.
