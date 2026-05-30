# Zotero Field Compatibility

## Goal

The goal is not merely to emit Word-valid fields.

The real goal is to emit fields that remain editable and stable under current Word + Zotero behavior.

## Canonical emitted payload shape

For emitted citation payloads, converge on this canonical shape:

- `citationID`: 8-character alphanumeric id
- `properties.unsorted`
- `properties.formattedCitation`
- `properties.plainCitation`
- `properties.noteIndex`
- `citationItems`
- `schema`

Current canonical rules:

- keep `formattedCitation` equal to the visible citation text
- for parenthetical citations, keep `plainCitation` bracket-preserving as well
- omit `dontUpdate`
- preserve native-looking 8-character citation ids from clean shells
- repair legacy `cite-...` ids into deterministic 8-character ids

## Why this canonicalization exists

Observed real-world problem:

- older source shells often used synthetic `cite-...` ids
- some parenthetical citations stored `plainCitation` without brackets even though the visible result was bracketed
- current Zotero refresh behavior can rewrite those payloads on open/edit

So the stable strategy is:

- preserve rich shell structure where it helps compatibility
- normalize obviously stale payload internals where they are known to trigger rewrites

## Direct builder vs preserved shell

There are two ways a citation field can be emitted:

1. direct construction from resolved citation metadata
2. shell reuse from `dotex_zotero_items.json`

These must not diverge.

If the direct builder emits one payload shape but preserved shells emit another, users will see unstable behavior between:

- direct TeX -> DOCX
- DOCX -> TeX -> DOCX

Canonicalization must therefore happen on both paths.

## Occurrence-level shell preservation

Preserving only bibliography item data is insufficient.

The roundtrip path must preserve occurrence-level shells because:

- different appearances of the same citation text can have different original field shells
- repeated citations may share the same bibliography target but differ in XML details

Current storage model:

- `dotex_zotero_items.json` contains a `citations` list
- each entry stores:
  - `source_keys`
  - `formatted_citation`
  - `field_nodes_xml`

## Known weakness to remember

The current matching key is still relatively weak:

- `source_keys + formatted_citation`

This is enough for many cases, but repeated citations can still collide if they share the same visible text while having distinct historical shells.

If future stability work is needed, the next improvement target is a stronger occurrence-level key.

## Package-level parts that matter to Zotero

Field XML is not the whole story.

The following package-level parts affect Word/Zotero behavior and must be copied from the reference/source DOCX when present:

- `docProps/custom.xml`
- `customXml/item1.xml`
- `customXml/itemProps1.xml`
- `customXml/_rels/item1.xml.rels`
- customXml relationship in `word/_rels/document.xml.rels`
- `/customXml/itemProps1.xml` override in `[Content_Types].xml`

Interpretation:

- if field payloads look correct but Word/Zotero still behaves differently, inspect package metadata next

## Dirty source detection mindset

Do not assume the reference/source DOCX is a pure Zotero document.

Real manuscript finding:

- the source DOCX contained a large number of already synthetic `cite-...` citation ids

Therefore:

- a preserve-shell pipeline can faithfully preserve bad shells as well as good ones
- the tool should prefer canonical repair over naive fidelity when a shell clearly looks synthetic or stale

## Display-layer contamination in preserved shells

Payload fidelity is not enough.

Real manuscript finding:

- many source citation shells were payload-valid but carried explicit `w:sz=21` and `w:szCs=21`
- many source citation runs had no `w:color`
- roundtrip output therefore inherited smaller black citations even though the Zotero payload itself was correct and editable

Stable rule:

- normalize preserved-shell display properties as well as payload internals
- keep direct-builder fields and preserved-shell fields in the same display compatibility class

Current display normalization target for every run inside a Zotero field:

- remove `w:rStyle`
- remove `w:rFonts`
- remove `w:sz`
- remove `w:szCs`
- force `w:color` to `003399`

Visible-reference convergence rule:

- TeX -> DOCX should not let citation styling diverge across direct Zotero fields, preserved Zotero shells, and citation hyperlink fallbacks
- if a visible citation-like reference survives as a plain internal-link fallback, it should still converge to the same `003399` display class

Interpretation:

- source-shell fidelity is not the goal when the source display shell is already dirty
- canonical display cleanup must stay scoped to Zotero field runs only, not body text around them

## What not to do

- do not inject `w:updateFields`
- do not reduce item payloads to title-only minimal shells if richer item data is available
- do not mix Zotero fields and plain internal-anchor citations in the same rich-output manuscript
- do not treat Word showing field code as proof that XML is wrong; check the editor/view state first

## Best discriminating checks

When a Zotero warning appears, inspect in this order:

1. visible citation text
2. `formattedCitation`
3. `plainCitation`
4. `citationID` shape: native-looking 8 chars vs legacy `cite-...`
5. presence/absence of `dontUpdate`
6. `citationItems[].id` and `uris`
7. package-level Zotero metadata parts
