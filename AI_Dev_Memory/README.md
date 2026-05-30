# AI Dev Memory

This folder is the handoff memory for AI developers working on dotex.

It is not end-user documentation.
It records stable conclusions, dangerous assumptions, debug tactics, and the control paths that matter when editing the converter.

Read these files in order when taking over the repo:

1. `PIPELINES_AND_INVARIANTS.md`
   - high-level model of the two one-way chains
   - stability targets that must hold for both chains
   - where each behavior is implemented in code
2. `ZOTERO_FIELD_COMPATIBILITY.md`
   - canonical Zotero payload shape
   - package-level DOCX parts that affect Word/Zotero behavior
   - known contamination patterns in real source DOCX files
3. `STYLE_CAPTION_AND_XREF.md`
   - Word-native caption and cross-reference model
   - style and bookmark invariants
   - what to flatten and what never to flatten
4. `DEBUG_PLAYBOOK.md`
   - concrete triage order
   - minimal discriminating checks
   - safe validation commands and anti-patterns

Core rule:

- Treat real DOCX packages as the source of truth for Word behavior.
- Treat real user manuscripts as the only meaningful regression target.
- Do not trust old assumptions when package XML says otherwise.

Current stable target:

- DOCX -> TeX -> DOCX and direct TeX -> DOCX should converge to the same canonical Zotero citation payload shape.
- The document should preserve editable Zotero fields, Word-native REF/SEQ caption references, hidden `_Ref...` bookmarks, and package-level Zotero metadata.
- Plain output is opt-in via downgrade flags only.
