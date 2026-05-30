# Style, Captions, And Cross-References

## The core Word model

Word-native caption cross-references are not ordinary hyperlinks.

They depend on:

- hidden `_Ref...` bookmarks
- `REF _Ref... \h` fields in body text
- `SEQ 图`, `SEQ 表`, `SEQ 公式` fields for numbering

Important consequence:

- `internal hyperlinks == 0` does not mean caption references are broken
- clickable Word-native references can exist with zero internal hyperlink nodes

## Stable caption rules

- figures use `SEQ 图`
- tables use `SEQ 表`
- equations use `SEQ 公式`
- do not use custom sequence names as the main implementation

## Stable bookmark rules

- keep hidden `_Ref...` caption bookmarks
- remove unrelated non-hidden bookmark noise where possible
- do not pursue a "zero bookmarks" target in rich Zotero mode

Correct target:

- zero internal hyperlink nodes is acceptable
- zero non-hidden ordinary bookmarks is desirable
- hidden `_Ref...` cross-reference bookmarks must remain

## Stable ordering rule

In Zotero-rich DOCX postprocess, the order matters:

1. apply bibliography hints
2. convert citation hyperlinks into Zotero fields
3. emit native caption bookmarks and REF/SEQ fields
4. normalize document styling
5. strip internal hyperlink styling
6. clean non-`_Ref...` bookmarks

Why:

- if REF insertion happens too early, paragraphs containing fresh `fldChar` nodes can be skipped by later citation conversion logic
- that can silently reduce citation field counts

## Style handling rules

- do not rewrite `word/styles.xml`
- prefer localized XML changes in `document.xml` and related parts
- if font normalization adds `w:rFonts` into Zotero field runs, strip them back out inside the field block
- keep Zotero field runs source-like, including blue run color

## Figure/table layout guidance

- table notices are acceptable auxiliary output, not necessarily pipeline failure
- missing source labels are a content problem, not always a converter bug
- if longtable interacts badly with figure placement, fix figure placement at the TeX-prep stage rather than patching Word XML blindly

## Rich vs plain behavior

Rich mode:

- preserves editable Zotero fields
- preserves native Word caption structure

Plain mode:

- flattens citations only when `--plaincitation` is passed
- flattens caption references only when `--plainref` is passed

If rich mode starts flattening any of these by default, that is a regression.

## Fast caption triage

If figure/table references do not click or update correctly:

1. inspect hidden `_Ref...` bookmarks
2. inspect `REF _Ref... \h` body fields
3. inspect `SEQ 图/表/公式` numbering fields
4. only then inspect visible styles or hyperlink nodes
