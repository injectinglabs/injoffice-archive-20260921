# Bounded saved-literal COUNTA arguments for print areas

This profile resolves a finite same-sheet `COUNTA` range used as an `OFFSET` numeric argument. It reads the exact saved package and grants no mutation authority. It does not calculate formulas, follow external references, repair data, or claim Excel printer calibration.

For example, `OFFSET('Data Set'!$A$1,0,0,COUNTA('Data Set'!$A$1:$A$1000),4)` describes a four-column list beginning at A1. COUNTA counts populated key cells; it does not find the last occupied row. An interior gap can therefore leave a later record outside the resulting print rectangle. Reinspect after changing the source range.

## Closed syntax and limits

A COUNTA argument contains one explicitly sheet-qualified absolute finite rectangle or absolute cell. Existing OFFSET integer literals and qualified saved integer-cell arguments remain available. No whole-column shorthand, multiple COUNTA arguments, arithmetic, nested functions, relative/external references, named ranges, general expressions or leading equals are accepted. Function spelling is uppercase and syntax padding is ASCII space.

The complete source formula is limited to 2,048 bytes and at most four scalar dependency occurrences, counting each COUNTA call and each saved integer-cell occurrence. Repeated occurrences count separately. COUNTA range coordinate counts total at most 100,000, and a certified worksheet may contain at most 100,000 extracted cells. Counting scans only the complete extracted cell inventory; it never allocates the missing grid coordinates. OFFSET's existing positive-dimension and Excel-grid checks remain in force. A zero count cannot supply a height or width.

The existing union profile remains at most 16 disjoint areas, with all-or-nothing resolution and source order. Dependency and coordinate limits apply across the complete union. One saved count warning per dependency, the complete source expression and an approximation warning fit within the existing warning transport budget. Existing literal-only and single OFFSET precedence remains unchanged.

## What counts

Numbers (including zero, decimal and exponent literals), booleans (including false), saved errors, validated dates and plain nonempty text each count once. Whitespace-only text counts. Explicit empty shared-string text counts once. Ordinary absent cells and explicit numeric/default blank cells count zero; formatting alone does not make a cell populated.

This follows Microsoft's [COUNTA description and example](https://support.microsoft.com/en-us/excel/functions/counta-function). The five values 39790, 19, 22.24, TRUE and #DIV/0! each count once. Microsoft's [OFFSET reference](https://support.microsoft.com/en-us/excel/functions/offset-function) supplies the existing displacement and positive-dimension semantics.

A deliberate conservative exception is empty inline text. The current extractor represents an inline-string cell without an `is` child and an explicit empty inline string identically. The preview refuses this ambiguous case instead of guessing whether it is blank. Formula-string storage without a formula is also refused.

Any formula input, with or without a cache, refuses the count. This includes formulas returning empty text, even though Excel can count that result. Shared/array/data-table formula groups refuse source certification, including groups with absent followers. Rich or opaque input values, metadata and intersecting merges are unsupported.

## Complete saved-source requirement

COUNTA can reason about absence only using the actual complete extraction created inside native Inspect. A private synchronous context retains the exact workbook's worksheet objects, package hash and identity. It is not accepted through a public API and cannot be supplied as a partial client workbook projection. Worksheet/name/order/source-part joins remain required. Extraction limits fail rather than truncate; the worksheet dimension and visible viewport are never used as completeness evidence.

Editable status alone is insufficient because the extractor may preserve unknown worksheet children outside its cell projection. The certificate rejects selected-worksheet unsupported records and unknown or ambiguous global records. A small explicit code/capability/scope/part allowlist admits audited source-neutral style, unused rich shared-string, chart and non-core relationship/package records. It does not follow those relationships or use their caches. Unsupported shared-string table/item semantics remain refused.

The generic workbook-feature inventory code is disambiguated using a closed workbook XML gate: exact namespace `sheets` and `definedNames`, an empty `bookViews/workbookView` shape without semantic attributes, and empty `calcPr` with only a decimal uint32 `calcId` attribute. Unknown children, alternative content and semantic/foreign attributes are refused. This intentionally excludes some otherwise valid workbooks and does not imply that the rejected features change occupancy.

## Qualification

Dedicated producer tests cover strict/transitional extraction, the Microsoft five-value example, empty/absent/literal variants, source-completeness counterexamples, exact budgets and atomic refusal. Public Inspect and consumer tests verify exact source ownership, complete-set refusal and retained count provenance. Actual Go-produced inspection envelopes and browser qualification use local synthetic fixtures. Local source fixtures and browser artifacts stay outside git and CI.
