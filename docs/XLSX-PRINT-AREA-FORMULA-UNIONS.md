# Mixed rectangle and OFFSET print-area unions

A saved worksheet-local `_xlnm.Print_Area` can contain a comma-separated union
of qualified absolute same-sheet rectangles and bounded `OFFSET` components.
For example:

```text
'Data Set'!$F$6:$G$8,OFFSET('Data Set'!$D$3,3,-2),OFFSET('Data Set'!$A$10,0,0,'Data Set'!$H$1,2)
```

When the saved numeric value in H1 is `2`, the ordered areas are F6:G8, B6,
and A10:B11. The preview keeps that order and starts a separate approximate
page sequence for each area. It does not paint the cells between those areas
or combine them into one bounding rectangle.

## Bounded grammar and atomic refusal

- A formula union is at most 2,048 bytes and contains two to sixteen components,
  including at least one supported OFFSET. ASCII spaces may surround components.
- Only commas outside quoted sheet tokens and OFFSET parentheses separate
  components. Doubled apostrophes and literal commas/parentheses inside quoted
  sheet names retain their source meaning. Unbalanced tokens, empty components,
  nested unquoted parentheses, a leading `=`, and general enclosing expressions
  are unsupported.
- A component is one absolute same-sheet A1 rectangle, one constant OFFSET,
  or one [source-cell OFFSET](XLSX-DYNAMIC-PRINT-AREAS.md). No other functions,
  expression arithmetic, names, relative references or cross-sheet/external
  references are evaluated.
- At most four saved-cell argument **occurrences** are allowed across the whole
  union. Repeating the same input reference consumes another occurrence. This
  is a complete-set budget, not four inputs per component.
- Every component must resolve, and all resulting rectangles must be disjoint.
  An overlap discovered after OFFSET resolution also refuses the entire set.
  A bad later component never leaves an earlier partial selection available.

Worksheet/name/source-part ownership, package-hash joins and supported saved
print-title requirements are unchanged. Existing geometry and aggregate
100,000-cell/100-page limits continue to apply after source selection, without
truncation or guessed ranges. Source bytes and mutation authority are unchanged.

The existing literal-only union parser remains the first path, with its
4,096-byte budget and original grammar. The single constant or source-cell
OFFSET path also keeps its existing behavior. This fallback does not broaden
literal-only whitespace or single-call expression syntax.

## Provenance and public preview

The existing `print_area_sets` inspection projection carries one whole-source
formula warning, one approximation/limitations warning and up to four saved
input warnings. This stays within the existing eight-warning decoder limit,
and the 2,048-byte formula bound keeps individual warnings below 4,096 bytes.
The original lexical input values are retained; formula caches are excluded.

No new consumer API is required. Decode the inspection against the current
workbook source hash, select with `selectNativeSheetPrintAreaSetV1`, compile
geometry for every selected range and call `compileNativeSheetPrintAreaSetPreviewV1`.
Each resulting area's plan retains the full provenance. The playground's
**Use saved print area** mode displays the same source/dependency warnings and
the separate area previews. Reinspect after changing source inputs.

Microsoft documents comma as a reference union operator and OFFSET as a
function returning an offset reference. This implementation combines only its
existing qualified reference subsets. See [Microsoft reference operators](https://support.microsoft.com/en-us/excel/calculation-operators-and-precedence-in-excel)
and [OFFSET semantics](https://support.microsoft.com/en-us/excel/functions/offset-function).
Source-order separate page sequences remain an explicit approximate preview
policy, not a claim of Excel page fitting or printer parity. Synthetic tests
and browser/WASM proof do not replace independent Excel print references.
