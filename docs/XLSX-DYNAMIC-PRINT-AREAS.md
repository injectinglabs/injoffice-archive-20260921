# Source-cell OFFSET print areas

The read-only `print_area_sets` projection can resolve one saved worksheet-local
`_xlnm.Print_Area` whose `OFFSET` numeric arguments use qualified numeric source
cells. For example:

```text
OFFSET('Data Set'!$D$3:$F$5,'Data Set'!$H$1,-2,'Data Set'!$J$1,3)
```

If the saved numeric values of H1 and J1 are `3` and `2`, this resolves to B6:D7.
Changing J1 to `4` in the source package and reinspecting resolves to B6:D9.
The inspection does not change the source, calculate formula cells, or observe
unsaved host edits. Each result remains joined to its package SHA-256, worksheet
ID, and worksheet source part. The legacy `print_areas` projection remains
literal-only.

## Supported subset

- One uppercase `OFFSET` call, without a leading `=`, nested calls or unions.
  Its base is an absolute same-sheet A1 cell or rectangle with explicit sheet
  qualification. The formula is at most 2,048 bytes and excludes control
  characters. ASCII spaces may surround arguments.
- Rows, columns, and optional height/width are either signed integer literals
  or explicit same-sheet absolute scalar references such as `'Data Set'!$H$1`.
  At most four scalar references can occur. Omitted trailing height/width use
  the base rectangle's dimensions; empty argument slots are unsupported.
- A referenced cell must exist exactly once in the extracted worksheet, be a
  non-formula stored numeric value, and have an optional sign followed by one
  to seven ASCII digits. Values are bounded to -1,048,576 through 1,048,576.
  Decimal and exponent spellings, including `2.0` and `2e0`, remain unsupported.
  No text/boolean/date coercion, fractional truncation or number-format parsing
  is performed. Authored numeric formatting does not alter the stored argument.
- Source sheet and cell qualification flags must permit known native semantics.
  Those flags are used conservatively for inspection; no mutation authority is
  created. Protected or otherwise refused worksheets, metadata/opaque cells,
  merged owners and covered cells, and shared/array/data-table formula group
  cells are excluded. Formula caches are never accepted as literal inputs.
- Height/width must be positive and the entire result must fit the worksheet
  grid. Existing viewport, aggregate 100,000-cell and 100-page planner limits
  still apply. A valid source rectangle can exceed the preview budget; the
  planner then refuses it as a whole.

Missing or ambiguous cells, unqualified or relative references, other-sheet or
external references, names, arithmetic expressions, dynamic functions such as
`COUNTA`, and unsupported print titles refuse the entire saved area. No partial
reference or guessed fallback area is returned.

## Public preview path and provenance

Use the unchanged inspection and consumer APIs:

1. Inspect the current workbook with `InspectNativeWorkbookObjectsV1` (Go), or
   the corresponding XLSX WASM object inspection operation.
2. Decode against the projected workbook's source hash using
   `decodeNativeWorkbookObjectsV1`, then select with
   `selectNativeSheetPrintAreaSetV1`.
3. Compile source-qualified geometry and call
   `compileNativeSheetPrintAreaSetPreviewV1`.

The source formula and each referenced cell's saved lexical value are included
in `print_area_sets[].warnings`. These warnings survive decoding and page-plan
compilation. The playground's **Use saved print area** path displays them,
including the instruction to reinspect after changing source inputs. Changing
source bytes invalidates the old package join; hosts must inspect the new
package rather than carry forward the old coordinates.

## Semantic basis and qualification limit

Microsoft documents that OFFSET shifts the reference's upper-left cell by its
row/column arguments, that dimensions must be positive, and that omitted
dimensions inherit the reference size. Its formula grammar permits cell
references as function argument expressions. The implementation applies this
reference arithmetic only after reading an exact numeric scalar from the
qualified source projection. See [Microsoft OFFSET documentation](https://support.microsoft.com/en-us/excel/functions/offset-function)
and [MS-XLSX section 2.2.2 formula grammar](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/3d025add-118d-4413-9856-ab65712ec1b0).

This bounded source-cell dependency is not a workbook calculation engine.
Generated producer/consumer tests and actual WASM/browser proof qualify this
software path. They do not establish independent Excel printer calibration,
general dynamic print-formula support, or freshness of saved formula results.
The [print calibration contract](XLSX-PRINT-CALIBRATION.md) remains unchanged.
