package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"reflect"
	"sort"
)

// NativeWorkbookMutationTransactionV1 is the atomic save boundary for the
// native XLSX v1 projection. ExpectedRevision must fingerprint the exact input
// package. The three operation families are applied only after every target
// has passed strict native extraction and mutation-refusal preflight.
type NativeWorkbookMutationTransactionV1 struct {
	ExpectedRevision string               `json:"expected_revision"`
	Cells            []CellMutation       `json:"cells,omitempty"`
	Styles           []StylePatchMutation `json:"styles,omitempty"`
	Layout           []LayoutMutation     `json:"layout,omitempty"`
}

// NativeWorkbookMutationResultV1 returns both the exact saved package and its
// reopened, validated native contract. Callers must persist Package rather
// than reconstructing an archive from Workbook.
type NativeWorkbookMutationResultV1 struct {
	Package  []byte
	Workbook *NativeWorkbookV1
}

const maxNativeWorkbookTransactionOperations = maxCellMutations
const maxNativeWorkbookTransactionStyleTargets = 10_000

// MaxNativeWorkbookMutationPayloadBytes bounds the format-specific JSON object
// carried by the renderer-neutral Office mutation envelope.
const MaxNativeWorkbookMutationPayloadBytes = 3 << 20

// DecodeNativeWorkbookMutationTransactionV1 strictly decodes one bounded XLSX
// mutation payload and binds its native rev:<digest> CAS to the outer
// sha256:<digest> CAS. Duplicate keys at any depth, unknown fields, and trailing
// JSON are refused before any mutation can run.
func DecodeNativeWorkbookMutationTransactionV1(data []byte, outerExpectedRevision string) (NativeWorkbookMutationTransactionV1, error) {
	if len(data) == 0 || len(data) > MaxNativeWorkbookMutationPayloadBytes {
		return NativeWorkbookMutationTransactionV1{}, fmt.Errorf("xlsxpatch: decode native mutation: JSON size must be 1..%d bytes", MaxNativeWorkbookMutationPayloadBytes)
	}
	if !nativeWorkbookSHA.MatchString(outerExpectedRevision) {
		return NativeWorkbookMutationTransactionV1{}, fmt.Errorf("xlsxpatch: decode native mutation: outer expected revision must be sha256 followed by the full lowercase exact-byte digest")
	}
	if err := rejectDuplicateNativeWorkbookJSONKeys(data); err != nil {
		return NativeWorkbookMutationTransactionV1{}, fmt.Errorf("xlsxpatch: decode native mutation: %w", err)
	}
	decoder := json.NewDecoder(bytes.NewReader(data))
	decoder.DisallowUnknownFields()
	decoder.UseNumber()
	var transaction NativeWorkbookMutationTransactionV1
	if err := decoder.Decode(&transaction); err != nil {
		return NativeWorkbookMutationTransactionV1{}, fmt.Errorf("xlsxpatch: decode native mutation: %w", err)
	}
	var trailing any
	if err := decoder.Decode(&trailing); err != io.EOF {
		if err == nil {
			return NativeWorkbookMutationTransactionV1{}, fmt.Errorf("xlsxpatch: decode native mutation: trailing JSON value")
		}
		return NativeWorkbookMutationTransactionV1{}, fmt.Errorf("xlsxpatch: decode native mutation: %w", err)
	}
	nativeRevision := "rev:" + outerExpectedRevision[len("sha256:"):]
	if transaction.ExpectedRevision != nativeRevision {
		return NativeWorkbookMutationTransactionV1{}, fmt.Errorf("xlsxpatch: decode native mutation: expected_revision %q does not join outer CAS %q", transaction.ExpectedRevision, outerExpectedRevision)
	}
	if err := validateNativeWorkbookMutationTransaction(transaction); err != nil {
		return NativeWorkbookMutationTransactionV1{}, err
	}
	return transaction, nil
}

// ApplyNativeWorkbookMutationPayloadV1 is the envelope-facing XLSX entrypoint.
// It proves the outer CAS against the exact source bytes, strictly decodes and
// joins the native revision, and only then enters the typed transaction path.
func ApplyNativeWorkbookMutationPayloadV1(original, payload []byte, outerExpectedRevision string) (*NativeWorkbookMutationResultV1, error) {
	actualOuterRevision := nativeWorkbookDigest(original)
	if outerExpectedRevision != actualOuterRevision {
		return nil, fmt.Errorf("xlsxpatch: native mutation: stale outer revision: expected %q, exact package is %q", outerExpectedRevision, actualOuterRevision)
	}
	transaction, err := DecodeNativeWorkbookMutationTransactionV1(payload, outerExpectedRevision)
	if err != nil {
		return nil, err
	}
	return ApplyNativeWorkbookMutationTransactionV1(original, transaction)
}

// ApplyNativeWorkbookMutationTransactionV1 applies cell/formula, style, and
// row/column mutations as one fail-closed transaction over native OOXML bytes.
// It rejects stale revisions and unsupported targets before producing output,
// then reopens, extracts, validates, and verifies the result and proves that
// every unrelated OPC entry retained its original raw bytes and metadata.
func ApplyNativeWorkbookMutationTransactionV1(original []byte, transaction NativeWorkbookMutationTransactionV1) (*NativeWorkbookMutationResultV1, error) {
	if err := validateNativeWorkbookMutationTransaction(transaction); err != nil {
		return nil, err
	}
	wantRevision := "rev:" + nativeWorkbookDigestHex(original)
	if transaction.ExpectedRevision != wantRevision {
		return nil, fmt.Errorf("xlsxpatch: native mutation: stale revision: expected %q, exact package is %q", transaction.ExpectedRevision, wantRevision)
	}

	before, err := ExtractNativeWorkbookV1(original)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: native mutation: extract source: %w", err)
	}
	if before.Revision != transaction.ExpectedRevision {
		return nil, fmt.Errorf("xlsxpatch: native mutation: extracted revision %q disagrees with expected revision", before.Revision)
	}
	if err := preflightNativeWorkbookMutationTargets(before, transaction); err != nil {
		return nil, err
	}
	expectedStyles, err := expectedNativeStyleMutationProjections(original, before, transaction.Styles)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: native mutation: style readback plan: %w", err)
	}

	produced := bytes.Clone(original)
	if len(transaction.Cells) != 0 {
		produced, err = ApplyCellMutations(produced, transaction.Cells)
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: native mutation: cells: %w", err)
		}
	}
	if len(transaction.Styles) != 0 {
		produced, err = ApplyStyleMutations(produced, transaction.Styles)
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: native mutation: styles: %w", err)
		}
	}
	if len(transaction.Layout) != 0 {
		produced, err = ApplyLayoutMutations(produced, transaction.Layout)
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: native mutation: layout: %w", err)
		}
	}
	if bytes.Equal(original, produced) {
		return nil, fmt.Errorf("xlsxpatch: native mutation: semantic no-op produced no authoritative package change")
	}

	after, err := ExtractNativeWorkbookV1WithOptions(produced, NativeWorkbookExtractionOptions{Previous: before})
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: native mutation: reopen/extract result: %w", err)
	}
	if issues := ValidateNativeWorkbookV1(after); len(issues) != 0 {
		return nil, fmt.Errorf("xlsxpatch: native mutation: validate result: %w", &NativeWorkbookValidationError{Issues: issues})
	}
	if after.DocumentID != before.DocumentID {
		return nil, fmt.Errorf("xlsxpatch: native mutation: post-save document identity changed")
	}
	if err := verifyNativeWorkbookMutationResult(before, after, transaction, expectedStyles); err != nil {
		return nil, err
	}
	if err := verifyNativeStyleClearPostconditions(produced, after, transaction.Styles); err != nil {
		return nil, err
	}
	if err := verifyNativeWorkbookTransactionRawPreservation(original, produced, before, transaction); err != nil {
		return nil, err
	}
	return &NativeWorkbookMutationResultV1{Package: produced, Workbook: after}, nil
}

func nativeWorkbookDigestHex(data []byte) string {
	return nativeWorkbookDigest(data)[len("sha256:"):]
}

func validateNativeWorkbookMutationTransaction(transaction NativeWorkbookMutationTransactionV1) error {
	if !nativeWorkbookRevision.MatchString(transaction.ExpectedRevision) {
		return fmt.Errorf("xlsxpatch: native mutation: expected_revision must contain a full SHA-256")
	}
	total := len(transaction.Cells) + len(transaction.Styles) + len(transaction.Layout)
	if total == 0 {
		return fmt.Errorf("xlsxpatch: native mutation: empty transaction")
	}
	if total > maxNativeWorkbookTransactionOperations {
		return fmt.Errorf("xlsxpatch: native mutation: transaction exceeds %d operations", maxNativeWorkbookTransactionOperations)
	}
	seen := make(map[string]bool, total)
	for _, group := range [][]string{cellOperationIDs(transaction.Cells), styleOperationIDs(transaction.Styles), layoutOperationIDs(transaction.Layout)} {
		for _, id := range group {
			if seen[id] {
				return fmt.Errorf("xlsxpatch: native mutation: duplicate operation_id %q across transaction", id)
			}
			seen[id] = true
		}
	}
	if len(transaction.Cells) != 0 {
		if _, err := validateCellMutations(transaction.Cells); err != nil {
			return err
		}
	}
	if len(transaction.Styles) != 0 {
		if _, err := validateStyleMutations(transaction.Styles); err != nil {
			return err
		}
		var targets uint64
		for _, mutation := range transaction.Styles {
			area := uint64(mutation.Range.EndRow-mutation.Range.Row+1) * uint64(mutation.Range.EndColumn-mutation.Range.Column+1)
			if area > maxNativeWorkbookTransactionStyleTargets-targets {
				return fmt.Errorf("xlsxpatch: native mutation: style target expansion exceeds %d cell visits", maxNativeWorkbookTransactionStyleTargets)
			}
			targets += area
		}
	}
	if len(transaction.Layout) != 0 {
		if _, err := validateLayoutMutations(transaction.Layout); err != nil {
			return err
		}
	}
	return nil
}

func cellOperationIDs(items []CellMutation) []string {
	ids := make([]string, len(items))
	for index := range items {
		ids[index] = items[index].OperationID
	}
	return ids
}

func styleOperationIDs(items []StylePatchMutation) []string {
	ids := make([]string, len(items))
	for index := range items {
		ids[index] = items[index].OperationID
	}
	return ids
}

func layoutOperationIDs(items []LayoutMutation) []string {
	ids := make([]string, len(items))
	for index := range items {
		ids[index] = items[index].OperationID
	}
	return ids
}

type nativeMutationSheetIndex struct {
	sheet *NativeWorkbookSheetV1
	cells map[cellKey]*NativeWorkbookCellV1
}

func indexNativeMutationSheets(workbook *NativeWorkbookV1) map[string]nativeMutationSheetIndex {
	result := make(map[string]nativeMutationSheetIndex, len(workbook.Sheets))
	for sheetIndex := range workbook.Sheets {
		sheet := &workbook.Sheets[sheetIndex]
		cells := make(map[cellKey]*NativeWorkbookCellV1, len(sheet.Cells))
		for cellIndex := range sheet.Cells {
			cell := &sheet.Cells[cellIndex]
			cells[cellKey{row: cell.Row, column: cell.Column}] = cell
		}
		result[sheet.ID] = nativeMutationSheetIndex{sheet: sheet, cells: cells}
	}
	return result
}

func preflightNativeWorkbookMutationTargets(workbook *NativeWorkbookV1, transaction NativeWorkbookMutationTransactionV1) error {
	sheets := indexNativeMutationSheets(workbook)
	mergedTargets := make(map[string]map[cellKey]string)
	addMergedTarget := func(sheetID, operationID string, key cellKey) {
		if mergedTargets[sheetID] == nil {
			mergedTargets[sheetID] = map[cellKey]string{}
		}
		if _, exists := mergedTargets[sheetID][key]; !exists {
			mergedTargets[sheetID][key] = operationID
		}
	}
	requireSheet := func(operationID, sheetID string) (nativeMutationSheetIndex, error) {
		indexed, ok := sheets[sheetID]
		if !ok {
			return nativeMutationSheetIndex{}, fmt.Errorf("xlsxpatch: native mutation: operation %q references missing sheet %q", operationID, sheetID)
		}
		if !indexed.sheet.Editable {
			code := "UNSUPPORTED_SHEET"
			if indexed.sheet.RefusalCode != nil {
				code = *indexed.sheet.RefusalCode
			}
			return nativeMutationSheetIndex{}, fmt.Errorf("xlsxpatch: native mutation: operation %q targets mutation-refused sheet %q (%s)", operationID, sheetID, code)
		}
		return indexed, nil
	}
	for _, mutation := range transaction.Cells {
		indexed, err := requireSheet(mutation.OperationID, mutation.SheetID)
		if err != nil {
			return err
		}
		addMergedTarget(mutation.SheetID, mutation.OperationID, cellKey{row: mutation.Cell.Row, column: mutation.Cell.Column})
		if cell := indexed.cells[cellKey{row: mutation.Cell.Row, column: mutation.Cell.Column}]; cell != nil && !cell.Editable {
			return fmt.Errorf("xlsxpatch: native mutation: operation %q targets mutation-refused cell %s!%s", mutation.OperationID, mutation.SheetID, cell.Ref)
		}
	}
	for _, mutation := range transaction.Styles {
		indexed, err := requireSheet(mutation.OperationID, mutation.SheetID)
		if err != nil {
			return err
		}
		for row := mutation.Range.Row; row <= mutation.Range.EndRow; row++ {
			for column := mutation.Range.Column; column <= mutation.Range.EndColumn; column++ {
				addMergedTarget(mutation.SheetID, mutation.OperationID, cellKey{row: row, column: column})
				styleID := uint32(0)
				if cell := indexed.cells[cellKey{row: row, column: column}]; cell != nil {
					if !cell.Editable {
						return fmt.Errorf("xlsxpatch: native mutation: operation %q targets mutation-refused cell %s!%s", mutation.OperationID, mutation.SheetID, cell.Ref)
					}
					styleID = cell.StyleID
				}
				if int(styleID) >= len(workbook.Styles) || workbook.Styles[styleID].Effective.Projection != "full" {
					return fmt.Errorf("xlsxpatch: native mutation: operation %q targets unsupported style %d at %s!%s", mutation.OperationID, styleID, mutation.SheetID, cellReference(row, column))
				}
			}
		}
	}
	for _, mutation := range transaction.Layout {
		if _, err := requireSheet(mutation.OperationID, mutation.SheetID); err != nil {
			return err
		}
	}
	mergedSheetIDs := make([]string, 0, len(mergedTargets))
	for sheetID := range mergedTargets {
		mergedSheetIDs = append(mergedSheetIDs, sheetID)
	}
	sort.Strings(mergedSheetIDs)
	for _, sheetID := range mergedSheetIDs {
		targets := mergedTargets[sheetID]
		if err := preflightNativeMergedMutationTargets(sheets[sheetID].sheet, targets); err != nil {
			return err
		}
	}
	return nil
}

// preflightNativeMergedMutationTargets checks sparse/blank targets as well as
// modeled cells without expanding a merged rectangle. Both the merged-range
// and transaction target inventories are bounded, and the row sweep is
// O((ranges+targets) log ExcelColumns).
func preflightNativeMergedMutationTargets(sheet *NativeWorkbookSheetV1, targets map[cellKey]string) error {
	if len(sheet.MergedRanges) == 0 || len(targets) == 0 {
		return nil
	}
	type mergedEvent struct {
		row, minimumColumn, maximumColumn, delta int
	}
	events := make([]mergedEvent, 0, len(sheet.MergedRanges)*2)
	for _, merged := range sheet.MergedRanges {
		events = append(events,
			mergedEvent{row: merged.Row, minimumColumn: merged.Column, maximumColumn: merged.EndColumn, delta: 1},
			mergedEvent{row: merged.EndRow + 1, minimumColumn: merged.Column, maximumColumn: merged.EndColumn, delta: -1},
		)
	}
	sort.Slice(events, func(left, right int) bool {
		if events[left].row != events[right].row {
			return events[left].row < events[right].row
		}
		if events[left].delta != events[right].delta {
			return events[left].delta < events[right].delta
		}
		if events[left].minimumColumn != events[right].minimumColumn {
			return events[left].minimumColumn < events[right].minimumColumn
		}
		return events[left].maximumColumn < events[right].maximumColumn
	})
	orderedTargets := make([]cellKey, 0, len(targets))
	for key := range targets {
		orderedTargets = append(orderedTargets, key)
	}
	sort.Slice(orderedTargets, func(left, right int) bool {
		if orderedTargets[left].row != orderedTargets[right].row {
			return orderedTargets[left].row < orderedTargets[right].row
		}
		return orderedTargets[left].column < orderedTargets[right].column
	})
	tree, eventIndex := newNativeColumnRangeTree(excelMaxColumns), 0
	for _, target := range orderedTargets {
		for eventIndex < len(events) && events[eventIndex].row <= target.row {
			item := events[eventIndex]
			tree.add(item.minimumColumn, item.maximumColumn, item.delta)
			eventIndex++
		}
		if tree.maximum(target.column, target.column) > 0 {
			return fmt.Errorf("xlsxpatch: native mutation: operation %q targets mutation-refused merged cell %s!%s", targets[target], sheet.ID, cellReference(target.row, target.column))
		}
	}
	return nil
}

type nativeExpectedStyleProjections map[string]map[cellKey]NativeWorkbookEffectiveStyleV1

func expectedNativeStyleMutationProjections(original []byte, workbook *NativeWorkbookV1, mutations []StylePatchMutation) (nativeExpectedStyleProjections, error) {
	result := nativeExpectedStyleProjections{}
	if len(mutations) == 0 {
		return result, nil
	}
	pkg, err := openNativeWorkbookPackage(original)
	if err != nil {
		return nil, err
	}
	read := func(name string) ([]byte, bool) { data, ok := pkg.files[name]; return data, ok }
	workbookPart, err := locateWorkbookPartBytes(pkg.index, read)
	if err != nil {
		return nil, err
	}
	stylesPart, err := locateStylesPart(pkg.index, read, workbookPart)
	if err != nil {
		return nil, err
	}
	registry, err := newStyleRegistry(pkg.files[stylesPart])
	if err != nil {
		return nil, err
	}
	sheets := indexNativeMutationSheets(workbook)
	final := map[string]map[cellKey]StyleDelta{}
	for _, mutation := range mutations {
		if final[mutation.SheetID] == nil {
			final[mutation.SheetID] = map[cellKey]StyleDelta{}
		}
		for row := mutation.Range.Row; row <= mutation.Range.EndRow; row++ {
			for column := mutation.Range.Column; column <= mutation.Range.EndColumn; column++ {
				key := cellKey{row: row, column: column}
				final[mutation.SheetID][key] = mergeStyleDelta(final[mutation.SheetID][key], mutation.Style)
			}
		}
	}
	for sheetID, cells := range final {
		result[sheetID] = map[cellKey]NativeWorkbookEffectiveStyleV1{}
		for _, key := range sortedStyleKeys(cells) {
			sourceStyle := 0
			if cell := sheets[sheetID].cells[key]; cell != nil {
				sourceStyle = int(cell.StyleID)
			}
			resolved, err := registry.resolveStyle(sourceStyle, cells[key])
			if err != nil {
				return nil, fmt.Errorf("%s!%s: %w", sheetID, cellReference(key.row, key.column), err)
			}
			result[sheetID][key] = projectNativeWorkbookStyle(registry, resolved)
		}
	}
	return result, nil
}

func verifyNativeWorkbookMutationResult(before, after *NativeWorkbookV1, transaction NativeWorkbookMutationTransactionV1, expectedStyles nativeExpectedStyleProjections) error {
	if err := verifyNativeWorkbookSheetTopologyPreserved(before, after); err != nil {
		return err
	}
	if err := verifyNativeWorkbookMergedRangesPreserved(before, after); err != nil {
		return err
	}
	if err := verifyNativeWorkbookStyleTablePreserved(before, after); err != nil {
		return err
	}
	beforeSheets, afterSheets := indexNativeMutationSheets(before), indexNativeMutationSheets(after)
	if err := verifyNativeUnsupportedInventoryPreserved(before.Unsupported, after.Unsupported); err != nil {
		return err
	}
	cellTouched := make(map[string]map[cellKey]bool)
	styleTouched := make(map[string]map[cellKey]bool)
	allowedAfterOnly := make(map[string]map[cellKey]bool)
	for _, mutation := range transaction.Cells {
		if cellTouched[mutation.SheetID] == nil {
			cellTouched[mutation.SheetID] = map[cellKey]bool{}
		}
		if allowedAfterOnly[mutation.SheetID] == nil {
			allowedAfterOnly[mutation.SheetID] = map[cellKey]bool{}
		}
		key := cellKey{row: mutation.Cell.Row, column: mutation.Cell.Column}
		cellTouched[mutation.SheetID][key], allowedAfterOnly[mutation.SheetID][key] = true, true
	}
	for _, mutation := range transaction.Styles {
		if styleTouched[mutation.SheetID] == nil {
			styleTouched[mutation.SheetID] = map[cellKey]bool{}
		}
		if allowedAfterOnly[mutation.SheetID] == nil {
			allowedAfterOnly[mutation.SheetID] = map[cellKey]bool{}
		}
		for row := mutation.Range.Row; row <= mutation.Range.EndRow; row++ {
			for column := mutation.Range.Column; column <= mutation.Range.EndColumn; column++ {
				key := cellKey{row: row, column: column}
				styleTouched[mutation.SheetID][key], allowedAfterOnly[mutation.SheetID][key] = true, true
			}
		}
	}
	for sheetID, indexed := range beforeSheets {
		for key, cell := range indexed.cells {
			if cellTouched[sheetID][key] {
				continue
			}
			candidate := afterSheets[sheetID].cells[key]
			if candidate == nil || candidate.Editable != cell.Editable {
				return fmt.Errorf("xlsxpatch: native mutation: cell authority changed outside mutation at %s!%s", sheetID, cell.Ref)
			}
			if candidate == nil || !reflect.DeepEqual(cell.Value, candidate.Value) || !reflect.DeepEqual(cell.Formula, candidate.Formula) || !reflect.DeepEqual(cell.OOXMLType, candidate.OOXMLType) {
				return fmt.Errorf("xlsxpatch: native mutation: lexical content changed outside cell mutation at %s!%s", sheetID, cell.Ref)
			}
		}
		for key, cell := range indexed.cells {
			if styleTouched[sheetID][key] {
				continue
			}
			candidate := afterSheets[sheetID].cells[key]
			if (candidate == nil && cell.StyleID != 0) || (candidate != nil && candidate.StyleID != cell.StyleID) {
				return fmt.Errorf("xlsxpatch: native mutation: style changed outside style mutation at %s!%s", sheetID, cell.Ref)
			}
		}
	}
	for sheetID, indexed := range afterSheets {
		for key, cell := range indexed.cells {
			_, existedBefore := beforeSheets[sheetID].cells[key]
			if !existedBefore && !allowedAfterOnly[sheetID][key] {
				return fmt.Errorf("xlsxpatch: native mutation: unexpected after-only cell at %s!%s", sheetID, cell.Ref)
			}
			if allowedAfterOnly[sheetID][key] && !cell.Editable {
				return fmt.Errorf("xlsxpatch: native mutation: result cell became mutation-refused at %s!%s", sheetID, cell.Ref)
			}
		}
	}
	if err := verifyNativeWorkbookDimensionsPreserved(beforeSheets, afterSheets, transaction.Layout); err != nil {
		return err
	}
	if err := verifyNativeCellMutationPostconditions(afterSheets, transaction.Cells); err != nil {
		return err
	}
	if err := verifyNativeLayoutMutationPostconditions(afterSheets, transaction.Layout); err != nil {
		return err
	}
	if err := verifyNativeStyleMutationPostconditions(after, afterSheets, expectedStyles); err != nil {
		return err
	}
	return nil
}

func verifyNativeWorkbookMergedRangesPreserved(before, after *NativeWorkbookV1) error {
	for index := range before.Sheets {
		if !reflect.DeepEqual(before.Sheets[index].MergedRanges, after.Sheets[index].MergedRanges) {
			return fmt.Errorf("xlsxpatch: native mutation: merged-range projection changed on sheet %s", before.Sheets[index].ID)
		}
	}
	return nil
}

func verifyNativeWorkbookStyleTablePreserved(before, after *NativeWorkbookV1) error {
	if len(after.Styles) < len(before.Styles) {
		return fmt.Errorf("xlsxpatch: native mutation: style table shrank from %d to %d records", len(before.Styles), len(after.Styles))
	}
	for index := range before.Styles {
		if !reflect.DeepEqual(before.Styles[index], after.Styles[index]) {
			return fmt.Errorf("xlsxpatch: native mutation: existing style %d changed", index)
		}
	}
	return nil
}

func verifyNativeWorkbookDimensionsPreserved(before, after map[string]nativeMutationSheetIndex, mutations []LayoutMutation) error {
	rows, columns := map[string]map[int]bool{}, map[string]map[int]bool{}
	for _, mutation := range mutations {
		if mutation.Kind == RowSetHeight {
			if rows[mutation.SheetID] == nil {
				rows[mutation.SheetID] = map[int]bool{}
			}
			rows[mutation.SheetID][mutation.Row] = true
		} else {
			if columns[mutation.SheetID] == nil {
				columns[mutation.SheetID] = map[int]bool{}
			}
			columns[mutation.SheetID][mutation.Column] = true
		}
	}
	for sheetID, source := range before {
		candidate := after[sheetID]
		if err := verifyNativeWorkbookRowsPreserved(sheetID, source.sheet.Rows, candidate.sheet.Rows, rows[sheetID]); err != nil {
			return err
		}
		if err := verifyNativeWorkbookColumnsPreserved(sheetID, source.sheet.Columns, candidate.sheet.Columns, columns[sheetID]); err != nil {
			return err
		}
	}
	return nil
}

func verifyNativeWorkbookRowsPreserved(sheetID string, before, after []NativeWorkbookRowDimensionV1, touched map[int]bool) error {
	if len(touched) == 0 {
		if !reflect.DeepEqual(before, after) {
			return fmt.Errorf("xlsxpatch: native mutation: row dimensions changed without a row mutation on sheet %s", sheetID)
		}
		return nil
	}
	beforeByRow, afterByRow := map[int]NativeWorkbookRowDimensionV1{}, map[int]NativeWorkbookRowDimensionV1{}
	for _, row := range before {
		beforeByRow[row.Row] = row
	}
	for _, row := range after {
		afterByRow[row.Row] = row
	}
	for row, source := range beforeByRow {
		candidate, ok := afterByRow[row]
		if !touched[row] {
			if !ok || !reflect.DeepEqual(source, candidate) {
				return fmt.Errorf("xlsxpatch: native mutation: row dimension changed outside layout mutation at %s!%d", sheetID, row)
			}
		} else if !ok || !reflect.DeepEqual(source.StyleID, candidate.StyleID) {
			return fmt.Errorf("xlsxpatch: native mutation: row style changed during layout mutation at %s!%d", sheetID, row)
		}
	}
	for row, candidate := range afterByRow {
		if _, ok := beforeByRow[row]; ok {
			continue
		}
		if !touched[row] || candidate.StyleID != nil {
			return fmt.Errorf("xlsxpatch: native mutation: unexpected row dimension at %s!%d", sheetID, row)
		}
	}
	return nil
}

type nativeColumnDimensionSemantics struct {
	Width       *float64
	Hidden      bool
	CustomWidth bool
	BestFit     bool
	StyleID     *uint32
}

func verifyNativeWorkbookColumnsPreserved(sheetID string, before, after []NativeWorkbookColumnDimensionV1, touched map[int]bool) error {
	if len(touched) == 0 {
		if !reflect.DeepEqual(before, after) {
			return fmt.Errorf("xlsxpatch: native mutation: column dimensions changed without a column mutation on sheet %s", sheetID)
		}
		return nil
	}
	beforeCursor, afterCursor := 0, 0
	for column := 0; column < excelMaxColumns; column++ {
		source := nativeColumnDimensionAt(before, column, &beforeCursor)
		candidate := nativeColumnDimensionAt(after, column, &afterCursor)
		if touched[column] {
			if source.BestFit != candidate.BestFit || !reflect.DeepEqual(source.StyleID, candidate.StyleID) {
				return fmt.Errorf("xlsxpatch: native mutation: column metadata changed during layout mutation at %s!%d", sheetID, column)
			}
			continue
		}
		if !reflect.DeepEqual(source, candidate) {
			return fmt.Errorf("xlsxpatch: native mutation: column dimension changed outside layout mutation at %s!%d", sheetID, column)
		}
	}
	return nil
}

func nativeColumnDimensionAt(dimensions []NativeWorkbookColumnDimensionV1, column int, cursor *int) nativeColumnDimensionSemantics {
	for *cursor < len(dimensions) && dimensions[*cursor].EndColumn < column {
		*cursor = *cursor + 1
	}
	if *cursor < len(dimensions) {
		dimension := dimensions[*cursor]
		if column >= dimension.Column && column <= dimension.EndColumn {
			return nativeColumnDimensionSemantics{
				Width: dimension.Width, Hidden: dimension.Hidden, CustomWidth: dimension.CustomWidth,
				BestFit: dimension.BestFit, StyleID: dimension.StyleID,
			}
		}
	}
	return nativeColumnDimensionSemantics{}
}

func verifyNativeWorkbookSheetTopologyPreserved(before, after *NativeWorkbookV1) error {
	if len(before.Sheets) != len(after.Sheets) {
		return fmt.Errorf("xlsxpatch: native mutation: sheet topology changed from %d to %d sheets", len(before.Sheets), len(after.Sheets))
	}
	for index := range before.Sheets {
		left, right := before.Sheets[index], after.Sheets[index]
		if left.ID != right.ID || left.Name != right.Name || left.Order != right.Order || left.State != right.State || left.PartName != right.PartName ||
			left.Editable != right.Editable || !reflect.DeepEqual(left.RefusalCode, right.RefusalCode) {
			return fmt.Errorf("xlsxpatch: native mutation: sheet topology changed at order %d", index)
		}
	}
	return nil
}

func verifyNativeUnsupportedInventoryPreserved(before, after []NativeWorkbookUnsupportedV1) error {
	if len(after) != len(before) {
		return fmt.Errorf("xlsxpatch: native mutation: unsupported inventory changed from %d to %d items", len(before), len(after))
	}
	index := func(label string, items []NativeWorkbookUnsupportedV1) (map[string]NativeWorkbookUnsupportedV1, error) {
		result := make(map[string]NativeWorkbookUnsupportedV1, len(items))
		for _, item := range items {
			if _, duplicate := result[item.ID]; duplicate {
				return nil, fmt.Errorf("xlsxpatch: native mutation: %s unsupported inventory duplicates id %q", label, item.ID)
			}
			result[item.ID] = item
		}
		return result, nil
	}
	beforeByID, err := index("source", before)
	if err != nil {
		return err
	}
	afterByID, err := index("result", after)
	if err != nil {
		return err
	}
	for id, item := range beforeByID {
		candidate, ok := afterByID[id]
		if !ok || !reflect.DeepEqual(candidate, item) {
			return fmt.Errorf("xlsxpatch: native mutation: source-authoritative unsupported item %q was lost or remapped", id)
		}
	}
	for id := range afterByID {
		if _, ok := beforeByID[id]; !ok {
			return fmt.Errorf("xlsxpatch: native mutation: result introduced unsupported item %q", id)
		}
	}
	return nil
}

func verifyNativeCellMutationPostconditions(sheets map[string]nativeMutationSheetIndex, mutations []CellMutation) error {
	final := map[string]map[cellKey]normalizedCellMutation{}
	normalized, _ := validateCellMutations(mutations)
	for _, mutation := range normalized {
		if final[mutation.SheetID] == nil {
			final[mutation.SheetID] = map[cellKey]normalizedCellMutation{}
		}
		final[mutation.SheetID][cellKey{row: mutation.Cell.Row, column: mutation.Cell.Column}] = mutation
	}
	for sheetID, cells := range final {
		for key, mutation := range cells {
			cell := sheets[sheetID].cells[key]
			switch mutation.Kind {
			case CellClearValue, CellClearFormula:
				if cell != nil && (cell.Value != nil || cell.Formula != nil) {
					return fmt.Errorf("xlsxpatch: native mutation: clear postcondition failed at %s!%s", sheetID, cellReference(key.row, key.column))
				}
			case CellSetFormula:
				if cell == nil || cell.Formula == nil || cell.Formula.Type != "normal" || cell.Formula.Text != mutation.payload.text || cell.Formula.Cached != nil || cell.Value != nil {
					return fmt.Errorf("xlsxpatch: native mutation: formula postcondition failed at %s!%s", sheetID, cellReference(key.row, key.column))
				}
			case CellSetValue:
				if cell == nil || cell.Formula != nil || cell.Value == nil || !nativeMutationLiteralMatches(cell.Value, mutation.payload) {
					return fmt.Errorf("xlsxpatch: native mutation: value postcondition failed at %s!%s", sheetID, cellReference(key.row, key.column))
				}
			}
		}
	}
	return nil
}

func nativeMutationLiteralMatches(value *NativeWorkbookValueV1, payload cellPayload) bool {
	switch payload.literalKind {
	case literalString:
		return value.Kind == "string" && value.Storage == "inline" && value.Text != nil && *value.Text == payload.text
	case literalNumber:
		return value.Kind == "number" && value.Storage == "number" && value.Lexical != nil && *value.Lexical == payload.number
	case literalBoolean:
		lexical := "0"
		if payload.boolean {
			lexical = "1"
		}
		return value.Kind == "boolean" && value.Storage == "boolean" && value.Lexical != nil && *value.Lexical == lexical
	default:
		return false
	}
}

func verifyNativeLayoutMutationPostconditions(sheets map[string]nativeMutationSheetIndex, mutations []LayoutMutation) error {
	rows := map[string]map[int]float64{}
	columns := map[string]map[int]float64{}
	for _, mutation := range mutations {
		if mutation.Kind == RowSetHeight {
			if rows[mutation.SheetID] == nil {
				rows[mutation.SheetID] = map[int]float64{}
			}
			rows[mutation.SheetID][mutation.Row] = mutation.HeightPoints
		} else {
			if columns[mutation.SheetID] == nil {
				columns[mutation.SheetID] = map[int]float64{}
			}
			columns[mutation.SheetID][mutation.Column] = mutation.Width
		}
	}
	for sheetID, expected := range rows {
		for row, height := range expected {
			matched := false
			for _, dimension := range sheets[sheetID].sheet.Rows {
				if dimension.Row == row && dimension.HeightPoints != nil && *dimension.HeightPoints == height && dimension.Hidden == (height == 0) && dimension.CustomHeight {
					matched = true
				}
			}
			if !matched {
				return fmt.Errorf("xlsxpatch: native mutation: row layout postcondition failed at %s!%d", sheetID, row)
			}
		}
	}
	for sheetID, expected := range columns {
		for column, width := range expected {
			matched := false
			for _, dimension := range sheets[sheetID].sheet.Columns {
				if column >= dimension.Column && column <= dimension.EndColumn && dimension.Width != nil && *dimension.Width == width && dimension.Hidden == (width == 0) && dimension.CustomWidth {
					matched = true
				}
			}
			if !matched {
				return fmt.Errorf("xlsxpatch: native mutation: column layout postcondition failed at %s!%d", sheetID, column)
			}
		}
	}
	return nil
}

func verifyNativeStyleMutationPostconditions(workbook *NativeWorkbookV1, sheets map[string]nativeMutationSheetIndex, expected nativeExpectedStyleProjections) error {
	for sheetID, cells := range expected {
		for key, projection := range cells {
			cell := sheets[sheetID].cells[key]
			if cell == nil || int(cell.StyleID) >= len(workbook.Styles) || !reflect.DeepEqual(workbook.Styles[cell.StyleID].Effective, projection) {
				return fmt.Errorf("xlsxpatch: native mutation: style postcondition failed at %s!%s", sheetID, cellReference(key.row, key.column))
			}
		}
	}
	return nil
}

func verifyNativeStyleClearPostconditions(produced []byte, workbook *NativeWorkbookV1, mutations []StylePatchMutation) error {
	if len(mutations) == 0 {
		return nil
	}
	final := map[string]map[cellKey]StyleDelta{}
	for _, mutation := range mutations {
		if final[mutation.SheetID] == nil {
			final[mutation.SheetID] = map[cellKey]StyleDelta{}
		}
		for row := mutation.Range.Row; row <= mutation.Range.EndRow; row++ {
			for column := mutation.Range.Column; column <= mutation.Range.EndColumn; column++ {
				key := cellKey{row: row, column: column}
				final[mutation.SheetID][key] = mergeStyleDelta(final[mutation.SheetID][key], mutation.Style)
			}
		}
	}
	if !nativeStyleDeltasContainClear(final) {
		return nil
	}
	pkg, err := openNativeWorkbookPackage(produced)
	if err != nil {
		return fmt.Errorf("xlsxpatch: native mutation: clear-style postcondition package: %w", err)
	}
	read := func(name string) ([]byte, bool) { data, ok := pkg.files[name]; return data, ok }
	workbookPart, err := locateWorkbookPartBytes(pkg.index, read)
	if err != nil {
		return fmt.Errorf("xlsxpatch: native mutation: clear-style postcondition workbook: %w", err)
	}
	stylesPart, err := locateStylesPart(pkg.index, read, workbookPart)
	if err != nil {
		return fmt.Errorf("xlsxpatch: native mutation: clear-style postcondition styles: %w", err)
	}
	registry, err := newStyleRegistry(pkg.files[stylesPart])
	if err != nil {
		return fmt.Errorf("xlsxpatch: native mutation: clear-style postcondition styles: %w", err)
	}
	sheets := indexNativeMutationSheets(workbook)
	for sheetID, cells := range final {
		for key, delta := range cells {
			cell := sheets[sheetID].cells[key]
			if cell == nil || int(cell.StyleID) >= len(registry.cellXfs) {
				return fmt.Errorf("xlsxpatch: native mutation: clear-style target missing at %s!%s", sheetID, cellReference(key.row, key.column))
			}
			if err := verifyNativeStyleClearDelta(registry, int(cell.StyleID), delta); err != nil {
				return fmt.Errorf("xlsxpatch: native mutation: clear-style postcondition failed at %s!%s: %w", sheetID, cell.Ref, err)
			}
		}
	}
	return nil
}

func nativeStyleDeltasContainClear(deltas map[string]map[cellKey]StyleDelta) bool {
	for _, cells := range deltas {
		for _, delta := range cells {
			if nativeStyleDeltaContainsClear(delta) {
				return true
			}
		}
	}
	return false
}

func nativeStyleDeltaContainsClear(delta StyleDelta) bool {
	return stylePropertyCleared(delta.NumberFormat) || stylePropertyCleared(delta.FontName) || stylePropertyCleared(delta.FontSizePoints) ||
		stylePropertyCleared(delta.Bold) || stylePropertyCleared(delta.Italic) || stylePropertyCleared(delta.FontColor) ||
		stylePropertyCleared(delta.FillColor) || stylePropertyCleared(delta.HorizontalAlignment) ||
		stylePropertyCleared(delta.VerticalAlignment) || stylePropertyCleared(delta.WrapText)
}

func stylePropertyCleared[T any](property StyleProperty[T]) bool {
	return property.Present && property.Value == nil
}

func verifyNativeStyleClearDelta(registry *styleRegistry, styleID int, delta StyleDelta) error {
	source := registry.cellXfs[styleID]
	base := effectiveCellStyleXF(registry.styleXfs[source.xfID])
	actualFontID := effectiveStyleComponent(source.fontID, base.fontID, source.applyFont)
	actualFillID := effectiveStyleComponent(source.fillID, base.fillID, source.applyFill)
	actualNumFmtID := effectiveStyleComponent(source.numFmtID, base.numFmtID, source.applyNumberFormat)
	actualFont, baseFont := registry.fonts[actualFontID], registry.fonts[base.fontID]
	if stylePropertyCleared(delta.NumberFormat) && (actualNumFmtID != base.numFmtID || nativeStyleApplyEnabled(source.applyNumberFormat)) {
		return fmt.Errorf("number_format remains directly applied")
	}
	if stylePropertyCleared(delta.FillColor) && (actualFillID != base.fillID || nativeStyleApplyEnabled(source.applyFill)) {
		return fmt.Errorf("fill_color remains directly applied")
	}
	if stylePropertyCleared(delta.FontName) && !reflect.DeepEqual(actualFont.name, baseFont.name) {
		return fmt.Errorf("font_name does not match inherited value")
	}
	if stylePropertyCleared(delta.FontSizePoints) && !reflect.DeepEqual(actualFont.size, baseFont.size) {
		return fmt.Errorf("font_size_points does not match inherited value")
	}
	if stylePropertyCleared(delta.Bold) && actualFont.bold != baseFont.bold {
		return fmt.Errorf("bold does not match inherited value")
	}
	if stylePropertyCleared(delta.Italic) && actualFont.italic != baseFont.italic {
		return fmt.Errorf("italic does not match inherited value")
	}
	if stylePropertyCleared(delta.FontColor) && (actualFont.colorSafe != baseFont.colorSafe || !reflect.DeepEqual(actualFont.color, baseFont.color)) {
		return fmt.Errorf("font_color does not match inherited value")
	}
	if stylePropertyCleared(delta.FontName) && stylePropertyCleared(delta.FontSizePoints) && stylePropertyCleared(delta.Bold) &&
		stylePropertyCleared(delta.Italic) && stylePropertyCleared(delta.FontColor) &&
		(actualFontID != base.fontID || nativeStyleApplyEnabled(source.applyFont)) {
		return fmt.Errorf("fully cleared font remains directly applied")
	}
	actualAlignment := base.alignment
	if nativeStyleApplyEnabled(source.applyAlignment) {
		actualAlignment = source.alignment
	}
	if stylePropertyCleared(delta.HorizontalAlignment) && effectiveHorizontal(actualAlignment.horizontal, nil) != effectiveHorizontal(base.alignment.horizontal, nil) {
		return fmt.Errorf("horizontal_alignment does not match inherited value")
	}
	if stylePropertyCleared(delta.VerticalAlignment) && effectiveVertical(actualAlignment.vertical, nil) != effectiveVertical(base.alignment.vertical, nil) {
		return fmt.Errorf("vertical_alignment does not match inherited value")
	}
	if stylePropertyCleared(delta.WrapText) && effectiveWrap(actualAlignment.wrap, nil) != effectiveWrap(base.alignment.wrap, nil) {
		return fmt.Errorf("wrap_text does not match inherited value")
	}
	if stylePropertyCleared(delta.HorizontalAlignment) && stylePropertyCleared(delta.VerticalAlignment) && stylePropertyCleared(delta.WrapText) && nativeStyleApplyEnabled(source.applyAlignment) {
		return fmt.Errorf("fully cleared alignment remains directly applied")
	}
	return nil
}

func nativeStyleApplyEnabled(value *bool) bool { return value != nil && *value }

func verifyNativeWorkbookTransactionRawPreservation(original, produced []byte, workbook *NativeWorkbookV1, transaction NativeWorkbookMutationTransactionV1) error {
	allowed := map[string]bool{}
	for _, mutation := range transaction.Cells {
		allowed[workbookSheetPart(workbook, mutation.SheetID)] = true
	}
	for _, mutation := range transaction.Styles {
		allowed[workbookSheetPart(workbook, mutation.SheetID)] = true
	}
	for _, mutation := range transaction.Layout {
		allowed[workbookSheetPart(workbook, mutation.SheetID)] = true
	}

	beforePackage, err := openNativeWorkbookPackage(original)
	if err != nil {
		return err
	}
	read := func(name string) ([]byte, bool) { data, ok := beforePackage.files[name]; return data, ok }
	workbookPart, err := locateWorkbookPartBytes(beforePackage.index, read)
	if err != nil {
		return err
	}
	if len(transaction.Styles) != 0 {
		stylesPart, err := locateStylesPart(beforePackage.index, read, workbookPart)
		if err != nil {
			return err
		}
		allowed[stylesPart] = true
	}
	if len(transaction.Cells) != 0 {
		allowed[workbookPart.part] = true
		_, calcTarget, found, err := workbookRelationshipsWithoutCalcChain(beforePackage.files[workbookPart.relsPart], workbookPart.baseDir)
		if err != nil {
			return err
		}
		if found {
			allowed[workbookPart.relsPart], allowed[beforePackage.contentTypesPart] = true, true
			calcPart, _, ok := beforePackage.index.lookupResolved(calcTarget)
			if !ok {
				return fmt.Errorf("xlsxpatch: native mutation: missing calc-chain part %q", calcTarget)
			}
			allowed[calcPart] = true
		}
	}
	return compareUntouchedRawZipEntries(original, produced, allowed)
}

func workbookSheetPart(workbook *NativeWorkbookV1, sheetID string) string {
	for _, sheet := range workbook.Sheets {
		if sheet.ID == sheetID {
			return sheet.PartName
		}
	}
	return ""
}

func compareUntouchedRawZipEntries(original, produced []byte, allowed map[string]bool) error {
	originalZip, err := zip.NewReader(bytes.NewReader(original), int64(len(original)))
	if err != nil {
		return err
	}
	producedZip, err := zip.NewReader(bytes.NewReader(produced), int64(len(produced)))
	if err != nil {
		return err
	}
	producedByName := make(map[string]*zip.File, len(producedZip.File))
	for _, file := range producedZip.File {
		producedByName[file.Name] = file
	}
	seen := make(map[string]bool, len(originalZip.File))
	for _, source := range originalZip.File {
		seen[source.Name] = true
		if allowed[source.Name] {
			continue
		}
		candidate := producedByName[source.Name]
		if candidate == nil || !sameRawZipHeader(source, candidate) {
			return fmt.Errorf("xlsxpatch: native mutation: untouched OPC entry %q metadata changed", source.Name)
		}
		sourceRaw, err := source.OpenRaw()
		if err != nil {
			return err
		}
		candidateRaw, err := candidate.OpenRaw()
		if err != nil {
			return err
		}
		left, err := io.ReadAll(sourceRaw)
		if err != nil {
			return err
		}
		right, err := io.ReadAll(candidateRaw)
		if err != nil {
			return err
		}
		if !bytes.Equal(left, right) {
			return fmt.Errorf("xlsxpatch: native mutation: untouched OPC entry %q raw bytes changed", source.Name)
		}
	}
	for _, file := range producedZip.File {
		if !seen[file.Name] && !allowed[file.Name] {
			return fmt.Errorf("xlsxpatch: native mutation: unexpected OPC entry %q was added", file.Name)
		}
	}
	return nil
}

func sameRawZipHeader(left, right *zip.File) bool {
	return left.Name == right.Name && left.Method == right.Method && left.Flags == right.Flags && left.CRC32 == right.CRC32 &&
		left.CompressedSize64 == right.CompressedSize64 && left.UncompressedSize64 == right.UncompressedSize64 &&
		left.Comment == right.Comment && bytes.Equal(left.Extra, right.Extra) && left.Modified.Equal(right.Modified) &&
		left.ModifiedTime == right.ModifiedTime && left.ModifiedDate == right.ModifiedDate && left.ExternalAttrs == right.ExternalAttrs &&
		left.CreatorVersion == right.CreatorVersion && left.ReaderVersion == right.ReaderVersion && left.NonUTF8 == right.NonUTF8
}
