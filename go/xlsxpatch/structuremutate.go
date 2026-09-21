package xlsxpatch

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"sort"
	"strconv"
	"strings"
	"unicode/utf8"
)

type StructureMutation struct {
	OperationID string `json:"operation_id"`
	SheetID     string `json:"sheet_id"`
	Kind        string `json:"kind"`
	Index       int    `json:"index"`
	Count       int    `json:"count"`
}

func (m StructureMutation) rows() bool     { return strings.HasPrefix(m.Kind, "row.") }
func (m StructureMutation) deleting() bool { return strings.HasSuffix(m.Kind, ".delete") }
func (m StructureMutation) limit() int {
	if m.rows() {
		return excelMaxRows
	}
	return excelMaxColumns
}
func validateStructureMutation(m StructureMutation) error {
	if !restrictedIDPattern.MatchString(m.OperationID) || len(m.OperationID) > maxOperationIDLen || m.SheetID == "" || strings.TrimSpace(m.SheetID) != m.SheetID || !utf8.ValidString(m.SheetID) || utf16Length(m.SheetID) > maxStableSheetIDLen {
		return fmt.Errorf("xlsxpatch: structure: invalid operation or sheet id")
	}
	if m.Kind != "row.insert" && m.Kind != "row.delete" && m.Kind != "column.insert" && m.Kind != "column.delete" {
		return fmt.Errorf("xlsxpatch: structure: invalid kind")
	}
	if m.Index < 0 || m.Index >= m.limit() || m.Count < 1 || m.Count > m.limit()-m.Index {
		return fmt.Errorf("xlsxpatch: structure: index/count exceeds worksheet bounds")
	}
	return nil
}

// shiftInterval contracts references across a deletion and expands them for
// insertion inside the range. It never wraps coordinates at worksheet limits.
func (m StructureMutation) shiftInterval(first, last int) (int, int, bool, error) {
	if !m.deleting() {
		if first >= m.Index {
			first += m.Count
		}
		if last >= m.Index {
			last += m.Count
		}
	} else {
		end := m.Index + m.Count
		if first >= m.Index && last < end {
			return 0, 0, false, nil
		}
		if first >= end {
			first -= m.Count
		} else if first >= m.Index {
			first = m.Index
		}
		if last >= end {
			last -= m.Count
		} else if last >= m.Index {
			last = m.Index - 1
		}
	}
	if first < 0 || last >= m.limit() {
		return 0, 0, false, fmt.Errorf("xlsxpatch: structure: occupied cells or references would exceed worksheet bounds")
	}
	return first, last, true, nil
}
func (m StructureMutation) shiftCell(row, column int) (int, int, bool, error) {
	index := column
	if m.rows() {
		index = row
	}
	next, _, keep, err := m.shiftInterval(index, index)
	if m.rows() {
		row = next
	} else {
		column = next
	}
	return row, column, keep, err
}
func (m StructureMutation) shiftRange(r StyleRange) (StyleRange, bool, error) {
	first, last := r.Column, r.EndColumn
	if m.rows() {
		first, last = r.Row, r.EndRow
	}
	first, last, keep, err := m.shiftInterval(first, last)
	if m.rows() {
		r.Row, r.EndRow = first, last
	} else {
		r.Column, r.EndColumn = first, last
	}
	return r, keep, err
}

func applyNativeStructureTransaction(original []byte, before *NativeWorkbookV1, m StructureMutation) (*NativeWorkbookMutationResultV1, error) {
	var target *NativeWorkbookSheetV1
	for i := range before.Sheets {
		if before.Sheets[i].ID == m.SheetID {
			target = &before.Sheets[i]
		}
	}
	if target == nil || !target.Editable {
		return nil, fmt.Errorf("xlsxpatch: structure: missing or read-only sheet")
	}
	// Fail closed on features whose coordinates are not modeled by this writer.
	// This includes names, tables, drawings, filters, validations and formula groups.
	for _, u := range before.Unsupported {
		if u.Capability == "styles" || u.Code == "RICH_SHARED_STRING" {
			continue
		}
		switch u.Code {
		case "WORKSHEET_DIMENSION_METADATA", "WORKBOOK_VIEW_METADATA", "MERGED_CELLS", "UNMODELED_WORKBOOK_FEATURE":
		default:
			return nil, fmt.Errorf("xlsxpatch: structure: cannot preserve references in %s", u.Code)
		}
	}
	source, err := openNativeWorkbookPackage(original)
	if err != nil {
		return nil, err
	}
	if err := validateStructureWorkbook(source.files[before.Source.WorkbookPart]); err != nil {
		return nil, err
	}
	formulas := []CellMutation{}
	for _, sheet := range before.Sheets {
		if !sheet.Editable {
			return nil, fmt.Errorf("xlsxpatch: structure: all referencing sheets must be editable")
		}
		for _, c := range sheet.Cells {
			if c.Formula == nil {
				continue
			}
			if !c.Editable || c.Formula.Type != "normal" {
				return nil, fmt.Errorf("xlsxpatch: structure: grouped or read-only formula")
			}
			if sheet.ID == m.SheetID {
				_, _, keep, err := m.shiftCell(c.Row, c.Column)
				if err != nil {
					return nil, err
				}
				if !keep {
					continue
				}
			}
			formula, err := shiftStructureFormula(c.Formula.Text, sheet.Name, target.Name, m)
			if err != nil {
				return nil, err
			}
			formulas = append(formulas, CellMutation{OperationID: fmt.Sprintf("structure-formula-%d", len(formulas)), SheetID: sheet.ID, Kind: CellSetFormula, Cell: CellRef{Row: c.Row, Column: c.Column}, Formula: "=" + formula})
		}
	}
	produced := original
	if len(formulas) > 0 {
		produced, err = ApplyCellMutations(produced, formulas)
		if err != nil {
			return nil, err
		}
	}
	pkg, err := openNativeWorkbookPackage(produced)
	if err != nil {
		return nil, err
	}
	updated, err := shiftStructureWorksheet(pkg.files[target.PartName], target, m)
	if err != nil {
		return nil, err
	}
	patch := Patch{Replace: map[string][]byte{target.PartName: updated}}
	read := func(name string) (string, bool) { value, ok := pkg.files[name]; return string(value), ok }
	location, err := locateWorkbookPart(pkg.index, read)
	if err != nil {
		return nil, err
	}
	if err := applyFullRecalculationPatch(pkg.index, read, location, &patch); err != nil {
		return nil, err
	}
	produced, err = Apply(produced, patch)
	if err != nil {
		return nil, err
	}
	after, err := ExtractNativeWorkbookV1WithOptions(produced, NativeWorkbookExtractionOptions{Previous: before})
	if err != nil {
		return nil, err
	}
	if issues := ValidateNativeWorkbookV1(after); len(issues) > 0 {
		return nil, &NativeWorkbookValidationError{Issues: issues}
	}
	if err := verifyNativeWorkbookSheetTopologyPreserved(before, after); err != nil {
		return nil, err
	}
	// Independently verify cell locations, lexical values, styles, and formulas.
	actualSheets := indexNativeMutationSheets(after)
	for i, sheet := range before.Sheets {
		actual := actualSheets[sheet.ID].cells
		kept := 0
		for _, c := range sheet.Cells {
			row, column, keep := c.Row, c.Column, true
			if sheet.ID == m.SheetID {
				row, column, keep, err = m.shiftCell(row, column)
				if err != nil {
					return nil, err
				}
			}
			if !keep {
				continue
			}
			kept++
			got := actual[cellKey{row: row, column: column}]
			if got == nil || got.StyleID != c.StyleID || !structureValuesEqual(got.Value, c.Value) {
				return nil, fmt.Errorf("xlsxpatch: structure: cell readback mismatch")
			}
			if c.Formula != nil {
				want, err := shiftStructureFormula(c.Formula.Text, sheet.Name, target.Name, m)
				if err != nil {
					return nil, err
				}
				if got.Formula == nil || got.Formula.Text != want || got.Formula.Cached != nil {
					return nil, fmt.Errorf("xlsxpatch: structure: formula readback mismatch")
				}
			} else if got.Formula != nil {
				return nil, fmt.Errorf("xlsxpatch: structure: unexpected formula")
			}
		}
		if len(after.Sheets[i].Cells) != kept {
			return nil, fmt.Errorf("xlsxpatch: structure: cell inventory mismatch")
		}
	}
	return &NativeWorkbookMutationResultV1{Package: produced, Workbook: after}, nil
}
func structureValuesEqual(a, b *NativeWorkbookValueV1) bool {
	if a == nil || b == nil {
		return a == b
	}
	return a.Kind == b.Kind && a.Storage == b.Storage && sameStringPointer(a.Text, b.Text) && sameStringPointer(a.Lexical, b.Lexical) && a.Rich == b.Rich
}
func sameStringPointer(a, b *string) bool {
	if a == nil || b == nil {
		return a == b
	}
	return *a == *b
}

type structureEdit struct {
	start, end int
	data       []byte
}

func applyStructureEdits(data []byte, edits []structureEdit) []byte {
	sort.Slice(edits, func(i, j int) bool { return edits[i].start < edits[j].start })
	result := make([]byte, 0, len(data))
	cursor := 0
	for _, edit := range edits {
		result = append(result, data[cursor:edit.start]...)
		result = append(result, edit.data...)
		cursor = edit.end
	}
	return append(result, data[cursor:]...)
}

func shiftStructureWorksheet(data []byte, sheet *NativeWorkbookSheetV1, m StructureMutation) ([]byte, error) {
	index, err := indexStyleWorksheet(data)
	if err != nil {
		return nil, err
	}
	edits := []structureEdit{}
	for _, row := range index.rows {
		nextRow := row.row
		if m.rows() {
			var keep bool
			nextRow, _, keep, err = m.shiftInterval(row.row, row.row)
			if err != nil {
				return nil, err
			}
			if !keep {
				edits = append(edits, structureEdit{row.start, row.end, nil})
				continue
			}
		}
		tag := bytes.Clone(data[row.start:row.startTagEnd])
		if m.rows() {
			ref := strconv.Itoa(nextRow + 1)
			tag, err = rewriteUnqualifiedAttribute(tag, "r", &ref)
		} else {
			tag, err = rewriteUnqualifiedAttribute(tag, "spans", nil)
		}
		if err != nil {
			return nil, err
		}
		edits = append(edits, structureEdit{row.start, row.startTagEnd, tag})
		for _, cell := range row.cells {
			r, c, keep, err := m.shiftCell(row.row, cell.column)
			if err != nil {
				return nil, err
			}
			if !keep {
				edits = append(edits, structureEdit{cell.start, cell.end, nil})
				continue
			}
			ref := cellReference(r, c)
			tag, err := rewriteUnqualifiedAttribute(data[cell.start:cell.startTagEnd], "r", &ref)
			if err != nil {
				return nil, err
			}
			edits = append(edits, structureEdit{cell.start, cell.startTagEnd, tag})
		}
	}
	if index.dimension != nil {
		r, c, er, ec, err := parseDimensionReference(index.dimensionRef)
		if err != nil {
			return nil, err
		}
		shifted, keep, err := m.shiftRange(StyleRange{r, c, er, ec})
		if err != nil {
			return nil, err
		}
		ref := "A1"
		if keep {
			ref = cellReference(shifted.Row, shifted.Column) + ":" + cellReference(shifted.EndRow, shifted.EndColumn)
		}
		tag, err := rewriteUnqualifiedAttribute(data[index.dimension.start:index.dimension.startTagEnd], "ref", &ref)
		if err != nil {
			return nil, err
		}
		edits = append(edits, structureEdit{index.dimension.start, index.dimension.startTagEnd, tag})
	}
	if !m.rows() {
		cols, _, err := directChildElements(data, "worksheet", "cols")
		if err != nil {
			return nil, err
		}
		if len(cols) > 0 {
			span := cols[0].span
			raw := data[span.start:span.end]
			children, _, err := directChildElements(raw, "cols", "col")
			if err != nil {
				return nil, err
			}
			colEdits := []structureEdit{}
			remaining := 0
			for _, child := range children {
				start := child.span
				tag := raw[start.start:start.startTagEnd]
				min, err := canonicalColumnIndexAttribute(tag, "min")
				if err != nil {
					return nil, err
				}
				max, err := canonicalColumnIndexAttribute(tag, "max")
				if err != nil {
					return nil, err
				}
				first, last, keep, err := m.shiftInterval(min-1, max-1)
				if err != nil {
					return nil, err
				}
				if !keep {
					colEdits = append(colEdits, structureEdit{start.start, start.end, nil})
					continue
				}
				remaining++
				lo, hi := strconv.Itoa(first+1), strconv.Itoa(last+1)
				tag, err = rewriteUnqualifiedAttribute(tag, "min", &lo)
				if err != nil {
					return nil, err
				}
				tag, err = rewriteUnqualifiedAttribute(tag, "max", &hi)
				if err != nil {
					return nil, err
				}
				colEdits = append(colEdits, structureEdit{start.start, start.startTagEnd, tag})
			}
			var updated []byte
			if remaining > 0 {
				updated = applyStructureEdits(raw, colEdits)
			}
			edits = append(edits, structureEdit{span.start, span.end, updated})
		}
	}
	result := applyStructureEdits(data, edits)
	if len(sheet.MergedRanges) > 0 {
		ranges := []NativeWorkbookMergedRangeV1{}
		for _, v := range sheet.MergedRanges {
			r, keep, err := m.shiftRange(StyleRange{v.Row, v.Column, v.EndRow, v.EndColumn})
			if err != nil {
				return nil, err
			}
			if !keep || (r.Row == r.EndRow && r.Column == r.EndColumn) {
				continue
			}
			v.Row, v.Column, v.EndRow, v.EndColumn = r.Row, r.Column, r.EndRow, r.EndColumn
			v.Ref = cellReference(r.Row, r.Column) + ":" + cellReference(r.EndRow, r.EndColumn)
			ranges = append(ranges, v)
		}
		result, err = rewriteWorksheetMerges(result, ranges)
		if err != nil {
			return nil, err
		}
	}
	return result, nil
}

func validateStructureWorkbook(data []byte) error {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	depth := 0
	namespace := ""
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			return nil
		}
		if err != nil {
			return err
		}
		switch token := token.(type) {
		case xml.StartElement:
			depth++
			if depth == 1 {
				namespace = token.Name.Space
			}
			if depth == 2 {
				if token.Name.Space != namespace {
					return fmt.Errorf("xlsxpatch: structure: foreign workbook references")
				}
				switch token.Name.Local {
				case "bookViews":
					qualified, err := consumeNativeChartDataMetadata(decoder, token, namespace)
					if err != nil {
						return err
					}
					if !qualified {
						return fmt.Errorf("xlsxpatch: structure: unmodeled workbook view references")
					}
					depth--
				case "sheets": // Already validated and routed by native extraction.
				case "workbookPr", "calcPr":
					names := "date1904 showObjects showBorderUnselectedTables filterPrivacy promptedSolutions showInk backupFile saveExternalLinkValues updateLinks codeName hidePivotFieldList showPivotChartFilter allowRefreshQuery autoCompressPictures refreshAllConnections defaultThemeVersion"
					if token.Name.Local == "calcPr" {
						names = "calcId calcMode fullCalcOnLoad refMode iterate iterateCount iterateDelta fullPrecision calcCompleted calcOnSave concurrentCalc concurrentManualCount forceFullCalc"
					}
					allowed := []xml.Name{}
					for _, name := range strings.Fields(names) {
						allowed = append(allowed, xml.Name{Local: name})
					}
					if err := requireOnlySemanticXMLAttributes(token, allowed...); err != nil {
						return fmt.Errorf("xlsxpatch: structure: %w", err)
					}
					if mode, _, err := unqualifiedXMLAttribute(token, "refMode"); err != nil {
						return err
					} else if mode != "" && mode != "A1" {
						return fmt.Errorf("xlsxpatch: structure: only A1 formula reference mode is supported")
					}
					if err := requireNativeEmptyElement(decoder, token); err != nil {
						return err
					}
					depth--
				default:
					return fmt.Errorf("xlsxpatch: structure: workbook contains unmodeled references in %s", token.Name.Local)
				}
			}
		case xml.EndElement:
			depth--
		}
	}
}
