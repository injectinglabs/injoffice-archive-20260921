package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/xml"
	"fmt"
	"io"
	"sort"
	"strconv"
	"strings"
	"unicode/utf8"
)

// XLSXMaxOutlineDepth is the largest outlineLevel representable by
// SpreadsheetML. Excel's eight outline UI states include the ungrouped level;
// only seven nested groups can be stored in row/column outlineLevel metadata.
const XLSXMaxOutlineDepth = 7

const maxOutlineMaterializedRows = 100_000

// XLSXOutlineAxis identifies row or column outline metadata.
type XLSXOutlineAxis string

const (
	XLSXOutlineRows    XLSXOutlineAxis = "row"
	XLSXOutlineColumns XLSXOutlineAxis = "column"
)

// XLSXOutlineGroup is a zero-based, inclusive, properly nested group.
type XLSXOutlineGroup struct {
	ID        string          `json:"id"`
	SheetID   string          `json:"sheet_id"`
	Axis      XLSXOutlineAxis `json:"axis"`
	Start     int             `json:"start"`
	End       int             `json:"end"`
	Collapsed bool            `json:"collapsed"`
}

// XLSXOutlineBand exposes the exact effective metadata found on one or more
// adjacent rows or columns. Bands with identical metadata are coalesced.
type XLSXOutlineBand struct {
	Start        int  `json:"start"`
	End          int  `json:"end"`
	OutlineLevel int  `json:"outline_level"`
	Hidden       bool `json:"hidden"`
	Collapsed    bool `json:"collapsed"`
}

// WorksheetOutlineSnapshot is the native outline projection for one sheet.
// SummaryBelow and SummaryRight include the SpreadsheetML defaults (true)
// when sheetPr/outlinePr omits either property.
type WorksheetOutlineSnapshot struct {
	SheetID      string             `json:"sheet_id"`
	SummaryBelow bool               `json:"summary_below"`
	SummaryRight bool               `json:"summary_right"`
	Rows         []XLSXOutlineBand  `json:"rows"`
	Columns      []XLSXOutlineBand  `json:"columns"`
	Groups       []XLSXOutlineGroup `json:"groups"`
}

// WorksheetOutlineWrite replaces the outline-owned metadata for one sheet.
// Independently hidden level-zero rows and columns remain untouched.
type WorksheetOutlineWrite struct {
	SheetID      string             `json:"sheet_id"`
	SummaryBelow bool               `json:"summary_below"`
	SummaryRight bool               `json:"summary_right"`
	Groups       []XLSXOutlineGroup `json:"groups"`
}

type outlineCellState struct {
	level     int
	hidden    bool
	collapsed bool
}

// ReadWorksheetOutline reads raw outline bands and reconstructs a canonical,
// deterministic group tree without interpreting unrelated worksheet content.
func ReadWorksheetOutline(orig []byte, sheetID string) (WorksheetOutlineSnapshot, error) {
	part, data, err := readWorksheetPart(orig, sheetID)
	if err != nil {
		return WorksheetOutlineSnapshot{}, fmt.Errorf("xlsxpatch: read outlines: %w", err)
	}
	snapshot, err := readWorksheetOutlineXML(data, sheetID)
	if err != nil {
		return WorksheetOutlineSnapshot{}, fmt.Errorf("xlsxpatch: read outlines: worksheet %q: %w", part, err)
	}
	return snapshot, nil
}

// ApplyWorksheetOutline writes the complete group snapshot atomically. Only
// the target worksheet part is replaced; Apply raw-copies and verifies every
// unrelated OPC part.
func ApplyWorksheetOutline(orig []byte, write WorksheetOutlineWrite) ([]byte, error) {
	groups, rowStates, columnStates, err := validateOutlineWrite(write)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: write outlines: %w", err)
	}
	part, data, err := readWorksheetPart(orig, write.SheetID)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: write outlines: %w", err)
	}
	updated, err := patchWorksheetOutlineXML(data, write.SummaryBelow, write.SummaryRight, rowStates, columnStates)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: write outlines: worksheet %q: %w", part, err)
	}
	if bytes.Equal(updated, data) {
		return bytes.Clone(orig), nil
	}
	output, err := Apply(orig, Patch{Replace: map[string][]byte{part: updated}})
	if err != nil {
		return nil, err
	}
	// Re-read the result and prove that the serialized metadata reconstructs
	// the requested groups rather than accepting a merely well-formed XML edit.
	reopened, err := ReadWorksheetOutline(output, write.SheetID)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: write outlines: verify projection: %w", err)
	}
	if reopened.SummaryBelow != write.SummaryBelow || reopened.SummaryRight != write.SummaryRight || !outlineGroupsEquivalent(reopened.Groups, groups) {
		return nil, fmt.Errorf("xlsxpatch: write outlines: serialized projection does not match requested groups")
	}
	return output, nil
}

func readWorksheetPart(orig []byte, sheetID string) (string, []byte, error) {
	if sheetID == "" || !utf8.ValidString(sheetID) || utf16Length(sheetID) > maxStableSheetIDLen || strings.TrimSpace(sheetID) != sheetID {
		return "", nil, fmt.Errorf("invalid sheet_id")
	}
	zr, err := zip.NewReader(bytes.NewReader(orig), int64(len(orig)))
	if err != nil {
		return "", nil, fmt.Errorf("open original: %w", err)
	}
	index, err := newOPCPackageIndex(zr)
	if err != nil {
		return "", nil, err
	}
	cache := make(map[string]string)
	var readErr error
	read := func(name string) (string, bool) {
		if value, ok := cache[name]; ok {
			return value, true
		}
		file, ok := index.byExact[name]
		if !ok || readErr != nil {
			return "", false
		}
		value, err := readZipFile(file)
		if err != nil {
			readErr = err
			return "", false
		}
		cache[name] = string(value)
		return string(value), true
	}
	workbook, err := locateWorkbookPart(index, read)
	if readErr != nil {
		return "", nil, readErr
	}
	if err != nil {
		return "", nil, err
	}
	part, err := worksheetPartForID(index, read, workbook, sheetID)
	if readErr != nil {
		return "", nil, readErr
	}
	if err != nil {
		return "", nil, err
	}
	file, ok := index.byExact[part]
	if !ok {
		return "", nil, fmt.Errorf("worksheet part %q is missing", part)
	}
	data, err := readZipFile(file)
	if err != nil {
		return "", nil, fmt.Errorf("read worksheet part %q: %w", part, err)
	}
	return part, data, nil
}

func readWorksheetOutlineXML(data []byte, sheetID string) (WorksheetOutlineSnapshot, error) {
	index, err := indexWorksheetLayout(data)
	if err != nil {
		return WorksheetOutlineSnapshot{}, err
	}
	summaryBelow, summaryRight, _, _, _, err := indexOutlineProperties(data, index.namespace)
	if err != nil {
		return WorksheetOutlineSnapshot{}, err
	}
	rows := make([]XLSXOutlineBand, 0)
	rowStates := make(map[int]outlineCellState)
	for _, row := range index.rows {
		state, present, err := parseOutlineState(data[row.start:row.startTagEnd])
		if err != nil {
			return WorksheetOutlineSnapshot{}, fmt.Errorf("row %d: %w", row.row+1, err)
		}
		if present {
			rowStates[row.row] = state
			rows = appendOutlineBand(rows, row.row, row.row, state)
		}
	}
	columns := make([]XLSXOutlineBand, 0)
	columnStates := make(map[int]outlineCellState)
	if index.cols != nil {
		for _, column := range index.cols.columns {
			state, present, err := parseOutlineState(data[column.start:column.startTagEnd])
			if err != nil {
				return WorksheetOutlineSnapshot{}, fmt.Errorf("column %d:%d: %w", column.min, column.max, err)
			}
			if !present {
				continue
			}
			columns = appendOutlineBand(columns, column.min-1, column.max-1, state)
			for columnIndex := column.min - 1; columnIndex < column.max; columnIndex++ {
				columnStates[columnIndex] = state
			}
		}
	}
	rowGroups, err := reconstructOutlineGroups(sheetID, XLSXOutlineRows, rowStates, summaryBelow, excelMaxRows)
	if err != nil {
		return WorksheetOutlineSnapshot{}, err
	}
	columnGroups, err := reconstructOutlineGroups(sheetID, XLSXOutlineColumns, columnStates, summaryRight, excelMaxColumns)
	if err != nil {
		return WorksheetOutlineSnapshot{}, err
	}
	groups := append(rowGroups, columnGroups...)
	return WorksheetOutlineSnapshot{SheetID: sheetID, SummaryBelow: summaryBelow, SummaryRight: summaryRight, Rows: rows, Columns: columns, Groups: groups}, nil
}

func parseOutlineState(tag []byte) (outlineCellState, bool, error) {
	start, err := decodeStartElement(tag)
	if err != nil {
		return outlineCellState{}, false, err
	}
	levelText, levelFound, err := unqualifiedXMLAttribute(start, "outlineLevel")
	if err != nil {
		return outlineCellState{}, false, err
	}
	level := 0
	if levelFound {
		if levelText == "" || (len(levelText) > 1 && levelText[0] == '0') {
			return outlineCellState{}, false, fmt.Errorf("outlineLevel=%q is not canonical", levelText)
		}
		for _, character := range levelText {
			if character < '0' || character > '9' {
				return outlineCellState{}, false, fmt.Errorf("outlineLevel=%q is not canonical", levelText)
			}
		}
		level, err = strconv.Atoi(levelText)
		if err != nil || level < 0 || level > XLSXMaxOutlineDepth {
			return outlineCellState{}, false, fmt.Errorf("outlineLevel=%q is outside 0..%d", levelText, XLSXMaxOutlineDepth)
		}
	}
	hiddenText, hiddenFound, err := unqualifiedXMLAttribute(start, "hidden")
	if err != nil {
		return outlineCellState{}, false, err
	}
	hidden, err := ooxmlBoolean(hiddenText, hiddenFound)
	if err != nil {
		return outlineCellState{}, false, fmt.Errorf("invalid hidden=%q: %w", hiddenText, err)
	}
	collapsedText, collapsedFound, err := unqualifiedXMLAttribute(start, "collapsed")
	if err != nil {
		return outlineCellState{}, false, err
	}
	collapsed, err := ooxmlBoolean(collapsedText, collapsedFound)
	if err != nil {
		return outlineCellState{}, false, fmt.Errorf("invalid collapsed=%q: %w", collapsedText, err)
	}
	return outlineCellState{level: level, hidden: hidden, collapsed: collapsed}, levelFound || hiddenFound || collapsedFound, nil
}

func appendOutlineBand(bands []XLSXOutlineBand, start, end int, state outlineCellState) []XLSXOutlineBand {
	if len(bands) > 0 {
		last := &bands[len(bands)-1]
		if last.End+1 == start && last.OutlineLevel == state.level && last.Hidden == state.hidden && last.Collapsed == state.collapsed {
			last.End = end
			return bands
		}
	}
	return append(bands, XLSXOutlineBand{Start: start, End: end, OutlineLevel: state.level, Hidden: state.hidden, Collapsed: state.collapsed})
}

func reconstructOutlineGroups(sheetID string, axis XLSXOutlineAxis, states map[int]outlineCellState, summaryAfter bool, maximum int) ([]XLSXOutlineGroup, error) {
	groups := make([]XLSXOutlineGroup, 0)
	consumedCollapsed := make(map[int]bool)
	seenRanges := make(map[[2]int]bool)
	for level := 1; level <= XLSXMaxOutlineDepth; level++ {
		indices := make([]int, 0)
		for index, state := range states {
			if state.level >= level {
				indices = append(indices, index)
			}
		}
		sort.Ints(indices)
		for cursor := 0; cursor < len(indices); {
			start, end := indices[cursor], indices[cursor]
			cursor++
			for cursor < len(indices) && indices[cursor] == end+1 {
				end = indices[cursor]
				cursor++
			}
			key := [2]int{start, end}
			if seenRanges[key] {
				return nil, fmt.Errorf("%s outline levels produce duplicate range %d:%d", axis, start, end)
			}
			seenRanges[key] = true
			summary := end + 1
			if !summaryAfter {
				summary = start - 1
			}
			collapsed := false
			if summary >= 0 && summary < maximum {
				marker := states[summary]
				if marker.collapsed && marker.level == level-1 {
					for detail := start; detail <= end; detail++ {
						if !states[detail].hidden {
							return nil, fmt.Errorf("%s outline %d:%d has a collapsed summary but visible detail %d", axis, start, end, detail)
						}
					}
					collapsed = true
					consumedCollapsed[summary] = true
				}
			}
			groups = append(groups, XLSXOutlineGroup{ID: deterministicOutlineID(sheetID, axis, level, start, end), SheetID: sheetID, Axis: axis, Start: start, End: end, Collapsed: collapsed})
		}
	}
	for index, state := range states {
		if state.collapsed && !consumedCollapsed[index] {
			return nil, fmt.Errorf("%s %d has an orphan or ambiguous collapsed marker", axis, index)
		}
	}
	sort.Slice(groups, func(i, j int) bool {
		if groups[i].Start != groups[j].Start {
			return groups[i].Start < groups[j].Start
		}
		if groups[i].End != groups[j].End {
			return groups[i].End > groups[j].End
		}
		return groups[i].ID < groups[j].ID
	})
	return groups, nil
}

func deterministicOutlineID(sheetID string, axis XLSXOutlineAxis, level, start, end int) string {
	digest := sha256.Sum256([]byte(sheetID))
	return fmt.Sprintf("xlsx-%s-%s-%d-%d-%d", axis, hex.EncodeToString(digest[:6]), start, end, level)
}

func validateOutlineWrite(write WorksheetOutlineWrite) ([]XLSXOutlineGroup, map[int]outlineCellState, map[int]outlineCellState, error) {
	if write.SheetID == "" || !utf8.ValidString(write.SheetID) || utf16Length(write.SheetID) > maxStableSheetIDLen || strings.TrimSpace(write.SheetID) != write.SheetID {
		return nil, nil, nil, fmt.Errorf("invalid sheet_id")
	}
	if len(write.Groups) > maxCellMutations {
		return nil, nil, nil, fmt.Errorf("group collection exceeds %d entries", maxCellMutations)
	}
	groups := append([]XLSXOutlineGroup(nil), write.Groups...)
	ids := make(map[string]bool, len(groups))
	for index, group := range groups {
		if group.ID == "" || len(group.ID) > maxOperationIDLen || !restrictedIDPattern.MatchString(group.ID) {
			return nil, nil, nil, fmt.Errorf("group %d has invalid id %q", index, group.ID)
		}
		if ids[group.ID] {
			return nil, nil, nil, fmt.Errorf("duplicate group id %q", group.ID)
		}
		ids[group.ID] = true
		if group.SheetID != write.SheetID {
			return nil, nil, nil, fmt.Errorf("group %q sheet_id does not match write sheet_id", group.ID)
		}
		maximum := excelMaxRows
		if group.Axis == XLSXOutlineColumns {
			maximum = excelMaxColumns
		} else if group.Axis != XLSXOutlineRows {
			return nil, nil, nil, fmt.Errorf("group %q has invalid axis %q", group.ID, group.Axis)
		}
		if group.Start < 0 || group.End < group.Start || group.End >= maximum {
			return nil, nil, nil, fmt.Errorf("group %q range %d:%d is outside Excel limits", group.ID, group.Start, group.End)
		}
	}
	for i, left := range groups {
		depth := 1
		for j, right := range groups {
			if i == j || left.SheetID != right.SheetID || left.Axis != right.Axis {
				continue
			}
			overlaps := left.Start <= right.End && right.Start <= left.End
			leftContains := left.Start <= right.Start && left.End >= right.End
			rightContains := right.Start <= left.Start && right.End >= left.End
			if left.Start == right.Start && left.End == right.End {
				return nil, nil, nil, fmt.Errorf("groups %q and %q duplicate a range", left.ID, right.ID)
			}
			if overlaps && !leftContains && !rightContains {
				return nil, nil, nil, fmt.Errorf("groups %q and %q cross", left.ID, right.ID)
			}
			if right.Start <= left.Start && right.End >= left.End {
				depth++
			}
		}
		if depth > XLSXMaxOutlineDepth {
			return nil, nil, nil, fmt.Errorf("group %q requires depth %d; SpreadsheetML supports at most %d", left.ID, depth, XLSXMaxOutlineDepth)
		}
	}
	rowStates := make(map[int]outlineCellState)
	columnStates := make(map[int]outlineCellState)
	rowVisits := 0
	for _, group := range groups {
		states := rowStates
		maximum := excelMaxRows
		summaryAfter := write.SummaryBelow
		if group.Axis == XLSXOutlineColumns {
			states = columnStates
			maximum = excelMaxColumns
			summaryAfter = write.SummaryRight
		} else {
			rowVisits += group.End - group.Start + 1
			if rowVisits > maxOutlineMaterializedRows {
				return nil, nil, nil, fmt.Errorf("row outline expansion exceeds %d materialized records", maxOutlineMaterializedRows)
			}
		}
		for item := group.Start; item <= group.End; item++ {
			state := states[item]
			state.level++
			state.hidden = state.hidden || group.Collapsed
			states[item] = state
		}
		if group.Collapsed {
			summary := group.End + 1
			if !summaryAfter {
				summary = group.Start - 1
			}
			if summary < 0 || summary >= maximum {
				return nil, nil, nil, fmt.Errorf("collapsed group %q has no representable summary item", group.ID)
			}
			state := states[summary]
			if state.collapsed {
				return nil, nil, nil, fmt.Errorf("collapsed groups share summary item %d", summary)
			}
			state.collapsed = true
			states[summary] = state
		}
	}
	return groups, rowStates, columnStates, nil
}

func patchWorksheetOutlineXML(data []byte, summaryBelow, summaryRight bool, rowStates, columnStates map[int]outlineCellState) ([]byte, error) {
	index, err := indexWorksheetLayout(data)
	if err != nil {
		return nil, err
	}
	edits := make([]layoutTextEdit, 0)
	propertyEdit, err := buildOutlinePropertyEdit(data, index, summaryBelow, summaryRight)
	if err != nil {
		return nil, err
	}
	edits = append(edits, propertyEdit)
	rowEdits, err := buildRowOutlineEdits(data, index, rowStates)
	if err != nil {
		return nil, err
	}
	edits = append(edits, rowEdits...)
	columnEdit, changed, err := buildColumnOutlineEdit(data, index, columnStates)
	if err != nil {
		return nil, err
	}
	if changed {
		edits = append(edits, columnEdit)
	}
	sort.Slice(edits, func(i, j int) bool {
		if edits[i].start == edits[j].start {
			return edits[i].end > edits[j].end
		}
		return edits[i].start > edits[j].start
	})
	updated := bytes.Clone(data)
	for _, edit := range edits {
		updated = append(updated[:edit.start], append(edit.data, updated[edit.end:]...)...)
	}
	if err := validateWorksheetXML(updated); err != nil {
		return nil, fmt.Errorf("produced invalid worksheet XML: %w", err)
	}
	return updated, nil
}

type outlinePropertyIndex struct {
	sheetPr   *xmlSpan
	outlinePr *xmlSpan
	pageSetup *xmlSpan
}

func indexOutlineProperties(data []byte, namespace string) (bool, bool, *xmlSpan, *xmlSpan, *xmlSpan, error) {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	depth, sheetPrDepth := 0, 0
	properties := outlinePropertyIndex{}
	var activeSheetPr, activeOutline, activePage *xmlSpan
	var outlineStart *xml.StartElement
	for {
		before := int(decoder.InputOffset())
		token, err := decoder.Token()
		after := int(decoder.InputOffset())
		if err == io.EOF {
			break
		}
		if err != nil {
			return false, false, nil, nil, nil, err
		}
		switch token := token.(type) {
		case xml.StartElement:
			depth++
			if depth == 2 && token.Name.Space == namespace && token.Name.Local == "sheetPr" {
				if activeSheetPr != nil || properties.sheetPr != nil {
					return false, false, nil, nil, nil, fmt.Errorf("multiple sheetPr elements")
				}
				activeSheetPr = &xmlSpan{qname: rawQName(data, before, after), start: before, startTagEnd: after}
				sheetPrDepth = depth
			} else if sheetPrDepth != 0 && depth == sheetPrDepth+1 && token.Name.Space == namespace {
				switch token.Name.Local {
				case "outlinePr":
					if activeOutline != nil || properties.outlinePr != nil {
						return false, false, nil, nil, nil, fmt.Errorf("multiple outlinePr elements")
					}
					activeOutline = &xmlSpan{qname: rawQName(data, before, after), start: before, startTagEnd: after}
					copy := token
					outlineStart = &copy
				case "pageSetUpPr":
					if activePage != nil || properties.pageSetup != nil {
						return false, false, nil, nil, nil, fmt.Errorf("multiple pageSetUpPr elements")
					}
					activePage = &xmlSpan{qname: rawQName(data, before, after), start: before, startTagEnd: after}
				}
			}
		case xml.EndElement:
			if activeOutline != nil && depth == sheetPrDepth+1 && token.Name.Space == namespace && token.Name.Local == "outlinePr" {
				activeOutline.endStart, activeOutline.end = before, after
				properties.outlinePr, activeOutline = activeOutline, nil
			}
			if activePage != nil && depth == sheetPrDepth+1 && token.Name.Space == namespace && token.Name.Local == "pageSetUpPr" {
				activePage.endStart, activePage.end = before, after
				properties.pageSetup, activePage = activePage, nil
			}
			if activeSheetPr != nil && depth == sheetPrDepth && token.Name.Space == namespace && token.Name.Local == "sheetPr" {
				activeSheetPr.endStart, activeSheetPr.end = before, after
				properties.sheetPr, activeSheetPr = activeSheetPr, nil
				sheetPrDepth = 0
			}
			depth--
		}
	}
	below, right := true, true
	if outlineStart != nil {
		if properties.outlinePr == nil || len(bytes.TrimSpace(data[properties.outlinePr.startTagEnd:properties.outlinePr.endStart])) != 0 {
			return false, false, nil, nil, nil, fmt.Errorf("outlinePr has unsupported content")
		}
		value, found, err := unqualifiedXMLAttribute(*outlineStart, "summaryBelow")
		if err != nil {
			return false, false, nil, nil, nil, err
		}
		if found {
			below, err = ooxmlBoolean(value, true)
			if err != nil {
				return false, false, nil, nil, nil, fmt.Errorf("invalid outlinePr summaryBelow=%q", value)
			}
		}
		value, found, err = unqualifiedXMLAttribute(*outlineStart, "summaryRight")
		if err != nil {
			return false, false, nil, nil, nil, err
		}
		if found {
			right, err = ooxmlBoolean(value, true)
			if err != nil {
				return false, false, nil, nil, nil, fmt.Errorf("invalid outlinePr summaryRight=%q", value)
			}
		}
	}
	return below, right, properties.sheetPr, properties.outlinePr, properties.pageSetup, nil
}

func buildOutlinePropertyEdit(data []byte, index worksheetLayoutIndex, below, right bool) (layoutTextEdit, error) {
	_, _, sheetPr, outlinePr, pageSetup, err := indexOutlineProperties(data, index.namespace)
	if err != nil {
		return layoutTextEdit{}, err
	}
	value := func(flag bool) string {
		if flag {
			return "1"
		}
		return "0"
	}
	if outlinePr != nil {
		start := data[outlinePr.start:outlinePr.startTagEnd]
		start, err = rewriteUnqualifiedAttribute(start, "summaryBelow", stringPointer(value(below)))
		if err == nil {
			start, err = rewriteUnqualifiedAttribute(start, "summaryRight", stringPointer(value(right)))
		}
		if err != nil {
			return layoutTextEdit{}, err
		}
		return layoutTextEdit{start: outlinePr.start, end: outlinePr.startTagEnd, data: start}, nil
	}
	qname := prefixedLocal(index.root.qname, "outlinePr")
	element := []byte("<" + qname + ` summaryBelow="` + value(below) + `" summaryRight="` + value(right) + `"/>`)
	if sheetPr == nil {
		sheetQName := prefixedLocal(index.root.qname, "sheetPr")
		wrapped := append([]byte("<"+sheetQName+">"), element...)
		wrapped = append(wrapped, []byte("</"+sheetQName+">")...)
		return layoutTextEdit{start: index.root.startTagEnd, end: index.root.startTagEnd, data: wrapped}, nil
	}
	if sheetPr.selfClosing(data) {
		opened := openTag(data[sheetPr.start:sheetPr.startTagEnd])
		replacement := append(opened, element...)
		replacement = append(replacement, []byte("</"+sheetPr.qname+">")...)
		return layoutTextEdit{start: sheetPr.start, end: sheetPr.end, data: replacement}, nil
	}
	// outlinePr precedes pageSetUpPr in CT_SheetPr.
	insertAt := sheetPr.endStart
	if pageSetup != nil {
		insertAt = pageSetup.start
	}
	return layoutTextEdit{start: insertAt, end: insertAt, data: element}, nil
}

func buildRowOutlineEdits(data []byte, index worksheetLayoutIndex, desired map[int]outlineCellState) ([]layoutTextEdit, error) {
	existing := make(map[int]layoutRow, len(index.rows))
	targets := make(map[int]bool, len(desired))
	for item := range desired {
		targets[item] = true
	}
	for _, row := range index.rows {
		existing[row.row] = row
		state, present, err := parseOutlineState(data[row.start:row.startTagEnd])
		if err != nil {
			return nil, fmt.Errorf("row %d: %w", row.row+1, err)
		}
		if present && (state.level > 0 || state.collapsed) {
			targets[row.row] = true
		}
	}
	ordered := make([]int, 0, len(targets))
	for item := range targets {
		ordered = append(ordered, item)
	}
	sort.Ints(ordered)
	edits := make([]layoutTextEdit, 0, len(ordered))
	insertions := make(map[int][]int)
	for _, item := range ordered {
		if row, ok := existing[item]; ok {
			old, _, err := parseOutlineState(data[row.start:row.startTagEnd])
			if err != nil {
				return nil, err
			}
			updated, err := rewriteOutlineTag(data[row.start:row.startTagEnd], old, desired[item])
			if err != nil {
				return nil, fmt.Errorf("row %d: %w", item+1, err)
			}
			if !bytes.Equal(updated, data[row.start:row.startTagEnd]) {
				edits = append(edits, layoutTextEdit{start: row.start, end: row.startTagEnd, data: updated})
			}
			continue
		}
		if _, ok := desired[item]; !ok {
			continue
		}
		insertAt := index.sheetData.endStart
		for _, row := range index.rows {
			if row.row > item {
				insertAt = row.start
				break
			}
		}
		insertions[insertAt] = append(insertions[insertAt], item)
	}
	if index.sheetData.selfClosing(data) && len(insertions) > 0 {
		var rows bytes.Buffer
		for _, item := range ordered {
			state, ok := desired[item]
			if !ok {
				continue
			}
			built, err := newOutlineTag(prefixedLocal(index.sheetData.qname, "row"), "r", item+1, item+1, state)
			if err != nil {
				return nil, err
			}
			rows.Write(built)
		}
		opened := openTag(data[index.sheetData.start:index.sheetData.startTagEnd])
		replacement := append(opened, rows.Bytes()...)
		replacement = append(replacement, []byte("</"+index.sheetData.qname+">")...)
		return []layoutTextEdit{{start: index.sheetData.start, end: index.sheetData.end, data: replacement}}, nil
	}
	for offset, items := range insertions {
		sort.Ints(items)
		var rows bytes.Buffer
		for _, item := range items {
			built, err := newOutlineTag(prefixedLocal(index.sheetData.qname, "row"), "r", item+1, item+1, desired[item])
			if err != nil {
				return nil, err
			}
			rows.Write(built)
		}
		edits = append(edits, layoutTextEdit{start: offset, end: offset, data: rows.Bytes()})
	}
	return edits, nil
}

func rewriteOutlineTag(tag []byte, old, desired outlineCellState) ([]byte, error) {
	updated := bytes.Clone(tag)
	var err error
	if desired.level > 0 {
		updated, err = rewriteUnqualifiedAttribute(updated, "outlineLevel", stringPointer(strconv.Itoa(desired.level)))
	} else {
		updated, err = rewriteUnqualifiedAttribute(updated, "outlineLevel", nil)
	}
	if err != nil {
		return nil, err
	}
	if desired.hidden {
		updated, err = rewriteUnqualifiedAttribute(updated, "hidden", stringPointer("1"))
	} else if old.level > 0 {
		updated, err = rewriteUnqualifiedAttribute(updated, "hidden", nil)
	}
	if err != nil {
		return nil, err
	}
	if desired.collapsed {
		updated, err = rewriteUnqualifiedAttribute(updated, "collapsed", stringPointer("1"))
	} else {
		updated, err = rewriteUnqualifiedAttribute(updated, "collapsed", nil)
	}
	return updated, err
}

func newOutlineTag(qname, indexName string, minimum, maximum int, state outlineCellState) ([]byte, error) {
	tag := []byte("<" + qname + "/>")
	var err error
	if indexName == "r" {
		tag, err = rewriteUnqualifiedAttribute(tag, "r", stringPointer(strconv.Itoa(minimum)))
	} else {
		tag, err = rewriteUnqualifiedAttribute(tag, "min", stringPointer(strconv.Itoa(minimum)))
		if err == nil {
			tag, err = rewriteUnqualifiedAttribute(tag, "max", stringPointer(strconv.Itoa(maximum)))
		}
	}
	if err != nil {
		return nil, err
	}
	return rewriteOutlineTag(tag, outlineCellState{}, state)
}

func buildColumnOutlineEdit(data []byte, index worksheetLayoutIndex, desired map[int]outlineCellState) (layoutTextEdit, bool, error) {
	targets := make(map[int]bool, len(desired))
	for item := range desired {
		targets[item] = true
	}
	if index.cols != nil {
		for _, column := range index.cols.columns {
			old, present, err := parseOutlineState(data[column.start:column.startTagEnd])
			if err != nil {
				return layoutTextEdit{}, false, err
			}
			if present && (old.level > 0 || old.collapsed) {
				for item := column.min - 1; item < column.max; item++ {
					targets[item] = true
				}
			}
		}
	}
	if len(targets) == 0 {
		return layoutTextEdit{}, false, nil
	}
	ordered := make([]int, 0, len(targets))
	for item := range targets {
		ordered = append(ordered, item)
	}
	sort.Ints(ordered)
	colsQName := prefixedLocal(index.root.qname, "cols")
	colQName := prefixedLocal(colsQName, "col")
	if index.cols == nil {
		var out bytes.Buffer
		out.WriteString("<" + colsQName + ">")
		for _, item := range ordered {
			built, err := newOutlineTag(colQName, "min", item+1, item+1, desired[item])
			if err != nil {
				return layoutTextEdit{}, false, err
			}
			out.Write(built)
		}
		out.WriteString("</" + colsQName + ">")
		return layoutTextEdit{start: index.sheetData.start, end: index.sheetData.start, data: out.Bytes()}, true, nil
	}
	cols := *index.cols
	var out bytes.Buffer
	out.Write(data[cols.start:cols.startTagEnd])
	cursor, targetCursor := cols.startTagEnd, 0
	for _, column := range cols.columns {
		out.Write(data[cursor:column.start])
		for targetCursor < len(ordered) && ordered[targetCursor]+1 < column.min {
			item := ordered[targetCursor]
			built, err := newOutlineTag(colQName, "min", item+1, item+1, desired[item])
			if err != nil {
				return layoutTextEdit{}, false, err
			}
			out.Write(built)
			targetCursor++
		}
		first := targetCursor
		for targetCursor < len(ordered) && ordered[targetCursor]+1 <= column.max {
			targetCursor++
		}
		if first == targetCursor {
			out.Write(data[column.start:column.end])
		} else {
			old, _, err := parseOutlineState(data[column.start:column.startTagEnd])
			if err != nil {
				return layoutTextEdit{}, false, err
			}
			position := column.min - 1
			for _, item := range ordered[first:targetCursor] {
				if position < item {
					segment, err := rewriteColumnOutlineRange(data[column.start:column.startTagEnd], position, item-1, old, outlineCellState{})
					if err != nil {
						return layoutTextEdit{}, false, err
					}
					out.Write(segment)
				}
				segment, err := rewriteColumnOutlineRange(data[column.start:column.startTagEnd], item, item, old, desired[item])
				if err != nil {
					return layoutTextEdit{}, false, err
				}
				out.Write(segment)
				position = item + 1
			}
			if position < column.max {
				segment, err := rewriteColumnOutlineRange(data[column.start:column.startTagEnd], position, column.max-1, old, outlineCellState{})
				if err != nil {
					return layoutTextEdit{}, false, err
				}
				out.Write(segment)
			}
		}
		cursor = column.end
	}
	out.Write(data[cursor:cols.endStart])
	for targetCursor < len(ordered) {
		item := ordered[targetCursor]
		built, err := newOutlineTag(colQName, "min", item+1, item+1, desired[item])
		if err != nil {
			return layoutTextEdit{}, false, err
		}
		out.Write(built)
		targetCursor++
	}
	out.Write(data[cols.endStart:cols.end])
	if cols.selfClosing(data) {
		out.Reset()
		out.Write(openTag(data[cols.start:cols.startTagEnd]))
		for _, item := range ordered {
			built, err := newOutlineTag(colQName, "min", item+1, item+1, desired[item])
			if err != nil {
				return layoutTextEdit{}, false, err
			}
			out.Write(built)
		}
		out.WriteString("</" + cols.qname + ">")
	}
	return layoutTextEdit{start: cols.start, end: cols.end, data: out.Bytes()}, true, nil
}

func rewriteColumnOutlineRange(tag []byte, start, end int, old, desired outlineCellState) ([]byte, error) {
	updated, err := rewriteUnqualifiedAttribute(tag, "min", stringPointer(strconv.Itoa(start+1)))
	if err == nil {
		updated, err = rewriteUnqualifiedAttribute(updated, "max", stringPointer(strconv.Itoa(end+1)))
	}
	if err != nil {
		return nil, err
	}
	updated, err = canonicalEmptyElementTag(updated)
	if err != nil {
		return nil, err
	}
	return rewriteOutlineTag(updated, old, desired)
}

func outlineGroupsEquivalent(left, right []XLSXOutlineGroup) bool {
	if len(left) != len(right) {
		return false
	}
	normalize := func(groups []XLSXOutlineGroup) []XLSXOutlineGroup {
		copy := append([]XLSXOutlineGroup(nil), groups...)
		sort.Slice(copy, func(i, j int) bool {
			if copy[i].Axis != copy[j].Axis {
				return copy[i].Axis < copy[j].Axis
			}
			if copy[i].Start != copy[j].Start {
				return copy[i].Start < copy[j].Start
			}
			if copy[i].End != copy[j].End {
				return copy[i].End < copy[j].End
			}
			return !copy[i].Collapsed && copy[j].Collapsed
		})
		return copy
	}
	left, right = normalize(left), normalize(right)
	for index := range left {
		if left[index].SheetID != right[index].SheetID || left[index].Axis != right[index].Axis || left[index].Start != right[index].Start || left[index].End != right[index].End || left[index].Collapsed != right[index].Collapsed {
			return false
		}
	}
	return true
}
