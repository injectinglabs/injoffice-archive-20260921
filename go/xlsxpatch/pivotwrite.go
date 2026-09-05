package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"regexp"
	"strconv"
	"strings"
)

// Pivot WRITING: bake a native OOXML pivot table into the workbook. The
// strategy (used by long-standing SpreadsheetML tooling): write a
// pivotCacheDefinition with refreshOnLoad="1" and an EMPTY records part —
// Excel (and LibreOffice) rebuild the live pivot from the source range the
// moment the file opens, so the pivot is fully interactive there without us
// serializing cached record data. The InjOffice browser side independently
// bakes the same aggregation as real cells (PivotManager), so every consumer
// sees a result: dumb viewers see the baked grid, Excel sees a LIVE pivot.
//
// Composed as a Patch through Apply — untouched parts verified byte-identical,
// same as charts.

// PivotDataField is one value aggregation in the pivot.
type PivotDataField struct {
	Field string `json:"field"` // source column header
	Agg   string `json:"agg"`   // sum | count | avg | min | max (ChartSpec-style vocabulary)
}

type PivotPageField struct {
	Field        string  `json:"field"`
	SelectedItem *string `json:"selectedItem,omitempty"`
}

type PivotMemberFilter struct {
	Field         string   `json:"field"`
	ExcludedItems []string `json:"excludedItems"`
}

type PivotFieldSort struct {
	Field     string `json:"field"`
	Direction string `json:"direction"`
}

type PivotFieldMember struct {
	Value string `json:"value"`
	Kind  string `json:"kind"` // string | number | boolean | blank
}

type PivotFieldMembers struct {
	Field string             `json:"field"`
	Items []PivotFieldMember `json:"items"`
}

// PivotWriteSpec describes the pivot to bake.
type PivotWriteSpec struct {
	// SourceSheetName + SourceRef ("A1:D5", header row included) locate the data.
	SourceSheetName string
	SourceRef       string
	// Fields are the source column headers, in column order.
	Fields []string
	// TargetSheetName + TargetCellRef ("F1") place the pivot table.
	TargetSheetName string
	TargetCellRef   string
	RowFields       []string
	ColFields       []string
	DataFields      []PivotDataField
	PageFields      []PivotPageField
	MemberFilters   []PivotMemberFilter
	Sorts           []PivotFieldSort
	FieldMembers    []PivotFieldMembers
	GrandTotals     *bool
	Name            string // pivot table display name (default "InjOfficePivot")
}

var aggToSubtotal = map[string]string{
	"sum": "sum", "count": "count", "avg": "average", "min": "min", "max": "max",
}

var aggDisplay = map[string]string{
	"sum": "Sum", "count": "Count", "avg": "Average", "min": "Min", "max": "Max",
}

// AddPivot returns new workbook bytes containing the native pivot parts.
func AddPivot(orig []byte, spec PivotWriteSpec) ([]byte, error) {
	return addPivotAtPart(orig, spec, "")
}

func addPivotAtPart(orig []byte, spec PivotWriteSpec, preferredTablePart string) ([]byte, error) {
	if len(spec.Fields) == 0 {
		return nil, fmt.Errorf("xlsxpatch: pivot needs source fields")
	}
	if len(spec.DataFields) == 0 {
		return nil, fmt.Errorf("xlsxpatch: pivot needs at least one data field")
	}
	fieldIdx := map[string]int{}
	for i, f := range spec.Fields {
		if _, dup := fieldIdx[f]; !dup {
			fieldIdx[f] = i
		}
	}
	resolve := func(names []string, what string) ([]int, error) {
		out := make([]int, 0, len(names))
		for _, n := range names {
			i, ok := fieldIdx[n]
			if !ok {
				return nil, fmt.Errorf("xlsxpatch: pivot %s field %q not in source fields", what, n)
			}
			out = append(out, i)
		}
		return out, nil
	}
	rowIdx, err := resolve(spec.RowFields, "row")
	if err != nil {
		return nil, err
	}
	colIdx, err := resolve(spec.ColFields, "column")
	if err != nil {
		return nil, err
	}
	if len(colIdx) > 1 {
		return nil, fmt.Errorf("xlsxpatch: pivot supports at most one column field")
	}
	for _, df := range spec.DataFields {
		if _, ok := fieldIdx[df.Field]; !ok {
			return nil, fmt.Errorf("xlsxpatch: pivot data field %q not in source fields", df.Field)
		}
		if _, ok := aggToSubtotal[df.Agg]; !ok {
			return nil, fmt.Errorf("xlsxpatch: pivot aggregation %q not supported", df.Agg)
		}
	}
	if err := validatePivotState(spec, fieldIdx); err != nil {
		return nil, err
	}

	zr, err := zip.NewReader(bytes.NewReader(orig), int64(len(orig)))
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: add pivot: %w", err)
	}
	read := func(name string) (string, bool) {
		for _, f := range zr.File {
			if f.Name == name {
				rc, err := f.Open()
				if err != nil {
					return "", false
				}
				defer rc.Close()
				b, err := io.ReadAll(rc)
				if err != nil {
					return "", false
				}
				return string(b), true
			}
		}
		return "", false
	}

	targetSheetPart, err := worksheetPartFor(read, spec.TargetSheetName)
	if err != nil {
		return nil, err
	}
	// The source sheet must resolve too — fail before writing anything.
	if _, err := worksheetPartFor(read, spec.SourceSheetName); err != nil {
		return nil, err
	}
	wbXML, ok := read("xl/workbook.xml")
	if !ok {
		return nil, fmt.Errorf("xlsxpatch: missing xl/workbook.xml")
	}
	wbRels, ok := read("xl/_rels/workbook.xml.rels")
	if !ok {
		return nil, fmt.Errorf("xlsxpatch: missing workbook rels")
	}
	ct, ok := read("[Content_Types].xml")
	if !ok {
		return nil, fmt.Errorf("xlsxpatch: missing [Content_Types].xml")
	}

	patch := Patch{Replace: map[string][]byte{}, Add: map[string][]byte{}}
	targetSheetXML, ok := read(targetSheetPart)
	if !ok {
		return nil, fmt.Errorf("xlsxpatch: missing target worksheet %q", targetSheetPart)
	}

	cacheDefPart := nextFreePart(zr, "xl/pivotCache/pivotCacheDefinition", ".xml")
	cacheRecPart := nextFreePart(zr, "xl/pivotCache/pivotCacheRecords", ".xml")
	tablePart := preferredTablePart
	if tablePart == "" {
		tablePart = nextFreePart(zr, "xl/pivotTables/pivotTable", ".xml")
	} else {
		if _, err := canonicalOPCPartKey(tablePart); err != nil || !strings.HasPrefix(tablePart, "xl/pivotTables/") || !strings.HasSuffix(tablePart, ".xml") {
			return nil, fmt.Errorf("xlsxpatch: invalid preferred pivot table part %q", tablePart)
		}
		if _, exists := read(tablePart); exists {
			return nil, fmt.Errorf("xlsxpatch: preferred pivot table part %q already exists", tablePart)
		}
	}
	cacheID := nextFreePivotCacheID(wbXML)

	name := spec.Name
	if name == "" {
		name = "InjOfficePivot"
	}

	patch.Add[cacheDefPart] = []byte(pivotCacheDefXML(spec))
	patch.Add[cacheRecPart] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n" +
		`<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0"/>`)
	patch.Add[relsPartFor(cacheDefPart)] = []byte(relsDoc(relationshipXML("rId1", relTypePivotRecords, partRelTargetFrom(cacheDefPart, cacheRecPart))))
	patch.Add[tablePart] = []byte(pivotTableXML(spec, name, cacheID, rowIdx, colIdx, fieldIdx))
	patch.Add[relsPartFor(tablePart)] = []byte(relsDoc(relationshipXML("rId1", relTypePivotCacheDef, partRelTargetFrom(tablePart, cacheDefPart))))

	// Workbook: pivotCaches element + rel to the cache definition.
	wbRelID := nextFreeRelID(wbRels)
	newWbRels, err := appendRelationship(wbRels, wbRelID, relTypePivotCacheDef, strings.TrimPrefix(cacheDefPart, "xl/"))
	if err != nil {
		return nil, err
	}
	patch.Replace["xl/_rels/workbook.xml.rels"] = []byte(newWbRels)
	newWb, err := workbookWithPivotCache(wbXML, cacheID, wbRelID)
	if err != nil {
		return nil, err
	}
	patch.Replace["xl/workbook.xml"] = []byte(newWb)

	// Target worksheet rels: relationship to the pivot table (no sheet-XML
	// element needed — the relationship alone binds it).
	sheetRelsPart := relsPartFor(targetSheetPart)
	sheetRels, hasSheetRels := read(sheetRelsPart)
	if !hasSheetRels {
		sheetRels = emptyRelsXML
	}
	sheetRelID := nextFreeRelID(sheetRels)
	newSheetRels, err := appendRelationship(sheetRels, sheetRelID, relTypePivotTable, partRelTargetFrom(targetSheetPart, tablePart))
	if err != nil {
		return nil, err
	}
	if hasSheetRels {
		patch.Replace[sheetRelsPart] = []byte(newSheetRels)
	} else {
		patch.Add[sheetRelsPart] = []byte(newSheetRels)
	}
	newTargetSheet, err := worksheetWithPivotTable(targetSheetXML, sheetRelID)
	if err != nil {
		return nil, err
	}
	patch.Replace[targetSheetPart] = []byte(newTargetSheet)

	newCT, err := contentTypesWith(ct, map[string]string{
		"/" + cacheDefPart: ctPivotCacheDef,
		"/" + cacheRecPart: ctPivotRecords,
		"/" + tablePart:    ctPivotTable,
	})
	if err != nil {
		return nil, err
	}
	patch.Replace["[Content_Types].xml"] = []byte(newCT)

	return Apply(orig, patch)
}

func validatePivotState(spec PivotWriteSpec, fieldIdx map[string]int) error {
	axes := map[string]bool{}
	for _, field := range append(append([]string{}, spec.RowFields...), spec.ColFields...) {
		axes[field] = true
	}
	members := map[string][]PivotFieldMember{}
	for _, inventory := range spec.FieldMembers {
		if _, ok := fieldIdx[inventory.Field]; !ok {
			return fmt.Errorf("xlsxpatch: pivot member field %q not in source fields", inventory.Field)
		}
		if _, duplicate := members[inventory.Field]; duplicate {
			return fmt.Errorf("xlsxpatch: duplicate pivot member inventory for %q", inventory.Field)
		}
		seen := map[string]bool{}
		for _, item := range inventory.Items {
			if seen[item.Value] {
				return fmt.Errorf("xlsxpatch: duplicate pivot member %q in field %q", item.Value, inventory.Field)
			}
			seen[item.Value] = true
			switch item.Kind {
			case "string":
			case "number":
				if _, err := strconv.ParseFloat(item.Value, 64); err != nil {
					return fmt.Errorf("xlsxpatch: invalid numeric pivot member %q in field %q", item.Value, inventory.Field)
				}
			case "boolean":
				if item.Value != "true" && item.Value != "false" {
					return fmt.Errorf("xlsxpatch: invalid boolean pivot member %q in field %q", item.Value, inventory.Field)
				}
			case "blank":
				if item.Value != "" {
					return fmt.Errorf("xlsxpatch: blank pivot member in field %q has a value", inventory.Field)
				}
			default:
				return fmt.Errorf("xlsxpatch: unsupported pivot member kind %q in field %q", item.Kind, inventory.Field)
			}
		}
		members[inventory.Field] = inventory.Items
	}
	pageSeen := map[string]bool{}
	for _, page := range spec.PageFields {
		if _, ok := fieldIdx[page.Field]; !ok {
			return fmt.Errorf("xlsxpatch: pivot page field %q not in source fields", page.Field)
		}
		if pageSeen[page.Field] || axes[page.Field] {
			return fmt.Errorf("xlsxpatch: pivot page field %q is duplicated or also an axis field", page.Field)
		}
		pageSeen[page.Field] = true
		if page.SelectedItem != nil && memberIndex(members[page.Field], *page.SelectedItem) < 0 {
			return fmt.Errorf("xlsxpatch: pivot page selection %q not found in member inventory for %q", *page.SelectedItem, page.Field)
		}
	}
	filterSeen := map[string]bool{}
	for _, filter := range spec.MemberFilters {
		if _, ok := fieldIdx[filter.Field]; !ok {
			return fmt.Errorf("xlsxpatch: pivot filter field %q not in source fields", filter.Field)
		}
		if filterSeen[filter.Field] {
			return fmt.Errorf("xlsxpatch: duplicate pivot filter for %q", filter.Field)
		}
		filterSeen[filter.Field] = true
		for _, page := range spec.PageFields {
			if page.Field == filter.Field && page.SelectedItem != nil {
				return fmt.Errorf("xlsxpatch: pivot page field %q cannot combine a single selection and a member filter", filter.Field)
			}
		}
		inventory, ok := members[filter.Field]
		if !ok {
			return fmt.Errorf("xlsxpatch: pivot filter %q needs member inventory", filter.Field)
		}
		for _, excluded := range filter.ExcludedItems {
			if memberIndex(inventory, excluded) < 0 {
				return fmt.Errorf("xlsxpatch: excluded pivot member %q not found in field %q", excluded, filter.Field)
			}
		}
	}
	sortSeen := map[string]bool{}
	for _, rule := range spec.Sorts {
		if !axes[rule.Field] {
			return fmt.Errorf("xlsxpatch: pivot sort field %q is not a row or column field", rule.Field)
		}
		if sortSeen[rule.Field] || (rule.Direction != "ascending" && rule.Direction != "descending") {
			return fmt.Errorf("xlsxpatch: duplicate or unsupported pivot sort for %q", rule.Field)
		}
		sortSeen[rule.Field] = true
	}
	return nil
}

func memberIndex(items []PivotFieldMember, value string) int {
	for index, item := range items {
		if item.Value == value {
			return index
		}
	}
	return -1
}

var pivotTablePartsRe = regexp.MustCompile(`(?s)<pivotTableParts\b([^>]*)>(.*?)</pivotTableParts>`)
var pivotTablePartsEmptyRe = regexp.MustCompile(`<pivotTableParts\b([^>]*)/>`)
var countAttrRe = regexp.MustCompile(`\s+count="[0-9]+"`)

// worksheetWithPivotTable wires a pivot relationship into worksheet XML.
// The relationship alone is not enough: ECMA-376 requires a pivotTableParts
// collection in the worksheet that references it.
func worksheetWithPivotTable(sheetXML, relID string) (string, error) {
	entry := `<pivotTablePart xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="` + esc(relID) + `"/>`
	if match := pivotTablePartsRe.FindStringSubmatchIndex(sheetXML); match != nil {
		whole := sheetXML[match[0]:match[1]]
		openEnd := strings.Index(whole, ">")
		closeAt := strings.LastIndex(whole, "</pivotTableParts>")
		if openEnd < 0 || closeAt < 0 {
			return "", fmt.Errorf("xlsxpatch: malformed pivotTableParts")
		}
		open := whole[:openEnd+1]
		count := strings.Count(whole[openEnd+1:closeAt], "<pivotTablePart") + 1
		if countAttrRe.MatchString(open) {
			open = countAttrRe.ReplaceAllString(open, fmt.Sprintf(` count="%d"`, count))
		} else {
			open = strings.TrimSuffix(open, ">") + fmt.Sprintf(` count="%d">`, count)
		}
		replacement := open + whole[openEnd+1:closeAt] + entry + `</pivotTableParts>`
		return sheetXML[:match[0]] + replacement + sheetXML[match[1]:], nil
	}
	if match := pivotTablePartsEmptyRe.FindStringSubmatchIndex(sheetXML); match != nil {
		whole := sheetXML[match[0]:match[1]]
		open := strings.TrimSuffix(whole, "/>")
		open = countAttrRe.ReplaceAllString(open, "")
		replacement := open + ` count="1">` + entry + `</pivotTableParts>`
		return sheetXML[:match[0]] + replacement + sheetXML[match[1]:], nil
	}
	insertAt := strings.LastIndex(sheetXML, "</worksheet>")
	if insertAt < 0 {
		return "", fmt.Errorf("xlsxpatch: malformed target worksheet")
	}
	if i := strings.Index(sheetXML, "<extLst"); i >= 0 && i < insertAt {
		insertAt = i
	}
	collection := `<pivotTableParts count="1">` + entry + `</pivotTableParts>`
	return sheetXML[:insertAt] + collection + sheetXML[insertAt:], nil
}

const (
	relTypePivotCacheDef = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition"
	relTypePivotRecords  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords"
	relTypePivotTable    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"
	ctPivotCacheDef      = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"
	ctPivotRecords       = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"
	ctPivotTable         = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"
)

var pivotCacheIDRe = regexp.MustCompile(`cacheId="(\d+)"`)

func nextFreePivotCacheID(workbookXML string) int {
	max := 0
	for _, m := range pivotCacheIDRe.FindAllStringSubmatch(workbookXML, -1) {
		n := 0
		fmt.Sscanf(m[1], "%d", &n) //nolint:errcheck
		if n > max {
			max = n
		}
	}
	return max + 1
}

// workbookWithPivotCache inserts (or extends) <pivotCaches> in workbook.xml.
// Schema position: pivotCaches follows definedNames/calcPr region; inserting
// immediately before </workbook> is valid when none of the later elements
// (smartTagPr, webPublishing, extLst...) are present — extend before extLst
// when it exists.
func workbookWithPivotCache(wbXML string, cacheID int, relID string) (string, error) {
	// xmlns:r declared LOCALLY on the element: documents in the wild (openpyxl
	// among them) bind the r prefix on arbitrary inner elements rather than
	// the root, so a "does the file mention xmlns:r somewhere" check is a
	// false-positive machine — a local declaration is valid everywhere and
	// depends on nothing. (Caught by external openpyxl validation.)
	entry := fmt.Sprintf(`<pivotCache xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" cacheId="%d" r:id=%q/>`, cacheID, relID)
	if idx := strings.Index(wbXML, "<pivotCaches>"); idx >= 0 {
		return wbXML[:idx+len("<pivotCaches>")] + entry + wbXML[idx+len("<pivotCaches>"):], nil
	}
	insertAt := strings.LastIndex(wbXML, "</workbook>")
	if insertAt < 0 {
		return "", fmt.Errorf("xlsxpatch: malformed workbook.xml")
	}
	if i := strings.Index(wbXML, "<extLst"); i >= 0 && i < insertAt {
		insertAt = i
	}
	return wbXML[:insertAt] + `<pivotCaches>` + entry + `</pivotCaches>` + wbXML[insertAt:], nil
}

func pivotCacheDefXML(spec PivotWriteSpec) string {
	members := map[string][]PivotFieldMember{}
	for _, inventory := range spec.FieldMembers {
		members[inventory.Field] = inventory.Items
	}
	var fields strings.Builder
	for _, f := range spec.Fields {
		items := members[f]
		var shared strings.Builder
		for _, item := range items {
			switch item.Kind {
			case "string":
				fmt.Fprintf(&shared, `<s v=%q/>`, esc(item.Value))
			case "number":
				fmt.Fprintf(&shared, `<n v=%q/>`, esc(item.Value))
			case "boolean":
				value := "0"
				if item.Value == "true" {
					value = "1"
				}
				fmt.Fprintf(&shared, `<b v=%q/>`, value)
			case "blank":
				shared.WriteString(`<m/>`)
			}
		}
		if len(items) == 0 {
			fmt.Fprintf(&fields, `<cacheField name=%q numFmtId="0"><sharedItems/></cacheField>`, esc(f))
		} else {
			fmt.Fprintf(&fields, `<cacheField name=%q numFmtId="0"><sharedItems count="%d">%s</sharedItems></cacheField>`, esc(f), len(items), shared.String())
		}
	}
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n" +
		fmt.Sprintf(`<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" refreshOnLoad="1" refreshedBy="InjOffice" refreshedVersion="6" minRefreshableVersion="3" createdVersion="6" recordCount="0">`+
			`<cacheSource type="worksheet"><worksheetSource ref=%q sheet=%q/></cacheSource>`+
			`<cacheFields count="%d">%s</cacheFields>`+
			`</pivotCacheDefinition>`,
			esc(spec.SourceRef), esc(spec.SourceSheetName), len(spec.Fields), fields.String())
}

func pivotTableXML(spec PivotWriteSpec, name string, cacheID int, rowIdx, colIdx []int, fieldIdx map[string]int) string {
	axisFor := map[int]string{}
	for _, i := range rowIdx {
		axisFor[i] = "axisRow"
	}
	for _, i := range colIdx {
		axisFor[i] = "axisCol"
	}
	for _, page := range spec.PageFields {
		axisFor[fieldIdx[page.Field]] = "axisPage"
	}
	pageFor := map[int]PivotPageField{}
	for _, page := range spec.PageFields {
		pageFor[fieldIdx[page.Field]] = page
	}
	sortFor := map[int]string{}
	for _, rule := range spec.Sorts {
		sortFor[fieldIdx[rule.Field]] = rule.Direction
	}
	membersFor := map[int][]PivotFieldMember{}
	for _, inventory := range spec.FieldMembers {
		membersFor[fieldIdx[inventory.Field]] = inventory.Items
	}
	hiddenFor := map[int]map[string]bool{}
	for _, filter := range spec.MemberFilters {
		values := map[string]bool{}
		for _, value := range filter.ExcludedItems {
			values[value] = true
		}
		hiddenFor[fieldIdx[filter.Field]] = values
	}
	dataFieldSet := map[int]bool{}
	for _, df := range spec.DataFields {
		dataFieldSet[fieldIdx[df.Field]] = true
	}

	var pf strings.Builder
	for i := range spec.Fields {
		attrs := ""
		if axis, ok := axisFor[i]; ok {
			attrs += fmt.Sprintf(` axis=%q`, axis)
		}
		if dataFieldSet[i] {
			attrs += ` dataField="1"`
		}
		if direction := sortFor[i]; direction != "" {
			attrs += fmt.Sprintf(` sortType=%q`, direction)
		}
		if page, ok := pageFor[i]; ok && page.SelectedItem == nil && len(hiddenFor[i]) > 0 {
			attrs += ` multipleItemSelectionAllowed="1"`
		}
		members := membersFor[i]
		if len(members) > 0 {
			var items strings.Builder
			for index, member := range members {
				hidden := ""
				if hiddenFor[i][member.Value] {
					hidden = ` h="1"`
				}
				fmt.Fprintf(&items, `<item x="%d"%s/>`, index, hidden)
			}
			items.WriteString(`<item t="default"/>`)
			fmt.Fprintf(&pf, `<pivotField%s showAll="0"><items count="%d">%s</items></pivotField>`, attrs, len(members)+1, items.String())
		} else if attrs == "" {
			fmt.Fprintf(&pf, `<pivotField showAll="0"/>`)
		} else {
			// items list with default subtotal item keeps strict readers happy.
			fmt.Fprintf(&pf, `<pivotField%s showAll="0"><items count="1"><item t="default"/></items></pivotField>`, attrs)
		}
	}

	var rows, cols strings.Builder
	for _, i := range rowIdx {
		fmt.Fprintf(&rows, `<field x="%d"/>`, i)
	}
	for _, i := range colIdx {
		fmt.Fprintf(&cols, `<field x="%d"/>`, i)
	}
	rowsXML := ""
	if len(rowIdx) > 0 {
		rowsXML = fmt.Sprintf(`<rowFields count="%d">%s</rowFields>`, len(rowIdx), rows.String())
	}
	colsXML := ""
	if len(colIdx) > 0 {
		colsXML = fmt.Sprintf(`<colFields count="%d">%s</colFields>`, len(colIdx), cols.String())
	}
	var pages strings.Builder
	for _, page := range spec.PageFields {
		selected := ""
		if page.SelectedItem != nil {
			selected = fmt.Sprintf(` item="%d"`, memberIndex(membersFor[fieldIdx[page.Field]], *page.SelectedItem))
		}
		fmt.Fprintf(&pages, `<pageField fld="%d" hier="-1"%s/>`, fieldIdx[page.Field], selected)
	}
	pagesXML := ""
	if len(spec.PageFields) > 0 {
		pagesXML = fmt.Sprintf(`<pageFields count="%d">%s</pageFields>`, len(spec.PageFields), pages.String())
	}

	var dfs strings.Builder
	for _, df := range spec.DataFields {
		sub := aggToSubtotal[df.Agg]
		label := fmt.Sprintf("%s of %s", aggDisplay[df.Agg], df.Field)
		attr := ""
		if sub != "sum" {
			attr = fmt.Sprintf(` subtotal=%q`, sub)
		}
		fmt.Fprintf(&dfs, `<dataField name=%q fld="%d" baseField="0" baseItem="0"%s/>`, esc(label), fieldIdx[df.Field], attr)
	}

	// Location ref is a placeholder extent from the target cell — refreshOnLoad
	// re-lays the pivot out on open, so the exact size doesn't matter.
	loc := spec.TargetCellRef
	if loc == "" {
		loc = "H1"
	}

	grandTotals := ""
	if spec.GrandTotals != nil && !*spec.GrandTotals {
		grandTotals = ` rowGrandTotals="0" colGrandTotals="0"`
	}
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n" +
		fmt.Sprintf(`<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name=%q cacheId="%d" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" createdVersion="6" useAutoFormatting="1" itemPrintTitles="1" indent="0" outline="1" outlineData="1" multipleFieldFilters="0"%s>`+
			`<location ref=%q firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>`+
			`<pivotFields count="%d">%s</pivotFields>`+
			`%s%s%s`+
			`<dataFields count="%d">%s</dataFields>`+
			`<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>`+
			`</pivotTableDefinition>`,
			esc(name), cacheID, grandTotals, esc(loc), len(spec.Fields), pf.String(), rowsXML, colsXML, pagesXML, len(spec.DataFields), dfs.String())
}
