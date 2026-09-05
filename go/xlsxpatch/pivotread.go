package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"sort"
	"strconv"
	"strings"
)

// PivotInfo is the editable subset of one native OOXML pivot table. It is
// intentionally renderer-neutral and maps directly onto @injoffice/pivots.
// Warnings describe features that were present but could not be represented;
// callers must surface or reject them rather than silently dropping state.
type PivotInfo struct {
	Part            string              `json:"part"`
	Identity        PivotIdentity       `json:"identity"`
	Name            string              `json:"name,omitempty"`
	CacheID         int                 `json:"cacheId"`
	SourceSheetName string              `json:"sourceSheetName,omitempty"`
	SourceRef       string              `json:"sourceRef,omitempty"`
	Fields          []string            `json:"fields"`
	TargetSheetName string              `json:"targetSheetName,omitempty"`
	TargetRef       string              `json:"targetRef,omitempty"`
	RowFields       []string            `json:"rowFields"`
	ColFields       []string            `json:"colFields"`
	DataFields      []PivotDataField    `json:"dataFields"`
	PageFields      []PivotPageField    `json:"pageFields,omitempty"`
	MemberFilters   []PivotMemberFilter `json:"memberFilters,omitempty"`
	Sorts           []PivotFieldSort    `json:"sorts,omitempty"`
	FieldMembers    []PivotFieldMembers `json:"fieldMembers,omitempty"`
	GrandTotals     bool                `json:"grandTotals"`
	Warnings        []string            `json:"warnings,omitempty"`
}

// PivotIdentity is the stable native identity of a pivot table. The part name
// is package-local, immutable for the lifetime of the hydrated object, and
// does not depend on a user-editable display name or a recyclable rId.
type PivotIdentity struct {
	Part string `json:"part"`
}

type packageRelationship struct {
	ID         string
	Type       string
	Target     string
	TargetMode string
}

type workbookSheet struct {
	Name string
	Part string
}

// ReadPivots hydrates native pivot-table/cache definitions. Cached record
// values are deliberately not loaded: the source range and field model are
// the editable contract, and native applications can refresh cached values.
func ReadPivots(data []byte) ([]PivotInfo, error) {
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: read pivots: %w", err)
	}
	files := make(map[string]*zip.File, len(zr.File))
	for _, file := range zr.File {
		files[file.Name] = file
	}
	read := func(name string) (string, bool) {
		file, ok := files[name]
		if !ok {
			return "", false
		}
		rc, err := file.Open()
		if err != nil {
			return "", false
		}
		defer rc.Close()
		value, err := io.ReadAll(rc)
		return string(value), err == nil
	}

	var tableParts []string
	for name := range files {
		if strings.HasPrefix(name, "xl/pivotTables/") && strings.HasSuffix(name, ".xml") && !strings.Contains(name, "/_rels/") {
			tableParts = append(tableParts, name)
		}
	}
	if len(tableParts) == 0 {
		return nil, nil
	}
	sort.Strings(tableParts)

	sheets, err := readWorkbookSheets(read)
	if err != nil {
		return nil, err
	}
	targetSheetByTable := map[string]string{}
	for _, sheet := range sheets {
		relsXML, ok := read(relsPartFor(sheet.Part))
		if !ok {
			continue
		}
		rels, err := parsePackageRelationships(relsXML)
		if err != nil {
			continue
		}
		base := sheet.Part[:strings.LastIndex(sheet.Part, "/")]
		for _, rel := range rels {
			if rel.Type == relTypePivotTable || strings.HasSuffix(rel.Type, "/pivotTable") {
				external, modeErr := packageRelationshipIsExternal(rel)
				if modeErr != nil {
					return nil, fmt.Errorf("xlsxpatch: pivot table relationship %q: %w", rel.ID, modeErr)
				}
				if external {
					continue
				}
				target, resolveErr := resolveRelPath(base, rel.Target)
				if resolveErr != nil {
					return nil, fmt.Errorf("xlsxpatch: pivot table relationship %q: %w", rel.ID, resolveErr)
				}
				targetSheetByTable[target] = sheet.Name
			}
		}
	}

	out := make([]PivotInfo, 0, len(tableParts))
	for _, part := range tableParts {
		tableXML, ok := read(part)
		if !ok {
			continue
		}
		parsed, err := parsePivotTableDefinition(tableXML)
		if err != nil {
			out = append(out, PivotInfo{Part: part, Identity: PivotIdentity{Part: part}, TargetSheetName: targetSheetByTable[part], GrandTotals: true, Warnings: []string{err.Error()}})
			continue
		}
		parsed.Part = part
		parsed.Identity = PivotIdentity{Part: part}
		parsed.TargetSheetName = targetSheetByTable[part]
		if parsed.TargetSheetName == "" {
			parsed.Warnings = append(parsed.Warnings, "pivot table is not linked from a worksheet")
		}

		cachePart := ""
		if relsXML, ok := read(relsPartFor(part)); ok {
			rels, relErr := parsePackageRelationships(relsXML)
			if relErr == nil {
				base := part[:strings.LastIndex(part, "/")]
				for _, rel := range rels {
					if rel.Type == relTypePivotCacheDef || strings.HasSuffix(rel.Type, "/pivotCacheDefinition") {
						external, modeErr := packageRelationshipIsExternal(rel)
						if modeErr != nil {
							return nil, fmt.Errorf("xlsxpatch: pivot cache relationship %q: %w", rel.ID, modeErr)
						}
						if external {
							continue
						}
						cachePart, relErr = resolveRelPath(base, rel.Target)
						if relErr != nil {
							return nil, fmt.Errorf("xlsxpatch: pivot cache relationship %q: %w", rel.ID, relErr)
						}
						break
					}
				}
			}
		}
		if cachePart == "" {
			parsed.Warnings = append(parsed.Warnings, "pivot cache relationship is missing")
		} else if cacheXML, ok := read(cachePart); !ok {
			parsed.Warnings = append(parsed.Warnings, "pivot cache definition is missing")
		} else {
			cache, cacheWarnings, cacheErr := parsePivotCacheDefinition(cacheXML)
			if cacheErr != nil {
				parsed.Warnings = append(parsed.Warnings, cacheErr.Error())
			} else {
				parsed.SourceSheetName = cache.SourceSheetName
				parsed.SourceRef = cache.SourceRef
				parsed.Fields = cache.Fields
				parsed.fieldMembers = cache.FieldMembers
				for index, field := range cache.Fields {
					if index < len(cache.FieldMembers) && len(cache.FieldMembers[index]) > 0 {
						parsed.FieldMembers = append(parsed.FieldMembers, PivotFieldMembers{Field: field, Items: cache.FieldMembers[index]})
					}
				}
				parsed.Warnings = append(parsed.Warnings, cacheWarnings...)
			}
		}
		parsed.resolveFieldIndexes()
		out = append(out, parsed.PivotInfo)
	}
	return out, nil
}

func readWorkbookSheets(read func(string) (string, bool)) ([]workbookSheet, error) {
	workbookXML, ok := read("xl/workbook.xml")
	if !ok {
		return nil, fmt.Errorf("xlsxpatch: missing xl/workbook.xml")
	}
	relsXML, ok := read("xl/_rels/workbook.xml.rels")
	if !ok {
		return nil, fmt.Errorf("xlsxpatch: missing workbook rels")
	}
	rels, err := parsePackageRelationships(relsXML)
	if err != nil {
		return nil, err
	}
	byID := map[string]packageRelationship{}
	for _, rel := range rels {
		byID[rel.ID] = rel
	}
	dec := xml.NewDecoder(strings.NewReader(workbookXML))
	var out []workbookSheet
	for {
		token, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: parse workbook: %w", err)
		}
		start, ok := token.(xml.StartElement)
		if !ok || start.Name.Local != "sheet" {
			continue
		}
		name, relID := attrVal(start, "name"), attrVal(start, "id")
		if rel, ok := byID[relID]; ok && name != "" {
			external, modeErr := packageRelationshipIsExternal(rel)
			if modeErr != nil {
				return nil, fmt.Errorf("xlsxpatch: worksheet relationship %q: %w", rel.ID, modeErr)
			}
			if external {
				return nil, fmt.Errorf("xlsxpatch: worksheet relationship %q has an external target", rel.ID)
			}
			part, resolveErr := resolveRelPath("xl", rel.Target)
			if resolveErr != nil {
				return nil, fmt.Errorf("xlsxpatch: worksheet relationship %q: %w", rel.ID, resolveErr)
			}
			out = append(out, workbookSheet{Name: name, Part: part})
		}
	}
	return out, nil
}

func parsePackageRelationships(value string) ([]packageRelationship, error) {
	dec := xml.NewDecoder(strings.NewReader(value))
	var out []packageRelationship
	for {
		token, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: parse relationships: %w", err)
		}
		start, ok := token.(xml.StartElement)
		if !ok || start.Name.Local != "Relationship" {
			continue
		}
		out = append(out, packageRelationship{
			ID:         attrVal(start, "Id"),
			Type:       attrVal(start, "Type"),
			Target:     attrVal(start, "Target"),
			TargetMode: attrVal(start, "TargetMode"),
		})
	}
	return out, nil
}

func packageRelationshipIsExternal(rel packageRelationship) (bool, error) {
	mode := strings.TrimSpace(rel.TargetMode)
	if mode == "" || strings.EqualFold(mode, "Internal") {
		return false, nil
	}
	if strings.EqualFold(mode, "External") {
		return true, nil
	}
	return false, fmt.Errorf("unsupported target mode %q", rel.TargetMode)
}

type rawPivotTable struct {
	PivotInfo
	rowIndexes   []int
	colIndexes   []int
	dataIndexes  []rawDataField
	pivotFields  []rawPivotField
	pageIndexes  []rawPageField
	fieldMembers [][]PivotFieldMember
}

type rawDataField struct {
	Index int
	Agg   string
}

type rawPivotField struct {
	Sort      string
	MultiPage bool
	Items     []rawPivotItem
}

type rawPivotItem struct {
	CacheIndex int
	Hidden     bool
}

type rawPageField struct {
	Index int
	Item  *int
}

func parsePivotTableDefinition(value string) (rawPivotTable, error) {
	dec := xml.NewDecoder(strings.NewReader(value))
	result := rawPivotTable{PivotInfo: PivotInfo{GrandTotals: true}}
	var stack []string
	inside := func(name string) bool {
		for _, item := range stack {
			if item == name {
				return true
			}
		}
		return false
	}
	seenRoot := false
	currentPivotField := -1
	warnedPivotFilters := false
	warnedAutoSort := false
	for {
		token, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return rawPivotTable{}, fmt.Errorf("malformed pivot table definition: %w", err)
		}
		switch item := token.(type) {
		case xml.StartElement:
			stack = append(stack, item.Name.Local)
			switch item.Name.Local {
			case "pivotTableDefinition":
				seenRoot = true
				result.Name = attrVal(item, "name")
				result.CacheID, _ = strconv.Atoi(attrVal(item, "cacheId"))
				if attrVal(item, "rowGrandTotals") == "0" {
					result.GrandTotals = false
				}
				rowGrand, colGrand := attrVal(item, "rowGrandTotals"), attrVal(item, "colGrandTotals")
				if rowGrand == "" {
					rowGrand = "1"
				}
				if colGrand == "" {
					colGrand = "1"
				}
				if rowGrand != colGrand {
					result.Warnings = append(result.Warnings, "different row and column grand-total settings are not representable")
				}
			case "location":
				result.TargetRef = attrVal(item, "ref")
			case "pivotField":
				if inside("pivotFields") {
					currentPivotField++
					result.pivotFields = append(result.pivotFields, rawPivotField{Sort: attrVal(item, "sortType"), MultiPage: attrVal(item, "multipleItemSelectionAllowed") == "1"})
				}
			case "item":
				if currentPivotField >= 0 && inside("pivotFields") && inside("pivotField") {
					index := -1
					if raw := attrVal(item, "x"); raw != "" {
						if parsed, parseErr := strconv.Atoi(raw); parseErr == nil {
							index = parsed
						}
					}
					result.pivotFields[currentPivotField].Items = append(result.pivotFields[currentPivotField].Items, rawPivotItem{CacheIndex: index, Hidden: attrVal(item, "h") == "1"})
				}
			case "field":
				index, parseErr := strconv.Atoi(attrVal(item, "x"))
				if parseErr == nil {
					if inside("rowFields") {
						result.rowIndexes = append(result.rowIndexes, index)
					} else if inside("colFields") {
						result.colIndexes = append(result.colIndexes, index)
					}
				}
			case "dataField":
				index, parseErr := strconv.Atoi(attrVal(item, "fld"))
				if parseErr == nil {
					result.dataIndexes = append(result.dataIndexes, rawDataField{Index: index, Agg: pivotAggFromOOXML(attrVal(item, "subtotal"))})
				}
			case "pageField":
				index, parseErr := strconv.Atoi(attrVal(item, "fld"))
				if parseErr == nil {
					page := rawPageField{Index: index}
					if raw := attrVal(item, "item"); raw != "" {
						if parsed, itemErr := strconv.Atoi(raw); itemErr == nil {
							page.Item = &parsed
						} else {
							result.Warnings = append(result.Warnings, "page field has an invalid selected item index")
						}
					}
					result.pageIndexes = append(result.pageIndexes, page)
				}
			case "pivotFilters":
				if !warnedPivotFilters {
					result.Warnings = append(result.Warnings, "value and label pivot filters are not supported")
					warnedPivotFilters = true
				}
			case "autoSortScope":
				if !warnedAutoSort {
					result.Warnings = append(result.Warnings, "value-based pivot autosort scopes are not supported")
					warnedAutoSort = true
				}
			}
		case xml.EndElement:
			if len(stack) > 0 {
				stack = stack[:len(stack)-1]
			}
		}
	}
	if !seenRoot {
		return rawPivotTable{}, fmt.Errorf("malformed pivot table definition: root element is missing")
	}
	return result, nil
}

type rawPivotCache struct {
	SourceSheetName string
	SourceRef       string
	Fields          []string
	FieldMembers    [][]PivotFieldMember
}

func parsePivotCacheDefinition(value string) (rawPivotCache, []string, error) {
	dec := xml.NewDecoder(strings.NewReader(value))
	var result rawPivotCache
	var warnings []string
	seenRoot := false
	currentField := -1
	sharedDepth := 0
	for {
		token, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return rawPivotCache{}, warnings, fmt.Errorf("malformed pivot cache definition: %w", err)
		}
		switch item := token.(type) {
		case xml.StartElement:
			switch item.Name.Local {
			case "pivotCacheDefinition":
				seenRoot = true
			case "worksheetSource":
				result.SourceSheetName = attrVal(item, "sheet")
				result.SourceRef = attrVal(item, "ref")
				if name := attrVal(item, "name"); result.SourceRef == "" && name != "" {
					warnings = append(warnings, "named-range pivot sources are not hydrated")
				}
			case "cacheField":
				currentField++
				result.Fields = append(result.Fields, attrVal(item, "name"))
				result.FieldMembers = append(result.FieldMembers, nil)
			case "sharedItems":
				sharedDepth++
			case "s", "n", "b", "m", "d", "e":
				if sharedDepth == 0 || currentField < 0 {
					continue
				}
				member := PivotFieldMember{Value: attrVal(item, "v")}
				switch item.Name.Local {
				case "s":
					member.Kind = "string"
				case "n":
					member.Kind = "number"
				case "b":
					member.Kind = "boolean"
					if member.Value == "1" {
						member.Value = "true"
					} else if member.Value == "0" {
						member.Value = "false"
					} else {
						warnings = append(warnings, fmt.Sprintf("field %q has an invalid boolean member", result.Fields[currentField]))
						continue
					}
				case "m":
					member.Kind, member.Value = "blank", ""
				default:
					warnings = append(warnings, fmt.Sprintf("field %q has unsupported %s cache members", result.Fields[currentField], item.Name.Local))
					continue
				}
				result.FieldMembers[currentField] = append(result.FieldMembers[currentField], member)
			}
		case xml.EndElement:
			if item.Name.Local == "sharedItems" && sharedDepth > 0 {
				sharedDepth--
			}
		}
	}
	if !seenRoot {
		return rawPivotCache{}, warnings, fmt.Errorf("malformed pivot cache definition: root element is missing")
	}
	if result.SourceSheetName == "" || result.SourceRef == "" {
		warnings = append(warnings, "pivot source is not a direct worksheet range")
	}
	for index, members := range result.FieldMembers {
		seen := map[string]bool{}
		for _, member := range members {
			if seen[member.Value] {
				warnings = append(warnings, fmt.Sprintf("field %q has ambiguous members with display value %q", result.Fields[index], member.Value))
			}
			seen[member.Value] = true
		}
	}
	return result, warnings, nil
}

func pivotAggFromOOXML(value string) string {
	switch value {
	case "", "sum":
		return "sum"
	case "average":
		return "avg"
	case "count", "min", "max":
		return value
	default:
		return value
	}
}

func (result *rawPivotTable) resolveFieldIndexes() {
	fieldAt := func(index int, kind string) (string, bool) {
		if index < 0 || index >= len(result.Fields) {
			result.Warnings = append(result.Warnings, fmt.Sprintf("%s field index %d is not representable", kind, index))
			return "", false
		}
		return result.Fields[index], true
	}
	for _, index := range result.rowIndexes {
		if field, ok := fieldAt(index, "row"); ok {
			result.RowFields = append(result.RowFields, field)
		}
	}
	for _, index := range result.colIndexes {
		if field, ok := fieldAt(index, "column"); ok {
			result.ColFields = append(result.ColFields, field)
		}
	}
	for _, data := range result.dataIndexes {
		if field, ok := fieldAt(data.Index, "data"); ok {
			result.DataFields = append(result.DataFields, PivotDataField{Field: field, Agg: data.Agg})
			if _, supported := aggToSubtotal[data.Agg]; !supported {
				result.Warnings = append(result.Warnings, fmt.Sprintf("data aggregation %q is not supported", data.Agg))
			}
		}
	}
	memberAt := func(fieldIndex, cacheIndex int, kind string) (string, bool) {
		if fieldIndex < 0 || fieldIndex >= len(result.fieldMembers) || cacheIndex < 0 || cacheIndex >= len(result.fieldMembers[fieldIndex]) {
			result.Warnings = append(result.Warnings, fmt.Sprintf("%s member index %d for field %d is not representable", kind, cacheIndex, fieldIndex))
			return "", false
		}
		return result.fieldMembers[fieldIndex][cacheIndex].Value, true
	}
	for index, pivotField := range result.pivotFields {
		if pivotField.Sort != "" && pivotField.Sort != "manual" {
			if pivotField.Sort == "ascending" || pivotField.Sort == "descending" {
				if field, ok := fieldAt(index, "sort"); ok {
					axis := false
					for _, candidate := range append(append([]string{}, result.RowFields...), result.ColFields...) {
						if candidate == field {
							axis = true
						}
					}
					if axis {
						result.Sorts = append(result.Sorts, PivotFieldSort{Field: field, Direction: pivotField.Sort})
					} else {
						result.Warnings = append(result.Warnings, fmt.Sprintf("sort field %q is not a row or column field", field))
					}
				}
			} else {
				result.Warnings = append(result.Warnings, fmt.Sprintf("pivot sort type %q is not supported", pivotField.Sort))
			}
		}
		var excluded []string
		for _, item := range pivotField.Items {
			if !item.Hidden {
				continue
			}
			if value, ok := memberAt(index, item.CacheIndex, "hidden"); ok {
				excluded = append(excluded, value)
			}
		}
		if len(excluded) > 0 {
			if field, ok := fieldAt(index, "filter"); ok {
				result.MemberFilters = append(result.MemberFilters, PivotMemberFilter{Field: field, ExcludedItems: excluded})
			}
		}
	}
	for _, page := range result.pageIndexes {
		field, ok := fieldAt(page.Index, "page")
		if !ok {
			continue
		}
		hydrated := PivotPageField{Field: field}
		if page.Item != nil {
			if page.Index >= 0 && page.Index < len(result.pivotFields) && result.pivotFields[page.Index].MultiPage {
				result.Warnings = append(result.Warnings, fmt.Sprintf("page field %q combines a single selected item with multiple-item mode", field))
				continue
			}
			if page.Index < 0 || page.Index >= len(result.pivotFields) || *page.Item < 0 || *page.Item >= len(result.pivotFields[page.Index].Items) {
				result.Warnings = append(result.Warnings, fmt.Sprintf("page field %q selected item index is not representable", field))
				continue
			}
			cacheIndex := result.pivotFields[page.Index].Items[*page.Item].CacheIndex
			if value, memberOK := memberAt(page.Index, cacheIndex, "page"); memberOK {
				hydrated.SelectedItem = &value
			} else {
				continue
			}
		}
		result.PageFields = append(result.PageFields, hydrated)
	}
}
