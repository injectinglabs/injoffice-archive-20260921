package xlsxpatch

import (
	"regexp"
	"strconv"
	"strings"
)

// A read-only, single-rule supplement. It never changes the native style or
// conditional-formatting preservation inventory.
type NativeConditionalFillPreviewV1 struct {
	SheetID   string                        `json:"sheet_id"`
	SheetPart string                        `json:"sheet_part"`
	Status    string                        `json:"status"`
	Warnings  []string                      `json:"warnings"`
	Rule      *NativeConditionalFillRuleV1  `json:"rule,omitempty"`
	Cells     []NativeConditionalFillCellV1 `json:"cells,omitempty"`
}
type NativeConditionalFillRuleV1 struct {
	Ref        string `json:"ref"`
	Operator   string `json:"operator"`
	Operand    string `json:"operand"`
	Priority   int    `json:"priority"`
	StopIfTrue bool   `json:"stop_if_true"`
	DxfID      int    `json:"dxf_id"`
	Fill       string `json:"fill"`
}
type NativeConditionalFillCellV1 struct {
	Row     int    `json:"row"`
	Column  int    `json:"column"`
	Lexical string `json:"lexical"`
	Cached  bool   `json:"cached"`
	Matches bool   `json:"matches"`
}

var nativeConditionalInteger = regexp.MustCompile(`^-?(0|[1-9][0-9]{0,14})$`)
var nativeConditionalRange = regexp.MustCompile(`^[A-Z]{1,3}[1-9][0-9]{0,6}(:[A-Z]{1,3}[1-9][0-9]{0,6})?$`)
var nativeConditionalRGB = regexp.MustCompile(`^FF[0-9A-Fa-f]{6}$`)

// Numeric comparison is restricted to exactly representable integers. Decimal,
// exponent, locale, coercion and general formula semantics are not inferred.
func nativeConditionalCompare(value, operand int64, operator string) (bool, bool) {
	switch operator {
	case "equal":
		return value == operand, true
	case "notEqual":
		return value != operand, true
	case "lessThan":
		return value < operand, true
	case "lessThanOrEqual":
		return value <= operand, true
	case "greaterThan":
		return value > operand, true
	case "greaterThanOrEqual":
		return value >= operand, true
	}
	return false, false
}

func nativeConditionalNode(n *previewXML, namespace, name string, allowed ...string) bool {
	if n == nil || n.name.Space != namespace || n.name.Local != name || strings.TrimSpace(n.text) != "" {
		return false
	}
	for _, a := range n.attrs {
		if a.Name.Space == "http://www.w3.org/2000/xmlns/" || a.Name.Space == "" && a.Name.Local == "xmlns" {
			continue
		}
		found := false
		for _, key := range allowed {
			if a.Name.Space == "" && a.Name.Local == key {
				found = true
			}
		}
		if !found {
			return false
		}
	}
	return true
}

func previewNativeConditionalFills(pkg *nativeWorkbookPackage, sheets []NativeWorkbookSheetV2) []NativeConditionalFillPreviewV1 {
	result := []NativeConditionalFillPreviewV1{}
	budget := 16384
	for _, sheet := range sheets {
		if len(result) == 64 {
			break
		}
		entry := previewNativeConditionalFill(pkg, sheet, budget)
		if entry != nil {
			result = append(result, *entry)
			budget -= len(entry.Cells)
		}
	}
	return result
}

func previewNativeConditionalFill(pkg *nativeWorkbookPackage, sheet NativeWorkbookSheetV2, budget int) *NativeConditionalFillPreviewV1 {
	root, err := parsePreviewXML(pkg.files[sheet.PartName])
	if err != nil {
		return &NativeConditionalFillPreviewV1{SheetID: sheet.ID, SheetPart: sheet.PartName, Status: "unavailable", Warnings: []string{"Conditional fill preview unavailable: worksheet XML could not be qualified within the preview bounds. No conditional fills were applied."}}
	}
	var formats []*previewXML
	unresolved := false
	for _, child := range root.children {
		if child.name.Local == "conditionalFormatting" {
			formats = append(formats, child)
		}
		if child.name.Local == "extLst" || child.name.Local == "AlternateContent" {
			unresolved = true
		}
	}
	formatCount, ruleCount := 0, 0
	var inspect func(*previewXML)
	inspect = func(n *previewXML) {
		if n.name.Local == "conditionalFormatting" {
			formatCount++
		}
		if n.name.Local == "cfRule" {
			ruleCount++
		}
		for _, child := range n.children {
			inspect(child)
		}
	}
	inspect(root)
	if formatCount == 0 && ruleCount == 0 && !unresolved {
		return nil
	}
	entry := &NativeConditionalFillPreviewV1{SheetID: sheet.ID, SheetPart: sheet.PartName, Status: "unavailable"}
	refuse := func(reason string) *NativeConditionalFillPreviewV1 {
		entry.Warnings = []string{"Conditional fill preview unavailable: " + reason + " No conditional fills were applied; source rules remain preserved."}
		return entry
	}
	if unresolved || len(formats) != 1 || formatCount != 1 || ruleCount != 1 {
		return refuse("requires exactly one rule without worksheet extensions or alternate content.")
	}
	ns := root.name.Space
	if root.name.Local != "worksheet" || ns != spreadsheetMLTransitional && ns != spreadsheetMLStrict {
		return refuse("unqualified worksheet namespace.")
	}
	for _, child := range root.children {
		if child.name.Local == "tableParts" {
			return refuse("table and conditional-format combinations are outside this subset.")
		}
	}
	cf := formats[0]
	if !nativeConditionalNode(cf, ns, "conditionalFormatting", "sqref") || len(cf.children) != 1 {
		return refuse("requires one plain conditionalFormatting range and one rule.")
	}
	top, left, bottom, right, err := parseDimensionReference(cf.attr("sqref"))
	if err != nil || !nativeConditionalRange.MatchString(cf.attr("sqref")) || (bottom-top+1)*(right-left+1) > 4096 || (bottom-top+1)*(right-left+1) > budget {
		return refuse("range is invalid or exceeds the 4096-cell sheet / 16384-cell workbook bound.")
	}
	for _, merge := range sheet.MergedRanges {
		if merge.Row <= bottom && merge.EndRow >= top && merge.Column <= right && merge.EndColumn >= left {
			return refuse("the conditional range intersects merged cells.")
		}
	}
	rule := cf.children[0]
	if !nativeConditionalNode(rule, ns, "cfRule", "type", "operator", "dxfId", "priority", "stopIfTrue") || len(rule.children) != 1 || rule.attr("type") != "cellIs" {
		return refuse("requires one cellIs rule with only supported attributes.")
	}
	formula := rule.children[0]
	if formula.name.Space != ns || formula.name.Local != "formula" || len(formula.children) != 0 || len(formula.attrs) != 0 || !nativeConditionalInteger.MatchString(formula.text) {
		return refuse("comparison operand must be a literal integer of at most 15 digits.")
	}
	operand, _ := strconv.ParseInt(formula.text, 10, 64)
	if _, ok := nativeConditionalCompare(0, operand, rule.attr("operator")); !ok {
		return refuse("comparison operator is outside the six single-operand comparisons.")
	}
	priority, err := strconv.Atoi(rule.attr("priority"))
	if err != nil || priority < 1 || priority > 2147483647 {
		return refuse("priority is missing or invalid.")
	}
	stop := rule.attr("stopIfTrue")
	if stop != "" && stop != "0" && stop != "1" && stop != "true" && stop != "false" {
		return refuse("stopIfTrue is invalid.")
	}
	dxfID, err := strconv.Atoi(rule.attr("dxfId"))
	if err != nil || dxfID < 0 || dxfID > 4095 {
		return refuse("differential style index is missing or outside the bound.")
	}
	fill := nativeConditionalSolidFill(pkg, ns, dxfID)
	if fill == "" {
		return refuse("differential style must contain only one explicit opaque RGB solid foreground fill.")
	}
	if !nativeConditionalSourceCells(root, top, left, bottom, right) {
		return refuse("source cells have ambiguous identities, metadata or unsupported markup.")
	}
	cells := map[[2]int]NativeWorkbookCellV2{}
	identities := map[[2]int]bool{}
	for _, cell := range sheet.Cells {
		key := [2]int{cell.Row, cell.Column}
		if identities[key] || cell.Ref != cellReference(cell.Row, cell.Column) {
			return refuse("source cell identities are ambiguous.")
		}
		identities[key] = true
		if cell.Formula != nil && cell.Formula.Type != "normal" {
			return refuse("formula groups anywhere on the worksheet are outside this subset.")
		}
		if cell.Row >= top && cell.Row <= bottom && cell.Column >= left && cell.Column <= right {
			cells[[2]int{cell.Row, cell.Column}] = cell
		}
	}
	evaluated := []NativeConditionalFillCellV1{}
	for row := top; row <= bottom; row++ {
		for col := left; col <= right; col++ {
			cell, ok := cells[[2]int{row, col}]
			value := cell.Value
			if cell.Formula != nil {
				if cell.Formula.Type != "normal" {
					return refuse("formula groups are outside this subset.")
				}
				value = cell.Formula.Cached
			}
			if !ok || value == nil || value.Kind != "number" || value.Storage != "number" || value.Lexical == nil || !nativeConditionalInteger.MatchString(*value.Lexical) {
				return refuse("every target cell must have a saved integer numeric value; blanks, text, errors and unavailable caches are not coerced.")
			}
			number, _ := strconv.ParseInt(*value.Lexical, 10, 64)
			matches, _ := nativeConditionalCompare(number, operand, rule.attr("operator"))
			evaluated = append(evaluated, NativeConditionalFillCellV1{row, col, *value.Lexical, cell.Formula != nil, matches})
		}
	}
	entry.Status = "available"
	entry.Rule = &NativeConditionalFillRuleV1{cf.attr("sqref"), rule.attr("operator"), formula.text, priority, stop == "1" || stop == "true", dxfID, fill}
	entry.Cells = evaluated
	entry.Warnings = []string{"Read-only conditional fill preview: one integer cellIs rule and an explicit RGB solid fill. Saved formula caches are used as written; freshness is unknown and formulas are never recalculated. Multiple rules, other styles and Excel print fidelity remain unqualified."}
	return entry
}

func nativeConditionalSolidFill(pkg *nativeWorkbookPackage, ns string, id int) string {
	location, err := locateWorkbookPartBytes(pkg.index, func(name string) ([]byte, bool) { v, ok := pkg.files[name]; return v, ok })
	if err != nil {
		return ""
	}
	expected, opposing, relNS := relTypeStylesTransitional, relTypeStylesStrict, officeRelNamespaceTransitional
	if location.strict {
		expected, opposing, relNS = relTypeStylesStrict, relTypeStylesTransitional, officeRelNamespaceStrict
	}
	extractor := &nativeWorkbookExtractor{pkg: pkg, workbook: location, namespace: ns, relNamespace: relNS, modeled: map[string]bool{}, unsupportedKeys: map[string]bool{}, claimedXML: map[string]bool{}}
	part, err := extractor.relatedCorePart(expected, opposing, stylesPartContentType, "styles", false)
	if err != nil || part == "" {
		return ""
	}
	styles, err := parsePreviewXML(pkg.files[part])
	if err != nil || styles.name.Space != ns || styles.name.Local != "styleSheet" {
		return ""
	}
	for _, child := range styles.children {
		if child.name.Local == "extLst" || child.name.Local == "AlternateContent" {
			return ""
		}
	}
	dxfContainers := 0
	var inspect func(*previewXML)
	inspect = func(n *previewXML) {
		if n.name.Local == "dxfs" {
			dxfContainers++
		}
		for _, child := range n.children {
			inspect(child)
		}
	}
	inspect(styles)
	if dxfContainers != 1 {
		return ""
	}
	dxfs := styles.child("dxfs")
	if !nativeConditionalNode(dxfs, ns, "dxfs", "count") || len(dxfs.children) > 4096 || id >= len(dxfs.children) {
		return ""
	}
	if count := dxfs.attr("count"); count != "" {
		n, err := strconv.Atoi(count)
		if err != nil || n != len(dxfs.children) {
			return ""
		}
	}
	for _, dxf := range dxfs.children {
		if dxf.name.Space != ns || dxf.name.Local != "dxf" {
			return ""
		}
	}
	dxf := dxfs.children[id]
	if !nativeConditionalNode(dxf, ns, "dxf") || len(dxf.children) != 1 {
		return ""
	}
	fill := dxf.children[0]
	if !nativeConditionalNode(fill, ns, "fill") || len(fill.children) != 1 {
		return ""
	}
	pattern := fill.children[0]
	if !nativeConditionalNode(pattern, ns, "patternFill", "patternType") || pattern.attr("patternType") != "solid" || len(pattern.children) != 1 {
		return ""
	}
	fg := pattern.children[0]
	if !nativeConditionalNode(fg, ns, "fgColor", "rgb") || len(fg.children) != 0 || !nativeConditionalRGB.MatchString(fg.attr("rgb")) {
		return ""
	}
	return "#" + strings.ToUpper(fg.attr("rgb")[2:])
}

// Inspect the original cell payload as well as its numeric projection. Metadata
// and opaque children cannot gain display authority through a surviving <v>.
func nativeConditionalSourceCells(root *previewXML, top, left, bottom, right int) bool {
	ns := root.name.Space
	data := root.child("sheetData")
	if !nativeConditionalNode(data, ns, "sheetData") {
		return false
	}
	seen := map[string]bool{}
	for _, row := range data.children {
		if row.name.Space != ns || row.name.Local != "row" {
			return false
		}
		for _, cell := range row.children {
			if cell.name.Local != "c" {
				return false
			}
			r, c, err := parseCellReference(cell.attr("r"))
			if err != nil || seen[cell.attr("r")] || cell.attr("r") != cellReference(r, c) {
				return false
			}
			seen[cell.attr("r")] = true
			if r < top || r > bottom || c < left || c > right {
				continue
			}
			if !nativeConditionalNode(cell, ns, "c", "r", "s", "t") || (cell.attr("t") != "" && cell.attr("t") != "n") {
				return false
			}
			values, formulas := 0, 0
			for _, child := range cell.children {
				if child.name.Space != ns || len(child.children) != 0 {
					return false
				}
				switch child.name.Local {
				case "v":
					if len(child.attrs) != 0 {
						return false
					}
					values++
				case "f":
					for _, a := range child.attrs {
						if a.Name.Space != "" || a.Name.Local != "t" || a.Value != "normal" {
							return false
						}
					}
					formulas++
				default:
					return false
				}
			}
			if values != 1 || formulas > 1 {
				return false
			}
		}
	}
	return true
}
