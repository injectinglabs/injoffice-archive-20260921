package xlsxpatch

import (
	"fmt"
	"math"
	"sort"
	"strconv"
	"strings"
)

// NativeConditionalBarFillPreviewV1 is a read-only data-bar overlay. Like the
// colour-scale tier it never changes the native style inventory and never
// removes the sheet's CONDITIONAL_FORMATTING preservation entry.
type NativeConditionalBarFillPreviewV1 struct {
	SheetID   string                            `json:"sheet_id"`
	SheetPart string                            `json:"sheet_part"`
	Status    string                            `json:"status"`
	Warnings  []string                          `json:"warnings"`
	Cells     []NativeConditionalBarFillCellV1  `json:"cells,omitempty"`
	Ranges    []NativeConditionalBarFillRangeV1 `json:"ranges,omitempty"`
}

// NativeConditionalBarFillRangeV1 discloses which source range each painted
// bar came from.
type NativeConditionalBarFillRangeV1 struct {
	Ref      string `json:"ref"`
	Priority int    `json:"priority"`
}

// NativeConditionalBarFillCellV1 is one painted bar. Coordinates are
// zero-based. The span and the axis are thousandths of the cell's width, so
// the renderer needs no source units: Start is always less than End.
type NativeConditionalBarFillCellV1 struct {
	Row         int    `json:"row"`
	Column      int    `json:"column"`
	Start       int    `json:"start_permille"`
	End         int    `json:"end_permille"`
	Axis        int    `json:"axis_permille"`
	Color       string `json:"color"`
	BorderColor string `json:"border_color,omitempty"`
	AxisColor   string `json:"axis_color,omitempty"`
}

const nativeBarDataBarExtURI = "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"

func previewNativeConditionalBarFills(pkg *nativeWorkbookPackage, sheets []NativeWorkbookSheetV2) []NativeConditionalBarFillPreviewV1 {
	theme, themeOK := nativeScaleWorkbookTheme(pkg)
	result := []NativeConditionalBarFillPreviewV1{}
	budget := 16384
	for _, sheet := range sheets {
		if len(result) == 64 || budget <= 0 {
			break
		}
		entry := previewNativeConditionalBarFill(pkg, sheet, theme, themeOK, budget)
		if entry != nil {
			result = append(result, *entry)
			budget -= len(entry.Cells)
		}
	}
	return result
}

type nativeBarRule struct {
	rect     nativeScaleRange
	ref      string
	priority int
	legacy   *previewXML
	extID    string
}

func previewNativeConditionalBarFill(pkg *nativeWorkbookPackage, sheet NativeWorkbookSheetV2, theme nativeWorkbookThemeDisplay, themeOK bool, budget int) *NativeConditionalBarFillPreviewV1 {
	root, err := parsePreviewXML(pkg.files[sheet.PartName])
	if err != nil {
		return nil
	}
	ns := root.name.Space
	if root.name.Local != "worksheet" || ns != spreadsheetMLTransitional && ns != spreadsheetMLStrict {
		return nil
	}
	var bars []nativeBarRule
	var occupied []nativeScaleRange
	for _, format := range root.children {
		if format.name.Local != "conditionalFormatting" || format.name.Space != ns {
			continue
		}
		rect, ok := nativeScaleParseRange(format.attr("sqref"))
		if !ok {
			return nativeBarRefusal(sheet, "a conditional range is outside the single-rectangle A1 subset.")
		}
		occupied = append(occupied, rect)
		for _, rule := range format.children {
			if rule.name.Local != "cfRule" || rule.name.Space != ns || rule.attr("type") != "dataBar" {
				continue
			}
			legacy := rule.child("dataBar")
			if legacy == nil {
				continue
			}
			priority, err := strconv.Atoi(rule.attr("priority"))
			if err != nil || priority < 1 || priority > 2147483647 {
				continue
			}
			bars = append(bars, nativeBarRule{rect: rect, ref: format.attr("sqref"), priority: priority, legacy: legacy, extID: nativeBarRuleExtensionID(rule)})
		}
	}
	if len(bars) == 0 {
		return nil
	}
	extensions := nativeBarExtensionRules(root, ns)
	// An x14 range this tier did not match to one of its own rules still
	// claims its cells, exactly as it does for a colour scale.
	occupied = append(occupied, nativeScaleExtensionRanges(root)...)
	cells := map[[2]int]float64{}
	for _, cell := range sheet.Cells {
		if cell.Ref != cellReference(cell.Row, cell.Column) {
			return nativeBarRefusal(sheet, "source cell identities are ambiguous.")
		}
		if number, ok := nativeScaleCellNumber(cell); ok {
			cells[[2]int{cell.Row, cell.Column}] = number
		}
	}
	entry := &NativeConditionalBarFillPreviewV1{SheetID: sheet.ID, SheetPart: sheet.PartName, Status: "unavailable"}
	unpainted := 0
	sort.SliceStable(bars, func(i, j int) bool { return bars[i].priority < bars[j].priority })
	for _, bar := range bars {
		// One rule owns its own range twice here -- once as the legacy rule and
		// once as its x14 twin -- so a range that meets anything else is left
		// to priority and stopIfTrue, which are not evaluated.
		claims := 0
		for _, claimed := range occupied {
			if bar.rect.intersects(claimed) {
				claims++
			}
		}
		for _, merge := range sheet.MergedRanges {
			if bar.rect.intersects(nativeScaleRange{merge.Row, merge.Column, merge.EndRow, merge.EndColumn}) {
				claims = 0
			}
		}
		size := (bar.rect.bottom - bar.rect.top + 1) * (bar.rect.right - bar.rect.left + 1)
		if claims != 2 || size > nativeScaleMaximumRangeCells || len(entry.Cells)+size > budget {
			unpainted++
			continue
		}
		painted, ok := nativeBarPaintRange(bar, extensions, cells, ns, theme, themeOK)
		if !ok {
			unpainted++
			continue
		}
		entry.Cells = append(entry.Cells, painted...)
		entry.Ranges = append(entry.Ranges, NativeConditionalBarFillRangeV1{Ref: bar.ref, Priority: bar.priority})
	}
	if len(entry.Cells) == 0 {
		return nativeBarRefusal(sheet, "no data-bar rule resolved to an exact source-qualified bar.")
	}
	sort.SliceStable(entry.Cells, func(i, j int) bool {
		if entry.Cells[i].Row != entry.Cells[j].Row {
			return entry.Cells[i].Row < entry.Cells[j].Row
		}
		return entry.Cells[i].Column < entry.Cells[j].Column
	})
	entry.Status = "available"
	warning := fmt.Sprintf("Read-only data-bar preview: %d bars across %d source ranges. Solid bars with explicit zero and full-length bounds are measured from saved cell values; formulas are never recalculated. The bar spans the cell, so Excel's own inset is not reproduced.", len(entry.Cells), len(entry.Ranges))
	if unpainted > 0 {
		warning += fmt.Sprintf(" %d data-bar rule(s) were not painted because a bound, a length, a gradient, an axis position, a colour or an overlapping rule is outside this subset; those source rules remain preserved.", unpainted)
	}
	entry.Warnings = []string{warning}
	return entry
}

func nativeBarRefusal(sheet NativeWorkbookSheetV2, reason string) *NativeConditionalBarFillPreviewV1 {
	return &NativeConditionalBarFillPreviewV1{
		SheetID: sheet.ID, SheetPart: sheet.PartName, Status: "unavailable",
		Warnings: []string{"Data-bar preview unavailable: " + reason + " No data bars were applied; source rules remain preserved."},
	}
}

// nativeBarRuleExtensionID reads the x14 identifier a downlevel dataBar rule
// carries, which is how the rule joins its real definition.
func nativeBarRuleExtensionID(rule *previewXML) string {
	extLst := rule.child("extLst")
	if extLst == nil {
		return ""
	}
	for _, ext := range extLst.children {
		if ext.name.Local != "ext" || ext.attr("uri") != nativeBarDataBarExtURI {
			continue
		}
		for _, child := range ext.children {
			if child.name.Local == "id" {
				return strings.TrimSpace(child.text)
			}
		}
	}
	return ""
}

// nativeBarExtensionRules indexes the worksheet's x14 dataBar definitions by
// identifier.
func nativeBarExtensionRules(root *previewXML, ns string) map[string]*previewXML {
	rules := map[string]*previewXML{}
	var walk func(*previewXML)
	walk = func(n *previewXML) {
		if n.name.Local == "cfRule" && n.name.Space != ns && n.attr("type") == "dataBar" {
			id := strings.TrimSpace(n.attr("id"))
			for _, child := range n.children {
				if child.name.Local == "dataBar" && id != "" {
					rules[id] = child
				}
			}
			return
		}
		for _, child := range n.children {
			walk(child)
		}
	}
	for _, child := range root.children {
		if child.name.Local == "extLst" || child.name.Local == "AlternateContent" {
			walk(child)
		}
	}
	return rules
}

func nativeBarPaintRange(bar nativeBarRule, extensions map[string]*previewXML, cells map[[2]int]float64, ns string, theme nativeWorkbookThemeDisplay, themeOK bool) ([]NativeConditionalBarFillCellV1, bool) {
	// The downlevel element carries the positive fill; everything that decides
	// the geometry lives in the x14 twin, so a rule without one is not painted.
	extension := extensions[bar.extID]
	if bar.extID == "" || extension == nil {
		return nil, false
	}
	var positive [3]int
	seen := 0
	for _, child := range bar.legacy.children {
		switch child.name.Local {
		case "cfvo":
			if !nativeConditionalNode(child, ns, "cfvo", "type", "val") {
				return nil, false
			}
		case "color":
			rgb, ok := nativeBarColor(child, theme, themeOK)
			if !ok {
				return nil, false
			}
			positive = rgb
			seen++
		case "extLst":
		default:
			return nil, false
		}
	}
	if seen != 1 {
		return nil, false
	}
	// Excel writes these three on every x14 bar it authors. A gradient, a
	// clipped length or a middle axis changes the painted geometry and is left
	// to a later tier rather than approximated.
	if extension.attr("gradient") != "0" || extension.attr("minLength") != "0" || extension.attr("maxLength") != "100" {
		return nil, false
	}
	axisPosition := extension.attr("axisPosition")
	if axisPosition != "" && axisPosition != "automatic" && axisPosition != "none" {
		return nil, false
	}
	var bounds []*previewXML
	negative, border, negativeBorder, axisColor := positive, [3]int{}, [3]int{}, [3]int{}
	hasBorder, hasNegative, hasNegativeBorder, hasAxisColor := false, false, false, false
	for _, child := range extension.children {
		rgb, ok := [3]int{}, false
		if child.name.Local != "cfvo" {
			rgb, ok = nativeBarColor(child, theme, themeOK)
			if !ok {
				return nil, false
			}
		}
		switch child.name.Local {
		case "cfvo":
			bounds = append(bounds, child)
		case "negativeFillColor":
			negative, hasNegative = rgb, true
		case "borderColor":
			border, hasBorder = rgb, true
		case "negativeBorderColor":
			negativeBorder, hasNegativeBorder = rgb, true
		case "axisColor":
			axisColor, hasAxisColor = rgb, true
		default:
			return nil, false
		}
	}
	if len(bounds) != 2 {
		return nil, false
	}
	if extension.attr("border") == "1" && !hasBorder {
		return nil, false
	}
	if extension.attr("border") != "1" {
		hasBorder = false
	}
	if !hasNegative {
		negative = positive
	}
	if !hasNegativeBorder || extension.attr("negativeBarBorderColorSameAsPositive") != "0" {
		negativeBorder = border
	}

	var values []float64
	for row := bar.rect.top; row <= bar.rect.bottom; row++ {
		for column := bar.rect.left; column <= bar.rect.right; column++ {
			if number, ok := cells[[2]int{row, column}]; ok {
				values = append(values, number)
			}
		}
	}
	if len(values) == 0 {
		return nil, false
	}
	sorted := append([]float64(nil), values...)
	sort.Float64s(sorted)
	lowest, highest := sorted[0], sorted[len(sorted)-1]
	low, ok := nativeBarBound(bounds[0], lowest, highest, sorted, cells, math.Min(0, lowest))
	if !ok {
		return nil, false
	}
	high, ok := nativeBarBound(bounds[1], lowest, highest, sorted, cells, math.Max(0, highest))
	if !ok || !(high > low) {
		return nil, false
	}
	// The whole bound span maps onto the cell, so the axis is wherever zero
	// falls inside it and every bar runs from the axis to its own value.
	axisValue := math.Min(math.Max(0, low), high)
	if axisPosition == "none" {
		axisValue = low
	}
	axis := nativeBarPermille((axisValue - low) / (high - low))
	drawAxis := axisPosition != "none" && axis > 0 && axis < 1000 && hasAxisColor

	painted := make([]NativeConditionalBarFillCellV1, 0, len(values))
	for row := bar.rect.top; row <= bar.rect.bottom; row++ {
		for column := bar.rect.left; column <= bar.rect.right; column++ {
			value, ok := cells[[2]int{row, column}]
			if !ok {
				continue
			}
			edge := nativeBarPermille((math.Min(math.Max(value, low), high) - low) / (high - low))
			cell := NativeConditionalBarFillCellV1{Row: row, Column: column, Start: axis, End: edge, Axis: -1, Color: nativeScaleHex(positive)}
			if hasBorder {
				cell.BorderColor = nativeScaleHex(border)
			}
			if edge < axis {
				cell.Start, cell.End = edge, axis
				cell.Color = nativeScaleHex(negative)
				if hasBorder {
					cell.BorderColor = nativeScaleHex(negativeBorder)
				}
			}
			if drawAxis {
				cell.Axis, cell.AxisColor = axis, nativeScaleHex(axisColor)
			}
			if cell.End == cell.Start && !drawAxis {
				continue
			}
			painted = append(painted, cell)
		}
	}
	if len(painted) == 0 {
		return nil, false
	}
	return painted, true
}

// A data bar names its colours with five different element names, so the
// colour reader is keyed on the attributes rather than on the local name.
func nativeBarColor(node *previewXML, theme nativeWorkbookThemeDisplay, themeOK bool) ([3]int, bool) {
	if len(node.children) != 0 || strings.TrimSpace(node.text) != "" {
		return [3]int{}, false
	}
	rgb, index, tint := "", "", ""
	for _, attribute := range node.attrs {
		if attribute.Name.Space == "http://www.w3.org/2000/xmlns/" || attribute.Name.Space == "" && attribute.Name.Local == "xmlns" {
			continue
		}
		if attribute.Name.Space != "" {
			return [3]int{}, false
		}
		switch attribute.Name.Local {
		case "rgb":
			rgb = attribute.Value
		case "theme":
			index = attribute.Value
		case "tint":
			tint = attribute.Value
		default:
			return [3]int{}, false
		}
	}
	if rgb != "" {
		if index != "" || tint != "" || !nativeConditionalRGB.MatchString(rgb) {
			return [3]int{}, false
		}
		return nativeScaleChannels("#" + strings.ToUpper(rgb[2:]))
	}
	if index == "" || !themeOK {
		return [3]int{}, false
	}
	slot, err := strconv.Atoi(index)
	if err != nil || slot < 0 || slot > 11 {
		return [3]int{}, false
	}
	var shade *float64
	if tint != "" {
		value, err := strconv.ParseFloat(tint, 64)
		if err != nil || value < -1 || value > 1 {
			return [3]int{}, false
		}
		shade = &value
	}
	resolved, ok := theme.resolve(slot, shade)
	if !ok {
		return [3]int{}, false
	}
	return nativeScaleChannels(resolved)
}

func nativeBarBound(cfvo *previewXML, lowest, highest float64, sorted []float64, cells map[[2]int]float64, automatic float64) (float64, bool) {
	if !nativeConditionalNode(cfvo, cfvo.name.Space, "cfvo", "type", "val", "value") {
		return 0, false
	}
	raw := cfvo.attr("val")
	if raw == "" {
		raw = cfvo.attr("value")
	}
	switch cfvo.attr("type") {
	case "autoMin", "autoMax":
		return automatic, true
	case "min", "max", "num", "percent", "percentile":
		return nativeScaleBound(cfvo.attr("type"), raw, sorted, lowest, highest, cells)
	}
	return 0, false
}

func nativeBarPermille(fraction float64) int {
	if math.IsNaN(fraction) {
		return 0
	}
	value := int(math.Floor(fraction*1000 + 0.5))
	if value < 0 {
		return 0
	}
	if value > 1000 {
		return 1000
	}
	return value
}
