package xlsxpatch

import (
	"fmt"
	"math"
	"regexp"
	"sort"
	"strconv"
	"strings"
)

// NativeConditionalScaleFillPreviewV1 is a read-only colour-scale overlay. It
// never changes the native style inventory and never removes the sheet's
// CONDITIONAL_FORMATTING preservation entry: the source rules stay preserved
// exactly and this tier only reports the colour Excel would paint.
type NativeConditionalScaleFillPreviewV1 struct {
	SheetID   string                              `json:"sheet_id"`
	SheetPart string                              `json:"sheet_part"`
	Status    string                              `json:"status"`
	Warnings  []string                            `json:"warnings"`
	Cells     []NativeConditionalScaleFillCellV1  `json:"cells,omitempty"`
	Ranges    []NativeConditionalScaleFillRangeV1 `json:"ranges,omitempty"`
}

// NativeConditionalScaleFillRangeV1 discloses which source range each painted
// colour came from, so the overlay can be audited against the worksheet.
type NativeConditionalScaleFillRangeV1 struct {
	Ref      string `json:"ref"`
	Priority int    `json:"priority"`
	Stops    int    `json:"stops"`
}

// NativeConditionalScaleFillCellV1 is one painted cell. Coordinates are
// zero-based; Color is an opaque #RRGGBB.
type NativeConditionalScaleFillCellV1 struct {
	Row    int    `json:"row"`
	Column int    `json:"column"`
	Color  string `json:"color"`
}

// Excel writes a cfvo bound either as a decimal literal or, for the `num`
// type, as a reference to one cell holding the bound. Only these two exact
// spellings are resolved; a formula, a name or a cross-sheet reference is not
// evaluated and its rule is left unpainted.
var nativeScaleDecimal = regexp.MustCompile(`^-?(0|[1-9][0-9]{0,14})(\.[0-9]{1,10})?$`)
var nativeScaleCellReference = regexp.MustCompile(`^\$?([A-Z]{1,3})\$?([1-9][0-9]{0,6})$`)

const nativeScaleMaximumRangeCells = 4096

type nativeScaleRange struct {
	top, left, bottom, right int
}

func (r nativeScaleRange) intersects(other nativeScaleRange) bool {
	return r.top <= other.bottom && r.bottom >= other.top && r.left <= other.right && r.right >= other.left
}

func (r nativeScaleRange) contains(row, column int) bool {
	return row >= r.top && row <= r.bottom && column >= r.left && column <= r.right
}

func previewNativeConditionalScaleFills(pkg *nativeWorkbookPackage, sheets []NativeWorkbookSheetV2) []NativeConditionalScaleFillPreviewV1 {
	theme, ok := nativeScaleWorkbookTheme(pkg)
	result := []NativeConditionalScaleFillPreviewV1{}
	budget := 16384
	for _, sheet := range sheets {
		if len(result) == 64 || budget <= 0 {
			break
		}
		entry := previewNativeConditionalScaleFill(pkg, sheet, theme, ok, budget)
		if entry != nil {
			result = append(result, *entry)
			budget -= len(entry.Cells)
		}
	}
	return result
}

func nativeScaleWorkbookTheme(pkg *nativeWorkbookPackage) (nativeWorkbookThemeDisplay, bool) {
	location, err := locateWorkbookPartBytes(pkg.index, func(name string) ([]byte, bool) { v, ok := pkg.files[name]; return v, ok })
	if err != nil {
		return nativeWorkbookThemeDisplay{}, false
	}
	namespace, relNamespace := spreadsheetMLTransitional, officeRelNamespaceTransitional
	expected, opposing := relTypeThemeTransitional, relTypeThemeStrict
	if location.strict {
		namespace, relNamespace = spreadsheetMLStrict, officeRelNamespaceStrict
		expected, opposing = relTypeThemeStrict, relTypeThemeTransitional
	}
	extractor := &nativeWorkbookExtractor{
		pkg: pkg, workbook: location, namespace: namespace, relNamespace: relNamespace,
		modeled: map[string]bool{}, unsupportedKeys: map[string]bool{}, claimedXML: map[string]bool{},
	}
	theme, err := parseWorkbookThemeDisplay(extractor, expected, opposing)
	if err != nil {
		return nativeWorkbookThemeDisplay{}, false
	}
	return theme, true
}

func previewNativeConditionalScaleFill(pkg *nativeWorkbookPackage, sheet NativeWorkbookSheetV2, theme nativeWorkbookThemeDisplay, themeOK bool, budget int) *NativeConditionalScaleFillPreviewV1 {
	root, err := parsePreviewXML(pkg.files[sheet.PartName])
	if err != nil {
		return nil
	}
	ns := root.name.Space
	if root.name.Local != "worksheet" || ns != spreadsheetMLTransitional && ns != spreadsheetMLStrict {
		return nil
	}
	type candidate struct {
		rect     nativeScaleRange
		ref      string
		priority int
		rule     *previewXML
		scale    *previewXML
	}
	var scales []candidate
	var occupied []nativeScaleRange
	// Every rule on the sheet claims its range, including the ones this tier
	// does not paint: a colour scale that shares cells with another rule has a
	// priority-dependent result that is not inferred here.
	for _, format := range root.children {
		if format.name.Local != "conditionalFormatting" || format.name.Space != ns {
			continue
		}
		rect, ok := nativeScaleParseRange(format.attr("sqref"))
		if !ok {
			return nativeScaleRefusal(sheet, "a conditional range is outside the single-rectangle A1 subset.")
		}
		occupied = append(occupied, rect)
		for _, rule := range format.children {
			if rule.name.Local != "cfRule" || rule.name.Space != ns || rule.attr("type") != "colorScale" {
				continue
			}
			scale := rule.child("colorScale")
			if scale == nil {
				continue
			}
			priority, err := strconv.Atoi(rule.attr("priority"))
			if err != nil || priority < 1 || priority > 2147483647 {
				continue
			}
			scales = append(scales, candidate{rect: rect, ref: format.attr("sqref"), priority: priority, rule: rule, scale: scale})
		}
	}
	if len(scales) == 0 {
		return nil
	}
	// Worksheet extensions carry x14 rules whose ranges are written as an
	// `xm:sqref`. A colour scale that meets one of them is left unpainted.
	for _, ext := range nativeScaleExtensionRanges(root) {
		occupied = append(occupied, ext)
	}
	cells := map[[2]int]float64{}
	for _, cell := range sheet.Cells {
		if cell.Ref != cellReference(cell.Row, cell.Column) {
			return nativeScaleRefusal(sheet, "source cell identities are ambiguous.")
		}
		if number, ok := nativeScaleCellNumber(cell); ok {
			cells[[2]int{cell.Row, cell.Column}] = number
		}
	}
	entry := &NativeConditionalScaleFillPreviewV1{SheetID: sheet.ID, SheetPart: sheet.PartName, Status: "unavailable"}
	unpainted := 0
	sort.SliceStable(scales, func(i, j int) bool { return scales[i].priority < scales[j].priority })
	for _, scale := range scales {
		overlaps := 0
		for _, claimed := range occupied {
			if scale.rect.intersects(claimed) {
				overlaps++
			}
		}
		if overlaps != 1 {
			unpainted++
			continue
		}
		for _, merge := range sheet.MergedRanges {
			if scale.rect.intersects(nativeScaleRange{merge.Row, merge.Column, merge.EndRow, merge.EndColumn}) {
				overlaps = 0
			}
		}
		if overlaps != 1 {
			unpainted++
			continue
		}
		size := (scale.rect.bottom - scale.rect.top + 1) * (scale.rect.right - scale.rect.left + 1)
		if size > nativeScaleMaximumRangeCells || len(entry.Cells)+size > budget {
			unpainted++
			continue
		}
		painted, ok := nativeScalePaintRange(scale.scale, scale.rect, cells, theme, themeOK)
		if !ok {
			unpainted++
			continue
		}
		entry.Cells = append(entry.Cells, painted...)
		entry.Ranges = append(entry.Ranges, NativeConditionalScaleFillRangeV1{Ref: scale.ref, Priority: scale.priority, Stops: len(scale.scale.children) / 2})
	}
	if len(entry.Cells) == 0 {
		return nativeScaleRefusal(sheet, "no colour-scale rule resolved to an exact source-qualified colour.")
	}
	sort.SliceStable(entry.Cells, func(i, j int) bool {
		if entry.Cells[i].Row != entry.Cells[j].Row {
			return entry.Cells[i].Row < entry.Cells[j].Row
		}
		return entry.Cells[i].Column < entry.Cells[j].Column
	})
	entry.Status = "available"
	warning := fmt.Sprintf("Read-only colour-scale preview: %d cells across %d source ranges. Two- and three-stop scales are interpolated between explicit RGB or theme stops from saved cell values; formulas are never recalculated.", len(entry.Cells), len(entry.Ranges))
	if unpainted > 0 {
		warning += fmt.Sprintf(" %d colour-scale rule(s) were not painted because a bound, a stop colour or an overlapping rule is outside this subset; those source rules remain preserved.", unpainted)
	}
	entry.Warnings = []string{warning}
	return entry
}

func nativeScaleRefusal(sheet NativeWorkbookSheetV2, reason string) *NativeConditionalScaleFillPreviewV1 {
	return &NativeConditionalScaleFillPreviewV1{
		SheetID: sheet.ID, SheetPart: sheet.PartName, Status: "unavailable",
		Warnings: []string{"Colour-scale preview unavailable: " + reason + " No colour-scale fills were applied; source rules remain preserved."},
	}
}

// nativeScaleExtensionRanges reports the ranges claimed by x14 worksheet
// extensions. An unreadable extension claims the whole sheet, so a colour
// scale never paints over a rule this tier cannot see.
func nativeScaleExtensionRanges(root *previewXML) []nativeScaleRange {
	var ranges []nativeScaleRange
	var walk func(*previewXML)
	walk = func(n *previewXML) {
		if n.name.Local == "conditionalFormatting" && n.name.Space != root.name.Space {
			found := false
			for _, child := range n.children {
				if child.name.Local != "sqref" {
					continue
				}
				if rect, ok := nativeScaleParseRange(child.text); ok {
					ranges = append(ranges, rect)
					found = true
				}
			}
			if !found {
				ranges = append(ranges, nativeScaleRange{0, 0, 1048575, 16383})
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
	return ranges
}

func nativeScaleParseRange(sqref string) (nativeScaleRange, bool) {
	sqref = strings.TrimSpace(sqref)
	if sqref == "" || !nativeConditionalRange.MatchString(sqref) {
		return nativeScaleRange{}, false
	}
	top, left, bottom, right, err := parseDimensionReference(sqref)
	if err != nil || top > bottom || left > right {
		return nativeScaleRange{}, false
	}
	return nativeScaleRange{top, left, bottom, right}, true
}

func nativeScaleCellNumber(cell NativeWorkbookCellV2) (float64, bool) {
	value := cell.Value
	if cell.Formula != nil {
		value = cell.Formula.Cached
	}
	if value == nil || value.Kind != "number" || value.Storage != "number" || value.Lexical == nil || !nativeScaleDecimal.MatchString(*value.Lexical) {
		return 0, false
	}
	number, err := strconv.ParseFloat(*value.Lexical, 64)
	if err != nil || math.IsNaN(number) || math.IsInf(number, 0) {
		return 0, false
	}
	return number, true
}

type nativeScaleStop struct {
	kind  string
	value float64
	rgb   [3]int
}

func nativeScalePaintRange(scale *previewXML, rect nativeScaleRange, cells map[[2]int]float64, theme nativeWorkbookThemeDisplay, themeOK bool) ([]NativeConditionalScaleFillCellV1, bool) {
	var cfvos, colors []*previewXML
	for _, child := range scale.children {
		switch child.name.Local {
		case "cfvo":
			cfvos = append(cfvos, child)
		case "color":
			colors = append(colors, child)
		default:
			return nil, false
		}
	}
	if len(cfvos) != len(colors) || len(cfvos) < 2 || len(cfvos) > 3 {
		return nil, false
	}
	// Excel excludes blanks, text, booleans and errors from a colour scale's
	// own minimum and maximum and leaves those cells unfilled.
	var values []float64
	for row := rect.top; row <= rect.bottom; row++ {
		for column := rect.left; column <= rect.right; column++ {
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
	low, high := sorted[0], sorted[len(sorted)-1]

	stops := make([]nativeScaleStop, len(cfvos))
	for index, cfvo := range cfvos {
		if !nativeConditionalNode(cfvo, scale.name.Space, "cfvo", "type", "val") {
			return nil, false
		}
		bound, ok := nativeScaleBound(cfvo.attr("type"), cfvo.attr("val"), sorted, low, high, cells)
		if !ok {
			return nil, false
		}
		rgb, ok := nativeScaleColor(colors[index], scale.name.Space, theme, themeOK)
		if !ok {
			return nil, false
		}
		stops[index] = nativeScaleStop{kind: cfvo.attr("type"), value: bound, rgb: rgb}
	}
	for index := 1; index < len(stops); index++ {
		if stops[index].value < stops[index-1].value {
			return nil, false
		}
	}
	painted := make([]NativeConditionalScaleFillCellV1, 0, len(values))
	for row := rect.top; row <= rect.bottom; row++ {
		for column := rect.left; column <= rect.right; column++ {
			number, ok := cells[[2]int{row, column}]
			if !ok {
				continue
			}
			painted = append(painted, NativeConditionalScaleFillCellV1{Row: row, Column: column, Color: nativeScaleInterpolate(stops, number)})
		}
	}
	return painted, true
}

func nativeScaleBound(kind, raw string, sorted []float64, low, high float64, cells map[[2]int]float64) (float64, bool) {
	switch kind {
	case "min":
		return low, true
	case "max":
		return high, true
	case "num":
		if nativeScaleDecimal.MatchString(raw) {
			number, err := strconv.ParseFloat(raw, 64)
			return number, err == nil
		}
		match := nativeScaleCellReference.FindStringSubmatch(raw)
		if match == nil {
			return 0, false
		}
		row, err := strconv.Atoi(match[2])
		if err != nil {
			return 0, false
		}
		column := 0
		for _, letter := range match[1] {
			column = column*26 + int(letter-'A') + 1
		}
		number, ok := cells[[2]int{row - 1, column - 1}]
		return number, ok
	case "percent":
		percent, ok := nativeScalePercent(raw)
		if !ok {
			return 0, false
		}
		return low + percent/100*(high-low), true
	case "percentile":
		percent, ok := nativeScalePercent(raw)
		if !ok {
			return 0, false
		}
		// PERCENTILE.INC over the range's own numeric values, which is the
		// function Excel documents for a percentile colour-scale bound.
		position := percent / 100 * float64(len(sorted)-1)
		lower := int(math.Floor(position))
		if lower >= len(sorted)-1 {
			return sorted[len(sorted)-1], true
		}
		return sorted[lower] + (position-float64(lower))*(sorted[lower+1]-sorted[lower]), true
	}
	return 0, false
}

func nativeScalePercent(raw string) (float64, bool) {
	if !nativeScaleDecimal.MatchString(raw) {
		return 0, false
	}
	percent, err := strconv.ParseFloat(raw, 64)
	if err != nil || percent < 0 || percent > 100 {
		return 0, false
	}
	return percent, true
}

func nativeScaleColor(color *previewXML, namespace string, theme nativeWorkbookThemeDisplay, themeOK bool) ([3]int, bool) {
	if color.name.Space != namespace || len(color.children) != 0 || strings.TrimSpace(color.text) != "" {
		return [3]int{}, false
	}
	if rgb := color.attr("rgb"); rgb != "" {
		if !nativeConditionalNode(color, namespace, "color", "rgb") || !nativeConditionalRGB.MatchString(rgb) {
			return [3]int{}, false
		}
		return nativeScaleChannels("#" + strings.ToUpper(rgb[2:]))
	}
	if index := color.attr("theme"); index != "" {
		if !nativeConditionalNode(color, namespace, "color", "theme", "tint") || !themeOK {
			return [3]int{}, false
		}
		slot, err := strconv.Atoi(index)
		if err != nil || slot < 0 || slot > 11 {
			return [3]int{}, false
		}
		var tint *float64
		if raw := color.attr("tint"); raw != "" {
			value, err := strconv.ParseFloat(raw, 64)
			if err != nil || value < -1 || value > 1 {
				return [3]int{}, false
			}
			tint = &value
		}
		resolved, ok := theme.resolve(slot, tint)
		if !ok {
			return [3]int{}, false
		}
		return nativeScaleChannels(resolved)
	}
	return [3]int{}, false
}

func nativeScaleChannels(rgb string) ([3]int, bool) {
	if len(rgb) != 7 || rgb[0] != '#' {
		return [3]int{}, false
	}
	var channels [3]int
	for index := 0; index < 3; index++ {
		value, err := strconv.ParseUint(rgb[1+index*2:3+index*2], 16, 8)
		if err != nil {
			return [3]int{}, false
		}
		channels[index] = int(value)
	}
	return channels, true
}

// Excel interpolates a colour scale channel-wise in sRGB between the two
// bounds that surround the value, and clamps outside the end bounds.
func nativeScaleInterpolate(stops []nativeScaleStop, value float64) string {
	last := len(stops) - 1
	if value <= stops[0].value {
		return nativeScaleHex(stops[0].rgb)
	}
	if value >= stops[last].value {
		return nativeScaleHex(stops[last].rgb)
	}
	for index := 1; index <= last; index++ {
		if value > stops[index].value {
			continue
		}
		span := stops[index].value - stops[index-1].value
		if span <= 0 {
			return nativeScaleHex(stops[index].rgb)
		}
		ratio := (value - stops[index-1].value) / span
		var mixed [3]int
		for channel := 0; channel < 3; channel++ {
			from := float64(stops[index-1].rgb[channel])
			to := float64(stops[index].rgb[channel])
			mixed[channel] = int(math.Floor(from + ratio*(to-from) + 0.5))
		}
		return nativeScaleHex(mixed)
	}
	return nativeScaleHex(stops[last].rgb)
}

func nativeScaleHex(rgb [3]int) string {
	return fmt.Sprintf("#%02X%02X%02X", rgb[0], rgb[1], rgb[2])
}
