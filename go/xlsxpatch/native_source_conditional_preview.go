package xlsxpatch

import (
	"encoding/xml"
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

// NativeSourceStylePreviewV2 adds bounded conditional evidence to a read-only
// V1 base grid. It is not an editable workbook or a calculation result.
type NativeSourceStylePreviewV2 struct {
	Protocol        string                      `json:"protocol"`
	Version         int                         `json:"version"`
	ReadOnly        bool                        `json:"read_only"`
	Fidelity        string                      `json:"fidelity"`
	Grid            *NativeSourceStylePreviewV1 `json:"grid"`
	StylesSHA256    string                      `json:"styles_sha256"`
	WorksheetSHA256 string                      `json:"worksheet_sha256"`
	Warnings        []string                    `json:"warnings"`
	TextRule        *NativeSourceTextRuleV2     `json:"text_rule"`
	DataBar         *NativeSourceDataBarV2      `json:"data_bar"`
	FrozenView      *NativeSourceFrozenViewV2   `json:"frozen_view"`
}
type NativeSourceFrozenViewV2 struct {
	FrozenRows  int    `json:"frozen_rows"`
	TopLeftCell string `json:"top_left_cell"`
	ActivePane  string `json:"active_pane"`
}
type NativeSourceTextRuleV2 struct {
	Range           string                     `json:"range"`
	Priority        int                        `json:"priority"`
	DXFID           int                        `json:"dxf_id"`
	Text            string                     `json:"text"`
	Bold            bool                       `json:"bold"`
	FontColor       string                     `json:"font_color"`
	BackgroundColor string                     `json:"background_color"`
	Cells           []NativeSourceTextEffectV2 `json:"cells"`
}
type NativeSourceTextEffectV2 struct {
	Ref     string `json:"ref"`
	Value   string `json:"value"`
	Cached  bool   `json:"cached"`
	Matched bool   `json:"matched"`
}
type NativeSourceDataBarV2 struct {
	Range       string                    `json:"range"`
	Priority    int                       `json:"priority"`
	ExtensionID string                    `json:"extension_id"`
	Minimum     int                       `json:"minimum"`
	Maximum     int                       `json:"maximum"`
	MinLength   int                       `json:"min_length"`
	MaxLength   int                       `json:"max_length"`
	Color       string                    `json:"color"`
	Gradient    bool                      `json:"gradient"`
	Cells       []NativeSourceBarEffectV2 `json:"cells"`
}
type NativeSourceBarEffectV2 struct {
	Ref           string  `json:"ref"`
	Value         int     `json:"value"`
	LengthPercent float64 `json:"length_percent"`
}

const sourceX14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
const sourceXM = "http://schemas.microsoft.com/office/excel/2006/main"
const sourceBarRuleURI = "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"
const sourceBarSheetURI = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}"

var sourceConditionalText = regexp.MustCompile(`^[A-Z]{1,64}$`)
var sourceConditionalGUID = regexp.MustCompile(`^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$`)

//go:noinline
func sourceConditionalError() error {
	return fmt.Errorf("xlsxpatch: source conditional preview: unqualified rule, extension, differential style or input")
}

// PreviewNativeSourceStylesV2 qualifies one text-equality rule and/or one fixed
// numeric data bar over disjoint, fully populated, unmerged source ranges.
// Saved string caches are disclosed; formulas are never evaluated. V1 and all
// strict extraction/mutation entry points retain their existing boundaries.
func PreviewNativeSourceStylesV2(data []byte) (*NativeSourceStylePreviewV2, error) {
	out := &NativeSourceStylePreviewV2{Protocol: "injoffice.xlsx.source-style-preview", Version: 2, ReadOnly: true, Fidelity: "approximate", Warnings: []string{
		"The nested grid contains base styles. Conditional effects are separate source-qualified records; no workbook calculation or mutation is performed.",
		"Text equality uses stored uppercase ASCII strings, including explicitly marked saved formula caches. Differential background selection and browser data-bar geometry are compatibility approximations, not Excel print calibration.",
	}}
	grid, err := previewNativeSourceStyles(data, func(root *previewXML, styles, worksheet []byte) error {
		if err := out.qualifySourceConditions(root, styles); err != nil {
			return err
		}
		out.StylesSHA256 = nativeWorkbookDigest(styles)
		out.WorksheetSHA256 = nativeWorkbookDigest(worksheet)
		return nil
	})
	if err != nil {
		return nil, err
	}
	out.Grid = grid
	if err = out.qualifySourceConditionalInputs(); err != nil {
		return nil, err
	}
	return out, nil
}

func scNode(n *previewXML, ns, name string, attrs ...string) bool {
	keys := make([]xml.Name, len(attrs))
	for i, a := range attrs {
		keys[i] = xml.Name{Local: a}
	}
	return nativeRichNode(n, ns, name, keys...) && strings.TrimSpace(n.text) == ""
}

//go:noinline
func scLeaf(n *previewXML, ns, name string, attrs ...string) bool {
	return scNode(n, ns, name, attrs...) && len(n.children) == 0
}

//go:noinline
func scChildren(n *previewXML, ns string, names ...string) bool {
	if n == nil || len(n.children) != len(names) {
		return false
	}
	for i, name := range names {
		if n.children[i].name != (xml.Name{Space: ns, Local: name}) {
			return false
		}
	}
	return true
}
func scText(n *previewXML, ns, name string) (string, bool) {
	if n == nil {
		return "", false
	}
	copy := *n
	copy.text = ""
	return n.text, scLeaf(&copy, ns, name)
}
func scInt(s string, low, high int) (int, bool) {
	n, e := strconv.Atoi(s)
	return n, e == nil && n >= low && n <= high && strconv.Itoa(n) == s
}
func scRGB(n *previewXML, ns, name string) (string, bool) {
	if !scLeaf(n, ns, name, "rgb") || !nativeConditionalRGB.MatchString(n.attr("rgb")) {
		return "", false
	}
	return "#" + strings.ToUpper(n.attr("rgb")[2:]), true
}

func (out *NativeSourceStylePreviewV2) qualifySourceConditions(root *previewXML, styles []byte) error {
	ns := spreadsheetMLTransitional
	admitted := map[*previewXML]bool{}
	var extension *previewXML
	for _, cf := range root.children {
		if cf.name == (xml.Name{Space: ns, Local: "extLst"}) {
			if extension != nil {
				return sourceConditionalError()
			}
			extension = cf
			continue
		}
		if cf.name != (xml.Name{Space: ns, Local: "conditionalFormatting"}) {
			continue
		}
		if !scNode(cf, ns, "conditionalFormatting", "sqref") || !scChildren(cf, ns, "cfRule") {
			return sourceConditionalError()
		}
		rule := cf.children[0]
		priority, ok := scInt(rule.attr("priority"), 1, 65535)
		if !ok {
			return sourceConditionalError()
		}
		switch rule.attr("type") {
		case "cellIs":
			if out.TextRule != nil || !scNode(rule, ns, "cfRule", "type", "priority", "operator", "aboveAverage", "equalAverage", "bottom", "percent", "rank", "text", "dxfId") || rule.attr("operator") != "equal" || !scChildren(rule, ns, "formula") {
				return sourceConditionalError()
			}
			for _, a := range rule.attrs {
				switch a.Name.Local {
				case "aboveAverage", "equalAverage", "bottom", "percent", "rank":
					if a.Value != "0" {
						return sourceConditionalError()
					}
				case "text":
					if a.Value != "" {
						return sourceConditionalError()
					}
				}
			}
			literal, ok := scText(rule.children[0], ns, "formula")
			if !ok || len(literal) < 3 || literal[0] != '"' || literal[len(literal)-1] != '"' || !sourceConditionalText.MatchString(literal[1:len(literal)-1]) {
				return sourceConditionalError()
			}
			id, ok := scInt(rule.attr("dxfId"), 0, 0)
			if !ok {
				return sourceConditionalError()
			}
			out.TextRule = &NativeSourceTextRuleV2{Range: cf.attr("sqref"), Priority: priority, DXFID: id, Text: literal[1 : len(literal)-1], Cells: []NativeSourceTextEffectV2{}}
		case "dataBar":
			if out.DataBar != nil || !scNode(rule, ns, "cfRule", "type", "priority") || !scChildren(rule, ns, "dataBar", "extLst") {
				return sourceConditionalError()
			}
			bar := rule.children[0]
			if !scNode(bar, ns, "dataBar", "showValue", "minLength", "maxLength") || bar.attr("showValue") != "1" || !scChildren(bar, ns, "cfvo", "cfvo", "color") {
				return sourceConditionalError()
			}
			lo, ok := scInt(bar.attr("minLength"), 0, 100)
			hi, ok2 := scInt(bar.attr("maxLength"), 0, 100)
			if !ok || !ok2 || lo >= hi {
				return sourceConditionalError()
			}
			bounds := [2]int{}
			for i := 0; i < 2; i++ {
				n := bar.children[i]
				if !scLeaf(n, ns, "cfvo", "type", "val") || n.attr("type") != "num" {
					return sourceConditionalError()
				}
				v, ok := scInt(n.attr("val"), 0, 1000000000)
				if !ok {
					return sourceConditionalError()
				}
				bounds[i] = v
			}
			color, ok := scRGB(bar.children[2], ns, "color")
			if !ok || bounds[0] >= bounds[1] {
				return sourceConditionalError()
			}
			list := rule.children[1]
			if !scNode(list, ns, "extLst") || !scChildren(list, ns, "ext") {
				return sourceConditionalError()
			}
			ext := list.children[0]
			if !scNode(ext, ns, "ext", "uri") || ext.attr("uri") != sourceBarRuleURI || !scChildren(ext, sourceX14, "id") {
				return sourceConditionalError()
			}
			id, ok := scText(ext.children[0], sourceX14, "id")
			if !ok || !sourceConditionalGUID.MatchString(id) {
				return sourceConditionalError()
			}
			out.DataBar = &NativeSourceDataBarV2{Range: cf.attr("sqref"), Priority: priority, ExtensionID: id, Minimum: bounds[0], Maximum: bounds[1], MinLength: lo, MaxLength: hi, Color: color, Cells: []NativeSourceBarEffectV2{}}
		default:
			return sourceConditionalError()
		}
		admitted[cf] = true
	}
	if out.TextRule == nil && out.DataBar == nil {
		return sourceConditionalError()
	}
	if out.TextRule != nil && out.DataBar != nil && out.TextRule.Priority == out.DataBar.Priority {
		return sourceConditionalError()
	}
	if out.DataBar != nil {
		if err := out.qualifySourceBarExtension(extension); err != nil {
			return err
		}
		admitted[extension] = true
	} else if extension != nil {
		return sourceConditionalError()
	}
	if err := out.qualifySourceDXF(styles); err != nil {
		return err
	}
	if err := out.qualifySourceFrozenView(root, admitted); err != nil {
		return err
	}
	return qualifySourceStyleSheetWithNodes(root, ns, admitted)
}

func (out *NativeSourceStylePreviewV2) qualifySourceBarExtension(list *previewXML) error {
	ns := spreadsheetMLTransitional
	bar := out.DataBar
	if !scNode(list, ns, "extLst") || !scChildren(list, ns, "ext") {
		return sourceConditionalError()
	}
	ext := list.children[0]
	if !scNode(ext, ns, "ext", "uri") || ext.attr("uri") != sourceBarSheetURI || !scChildren(ext, sourceX14, "conditionalFormattings") {
		return sourceConditionalError()
	}
	group := ext.children[0]
	if !scNode(group, sourceX14, "conditionalFormattings") || !scChildren(group, sourceX14, "conditionalFormatting") {
		return sourceConditionalError()
	}
	cf := group.children[0]
	if !scNode(cf, sourceX14, "conditionalFormatting") || len(cf.children) != 2 || cf.children[0].name != (xml.Name{Space: sourceX14, Local: "cfRule"}) {
		return sourceConditionalError()
	}
	ref, ok := scText(cf.children[1], sourceXM, "sqref")
	if !ok || ref != bar.Range {
		return sourceConditionalError()
	}
	rule := cf.children[0]
	if !scNode(rule, sourceX14, "cfRule", "type", "id") || rule.attr("type") != "dataBar" || rule.attr("id") != bar.ExtensionID || !scChildren(rule, sourceX14, "dataBar") {
		return sourceConditionalError()
	}
	b := rule.children[0]
	if !scNode(b, sourceX14, "dataBar", "minLength", "maxLength", "axisPosition", "gradient") || b.attr("axisPosition") != "none" || b.attr("gradient") != "true" || b.attr("minLength") != strconv.Itoa(bar.MinLength) || b.attr("maxLength") != strconv.Itoa(bar.MaxLength) || !scChildren(b, sourceX14, "cfvo", "cfvo", "negativeFillColor", "axisColor") {
		return sourceConditionalError()
	}
	for i, v := range []int{bar.Minimum, bar.Maximum} {
		n := b.children[i]
		if !scNode(n, sourceX14, "cfvo", "type") || n.attr("type") != "num" || !scChildren(n, sourceXM, "f") {
			return sourceConditionalError()
		}
		val, ok := scText(n.children[0], sourceXM, "f")
		if !ok || val != strconv.Itoa(v) {
			return sourceConditionalError()
		}
	}
	negative, ok := scRGB(b.children[2], sourceX14, "negativeFillColor")
	axis, ok2 := scRGB(b.children[3], sourceX14, "axisColor")
	// Negative values and axes are excluded; admit only this explicit inert form.
	if !ok || !ok2 || negative != bar.Color || axis != "#000000" {
		return sourceConditionalError()
	}
	bar.Gradient = true
	return nil
}

func (out *NativeSourceStylePreviewV2) qualifySourceDXF(data []byte) error {
	root, err := parsePreviewXML(data)
	if err != nil {
		return err
	}
	ns := spreadsheetMLTransitional
	count := 0
	var walk func(*previewXML) bool
	walk = func(n *previewXML) bool {
		if n.name.Local == "extLst" || n.name.Local == "AlternateContent" {
			return false
		}
		if n.name.Local == "dxfs" {
			count++
		}
		for _, c := range n.children {
			if !walk(c) {
				return false
			}
		}
		return true
	}
	if !walk(root) {
		return sourceConditionalError()
	}
	dxfs := root.child("dxfs")
	if out.TextRule == nil {
		if count == 0 {
			return nil
		}
		if count == 1 && scLeaf(dxfs, ns, "dxfs", "count") && dxfs.attr("count") == "0" {
			return nil
		}
		return sourceConditionalError()
	}
	if count != 1 || !scNode(dxfs, ns, "dxfs", "count") || dxfs.attr("count") != "1" || !scChildren(dxfs, ns, "dxf") {
		return sourceConditionalError()
	}
	dxf := dxfs.children[0]
	if !scNode(dxf, ns, "dxf") || !scChildren(dxf, ns, "font", "fill") {
		return sourceConditionalError()
	}
	font, fill := dxf.children[0], dxf.children[1]
	if !scNode(font, ns, "font") || !scChildren(font, ns, "b", "color") || !scLeaf(font.children[0], ns, "b", "val") || font.children[0].attr("val") != "1" || !scNode(fill, ns, "fill") || !scChildren(fill, ns, "patternFill") {
		return sourceConditionalError()
	}
	pattern := fill.children[0]
	if !scNode(pattern, ns, "patternFill") || !scChildren(pattern, ns, "bgColor") {
		return sourceConditionalError()
	}
	fg, ok := scRGB(font.children[1], ns, "color")
	bg, ok2 := scRGB(pattern.children[0], ns, "bgColor")
	if !ok || !ok2 {
		return sourceConditionalError()
	}
	out.TextRule.Bold = true
	out.TextRule.FontColor = fg
	out.TextRule.BackgroundColor = bg
	return nil
}

func (out *NativeSourceStylePreviewV2) qualifySourceConditionalInputs() error {
	sheet := out.Grid.Sheets[0]
	if out.FrozenView != nil && out.FrozenView.FrozenRows >= len(sheet.RowHeights) {
		return sourceConditionalError()
	}
	cells := map[string]NativeSourceStyleCellV1{}
	styles := map[int]NativeSourceStyleV1{}
	for _, c := range sheet.Cells {
		cells[c.Ref] = c
	}
	for _, s := range out.Grid.Styles {
		styles[s.ID] = s
	}
	occupied := map[string]bool{}
	visit := func(ref string, apply func(NativeSourceStyleCellV1) error) error {
		top, left, bottom, right, err := parseDimensionReference(ref)
		if err != nil || !nativeConditionalRange.MatchString(ref) || top < 0 || left < 0 || bottom >= len(sheet.RowHeights) || right >= len(sheet.ColumnWidths) || bottom < top || right < left || (bottom-top+1)*(right-left+1) > 4096 {
			return sourceConditionalError()
		}
		for row := top; row <= bottom; row++ {
			for col := left; col <= right; col++ {
				for _, m := range sheet.Merges {
					if row >= m.Row && row <= m.EndRow && col >= m.Column && col <= m.EndColumn {
						return sourceConditionalError()
					}
				}
				cellRef := cellReference(row, col)
				c, ok := cells[cellRef]
				if !ok || occupied[cellRef] {
					return sourceConditionalError()
				}
				occupied[cellRef] = true
				if err := apply(c); err != nil {
					return err
				}
			}
		}
		return nil
	}
	if rule := out.TextRule; rule != nil {
		if err := visit(rule.Range, func(c NativeSourceStyleCellV1) error {
			if c.Kind != "string" || !sourceConditionalText.MatchString(c.Text) {
				return sourceConditionalError()
			}
			rule.Cells = append(rule.Cells, NativeSourceTextEffectV2{c.Ref, c.Text, c.Cached, c.Text == rule.Text})
			return nil
		}); err != nil {
			return err
		}
	}
	if bar := out.DataBar; bar != nil {
		if err := visit(bar.Range, func(c NativeSourceStyleCellV1) error {
			v, ok := scInt(c.Lexical, bar.Minimum, bar.Maximum)
			if !ok || c.Kind != "number" || c.Formula != "" || c.Cached || styles[int(c.StyleID)].FillColor != "" {
				return sourceConditionalError()
			}
			length := float64(bar.MinLength) + float64(v-bar.Minimum)/float64(bar.Maximum-bar.Minimum)*float64(bar.MaxLength-bar.MinLength)
			bar.Cells = append(bar.Cells, NativeSourceBarEffectV2{c.Ref, v, length})
			return nil
		}); err != nil {
			return err
		}
	}
	return nil
}

// Frozen panes affect the interactive viewport, not stored grid coordinates.
// Record and disclose the bounded source form; do not silently simulate it.
func (out *NativeSourceStylePreviewV2) qualifySourceFrozenView(root *previewXML, admitted map[*previewXML]bool) error {
	ns := spreadsheetMLTransitional
	view := root.child("sheetViews").child("sheetView")
	if view == nil {
		return nil
	}
	pane := view.child("pane")
	if pane == nil {
		return nil
	}
	if !scLeaf(pane, ns, "pane", "xSplit", "ySplit", "topLeftCell", "activePane", "state") || pane.attr("xSplit") != "0" || pane.attr("state") != "frozen" || pane.attr("activePane") != "bottomLeft" || !scChildren(view, ns, "pane", "selection", "selection") {
		return sourceConditionalError()
	}
	rows, ok := scInt(pane.attr("ySplit"), 1, 127)
	if !ok || pane.attr("topLeftCell") != cellReference(rows, 0) {
		return sourceConditionalError()
	}
	for i, which := range []string{"topLeft", "bottomLeft"} {
		selection := view.children[i+1]
		if !scLeaf(selection, ns, "selection", "pane", "activeCell", "activeCellId", "sqref") || selection.attr("pane") != which || selection.attr("activeCell") != "A1" || selection.attr("activeCellId") != "0" || selection.attr("sqref") != "A1" {
			return sourceConditionalError()
		}
		admitted[selection] = true
	}
	admitted[pane] = true
	out.FrozenView = &NativeSourceFrozenViewV2{rows, pane.attr("topLeftCell"), pane.attr("activePane")}
	out.Warnings = append(out.Warnings, "Source frozen rows and pane selection are recorded but not applied to the read-only grid viewport.")
	return nil
}
