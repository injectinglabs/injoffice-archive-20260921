package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"math"
	"regexp"
	"strconv"
	"strings"
)

const (
	sparklineExtensionURI = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"
	x14Namespace          = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
	xmNamespace           = "http://schemas.microsoft.com/office/excel/2006/main"
)

// SparklineOptions is the losslessly editable x14 sparkline option subset.
// Unsupported x14 state is reported through SparklineInfo.Warnings and blocks
// SetSparklines, so a native save never silently normalizes it away.
type SparklineOptions struct {
	EmptyCells   string          `json:"emptyCells,omitempty"`
	RightToLeft  bool            `json:"rightToLeft,omitempty"`
	ShowMarkers  bool            `json:"showMarkers,omitempty"`
	ShowHigh     bool            `json:"showHigh,omitempty"`
	ShowLow      bool            `json:"showLow,omitempty"`
	ShowFirst    bool            `json:"showFirst,omitempty"`
	ShowLast     bool            `json:"showLast,omitempty"`
	ShowNegative bool            `json:"showNegative,omitempty"`
	LineWeight   float64         `json:"lineWeight,omitempty"`
	Min          *float64        `json:"min,omitempty"`
	Max          *float64        `json:"max,omitempty"`
	Colors       SparklineColors `json:"colors,omitempty"`
}

type SparklineColors struct {
	Series   string `json:"series,omitempty"`
	Negative string `json:"negative,omitempty"`
	Markers  string `json:"markers,omitempty"`
	High     string `json:"high,omitempty"`
	Low      string `json:"low,omitempty"`
	First    string `json:"first,omitempty"`
	Last     string `json:"last,omitempty"`
	Axis     string `json:"axis,omitempty"`
}

// SparklineInfo mirrors one x14:sparkline. GroupID is deterministic within a
// file (worksheet part plus group ordinal); it is not a native stable ID.
type SparklineInfo struct {
	ID              string           `json:"id"`
	GroupID         string           `json:"groupId,omitempty"`
	Type            string           `json:"type"`
	SourceSheetName string           `json:"sourceSheetName,omitempty"`
	SourceRef       string           `json:"sourceRef,omitempty"`
	TargetSheetName string           `json:"targetSheetName"`
	TargetRef       string           `json:"targetRef"`
	Options         SparklineOptions `json:"options"`
	Warnings        []string         `json:"warnings,omitempty"`
}

type SparklineWriteItem struct {
	SourceSheetName string `json:"sourceSheetName"`
	SourceRef       string `json:"sourceRef"`
	TargetCellRef   string `json:"targetCellRef"`
}

// SparklineWriteGroup is one native x14 group. Options are shared by every
// item, matching the file format rather than pretending per-item styles exist.
type SparklineWriteGroup struct {
	ID              string               `json:"id,omitempty"`
	TargetSheetName string               `json:"targetSheetName"`
	Type            string               `json:"type"`
	Options         SparklineOptions     `json:"options,omitempty"`
	Sparklines      []SparklineWriteItem `json:"sparklines"`
}

type sparkXMLNode struct {
	XMLName  xml.Name
	Attrs    []xml.Attr     `xml:",any,attr"`
	Text     string         `xml:",chardata"`
	Children []sparkXMLNode `xml:",any"`
}

// ReadSparklines reads the standard Microsoft x14 worksheet extension used by
// Excel. Structurally malformed XML is an error. Valid but unrepresentable
// state is returned with explicit warnings so callers can skip it visibly.
func ReadSparklines(data []byte) ([]SparklineInfo, error) {
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: read sparklines: %w", err)
	}
	files := make(map[string]*zip.File, len(zr.File))
	for _, file := range zr.File {
		if _, exists := files[file.Name]; exists {
			return nil, fmt.Errorf("xlsxpatch: read sparklines: duplicate entry %q", file.Name)
		}
		files[file.Name] = file
	}
	read := func(name string) (string, bool) {
		file, ok := files[name]
		if !ok {
			return "", false
		}
		rc, openErr := file.Open()
		if openErr != nil {
			return "", false
		}
		defer rc.Close()
		value, readErr := io.ReadAll(rc)
		return string(value), readErr == nil
	}
	sheets, err := readWorkbookSheets(read)
	if err != nil {
		return nil, err
	}
	var result []SparklineInfo
	for _, sheet := range sheets {
		value, ok := read(sheet.Part)
		if !ok {
			return nil, fmt.Errorf("xlsxpatch: read sparklines: missing worksheet %q", sheet.Part)
		}
		extensions, err := sparklineExtensions(value)
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: read sparklines: %s: %w", sheet.Part, err)
		}
		if len(extensions) > 1 {
			return nil, fmt.Errorf("xlsxpatch: read sparklines: %s has multiple sparkline extensions", sheet.Part)
		}
		if len(extensions) == 0 {
			continue
		}
		var ext sparkXMLNode
		if err := xml.Unmarshal([]byte(extensions[0].xml), &ext); err != nil {
			return nil, fmt.Errorf("xlsxpatch: read sparklines: %s: %w", sheet.Part, err)
		}
		groupCollections := children(ext, "sparklineGroups")
		if len(groupCollections) != 1 || len(ext.Children) != 1 {
			return nil, fmt.Errorf("xlsxpatch: read sparklines: %s extension must contain exactly one sparklineGroups child", sheet.Part)
		}
		if names := unsupportedAttrNames(ext, map[string]bool{"uri": true}); len(names) != 0 {
			return nil, fmt.Errorf("xlsxpatch: read sparklines: %s extension has unsupported attributes %s", sheet.Part, strings.Join(names, ", "))
		}
		groups := groupCollections[0]
		if names := unsupportedAttrNames(*groups, nil); len(names) != 0 {
			return nil, fmt.Errorf("xlsxpatch: read sparklines: %s sparklineGroups has unsupported attributes %s", sheet.Part, strings.Join(names, ", "))
		}
		for _, nested := range groups.Children {
			if nested.XMLName.Local != "sparklineGroup" {
				return nil, fmt.Errorf("xlsxpatch: read sparklines: %s sparklineGroups contains unsupported child %s", sheet.Part, nested.XMLName.Local)
			}
		}
		for groupIndex := range groups.Children {
			group := &groups.Children[groupIndex]
			if group.XMLName.Local != "sparklineGroup" {
				continue
			}
			infos, err := parseSparklineGroup(*group, sheet, groupIndex)
			if err != nil {
				return nil, fmt.Errorf("xlsxpatch: read sparklines: %s group %d: %w", sheet.Part, groupIndex+1, err)
			}
			result = append(result, infos...)
		}
	}
	return result, nil
}

func parseSparklineGroup(group sparkXMLNode, sheet workbookSheet, groupIndex int) ([]SparklineInfo, error) {
	warnings := []string{}
	typeName := attr(group, "type")
	switch typeName {
	case "", "line":
		typeName = "line"
	case "column":
	case "stacked":
		typeName = "win-loss"
	default:
		warnings = append(warnings, fmt.Sprintf("unsupported sparkline type %q", typeName))
	}
	options := SparklineOptions{EmptyCells: "gap"}
	known := map[string]bool{"type": true, "displayEmptyCellsAs": true, "markers": true, "high": true, "low": true, "first": true, "last": true, "negative": true, "rightToLeft": true, "lineWeight": true, "manualMin": true, "manualMax": true, "minAxisType": true, "maxAxisType": true, "dateAxis": true, "displayHidden": true, "displayXAxis": true}
	for _, a := range group.Attrs {
		if a.Name.Space == "xmlns" || a.Name.Local == "xmlns" {
			continue
		}
		if !known[a.Name.Local] {
			warnings = append(warnings, fmt.Sprintf("unsupported group attribute %s", a.Name.Local))
		}
	}
	empty := attr(group, "displayEmptyCellsAs")
	switch empty {
	case "", "gap":
		options.EmptyCells = "gap"
	case "zero":
		options.EmptyCells = "zero"
	case "span":
		options.EmptyCells = "connect"
	default:
		warnings = append(warnings, fmt.Sprintf("unsupported empty-cell mode %q", empty))
	}
	boolAttrs := []struct {
		name string
		out  *bool
	}{{"rightToLeft", &options.RightToLeft}, {"markers", &options.ShowMarkers}, {"high", &options.ShowHigh}, {"low", &options.ShowLow}, {"first", &options.ShowFirst}, {"last", &options.ShowLast}, {"negative", &options.ShowNegative}}
	for _, entry := range boolAttrs {
		if raw := attr(group, entry.name); raw != "" {
			value, ok := parseOOXMLBool(raw)
			if !ok {
				warnings = append(warnings, fmt.Sprintf("invalid %s boolean %q", entry.name, raw))
			} else {
				*entry.out = value
			}
		}
	}
	for _, name := range []string{"dateAxis", "displayHidden", "displayXAxis"} {
		if raw := attr(group, name); raw != "" {
			value, ok := parseOOXMLBool(raw)
			if !ok {
				warnings = append(warnings, fmt.Sprintf("invalid %s boolean %q", name, raw))
			} else if value {
				warnings = append(warnings, fmt.Sprintf("%s is not representable", name))
			}
		}
	}
	if raw := attr(group, "lineWeight"); raw != "" {
		value, err := strconv.ParseFloat(raw, 64)
		if err != nil || !isFinitePositive(value) {
			warnings = append(warnings, fmt.Sprintf("invalid lineWeight %q", raw))
		} else {
			options.LineWeight = value
		}
	}
	minMode, maxMode := attr(group, "minAxisType"), attr(group, "maxAxisType")
	if minMode == "" {
		minMode = "individual"
	}
	if maxMode == "" {
		maxMode = "individual"
	}
	if minMode == "custom" {
		value, ok := finiteAttr(group, "manualMin")
		if !ok {
			warnings = append(warnings, "custom minimum axis has no finite manualMin")
		} else {
			options.Min = &value
		}
	} else if minMode != "individual" && minMode != "group" {
		warnings = append(warnings, fmt.Sprintf("unsupported minAxisType %q", minMode))
	}
	if maxMode == "custom" {
		value, ok := finiteAttr(group, "manualMax")
		if !ok {
			warnings = append(warnings, "custom maximum axis has no finite manualMax")
		} else {
			options.Max = &value
		}
	} else if maxMode != "individual" && maxMode != "group" {
		warnings = append(warnings, fmt.Sprintf("unsupported maxAxisType %q", maxMode))
	}
	if attr(group, "manualMin") != "" && minMode != "custom" {
		warnings = append(warnings, "manualMin without custom minimum axis is not representable")
	}
	if attr(group, "manualMax") != "" && maxMode != "custom" {
		warnings = append(warnings, "manualMax without custom maximum axis is not representable")
	}
	if (minMode == "group") != (maxMode == "group") && (options.Min == nil || options.Max == nil) {
		warnings = append(warnings, "mixed individual/group axis modes are not representable")
	}
	if options.Min != nil && options.Max != nil && *options.Min >= *options.Max {
		warnings = append(warnings, "manualMin must be less than manualMax")
	}
	colorMap := map[string]*string{"colorSeries": &options.Colors.Series, "colorNegative": &options.Colors.Negative, "colorMarkers": &options.Colors.Markers, "colorHigh": &options.Colors.High, "colorLow": &options.Colors.Low, "colorFirst": &options.Colors.First, "colorLast": &options.Colors.Last, "colorAxis": &options.Colors.Axis}
	var list *sparkXMLNode
	for index := range group.Children {
		part := &group.Children[index]
		if destination, ok := colorMap[part.XMLName.Local]; ok {
			color, warning := parseSparklineColor(*part)
			if warning != "" {
				warnings = append(warnings, warning)
			} else {
				*destination = color
			}
			continue
		}
		if part.XMLName.Local == "sparklines" {
			if list != nil {
				return nil, fmt.Errorf("multiple sparklines collections")
			}
			list = part
			continue
		}
		warnings = append(warnings, fmt.Sprintf("unsupported group child %s", part.XMLName.Local))
	}
	if list == nil {
		return nil, fmt.Errorf("missing sparklines collection")
	}
	if names := unsupportedAttrNames(*list, nil); len(names) != 0 {
		warnings = append(warnings, fmt.Sprintf("sparklines collection has unsupported attributes %s", strings.Join(names, ", ")))
	}
	for _, item := range list.Children {
		if item.XMLName.Local != "sparkline" {
			warnings = append(warnings, fmt.Sprintf("unsupported sparklines child %s", item.XMLName.Local))
		}
	}
	groupID := fmt.Sprintf("%s#sparkline-group-%d", sheet.Part, groupIndex+1)
	items := []SparklineInfo{}
	for itemIndex, item := range list.Children {
		if item.XMLName.Local != "sparkline" {
			continue
		}
		formulaNodes, targetNodes := children(item, "f"), children(item, "sqref")
		if len(formulaNodes) != 1 || len(targetNodes) != 1 || len(item.Children) != 2 {
			return nil, fmt.Errorf("sparkline %d is missing f or sqref", itemIndex+1)
		}
		formulaNode, targetNode := formulaNodes[0], targetNodes[0]
		itemWarnings := append([]string{}, warnings...)
		if names := unsupportedAttrNames(item, nil); len(names) != 0 {
			itemWarnings = append(itemWarnings, fmt.Sprintf("sparkline %d has unsupported attributes %s", itemIndex+1, strings.Join(names, ", ")))
		}
		if len(formulaNode.Children) != 0 || len(targetNode.Children) != 0 || len(unsupportedAttrNames(*formulaNode, nil)) != 0 || len(unsupportedAttrNames(*targetNode, nil)) != 0 {
			itemWarnings = append(itemWarnings, fmt.Sprintf("sparkline %d formula or target has unsupported markup", itemIndex+1))
		}
		sourceSheet, sourceRef, sourceWarning := parseSparklineFormula(strings.TrimSpace(formulaNode.Text), sheet.Name)
		targetRef := strings.TrimSpace(targetNode.Text)
		if sourceWarning != "" {
			itemWarnings = append(itemWarnings, sourceWarning)
		}
		if validateSingleCell(targetRef) != "" {
			itemWarnings = append(itemWarnings, fmt.Sprintf("unsupported target reference %q", targetRef))
		}
		if minMode == "group" || maxMode == "group" || options.Min != nil || options.Max != nil { /* retain native grouping */
		} else {
			groupID = ""
		}
		items = append(items, SparklineInfo{ID: fmt.Sprintf("%s#sparkline-%d-%d", sheet.Part, groupIndex+1, itemIndex+1), GroupID: groupID, Type: typeName, SourceSheetName: sourceSheet, SourceRef: sourceRef, TargetSheetName: sheet.Name, TargetRef: targetRef, Options: options, Warnings: itemWarnings})
	}
	if len(items) == 0 {
		return nil, fmt.Errorf("sparkline group is empty")
	}
	return items, nil
}

// SetSparklines atomically replaces the workbook's supported x14 sparkline
// metadata. It refuses to overwrite pre-existing unsupported state. Only
// worksheets whose extension changes are patched; every other OPC entry is
// raw-copied and verified byte-identical by Apply.
func SetSparklines(orig []byte, groups []SparklineWriteGroup) ([]byte, error) {
	existing, err := ReadSparklines(orig)
	if err != nil {
		return nil, err
	}
	for _, item := range existing {
		if len(item.Warnings) > 0 {
			return nil, fmt.Errorf("xlsxpatch: set sparklines refused: %s: %s", item.ID, strings.Join(item.Warnings, "; "))
		}
	}
	if err := validateSparklineWriteGroups(groups); err != nil {
		return nil, err
	}
	zr, err := zip.NewReader(bytes.NewReader(orig), int64(len(orig)))
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: set sparklines: %w", err)
	}
	files := map[string]*zip.File{}
	for _, file := range zr.File {
		files[file.Name] = file
	}
	read := func(name string) (string, bool) {
		file, ok := files[name]
		if !ok {
			return "", false
		}
		rc, e := file.Open()
		if e != nil {
			return "", false
		}
		defer rc.Close()
		b, e := io.ReadAll(rc)
		return string(b), e == nil
	}
	sheets, err := readWorkbookSheets(read)
	if err != nil {
		return nil, err
	}
	byName := map[string]workbookSheet{}
	for _, sheet := range sheets {
		byName[sheet.Name] = sheet
	}
	grouped := map[string][]SparklineWriteGroup{}
	for _, group := range groups {
		if _, ok := byName[group.TargetSheetName]; !ok {
			return nil, fmt.Errorf("xlsxpatch: set sparklines: target sheet %q not found", group.TargetSheetName)
		}
		for _, item := range group.Sparklines {
			if _, ok := byName[item.SourceSheetName]; !ok {
				return nil, fmt.Errorf("xlsxpatch: set sparklines: source sheet %q not found", item.SourceSheetName)
			}
		}
		grouped[group.TargetSheetName] = append(grouped[group.TargetSheetName], group)
	}
	patch := Patch{Replace: map[string][]byte{}}
	for _, sheet := range sheets {
		value, ok := read(sheet.Part)
		if !ok {
			return nil, fmt.Errorf("xlsxpatch: set sparklines: missing worksheet %q", sheet.Part)
		}
		updated, changed, err := replaceSparklineExtension(value, grouped[sheet.Name])
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: set sparklines: %s: %w", sheet.Part, err)
		}
		if changed {
			patch.Replace[sheet.Part] = []byte(updated)
		}
	}
	return Apply(orig, patch)
}

type sparkExtension struct {
	start, end int
	xml        string
}

func sparklineExtensions(value string) ([]sparkExtension, error) {
	decoder := xml.NewDecoder(strings.NewReader(value))
	out := []sparkExtension{}
	for {
		tok, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}
		start, ok := tok.(xml.StartElement)
		if !ok || start.Name.Local != "ext" || !strings.EqualFold(attrElement(start, "uri"), sparklineExtensionURI) {
			continue
		}
		after := int(decoder.InputOffset())
		begin := strings.LastIndex(value[:after], "<")
		if begin < 0 {
			return nil, fmt.Errorf("cannot locate sparkline extension")
		}
		depth := 1
		for depth > 0 {
			inner, e := decoder.Token()
			if e != nil {
				return nil, e
			}
			switch inner.(type) {
			case xml.StartElement:
				depth++
			case xml.EndElement:
				depth--
			}
		}
		end := int(decoder.InputOffset())
		out = append(out, sparkExtension{begin, end, value[begin:end]})
	}
	return out, nil
}

func replaceSparklineExtension(value string, groups []SparklineWriteGroup) (string, bool, error) {
	exts, err := sparklineExtensions(value)
	if err != nil {
		return "", false, err
	}
	if len(exts) > 1 {
		return "", false, fmt.Errorf("multiple sparkline extensions")
	}
	newExt := ""
	if len(groups) > 0 {
		newExt = marshalSparklineExtension(groups)
	}
	if len(exts) == 1 {
		if exts[0].xml == newExt {
			return value, false, nil
		}
		value = value[:exts[0].start] + newExt + value[exts[0].end:]
		if newExt == "" {
			empty := regexp.MustCompile(`<extLst\b[^>]*>\s*</extLst>`)
			value = empty.ReplaceAllString(value, "")
		}
		return value, true, nil
	}
	if newExt == "" {
		return value, false, nil
	}
	if match := regexp.MustCompile(`<extLst\b[^>]*/>`).FindStringIndex(value); match != nil {
		replacement := `<extLst>` + newExt + `</extLst>`
		return value[:match[0]] + replacement + value[match[1]:], true, nil
	}
	if closeAt := strings.LastIndex(value, "</extLst>"); closeAt >= 0 {
		return value[:closeAt] + newExt + value[closeAt:], true, nil
	}
	closeAt := strings.LastIndex(value, "</worksheet>")
	if closeAt < 0 {
		return "", false, fmt.Errorf("malformed worksheet")
	}
	return value[:closeAt] + `<extLst>` + newExt + `</extLst>` + value[closeAt:], true, nil
}

func marshalSparklineExtension(groups []SparklineWriteGroup) string {
	var b strings.Builder
	b.WriteString(`<ext uri="` + sparklineExtensionURI + `" xmlns:x14="` + x14Namespace + `" xmlns:xm="` + xmNamespace + `"><x14:sparklineGroups>`)
	for _, group := range groups {
		b.WriteString(`<x14:sparklineGroup`)
		writeAttr(&b, "type", map[string]string{"line": "line", "column": "column", "win-loss": "stacked"}[group.Type])
		if group.Options.EmptyCells != "" && group.Options.EmptyCells != "gap" {
			writeAttr(&b, "displayEmptyCellsAs", map[string]string{"zero": "zero", "connect": "span"}[group.Options.EmptyCells])
		}
		boolAttrs := []struct {
			name  string
			value bool
		}{{"markers", group.Options.ShowMarkers}, {"high", group.Options.ShowHigh}, {"low", group.Options.ShowLow}, {"first", group.Options.ShowFirst}, {"last", group.Options.ShowLast}, {"negative", group.Options.ShowNegative}, {"rightToLeft", group.Options.RightToLeft}}
		for _, a := range boolAttrs {
			if a.value {
				writeAttr(&b, a.name, "1")
			}
		}
		if group.Options.LineWeight != 0 {
			writeAttr(&b, "lineWeight", strconv.FormatFloat(group.Options.LineWeight, 'g', -1, 64))
		}
		mode := "individual"
		if len(group.Sparklines) > 1 {
			mode = "group"
		}
		if group.Options.Min != nil {
			writeAttr(&b, "minAxisType", "custom")
			writeAttr(&b, "manualMin", strconv.FormatFloat(*group.Options.Min, 'g', -1, 64))
		} else {
			writeAttr(&b, "minAxisType", mode)
		}
		if group.Options.Max != nil {
			writeAttr(&b, "maxAxisType", "custom")
			writeAttr(&b, "manualMax", strconv.FormatFloat(*group.Options.Max, 'g', -1, 64))
		} else {
			writeAttr(&b, "maxAxisType", mode)
		}
		b.WriteByte('>')
		colors := []struct{ name, value string }{{"colorSeries", group.Options.Colors.Series}, {"colorNegative", group.Options.Colors.Negative}, {"colorMarkers", group.Options.Colors.Markers}, {"colorHigh", group.Options.Colors.High}, {"colorLow", group.Options.Colors.Low}, {"colorFirst", group.Options.Colors.First}, {"colorLast", group.Options.Colors.Last}, {"colorAxis", group.Options.Colors.Axis}}
		for _, color := range colors {
			if color.value != "" {
				b.WriteString(`<x14:` + color.name + ` rgb="` + normalizeARGB(color.value) + `"/>`)
			}
		}
		b.WriteString(`<x14:sparklines>`)
		for _, item := range group.Sparklines {
			b.WriteString(`<x14:sparkline><xm:f>` + esc(formatSparklineFormula(item.SourceSheetName, item.SourceRef)) + `</xm:f><xm:sqref>` + esc(item.TargetCellRef) + `</xm:sqref></x14:sparkline>`)
		}
		b.WriteString(`</x14:sparklines></x14:sparklineGroup>`)
	}
	b.WriteString(`</x14:sparklineGroups></ext>`)
	return b.String()
}

var rangeRefRE = regexp.MustCompile(`^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$`)
var singleCellRefRE = regexp.MustCompile(`^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)$`)
var colorRE = regexp.MustCompile(`^#?([0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$`)

func validateSparklineWriteGroups(groups []SparklineWriteGroup) error {
	targets := map[string]bool{}
	ids := map[string]bool{}
	for gi, g := range groups {
		label := fmt.Sprintf("group %d", gi+1)
		if g.ID != "" {
			if ids[g.ID] {
				return fmt.Errorf("xlsxpatch: set sparklines: duplicate group id %q", g.ID)
			}
			ids[g.ID] = true
			label = "group " + g.ID
		}
		if g.TargetSheetName == "" {
			return fmt.Errorf("xlsxpatch: set sparklines: %s has empty target sheet", label)
		}
		if g.Type != "line" && g.Type != "column" && g.Type != "win-loss" {
			return fmt.Errorf("xlsxpatch: set sparklines: %s has unsupported type %q", label, g.Type)
		}
		if len(g.Sparklines) == 0 {
			return fmt.Errorf("xlsxpatch: set sparklines: %s is empty", label)
		}
		if g.Options.EmptyCells != "" && g.Options.EmptyCells != "gap" && g.Options.EmptyCells != "zero" && g.Options.EmptyCells != "connect" {
			return fmt.Errorf("xlsxpatch: set sparklines: %s has unsupported empty-cell mode %q", label, g.Options.EmptyCells)
		}
		if g.Options.LineWeight < 0 || math.IsNaN(g.Options.LineWeight) || math.IsInf(g.Options.LineWeight, 0) {
			return fmt.Errorf("xlsxpatch: set sparklines: %s has invalid line weight", label)
		}
		if g.Options.Min != nil && (!isFinite(*g.Options.Min)) {
			return fmt.Errorf("xlsxpatch: set sparklines: %s has invalid minimum", label)
		}
		if g.Options.Max != nil && (!isFinite(*g.Options.Max)) {
			return fmt.Errorf("xlsxpatch: set sparklines: %s has invalid maximum", label)
		}
		if g.Options.Min != nil && g.Options.Max != nil && *g.Options.Min >= *g.Options.Max {
			return fmt.Errorf("xlsxpatch: set sparklines: %s minimum must be less than maximum", label)
		}
		for _, color := range []string{g.Options.Colors.Series, g.Options.Colors.Negative, g.Options.Colors.Markers, g.Options.Colors.High, g.Options.Colors.Low, g.Options.Colors.First, g.Options.Colors.Last, g.Options.Colors.Axis} {
			if color != "" && !validNativeRGB(color) {
				return fmt.Errorf("xlsxpatch: set sparklines: %s has unsupported color %q", label, color)
			}
		}
		for ii, item := range g.Sparklines {
			if item.SourceSheetName == "" {
				return fmt.Errorf("xlsxpatch: set sparklines: %s item %d has empty source sheet", label, ii+1)
			}
			if warning := validateOneDimensionalRange(item.SourceRef); warning != "" {
				return fmt.Errorf("xlsxpatch: set sparklines: %s item %d: %s", label, ii+1, warning)
			}
			if warning := validateSingleCell(item.TargetCellRef); warning != "" {
				return fmt.Errorf("xlsxpatch: set sparklines: %s item %d has unsupported target reference %q", label, ii+1, item.TargetCellRef)
			}
			if item.SourceSheetName == g.TargetSheetName && rangeContainsCell(item.SourceRef, item.TargetCellRef) {
				return fmt.Errorf("xlsxpatch: set sparklines: %s item %d target overlaps its source", label, ii+1)
			}
			key := g.TargetSheetName + "\x00" + strings.ToUpper(strings.ReplaceAll(item.TargetCellRef, "$", ""))
			if targets[key] {
				return fmt.Errorf("xlsxpatch: set sparklines: duplicate target %s!%s", g.TargetSheetName, item.TargetCellRef)
			}
			targets[key] = true
		}
	}
	return nil
}

func validateOneDimensionalRange(value string) string {
	m := rangeRefRE.FindStringSubmatch(strings.TrimSpace(value))
	if m == nil {
		return fmt.Sprintf("unsupported source reference %q", value)
	}
	r1, _ := strconv.Atoi(m[2])
	r2 := r1
	if m[4] != "" {
		r2, _ = strconv.Atoi(m[4])
	}
	if r1 < 1 || r1 > 1048576 || r2 > 1048576 {
		return fmt.Sprintf("source reference %q exceeds Excel row bounds", value)
	}
	c1 := columnNumber(m[1])
	c2 := c1
	if m[3] != "" {
		c2 = columnNumber(m[3])
	}
	if c1 < 1 || c1 > 16384 || c2 > 16384 {
		return fmt.Sprintf("source reference %q exceeds Excel column bounds", value)
	}
	if r2 < r1 || c2 < c1 {
		return fmt.Sprintf("reversed source reference %q", value)
	}
	if r2 > r1 && c2 > c1 {
		return fmt.Sprintf("two-dimensional source reference %q", value)
	}
	return ""
}

func parseSparklineFormula(value, defaultSheet string) (string, string, string) {
	if strings.ContainsAny(value, "[]#") {
		return "", "", fmt.Sprintf("external or error source reference %q is unsupported", value)
	}
	bang := strings.LastIndex(value, "!")
	sheet, ref := defaultSheet, value
	if bang >= 0 {
		sheet, ref = value[:bang], value[bang+1:]
		if strings.HasPrefix(sheet, "'") {
			if !strings.HasSuffix(sheet, "'") || len(sheet) < 2 {
				return "", "", fmt.Sprintf("malformed quoted sheet in %q", value)
			}
			sheet = strings.ReplaceAll(sheet[1:len(sheet)-1], "''", "'")
		} else if strings.Contains(sheet, "'") {
			return "", "", fmt.Sprintf("malformed sheet in %q", value)
		}
	}
	if sheet == "" {
		return "", "", fmt.Sprintf("empty source sheet in %q", value)
	}
	if warning := validateOneDimensionalRange(ref); warning != "" {
		return sheet, "", warning
	}
	return sheet, ref, ""
}

func formatSparklineFormula(sheet, ref string) string {
	sheet = "'" + strings.ReplaceAll(sheet, "'", "''") + "'"
	return sheet + "!" + ref
}
func parseSparklineColor(node sparkXMLNode) (string, string) {
	rgb := attr(node, "rgb")
	if rgb == "" || !colorRE.MatchString(rgb) {
		return "", fmt.Sprintf("%s uses an unsupported non-RGB color", node.XMLName.Local)
	}
	for _, a := range node.Attrs {
		if a.Name.Space == "xmlns" || a.Name.Local == "xmlns" || a.Name.Local == "rgb" {
			continue
		}
		return "", fmt.Sprintf("%s uses unsupported color attribute %s", node.XMLName.Local, a.Name.Local)
	}
	if len(rgb) == 8 && !strings.EqualFold(rgb[:2], "FF") {
		return "", fmt.Sprintf("%s uses unsupported alpha %s", node.XMLName.Local, rgb[:2])
	}
	if len(rgb) == 8 {
		rgb = rgb[2:]
	}
	return "#" + strings.ToUpper(rgb), ""
}
func normalizeARGB(value string) string {
	value = strings.TrimPrefix(value, "#")
	value = strings.ToUpper(value)
	if len(value) == 6 {
		return "FF" + value
	}
	return value
}
func validNativeRGB(value string) bool {
	if !colorRE.MatchString(value) {
		return false
	}
	value = strings.TrimPrefix(value, "#")
	return len(value) == 6 || strings.EqualFold(value[:2], "FF")
}
func attr(node sparkXMLNode, name string) string {
	for _, a := range node.Attrs {
		if a.Name.Local == name {
			return a.Value
		}
	}
	return ""
}
func attrElement(node xml.StartElement, name string) string {
	for _, a := range node.Attr {
		if a.Name.Local == name {
			return a.Value
		}
	}
	return ""
}
func child(node sparkXMLNode, name string) *sparkXMLNode {
	for i := range node.Children {
		if node.Children[i].XMLName.Local == name {
			return &node.Children[i]
		}
	}
	return nil
}
func children(node sparkXMLNode, name string) []*sparkXMLNode {
	var result []*sparkXMLNode
	for i := range node.Children {
		if node.Children[i].XMLName.Local == name {
			result = append(result, &node.Children[i])
		}
	}
	return result
}
func unsupportedAttrNames(node sparkXMLNode, known map[string]bool) []string {
	var result []string
	for _, attribute := range node.Attrs {
		if attribute.Name.Space == "xmlns" || attribute.Name.Local == "xmlns" || known != nil && known[attribute.Name.Local] {
			continue
		}
		result = append(result, attribute.Name.Local)
	}
	return result
}
func validateSingleCell(value string) string {
	match := singleCellRefRE.FindStringSubmatch(strings.TrimSpace(value))
	if match == nil {
		return "unsupported cell reference"
	}
	row, _ := strconv.Atoi(match[2])
	if row < 1 || row > 1048576 || columnNumber(match[1]) > 16384 {
		return "cell reference exceeds Excel bounds"
	}
	return ""
}
func rangeContainsCell(rangeValue, cellValue string) bool {
	rangeMatch, cellMatch := rangeRefRE.FindStringSubmatch(strings.TrimSpace(rangeValue)), singleCellRefRE.FindStringSubmatch(strings.TrimSpace(cellValue))
	if rangeMatch == nil || cellMatch == nil {
		return false
	}
	r1, _ := strconv.Atoi(rangeMatch[2])
	r2 := r1
	if rangeMatch[4] != "" {
		r2, _ = strconv.Atoi(rangeMatch[4])
	}
	c1, c2, row, col := columnNumber(rangeMatch[1]), columnNumber(rangeMatch[1]), 0, columnNumber(cellMatch[1])
	if rangeMatch[3] != "" {
		c2 = columnNumber(rangeMatch[3])
	}
	row, _ = strconv.Atoi(cellMatch[2])
	return row >= r1 && row <= r2 && col >= c1 && col <= c2
}
func parseOOXMLBool(raw string) (bool, bool) {
	switch raw {
	case "1", "true", "on":
		return true, true
	case "0", "false", "off":
		return false, true
	default:
		return false, false
	}
}
func finiteAttr(node sparkXMLNode, name string) (float64, bool) {
	raw := attr(node, name)
	value, err := strconv.ParseFloat(raw, 64)
	return value, err == nil && isFinite(value)
}
func isFinite(value float64) bool         { return !math.IsNaN(value) && !math.IsInf(value, 0) }
func isFinitePositive(value float64) bool { return isFinite(value) && value > 0 }
func columnNumber(value string) int {
	n := 0
	for _, r := range strings.ToUpper(value) {
		n = n*26 + int(r-'A'+1)
	}
	return n
}
func writeAttr(b *strings.Builder, name, value string) { fmt.Fprintf(b, " %s=%q", name, value) }
