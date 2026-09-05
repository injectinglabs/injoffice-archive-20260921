package xlsxpatch

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"math"
	"sort"
	"strconv"
	"strings"
)

type styleFont struct {
	raw, rootStart []byte
	qname          string
	name           *string
	size           *float64
	bold, italic   bool
	color          *string
	colorSafe      bool
	themeIndex     *int
	themeTint      *float64
	scheme         *string
	children       []styleFontChild
	key            string
}

type styleFontChild struct {
	local string
	raw   []byte
}

type styleFill struct {
	raw, rootStart []byte
	qname          string
	color          *string
	supported      bool
	background64   bool
	key            string
}

type styleBorderSide struct {
	style string
	color string
}

type styleBorder struct {
	raw       []byte
	supported bool
	left      *styleBorderSide
	right     *styleBorderSide
	top       *styleBorderSide
	bottom    *styleBorderSide
}

type styleAlignmentRecord struct {
	present                    bool
	raw                        []byte
	rawStart                   []byte
	qname                      string
	horizontal                 *string
	vertical                   *string
	wrap                       *bool
	shrinkToFit                *bool
	textRotation               *int
	otherAttrsKey              string
	otherAttrsExceptPaintExact string
}

type styleXF struct {
	raw, rootStart    []byte
	body              []byte
	qname             string
	selfClosing       bool
	alignmentStart    int
	alignmentEnd      int
	numFmtID          int
	fontID            int
	fillID            int
	borderID          int
	xfID              int
	applyNumberFormat *bool
	applyFont         *bool
	applyFill         *bool
	applyBorder       *bool
	applyAlignment    *bool
	alignment         styleAlignmentRecord
}

type styleRegistry struct {
	data  []byte
	index styleTableIndex

	numFmtByID   map[int]string
	numFmtByCode map[string]int
	usedNumFmtID map[int]bool
	fonts        []styleFont
	fontByKey    map[string]int
	fills        []styleFill
	fillByKey    map[string]int
	borders      []styleBorder
	styleXfs     []styleXF
	cellXfs      []styleXF
	xfByRaw      map[string]int

	appendedNumFmts [][]byte
	appendedFonts   [][]byte
	appendedFills   [][]byte
	appendedCellXfs [][]byte
}

func newStyleRegistry(data []byte) (*styleRegistry, error) {
	index, err := parseStyleTable(data)
	if err != nil {
		return nil, err
	}
	registry := &styleRegistry{
		data: data, index: index,
		numFmtByID: make(map[int]string), numFmtByCode: make(map[string]int), usedNumFmtID: make(map[int]bool),
		fontByKey: make(map[string]int), fillByKey: make(map[string]int), xfByRaw: make(map[string]int),
	}
	if index.numFmts != nil {
		for _, entry := range index.numFmts.entries {
			if err := requireEmptyStyleEntry(data, entry); err != nil {
				return nil, fmt.Errorf("numFmt: %w", err)
			}
			id, err := requiredUnsignedStyleAttribute(entry.start, "numFmtId")
			if err != nil {
				return nil, fmt.Errorf("numFmt: %w", err)
			}
			formatCode, found, err := unqualifiedXMLAttribute(entry.start, "formatCode")
			if err != nil {
				return nil, fmt.Errorf("numFmt %d: %w", id, err)
			}
			if !found || formatCode == "" {
				return nil, fmt.Errorf("numFmt %d requires formatCode", id)
			}
			if _, duplicate := registry.numFmtByID[id]; duplicate {
				return nil, fmt.Errorf("duplicate numFmtId %d", id)
			}
			registry.numFmtByID[id] = formatCode
			registry.usedNumFmtID[id] = true
			if prior, found := registry.numFmtByCode[formatCode]; !found || id < prior {
				registry.numFmtByCode[formatCode] = id
			}
		}
	}
	for id, format := range builtinNumberFormats {
		registry.usedNumFmtID[id] = true
		if _, explicitlyDeclared := registry.numFmtByID[id]; explicitlyDeclared {
			continue
		}
		if prior, found := registry.numFmtByCode[format]; !found || id < prior {
			registry.numFmtByCode[format] = id
		}
	}
	for position, entry := range index.fonts.entries {
		font, err := parseStyleFont(data, entry, index.namespace)
		if err != nil {
			return nil, fmt.Errorf("font %d: %w", position, err)
		}
		registry.fonts = append(registry.fonts, font)
		if _, found := registry.fontByKey[font.key]; !found {
			registry.fontByKey[font.key] = position
		}
	}
	for position, entry := range index.fills.entries {
		fill, err := parseStyleFill(data, entry, index.namespace)
		if err != nil {
			return nil, fmt.Errorf("fill %d: %w", position, err)
		}
		registry.fills = append(registry.fills, fill)
		if fill.supported {
			if _, found := registry.fillByKey[fill.key]; !found {
				registry.fillByKey[fill.key] = position
			}
		}
	}
	for position, entry := range index.borders.entries {
		border, err := parseStyleBorder(data, entry, index.namespace)
		if err != nil {
			return nil, fmt.Errorf("border %d: %w", position, err)
		}
		registry.borders = append(registry.borders, border)
	}
	for position, entry := range index.cellStyleXfs.entries {
		xf, err := parseStyleXF(data, entry, index.namespace, false)
		if err != nil {
			return nil, fmt.Errorf("cellStyleXf %d: %w", position, err)
		}
		registry.styleXfs = append(registry.styleXfs, xf)
	}
	for position, entry := range index.cellXfs.entries {
		xf, err := parseStyleXF(data, entry, index.namespace, true)
		if err != nil {
			return nil, fmt.Errorf("cellXf %d: %w", position, err)
		}
		registry.cellXfs = append(registry.cellXfs, xf)
		registry.xfByRaw[string(xf.raw)] = position
	}
	for position := range registry.styleXfs {
		if err := registry.validateXFReferences(registry.styleXfs[position], false); err != nil {
			return nil, fmt.Errorf("cellStyleXf %d: %w", position, err)
		}
	}
	for position := range registry.cellXfs {
		if err := registry.validateXFReferences(registry.cellXfs[position], true); err != nil {
			return nil, fmt.Errorf("cellXf %d: %w", position, err)
		}
	}
	return registry, nil
}

var builtinNumberFormats = map[int]string{
	0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00",
	9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??",
	14: "mm-dd-yy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy",
	18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss",
	22: "m/d/yy h:mm", 37: "#,##0 ;(#,##0)", 38: "#,##0 ;[Red](#,##0)",
	39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)",
	45: "mm:ss", 46: "[h]:mm:ss", 47: "mmss.0", 48: "##0.0E+0", 49: "@",
}

func requiredUnsignedStyleAttribute(start xml.StartElement, name string) (int, error) {
	raw, found, err := unqualifiedXMLAttribute(start, name)
	if err != nil {
		return 0, err
	}
	if !found || raw == "" {
		return 0, fmt.Errorf("requires %s", name)
	}
	for _, character := range raw {
		if character < '0' || character > '9' {
			return 0, fmt.Errorf("%s=%q is not an unsigned integer", name, raw)
		}
	}
	value, err := strconv.Atoi(raw)
	if err != nil {
		return 0, fmt.Errorf("%s=%q is outside platform limits", name, raw)
	}
	return value, nil
}

func optionalUnsignedStyleAttribute(start xml.StartElement, name string, fallback int) (int, error) {
	raw, found, err := unqualifiedXMLAttribute(start, name)
	if err != nil {
		return 0, err
	}
	if !found {
		return fallback, nil
	}
	if raw == "" {
		return 0, fmt.Errorf("%s is empty", name)
	}
	for _, character := range raw {
		if character < '0' || character > '9' {
			return 0, fmt.Errorf("%s=%q is not an unsigned integer", name, raw)
		}
	}
	value, err := strconv.Atoi(raw)
	if err != nil {
		return 0, fmt.Errorf("%s=%q is outside platform limits", name, raw)
	}
	return value, nil
}

func parseStyleFont(data []byte, entry styleTableEntry, namespace string) (styleFont, error) {
	if err := requireWhitespaceAroundStyleChildren(data, entry.span, entry.children); err != nil {
		return styleFont{}, err
	}
	font := styleFont{
		raw:       bytes.Clone(data[entry.span.start:entry.span.end]),
		rootStart: bytes.Clone(data[entry.span.start:entry.span.startTagEnd]),
		qname:     entry.span.qname, colorSafe: true,
	}
	seen := make(map[string]bool)
	for _, child := range entry.children {
		raw := bytes.Clone(data[child.span.start:child.span.end])
		font.children = append(font.children, styleFontChild{raw: raw})
		if child.start.Name.Space != namespace {
			continue
		}
		local := child.start.Name.Local
		font.children[len(font.children)-1].local = local
		switch local {
		case "name", "sz", "b", "i", "color", "scheme":
			if seen[local] {
				return styleFont{}, fmt.Errorf("duplicate %s child", local)
			}
			seen[local] = true
			if err := requireEmptyStyleNode(data, child); err != nil {
				return styleFont{}, fmt.Errorf("%s: %w", local, err)
			}
		default:
			continue
		}
		switch local {
		case "name":
			value, err := styleNodeValue(child.start, "val")
			if err != nil || value == "" {
				return styleFont{}, fmt.Errorf("name requires a non-empty val")
			}
			font.name = stringPointer(value)
		case "sz":
			value, err := styleNodeValue(child.start, "val")
			if err != nil {
				return styleFont{}, fmt.Errorf("sz: %w", err)
			}
			parsed, parseErr := strconv.ParseFloat(value, 64)
			if parseErr != nil || math.IsNaN(parsed) || math.IsInf(parsed, 0) || parsed <= 0 {
				return styleFont{}, fmt.Errorf("sz val=%q is invalid", value)
			}
			font.size = floatPointer(parsed)
		case "b", "i":
			value, found, err := unqualifiedXMLAttribute(child.start, "val")
			if err != nil {
				return styleFont{}, err
			}
			enabled := true
			if found {
				enabled, err = ooxmlBoolean(value, true)
				if err != nil {
					return styleFont{}, fmt.Errorf("%s val=%q is invalid", local, value)
				}
			}
			if local == "b" {
				font.bold = enabled
			} else {
				font.italic = enabled
			}
		case "color":
			color, safe := supportedRGBStyleColor(child.start)
			font.color, font.colorSafe = color, safe
			if !safe {
				font.themeIndex, font.themeTint = exactThemeColor(child.start)
			}
		case "scheme":
			value, err := styleNodeValue(child.start, "val")
			if err == nil && (value == "major" || value == "minor") {
				font.scheme = stringPointer(value)
			}
		}
	}
	font.key = fontSemanticKey(font)
	return font, nil
}

func parseStyleFill(data []byte, entry styleTableEntry, namespace string) (styleFill, error) {
	if err := requireWhitespaceAroundStyleChildren(data, entry.span, entry.children); err != nil {
		return styleFill{}, err
	}
	fill := styleFill{raw: bytes.Clone(data[entry.span.start:entry.span.end]), rootStart: bytes.Clone(data[entry.span.start:entry.span.startTagEnd]), qname: entry.span.qname}
	if len(entry.children) != 1 {
		fill.key = "unsupported:" + string(fill.raw)
		return fill, nil
	}
	pattern := entry.children[0]
	if err := requireWhitespaceAroundStyleChildren(data, pattern.span, pattern.children); err != nil {
		return styleFill{}, fmt.Errorf("patternFill: %w", err)
	}
	if pattern.start.Name.Space != namespace || pattern.start.Name.Local != "patternFill" {
		fill.key = "unsupported:" + string(fill.raw)
		return fill, nil
	}
	patternType, found, err := unqualifiedXMLAttribute(pattern.start, "patternType")
	if err != nil {
		return styleFill{}, err
	}
	if !found || (patternType != "none" && patternType != "solid") {
		fill.key = "unsupported:" + string(fill.raw)
		return fill, nil
	}
	if !styleAttributesOnly(pattern.start, "patternType") {
		fill.key = "unsupported:" + string(fill.raw)
		return fill, nil
	}
	if patternType == "none" {
		if len(pattern.children) != 0 {
			fill.key = "unsupported:" + string(fill.raw)
			return fill, nil
		}
		fill.supported, fill.key = true, fillSemanticKey(fill.rootStart, nil)
		return fill, nil
	}
	if len(pattern.children) < 1 || len(pattern.children) > 2 || pattern.children[0].start.Name.Space != namespace || pattern.children[0].start.Name.Local != "fgColor" {
		fill.key = "unsupported:" + string(fill.raw)
		return fill, nil
	}
	if err := requireEmptyStyleNode(data, pattern.children[0]); err != nil {
		return styleFill{}, err
	}
	color, safe := supportedRGBStyleColor(pattern.children[0].start)
	if !safe || color == nil {
		fill.key = "unsupported:" + string(fill.raw)
		return fill, nil
	}
	if len(pattern.children) == 2 {
		background := pattern.children[1]
		if background.start.Name.Space != namespace || background.start.Name.Local != "bgColor" {
			fill.key = "unsupported:" + string(fill.raw)
			return fill, nil
		}
		if err := requireEmptyStyleNode(data, background); err != nil {
			return styleFill{}, err
		}
		indexed, indexedFound, err := unqualifiedXMLAttribute(background.start, "indexed")
		if err != nil {
			return styleFill{}, err
		}
		backgroundColor, backgroundSafe := supportedRGBStyleColor(background.start)
		indexed64 := indexedFound && indexed == "64" && styleAttributesOnly(background.start, "indexed")
		sameOpaqueRGB := backgroundSafe && backgroundColor != nil && *backgroundColor == *color
		if !indexed64 && !sameOpaqueRGB {
			fill.key = "unsupported:" + string(fill.raw)
			return fill, nil
		}
		fill.background64 = indexed64
	}
	fill.color, fill.supported = color, true
	fill.key = fillSemanticKey(fill.rootStart, color)
	return fill, nil
}

var supportedBorderStyleTokens = map[string]bool{
	"dashDot": true, "dashDotDot": true, "dashed": true, "dotted": true, "double": true, "hair": true,
	"medium": true, "mediumDashDot": true, "mediumDashDotDot": true, "mediumDashed": true,
	"slantDashDot": true, "thick": true, "thin": true,
}

func parseStyleBorder(data []byte, entry styleTableEntry, namespace string) (styleBorder, error) {
	if err := requireWhitespaceAroundStyleChildren(data, entry.span, entry.children); err != nil {
		return styleBorder{}, err
	}
	border := styleBorder{raw: bytes.Clone(data[entry.span.start:entry.span.end]), supported: true}
	if !styleAttributesOnly(entry.start) {
		border.supported = false
		return border, nil
	}
	seen := make(map[string]bool)
	for _, child := range entry.children {
		if child.start.Name.Space != namespace {
			border.supported = false
			continue
		}
		local := child.start.Name.Local
		if seen[local] {
			return styleBorder{}, fmt.Errorf("duplicate %s child", local)
		}
		seen[local] = true
		switch local {
		case "left", "right", "top", "bottom":
			side, supported, err := parseStyleBorderSide(data, child, namespace)
			if err != nil {
				return styleBorder{}, fmt.Errorf("%s: %w", local, err)
			}
			if !supported {
				border.supported = false
				continue
			}
			switch local {
			case "left":
				border.left = side
			case "right":
				border.right = side
			case "top":
				border.top = side
			case "bottom":
				border.bottom = side
			}
		case "diagonal":
			if err := requireEmptyStyleNode(data, child); err != nil || !styleAttributesOnly(child.start) {
				border.supported = false
			}
		default:
			border.supported = false
		}
	}
	return border, nil
}

func parseStyleBorderSide(data []byte, node styleTableNode, namespace string) (*styleBorderSide, bool, error) {
	if err := requireWhitespaceAroundStyleChildren(data, node.span, node.children); err != nil {
		return nil, false, err
	}
	style, found, err := unqualifiedXMLAttribute(node.start, "style")
	if err != nil {
		return nil, false, err
	}
	if !styleAttributesOnly(node.start, "style") {
		return nil, false, nil
	}
	if !found {
		if err := requireEmptyStyleNode(data, node); err != nil {
			return nil, false, nil
		}
		return nil, true, nil
	}
	if !supportedBorderStyleTokens[style] || len(node.children) != 1 {
		return nil, false, nil
	}
	colorNode := node.children[0]
	if colorNode.start.Name != (xml.Name{Space: namespace, Local: "color"}) {
		return nil, false, nil
	}
	if err := requireEmptyStyleNode(data, colorNode); err != nil {
		return nil, false, err
	}
	color, safe := supportedRGBStyleColor(colorNode.start)
	if !safe || color == nil {
		return nil, false, nil
	}
	return &styleBorderSide{style: style, color: *color}, true, nil
}

func parseStyleXF(data []byte, entry styleTableEntry, namespace string, cellXF bool) (styleXF, error) {
	if err := requireWhitespaceAroundStyleChildren(data, entry.span, entry.children); err != nil {
		return styleXF{}, err
	}
	raw := bytes.Clone(data[entry.span.start:entry.span.end])
	xf := styleXF{
		raw: raw, rootStart: bytes.Clone(data[entry.span.start:entry.span.startTagEnd]),
		body: bytes.Clone(data[entry.span.startTagEnd:entry.span.endStart]), qname: entry.span.qname,
		selfClosing: bytes.HasSuffix(bytes.TrimSpace(raw), []byte("/>")), alignmentStart: -1, alignmentEnd: -1,
	}
	var err error
	if xf.numFmtID, err = optionalUnsignedStyleAttribute(entry.start, "numFmtId", 0); err != nil {
		return styleXF{}, err
	}
	if xf.fontID, err = optionalUnsignedStyleAttribute(entry.start, "fontId", 0); err != nil {
		return styleXF{}, err
	}
	if xf.fillID, err = optionalUnsignedStyleAttribute(entry.start, "fillId", 0); err != nil {
		return styleXF{}, err
	}
	if xf.borderID, err = optionalUnsignedStyleAttribute(entry.start, "borderId", 0); err != nil {
		return styleXF{}, err
	}
	if cellXF {
		if xf.xfID, err = optionalUnsignedStyleAttribute(entry.start, "xfId", 0); err != nil {
			return styleXF{}, err
		}
	}
	if xf.applyNumberFormat, err = optionalStyleBoolean(entry.start, "applyNumberFormat"); err != nil {
		return styleXF{}, err
	}
	if xf.applyFont, err = optionalStyleBoolean(entry.start, "applyFont"); err != nil {
		return styleXF{}, err
	}
	if xf.applyFill, err = optionalStyleBoolean(entry.start, "applyFill"); err != nil {
		return styleXF{}, err
	}
	if xf.applyBorder, err = optionalStyleBoolean(entry.start, "applyBorder"); err != nil {
		return styleXF{}, err
	}
	if xf.applyAlignment, err = optionalStyleBoolean(entry.start, "applyAlignment"); err != nil {
		return styleXF{}, err
	}
	for _, child := range entry.children {
		if child.start.Name.Space == namespace && child.start.Name.Local == "alignment" {
			if xf.alignment.present {
				return styleXF{}, fmt.Errorf("duplicate alignment child")
			}
			alignment, err := parseStyleAlignment(data, child)
			if err != nil {
				return styleXF{}, err
			}
			xf.alignment = alignment
			xf.alignmentStart = child.span.start - entry.span.startTagEnd
			xf.alignmentEnd = child.span.end - entry.span.startTagEnd
			continue
		}
	}
	return xf, nil
}

func parseStyleAlignment(data []byte, node styleTableNode) (styleAlignmentRecord, error) {
	if err := requireEmptyStyleNode(data, node); err != nil {
		return styleAlignmentRecord{}, fmt.Errorf("alignment: %w", err)
	}
	alignment := styleAlignmentRecord{present: true, raw: bytes.Clone(data[node.span.start:node.span.end]), rawStart: bytes.Clone(data[node.span.start:node.span.startTagEnd]), qname: node.span.qname}
	if value, found, err := unqualifiedXMLAttribute(node.start, "horizontal"); err != nil {
		return alignment, err
	} else if found {
		alignment.horizontal = stringPointer(value)
	}
	if value, found, err := unqualifiedXMLAttribute(node.start, "vertical"); err != nil {
		return alignment, err
	} else if found {
		alignment.vertical = stringPointer(value)
	}
	if value, found, err := unqualifiedXMLAttribute(node.start, "wrapText"); err != nil {
		return alignment, err
	} else if found {
		parsed, err := ooxmlBoolean(value, true)
		if err != nil {
			return alignment, fmt.Errorf("alignment wrapText=%q is invalid", value)
		}
		alignment.wrap = boolPointer(parsed)
	}
	if value, found, err := unqualifiedXMLAttribute(node.start, "shrinkToFit"); err != nil {
		return alignment, err
	} else if found {
		parsed, err := ooxmlBoolean(value, true)
		if err != nil {
			return alignment, fmt.Errorf("alignment shrinkToFit=%q is invalid", value)
		}
		alignment.shrinkToFit = boolPointer(parsed)
	}
	if value, found, err := unqualifiedXMLAttribute(node.start, "textRotation"); err != nil {
		return alignment, err
	} else if found {
		parsed, err := strconv.Atoi(value)
		if err != nil || parsed < 0 || parsed > 255 {
			return alignment, fmt.Errorf("alignment textRotation=%q is invalid", value)
		}
		alignment.textRotation = intPointer(parsed)
	}
	attributes, err := rawStartTagAttributes(alignment.rawStart)
	if err != nil {
		return alignment, err
	}
	var other, otherExcept strings.Builder
	for _, attribute := range attributes {
		if attribute.qname == "horizontal" || attribute.qname == "vertical" || attribute.qname == "wrapText" {
			continue
		}
		other.Write(alignment.rawStart[attribute.leading:attribute.end])
		if attribute.qname == "shrinkToFit" || attribute.qname == "textRotation" {
			continue
		}
		otherExcept.Write(alignment.rawStart[attribute.leading:attribute.end])
	}
	alignment.otherAttrsKey = other.String()
	alignment.otherAttrsExceptPaintExact = otherExcept.String()
	return alignment, nil
}

func optionalStyleBoolean(start xml.StartElement, name string) (*bool, error) {
	raw, found, err := unqualifiedXMLAttribute(start, name)
	if err != nil || !found {
		return nil, err
	}
	value, err := ooxmlBoolean(raw, true)
	if err != nil {
		return nil, fmt.Errorf("%s=%q is invalid", name, raw)
	}
	return boolPointer(value), nil
}

func styleNodeValue(start xml.StartElement, name string) (string, error) {
	value, found, err := unqualifiedXMLAttribute(start, name)
	if err != nil {
		return "", err
	}
	if !found {
		return "", fmt.Errorf("requires %s", name)
	}
	return value, nil
}

func requireEmptyStyleNode(data []byte, node styleTableNode) error {
	if len(node.children) != 0 || len(bytes.TrimSpace(data[node.span.startTagEnd:node.span.endStart])) != 0 {
		return fmt.Errorf("element must be empty")
	}
	return nil
}

func requireEmptyStyleEntry(data []byte, entry styleTableEntry) error {
	if len(entry.children) != 0 || len(bytes.TrimSpace(data[entry.span.startTagEnd:entry.span.endStart])) != 0 {
		return fmt.Errorf("element must be empty")
	}
	return nil
}

func requireWhitespaceAroundStyleChildren(data []byte, span xmlSpan, children []styleTableNode) error {
	cursor := span.startTagEnd
	for _, child := range children {
		if child.span.start < cursor || child.span.end > span.endStart {
			return fmt.Errorf("child spans are not ordered within their parent")
		}
		if len(bytes.TrimSpace(data[cursor:child.span.start])) != 0 {
			return fmt.Errorf("element has unsupported direct text content")
		}
		cursor = child.span.end
	}
	if cursor > span.endStart || len(bytes.TrimSpace(data[cursor:span.endStart])) != 0 {
		return fmt.Errorf("element has unsupported direct text content")
	}
	return nil
}

func supportedRGBStyleColor(start xml.StartElement) (*string, bool) {
	rgb, found, err := unqualifiedXMLAttribute(start, "rgb")
	if err != nil || !found || len(rgb) != 8 || !strings.EqualFold(rgb[:2], "FF") {
		return nil, false
	}
	for _, character := range rgb[2:] {
		if !((character >= '0' && character <= '9') || (character >= 'a' && character <= 'f') || (character >= 'A' && character <= 'F')) {
			return nil, false
		}
	}
	if !styleAttributesOnly(start, "rgb") {
		return nil, false
	}
	value := "#" + strings.ToUpper(rgb[2:])
	return &value, true
}

func styleAttributesOnly(start xml.StartElement, allowed ...string) bool {
	allowedSet := make(map[string]bool, len(allowed))
	for _, name := range allowed {
		allowedSet[name] = true
	}
	for _, attribute := range start.Attr {
		if attribute.Name.Space != "" || !allowedSet[attribute.Name.Local] {
			return false
		}
	}
	return true
}

func fontSemanticKey(font styleFont) string {
	var out strings.Builder
	out.WriteString("font|")
	writeOptionalStringKey(&out, font.name)
	if font.size == nil {
		out.WriteString("-|")
	} else {
		out.WriteString(strconv.FormatFloat(*font.size, 'g', -1, 64) + "|")
	}
	out.WriteString(strconv.FormatBool(font.bold) + "|" + strconv.FormatBool(font.italic) + "|")
	writeOptionalStringKey(&out, font.color)
	out.WriteString(strconv.FormatBool(font.colorSafe) + "|")
	if attributes, err := rawStartTagAttributes(font.rootStart); err == nil {
		for _, attribute := range attributes {
			out.Write(font.rootStart[attribute.leading:attribute.end])
			out.WriteByte('|')
		}
	}
	for _, child := range font.children {
		switch child.local {
		case "name", "sz", "b", "i":
			writeStyleFontChildOpaqueAttributes(&out, child, "val")
			continue
		case "color":
			if font.colorSafe {
				writeStyleFontChildOpaqueAttributes(&out, child, "rgb")
				continue
			}
		}
		out.Write(child.raw)
		out.WriteByte('|')
	}
	return out.String()
}

func writeStyleFontChildOpaqueAttributes(out *strings.Builder, child styleFontChild, known string) {
	attributes, err := rawStartTagAttributes(child.raw)
	if err != nil {
		out.WriteString("invalid-child|")
		out.Write(child.raw)
		return
	}
	for _, attribute := range attributes {
		if attribute.qname == known {
			continue
		}
		out.Write(child.raw[attribute.leading:attribute.end])
		out.WriteByte('|')
	}
}

func fillSemanticKey(rootStart []byte, color *string) string {
	var out strings.Builder
	out.WriteString("fill|")
	writeOptionalStringKey(&out, color)
	if attributes, err := rawStartTagAttributes(rootStart); err == nil {
		for _, attribute := range attributes {
			out.Write(rootStart[attribute.leading:attribute.end])
			out.WriteByte('|')
		}
	}
	return out.String()
}

func writeOptionalStringKey(out *strings.Builder, value *string) {
	if value == nil {
		out.WriteString("-|")
		return
	}
	out.WriteString(strconv.Itoa(len(*value)))
	out.WriteByte(':')
	out.WriteString(*value)
	out.WriteByte('|')
}

func stringPointer(value string) *string  { return &value }
func floatPointer(value float64) *float64 { return &value }
func boolPointer(value bool) *bool        { return &value }
func intPointer(value int) *int           { return &value }

func exactThemeColorIndex(start xml.StartElement) *int {
	index, _ := exactThemeColor(start)
	return index
}

func exactThemeColor(start xml.StartElement) (*int, *float64) {
	raw, found, err := unqualifiedXMLAttribute(start, "theme")
	if err != nil || !found {
		return nil, nil
	}
	if !styleAttributesOnly(start, "theme") && !styleAttributesOnly(start, "theme", "tint") {
		return nil, nil
	}
	value, err := strconv.Atoi(raw)
	if err != nil || value < 0 || value > 11 {
		return nil, nil
	}
	tintRaw, tintFound, tintErr := unqualifiedXMLAttribute(start, "tint")
	if tintErr != nil {
		return nil, nil
	}
	var tint *float64
	if tintFound {
		parsed, ok := parseSpreadsheetTint(tintRaw)
		if !ok {
			return nil, nil
		}
		tint = parsed
	}
	return intPointer(value), tint
}

func parseSpreadsheetTint(raw string) (*float64, bool) {
	if raw == "" {
		return nil, false
	}
	value, err := strconv.ParseFloat(raw, 64)
	if err != nil || math.IsNaN(value) || math.IsInf(value, 0) || value < -1 || value > 1 {
		return nil, false
	}
	return floatPointer(value), true
}

func styleFontChildRaw(font styleFont, local string) []byte {
	for _, child := range font.children {
		if child.local == local {
			return bytes.Clone(child.raw)
		}
	}
	return nil
}

func styleFontValueChild(fontQName, local, value string) []byte {
	escaped, _ := escapeXMLAttribute(value)
	return []byte(fmt.Sprintf(`<%s val="%s"/>`, prefixedLocal(fontQName, local), escaped))
}

func replaceStyleFontChild(font *styleFont, local string, raw []byte) {
	updated := make([]styleFontChild, 0, len(font.children)+1)
	replaced := false
	for _, child := range font.children {
		if child.local != local {
			updated = append(updated, child)
			continue
		}
		if !replaced && len(raw) != 0 {
			updated = append(updated, styleFontChild{local: local, raw: bytes.Clone(raw)})
		}
		replaced = true
	}
	if !replaced && len(raw) != 0 {
		insertAt := len(updated)
		targetRank := styleFontChildRank(local)
		for index, child := range updated {
			if styleFontChildRank(child.local) > targetRank {
				insertAt = index
				break
			}
		}
		updated = append(updated, styleFontChild{})
		copy(updated[insertAt+1:], updated[insertAt:])
		updated[insertAt] = styleFontChild{local: local, raw: bytes.Clone(raw)}
	}
	font.children = updated
}

func styleFontChildRank(local string) int {
	order := map[string]int{
		"name": 0, "charset": 1, "family": 2, "b": 3, "i": 4, "strike": 5,
		"outline": 6, "shadow": 7, "condense": 8, "extend": 9, "color": 10,
		"sz": 11, "u": 12, "vertAlign": 13, "scheme": 14,
	}
	if rank, found := order[local]; found {
		return rank
	}
	return 100
}

func (registry *styleRegistry) validateXFReferences(xf styleXF, cellXF bool) error {
	if xf.fontID < 0 || xf.fontID >= len(registry.fonts) {
		return fmt.Errorf("fontId %d is outside the fonts table", xf.fontID)
	}
	if xf.fillID < 0 || xf.fillID >= len(registry.fills) {
		return fmt.Errorf("fillId %d is outside the fills table", xf.fillID)
	}
	if xf.borderID < 0 || xf.borderID >= len(registry.borders) {
		return fmt.Errorf("borderId %d is outside the borders table", xf.borderID)
	}
	if xf.numFmtID >= 164 {
		if _, found := registry.numFmtByID[xf.numFmtID]; !found {
			return fmt.Errorf("numFmtId %d is missing", xf.numFmtID)
		}
	}
	if cellXF && (xf.xfID < 0 || xf.xfID >= len(registry.styleXfs)) {
		return fmt.Errorf("xfId %d is outside cellStyleXfs", xf.xfID)
	}
	if cellXF {
		base := effectiveCellStyleXF(registry.styleXfs[xf.xfID])
		for _, component := range []struct {
			name              string
			direct, inherited int
			apply             *bool
		}{
			{name: "font", direct: xf.fontID, inherited: base.fontID, apply: xf.applyFont},
			{name: "fill", direct: xf.fillID, inherited: base.fillID, apply: xf.applyFill},
			{name: "border", direct: xf.borderID, inherited: base.borderID, apply: xf.applyBorder},
			{name: "number format", direct: xf.numFmtID, inherited: base.numFmtID, apply: xf.applyNumberFormat},
		} {
			if (component.apply == nil || !*component.apply) && component.direct != component.inherited {
				return fmt.Errorf("%s id differs from its cellStyleXf parent while its apply flag is false or absent", component.name)
			}
		}
		if (xf.applyAlignment == nil || !*xf.applyAlignment) && xf.alignment.present && !styleAlignmentsEquivalent(xf.alignment, base.alignment) {
			return fmt.Errorf("alignment differs from its cellStyleXf parent while applyAlignment is false or absent")
		}
	}
	return nil
}

func styleAlignmentsEquivalent(left, right styleAlignmentRecord) bool {
	return effectiveHorizontal(left.horizontal, nil) == effectiveHorizontal(right.horizontal, nil) &&
		effectiveVertical(left.vertical, nil) == effectiveVertical(right.vertical, nil) &&
		effectiveWrap(left.wrap, nil) == effectiveWrap(right.wrap, nil) &&
		left.otherAttrsKey == right.otherAttrsKey
}

func cloneString(value *string) *string {
	if value == nil {
		return nil
	}
	return stringPointer(*value)
}
func cloneFloat(value *float64) *float64 {
	if value == nil {
		return nil
	}
	return floatPointer(*value)
}
func cloneBool(value *bool) *bool {
	if value == nil {
		return nil
	}
	return boolPointer(*value)
}
func cloneInt(value *int) *int {
	if value == nil {
		return nil
	}
	return intPointer(*value)
}

func sortedStyleKeys(values map[cellKey]StyleDelta) []cellKey {
	keys := make([]cellKey, 0, len(values))
	for key := range values {
		keys = append(keys, key)
	}
	sort.Slice(keys, func(i, j int) bool {
		if keys[i].row == keys[j].row {
			return keys[i].column < keys[j].column
		}
		return keys[i].row < keys[j].row
	})
	return keys
}

func (registry *styleRegistry) resolveStyle(sourceIndex int, delta StyleDelta) (int, error) {
	if sourceIndex < 0 || sourceIndex >= len(registry.cellXfs) {
		return 0, fmt.Errorf("cell style index %d is outside cellXfs", sourceIndex)
	}
	source := registry.cellXfs[sourceIndex]
	base := effectiveCellStyleXF(registry.styleXfs[source.xfID])
	fontID, fontChanged, err := registry.resolveFont(source, base, delta)
	if err != nil {
		return 0, err
	}
	fillID, fillChanged, err := registry.resolveFill(source, base, delta.FillColor)
	if err != nil {
		return 0, err
	}
	numFmtID, numFmtChanged, err := registry.resolveNumberFormat(source, base, delta.NumberFormat)
	if err != nil {
		return 0, err
	}
	alignment, alignmentChanged, alignmentInherited, err := registry.resolveAlignment(source, base, delta)
	if err != nil {
		return 0, err
	}
	if !fontChanged && !fillChanged && !numFmtChanged && !alignmentChanged {
		return sourceIndex, nil
	}
	start := bytes.Clone(source.rootStart)
	if fontChanged {
		start, err = rewriteStyleXFComponent(start, "fontId", "applyFont", fontID, base.fontID)
		if err != nil {
			return 0, err
		}
	}
	if fillChanged {
		start, err = rewriteStyleXFComponent(start, "fillId", "applyFill", fillID, base.fillID)
		if err != nil {
			return 0, err
		}
	}
	if numFmtChanged {
		start, err = rewriteStyleXFComponent(start, "numFmtId", "applyNumberFormat", numFmtID, base.numFmtID)
		if err != nil {
			return 0, err
		}
	}
	if alignmentChanged {
		if alignmentInherited {
			start, err = rewriteUnqualifiedAttribute(start, "applyAlignment", nil)
		} else {
			one := "1"
			start, err = rewriteUnqualifiedAttribute(start, "applyAlignment", &one)
		}
		if err != nil {
			return 0, fmt.Errorf("rewrite applyAlignment: %w", err)
		}
	}
	start, err = canonicalizeStyleXFStart(start)
	if err != nil {
		return 0, err
	}
	candidate, body, alignmentStart, alignmentEnd := buildStyleXFRecord(start, source, alignment, alignmentChanged)
	if existing, found := registry.xfByRaw[string(candidate)]; found {
		return existing, nil
	}
	parsed := source
	parsed.raw, parsed.rootStart, parsed.body, parsed.alignment = candidate, start, body, alignment
	parsed.selfClosing = bytes.HasSuffix(bytes.TrimSpace(candidate), []byte("/>"))
	parsed.alignmentStart, parsed.alignmentEnd = alignmentStart, alignmentEnd
	if fontChanged {
		parsed.fontID, parsed.applyFont = fontID, appliedUnlessInherited(fontID, base.fontID)
	}
	if fillChanged {
		parsed.fillID, parsed.applyFill = fillID, appliedUnlessInherited(fillID, base.fillID)
	}
	if numFmtChanged {
		parsed.numFmtID, parsed.applyNumberFormat = numFmtID, appliedUnlessInherited(numFmtID, base.numFmtID)
	}
	if alignmentChanged {
		if alignmentInherited {
			parsed.applyAlignment = nil
		} else {
			parsed.applyAlignment = boolPointer(true)
		}
	}
	if err := registry.validateXFReferences(parsed, true); err != nil {
		return 0, err
	}
	if err := registry.ensureStyleRecordCapacity(); err != nil {
		return 0, err
	}
	index := len(registry.cellXfs)
	registry.cellXfs = append(registry.cellXfs, parsed)
	registry.appendedCellXfs = append(registry.appendedCellXfs, candidate)
	registry.xfByRaw[string(candidate)] = index
	return index, nil
}

func (registry *styleRegistry) resolveFont(source, base styleXF, delta StyleDelta) (int, bool, error) {
	touched := delta.FontName.Present || delta.FontSizePoints.Present || delta.Bold.Present || delta.Italic.Present || delta.FontColor.Present
	currentID := effectiveStyleComponent(source.fontID, base.fontID, source.applyFont)
	if !touched {
		return currentID, false, nil
	}
	if delta.FontName.Present && delta.FontName.Value == nil && delta.FontSizePoints.Present && delta.FontSizePoints.Value == nil &&
		delta.Bold.Present && delta.Bold.Value == nil && delta.Italic.Present && delta.Italic.Value == nil &&
		delta.FontColor.Present && delta.FontColor.Value == nil {
		return base.fontID, currentID != base.fontID, nil
	}
	current, inherited := registry.fonts[currentID], registry.fonts[base.fontID]
	desired := current
	desired.name, desired.size, desired.color = cloneString(current.name), cloneFloat(current.size), cloneString(current.color)
	if delta.FontName.Present {
		if delta.FontName.Value == nil {
			desired.name = cloneString(inherited.name)
			replaceStyleFontChild(&desired, "name", styleFontChildRaw(inherited, "name"))
		} else {
			desired.name = cloneString(delta.FontName.Value)
			replaceStyleFontChild(&desired, "name", styleFontValueChild(desired.qname, "name", *delta.FontName.Value))
		}
	}
	if delta.FontSizePoints.Present {
		if delta.FontSizePoints.Value == nil {
			desired.size = cloneFloat(inherited.size)
			replaceStyleFontChild(&desired, "sz", styleFontChildRaw(inherited, "sz"))
		} else {
			desired.size = cloneFloat(delta.FontSizePoints.Value)
			replaceStyleFontChild(&desired, "sz", styleFontValueChild(desired.qname, "sz", strconv.FormatFloat(*delta.FontSizePoints.Value, 'g', -1, 64)))
		}
	}
	if delta.Bold.Present {
		if delta.Bold.Value == nil {
			desired.bold = inherited.bold
			replaceStyleFontChild(&desired, "b", styleFontChildRaw(inherited, "b"))
		} else {
			desired.bold = *delta.Bold.Value
			var raw []byte
			if desired.bold {
				raw = []byte("<" + prefixedLocal(desired.qname, "b") + "/>")
			}
			replaceStyleFontChild(&desired, "b", raw)
		}
	}
	if delta.Italic.Present {
		if delta.Italic.Value == nil {
			desired.italic = inherited.italic
			replaceStyleFontChild(&desired, "i", styleFontChildRaw(inherited, "i"))
		} else {
			desired.italic = *delta.Italic.Value
			var raw []byte
			if desired.italic {
				raw = []byte("<" + prefixedLocal(desired.qname, "i") + "/>")
			}
			replaceStyleFontChild(&desired, "i", raw)
		}
	}
	if delta.FontColor.Present {
		if delta.FontColor.Value == nil {
			desired.color = cloneString(inherited.color)
			desired.colorSafe = inherited.colorSafe
			replaceStyleFontChild(&desired, "color", styleFontChildRaw(inherited, "color"))
		} else {
			desired.color = cloneString(delta.FontColor.Value)
			desired.colorSafe = true
			raw := []byte(fmt.Sprintf(`<%s rgb="FF%s"/>`, prefixedLocal(desired.qname, "color"), (*delta.FontColor.Value)[1:]))
			replaceStyleFontChild(&desired, "color", raw)
		}
	}
	desired.key = fontSemanticKey(desired)
	if desired.key == current.key {
		return currentID, false, nil
	}
	if existing, found := registry.fontByKey[desired.key]; found {
		return existing, true, nil
	}
	desired.raw = buildStyleFontRecord(desired)
	if err := registry.ensureStyleRecordCapacity(); err != nil {
		return 0, false, err
	}
	index := len(registry.fonts)
	registry.fonts = append(registry.fonts, desired)
	registry.appendedFonts = append(registry.appendedFonts, desired.raw)
	registry.fontByKey[desired.key] = index
	return index, true, nil
}

func (registry *styleRegistry) resolveFill(source, base styleXF, property StyleProperty[string]) (int, bool, error) {
	currentID := effectiveStyleComponent(source.fillID, base.fillID, source.applyFill)
	if !property.Present {
		return currentID, false, nil
	}
	current, inherited := registry.fills[currentID], registry.fills[base.fillID]
	if property.Value == nil {
		if !current.supported || !inherited.supported {
			return 0, false, fmt.Errorf("fill clear cannot safely inherit an unsupported gradient, non-solid pattern, theme, indexed, tint, or auto color")
		}
		if currentID == base.fillID {
			return currentID, false, nil
		}
		return base.fillID, true, nil
	}
	qname := current.qname
	rootStart := []byte("<" + qname + ">")
	key := fillSemanticKey(rootStart, property.Value)
	if current.supported && current.key == key {
		return currentID, false, nil
	}
	if existing, found := registry.fillByKey[key]; found {
		return existing, true, nil
	}
	fill := styleFill{qname: qname, rootStart: rootStart, color: cloneString(property.Value), supported: true, key: key}
	fill.raw = buildStyleFillRecord(fill)
	if err := registry.ensureStyleRecordCapacity(); err != nil {
		return 0, false, err
	}
	index := len(registry.fills)
	registry.fills = append(registry.fills, fill)
	registry.appendedFills = append(registry.appendedFills, fill.raw)
	registry.fillByKey[key] = index
	return index, true, nil
}

func (registry *styleRegistry) resolveNumberFormat(source, base styleXF, property StyleProperty[string]) (int, bool, error) {
	currentID := effectiveStyleComponent(source.numFmtID, base.numFmtID, source.applyNumberFormat)
	if !property.Present {
		return currentID, false, nil
	}
	if property.Value == nil {
		if currentID == base.numFmtID {
			return currentID, false, nil
		}
		return base.numFmtID, true, nil
	}
	if currentCode, found := registry.numberFormatCode(currentID); found && currentCode == *property.Value {
		return currentID, false, nil
	}
	if existing, found := registry.numFmtByCode[*property.Value]; found {
		return existing, existing != currentID, nil
	}
	newID := 164
	for registry.usedNumFmtID[newID] {
		newID++
	}
	qname := prefixedLocal(registry.index.root.qname, "numFmt")
	if registry.index.numFmts != nil && len(registry.index.numFmts.entries) > 0 {
		qname = registry.index.numFmts.entries[0].span.qname
	}
	escaped, err := escapeXMLAttribute(*property.Value)
	if err != nil {
		return 0, false, err
	}
	raw := []byte(fmt.Sprintf(`<%s numFmtId="%d" formatCode="%s"/>`, qname, newID, escaped))
	if err := registry.ensureStyleRecordCapacity(); err != nil {
		return 0, false, err
	}
	registry.usedNumFmtID[newID] = true
	registry.numFmtByID[newID] = *property.Value
	registry.numFmtByCode[*property.Value] = newID
	registry.appendedNumFmts = append(registry.appendedNumFmts, raw)
	return newID, true, nil
}

func (registry *styleRegistry) resolveAlignment(source, base styleXF, delta StyleDelta) (styleAlignmentRecord, bool, bool, error) {
	touched := delta.HorizontalAlignment.Present || delta.VerticalAlignment.Present || delta.WrapText.Present
	if !touched {
		return source.alignment, false, false, nil
	}
	baseEffective := base.alignment
	directApplied := source.applyAlignment != nil && *source.applyAlignment
	current := baseEffective
	if directApplied {
		current = source.alignment
	}
	// applyAlignment governs the complete alignment component. Once direct
	// alignment is applied, omitted attributes take their schema defaults; they
	// do not individually inherit from the cellStyleXf alignment.
	currentHorizontal := effectiveHorizontal(current.horizontal, nil)
	currentVertical := effectiveVertical(current.vertical, nil)
	currentWrap := effectiveWrap(current.wrap, nil)
	baseHorizontal := effectiveHorizontal(baseEffective.horizontal, nil)
	baseVertical := effectiveVertical(baseEffective.vertical, nil)
	baseWrap := effectiveWrap(baseEffective.wrap, nil)
	desiredHorizontal, desiredVertical, desiredWrap := currentHorizontal, currentVertical, currentWrap
	if delta.HorizontalAlignment.Present {
		if delta.HorizontalAlignment.Value == nil {
			desiredHorizontal = baseHorizontal
		} else {
			desiredHorizontal = protocolHorizontal(*delta.HorizontalAlignment.Value)
		}
	}
	if delta.VerticalAlignment.Present {
		if delta.VerticalAlignment.Value == nil {
			desiredVertical = baseVertical
		} else {
			desiredVertical = protocolVertical(*delta.VerticalAlignment.Value)
		}
	}
	if delta.WrapText.Present {
		if delta.WrapText.Value == nil {
			desiredWrap = baseWrap
		} else {
			desiredWrap = *delta.WrapText.Value
		}
	}
	if desiredHorizontal == currentHorizontal && desiredVertical == currentVertical && desiredWrap == currentWrap {
		return source.alignment, false, false, nil
	}
	working := current
	if !working.present {
		qname := prefixedLocal(source.qname, "alignment")
		working = styleAlignmentRecord{present: true, qname: qname, rawStart: []byte("<" + qname + "/>")}
	}
	start := bytes.Clone(working.rawStart)
	var err error
	if delta.HorizontalAlignment.Present {
		if delta.HorizontalAlignment.Value == nil {
			if baseHorizontal == "general" {
				start, err = rewriteUnqualifiedAttribute(start, "horizontal", nil)
			} else {
				start, err = rewriteUnqualifiedAttribute(start, "horizontal", &baseHorizontal)
			}
		} else {
			value := protocolHorizontal(*delta.HorizontalAlignment.Value)
			start, err = rewriteUnqualifiedAttribute(start, "horizontal", &value)
		}
		if err != nil {
			return working, false, false, err
		}
	}
	if delta.VerticalAlignment.Present {
		if delta.VerticalAlignment.Value == nil {
			if baseVertical == "bottom" {
				start, err = rewriteUnqualifiedAttribute(start, "vertical", nil)
			} else {
				start, err = rewriteUnqualifiedAttribute(start, "vertical", &baseVertical)
			}
		} else {
			value := protocolVertical(*delta.VerticalAlignment.Value)
			start, err = rewriteUnqualifiedAttribute(start, "vertical", &value)
		}
		if err != nil {
			return working, false, false, err
		}
	}
	if delta.WrapText.Present {
		if delta.WrapText.Value == nil {
			if !baseWrap {
				start, err = rewriteUnqualifiedAttribute(start, "wrapText", nil)
			} else {
				one := "1"
				start, err = rewriteUnqualifiedAttribute(start, "wrapText", &one)
			}
		} else {
			value := "0"
			if *delta.WrapText.Value {
				value = "1"
			}
			start, err = rewriteUnqualifiedAttribute(start, "wrapText", &value)
		}
		if err != nil {
			return working, false, false, err
		}
	}
	start, err = canonicalizeStyleAlignmentStart(start)
	if err != nil {
		return working, false, false, err
	}
	parsed, err := parseAlignmentStart(start)
	if err != nil {
		return working, false, false, err
	}
	if styleAlignmentsEquivalent(parsed, baseEffective) {
		return styleAlignmentRecord{}, true, true, nil
	}
	parsed.raw = canonicalStyleEmptyElement(start)
	parsed.rawStart = parsed.raw
	return parsed, true, false, nil
}

func effectiveStyleComponent(direct, inherited int, apply *bool) int {
	if apply == nil || !*apply {
		return inherited
	}
	return direct
}

func effectiveCellStyleXF(base styleXF) styleXF {
	effective := base
	if base.applyNumberFormat != nil && !*base.applyNumberFormat {
		effective.numFmtID = 0
	}
	if base.applyFont != nil && !*base.applyFont {
		effective.fontID = 0
	}
	if base.applyFill != nil && !*base.applyFill {
		effective.fillID = 0
	}
	if base.applyBorder != nil && !*base.applyBorder {
		effective.borderID = 0
	}
	if base.applyAlignment != nil && !*base.applyAlignment {
		effective.alignment = styleAlignmentRecord{}
	}
	return effective
}

func effectiveStyleAlignment(direct, inherited styleXF) styleAlignmentRecord {
	if direct.applyAlignment == nil || !*direct.applyAlignment {
		return inherited.alignment
	}
	return direct.alignment
}

func effectiveHorizontal(value, inherited *string) string {
	if value != nil {
		return *value
	}
	if inherited != nil {
		return *inherited
	}
	return "general"
}

func effectiveVertical(value, inherited *string) string {
	if value != nil {
		return *value
	}
	if inherited != nil {
		return *inherited
	}
	return "bottom"
}

func effectiveWrap(value, inherited *bool) bool {
	if value != nil {
		return *value
	}
	if inherited != nil {
		return *inherited
	}
	return false
}

func protocolHorizontal(value HorizontalAlignment) string { return string(value) }
func protocolVertical(value VerticalAlignment) string {
	if value == VerticalMiddle {
		return "center"
	}
	return string(value)
}

func (registry *styleRegistry) numberFormatCode(id int) (string, bool) {
	// Real Microsoft Excel output can declare an explicit formatCode for an ID
	// in the nominal built-in range. The authored numFmt table is source
	// authority and therefore overrides our fallback map for that exact ID.
	if value, found := registry.numFmtByID[id]; found {
		return value, true
	}
	if id < 164 {
		value, found := builtinNumberFormats[id]
		return value, found
	}
	value, found := registry.numFmtByID[id]
	return value, found
}

func rewriteStyleXFComponent(start []byte, idName, applyName string, id, inherited int) ([]byte, error) {
	value := strconv.Itoa(id)
	updated, err := rewriteUnqualifiedAttribute(start, idName, &value)
	if err != nil {
		return nil, fmt.Errorf("rewrite %s: %w", idName, err)
	}
	if id == inherited {
		updated, err = rewriteUnqualifiedAttribute(updated, applyName, nil)
	} else {
		one := "1"
		updated, err = rewriteUnqualifiedAttribute(updated, applyName, &one)
	}
	if err != nil {
		return nil, fmt.Errorf("rewrite %s: %w", applyName, err)
	}
	return updated, nil
}

func appliedUnlessInherited(id, inherited int) *bool {
	if id == inherited {
		return nil
	}
	return boolPointer(true)
}

func canonicalizeStyleXFStart(start []byte) ([]byte, error) {
	element, err := decodeStartElement(start)
	if err != nil {
		return nil, err
	}
	type attributeValue struct {
		name  string
		value string
		found bool
	}
	values := make([]attributeValue, 0, 9)
	for _, name := range []string{"numFmtId", "fontId", "fillId", "borderId", "xfId"} {
		value, found, err := unqualifiedXMLAttribute(element, name)
		if err != nil {
			return nil, err
		}
		if found {
			parsed, err := optionalUnsignedStyleAttribute(element, name, 0)
			if err != nil {
				return nil, err
			}
			value = strconv.Itoa(parsed)
		}
		values = append(values, attributeValue{name: name, value: value, found: found})
	}
	for _, name := range []string{"applyNumberFormat", "applyFont", "applyFill", "applyBorder", "applyAlignment"} {
		value, found, err := unqualifiedXMLAttribute(element, name)
		if err != nil {
			return nil, err
		}
		if found {
			parsed, err := ooxmlBoolean(value, true)
			if err != nil {
				return nil, fmt.Errorf("%s=%q is invalid", name, value)
			}
			value = "0"
			if parsed {
				value = "1"
			}
		}
		values = append(values, attributeValue{name: name, value: value, found: found})
	}
	updated := bytes.Clone(start)
	for _, value := range values {
		updated, err = rewriteUnqualifiedAttribute(updated, value.name, nil)
		if err != nil {
			return nil, err
		}
	}
	for _, value := range values {
		if !value.found {
			continue
		}
		updated, err = rewriteUnqualifiedAttribute(updated, value.name, &value.value)
		if err != nil {
			return nil, err
		}
	}
	return updated, nil
}

func buildStyleFontRecord(font styleFont) []byte {
	var out bytes.Buffer
	out.Write(openTag(font.rootStart))
	for _, child := range font.children {
		out.Write(child.raw)
	}
	out.WriteString("</" + font.qname + ">")
	return out.Bytes()
}

func buildStyleFillRecord(fill styleFill) []byte {
	var out bytes.Buffer
	out.Write(openTag(fill.rootStart))
	patternQName := prefixedLocal(fill.qname, "patternFill")
	if fill.color == nil {
		fmt.Fprintf(&out, `<%s patternType="none"/>`, patternQName)
	} else {
		fmt.Fprintf(&out, `<%s patternType="solid"><%s rgb="FF%s"/>`, patternQName, prefixedLocal(fill.qname, "fgColor"), (*fill.color)[1:])
		if fill.background64 {
			fmt.Fprintf(&out, `<%s indexed="64"/>`, prefixedLocal(fill.qname, "bgColor"))
		}
		fmt.Fprintf(&out, `</%s>`, patternQName)
	}
	out.WriteString("</" + fill.qname + ">")
	return out.Bytes()
}

func buildStyleXFRecord(start []byte, source styleXF, alignment styleAlignmentRecord, alignmentChanged bool) ([]byte, []byte, int, int) {
	body := bytes.Clone(source.body)
	alignmentStart, alignmentEnd := source.alignmentStart, source.alignmentEnd
	if alignmentChanged {
		var replacement []byte
		if alignment.present {
			if len(alignment.raw) > 0 {
				replacement = bytes.Clone(alignment.raw)
			} else {
				replacement = canonicalStyleEmptyElement(alignment.rawStart)
			}
		}
		if source.alignmentStart >= 0 {
			rebuilt := make([]byte, 0, len(body)-(source.alignmentEnd-source.alignmentStart)+len(replacement))
			rebuilt = append(rebuilt, body[:source.alignmentStart]...)
			rebuilt = append(rebuilt, replacement...)
			rebuilt = append(rebuilt, body[source.alignmentEnd:]...)
			body = rebuilt
			if len(replacement) == 0 {
				alignmentStart, alignmentEnd = -1, -1
			} else {
				alignmentStart, alignmentEnd = source.alignmentStart, source.alignmentStart+len(replacement)
			}
		} else if len(replacement) > 0 {
			body = append(replacement, body...)
			alignmentStart, alignmentEnd = 0, len(replacement)
		}
	}
	if len(body) == 0 && (source.selfClosing || alignmentChanged) {
		raw := canonicalStyleEmptyElement(start)
		return raw, body, alignmentStart, alignmentEnd
	}
	var out bytes.Buffer
	out.Write(openTag(start))
	out.Write(body)
	out.WriteString("</" + source.qname + ">")
	return out.Bytes(), body, alignmentStart, alignmentEnd
}

func canonicalStyleEmptyElement(start []byte) []byte {
	trimmed := bytes.TrimSpace(start)
	if bytes.HasSuffix(trimmed, []byte("/>")) {
		return bytes.Clone(start)
	}
	closeAt := bytes.LastIndexByte(start, '>')
	if closeAt < 0 {
		return bytes.Clone(start)
	}
	return append(bytes.Clone(start[:closeAt]), append([]byte{'/'}, start[closeAt:]...)...)
}

func parseAlignmentStart(start []byte) (styleAlignmentRecord, error) {
	element, err := decodeStartElement(start)
	if err != nil {
		return styleAlignmentRecord{}, err
	}
	record := styleAlignmentRecord{present: true, rawStart: bytes.Clone(start), qname: rawQName(start, 0, len(start))}
	if value, found, err := unqualifiedXMLAttribute(element, "horizontal"); err != nil {
		return record, err
	} else if found {
		record.horizontal = stringPointer(value)
	}
	if value, found, err := unqualifiedXMLAttribute(element, "vertical"); err != nil {
		return record, err
	} else if found {
		record.vertical = stringPointer(value)
	}
	if value, found, err := unqualifiedXMLAttribute(element, "wrapText"); err != nil {
		return record, err
	} else if found {
		parsed, err := ooxmlBoolean(value, true)
		if err != nil {
			return record, err
		}
		record.wrap = boolPointer(parsed)
	}
	attributes, err := rawStartTagAttributes(start)
	if err != nil {
		return record, err
	}
	var other strings.Builder
	for _, attribute := range attributes {
		if attribute.qname != "horizontal" && attribute.qname != "vertical" && attribute.qname != "wrapText" {
			other.Write(start[attribute.leading:attribute.end])
		}
	}
	record.otherAttrsKey = other.String()
	return record, nil
}

func canonicalizeStyleAlignmentStart(start []byte) ([]byte, error) {
	parsed, err := parseAlignmentStart(start)
	if err != nil {
		return nil, err
	}
	updated := bytes.Clone(start)
	for _, name := range []string{"horizontal", "vertical", "wrapText"} {
		updated, err = rewriteUnqualifiedAttribute(updated, name, nil)
		if err != nil {
			return nil, err
		}
	}
	if parsed.horizontal != nil {
		updated, err = rewriteUnqualifiedAttribute(updated, "horizontal", parsed.horizontal)
		if err != nil {
			return nil, err
		}
	}
	if parsed.vertical != nil {
		updated, err = rewriteUnqualifiedAttribute(updated, "vertical", parsed.vertical)
		if err != nil {
			return nil, err
		}
	}
	if parsed.wrap != nil {
		value := "0"
		if *parsed.wrap {
			value = "1"
		}
		updated, err = rewriteUnqualifiedAttribute(updated, "wrapText", &value)
		if err != nil {
			return nil, err
		}
	}
	return updated, nil
}

func (registry *styleRegistry) changed() bool {
	return len(registry.appendedNumFmts)+len(registry.appendedFonts)+len(registry.appendedFills)+len(registry.appendedCellXfs) > 0
}

func (registry *styleRegistry) ensureStyleRecordCapacity() error {
	total := len(registry.numFmtByID) + len(registry.fonts) + len(registry.fills) +
		len(registry.index.borders.entries) + len(registry.styleXfs) + len(registry.cellXfs)
	return ensureStyleRecordCountCapacity(total)
}

func ensureStyleRecordCountCapacity(total int) error {
	if total >= maxStyleTableRecords {
		return fmt.Errorf("appending a style record would exceed the %d-record safety limit", maxStyleTableRecords)
	}
	return nil
}

func (registry *styleRegistry) render() ([]byte, error) {
	type edit struct {
		start, end int
		data       []byte
	}
	edits := make([]edit, 0, 8)
	appendContainer := func(container *styleTableContainer, additions [][]byte) error {
		if len(additions) == 0 {
			return nil
		}
		count := strconv.Itoa(container.count + len(additions))
		start, err := rewriteUnqualifiedAttribute(registry.data[container.start:container.startTagEnd], "count", &count)
		if err != nil {
			return err
		}
		edits = append(edits, edit{start: container.start, end: container.startTagEnd, data: start})
		var added bytes.Buffer
		for _, record := range additions {
			added.Write(record)
		}
		if container.selfClosing(registry.data) {
			open := openTag(start)
			open = append(open, added.Bytes()...)
			open = append(open, []byte("</"+container.qname+">")...)
			edits[len(edits)-1] = edit{start: container.start, end: container.end, data: open}
		} else {
			edits = append(edits, edit{start: container.endStart, end: container.endStart, data: added.Bytes()})
		}
		return nil
	}
	if len(registry.appendedNumFmts) > 0 {
		if registry.index.numFmts == nil {
			qname := prefixedLocal(registry.index.root.qname, "numFmts")
			var added bytes.Buffer
			fmt.Fprintf(&added, `<%s count="%d">`, qname, len(registry.appendedNumFmts))
			for _, record := range registry.appendedNumFmts {
				added.Write(record)
			}
			added.WriteString("</" + qname + ">")
			edits = append(edits, edit{start: registry.index.fonts.start, end: registry.index.fonts.start, data: added.Bytes()})
		} else if err := appendContainer(registry.index.numFmts, registry.appendedNumFmts); err != nil {
			return nil, err
		}
	}
	if err := appendContainer(registry.index.fonts, registry.appendedFonts); err != nil {
		return nil, err
	}
	if err := appendContainer(registry.index.fills, registry.appendedFills); err != nil {
		return nil, err
	}
	if err := appendContainer(registry.index.cellXfs, registry.appendedCellXfs); err != nil {
		return nil, err
	}
	sort.Slice(edits, func(i, j int) bool {
		if edits[i].start == edits[j].start {
			return edits[i].end > edits[j].end
		}
		return edits[i].start > edits[j].start
	})
	updated := bytes.Clone(registry.data)
	for _, edit := range edits {
		updated = append(updated[:edit.start], append(edit.data, updated[edit.end:]...)...)
	}
	if _, err := parseStyleTable(updated); err != nil {
		return nil, fmt.Errorf("generated style table is invalid: %w", err)
	}
	return updated, nil
}
