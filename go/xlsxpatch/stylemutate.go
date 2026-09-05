package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"io"
	"math"
	"strings"
	"unicode/utf8"
)

const (
	StylePatch            = "style.patch"
	maxStyleRangeCells    = 100_000
	maxNumberFormatLength = 255
	maxFontNameLength     = 255
	maxFontSizePoints     = 409.0
	contentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types"
	stylesPartContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
)

// StyleProperty represents the three wire states of a style.patch property.
// Present=false means omitted/unchanged; Present=true with Value=nil means
// null/clear; Present=true with Value!=nil sets the supplied value.
type StyleProperty[T any] struct {
	Present bool
	Value   *T
}

func SetStyleProperty[T any](value T) StyleProperty[T] {
	return StyleProperty[T]{Present: true, Value: &value}
}

func ClearStyleProperty[T any]() StyleProperty[T] {
	return StyleProperty[T]{Present: true}
}

func (property *StyleProperty[T]) UnmarshalJSON(data []byte) error {
	property.Present = true
	if bytes.Equal(bytes.TrimSpace(data), []byte("null")) {
		property.Value = nil
		return nil
	}
	var value T
	if err := json.Unmarshal(data, &value); err != nil {
		return err
	}
	property.Value = &value
	return nil
}

type StyleRange struct {
	Row       int `json:"row"`
	Column    int `json:"column"`
	EndRow    int `json:"end_row"`
	EndColumn int `json:"end_column"`
}

type HorizontalAlignment string
type VerticalAlignment string

const (
	HorizontalGeneral HorizontalAlignment = "general"
	HorizontalLeft    HorizontalAlignment = "left"
	HorizontalCenter  HorizontalAlignment = "center"
	HorizontalRight   HorizontalAlignment = "right"

	VerticalTop    VerticalAlignment = "top"
	VerticalMiddle VerticalAlignment = "middle"
	VerticalBottom VerticalAlignment = "bottom"
)

type StyleDelta struct {
	NumberFormat        StyleProperty[string]              `json:"number_format"`
	FontName            StyleProperty[string]              `json:"font_name"`
	FontSizePoints      StyleProperty[float64]             `json:"font_size_points"`
	Bold                StyleProperty[bool]                `json:"bold"`
	Italic              StyleProperty[bool]                `json:"italic"`
	FontColor           StyleProperty[string]              `json:"font_color"`
	FillColor           StyleProperty[string]              `json:"fill_color"`
	HorizontalAlignment StyleProperty[HorizontalAlignment] `json:"horizontal_alignment"`
	VerticalAlignment   StyleProperty[VerticalAlignment]   `json:"vertical_alignment"`
	WrapText            StyleProperty[bool]                `json:"wrap_text"`
}

type StylePatchMutation struct {
	OperationID string     `json:"operation_id"`
	SheetID     string     `json:"sheet_id"`
	Kind        string     `json:"kind"`
	Range       StyleRange `json:"range"`
	Style       StyleDelta `json:"style"`
}

type normalizedStyleMutation struct {
	StylePatchMutation
}

// ApplyStyleMutations applies exactly @injoffice/sheets v1 style.patch to
// bounded worksheet ranges. The package, style table, and every target sheet
// are validated before Apply writes an atomic result.
func ApplyStyleMutations(orig []byte, mutations []StylePatchMutation) ([]byte, error) {
	normalized, err := validateStyleMutations(mutations)
	if err != nil {
		return nil, err
	}
	zr, err := zip.NewReader(bytes.NewReader(orig), int64(len(orig)))
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: style mutations: open original: %w", err)
	}
	packageIndex, err := newOPCPackageIndex(zr)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: style mutations: %w", err)
	}
	files := packageIndex.byExact
	cache := make(map[string][]byte)
	var readErr error
	readBytes := func(name string) ([]byte, bool) {
		if value, ok := cache[name]; ok {
			return value, true
		}
		file, ok := files[name]
		if !ok || readErr != nil {
			return nil, false
		}
		data, err := readZipFile(file)
		if err != nil {
			readErr = fmt.Errorf("read %q: %w", name, err)
			return nil, false
		}
		cache[name] = data
		return data, true
	}
	readText := func(name string) (string, bool) {
		data, ok := readBytes(name)
		return string(data), ok
	}
	workbook, err := locateWorkbookPart(packageIndex, readText)
	if readErr != nil {
		return nil, fmt.Errorf("xlsxpatch: style mutations: %w", readErr)
	}
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: style mutations: %w", err)
	}
	stylesPart, err := locateStylesPart(packageIndex, readBytes, workbook)
	if readErr != nil {
		return nil, fmt.Errorf("xlsxpatch: style mutations: %w", readErr)
	}
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: style mutations: %w", err)
	}
	stylesXML, ok := readBytes(stylesPart)
	if !ok {
		return nil, fmt.Errorf("xlsxpatch: style mutations: missing styles part %q", stylesPart)
	}
	registry, err := newStyleRegistry(stylesXML)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: style mutations: styles %q: %w", stylesPart, err)
	}

	byPart := make(map[string]map[cellKey]StyleDelta)
	partOrder := make([]string, 0)
	resolvedSheets := make(map[string]string)
	for _, mutation := range normalized {
		part, found := resolvedSheets[mutation.SheetID]
		if !found {
			part, err = worksheetPartForID(packageIndex, readText, workbook, mutation.SheetID)
			if readErr != nil {
				return nil, fmt.Errorf("xlsxpatch: style mutations: operation %q: %w", mutation.OperationID, readErr)
			}
			if err != nil {
				return nil, fmt.Errorf("xlsxpatch: style mutations: operation %q: %w", mutation.OperationID, err)
			}
			resolvedSheets[mutation.SheetID] = part
			if _, seen := byPart[part]; !seen {
				byPart[part] = make(map[cellKey]StyleDelta)
				partOrder = append(partOrder, part)
			}
		}
		for row := mutation.Range.Row; row <= mutation.Range.EndRow; row++ {
			for column := mutation.Range.Column; column <= mutation.Range.EndColumn; column++ {
				key := cellKey{row: row, column: column}
				byPart[part][key] = mergeStyleDelta(byPart[part][key], mutation.Style)
			}
		}
	}

	patch := Patch{Replace: make(map[string][]byte, len(partOrder)+1)}
	for _, part := range partOrder {
		sheetXML, ok := readBytes(part)
		if !ok {
			return nil, fmt.Errorf("xlsxpatch: style mutations: worksheet part %q is missing", part)
		}
		updated, err := patchWorksheetStyles(sheetXML, byPart[part], registry)
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: style mutations: worksheet %q: %w", part, err)
		}
		if !bytes.Equal(updated, sheetXML) {
			patch.Replace[part] = updated
		}
	}
	if registry.changed() {
		updatedStyles, err := registry.render()
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: style mutations: render styles: %w", err)
		}
		patch.Replace[stylesPart] = updatedStyles
	}
	if len(patch.Replace) == 0 {
		return bytes.Clone(orig), nil
	}
	return Apply(orig, patch)
}

func validateStyleMutations(mutations []StylePatchMutation) ([]normalizedStyleMutation, error) {
	if len(mutations) == 0 {
		return nil, fmt.Errorf("xlsxpatch: style mutations: empty batch")
	}
	if len(mutations) > maxCellMutations {
		return nil, fmt.Errorf("xlsxpatch: style mutations: batch exceeds %d operations", maxCellMutations)
	}
	seenIDs := make(map[string]bool, len(mutations))
	normalized := make([]normalizedStyleMutation, 0, len(mutations))
	totalCells := uint64(0)
	for index, mutation := range mutations {
		prefix := fmt.Sprintf("xlsxpatch: style mutations: operation %d", index)
		if mutation.OperationID == "" || len(mutation.OperationID) > maxOperationIDLen || !restrictedIDPattern.MatchString(mutation.OperationID) {
			return nil, fmt.Errorf("%s: invalid operation_id %q", prefix, mutation.OperationID)
		}
		if seenIDs[mutation.OperationID] {
			return nil, fmt.Errorf("%s: duplicate operation_id %q", prefix, mutation.OperationID)
		}
		seenIDs[mutation.OperationID] = true
		if mutation.SheetID == "" || !utf8.ValidString(mutation.SheetID) || utf16Length(mutation.SheetID) > maxStableSheetIDLen || strings.TrimSpace(mutation.SheetID) != mutation.SheetID {
			return nil, fmt.Errorf("%s: invalid sheet_id", prefix)
		}
		if mutation.Kind != StylePatch {
			return nil, fmt.Errorf("%s: unsupported kind %q", prefix, mutation.Kind)
		}
		rangeRef := mutation.Range
		if rangeRef.Row < 0 || rangeRef.Row >= excelMaxRows || rangeRef.EndRow < rangeRef.Row || rangeRef.EndRow >= excelMaxRows || rangeRef.Column < 0 || rangeRef.Column >= excelMaxColumns || rangeRef.EndColumn < rangeRef.Column || rangeRef.EndColumn >= excelMaxColumns {
			return nil, fmt.Errorf("%s: range is outside Excel limits or reversed", prefix)
		}
		area := uint64(rangeRef.EndRow-rangeRef.Row+1) * uint64(rangeRef.EndColumn-rangeRef.Column+1)
		totalCells += area
		if totalCells > maxStyleRangeCells {
			return nil, fmt.Errorf("%s: cumulative style range exceeds %d cells", prefix, maxStyleRangeCells)
		}
		if err := validateStyleDelta(mutation.Style); err != nil {
			return nil, fmt.Errorf("%s: %w", prefix, err)
		}
		normalized = append(normalized, normalizedStyleMutation{StylePatchMutation: mutation})
	}
	return normalized, nil
}

func validateStyleDelta(delta StyleDelta) error {
	properties := []bool{
		delta.NumberFormat.Present, delta.FontName.Present, delta.FontSizePoints.Present,
		delta.Bold.Present, delta.Italic.Present, delta.FontColor.Present,
		delta.FillColor.Present, delta.HorizontalAlignment.Present,
		delta.VerticalAlignment.Present, delta.WrapText.Present,
	}
	anyPresent := false
	for _, present := range properties {
		anyPresent = anyPresent || present
	}
	if !anyPresent {
		return fmt.Errorf("style delta is empty")
	}
	if err := validateStylePropertyState(delta); err != nil {
		return err
	}
	if value := delta.NumberFormat.Value; delta.NumberFormat.Present && value != nil && (*value == "" || utf16Length(*value) > maxNumberFormatLength || !utf8.ValidString(*value)) {
		return fmt.Errorf("number_format must contain 1..%d UTF-16 code units", maxNumberFormatLength)
	}
	if value := delta.NumberFormat.Value; delta.NumberFormat.Present && value != nil && !styleXMLString(*value) {
		return fmt.Errorf("number_format contains a character forbidden by XML 1.0")
	}
	if value := delta.FontName.Value; delta.FontName.Present && value != nil && (*value == "" || utf16Length(*value) > maxFontNameLength || !utf8.ValidString(*value)) {
		return fmt.Errorf("font_name must contain 1..%d UTF-16 code units", maxFontNameLength)
	}
	if value := delta.FontName.Value; delta.FontName.Present && value != nil && !styleXMLString(*value) {
		return fmt.Errorf("font_name contains a character forbidden by XML 1.0")
	}
	if value := delta.FontSizePoints.Value; delta.FontSizePoints.Present && value != nil && (math.IsNaN(*value) || math.IsInf(*value, 0) || *value < 1 || *value > maxFontSizePoints) {
		return fmt.Errorf("font_size_points must be finite and within 1..%g", maxFontSizePoints)
	}
	for name, property := range map[string]StyleProperty[string]{"font_color": delta.FontColor, "fill_color": delta.FillColor} {
		if property.Present && property.Value != nil && !canonicalRGBColor(*property.Value) {
			return fmt.Errorf("%s must be canonical uppercase #RRGGBB", name)
		}
	}
	if value := delta.HorizontalAlignment.Value; delta.HorizontalAlignment.Present && value != nil && *value != HorizontalGeneral && *value != HorizontalLeft && *value != HorizontalCenter && *value != HorizontalRight {
		return fmt.Errorf("unsupported horizontal_alignment %q", *value)
	}
	if value := delta.VerticalAlignment.Value; delta.VerticalAlignment.Present && value != nil && *value != VerticalTop && *value != VerticalMiddle && *value != VerticalBottom {
		return fmt.Errorf("unsupported vertical_alignment %q", *value)
	}
	return nil
}

func styleXMLString(value string) bool {
	for _, character := range value {
		if xmlForbiddenRune(character) {
			return false
		}
	}
	return true
}

func validateStylePropertyState(delta StyleDelta) error {
	// A value without Present=true cannot come from the wire and would make
	// omitted/null semantics ambiguous in direct Go callers.
	if (!delta.NumberFormat.Present && delta.NumberFormat.Value != nil) ||
		(!delta.FontName.Present && delta.FontName.Value != nil) ||
		(!delta.FontSizePoints.Present && delta.FontSizePoints.Value != nil) ||
		(!delta.Bold.Present && delta.Bold.Value != nil) ||
		(!delta.Italic.Present && delta.Italic.Value != nil) ||
		(!delta.FontColor.Present && delta.FontColor.Value != nil) ||
		(!delta.FillColor.Present && delta.FillColor.Value != nil) ||
		(!delta.HorizontalAlignment.Present && delta.HorizontalAlignment.Value != nil) ||
		(!delta.VerticalAlignment.Present && delta.VerticalAlignment.Value != nil) ||
		(!delta.WrapText.Present && delta.WrapText.Value != nil) {
		return fmt.Errorf("style property value requires Present=true")
	}
	return nil
}

func canonicalRGBColor(value string) bool {
	if len(value) != 7 || value[0] != '#' {
		return false
	}
	for _, character := range value[1:] {
		if !((character >= '0' && character <= '9') || (character >= 'A' && character <= 'F')) {
			return false
		}
	}
	return true
}

func mergeStyleDelta(current, next StyleDelta) StyleDelta {
	if next.NumberFormat.Present {
		current.NumberFormat = next.NumberFormat
	}
	if next.FontName.Present {
		current.FontName = next.FontName
	}
	if next.FontSizePoints.Present {
		current.FontSizePoints = next.FontSizePoints
	}
	if next.Bold.Present {
		current.Bold = next.Bold
	}
	if next.Italic.Present {
		current.Italic = next.Italic
	}
	if next.FontColor.Present {
		current.FontColor = next.FontColor
	}
	if next.FillColor.Present {
		current.FillColor = next.FillColor
	}
	if next.HorizontalAlignment.Present {
		current.HorizontalAlignment = next.HorizontalAlignment
	}
	if next.VerticalAlignment.Present {
		current.VerticalAlignment = next.VerticalAlignment
	}
	if next.WrapText.Present {
		current.WrapText = next.WrapText
	}
	return current
}

func locateStylesPart(index *opcPackageIndex, read func(string) ([]byte, bool), workbook workbookPartLocation) (string, error) {
	relsXML, ok := read(workbook.relsPart)
	if !ok {
		return "", fmt.Errorf("missing workbook relationships")
	}
	relationships, err := parseRoutingRelationships(relsXML)
	if err != nil {
		return "", fmt.Errorf("parse workbook relationships: %w", err)
	}
	var match *routingRelationship
	expectedType, opposingType := relTypeStylesTransitional, relTypeStylesStrict
	if workbook.strict {
		expectedType, opposingType = relTypeStylesStrict, relTypeStylesTransitional
	}
	for position := range relationships {
		candidate := &relationships[position]
		if candidate.relType == opposingType {
			return "", fmt.Errorf("workbook and styles relationship use opposing Strict/Transitional dialects")
		}
		if candidate.relType != expectedType {
			continue
		}
		if match != nil {
			return "", fmt.Errorf("multiple workbook styles relationships")
		}
		match = candidate
	}
	if match == nil {
		return "", fmt.Errorf("workbook has no styles relationship")
	}
	switch match.targetMode {
	case "", "Internal":
	case "External":
		return "", fmt.Errorf("styles relationship is external")
	default:
		return "", fmt.Errorf("styles relationship has unsupported target mode %q", match.targetMode)
	}
	resolved, err := resolveRelPath(workbook.baseDir, match.target)
	if err != nil {
		return "", fmt.Errorf("styles relationship target %q: %w", match.target, err)
	}
	part, _, ok := index.lookupResolved(resolved)
	if !ok {
		return "", fmt.Errorf("styles relationship targets missing part %q", resolved)
	}
	contentTypesPart, _, ok := index.lookupSpelling("[Content_Types].xml")
	if !ok {
		return "", fmt.Errorf("missing [Content_Types].xml")
	}
	contentTypes, ok := read(contentTypesPart)
	if !ok {
		return "", fmt.Errorf("unreadable [Content_Types].xml")
	}
	if err := validateStylesContentType(contentTypes, resolved); err != nil {
		return "", err
	}
	return part, nil
}

func validateStylesContentType(data []byte, stylePart string) error {
	targetKey, err := canonicalOPCPartKey(stylePart)
	if err != nil {
		return fmt.Errorf("invalid styles part %q: %w", stylePart, err)
	}
	decoder := xml.NewDecoder(bytes.NewReader(data))
	extension := stylePartExtension(stylePart)
	if extension == "" {
		return fmt.Errorf("styles part %q has no content-type extension", stylePart)
	}
	depth := 0
	rootSeen, rootClosed := false, false
	overrides := make(map[string]string)
	defaults := make(map[string]string)
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return fmt.Errorf("parse content types: %w", err)
		}
		switch token := token.(type) {
		case xml.StartElement:
			depth++
			if depth == 1 {
				if rootSeen {
					return fmt.Errorf("content types has multiple root elements")
				}
				if token.Name.Space != contentTypesNamespace || token.Name.Local != "Types" {
					return fmt.Errorf("content types root is not package Types")
				}
				rootSeen = true
				continue
			}
			if depth != 2 || token.Name.Space != contentTypesNamespace || (token.Name.Local != "Override" && token.Name.Local != "Default") {
				return fmt.Errorf("content types has unsupported direct or nested element %q", token.Name.Local)
			}
			contentType, typeFound, attrErr := unqualifiedXMLAttribute(token, "ContentType")
			if attrErr != nil {
				return attrErr
			}
			if !typeFound || contentType == "" {
				return fmt.Errorf("content-type declaration requires ContentType")
			}
			if token.Name.Local == "Override" {
				partName, found, attrErr := unqualifiedXMLAttribute(token, "PartName")
				if attrErr != nil {
					return attrErr
				}
				if !found || partName == "" {
					return fmt.Errorf("content-type Override requires PartName")
				}
				key, keyErr := canonicalOPCPartKey(partName)
				if keyErr != nil {
					return fmt.Errorf("invalid content-type PartName %q: %w", partName, keyErr)
				}
				if _, duplicate := overrides[key]; duplicate {
					return fmt.Errorf("duplicate content-type Override for %q", partName)
				}
				overrides[key] = contentType
			} else {
				extensionValue, found, attrErr := unqualifiedXMLAttribute(token, "Extension")
				if attrErr != nil {
					return attrErr
				}
				if !found || extensionValue == "" {
					return fmt.Errorf("content-type Default requires Extension")
				}
				key := asciiLower(extensionValue)
				if _, duplicate := defaults[key]; duplicate {
					return fmt.Errorf("duplicate content-type Default for extension %q", extensionValue)
				}
				defaults[key] = contentType
			}
		case xml.EndElement:
			if depth == 1 && token.Name.Space == contentTypesNamespace && token.Name.Local == "Types" {
				rootClosed = true
			}
			depth--
		case xml.CharData:
			if depth > 0 && len(bytes.TrimSpace(token)) != 0 {
				return fmt.Errorf("content types has unsupported text")
			}
		case xml.ProcInst:
			if depth == 0 && token.Target == "xml" {
				continue
			}
			return fmt.Errorf("content types has unsupported processing instruction")
		case xml.Directive:
			return fmt.Errorf("content types has unsupported directive")
		}
	}
	if depth != 0 || !rootSeen || !rootClosed {
		return fmt.Errorf("missing complete content types root")
	}
	effective, found := overrides[targetKey]
	if !found {
		effective, found = defaults[asciiLower(extension)]
	}
	if !found {
		return fmt.Errorf("styles part has no effective content type")
	}
	if !asciiEqualFold(effective, stylesPartContentType) {
		return fmt.Errorf("styles part has unexpected content type %q", effective)
	}
	return nil
}

func stylePartExtension(part string) string {
	base := part
	if slash := strings.LastIndexByte(base, '/'); slash >= 0 {
		base = base[slash+1:]
	}
	if dot := strings.LastIndexByte(base, '.'); dot >= 0 && dot+1 < len(base) {
		return base[dot+1:]
	}
	return ""
}
