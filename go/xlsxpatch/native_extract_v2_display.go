package xlsxpatch

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"math"
	"strconv"
	"strings"
)

const (
	relTypeThemeTransitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
	relTypeThemeStrict       = "http://purl.oclc.org/ooxml/officeDocument/relationships/theme"
	drawingMLNamespace       = "http://schemas.openxmlformats.org/drawingml/2006/main"
	themePartContentType     = "application/vnd.openxmlformats-officedocument.theme+xml"
)

var spreadsheetThemeColorSlots = []string{"lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"}

type nativeWorkbookThemeDisplay struct {
	colors     map[int]string
	majorLatin string
	minorLatin string
}

func (theme nativeWorkbookThemeDisplay) typeface(scheme string) string {
	switch scheme {
	case "major":
		return theme.majorLatin
	case "minor":
		return theme.minorLatin
	default:
		return ""
	}
}

func (theme nativeWorkbookThemeDisplay) resolve(index int, tint *float64) (string, bool) {
	rgb, ok := theme.colors[index]
	if !ok || rgb == "" {
		return "", false
	}
	if tint == nil {
		return rgb, true
	}
	return applyLinearSpreadsheetTint(rgb, *tint), true
}

func enrichNativeWorkbookV2Display(data []byte, workbook *NativeWorkbookV2) error {
	pkg, err := openNativeWorkbookPackage(data)
	if err != nil {
		return err
	}
	location, err := locateWorkbookPartBytes(pkg.index, func(name string) ([]byte, bool) {
		value, ok := pkg.files[name]
		return value, ok
	})
	if err != nil {
		return fmt.Errorf("xlsxpatch: native v2 display: %w", err)
	}
	namespace, relNamespace := spreadsheetMLTransitional, officeRelNamespaceTransitional
	expectedTheme, opposingTheme := relTypeThemeTransitional, relTypeThemeStrict
	if location.strict {
		namespace, relNamespace = spreadsheetMLStrict, officeRelNamespaceStrict
		expectedTheme, opposingTheme = relTypeThemeStrict, relTypeThemeTransitional
	}
	extractor := &nativeWorkbookExtractor{
		pkg: pkg, workbook: location, namespace: namespace, relNamespace: relNamespace,
		modeled: map[string]bool{}, unsupportedKeys: map[string]bool{}, claimedXML: map[string]bool{},
	}
	date1904, err := parseWorkbookDate1904(pkg.files[location.part], namespace)
	if err != nil {
		return err
	}
	workbook.Date1904 = nativeWorkbookBool(date1904)

	theme, _ := parseWorkbookThemeDisplay(extractor, expectedTheme, opposingTheme)

	expectedStyles, opposingStyles := relTypeStylesTransitional, relTypeStylesStrict
	if location.strict {
		expectedStyles, opposingStyles = relTypeStylesStrict, relTypeStylesTransitional
	}
	stylesPart, err := extractor.relatedCorePart(expectedStyles, opposingStyles, stylesPartContentType, "styles", false)
	if err != nil {
		return err
	}
	if stylesPart != "" {
		registry, registryErr := newStyleRegistry(pkg.files[stylesPart])
		if registryErr != nil {
			return fmt.Errorf("xlsxpatch: native v2 display: styles %q: %w", stylesPart, registryErr)
		}
		if err := enrichNativeWorkbookV2Styles(workbook, registry, theme); err != nil {
			return err
		}
	}
	enrichNativeWorkbookV2RichRuns(workbook, theme)
	return nil
}

func parseWorkbookDate1904(data []byte, namespace string) (bool, error) {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	depth := 0
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			return false, nil
		}
		if err != nil {
			return false, fmt.Errorf("xlsxpatch: native v2 display: parse workbookPr: %w", err)
		}
		switch token := token.(type) {
		case xml.StartElement:
			depth++
			if depth == 2 && token.Name.Space == namespace && token.Name.Local == "workbookPr" {
				raw, found, attrErr := unqualifiedXMLAttribute(token, "date1904")
				if attrErr != nil {
					return false, attrErr
				}
				if !found {
					return false, nil
				}
				value, boolErr := ooxmlBoolean(raw, true)
				if boolErr != nil {
					return false, fmt.Errorf("xlsxpatch: native v2 display: workbookPr date1904=%q is invalid", raw)
				}
				return value, nil
			}
		case xml.EndElement:
			depth--
		}
	}
}

func parseWorkbookThemeDisplay(extractor *nativeWorkbookExtractor, expected, opposing string) (nativeWorkbookThemeDisplay, error) {
	theme := nativeWorkbookThemeDisplay{colors: map[int]string{}}
	part, err := extractor.relatedCorePart(expected, opposing, themePartContentType, "theme", false)
	if err != nil || part == "" {
		return theme, err
	}
	data := extractor.pkg.files[part]
	decoder := xml.NewDecoder(bytes.NewReader(data))
	depth, schemeDepth, slotDepth, fontDepth, fontKindDepth := 0, 0, 0, 0, 0
	slotName, fontKind := "", ""
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			return theme, nil
		}
		if err != nil {
			return nativeWorkbookThemeDisplay{}, fmt.Errorf("xlsxpatch: native v2 display: parse theme: %w", err)
		}
		switch token := token.(type) {
		case xml.StartElement:
			depth++
			if depth == 3 && token.Name.Space == drawingMLNamespace && token.Name.Local == "clrScheme" {
				schemeDepth = depth
				continue
			}
			if depth == 3 && token.Name.Space == drawingMLNamespace && token.Name.Local == "fontScheme" {
				fontDepth = depth
				continue
			}
			if schemeDepth != 0 && depth == schemeDepth+1 && token.Name.Space == drawingMLNamespace {
				slotName, slotDepth = token.Name.Local, depth
				continue
			}
			if fontDepth != 0 && depth == fontDepth+1 && token.Name.Space == drawingMLNamespace && (token.Name.Local == "majorFont" || token.Name.Local == "minorFont") {
				fontKind, fontKindDepth = token.Name.Local, depth
				continue
			}
			if fontKindDepth != 0 && depth == fontKindDepth+1 && token.Name.Space == drawingMLNamespace && token.Name.Local == "latin" {
				typeface, _, typeErr := unqualifiedXMLAttribute(token, "typeface")
				if typeErr != nil {
					return nativeWorkbookThemeDisplay{}, typeErr
				}
				if err := skipNativeXMLElement(decoder, token, 1); err != nil {
					return nativeWorkbookThemeDisplay{}, err
				}
				depth--
				if typeface != "" {
					if fontKind == "majorFont" {
						theme.majorLatin = typeface
					} else if fontKind == "minorFont" {
						theme.minorLatin = typeface
					}
				}
				continue
			}
			if slotDepth != 0 && depth == slotDepth+1 && token.Name.Space == drawingMLNamespace && (token.Name.Local == "srgbClr" || token.Name.Local == "sysClr") {
				rgb, colorErr := parseDrawingMLThemeColor(decoder, token)
				if colorErr != nil {
					return nativeWorkbookThemeDisplay{}, colorErr
				}
				depth--
				if rgb == "" {
					slotName = ""
					continue
				}
				if index := spreadsheetThemeSlotIndex(slotName); index >= 0 {
					theme.colors[index] = rgb
				}
			} else if slotDepth != 0 && depth == slotDepth+1 {
				slotName = ""
			}
		case xml.EndElement:
			if slotDepth != 0 && depth == slotDepth {
				slotDepth, slotName = 0, ""
			}
			if schemeDepth != 0 && depth == schemeDepth {
				schemeDepth = 0
			}
			if fontKindDepth != 0 && depth == fontKindDepth {
				fontKindDepth, fontKind = 0, ""
			}
			if fontDepth != 0 && depth == fontDepth {
				fontDepth = 0
			}
			depth--
		}
	}
}

func parseDrawingMLThemeColor(decoder *xml.Decoder, root xml.StartElement) (string, error) {
	base := ""
	switch root.Name.Local {
	case "srgbClr":
		hex, found, err := unqualifiedXMLAttribute(root, "val")
		if err != nil {
			return "", err
		}
		hex = strings.ToUpper(hex)
		if !found || !drawingMLHex6(hex) || !drawingMLColorAttributesOnly(root, "val") {
			if err := skipNativeXMLElement(decoder, root, 1); err != nil {
				return "", err
			}
			return "", nil
		}
		base = hex
	case "sysClr":
		last, found, err := unqualifiedXMLAttribute(root, "lastClr")
		if err != nil {
			return "", err
		}
		last = strings.ToUpper(last)
		if !found || !drawingMLHex6(last) {
			if err := skipNativeXMLElement(decoder, root, 1); err != nil {
				return "", err
			}
			return "", nil
		}
		base = last
	default:
		if err := skipNativeXMLElement(decoder, root, 1); err != nil {
			return "", err
		}
		return "", nil
	}
	var tint, shade *int
	level := 1
	for level > 0 {
		token, err := decoder.Token()
		if err != nil {
			return "", err
		}
		switch token := token.(type) {
		case xml.StartElement:
			level++
			if level == 2 && token.Name.Space == drawingMLNamespace && (token.Name.Local == "tint" || token.Name.Local == "shade") {
				raw, found, attrErr := unqualifiedXMLAttribute(token, "val")
				empty, skipErr := consumeEmptyXMLElement(decoder, token)
				if attrErr != nil {
					return "", attrErr
				}
				if skipErr != nil {
					return "", skipErr
				}
				level--
				percent, ok := parseDrawingMLPercent(raw)
				if !found || !empty || !ok || !drawingMLColorAttributesOnly(token, "val") {
					if err := skipNativeXMLElementRemaining(decoder, root, level); err != nil {
						return "", err
					}
					return "", nil
				}
				if token.Name.Local == "tint" {
					if tint != nil {
						if err := skipNativeXMLElementRemaining(decoder, root, level); err != nil {
							return "", err
						}
						return "", nil
					}
					tint = intPointer(percent)
				} else {
					if shade != nil {
						if err := skipNativeXMLElementRemaining(decoder, root, level); err != nil {
							return "", err
						}
						return "", nil
					}
					shade = intPointer(percent)
				}
				continue
			}
			if err := skipNativeXMLElement(decoder, token, 1); err != nil {
				return "", err
			}
			level--
			if err := skipNativeXMLElementRemaining(decoder, root, level); err != nil {
				return "", err
			}
			return "", nil
		case xml.EndElement:
			if level == 1 && token.Name != root.Name {
				return "", fmt.Errorf("xlsxpatch: native v2 display: theme color has mismatched closing element")
			}
			level--
		case xml.CharData:
			if len(bytes.TrimSpace(token)) != 0 {
				return "", nil
			}
		}
	}
	if tint != nil && shade != nil {
		return "", nil
	}
	return applyLinearDrawingMLTintShade("#"+base, tint, shade), nil
}

func skipNativeXMLElementRemaining(decoder *xml.Decoder, root xml.StartElement, level int) error {
	for level > 0 {
		token, err := decoder.Token()
		if err != nil {
			return err
		}
		switch token := token.(type) {
		case xml.StartElement:
			level++
		case xml.EndElement:
			if level == 1 && token.Name != root.Name {
				return fmt.Errorf("xlsxpatch: native v2 display: theme color has mismatched closing element")
			}
			level--
		}
	}
	return nil
}

func drawingMLHex6(value string) bool {
	if len(value) != 6 {
		return false
	}
	return strings.TrimLeft(value, "0123456789ABCDEF") == ""
}

func drawingMLColorAttributesOnly(start xml.StartElement, allowed ...string) bool {
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

func parseDrawingMLPercent(raw string) (int, bool) {
	if raw == "" {
		return 0, false
	}
	for index, character := range raw {
		if character == '-' && index == 0 {
			continue
		}
		if character < '0' || character > '9' {
			return 0, false
		}
	}
	value, err := strconv.Atoi(raw)
	if err != nil || value < -100000 || value > 100000 {
		return 0, false
	}
	return value, true
}

func applyLinearDrawingMLTintShade(rgb string, tint, shade *int) string {
	if tint == nil && shade == nil {
		return rgb
	}
	if shade != nil {
		amount := float64(*shade) / 100000
		return applyLinearRGB(rgb, func(channel int) int {
			return roundColorChannel(float64(channel) * amount)
		})
	}
	amount := float64(*tint) / 100000
	return applyLinearSpreadsheetTint(rgb, amount)
}

func applyLinearSpreadsheetTint(rgb string, tint float64) string {
	return applyLinearRGB(rgb, func(channel int) int {
		if tint < 0 {
			return roundColorChannel(float64(channel) * (1 + tint))
		}
		return roundColorChannel(float64(channel)*(1-tint) + 255*tint)
	})
}

func applyLinearRGB(rgb string, each func(int) int) string {
	if len(rgb) != 7 || rgb[0] != '#' {
		return rgb
	}
	parse := func(start int) int {
		value, _ := strconv.ParseUint(rgb[start:start+2], 16, 8)
		return int(value)
	}
	r, g, b := each(parse(1)), each(parse(3)), each(parse(5))
	return fmt.Sprintf("#%02X%02X%02X", r, g, b)
}

func roundColorChannel(value float64) int {
	if value < 0 {
		return 0
	}
	if value > 255 {
		return 255
	}
	return int(math.Floor(value + 0.5))
}

func consumeEmptyXMLElement(decoder *xml.Decoder, root xml.StartElement) (bool, error) {
	empty, level := true, 1
	for level > 0 {
		token, err := decoder.Token()
		if err != nil {
			return false, err
		}
		switch token := token.(type) {
		case xml.StartElement:
			empty = false
			level++
		case xml.EndElement:
			if level == 1 && token.Name != root.Name {
				return false, fmt.Errorf("xlsxpatch: native v2 display: theme color has mismatched closing element")
			}
			level--
		case xml.CharData:
			if len(bytes.TrimSpace(token)) != 0 {
				empty = false
			}
		}
	}
	return empty, nil
}

func spreadsheetThemeSlotIndex(name string) int {
	for index, slot := range spreadsheetThemeColorSlots {
		if slot == name {
			return index
		}
	}
	return -1
}

func enrichNativeWorkbookV2Styles(workbook *NativeWorkbookV2, registry *styleRegistry, theme nativeWorkbookThemeDisplay) error {
	if workbook.NormalStyle != nil && int(workbook.NormalStyle.FontID) >= 0 && int(workbook.NormalStyle.FontID) < len(registry.fonts) {
		font := registry.fonts[workbook.NormalStyle.FontID]
		if font.scheme != nil {
			if typeface := theme.typeface(*font.scheme); typeface != "" {
				workbook.NormalStyle.FontName = typeface
			}
		}
	}
	if len(workbook.Styles) != len(registry.cellXfs) {
		return nil
	}
	for index := range workbook.Styles {
		style := &workbook.Styles[index]
		source := registry.cellXfs[index]
		base := effectiveCellStyleXF(registry.styleXfs[source.xfID])
		fontID := effectiveStyleComponent(source.fontID, base.fontID, source.applyFont)
		alignment := effectiveStyleAlignment(source, base)
		font := registry.fonts[fontID]
		changed := false
		if font.scheme != nil {
			if typeface := theme.typeface(*font.scheme); typeface != "" && (style.Effective.FontName == nil || *style.Effective.FontName != typeface) {
				style.Effective.FontName = nativeWorkbookString(typeface)
				changed = true
			}
		}
		if containsNativeStyleUnsupported(style.Effective.Unsupported, "font-color") && font.themeIndex != nil {
			if rgb, ok := theme.resolve(*font.themeIndex, font.themeTint); ok {
				style.Effective.FontColor = nativeWorkbookString(rgb)
				style.Effective.Unsupported = removeNativeStyleUnsupported(style.Effective.Unsupported, "font-color")
				changed = true
			}
		}
		if containsNativeStyleUnsupported(style.Effective.Unsupported, "alignment-extended") && alignment.otherAttrsExceptPaintExact == "" {
			mapped := mapSpreadsheetTextRotation(alignment.textRotation)
			if alignment.textRotation != nil && mapped == nil {
				// unmodeled rotation angle remains alignment-extended
			} else {
				if alignment.shrinkToFit != nil && *alignment.shrinkToFit {
					style.Effective.ShrinkToFit = nativeWorkbookBool(true)
				}
				if mapped != nil && *mapped != 0 {
					style.Effective.TextRotation = mapped
				}
				style.Effective.Unsupported = removeNativeStyleUnsupported(style.Effective.Unsupported, "alignment-extended")
				changed = true
			}
		}
		if !changed {
			continue
		}
		if len(style.Effective.Unsupported) == 0 {
			style.Effective.Projection = "full"
		} else {
			style.Effective.Projection = "partial"
		}
		digest, err := nativeRawStyleProjectionDigestV2(style.Effective)
		if err != nil {
			return err
		}
		style.RawProjectionSHA256 = digest
	}
	workbook.Unsupported = filterNativeWorkbookUnsupportedV2(workbook.Unsupported, workbook)
	return nil
}

func enrichNativeWorkbookV2RichRuns(workbook *NativeWorkbookV2, theme nativeWorkbookThemeDisplay) {
	for sheetIndex := range workbook.Sheets {
		for cellIndex := range workbook.Sheets[sheetIndex].Cells {
			cell := &workbook.Sheets[sheetIndex].Cells[cellIndex]
			if cell.Value != nil {
				enrichNativeRichRunSlice(cell.Value.Runs, theme)
			}
			if cell.Formula != nil && cell.Formula.Cached != nil {
				enrichNativeRichRunSlice(cell.Formula.Cached.Runs, theme)
			}
		}
	}
}

func enrichNativeRichRunSlice(runs []NativeWorkbookRichRunV2, theme nativeWorkbookThemeDisplay) {
	for index := range runs {
		run := &runs[index]
		if run.scheme != nil {
			if typeface := theme.typeface(*run.scheme); typeface != "" {
				run.FontName = nativeWorkbookString(typeface)
			}
		}
		if run.FontColor == nil && run.themeIndex != nil {
			if rgb, ok := theme.resolve(*run.themeIndex, run.themeTint); ok {
				run.FontColor = nativeWorkbookString(rgb)
			}
		}
	}
}

func mapSpreadsheetTextRotation(value *int) *int {
	if value == nil {
		return nil
	}
	switch *value {
	case 0:
		return intPointer(0)
	case 90:
		return intPointer(90)
	case 180:
		return intPointer(270)
	default:
		return nil
	}
}

func containsNativeStyleUnsupported(values []string, code string) bool {
	for _, value := range values {
		if value == code {
			return true
		}
	}
	return false
}

func removeNativeStyleUnsupported(values []string, code string) []string {
	out := make([]string, 0, len(values))
	for _, value := range values {
		if value != code {
			out = append(out, value)
		}
	}
	if len(out) == 0 {
		return []string{}
	}
	return out
}

func filterNativeWorkbookUnsupportedV2(items []NativeWorkbookUnsupportedV2, workbook *NativeWorkbookV2) []NativeWorkbookUnsupportedV2 {
	needed := map[string]bool{}
	for index, style := range workbook.Styles {
		for _, code := range style.Effective.Unsupported {
			needed["STYLE_"+strings.ToUpper(strings.ReplaceAll(code, "-", "_"))+"\x00style:"+strconv.Itoa(index)] = true
		}
	}
	out := items[:0]
	for _, item := range items {
		if strings.HasPrefix(item.Code, "STYLE_") && (item.Code == "STYLE_FONT_COLOR" || item.Code == "STYLE_ALIGNMENT_EXTENDED") {
			if !needed[item.Code+"\x00"+item.ScopeID] {
				continue
			}
		}
		out = append(out, item)
	}
	if len(out) == 0 {
		return []NativeWorkbookUnsupportedV2{}
	}
	return out
}
