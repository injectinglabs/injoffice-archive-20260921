package xlsxpatch

import (
	"bytes"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"io"
	"sort"
	"strconv"
	"strings"
)

type nativeSharedString struct {
	text string
	rich bool
	runs []NativeWorkbookRichRunV2
}

func (extractor *nativeWorkbookExtractor) relatedCorePart(expectedType, opposingType, expectedContentType, label string, required bool) (string, error) {
	relationships, err := parseRoutingRelationships(extractor.pkg.files[extractor.workbook.relsPart])
	if err != nil {
		return "", fmt.Errorf("xlsxpatch: native extract: parse workbook relationships: %w", err)
	}
	var match *routingRelationship
	for index := range relationships {
		candidate := &relationships[index]
		if candidate.relType == opposingType {
			return "", fmt.Errorf("xlsxpatch: native extract: workbook and %s relationship use opposing Strict/Transitional dialects", label)
		}
		if candidate.relType != expectedType {
			continue
		}
		if match != nil {
			return "", fmt.Errorf("xlsxpatch: native extract: multiple workbook %s relationships", label)
		}
		match = candidate
	}
	if match == nil {
		if required {
			return "", fmt.Errorf("xlsxpatch: native extract: workbook has no %s relationship", label)
		}
		return "", nil
	}
	if match.targetMode != "" && match.targetMode != "Internal" {
		return "", fmt.Errorf("xlsxpatch: native extract: %s relationship has unsupported TargetMode %q", label, match.targetMode)
	}
	resolved, err := resolveRelPath(extractor.workbook.baseDir, match.target)
	if err != nil {
		return "", fmt.Errorf("xlsxpatch: native extract: %s relationship target %q: %w", label, match.target, err)
	}
	part, _, found := extractor.pkg.index.lookupResolved(resolved)
	if !found {
		return "", fmt.Errorf("xlsxpatch: native extract: %s relationship targets missing part %q", label, resolved)
	}
	if err := requireNativeContentType(extractor.pkg, part, expectedContentType); err != nil {
		return "", fmt.Errorf("xlsxpatch: native extract: %s: %w", label, err)
	}
	return part, nil
}

func (extractor *nativeWorkbookExtractor) extractSharedStrings() ([]nativeSharedString, string, error) {
	expected, opposing := relTypeSharedStringsTransitional, relTypeSharedStringsStrict
	if extractor.workbook.strict {
		expected, opposing = relTypeSharedStringsStrict, relTypeSharedStringsTransitional
	}
	part, err := extractor.relatedCorePart(expected, opposing, nativeSharedStringsType, "sharedStrings", false)
	if err != nil || part == "" {
		return nil, part, err
	}
	data := extractor.pkg.files[part]
	if err := extractor.claimCoreXML(part); err != nil {
		return nil, part, err
	}
	if len(data) == 0 || len(data) > NativeXLSXMaxXMLPartBytes {
		return nil, part, fmt.Errorf("xlsxpatch: native extract: sharedStrings part %q size must be 1..%d bytes", part, NativeXLSXMaxXMLPartBytes)
	}
	decoder := xml.NewDecoder(bytes.NewReader(data))
	root, err := nativeReadXMLRoot(decoder)
	if err != nil {
		return nil, part, fmt.Errorf("xlsxpatch: native extract: sharedStrings %q: %w", part, err)
	}
	if root.Name != (xml.Name{Space: extractor.namespace, Local: "sst"}) {
		return nil, part, fmt.Errorf("xlsxpatch: native extract: sharedStrings root uses an unexpected namespace")
	}
	if len(unexpectedSemanticXMLAttributes(root, xml.Name{Local: "count"}, xml.Name{Local: "uniqueCount"})) != 0 {
		if err := extractor.addUnsupported("SHARED_STRING_TABLE_ATTRIBUTES", "rich-text", "workbook", part, "", "unmodeled shared-string table attributes remain authority-bound to the source package"); err != nil {
			return nil, part, err
		}
	}
	declaredCount, countFound, err := optionalNativeUintAttribute(root, "count", uint64(NativeXLSXMaxCells))
	if err != nil {
		return nil, part, fmt.Errorf("xlsxpatch: native extract: sharedStrings count: %w", err)
	}
	declaredUnique, uniqueFound, err := optionalNativeUintAttribute(root, "uniqueCount", uint64(NativeXLSXMaxCells))
	if err != nil {
		return nil, part, fmt.Errorf("xlsxpatch: native extract: sharedStrings uniqueCount: %w", err)
	}
	values := make([]nativeSharedString, 0)
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			return nil, part, fmt.Errorf("xlsxpatch: native extract: sharedStrings has no closing sst")
		}
		if err != nil {
			return nil, part, fmt.Errorf("xlsxpatch: native extract: parse sharedStrings: %w", err)
		}
		switch token := token.(type) {
		case xml.StartElement:
			if token.Name == (xml.Name{Space: extractor.namespace, Local: "extLst"}) {
				if err := extractor.addUnsupported("SHARED_STRING_TABLE_OPAQUE_CONTENT", "rich-text", "workbook", part, "", "shared-string extensions remain authority-bound to the source package"); err != nil {
					return nil, part, err
				}
				if err := skipNativeXMLElement(decoder, token, 1); err != nil {
					return nil, part, err
				}
				continue
			}
			if token.Name != (xml.Name{Space: extractor.namespace, Local: "si"}) {
				return nil, part, fmt.Errorf("xlsxpatch: native extract: sharedStrings has unsupported direct child {%s}%s", token.Name.Space, token.Name.Local)
			}
			if len(unexpectedSemanticXMLAttributes(token)) != 0 {
				if err := extractor.addUnsupported("SHARED_STRING_ITEM_ATTRIBUTES", "rich-text", "workbook", part, "", "unmodeled shared-string item attributes remain authority-bound to the source package"); err != nil {
					return nil, part, err
				}
			}
			value, parseErr := parseNativeStringItem(decoder, token, extractor.namespace)
			if parseErr != nil {
				return nil, part, fmt.Errorf("xlsxpatch: native extract: shared string %d: %w", len(values), parseErr)
			}
			values = append(values, value)
			if err := extractor.claimNativeText(value.text, maxCellTextLen); err != nil {
				return nil, part, fmt.Errorf("xlsxpatch: native extract: shared string %d: %w", len(values)-1, err)
			}
			if len(values) > NativeXLSXMaxCells {
				return nil, part, fmt.Errorf("xlsxpatch: native extract: shared string table exceeds resource limits")
			}
			if value.rich {
				if err := extractor.addUnsupported("RICH_SHARED_STRING", "rich-text", "workbook", part, "", "rich shared-string formatting is preserved exactly and exposed as read-only text"); err != nil {
					return nil, part, err
				}
			}
		case xml.EndElement:
			if token.Name != root.Name {
				return nil, part, fmt.Errorf("xlsxpatch: native extract: sharedStrings has mismatched closing element")
			}
			if uniqueFound && declaredUnique != uint64(len(values)) {
				return nil, part, fmt.Errorf("xlsxpatch: native extract: sharedStrings uniqueCount=%d does not match %d si elements", declaredUnique, len(values))
			}
			if countFound && declaredCount < uint64(len(values)) {
				return nil, part, fmt.Errorf("xlsxpatch: native extract: sharedStrings count=%d is smaller than uniqueCount %d", declaredCount, len(values))
			}
			if err := nativeRequireXMLEOF(decoder); err != nil {
				return nil, part, err
			}
			return values, part, nil
		case xml.CharData:
			if len(bytes.TrimSpace(token)) != 0 {
				return nil, part, fmt.Errorf("xlsxpatch: native extract: sharedStrings contains direct text")
			}
		case xml.ProcInst:
			return nil, part, fmt.Errorf("xlsxpatch: native extract: sharedStrings contains unsupported processing instruction %q", token.Target)
		case xml.Directive:
			return nil, part, fmt.Errorf("xlsxpatch: native extract: sharedStrings contains unsupported XML directive")
		}
	}
}

func parseNativeStringItem(decoder *xml.Decoder, root xml.StartElement, namespace string) (nativeSharedString, error) {
	result := nativeSharedString{}
	plainSeen := false
	for {
		token, err := decoder.Token()
		if err != nil {
			return result, err
		}
		switch token := token.(type) {
		case xml.StartElement:
			if token.Name.Space != namespace {
				result.rich = true
				if err := skipNativeXMLElement(decoder, token, 1); err != nil {
					return result, err
				}
				continue
			}
			switch token.Name.Local {
			case "t":
				if err := requireOnlySemanticXMLAttributes(token, xml.Name{Space: "http://www.w3.org/XML/1998/namespace", Local: "space"}); err != nil {
					return result, err
				}
				if plainSeen {
					result.rich = true
				}
				text, err := readNativeTextElement(decoder, token)
				if err != nil {
					return result, err
				}
				result.text += text
				if text != "" {
					result.runs = append(result.runs, NativeWorkbookRichRunV2{Text: text})
				}
				plainSeen = true
			case "r":
				if err := requireOnlySemanticXMLAttributes(token); err != nil {
					return result, err
				}
				result.rich = true
				run, parseErr := parseNativeRichStringRun(decoder, token, namespace)
				if parseErr != nil {
					return result, parseErr
				}
				result.text += run.Text
				if run.Text != "" {
					result.runs = append(result.runs, run.value)
				}
			case "rPh", "phoneticPr", "extLst":
				result.rich = true
				if err := skipNativeXMLElement(decoder, token, 1); err != nil {
					return result, err
				}
			default:
				result.rich = true
				if err := skipNativeXMLElement(decoder, token, 1); err != nil {
					return result, err
				}
			}
		case xml.EndElement:
			if token.Name != root.Name {
				return result, fmt.Errorf("string item has mismatched closing element")
			}
			result.runs = reconstructNativeRichRuns(result.rich, result.text, result.runs)
			return result, nil
		case xml.CharData:
			if len(bytes.TrimSpace(token)) != 0 {
				return result, fmt.Errorf("string item contains unsupported direct text")
			}
		case xml.ProcInst:
			return result, fmt.Errorf("string item contains unsupported processing instruction %q", token.Target)
		case xml.Directive:
			return result, fmt.Errorf("string item contains unsupported XML directive")
		}
	}
}

func reconstructNativeRichRuns(rich bool, text string, runs []NativeWorkbookRichRunV2) []NativeWorkbookRichRunV2 {
	if !rich {
		return nil
	}
	joined := ""
	for _, run := range runs {
		joined += run.Text
	}
	if len(runs) != 0 && joined == text {
		return runs
	}
	if text == "" {
		return nil
	}
	return []NativeWorkbookRichRunV2{{Text: text}}
}

type nativeParsedRichRun struct {
	Text  string
	value NativeWorkbookRichRunV2
	exact bool
}

func parseNativeRichStringRun(decoder *xml.Decoder, root xml.StartElement, namespace string) (nativeParsedRichRun, error) {
	result := nativeParsedRichRun{exact: true}
	seenText := false
	seenProps := false
	for {
		token, err := decoder.Token()
		if err != nil {
			return result, err
		}
		switch token := token.(type) {
		case xml.StartElement:
			if token.Name == (xml.Name{Space: namespace, Local: "t"}) {
				if err := requireOnlySemanticXMLAttributes(token, xml.Name{Space: "http://www.w3.org/XML/1998/namespace", Local: "space"}); err != nil {
					return result, err
				}
				if seenText {
					return result, fmt.Errorf("rich string run has multiple text children")
				}
				value, err := readNativeTextElement(decoder, token)
				if err != nil {
					return result, err
				}
				result.Text, seenText = value, true
				result.value.Text = value
				continue
			}
			if token.Name == (xml.Name{Space: namespace, Local: "rPr"}) {
				if seenProps {
					return result, fmt.Errorf("rich string run has multiple rPr children")
				}
				seenProps = true
				props, err := parseNativeRichRunProperties(decoder, token, namespace)
				if err != nil {
					return result, err
				}
				result.value.FontName = props.FontName
				result.value.Bold = props.Bold
				result.value.Italic = props.Italic
				result.value.FontSizePoints = props.FontSizePoints
				result.value.FontColor = props.FontColor
				result.value.scheme = props.scheme
				result.value.themeIndex = props.themeIndex
				result.value.themeTint = props.themeTint
				continue
			}
			return result, fmt.Errorf("rich string run has unsupported child {%s}%s", token.Name.Space, token.Name.Local)
		case xml.EndElement:
			if token.Name != root.Name {
				return result, fmt.Errorf("rich string run has mismatched closing element")
			}
			if result.Text == "" {
				result.exact = false
			}
			return result, nil
		case xml.CharData:
			if len(bytes.TrimSpace(token)) != 0 {
				return result, fmt.Errorf("rich string run contains direct text")
			}
		case xml.ProcInst:
			return result, fmt.Errorf("rich string run contains processing instruction")
		case xml.Directive:
			return result, fmt.Errorf("rich string run contains XML directive")
		}
	}
}

func parseNativeRichRunProperties(decoder *xml.Decoder, root xml.StartElement, namespace string) (NativeWorkbookRichRunV2, error) {
	if err := requireOnlySemanticXMLAttributes(root); err != nil {
		return NativeWorkbookRichRunV2{}, err
	}
	result := NativeWorkbookRichRunV2{}
	seen := map[string]bool{}
	for {
		token, err := decoder.Token()
		if err != nil {
			return result, err
		}
		switch token := token.(type) {
		case xml.StartElement:
			if token.Name.Space != namespace {
				if err := skipNativeXMLElement(decoder, token, 1); err != nil {
					return result, err
				}
				continue
			}
			switch token.Name.Local {
			case "rFont", "sz", "b", "i", "color", "scheme":
				if seen[token.Name.Local] {
					return result, fmt.Errorf("duplicate rPr %s", token.Name.Local)
				}
				seen[token.Name.Local] = true
			default:
				if err := skipNativeXMLElement(decoder, token, 1); err != nil {
					return result, err
				}
				continue
			}
			switch token.Name.Local {
			case "rFont":
				value, err := styleNodeValue(token, "val")
				if err != nil || value == "" {
					return result, fmt.Errorf("rFont requires a non-empty val")
				}
				if err := skipNativeXMLElement(decoder, token, 1); err != nil {
					return result, err
				}
				result.FontName = nativeWorkbookString(value)
			case "sz":
				value, err := styleNodeValue(token, "val")
				if err != nil {
					return result, err
				}
				parsed, parseErr := strconv.ParseFloat(value, 64)
				if parseErr != nil || parsed <= 0 {
					return result, fmt.Errorf("sz val=%q is invalid", value)
				}
				if err := skipNativeXMLElement(decoder, token, 1); err != nil {
					return result, err
				}
				result.FontSizePoints = cloneFloat(&parsed)
			case "b", "i":
				value, found, err := unqualifiedXMLAttribute(token, "val")
				if err != nil {
					return result, err
				}
				enabled := true
				if found {
					enabled, err = ooxmlBoolean(value, true)
					if err != nil {
						return result, err
					}
				}
				if err := skipNativeXMLElement(decoder, token, 1); err != nil {
					return result, err
				}
				if token.Name.Local == "b" {
					result.Bold = nativeWorkbookBool(enabled)
				} else {
					result.Italic = nativeWorkbookBool(enabled)
				}
			case "color":
				color, safe := supportedRGBStyleColor(token)
				if !safe {
					result.themeIndex, result.themeTint = exactThemeColor(token)
				} else {
					result.FontColor = color
				}
				if err := skipNativeXMLElement(decoder, token, 1); err != nil {
					return result, err
				}
			case "scheme":
				value, err := styleNodeValue(token, "val")
				if err := skipNativeXMLElement(decoder, token, 1); err != nil {
					return result, err
				}
				if err == nil && (value == "major" || value == "minor") {
					result.scheme = nativeWorkbookString(value)
				}
			}
		case xml.EndElement:
			if token.Name != root.Name {
				return result, fmt.Errorf("rPr has mismatched closing element")
			}
			return result, nil
		case xml.CharData:
			if len(bytes.TrimSpace(token)) != 0 {
				return result, fmt.Errorf("rPr contains direct text")
			}
		case xml.ProcInst:
			return result, fmt.Errorf("rPr contains processing instruction")
		case xml.Directive:
			return result, fmt.Errorf("rPr contains XML directive")
		}
	}
}

func readNativeTextElement(decoder *xml.Decoder, root xml.StartElement) (string, error) {
	if space, found, err := namespacedXMLAttribute(root, "http://www.w3.org/XML/1998/namespace", "space"); err != nil {
		return "", err
	} else if found && space != "default" && space != "preserve" {
		return "", fmt.Errorf("text element has invalid xml:space=%q", space)
	}
	var text strings.Builder
	for {
		token, err := decoder.Token()
		if err != nil {
			return "", err
		}
		switch token := token.(type) {
		case xml.CharData:
			text.Write(token)
		case xml.EndElement:
			if token.Name != root.Name {
				return "", fmt.Errorf("text element has mismatched closing element")
			}
			return decodeSpreadsheetString(text.String())
		case xml.StartElement:
			return "", fmt.Errorf("text element contains nested markup")
		case xml.ProcInst:
			return "", fmt.Errorf("text element contains processing instruction")
		case xml.Directive:
			return "", fmt.Errorf("text element contains XML directive")
		}
	}
}

func skipNativeXMLElement(decoder *xml.Decoder, root xml.StartElement, depth int) error {
	if depth > NativeXLSXMaxXMLDepth {
		return fmt.Errorf("XML nesting exceeds %d", NativeXLSXMaxXMLDepth)
	}
	level := 1
	for level > 0 {
		token, err := decoder.Token()
		if err != nil {
			return err
		}
		switch token := token.(type) {
		case xml.StartElement:
			level++
			if depth+level-1 > NativeXLSXMaxXMLDepth {
				return fmt.Errorf("XML nesting exceeds %d", NativeXLSXMaxXMLDepth)
			}
		case xml.EndElement:
			level--
		case xml.ProcInst:
			return fmt.Errorf("XML contains unsupported processing instruction %q", token.Target)
		case xml.Directive:
			return fmt.Errorf("XML contains unsupported directive")
		}
	}
	return nil
}

func nativeReadXMLRoot(decoder *xml.Decoder) (xml.StartElement, error) {
	for {
		token, err := decoder.Token()
		if err != nil {
			return xml.StartElement{}, err
		}
		switch token := token.(type) {
		case xml.StartElement:
			return token, nil
		case xml.CharData:
			if len(bytes.TrimSpace(token)) != 0 {
				return xml.StartElement{}, fmt.Errorf("XML contains text before its root")
			}
		case xml.ProcInst:
			if token.Target != "xml" {
				return xml.StartElement{}, fmt.Errorf("XML contains unsupported processing instruction %q", token.Target)
			}
		case xml.Directive:
			return xml.StartElement{}, fmt.Errorf("XML contains unsupported directive before root")
		}
	}
}

func nativeRequireXMLEOF(decoder *xml.Decoder) error {
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			return nil
		}
		if err != nil {
			return err
		}
		switch token := token.(type) {
		case xml.CharData:
			if len(bytes.TrimSpace(token)) == 0 {
				continue
			}
			return fmt.Errorf("XML has trailing text")
		case xml.Comment:
			continue
		default:
			return fmt.Errorf("XML has trailing token %T", token)
		}
	}
}

func optionalNativeUintAttribute(start xml.StartElement, name string, maximum uint64) (uint64, bool, error) {
	raw, found, err := unqualifiedXMLAttribute(start, name)
	if err != nil || !found {
		return 0, found, err
	}
	if raw == "" {
		return 0, true, fmt.Errorf("%s is empty", name)
	}
	for _, character := range raw {
		if character < '0' || character > '9' {
			return 0, true, fmt.Errorf("%s=%q is not an unsigned integer", name, raw)
		}
	}
	value, err := strconv.ParseUint(raw, 10, 64)
	if err != nil || value > maximum {
		return 0, true, fmt.Errorf("%s=%q exceeds %d", name, raw, maximum)
	}
	return value, true, nil
}

func (extractor *nativeWorkbookExtractor) extractStyles() ([]NativeWorkbookStyleV1, *NativeWorkbookNormalStyleV1, *styleRegistry, string, error) {
	expected, opposing := relTypeStylesTransitional, relTypeStylesStrict
	if extractor.workbook.strict {
		expected, opposing = relTypeStylesStrict, relTypeStylesTransitional
	}
	part, err := extractor.relatedCorePart(expected, opposing, stylesPartContentType, "styles", false)
	if err != nil {
		return nil, nil, nil, part, err
	}
	if part == "" {
		fallback := defaultNativeStyleProjection()
		digest, err := nativeRawStyleProjectionDigest(fallback)
		if err != nil {
			return nil, nil, nil, part, err
		}
		return []NativeWorkbookStyleV1{{ID: 0, Effective: fallback, RawProjectionSHA256: digest}}, nil, nil, "", nil
	}
	data := extractor.pkg.files[part]
	if err := extractor.claimCoreXML(part); err != nil {
		return nil, nil, nil, part, err
	}
	if len(data) == 0 || len(data) > NativeXLSXMaxXMLPartBytes {
		return nil, nil, nil, part, fmt.Errorf("xlsxpatch: native extract: styles part %q size must be 1..%d bytes", part, NativeXLSXMaxXMLPartBytes)
	}
	if err := extractor.inventoryNativeStylesMarkup(data, part); err != nil {
		return nil, nil, nil, part, err
	}
	registry, err := newStyleRegistry(data)
	if err != nil {
		return nil, nil, nil, part, fmt.Errorf("xlsxpatch: native extract: styles %q: %w", part, err)
	}
	if registry.index.namespace != extractor.namespace {
		return nil, nil, nil, part, fmt.Errorf("xlsxpatch: native extract: styles and workbook use opposing Strict/Transitional namespaces")
	}
	normalStyle := projectNativeWorkbookNormalStyle(data, registry)
	styles := make([]NativeWorkbookStyleV1, 0, len(registry.cellXfs))
	for index := range registry.cellXfs {
		projection := projectNativeWorkbookStyle(registry, index)
		digest, digestErr := nativeRawStyleProjectionDigest(projection)
		if digestErr != nil {
			return nil, nil, nil, part, fmt.Errorf("xlsxpatch: native extract: style %d projection: %w", index, digestErr)
		}
		styles = append(styles, NativeWorkbookStyleV1{ID: uint32(index), Effective: projection, RawProjectionSHA256: digest})
		for _, code := range projection.Unsupported {
			if err := extractor.addUnsupported("STYLE_"+strings.ToUpper(strings.ReplaceAll(code, "-", "_")), "styles", "style:"+strconv.Itoa(index), part, "", "style component is preserved but cannot be projected safely into the v1 mutation vocabulary"); err != nil {
				return nil, nil, nil, part, err
			}
		}
	}
	return styles, normalStyle, registry, part, nil
}

// nativeRawStyleProjectionDigest fingerprints the canonical compact JSON bytes
// of the exact effective-style projection placed on the wire. Consumers can
// recompute it without access to styles.xml, so a color/token edit cannot retain
// stale source provenance and be mistaken for an extracted projection.
func nativeRawStyleProjectionDigest(projection NativeWorkbookEffectiveStyleV1) (string, error) {
	var buffer bytes.Buffer
	encoder := json.NewEncoder(&buffer)
	encoder.SetEscapeHTML(false)
	err := encoder.Encode(projection)
	if err != nil {
		return "", err
	}
	raw := bytes.TrimSuffix(buffer.Bytes(), []byte{'\n'})
	return nativeWorkbookDigest(raw), nil
}

// projectNativeWorkbookNormalStyle binds grid font metrics to the declared
// built-in Normal style, not to cellXf index zero. Ambiguous or incomplete
// declarations intentionally produce no authority and make geometry refuse.
func projectNativeWorkbookNormalStyle(data []byte, registry *styleRegistry) *NativeWorkbookNormalStyleV1 {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	depth, cellStylesDepth := 0, 0
	var result *NativeWorkbookNormalStyleV1
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			return result
		}
		if err != nil {
			return nil
		}
		switch token := token.(type) {
		case xml.StartElement:
			depth++
			if depth == 2 && token.Name == (xml.Name{Space: registry.index.namespace, Local: "cellStyles"}) {
				if cellStylesDepth != 0 {
					return nil
				}
				cellStylesDepth = depth
				continue
			}
			if cellStylesDepth == 0 || depth != cellStylesDepth+1 || token.Name != (xml.Name{Space: registry.index.namespace, Local: "cellStyle"}) {
				continue
			}
			builtinID, err := optionalUnsignedStyleAttribute(token, "builtinId", -1)
			if err != nil || builtinID != 0 {
				continue
			}
			if result != nil {
				return nil
			}
			xfID, err := requiredUnsignedStyleAttribute(token, "xfId")
			if err != nil || xfID < 0 || xfID >= len(registry.styleXfs) {
				return nil
			}
			xf := effectiveCellStyleXF(registry.styleXfs[xfID])
			if xf.fontID < 0 || xf.fontID >= len(registry.fonts) {
				return nil
			}
			font := registry.fonts[xf.fontID]
			if font.name == nil || *font.name == "" || font.size == nil || *font.size <= 0 {
				return nil
			}
			result = &NativeWorkbookNormalStyleV1{
				StyleXFID: uint32(xfID), FontID: uint32(xf.fontID), FontName: *font.name, FontSizePoints: *font.size,
				FontBold: font.bold, FontItalic: font.italic, FontRecordSHA256: nativeWorkbookDigest(font.raw),
			}
		case xml.EndElement:
			if cellStylesDepth != 0 && depth == cellStylesDepth && token.Name == (xml.Name{Space: registry.index.namespace, Local: "cellStyles"}) {
				cellStylesDepth = 0
			}
			depth--
		}
	}
}

func (extractor *nativeWorkbookExtractor) inventoryNativeStylesMarkup(data []byte, part string) error {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	depth := 0
	containers := []string{}
	modeledContainers := map[string]bool{"numFmts": true, "fonts": true, "fills": true, "borders": true, "cellStyleXfs": true, "cellXfs": true}
	modeledFontChildren := map[string]bool{"name": true, "sz": true, "b": true, "i": true, "color": true}
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			return nil
		}
		if err != nil {
			return err
		}
		switch token := token.(type) {
		case xml.StartElement:
			depth++
			if depth == 1 {
				if len(unexpectedSemanticXMLAttributes(token)) != 0 {
					if err := extractor.addUnsupported("STYLE_SHEET_ATTRIBUTES", "styles", "workbook", part, "", "unmodeled styleSheet attributes remain authority-bound to the source package"); err != nil {
						return err
					}
				}
				continue
			}
			if depth == 2 {
				containers = append(containers, token.Name.Local)
				if token.Name.Space != extractor.namespace || !modeledContainers[token.Name.Local] {
					if err := extractor.addUnsupported("STYLE_TABLE_OPAQUE_CONTENT", "styles", "workbook", part, "", "unmodeled style-table content remains authority-bound to the source package"); err != nil {
						return err
					}
				} else if len(unexpectedSemanticXMLAttributes(token, xml.Name{Local: "count"})) != 0 {
					if err := extractor.addUnsupported("STYLE_TABLE_ATTRIBUTES", "styles", "workbook", part, "", "unmodeled style-table container attributes remain authority-bound to the source package"); err != nil {
						return err
					}
				}
				continue
			}
			container := ""
			if len(containers) != 0 {
				container = containers[len(containers)-1]
			}
			if depth == 3 && modeledContainers[container] {
				allowed := []xml.Name{}
				if container == "fonts" || container == "fills" || container == "borders" {
					// These records have no modeled attributes; all semantics live in children.
				} else if container == "numFmts" {
					allowed = []xml.Name{{Local: "numFmtId"}, {Local: "formatCode"}}
				} else {
					allowed = []xml.Name{{Local: "numFmtId"}, {Local: "fontId"}, {Local: "fillId"}, {Local: "borderId"}, {Local: "xfId"}, {Local: "applyNumberFormat"}, {Local: "applyFont"}, {Local: "applyFill"}, {Local: "applyBorder"}, {Local: "applyAlignment"}, {Local: "applyProtection"}, {Local: "pivotButton"}, {Local: "quotePrefix"}}
				}
				if len(unexpectedSemanticXMLAttributes(token, allowed...)) != 0 {
					if err := extractor.addUnsupported("STYLE_RECORD_ATTRIBUTES", "styles", "workbook", part, "", "unmodeled style-record attributes remain authority-bound to the source package"); err != nil {
						return err
					}
				}
				if (container == "cellXfs" || container == "cellStyleXfs") && token.Name.Local == "xf" {
					if !nativeDefaultOffXMLFlag(token, "pivotButton") || !nativeDefaultOffXMLFlag(token, "quotePrefix") {
						if err := extractor.addUnsupported("STYLE_RECORD_ATTRIBUTES", "styles", "workbook", part, "", "unmodeled style-record attributes remain authority-bound to the source package"); err != nil {
							return err
						}
					}
				}
			}
			if depth == 4 && container == "fonts" && (token.Name.Space != extractor.namespace || !modeledFontChildren[token.Name.Local]) {
				if err := extractor.addUnsupported("STYLE_FONT_OPAQUE_CONTENT", "styles", "workbook", part, "", "unmodeled font children remain authority-bound to the source package"); err != nil {
					return err
				}
			}
			if depth == 4 && (container == "cellXfs" || container == "cellStyleXfs") && (token.Name.Space != extractor.namespace || token.Name.Local != "alignment") {
				if err := extractor.addUnsupported("STYLE_XF_OPAQUE_CONTENT", "styles", "workbook", part, "", "protection, extensions, and unmodeled XF children remain authority-bound to the source package"); err != nil {
					return err
				}
			}
		case xml.EndElement:
			if depth == 2 && len(containers) != 0 {
				containers = containers[:len(containers)-1]
			}
			depth--
		case xml.CharData:
			if depth > 0 && depth <= 3 && len(bytes.TrimSpace(token)) != 0 {
				return fmt.Errorf("styles XML contains unsupported direct text")
			}
		}
	}
}

func defaultNativeStyleProjection() NativeWorkbookEffectiveStyleV1 {
	general, bottom := "General", "general"
	vertical := "bottom"
	falseValue := false
	return NativeWorkbookEffectiveStyleV1{
		NumberFormat: &general, Bold: &falseValue, Italic: &falseValue,
		Fill:                &NativeWorkbookFillV1{Origin: "implicit-default"},
		Border:              &NativeWorkbookBorderV1{Origin: "implicit-default"},
		HorizontalAlignment: &bottom, VerticalAlignment: &vertical, WrapText: &falseValue,
		Projection: "full", Unsupported: []string{},
	}
}

func projectNativeWorkbookStyle(registry *styleRegistry, index int) NativeWorkbookEffectiveStyleV1 {
	source := registry.cellXfs[index]
	base := effectiveCellStyleXF(registry.styleXfs[source.xfID])
	fontID := effectiveStyleComponent(source.fontID, base.fontID, source.applyFont)
	fillID := effectiveStyleComponent(source.fillID, base.fillID, source.applyFill)
	borderID := effectiveStyleComponent(source.borderID, base.borderID, source.applyBorder)
	numFmtID := effectiveStyleComponent(source.numFmtID, base.numFmtID, source.applyNumberFormat)
	alignment := effectiveStyleAlignment(source, base)
	font, fill, border := registry.fonts[fontID], registry.fills[fillID], registry.borders[borderID]
	projection := NativeWorkbookEffectiveStyleV1{
		FontName: cloneString(font.name), FontSizePoints: cloneFloat(font.size),
		Bold: nativeWorkbookBool(font.bold), Italic: nativeWorkbookBool(font.italic),
		HorizontalAlignment: nativeWorkbookString("general"), VerticalAlignment: nativeWorkbookString("bottom"),
		WrapText: nativeWorkbookBool(effectiveWrap(alignment.wrap, nil)), Projection: "full", Unsupported: []string{},
	}
	if format, found := registry.numberFormatCode(numFmtID); found {
		projection.NumberFormat = nativeWorkbookString(format)
	} else {
		projection.Unsupported = append(projection.Unsupported, "number-format")
	}
	if font.colorSafe {
		projection.FontColor = cloneString(font.color)
	} else {
		projection.Unsupported = append(projection.Unsupported, "font-color")
	}
	if fill.supported {
		projection.FillColor = cloneString(fill.color)
		fillIDValue := uint32(fillID)
		fillDigest := nativeWorkbookDigest(fill.raw)
		projection.Fill = &NativeWorkbookFillV1{
			Origin: "styles-record", FillID: &fillIDValue, RecordSHA256: &fillDigest, Color: cloneString(fill.color),
		}
	} else {
		projection.Unsupported = append(projection.Unsupported, "fill")
	}
	if border.supported {
		borderIDValue := uint32(borderID)
		recordDigest := nativeWorkbookDigest(border.raw)
		projection.Border = &NativeWorkbookBorderV1{
			Origin: "styles-record", BorderID: &borderIDValue, RecordSHA256: &recordDigest,
			Left: projectNativeWorkbookBorderSide(border.left), Right: projectNativeWorkbookBorderSide(border.right),
			Top: projectNativeWorkbookBorderSide(border.top), Bottom: projectNativeWorkbookBorderSide(border.bottom),
		}
	} else {
		projection.Unsupported = append(projection.Unsupported, "border")
	}
	horizontal := effectiveHorizontal(alignment.horizontal, nil)
	switch horizontal {
	case "general", "left", "center", "right":
		projection.HorizontalAlignment = nativeWorkbookString(horizontal)
	default:
		projection.HorizontalAlignment = nil
		projection.Unsupported = append(projection.Unsupported, "horizontal-alignment")
	}
	vertical := effectiveVertical(alignment.vertical, nil)
	if vertical == "center" {
		vertical = "middle"
	}
	switch vertical {
	case "top", "middle", "bottom":
		projection.VerticalAlignment = nativeWorkbookString(vertical)
	default:
		projection.VerticalAlignment = nil
		projection.Unsupported = append(projection.Unsupported, "vertical-alignment")
	}
	if alignment.otherAttrsKey != "" {
		projection.Unsupported = append(projection.Unsupported, "alignment-extended")
	}
	sort.Strings(projection.Unsupported)
	if len(projection.Unsupported) != 0 {
		projection.Projection = "partial"
	}
	return projection
}

func projectNativeWorkbookBorderSide(side *styleBorderSide) *NativeWorkbookBorderSideV1 {
	if side == nil {
		return nil
	}
	return &NativeWorkbookBorderSideV1{Style: side.style, Color: side.color}
}
