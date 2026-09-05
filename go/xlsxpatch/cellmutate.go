package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"io"
	"math"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"unicode/utf16"
	"unicode/utf8"
)

// CellMutationKind is the cell subset of @injoffice/sheets protocol v1.
// Formatting, dimensions, merges, and structural edits have separate native
// appliers; this API refuses values outside the four kinds below.
type CellMutationKind string

const (
	CellSetValue     CellMutationKind = "cell.set_value"
	CellClearValue   CellMutationKind = "cell.clear_value"
	CellSetFormula   CellMutationKind = "cell.set_formula"
	CellClearFormula CellMutationKind = "cell.clear_formula"
)

const (
	excelMaxRows                   = 1_048_576
	excelMaxColumns                = 16_384
	maxCellMutations               = 10_000
	maxOperationIDLen              = 128
	maxStableSheetIDLen            = 256
	maxCellTextLen                 = 32_767
	maxFormulaLen                  = 8_192
	spreadsheetMLTransitional      = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
	spreadsheetMLStrict            = "http://purl.oclc.org/ooxml/spreadsheetml/main"
	relTypeCalcChainTransitional   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"
	relTypeCalcChainStrict         = "http://purl.oclc.org/ooxml/officeDocument/relationships/calcChain"
	officeRelNamespaceTransitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	officeRelNamespaceStrict       = "http://purl.oclc.org/ooxml/officeDocument/relationships"
)

// CellRef identifies a zero-based XLSX cell.
type CellRef struct {
	Row    int `json:"row"`
	Column int `json:"column"`
}

// CellMutation is a validated engine input mirroring the cell operations in
// @injoffice/sheets protocol v1. Value accepts strings, booleans, and finite Go
// numeric values. Formula includes the leading '='.
type CellMutation struct {
	OperationID string           `json:"operation_id"`
	SheetID     string           `json:"sheet_id"`
	Kind        CellMutationKind `json:"kind"`
	Cell        CellRef          `json:"cell"`
	Value       any              `json:"value,omitempty"`
	Formula     string           `json:"formula,omitempty"`
}

type normalizedCellMutation struct {
	CellMutation
	payload cellPayload
}

type cellPayload struct {
	kind        CellMutationKind
	literalKind literalKind
	text        string
	number      string
	boolean     bool
}

type literalKind string

const (
	literalString  literalKind = "string"
	literalNumber  literalKind = "number"
	literalBoolean literalKind = "boolean"
)

type cellKey struct {
	row, column int
}

// ApplyCellMutations applies literal, formula, and clear operations directly
// to worksheet XML parts in the original XLSX package. Operations are
// validated atomically, applied in listed order, and collapsed to their final
// value per cell. Every unrelated OPC part is raw-copied and verified by
// Apply. Existing cell attributes and extLst content are preserved.
func ApplyCellMutations(orig []byte, mutations []CellMutation) ([]byte, error) {
	normalized, err := validateCellMutations(mutations)
	if err != nil {
		return nil, err
	}

	zr, err := zip.NewReader(bytes.NewReader(orig), int64(len(orig)))
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: cell mutations: open original: %w", err)
	}
	packageIndex, err := newOPCPackageIndex(zr)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: cell mutations: %w", err)
	}
	files := packageIndex.byExact
	metadata := make(map[string]string, 2)
	var metadataReadErr error
	read := func(name string) (string, bool) {
		if value, ok := metadata[name]; ok {
			return value, true
		}
		file, ok := files[name]
		if !ok || metadataReadErr != nil {
			return "", false
		}
		data, err := readZipFile(file)
		if err != nil {
			metadataReadErr = fmt.Errorf("read %q: %w", name, err)
			return "", false
		}
		value := string(data)
		metadata[name] = value
		return value, true
	}
	workbook, err := locateWorkbookPart(packageIndex, read)
	if metadataReadErr != nil {
		return nil, fmt.Errorf("xlsxpatch: cell mutations: %w", metadataReadErr)
	}
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: cell mutations: %w", err)
	}

	byPart := make(map[string][]normalizedCellMutation)
	partOrder := make([]string, 0)
	resolvedSheets := make(map[string]string)
	for _, mutation := range normalized {
		part, ok := resolvedSheets[mutation.SheetID]
		if !ok {
			part, err = worksheetPartForID(packageIndex, read, workbook, mutation.SheetID)
			if metadataReadErr != nil {
				return nil, fmt.Errorf("xlsxpatch: cell mutations: operation %q: %w", mutation.OperationID, metadataReadErr)
			}
			if err != nil {
				return nil, fmt.Errorf("xlsxpatch: cell mutations: operation %q: %w", mutation.OperationID, err)
			}
			resolvedSheets[mutation.SheetID] = part
			if _, seen := byPart[part]; !seen {
				partOrder = append(partOrder, part)
			}
		}
		byPart[part] = append(byPart[part], mutation)
	}

	patch := Patch{Replace: make(map[string][]byte, len(byPart))}
	for _, part := range partOrder {
		file, ok := files[part]
		if !ok {
			return nil, fmt.Errorf("xlsxpatch: cell mutations: worksheet part %q is missing", part)
		}
		sheetXML, err := readZipFile(file)
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: cell mutations: read worksheet part %q: %w", part, err)
		}
		updated, err := patchWorksheetCells(sheetXML, byPart[part])
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: cell mutations: worksheet %q: %w", part, err)
		}
		if !bytes.Equal(updated, sheetXML) {
			patch.Replace[part] = updated
		}
	}
	if len(patch.Replace) == 0 {
		return bytes.Clone(orig), nil
	}
	if err := applyFullRecalculationPatch(packageIndex, read, workbook, &patch); err != nil {
		return nil, fmt.Errorf("xlsxpatch: cell mutations: calculation state: %w", err)
	}
	return Apply(orig, patch)
}

func readZipFile(file *zip.File) ([]byte, error) {
	rc, err := file.Open()
	if err != nil {
		return nil, err
	}
	data, readErr := io.ReadAll(rc)
	closeErr := rc.Close()
	if readErr != nil {
		return nil, readErr
	}
	if closeErr != nil {
		return nil, closeErr
	}
	return data, nil
}

var restrictedIDPattern = regexp.MustCompile(`^[A-Za-z0-9][A-Za-z0-9._:-]*$`)

func validateCellMutations(mutations []CellMutation) ([]normalizedCellMutation, error) {
	if len(mutations) == 0 {
		return nil, fmt.Errorf("xlsxpatch: cell mutations: empty batch")
	}
	if len(mutations) > maxCellMutations {
		return nil, fmt.Errorf("xlsxpatch: cell mutations: batch exceeds %d operations", maxCellMutations)
	}
	seenIDs := make(map[string]bool, len(mutations))
	normalized := make([]normalizedCellMutation, 0, len(mutations))
	for index, mutation := range mutations {
		prefix := fmt.Sprintf("xlsxpatch: cell mutations: operation %d", index)
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
		if mutation.Cell.Row < 0 || mutation.Cell.Row >= excelMaxRows || mutation.Cell.Column < 0 || mutation.Cell.Column >= excelMaxColumns {
			return nil, fmt.Errorf("%s: cell (%d,%d) is outside Excel limits", prefix, mutation.Cell.Row, mutation.Cell.Column)
		}

		payload := cellPayload{kind: mutation.Kind}
		switch mutation.Kind {
		case CellSetValue:
			if mutation.Formula != "" {
				return nil, fmt.Errorf("%s: cell.set_value must not include formula", prefix)
			}
			var err error
			payload, err = normalizeLiteral(mutation.Value)
			if err != nil {
				return nil, fmt.Errorf("%s: %w", prefix, err)
			}
		case CellSetFormula:
			if mutation.Value != nil {
				return nil, fmt.Errorf("%s: cell.set_formula must not include value", prefix)
			}
			if !strings.HasPrefix(mutation.Formula, "=") || utf16Length(mutation.Formula) > maxFormulaLen || len(mutation.Formula) == 1 {
				return nil, fmt.Errorf("%s: formula must start with '=' and contain 1 to %d UTF-16 code units after it", prefix, maxFormulaLen-1)
			}
			if _, err := escapeXMLText(strings.TrimPrefix(mutation.Formula, "=")); err != nil {
				return nil, fmt.Errorf("%s: invalid formula text: %w", prefix, err)
			}
			payload.text = strings.TrimPrefix(mutation.Formula, "=")
		case CellClearValue, CellClearFormula:
			if mutation.Value != nil || mutation.Formula != "" {
				return nil, fmt.Errorf("%s: clear operation must not include value or formula", prefix)
			}
		default:
			return nil, fmt.Errorf("%s: unsupported kind %q", prefix, mutation.Kind)
		}
		normalized = append(normalized, normalizedCellMutation{CellMutation: mutation, payload: payload})
	}
	return normalized, nil
}

func normalizeLiteral(value any) (cellPayload, error) {
	payload := cellPayload{kind: CellSetValue}
	switch value := value.(type) {
	case string:
		if utf16Length(value) > maxCellTextLen {
			return cellPayload{}, fmt.Errorf("string value exceeds %d UTF-16 code units", maxCellTextLen)
		}
		if _, err := encodeSpreadsheetString(value); err != nil {
			return cellPayload{}, fmt.Errorf("invalid string value: %w", err)
		}
		payload.literalKind, payload.text = literalString, value
	case bool:
		payload.literalKind, payload.boolean = literalBoolean, value
	case float64:
		if math.IsNaN(value) || math.IsInf(value, 0) {
			return cellPayload{}, fmt.Errorf("number value must be finite")
		}
		payload.literalKind, payload.number = literalNumber, strconv.FormatFloat(value, 'g', -1, 64)
	case float32:
		if math.IsNaN(float64(value)) || math.IsInf(float64(value), 0) {
			return cellPayload{}, fmt.Errorf("number value must be finite")
		}
		payload.literalKind, payload.number = literalNumber, strconv.FormatFloat(float64(value), 'g', -1, 32)
	case json.Number:
		parsed, err := strconv.ParseFloat(value.String(), 64)
		if err != nil || math.IsNaN(parsed) || math.IsInf(parsed, 0) {
			return cellPayload{}, fmt.Errorf("number value must be finite")
		}
		payload.literalKind, payload.number = literalNumber, value.String()
	case int:
		payload.literalKind, payload.number = literalNumber, strconv.FormatInt(int64(value), 10)
	case int8:
		payload.literalKind, payload.number = literalNumber, strconv.FormatInt(int64(value), 10)
	case int16:
		payload.literalKind, payload.number = literalNumber, strconv.FormatInt(int64(value), 10)
	case int32:
		payload.literalKind, payload.number = literalNumber, strconv.FormatInt(int64(value), 10)
	case int64:
		payload.literalKind, payload.number = literalNumber, strconv.FormatInt(value, 10)
	case uint:
		payload.literalKind, payload.number = literalNumber, strconv.FormatUint(uint64(value), 10)
	case uint8:
		payload.literalKind, payload.number = literalNumber, strconv.FormatUint(uint64(value), 10)
	case uint16:
		payload.literalKind, payload.number = literalNumber, strconv.FormatUint(uint64(value), 10)
	case uint32:
		payload.literalKind, payload.number = literalNumber, strconv.FormatUint(uint64(value), 10)
	case uint64:
		payload.literalKind, payload.number = literalNumber, strconv.FormatUint(value, 10)
	default:
		return cellPayload{}, fmt.Errorf("value must be a string, boolean, or finite number")
	}
	return payload, nil
}

func utf16Length(value string) int { return len(utf16.Encode([]rune(value))) }

func escapeXMLText(value string) (string, error) {
	if !utf8.ValidString(value) {
		return "", fmt.Errorf("text is not valid UTF-8")
	}
	for _, character := range value {
		if character == '\t' || character == '\n' || character == '\r' ||
			(character >= 0x20 && character <= 0xD7FF) ||
			(character >= 0xE000 && character <= 0xFFFD) ||
			(character >= 0x10000 && character <= 0x10FFFF) {
			continue
		}
		return "", fmt.Errorf("text contains XML-forbidden character U+%04X", character)
	}
	var out bytes.Buffer
	if err := xml.EscapeText(&out, []byte(value)); err != nil {
		return "", err
	}
	return out.String(), nil
}

// encodeSpreadsheetString implements the SpreadsheetML ST_Xstring escape
// layer before ordinary XML escaping. XML-forbidden UTF-16 code units use
// _xHHHH_, and a literal substring that already looks like an Office escape
// has its leading underscore escaped so Excel does not reinterpret user text.
func encodeSpreadsheetString(value string) (string, error) {
	if !utf8.ValidString(value) {
		return "", fmt.Errorf("text is not valid UTF-8")
	}
	var encoded strings.Builder
	for offset := 0; offset < len(value); {
		if value[offset] == '_' && looksLikeSpreadsheetEscape(value[offset:]) {
			encoded.WriteString("_x005F_")
			offset++
			continue
		}
		character, size := utf8.DecodeRuneInString(value[offset:])
		offset += size
		if xmlForbiddenRune(character) {
			if character > 0xffff {
				return "", fmt.Errorf("cannot encode XML-forbidden character U+%04X", character)
			}
			fmt.Fprintf(&encoded, "_x%04X_", character)
			continue
		}
		encoded.WriteRune(character)
	}
	var escaped bytes.Buffer
	if err := xml.EscapeText(&escaped, []byte(encoded.String())); err != nil {
		return "", err
	}
	return escaped.String(), nil
}

func looksLikeSpreadsheetEscape(value string) bool {
	if len(value) < 7 || value[0] != '_' || (value[1] != 'x' && value[1] != 'X') || value[6] != '_' {
		return false
	}
	for index := 2; index < 6; index++ {
		character := value[index]
		if !((character >= '0' && character <= '9') || (character >= 'a' && character <= 'f') || (character >= 'A' && character <= 'F')) {
			return false
		}
	}
	return true
}

func xmlForbiddenRune(character rune) bool {
	return character != '\t' && character != '\n' && character != '\r' &&
		(character < 0x20 || (character >= 0xd800 && character <= 0xdfff) || character == 0xfffe || character == 0xffff)
}

func worksheetPartForID(index *opcPackageIndex, read func(string) (string, bool), workbook workbookPartLocation, sheetID string) (string, error) {
	canonicalSheetID, err := canonicalNativeSheetID(sheetID)
	if err != nil {
		return "", fmt.Errorf("stable sheet id %q is invalid: %w", sheetID, err)
	}
	workbookXML, ok := read(workbook.part)
	if !ok {
		return "", fmt.Errorf("missing workbook part %q", workbook.part)
	}
	relID, rootNamespace := "", ""
	matchCount, depth, sheetsDepth := 0, 0, 0
	rootClosed := false
	decoder := xml.NewDecoder(strings.NewReader(workbookXML))
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return "", fmt.Errorf("parse workbook.xml: %w", err)
		}
		switch token := token.(type) {
		case xml.StartElement:
			depth++
			if depth == 1 {
				if token.Name.Local != "workbook" || !isSpreadsheetMLNamespace(token.Name.Space) {
					return "", fmt.Errorf("root element is not a supported SpreadsheetML workbook")
				}
				rootNamespace = token.Name.Space
				continue
			}
			if depth == 2 && token.Name.Space == rootNamespace && token.Name.Local == "sheets" {
				sheetsDepth = depth
				continue
			}
			if sheetsDepth == 0 || depth != sheetsDepth+1 || token.Name.Space != rootNamespace || token.Name.Local != "sheet" {
				continue
			}
			candidateID, candidateRel := "", ""
			for _, attr := range token.Attr {
				switch {
				case attr.Name.Space == "" && attr.Name.Local == "sheetId":
					candidateID = attr.Value
				case attr.Name.Local == "id" && isOfficeRelationshipNamespace(attr.Name.Space):
					candidateRel = attr.Value
				}
			}
			candidateCanonical, candidateErr := canonicalNativeSheetID(candidateID)
			if candidateErr != nil {
				return "", fmt.Errorf("workbook sheetId=%q is invalid: %w", candidateID, candidateErr)
			}
			if candidateCanonical == canonicalSheetID {
				matchCount++
				relID = candidateRel
			}
		case xml.EndElement:
			if sheetsDepth != 0 && depth == sheetsDepth && token.Name.Space == rootNamespace && token.Name.Local == "sheets" {
				sheetsDepth = 0
			}
			if depth == 1 && token.Name.Space == rootNamespace && token.Name.Local == "workbook" {
				rootClosed = true
			}
			depth--
		}
	}
	if rootNamespace == "" || !rootClosed || depth != 0 {
		return "", fmt.Errorf("missing complete workbook root")
	}
	if matchCount == 0 {
		return "", fmt.Errorf("stable sheet id %q not found", sheetID)
	}
	if matchCount != 1 {
		return "", fmt.Errorf("stable sheet id %q is duplicated", sheetID)
	}
	if relID == "" {
		return "", fmt.Errorf("stable sheet id %q has no relationship id in a supported namespace", sheetID)
	}
	relsXML, ok := read(workbook.relsPart)
	if !ok {
		return "", fmt.Errorf("missing workbook relationships")
	}
	resolved, err := relTargetOfType(relsXML, relID, workbook.baseDir, relTypeWorksheetTransitional, relTypeWorksheetStrict)
	if err != nil {
		return "", err
	}
	part, _, ok := index.lookupResolved(resolved)
	if !ok {
		return "", fmt.Errorf("worksheet relationship %q targets missing part %q", relID, resolved)
	}
	return part, nil
}

func isOfficeRelationshipNamespace(namespace string) bool {
	return namespace == officeRelNamespaceTransitional || namespace == officeRelNamespaceStrict
}

type xmlSpan struct {
	qname       string
	start       int
	startTagEnd int
	endStart    int
	end         int
}

type rawAttributeSpan struct {
	qname      string
	leading    int
	start      int
	end        int
	valueStart int
	valueEnd   int
}

func rawStartTagAttributes(tag []byte) ([]rawAttributeSpan, error) {
	left := bytes.IndexByte(tag, '<')
	if left < 0 {
		return nil, fmt.Errorf("missing start-tag opener")
	}
	index := left + 1
	if index >= len(tag) || tag[index] == '/' || tag[index] == '!' || tag[index] == '?' {
		return nil, fmt.Errorf("not an element start tag")
	}
	for index < len(tag) && !isXMLSpace(tag[index]) && tag[index] != '/' && tag[index] != '>' {
		index++
	}
	if index == left+1 {
		return nil, fmt.Errorf("missing element name")
	}
	attributes := make([]rawAttributeSpan, 0)
	for {
		leading := index
		for index < len(tag) && isXMLSpace(tag[index]) {
			index++
		}
		if index >= len(tag) || tag[index] == '>' || (tag[index] == '/' && index+1 < len(tag) && tag[index+1] == '>') {
			return attributes, nil
		}
		start := index
		for index < len(tag) && !isXMLSpace(tag[index]) && tag[index] != '=' && tag[index] != '/' && tag[index] != '>' {
			index++
		}
		if start == index {
			return nil, fmt.Errorf("malformed attribute name")
		}
		qname := string(tag[start:index])
		for index < len(tag) && isXMLSpace(tag[index]) {
			index++
		}
		if index >= len(tag) || tag[index] != '=' {
			return nil, fmt.Errorf("attribute %q is missing '='", qname)
		}
		index++
		for index < len(tag) && isXMLSpace(tag[index]) {
			index++
		}
		if index >= len(tag) || (tag[index] != '\'' && tag[index] != '"') {
			return nil, fmt.Errorf("attribute %q is missing a quoted value", qname)
		}
		quote := tag[index]
		index++
		valueStart := index
		for index < len(tag) && tag[index] != quote {
			index++
		}
		if index >= len(tag) {
			return nil, fmt.Errorf("attribute %q has an unterminated value", qname)
		}
		valueEnd := index
		index++
		attributes = append(attributes, rawAttributeSpan{qname: qname, leading: leading, start: start, end: index, valueStart: valueStart, valueEnd: valueEnd})
	}
}

func isXMLSpace(character byte) bool {
	return character == ' ' || character == '\t' || character == '\r' || character == '\n'
}

func rawUnqualifiedAttribute(tag []byte, name string) (rawAttributeSpan, bool, error) {
	attributes, err := rawStartTagAttributes(tag)
	if err != nil {
		return rawAttributeSpan{}, false, err
	}
	var found rawAttributeSpan
	seen := false
	for _, attribute := range attributes {
		if attribute.qname != name {
			continue
		}
		if seen {
			return rawAttributeSpan{}, false, fmt.Errorf("duplicate unqualified attribute %q", name)
		}
		found, seen = attribute, true
	}
	return found, seen, nil
}

// rewriteUnqualifiedAttribute changes exactly one parsed, unqualified XML
// attribute while preserving every other byte of the start tag. A nil value
// removes the attribute; a non-nil value replaces or appends it.
func rewriteUnqualifiedAttribute(tag []byte, name string, value *string) ([]byte, error) {
	attribute, found, err := rawUnqualifiedAttribute(tag, name)
	if err != nil {
		return nil, err
	}
	if found {
		if value == nil {
			return append(bytes.Clone(tag[:attribute.leading]), tag[attribute.end:]...), nil
		}
		escaped, err := escapeXMLAttribute(*value)
		if err != nil {
			return nil, err
		}
		return append(bytes.Clone(tag[:attribute.valueStart]), append([]byte(escaped), tag[attribute.valueEnd:]...)...), nil
	}
	if value == nil {
		return bytes.Clone(tag), nil
	}
	escaped, err := escapeXMLAttribute(*value)
	if err != nil {
		return nil, err
	}
	closeAt := bytes.LastIndexByte(tag, '>')
	if closeAt < 0 {
		return nil, fmt.Errorf("missing start-tag closer")
	}
	if closeAt > 0 && tag[closeAt-1] == '/' {
		closeAt--
	}
	insert := []byte(" " + name + "=\"" + escaped + "\"")
	return append(bytes.Clone(tag[:closeAt]), append(insert, tag[closeAt:]...)...), nil
}

func escapeXMLAttribute(value string) (string, error) {
	var out bytes.Buffer
	if err := xml.EscapeText(&out, []byte(value)); err != nil {
		return "", err
	}
	return strings.ReplaceAll(out.String(), "\"", "&#34;"), nil
}

type directChildElement struct {
	span  xmlSpan
	start xml.StartElement
}

func directChildElements(data []byte, rootLocal, childLocal string) ([]directChildElement, xmlSpan, error) {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	depth := 0
	root := xmlSpan{}
	rootNamespace := ""
	children := make([]directChildElement, 0)
	var active *directChildElement
	for {
		before := int(decoder.InputOffset())
		token, err := decoder.Token()
		after := int(decoder.InputOffset())
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, xmlSpan{}, err
		}
		switch token := token.(type) {
		case xml.StartElement:
			depth++
			if depth == 1 {
				if token.Name.Local != rootLocal {
					return nil, xmlSpan{}, fmt.Errorf("root element is %q, expected %q", token.Name.Local, rootLocal)
				}
				rootNamespace = token.Name.Space
				root = xmlSpan{qname: rawQName(data, before, after), start: before, startTagEnd: after}
			} else if depth == 2 && token.Name.Space == rootNamespace && token.Name.Local == childLocal {
				if active != nil {
					return nil, xmlSpan{}, fmt.Errorf("nested %s element", childLocal)
				}
				active = &directChildElement{span: xmlSpan{qname: rawQName(data, before, after), start: before, startTagEnd: after}, start: token}
			}
		case xml.EndElement:
			if active != nil && depth == 2 && token.Name.Space == rootNamespace && token.Name.Local == childLocal {
				active.span.endStart, active.span.end = before, after
				children = append(children, *active)
				active = nil
			}
			if depth == 1 && token.Name.Space == rootNamespace && token.Name.Local == rootLocal {
				root.endStart, root.end = before, after
			}
			depth--
		}
	}
	if root.end == 0 {
		return nil, xmlSpan{}, fmt.Errorf("missing complete %s root", rootLocal)
	}
	return children, root, nil
}

func removeElementSpans(data []byte, spans []xmlSpan) []byte {
	sort.Slice(spans, func(i, j int) bool { return spans[i].start > spans[j].start })
	updated := bytes.Clone(data)
	for _, span := range spans {
		updated = append(updated[:span.start], updated[span.end:]...)
	}
	return updated
}

func applyFullRecalculationPatch(index *opcPackageIndex, read func(string) (string, bool), workbook workbookPartLocation, patch *Patch) error {
	workbookXML, ok := read(workbook.part)
	if !ok {
		return fmt.Errorf("missing workbook part %q", workbook.part)
	}
	updatedWorkbook, err := workbookWithFullRecalculation([]byte(workbookXML))
	if err != nil {
		return err
	}
	patch.Replace[workbook.part] = updatedWorkbook

	relsXML, ok := read(workbook.relsPart)
	if !ok {
		return fmt.Errorf("missing workbook relationships")
	}
	updatedRels, calcChainTarget, found, err := workbookRelationshipsWithoutCalcChain([]byte(relsXML), workbook.baseDir)
	if err != nil {
		return err
	}
	if !found {
		orphanCandidate := "calcChain.xml"
		if workbook.baseDir != "" {
			orphanCandidate = workbook.baseDir + "/" + orphanCandidate
		}
		if orphan, _, exists := index.lookupResolved(orphanCandidate); exists {
			return fmt.Errorf("orphan %s has no workbook relationship", orphan)
		}
		return nil
	}
	calcChainPart, _, ok := index.lookupResolved(calcChainTarget)
	if !ok {
		return fmt.Errorf("calc-chain relationship targets missing part %q", calcChainTarget)
	}
	contentTypesPart, _, ok := index.lookupSpelling("[Content_Types].xml")
	if !ok {
		return fmt.Errorf("missing [Content_Types].xml")
	}
	contentTypesXML, ok := read(contentTypesPart)
	if !ok {
		return fmt.Errorf("missing [Content_Types].xml")
	}
	updatedContentTypes, err := contentTypesWithoutPart([]byte(contentTypesXML), calcChainPart)
	if err != nil {
		return err
	}
	patch.Replace[workbook.relsPart] = updatedRels
	patch.Replace[contentTypesPart] = updatedContentTypes
	if patch.Delete == nil {
		patch.Delete = make(map[string]bool)
	}
	patch.Delete[calcChainPart] = true
	return nil
}

func workbookRelationshipsWithoutCalcChain(data []byte, baseDir string) ([]byte, string, bool, error) {
	relationships, _, err := directChildElements(data, "Relationships", "Relationship")
	if err != nil {
		return nil, "", false, fmt.Errorf("parse workbook relationships: %w", err)
	}
	var match *directChildElement
	for index := range relationships {
		relType := attribute(relationships[index].start, "Type")
		if relType != relTypeCalcChainTransitional && relType != relTypeCalcChainStrict {
			continue
		}
		if match != nil {
			return nil, "", false, fmt.Errorf("multiple calc-chain relationships")
		}
		match = &relationships[index]
	}
	if match == nil {
		return bytes.Clone(data), "", false, nil
	}
	mode := strings.TrimSpace(attribute(match.start, "TargetMode"))
	if mode != "" && !strings.EqualFold(mode, "Internal") {
		if strings.EqualFold(mode, "External") {
			return nil, "", false, fmt.Errorf("calc-chain relationship is external")
		}
		return nil, "", false, fmt.Errorf("calc-chain relationship has unsupported target mode %q", mode)
	}
	target := attribute(match.start, "Target")
	if target == "" {
		return nil, "", false, fmt.Errorf("calc-chain relationship has no target")
	}
	part, err := resolveRelPath(baseDir, target)
	if err != nil {
		return nil, "", false, fmt.Errorf("calc-chain relationship has invalid target %q: %w", target, err)
	}
	return removeElementSpans(data, []xmlSpan{match.span}), part, true, nil
}

func contentTypesWithoutPart(data []byte, partName string) ([]byte, error) {
	overrides, _, err := directChildElements(data, "Types", "Override")
	if err != nil {
		return nil, fmt.Errorf("parse content types: %w", err)
	}
	targetKey, err := canonicalOPCPartKey(partName)
	if err != nil {
		return nil, fmt.Errorf("invalid calc-chain part %q: %w", partName, err)
	}
	var matches []xmlSpan
	for _, override := range overrides {
		overrideName := attribute(override.start, "PartName")
		overrideKey, keyErr := canonicalOPCPartKey(overrideName)
		if keyErr != nil {
			return nil, fmt.Errorf("invalid content-type PartName %q: %w", overrideName, keyErr)
		}
		if overrideKey == targetKey {
			matches = append(matches, override.span)
		}
	}
	if len(matches) != 1 {
		return nil, fmt.Errorf("calc-chain part %q has %d content-type overrides, expected 1", partName, len(matches))
	}
	return removeElementSpans(data, matches), nil
}

func workbookWithFullRecalculation(data []byte) ([]byte, error) {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	depth := 0
	root := xmlSpan{}
	rootNamespace := ""
	insertAt := -1
	var calcPr *xmlSpan
	followers := map[string]bool{
		"oleSize": true, "customWorkbookViews": true, "pivotCaches": true,
		"smartTagPr": true, "smartTagTypes": true, "webPublishing": true,
		"fileRecoveryPr": true, "webPublishObjects": true, "extLst": true,
	}
	for {
		before := int(decoder.InputOffset())
		token, err := decoder.Token()
		after := int(decoder.InputOffset())
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("parse workbook.xml: %w", err)
		}
		switch token := token.(type) {
		case xml.StartElement:
			depth++
			if depth == 1 {
				if token.Name.Local != "workbook" {
					return nil, fmt.Errorf("root element is not workbook")
				}
				rootNamespace = token.Name.Space
				root = xmlSpan{qname: rawQName(data, before, after), start: before, startTagEnd: after}
				continue
			}
			if depth != 2 || token.Name.Space != rootNamespace {
				continue
			}
			if token.Name.Local == "calcPr" {
				if calcPr != nil {
					return nil, fmt.Errorf("multiple calcPr elements")
				}
				span := xmlSpan{qname: rawQName(data, before, after), start: before, startTagEnd: after}
				calcPr = &span
			} else if insertAt < 0 && followers[token.Name.Local] {
				insertAt = before
			}
		case xml.EndElement:
			if calcPr != nil && calcPr.end == 0 && depth == 2 && token.Name.Space == rootNamespace && token.Name.Local == "calcPr" {
				calcPr.endStart, calcPr.end = before, after
			}
			if depth == 1 && token.Name.Space == rootNamespace && token.Name.Local == "workbook" {
				root.endStart, root.end = before, after
			}
			depth--
		}
	}
	if root.end == 0 {
		return nil, fmt.Errorf("missing complete workbook root")
	}
	if calcPr == nil {
		if insertAt < 0 {
			insertAt = root.endStart
		}
		qname := prefixedLocal(root.qname, "calcPr")
		element := []byte("<" + qname + ` calcId="0" fullCalcOnLoad="1" forceFullCalc="1" calcCompleted="0"/>`)
		return append(bytes.Clone(data[:insertAt]), append(element, data[insertAt:]...)...), nil
	}
	start := bytes.Clone(data[calcPr.start:calcPr.startTagEnd])
	for _, setting := range []struct{ name, value string }{
		{name: "calcId", value: "0"},
		{name: "fullCalcOnLoad", value: "1"},
		{name: "forceFullCalc", value: "1"},
		{name: "calcCompleted", value: "0"},
	} {
		updated, err := rewriteUnqualifiedAttribute(start, setting.name, &setting.value)
		if err != nil {
			return nil, fmt.Errorf("rewrite calcPr %s: %w", setting.name, err)
		}
		start = updated
	}
	return append(bytes.Clone(data[:calcPr.start]), append(start, data[calcPr.startTagEnd:]...)...), nil
}

func (span xmlSpan) selfClosing(data []byte) bool {
	return bytes.HasSuffix(bytes.TrimSpace(data[span.start:span.startTagEnd]), []byte("/>"))
}

type cellSpan struct {
	xmlSpan
	column   int
	space    string
	children []namedSpan
}

type namedSpan struct {
	local string
	space string
	xmlSpan
}

type rowSpan struct {
	xmlSpan
	row   int
	cells []cellSpan
}

type worksheetIndex struct {
	sheetData    xmlSpan
	dimension    *xmlSpan
	dimensionRef string
	namespace    string
	existing     map[int]rowSpan
	insertions   map[int]int
}

type activeRow struct {
	span      rowSpan
	depth     int
	touched   bool
	lastCol   int
	hasColumn bool
}

type activeCell struct {
	span  cellSpan
	depth int
}

type activeChild struct {
	span  namedSpan
	depth int
}

func patchWorksheetCells(sheetXML []byte, mutations []normalizedCellMutation) ([]byte, error) {
	final := make(map[cellKey]normalizedCellMutation, len(mutations))
	finalByRow := make(map[int]map[int]normalizedCellMutation)
	for _, mutation := range mutations {
		key := cellKey{row: mutation.Cell.Row, column: mutation.Cell.Column}
		final[key] = mutation
		if finalByRow[key.row] == nil {
			finalByRow[key.row] = make(map[int]normalizedCellMutation)
		}
		finalByRow[key.row][key.column] = mutation
	}
	if err := validateMergedCellTargets(sheetXML, final); err != nil {
		return nil, err
	}
	if err := validateFormulaGroupTargets(sheetXML, final); err != nil {
		return nil, err
	}
	targetRows := make([]int, 0, len(finalByRow))
	for row := range finalByRow {
		targetRows = append(targetRows, row)
	}
	sort.Ints(targetRows)

	index, err := indexWorksheet(sheetXML, targetRows)
	if err != nil {
		return nil, err
	}

	type textEdit struct {
		start, end int
		data       []byte
	}
	edits := make([]textEdit, 0, len(targetRows))
	inserts := make(map[int][]int)
	for _, rowNumber := range targetRows {
		rowMutations := finalByRow[rowNumber]
		if row, exists := index.existing[rowNumber]; exists {
			updated, err := rebuildExistingRow(sheetXML, row, rowMutations)
			if err != nil {
				return nil, fmt.Errorf("row %d: %w", rowNumber+1, err)
			}
			if !bytes.Equal(updated, sheetXML[row.start:row.end]) {
				edits = append(edits, textEdit{start: row.start, end: row.end, data: updated})
			}
			continue
		}
		if hasMaterialMutation(rowMutations) {
			inserts[index.insertions[rowNumber]] = append(inserts[index.insertions[rowNumber]], rowNumber)
		}
	}
	for offset, rows := range inserts {
		sort.Ints(rows)
		var data bytes.Buffer
		for _, rowNumber := range rows {
			built, err := buildNewRow(index.sheetData.qname, rowNumber, finalByRow[rowNumber])
			if err != nil {
				return nil, fmt.Errorf("row %d: %w", rowNumber+1, err)
			}
			data.Write(built)
		}
		edits = append(edits, textEdit{start: offset, end: offset, data: data.Bytes()})
	}

	if len(edits) == 0 {
		return bytes.Clone(sheetXML), nil
	}
	if index.sheetData.selfClosing(sheetXML) {
		// A self-closing sheetData cannot contain an existing row. Replace it
		// once so same-offset insertions do not produce invalid XML.
		var rows bytes.Buffer
		rowNumbers := make([]int, 0, len(finalByRow))
		for row, rowMutations := range finalByRow {
			if hasMaterialMutation(rowMutations) {
				rowNumbers = append(rowNumbers, row)
			}
		}
		sort.Ints(rowNumbers)
		for _, row := range rowNumbers {
			built, err := buildNewRow(index.sheetData.qname, row, finalByRow[row])
			if err != nil {
				return nil, err
			}
			rows.Write(built)
		}
		if rows.Len() == 0 {
			return bytes.Clone(sheetXML), nil
		}
		start := openTag(sheetXML[index.sheetData.start:index.sheetData.startTagEnd])
		replacement := append(start, rows.Bytes()...)
		replacement = append(replacement, []byte("</"+index.sheetData.qname+">")...)
		edits = []textEdit{{start: index.sheetData.start, end: index.sheetData.end, data: replacement}}
	}
	if index.dimension != nil {
		expanded, err := expandDimensionTag(sheetXML[index.dimension.start:index.dimension.startTagEnd], index.dimensionRef, final)
		if err != nil {
			return nil, err
		}
		if !bytes.Equal(expanded, sheetXML[index.dimension.start:index.dimension.startTagEnd]) {
			edits = append(edits, textEdit{start: index.dimension.start, end: index.dimension.startTagEnd, data: expanded})
		}
	}

	sort.Slice(edits, func(i, j int) bool {
		if edits[i].start == edits[j].start {
			return edits[i].end > edits[j].end
		}
		return edits[i].start > edits[j].start
	})
	updated := bytes.Clone(sheetXML)
	for _, edit := range edits {
		updated = append(updated[:edit.start], append(edit.data, updated[edit.end:]...)...)
	}
	if err := validateWorksheetXML(updated); err != nil {
		return nil, fmt.Errorf("produced invalid worksheet XML: %w", err)
	}
	return updated, nil
}

func indexWorksheet(data []byte, targetRows []int) (worksheetIndex, error) {
	index := worksheetIndex{existing: make(map[int]rowSpan), insertions: make(map[int]int)}
	targetSet := make(map[int]bool, len(targetRows))
	for _, row := range targetRows {
		targetSet[row] = true
	}
	decoder := xml.NewDecoder(bytes.NewReader(data))
	depth := 0
	worksheetDepth := 0
	sheetDepth := 0
	foundSheetData := false
	foundDimension := false
	lastRow := -1
	targetCursor := 0
	var row *activeRow
	var cell *activeCell
	var child *activeChild
	for {
		before := int(decoder.InputOffset())
		token, err := decoder.Token()
		after := int(decoder.InputOffset())
		if err == io.EOF {
			break
		}
		if err != nil {
			return worksheetIndex{}, fmt.Errorf("parse worksheet XML: %w", err)
		}
		switch token := token.(type) {
		case xml.StartElement:
			depth++
			if worksheetDepth == 0 {
				if depth != 1 || token.Name.Local != "worksheet" || !isSpreadsheetMLNamespace(token.Name.Space) {
					return worksheetIndex{}, fmt.Errorf("root element is not a supported SpreadsheetML worksheet")
				}
				worksheetDepth = depth
				index.namespace = token.Name.Space
				continue
			}
			if worksheetDepth > 0 && depth == worksheetDepth+1 && token.Name.Space == index.namespace && token.Name.Local == "dimension" {
				if foundDimension {
					return worksheetIndex{}, fmt.Errorf("multiple dimension elements")
				}
				foundDimension = true
				span := xmlSpan{qname: rawQName(data, before, after), start: before, startTagEnd: after}
				index.dimension = &span
				index.dimensionRef = attribute(token, "ref")
				continue
			}
			if token.Name.Space == index.namespace && token.Name.Local == "sheetData" {
				if foundSheetData {
					return worksheetIndex{}, fmt.Errorf("multiple sheetData elements")
				}
				foundSheetData = true
				sheetDepth = depth
				index.sheetData = xmlSpan{qname: rawQName(data, before, after), start: before, startTagEnd: after}
				continue
			}
			if foundSheetData && depth == sheetDepth+1 && token.Name.Space == index.namespace && token.Name.Local == "row" {
				rowNumber, err := positiveIndexAttribute(token, "r", excelMaxRows)
				if err != nil {
					return worksheetIndex{}, fmt.Errorf("row: %w", err)
				}
				rowNumber--
				if rowNumber <= lastRow {
					return worksheetIndex{}, fmt.Errorf("rows are not strictly ordered at row %d", rowNumber+1)
				}
				for targetCursor < len(targetRows) && targetRows[targetCursor] < rowNumber {
					index.insertions[targetRows[targetCursor]] = before
					targetCursor++
				}
				touched := targetSet[rowNumber]
				if touched && targetCursor < len(targetRows) && targetRows[targetCursor] == rowNumber {
					targetCursor++
				}
				lastRow = rowNumber
				row = &activeRow{span: rowSpan{xmlSpan: xmlSpan{qname: rawQName(data, before, after), start: before, startTagEnd: after}, row: rowNumber}, depth: depth, touched: touched}
				continue
			}
			if row != nil && row.touched && depth == row.depth+1 && token.Name.Space == index.namespace && token.Name.Local == "c" {
				ref := attribute(token, "r")
				cellRow, column, err := parseCellReference(ref)
				if err != nil {
					return worksheetIndex{}, fmt.Errorf("row %d cell: %w", row.span.row+1, err)
				}
				if cellRow != row.span.row {
					return worksheetIndex{}, fmt.Errorf("cell %q is stored in row %d", ref, row.span.row+1)
				}
				if row.hasColumn && column <= row.lastCol {
					return worksheetIndex{}, fmt.Errorf("cells are not strictly ordered in row %d", row.span.row+1)
				}
				row.hasColumn, row.lastCol = true, column
				cell = &activeCell{span: cellSpan{xmlSpan: xmlSpan{qname: rawQName(data, before, after), start: before, startTagEnd: after}, column: column, space: token.Name.Space}, depth: depth}
				continue
			}
			if cell != nil && depth == cell.depth+1 {
				child = &activeChild{span: namedSpan{local: token.Name.Local, space: token.Name.Space, xmlSpan: xmlSpan{qname: rawQName(data, before, after), start: before, startTagEnd: after}}, depth: depth}
			}
		case xml.EndElement:
			if index.dimension != nil && index.dimension.end == 0 && depth == worksheetDepth+1 && token.Name.Space == index.namespace && token.Name.Local == "dimension" {
				index.dimension.endStart, index.dimension.end = before, after
			}
			if child != nil && depth == child.depth {
				child.span.endStart, child.span.end = before, after
				cell.span.children = append(cell.span.children, child.span)
				child = nil
			}
			if cell != nil && depth == cell.depth {
				cell.span.endStart, cell.span.end = before, after
				row.span.cells = append(row.span.cells, cell.span)
				cell = nil
			}
			if row != nil && depth == row.depth {
				row.span.endStart, row.span.end = before, after
				if row.touched {
					index.existing[row.span.row] = row.span
				}
				row = nil
			}
			if foundSheetData && depth == sheetDepth && token.Name.Space == index.namespace && token.Name.Local == "sheetData" {
				index.sheetData.endStart, index.sheetData.end = before, after
				for targetCursor < len(targetRows) {
					index.insertions[targetRows[targetCursor]] = before
					targetCursor++
				}
			}
			depth--
		}
	}
	if !foundSheetData || index.sheetData.end == 0 {
		return worksheetIndex{}, fmt.Errorf("missing complete sheetData element")
	}
	return index, nil
}

func expandDimensionTag(tag []byte, ref string, mutations map[cellKey]normalizedCellMutation) ([]byte, error) {
	material := false
	for _, mutation := range mutations {
		if mutation.payload.kind != CellClearValue && mutation.payload.kind != CellClearFormula {
			material = true
			break
		}
	}
	if !material {
		return bytes.Clone(tag), nil
	}
	_, hasRef, err := rawUnqualifiedAttribute(tag, "ref")
	if err != nil {
		return nil, fmt.Errorf("parse dimension start tag: %w", err)
	}
	if ref == "" || !hasRef {
		return nil, fmt.Errorf("dimension element is missing a ref attribute")
	}
	minRow, minColumn, maxRow, maxColumn, err := parseDimensionReference(ref)
	if err != nil {
		return nil, err
	}
	for key, mutation := range mutations {
		if mutation.payload.kind == CellClearValue || mutation.payload.kind == CellClearFormula {
			continue
		}
		minRow = min(minRow, key.row)
		minColumn = min(minColumn, key.column)
		maxRow = max(maxRow, key.row)
		maxColumn = max(maxColumn, key.column)
	}
	expanded := cellReference(minRow, minColumn)
	if minRow != maxRow || minColumn != maxColumn {
		expanded += ":" + cellReference(maxRow, maxColumn)
	}
	if expanded == strings.ToUpper(strings.ReplaceAll(ref, "$", "")) {
		return bytes.Clone(tag), nil
	}
	return rewriteUnqualifiedAttribute(tag, "ref", &expanded)
}

func parseDimensionReference(ref string) (minRow, minColumn, maxRow, maxColumn int, err error) {
	ref = strings.ReplaceAll(ref, "$", "")
	parts := strings.Split(ref, ":")
	if len(parts) < 1 || len(parts) > 2 {
		return 0, 0, 0, 0, fmt.Errorf("invalid worksheet dimension %q", ref)
	}
	firstRow, firstColumn, err := parseCellReference(parts[0])
	if err != nil {
		return 0, 0, 0, 0, fmt.Errorf("invalid worksheet dimension %q: %w", ref, err)
	}
	lastRow, lastColumn := firstRow, firstColumn
	if len(parts) == 2 {
		lastRow, lastColumn, err = parseCellReference(parts[1])
		if err != nil {
			return 0, 0, 0, 0, fmt.Errorf("invalid worksheet dimension %q: %w", ref, err)
		}
	}
	if firstRow > lastRow || firstColumn > lastColumn {
		return 0, 0, 0, 0, fmt.Errorf("invalid worksheet dimension %q: bounds are reversed", ref)
	}
	return firstRow, firstColumn, lastRow, lastColumn, nil
}

func rebuildExistingRow(data []byte, row rowSpan, mutations map[int]normalizedCellMutation) ([]byte, error) {
	existing := make(map[int]bool, len(row.cells))
	for _, cell := range row.cells {
		existing[cell.column] = true
	}
	effective := false
	for column, mutation := range mutations {
		if existing[column] || (mutation.payload.kind != CellClearValue && mutation.payload.kind != CellClearFormula) {
			effective = true
			break
		}
	}
	if !effective {
		return bytes.Clone(data[row.start:row.end]), nil
	}

	columns := make([]int, 0, len(mutations))
	for column, mutation := range mutations {
		if mutation.payload.kind == CellClearValue || mutation.payload.kind == CellClearFormula {
			continue
		}
		columns = append(columns, column)
	}
	sort.Ints(columns)
	nextNew := 0
	insertingCell := false
	for _, column := range columns {
		if !existing[column] {
			insertingCell = true
			break
		}
	}

	rowStart := data[row.start:row.startTagEnd]
	if insertingCell {
		var err error
		rowStart, err = rewriteUnqualifiedAttribute(rowStart, "spans", nil)
		if err != nil {
			return nil, fmt.Errorf("rewrite row spans: %w", err)
		}
	}
	var out bytes.Buffer
	out.Write(openTag(rowStart))
	cursor := row.startTagEnd
	for _, cell := range row.cells {
		for nextNew < len(columns) && columns[nextNew] < cell.column {
			column := columns[nextNew]
			if !existing[column] {
				built, err := buildCell(prefixedLocal(row.qname, "c"), row.row, column, mutations[column].payload)
				if err != nil {
					return nil, err
				}
				out.Write(built)
			}
			nextNew++
		}
		out.Write(data[cursor:cell.start])
		if mutation, ok := mutations[cell.column]; ok {
			built, err := buildExistingCell(data, row.row, cell, mutation.payload)
			if err != nil {
				return nil, err
			}
			out.Write(built)
		} else {
			out.Write(data[cell.start:cell.end])
		}
		cursor = cell.end
		for nextNew < len(columns) && columns[nextNew] == cell.column {
			nextNew++
		}
	}
	for nextNew < len(columns) {
		column := columns[nextNew]
		if !existing[column] {
			built, err := buildCell(prefixedLocal(row.qname, "c"), row.row, column, mutations[column].payload)
			if err != nil {
				return nil, err
			}
			out.Write(built)
		}
		nextNew++
	}
	if !row.selfClosing(data) {
		out.Write(data[cursor:row.endStart])
	}
	out.WriteString("</" + row.qname + ">")
	return out.Bytes(), nil
}

func buildNewRow(sheetDataQName string, rowNumber int, mutations map[int]normalizedCellMutation) ([]byte, error) {
	columns := make([]int, 0)
	for column, mutation := range mutations {
		if mutation.payload.kind != CellClearValue && mutation.payload.kind != CellClearFormula {
			columns = append(columns, column)
		}
	}
	sort.Ints(columns)
	rowQName := prefixedLocal(sheetDataQName, "row")
	cellQName := prefixedLocal(sheetDataQName, "c")
	var out bytes.Buffer
	fmt.Fprintf(&out, "<%s r=\"%d\">", rowQName, rowNumber+1)
	for _, column := range columns {
		built, err := buildCell(cellQName, rowNumber, column, mutations[column].payload)
		if err != nil {
			return nil, err
		}
		out.Write(built)
	}
	out.WriteString("</" + rowQName + ">")
	return out.Bytes(), nil
}

func buildCell(qname string, row, column int, payload cellPayload) ([]byte, error) {
	start := []byte(fmt.Sprintf("<%s r=\"%s\">", qname, cellReference(row, column)))
	return assembleCell(start, qname, payload, nil)
}

func buildExistingCell(data []byte, row int, cell cellSpan, payload cellPayload) ([]byte, error) {
	start := bytes.Clone(data[cell.start:cell.startTagEnd])
	for _, name := range []string{"cm", "vm"} {
		if _, found, err := rawUnqualifiedAttribute(start, name); err != nil {
			return nil, fmt.Errorf("parse cell %s attributes: %w", cellReference(row, cell.column), err)
		} else if found {
			return nil, fmt.Errorf("cell %s carries %s metadata requiring a metadata-aware edit", cellReference(row, cell.column), name)
		}
	}
	for _, child := range cell.children {
		if child.space != cell.space {
			return nil, fmt.Errorf("cell %s has unsupported foreign-namespace direct child %q", cellReference(row, cell.column), child.qname)
		}
		switch child.local {
		case "f":
			formulaStart, err := decodeStartElement(data[child.start:child.startTagEnd])
			if err != nil {
				return nil, fmt.Errorf("parse existing formula element: %w", err)
			}
			if len(formulaStart.Attr) != 0 {
				return nil, fmt.Errorf("cell %s has a shared, array, or otherwise attributed formula requiring a group-aware edit", cellReference(row, cell.column))
			}
		case "v", "is":
		case "extLst":
		default:
			return nil, fmt.Errorf("cell %s has unsupported direct child %q", cell.qname, child.local)
		}
	}
	extensions := make([][]byte, 0, 1)
	for _, child := range cell.children {
		if child.local == "extLst" {
			extensions = append(extensions, data[child.start:child.end])
		}
	}
	return assembleCell(start, cell.qname, payload, extensions)
}

func decodeStartElement(data []byte) (xml.StartElement, error) {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	token, err := decoder.Token()
	if err != nil {
		return xml.StartElement{}, err
	}
	start, ok := token.(xml.StartElement)
	if !ok {
		return xml.StartElement{}, fmt.Errorf("expected start element")
	}
	return start, nil
}

func isSpreadsheetMLNamespace(namespace string) bool {
	return namespace == "" || namespace == spreadsheetMLTransitional || namespace == spreadsheetMLStrict
}

func worksheetNamespace(data []byte) (string, error) {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	for {
		token, err := decoder.Token()
		if err != nil {
			return "", fmt.Errorf("parse worksheet root: %w", err)
		}
		start, ok := token.(xml.StartElement)
		if !ok {
			continue
		}
		if start.Name.Local != "worksheet" || !isSpreadsheetMLNamespace(start.Name.Space) {
			return "", fmt.Errorf("root element is not a supported SpreadsheetML worksheet")
		}
		return start.Name.Space, nil
	}
}

func validateMergedCellTargets(data []byte, mutations map[cellKey]normalizedCellMutation) error {
	namespace, err := worksheetNamespace(data)
	if err != nil {
		return err
	}
	decoder := xml.NewDecoder(bytes.NewReader(data))
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			return nil
		}
		if err != nil {
			return fmt.Errorf("parse worksheet merges: %w", err)
		}
		start, ok := token.(xml.StartElement)
		if !ok || start.Name.Space != namespace || start.Name.Local != "mergeCell" {
			continue
		}
		ref := attribute(start, "ref")
		minRow, minColumn, maxRow, maxColumn, err := parseDimensionReference(ref)
		if err != nil {
			return fmt.Errorf("invalid merged range %q: %w", ref, err)
		}
		for key, mutation := range mutations {
			if key.row < minRow || key.row > maxRow || key.column < minColumn || key.column > maxColumn || (key.row == minRow && key.column == minColumn) {
				continue
			}
			return fmt.Errorf("operation %q targets %s inside merged range %s; only the top-left cell is editable", mutation.OperationID, cellReference(key.row, key.column), ref)
		}
	}
}

func validateFormulaGroupTargets(data []byte, mutations map[cellKey]normalizedCellMutation) error {
	namespace, err := worksheetNamespace(data)
	if err != nil {
		return err
	}
	decoder := xml.NewDecoder(bytes.NewReader(data))
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			return nil
		}
		if err != nil {
			return fmt.Errorf("parse worksheet formula groups: %w", err)
		}
		start, ok := token.(xml.StartElement)
		if !ok || start.Name.Space != namespace || start.Name.Local != "f" {
			continue
		}
		formulaType := attribute(start, "t")
		if formulaType != "array" && formulaType != "dataTable" && formulaType != "shared" {
			continue
		}
		ref := attribute(start, "ref")
		if ref == "" {
			continue
		}
		minRow, minColumn, maxRow, maxColumn, err := parseDimensionReference(ref)
		if err != nil {
			return fmt.Errorf("invalid %s formula range %q: %w", formulaType, ref, err)
		}
		for key, mutation := range mutations {
			if key.row < minRow || key.row > maxRow || key.column < minColumn || key.column > maxColumn {
				continue
			}
			return fmt.Errorf("operation %q targets %s inside %s formula range %s requiring a group-aware edit", mutation.OperationID, cellReference(key.row, key.column), formulaType, ref)
		}
	}
}

func assembleCell(start []byte, qname string, payload cellPayload, extensions [][]byte) ([]byte, error) {
	typeValue := ""
	valueXML := ""
	switch payload.kind {
	case CellSetValue:
		switch payload.literalKind {
		case literalNumber:
			valueXML = "<" + prefixedLocal(qname, "v") + ">" + payload.number + "</" + prefixedLocal(qname, "v") + ">"
		case literalString:
			escaped, err := encodeSpreadsheetString(payload.text)
			if err != nil {
				return nil, err
			}
			typeValue = "inlineStr"
			space := ""
			if strings.TrimSpace(payload.text) != payload.text {
				space = ` xml:space="preserve"`
			}
			isQName, textQName := prefixedLocal(qname, "is"), prefixedLocal(qname, "t")
			valueXML = "<" + isQName + "><" + textQName + space + ">" + escaped + "</" + textQName + "></" + isQName + ">"
		case literalBoolean:
			typeValue = "b"
			boolean := "0"
			if payload.boolean {
				boolean = "1"
			}
			vQName := prefixedLocal(qname, "v")
			valueXML = "<" + vQName + ">" + boolean + "</" + vQName + ">"
		default:
			return nil, fmt.Errorf("unsupported literal kind %q", payload.literalKind)
		}
	case CellSetFormula:
		escaped, err := escapeXMLText(payload.text)
		if err != nil {
			return nil, err
		}
		fQName := prefixedLocal(qname, "f")
		valueXML = "<" + fQName + ">" + escaped + "</" + fQName + ">"
	case CellClearValue, CellClearFormula:
	default:
		return nil, fmt.Errorf("unsupported payload kind %q", payload.kind)
	}

	var err error
	start, err = rewriteUnqualifiedAttribute(start, "t", nil)
	if err != nil {
		return nil, fmt.Errorf("rewrite cell type: %w", err)
	}
	if typeValue != "" {
		start, err = rewriteUnqualifiedAttribute(start, "t", &typeValue)
		if err != nil {
			return nil, fmt.Errorf("write cell type: %w", err)
		}
	}
	start = openTag(start)
	if valueXML == "" && len(extensions) == 0 {
		return append(start[:len(start)-1], []byte("/>")...), nil
	}
	var out bytes.Buffer
	out.Write(start)
	out.WriteString(valueXML)
	for _, extension := range extensions {
		out.Write(extension)
	}
	out.WriteString("</" + qname + ">")
	return out.Bytes(), nil
}

func hasMaterialMutation(mutations map[int]normalizedCellMutation) bool {
	for _, mutation := range mutations {
		if mutation.payload.kind != CellClearValue && mutation.payload.kind != CellClearFormula {
			return true
		}
	}
	return false
}

func openTag(tag []byte) []byte {
	trimmed := bytes.TrimSpace(tag)
	if bytes.HasSuffix(trimmed, []byte("/>")) {
		end := bytes.LastIndex(tag, []byte("/>"))
		return append(bytes.Clone(tag[:end]), '>')
	}
	return bytes.Clone(tag)
}

func rawQName(data []byte, start, end int) string {
	tag := data[start:end]
	left := bytes.IndexByte(tag, '<')
	if left < 0 {
		return ""
	}
	i := left + 1
	for i < len(tag) && (tag[i] == '/' || tag[i] == '!' || tag[i] == '?') {
		i++
	}
	j := i
	for j < len(tag) && tag[j] != ' ' && tag[j] != '\t' && tag[j] != '\r' && tag[j] != '\n' && tag[j] != '/' && tag[j] != '>' {
		j++
	}
	return string(tag[i:j])
}

func prefixedLocal(qname, local string) string {
	if index := strings.IndexByte(qname, ':'); index >= 0 {
		return qname[:index+1] + local
	}
	return local
}

func attribute(start xml.StartElement, local string) string {
	for _, attr := range start.Attr {
		if attr.Name.Local == local {
			return attr.Value
		}
	}
	return ""
}

func positiveIndexAttribute(start xml.StartElement, local string, max int) (int, error) {
	raw := attribute(start, local)
	value, err := strconv.Atoi(raw)
	if err != nil || value < 1 || value > max {
		return 0, fmt.Errorf("attribute %s=%q is outside 1..%d", local, raw, max)
	}
	return value, nil
}

func parseCellReference(ref string) (row, column int, err error) {
	if ref == "" {
		return 0, 0, fmt.Errorf("missing r attribute")
	}
	index := 0
	columnValue := 0
	for index < len(ref) && ((ref[index] >= 'A' && ref[index] <= 'Z') || (ref[index] >= 'a' && ref[index] <= 'z')) {
		ch := ref[index]
		if ch >= 'a' {
			ch -= 'a' - 'A'
		}
		columnValue = columnValue*26 + int(ch-'A'+1)
		index++
	}
	if index == 0 || index == len(ref) {
		return 0, 0, fmt.Errorf("invalid cell reference %q", ref)
	}
	rowValue, parseErr := strconv.Atoi(ref[index:])
	if parseErr != nil || rowValue < 1 || rowValue > excelMaxRows || columnValue < 1 || columnValue > excelMaxColumns {
		return 0, 0, fmt.Errorf("invalid cell reference %q", ref)
	}
	return rowValue - 1, columnValue - 1, nil
}

func cellReference(row, column int) string { return columnName(column) + strconv.Itoa(row+1) }

func columnName(column int) string {
	column++
	var reversed [4]byte
	length := 0
	for column > 0 {
		column--
		reversed[length] = byte('A' + column%26)
		length++
		column /= 26
	}
	var out [4]byte
	for i := 0; i < length; i++ {
		out[i] = reversed[length-i-1]
	}
	return string(out[:length])
}

func validateWorksheetXML(data []byte) error {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	for {
		_, err := decoder.Token()
		if err == io.EOF {
			return nil
		}
		if err != nil {
			return err
		}
	}
}
