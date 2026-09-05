package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/xml"
	"fmt"
	"io"
	"math"
	"net/url"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"unicode/utf16"
	"unicode/utf8"
)

var nativeXMLFiniteDoublePattern = regexp.MustCompile(`^[+-]?(?:(?:[0-9]+(?:\.[0-9]*)?)|(?:\.[0-9]+))(?:[eE][+-]?[0-9]+)?$`)

const (
	NativeXLSXMaxPackageBytes       = 128 * 1024 * 1024
	NativeXLSXMaxPartBytes          = 64 * 1024 * 1024
	NativeXLSXMaxXMLPartBytes       = 64 * 1024 * 1024
	NativeXLSXMaxUncompressedBytes  = 256 * 1024 * 1024
	NativeXLSXMaxPackageParts       = 10_000
	NativeXLSXMaxCompressionRatio   = 200
	NativeXLSXCompressionRatioSlack = 1 * 1024 * 1024
	NativeXLSXMaxRelationships      = 100_000
	NativeXLSXMaxXMLDepth           = 128
)

const (
	nativeContentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types"
	nativeWorkbookContentType   = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
	nativeWorksheetContentType  = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
	nativeSharedStringsType     = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
	nativeRelationshipsType     = "application/vnd.openxmlformats-package.relationships+xml"

	relTypeSharedStringsTransitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
	relTypeSharedStringsStrict       = "http://purl.oclc.org/ooxml/officeDocument/relationships/sharedStrings"
)

type NativeWorkbookExtractionOptions struct {
	Previous   *NativeWorkbookV1
	DocumentID string
}

type nativeWorkbookPackage struct {
	index            *opcPackageIndex
	files            map[string][]byte
	contentTypesPart string
	contentTypes     nativeContentTypeRegistry
}

type nativeContentTypeRegistry struct {
	defaults  map[string]string
	overrides map[string]string
}

type nativeWorkbookSheetRoute struct {
	id, name, state, relID, part, refusalCode string
}

type nativeWorkbookExtractor struct {
	pkg               *nativeWorkbookPackage
	workbook          workbookPartLocation
	namespace         string
	relNamespace      string
	modeled           map[string]bool
	unsupported       []NativeWorkbookUnsupportedV1
	unsupportedKeys   map[string]bool
	textLength        int
	cellCount         int
	mergedRangeCount  int
	relationshipCount int
	xmlTokens         int
	xmlElements       int
	claimedXML        map[string]bool
}

// ExtractNativeWorkbookV1 is the renderer-free XLSX import boundary. It reads
// OOXML directly, never evaluates formulas, and leaves the source archive as
// the authority for unsupported content.
func ExtractNativeWorkbookV1(data []byte) (*NativeWorkbookV1, error) {
	return ExtractNativeWorkbookV1WithOptions(data, NativeWorkbookExtractionOptions{})
}

// ExtractNativeWorkbookV1WithOptions retains application/document identity
// across revisions. Previous is accepted only when it is a valid v1 contract
// for the same canonical workbook and does not remap an existing sheet id.
func ExtractNativeWorkbookV1WithOptions(data []byte, options NativeWorkbookExtractionOptions) (*NativeWorkbookV1, error) {
	pkg, err := openNativeWorkbookPackage(data)
	if err != nil {
		return nil, err
	}
	rootRelationshipsPart, _, found := pkg.index.lookupSpelling("_rels/.rels")
	if !found {
		return nil, fmt.Errorf("xlsxpatch: native extract: missing root package relationships")
	}
	if _, _, err := preflightNativeCoreXML(pkg.files[rootRelationshipsPart]); err != nil {
		return nil, fmt.Errorf("xlsxpatch: native extract: root package relationships: %w", err)
	}
	readBytes := func(name string) ([]byte, bool) {
		value, ok := pkg.files[name]
		return value, ok
	}
	workbook, err := locateWorkbookPartBytes(pkg.index, readBytes)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: native extract: %w", err)
	}
	if err := requireNativeContentType(pkg, workbook.part, nativeWorkbookContentType); err != nil {
		return nil, fmt.Errorf("xlsxpatch: native extract: workbook: %w", err)
	}
	workbookXML := pkg.files[workbook.part]
	if len(workbookXML) == 0 || len(workbookXML) > NativeXLSXMaxXMLPartBytes {
		return nil, fmt.Errorf("xlsxpatch: native extract: workbook XML size must be 1..%d bytes", NativeXLSXMaxXMLPartBytes)
	}

	namespace, relNamespace := spreadsheetMLTransitional, officeRelNamespaceTransitional
	if workbook.strict {
		namespace, relNamespace = spreadsheetMLStrict, officeRelNamespaceStrict
	}
	extractor := &nativeWorkbookExtractor{
		pkg: pkg, workbook: workbook, namespace: namespace, relNamespace: relNamespace,
		modeled: map[string]bool{}, unsupported: []NativeWorkbookUnsupportedV1{}, unsupportedKeys: map[string]bool{}, claimedXML: map[string]bool{},
	}
	for _, part := range []string{pkg.contentTypesPart, workbook.part, workbook.relsPart} {
		if err := extractor.claimCoreXML(part); err != nil {
			return nil, err
		}
	}
	if rootPart, _, ok := pkg.index.lookupSpelling("_rels/.rels"); ok {
		if err := extractor.claimCoreXML(rootPart); err != nil {
			return nil, err
		}
	}
	for _, part := range []string{pkg.contentTypesPart, workbook.part, workbook.relsPart} {
		extractor.markModeled(part)
	}
	if rootPart, _, ok := pkg.index.lookupSpelling("_rels/.rels"); ok {
		extractor.markModeled(rootPart)
	}
	if err := extractor.validateAllRelationships(); err != nil {
		return nil, err
	}
	routes, err := extractor.extractWorkbookRoutes(workbookXML)
	if err != nil {
		return nil, err
	}
	if err := validatePreviousNativeWorkbook(options, workbook, routes); err != nil {
		return nil, err
	}

	sharedStrings, sharedPart, err := extractor.extractSharedStrings()
	if err != nil {
		return nil, err
	}
	if sharedPart != "" {
		extractor.markModeled(sharedPart)
	}
	styles, normalStyle, registry, stylesPart, err := extractor.extractStyles()
	if err != nil {
		return nil, err
	}
	if stylesPart != "" {
		extractor.markModeled(stylesPart)
	}

	sheets := make([]NativeWorkbookSheetV1, 0, len(routes))
	seenParts := make(map[string]string, len(routes))
	for order, route := range routes {
		key, keyErr := canonicalOPCPartKey(route.part)
		if keyErr != nil {
			return nil, fmt.Errorf("xlsxpatch: native extract: worksheet %q: %w", route.part, keyErr)
		}
		if prior, duplicate := seenParts[key]; duplicate {
			return nil, fmt.Errorf("xlsxpatch: native extract: sheets %q and %q route to the same worksheet part %q", prior, route.id, route.part)
		}
		seenParts[key] = route.id
		extractor.markModeled(route.part)
		sheetXML := pkg.files[route.part]
		if err := extractor.claimCoreXML(route.part); err != nil {
			return nil, err
		}
		if len(sheetXML) == 0 || len(sheetXML) > NativeXLSXMaxXMLPartBytes {
			return nil, fmt.Errorf("xlsxpatch: native extract: worksheet %q size must be 1..%d bytes", route.part, NativeXLSXMaxXMLPartBytes)
		}
		sheet, sheetErr := extractor.extractWorksheet(route, order, sheetXML, sharedStrings, len(styles), registry != nil)
		if sheetErr != nil {
			return nil, fmt.Errorf("xlsxpatch: native extract: worksheet %q: %w", route.part, sheetErr)
		}
		sheets = append(sheets, sheet)
	}
	if extractor.cellCount > NativeXLSXMaxCells {
		return nil, fmt.Errorf("xlsxpatch: native extract: workbook exceeds %d modeled cells", NativeXLSXMaxCells)
	}

	passthrough, err := extractor.passthroughParts()
	if err != nil {
		return nil, err
	}
	digest := nativeWorkbookDigest(data)
	documentID := options.DocumentID
	if documentID == "" && options.Previous != nil {
		documentID = options.Previous.DocumentID
	}
	if documentID == "" {
		identitySeed := append([]byte(asciiLower(workbook.part)+"\x00"), workbookXML...)
		documentID = "workbook:" + strings.TrimPrefix(nativeWorkbookDigest(identitySeed), "sha256:")
	}
	dialect := "transitional"
	if workbook.strict {
		dialect = "strict"
	}
	result := &NativeWorkbookV1{
		Protocol: NativeXLSXProtocol, Version: NativeXLSXVersion,
		DocumentID: documentID, Revision: "rev:" + strings.TrimPrefix(digest, "sha256:"),
		Source:      NativeWorkbookSourceV1{PackageSHA256: digest, WorkbookPart: workbook.part, Dialect: dialect, Authority: "exact-package-bytes"},
		NormalStyle: normalStyle, Sheets: sheets, Styles: styles,
		Capabilities: []NativeWorkbookCapabilityV1{
			{Name: "native-ooxml-parse", Level: "read-only", Detail: nativeWorkbookString("HTML-free SpreadsheetML extraction with lexical values and formula caches")},
			{Name: "native-v1-mutations", Level: "partial", Detail: nativeWorkbookString("Cells, formulas, row/column dimensions, and the bounded style.patch projection")},
			{Name: "unsupported-content", Level: "preserve-exact", Detail: nativeWorkbookString("Unmodeled OOXML remains authoritative in the original package")},
		},
		PassthroughParts: passthrough, Unsupported: extractor.unsupported,
	}
	if issues := ValidateNativeWorkbookV1(result); len(issues) != 0 {
		return nil, fmt.Errorf("xlsxpatch: native extract produced invalid contract: %w", &NativeWorkbookValidationError{Issues: issues})
	}
	if _, err := EncodeNativeWorkbookV1(result); err != nil {
		return nil, err
	}
	return result, nil
}

func openNativeWorkbookPackage(data []byte) (*nativeWorkbookPackage, error) {
	if len(data) == 0 || len(data) > NativeXLSXMaxPackageBytes {
		return nil, fmt.Errorf("xlsxpatch: native extract: package size must be 1..%d bytes", NativeXLSXMaxPackageBytes)
	}
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: native extract: invalid ZIP: %w", err)
	}
	if len(zr.File) == 0 || len(zr.File) > NativeXLSXMaxPackageParts {
		return nil, fmt.Errorf("xlsxpatch: native extract: ZIP contains %d entries; limit is %d", len(zr.File), NativeXLSXMaxPackageParts)
	}
	index, err := newOPCPackageIndex(zr)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: native extract: %w", err)
	}
	files := make(map[string][]byte, len(zr.File))
	var declaredTotal, actualTotal uint64
	for _, file := range zr.File {
		if file.FileInfo().IsDir() || strings.HasSuffix(file.Name, "/") {
			trimmed := strings.TrimSuffix(file.Name, "/")
			if trimmed == "" {
				return nil, fmt.Errorf("xlsxpatch: native extract: invalid root ZIP directory")
			}
			if _, err := canonicalOPCPartKey(trimmed); err != nil {
				return nil, fmt.Errorf("xlsxpatch: native extract: non-conforming ZIP directory %q: %w", file.Name, err)
			}
			continue
		}
		if file.Flags&0x1 != 0 {
			return nil, fmt.Errorf("xlsxpatch: native extract: encrypted ZIP entry %q is unsupported", file.Name)
		}
		if file.Method != zip.Store && file.Method != zip.Deflate {
			return nil, fmt.Errorf("xlsxpatch: native extract: ZIP entry %q uses unsupported compression method %d", file.Name, file.Method)
		}
		if file.UncompressedSize64 > NativeXLSXMaxPartBytes {
			return nil, fmt.Errorf("xlsxpatch: native extract: ZIP entry %q exceeds %d bytes", file.Name, NativeXLSXMaxPartBytes)
		}
		if file.CompressedSize64 == 0 {
			if file.UncompressedSize64 > NativeXLSXCompressionRatioSlack {
				return nil, fmt.Errorf("xlsxpatch: native extract: ZIP entry %q exceeds compression-ratio limit", file.Name)
			}
		} else if file.UncompressedSize64 > file.CompressedSize64*NativeXLSXMaxCompressionRatio+NativeXLSXCompressionRatioSlack {
			return nil, fmt.Errorf("xlsxpatch: native extract: ZIP entry %q exceeds compression-ratio limit", file.Name)
		}
		declaredTotal += file.UncompressedSize64
		if declaredTotal > NativeXLSXMaxUncompressedBytes {
			return nil, fmt.Errorf("xlsxpatch: native extract: ZIP exceeds %d declared uncompressed bytes", NativeXLSXMaxUncompressedBytes)
		}
		rc, err := file.Open()
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: native extract: open ZIP entry %q: %w", file.Name, err)
		}
		content, readErr := io.ReadAll(io.LimitReader(rc, NativeXLSXMaxPartBytes+1))
		closeErr := rc.Close()
		if readErr != nil {
			return nil, fmt.Errorf("xlsxpatch: native extract: read ZIP entry %q: %w", file.Name, readErr)
		}
		if closeErr != nil {
			return nil, fmt.Errorf("xlsxpatch: native extract: close ZIP entry %q: %w", file.Name, closeErr)
		}
		if len(content) > NativeXLSXMaxPartBytes {
			return nil, fmt.Errorf("xlsxpatch: native extract: ZIP entry %q exceeds %d actual bytes", file.Name, NativeXLSXMaxPartBytes)
		}
		if uint64(len(content)) != file.UncompressedSize64 {
			return nil, fmt.Errorf("xlsxpatch: native extract: ZIP entry %q actual length %d does not match declared uncompressed size %d", file.Name, len(content), file.UncompressedSize64)
		}
		actualTotal += uint64(len(content))
		if actualTotal > NativeXLSXMaxUncompressedBytes {
			return nil, fmt.Errorf("xlsxpatch: native extract: ZIP exceeds %d actual uncompressed bytes", NativeXLSXMaxUncompressedBytes)
		}
		files[file.Name] = content
	}
	contentTypesPart, _, ok := index.lookupSpelling("[Content_Types].xml")
	if !ok {
		return nil, fmt.Errorf("xlsxpatch: native extract: missing [Content_Types].xml")
	}
	contentTypesData := files[contentTypesPart]
	if len(contentTypesData) == 0 || len(contentTypesData) > NativeXLSXMaxXMLPartBytes {
		return nil, fmt.Errorf("xlsxpatch: native extract: content-types XML size must be 1..%d bytes", NativeXLSXMaxXMLPartBytes)
	}
	contentTypes, err := parseNativeContentTypes(contentTypesData)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: native extract: %w", err)
	}
	pkg := &nativeWorkbookPackage{index: index, files: files, contentTypesPart: contentTypesPart, contentTypes: contentTypes}
	overrideKeys := make([]string, 0, len(contentTypes.overrides))
	for overrideKey := range contentTypes.overrides {
		overrideKeys = append(overrideKeys, overrideKey)
	}
	sort.Strings(overrideKeys)
	for _, overrideKey := range overrideKeys {
		if _, found := index.byKey[overrideKey]; !found {
			return nil, fmt.Errorf("xlsxpatch: native extract: content type Override targets missing part %q", overrideKey)
		}
	}
	fileNames := make([]string, 0, len(files))
	for name := range files {
		fileNames = append(fileNames, name)
	}
	sort.Slice(fileNames, func(i, j int) bool { return asciiLower(fileNames[i]) < asciiLower(fileNames[j]) })
	for _, name := range fileNames {
		if name == contentTypesPart {
			continue
		}
		if _, err := pkg.contentType(name); err != nil {
			return nil, fmt.Errorf("xlsxpatch: native extract: %w", err)
		}
	}
	return pkg, nil
}

func parseNativeContentTypes(data []byte) (nativeContentTypeRegistry, error) {
	registry := nativeContentTypeRegistry{defaults: map[string]string{}, overrides: map[string]string{}}
	if _, _, err := preflightNativeCoreXML(data); err != nil {
		return registry, fmt.Errorf("preflight [Content_Types].xml: %w", err)
	}
	decoder := xml.NewDecoder(bytes.NewReader(data))
	depth, entries := 0, 0
	rootSeen, rootClosed := false, false
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return registry, fmt.Errorf("parse [Content_Types].xml: %w", err)
		}
		switch token := token.(type) {
		case xml.StartElement:
			depth++
			if depth == 1 {
				if rootSeen || token.Name != (xml.Name{Space: nativeContentTypesNamespace, Local: "Types"}) {
					return registry, fmt.Errorf("content types root is not package Types")
				}
				rootSeen = true
				if err := requireOnlySemanticXMLAttributes(token); err != nil {
					return registry, err
				}
				continue
			}
			if depth != 2 || token.Name.Space != nativeContentTypesNamespace || (token.Name.Local != "Default" && token.Name.Local != "Override") {
				return registry, fmt.Errorf("content types has unsupported nested/direct element {%s}%s", token.Name.Space, token.Name.Local)
			}
			entries++
			if entries > NativeXLSXMaxPackageParts*2 {
				return registry, fmt.Errorf("content types exceeds %d declarations", NativeXLSXMaxPackageParts*2)
			}
			contentType, found, attrErr := unqualifiedXMLAttribute(token, "ContentType")
			if attrErr != nil || !found || contentType == "" || len(contentType) > 1024 || strings.TrimSpace(contentType) != contentType {
				return registry, fmt.Errorf("content type declaration requires one non-empty ContentType")
			}
			if token.Name.Local == "Default" {
				if err := requireOnlySemanticXMLAttributes(token, xml.Name{Local: "Extension"}, xml.Name{Local: "ContentType"}); err != nil {
					return registry, err
				}
				extension, found, attrErr := unqualifiedXMLAttribute(token, "Extension")
				if attrErr != nil || !found || extension == "" || len(extension) > 255 || strings.ContainsAny(extension, "/\\") {
					return registry, fmt.Errorf("content type Default requires a valid Extension")
				}
				key := asciiLower(extension)
				if _, duplicate := registry.defaults[key]; duplicate {
					return registry, fmt.Errorf("duplicate case-equivalent content type Default for %q", extension)
				}
				registry.defaults[key] = contentType
			} else {
				if err := requireOnlySemanticXMLAttributes(token, xml.Name{Local: "PartName"}, xml.Name{Local: "ContentType"}); err != nil {
					return registry, err
				}
				partName, found, attrErr := unqualifiedXMLAttribute(token, "PartName")
				if attrErr != nil || !found || len(partName) > maxRoutingRelationshipTargetLength || !strings.HasPrefix(partName, "/") {
					return registry, fmt.Errorf("content type Override requires a root-relative PartName")
				}
				key, keyErr := canonicalOPCPartKey(partName)
				if keyErr != nil {
					return registry, fmt.Errorf("content type Override PartName %q: %w", partName, keyErr)
				}
				if _, duplicate := registry.overrides[key]; duplicate {
					return registry, fmt.Errorf("duplicate case/escape-equivalent content type Override for %q", partName)
				}
				registry.overrides[key] = contentType
			}
		case xml.EndElement:
			if depth == 1 && token.Name == (xml.Name{Space: nativeContentTypesNamespace, Local: "Types"}) {
				rootClosed = true
			}
			depth--
		case xml.CharData:
			if depth > 0 && len(bytes.TrimSpace(token)) != 0 {
				return registry, fmt.Errorf("content types contains unsupported text")
			}
		case xml.ProcInst:
			if depth == 0 && !rootSeen && validNativeXMLDeclaration(token) {
				continue
			}
			return registry, fmt.Errorf("content types contains unsupported processing instruction %q", token.Target)
		case xml.Directive:
			return registry, fmt.Errorf("content types contains unsupported XML directive")
		}
	}
	if depth != 0 || !rootSeen || !rootClosed {
		return registry, fmt.Errorf("content types has no complete Types root")
	}
	return registry, nil
}

func (pkg *nativeWorkbookPackage) contentType(partName string) (string, error) {
	key, err := canonicalOPCPartKey(partName)
	if err != nil {
		return "", fmt.Errorf("invalid part %q for content-type lookup: %w", partName, err)
	}
	if value, found := pkg.contentTypes.overrides[key]; found {
		return value, nil
	}
	decoded := strings.TrimPrefix(partName, "/")
	if parsed, parseErr := url.Parse(decoded); parseErr == nil {
		decoded = parsed.Path
	}
	lastSlash, lastDot := strings.LastIndex(decoded, "/"), strings.LastIndex(decoded, ".")
	if lastDot <= lastSlash || lastDot == len(decoded)-1 {
		return "", fmt.Errorf("part %q has no effective content type", partName)
	}
	value, found := pkg.contentTypes.defaults[asciiLower(decoded[lastDot+1:])]
	if !found {
		return "", fmt.Errorf("part %q has no effective content type", partName)
	}
	return value, nil
}

func requireNativeContentType(pkg *nativeWorkbookPackage, partName, expected string) error {
	actual, err := pkg.contentType(partName)
	if err != nil {
		return err
	}
	if !asciiEqualFold(actual, expected) {
		return fmt.Errorf("part %q has content type %q; expected %q", partName, actual, expected)
	}
	return nil
}

func validatePreviousNativeWorkbook(options NativeWorkbookExtractionOptions, workbook workbookPartLocation, routes []nativeWorkbookSheetRoute) error {
	if options.DocumentID != "" && !validNativeWorkbookIdentifier(options.DocumentID) {
		return fmt.Errorf("xlsxpatch: native extract: invalid explicit document identity %q", options.DocumentID)
	}
	if options.Previous == nil {
		return nil
	}
	if issues := ValidateNativeWorkbookV1(options.Previous); len(issues) != 0 {
		return fmt.Errorf("xlsxpatch: native extract: previous identity contract is invalid: %w", &NativeWorkbookValidationError{Issues: issues})
	}
	previousKey, previousErr := canonicalOPCPartKey(options.Previous.Source.WorkbookPart)
	currentKey, currentErr := canonicalOPCPartKey(workbook.part)
	if previousErr != nil || currentErr != nil || previousKey != currentKey {
		return fmt.Errorf("xlsxpatch: native extract: previous workbook part %q does not match %q", options.Previous.Source.WorkbookPart, workbook.part)
	}
	currentDialect := "transitional"
	if workbook.strict {
		currentDialect = "strict"
	}
	if options.Previous.Source.Dialect != currentDialect {
		return fmt.Errorf("xlsxpatch: native extract: Previous and current workbook use opposing Strict/Transitional dialects")
	}
	if options.DocumentID != "" && options.DocumentID != options.Previous.DocumentID {
		return fmt.Errorf("xlsxpatch: native extract: explicit document identity conflicts with Previous")
	}
	previousByID := make(map[string]NativeWorkbookSheetV1, len(options.Previous.Sheets))
	previousByPart := make(map[string]NativeWorkbookSheetV1, len(options.Previous.Sheets))
	for _, sheet := range options.Previous.Sheets {
		previousByID[sheet.ID] = sheet
		partKey, err := canonicalOPCPartKey(sheet.PartName)
		if err == nil {
			previousByPart[partKey] = sheet
		}
	}
	for _, route := range routes {
		currentPart, currentErr := canonicalOPCPartKey(route.part)
		if currentErr != nil {
			return fmt.Errorf("xlsxpatch: native extract: current sheet part %q is invalid: %w", route.part, currentErr)
		}
		previous, found := previousByID[route.id]
		if found {
			previousPart, errA := canonicalOPCPartKey(previous.PartName)
			if errA != nil || previousPart != currentPart {
				return fmt.Errorf("xlsxpatch: native extract: Previous sheet id %q maps to %q, not current part %q", route.id, previous.PartName, route.part)
			}
		}
		if previous, found := previousByPart[currentPart]; found && previous.ID != route.id {
			return fmt.Errorf("xlsxpatch: native extract: Previous worksheet part %q has sheet id %q, not current id %q", route.part, previous.ID, route.id)
		}
	}
	return nil
}

func nativeWorkbookDigest(data []byte) string {
	digest := sha256.Sum256(data)
	return "sha256:" + hex.EncodeToString(digest[:])
}

func (extractor *nativeWorkbookExtractor) markModeled(part string) {
	key, err := canonicalOPCPartKey(part)
	if err == nil {
		extractor.modeled[key] = true
	}
}

func (extractor *nativeWorkbookExtractor) addUnsupported(code, capability, scope, part, cell, message string) error {
	return extractor.addUnsupportedLocation(code, capability, scope, part, cell, "", message)
}

func (extractor *nativeWorkbookExtractor) addUnsupportedRange(code, capability, scope, part, rangeRef, message string) error {
	return extractor.addUnsupportedLocation(code, capability, scope, part, "", rangeRef, message)
}

func (extractor *nativeWorkbookExtractor) addUnsupportedLocation(code, capability, scope, part, cell, rangeRef, message string) error {
	key := strings.Join([]string{code, capability, scope, part, cell, rangeRef}, "\x00")
	if extractor.unsupportedKeys[key] {
		return nil
	}
	if len(extractor.unsupported) >= NativeXLSXMaxInventory {
		return fmt.Errorf("xlsxpatch: native extract: unsupported inventory exceeds %d entries", NativeXLSXMaxInventory)
	}
	extractor.unsupportedKeys[key] = true
	seed := sha256.Sum256([]byte(key))
	item := NativeWorkbookUnsupportedV1{
		ID: "unsupported:" + hex.EncodeToString(seed[:]), Code: code, Capability: capability,
		ScopeID: scope, Preservation: "preserve-exact", Message: message,
	}
	if part != "" {
		item.PartName = nativeWorkbookString(part)
	}
	if cell != "" {
		item.CellRef = nativeWorkbookString(cell)
	}
	if rangeRef != "" {
		item.RangeRef = nativeWorkbookString(rangeRef)
	}
	extractor.unsupported = append(extractor.unsupported, item)
	return nil
}

func (extractor *nativeWorkbookExtractor) claimNativeText(value string, maximumUTF16 int) error {
	if utf16Length(value) > maximumUTF16 {
		return fmt.Errorf("text exceeds %d UTF-16 code units", maximumUTF16)
	}
	extractor.textLength += len(value)
	if extractor.textLength > NativeXLSXMaxTextLength {
		return fmt.Errorf("workbook text exceeds %d bytes", NativeXLSXMaxTextLength)
	}
	return nil
}

func (extractor *nativeWorkbookExtractor) claimCoreXML(partName string) error {
	key, err := canonicalOPCPartKey(partName)
	if err != nil {
		return err
	}
	if extractor.claimedXML[key] {
		return nil
	}
	data, found := extractor.pkg.files[partName]
	if !found {
		return fmt.Errorf("xlsxpatch: native extract: core XML part %q is missing", partName)
	}
	tokens, elements, err := preflightNativeCoreXML(data)
	if err != nil {
		return fmt.Errorf("xlsxpatch: native extract: core XML part %q: %w", partName, err)
	}
	if tokens > NativeXLSXMaxXMLTokens-extractor.xmlTokens || elements > NativeXLSXMaxXMLElements-extractor.xmlElements {
		return fmt.Errorf("xlsxpatch: native extract: routed/core XML exceeds cumulative %d-token/%d-element limits", NativeXLSXMaxXMLTokens, NativeXLSXMaxXMLElements)
	}
	extractor.xmlTokens += tokens
	extractor.xmlElements += elements
	extractor.claimedXML[key] = true
	return nil
}

func (extractor *nativeWorkbookExtractor) passthroughParts() ([]NativeWorkbookPassthroughPartV1, error) {
	parts := make([]NativeWorkbookPassthroughPartV1, 0)
	names := make([]string, 0, len(extractor.pkg.files))
	for name := range extractor.pkg.files {
		key, err := canonicalOPCPartKey(name)
		if err != nil || extractor.modeled[key] {
			continue
		}
		names = append(names, name)
	}
	sort.Slice(names, func(i, j int) bool { return asciiLower(names[i]) < asciiLower(names[j]) })
	for _, name := range names {
		if len(parts) >= NativeXLSXMaxInventory {
			return nil, fmt.Errorf("xlsxpatch: native extract: passthrough inventory exceeds %d entries", NativeXLSXMaxInventory)
		}
		contentType, err := extractor.pkg.contentType(name)
		if err != nil {
			return nil, err
		}
		content := extractor.pkg.files[name]
		parts = append(parts, NativeWorkbookPassthroughPartV1{
			PartName: name, ContentType: contentType, ByteLength: nativeWorkbookInt64(int64(len(content))),
			SHA256: nativeWorkbookDigest(content), Policy: "preserve-exact",
		})
		if code, capability := unsupportedPartCapability(contentType, name); code != "" {
			if err := extractor.addUnsupported(code, capability, "workbook", name, "", "part is preserved exactly and is outside the v1 native mutation projection"); err != nil {
				return nil, err
			}
		}
	}
	return parts, nil
}

func unsupportedPartCapability(contentType, name string) (string, string) {
	value := asciiLower(contentType + " " + name)
	switch {
	case strings.Contains(value, "chart"):
		return "CHART_CONTENT", "charts"
	case strings.Contains(value, "pivot"), strings.Contains(value, "slicer"), strings.Contains(value, "timeline"):
		return "PIVOT_OR_SLICER_CONTENT", "pivots"
	case strings.Contains(value, "drawing"), strings.Contains(value, "vml"), strings.Contains(value, "image"):
		return "DRAWING_OR_MEDIA_CONTENT", "drawings"
	case strings.Contains(value, "table"):
		return "TABLE_CONTENT", "tables"
	case strings.Contains(value, "comment"):
		return "COMMENT_CONTENT", "comments"
	case strings.Contains(value, "externallink"):
		return "EXTERNAL_LINK_CONTENT", "external-links"
	case strings.Contains(value, "vba"), strings.Contains(value, "macro"):
		return "MACRO_CONTENT", "macros"
	default:
		return "OPAQUE_PACKAGE_PART", "opaque-parts"
	}
}

func (extractor *nativeWorkbookExtractor) validateAllRelationships() error {
	names := make([]string, 0)
	for partName := range extractor.pkg.files {
		key, err := canonicalOPCPartKey(partName)
		if err == nil && strings.HasSuffix(key, ".rels") {
			names = append(names, partName)
		}
	}
	sort.Slice(names, func(i, j int) bool { return asciiLower(names[i]) < asciiLower(names[j]) })
	for _, partName := range names {
		data := extractor.pkg.files[partName]
		if err := extractor.claimCoreXML(partName); err != nil {
			return err
		}
		if len(data) > NativeXLSXMaxXMLPartBytes {
			return fmt.Errorf("xlsxpatch: native extract: relationship part %q exceeds %d bytes", partName, NativeXLSXMaxXMLPartBytes)
		}
		if err := requireNativeContentType(extractor.pkg, partName, nativeRelationshipsType); err != nil {
			return fmt.Errorf("xlsxpatch: native extract: relationship part %q: %w", partName, err)
		}
		relationships, err := parseRoutingRelationships(data)
		if err != nil {
			return fmt.Errorf("xlsxpatch: native extract: relationship part %q: %w", partName, err)
		}
		if len(relationships) > NativeXLSXMaxRelationships {
			return fmt.Errorf("xlsxpatch: native extract: relationship part %q exceeds %d relationships", partName, NativeXLSXMaxRelationships)
		}
		if len(relationships) > NativeXLSXMaxRelationships-extractor.relationshipCount {
			return fmt.Errorf("xlsxpatch: native extract: package relationships exceed cumulative limit %d", NativeXLSXMaxRelationships)
		}
		extractor.relationshipCount += len(relationships)
		owner, base, err := nativeRelationshipOwner(partName)
		if err != nil {
			return fmt.Errorf("xlsxpatch: native extract: relationship part %q: %w", partName, err)
		}
		if owner != "" {
			if _, _, found := extractor.pkg.index.lookupResolved(owner); !found {
				return fmt.Errorf("xlsxpatch: native extract: relationship part %q has missing owner %q", partName, owner)
			}
		}
		for _, relationship := range relationships {
			if strings.HasPrefix(relationship.relType, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/") && extractor.workbook.strict {
				return fmt.Errorf("xlsxpatch: native extract: opposing Strict/Transitional relationship dialect: relationship %q in %q uses a Transitional Office relationship type in a Strict workbook", relationship.id, partName)
			}
			if strings.HasPrefix(relationship.relType, "http://purl.oclc.org/ooxml/officeDocument/relationships/") && !extractor.workbook.strict {
				return fmt.Errorf("xlsxpatch: native extract: opposing Strict/Transitional relationship dialect: relationship %q in %q uses a Strict Office relationship type in a Transitional workbook", relationship.id, partName)
			}
			switch relationship.targetMode {
			case "", "Internal":
				resolved, resolveErr := resolveNativeRelationshipTarget(owner, base, relationship.target)
				if resolveErr != nil {
					return fmt.Errorf("xlsxpatch: native extract: relationship %q in %q: %w", relationship.id, partName, resolveErr)
				}
				if _, _, found := extractor.pkg.index.lookupResolved(resolved); !found {
					return fmt.Errorf("xlsxpatch: native extract: relationship %q in %q targets missing part %q", relationship.id, partName, resolved)
				}
			case "External":
				if _, parseErr := url.Parse(relationship.target); parseErr != nil {
					return fmt.Errorf("xlsxpatch: native extract: external relationship %q in %q has invalid target: %w", relationship.id, partName, parseErr)
				}
				if err := extractor.addUnsupported("EXTERNAL_RELATIONSHIP", "external-links", "workbook", partName, "", "external relationship is preserved but is outside native mutation scope"); err != nil {
					return err
				}
			default:
				return fmt.Errorf("xlsxpatch: native extract: relationship %q in %q has unsupported TargetMode %q", relationship.id, partName, relationship.targetMode)
			}
		}
	}
	return nil
}

func nativeRelationshipOwner(relsPart string) (owner, base string, err error) {
	reference, err := url.Parse(relsPart)
	if err != nil {
		return "", "", err
	}
	decoded := strings.TrimPrefix(reference.Path, "/")
	segments := strings.Split(decoded, "/")
	if len(segments) == 2 && asciiEqualFold(segments[0], "_rels") && asciiEqualFold(segments[1], ".rels") {
		return "", "", nil
	}
	if len(segments) < 2 || !asciiEqualFold(segments[len(segments)-2], "_rels") || !strings.HasSuffix(asciiLower(segments[len(segments)-1]), ".rels") {
		return "", "", fmt.Errorf("non-canonical relationship part location")
	}
	relsName := segments[len(segments)-1]
	ownerName := relsName[:len(relsName)-len(".rels")]
	ownerSegments := append([]string{}, segments[:len(segments)-2]...)
	ownerSegments = append(ownerSegments, ownerName)
	owner = strings.Join(ownerSegments, "/")
	base = strings.Join(ownerSegments[:len(ownerSegments)-1], "/")
	return owner, base, nil
}

func resolveNativeRelationshipTarget(owner, base, target string) (string, error) {
	reference, err := url.Parse(target)
	if err != nil {
		return "", fmt.Errorf("invalid relationship target: %w", err)
	}
	if reference.Path == "" && reference.Fragment != "" && owner != "" && reference.RawQuery == "" && !reference.ForceQuery {
		return owner, nil
	}
	return resolveRelPath(base, target)
}

func (extractor *nativeWorkbookExtractor) extractWorkbookRoutes(data []byte) ([]nativeWorkbookSheetRoute, error) {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	depth, sheetsDepth := 0, 0
	rootSeen, rootClosed, sheetsSeen := false, false, false
	routes := make([]nativeWorkbookSheetRoute, 0)
	seenIDs, seenNames := map[string]bool{}, map[string]bool{}
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: native extract: parse workbook: %w", err)
		}
		switch token := token.(type) {
		case xml.StartElement:
			depth++
			if depth == 1 {
				if token.Name != (xml.Name{Space: extractor.namespace, Local: "workbook"}) {
					return nil, fmt.Errorf("xlsxpatch: native extract: workbook root uses an unexpected namespace")
				}
				rootSeen = true
				if len(unexpectedSemanticXMLAttributes(token)) != 0 {
					if err := extractor.addUnsupported("WORKBOOK_ATTRIBUTES", "workbook-features", "workbook", extractor.workbook.part, "", "unmodeled workbook attributes remain authority-bound to the source package"); err != nil {
						return nil, err
					}
				}
				continue
			}
			if depth == 2 && token.Name.Space == extractor.namespace && token.Name.Local == "sheets" {
				if sheetsSeen {
					return nil, fmt.Errorf("xlsxpatch: native extract: workbook has multiple sheets elements")
				}
				sheetsSeen, sheetsDepth = true, depth
				if len(unexpectedSemanticXMLAttributes(token)) != 0 {
					if err := extractor.addUnsupported("SHEETS_ATTRIBUTES", "workbook-features", "workbook", extractor.workbook.part, "", "unmodeled sheets collection attributes remain authority-bound to the source package"); err != nil {
						return nil, err
					}
				}
				continue
			}
			if depth == 2 {
				if err := extractor.addUnsupported("UNMODELED_WORKBOOK_FEATURE", "workbook-features", "workbook", extractor.workbook.part, "", "workbook feature is preserved exactly outside the v1 sheet projection"); err != nil {
					return nil, err
				}
				if err := skipNativeXMLElement(decoder, token, 1); err != nil {
					return nil, fmt.Errorf("xlsxpatch: native extract: workbook feature %s: %w", token.Name.Local, err)
				}
				depth--
				continue
			}
			if sheetsDepth != 0 && depth == sheetsDepth+1 {
				if token.Name.Space != extractor.namespace || token.Name.Local != "sheet" {
					return nil, fmt.Errorf("xlsxpatch: native extract: sheets has unsupported direct child {%s}%s", token.Name.Space, token.Name.Local)
				}
				rawID, idFound, attrErr := unqualifiedXMLAttribute(token, "sheetId")
				if attrErr != nil || !idFound {
					return nil, fmt.Errorf("xlsxpatch: native extract: sheet requires a unique valid sheetId")
				}
				id, idErr := canonicalNativeSheetID(rawID)
				if idErr != nil {
					return nil, fmt.Errorf("xlsxpatch: native extract: sheet sheetId=%q: %w", rawID, idErr)
				}
				if seenIDs[id] {
					return nil, fmt.Errorf("xlsxpatch: native extract: stable sheet id %q is duplicated", id)
				}
				seenIDs[id] = true
				name, nameFound, attrErr := unqualifiedXMLAttribute(token, "name")
				if attrErr != nil || !nameFound || name == "" {
					return nil, fmt.Errorf("xlsxpatch: native extract: sheet %q requires a name", id)
				}
				name, err = decodeSpreadsheetString(name)
				if err != nil {
					return nil, fmt.Errorf("xlsxpatch: native extract: sheet %q name: %w", id, err)
				}
				if strings.HasPrefix(name, "'") || strings.HasSuffix(name, "'") {
					return nil, fmt.Errorf("xlsxpatch: native extract: sheet %q name cannot begin or end with an apostrophe", id)
				}
				nameKey := nativeSheetNameCaseKey(name)
				if seenNames[nameKey] {
					return nil, fmt.Errorf("xlsxpatch: native extract: sheet name %q is case-insensitively duplicated", name)
				}
				seenNames[nameKey] = true
				state, stateFound, attrErr := unqualifiedXMLAttribute(token, "state")
				if attrErr != nil {
					return nil, fmt.Errorf("xlsxpatch: native extract: sheet %q state: %w", id, attrErr)
				}
				if !stateFound {
					state = "visible"
				}
				if state != "visible" && state != "hidden" && state != "veryHidden" {
					return nil, fmt.Errorf("xlsxpatch: native extract: sheet %q has unsupported state %q", id, state)
				}
				relID, relFound, attrErr := namespacedXMLAttribute(token, extractor.relNamespace, "id")
				if attrErr != nil || !relFound || relID == "" {
					return nil, fmt.Errorf("xlsxpatch: native extract: sheet %q requires one relationship id in the workbook dialect", id)
				}
				refusalCode := ""
				if len(unexpectedSemanticXMLAttributes(token,
					xml.Name{Local: "sheetId"}, xml.Name{Local: "name"}, xml.Name{Local: "state"},
					xml.Name{Space: extractor.relNamespace, Local: "id"},
				)) != 0 {
					refusalCode = "SHEET_DECLARATION_ATTRIBUTES"
					if err := extractor.addUnsupported("SHEET_DECLARATION_ATTRIBUTES", "workbook-features", "sheet:"+id, extractor.workbook.part, "", "unmodeled sheet declaration attributes remain authority-bound to the source package and mutation is refused"); err != nil {
						return nil, err
					}
				}
				routes = append(routes, nativeWorkbookSheetRoute{id: id, name: name, state: state, relID: relID, refusalCode: refusalCode})
				continue
			}
			if sheetsDepth != 0 && depth > sheetsDepth+1 {
				return nil, fmt.Errorf("xlsxpatch: native extract: sheet declaration has unsupported nested markup")
			}
		case xml.EndElement:
			if sheetsDepth != 0 && depth == sheetsDepth && token.Name == (xml.Name{Space: extractor.namespace, Local: "sheets"}) {
				sheetsDepth = 0
			}
			if depth == 1 && token.Name == (xml.Name{Space: extractor.namespace, Local: "workbook"}) {
				rootClosed = true
			}
			depth--
		case xml.CharData:
			if depth > 0 && len(bytes.TrimSpace(token)) != 0 {
				return nil, fmt.Errorf("xlsxpatch: native extract: workbook contains unsupported direct text")
			}
		case xml.ProcInst:
			if depth == 0 && !rootSeen && token.Target == "xml" {
				continue
			}
			return nil, fmt.Errorf("xlsxpatch: native extract: workbook contains unsupported processing instruction %q", token.Target)
		case xml.Directive:
			return nil, fmt.Errorf("xlsxpatch: native extract: workbook contains unsupported XML directive")
		}
	}
	if depth != 0 || !rootSeen || !rootClosed || !sheetsSeen || len(routes) == 0 {
		return nil, fmt.Errorf("xlsxpatch: native extract: workbook requires a complete non-empty sheets collection")
	}
	if len(routes) > NativeXLSXMaxSheets {
		return nil, fmt.Errorf("xlsxpatch: native extract: workbook exceeds %d sheets", NativeXLSXMaxSheets)
	}
	relsData := extractor.pkg.files[extractor.workbook.relsPart]
	relationships, err := parseRoutingRelationships(relsData)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: native extract: workbook relationships: %w", err)
	}
	byID := make(map[string]routingRelationship, len(relationships))
	for _, relationship := range relationships {
		byID[relationship.id] = relationship
	}
	expectedType, opposingType := relTypeWorksheetTransitional, relTypeWorksheetStrict
	if extractor.workbook.strict {
		expectedType, opposingType = relTypeWorksheetStrict, relTypeWorksheetTransitional
	}
	for index := range routes {
		route := &routes[index]
		relationship, found := byID[route.relID]
		if !found {
			return nil, fmt.Errorf("xlsxpatch: native extract: sheet %q relationship %q is missing", route.id, route.relID)
		}
		if relationship.relType == opposingType {
			return nil, fmt.Errorf("xlsxpatch: native extract: sheet %q relationship uses the opposing Strict/Transitional dialect", route.id)
		}
		if relationship.relType != expectedType {
			return nil, fmt.Errorf("xlsxpatch: native extract: sheet %q relationship has non-worksheet type %q", route.id, relationship.relType)
		}
		if relationship.targetMode != "" && relationship.targetMode != "Internal" {
			return nil, fmt.Errorf("xlsxpatch: native extract: sheet %q relationship has unsupported TargetMode %q", route.id, relationship.targetMode)
		}
		resolved, err := resolveRelPath(extractor.workbook.baseDir, relationship.target)
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: native extract: sheet %q relationship target: %w", route.id, err)
		}
		part, _, found := extractor.pkg.index.lookupResolved(resolved)
		if !found {
			return nil, fmt.Errorf("xlsxpatch: native extract: sheet %q targets missing worksheet %q", route.id, resolved)
		}
		if err := requireNativeContentType(extractor.pkg, part, nativeWorksheetContentType); err != nil {
			return nil, fmt.Errorf("xlsxpatch: native extract: sheet %q: %w", route.id, err)
		}
		route.part = part
	}
	return routes, nil
}

func namespacedXMLAttribute(start xml.StartElement, namespace, local string) (string, bool, error) {
	value, found := "", false
	for _, attribute := range start.Attr {
		if attribute.Name.Space != namespace || attribute.Name.Local != local {
			continue
		}
		if found {
			return "", false, fmt.Errorf("duplicate {%s}%s attribute", namespace, local)
		}
		value, found = attribute.Value, true
	}
	return value, found, nil
}

func decodeSpreadsheetString(value string) (string, error) {
	if !utf8.ValidString(value) {
		return "", fmt.Errorf("string is not valid UTF-8")
	}
	var out strings.Builder
	for index := 0; index < len(value); {
		if index+7 <= len(value) && value[index] == '_' && (value[index+1] == 'x' || value[index+1] == 'X') && value[index+6] == '_' {
			unit, ok := parseSpreadsheetEscapeUnit(value[index+2 : index+6])
			if ok {
				index += 7
				if unit >= 0xD800 && unit <= 0xDBFF {
					if index+7 > len(value) || value[index] != '_' || (value[index+1] != 'x' && value[index+1] != 'X') || value[index+6] != '_' {
						return "", fmt.Errorf("unpaired high surrogate in ST_Xstring escape")
					}
					low, lowOK := parseSpreadsheetEscapeUnit(value[index+2 : index+6])
					if !lowOK || low < 0xDC00 || low > 0xDFFF {
						return "", fmt.Errorf("unpaired high surrogate in ST_Xstring escape")
					}
					out.WriteRune(utf16.DecodeRune(rune(unit), rune(low)))
					index += 7
					continue
				}
				if unit >= 0xDC00 && unit <= 0xDFFF {
					return "", fmt.Errorf("unpaired low surrogate in ST_Xstring escape")
				}
				out.WriteRune(rune(unit))
				continue
			}
		}
		r, size := utf8.DecodeRuneInString(value[index:])
		out.WriteRune(r)
		index += size
	}
	return out.String(), nil
}

func parseSpreadsheetEscapeUnit(value string) (uint16, bool) {
	if len(value) != 4 {
		return 0, false
	}
	parsed, err := strconv.ParseUint(value, 16, 16)
	return uint16(parsed), err == nil
}

func finiteNativeFloat(value string, minimum, maximum float64) (float64, error) {
	if !nativeXMLFiniteDoublePattern.MatchString(value) {
		return 0, fmt.Errorf("%q is not a canonical finite XML numeric lexical", value)
	}
	parsed, err := strconv.ParseFloat(value, 64)
	if err != nil || math.IsNaN(parsed) || math.IsInf(parsed, 0) || parsed < minimum || parsed > maximum {
		return 0, fmt.Errorf("%q is not finite within %g..%g", value, minimum, maximum)
	}
	return parsed, nil
}
