package pptxpatch

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"net/url"
	"path"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"unicode"
	"unicode/utf8"
)

const (
	NativePPTXMaxPackageBytes          = 512 << 20
	nativeExtractMaxParts              = 20_000
	nativeExtractMaxPartBytes          = 64 << 20
	nativeExtractMaxXMLBytes           = 16 << 20
	nativeExtractMaxTotalExpandedBytes = 512 << 20
	nativeExtractMaxCompressionRatio   = 200
	nativeExtractMaxXMLDepth           = 128
	nativeExtractMaxXMLNodes           = 250_000
	nativeExtractMaxRelationships      = 100_000
	nativeExtractMaxPassthrough        = 10_000
	nativeExtractMaxEmittedPassthrough = 100_000
	nativeExtractMaxTotalMediaBytes    = 512 << 20
)

const (
	nsContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types"
	nsPackageRels  = "http://schemas.openxmlformats.org/package/2006/relationships"

	nsPresentationTransitional = "http://schemas.openxmlformats.org/presentationml/2006/main"
	nsPresentationStrict       = "http://purl.oclc.org/ooxml/presentationml/main"
	nsDrawingTransitional      = "http://schemas.openxmlformats.org/drawingml/2006/main"
	nsDrawingStrict            = "http://purl.oclc.org/ooxml/drawingml/main"
	nsOfficeRelsTransitional   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	nsOfficeRelsStrict         = "http://purl.oclc.org/ooxml/officeDocument/relationships"
	nsChartTransitional        = "http://schemas.openxmlformats.org/drawingml/2006/chart"
	nsChartStrict              = "http://purl.oclc.org/ooxml/drawingml/chart"

	relOfficeDocumentTransitional = nsOfficeRelsTransitional + "/officeDocument"
	relOfficeDocumentStrict       = nsOfficeRelsStrict + "/officeDocument"
	relSlideTransitional          = nsOfficeRelsTransitional + "/slide"
	relSlideStrict                = nsOfficeRelsStrict + "/slide"
	relSlideLayoutTransitional    = nsOfficeRelsTransitional + "/slideLayout"
	relSlideLayoutStrict          = nsOfficeRelsStrict + "/slideLayout"
	relSlideMasterTransitional    = nsOfficeRelsTransitional + "/slideMaster"
	relSlideMasterStrict          = nsOfficeRelsStrict + "/slideMaster"
	relThemeTransitional          = nsOfficeRelsTransitional + "/theme"
	relThemeStrict                = nsOfficeRelsStrict + "/theme"
	relImageTransitional          = nsOfficeRelsTransitional + "/image"
	relImageStrict                = nsOfficeRelsStrict + "/image"
	relChartTransitional          = nsOfficeRelsTransitional + "/chart"
	relChartStrict                = nsOfficeRelsStrict + "/chart"

	contentTypePresentation = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
	contentTypeSlide        = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
	contentTypeSlideLayout  = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
	contentTypeSlideMaster  = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"
	contentTypeTheme        = "application/vnd.openxmlformats-officedocument.theme+xml"
	contentTypeChart        = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
)

var nativeXMLDeclarationPattern = regexp.MustCompile(`^version="1\.0"(?: encoding="UTF-8")?(?: standalone="(?:yes|no)")?$`)

// NativePassthroughTokenRequest is delivered only to a trusted server-side
// capability issuer. Payload is bounded source-package data and must never be
// used as the token itself or returned to an untrusted client. ByteLength is
// repeated explicitly so an issuer can bind asset reads without trusting a
// caller-computed payload length.
type NativePassthroughTokenRequest struct {
	SourceRevision    string
	OwnerPart         string
	ObjectID          string
	FingerprintSHA256 string
	ByteLength        int64
	Reason            string
	Payload           []byte
}

type NativePassthroughTokenFactory interface {
	IssueNativePassthroughToken(NativePassthroughTokenRequest) (string, error)
}

// NativePassthroughTokenTransaction is an all-or-nothing capability issuance
// boundary. IssueNativePassthroughToken stages a token without publishing it.
// A successful CommitNativePassthroughTokens publishes every staged token at
// once; a failed commit MUST publish none. RollbackNativePassthroughTokens is
// idempotent and MUST leave every staged token unavailable.
type NativePassthroughTokenTransaction interface {
	NativePassthroughTokenFactory
	CommitNativePassthroughTokens() error
	RollbackNativePassthroughTokens()
}

// NativePassthroughTransactionalTokenFactory starts a bounded atomic issuance
// transaction. Native extraction requires this stronger boundary whenever a
// parsed DrawingML group is present, because one opaque descendant must roll
// back capabilities staged for every otherwise-projectable sibling.
type NativePassthroughTransactionalTokenFactory interface {
	BeginNativePassthroughTokenTransaction() (NativePassthroughTokenTransaction, error)
}

type NativePassthroughTokenFactoryFunc func(NativePassthroughTokenRequest) (string, error)

func (f NativePassthroughTokenFactoryFunc) IssueNativePassthroughToken(request NativePassthroughTokenRequest) (string, error) {
	return f(request)
}

type NativePPTXExtractOptions struct {
	Previous     *NativePPTXDeck
	TokenFactory NativePassthroughTokenFactory
	// Opt-in read-only projection of spAutoFit in its saved source frame.
	// This does not implement content-dependent resizing or qualify Office fidelity.
	AllowSourceFrameAutoFitPreview bool
}

type nativeExtractPackage struct {
	parts        map[string][]byte
	aliases      map[string]string
	contentTypes nativeExtractContentTypes
}

type nativeExtractContentTypes struct {
	defaults  map[string]string
	overrides map[string]string
}

type nativeExtractRelationship struct {
	ID         string
	Type       string
	Target     string
	TargetMode string
	Part       string
	Namespace  string
}

type nativeExtractDialect struct {
	presentation   string
	drawing        string
	rels           string
	relSlide       string
	relSlideLayout string
	relSlideMaster string
	relTheme       string
	relImage       string
	relChart       string
	chart          string
	packageRels    string
}

type nativeXMLNode struct {
	Name     xml.Name
	Attrs    []xml.Attr
	Children []*nativeXMLNode
	Text     string
	RawStart int64
	RawEnd   int64
}

type nativeIdentityIndex struct {
	documentID      string
	assets          map[string]string
	slides          map[string]string
	elements        map[string]string
	elementsBySlide map[string]map[string]string
}

type nativeExtractor struct {
	pkg                     nativeExtractPackage
	options                 NativePPTXExtractOptions
	sourceRevision          string
	identities              nativeIdentityIndex
	passthroughUsed         int
	relationshipsUsed       int
	relationshipCache       map[string][]nativeExtractRelationship
	passthroughCache        map[string]NativePassthroughRef
	tokenOwners             map[string]string
	documentID              string
	passthroughRefsEmitted  int
	elementsEmitted         int
	outputNodesEmitted      int
	textCodeUnitsEmitted    int64
	tableCellsEmitted       int
	assets                  []NativeAsset
	assetByAlias            map[string]int
	assetIDOwners           map[string]string
	mediaBytesEmitted       int64
	assetBase64Emitted      int64
	mediaBytesInspected     int64
	picturePartInspections  map[string]nativePicturePartInspection
	passthroughCacheJournal []string
	tokenOwnerJournal       []string
	assetAliasJournal       []string
	assetIDOwnerJournal     []string
	tokenStager             *nativePassthroughTokenStager
	groupProjectionSeen     bool
	theme                   nativeResolvedTheme
	slideDependencies       nativeSlideDependencyGraph
}

type nativeStagedPassthroughToken struct {
	placeholder string
	request     NativePassthroughTokenRequest
}

// nativePassthroughTokenStager is extraction-local. It retains one bounded
// reference to package bytes and never calls an external issuer. Payload bytes
// are copied exactly once, at the eventual provider boundary.
type nativePassthroughTokenStager struct {
	requests     []nativeStagedPassthroughToken
	payloadBytes int64
	next         uint64
}

func (stager *nativePassthroughTokenStager) IssueNativePassthroughToken(request NativePassthroughTokenRequest) (string, error) {
	if len(stager.requests) >= nativeExtractMaxPassthrough {
		return "", fmt.Errorf("pptxpatch: native extract: staged passthrough request budget exceeded")
	}
	length := int64(len(request.Payload))
	if request.ByteLength != length || length < 0 || length > nativeExtractMaxPartBytes || length > int64(nativeExtractMaxTotalExpandedBytes)-stager.payloadBytes {
		return "", fmt.Errorf("pptxpatch: native extract: staged passthrough payload budget exceeded")
	}
	stager.next++
	placeholder := fmt.Sprintf("staged-capability-%d", stager.next)
	stager.requests = append(stager.requests, nativeStagedPassthroughToken{placeholder: placeholder, request: request})
	stager.payloadBytes += length
	return placeholder, nil
}

func (stager *nativePassthroughTokenStager) rollback(requestCount int, payloadBytes int64) {
	if requestCount < 0 || requestCount > len(stager.requests) || payloadBytes < 0 || payloadBytes > stager.payloadBytes {
		panic("pptxpatch: invalid native passthrough staging checkpoint")
	}
	clear(stager.requests[requestCount:])
	stager.requests = stager.requests[:requestCount]
	stager.payloadBytes = payloadBytes
}

type nativeUnsupportedSource struct {
	part        string
	objectID    string
	fingerprint string
	payload     []byte
	code        string
	message     string
}

type nativeDuplicateSingletonError struct {
	space string
	local string
}

func (err nativeDuplicateSingletonError) Error() string {
	return fmt.Sprintf("pptxpatch: native extract: duplicate {%s}%s", err.space, err.local)
}

// ExtractNativePPTX parses a bounded OPC package through exact relationships
// and namespaces into the validated native v1 contract. It is intentionally
// independent from the permissive reconstructive ParsePPTX reader.
func ExtractNativePPTX(data []byte, options NativePPTXExtractOptions) (NativePPTXDeck, error) {
	if len(data) == 0 || len(data) > NativePPTXMaxPackageBytes {
		return NativePPTXDeck{}, fmt.Errorf("pptxpatch: native extract: package size must be 1..%d bytes", NativePPTXMaxPackageBytes)
	}
	pkg, err := openNativeExtractPackage(data)
	if err != nil {
		return NativePPTXDeck{}, err
	}
	packageDigest := sha256.Sum256(data)
	revision := "rev-" + hex.EncodeToString(packageDigest[:])
	externalTokenFactory := options.TokenFactory
	tokenStager := &nativePassthroughTokenStager{requests: []nativeStagedPassthroughToken{}}
	options.TokenFactory = tokenStager
	extractor := nativeExtractor{
		pkg: pkg, options: options, sourceRevision: revision,
		relationshipCache:      map[string][]nativeExtractRelationship{},
		passthroughCache:       map[string]NativePassthroughRef{},
		tokenOwners:            map[string]string{},
		assets:                 []NativeAsset{},
		assetByAlias:           map[string]int{},
		assetIDOwners:          map[string]string{},
		picturePartInspections: map[string]nativePicturePartInspection{},
		tokenStager:            tokenStager,
	}
	extractor.identities, err = previousNativeIdentities(options.Previous)
	if err != nil {
		return NativePPTXDeck{}, err
	}
	deck, err := extractor.extract()
	if err != nil {
		return NativePPTXDeck{}, err
	}
	if issues := ValidateNativePPTX(deck); len(issues) != 0 {
		return NativePPTXDeck{}, NativeContractValidationError{Issues: issues}
	}
	if _, err := MarshalNativePPTXJSON(deck); err != nil {
		return NativePPTXDeck{}, fmt.Errorf("pptxpatch: native extract canonical validation: %w", err)
	}
	return extractor.publishNativePassthroughTokens(deck, externalTokenFactory)
}

func (extractor *nativeExtractor) publishNativePassthroughTokens(deck NativePPTXDeck, factory NativePassthroughTokenFactory) (NativePPTXDeck, error) {
	if extractor.tokenStager == nil || len(extractor.tokenStager.requests) == 0 {
		return deck, nil
	}
	if factory == nil {
		return NativePPTXDeck{}, fmt.Errorf("pptxpatch: native extract: unsupported content requires a trusted passthrough token factory")
	}
	transactional, supportsTransaction := factory.(NativePassthroughTransactionalTokenFactory)
	if extractor.groupProjectionSeen && !supportsTransaction {
		return NativePPTXDeck{}, fmt.Errorf("pptxpatch: native extract: grouped content requires an atomic passthrough token transaction factory")
	}
	if supportsTransaction {
		return extractor.publishNativePassthroughTransaction(deck, transactional)
	}
	return extractor.publishNativePassthroughLegacy(deck, factory)
}

func (extractor *nativeExtractor) publishNativePassthroughTransaction(deck NativePPTXDeck, factory NativePassthroughTransactionalTokenFactory) (result NativePPTXDeck, resultErr error) {
	transaction, err := factory.BeginNativePassthroughTokenTransaction()
	if err != nil || transaction == nil {
		return NativePPTXDeck{}, fmt.Errorf("pptxpatch: native extract: begin passthrough token transaction failed")
	}
	committed := false
	defer func() {
		if !committed {
			transaction.RollbackNativePassthroughTokens()
		}
	}()
	replacements, err := extractor.issueNativeStagedTokens(transaction)
	if err != nil {
		return NativePPTXDeck{}, err
	}
	if err := replaceNativeStagedTokens(&deck, replacements); err != nil {
		return NativePPTXDeck{}, err
	}
	if issues := ValidateNativePPTX(deck); len(issues) != 0 {
		return NativePPTXDeck{}, NativeContractValidationError{Issues: issues}
	}
	if _, err := MarshalNativePPTXJSON(deck); err != nil {
		return NativePPTXDeck{}, fmt.Errorf("pptxpatch: native extract canonical validation after capability issuance: %w", err)
	}
	if err := transaction.CommitNativePassthroughTokens(); err != nil {
		return NativePPTXDeck{}, fmt.Errorf("pptxpatch: native extract: atomic passthrough token commit failed")
	}
	committed = true
	return deck, nil
}

func (extractor *nativeExtractor) publishNativePassthroughLegacy(deck NativePPTXDeck, factory NativePassthroughTokenFactory) (NativePPTXDeck, error) {
	replacements, err := extractor.issueNativeStagedTokens(factory)
	if err != nil {
		return NativePPTXDeck{}, err
	}
	if err := replaceNativeStagedTokens(&deck, replacements); err != nil {
		return NativePPTXDeck{}, err
	}
	if issues := ValidateNativePPTX(deck); len(issues) != 0 {
		return NativePPTXDeck{}, NativeContractValidationError{Issues: issues}
	}
	if _, err := MarshalNativePPTXJSON(deck); err != nil {
		return NativePPTXDeck{}, fmt.Errorf("pptxpatch: native extract canonical validation after capability issuance: %w", err)
	}
	return deck, nil
}

func (extractor *nativeExtractor) issueNativeStagedTokens(factory NativePassthroughTokenFactory) (map[string]string, error) {
	replacements := make(map[string]string, len(extractor.tokenStager.requests))
	owners := make(map[string]string, len(extractor.tokenStager.requests))
	for _, staged := range extractor.tokenStager.requests {
		request := staged.request
		request.Payload = append([]byte(nil), request.Payload...)
		token, err := factory.IssueNativePassthroughToken(request)
		if err != nil || !nativeIDPattern.MatchString(token) {
			return nil, fmt.Errorf("pptxpatch: native extract: token factory returned invalid capability")
		}
		owner := request.OwnerPart + "\x00" + request.ObjectID + "\x00" + request.FingerprintSHA256 + "\x00" + request.Reason
		if previousOwner, exists := owners[token]; exists && previousOwner != owner {
			return nil, fmt.Errorf("pptxpatch: native extract: token factory reused one capability for distinct source objects")
		}
		owners[token] = owner
		replacements[staged.placeholder] = token
	}
	return replacements, nil
}

func replaceNativeStagedTokens(deck *NativePPTXDeck, replacements map[string]string) error {
	replaceRefs := func(refs []NativePassthroughRef) error {
		for index := range refs {
			replacement, ok := replacements[refs[index].Token]
			if !ok {
				return fmt.Errorf("pptxpatch: native extract: staged capability has no issued replacement")
			}
			refs[index].Token = replacement
		}
		return nil
	}
	var replaceElement func(*NativeElement) error
	replaceElement = func(element *NativeElement) error {
		if err := replaceRefs(element.Passthrough); err != nil {
			return err
		}
		if element.Chart != nil {
			replacement, ok := replacements[element.Chart.OpaqueRef.Token]
			if !ok {
				return fmt.Errorf("pptxpatch: native extract: staged chart capability has no issued replacement")
			}
			element.Chart.OpaqueRef.Token = replacement
		}
		for index := range element.Children {
			if err := replaceElement(&element.Children[index]); err != nil {
				return err
			}
		}
		return nil
	}
	for index := range deck.Assets {
		if err := replaceRefs(deck.Assets[index].Passthrough); err != nil {
			return err
		}
	}
	for slideIndex := range deck.Slides {
		if err := replaceRefs(deck.Slides[slideIndex].Passthrough); err != nil {
			return err
		}
		for elementIndex := range deck.Slides[slideIndex].Elements {
			if err := replaceElement(&deck.Slides[slideIndex].Elements[elementIndex]); err != nil {
				return err
			}
		}
	}
	return nil
}

func openNativeExtractPackage(data []byte) (nativeExtractPackage, error) {
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nativeExtractPackage{}, fmt.Errorf("pptxpatch: native extract OPC: %w", err)
	}
	if len(zr.File) == 0 || len(zr.File) > nativeExtractMaxParts {
		return nativeExtractPackage{}, fmt.Errorf("pptxpatch: native extract OPC: part count must be 1..%d", nativeExtractMaxParts)
	}
	parts := make(map[string][]byte, len(zr.File))
	aliases := map[string]string{}
	var total uint64
	for _, file := range zr.File {
		name := file.Name
		if file.FileInfo().IsDir() {
			trimmed := strings.TrimSuffix(name, "/")
			if trimmed == "" || !secureNativePartName(trimmed) {
				return nativeExtractPackage{}, fmt.Errorf("pptxpatch: native extract OPC: unsafe directory %q", name)
			}
			continue
		}
		if !secureNativePartName(name) {
			return nativeExtractPackage{}, fmt.Errorf("pptxpatch: native extract OPC: unsafe part name %q", name)
		}
		if _, exists := parts[name]; exists {
			return nativeExtractPackage{}, fmt.Errorf("pptxpatch: native extract OPC: duplicate part %q", name)
		}
		alias, err := nativePartAlias(name)
		if err != nil {
			return nativeExtractPackage{}, fmt.Errorf("pptxpatch: native extract OPC: part alias %q: %w", name, err)
		}
		if previous, exists := aliases[alias]; exists {
			return nativeExtractPackage{}, fmt.Errorf("pptxpatch: native extract OPC: case/percent alias collision %q and %q", previous, name)
		}
		aliases[alias] = name
		if file.Flags&1 != 0 || (file.Method != zip.Store && file.Method != zip.Deflate) {
			return nativeExtractPackage{}, fmt.Errorf("pptxpatch: native extract OPC: encrypted or unsupported compression for %q", name)
		}
		if file.UncompressedSize64 > nativeExtractMaxPartBytes {
			return nativeExtractPackage{}, fmt.Errorf("pptxpatch: native extract OPC: part %q exceeds %d bytes", name, nativeExtractMaxPartBytes)
		}
		if file.UncompressedSize64 > 0 && (file.CompressedSize64 == 0 || file.UncompressedSize64 > file.CompressedSize64*nativeExtractMaxCompressionRatio) {
			return nativeExtractPackage{}, fmt.Errorf("pptxpatch: native extract OPC: suspicious compression ratio for %q", name)
		}
		total += file.UncompressedSize64
		if total > nativeExtractMaxTotalExpandedBytes {
			return nativeExtractPackage{}, fmt.Errorf("pptxpatch: native extract OPC: expanded package exceeds %d bytes", nativeExtractMaxTotalExpandedBytes)
		}
		rc, err := file.Open()
		if err != nil {
			return nativeExtractPackage{}, fmt.Errorf("pptxpatch: native extract OPC: open %q: %w", name, err)
		}
		limited := io.LimitReader(rc, int64(nativeExtractMaxPartBytes)+1)
		payload, readErr := io.ReadAll(limited)
		closeErr := rc.Close()
		if readErr != nil || closeErr != nil || len(payload) > nativeExtractMaxPartBytes || uint64(len(payload)) != file.UncompressedSize64 {
			return nativeExtractPackage{}, fmt.Errorf("pptxpatch: native extract OPC: bounded read failed for %q", name)
		}
		parts[name] = payload
	}
	contentTypesAlias, _ := nativePartAlias("[Content_Types].xml")
	contentTypesPart, ok := aliases[contentTypesAlias]
	if !ok {
		return nativeExtractPackage{}, fmt.Errorf("pptxpatch: native extract OPC: missing [Content_Types].xml")
	}
	contentTypes, err := parseNativeContentTypes(parts[contentTypesPart], aliases)
	if err != nil {
		return nativeExtractPackage{}, err
	}
	for actualPart := range parts {
		if actualPart == contentTypesPart {
			continue
		}
		effective := contentTypes.forPart(actualPart)
		if effective == "" {
			return nativeExtractPackage{}, fmt.Errorf("pptxpatch: native extract OPC: part %q has no effective content type", actualPart)
		}
		if asciiEqualFoldNative(path.Ext(actualPart), ".rels") && !asciiEqualFoldNative(effective, "application/vnd.openxmlformats-package.relationships+xml") {
			return nativeExtractPackage{}, fmt.Errorf("pptxpatch: native extract OPC: relationships part %q has invalid effective content type", actualPart)
		}
	}
	return nativeExtractPackage{parts: parts, aliases: aliases, contentTypes: contentTypes}, nil
}

func secureNativePartName(name string) bool {
	if name == "" || len(name) > 1024 || strings.HasPrefix(name, "/") || strings.HasSuffix(name, "/") || strings.Contains(name, "\\") || !utf8.ValidString(name) {
		return false
	}
	for _, segment := range strings.Split(name, "/") {
		if segment == "" {
			return false
		}
		for index := 0; index < len(segment); index++ {
			if segment[index] == '?' || segment[index] == '#' || segment[index] <= 0x20 || segment[index] == 0x7f {
				return false
			}
			if segment[index] != '%' {
				continue
			}
			if index+2 >= len(segment) || !isUpperHex(segment[index+1]) || !isUpperHex(segment[index+2]) {
				return false
			}
			index += 2
		}
		decoded, err := url.PathUnescape(segment)
		if err != nil || !utf8.ValidString(decoded) || decoded == "" || decoded == "." || decoded == ".." || strings.HasSuffix(decoded, ".") || strings.ContainsAny(decoded, "/\\") {
			return false
		}
		for _, runeValue := range decoded {
			if runeValue < 0x20 || runeValue == 0x7f || (!strings.Contains(segment, "%") && unicode.IsSpace(runeValue)) {
				return false
			}
		}
	}
	return true
}

func nativePartAlias(name string) (string, error) {
	segments := strings.Split(name, "/")
	for index, segment := range segments {
		decoded, err := url.PathUnescape(segment)
		if err != nil {
			return "", err
		}
		segments[index] = asciiLowerNative(decoded)
	}
	return strings.Join(segments, "/"), nil
}

func asciiLowerNative(value string) string {
	buffer := []byte(value)
	for index, current := range buffer {
		if current >= 'A' && current <= 'Z' {
			buffer[index] = current + ('a' - 'A')
		}
	}
	return string(buffer)
}

func asciiEqualFoldNative(left, right string) bool {
	return asciiLowerNative(left) == asciiLowerNative(right)
}

func (types nativeExtractContentTypes) forPart(part string) string {
	alias, err := nativePartAlias(part)
	if err != nil {
		return ""
	}
	if value := types.overrides[alias]; value != "" {
		return value
	}
	extension := ""
	if dot := strings.LastIndexByte(part, '.'); dot >= 0 {
		extension = asciiLowerNative(part[dot+1:])
	}
	return types.defaults[extension]
}

func parseNativeContentTypes(data []byte, packageAliases map[string]string) (nativeExtractContentTypes, error) {
	root, err := parseNativeXML(data, "[Content_Types].xml")
	if err != nil {
		return nativeExtractContentTypes{}, err
	}
	if root.Name != (xml.Name{Space: nsContentTypes, Local: "Types"}) {
		return nativeExtractContentTypes{}, fmt.Errorf("pptxpatch: native extract OPC: invalid content-types root")
	}
	if err := requireOnlyNativeAttrs(root); err != nil {
		return nativeExtractContentTypes{}, fmt.Errorf("pptxpatch: native extract OPC: content-types root: %w", err)
	}
	result := nativeExtractContentTypes{defaults: map[string]string{}, overrides: map[string]string{}}
	overrideAliases := map[string]string{}
	for _, child := range root.Children {
		switch child.Name {
		case xml.Name{Space: nsContentTypes, Local: "Default"}:
			if err := requireOnlyNativeAttrs(child, xml.Name{Local: "Extension"}, xml.Name{Local: "ContentType"}); err != nil {
				return nativeExtractContentTypes{}, fmt.Errorf("pptxpatch: native extract OPC: content-type default: %w", err)
			}
			extension, ok := exactNativeAttr(child, "", "Extension")
			contentType, ctOK := exactNativeAttr(child, "", "ContentType")
			extension = asciiLowerNative(extension)
			if err := requireOnlyNativeChildren(child); err != nil {
				return nativeExtractContentTypes{}, fmt.Errorf("pptxpatch: native extract OPC: content-type default: %w", err)
			}
			if !ok || !ctOK || extension == "" || strings.Contains(extension, ".") || strings.TrimSpace(contentType) == "" || contentType != strings.TrimSpace(contentType) || result.defaults[extension] != "" {
				return nativeExtractContentTypes{}, fmt.Errorf("pptxpatch: native extract OPC: invalid or duplicate content-type default")
			}
			result.defaults[extension] = contentType
		case xml.Name{Space: nsContentTypes, Local: "Override"}:
			if err := requireOnlyNativeAttrs(child, xml.Name{Local: "PartName"}, xml.Name{Local: "ContentType"}); err != nil {
				return nativeExtractContentTypes{}, fmt.Errorf("pptxpatch: native extract OPC: content-type override: %w", err)
			}
			part, ok := exactNativeAttr(child, "", "PartName")
			contentType, ctOK := exactNativeAttr(child, "", "ContentType")
			if err := requireOnlyNativeChildren(child); err != nil {
				return nativeExtractContentTypes{}, fmt.Errorf("pptxpatch: native extract OPC: content-type override: %w", err)
			}
			if !ok || !ctOK || !strings.HasPrefix(part, "/") || strings.TrimSpace(contentType) == "" || contentType != strings.TrimSpace(contentType) {
				return nativeExtractContentTypes{}, fmt.Errorf("pptxpatch: native extract OPC: invalid content-type override")
			}
			part = strings.TrimPrefix(part, "/")
			alias, aliasErr := nativePartAlias(part)
			if !secureNativePartName(part) || aliasErr != nil || result.overrides[alias] != "" {
				return nativeExtractContentTypes{}, fmt.Errorf("pptxpatch: native extract OPC: unsafe or duplicate content-type override %q", part)
			}
			if previous, exists := overrideAliases[alias]; exists {
				return nativeExtractContentTypes{}, fmt.Errorf("pptxpatch: native extract OPC: content-type override aliases %q and %q", previous, part)
			}
			if _, exists := packageAliases[alias]; !exists {
				return nativeExtractContentTypes{}, fmt.Errorf("pptxpatch: native extract OPC: content-type override targets missing part %q", part)
			}
			overrideAliases[alias] = part
			result.overrides[alias] = contentType
		default:
			return nativeExtractContentTypes{}, fmt.Errorf("pptxpatch: native extract OPC: unknown content-types element %s", child.Name.Local)
		}
	}
	for extension := range result.defaults {
		used := false
		for _, actualPart := range packageAliases {
			if asciiEqualFoldNative(strings.TrimPrefix(path.Ext(actualPart), "."), extension) {
				used = true
				break
			}
		}
		if !used {
			return nativeExtractContentTypes{}, fmt.Errorf("pptxpatch: native extract OPC: dangling content-type default for %q", extension)
		}
	}
	return result, nil
}

func parseNativeXML(data []byte, part string) (*nativeXMLNode, error) {
	if len(data) == 0 || len(data) > nativeExtractMaxXMLBytes {
		return nil, fmt.Errorf("pptxpatch: native extract XML: %q size must be 1..%d", part, nativeExtractMaxXMLBytes)
	}
	decoder := xml.NewDecoder(bytes.NewReader(data))
	decoder.Strict = true
	var stack []*nativeXMLNode
	var root *nativeXMLNode
	seenXMLDeclaration := false
	seenPreRootContent := false
	nodes := 0
	for {
		tokenStart := decoder.InputOffset()
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("pptxpatch: native extract XML %q: %w", part, err)
		}
		switch typed := token.(type) {
		case xml.Directive:
			return nil, fmt.Errorf("pptxpatch: native extract XML %q: directives are forbidden", part)
		case xml.ProcInst:
			declaration := string(typed.Inst)
			if typed.Target != "xml" || seenXMLDeclaration || seenPreRootContent || root != nil || len(stack) != 0 || !nativeXMLDeclarationPattern.MatchString(declaration) {
				return nil, fmt.Errorf("pptxpatch: native extract XML %q: non-canonical declaration or processing instruction", part)
			}
			seenXMLDeclaration = true
		case xml.StartElement:
			nodes++
			if nodes > nativeExtractMaxXMLNodes || len(stack)+1 > nativeExtractMaxXMLDepth {
				return nil, fmt.Errorf("pptxpatch: native extract XML %q: resource budget exceeded", part)
			}
			if duplicateNativeAttrs(typed.Attr) {
				return nil, fmt.Errorf("pptxpatch: native extract XML %q: duplicate attribute", part)
			}
			node := &nativeXMLNode{Name: typed.Name, Attrs: append([]xml.Attr(nil), typed.Attr...), Children: []*nativeXMLNode{}, RawStart: tokenStart}
			if len(stack) == 0 {
				if root != nil {
					return nil, fmt.Errorf("pptxpatch: native extract XML %q: multiple roots", part)
				}
				root = node
				seenPreRootContent = true
			} else {
				stack[len(stack)-1].Children = append(stack[len(stack)-1].Children, node)
			}
			stack = append(stack, node)
		case xml.CharData:
			if len(stack) == 0 {
				if !onlyNativeXMLSpace(string(typed)) {
					return nil, fmt.Errorf("pptxpatch: native extract XML %q: non-whitespace text outside root", part)
				}
				if root == nil && len(typed) != 0 {
					seenPreRootContent = true
				}
			} else {
				stack[len(stack)-1].Text += string(typed)
				if len(stack[len(stack)-1].Text) > nativeMaxTextCodeUnits*4 {
					return nil, fmt.Errorf("pptxpatch: native extract XML %q: text budget exceeded", part)
				}
			}
		case xml.EndElement:
			if len(stack) == 0 || stack[len(stack)-1].Name != typed.Name {
				return nil, fmt.Errorf("pptxpatch: native extract XML %q: mismatched close", part)
			}
			stack[len(stack)-1].RawEnd = decoder.InputOffset()
			stack = stack[:len(stack)-1]
		case xml.Comment:
			if root == nil {
				seenPreRootContent = true
			}
		}
	}
	if root == nil || len(stack) != 0 {
		return nil, fmt.Errorf("pptxpatch: native extract XML %q: missing or incomplete root", part)
	}
	return root, nil
}

func duplicateNativeAttrs(attrs []xml.Attr) bool {
	seen := map[xml.Name]bool{}
	for _, attr := range attrs {
		if attr.Name.Space == "xmlns" || attr.Name.Local == "xmlns" {
			continue
		}
		if seen[attr.Name] {
			return true
		}
		seen[attr.Name] = true
	}
	return false
}

// XML S is deliberately narrower than Unicode whitespace. Native extraction
// must not silently discard NBSP or other visible/non-XML spacing characters
// when attesting that element text is empty.
func onlyNativeXMLSpace(value string) bool {
	for index := 0; index < len(value); index++ {
		switch value[index] {
		case 0x20, 0x09, 0x0a, 0x0d:
		default:
			return false
		}
	}
	return true
}

func exactNativeAttr(node *nativeXMLNode, space, local string) (string, bool) {
	if node == nil {
		return "", false
	}
	for _, attr := range node.Attrs {
		if attr.Name.Space == space && attr.Name.Local == local {
			return attr.Value, true
		}
	}
	return "", false
}

func requireOnlyNativeAttrs(node *nativeXMLNode, allowed ...xml.Name) error {
	if node == nil {
		return fmt.Errorf("missing element")
	}
	permitted := make(map[xml.Name]bool, len(allowed))
	for _, name := range allowed {
		permitted[name] = true
	}
	for _, attr := range node.Attrs {
		if attr.Name.Space == "xmlns" || attr.Name.Local == "xmlns" {
			continue
		}
		if !permitted[attr.Name] {
			return fmt.Errorf("unsupported attribute {%s}%s", attr.Name.Space, attr.Name.Local)
		}
	}
	return nil
}

func hasNativeSemanticAttrs(node *nativeXMLNode) bool {
	if node == nil {
		return false
	}
	for _, attr := range node.Attrs {
		if attr.Name.Space != "xmlns" && attr.Name.Local != "xmlns" {
			return true
		}
	}
	return false
}

func requireOnlyNativeChildren(node *nativeXMLNode, allowed ...xml.Name) error {
	if node == nil {
		return fmt.Errorf("missing element")
	}
	permitted := make(map[xml.Name]bool, len(allowed))
	for _, name := range allowed {
		permitted[name] = true
	}
	for _, child := range node.Children {
		if !permitted[child.Name] {
			return fmt.Errorf("unsupported child {%s}%s", child.Name.Space, child.Name.Local)
		}
	}
	if !onlyNativeXMLSpace(node.Text) {
		return fmt.Errorf("unsupported direct text")
	}
	return nil
}

func requireEmptyNativeElement(node *nativeXMLNode) error {
	if err := requireOnlyNativeAttrs(node); err != nil {
		return err
	}
	return requireOnlyNativeChildren(node)
}

func nativeChild(node *nativeXMLNode, space, local string) *nativeXMLNode {
	if node == nil {
		return nil
	}
	for _, child := range node.Children {
		if child.Name.Space == space && child.Name.Local == local {
			return child
		}
	}
	return nil
}

func nativeChildren(node *nativeXMLNode, space, local string) []*nativeXMLNode {
	if node == nil {
		return nil
	}
	var children []*nativeXMLNode
	for _, child := range node.Children {
		if child.Name.Space == space && child.Name.Local == local {
			children = append(children, child)
		}
	}
	return children
}

func nativeSingleton(node *nativeXMLNode, space, local string, required bool) (*nativeXMLNode, error) {
	children := nativeChildren(node, space, local)
	if len(children) > 1 {
		return nil, nativeDuplicateSingletonError{space: space, local: local}
	}
	if len(children) == 0 {
		if required {
			return nil, fmt.Errorf("pptxpatch: native extract: missing {%s}%s", space, local)
		}
		return nil, nil
	}
	return children[0], nil
}

func rawNativeNode(data []byte, node *nativeXMLNode) ([]byte, error) {
	if node == nil || node.RawStart < 0 || node.RawEnd <= node.RawStart || node.RawEnd > int64(len(data)) {
		return nil, fmt.Errorf("pptxpatch: native extract: invalid bounded XML source span")
	}
	return data[node.RawStart:node.RawEnd], nil
}

func makeNativeUnsupportedSource(data []byte, node *nativeXMLNode, part, objectID, code, message string) (nativeUnsupportedSource, error) {
	raw, err := rawNativeNode(data, node)
	if err != nil {
		return nativeUnsupportedSource{}, err
	}
	return nativeUnsupportedSource{
		part: part, objectID: objectID, fingerprint: nativeSHA256(raw), payload: raw,
		code: code, message: message,
	}, nil
}

func makeNativeUnsupportedPayload(payload []byte, part, objectID, code, message string) nativeUnsupportedSource {
	return nativeUnsupportedSource{
		part: part, objectID: objectID, fingerprint: nativeSHA256(payload), payload: payload,
		code: code, message: message,
	}
}

func (extractor *nativeExtractor) parseRelationships(sourcePart string) ([]nativeExtractRelationship, error) {
	if cached, ok := extractor.relationshipCache[sourcePart]; ok {
		return cached, nil
	}
	relsRequest := nativeRelationshipsPart(sourcePart)
	relsAlias, aliasErr := nativePartAlias(relsRequest)
	relsPart, ok := extractor.pkg.aliases[relsAlias]
	if !ok {
		return nil, fmt.Errorf("pptxpatch: native extract OPC: missing relationships %q", relsRequest)
	}
	if aliasErr != nil || !asciiEqualFoldNative(extractor.pkg.contentTypes.forPart(relsPart), "application/vnd.openxmlformats-package.relationships+xml") {
		return nil, fmt.Errorf("pptxpatch: native extract OPC: relationships %q has invalid effective content type", relsPart)
	}
	data := extractor.pkg.parts[relsPart]
	root, err := parseNativeXML(data, relsPart)
	if err != nil {
		return nil, err
	}
	if root.Name != (xml.Name{Space: nsPackageRels, Local: "Relationships"}) {
		return nil, fmt.Errorf("pptxpatch: native extract OPC: invalid relationships root %q", relsPart)
	}
	if len(root.Children) > nativeExtractMaxRelationships || extractor.relationshipsUsed > nativeExtractMaxRelationships-len(root.Children) {
		return nil, fmt.Errorf("pptxpatch: native extract OPC: relationship budget exceeded")
	}
	extractor.relationshipsUsed += len(root.Children)
	seen := map[string]bool{}
	result := make([]nativeExtractRelationship, 0, len(root.Children))
	for _, child := range root.Children {
		if child.Name != (xml.Name{Space: root.Name.Space, Local: "Relationship"}) {
			return nil, fmt.Errorf("pptxpatch: native extract OPC: unknown relationships element")
		}
		if err := requireOnlyNativeAttrs(child, xml.Name{Local: "Id"}, xml.Name{Local: "Type"}, xml.Name{Local: "Target"}, xml.Name{Local: "TargetMode"}); err != nil {
			return nil, fmt.Errorf("pptxpatch: native extract OPC: relationship in %q: %w", relsPart, err)
		}
		id, idOK := exactNativeAttr(child, "", "Id")
		relType, typeOK := exactNativeAttr(child, "", "Type")
		target, targetOK := exactNativeAttr(child, "", "Target")
		targetMode, _ := exactNativeAttr(child, "", "TargetMode")
		if !idOK || !typeOK || !targetOK || id == "" || relType == "" || target == "" || seen[id] {
			return nil, fmt.Errorf("pptxpatch: native extract OPC: invalid or duplicate relationship in %q", relsPart)
		}
		seen[id] = true
		if targetMode != "" && targetMode != "Internal" && targetMode != "External" {
			return nil, fmt.Errorf("pptxpatch: native extract OPC: invalid TargetMode %q in %q", targetMode, relsPart)
		}
		part := ""
		if targetMode != "External" {
			part, err = resolveNativeRelationshipTarget(sourcePart, target)
			if err != nil {
				return nil, fmt.Errorf("pptxpatch: native extract OPC: relationship %s in %q: %w", id, relsPart, err)
			}
			partAlias, aliasErr := nativePartAlias(part)
			actualPart, exists := extractor.pkg.aliases[partAlias]
			if aliasErr != nil || !exists {
				return nil, fmt.Errorf("pptxpatch: native extract OPC: relationship %s targets missing OPC part %q", id, part)
			}
			part = actualPart
		}
		result = append(result, nativeExtractRelationship{ID: id, Type: relType, Target: target, TargetMode: targetMode, Part: part, Namespace: root.Name.Space})
	}
	extractor.relationshipCache[sourcePart] = result
	return result, nil
}

func (relationship nativeExtractRelationship) internal() bool {
	return relationship.TargetMode == "" || relationship.TargetMode == "Internal"
}

func nativeRelationshipsPart(sourcePart string) string {
	if sourcePart == "" {
		return "_rels/.rels"
	}
	directory, base := path.Split(sourcePart)
	return directory + "_rels/" + base + ".rels"
}

func (extractor *nativeExtractor) actualRelationshipsPart(sourcePart string) (string, error) {
	requested := nativeRelationshipsPart(sourcePart)
	alias, err := nativePartAlias(requested)
	if err != nil {
		return "", err
	}
	actual, ok := extractor.pkg.aliases[alias]
	if !ok {
		return "", fmt.Errorf("pptxpatch: native extract OPC: missing relationships %q", requested)
	}
	return actual, nil
}

func (extractor *nativeExtractor) optionalRelationships(sourcePart string) ([]nativeExtractRelationship, bool, error) {
	requested := nativeRelationshipsPart(sourcePart)
	alias, err := nativePartAlias(requested)
	if err != nil {
		return nil, false, err
	}
	if _, ok := extractor.pkg.aliases[alias]; !ok {
		return nil, false, nil
	}
	relationships, err := extractor.parseRelationships(sourcePart)
	return relationships, true, err
}

func resolveNativeRelationshipTarget(sourcePart, target string) (string, error) {
	parsed, err := url.Parse(target)
	if err != nil || parsed.Scheme != "" || parsed.Host != "" || parsed.RawQuery != "" || parsed.Fragment != "" || strings.Contains(target, "\\") {
		return "", fmt.Errorf("unsafe relationship target %q", target)
	}
	rawPath := parsed.EscapedPath()
	if rawPath == "" || strings.Contains(rawPath, "//") || strings.HasSuffix(rawPath, "/") {
		return "", fmt.Errorf("empty relationship target")
	}
	absolute := strings.HasPrefix(rawPath, "/")
	segments := []string{}
	if !absolute {
		directory, _ := path.Split(sourcePart)
		for _, segment := range strings.Split(strings.TrimSuffix(directory, "/"), "/") {
			if segment != "" {
				segments = append(segments, segment)
			}
		}
	}
	for _, segment := range strings.Split(strings.TrimPrefix(rawPath, "/"), "/") {
		if segment == "" || segment == "." {
			return "", fmt.Errorf("non-canonical relationship target")
		}
		if segment == ".." {
			if len(segments) == 0 {
				return "", fmt.Errorf("relationship target escapes package root")
			}
			segments = segments[:len(segments)-1]
			continue
		}
		decoded, decodeErr := url.PathUnescape(segment)
		if decodeErr != nil || decoded == "." || decoded == ".." || strings.ContainsAny(decoded, "/\\") {
			return "", fmt.Errorf("encoded traversal or separator in target")
		}
		segments = append(segments, segment)
	}
	resolved := strings.Join(segments, "/")
	if !secureNativePartName(resolved) {
		return "", fmt.Errorf("resolved target is not a canonical OPC part")
	}
	return resolved, nil
}

func previousNativeIdentities(previous *NativePPTXDeck) (nativeIdentityIndex, error) {
	index := emptyNativeIdentityIndex()
	if previous == nil {
		return index, nil
	}
	if issues := ValidateNativePPTX(*previous); len(issues) != 0 {
		return nativeIdentityIndex{}, fmt.Errorf("pptxpatch: native extract: invalid previous native deck: %w", NativeContractValidationError{Issues: issues})
	}
	if previous.Origin != NativeOriginParsed {
		return index, nil
	}
	index.documentID = previous.DocumentID
	for _, asset := range previous.Assets {
		if asset.Source == nil {
			continue
		}
		assetKey := nativeIdentityKey(asset.Source.PartName, asset.Source.ObjectID)
		if existing, ok := index.assets[assetKey]; ok && existing != asset.ID {
			return nativeIdentityIndex{}, fmt.Errorf("pptxpatch: native extract: Previous maps one asset source anchor to distinct IDs")
		}
		index.assets[assetKey] = asset.ID
	}
	for _, slide := range previous.Slides {
		if slide.Source == nil {
			continue
		}
		slideKey := nativeIdentityKey(slide.Source.PartName, slide.Source.ObjectID)
		if existing, ok := index.slides[slideKey]; ok && existing != slide.ID {
			return nativeIdentityIndex{}, fmt.Errorf("pptxpatch: native extract: Previous maps one slide source anchor to distinct IDs")
		}
		index.slides[slideKey] = slide.ID
		slideElements := map[string]string{}
		index.elementsBySlide[slideKey] = slideElements
		var walk func([]NativeElement) error
		walk = func(elements []NativeElement) error {
			for _, element := range elements {
				if element.Source != nil {
					elementKey := nativeIdentityKey(element.Source.PartName, element.Source.ObjectID)
					if existing, ok := index.elements[elementKey]; ok && existing != element.ID {
						return fmt.Errorf("pptxpatch: native extract: Previous maps one element source anchor to distinct IDs")
					}
					index.elements[elementKey] = element.ID
					slideElements[elementKey] = element.ID
				}
				if err := walk(element.Children); err != nil {
					return err
				}
			}
			return nil
		}
		if err := walk(slide.Elements); err != nil {
			return nativeIdentityIndex{}, err
		}
	}
	return index, nil
}

func emptyNativeIdentityIndex() nativeIdentityIndex {
	return nativeIdentityIndex{assets: map[string]string{}, slides: map[string]string{}, elements: map[string]string{}, elementsBySlide: map[string]map[string]string{}}
}

// Previous is an identity hint, never byte/capability authority. Establish deck
// continuity with at least one current part+sldId anchor, then retain only the
// identities scoped to exact matched slides and relationship-routed exact assets.
// The retained document ID is only a durable namespace label; every source anchor,
// digest, relationship closure, and passthrough capability is rebuilt from current
// bytes. This keeps edits/reorder/additions durable without cross-slide ID leakage.
func (extractor *nativeExtractor) bindPreviousIdentitiesToSlideList(slideList *nativeXMLNode, dialect nativeExtractDialect, relationships map[string]nativeExtractRelationship) error {
	if extractor.identities.documentID == "" {
		return nil
	}
	matchedSlideKeys := map[string]bool{}
	matchedSlideParts := map[string]string{}
	for _, slideIDNode := range slideList.Children {
		if slideIDNode.Name != (xml.Name{Space: dialect.presentation, Local: "sldId"}) {
			continue
		}
		nativeObjectID, err := canonicalNativeUnsignedID(slideIDNode, "", "id", 256)
		if err != nil {
			continue
		}
		relationshipID, ok := exactNativeAttr(slideIDNode, dialect.rels, "id")
		relationship, exists := relationships[relationshipID]
		if !ok || !exists || relationship.Type != dialect.relSlide || !relationship.internal() {
			continue
		}
		slideKey := nativeIdentityKey(relationship.Part, "sldId-"+nativeObjectID)
		if extractor.identities.slides[slideKey] != "" {
			matchedSlideKeys[slideKey] = true
			partAlias, _ := nativePartAlias(relationship.Part)
			matchedSlideParts[partAlias] = relationship.Part
		}
	}
	if len(matchedSlideKeys) == 0 {
		extractor.identities = emptyNativeIdentityIndex()
		return nil
	}

	filtered := emptyNativeIdentityIndex()
	filtered.documentID = extractor.identities.documentID
	for key := range matchedSlideKeys {
		filtered.slides[key] = extractor.identities.slides[key]
		filtered.elementsBySlide[key] = map[string]string{}
		for elementKey, id := range extractor.identities.elementsBySlide[key] {
			filtered.elements[elementKey] = id
			filtered.elementsBySlide[key][elementKey] = id
		}
	}
	allowedAssetParts := map[string]bool{}
	for _, part := range matchedSlideParts {
		slideRelationships, err := extractor.parseRelationships(part)
		if err != nil {
			return err
		}
		if err := extractor.validateRelationshipSet(slideRelationships, dialect, part); err != nil {
			return err
		}
		for _, relationship := range slideRelationships {
			if relationship.Type == dialect.relImage && relationship.internal() {
				alias, _ := nativePartAlias(relationship.Part)
				allowedAssetParts[alias] = true
			}
		}
	}
	if extractor.options.Previous != nil {
		for _, asset := range extractor.options.Previous.Assets {
			if asset.Source == nil || asset.Source.ObjectID != "asset-part" || asset.ByteLength == nil {
				continue
			}
			alias, err := nativePartAlias(asset.Source.PartName)
			if err != nil || !allowedAssetParts[alias] {
				continue
			}
			actualPart, exists := extractor.pkg.aliases[alias]
			payload, hasPayload := extractor.pkg.parts[actualPart]
			if !exists || !hasPayload || int64(len(payload)) != *asset.ByteLength {
				continue
			}
			digest := nativeSHA256(payload)
			contentType, _, contentErr := nativePictureContentType(extractor.pkg.contentTypes.forPart(actualPart))
			if contentErr != nil || asset.SHA256 != digest || asset.Source.FingerprintSHA256 != digest || asset.ContentType != contentType {
				continue
			}
			filtered.assets[nativeIdentityKey(actualPart, asset.Source.ObjectID)] = asset.ID
		}
	}
	extractor.identities = filtered
	return nil
}

func (extractor *nativeExtractor) extract() (NativePPTXDeck, error) {
	rootRelationships, err := extractor.parseRelationships("")
	if err != nil {
		return NativePPTXDeck{}, err
	}
	var officeDocument *nativeExtractRelationship
	for index := range rootRelationships {
		relationship := &rootRelationships[index]
		if relationship.Type == relOfficeDocumentTransitional || relationship.Type == relOfficeDocumentStrict {
			if officeDocument != nil || !relationship.internal() {
				return NativePPTXDeck{}, fmt.Errorf("pptxpatch: native extract OPC: officeDocument relationship must be unique and internal")
			}
			officeDocument = relationship
		}
	}
	if officeDocument == nil || !asciiEqualFoldNative(extractor.pkg.contentTypes.forPart(officeDocument.Part), contentTypePresentation) {
		return NativePPTXDeck{}, fmt.Errorf("pptxpatch: native extract OPC: missing presentation officeDocument with effective content type")
	}
	presentationRoot, err := parseNativeXML(extractor.pkg.parts[officeDocument.Part], officeDocument.Part)
	if err != nil {
		return NativePPTXDeck{}, err
	}
	dialect, err := nativeDialectForPresentation(presentationRoot.Name)
	if err != nil {
		return NativePPTXDeck{}, err
	}
	if (officeDocument.Type == relOfficeDocumentTransitional && dialect.presentation != nsPresentationTransitional) ||
		(officeDocument.Type == relOfficeDocumentStrict && dialect.presentation != nsPresentationStrict) || officeDocument.Namespace != dialect.packageRels {
		return NativePPTXDeck{}, fmt.Errorf("pptxpatch: native extract: officeDocument relationship dialect does not match presentation root")
	}
	if err := extractor.validateRelationshipSet(rootRelationships, dialect, "package root"); err != nil {
		return NativePPTXDeck{}, err
	}
	presentationUnsupported := []nativeUnsupportedSource{}
	if hasNativeSemanticAttrs(presentationRoot) || !onlyNativeXMLSpace(presentationRoot.Text) {
		unsupported, unsupportedErr := makeNativeUnsupportedSource(extractor.pkg.parts[officeDocument.Part], presentationRoot, officeDocument.Part, "presentation-root", "pptx.unsupported-presentation-markup", "presentation root contains unmodeled markup")
		if unsupportedErr != nil {
			return NativePPTXDeck{}, unsupportedErr
		}
		presentationUnsupported = append(presentationUnsupported, unsupported)
	}
	for index, child := range presentationRoot.Children {
		if child.Name != (xml.Name{Space: dialect.presentation, Local: "sldIdLst"}) && child.Name != (xml.Name{Space: dialect.presentation, Local: "sldSz"}) {
			unsupported, unsupportedErr := makeNativeUnsupportedSource(extractor.pkg.parts[officeDocument.Part], child, officeDocument.Part, fmt.Sprintf("presentation-child-%d-%d", index, child.RawStart), "pptx.unsupported-presentation-child", "presentation contains an unmodeled child")
			if unsupportedErr != nil {
				return NativePPTXDeck{}, unsupportedErr
			}
			if len(presentationUnsupported) >= nativeExtractMaxPassthrough {
				return NativePPTXDeck{}, fmt.Errorf("pptxpatch: native extract: presentation passthrough budget exceeded")
			}
			presentationUnsupported = append(presentationUnsupported, unsupported)
		}
	}
	slideSize, err := nativeSingleton(presentationRoot, dialect.presentation, "sldSz", true)
	if err != nil {
		return NativePPTXDeck{}, err
	}
	if err := requireOnlyNativeAttrs(slideSize, xml.Name{Local: "cx"}, xml.Name{Local: "cy"}); err != nil {
		unsupported, unsupportedErr := makeNativeUnsupportedSource(extractor.pkg.parts[officeDocument.Part], slideSize, officeDocument.Part, "presentation-slide-size", "pptx.unsupported-slide-size-metadata", err.Error())
		if unsupportedErr != nil {
			return NativePPTXDeck{}, unsupportedErr
		}
		presentationUnsupported = append(presentationUnsupported, unsupported)
	}
	if err := requireOnlyNativeChildren(slideSize); err != nil {
		return NativePPTXDeck{}, fmt.Errorf("pptxpatch: native extract: invalid slide size: %w", err)
	}
	cx, err := requiredNativePositiveInt64(slideSize, "", "cx")
	if err != nil {
		return NativePPTXDeck{}, err
	}
	cy, err := requiredNativePositiveInt64(slideSize, "", "cy")
	if err != nil {
		return NativePPTXDeck{}, err
	}
	presentationRelationships, err := extractor.parseRelationships(officeDocument.Part)
	if err != nil {
		return NativePPTXDeck{}, err
	}
	if err := extractor.validateRelationshipSet(presentationRelationships, dialect, officeDocument.Part); err != nil {
		return NativePPTXDeck{}, err
	}
	for _, relationshipOwner := range []struct {
		part string
		id   string
	}{
		{part: "", id: "package-relationships"},
		{part: officeDocument.Part, id: "presentation-relationships"},
	} {
		relsPart, relsErr := extractor.actualRelationshipsPart(relationshipOwner.part)
		if relsErr != nil {
			return NativePPTXDeck{}, relsErr
		}
		if len(presentationUnsupported) >= nativeExtractMaxPassthrough {
			return NativePPTXDeck{}, fmt.Errorf("pptxpatch: native extract: presentation relationship closure budget exceeded")
		}
		relsPayload := extractor.pkg.parts[relsPart]
		presentationUnsupported = append(presentationUnsupported, makeNativeUnsupportedPayload(relsPayload, relsPart, relationshipOwner.id, "pptx.relationship-map-preserve", "package relationship map is capability-bound to the source revision"))
	}
	relsByID := map[string]nativeExtractRelationship{}
	for _, relationship := range presentationRelationships {
		relsByID[relationship.ID] = relationship
	}
	slideList, err := nativeSingleton(presentationRoot, dialect.presentation, "sldIdLst", true)
	if err != nil {
		return NativePPTXDeck{}, err
	}
	if err := requireOnlyNativeAttrs(slideList); err != nil {
		return NativePPTXDeck{}, fmt.Errorf("pptxpatch: native extract: slide list: %w", err)
	}
	if !onlyNativeXMLSpace(slideList.Text) {
		return NativePPTXDeck{}, fmt.Errorf("pptxpatch: native extract: slide list contains direct text")
	}
	if len(slideList.Children) > nativeMaxSlides {
		return NativePPTXDeck{}, fmt.Errorf("pptxpatch: native extract: slide count exceeds %d", nativeMaxSlides)
	}
	extractor.outputNodesEmitted += len(slideList.Children)
	if err := extractor.bindPreviousIdentitiesToSlideList(slideList, dialect, relsByID); err != nil {
		return NativePPTXDeck{}, err
	}
	documentID := extractor.identities.documentID
	if documentID == "" {
		documentID = stableNativeID("deck", officeDocument.Part, nativeSHA256(extractor.pkg.parts[officeDocument.Part]))
	}
	extractor.documentID = documentID
	slides := make([]NativeSlide, 0, len(slideList.Children))
	slideObjectIDs := map[string]bool{}
	usedSlideRelationshipIDs := map[string]bool{}
	usedSlideParts := map[string]bool{}
	for _, slideIDNode := range slideList.Children {
		if slideIDNode.Name != (xml.Name{Space: dialect.presentation, Local: "sldId"}) {
			return NativePPTXDeck{}, fmt.Errorf("pptxpatch: native extract: unknown slide-list element")
		}
		if err := requireOnlyNativeAttrs(slideIDNode, xml.Name{Local: "id"}, xml.Name{Space: dialect.rels, Local: "id"}); err != nil {
			return NativePPTXDeck{}, fmt.Errorf("pptxpatch: native extract: slide id: %w", err)
		}
		if err := requireOnlyNativeChildren(slideIDNode); err != nil {
			return NativePPTXDeck{}, fmt.Errorf("pptxpatch: native extract: slide id: %w", err)
		}
		nativeObjectID, idErr := canonicalNativeUnsignedID(slideIDNode, "", "id", 256)
		relationshipID, relOK := exactNativeAttr(slideIDNode, dialect.rels, "id")
		relationship, exists := relsByID[relationshipID]
		if idErr != nil || slideObjectIDs[nativeObjectID] {
			return NativePPTXDeck{}, fmt.Errorf("pptxpatch: native extract: invalid or duplicate slide native id")
		}
		slideObjectIDs[nativeObjectID] = true
		if !relOK || !exists || relationship.Type != dialect.relSlide || !relationship.internal() || !asciiEqualFoldNative(extractor.pkg.contentTypes.forPart(relationship.Part), contentTypeSlide) {
			return NativePPTXDeck{}, fmt.Errorf("pptxpatch: native extract: slide relationship is missing, external, or has wrong type/content type")
		}
		partAlias, _ := nativePartAlias(relationship.Part)
		if usedSlideRelationshipIDs[relationshipID] || usedSlideParts[partAlias] {
			return NativePPTXDeck{}, fmt.Errorf("pptxpatch: native extract: duplicate slide relationship or target part in slide order")
		}
		usedSlideRelationshipIDs[relationshipID] = true
		usedSlideParts[partAlias] = true
		slide, slideErr := extractor.extractSlide(relationship.Part, nativeObjectID, relationship.ID, dialect)
		if slideErr != nil {
			return NativePPTXDeck{}, slideErr
		}
		slides = append(slides, slide)
	}
	if len(presentationUnsupported) != 0 && len(slides) == 0 {
		return NativePPTXDeck{}, fmt.Errorf("pptxpatch: native extract: unsupported presentation markup cannot be capability-bound without an owning slide")
	}
	for index := range slides {
		for _, unsupported := range presentationUnsupported {
			if err := extractor.markSlideUnsupported(&slides[index], unsupported.part, unsupported.objectID, unsupported.fingerprint, unsupported.payload, unsupported.code, unsupported.message); err != nil {
				return NativePPTXDeck{}, err
			}
		}
	}
	compatibility := NativeCompatibility{Status: NativeCompatibilityStatusEditable, Diagnostics: []NativeDiagnostic{}}
	for _, slide := range slides {
		compatibility.Status = worseNativeStatus(compatibility.Status, slide.Compatibility.Status)
		if len(compatibility.Diagnostics) > nativeMaxDiagnosticsPerScope-len(slide.Compatibility.Diagnostics) {
			return NativePPTXDeck{}, fmt.Errorf("pptxpatch: native extract: deck diagnostic budget exceeded")
		}
		compatibility.Diagnostics = append(compatibility.Diagnostics, slide.Compatibility.Diagnostics...)
	}
	revision := extractor.sourceRevision
	assets := make([]NativeAsset, len(extractor.assets))
	copy(assets, extractor.assets)
	sort.Slice(assets, func(left, right int) bool { return assets[left].ID < assets[right].ID })
	return NativePPTXDeck{
		ContractVersion: NativePPTXContractVersion,
		DocumentID:      documentID,
		Origin:          NativeOriginParsed,
		SourceRevision:  &revision,
		Size:            NativeSize{Cx: int64Pointer(cx), Cy: int64Pointer(cy)},
		Assets:          assets, Slides: slides,
		Compatibility: compatibility,
	}, nil
}

func nativeDialectForPresentation(name xml.Name) (nativeExtractDialect, error) {
	switch name {
	case xml.Name{Space: nsPresentationTransitional, Local: "presentation"}:
		return nativeExtractDialect{presentation: nsPresentationTransitional, drawing: nsDrawingTransitional, rels: nsOfficeRelsTransitional, relSlide: relSlideTransitional, relSlideLayout: relSlideLayoutTransitional, relSlideMaster: relSlideMasterTransitional, relTheme: relThemeTransitional, relImage: relImageTransitional, relChart: relChartTransitional, chart: nsChartTransitional, packageRels: nsPackageRels}, nil
	case xml.Name{Space: nsPresentationStrict, Local: "presentation"}:
		return nativeExtractDialect{presentation: nsPresentationStrict, drawing: nsDrawingStrict, rels: nsOfficeRelsStrict, relSlide: relSlideStrict, relSlideLayout: relSlideLayoutStrict, relSlideMaster: relSlideMasterStrict, relTheme: relThemeStrict, relImage: relImageStrict, relChart: relChartStrict, chart: nsChartStrict, packageRels: nsPackageRels}, nil
	default:
		return nativeExtractDialect{}, fmt.Errorf("pptxpatch: native extract: unsupported presentation namespace/root")
	}
}

func (extractor *nativeExtractor) extractSlide(part, objectID, relationshipID string, dialect nativeExtractDialect) (NativeSlide, error) {
	payload := extractor.pkg.parts[part]
	root, err := parseNativeXML(payload, part)
	if err != nil {
		return NativeSlide{}, err
	}
	if root.Name != (xml.Name{Space: dialect.presentation, Local: "sld"}) {
		return NativeSlide{}, fmt.Errorf("pptxpatch: native extract: slide %q has wrong namespace/root", part)
	}
	rootRaw, err := rawNativeNode(payload, root)
	if err != nil {
		return NativeSlide{}, err
	}
	fingerprint := nativeSHA256(rootRaw)
	sourceObjectID := "sldId-" + objectID
	slideID := extractor.identities.slides[nativeIdentityKey(part, sourceObjectID)]
	if slideID == "" {
		slideID = stableNativeID("slide", extractor.documentID+"\x00"+part, sourceObjectID)
	}
	relID := relationshipID
	slide := NativeSlide{
		ID: slideID, Provenance: NativeProvenanceParsed,
		Elements: []NativeElement{}, Passthrough: []NativePassthroughRef{},
		Source:        &NativeSourceAnchor{PartName: part, ObjectID: sourceObjectID, RelationshipID: &relID, FingerprintSHA256: fingerprint},
		Compatibility: NativeCompatibility{Status: NativeCompatibilityStatusEditable, Diagnostics: []NativeDiagnostic{}},
	}
	graph, err := extractor.resolveSlideDependencyGraph(part, dialect)
	if err != nil {
		return NativeSlide{}, err
	}
	theme, err := resolveNativeTheme(graph, dialect)
	if err != nil {
		return NativeSlide{}, err
	}
	extractor.theme = theme
	extractor.slideDependencies = graph
	for _, unsupported := range graph.unsupported {
		if err := extractor.markSlideUnsupported(&slide, unsupported.part, unsupported.objectID, unsupported.fingerprint, unsupported.payload, unsupported.code, unsupported.message); err != nil {
			return NativeSlide{}, err
		}
	}
	if hasNativeSemanticAttrs(root) || !onlyNativeXMLSpace(root.Text) {
		if err := extractor.markSlideUnsupported(&slide, part, sourceObjectID, fingerprint, rootRaw, "pptx.unsupported-slide-markup", "slide root contains unmodeled markup"); err != nil {
			return NativeSlide{}, err
		}
	}
	for index, child := range root.Children {
		if child.Name == (xml.Name{Space: dialect.presentation, Local: "cSld"}) {
			continue
		}
		unsupported, unsupportedErr := makeNativeUnsupportedSource(payload, child, part, fmt.Sprintf("%s-property-%d-%d", sourceObjectID, index, child.RawStart), "pptx.unsupported-slide-property", "slide transition, timing, color-map, or extension markup is not representable in native PPTX v1")
		if unsupportedErr != nil {
			return NativeSlide{}, unsupportedErr
		}
		if err := extractor.markSlideUnsupported(&slide, unsupported.part, unsupported.objectID, unsupported.fingerprint, unsupported.payload, unsupported.code, unsupported.message); err != nil {
			return NativeSlide{}, err
		}
	}
	cSld, err := nativeSingleton(root, dialect.presentation, "cSld", true)
	if err != nil {
		return NativeSlide{}, err
	}
	if hasNativeSemanticAttrs(cSld) || !onlyNativeXMLSpace(cSld.Text) {
		unsupported, unsupportedErr := makeNativeUnsupportedSource(payload, cSld, part, sourceObjectID+"-common", "pptx.unsupported-common-slide-data", "common slide data contains unmodeled name or text metadata")
		if unsupportedErr != nil {
			return NativeSlide{}, unsupportedErr
		}
		if err := extractor.markSlideUnsupported(&slide, unsupported.part, unsupported.objectID, unsupported.fingerprint, unsupported.payload, unsupported.code, unsupported.message); err != nil {
			return NativeSlide{}, err
		}
	}
	if _, err := nativeSingleton(cSld, dialect.presentation, "bg", false); err != nil {
		return NativeSlide{}, err
	}
	for index, child := range cSld.Children {
		switch child.Name {
		case xml.Name{Space: dialect.presentation, Local: "spTree"}:
		case xml.Name{Space: dialect.presentation, Local: "bg"}:
			background, backgroundErr := extractNativeSlideBackground(child, dialect)
			if backgroundErr != nil {
				unsupported, unsupportedErr := makeNativeUnsupportedSource(payload, child, part, sourceObjectID+"-background", "pptx.unsupported-background", backgroundErr.Error())
				if unsupportedErr != nil {
					return NativeSlide{}, unsupportedErr
				}
				if err := extractor.markSlideUnsupported(&slide, unsupported.part, unsupported.objectID, unsupported.fingerprint, unsupported.payload, unsupported.code, unsupported.message); err != nil {
					return NativeSlide{}, err
				}
			} else {
				slide.Background = &background
			}
		default:
			unsupported, unsupportedErr := makeNativeUnsupportedSource(payload, child, part, fmt.Sprintf("%s-common-child-%d-%d", sourceObjectID, index, child.RawStart), "pptx.unsupported-common-slide-child", "common slide data contains unmodeled content")
			if unsupportedErr != nil {
				return NativeSlide{}, unsupportedErr
			}
			if err := extractor.markSlideUnsupported(&slide, unsupported.part, unsupported.objectID, unsupported.fingerprint, unsupported.payload, unsupported.code, unsupported.message); err != nil {
				return NativeSlide{}, err
			}
		}
	}
	spTree, err := nativeSingleton(cSld, dialect.presentation, "spTree", true)
	if err != nil {
		return NativeSlide{}, err
	}
	objectIDs, err := collectNativeShapeTreeObjectIDs(spTree, dialect)
	if err != nil {
		return NativeSlide{}, err
	}
	if len(objectIDs) > nativeMaxTotalElements-extractor.elementsEmitted || len(objectIDs) > nativeMaxNodes-extractor.outputNodesEmitted {
		return NativeSlide{}, fmt.Errorf("pptxpatch: native extract: element/output node budget exceeded")
	}
	extractor.elementsEmitted += len(objectIDs)
	extractor.outputNodesEmitted += len(objectIDs)
	rootObjectID, err := validateNativeRootGroupScaffold(spTree, dialect)
	if err != nil {
		return NativeSlide{}, err
	}
	rootGroupProperties, err := nativeSingleton(spTree, dialect.presentation, "grpSpPr", true)
	if err != nil {
		return NativeSlide{}, err
	}
	if hasNativeSemanticAttrs(rootGroupProperties) || len(rootGroupProperties.Children) != 0 || !onlyNativeXMLSpace(rootGroupProperties.Text) {
		raw, rawErr := rawNativeNode(payload, rootGroupProperties)
		if rawErr != nil {
			return NativeSlide{}, rawErr
		}
		if err := extractor.markSlideUnsupported(&slide, part, rootObjectID, nativeSHA256(raw), raw, "pptx.unsupported-root-group-transform", "slide root group transform is not representable in native PPTX v1"); err != nil {
			return NativeSlide{}, err
		}
	}
	usedPictureRelationships := map[string]bool{}
	for _, child := range spTree.Children {
		switch child.Name {
		case xml.Name{Space: dialect.presentation, Local: "nvGrpSpPr"}, xml.Name{Space: dialect.presentation, Local: "grpSpPr"}:
			continue
		case xml.Name{Space: dialect.presentation, Local: "sp"}:
			elementObjectID := objectIDs[child]
			textBox, textBoxErr := nativeShapeIsTextBox(child, dialect)
			if textBoxErr != nil {
				return NativeSlide{}, textBoxErr
			}
			var element NativeElement
			var elementErr error
			if textBox {
				element, elementErr = extractor.extractTextShape(child, part, slideID, fingerprint, len(slide.Elements), dialect)
			} else {
				element, elementErr = extractor.extractAutoShape(child, part, slideID, dialect)
			}
			if elementErr != nil {
				if !textBox {
					return NativeSlide{}, elementErr
				}
				var duplicate nativeDuplicateSingletonError
				if errors.As(elementErr, &duplicate) {
					return NativeSlide{}, elementErr
				}
				raw, rawErr := rawNativeNode(payload, child)
				if rawErr != nil {
					return NativeSlide{}, rawErr
				}
				if err := extractor.markSlideUnsupported(&slide, part, elementObjectID, nativeSHA256(raw), raw, "pptx.unsupported-shape", elementErr.Error()); err != nil {
					return NativeSlide{}, err
				}
				continue
			}
			slide.Elements = append(slide.Elements, element)
			slide.Compatibility.Status = worseNativeStatus(slide.Compatibility.Status, element.Compatibility.Status)
			if element.Compatibility.Status != NativeCompatibilityStatusEditable {
				slide.Compatibility.Diagnostics = append(slide.Compatibility.Diagnostics, element.Compatibility.Diagnostics...)
			}
		case xml.Name{Space: dialect.presentation, Local: "pic"}:
			element, elementErr := extractor.extractPicture(child, part, slideID, graph.relationships, dialect)
			if elementErr != nil {
				return NativeSlide{}, elementErr
			}
			usedPictureRelationships[*element.Source.RelationshipID] = true
			slide.Elements = append(slide.Elements, element)
			slide.Compatibility.Status = worseNativeStatus(slide.Compatibility.Status, element.Compatibility.Status)
			if element.Compatibility.Status != NativeCompatibilityStatusEditable {
				slide.Compatibility.Diagnostics = append(slide.Compatibility.Diagnostics, element.Compatibility.Diagnostics...)
			}
		case xml.Name{Space: dialect.presentation, Local: "cxnSp"}:
			element, elementErr := extractor.extractConnector(child, part, slideID, dialect)
			if elementErr != nil {
				return NativeSlide{}, elementErr
			}
			slide.Elements = append(slide.Elements, element)
			slide.Compatibility.Status = worseNativeStatus(slide.Compatibility.Status, element.Compatibility.Status)
			if element.Compatibility.Status != NativeCompatibilityStatusEditable {
				slide.Compatibility.Diagnostics = append(slide.Compatibility.Diagnostics, element.Compatibility.Diagnostics...)
			}
		case xml.Name{Space: dialect.presentation, Local: "graphicFrame"}:
			element, elementErr := extractor.extractNativeGraphicFrame(child, part, slideID, graph.relationships, dialect)
			if elementErr != nil {
				var refusal nativeGraphicFrameProjectionRefusal
				if !errors.As(elementErr, &refusal) {
					return NativeSlide{}, elementErr
				}
				raw, rawErr := rawNativeNode(payload, child)
				if rawErr != nil {
					return NativeSlide{}, rawErr
				}
				if err := extractor.markSlideUnsupported(&slide, part, objectIDs[child], nativeSHA256(raw), raw, refusal.code, refusal.message); err != nil {
					return NativeSlide{}, err
				}
				continue
			}
			slide.Elements = append(slide.Elements, element)
			slide.Compatibility.Status = worseNativeStatus(slide.Compatibility.Status, element.Compatibility.Status)
			if element.Compatibility.Status != NativeCompatibilityStatusEditable {
				slide.Compatibility.Diagnostics = append(slide.Compatibility.Diagnostics, element.Compatibility.Diagnostics...)
			}
		case xml.Name{Space: dialect.presentation, Local: "grpSp"}:
			group, groupErr := extractor.extractNativeGroupAtomically(child, part, slideID, fingerprint, dialect, graph.relationships, nativeMaxDiagnosticsPerScope-len(slide.Compatibility.Diagnostics))
			if groupErr != nil {
				var refusal nativeGroupProjectionRefusal
				if !errors.As(groupErr, &refusal) {
					return NativeSlide{}, groupErr
				}
				raw, rawErr := rawNativeNode(payload, child)
				if rawErr != nil {
					return NativeSlide{}, rawErr
				}
				if err := extractor.markSlideUnsupported(&slide, part, objectIDs[child], nativeSHA256(raw), raw, refusal.code, refusal.message); err != nil {
					return NativeSlide{}, err
				}
				continue
			}
			for usedRelationship := range group.usedPictureRelationships {
				usedPictureRelationships[usedRelationship] = true
			}
			slide.Elements = append(slide.Elements, group.element)
			slide.Compatibility.Status = worseNativeStatus(slide.Compatibility.Status, group.element.Compatibility.Status)
			if group.element.Compatibility.Status != NativeCompatibilityStatusEditable {
				slide.Compatibility.Diagnostics = append(slide.Compatibility.Diagnostics, group.element.Compatibility.Diagnostics...)
			}
		default:
			raw, rawErr := rawNativeNode(payload, child)
			if rawErr != nil {
				return NativeSlide{}, rawErr
			}
			unknownObjectID := objectIDs[child]
			if unknownObjectID == "" {
				unknownObjectID = "xml-" + nativeSHA256(raw)[:24]
			}
			if err := extractor.markSlideUnsupported(&slide, part, unknownObjectID, nativeSHA256(raw), raw, "pptx.unsupported-slide-child", "slide contains content outside the native v1 subset"); err != nil {
				return NativeSlide{}, err
			}
		}
	}
	for _, relationship := range graph.relationships {
		if relationship.Type != dialect.relImage || usedPictureRelationships[relationship.ID] || !relationship.internal() {
			continue
		}
		payload := extractor.pkg.parts[relationship.Part]
		objectID := "unused-image-rel-" + nativeSHA256([]byte(relationship.ID))[:24]
		if err := extractor.markSlideUnsupported(&slide, relationship.Part, objectID, nativeSHA256(payload), payload, "pptx.unreferenced-image-relationship", "slide image relationship target is not referenced by a modeled native picture"); err != nil {
			return NativeSlide{}, err
		}
	}
	return slide, nil
}

func (extractor *nativeExtractor) extractTextShape(node *nativeXMLNode, part, slideID, fingerprint string, zIndex int, dialect nativeExtractDialect) (NativeElement, error) {
	resolved, inheritedPlaceholder, inheritanceErr := extractor.resolveNativePlaceholder(node, dialect)
	if inheritanceErr != nil {
		return NativeElement{}, inheritanceErr
	}
	node = resolved
	if err := requireOnlyNativeAttrs(node); err != nil {
		return NativeElement{}, err
	}
	if err := requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.presentation, Local: "nvSpPr"},
		xml.Name{Space: dialect.presentation, Local: "spPr"},
		xml.Name{Space: dialect.presentation, Local: "txBody"}); err != nil {
		return NativeElement{}, err
	}
	nvSpPr, err := nativeSingleton(node, dialect.presentation, "nvSpPr", true)
	if err != nil {
		return NativeElement{}, err
	}
	cNvPr, err := nativeSingleton(nvSpPr, dialect.presentation, "cNvPr", true)
	if err != nil {
		return NativeElement{}, err
	}
	cNvSpPr, err := nativeSingleton(nvSpPr, dialect.presentation, "cNvSpPr", true)
	if err != nil {
		return NativeElement{}, err
	}
	nvPr, err := nativeSingleton(nvSpPr, dialect.presentation, "nvPr", true)
	if err != nil {
		return NativeElement{}, err
	}
	if err := requireOnlyNativeAttrs(nvSpPr); err != nil {
		return NativeElement{}, err
	}
	if err := requireOnlyNativeChildren(nvSpPr,
		xml.Name{Space: dialect.presentation, Local: "cNvPr"},
		xml.Name{Space: dialect.presentation, Local: "cNvSpPr"},
		xml.Name{Space: dialect.presentation, Local: "nvPr"}); err != nil {
		return NativeElement{}, err
	}
	if err := requireEmptyNativeElement(nvPr); err != nil {
		return NativeElement{}, fmt.Errorf("pptxpatch: native extract: placeholder/inheritance metadata is unsupported: %w", err)
	}
	if err := requireOnlyNativeAttrs(cNvPr, xml.Name{Local: "id"}, xml.Name{Local: "name"}); err != nil {
		return NativeElement{}, err
	}
	if err := requireOnlyNativeChildren(cNvPr); err != nil {
		return NativeElement{}, err
	}
	if err := requireOnlyNativeAttrs(cNvSpPr, xml.Name{Local: "txBox"}); err != nil {
		return NativeElement{}, err
	}
	if err := requireOnlyNativeChildren(cNvSpPr); err != nil {
		return NativeElement{}, err
	}
	nativeObjectID, err := canonicalNativeUnsignedID(cNvPr, "", "id", 1)
	if err != nil {
		return NativeElement{}, err
	}
	objectID := "cNvPr-" + nativeObjectID
	elementID := extractor.identities.elements[nativeIdentityKey(part, objectID)]
	if elementID == "" {
		elementID = stableNativeID("element", extractor.documentID+"\x00"+part, objectID)
	}
	name, _ := exactNativeAttr(cNvPr, "", "name")
	spPr, err := nativeSingleton(node, dialect.presentation, "spPr", true)
	if err != nil {
		return NativeElement{}, err
	}
	if err := requireOnlyNativeAttrs(spPr); err != nil {
		return NativeElement{}, err
	}
	if err := validateNativeTextBoxShapeProperties(spPr, dialect); err != nil {
		return NativeElement{}, err
	}
	xfrm, err := nativeSingleton(spPr, dialect.drawing, "xfrm", true)
	if err != nil {
		return NativeElement{}, err
	}
	if err := requireOnlyNativeAttrs(xfrm); err != nil {
		return NativeElement{}, fmt.Errorf("pptxpatch: native extract: unsupported transform: %w", err)
	}
	if err := requireOnlyNativeChildren(xfrm,
		xml.Name{Space: dialect.drawing, Local: "off"},
		xml.Name{Space: dialect.drawing, Local: "ext"}); err != nil {
		return NativeElement{}, fmt.Errorf("pptxpatch: native extract: unsupported transform: %w", err)
	}
	off, err := nativeSingleton(xfrm, dialect.drawing, "off", true)
	if err != nil {
		return NativeElement{}, err
	}
	ext, err := nativeSingleton(xfrm, dialect.drawing, "ext", true)
	if err != nil {
		return NativeElement{}, err
	}
	if err := requireOnlyNativeAttrs(off, xml.Name{Local: "x"}, xml.Name{Local: "y"}); err != nil {
		return NativeElement{}, err
	}
	if err := requireOnlyNativeChildren(off); err != nil {
		return NativeElement{}, err
	}
	if err := requireOnlyNativeAttrs(ext, xml.Name{Local: "cx"}, xml.Name{Local: "cy"}); err != nil {
		return NativeElement{}, err
	}
	if err := requireOnlyNativeChildren(ext); err != nil {
		return NativeElement{}, err
	}
	x, err := requiredNativeInt64(off, "", "x")
	if err != nil {
		return NativeElement{}, err
	}
	y, err := requiredNativeInt64(off, "", "y")
	if err != nil {
		return NativeElement{}, err
	}
	cx, err := requiredNativePositiveInt64(ext, "", "cx")
	if err != nil {
		return NativeElement{}, err
	}
	cy, err := requiredNativePositiveInt64(ext, "", "cy")
	if err != nil {
		return NativeElement{}, err
	}
	txBox, _ := exactNativeAttr(cNvSpPr, "", "txBox")
	if txBox != "1" && txBox != "true" {
		return NativeElement{}, fmt.Errorf("pptxpatch: native extract: shape %s is not an exact text box in the checkpoint subset", objectID)
	}
	txBody, err := nativeSingleton(node, dialect.presentation, "txBody", true)
	if err != nil {
		return NativeElement{}, err
	}
	if err := extractor.reserveNativeTextOutput(txBody, dialect); err != nil {
		return NativeElement{}, err
	}
	textBody, textLayoutErr := extractNativeTextBodyLayoutPolicy(txBody, dialect, extractor.options.AllowSourceFrameAutoFitPreview)
	textLayoutMessage := ""
	if textLayoutErr != nil {
		if !isNativeTextLayoutUnsupported(textLayoutErr) {
			return NativeElement{}, textLayoutErr
		}
		textLayoutMessage = textLayoutErr.Error()
		textBody = nil
	}
	paragraphs, paragraphErr := extractor.extractNativeParagraphs(txBody, dialect)
	textContentMessage := ""
	if paragraphErr != nil {
		if !isNativeTextContentUnsupported(paragraphErr) {
			return NativeElement{}, paragraphErr
		}
		textContentMessage = paragraphErr.Error()
		paragraphs = []NativeParagraph{}
	}
	paragraphPointer := &paragraphs
	raw, err := rawNativeNode(extractor.pkg.parts[part], node)
	if err != nil {
		return NativeElement{}, err
	}
	elementFingerprint := nativeSHA256(raw)
	transform := NativeTransform{X: int64Pointer(x), Y: int64Pointer(y), Cx: int64Pointer(cx), Cy: int64Pointer(cy)}
	if textLayoutMessage == "" {
		if boundsErr := validateNativeTextBodyBounds(textBody, transform); boundsErr != nil {
			if !isNativeTextLayoutUnsupported(boundsErr) {
				return NativeElement{}, boundsErr
			}
			textLayoutMessage = boundsErr.Error()
			textBody = nil
		}
	}
	element := NativeElement{
		Kind: NativeElementKindText, ID: elementID, Provenance: NativeProvenanceParsed,
		Placeholder: inheritedPlaceholder,
		Transform:   transform, Paragraphs: paragraphPointer, TextBody: textBody,
		Passthrough: []NativePassthroughRef{}, Children: nil,
		Source:        &NativeSourceAnchor{PartName: part, ObjectID: objectID, FingerprintSHA256: elementFingerprint},
		Compatibility: NativeCompatibility{Status: NativeCompatibilityStatusEditable, Diagnostics: []NativeDiagnostic{}},
	}
	if inheritedPlaceholder != nil {
		element.Compatibility.Status = NativeCompatibilityStatusPreserveOnly
		element.Compatibility.Diagnostics = append(element.Compatibility.Diagnostics, NativeDiagnostic{Severity: NativeDiagnosticSeverityWarning, Code: "pptx.inherited-placeholder-preview", Message: "title/body placeholder geometry and styles resolve through the exact layout/master relationship chain; inherited targets remain preserve-only", Scope: &NativeDiagnosticScope{SlideID: &slideID, ElementID: &elementID, PartName: &part}})
	}
	if name != "" {
		element.Name = stringPointer(name)
	}
	// The retained source-frame layout is approximate even when an independent
	// content refusal prevents painting text. Keep that provenance in both paths.
	nativeMarkSourceFrameAutoFit(&element)
	if textLayoutMessage == "" && textContentMessage == "" {
		nativeMarkVerticalTextPreview(&element)
		nativePreserveTextCheckingMetadata(&element, txBody, dialect)
		_ = fingerprint
		_ = zIndex
		return element, nil
	}
	if extractor.passthroughRefsEmitted >= nativeExtractMaxEmittedPassthrough {
		return NativeElement{}, fmt.Errorf("pptxpatch: native extract: cumulative emitted passthrough reference budget exceeded")
	}
	reason := "pptx.text-refused"
	nativeMarkVerticalTextPreview(&element)
	if textContentMessage == "" {
		reason = "pptx.text-layout-refused"
	} else if textLayoutMessage == "" {
		reason = "pptx.text-content-refused"
	}
	passthrough, err := extractor.issuePassthrough(part, objectID, elementFingerprint, raw, reason)
	if err != nil {
		return NativeElement{}, err
	}
	if err := extractor.reserveNativePassthroughReference(); err != nil {
		return NativeElement{}, err
	}
	element.Passthrough = append(element.Passthrough, passthrough)
	element.Compatibility.Status = NativeCompatibilityStatusRefused
	partName := part
	if textLayoutMessage != "" {
		element.Compatibility.Diagnostics = append(element.Compatibility.Diagnostics, NativeDiagnostic{
			Severity: NativeDiagnosticSeverityRefusal,
			Code:     "pptx.text-layout-unavailable",
			Message:  textLayoutMessage,
			Scope:    &NativeDiagnosticScope{SlideID: &slideID, ElementID: &elementID, PartName: &partName},
		})
	}
	if textContentMessage != "" {
		element.Compatibility.Diagnostics = append(element.Compatibility.Diagnostics, NativeDiagnostic{
			Severity: NativeDiagnosticSeverityRefusal,
			Code:     "pptx.text-content-unavailable",
			Message:  textContentMessage,
			Scope:    &NativeDiagnosticScope{SlideID: &slideID, ElementID: &elementID, PartName: &partName},
		})
	}
	_ = fingerprint
	_ = zIndex
	return element, nil
}

func (extractor *nativeExtractor) extractNativeParagraphs(txBody *nativeXMLNode, dialect nativeExtractDialect) ([]NativeParagraph, error) {
	if txBody == nil {
		return []NativeParagraph{}, nil
	}
	resolvedBody, styleErr := resolveNativeLocalTextStyles(txBody, dialect, extractor.theme)
	if styleErr != nil {
		return nil, styleErr
	}
	txBody = resolvedBody
	if err := requireOnlyNativeAttrs(txBody); err != nil {
		return nil, err
	}
	if err := requireOnlyNativeChildren(txBody,
		xml.Name{Space: dialect.drawing, Local: "bodyPr"},
		xml.Name{Space: dialect.drawing, Local: "lstStyle"},
		xml.Name{Space: dialect.drawing, Local: "p"}); err != nil {
		return nil, err
	}
	if _, err := nativeSingleton(txBody, dialect.drawing, "bodyPr", true); err != nil {
		return nil, err
	}
	lstStyle, err := nativeSingleton(txBody, dialect.drawing, "lstStyle", true)
	if err != nil {
		return nil, err
	}
	if err := requireEmptyNativeElement(lstStyle); err != nil {
		return nil, fmt.Errorf("pptxpatch: native extract: unmodeled list style metadata: %w", err)
	}
	paragraphs := []NativeParagraph{}
	for _, paragraphNode := range txBody.Children {
		if paragraphNode.Name == (xml.Name{Space: dialect.drawing, Local: "bodyPr"}) || paragraphNode.Name == (xml.Name{Space: dialect.drawing, Local: "lstStyle"}) {
			continue
		}
		if paragraphNode.Name != (xml.Name{Space: dialect.drawing, Local: "p"}) {
			return nil, fmt.Errorf("pptxpatch: native extract: unsupported txBody element %s", paragraphNode.Name.Local)
		}
		if err := requireOnlyNativeAttrs(paragraphNode); err != nil {
			return nil, err
		}
		for _, child := range paragraphNode.Children {
			if child.Name == (xml.Name{Space: dialect.drawing, Local: "br"}) {
				return nil, unsupportedNativeTextContent("DrawingML a:br requires modeled hard-break line metrics")
			}
		}
		if err := requireOnlyNativeChildren(paragraphNode,
			xml.Name{Space: dialect.drawing, Local: "pPr"},
			xml.Name{Space: dialect.drawing, Local: "r"},
			xml.Name{Space: dialect.drawing, Local: "endParaRPr"}); err != nil {
			return nil, err
		}
		if _, err := nativeSingleton(paragraphNode, dialect.drawing, "pPr", true); err != nil {
			return nil, err
		}
		endParaRPr, err := nativeSingleton(paragraphNode, dialect.drawing, "endParaRPr", false)
		if err != nil {
			return nil, err
		}
		paragraph := NativeParagraph{Runs: []NativeTextRun{}}
		for _, child := range paragraphNode.Children {
			switch child.Name {
			case xml.Name{Space: dialect.drawing, Local: "pPr"}:
				if err := requireOnlyNativeAttrs(child, xml.Name{Local: "algn"}, xml.Name{Local: "lvl"}, xml.Name{Local: "marL"}, xml.Name{Local: "indent"}); err != nil {
					return nil, fmt.Errorf("pptxpatch: native extract: unmodeled paragraph metadata: %w", err)
				}
				if err := requireOnlyNativeChildren(child,
					xml.Name{Space: dialect.drawing, Local: "buNone"},
					xml.Name{Space: dialect.drawing, Local: "buFont"},
					xml.Name{Space: dialect.drawing, Local: "buChar"}); err != nil {
					return nil, fmt.Errorf("pptxpatch: native extract: unmodeled paragraph metadata: %w", err)
				}
				buNone, noneErr := nativeSingleton(child, dialect.drawing, "buNone", false)
				buChar, charErr := nativeSingleton(child, dialect.drawing, "buChar", false)
				if noneErr != nil || charErr != nil || (buNone != nil && buChar != nil) {
					return nil, fmt.Errorf("pptxpatch: native extract: duplicate or conflicting bullet metadata")
				}
				if buNone != nil {
					if err := requireEmptyNativeElement(buNone); err != nil {
						return nil, err
					}
				}
				if buChar != nil {
					if requireOnlyNativeAttrs(buChar, xml.Name{Local: "char"}) != nil || requireOnlyNativeChildren(buChar) != nil {
						return nil, unsupportedNativeTextContent("unmodeled bullet character metadata")
					}
					marker, ok := exactNativeAttr(buChar, "", "char")
					if !ok || utf8.RuneCountInString(marker) != 1 || strings.IndexFunc(marker, unicode.IsControl) >= 0 {
						return nil, unsupportedNativeTextContent("bullet requires one non-control authored character")
					}
					paragraph.Bullet = boolPointer(true)
					paragraph.BulletCharacter = &marker
				}
				bulletFont, fontErr := nativeSingleton(child, dialect.drawing, "buFont", false)
				if fontErr != nil {
					return nil, fontErr
				}
				if bulletFont != nil {
					if buChar == nil {
						return nil, unsupportedNativeTextContent("authored bullet font requires an explicit character")
					}
					family, err := nativeBulletFontFamily(bulletFont)
					if err != nil {
						return nil, err
					}
					paragraph.BulletFontFamily = &family
					if charset, ok := exactNativeAttr(bulletFont, "", "charset"); ok && charset == "2" {
						if paragraph.BulletCharacter == nil || len(*paragraph.BulletCharacter) != 1 || (*paragraph.BulletCharacter)[0] < 32 || (*paragraph.BulletCharacter)[0] > 126 {
							return nil, unsupportedNativeTextContent("buFont symbol preview requires an ASCII graphic source byte")
						}
						paragraph.BulletFontEncoding = stringPointer("windows-symbol-byte-v1")
					}
				}
				for _, field := range []struct {
					name   string
					target **int64
					min    int64
				}{{"marL", &paragraph.MarginLeftEmu, 0}, {"indent", &paragraph.IndentEmu, -51206400}} {
					if value, ok := exactNativeAttr(child, "", field.name); ok {
						parsed, err := parseCanonicalNativeInt(value, field.min, 51206400)
						if err != nil {
							return nil, err
						}
						*field.target = &parsed
					}
				}
				if value, ok := exactNativeAttr(child, "", "algn"); ok {
					align, alignErr := nativeTextAlign(value)
					if alignErr != nil {
						return nil, alignErr
					}
					paragraph.Align = &align
				}
				if value, ok := exactNativeAttr(child, "", "lvl"); ok {
					level, parseErr := strconv.ParseInt(value, 10, 64)
					if parseErr != nil || level < 0 || level > 8 {
						return nil, fmt.Errorf("pptxpatch: native extract: invalid paragraph level")
					}
					paragraph.Level = int64Pointer(level)
				}
				if buNone != nil {
					paragraph.Bullet = boolPointer(false)
				}
			case xml.Name{Space: dialect.drawing, Local: "r"}:
				run, runErr := extractor.extractNativeTextRun(child, dialect)
				if runErr != nil {
					return nil, runErr
				}
				paragraph.Runs = append(paragraph.Runs, run)
			case xml.Name{Space: dialect.drawing, Local: "endParaRPr"}:
				continue
			default:
				return nil, fmt.Errorf("pptxpatch: native extract: unsupported paragraph child %s", child.Name.Local)
			}
		}
		if paragraph.Align == nil || paragraph.Level == nil || paragraph.Bullet == nil {
			return nil, fmt.Errorf("pptxpatch: native extract: paragraph formatting is not self-contained")
		}
		if len(paragraph.Runs) == 0 {
			return nil, fmt.Errorf("pptxpatch: native extract: empty paragraph lacks self-contained end-paragraph font metrics")
		}
		if err := qualifyNativeEndParagraphMetadata(endParaRPr, paragraph); err != nil {
			return nil, err
		}
		paragraphs = append(paragraphs, paragraph)
	}
	if len(paragraphs) == 0 {
		return nil, fmt.Errorf("pptxpatch: native extract: text body lacks a self-contained paragraph")
	}
	return paragraphs, nil
}

func extractNativeSlideBackground(node *nativeXMLNode, dialect nativeExtractDialect) (string, error) {
	if err := requireOnlyNativeAttrs(node); err != nil {
		return "", err
	}
	if err := requireOnlyNativeChildren(node, xml.Name{Space: dialect.presentation, Local: "bgPr"}); err != nil {
		return "", err
	}
	bgPr, err := nativeSingleton(node, dialect.presentation, "bgPr", true)
	if err != nil {
		return "", err
	}
	if err := requireOnlyNativeAttrs(bgPr); err != nil {
		return "", err
	}
	if err := requireOnlyNativeChildren(bgPr, xml.Name{Space: dialect.drawing, Local: "solidFill"}); err != nil {
		return "", fmt.Errorf("pptxpatch: native extract: only explicit solid slide backgrounds are representable: %w", err)
	}
	solidFill, err := nativeSingleton(bgPr, dialect.drawing, "solidFill", true)
	if err != nil {
		return "", err
	}
	if err := requireOnlyNativeAttrs(solidFill); err != nil {
		return "", err
	}
	if err := requireOnlyNativeChildren(solidFill, xml.Name{Space: dialect.drawing, Local: "srgbClr"}); err != nil {
		return "", err
	}
	colorNode, err := nativeSingleton(solidFill, dialect.drawing, "srgbClr", true)
	if err != nil {
		return "", err
	}
	if err := requireOnlyNativeAttrs(colorNode, xml.Name{Local: "val"}); err != nil {
		return "", err
	}
	if err := requireOnlyNativeChildren(colorNode); err != nil {
		return "", err
	}
	color, ok := exactNativeAttr(colorNode, "", "val")
	color = strings.ToUpper(color)
	if !ok || !colorPattern.MatchString(color) {
		return "", fmt.Errorf("pptxpatch: native extract: invalid explicit background color")
	}
	return color, nil
}

func (extractor *nativeExtractor) extractNativeTextRun(node *nativeXMLNode, dialect nativeExtractDialect) (NativeTextRun, error) {
	if err := requireOnlyNativeAttrs(node); err != nil {
		return NativeTextRun{}, err
	}
	if err := requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.drawing, Local: "rPr"},
		xml.Name{Space: dialect.drawing, Local: "t"}); err != nil {
		return NativeTextRun{}, err
	}
	textNode, err := nativeSingleton(node, dialect.drawing, "t", true)
	if err != nil {
		return NativeTextRun{}, err
	}
	if len(textNode.Children) != 0 {
		return NativeTextRun{}, fmt.Errorf("pptxpatch: native extract: text run lacks one exact a:t")
	}
	if err := requireOnlyNativeAttrs(textNode, xml.Name{Space: "http://www.w3.org/XML/1998/namespace", Local: "space"}); err != nil {
		return NativeTextRun{}, err
	}
	spaceMode, hasSpaceMode := exactNativeAttr(textNode, "http://www.w3.org/XML/1998/namespace", "space")
	if hasSpaceMode && spaceMode != "default" && spaceMode != "preserve" {
		return NativeTextRun{}, fmt.Errorf("pptxpatch: native extract: invalid xml:space %q", spaceMode)
	}
	text := textNode.Text
	if strings.ContainsAny(text, "\t\r\n") {
		return NativeTextRun{}, unsupportedNativeTextContent("literal tab or line-break characters require modeled DrawingML tab/break semantics")
	}
	if hasNativeEdgeXMLSpace(text) && (!hasSpaceMode || spaceMode != "preserve") {
		return NativeTextRun{}, fmt.Errorf("pptxpatch: native extract: significant edge whitespace requires xml:space=preserve")
	}
	run := NativeTextRun{Text: &text}
	rPr, err := nativeSingleton(node, dialect.drawing, "rPr", true)
	if err != nil {
		return NativeTextRun{}, err
	}
	if err := requireOnlyNativeAttrs(rPr, xml.Name{Local: "b"}, xml.Name{Local: "i"}, xml.Name{Local: "sz"}, xml.Name{Local: "lang"}); err != nil {
		return NativeTextRun{}, fmt.Errorf("pptxpatch: native extract: unmodeled run metadata: %w", err)
	}
	if err := requireOnlyNativeChildren(rPr,
		xml.Name{Space: dialect.drawing, Local: "latin"},
		xml.Name{Space: dialect.drawing, Local: "ea"},
		xml.Name{Space: dialect.drawing, Local: "cs"},
		xml.Name{Space: dialect.drawing, Local: "solidFill"}); err != nil {
		return NativeTextRun{}, fmt.Errorf("pptxpatch: native extract: unmodeled run metadata: %w", err)
	}
	if _, err := nativeSingleton(rPr, dialect.drawing, "latin", false); err != nil {
		return NativeTextRun{}, err
	}
	if _, err := nativeSingleton(rPr, dialect.drawing, "ea", false); err != nil {
		return NativeTextRun{}, err
	}
	if _, err := nativeSingleton(rPr, dialect.drawing, "cs", false); err != nil {
		return NativeTextRun{}, err
	}
	if _, err := nativeSingleton(rPr, dialect.drawing, "solidFill", false); err != nil {
		return NativeTextRun{}, err
	}
	if value, ok := exactNativeAttr(rPr, "", "b"); ok {
		parsed, err := nativeBool(value)
		if err != nil {
			return NativeTextRun{}, err
		}
		run.Bold = &parsed
	}
	if value, ok := exactNativeAttr(rPr, "", "i"); ok {
		parsed, err := nativeBool(value)
		if err != nil {
			return NativeTextRun{}, err
		}
		run.Italic = &parsed
	}
	if value, ok := exactNativeAttr(rPr, "", "lang"); ok {
		if !validNativeLanguage(value) {
			return NativeTextRun{}, fmt.Errorf("invalid authored language tag")
		}
		run.Language = &value
	}
	if value, ok := exactNativeAttr(rPr, "", "sz"); ok {
		size, err := strconv.ParseInt(value, 10, 64)
		if err != nil || size <= 0 {
			return NativeTextRun{}, fmt.Errorf("pptxpatch: native extract: invalid run size")
		}
		run.FontSizeHundredthPt = &size
	}
	var latinFamily, eaFamily, csFamily string
	for _, child := range rPr.Children {
		switch child.Name {
		case xml.Name{Space: dialect.drawing, Local: "latin"}:
			family, familyErr := extractor.exactNativeRunTypeface(child)
			if familyErr != nil {
				return NativeTextRun{}, familyErr
			}
			latinFamily = family
			run.FontFamily = &family
		case xml.Name{Space: dialect.drawing, Local: "ea"}:
			family, familyErr := extractor.exactNativeRunTypeface(child)
			if familyErr != nil {
				return NativeTextRun{}, familyErr
			}
			eaFamily = family
		case xml.Name{Space: dialect.drawing, Local: "cs"}:
			family, familyErr := extractor.exactNativeRunTypeface(child)
			if familyErr != nil {
				return NativeTextRun{}, familyErr
			}
			csFamily = family
		case xml.Name{Space: dialect.drawing, Local: "solidFill"}:
			color, colorErr := exactNativeSolidColor(child, dialect, extractor.theme)
			if colorErr != nil {
				return NativeTextRun{}, unsupportedNativeTextContent("run color is not an exact sRGB or documented theme color transform")
			}
			run.Color = &color
		default:
			return NativeTextRun{}, fmt.Errorf("pptxpatch: native extract: unsupported run property %s", child.Name.Local)
		}
	}
	if latinFamily != "" {
		if (eaFamily != "" && eaFamily != latinFamily) || (csFamily != "" && csFamily != latinFamily) {
			return NativeTextRun{}, unsupportedNativeTextContent("script-specific run typefaces require modeled font slots")
		}
	}
	if run.Bold == nil || run.Italic == nil || run.FontSizeHundredthPt == nil || run.Color == nil || run.FontFamily == nil {
		return NativeTextRun{}, fmt.Errorf("pptxpatch: native extract: text run formatting is not self-contained")
	}
	return run, nil
}

func (extractor *nativeExtractor) exactNativeRunTypeface(node *nativeXMLNode) (string, error) {
	if requireOnlyNativeAttrs(node, xml.Name{Local: "typeface"}) != nil || requireOnlyNativeChildren(node) != nil {
		return "", fmt.Errorf("pptxpatch: native extract: invalid run typeface")
	}
	family, ok := exactNativeAttr(node, "", "typeface")
	if !ok || family == "" {
		return "", fmt.Errorf("pptxpatch: native extract: invalid run typeface")
	}
	resolved, err := extractor.theme.resolveTypeface(family)
	if err != nil {
		return "", unsupportedNativeTextContent("theme font token cannot be resolved to an exact typeface")
	}
	return resolved, nil
}

func hasNativeEdgeXMLSpace(value string) bool {
	if value == "" {
		return false
	}
	isSpace := func(value byte) bool {
		switch value {
		case 0x20, 0x09, 0x0a, 0x0d:
			return true
		default:
			return false
		}
	}
	return isSpace(value[0]) || isSpace(value[len(value)-1])
}

func (extractor *nativeExtractor) markSlideUnsupported(slide *NativeSlide, part, objectID, fingerprint string, payload []byte, code, message string) error {
	if len(slide.Passthrough) >= nativeMaxPassthroughPerObject || extractor.passthroughRefsEmitted >= nativeExtractMaxEmittedPassthrough {
		return fmt.Errorf("pptxpatch: native extract: passthrough reference budget exceeded before capability issuance")
	}
	passthrough, err := extractor.issuePassthrough(part, objectID, fingerprint, payload, code)
	if err != nil {
		return err
	}
	foundPassthrough := false
	for _, existing := range slide.Passthrough {
		if existing.Token == passthrough.Token {
			foundPassthrough = true
			break
		}
	}
	if !foundPassthrough {
		if err := extractor.reserveNativePassthroughReference(); err != nil {
			return err
		}
		slide.Passthrough = append(slide.Passthrough, passthrough)
	}
	slide.Compatibility.Status = worseNativeStatus(slide.Compatibility.Status, NativeCompatibilityStatusPreserveOnly)
	slideID := slide.ID
	partName := part
	for _, diagnostic := range slide.Compatibility.Diagnostics {
		if diagnostic.Code == code && diagnostic.Scope != nil && diagnostic.Scope.PartName != nil && *diagnostic.Scope.PartName == part {
			return nil
		}
	}
	if len(slide.Compatibility.Diagnostics) >= nativeMaxDiagnosticsPerScope {
		return fmt.Errorf("pptxpatch: native extract: slide diagnostic budget exceeded")
	}
	slide.Compatibility.Diagnostics = append(slide.Compatibility.Diagnostics, NativeDiagnostic{
		Severity: NativeDiagnosticSeverityWarning, Code: code, Message: message,
		Scope: &NativeDiagnosticScope{SlideID: &slideID, PartName: &partName},
	})
	return nil
}

func (extractor *nativeExtractor) reserveNativePassthroughReference() error {
	if extractor.passthroughRefsEmitted >= nativeExtractMaxEmittedPassthrough {
		return fmt.Errorf("pptxpatch: native extract: cumulative emitted passthrough reference budget exceeded")
	}
	extractor.passthroughRefsEmitted++
	return nil
}

func (extractor *nativeExtractor) reserveNativeTextOutput(txBody *nativeXMLNode, dialect nativeExtractDialect) error {
	paragraphs := nativeChildren(txBody, dialect.drawing, "p")
	if len(paragraphs) > nativeMaxParagraphsPerElement {
		return fmt.Errorf("pptxpatch: native extract: paragraph budget exceeded")
	}
	nodes := len(paragraphs)
	var textUnits int64
	for _, paragraph := range paragraphs {
		runs := nativeChildren(paragraph, dialect.drawing, "r")
		if len(runs) > nativeMaxRunsPerParagraph {
			return fmt.Errorf("pptxpatch: native extract: run budget exceeded")
		}
		nodes += len(runs)
		for _, run := range runs {
			textNode, err := nativeSingleton(run, dialect.drawing, "t", true)
			if err != nil {
				return err
			}
			units := int64(utf16CodeUnitLengthBounded(textNode.Text, nativeMaxTextCodeUnits+1))
			if units > nativeMaxTextCodeUnits {
				return fmt.Errorf("pptxpatch: native extract: run text budget exceeded")
			}
			textUnits += units
		}
	}
	if nodes > nativeMaxNodes-extractor.outputNodesEmitted {
		return fmt.Errorf("pptxpatch: native extract: output node budget exceeded")
	}
	if textUnits > int64(nativeMaxTotalTextCodeUnits)-extractor.textCodeUnitsEmitted {
		return fmt.Errorf("pptxpatch: native extract: cumulative text budget exceeded")
	}
	extractor.outputNodesEmitted += nodes
	extractor.textCodeUnitsEmitted += textUnits
	return nil
}

func (extractor *nativeExtractor) issuePassthrough(part, objectID, fingerprint string, payload []byte, reason string) (NativePassthroughRef, error) {
	if extractor.options.TokenFactory == nil {
		return NativePassthroughRef{}, fmt.Errorf("pptxpatch: native extract: unsupported content requires a trusted passthrough token factory")
	}
	cacheKey := part + "\x00" + objectID + "\x00" + fingerprint + "\x00" + reason
	if cached, ok := extractor.passthroughCache[cacheKey]; ok {
		return cached, nil
	}
	extractor.passthroughUsed++
	if extractor.passthroughUsed > nativeExtractMaxPassthrough {
		return NativePassthroughRef{}, fmt.Errorf("pptxpatch: native extract: passthrough budget exceeded")
	}
	token, err := extractor.options.TokenFactory.IssueNativePassthroughToken(NativePassthroughTokenRequest{
		SourceRevision: extractor.sourceRevision, OwnerPart: part, ObjectID: objectID,
		FingerprintSHA256: fingerprint, ByteLength: int64(len(payload)), Reason: reason, Payload: payload,
	})
	if err != nil || !nativeIDPattern.MatchString(token) {
		return NativePassthroughRef{}, fmt.Errorf("pptxpatch: native extract: token factory returned invalid capability")
	}
	if previousOwner, exists := extractor.tokenOwners[token]; exists && previousOwner != cacheKey {
		return NativePassthroughRef{}, fmt.Errorf("pptxpatch: native extract: token factory reused one capability for distinct source objects")
	}
	if _, exists := extractor.tokenOwners[token]; !exists {
		extractor.tokenOwnerJournal = append(extractor.tokenOwnerJournal, token)
	}
	extractor.tokenOwners[token] = cacheKey
	result := NativePassthroughRef{Token: token, OwnerPart: part, FingerprintSHA256: fingerprint, Disposition: NativePassthroughDispositionPreserve}
	if _, exists := extractor.passthroughCache[cacheKey]; !exists {
		extractor.passthroughCacheJournal = append(extractor.passthroughCacheJournal, cacheKey)
	}
	extractor.passthroughCache[cacheKey] = result
	return result, nil
}

func requiredNativeInt64(node *nativeXMLNode, space, local string) (int64, error) {
	if node == nil {
		return 0, fmt.Errorf("pptxpatch: native extract: missing %s", local)
	}
	value, ok := exactNativeAttr(node, space, local)
	if !ok {
		return 0, fmt.Errorf("pptxpatch: native extract: missing %s", local)
	}
	parsed, err := parseCanonicalNativeInt(value, -nativeMaxSafeInteger, nativeMaxSafeInteger)
	if err != nil {
		return 0, fmt.Errorf("pptxpatch: native extract: invalid integer %s", local)
	}
	return parsed, nil
}

func requiredNativePositiveInt64(node *nativeXMLNode, space, local string) (int64, error) {
	value, err := requiredNativeInt64(node, space, local)
	if err != nil || value <= 0 {
		return 0, fmt.Errorf("pptxpatch: native extract: %s must be positive", local)
	}
	return value, nil
}

func canonicalNativeUnsignedID(node *nativeXMLNode, space, local string, minimum uint64) (string, error) {
	value, ok := exactNativeAttr(node, space, local)
	if !ok || value == "" || strings.HasPrefix(value, "+") {
		return "", fmt.Errorf("pptxpatch: native extract: missing canonical unsigned %s", local)
	}
	parsed, err := strconv.ParseUint(value, 10, 32)
	if err != nil || parsed < minimum || strconv.FormatUint(parsed, 10) != value {
		return "", fmt.Errorf("pptxpatch: native extract: invalid canonical unsigned %s", local)
	}
	return value, nil
}

func nativeShapeObjectID(node *nativeXMLNode, dialect nativeExtractDialect) (string, error) {
	return nativeObjectIDFromContainer(node, dialect, "nvSpPr")
}

func nativeObjectIDFromContainer(node *nativeXMLNode, dialect nativeExtractDialect, containerName string) (string, error) {
	nonVisual, err := nativeSingleton(node, dialect.presentation, containerName, true)
	if err != nil {
		return "", err
	}
	cNvPr, err := nativeSingleton(nonVisual, dialect.presentation, "cNvPr", true)
	if err != nil {
		return "", err
	}
	value, err := canonicalNativeUnsignedID(cNvPr, "", "id", 1)
	if err != nil {
		return "", err
	}
	return "cNvPr-" + value, nil
}

func collectNativeShapeTreeObjectIDs(spTree *nativeXMLNode, dialect nativeExtractDialect) (map[*nativeXMLNode]string, error) {
	result := map[*nativeXMLNode]string{}
	seen := map[string]bool{}
	rootID, err := nativeObjectIDFromContainer(spTree, dialect, "nvGrpSpPr")
	if err != nil {
		return nil, fmt.Errorf("pptxpatch: native extract: invalid root group cNvPr id: %w", err)
	}
	seen[rootID] = true
	var walk func(*nativeXMLNode) error
	walk = func(container *nativeXMLNode) error {
		for _, child := range container.Children {
			var nonVisual string
			switch child.Name {
			case xml.Name{Space: dialect.presentation, Local: "sp"}:
				nonVisual = "nvSpPr"
			case xml.Name{Space: dialect.presentation, Local: "pic"}:
				nonVisual = "nvPicPr"
			case xml.Name{Space: dialect.presentation, Local: "graphicFrame"}:
				nonVisual = "nvGraphicFramePr"
			case xml.Name{Space: dialect.presentation, Local: "cxnSp"}:
				nonVisual = "nvCxnSpPr"
			case xml.Name{Space: dialect.presentation, Local: "grpSp"}:
				nonVisual = "nvGrpSpPr"
			default:
				continue
			}
			objectID, err := nativeObjectIDFromContainer(child, dialect, nonVisual)
			if err != nil {
				return fmt.Errorf("pptxpatch: native extract: invalid shape-tree cNvPr id: %w", err)
			}
			if seen[objectID] {
				return fmt.Errorf("pptxpatch: native extract: duplicate cNvPr id %s in shape tree", objectID)
			}
			if len(seen) >= nativeMaxTotalElements {
				return fmt.Errorf("pptxpatch: native extract: shape-tree object budget exceeded")
			}
			seen[objectID] = true
			result[child] = objectID
			if child.Name == (xml.Name{Space: dialect.presentation, Local: "grpSp"}) {
				if err := walk(child); err != nil {
					return err
				}
			}
		}
		return nil
	}
	if err := walk(spTree); err != nil {
		return nil, err
	}
	return result, nil
}

func validateNativeRootGroupScaffold(spTree *nativeXMLNode, dialect nativeExtractDialect) (string, error) {
	if err := requireOnlyNativeAttrs(spTree); err != nil {
		return "", fmt.Errorf("pptxpatch: native extract: spTree: %w", err)
	}
	nonVisual, err := nativeSingleton(spTree, dialect.presentation, "nvGrpSpPr", true)
	if err != nil {
		return "", err
	}
	if _, err := nativeSingleton(spTree, dialect.presentation, "grpSpPr", true); err != nil {
		return "", err
	}
	if err := requireOnlyNativeAttrs(nonVisual); err != nil {
		return "", err
	}
	if err := requireOnlyNativeChildren(nonVisual,
		xml.Name{Space: dialect.presentation, Local: "cNvPr"},
		xml.Name{Space: dialect.presentation, Local: "cNvGrpSpPr"},
		xml.Name{Space: dialect.presentation, Local: "nvPr"}); err != nil {
		return "", err
	}
	cNvPr, err := nativeSingleton(nonVisual, dialect.presentation, "cNvPr", true)
	if err != nil {
		return "", err
	}
	cNvGrpSpPr, err := nativeSingleton(nonVisual, dialect.presentation, "cNvGrpSpPr", true)
	if err != nil {
		return "", err
	}
	nvPr, err := nativeSingleton(nonVisual, dialect.presentation, "nvPr", true)
	if err != nil {
		return "", err
	}
	if err := requireOnlyNativeAttrs(cNvPr, xml.Name{Local: "id"}, xml.Name{Local: "name"}); err != nil {
		return "", err
	}
	if err := requireOnlyNativeChildren(cNvPr); err != nil {
		return "", err
	}
	if err := requireEmptyNativeElement(cNvGrpSpPr); err != nil {
		return "", err
	}
	if err := requireEmptyNativeElement(nvPr); err != nil {
		return "", err
	}
	return nativeObjectIDFromContainer(spTree, dialect, "nvGrpSpPr")
}

func nativeBool(value string) (bool, error) {
	switch value {
	case "1", "true":
		return true, nil
	case "0", "false":
		return false, nil
	default:
		return false, fmt.Errorf("pptxpatch: native extract: invalid boolean %q", value)
	}
}

func nativeTextAlign(value string) (NativeTextAlign, error) {
	switch value {
	case "l":
		return NativeTextAlignLeft, nil
	case "ctr":
		return NativeTextAlignCenter, nil
	case "r":
		return NativeTextAlignRight, nil
	default:
		return "", fmt.Errorf("pptxpatch: native extract: unsupported paragraph alignment %q", value)
	}
}

func nativeSHA256(data []byte) string {
	digest := sha256.Sum256(data)
	return hex.EncodeToString(digest[:])
}

func stableNativeID(prefix, part, objectID string) string {
	digest := sha256.Sum256([]byte(part + "\x00" + objectID))
	return prefix + "-" + hex.EncodeToString(digest[:12])
}

func nativeIdentityKey(part, objectID string) string {
	alias, err := nativePartAlias(part)
	if err != nil {
		alias = part
	}
	return alias + "\x00" + objectID
}

func int64Pointer(value int64) *int64    { return &value }
func stringPointer(value string) *string { return &value }
func boolPointer(value bool) *bool       { return &value }
