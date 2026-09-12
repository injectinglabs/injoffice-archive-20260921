package docxpatch

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/xml"
	"fmt"
	"io"
	"net/url"
	"path"
	"sort"
	"strconv"
	"strings"
	"unicode"
	"unicode/utf8"
)

const (
	NativeDOCXMaxPackageBytes       = 128 * 1024 * 1024
	NativeDOCXMaxPartBytes          = 64 * 1024 * 1024
	NativeDOCXMaxXMLPartBytes       = 16 * 1024 * 1024
	NativeDOCXMaxUncompressedBytes  = 256 * 1024 * 1024
	NativeDOCXMaxPackageParts       = 10_000
	NativeDOCXMaxCompressionRatio   = 200
	NativeDOCXCompressionRatioSlack = 1 * 1024 * 1024
)

const (
	opcContentTypesNS  = "http://schemas.openxmlformats.org/package/2006/content-types"
	opcRelationshipsNS = "http://schemas.openxmlformats.org/package/2006/relationships"

	wordMLTransitional      = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	wordMLStrict            = "http://purl.oclc.org/ooxml/wordprocessingml/main"
	relNSTransitional       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	relNSStrict             = "http://purl.oclc.org/ooxml/officeDocument/relationships"
	relBaseTransitional     = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
	relBaseStrict           = "http://purl.oclc.org/ooxml/officeDocument/relationships/"
	wordDrawingTransitional = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
	wordDrawingStrict       = "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing"
	drawingMLTransitional   = "http://schemas.openxmlformats.org/drawingml/2006/main"
	drawingMLStrict         = "http://purl.oclc.org/ooxml/drawingml/main"
	pictureMLTransitional   = "http://schemas.openxmlformats.org/drawingml/2006/picture"
	pictureMLStrict         = "http://purl.oclc.org/ooxml/drawingml/picture"
)

type nativePackage struct {
	raw          []byte
	files        map[string][]byte
	partByKey    map[string]string
	contentTypes map[string]string
	rels         map[string][]nativeRelationship
	relsPart     map[string]string
}

type nativeRelationship struct {
	ID       string
	Type     string
	Target   string
	External bool
	PartName string
}

type nativeXMLNode struct {
	Name     xml.Name
	Attrs    []xml.Attr
	Start    int64
	End      int64
	Path     string
	Text     string
	Children []*nativeXMLNode
	parent   *nativeXMLNode
}

type nativeExtractor struct {
	pkg                 *nativePackage
	wordNS              string
	relNS               string
	mainPart            string
	mainRoot            *nativeXMLNode
	bodyNode            *nativeXMLNode
	bodyID              string
	modeledParts        map[string]bool
	unsupported         []NativeUnsupportedCapabilityV1
	unsupportedSet      map[string]bool
	storyByRel          map[string]string
	storyByNative       map[string]string
	activeNoteKind      string
	activeNoteStory     string
	activeNoteRole      string
	commentByNative     map[string]string
	headers             []NativeStoryV1
	footers             []NativeStoryV1
	notes               []NativeStoryV1
	commentStories      []NativeStoryV1
	comments            []NativeCommentV1
	previous            *NativeDocumentV1
	previousIDs         map[string][]string
	previousUsed        map[string]int
	previousNative      map[string]string
	previousPath        map[string]string
	reservedIDs         map[string]bool
	allocatedIDs        map[string]bool
	retainPathIdentity  bool
	unsupportedOverflow bool
	seenParaIDs         map[string]string
	themeSrgbColors     map[string]string
	themeSrgbLoaded     bool
	tableLookLoaded     bool
	tableLookStyles     map[string]*nativeXMLNode
	noEndnotesChecked   bool
	noEndnotesProven    bool
}

// NativeExtractionOptions lets a caller retain durable identity across source
// revisions. Previous must be a valid v1 contract for the same main part;
// unchanged anchored objects are matched by type, part, and XML fingerprint.
// DocumentID can establish an application-owned durable identity on first
// import. When omitted, the extractor derives an initial ID from the main XML
// only, so unrelated package-part changes never churn document identity.
type NativeExtractionOptions struct {
	Previous   *NativeDocumentV1
	DocumentID string
	// RetainPathIdentity may be set only after the caller independently proves
	// that native topology is unchanged. Content-changed objects at the same
	// kind, canonical part, and XML path then retain their previous IDs after
	// fingerprint matching has had first priority.
	RetainPathIdentity bool
}

// ExtractNativeDocumentV1 parses a DOCX directly from its OPC and
// WordprocessingML bytes. It never renders HTML and never regenerates the
// package. The result is validated against the versioned native contract and
// carries byte-exact anchors for future surgical mutations.
func ExtractNativeDocumentV1(data []byte) (*NativeDocumentV1, error) {
	return ExtractNativeDocumentV1WithOptions(data, NativeExtractionOptions{})
}

// ExtractNativeDocumentV1WithOptions is ExtractNativeDocumentV1 with explicit
// identity continuity for incremental imports.
func ExtractNativeDocumentV1WithOptions(data []byte, options NativeExtractionOptions) (*NativeDocumentV1, error) {
	pkg, err := openNativeDOCXPackage(data)
	if err != nil {
		return nil, err
	}
	mainPart, strict, err := pkg.officeDocumentPart()
	if err != nil {
		return nil, err
	}
	mainXML := pkg.files[mainPart]
	mainRoot, err := parseNativeXML(mainPart, mainXML)
	if err != nil {
		return nil, err
	}
	wordNS, relNS := wordMLTransitional, relNSTransitional
	if strict {
		wordNS, relNS = wordMLStrict, relNSStrict
	}
	if mainRoot.Name != (xml.Name{Space: wordNS, Local: "document"}) {
		return nil, fmt.Errorf("docxpatch: native extract: main part %q has root {%s}%s; expected {%s}document", mainPart, mainRoot.Name.Space, mainRoot.Name.Local, wordNS)
	}
	if err := rejectNativeNamespaceSpoofing(mainRoot, wordNS); err != nil {
		return nil, fmt.Errorf("docxpatch: native extract: main part %q: %w", mainPart, err)
	}
	body := firstDirectNativeChild(mainRoot, wordNS, "body")
	if body == nil {
		return nil, fmt.Errorf("docxpatch: native extract: main part %q has no w:body", mainPart)
	}

	if options.Previous != nil {
		if issues := ValidateNativeDocumentV1(options.Previous); len(issues) > 0 {
			return nil, fmt.Errorf("docxpatch: native extract: previous identity contract is invalid: %w", &NativeValidationError{Issues: issues})
		}
		previousMainKey, previousKeyErr := nativeDecodedPartKey(options.Previous.Source.MainPart)
		mainKey, mainKeyErr := nativeDecodedPartKey(mainPart)
		if previousKeyErr != nil || mainKeyErr != nil || previousMainKey != mainKey {
			return nil, fmt.Errorf("docxpatch: native extract: previous identity main part %q does not match %q", options.Previous.Source.MainPart, mainPart)
		}
		if options.DocumentID != "" && options.DocumentID != options.Previous.DocumentID {
			return nil, fmt.Errorf("docxpatch: native extract: explicit document identity conflicts with the previous contract")
		}
	}
	if options.DocumentID != "" && !nativeIDPattern.MatchString(options.DocumentID) {
		return nil, fmt.Errorf("docxpatch: native extract: invalid explicit document identity %q", options.DocumentID)
	}
	extractor := &nativeExtractor{
		pkg: pkg, wordNS: wordNS, relNS: relNS, mainPart: mainPart, mainRoot: mainRoot, bodyNode: body,
		modeledParts:   map[string]bool{mainPart: true},
		unsupportedSet: map[string]bool{}, storyByRel: map[string]string{}, storyByNative: map[string]string{}, commentByNative: map[string]string{},
		headers: []NativeStoryV1{}, footers: []NativeStoryV1{}, notes: []NativeStoryV1{}, commentStories: []NativeStoryV1{}, comments: []NativeCommentV1{}, unsupported: []NativeUnsupportedCapabilityV1{},
		previous: options.Previous, previousIDs: map[string][]string{}, previousUsed: map[string]int{}, previousNative: map[string]string{}, previousPath: map[string]string{}, reservedIDs: map[string]bool{}, allocatedIDs: map[string]bool{}, retainPathIdentity: options.RetainPathIdentity, seenParaIDs: map[string]string{},
	}
	extractor.indexPreviousIDs()
	if options.Previous != nil {
		extractor.bodyID = options.Previous.Body.ID
	} else {
		extractor.bodyID = nativeStableID("story", mainPart, "", "body:"+nativeSHA(mainXML))
	}
	if err := extractor.extractRelatedStories(); err != nil {
		return nil, err
	}
	bodyStory, sections, err := extractor.extractBody()
	if err != nil {
		return nil, err
	}
	if extractor.unsupportedOverflow {
		return nil, fmt.Errorf("docxpatch: native extract: unsupported construct inventory exceeds %d entries", NativeDOCXMaxCollectionItems)
	}
	packageDigest := nativeSHA(data)
	version := NativeDOCXVersion
	documentID := options.DocumentID
	if documentID == "" && options.Previous != nil {
		documentID = options.Previous.DocumentID
	}
	if documentID == "" {
		documentID = nativeStableID("document", mainPart, nativeSHA(mainXML), "")
	}
	doc := &NativeDocumentV1{
		Protocol: NativeDOCXProtocol, Version: &version,
		DocumentID: documentID,
		Revision:   "rev:" + strings.TrimPrefix(packageDigest, "sha256:")[:32],
		Source:     NativeSourcePackageV1{PackageSHA256: packageDigest, MainPart: mainPart},
		Body:       bodyStory, Sections: sections, Headers: extractor.headers, Footers: extractor.footers,
		Notes: extractor.notes, CommentStories: extractor.commentStories, Comments: extractor.comments,
		Capabilities: []NativeCapabilityV1{
			{Name: "native-ooxml-parse", Level: "read-only", Detail: nativeString("HTML-free, part-qualified WordprocessingML extraction")},
			{Name: "surgical-byte-anchors", Level: "read-write", Detail: nativeString("Exact XML byte ranges and fingerprints guard atomic native text mutations")},
			{Name: "native-text-mutation", Level: "read-write", Detail: nativeString("Guarded paragraph and text-run replacement with post-write native validation")},
			{Name: "full-document-regeneration", Level: "unsupported", Detail: nativeString("Unmodeled OOXML is preserved verbatim and must not be flattened")},
		},
		PassthroughParts: extractor.passthroughParts(), Unsupported: extractor.unsupported,
	}
	if issues := ValidateNativeDocumentV1(doc); len(issues) > 0 {
		return nil, fmt.Errorf("docxpatch: native extract produced invalid contract: %w", &NativeValidationError{Issues: issues})
	}
	if _, err := EncodeNativeDocumentV1(doc); err != nil {
		return nil, fmt.Errorf("docxpatch: native extract output exceeds contract limits: %w", err)
	}
	return doc, nil
}

// ExtractNativeDocument is the concise API alias for the v1 extractor. The
// returned contract still carries its explicit protocol and version.
func ExtractNativeDocument(data []byte) (*NativeDocumentV1, error) {
	return ExtractNativeDocumentV1(data)
}

func openNativeDOCXPackage(data []byte) (*nativePackage, error) {
	if len(data) == 0 || len(data) > NativeDOCXMaxPackageBytes {
		return nil, fmt.Errorf("docxpatch: native extract: package size must be 1..%d bytes", NativeDOCXMaxPackageBytes)
	}
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, fmt.Errorf("docxpatch: native extract: invalid ZIP: %w", err)
	}
	if len(zr.File) == 0 || len(zr.File) > NativeDOCXMaxPackageParts {
		return nil, fmt.Errorf("docxpatch: native extract: ZIP contains %d entries; limit is %d", len(zr.File), NativeDOCXMaxPackageParts)
	}
	files := make(map[string][]byte, len(zr.File))
	partByKey := make(map[string]string, len(zr.File))
	aliases := make(map[string]string, len(zr.File))
	var declaredTotal uint64
	var actualTotal uint64
	for _, file := range zr.File {
		if file.FileInfo().IsDir() {
			if err := validateNativeDirectoryName(file.Name); err != nil {
				return nil, fmt.Errorf("docxpatch: native extract: ZIP directory %q: %w", file.Name, err)
			}
			continue
		}
		if file.Name != "[Content_Types].xml" {
			if err := validateNativePartName(file.Name); err != nil {
				return nil, fmt.Errorf("docxpatch: native extract: ZIP entry %q: %w", file.Name, err)
			}
		}
		if _, duplicate := files[file.Name]; duplicate {
			return nil, fmt.Errorf("docxpatch: native extract: duplicate ZIP entry %q", file.Name)
		}
		alias, err := nativeDecodedPartKey(file.Name)
		if err != nil {
			return nil, fmt.Errorf("docxpatch: native extract: ZIP entry %q: %w", file.Name, err)
		}
		if original, duplicate := aliases[alias]; duplicate {
			return nil, fmt.Errorf("docxpatch: native extract: ambiguous ZIP entries %q and %q", original, file.Name)
		}
		aliases[alias] = file.Name
		partByKey[alias] = file.Name
		if file.Flags&0x1 != 0 {
			return nil, fmt.Errorf("docxpatch: native extract: encrypted ZIP entry %q is unsupported", file.Name)
		}
		if file.Method != zip.Store && file.Method != zip.Deflate {
			return nil, fmt.Errorf("docxpatch: native extract: ZIP entry %q uses unsupported compression method %d", file.Name, file.Method)
		}
		if file.UncompressedSize64 > NativeDOCXMaxPartBytes {
			return nil, fmt.Errorf("docxpatch: native extract: ZIP entry %q exceeds %d uncompressed bytes", file.Name, NativeDOCXMaxPartBytes)
		}
		if file.UncompressedSize64 > file.CompressedSize64*NativeDOCXMaxCompressionRatio+NativeDOCXCompressionRatioSlack {
			return nil, fmt.Errorf("docxpatch: native extract: ZIP entry %q exceeds compression-ratio limit", file.Name)
		}
		declaredTotal += file.UncompressedSize64
		if declaredTotal > NativeDOCXMaxUncompressedBytes {
			return nil, fmt.Errorf("docxpatch: native extract: ZIP exceeds %d total uncompressed bytes", NativeDOCXMaxUncompressedBytes)
		}
		rc, err := file.Open()
		if err != nil {
			return nil, fmt.Errorf("docxpatch: native extract: open ZIP entry %q: %w", file.Name, err)
		}
		part, readErr := io.ReadAll(io.LimitReader(rc, NativeDOCXMaxPartBytes+1))
		closeErr := rc.Close()
		if readErr != nil {
			return nil, fmt.Errorf("docxpatch: native extract: read ZIP entry %q: %w", file.Name, readErr)
		}
		if closeErr != nil {
			return nil, fmt.Errorf("docxpatch: native extract: close ZIP entry %q: %w", file.Name, closeErr)
		}
		if len(part) > NativeDOCXMaxPartBytes {
			return nil, fmt.Errorf("docxpatch: native extract: ZIP entry %q exceeds %d bytes", file.Name, NativeDOCXMaxPartBytes)
		}
		actualTotal += uint64(len(part))
		if actualTotal > NativeDOCXMaxUncompressedBytes {
			return nil, fmt.Errorf("docxpatch: native extract: ZIP exceeds %d actual uncompressed bytes", NativeDOCXMaxUncompressedBytes)
		}
		files[file.Name] = part
	}
	contentXML, ok := files["[Content_Types].xml"]
	if !ok {
		return nil, fmt.Errorf("docxpatch: native extract: [Content_Types].xml is missing")
	}
	contentTypes, err := parseNativeContentTypes(contentXML, files, partByKey)
	if err != nil {
		return nil, err
	}
	pkg := &nativePackage{raw: data, files: files, partByKey: partByKey, contentTypes: contentTypes, rels: map[string][]nativeRelationship{}, relsPart: map[string]string{}}
	if err := pkg.loadRelationships(); err != nil {
		return nil, err
	}
	return pkg, nil
}

func validateNativeDirectoryName(name string) error {
	if !strings.HasSuffix(name, "/") || name == "/" {
		return fmt.Errorf("invalid directory marker")
	}
	for _, segment := range strings.Split(strings.TrimSuffix(name, "/"), "/") {
		if err := validateNativePartSegment(segment); err != nil {
			return err
		}
	}
	return nil
}

func validateNativePartName(name string) error {
	if name == "" || len(name) > 4096 || strings.HasPrefix(name, "/") || strings.Contains(name, "\\") || strings.Contains(name, "//") || strings.HasSuffix(name, "/") {
		return fmt.Errorf("invalid OPC part name")
	}
	for _, segment := range strings.Split(name, "/") {
		if err := validateNativePartSegment(segment); err != nil {
			return err
		}
	}
	return nil
}

func validateNativePartSegment(segment string) error {
	if segment == "" || segment == "." || segment == ".." || !nativePartSegment.MatchString(segment) {
		return fmt.Errorf("non-canonical OPC part segment %q", segment)
	}
	decoded, err := url.PathUnescape(segment)
	if err != nil || !utf8.ValidString(decoded) || decoded == "." || decoded == ".." || strings.HasSuffix(decoded, ".") || strings.ContainsAny(decoded, `\/?#%`) {
		return fmt.Errorf("unsafe percent-encoded OPC part segment %q", segment)
	}
	for _, character := range decoded {
		if character < 0x20 || character == 0x7f {
			return fmt.Errorf("encoded control character in OPC part segment")
		}
	}
	return nil
}

func nativeDecodedPartKey(name string) (string, error) {
	if name == "[Content_Types].xml" {
		return name, nil
	}
	parts := strings.Split(name, "/")
	for index, segment := range parts {
		decoded, err := url.PathUnescape(segment)
		if err != nil {
			return "", err
		}
		parts[index] = decoded
	}
	return nativeASCIIFold(strings.Join(parts, "/")), nil
}

func nativeASCIIFold(value string) string {
	var folded strings.Builder
	folded.Grow(len(value))
	for index := 0; index < len(value); index++ {
		character := value[index]
		if character >= 'A' && character <= 'Z' {
			character += 'a' - 'A'
		}
		folded.WriteByte(character)
	}
	return folded.String()
}

func nativeASCIIEqual(left, right string) bool {
	return nativeASCIIFold(left) == nativeASCIIFold(right)
}

func parseNativeContentTypes(data []byte, files map[string][]byte, partByKey map[string]string) (map[string]string, error) {
	root, err := parseNativeXML("[Content_Types].xml", data)
	if err != nil {
		return nil, err
	}
	if root.Name != (xml.Name{Space: opcContentTypesNS, Local: "Types"}) {
		return nil, fmt.Errorf("docxpatch: native extract: [Content_Types].xml has spoofed or invalid root")
	}
	if !nativeExactContainer(root) {
		return nil, fmt.Errorf("docxpatch: native extract: [Content_Types].xml root has attributes or text outside the exact OPC subset")
	}
	defaults := map[string]string{}
	overrides := map[string]string{}
	for _, child := range root.Children {
		switch {
		case child.Name == (xml.Name{Space: opcContentTypesNS, Local: "Default"}):
			if !nativeExactLeaf(child, xml.Name{Local: "Extension"}, xml.Name{Local: "ContentType"}) {
				return nil, fmt.Errorf("docxpatch: native extract: content-type Default at %s has attributes or content outside the exact OPC subset", child.Path)
			}
			extension, okExt := nativeUnqualifiedAttr(child, "Extension")
			contentType, okType := nativeUnqualifiedAttr(child, "ContentType")
			extension = strings.ToLower(extension)
			if !okExt || !okType || extension == "" || contentType == "" || strings.ContainsAny(extension, `/\\.`) {
				return nil, fmt.Errorf("docxpatch: native extract: invalid content-type Default at %s", child.Path)
			}
			if _, duplicate := defaults[extension]; duplicate {
				return nil, fmt.Errorf("docxpatch: native extract: duplicate content-type Default for %q", extension)
			}
			defaults[extension] = contentType
		case child.Name == (xml.Name{Space: opcContentTypesNS, Local: "Override"}):
			if !nativeExactLeaf(child, xml.Name{Local: "PartName"}, xml.Name{Local: "ContentType"}) {
				return nil, fmt.Errorf("docxpatch: native extract: content-type Override at %s has attributes or content outside the exact OPC subset", child.Path)
			}
			partName, okPart := nativeUnqualifiedAttr(child, "PartName")
			contentType, okType := nativeUnqualifiedAttr(child, "ContentType")
			if !okPart || !strings.HasPrefix(partName, "/") || !okType || contentType == "" {
				return nil, fmt.Errorf("docxpatch: native extract: invalid content-type Override at %s", child.Path)
			}
			partName = strings.TrimPrefix(partName, "/")
			if err := validateNativePartName(partName); err != nil {
				return nil, fmt.Errorf("docxpatch: native extract: invalid content-type part %q: %w", partName, err)
			}
			key, err := nativeDecodedPartKey(partName)
			if err != nil {
				return nil, fmt.Errorf("docxpatch: native extract: invalid content-type part %q: %w", partName, err)
			}
			if _, duplicate := overrides[key]; duplicate {
				return nil, fmt.Errorf("docxpatch: native extract: duplicate content-type Override for %q", partName)
			}
			overrides[key] = contentType
		default:
			return nil, fmt.Errorf("docxpatch: native extract: unexpected content-type element {%s}%s", child.Name.Space, child.Name.Local)
		}
	}
	resolved := make(map[string]string, len(files))
	for name := range files {
		if name == "[Content_Types].xml" {
			continue
		}
		key, _ := nativeDecodedPartKey(name)
		if value := overrides[key]; value != "" {
			resolved[name] = value
			continue
		}
		extension := strings.TrimPrefix(strings.ToLower(path.Ext(name)), ".")
		if value := defaults[extension]; value != "" {
			resolved[name] = value
			continue
		}
		return nil, fmt.Errorf("docxpatch: native extract: part %q has no content type", name)
	}
	for key := range overrides {
		if _, exists := partByKey[key]; !exists {
			return nil, fmt.Errorf("docxpatch: native extract: content-type Override targets missing part %q", key)
		}
	}
	return resolved, nil
}

func (pkg *nativePackage) loadRelationships() error {
	for name, data := range pkg.files {
		if !nativeASCIIEqual(pkg.contentTypes[name], "application/vnd.openxmlformats-package.relationships+xml") {
			continue
		}
		owner, ok := nativeRelationshipOwner(name)
		if !ok {
			return fmt.Errorf("docxpatch: native extract: relationship part %q is not in a canonical _rels location", name)
		}
		if owner != "" {
			ownerKey, keyErr := nativeDecodedPartKey(owner)
			if keyErr != nil {
				return fmt.Errorf("docxpatch: native extract: relationship owner %q: %w", owner, keyErr)
			}
			actualOwner, exists := pkg.partByKey[ownerKey]
			if !exists {
				return fmt.Errorf("docxpatch: native extract: relationship part %q has no owner part %q", name, owner)
			}
			owner = actualOwner
		}
		if _, duplicate := pkg.rels[owner]; duplicate {
			return fmt.Errorf("docxpatch: native extract: multiple relationship parts for owner %q", owner)
		}
		rels, err := parseNativeRelationships(name, owner, data)
		if err != nil {
			return err
		}
		for index := range rels {
			if rels[index].External {
				continue
			}
			resolved, err := resolveNativeRelationshipTarget(owner, rels[index].Target)
			if err != nil {
				return fmt.Errorf("docxpatch: native extract: relationship %q in %q: %w", rels[index].ID, name, err)
			}
			key, keyErr := nativeDecodedPartKey(resolved)
			if keyErr != nil {
				return keyErr
			}
			actual, exists := pkg.partByKey[key]
			if !exists {
				return fmt.Errorf("docxpatch: native extract: relationship %q in %q targets missing part %q", rels[index].ID, name, resolved)
			}
			rels[index].PartName = actual
		}
		pkg.rels[owner] = rels
		pkg.relsPart[owner] = name
	}
	if _, ok := pkg.rels[""]; !ok {
		return fmt.Errorf("docxpatch: native extract: root relationship part _rels/.rels is missing")
	}
	return nil
}

func nativeRelationshipOwner(name string) (string, bool) {
	if nativeASCIIFold(name) == "_rels/.rels" {
		return "", true
	}
	dir, base := path.Dir(name), path.Base(name)
	if nativeASCIIFold(path.Base(dir)) != "_rels" || !strings.HasSuffix(nativeASCIIFold(base), ".rels") {
		return "", false
	}
	ownerDir := path.Dir(dir)
	if ownerDir == "." {
		ownerDir = ""
	}
	ownerBase := base[:len(base)-len(".rels")]
	owner := ownerBase
	if ownerDir != "" {
		owner = ownerDir + "/" + ownerBase
	}
	if err := validateNativePartName(owner); err != nil {
		return "", false
	}
	return owner, true
}

func parseNativeRelationships(partName, owner string, data []byte) ([]nativeRelationship, error) {
	root, err := parseNativeXML(partName, data)
	if err != nil {
		return nil, err
	}
	if root.Name != (xml.Name{Space: opcRelationshipsNS, Local: "Relationships"}) {
		return nil, fmt.Errorf("docxpatch: native extract: relationship part %q has spoofed or invalid root", partName)
	}
	if !nativeExactContainer(root) {
		return nil, fmt.Errorf("docxpatch: native extract: relationship part %q root has attributes or text outside the exact OPC subset", partName)
	}
	if len(root.Children) > NativeDOCXMaxCollectionItems {
		return nil, fmt.Errorf("docxpatch: native extract: relationship part %q exceeds %d relationships", partName, NativeDOCXMaxCollectionItems)
	}
	seen := map[string]bool{}
	rels := make([]nativeRelationship, 0, len(root.Children))
	for _, child := range root.Children {
		if child.Name != (xml.Name{Space: opcRelationshipsNS, Local: "Relationship"}) {
			return nil, fmt.Errorf("docxpatch: native extract: unexpected relationship element {%s}%s", child.Name.Space, child.Name.Local)
		}
		if !nativeExactLeaf(child, xml.Name{Local: "Id"}, xml.Name{Local: "Type"}, xml.Name{Local: "Target"}, xml.Name{Local: "TargetMode"}) {
			return nil, fmt.Errorf("docxpatch: native extract: relationship at %s has attributes or content outside the exact OPC subset", child.Path)
		}
		id, okID := nativeUnqualifiedAttr(child, "Id")
		typeURI, okType := nativeUnqualifiedAttr(child, "Type")
		target, okTarget := nativeUnqualifiedAttr(child, "Target")
		mode, hasMode := nativeUnqualifiedAttr(child, "TargetMode")
		if !okID || !nativeIDPattern.MatchString(id) || !okType || !nativeSafeAbsoluteURI(typeURI) || !okTarget || target == "" {
			return nil, fmt.Errorf("docxpatch: native extract: invalid relationship at %s", child.Path)
		}
		if seen[id] {
			return nil, fmt.Errorf("docxpatch: native extract: duplicate relationship id %q in %q", id, partName)
		}
		seen[id] = true
		if len(child.Children) != 0 || strings.TrimSpace(child.Text) != "" {
			return nil, fmt.Errorf("docxpatch: native extract: relationship %q in %q contains child content", id, partName)
		}
		external := false
		if hasMode {
			switch mode {
			case "Internal":
			case "External":
				external = true
			default:
				return nil, fmt.Errorf("docxpatch: native extract: relationship %q has invalid TargetMode %q", id, mode)
			}
		}
		if external && !nativeSafeExternalTarget(target) {
			return nil, fmt.Errorf("docxpatch: native extract: relationship %q has unsafe external target", id)
		}
		rels = append(rels, nativeRelationship{ID: id, Type: typeURI, Target: target, External: external})
	}
	return rels, nil
}

func nativeSafeAbsoluteURI(value string) bool {
	if strings.TrimSpace(value) != value || strings.ContainsAny(value, "\x00\r\n\t") || strings.IndexFunc(value, unicode.IsSpace) >= 0 {
		return false
	}
	parsed, err := url.Parse(value)
	return err == nil && parsed.IsAbs() && parsed.Scheme != ""
}

func nativeSafeExternalTarget(value string) bool {
	if value == "" || strings.TrimSpace(value) != value || strings.ContainsAny(value, "\\\x00\r\n\t") || strings.IndexFunc(value, unicode.IsSpace) >= 0 {
		return false
	}
	_, err := url.Parse(value)
	return err == nil
}

func resolveNativeRelationshipTarget(owner, target string) (string, error) {
	if target == "" || strings.HasPrefix(target, "/") || strings.ContainsAny(target, `\?#`) || strings.TrimSpace(target) != target {
		return "", fmt.Errorf("unsafe internal target %q", target)
	}
	segments := []string{}
	if owner != "" {
		directory := path.Dir(owner)
		if directory != "." {
			segments = append(segments, strings.Split(directory, "/")...)
		}
	}
	for _, segment := range strings.Split(target, "/") {
		switch segment {
		case "":
			return "", fmt.Errorf("target contains an empty path segment")
		case ".":
			continue
		case "..":
			if len(segments) == 0 {
				return "", fmt.Errorf("target escapes the package root")
			}
			segments = segments[:len(segments)-1]
		default:
			if err := validateNativePartSegment(segment); err != nil {
				return "", err
			}
			segments = append(segments, segment)
		}
	}
	resolved := strings.Join(segments, "/")
	if err := validateNativePartName(resolved); err != nil {
		return "", err
	}
	return resolved, nil
}

func (pkg *nativePackage) officeDocumentPart() (string, bool, error) {
	var main string
	strict := false
	for _, rel := range pkg.rels[""] {
		isStrict := rel.Type == relBaseStrict+"officeDocument"
		if rel.Type != relBaseTransitional+"officeDocument" && !isStrict {
			continue
		}
		if rel.External || rel.PartName == "" {
			return "", false, fmt.Errorf("docxpatch: native extract: officeDocument relationship must be internal")
		}
		if main != "" {
			return "", false, fmt.Errorf("docxpatch: native extract: multiple officeDocument relationships")
		}
		main, strict = rel.PartName, isStrict
	}
	if main == "" {
		return "", false, fmt.Errorf("docxpatch: native extract: root relationships contain no Word officeDocument")
	}
	if len(pkg.files[main]) > NativeDOCXMaxXMLPartBytes {
		return "", false, fmt.Errorf("docxpatch: native extract: main XML part %q exceeds %d bytes", main, NativeDOCXMaxXMLPartBytes)
	}
	contentType := pkg.contentTypes[main]
	if contentType != "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" &&
		contentType != "application/vnd.ms-word.document.macroEnabled.main+xml" {
		return "", false, fmt.Errorf("docxpatch: native extract: officeDocument part %q has non-Word content type %q", main, contentType)
	}
	return main, strict, nil
}

func parseNativeXML(partName string, data []byte) (*nativeXMLNode, error) {
	if len(data) == 0 || len(data) > NativeDOCXMaxXMLPartBytes {
		return nil, fmt.Errorf("docxpatch: native extract: XML part %q size must be 1..%d bytes", partName, NativeDOCXMaxXMLPartBytes)
	}
	decoder := xml.NewDecoder(bytes.NewReader(data))
	decoder.Strict = true
	var root *nativeXMLNode
	stack := []*nativeXMLNode{}
	childCounts := []map[xml.Name]int{}
	nodes := 0
	for {
		before := decoder.InputOffset()
		token, err := decoder.Token()
		after := decoder.InputOffset()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("docxpatch: native extract: parse XML part %q at byte %d: %w", partName, before, err)
		}
		switch value := token.(type) {
		case xml.StartElement:
			if value.Name.Local == "" || len(value.Name.Local) > 256 || len(value.Name.Space) > 2048 {
				return nil, fmt.Errorf("docxpatch: native extract: XML part %q contains an oversized expanded name", partName)
			}
			if len(value.Attr) > 256 {
				return nil, fmt.Errorf("docxpatch: native extract: XML part %q contains an element with more than 256 attributes", partName)
			}
			seenAttrs := map[xml.Name]bool{}
			for _, attr := range value.Attr {
				if seenAttrs[attr.Name] {
					return nil, fmt.Errorf("docxpatch: native extract: XML part %q contains duplicate attribute {%s}%s", partName, attr.Name.Space, attr.Name.Local)
				}
				seenAttrs[attr.Name] = true
			}
			nodes++
			if nodes > NativeDOCXMaxNodes {
				return nil, fmt.Errorf("docxpatch: native extract: XML part %q exceeds %d elements", partName, NativeDOCXMaxNodes)
			}
			if len(stack) >= NativeDOCXMaxDepth {
				return nil, fmt.Errorf("docxpatch: native extract: XML part %q exceeds depth %d", partName, NativeDOCXMaxDepth)
			}
			start := nativeTokenStart(data, before, after)
			node := &nativeXMLNode{Name: value.Name, Attrs: append([]xml.Attr(nil), value.Attr...), Start: start}
			if len(stack) == 0 {
				if root != nil {
					return nil, fmt.Errorf("docxpatch: native extract: XML part %q has multiple roots", partName)
				}
				root = node
				node.Path = "/" + nativeXMLQName(node.Name) + "[1]"
			} else {
				parent := stack[len(stack)-1]
				node.parent = parent
				parent.Children = append(parent.Children, node)
				counts := childCounts[len(childCounts)-1]
				counts[node.Name]++
				node.Path = fmt.Sprintf("%s/%s[%d]", parent.Path, nativeXMLQName(node.Name), counts[node.Name])
			}
			if len(node.Path) > 4096 {
				return nil, fmt.Errorf("docxpatch: native extract: XML path in %q exceeds 4096 bytes", partName)
			}
			stack = append(stack, node)
			childCounts = append(childCounts, map[xml.Name]int{})
		case xml.EndElement:
			if len(stack) == 0 {
				return nil, fmt.Errorf("docxpatch: native extract: XML part %q has unmatched end element", partName)
			}
			node := stack[len(stack)-1]
			node.End = after
			stack = stack[:len(stack)-1]
			childCounts = childCounts[:len(childCounts)-1]
		case xml.CharData:
			if len(stack) == 0 {
				if !nativeXMLWhitespaceOnly(string(value)) {
					return nil, fmt.Errorf("docxpatch: native extract: XML part %q has text outside its root", partName)
				}
			} else {
				stack[len(stack)-1].Text += string(value)
			}
		case xml.Directive:
			return nil, fmt.Errorf("docxpatch: native extract: XML directives/DOCTYPE are forbidden in %q", partName)
		case xml.ProcInst:
			if !strings.EqualFold(value.Target, "xml") || root != nil {
				return nil, fmt.Errorf("docxpatch: native extract: processing instruction %q is forbidden in %q", value.Target, partName)
			}
		case xml.Comment:
			// Comments are inert and remain part of the exact parent fingerprint.
		}
	}
	if root == nil || len(stack) != 0 || root.End <= root.Start {
		return nil, fmt.Errorf("docxpatch: native extract: XML part %q has no complete root", partName)
	}
	return root, nil
}

// XML S is intentionally narrower than Unicode whitespace. OOXML lexical
// structure must not treat NBSP or other Unicode separators as ignorable.
func nativeXMLWhitespaceOnly(value string) bool {
	for index := 0; index < len(value); index++ {
		switch value[index] {
		case ' ', '\t', '\r', '\n':
		default:
			return false
		}
	}
	return true
}

func nativeTokenStart(data []byte, before, after int64) int64 {
	if before < 0 || after < before || after > int64(len(data)) {
		return before
	}
	if offset := bytes.IndexByte(data[before:after], '<'); offset >= 0 {
		return before + int64(offset)
	}
	return before
}

func nativeXMLQName(name xml.Name) string {
	switch name.Space {
	case wordMLTransitional, wordMLStrict:
		return "w:" + name.Local
	case relNSTransitional, relNSStrict:
		return "r:" + name.Local
	case opcRelationshipsNS:
		return "pr:" + name.Local
	case opcContentTypesNS:
		return "ct:" + name.Local
	default:
		digest := sha256.Sum256([]byte(name.Space))
		return "ns" + hex.EncodeToString(digest[:4]) + ":" + name.Local
	}
}

func rejectNativeNamespaceSpoofing(root *nativeXMLNode, wordNS string) error {
	known := map[string]bool{"document": true, "body": true, "p": true, "r": true, "t": true, "tbl": true, "tr": true, "tc": true, "sectPr": true, "hdr": true, "ftr": true, "footnotes": true, "footnote": true, "endnotes": true, "endnote": true, "comments": true, "comment": true}
	var visit func(*nativeXMLNode) error
	visit = func(node *nativeXMLNode) error {
		// Genuine direct paragraph math is opaque source, not Word runs. The
		// paragraph extractor retains an unsupported diagnostic and read-only policy.
		if nativeDirectMathRoot(node, wordNS) {
			return nil
		}
		if known[node.Name.Local] && node.Name.Space != wordNS {
			return fmt.Errorf("namespace spoofing at %s: {%s}%s", node.Path, node.Name.Space, node.Name.Local)
		}
		for _, child := range node.Children {
			if err := visit(child); err != nil {
				return err
			}
		}
		return nil
	}
	return visit(root)
}

func nativeUnqualifiedAttr(node *nativeXMLNode, local string) (string, bool) {
	for _, attr := range node.Attrs {
		if attr.Name.Space == "" && attr.Name.Local == local {
			return attr.Value, true
		}
	}
	return "", false
}

func nativeAttr(node *nativeXMLNode, namespace, local string) (string, bool) {
	for _, attr := range node.Attrs {
		if attr.Name.Space == namespace && attr.Name.Local == local {
			return attr.Value, true
		}
	}
	return "", false
}

func firstDirectNativeChild(node *nativeXMLNode, namespace, local string) *nativeXMLNode {
	for _, child := range node.Children {
		if child.Name.Space == namespace && child.Name.Local == local {
			return child
		}
	}
	return nil
}

func directNativeChildren(node *nativeXMLNode, namespace, local string) []*nativeXMLNode {
	children := []*nativeXMLNode{}
	for _, child := range node.Children {
		if child.Name.Space == namespace && child.Name.Local == local {
			children = append(children, child)
		}
	}
	return children
}

func nativeExactContainer(node *nativeXMLNode, allowedAttrs ...xml.Name) bool {
	if !nativeXMLWhitespaceOnly(node.Text) {
		return false
	}
	allowed := make(map[xml.Name]bool, len(allowedAttrs))
	for _, name := range allowedAttrs {
		allowed[name] = true
	}
	for _, attr := range node.Attrs {
		if nativeSettingsNamespaceDeclaration(attr) {
			continue
		}
		if !allowed[attr.Name] {
			return false
		}
	}
	return true
}

func nativeExactLeaf(node *nativeXMLNode, allowedAttrs ...xml.Name) bool {
	return len(node.Children) == 0 && nativeExactContainer(node, allowedAttrs...)
}

func nativeExactRGB(value string) (string, bool) {
	upper := strings.ToUpper(value)
	if len(upper) != 6 || !nativeColor.MatchString(upper) {
		return "", false
	}
	return upper, true
}

func nativeThemeColorSlot(themeColor string) (string, bool) {
	switch themeColor {
	case "dark1", "text1":
		return "dk1", true
	case "light1", "background1":
		return "lt1", true
	case "dark2", "text2":
		return "dk2", true
	case "light2", "background2":
		return "lt2", true
	case "accent1", "accent2", "accent3", "accent4", "accent5", "accent6":
		return themeColor, true
	case "hyperlink":
		return "hlink", true
	case "followedHyperlink":
		return "folHlink", true
	default:
		return "", false
	}
}

func nativeThemeSchemeSlot(local string) bool {
	switch local {
	case "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink":
		return true
	default:
		return false
	}
}

func nativeDrawingAttr(node *nativeXMLNode, drawingNS, local string) (string, bool) {
	unqualified, hasUnqualified := nativeUnqualifiedAttr(node, local)
	namespaced, hasNamespaced := nativeAttr(node, drawingNS, local)
	if hasUnqualified && hasNamespaced && unqualified != namespaced {
		return "", false
	}
	if hasUnqualified {
		return unqualified, true
	}
	if hasNamespaced {
		return namespaced, true
	}
	return "", false
}

func nativeDrawingVal(node *nativeXMLNode, drawingNS string) (string, bool) {
	return nativeDrawingAttr(node, drawingNS, "val")
}

func nativeExactThemeSrgbSlot(slot *nativeXMLNode, drawingNS string) (string, bool) {
	if slot == nil || !nativeExactContainer(slot) || len(slot.Children) != 1 {
		return "", false
	}
	child := slot.Children[0]
	switch child.Name {
	case xml.Name{Space: drawingNS, Local: "srgbClr"}:
		if !nativeExactLeaf(child, xml.Name{Local: "val"}, xml.Name{Space: drawingNS, Local: "val"}) {
			return "", false
		}
		value, ok := nativeDrawingVal(child, drawingNS)
		if !ok {
			return "", false
		}
		return nativeExactRGB(value)
	case xml.Name{Space: drawingNS, Local: "sysClr"}:
		// Word's default theme freezes window/text colors as lastClr sRGB.
		// The system color name is not interpreted; missing/transformed lastClr stays unresolved.
		if !nativeExactLeaf(child,
			xml.Name{Local: "val"}, xml.Name{Space: drawingNS, Local: "val"},
			xml.Name{Local: "lastClr"}, xml.Name{Space: drawingNS, Local: "lastClr"},
		) {
			return "", false
		}
		if _, ok := nativeDrawingVal(child, drawingNS); !ok {
			return "", false
		}
		last, ok := nativeDrawingAttr(child, drawingNS, "lastClr")
		if !ok {
			return "", false
		}
		return nativeExactRGB(last)
	default:
		return "", false
	}
}

type nativeThemeLatinFonts struct {
	major string
	minor string
}

func nativeExactThemeLatinTypeface(fontSet *nativeXMLNode, drawingNS string) (string, bool) {
	if fontSet == nil {
		return "", false
	}
	latin := nativeUniqueThemeFontChild(fontSet, drawingNS, "latin")
	if latin == nil || !nativeExactLeaf(latin,
		xml.Name{Local: "typeface"}, xml.Name{Space: drawingNS, Local: "typeface"},
		xml.Name{Local: "panose"}, xml.Name{Space: drawingNS, Local: "panose"},
	) {
		return "", false
	}
	value, ok := nativeDrawingAttr(latin, drawingNS, "typeface")
	if !ok || value == "" || !nativeBoundedResolvedString(value, 256) {
		return "", false
	}
	return value, true
}

func nativeParseThemeLatinFonts(root *nativeXMLNode, drawingNS string) nativeThemeLatinFonts {
	fonts := nativeThemeLatinFonts{}
	if root == nil {
		return fonts
	}
	elements := nativeUniqueThemeFontChild(root, drawingNS, "themeElements")
	if elements == nil {
		return fonts
	}
	scheme := nativeUniqueThemeFontChild(elements, drawingNS, "fontScheme")
	if scheme == nil {
		return fonts
	}
	if major := nativeUniqueThemeFontChild(scheme, drawingNS, "majorFont"); major != nil {
		fonts.major, _ = nativeExactThemeLatinTypeface(major, drawingNS)
	}
	if minor := nativeUniqueThemeFontChild(scheme, drawingNS, "minorFont"); minor != nil {
		fonts.minor, _ = nativeExactThemeLatinTypeface(minor, drawingNS)
	}
	return fonts
}

// A duplicated or namespace-spoofed font branch has no single authored answer.
func nativeUniqueThemeFontChild(parent *nativeXMLNode, drawingNS, local string) *nativeXMLNode {
	var result *nativeXMLNode
	for _, child := range parent.Children {
		if child.Name.Local != local {
			continue
		}
		if child.Name.Space != drawingNS || result != nil {
			return nil
		}
		result = child
	}
	return result
}

func nativeThemeLatinTypeface(theme string, fonts nativeThemeLatinFonts) (string, bool) {
	switch theme {
	case "majorAscii", "majorHAnsi":
		return fonts.major, fonts.major != ""
	case "minorAscii", "minorHAnsi":
		return fonts.minor, fonts.minor != ""
	default:
		return "", false
	}
}

func nativeParseThemeSrgbColors(root *nativeXMLNode, drawingNS string) map[string]string {
	colors := map[string]string{}
	if root == nil {
		return colors
	}
	elements := firstDirectNativeChild(root, drawingNS, "themeElements")
	if elements == nil {
		return colors
	}
	scheme := firstDirectNativeChild(elements, drawingNS, "clrScheme")
	if scheme == nil {
		return colors
	}
	seen := map[string]bool{}
	blocked := map[string]bool{}
	for _, child := range scheme.Children {
		if child.Name.Space != drawingNS || !nativeThemeSchemeSlot(child.Name.Local) {
			continue
		}
		slot := child.Name.Local
		if seen[slot] {
			blocked[slot] = true
			delete(colors, slot)
			continue
		}
		seen[slot] = true
		if rgb, ok := nativeExactThemeSrgbSlot(child, drawingNS); ok {
			colors[slot] = rgb
		}
	}
	for slot := range blocked {
		delete(colors, slot)
	}
	return colors
}

func nativePackageThemeSrgbColors(pkg *nativePackage, mainPart, wordNS string) map[string]string {
	relBase := relBaseTransitional
	drawingNS := drawingMLTransitional
	if wordNS == wordMLStrict {
		relBase = relBaseStrict
		drawingNS = drawingMLStrict
	}
	partName := ""
	count := 0
	for _, rel := range pkg.rels[mainPart] {
		if rel.Type != relBase+"theme" {
			continue
		}
		if rel.External || rel.PartName == "" {
			return map[string]string{}
		}
		count++
		partName = rel.PartName
	}
	if count != 1 || !nativeASCIIEqual(pkg.contentTypes[partName], "application/vnd.openxmlformats-officedocument.theme+xml") {
		return map[string]string{}
	}
	root, err := parseNativeXML(partName, pkg.files[partName])
	if err != nil || root.Name != (xml.Name{Space: drawingNS, Local: "theme"}) {
		return map[string]string{}
	}
	return nativeParseThemeSrgbColors(root, drawingNS)
}

func (extractor *nativeExtractor) resolveThemeSrgb(themeColor string) (string, bool) {
	slot, ok := nativeThemeColorSlot(themeColor)
	if !ok {
		return "", false
	}
	if !extractor.themeSrgbLoaded {
		extractor.themeSrgbLoaded = true
		extractor.themeSrgbColors = nativePackageThemeSrgbColors(extractor.pkg, extractor.mainPart, extractor.wordNS)
	}
	rgb, ok := extractor.themeSrgbColors[slot]
	return rgb, ok
}

func nativeWordColorLeaf(node *nativeXMLNode, wordNS string) bool {
	return nativeExactLeaf(node,
		xml.Name{Space: wordNS, Local: "val"},
		xml.Name{Space: wordNS, Local: "themeColor"},
		xml.Name{Space: wordNS, Local: "themeTint"},
		xml.Name{Space: wordNS, Local: "themeShade"},
	)
}

func nativeExactRevisionContainer(node *nativeXMLNode, wordNS string, allowedAttrs ...string) bool {
	names := make([]xml.Name, 0, len(allowedAttrs))
	allowed := make(map[string]bool, len(allowedAttrs))
	for _, local := range allowedAttrs {
		names = append(names, xml.Name{Space: wordNS, Local: local})
		allowed[local] = true
	}
	if !nativeExactContainer(node, names...) {
		return false
	}
	for _, attr := range node.Attrs {
		if nativeSettingsNamespaceDeclaration(attr) {
			continue
		}
		if attr.Name.Space != wordNS || !allowed[attr.Name.Local] || !nativeValidLongHexNumber(attr.Value) {
			return false
		}
	}
	return true
}

func nativeSHA(data []byte) string {
	digest := sha256.Sum256(data)
	return "sha256:" + hex.EncodeToString(digest[:])
}

func nativeStableID(kind, partName, nodePath, nativeKey string) string {
	if canonical, err := nativeDecodedPartKey(partName); err == nil {
		partName = canonical
	}
	digest := sha256.Sum256([]byte(kind + "\x00" + partName + "\x00" + nodePath + "\x00" + nativeKey))
	return kind + ":" + hex.EncodeToString(digest[:12])
}

func nativeString(value string) *string { return &value }

func nativeInt(value int) *int       { return &value }
func nativeInt64(value int64) *int64 { return &value }
func nativeBool(value bool) *bool    { return &value }

func (extractor *nativeExtractor) extractRelatedStories() error {
	storyByPart := map[string]string{}
	storyKindByPart := map[string]string{}
	seenSingleton := map[string]bool{}
	for _, rel := range extractor.pkg.rels[extractor.mainPart] {
		if mismatch := extractor.mismatchedRelationshipKind(rel.Type); mismatch != "" {
			return fmt.Errorf("docxpatch: native extract: %s relationship %q uses the wrong Strict/Transitional relationship namespace", mismatch, rel.ID)
		}
		kind, modeled := extractor.relationshipKind(rel.Type)
		if !modeled {
			continue
		}
		if rel.External || rel.PartName == "" {
			return fmt.Errorf("docxpatch: native extract: %s relationship %q must be internal", kind, rel.ID)
		}
		if len(extractor.pkg.files[rel.PartName]) > NativeDOCXMaxXMLPartBytes {
			return fmt.Errorf("docxpatch: native extract: %s part %q exceeds %d bytes", kind, rel.PartName, NativeDOCXMaxXMLPartBytes)
		}
		switch kind {
		case "header", "footer":
			if storyID := storyByPart[rel.PartName]; storyID != "" {
				if storyKindByPart[rel.PartName] != kind {
					return fmt.Errorf("docxpatch: native extract: story part %q is ambiguously related as both %s and %s", rel.PartName, storyKindByPart[rel.PartName], kind)
				}
				extractor.storyByRel[rel.ID] = storyID
				continue
			}
			story, err := extractor.extractHeaderFooter(rel.PartName, kind)
			if err != nil {
				return err
			}
			storyByPart[rel.PartName] = story.ID
			storyKindByPart[rel.PartName] = kind
			extractor.storyByRel[rel.ID] = story.ID
			if kind == "header" {
				extractor.headers = append(extractor.headers, story)
			} else {
				extractor.footers = append(extractor.footers, story)
			}
		case "footnotes", "endnotes", "comments":
			if seenSingleton[kind] {
				return fmt.Errorf("docxpatch: native extract: multiple %s relationships in %q", kind, extractor.mainPart)
			}
			seenSingleton[kind] = true
			wantContentType := map[string]string{
				"footnotes": "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
				"endnotes":  "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
				"comments":  "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
			}[kind]
			if got := extractor.pkg.contentTypes[rel.PartName]; !nativeASCIIEqual(got, wantContentType) {
				return fmt.Errorf("docxpatch: native extract: %s part %q has content type %q; expected %q", kind, rel.PartName, got, wantContentType)
			}
			if kind == "comments" {
				if err := extractor.extractComments(rel.PartName); err != nil {
					return err
				}
			} else if err := extractor.extractNotes(rel.PartName, strings.TrimSuffix(kind, "s"), rel.ID); err != nil {
				return err
			}
		}
	}
	return nil
}

func (extractor *nativeExtractor) mismatchedRelationshipKind(typeURI string) string {
	other := relBaseStrict
	if extractor.wordNS == wordMLStrict {
		other = relBaseTransitional
	}
	for _, kind := range []string{"header", "footer", "footnotes", "endnotes", "comments", "image"} {
		if typeURI == other+kind {
			return kind
		}
	}
	return ""
}

func (extractor *nativeExtractor) relationshipKind(typeURI string) (string, bool) {
	base := relBaseTransitional
	other := relBaseStrict
	if extractor.wordNS == wordMLStrict {
		base, other = relBaseStrict, relBaseTransitional
	}
	for _, kind := range []string{"header", "footer", "footnotes", "endnotes", "comments"} {
		if typeURI == base+kind {
			return kind, true
		}
		if typeURI == other+kind {
			return "", false
		}
	}
	return "", false
}

func (extractor *nativeExtractor) extractHeaderFooter(partName, kind string) (NativeStoryV1, error) {
	root, err := parseNativeXML(partName, extractor.pkg.files[partName])
	if err != nil {
		return NativeStoryV1{}, err
	}
	wantRoot := "hdr"
	if kind == "footer" {
		wantRoot = "ftr"
	}
	if root.Name != (xml.Name{Space: extractor.wordNS, Local: wantRoot}) {
		return NativeStoryV1{}, fmt.Errorf("docxpatch: native extract: %s part %q has spoofed or invalid root", kind, partName)
	}
	if err := rejectNativeNamespaceSpoofing(root, extractor.wordNS); err != nil {
		return NativeStoryV1{}, fmt.Errorf("docxpatch: native extract: %s part %q: %w", kind, partName, err)
	}
	extractor.modeledParts[partName] = true
	storyID := extractor.objectID("story", partName, root, "")
	blocks, err := extractor.extractStoryBlocks(partName, root, storyID)
	if err != nil {
		return NativeStoryV1{}, err
	}
	anchor := extractor.anchor(partName, root)
	return NativeStoryV1{ID: storyID, Kind: kind, PartName: partName, Anchor: &anchor, Blocks: blocks}, nil
}

func (extractor *nativeExtractor) extractNotes(partName, kind, relationshipID string) error {
	root, err := parseNativeXML(partName, extractor.pkg.files[partName])
	if err != nil {
		return err
	}
	wantRoot, wantChild := kind+"s", kind
	if root.Name != (xml.Name{Space: extractor.wordNS, Local: wantRoot}) {
		return fmt.Errorf("docxpatch: native extract: %s part %q has spoofed or invalid root", wantRoot, partName)
	}
	// Word commonly declares ignorable extension namespaces on note roots.
	// Unknown child markup is still recorded below, not silently discarded.
	if !nativeExactContainer(root, xml.Name{Space: "http://schemas.openxmlformats.org/markup-compatibility/2006", Local: "Ignorable"}) {
		return fmt.Errorf("docxpatch: native extract: %s part %q root has attributes or text outside the exact note subset", wantRoot, partName)
	}
	if err := rejectNativeNamespaceSpoofing(root, extractor.wordNS); err != nil {
		return fmt.Errorf("docxpatch: native extract: %s part %q: %w", wantRoot, partName, err)
	}
	extractor.modeledParts[partName] = true
	for _, node := range root.Children {
		if node.Name != (xml.Name{Space: extractor.wordNS, Local: wantChild}) {
			extractor.addUnsupported("UNMODELED_NOTE_MARKUP", "notes", extractor.bodyID, partName, node, "Non-note markup in a note part is preserved verbatim")
			continue
		}
		nativeID, ok := nativeAttr(node, extractor.wordNS, "id")
		noteType, hasType := nativeAttr(node, extractor.wordNS, "type")
		if !nativeExactContainer(node, xml.Name{Space: extractor.wordNS, Local: "id"}, xml.Name{Space: extractor.wordNS, Local: "type"}) {
			return fmt.Errorf("docxpatch: native extract: %s at %s has attributes or direct text outside the exact note subset", wantChild, node.Path)
		}
		noteRole := "content"
		if hasType {
			switch noteType {
			case "separator":
				noteRole = "separator"
			case "continuationSeparator":
				noteRole = "continuation-separator"
			default:
				extractor.addUnsupported("SPECIAL_NOTE_STORY", "notes", extractor.bodyID, partName, node, "Unsupported special note story type is preserved but not exposed")
				continue
			}
		}
		validID := ok && (noteRole == "content" && !hasType && nativeNoteIDPattern.MatchString(nativeID) || noteRole == "separator" && hasType && nativeID == "-1" || noteRole == "continuation-separator" && hasType && nativeID == "0")
		if !validID {
			extractor.addUnsupported("SPECIAL_NOTE_STORY", "notes", extractor.bodyID, partName, node, "Invalid native note story identity/type pairing is preserved but not modeled")
			continue
		}
		storyID := extractor.objectID("story", partName, node, kind+":"+nativeID)
		key := nativeStoryKey(kind, nativeID)
		if extractor.storyByNative[key] != "" {
			return fmt.Errorf("docxpatch: native extract: duplicate %s id %q", kind, nativeID)
		}
		extractor.storyByNative[key] = storyID
		if noteRole != "content" && !nativeExactNoteSentinel(node, extractor.wordNS, noteRole) {
			extractor.addUnsupported("UNMODELED_NOTE_MARKUP", "notes", storyID, partName, node, "Reserved note separator stories must contain exactly one matching instruction leaf and no visible text")
		}
		previousActiveKind, previousActiveStory, previousActiveRole := extractor.activeNoteKind, extractor.activeNoteStory, extractor.activeNoteRole
		extractor.activeNoteKind, extractor.activeNoteStory, extractor.activeNoteRole = kind, storyID, noteRole
		blocks, err := extractor.extractStoryBlocks(partName, node, storyID)
		extractor.activeNoteKind, extractor.activeNoteStory, extractor.activeNoteRole = previousActiveKind, previousActiveStory, previousActiveRole
		if err != nil {
			return err
		}
		anchor := extractor.anchor(partName, node)
		idCopy := nativeID
		story := NativeStoryV1{ID: storyID, Kind: kind, PartName: partName, NativeStoryID: &idCopy, RelationshipID: nativeString(relationshipID), NoteRole: noteRole, Anchor: &anchor, Blocks: blocks}
		extractor.notes = append(extractor.notes, story)
	}
	return nil
}

func nativeExactNoteSentinel(node *nativeXMLNode, wordNS, role string) bool {
	instruction := "separator"
	if role == "continuation-separator" {
		instruction = "continuationSeparator"
	}
	paragraphs := directNativeChildren(node, wordNS, "p")
	// Word records revision-session identifiers on otherwise exact reserved
	// separators. These source-preserved hex IDs do not alter separator layout.
	if len(node.Children) != 1 || len(paragraphs) != 1 || !nativeExactRevisionContainer(paragraphs[0], wordNS, "rsidR", "rsidRDefault", "rsidP") {
		return false
	}
	runs := directNativeChildren(paragraphs[0], wordNS, "r")
	if len(runs) != 1 {
		return false
	}
	paragraph := paragraphs[0]
	properties := directNativeChildren(paragraph, wordNS, "pPr")
	if len(properties) > 1 {
		return false
	}
	if len(properties) == 1 {
		if !nativeExactContainer(properties[0]) {
			return false
		}
		spacing := directNativeChildren(properties[0], wordNS, "spacing")
		after, hasAfter := "", false
		line, hasLine := "", false
		rule, hasRule := "", false
		if len(properties[0].Children) != 1 || len(spacing) != 1 || !nativeExactLeaf(spacing[0], xml.Name{Space: wordNS, Local: "after"}, xml.Name{Space: wordNS, Local: "line"}, xml.Name{Space: wordNS, Local: "lineRule"}) {
			return false
		}
		after, hasAfter = nativeAttr(spacing[0], wordNS, "after")
		line, hasLine = nativeAttr(spacing[0], wordNS, "line")
		rule, hasRule = nativeAttr(spacing[0], wordNS, "lineRule")
		if !hasAfter || !hasLine || !hasRule || after != "0" || line != "240" || rule != "auto" {
			return false
		}
	}
	for _, child := range paragraph.Children {
		if child != runs[0] && (len(properties) == 0 || child != properties[0]) {
			return false
		}
	}
	run := runs[0]
	return nativeExactContainer(run) && len(run.Children) == 1 && run.Children[0].Name == (xml.Name{Space: wordNS, Local: instruction}) && nativeExactLeaf(run.Children[0])
}

func (extractor *nativeExtractor) extractComments(partName string) error {
	root, err := parseNativeXML(partName, extractor.pkg.files[partName])
	if err != nil {
		return err
	}
	if root.Name != (xml.Name{Space: extractor.wordNS, Local: "comments"}) {
		return fmt.Errorf("docxpatch: native extract: comments part %q has spoofed or invalid root", partName)
	}
	if err := rejectNativeNamespaceSpoofing(root, extractor.wordNS); err != nil {
		return fmt.Errorf("docxpatch: native extract: comments part %q: %w", partName, err)
	}
	extractor.modeledParts[partName] = true
	for _, node := range root.Children {
		if node.Name != (xml.Name{Space: extractor.wordNS, Local: "comment"}) {
			extractor.addUnsupported("UNMODELED_COMMENT_MARKUP", "comments", extractor.bodyID, partName, node, "Non-comment markup in the comments part is preserved verbatim")
			continue
		}
		nativeID, okID := nativeAttr(node, extractor.wordNS, "id")
		author, okAuthor := nativeAttr(node, extractor.wordNS, "author")
		if !okID || !nativeIDPattern.MatchString(nativeID) || !okAuthor || author == "" {
			return fmt.Errorf("docxpatch: native extract: invalid comment identity at %s", node.Path)
		}
		if extractor.commentByNative[nativeID] != "" {
			return fmt.Errorf("docxpatch: native extract: duplicate comment id %q", nativeID)
		}
		storyID := extractor.objectID("story", partName, node, "comment:"+nativeID)
		commentID := extractor.objectID("comment", partName, node, nativeID)
		blocks, err := extractor.extractStoryBlocks(partName, node, storyID)
		if err != nil {
			return err
		}
		anchor := extractor.anchor(partName, node)
		idCopy := nativeID
		extractor.commentStories = append(extractor.commentStories, NativeStoryV1{ID: storyID, Kind: "comment", PartName: partName, NativeStoryID: &idCopy, Anchor: &anchor, Blocks: blocks})
		comment := NativeCommentV1{ID: commentID, NativeCommentID: nativeID, Author: author, Anchor: &anchor, BodyStoryID: storyID}
		if value, ok := nativeAttr(node, extractor.wordNS, "initials"); ok && value != "" {
			comment.Initials = nativeString(value)
		}
		if value, ok := nativeAttr(node, extractor.wordNS, "date"); ok && value != "" {
			comment.CreatedAt = nativeString(value)
		}
		extractor.comments = append(extractor.comments, comment)
		extractor.commentByNative[nativeID] = commentID
		extractor.storyByNative[nativeStoryKey("comment", nativeID)] = storyID
	}
	return nil
}

func nativeStoryKey(kind, nativeID string) string { return kind + "\x00" + nativeID }

func (extractor *nativeExtractor) indexPreviousIDs() {
	if extractor.previous == nil {
		return
	}
	remember := func(kind, id string, anchor *NativeSourceAnchorV1, nativeKey string) {
		extractor.reservedIDs[id] = true
		partKey := anchorPart(anchor)
		if canonical, err := nativeDecodedPartKey(partKey); err == nil {
			partKey = canonical
		}
		if nativeKey != "" {
			extractor.previousNative[kind+"\x00"+partKey+"\x00"+nativeKey] = id
		}
		if anchor == nil || anchor.PartName == "" || anchor.XMLSHA256 == "" {
			return
		}
		key := kind + "\x00" + partKey + "\x00" + anchor.XMLSHA256
		extractor.previousIDs[key] = append(extractor.previousIDs[key], id)
		pathKey := kind + "\x00" + partKey + "\x00" + anchor.Path
		if _, duplicate := extractor.previousPath[pathKey]; duplicate {
			extractor.previousPath[pathKey] = ""
		} else {
			extractor.previousPath[pathKey] = id
		}
	}
	var paragraph func(NativeParagraphV1)
	paragraph = func(value NativeParagraphV1) {
		remember("paragraph", value.ID, &value.Anchor, "")
		for _, run := range value.Runs {
			remember("run", run.ID, &run.Anchor, "")
			if run.Drawing != nil {
				remember("drawing", run.Drawing.ID, &run.Drawing.Anchor, "")
			}
		}
	}
	var story func(NativeStoryV1)
	story = func(value NativeStoryV1) {
		nativeKey := ""
		if value.NativeStoryID != nil {
			nativeKey = value.Kind + ":" + *value.NativeStoryID
		}
		remember("story", value.ID, value.Anchor, nativeKey)
		for _, block := range value.Blocks {
			switch block.Kind {
			case "paragraph":
				if block.Paragraph != nil {
					paragraph(*block.Paragraph)
				}
			case "table":
				if block.Table == nil {
					continue
				}
				remember("table", block.Table.ID, &block.Table.Anchor, "")
				for _, row := range block.Table.Rows {
					remember("row", row.ID, &row.Anchor, "")
					for _, cell := range row.Cells {
						remember("cell", cell.ID, &cell.Anchor, "")
						for _, nested := range cell.Paragraphs {
							paragraph(nested)
						}
					}
				}
			}
		}
	}
	story(extractor.previous.Body)
	for _, collection := range [][]NativeStoryV1{extractor.previous.Headers, extractor.previous.Footers, extractor.previous.Notes, extractor.previous.CommentStories} {
		for _, value := range collection {
			story(value)
		}
	}
	for _, section := range extractor.previous.Sections {
		remember("section", section.ID, &section.Anchor, "")
	}
	for _, comment := range extractor.previous.Comments {
		remember("comment", comment.ID, comment.Anchor, comment.NativeCommentID)
	}
	for _, entry := range extractor.previous.Unsupported {
		nativeKey := entry.Code + "\x00" + entry.Capability + "\x00" + entry.ScopeID
		if entry.Anchor != nil {
			nativeKey += "\x00" + entry.Anchor.XMLSHA256
		}
		remember("unsupported", entry.ID, entry.Anchor, nativeKey)
	}
}

func anchorPart(anchor *NativeSourceAnchorV1) string {
	if anchor == nil {
		return ""
	}
	return anchor.PartName
}

func (extractor *nativeExtractor) objectID(kind, partName string, node *nativeXMLNode, nativeKey string) string {
	partKey := partName
	if canonical, err := nativeDecodedPartKey(partName); err == nil {
		partKey = canonical
	}
	if nativeKey != "" {
		key := kind + "\x00" + partKey + "\x00" + nativeKey
		if previous := extractor.previousNative[key]; previous != "" && !extractor.allocatedIDs[previous] {
			extractor.allocatedIDs[previous] = true
			return previous
		}
		if node != nil {
			anchor := extractor.anchor(partName, node)
			anchorKey := kind + "\x00" + partKey + "\x00" + anchor.XMLSHA256
			for index := extractor.previousUsed[anchorKey]; index < len(extractor.previousIDs[anchorKey]); index++ {
				extractor.previousUsed[anchorKey] = index + 1
				candidate := extractor.previousIDs[anchorKey][index]
				if !extractor.allocatedIDs[candidate] {
					extractor.allocatedIDs[candidate] = true
					return candidate
				}
			}
			if candidate := extractor.previousPathCandidate(kind, partKey, node); candidate != "" {
				return candidate
			}
		}
		candidate := nativeStableID(kind, partKey, "", nativeKey)
		for salt := 0; extractor.allocatedIDs[candidate]; salt++ {
			nodePath := ""
			if node != nil {
				nodePath = node.Path
			}
			candidate = nativeStableID(kind, partKey, nodePath, nativeKey+":"+strconv.Itoa(salt))
		}
		extractor.allocatedIDs[candidate] = true
		return candidate
	}
	if node != nil {
		anchor := extractor.anchor(partName, node)
		key := kind + "\x00" + partKey + "\x00" + anchor.XMLSHA256
		candidates := extractor.previousIDs[key]
		for index := extractor.previousUsed[key]; index < len(candidates); index++ {
			extractor.previousUsed[key] = index + 1
			if extractor.allocatedIDs[candidates[index]] {
				continue
			}
			extractor.allocatedIDs[candidates[index]] = true
			return candidates[index]
		}
		if candidate := extractor.previousPathCandidate(kind, partKey, node); candidate != "" {
			return candidate
		}
		candidate := nativeStableID(kind, partKey, node.Path, "")
		for salt := 0; extractor.reservedIDs[candidate] || extractor.allocatedIDs[candidate]; salt++ {
			candidate = nativeStableID(kind, partKey, node.Path, anchor.XMLSHA256+":"+strconv.Itoa(salt))
		}
		extractor.allocatedIDs[candidate] = true
		return candidate
	}
	candidate := nativeStableID(kind, partKey, "", "")
	extractor.allocatedIDs[candidate] = true
	return candidate
}

func (extractor *nativeExtractor) previousPathCandidate(kind, partKey string, node *nativeXMLNode) string {
	if !extractor.retainPathIdentity || node == nil {
		return ""
	}
	key := kind + "\x00" + partKey + "\x00" + node.Path
	candidate := extractor.previousPath[key]
	if candidate == "" || extractor.allocatedIDs[candidate] {
		return ""
	}
	extractor.allocatedIDs[candidate] = true
	return candidate
}

func (extractor *nativeExtractor) passthroughParts() []NativePassthroughPartV1 {
	names := make([]string, 0, len(extractor.pkg.files))
	for name := range extractor.pkg.files {
		if name != "[Content_Types].xml" && !extractor.modeledParts[name] {
			names = append(names, name)
		}
	}
	sort.Strings(names)
	parts := make([]NativePassthroughPartV1, 0, len(names))
	for _, name := range names {
		data := extractor.pkg.files[name]
		length := int64(len(data))
		parts = append(parts, NativePassthroughPartV1{PartName: name, ContentType: extractor.pkg.contentTypes[name], ByteLength: &length, SHA256: nativeSHA(data), Policy: "preserve-verbatim"})
	}
	return parts
}

func (extractor *nativeExtractor) anchor(partName string, node *nativeXMLNode) NativeSourceAnchorV1 {
	start, end := node.Start, node.End
	raw := extractor.pkg.files[partName]
	digest := "sha256:" + strings.Repeat("0", 64)
	if start >= 0 && end > start && end <= int64(len(raw)) {
		digest = nativeSHA(raw[start:end])
	}
	return NativeSourceAnchorV1{PartName: partName, Path: node.Path, StartByte: &start, EndByte: &end, XMLSHA256: digest}
}

func (extractor *nativeExtractor) addUnsupported(code, capability, scopeID, partName string, node *nativeXMLNode, message string) {
	if len(extractor.unsupported) >= NativeDOCXMaxCollectionItems {
		extractor.unsupportedOverflow = true
		return
	}
	key := code + "\x00" + scopeID + "\x00" + partName + "\x00"
	if node != nil {
		key += node.Path
	}
	if extractor.unsupportedSet[key] {
		return
	}
	extractor.unsupportedSet[key] = true
	nativeKey := code + "\x00" + capability + "\x00" + scopeID
	if node != nil {
		nativeKey += "\x00" + extractor.anchor(partName, node).XMLSHA256
	}
	id := extractor.objectID("unsupported", partName, node, nativeKey)
	entry := NativeUnsupportedCapabilityV1{ID: id, Code: code, Capability: capability, ScopeID: scopeID, Preservation: "refuse-mutation", Message: message}
	if node != nil {
		anchor := extractor.anchor(partName, node)
		entry.Anchor = &anchor
	}
	extractor.unsupported = append(extractor.unsupported, entry)
}

func nativeMutableParagraphPolicy() NativeEditPolicyV1 {
	return NativeEditPolicyV1{Mode: "read-write", AllowedOperations: []string{"text.replace"}}
}

func nativeExtractOnlyTablePolicy() NativeEditPolicyV1 {
	return nativeReadOnlyPolicy("EXTRACT_ONLY", "Native extraction does not yet expose a guarded contract mutation API")
}

func nativeReadOnlyPolicy(code, message string) NativeEditPolicyV1 {
	return NativeEditPolicyV1{Mode: "read-only", AllowedOperations: []string{}, Refusal: &NativeRefusalV1{Code: code, Message: message, Preservation: "refuse-mutation"}}
}

type nativeSectionMarker struct {
	node       *nativeXMLNode
	startBlock int
	synthetic  bool
}

func (extractor *nativeExtractor) extractBody() (NativeStoryV1, []NativeSectionV1, error) {
	blocks := []NativeBlockV1{}
	markers := []nativeSectionMarker{}
	sectionStart := 0
	finalSection := false
	for _, child := range extractor.bodyNode.Children {
		switch {
		case child.Name == (xml.Name{Space: extractor.wordNS, Local: "p"}):
			paragraph, err := extractor.extractParagraph(extractor.mainPart, child)
			if err != nil {
				return NativeStoryV1{}, nil, err
			}
			blocks = append(blocks, NativeBlockV1{Kind: "paragraph", ID: paragraph.ID, Paragraph: &paragraph})
			if pPr := firstDirectNativeChild(child, extractor.wordNS, "pPr"); pPr != nil {
				if sectPr := firstDirectNativeChild(pPr, extractor.wordNS, "sectPr"); sectPr != nil {
					markers = append(markers, nativeSectionMarker{node: sectPr, startBlock: sectionStart})
					sectionStart = len(blocks)
				}
			}
		case child.Name == (xml.Name{Space: extractor.wordNS, Local: "tbl"}):
			table, err := extractor.extractTable(extractor.mainPart, child)
			if err != nil {
				return NativeStoryV1{}, nil, err
			}
			blocks = append(blocks, NativeBlockV1{Kind: "table", ID: table.ID, Table: &table})
		case child.Name == (xml.Name{Space: extractor.wordNS, Local: "sectPr"}):
			if finalSection {
				return NativeStoryV1{}, nil, fmt.Errorf("docxpatch: native extract: body contains multiple final w:sectPr elements")
			}
			finalSection = true
			markers = append(markers, nativeSectionMarker{node: child, startBlock: sectionStart})
		default:
			extractor.addUnsupported("UNMODELED_BODY_BLOCK", "body-structure", extractor.bodyID, extractor.mainPart, child, "Body markup outside paragraphs, basic tables, and sections is preserved verbatim")
		}
		if len(blocks) > NativeDOCXMaxCollectionItems {
			return NativeStoryV1{}, nil, fmt.Errorf("docxpatch: native extract: body exceeds %d modeled blocks", NativeDOCXMaxCollectionItems)
		}
	}
	if len(blocks) == 0 {
		return NativeStoryV1{}, nil, fmt.Errorf("docxpatch: native extract: a native body requires at least one paragraph or table")
	}
	if !finalSection && sectionStart < len(blocks) {
		markers = append(markers, nativeSectionMarker{node: extractor.bodyNode, startBlock: sectionStart, synthetic: true})
	}
	sections := make([]NativeSectionV1, 0, len(markers))
	for _, marker := range markers {
		if marker.startBlock < 0 || marker.startBlock >= len(blocks) {
			return NativeStoryV1{}, nil, fmt.Errorf("docxpatch: native extract: empty trailing section at %s is not safely representable", marker.node.Path)
		}
		var section NativeSectionV1
		var err error
		if marker.synthetic {
			section = extractor.defaultSection(marker.node, blocks[marker.startBlock].ID)
		} else {
			section, err = extractor.extractSection(marker.node, blocks[marker.startBlock].ID)
			if err != nil {
				return NativeStoryV1{}, nil, err
			}
		}
		if marker.synthetic {
			extractor.addUnsupported("DEFAULT_SECTION_INFERRED", "sections", section.ID, extractor.mainPart, extractor.bodyNode, "The schema-optional final w:sectPr is absent; Word default page geometry is exposed with the body anchor")
		}
		sections = append(sections, section)
	}
	bodyAnchor := extractor.anchor(extractor.mainPart, extractor.bodyNode)
	story := NativeStoryV1{ID: extractor.bodyID, Kind: "body", PartName: extractor.mainPart, Anchor: &bodyAnchor, Blocks: blocks}
	return story, sections, nil
}

func (extractor *nativeExtractor) extractStoryBlocks(partName string, container *nativeXMLNode, storyID string) ([]NativeBlockV1, error) {
	blocks := []NativeBlockV1{}
	for _, child := range container.Children {
		switch {
		case child.Name == (xml.Name{Space: extractor.wordNS, Local: "p"}):
			paragraph, err := extractor.extractParagraph(partName, child)
			if err != nil {
				return nil, err
			}
			blocks = append(blocks, NativeBlockV1{Kind: "paragraph", ID: paragraph.ID, Paragraph: &paragraph})
		case child.Name == (xml.Name{Space: extractor.wordNS, Local: "tbl"}):
			table, err := extractor.extractTable(partName, child)
			if err != nil {
				return nil, err
			}
			blocks = append(blocks, NativeBlockV1{Kind: "table", ID: table.ID, Table: &table})
		default:
			extractor.addUnsupported("UNMODELED_STORY_BLOCK", "story-structure", storyID, partName, child, "Story markup outside paragraphs and basic tables is preserved verbatim")
		}
		if len(blocks) > NativeDOCXMaxCollectionItems {
			return nil, fmt.Errorf("docxpatch: native extract: story %q exceeds %d modeled blocks", storyID, NativeDOCXMaxCollectionItems)
		}
	}
	return blocks, nil
}

func (extractor *nativeExtractor) extractParagraph(partName string, node *nativeXMLNode) (NativeParagraphV1, error) {
	nativeKey := ""
	paraIDIssueCode := ""
	paraIDIssueMessage := ""
	for _, attr := range node.Attrs {
		if attr.Name.Space == "http://schemas.microsoft.com/office/word/2010/wordml" && attr.Name.Local == "paraId" && attr.Value != "" {
			if nativeValidOfficeHexID(attr.Value) {
				candidateKey := "para:" + strings.ToUpper(attr.Value)
				partKey, _ := nativeDecodedPartKey(partName)
				identityKey := partKey + "\x00" + candidateKey
				if originalPath := extractor.seenParaIDs[identityKey]; originalPath != "" {
					paraIDIssueCode = "DUPLICATE_NATIVE_PARAGRAPH_ID"
					paraIDIssueMessage = fmt.Sprintf("Duplicate w14:paraId %q (first seen at %s) is preserved but excluded from identity generation", attr.Value, originalPath)
				} else {
					nativeKey = candidateKey
					extractor.seenParaIDs[identityKey] = node.Path
				}
			} else {
				paraIDIssueCode = "INVALID_NATIVE_PARAGRAPH_ID"
				paraIDIssueMessage = "w14:paraId must be eight hexadecimal digits with a value from 00000001 through 7FFFFFFF; the source value is preserved but excluded from identity generation"
			}
		}
	}
	id := extractor.objectID("paragraph", partName, node, nativeKey)
	paragraph := NativeParagraphV1{ID: id, Anchor: extractor.anchor(partName, node), EditPolicy: nativeReadOnlyPolicy("NO_MUTABLE_TEXT", "Paragraph does not contain a safely addressable native text node"), Properties: &NativeParagraphPropertiesV1{}, Runs: []NativeRunV1{}}
	unsafe := false
	if paraIDIssueCode != "" {
		unsafe = true
		extractor.addUnsupported(paraIDIssueCode, "identity", id, partName, node, paraIDIssueMessage)
	}
	paragraphProperties := directNativeChildren(node, extractor.wordNS, "pPr")
	if len(paragraphProperties) > 1 {
		unsafe = true
		extractor.addUnsupported("DUPLICATE_PARAGRAPH_PROPERTIES", "paragraph-properties", id, partName, node, "Duplicate w:pPr elements make the effective paragraph pagination properties ambiguous")
	}
	if len(paragraphProperties) == 1 {
		pPr := paragraphProperties[0]
		properties, propertyUnsafe, err := extractor.extractParagraphProperties(partName, id, pPr)
		if err != nil {
			return NativeParagraphV1{}, err
		}
		paragraph.Properties = properties
		unsafe = unsafe || propertyUnsafe
	}
	runs, runUnsafe, err := extractor.extractParagraphRuns(partName, id, node)
	if err != nil {
		return NativeParagraphV1{}, err
	}
	paragraph.Runs = runs
	unsafe = unsafe || runUnsafe
	if unsafe {
		paragraph.EditPolicy = nativeReadOnlyPolicy("UNMODELED_PARAGRAPH_MARKUP", "Paragraph contains OOXML that the v1 extractor preserves but cannot safely mutate")
	} else {
		for _, run := range runs {
			if run.Kind == "text" && run.Text != nil {
				paragraph.EditPolicy = nativeMutableParagraphPolicy()
				break
			}
		}
	}
	return paragraph, nil
}

func nativeValidOfficeHexID(value string) bool {
	if len(value) != 8 {
		return false
	}
	for index := 0; index < len(value); index++ {
		character := value[index]
		if !((character >= '0' && character <= '9') || (character >= 'a' && character <= 'f') || (character >= 'A' && character <= 'F')) {
			return false
		}
	}
	parsed, err := strconv.ParseUint(value, 16, 32)
	return err == nil && parsed > 0 && parsed < 0x80000000
}

func nativeValidLongHexNumber(value string) bool {
	if len(value) != 8 {
		return false
	}
	for index := 0; index < len(value); index++ {
		character := value[index]
		if !((character >= '0' && character <= '9') || (character >= 'a' && character <= 'f') || (character >= 'A' && character <= 'F')) {
			return false
		}
	}
	return true
}

func (extractor *nativeExtractor) extractParagraphProperties(partName, paragraphID string, node *nativeXMLNode) (*NativeParagraphPropertiesV1, bool, error) {
	properties := &NativeParagraphPropertiesV1{}
	unsafe := false
	preserveOnly := false
	if !nativeExactRevisionContainer(node, extractor.wordNS, "rsidRPr") {
		unsafe = true
		extractor.addUnsupported("UNMODELED_PARAGRAPH_PROPERTY", "paragraph-properties", paragraphID, partName, node, "Paragraph-properties container has attributes or direct text outside the exact v1 subset")
	}
	seenSingleton := map[string]bool{}
	for _, child := range node.Children {
		if child.Name.Space != extractor.wordNS {
			unsafe = true
			extractor.addUnsupported("FOREIGN_PARAGRAPH_PROPERTY", "paragraph-properties", paragraphID, partName, child, "Foreign namespace paragraph property is preserved verbatim")
			continue
		}
		if child.Name.Local == "pStyle" || child.Name.Local == "numPr" || child.Name.Local == "jc" || child.Name.Local == "spacing" || child.Name.Local == "ind" || child.Name.Local == "keepNext" || child.Name.Local == "keepLines" || child.Name.Local == "pageBreakBefore" || child.Name.Local == "widowControl" || child.Name.Local == "bidi" || child.Name.Local == "sectPr" || child.Name.Local == "rPr" {
			if seenSingleton[child.Name.Local] {
				unsafe = true
				extractor.addUnsupported("DUPLICATE_PARAGRAPH_PROPERTY", "paragraph-properties", paragraphID, partName, child, "Duplicate modeled paragraph-property singletons are preserved but not treated as editable or pagination-safe")
				continue
			}
			seenSingleton[child.Name.Local] = true
		}
		switch child.Name.Local {
		case "pStyle":
			if value, ok := nativeAttr(child, extractor.wordNS, "val"); ok && nativeIDPattern.MatchString(value) && nativeExactLeaf(child, xml.Name{Space: extractor.wordNS, Local: "val"}) {
				properties.ParagraphStyleID = nativeString(value)
			} else {
				unsafe = true
			}
		case "numPr":
			numIDNodes := directNativeChildren(child, extractor.wordNS, "numId")
			levelNodes := directNativeChildren(child, extractor.wordNS, "ilvl")
			if len(numIDNodes) > 1 || len(levelNodes) > 1 {
				unsafe = true
				extractor.addUnsupported("DUPLICATE_NUMBERING_REFERENCE", "paragraph-properties", paragraphID, partName, child, "Duplicate numbering singleton children are preserved but excluded from the native projection")
				continue
			}
			var numIDNode, levelNode *nativeXMLNode
			if len(numIDNodes) == 1 {
				numIDNode = numIDNodes[0]
			}
			if len(levelNodes) == 1 {
				levelNode = levelNodes[0]
			}
			numID, okNum := "", false
			level := 0
			okLevel := false
			if numIDNode != nil {
				numID, okNum = nativeAttr(numIDNode, extractor.wordNS, "val")
			}
			if levelNode != nil {
				if raw, ok := nativeAttr(levelNode, extractor.wordNS, "val"); ok {
					parsed, parseErr := strconv.Atoi(raw)
					if parseErr == nil && parsed >= 0 {
						level, okLevel = parsed, true
					}
				}
			}
			exact := len(child.Attrs) == 0 && nativeXMLWhitespaceOnly(child.Text) && len(child.Children) == len(numIDNodes)+len(levelNodes)
			if numIDNode != nil {
				exact = exact && nativeExactLeaf(numIDNode, xml.Name{Space: extractor.wordNS, Local: "val"})
			}
			if levelNode != nil {
				exact = exact && nativeExactLeaf(levelNode, xml.Name{Space: extractor.wordNS, Local: "val"})
			}
			if okNum && nativeIDPattern.MatchString(numID) && okLevel && exact {
				properties.Numbering = &NativeNumberingReferenceV1{NumID: numID, Level: nativeInt(level)}
			} else {
				unsafe = true
			}
		case "jc":
			if value, ok := nativeAttr(child, extractor.wordNS, "val"); ok && nativeExactLeaf(child, xml.Name{Space: extractor.wordNS, Local: "val"}) && (value == "left" || value == "center" || value == "right" || value == "both" || value == "distribute") {
				properties.Alignment = nativeString(value)
			} else {
				unsafe = true
			}
		case "spacing":
			// The resolved-layout projection owns these values. Extraction still
			// proves the exact source shape so an honest direct-formatting layer is
			// not mislabeled as unmodeled before the resolver can attest it. Because
			// this projection cannot round-trip spacing, the paragraph remains
			// preservation-only even when the source shape is exact.
			preserveOnly = true
			if !nativeExactResolvedParagraphSpacing(child, extractor.wordNS) {
				unsafe = true
				extractor.addUnsupported("UNMODELED_PARAGRAPH_SPACING", "paragraph-properties", paragraphID, partName, child, "Paragraph spacing is invalid, automatic, line-unit based, or has structure outside the exact resolved-layout subset")
			}
		case "rPr":
			// Paragraph-mark metrics are owned by the resolved-layout projection,
			// including empty paragraphs. Admit only a closed direct-formatting
			// subset; do not make the paragraph writable or apply mark properties
			// to its body runs.
			preserveOnly = true
			if !nativeExactParagraphMarkProperties(child, extractor.wordNS) {
				unsafe = true
				extractor.addUnsupported("UNMODELED_PARAGRAPH_MARK_PROPERTIES", "paragraph-properties", paragraphID, partName, child, "Paragraph-mark formatting has unknown, duplicate, or noncanonical source structure")
			} else if _, invalid, _ := extractor.extractRunPropertiesState(partName, paragraphID, child); invalid {
				unsafe = true
			}
		case "ind":
			preserveOnly = true
			if !nativeExactResolvedParagraphIndent(child, extractor.wordNS) {
				unsafe = true
				extractor.addUnsupported("UNMODELED_PARAGRAPH_INDENT", "paragraph-properties", paragraphID, partName, child, "Paragraph indentation is invalid, conflicting, character-unit based, or has structure outside the exact resolved-layout subset")
			}
		case "keepNext", "keepLines", "pageBreakBefore", "widowControl":
			value, ok := nativeOnOff(child, extractor.wordNS)
			if !ok || !nativeExactLeaf(child, xml.Name{Space: extractor.wordNS, Local: "val"}) {
				unsafe = true
				break
			}
			switch child.Name.Local {
			case "keepNext":
				properties.KeepNext = nativeBool(value)
			case "keepLines":
				properties.KeepLines = nativeBool(value)
			case "pageBreakBefore":
				properties.PageBreakBefore = nativeBool(value)
			case "widowControl":
				properties.WidowControl = nativeBool(value)
			}
		case "autoSpaceDE", "autoSpaceDN", "adjustRightInd":
			preserveOnly = true
			if !nativeNeutralSourceProperty(child, node, extractor.wordNS) {
				unsafe = true
				extractor.addUnsupported("UNMODELED_PARAGRAPH_PROPERTY", "paragraph-properties", paragraphID, partName, child, "Automatic spacing or grid indent adjustment is supported only as an exact explicit disabled setting")
			}
		case "bidi":
			preserveOnly = true
			if _, ok := nativeOnOff(child, extractor.wordNS); !ok || !nativeExactLeaf(child, xml.Name{Space: extractor.wordNS, Local: "val"}) {
				unsafe = true
			}
		case "sectPr":
			// Modeled separately as a section.
		default:
			unsafe = true
			extractor.addUnsupported("UNMODELED_PARAGRAPH_PROPERTY", "paragraph-properties", paragraphID, partName, child, "This paragraph property is preserved verbatim")
		}
	}
	if unsafe {
		extractor.addUnsupported("PARTIAL_PARAGRAPH_PROPERTIES", "paragraph-properties", paragraphID, partName, node, "Only the conservative v1 paragraph-property subset is exposed")
	}
	return properties, unsafe || preserveOnly, nil
}

func nativeExactParagraphMarkProperties(node *nativeXMLNode, wordNS string) bool {
	if !nativeExactContainer(node) {
		return false
	}
	seen := map[string]bool{}
	for _, property := range node.Children {
		if property.Name.Space != wordNS || seen[property.Name.Local] {
			return false
		}
		seen[property.Name.Local] = true
		switch property.Name.Local {
		case "rFonts":
			// Preserve the bounded complex-script slot without treating it as
			// active. The resolver still qualifies the paragraph mark's script
			// context; RTL/mixed-script uncertainty retains its diagnostic.
			if !nativeExactLeaf(property, xml.Name{Space: wordNS, Local: "ascii"}, xml.Name{Space: wordNS, Local: "hAnsi"}, xml.Name{Space: wordNS, Local: "cs"}, xml.Name{Space: wordNS, Local: "eastAsia"}) {
				return false
			}
			for _, slot := range []string{"cs", "eastAsia"} {
				if value, present := nativeAttr(property, wordNS, slot); present && !nativeBoundedResolvedString(value, 256) {
					return false
				}
			}
		case "lang":
			if !nativeExactLeaf(property, xml.Name{Space: wordNS, Local: "val"}, xml.Name{Space: wordNS, Local: "eastAsia"}, xml.Name{Space: wordNS, Local: "bidi"}) {
				return false
			}
			for _, slot := range []string{"eastAsia", "bidi"} {
				if value, present := nativeAttr(property, wordNS, slot); present && !nativeScriptLanguageTag(value) {
					return false
				}
			}
		case "sz", "b", "i", "rtl", "vanish", "color":
			if !nativeExactLeaf(property, xml.Name{Space: wordNS, Local: "val"}) {
				return false
			}
		case "kern":
			if _, ok := nativeKerningThreshold(property, wordNS); !ok {
				return false
			}
		case "szCs":
			// Preserve the exact inactive slot. Resolved script context still
			// decides whether its size is active; RTL/mixed text keeps refusal.
			value, ok := nativePositiveIntAttr(property, wordNS, "val")
			if !ok || value > 3276 || !nativeExactLeaf(property, xml.Name{Space: wordNS, Local: "val"}) {
				return false
			}
		default:
			return false
		}
	}
	return true
}

func nativeExactResolvedParagraphSpacing(node *nativeXMLNode, wordNS string) bool {
	if !nativeExactLeaf(node,
		xml.Name{Space: wordNS, Local: "before"}, xml.Name{Space: wordNS, Local: "after"},
		xml.Name{Space: wordNS, Local: "line"}, xml.Name{Space: wordNS, Local: "lineRule"},
		xml.Name{Space: wordNS, Local: "beforeAutospacing"}, xml.Name{Space: wordNS, Local: "afterAutospacing"},
		xml.Name{Space: wordNS, Local: "beforeLines"}, xml.Name{Space: wordNS, Local: "afterLines"}) {
		return false
	}
	for _, name := range []string{"before", "after", "line"} {
		if _, present := nativeAttr(node, wordNS, name); present {
			if _, valid := nativeNonnegativeInt64Attr(node, wordNS, name); !valid {
				return false
			}
		}
	}
	for _, name := range []string{"beforeAutospacing", "afterAutospacing"} {
		if raw, present := nativeAttr(node, wordNS, name); present {
			value, valid := nativeLexicalOnOff(raw)
			if !valid || value {
				return false
			}
		}
	}
	if _, present := nativeAttr(node, wordNS, "beforeLines"); present {
		return false
	}
	if _, present := nativeAttr(node, wordNS, "afterLines"); present {
		return false
	}
	if rule, present := nativeAttr(node, wordNS, "lineRule"); present {
		_, linePresent := nativeAttr(node, wordNS, "line")
		if !linePresent || (rule != "auto" && rule != "exact" && rule != "atLeast") {
			return false
		}
	}
	return true
}

func nativeExactResolvedParagraphIndent(node *nativeXMLNode, wordNS string) bool {
	if !nativeExactLeaf(node,
		xml.Name{Space: wordNS, Local: "left"}, xml.Name{Space: wordNS, Local: "right"},
		xml.Name{Space: wordNS, Local: "start"}, xml.Name{Space: wordNS, Local: "end"},
		xml.Name{Space: wordNS, Local: "firstLine"}, xml.Name{Space: wordNS, Local: "hanging"},
		xml.Name{Space: wordNS, Local: "leftChars"}, xml.Name{Space: wordNS, Local: "rightChars"},
		xml.Name{Space: wordNS, Local: "startChars"}, xml.Name{Space: wordNS, Local: "endChars"},
		xml.Name{Space: wordNS, Local: "firstLineChars"}, xml.Name{Space: wordNS, Local: "hangingChars"}) {
		return false
	}
	for _, name := range []string{"left", "right", "start", "end"} {
		if _, present := nativeAttr(node, wordNS, name); present {
			if _, valid := nativeSignedInt64Attr(node, wordNS, name); !valid {
				return false
			}
		}
	}
	for _, name := range []string{"firstLine", "hanging"} {
		if _, present := nativeAttr(node, wordNS, name); present {
			if _, valid := nativeNonnegativeInt64Attr(node, wordNS, name); !valid {
				return false
			}
		}
	}
	if _, first := nativeAttr(node, wordNS, "firstLine"); first {
		if _, hanging := nativeAttr(node, wordNS, "hanging"); hanging {
			return false
		}
	}
	for _, name := range []string{"leftChars", "rightChars", "startChars", "endChars", "firstLineChars", "hangingChars"} {
		if _, present := nativeAttr(node, wordNS, name); present {
			return false
		}
	}
	return true
}

func (extractor *nativeExtractor) extractParagraphRuns(partName, paragraphID string, paragraph *nativeXMLNode) ([]NativeRunV1, bool, error) {
	runs := []NativeRunV1{}
	unsafe := false
	for childIndex := 0; childIndex < len(paragraph.Children); childIndex++ {
		child := paragraph.Children[childIndex]
		if child.Name == (xml.Name{Space: extractor.wordNS, Local: "pPr"}) {
			continue
		}
		switch {
		case child.Name == (xml.Name{Space: extractor.wordNS, Local: "bookmarkStart"}) && nativeExactEmptyBookmark(paragraph.Children[childIndex:], extractor.wordNS):
			// A closed, empty bookmark has no painted content. Its source is
			// retained and the containing paragraph remains non-editable.
			unsafe = true
			childIndex++
		case child.Name == (xml.Name{Space: extractor.wordNS, Local: "r"}):
			if hasNativeFieldBegin(child, extractor.wordNS) {
				unsafe = true
				field, ok, err := extractor.extractFlatPageField(partName, paragraphID, paragraph.Children[childIndex:])
				if err != nil {
					return nil, false, err
				}
				if ok {
					runs = append(runs, field)
					childIndex += 4
					continue
				}
				extractor.addUnsupported("FIELD_SEMANTICS", "fields", paragraphID, partName, child, "Complex page fields require an exact flat begin/instruction/separate/result/end run sequence")
			}
			extracted, runUnsafe, err := extractor.extractRunNode(partName, paragraphID, child)
			if err != nil {
				return nil, false, err
			}
			runs = append(runs, extracted...)
			unsafe = unsafe || runUnsafe
		case child.Name == (xml.Name{Space: extractor.wordNS, Local: "fldSimple"}):
			unsafe = true // Field results are never editable text.
			instruction, present := nativeAttr(child, extractor.wordNS, "instr")
			instruction, qualifiedInstruction := nativePageFieldInstruction(instruction)
			instructionAttrs := 0
			for _, attr := range child.Attrs {
				if attr.Name == (xml.Name{Space: extractor.wordNS, Local: "instr"}) {
					instructionAttrs++
				}
			}
			if !present || instructionAttrs != 1 || !qualifiedInstruction || !nativeExactContainer(child, xml.Name{Space: extractor.wordNS, Local: "instr"}) || len(child.Children) != 1 || child.Children[0].Name != (xml.Name{Space: extractor.wordNS, Local: "r"}) {
				extractor.addUnsupported("FIELD_SEMANTICS", "fields", paragraphID, partName, child, "Only an unlocked simple decimal PAGE or NUMPAGES field with one text result run is modeled")
				continue
			}
			extracted, runUnsafe, err := extractor.extractRunNode(partName, paragraphID, child.Children[0])
			if err != nil {
				return nil, false, err
			}
			if runUnsafe || len(extracted) != 1 || extracted[0].Kind != "text" || !nativeExactContainer(child.Children[0]) || len(directNativeChildren(child.Children[0], extractor.wordNS, "rPr")) > 1 {
				extractor.addUnsupported("FIELD_SEMANTICS", "fields", paragraphID, partName, child, "Page-field result must be one exact text run; nested fields and controls are refused")
				continue
			}
			extracted[0].PageField = instruction
			extracted[0].Text = nativeString("") // Ignore stale cache, including nonnumeric values.
			runs = append(runs, extracted[0])
		case child.Name == (xml.Name{Space: extractor.wordNS, Local: "hyperlink"}):
			unsafe = true
			extractor.addUnsupported("HYPERLINK_SEMANTICS", "hyperlinks", paragraphID, partName, child, "Visible hyperlink text is exposed, while relationship and field semantics remain preserve-only")
			for _, nested := range directNativeChildren(child, extractor.wordNS, "r") {
				extracted, _, err := extractor.extractRunNode(partName, paragraphID, nested)
				if err != nil {
					return nil, false, err
				}
				runs = append(runs, extracted...)
			}
		case child.Name == (xml.Name{Space: extractor.wordNS, Local: "commentRangeStart"}), child.Name == (xml.Name{Space: extractor.wordNS, Local: "commentRangeEnd"}):
			nativeID, ok := nativeAttr(child, extractor.wordNS, "id")
			targetID := extractor.commentByNative[nativeID]
			if !ok || targetID == "" {
				unsafe = true
				extractor.addUnsupported("UNRESOLVED_COMMENT_RANGE", "comments", paragraphID, partName, child, "Comment range marker does not resolve to a modeled comment")
				continue
			}
			kind := "comment-range-start"
			if child.Name.Local == "commentRangeEnd" {
				kind = "comment-range-end"
			}
			runs = append(runs, NativeRunV1{Kind: "reference", ID: extractor.objectID("run", partName, child, kind+":"+nativeID), Anchor: extractor.anchor(partName, child), Reference: &NativeReferenceV1{Kind: kind, TargetID: targetID}})
		case child.Name.Space == extractor.wordNS && (child.Name.Local == "sdt" || child.Name.Local == "smartTag" || child.Name.Local == "customXml" || child.Name.Local == "ins" || child.Name.Local == "moveTo"):
			unsafe = true
			extractor.addUnsupported("WRAPPED_RUN_MARKUP", "run-structure", paragraphID, partName, child, "Wrapped or revision-tracked runs are preserved and exposed read-only")
			for _, nested := range nativeDescendants(child, extractor.wordNS, "r") {
				extracted, _, err := extractor.extractRunNode(partName, paragraphID, nested)
				if err != nil {
					return nil, false, err
				}
				runs = append(runs, extracted...)
			}
		default:
			unsafe = true
			extractor.addUnsupported("UNMODELED_PARAGRAPH_CONTENT", "run-structure", paragraphID, partName, child, "Paragraph content outside the v1 run subset is preserved verbatim")
		}
		if len(runs) > NativeDOCXMaxCollectionItems {
			return nil, false, fmt.Errorf("docxpatch: native extract: paragraph %q exceeds %d runs", paragraphID, NativeDOCXMaxCollectionItems)
		}
	}
	return runs, unsafe, nil
}

func (extractor *nativeExtractor) extractRunNode(partName, paragraphID string, node *nativeXMLNode) ([]NativeRunV1, bool, error) {
	var properties *NativeRunPropertiesV1
	unsafe := false
	if rPr := firstDirectNativeChild(node, extractor.wordNS, "rPr"); rPr != nil {
		parsed, propertyUnsafe := extractor.extractRunProperties(partName, paragraphID, rPr)
		properties, unsafe = parsed, unsafe || propertyUnsafe
	}
	runs := []NativeRunV1{}
	for _, child := range node.Children {
		if child.Name == (xml.Name{Space: extractor.wordNS, Local: "rPr"}) {
			continue
		}
		base := NativeRunV1{ID: extractor.objectID("run", partName, child, ""), Anchor: extractor.anchor(partName, child), Properties: properties}
		switch {
		case child.Name == (xml.Name{Space: extractor.wordNS, Local: "t"}):
			if len(child.Children) != 0 {
				unsafe = true
				extractor.addUnsupported("COMPLEX_TEXT_NODE", "text", paragraphID, partName, child, "Nested markup inside w:t is preserved verbatim")
				continue
			}
			text := child.Text
			base.Kind, base.Text = "text", &text
			runs = append(runs, base)
		case child.Name == (xml.Name{Space: extractor.wordNS, Local: "tab"}):
			if !nativeExactLeaf(child) {
				unsafe = true
				extractor.addUnsupported("UNMODELED_CONTROL", "controls", paragraphID, partName, child, "Tab control has attributes, children, or text outside the exact v1 subset")
				continue
			}
			base.Kind, base.Control = "control", "tab"
			runs = append(runs, base)
		case child.Name == (xml.Name{Space: extractor.wordNS, Local: "br"}):
			if !nativeExactLeaf(child, xml.Name{Space: extractor.wordNS, Local: "type"}, xml.Name{Space: extractor.wordNS, Local: "clear"}) {
				unsafe = true
				extractor.addUnsupported("UNMODELED_BREAK", "controls", paragraphID, partName, child, "Break control has attributes, children, or text outside the exact v1 subset")
				continue
			}
			if clear, present := nativeAttr(child, extractor.wordNS, "clear"); present && clear != "none" {
				unsafe = true
				extractor.addUnsupported("BREAK_CLEAR_UNSUPPORTED", "controls", paragraphID, partName, child, "Text-wrapping break clear semantics require floating-layout authority")
				continue
			}
			control := "line-break"
			if breakType, ok := nativeAttr(child, extractor.wordNS, "type"); ok {
				switch breakType {
				case "textWrapping":
				case "page":
					control = "page-break"
				case "column":
					control = "column-break"
				default:
					unsafe = true
					extractor.addUnsupported("UNMODELED_BREAK", "controls", paragraphID, partName, child, "This break kind is preserved verbatim")
					continue
				}
			}
			base.Kind, base.Control = "control", control
			runs = append(runs, base)
		case child.Name == (xml.Name{Space: extractor.wordNS, Local: "cr"}):
			if !nativeExactLeaf(child) {
				unsafe = true
				extractor.addUnsupported("UNMODELED_BREAK", "controls", paragraphID, partName, child, "Carriage-return control has attributes, children, or text outside the exact v1 subset")
				continue
			}
			base.Kind, base.Control = "control", "line-break"
			runs = append(runs, base)
		case child.Name == (xml.Name{Space: extractor.wordNS, Local: "softHyphen"}):
			if !nativeExactLeaf(child) {
				unsafe = true
				extractor.addUnsupported("UNMODELED_CONTROL", "controls", paragraphID, partName, child, "Soft hyphen has attributes, children, or text outside the exact v1 subset")
				continue
			}
			base.Kind, base.Control = "control", "soft-hyphen"
			runs = append(runs, base)
		case child.Name.Space == extractor.wordNS && (child.Name.Local == "footnoteReference" || child.Name.Local == "endnoteReference"):
			kind := strings.TrimSuffix(child.Name.Local, "Reference")
			nativeID, ok := nativeAttr(child, extractor.wordNS, "id")
			target := extractor.storyByNative[nativeStoryKey(kind, nativeID)]
			if !ok || target == "" || !nativeExactLeaf(child, xml.Name{Space: extractor.wordNS, Local: "id"}) {
				unsafe = true
				extractor.addUnsupported("UNRESOLVED_NOTE_REFERENCE", "notes", paragraphID, partName, child, "Note reference must be an exact id-only leaf resolving to a modeled note story")
				continue
			}
			base.Kind, base.Reference = "reference", &NativeReferenceV1{Kind: kind, TargetID: target}
			runs = append(runs, base)
		case child.Name.Space == extractor.wordNS && (child.Name.Local == "footnoteRef" || child.Name.Local == "endnoteRef"):
			kind := strings.TrimSuffix(child.Name.Local, "Ref")
			if !nativeExactLeaf(child) || extractor.activeNoteKind != kind || extractor.activeNoteStory == "" {
				unsafe = true
				extractor.addUnsupported("UNMODELED_NOTE_MARKUP", "notes", paragraphID, partName, child, "Note label is only supported as an exact leaf in its owning note story")
				continue
			}
			base.Kind, base.Reference = "reference", &NativeReferenceV1{Kind: kind, TargetID: extractor.activeNoteStory, Role: "label"}
			runs = append(runs, base)
		case child.Name == (xml.Name{Space: extractor.wordNS, Local: "separator"}):
			if !nativeExactLeaf(child) || extractor.activeNoteRole != "separator" {
				unsafe = true
				extractor.addUnsupported("UNMODELED_NOTE_MARKUP", "notes", paragraphID, partName, child, "Separator instruction must be an exact leaf in the ordinary separator sentinel")
			}
		case child.Name == (xml.Name{Space: extractor.wordNS, Local: "continuationSeparator"}):
			if !nativeExactLeaf(child) || extractor.activeNoteRole != "continuation-separator" {
				unsafe = true
				extractor.addUnsupported("UNMODELED_NOTE_MARKUP", "notes", paragraphID, partName, child, "Continuation separator instruction must be an exact leaf in the continuation sentinel")
			}
		case child.Name == (xml.Name{Space: extractor.wordNS, Local: "commentReference"}):
			nativeID, ok := nativeAttr(child, extractor.wordNS, "id")
			target := extractor.commentByNative[nativeID]
			if !ok || target == "" {
				unsafe = true
				extractor.addUnsupported("UNRESOLVED_COMMENT_REFERENCE", "comments", paragraphID, partName, child, "Comment reference does not resolve to a modeled comment")
				continue
			}
			base.Kind, base.Reference = "reference", &NativeReferenceV1{Kind: "comment", TargetID: target}
			runs = append(runs, base)
		case child.Name.Space == extractor.wordNS && (child.Name.Local == "instrText" || child.Name.Local == "fldChar"):
			unsafe = true
			extractor.addUnsupported("FIELD_SEMANTICS", "fields", paragraphID, partName, child, "Field instructions and boundaries are preserved verbatim")
		case child.Name == (xml.Name{Space: extractor.wordNS, Local: "drawing"}):
			drawing, ok := extractor.extractDrawing(partName, paragraphID, child)
			if !ok {
				unsafe = true
				continue
			}
			base.Kind, base.Drawing = "drawing", drawing
			runs = append(runs, base)
		case child.Name.Space == extractor.wordNS && (child.Name.Local == "pict" || child.Name.Local == "object"):
			unsafe = true
			extractor.addUnsupported("UNMODELED_DRAWING", "drawings", paragraphID, partName, child, "Drawing/object markup and related media are preserved verbatim")
		default:
			unsafe = true
			extractor.addUnsupported("UNMODELED_RUN_CONTENT", "runs", paragraphID, partName, child, "Run content outside text, controls, and native references is preserved verbatim")
		}
	}
	return runs, unsafe, nil
}

func nativeDescendants(node *nativeXMLNode, namespace, local string) []*nativeXMLNode {
	result := []*nativeXMLNode{}
	var visit func(*nativeXMLNode)
	visit = func(current *nativeXMLNode) {
		for _, child := range current.Children {
			if child.Name.Space == namespace && child.Name.Local == local {
				result = append(result, child)
				continue
			}
			visit(child)
		}
	}
	visit(node)
	return result
}

func (extractor *nativeExtractor) drawingNamespaces() (string, string, string) {
	if extractor.wordNS == wordMLStrict {
		return wordDrawingStrict, drawingMLStrict, pictureMLStrict
	}
	return wordDrawingTransitional, drawingMLTransitional, pictureMLTransitional
}

func (extractor *nativeExtractor) extractDrawing(partName, paragraphID string, node *nativeXMLNode) (*NativeDrawingV1, bool) {
	wpNS, aNS, picNS := extractor.drawingNamespaces()
	containers := []*nativeXMLNode{}
	for _, child := range node.Children {
		if child.Name.Space == wpNS && (child.Name.Local == "inline" || child.Name.Local == "anchor") {
			containers = append(containers, child)
		}
	}
	refuse := func(code, message string, anchor *nativeXMLNode) (*NativeDrawingV1, bool) {
		extractor.addUnsupported(code, "drawings", paragraphID, partName, anchor, message)
		return nil, false
	}
	if len(containers) != 1 || len(node.Children) != 1 {
		return refuse("AMBIGUOUS_DRAWING", "A DrawingML run must contain exactly one wp:inline or wp:anchor picture", node)
	}
	container := containers[0]
	if container.Name.Local == "anchor" && !nativeExactPageAnchor(container, wpNS, aNS) {
		return refuse("FLOATING_DRAWING_SEMANTICS_PRESERVED", "Only exact page-relative wrapNone or bothSides wrapSquare anchors with explicit layering and overlap are projected", container)
	}
	if container.Name.Local == "inline" && !nativeExactInlinePictureContainer(container, wpNS, aNS) {
		return refuse("INLINE_DRAWING_SEMANTICS_PRESERVED", "Inline pictures with unmodeled container attributes or children remain preserve-only", container)
	}
	extent := firstDirectNativeChild(container, wpNS, "extent")
	if extent == nil {
		return refuse("DRAWING_EXTENT_REQUIRED", "Picture extent is missing", container)
	}
	width, okWidth := nativePositiveInt64Attr(extent, "", "cx")
	height, okHeight := nativePositiveInt64Attr(extent, "", "cy")
	if !okWidth || !okHeight {
		return refuse("INVALID_DRAWING_EXTENT", "Picture extent must contain positive safe cx/cy values", extent)
	}
	var inlineEffects *NativeDrawingCropV1
	if effects := directNativeChildren(container, wpNS, "effectExtent"); len(effects) > 0 {
		if len(effects) != 1 {
			return refuse("DRAWING_EFFECTS_PRESERVED", "Ambiguous drawing effect extents", container)
		}
		if !nativeZeroExtent(effects[0]) {
			if container.Name.Local != "inline" {
				return refuse("DRAWING_EFFECTS_PRESERVED", "Floating drawing effect extents remain unqualified", container)
			}
			var valid bool
			inlineEffects, valid = nativeInlineEffectExtents(effects[0])
			if !valid {
				return refuse("DRAWING_EFFECTS_PRESERVED", "Inline effect extents must be exact bounded nonnegative EMUs", container)
			}
		}
	}
	for _, attrName := range []string{"distT", "distB", "distL", "distR"} {
		if raw, present := nativeUnqualifiedAttr(container, attrName); present && raw != "0" {
			return refuse("DRAWING_DISTANCE_PRESERVED", "Non-zero picture wrap distances are preserved but not projected", container)
		}
	}
	docPr := firstDirectNativeChild(container, wpNS, "docPr")
	if docPr == nil {
		return refuse("DRAWING_METADATA_REQUIRED", "Picture docPr metadata is missing", container)
	}
	if !nativeExactLeaf(docPr, xml.Name{Local: "id"}, xml.Name{Local: "name"}, xml.Name{Local: "descr"}, xml.Name{Local: "title"}) {
		return refuse("DRAWING_METADATA_PRESERVED", "Picture metadata with visibility or unmodeled semantics remains preserve-only", docPr)
	}
	if _, valid := nativePositiveInt64Attr(docPr, "", "id"); !valid {
		return refuse("DRAWING_METADATA_PRESERVED", "Picture metadata requires a positive canonical docPr id", docPr)
	}
	if name, present := nativeUnqualifiedAttr(docPr, "name"); !present || name == "" {
		return refuse("DRAWING_METADATA_PRESERVED", "Picture metadata requires an explicit non-empty docPr name", docPr)
	}
	graphic := firstDirectNativeChild(container, aNS, "graphic")
	if graphic == nil || len(graphic.Children) != 1 || !nativeExactContainer(graphic) {
		return refuse("PICTURE_GRAPHIC_REQUIRED", "DrawingML graphic payload is missing", container)
	}
	graphicData := firstDirectNativeChild(graphic, aNS, "graphicData")
	if graphicData == nil || !nativeExactContainer(graphicData, xml.Name{Local: "uri"}) {
		return refuse("PICTURE_GRAPHIC_REQUIRED", "DrawingML graphicData payload is missing", graphic)
	}
	if uri, ok := nativeUnqualifiedAttr(graphicData, "uri"); !ok || uri != picNS {
		return refuse("PICTURE_GRAPHIC_REQUIRED", "DrawingML graphicData must declare the exact picture namespace URI", graphicData)
	}
	pictures := directNativeChildren(graphicData, picNS, "pic")
	if len(pictures) != 1 || len(graphicData.Children) != 1 {
		return refuse("AMBIGUOUS_PICTURE", "Only a single native DrawingML picture payload is modeled", graphicData)
	}
	picture := pictures[0]
	crop, cropOK := nativePictureSourceCrop(picture, aNS, picNS)
	if !cropOK {
		return refuse("PICTURE_CROP_PRESERVED", "Picture crop must be one exact source rectangle retaining at least one percent per axis", picture)
	}
	if !nativeExactPictureNonVisual(picture, aNS, picNS) {
		return refuse("PICTURE_NONVISUAL_PRESERVED", "Picture nonvisual properties with missing, hidden, or unmodeled semantics remain preserve-only", picture)
	}
	if !nativePictureBoundedTransform(picture, aNS, picNS, width, height) {
		return refuse("PICTURE_TRANSFORM_PRESERVED", "Only flips and quarter turns with exact rotated DrawingML/inline extents are projected", picture)
	}
	blips := nativeDescendants(picture, aNS, "blip")
	if len(blips) != 1 || !nativeExactLeaf(blips[0], xml.Name{Space: extractor.relNS, Local: "embed"}, xml.Name{Local: "cstate"}) {
		return refuse("PICTURE_EFFECTS_PRESERVED", "Pictures with missing, ambiguous, or effect-bearing blips remain preserve-only", picture)
	}
	if state, present := nativeUnqualifiedAttr(blips[0], "cstate"); present && state != "email" && state != "screen" && state != "print" && state != "hqprint" && state != "none" {
		return refuse("PICTURE_EFFECTS_PRESERVED", "Picture compression state has an unsupported lexical value", blips[0])
	}
	relID, okRelID := nativeAttr(blips[0], extractor.relNS, "embed")
	if !okRelID || relID == "" {
		return refuse("PICTURE_RELATIONSHIP_REQUIRED", "Picture blip has a spoofed or missing native relationship id", blips[0])
	}
	mediaPart, contentType, okImage := extractor.imageRelationship(partName, relID)
	if !okImage {
		return refuse("PICTURE_RELATIONSHIP_INVALID", "Picture relationship must resolve internally to an image part", blips[0])
	}
	drawing := &NativeDrawingV1{
		ID: extractor.objectID("drawing", partName, node, ""), Anchor: extractor.anchor(partName, node),
		RelationshipID: nativeString(relID), MediaPart: nativeString(mediaPart), ContentType: nativeString(contentType),
		Placement: "inline", WidthEMU: nativeInt64(width), HeightEMU: nativeInt64(height),
		SourceCrop:            crop,
		InlineEffectExtentEMU: inlineEffects,
		EditPolicy:            nativeReadOnlyPolicy("EXTRACT_ONLY", "Native picture extraction does not yet expose guarded drawing replacement"),
	}
	xfrm := firstDirectNativeChild(firstDirectNativeChild(picture, picNS, "spPr"), aNS, "xfrm")
	if rotation, ok := nativeUnqualifiedAttr(xfrm, "rot"); ok {
		angle, _ := strconv.ParseInt(rotation, 10, 64) // bounded lexical values qualified above
		degrees := angle / 60000
		drawing.RotationDegrees = nativeInt64(degrees)
	}
	if flip, ok := nativeUnqualifiedAttr(xfrm, "flipH"); ok {
		value := flip == "1" || flip == "true"
		drawing.FlipHorizontal = &value
	}
	if flip, ok := nativeUnqualifiedAttr(xfrm, "flipV"); ok {
		value := flip == "1" || flip == "true"
		drawing.FlipVertical = &value
	}
	if value, ok := nativeUnqualifiedAttr(docPr, "name"); ok && value != "" {
		drawing.Name = nativeString(value)
	}
	if value, ok := nativeUnqualifiedAttr(docPr, "descr"); ok && value != "" {
		drawing.AltText = nativeString(value)
	} else if value, ok := nativeUnqualifiedAttr(docPr, "title"); ok && value != "" {
		drawing.AltText = nativeString(value)
	}
	if container.Name.Local == "anchor" {
		drawing.Placement = "floating"
		layer := "front"
		if behind, _ := nativeUnqualifiedAttr(container, "behindDoc"); behind == "1" || behind == "true" {
			layer = "behind"
		}
		drawing.FloatingLayer = nativeString(layer)
		order, _ := nativeNonnegativeInt64Attr(container, "", "relativeHeight")
		drawing.StackingOrder = nativeInt64(order)
		positionH := firstDirectNativeChild(container, wpNS, "positionH")
		positionV := firstDirectNativeChild(container, wpNS, "positionV")
		if positionH == nil || positionV == nil {
			return refuse("FLOATING_POSITION_REQUIRED", "Floating pictures require explicit horizontal and vertical positions", container)
		}
		x, relativeH, okH := nativeDrawingPosition(positionH, wpNS)
		y, relativeV, okV := nativeDrawingPosition(positionV, wpNS)
		if !okH || !okV {
			return refuse("FLOATING_ALIGNMENT_PRESERVED", "Aligned/ambiguous floating positions are preserved but not projected as offsets", container)
		}
		drawing.XEMU, drawing.YEMU = nativeInt64(x), nativeInt64(y)
		if relativeH != "" {
			drawing.HorizontalRelativeFrom = nativeString(relativeH)
		}
		if relativeV != "" {
			drawing.VerticalRelativeFrom = nativeString(relativeV)
		}
		wrap, ok := nativeDrawingWrap(container, wpNS)
		if !ok {
			return refuse("DRAWING_WRAP_PRESERVED", "Floating picture wrap geometry is ambiguous or unsupported", container)
		}
		drawing.Wrap = nativeString(wrap)
	}
	return drawing, true
}

func nativeExactInlinePictureContainer(container *nativeXMLNode, wpNS, aNS string) bool {
	if !nativeExactContainer(container, xml.Name{Local: "distT"}, xml.Name{Local: "distB"}, xml.Name{Local: "distL"}, xml.Name{Local: "distR"}) {
		return false
	}
	allowed := map[xml.Name]bool{
		{Space: wpNS, Local: "extent"}:            true,
		{Space: wpNS, Local: "effectExtent"}:      true,
		{Space: wpNS, Local: "docPr"}:             true,
		{Space: wpNS, Local: "cNvGraphicFramePr"}: true,
		{Space: aNS, Local: "graphic"}:            true,
	}
	for _, child := range container.Children {
		if !allowed[child.Name] {
			return false
		}
	}
	if len(directNativeChildren(container, wpNS, "extent")) != 1 || len(directNativeChildren(container, wpNS, "effectExtent")) > 1 || len(directNativeChildren(container, wpNS, "docPr")) != 1 || len(directNativeChildren(container, wpNS, "cNvGraphicFramePr")) > 1 || len(directNativeChildren(container, aNS, "graphic")) != 1 {
		return false
	}
	extent := firstDirectNativeChild(container, wpNS, "extent")
	if !nativeExactLeaf(extent, xml.Name{Local: "cx"}, xml.Name{Local: "cy"}) {
		return false
	}
	if effect := firstDirectNativeChild(container, wpNS, "effectExtent"); effect != nil && !nativeExactLeaf(effect, xml.Name{Local: "l"}, xml.Name{Local: "t"}, xml.Name{Local: "r"}, xml.Name{Local: "b"}) {
		return false
	}
	if frame := firstDirectNativeChild(container, wpNS, "cNvGraphicFramePr"); frame != nil {
		if !nativeExactContainer(frame) || len(frame.Children) > 1 {
			return false
		}
		if len(frame.Children) == 1 && (frame.Children[0].Name != (xml.Name{Space: aNS, Local: "graphicFrameLocks"}) || !nativeExactLeaf(frame.Children[0], xml.Name{Local: "noChangeAspect"})) {
			return false
		}
	}
	return true
}

func nativeExactPictureNonVisual(picture *nativeXMLNode, aNS, picNS string) bool {
	nonVisual := firstDirectNativeChild(picture, picNS, "nvPicPr")
	if nonVisual == nil || !nativeExactContainer(nonVisual) || len(nonVisual.Children) != 2 {
		return false
	}
	common := firstDirectNativeChild(nonVisual, picNS, "cNvPr")
	pictureProperties := firstDirectNativeChild(nonVisual, picNS, "cNvPicPr")
	if common == nil || pictureProperties == nil || len(directNativeChildren(nonVisual, picNS, "cNvPr")) != 1 || len(directNativeChildren(nonVisual, picNS, "cNvPicPr")) != 1 {
		return false
	}
	if !nativeExactLeaf(common, xml.Name{Local: "id"}, xml.Name{Local: "name"}, xml.Name{Local: "descr"}, xml.Name{Local: "title"}) {
		return false
	}
	if _, valid := nativeNonnegativeInt64Attr(common, "", "id"); !valid {
		return false
	}
	if name, present := nativeUnqualifiedAttr(common, "name"); !present || name == "" {
		return false
	}
	if !nativeExactContainer(pictureProperties) || len(pictureProperties.Children) > 1 {
		return false
	}
	if len(pictureProperties.Children) == 0 {
		return true
	}
	locks := pictureProperties.Children[0]
	if locks.Name != (xml.Name{Space: aNS, Local: "picLocks"}) {
		return false
	}
	lockNames := []string{"noGrp", "noSelect", "noRot", "noChangeAspect", "noMove", "noResize", "noEditPoints", "noAdjustHandles", "noChangeArrowheads", "noChangeShapeType", "noCrop"}
	allowed := make([]xml.Name, 0, len(lockNames))
	for _, name := range lockNames {
		allowed = append(allowed, xml.Name{Local: name})
	}
	if !nativeExactLeaf(locks, allowed...) {
		return false
	}
	for _, attr := range locks.Attrs {
		if nativeSettingsNamespaceDeclaration(attr) {
			continue
		}
		if attr.Value != "true" && attr.Value != "false" && attr.Value != "1" && attr.Value != "0" {
			return false
		}
	}
	return true
}

func nativePictureSourceCrop(picture *nativeXMLNode, aNS, picNS string) (*NativeDrawingCropV1, bool) {
	crops := nativeDescendants(picture, aNS, "srcRect")
	if len(crops) == 0 {
		return nil, true
	}
	fill := firstDirectNativeChild(picture, picNS, "blipFill")
	if fill == nil || len(crops) != 1 {
		return nil, false
	}
	direct := directNativeChildren(fill, aNS, "srcRect")
	if len(direct) != 1 || direct[0] != crops[0] || !nativeExactLeaf(crops[0], xml.Name{Local: "l"}, xml.Name{Local: "t"}, xml.Name{Local: "r"}, xml.Name{Local: "b"}) {
		return nil, false
	}
	values := [4]int64{}
	for index, name := range []string{"l", "t", "r", "b"} {
		if _, present := nativeUnqualifiedAttr(crops[0], name); !present {
			continue
		}
		value, ok := nativeInt64Attr(crops[0], "", name)
		if !ok || value < 0 || value > 99000 {
			return nil, false
		}
		values[index] = value
	}
	if values[0]+values[2] > 99000 || values[1]+values[3] > 99000 {
		return nil, false
	}
	return &NativeDrawingCropV1{Left: nativeInt64(values[0]), Top: nativeInt64(values[1]), Right: nativeInt64(values[2]), Bottom: nativeInt64(values[3])}, true
}

func nativePictureBoundedTransform(picture *nativeXMLNode, aNS, picNS string, width, height int64) bool {
	if !nativeExactContainer(picture) || len(picture.Children) != 3 || len(directNativeChildren(picture, picNS, "nvPicPr")) != 1 || len(directNativeChildren(picture, picNS, "blipFill")) != 1 || len(directNativeChildren(picture, picNS, "spPr")) != 1 {
		return false
	}
	blipFill := firstDirectNativeChild(picture, picNS, "blipFill")
	blips := directNativeChildren(blipFill, aNS, "blip")
	stretches := directNativeChildren(blipFill, aNS, "stretch")
	crops := directNativeChildren(blipFill, aNS, "srcRect")
	if !nativeExactContainer(blipFill) || len(crops) > 1 || len(blipFill.Children) != 2+len(crops) || len(blips) != 1 || len(stretches) != 1 || len(blips[0].Children) != 0 || !nativeExactContainer(stretches[0]) || len(stretches[0].Children) != 1 {
		return false
	}
	fillRect := firstDirectNativeChild(stretches[0], aNS, "fillRect")
	if fillRect == nil || len(fillRect.Children) != 0 || len(fillRect.Attrs) != 0 {
		return false
	}
	spPr := firstDirectNativeChild(picture, picNS, "spPr")
	if spPr == nil || !nativeExactContainer(spPr) {
		return false
	}
	xfrms := directNativeChildren(spPr, aNS, "xfrm")
	geometries := directNativeChildren(spPr, aNS, "prstGeom")
	if len(spPr.Children) != 2 || len(xfrms) != 1 || len(geometries) != 1 {
		return false
	}
	geometry := geometries[0]
	if preset, ok := nativeUnqualifiedAttr(geometry, "prst"); !ok || preset != "rect" || !nativeExactContainer(geometry, xml.Name{Local: "prst"}) || len(geometry.Children) > 1 || len(geometry.Children) == 1 && (geometry.Children[0].Name != (xml.Name{Space: aNS, Local: "avLst"}) || !nativeExactLeaf(geometry.Children[0])) {
		return false
	}
	xfrm := xfrms[0]
	seenTransformAttrs := map[string]bool{}
	for _, attr := range xfrm.Attrs {
		if attr.Name.Space != "" || (attr.Name.Local != "rot" && attr.Name.Local != "flipH" && attr.Name.Local != "flipV") {
			return false
		}
		if seenTransformAttrs[attr.Name.Local] {
			return false
		}
		seenTransformAttrs[attr.Name.Local] = true
		if attr.Name.Local == "rot" && attr.Value != "0" && attr.Value != "5400000" && attr.Value != "10800000" && attr.Value != "16200000" || (attr.Name.Local == "flipH" || attr.Name.Local == "flipV") && attr.Value != "0" && attr.Value != "false" && attr.Value != "1" && attr.Value != "true" {
			return false
		}
	}
	if !nativeXMLWhitespaceOnly(xfrm.Text) {
		return false
	}
	rotation, _ := nativeUnqualifiedAttr(xfrm, "rot")
	quarterTurn := rotation == "5400000" || rotation == "16200000"
	if len(xfrm.Children) == 0 {
		return !quarterTurn // a quarter turn needs explicit original extents
	}
	if len(xfrm.Children) != 2 {
		return false
	}
	off := firstDirectNativeChild(xfrm, aNS, "off")
	ext := firstDirectNativeChild(xfrm, aNS, "ext")
	if off == nil || ext == nil || !nativeExactLeaf(off, xml.Name{Local: "x"}, xml.Name{Local: "y"}) || !nativeExactLeaf(ext, xml.Name{Local: "cx"}, xml.Name{Local: "cy"}) {
		return false
	}
	x, okX := nativeInt64Attr(off, "", "x")
	y, okY := nativeInt64Attr(off, "", "y")
	cx, okCX := nativePositiveInt64Attr(ext, "", "cx")
	cy, okCY := nativePositiveInt64Attr(ext, "", "cy")
	if quarterTurn {
		return okX && okY && x == 0 && y == 0 && okCX && okCY && cy == width && cx == height
	}
	return okX && okY && x == 0 && y == 0 && okCX && okCY && cx == width && cy == height
}

func nativeInt64Attr(node *nativeXMLNode, namespace, local string) (int64, bool) {
	raw, ok := nativeAttr(node, namespace, local)
	if !ok {
		return 0, false
	}
	value, err := strconv.ParseInt(raw, 10, 64)
	return value, err == nil && value >= -9007199254740991 && value <= 9007199254740991
}

func nativePositiveInt64Attr(node *nativeXMLNode, namespace, local string) (int64, bool) {
	raw, ok := nativeAttr(node, namespace, local)
	if !ok {
		return 0, false
	}
	value, err := strconv.ParseInt(raw, 10, 64)
	return value, err == nil && value > 0 && value <= 9007199254740991
}

func nativeZeroExtent(node *nativeXMLNode) bool {
	if !nativeExactLeaf(node, xml.Name{Local: "l"}, xml.Name{Local: "t"}, xml.Name{Local: "r"}, xml.Name{Local: "b"}) {
		return false
	}
	for _, name := range []string{"l", "t", "r", "b"} {
		value, ok := nativeUnqualifiedAttr(node, name)
		if !ok || value != "0" {
			return false
		}
	}
	return true
}

func (extractor *nativeExtractor) imageRelationship(ownerPart, relID string) (string, string, bool) {
	base := relBaseTransitional
	if extractor.wordNS == wordMLStrict {
		base = relBaseStrict
	}
	for _, rel := range extractor.pkg.rels[ownerPart] {
		if rel.ID != relID {
			continue
		}
		if rel.External || rel.Type != base+"image" || rel.PartName == "" {
			return "", "", false
		}
		contentType := extractor.pkg.contentTypes[rel.PartName]
		return rel.PartName, contentType, strings.HasPrefix(nativeASCIIFold(contentType), "image/")
	}
	return "", "", false
}

func nativeExactPageAnchor(node *nativeXMLNode, wpNS, aNS string) bool {
	if !nativeExactContainer(node, xml.Name{Local: "distT"}, xml.Name{Local: "distB"}, xml.Name{Local: "distL"}, xml.Name{Local: "distR"}, xml.Name{Local: "simplePos"}, xml.Name{Local: "relativeHeight"}, xml.Name{Local: "behindDoc"}, xml.Name{Local: "locked"}, xml.Name{Local: "layoutInCell"}, xml.Name{Local: "allowOverlap"}) {
		return false
	}
	for _, name := range []string{"distT", "distB", "distL", "distR"} {
		if value, present := nativeUnqualifiedAttr(node, name); present && value != "0" {
			return false
		}
	}
	for name, want := range map[string]bool{"simplePos": false, "locked": false, "layoutInCell": true, "allowOverlap": true} {
		value, present := nativeUnqualifiedAttr(node, name)
		if !present || (want && value != "1" && value != "true") || (!want && value != "0" && value != "false") {
			return false
		}
	}
	behind, present := nativeUnqualifiedAttr(node, "behindDoc")
	if !present || (behind != "0" && behind != "1" && behind != "true" && behind != "false") {
		return false
	}
	order, ok := nativeNonnegativeInt64Attr(node, "", "relativeHeight")
	if !ok || order > 4294967295 {
		return false
	}
	// Reuse the exact inline payload qualification after removing only the
	// anchor fields whose semantics are explicitly represented below.
	projection := *node
	projection.Attrs = nil
	projection.Children = nil
	seen := map[string]bool{}
	for _, child := range node.Children {
		if child.Name.Space == wpNS && (child.Name.Local == "simplePos" || child.Name.Local == "positionH" || child.Name.Local == "positionV" || child.Name.Local == "wrapNone" || child.Name.Local == "wrapSquare") {
			if seen[child.Name.Local] {
				return false
			}
			seen[child.Name.Local] = true
			switch child.Name.Local {
			case "simplePos":
				if !nativeExactLeaf(child, xml.Name{Local: "x"}, xml.Name{Local: "y"}) {
					return false
				}
				x, xok := nativeUnqualifiedAttr(child, "x")
				y, yok := nativeUnqualifiedAttr(child, "y")
				if !xok || !yok || x != "0" || y != "0" {
					return false
				}
			case "positionH", "positionV":
				value, relative, ok := nativeDrawingPosition(child, wpNS)
				if !ok || relative != "page" || value < 0 {
					return false
				}
			case "wrapNone":
				if !nativeExactLeaf(child) {
					return false
				}
			case "wrapSquare":
				if !nativeExactLeaf(child, xml.Name{Local: "wrapText"}) {
					return false
				}
				value, present := nativeUnqualifiedAttr(child, "wrapText")
				if !present || value != "bothSides" {
					return false
				}
			}
		} else {
			projection.Children = append(projection.Children, child)
		}
	}
	return len(seen) == 4 && nativeExactInlinePictureContainer(&projection, wpNS, aNS)
}

func nativeDrawingPosition(node *nativeXMLNode, wpNS string) (int64, string, bool) {
	relativeFrom, _ := nativeUnqualifiedAttr(node, "relativeFrom")
	position := firstDirectNativeChild(node, wpNS, "posOffset")
	if position == nil || len(node.Children) != 1 || !nativeExactContainer(node, xml.Name{Local: "relativeFrom"}) || len(position.Attrs) != 0 || len(position.Children) != 0 {
		return 0, "", false
	}
	value, err := strconv.ParseInt(strings.TrimSpace(position.Text), 10, 64)
	if err != nil || value < -9007199254740991 || value > 9007199254740991 {
		return 0, "", false
	}
	return value, relativeFrom, true
}

func nativeDrawingWrap(container *nativeXMLNode, wpNS string) (string, bool) {
	wraps := map[string]string{"wrapNone": "none", "wrapSquare": "square", "wrapTight": "tight", "wrapThrough": "through", "wrapTopAndBottom": "top-and-bottom"}
	value := ""
	for _, child := range container.Children {
		if candidate, ok := wraps[child.Name.Local]; ok && child.Name.Space == wpNS {
			if value != "" {
				return "", false
			}
			if len(child.Children) != 0 {
				return "", false
			}
			for _, attr := range child.Attrs {
				if child.Name.Local == "wrapSquare" && attr.Name.Space == "" && attr.Name.Local == "wrapText" && attr.Value == "bothSides" {
					continue
				}
				return "", false
			}
			value = candidate
		}
	}
	return value, value != ""
}

func nativeVerticalAlignmentValue(node *nativeXMLNode, wordNS string) (string, bool) {
	if !nativeExactLeaf(node, xml.Name{Space: wordNS, Local: "val"}) {
		return "", false
	}
	count, value := 0, ""
	for _, attr := range node.Attrs {
		if attr.Name == (xml.Name{Space: wordNS, Local: "val"}) {
			count++
			value = attr.Value
		}
	}
	return value, count == 1 && (value == "baseline" || value == "subscript" || value == "superscript")
}

func (extractor *nativeExtractor) extractRunProperties(partName, paragraphID string, node *nativeXMLNode) (*NativeRunPropertiesV1, bool) {
	properties, unsafe, preserveOnly := extractor.extractRunPropertiesState(partName, paragraphID, node)
	return properties, unsafe || preserveOnly
}

// Keep source-layout invalidity separate from valid but noneditable metadata.
// Paragraph marks already retain a read-only policy even when fully qualified.
func (extractor *nativeExtractor) extractRunPropertiesState(partName, paragraphID string, node *nativeXMLNode) (*NativeRunPropertiesV1, bool, bool) {
	properties := &NativeRunPropertiesV1{}
	unsafe := false
	preserveOnly := false
	for _, child := range node.Children {
		if child.Name.Space != extractor.wordNS {
			unsafe = true
			extractor.addUnsupported("FOREIGN_RUN_PROPERTY", "run-properties", paragraphID, partName, child, "Foreign namespace run property is preserved verbatim")
			continue
		}
		switch child.Name.Local {
		case "rStyle":
			if value, ok := nativeAttr(child, extractor.wordNS, "val"); ok && nativeIDPattern.MatchString(value) {
				properties.CharacterStyleID = nativeString(value)
			} else {
				unsafe = true
			}
		case "rFonts":
			value, ok := nativeAttr(child, extractor.wordNS, "ascii")
			if !ok {
				value, ok = nativeAttr(child, extractor.wordNS, "hAnsi")
			}
			if ok && value != "" {
				properties.FontFamily = nativeString(value)
			} else {
				unsafe = true
			}
		case "sz":
			value, ok := nativePositiveIntAttr(child, extractor.wordNS, "val")
			if ok {
				properties.FontSizeHalfPoint = nativeInt(value)
			} else {
				unsafe = true
			}
		case "b", "i", "rtl", "vanish":
			value, ok := nativeOnOff(child, extractor.wordNS)
			if !ok {
				unsafe = true
				break
			}
			switch child.Name.Local {
			case "b":
				properties.Bold = nativeBool(value)
			case "i":
				properties.Italic = nativeBool(value)
			case "rtl":
				properties.RTL = nativeBool(value)
			case "vanish":
				properties.Hidden = nativeBool(value)
			}
		case "u":
			value, ok := nativeAttr(child, extractor.wordNS, "val")
			if !ok {
				value, ok = "single", true
			}
			if value == "none" || value == "single" || value == "double" || value == "words" {
				properties.Underline = nativeString(value)
			} else {
				unsafe = true
			}
		case "color":
			if rgb, ok := extractor.extractExactColorRGB(child); ok {
				properties.Color = nativeString(rgb)
			} else {
				unsafe = true
			}
		case "highlight":
			if value, ok := nativeAttr(child, extractor.wordNS, "val"); ok && value != "" {
				properties.Highlight = nativeString(value)
			} else {
				unsafe = true
			}
		case "lang":
			value, ok := nativeAttr(child, extractor.wordNS, "val")
			if !ok {
				value, ok = nativeAttr(child, extractor.wordNS, "eastAsia")
			}
			if ok && value != "" {
				properties.Language = nativeString(value)
			} else {
				unsafe = true
			}
		case "kern":
			preserveOnly = true
			if _, ok := nativeKerningThreshold(child, extractor.wordNS); !ok || len(directNativeChildren(node, extractor.wordNS, "kern")) != 1 {
				unsafe = true
				extractor.addUnsupported("UNMODELED_RUN_PROPERTY", "run-properties", paragraphID, partName, child, "Kerning threshold is malformed, duplicate or outside the bounded whole half-point subset")
			}
		case "szCs":
			// The resolved source context, not extraction, determines whether
			// this preserved script slot is inactive. Never expose it for edits.
			preserveOnly = true
			value, ok := nativePositiveIntAttr(child, extractor.wordNS, "val")
			if !ok || value > 3276 || !nativeExactLeaf(child, xml.Name{Space: extractor.wordNS, Local: "val"}) || len(directNativeChildren(node, extractor.wordNS, "szCs")) != 1 {
				unsafe = true
				extractor.addUnsupported("UNMODELED_RUN_PROPERTY", "run-properties", paragraphID, partName, child, "Complex-script size is malformed, duplicate or outside the bounded whole half-point subset")
			}
		case "noProof":
			preserveOnly = true
			if !nativeNeutralSourceProperty(child, node, extractor.wordNS) {
				unsafe = true
				extractor.addUnsupported("UNMODELED_RUN_PROPERTY", "run-properties", paragraphID, partName, child, "Proofing metadata has malformed, duplicate or unknown source structure")
			}
		case "vertAlign":
			preserveOnly = true
			value, ok := nativeVerticalAlignmentValue(child, extractor.wordNS)
			if ok {
				properties.VerticalAlignment = nativeString(value)
			} else {
				unsafe = true
				extractor.addUnsupported("VERTICAL_ALIGNMENT_UNSUPPORTED", "run-properties", paragraphID, partName, child, "Vertical alignment requires an exact baseline, subscript or superscript value")
			}
		default:
			unsafe = true
			extractor.addUnsupported("UNMODELED_RUN_PROPERTY", "run-properties", paragraphID, partName, child, "This run property is preserved verbatim")
		}
	}
	if unsafe {
		extractor.addUnsupported("PARTIAL_RUN_PROPERTIES", "run-properties", paragraphID, partName, node, "Only the conservative v1 run-property subset is exposed")
	}
	return properties, unsafe, preserveOnly
}

func nativeOnOff(node *nativeXMLNode, namespace string) (bool, bool) {
	return nativeOnOffAttr(node, namespace, "val", true)
}

func nativeOnOffAttr(node *nativeXMLNode, namespace, local string, defaultValue bool) (bool, bool) {
	value, ok := nativeAttr(node, namespace, local)
	if !ok {
		return defaultValue, true
	}
	switch value {
	case "true", "1", "on":
		return true, true
	case "false", "0", "off":
		return false, true
	default:
		return false, false
	}
}

func nativePositiveIntAttr(node *nativeXMLNode, namespace, local string) (int, bool) {
	raw, ok := nativeAttr(node, namespace, local)
	if !ok {
		return 0, false
	}
	value, err := strconv.Atoi(raw)
	return value, err == nil && value > 0
}

func nativeNonnegativeInt64Attr(node *nativeXMLNode, namespace, local string) (int64, bool) {
	raw, ok := nativeAttr(node, namespace, local)
	if !ok {
		return 0, false
	}
	value, err := strconv.ParseInt(raw, 10, 64)
	return value, err == nil && value >= 0 && value <= 9007199254740991
}

func (extractor *nativeExtractor) extractExactColorRGB(node *nativeXMLNode) (string, bool) {
	if !nativeWordColorLeaf(node, extractor.wordNS) {
		return "", false
	}
	if _, tint := nativeAttr(node, extractor.wordNS, "themeTint"); tint {
		return "", false
	}
	if _, shade := nativeAttr(node, extractor.wordNS, "themeShade"); shade {
		return "", false
	}
	theme, hasTheme := nativeAttr(node, extractor.wordNS, "themeColor")
	if hasTheme {
		return extractor.resolveThemeSrgb(theme)
	}
	value, ok := nativeAttr(node, extractor.wordNS, "val")
	if !ok {
		return "", false
	}
	return nativeExactRGB(value)
}

func nativeZeroOrAbsentTwipAttr(node *nativeXMLNode, namespace, local string) bool {
	value, ok := nativeAttr(node, namespace, local)
	return !ok || value == "0"
}

func nativeResolveExactThemeOrRGB(node *nativeXMLNode, wordNS string, resolveTheme func(string) (string, bool)) (string, bool) {
	theme, hasTheme := nativeAttr(node, wordNS, "themeColor")
	if hasTheme {
		return resolveTheme(theme)
	}
	color, ok := nativeAttr(node, wordNS, "color")
	if !ok {
		return "", false
	}
	return nativeExactRGB(color)
}

func nativeExtractTableBorder(node *nativeXMLNode, wordNS string, resolveTheme func(string) (string, bool)) (*NativeTableBorderV1, bool) {
	value, ok := nativeAttr(node, wordNS, "val")
	if !ok || (value != "none" && value != "nil" && value != "single") {
		return nil, false
	}
	if !nativeExactLeaf(node,
		xml.Name{Space: wordNS, Local: "val"},
		xml.Name{Space: wordNS, Local: "sz"},
		xml.Name{Space: wordNS, Local: "color"},
		xml.Name{Space: wordNS, Local: "space"},
		xml.Name{Space: wordNS, Local: "themeColor"},
		xml.Name{Space: wordNS, Local: "themeTint"},
		xml.Name{Space: wordNS, Local: "themeShade"},
	) {
		return nil, false
	}
	if !nativeZeroOrAbsentTwipAttr(node, wordNS, "space") {
		return nil, false
	}
	if _, tint := nativeAttr(node, wordNS, "themeTint"); tint {
		return nil, false
	}
	if _, shade := nativeAttr(node, wordNS, "themeShade"); shade {
		return nil, false
	}
	if value == "none" || value == "nil" {
		if size, present := nativeNonnegativeInt64Attr(node, wordNS, "sz"); present && size != 0 {
			return nil, false
		}
		if color, present := nativeAttr(node, wordNS, "color"); present && !strings.EqualFold(color, "auto") {
			return nil, false
		}
		if _, theme := nativeAttr(node, wordNS, "themeColor"); theme {
			return nil, false
		}
		return &NativeTableBorderV1{Style: "none", SizeEighthPoints: 0}, true
	}
	size, sizeOK := nativeNonnegativeInt64Attr(node, wordNS, "sz")
	if !sizeOK || size <= 0 {
		return nil, false
	}
	rgb, ok := nativeResolveExactThemeOrRGB(node, wordNS, resolveTheme)
	if !ok {
		return nil, false
	}
	return &NativeTableBorderV1{Style: "single", SizeEighthPoints: size, ColorRGB: nativeString(rgb)}, true
}

func nativeExtractTableBorders(node *nativeXMLNode, wordNS string, resolveTheme func(string) (string, bool)) (*NativeTableBordersV1, bool) {
	if !nativeExactContainer(node) {
		return nil, false
	}
	result := &NativeTableBordersV1{}
	seen := map[string]bool{}
	for _, child := range node.Children {
		field := child.Name.Local
		if child.Name.Space != wordNS || seen[field] {
			return nil, false
		}
		border, ok := nativeExtractTableBorder(child, wordNS, resolveTheme)
		if !ok {
			return nil, false
		}
		seen[field] = true
		switch field {
		case "top":
			result.Top = border
		case "right":
			result.Right = border
		case "bottom":
			result.Bottom = border
		case "left":
			result.Left = border
		case "insideH":
			result.InsideHorizontal = border
		case "insideV":
			result.InsideVertical = border
		default:
			return nil, false
		}
	}
	return result, true
}

func nativeExtractCellShading(node *nativeXMLNode, wordNS string, resolveTheme func(string) (string, bool)) (*string, bool) {
	if !nativeExactLeaf(node,
		xml.Name{Space: wordNS, Local: "val"},
		xml.Name{Space: wordNS, Local: "color"},
		xml.Name{Space: wordNS, Local: "fill"},
		xml.Name{Space: wordNS, Local: "themeColor"},
		xml.Name{Space: wordNS, Local: "themeTint"},
		xml.Name{Space: wordNS, Local: "themeShade"},
		xml.Name{Space: wordNS, Local: "themeFill"},
		xml.Name{Space: wordNS, Local: "themeFillTint"},
		xml.Name{Space: wordNS, Local: "themeFillShade"},
	) {
		return nil, false
	}
	value, ok := nativeAttr(node, wordNS, "val")
	if !ok || value != "clear" {
		return nil, false
	}
	for _, attr := range []string{"themeTint", "themeShade", "themeFillTint", "themeFillShade", "themeColor"} {
		if _, present := nativeAttr(node, wordNS, attr); present {
			return nil, false
		}
	}
	if color, present := nativeAttr(node, wordNS, "color"); present && !strings.EqualFold(color, "auto") {
		return nil, false
	}
	if themeFill, present := nativeAttr(node, wordNS, "themeFill"); present {
		rgb, ok := resolveTheme(themeFill)
		if !ok {
			return nil, false
		}
		return nativeString(rgb), true
	}
	fill, fillOK := nativeAttr(node, wordNS, "fill")
	rgb, ok := nativeExactRGB(fill)
	if !fillOK || !ok {
		return nil, false
	}
	return nativeString(rgb), true
}

func (extractor *nativeExtractor) extractTableBorder(node *nativeXMLNode) (*NativeTableBorderV1, bool) {
	return nativeExtractTableBorder(node, extractor.wordNS, extractor.resolveThemeSrgb)
}

func (extractor *nativeExtractor) resolveExactThemeOrRGB(node *nativeXMLNode) (string, bool) {
	return nativeResolveExactThemeOrRGB(node, extractor.wordNS, extractor.resolveThemeSrgb)
}

func (extractor *nativeExtractor) extractTableBorders(node *nativeXMLNode) (*NativeTableBordersV1, bool) {
	return nativeExtractTableBorders(node, extractor.wordNS, extractor.resolveThemeSrgb)
}

func (extractor *nativeExtractor) extractCellShading(node *nativeXMLNode) (*string, bool) {
	return nativeExtractCellShading(node, extractor.wordNS, extractor.resolveThemeSrgb)
}

func (extractor *nativeExtractor) extractTableCellMargins(node *nativeXMLNode) (*NativeTableCellMarginsV1, bool) {
	if !nativeExactContainer(node) {
		return nil, false
	}
	values := map[string]int64{}
	for _, child := range node.Children {
		if child.Name.Space != extractor.wordNS || (child.Name.Local != "top" && child.Name.Local != "right" && child.Name.Local != "bottom" && child.Name.Local != "left") {
			return nil, false
		}
		if _, duplicate := values[child.Name.Local]; duplicate {
			return nil, false
		}
		width, ok := nativeNonnegativeInt64Attr(child, extractor.wordNS, "w")
		typeValue, hasType := nativeAttr(child, extractor.wordNS, "type")
		if !ok || (hasType && typeValue != "dxa") || !nativeExactLeaf(child, xml.Name{Space: extractor.wordNS, Local: "w"}, xml.Name{Space: extractor.wordNS, Local: "type"}) {
			return nil, false
		}
		values[child.Name.Local] = width
	}
	if len(values) != 4 {
		return nil, false
	}
	return &NativeTableCellMarginsV1{TopTwips: values["top"], RightTwips: values["right"], BottomTwips: values["bottom"], LeftTwips: values["left"]}, true
}

func (extractor *nativeExtractor) extractTable(partName string, node *nativeXMLNode) (NativeTableV1, error) {
	id := extractor.objectID("table", partName, node, "")
	table := NativeTableV1{ID: id, Anchor: extractor.anchor(partName, node), EditPolicy: nativeExtractOnlyTablePolicy(), GridWidthsTwips: []int64{}, Rows: []NativeTableRowV1{}}
	unsafe := false
	if tblPr := firstDirectNativeChild(node, extractor.wordNS, "tblPr"); tblPr != nil {
		if !nativeExactContainer(tblPr) {
			unsafe = true
		}
		seen := map[string]bool{}
		for _, property := range tblPr.Children {
			if seen[property.Name.Local] {
				unsafe = true
				extractor.addUnsupported("UNMODELED_TABLE_PROPERTY", "table-properties", id, partName, property, "Duplicate table properties are ambiguous and preserved verbatim")
				continue
			}
			seen[property.Name.Local] = true
			if property.Name == (xml.Name{Space: extractor.wordNS, Local: "tblStyle"}) {
				if value, ok := nativeAttr(property, extractor.wordNS, "val"); ok && nativeIDPattern.MatchString(value) {
					table.TableStyleID = nativeString(value)
				} else {
					unsafe = true
				}
			} else if property.Name == (xml.Name{Space: extractor.wordNS, Local: "tblLook"}) {
				// Conditional-style switches have no visual effect only when
				// the complete referenced style chain has no conditional layers.
				// The source remains read-only even for that qualified case.
				unsafe = true
				if !extractor.inactiveTableLook(tblPr, property) {
					extractor.addUnsupported("UNMODELED_TABLE_PROPERTY", "table-properties", id, partName, property, "Table look has active, ambiguous, or unqualified conditional-style semantics")
				}
			} else if property.Name == (xml.Name{Space: extractor.wordNS, Local: "tblW"}) {
				width, widthOK := nativeNonnegativeInt64Attr(property, extractor.wordNS, "w")
				typeValue, typeOK := nativeAttr(property, extractor.wordNS, "type")
				if widthOK && width > 0 && typeOK && typeValue == "dxa" && nativeExactLeaf(property, xml.Name{Space: extractor.wordNS, Local: "w"}, xml.Name{Space: extractor.wordNS, Local: "type"}) {
					table.WidthTwips = nativeInt64(width)
				} else if widthOK && width == 0 && typeOK && typeValue == "auto" && nativeExactLeaf(property, xml.Name{Space: extractor.wordNS, Local: "w"}, xml.Name{Space: extractor.wordNS, Local: "type"}) {
					// Auto width has no absolute preferred extent; the explicit
					// autofit layout is qualified by font-shaped content later.
				} else if widthOK && width > 0 && width <= 5000 && typeOK && typeValue == "pct" && nativeExactLeaf(property, xml.Name{Space: extractor.wordNS, Local: "w"}, xml.Name{Space: extractor.wordNS, Local: "type"}) {
					table.WidthPercentFiftieths = nativeInt64(width)
				} else {
					unsafe = true
				}
			} else if property.Name == (xml.Name{Space: extractor.wordNS, Local: "tblLayout"}) {
				value, ok := nativeAttr(property, extractor.wordNS, "type")
				if ok && (value == "fixed" || value == "autofit") && nativeExactLeaf(property, xml.Name{Space: extractor.wordNS, Local: "type"}) {
					table.Layout = nativeString(value)
				} else {
					unsafe = true
				}
			} else if property.Name == (xml.Name{Space: extractor.wordNS, Local: "jc"}) {
				value, ok := nativeAttr(property, extractor.wordNS, "val")
				if ok && value == "left" && nativeExactLeaf(property, xml.Name{Space: extractor.wordNS, Local: "val"}) {
					table.Alignment = nativeString(value)
				} else {
					unsafe = true
				}
			} else if property.Name == (xml.Name{Space: extractor.wordNS, Local: "tblInd"}) {
				width, widthOK := nativeNonnegativeInt64Attr(property, extractor.wordNS, "w")
				typeValue, typeOK := nativeAttr(property, extractor.wordNS, "type")
				if widthOK && typeOK && typeValue == "dxa" && nativeExactLeaf(property, xml.Name{Space: extractor.wordNS, Local: "w"}, xml.Name{Space: extractor.wordNS, Local: "type"}) {
					table.IndentTwips = nativeInt64(width)
				} else {
					unsafe = true
				}
			} else if property.Name == (xml.Name{Space: extractor.wordNS, Local: "tblCellMar"}) {
				if margins, ok := extractor.extractTableCellMargins(property); ok {
					table.CellMargins = margins
				} else {
					unsafe = true
				}
			} else if property.Name == (xml.Name{Space: extractor.wordNS, Local: "tblBorders"}) {
				if borders, ok := extractor.extractTableBorders(property); ok {
					table.Borders = borders
				} else {
					unsafe = true
					extractor.addUnsupported("UNMODELED_TABLE_PROPERTY", "table-properties", id, partName, property, "Table borders do not have an exact supported source color and structure")
				}
			} else {
				unsafe = true
				extractor.addUnsupported("UNMODELED_TABLE_PROPERTY", "table-properties", id, partName, property, "This table property is preserved verbatim")
			}
		}
	}
	for _, child := range node.Children {
		if child.Name == (xml.Name{Space: extractor.wordNS, Local: "tblPr"}) || child.Name == (xml.Name{Space: extractor.wordNS, Local: "tblGrid"}) {
			if child.Name.Local == "tblGrid" {
				if !nativeExactContainer(child) {
					unsafe = true
				}
				for _, column := range child.Children {
					width, ok := nativeNonnegativeInt64Attr(column, extractor.wordNS, "w")
					if column.Name != (xml.Name{Space: extractor.wordNS, Local: "gridCol"}) || !ok || width <= 0 || !nativeExactLeaf(column, xml.Name{Space: extractor.wordNS, Local: "w"}) {
						unsafe = true
						continue
					}
					table.GridWidthsTwips = append(table.GridWidthsTwips, width)
				}
			}
			continue
		}
		if child.Name != (xml.Name{Space: extractor.wordNS, Local: "tr"}) {
			unsafe = true
			extractor.addUnsupported("UNMODELED_TABLE_CONTENT", "table-structure", id, partName, child, "Table content outside direct rows is preserved verbatim")
			continue
		}
		row, rowUnsafe, err := extractor.extractTableRow(partName, id, child)
		if err != nil {
			return NativeTableV1{}, err
		}
		table.Rows = append(table.Rows, row)
		unsafe = unsafe || rowUnsafe
		if len(table.Rows) > NativeDOCXMaxCollectionItems {
			return NativeTableV1{}, fmt.Errorf("docxpatch: native extract: table %q exceeds %d rows", id, NativeDOCXMaxCollectionItems)
		}
	}
	if unsafe {
		table.EditPolicy = nativeReadOnlyPolicy("UNMODELED_TABLE_MARKUP", "Table contains OOXML that the v1 extractor preserves but cannot safely mutate")
	}
	return table, nil
}

func (extractor *nativeExtractor) extractTableRow(partName, tableID string, node *nativeXMLNode) (NativeTableRowV1, bool, error) {
	id := extractor.objectID("row", partName, node, "")
	row := NativeTableRowV1{ID: id, Anchor: extractor.anchor(partName, node), RepeatHeader: nativeBool(false), CantSplit: nativeBool(false), Cells: []NativeTableCellV1{}}
	unsafe := false
	if trPr := firstDirectNativeChild(node, extractor.wordNS, "trPr"); trPr != nil {
		if !nativeExactContainer(trPr) {
			unsafe = true
		}
		seen := map[string]bool{}
		for _, property := range trPr.Children {
			if seen[property.Name.Local] {
				unsafe = true
				extractor.addUnsupported("UNMODELED_ROW_PROPERTY", "table-properties", tableID, partName, property, "Duplicate row properties are ambiguous and preserved verbatim")
				continue
			}
			seen[property.Name.Local] = true
			switch {
			case property.Name == (xml.Name{Space: extractor.wordNS, Local: "trHeight"}):
				if !nativeExactLeaf(property, xml.Name{Space: extractor.wordNS, Local: "val"}, xml.Name{Space: extractor.wordNS, Local: "hRule"}) {
					unsafe = true
					break
				}
				rule, hasRule := nativeAttr(property, extractor.wordNS, "hRule")
				if !hasRule || rule == "auto" {
					// ST_HeightRule default is auto: val is not a layout constraint.
					break
				}
				value, ok := nativeNonnegativeInt64Attr(property, extractor.wordNS, "val")
				if !ok || (rule != "atLeast" && rule != "exact") {
					unsafe = true
					break
				}
				row.HeightTwips = nativeInt64(value)
				row.HeightRule = nativeString(rule)
			case property.Name == (xml.Name{Space: extractor.wordNS, Local: "tblHeader"}):
				if value, ok := nativeOnOff(property, extractor.wordNS); ok {
					row.RepeatHeader = nativeBool(value)
				} else {
					unsafe = true
				}
			case property.Name == (xml.Name{Space: extractor.wordNS, Local: "cantSplit"}):
				if value, ok := nativeOnOff(property, extractor.wordNS); ok {
					row.CantSplit = nativeBool(value)
				} else {
					unsafe = true
				}
			default:
				unsafe = true
				extractor.addUnsupported("UNMODELED_ROW_PROPERTY", "table-properties", tableID, partName, property, "This table-row property is preserved verbatim")
			}
		}
	}
	for _, child := range node.Children {
		if child.Name == (xml.Name{Space: extractor.wordNS, Local: "trPr"}) {
			continue
		}
		if child.Name != (xml.Name{Space: extractor.wordNS, Local: "tc"}) {
			unsafe = true
			extractor.addUnsupported("UNMODELED_ROW_CONTENT", "table-structure", tableID, partName, child, "Row content outside direct cells is preserved verbatim")
			continue
		}
		cell, cellUnsafe, err := extractor.extractTableCell(partName, tableID, child)
		if err != nil {
			return NativeTableRowV1{}, false, err
		}
		row.Cells = append(row.Cells, cell)
		unsafe = unsafe || cellUnsafe
		if len(row.Cells) > NativeDOCXMaxCollectionItems {
			return NativeTableRowV1{}, false, fmt.Errorf("docxpatch: native extract: row %q exceeds %d cells", id, NativeDOCXMaxCollectionItems)
		}
	}
	return row, unsafe, nil
}

func (extractor *nativeExtractor) extractTableCell(partName, tableID string, node *nativeXMLNode) (NativeTableCellV1, bool, error) {
	id := extractor.objectID("cell", partName, node, "")
	cell := NativeTableCellV1{ID: id, Anchor: extractor.anchor(partName, node), GridSpan: nativeInt(1), VerticalMerge: "none", Paragraphs: []NativeParagraphV1{}}
	unsafe := false
	if tcPr := firstDirectNativeChild(node, extractor.wordNS, "tcPr"); tcPr != nil {
		if !nativeExactContainer(tcPr) {
			unsafe = true
		}
		seen := map[string]bool{}
		for _, property := range tcPr.Children {
			if seen[property.Name.Local] {
				unsafe = true
				extractor.addUnsupported("UNMODELED_CELL_PROPERTY", "table-properties", tableID, partName, property, "Duplicate cell properties are ambiguous and preserved verbatim")
				continue
			}
			seen[property.Name.Local] = true
			switch {
			case property.Name == (xml.Name{Space: extractor.wordNS, Local: "tcW"}):
				value, widthOK := nativeNonnegativeInt64Attr(property, extractor.wordNS, "w")
				typeValue, typeOK := nativeAttr(property, extractor.wordNS, "type")
				// CT_TblWidth defaults an omitted w:type to dxa. Preserve the
				// exact width emitted by Word-compatible producers that rely on
				// that schema default instead of treating it as unmodeled markup.
				if widthOK && (!typeOK || typeValue == "dxa") && nativeExactLeaf(property, xml.Name{Space: extractor.wordNS, Local: "w"}, xml.Name{Space: extractor.wordNS, Local: "type"}) {
					cell.WidthTwips = nativeInt64(value)
				} else if widthOK && value == 0 && typeOK && typeValue == "auto" && nativeExactLeaf(property, xml.Name{Space: extractor.wordNS, Local: "w"}, xml.Name{Space: extractor.wordNS, Local: "type"}) {
					// No absolute preferred cell width.
				} else {
					unsafe = true
				}
			case property.Name == (xml.Name{Space: extractor.wordNS, Local: "gridSpan"}):
				if value, ok := nativePositiveIntAttr(property, extractor.wordNS, "val"); ok {
					cell.GridSpan = nativeInt(value)
				} else {
					unsafe = true
				}
			case property.Name == (xml.Name{Space: extractor.wordNS, Local: "vMerge"}):
				value, ok := nativeAttr(property, extractor.wordNS, "val")
				if !ok {
					value = "continue"
				}
				if value == "restart" || value == "continue" {
					cell.VerticalMerge = value
				} else {
					unsafe = true
				}
			case property.Name == (xml.Name{Space: extractor.wordNS, Local: "tcBorders"}):
				if borders, ok := extractor.extractTableBorders(property); ok {
					cell.Borders = borders
				} else {
					unsafe = true
				}
			case property.Name == (xml.Name{Space: extractor.wordNS, Local: "shd"}):
				if shading, ok := extractor.extractCellShading(property); ok {
					cell.ShadingRGB = shading
				} else {
					unsafe = true
				}
			default:
				unsafe = true
				extractor.addUnsupported("UNMODELED_CELL_PROPERTY", "table-properties", tableID, partName, property, "This table-cell property is preserved verbatim")
			}
		}
	}
	for _, child := range node.Children {
		if child.Name == (xml.Name{Space: extractor.wordNS, Local: "tcPr"}) {
			continue
		}
		if child.Name != (xml.Name{Space: extractor.wordNS, Local: "p"}) {
			unsafe = true
			extractor.addUnsupported("NESTED_TABLE_OR_CELL_MARKUP", "table-structure", tableID, partName, child, "Only direct cell paragraphs are modeled; nested content is preserved verbatim")
			continue
		}
		paragraph, err := extractor.extractParagraph(partName, child)
		if err != nil {
			return NativeTableCellV1{}, false, err
		}
		cell.Paragraphs = append(cell.Paragraphs, paragraph)
	}
	return cell, unsafe, nil
}

func (extractor *nativeExtractor) extractSection(node *nativeXMLNode, startsAtBlockID string) (NativeSectionV1, error) {
	section := extractor.defaultSection(node, startsAtBlockID)
	id := section.ID
	if !nativeExactRevisionContainer(node, extractor.wordNS, "rsidRPr", "rsidDel", "rsidR", "rsidSect") {
		extractor.addUnsupported("UNMODELED_SECTION_PROPERTY", "sections", id, extractor.mainPart, node, "Section-properties container has attributes or direct text outside the exact v1 subset")
	}
	seenRefs := map[string]bool{}
	seenSingleton := map[string]bool{}
	for _, child := range node.Children {
		if child.Name.Space != extractor.wordNS {
			extractor.addUnsupported("FOREIGN_SECTION_MARKUP", "sections", id, extractor.mainPart, child, "Foreign section markup is preserved verbatim")
			continue
		}
		if child.Name.Local == "type" || child.Name.Local == "titlePg" || child.Name.Local == "pgNumType" || child.Name.Local == "pgSz" || child.Name.Local == "pgMar" || child.Name.Local == "cols" || child.Name.Local == "docGrid" || child.Name.Local == "formProt" || child.Name.Local == "noEndnote" {
			if seenSingleton[child.Name.Local] {
				extractor.addUnsupported("DUPLICATE_SECTION_PROPERTY", "sections", id, extractor.mainPart, child, "Duplicate modeled section-property singletons make exact pagination geometry ambiguous")
				continue
			}
			seenSingleton[child.Name.Local] = true
		}
		switch child.Name.Local {
		case "formProt":
			value, present := nativeAttr(child, extractor.wordNS, "val")
			enabled, valid := nativeOnOff(child, extractor.wordNS)
			if !present || value == "" || !valid || enabled || !nativeExactLeaf(child, xml.Name{Space: extractor.wordNS, Local: "val"}) {
				extractor.addUnsupported("UNMODELED_SECTION_PROPERTY", "sections", id, extractor.mainPart, child, "Only exact explicitly disabled section form protection is layout-neutral")
			}
		case "noEndnote":
			_, valid := nativeOnOff(child, extractor.wordNS)
			if !valid || !nativeExactLeaf(child, xml.Name{Space: extractor.wordNS, Local: "val"}) || !extractor.proveNoContentEndnotes() {
				extractor.addUnsupported("UNMODELED_SECTION_PROPERTY", "sections", id, extractor.mainPart, child, "Endnote placement remains unqualified unless the package proves no content endnotes or references")
			}
		case "docGrid":
			if !nativeInactiveSectionGrid(child, extractor.wordNS) {
				extractor.addUnsupported("UNMODELED_SECTION_PROPERTY", "sections", id, extractor.mainPart, child, "Active or unqualified document-grid markup remains unsupported")
			}
		case "type":
			if !nativeExactLeaf(child, xml.Name{Space: extractor.wordNS, Local: "val"}) {
				extractor.addUnsupported("UNMODELED_SECTION_PROPERTY", "sections", id, extractor.mainPart, child, "Section break markup has attributes or children outside the exact v1 subset")
				continue
			}
			value, _ := nativeAttr(child, extractor.wordNS, "val")
			switch value {
			case "continuous", "evenPage", "oddPage", "nextColumn", "nextPage":
				section.BreakType = map[string]string{"continuous": "continuous", "evenPage": "even-page", "oddPage": "odd-page", "nextColumn": "next-column", "nextPage": "next-page"}[value]
			default:
				extractor.addUnsupported("UNMODELED_SECTION_BREAK", "sections", id, extractor.mainPart, child, "Missing or unknown section break type is preserved; next-page is exposed conservatively")
			}
		case "pgNumType":
			format, _ := nativeAttr(child, extractor.wordNS, "fmt")
			start, hasStart := nativeAttr(child, extractor.wordNS, "start")
			valid := nativeExactLeaf(child, xml.Name{Space: extractor.wordNS, Local: "fmt"}, xml.Name{Space: extractor.wordNS, Local: "start"}) && (format == "" || format == "decimal")
			var number int64
			if hasStart {
				var err error
				number, err = strconv.ParseInt(start, 10, 64)
				valid = valid && err == nil && number >= 0 && number <= 999999 && strconv.FormatInt(number, 10) == start
			}
			if !valid {
				extractor.addUnsupported("UNMODELED_SECTION_PROPERTY", "sections", id, extractor.mainPart, child, "Page numbering requires bounded decimal start and no chapter/switch attributes")
				continue
			}
			if hasStart {
				section.PageNumberStart = &number
			}
		case "titlePg":
			if !nativeExactLeaf(child, xml.Name{Space: extractor.wordNS, Local: "val"}) {
				extractor.addUnsupported("UNMODELED_SECTION_PROPERTY", "sections", id, extractor.mainPart, child, "Title-page markup has attributes or children outside the exact v1 subset")
				continue
			}
			value, ok := nativeOnOff(child, extractor.wordNS)
			if !ok {
				extractor.addUnsupported("UNMODELED_TITLE_PAGE", "sections", id, extractor.mainPart, child, "Title-page policy has an invalid lexical value")
				continue
			}
			section.TitlePage = nativeBool(value)
		case "pgSz":
			if !nativeExactLeaf(child, xml.Name{Space: extractor.wordNS, Local: "w"}, xml.Name{Space: extractor.wordNS, Local: "h"}, xml.Name{Space: extractor.wordNS, Local: "orient"}, xml.Name{Space: extractor.wordNS, Local: "code"}) {
				extractor.addUnsupported("UNMODELED_SECTION_PROPERTY", "sections", id, extractor.mainPart, child, "Page-size markup has attributes or children outside the exact v1 subset")
				continue
			}
			// MS-OE376 2.1.219: code selects printer paper; w/h remain the
			// authored page geometry. Retain the source code, never derive size
			// from a platform-specific printer table. Word bounds code to 0..118.
			if _, present := nativeAttr(child, extractor.wordNS, "code"); present {
				if code, ok := nativeNonnegativeInt64Attr(child, extractor.wordNS, "code"); !ok || code > 118 {
					extractor.addUnsupported("UNMODELED_SECTION_PROPERTY", "sections", id, extractor.mainPart, child, "Printer paper code is outside the qualified Word subset")
					continue
				}
			}
			if value, ok := nativeNonnegativeInt64Attr(child, extractor.wordNS, "w"); ok && value > 0 {
				section.Page.WidthTwips = nativeInt64(value)
			} else {
				extractor.addUnsupported("INVALID_PAGE_SIZE", "sections", id, extractor.mainPart, child, "Invalid page width is preserved; the Word default is exposed")
			}
			if value, ok := nativeNonnegativeInt64Attr(child, extractor.wordNS, "h"); ok && value > 0 {
				section.Page.HeightTwips = nativeInt64(value)
			} else {
				extractor.addUnsupported("INVALID_PAGE_SIZE", "sections", id, extractor.mainPart, child, "Invalid page height is preserved; the Word default is exposed")
			}
			if value, ok := nativeAttr(child, extractor.wordNS, "orient"); ok {
				if value == "portrait" || value == "landscape" {
					section.Page.Orientation = value
				} else {
					extractor.addUnsupported("UNMODELED_PAGE_ORIENTATION", "sections", id, extractor.mainPart, child, "Unknown page orientation is preserved; portrait is exposed")
				}
			}
		case "pgMar":
			if !nativeExactLeaf(child,
				xml.Name{Space: extractor.wordNS, Local: "top"}, xml.Name{Space: extractor.wordNS, Local: "right"},
				xml.Name{Space: extractor.wordNS, Local: "bottom"}, xml.Name{Space: extractor.wordNS, Local: "left"},
				xml.Name{Space: extractor.wordNS, Local: "header"}, xml.Name{Space: extractor.wordNS, Local: "footer"},
				xml.Name{Space: extractor.wordNS, Local: "gutter"}) {
				extractor.addUnsupported("UNMODELED_SECTION_PROPERTY", "sections", id, extractor.mainPart, child, "Page-margin markup has attributes or children outside the exact v1 subset")
				continue
			}
			marginTargets := []struct {
				name   string
				target **int64
			}{{"top", &section.Page.Margins.TopTwips}, {"right", &section.Page.Margins.RightTwips}, {"bottom", &section.Page.Margins.BottomTwips}, {"left", &section.Page.Margins.LeftTwips}, {"header", &section.Page.Margins.HeaderTwips}, {"footer", &section.Page.Margins.FooterTwips}, {"gutter", &section.Page.Margins.GutterTwips}}
			for _, margin := range marginTargets {
				if _, present := nativeAttr(child, extractor.wordNS, margin.name); !present {
					extractor.addUnsupported("MISSING_PAGE_MARGIN", "sections", id, extractor.mainPart, child, "Exact section geometry requires every modeled page-margin attribute")
					continue
				}
				if value, ok := nativeNonnegativeInt64Attr(child, extractor.wordNS, margin.name); ok {
					*margin.target = nativeInt64(value)
				} else {
					extractor.addUnsupported("UNMODELED_PAGE_MARGIN", "sections", id, extractor.mainPart, child, "Signed or invalid page margin is preserved; the Word default is exposed")
				}
			}
		case "cols":
			if !nativeExactContainer(child,
				xml.Name{Space: extractor.wordNS, Local: "num"}, xml.Name{Space: extractor.wordNS, Local: "space"},
				xml.Name{Space: extractor.wordNS, Local: "equalWidth"}, xml.Name{Space: extractor.wordNS, Local: "sep"}) {
				extractor.addUnsupported("MALFORMED_SECTION_COLUMNS", "sections", id, extractor.mainPart, child, "Column markup has attributes or text outside the exact WordprocessingML subset")
				continue
			}
			equalWidth := true
			if _, present := nativeAttr(child, extractor.wordNS, "equalWidth"); present {
				var valid bool
				equalWidth, valid = nativeOnOffAttr(child, extractor.wordNS, "equalWidth", true)
				if !valid {
					extractor.addUnsupported("INVALID_SECTION_COLUMNS", "sections", id, extractor.mainPart, child, "Invalid equalWidth value makes the column model ambiguous")
					continue
				}
			}
			if _, present := nativeAttr(child, extractor.wordNS, "sep"); present {
				separator, valid := nativeOnOffAttr(child, extractor.wordNS, "sep", false)
				if !valid || separator {
					extractor.addUnsupported("COLUMN_SEPARATOR_UNSUPPORTED", "sections", id, extractor.mainPart, child, "Column separators require paint geometry that v1 does not model")
				}
			}
			children := directNativeChildren(child, extractor.wordNS, "col")
			if len(children) != len(child.Children) {
				extractor.addUnsupported("MALFORMED_SECTION_COLUMNS", "sections", id, extractor.mainPart, child, "Only exact w:col children are allowed in column definitions")
			}
			if equalWidth {
				if len(children) != 0 {
					extractor.addUnsupported("CONFLICTING_COLUMN_DEFINITIONS", "sections", id, extractor.mainPart, child, "Equal-width columns cannot also carry explicit w:col geometry")
				}
				count := 1
				if _, present := nativeAttr(child, extractor.wordNS, "num"); present {
					value, ok := nativePositiveIntAttr(child, extractor.wordNS, "num")
					if !ok || value > 45 {
						extractor.addUnsupported("INVALID_SECTION_COLUMNS", "sections", id, extractor.mainPart, child, "Word equal-width column count must be from 1 through 45")
					} else {
						count = value
					}
				}
				spacing := int64(720)
				if _, present := nativeAttr(child, extractor.wordNS, "space"); present {
					value, ok := nativeNonnegativeInt64Attr(child, extractor.wordNS, "space")
					if !ok {
						extractor.addUnsupported("INVALID_COLUMN_SPACING", "sections", id, extractor.mainPart, child, "Invalid equal-width column spacing makes geometry ambiguous")
					} else {
						spacing = value
					}
				} else if count > 1 {
					extractor.addUnsupported("AMBIGUOUS_COLUMN_SPACING", "sections", id, extractor.mainPart, child, "Multi-column equal-width geometry requires explicit spacing for exact pagination")
				}
				section.Page.ColumnLayout = "equal-width"
				section.Page.Columns = nativeInt(count)
				section.Page.ColumnSpacingTwips = nativeInt64(spacing)
				section.Page.ColumnDefinitions = nativeColumnIdentities(id, count)
				continue
			}
			count := len(children)
			if count == 0 || count > 45 {
				extractor.addUnsupported("INVALID_SECTION_COLUMNS", "sections", id, extractor.mainPart, child, "Explicit Word columns require from 1 through 45 w:col children")
				count = 1
				children = nil
			}
			if _, present := nativeAttr(child, extractor.wordNS, "num"); present {
				parsed, ok := nativePositiveIntAttr(child, extractor.wordNS, "num")
				if !ok || parsed != len(children) {
					extractor.addUnsupported("COLUMN_COUNT_MISMATCH", "sections", id, extractor.mainPart, child, "Explicit column count must match the number of w:col children")
				}
			}
			definitions := make([]NativeColumnV1, 0, count)
			for index := 0; index < count; index++ {
				width, spacing := int64(1), int64(0)
				if index < len(children) {
					column := children[index]
					if !nativeExactLeaf(column, xml.Name{Space: extractor.wordNS, Local: "w"}, xml.Name{Space: extractor.wordNS, Local: "space"}) {
						extractor.addUnsupported("MALFORMED_SECTION_COLUMN", "sections", id, extractor.mainPart, column, "Explicit columns allow only exact width and following-space attributes")
					}
					parsedWidth, ok := nativeNonnegativeInt64Attr(column, extractor.wordNS, "w")
					if !ok || parsedWidth == 0 {
						extractor.addUnsupported("INVALID_COLUMN_WIDTH", "sections", id, extractor.mainPart, column, "Every explicit Word column requires a positive width")
					} else {
						width = parsedWidth
					}
					if _, present := nativeAttr(column, extractor.wordNS, "space"); present {
						parsedSpace, ok := nativeNonnegativeInt64Attr(column, extractor.wordNS, "space")
						if !ok {
							extractor.addUnsupported("INVALID_COLUMN_SPACING", "sections", id, extractor.mainPart, column, "Explicit following-column space must be a nonnegative twip integer")
						} else {
							spacing = parsedSpace
						}
					}
				}
				if index == count-1 && spacing != 0 {
					extractor.addUnsupported("TRAILING_COLUMN_SPACING", "sections", id, extractor.mainPart, children[index], "The final explicit column cannot have trailing inter-column space")
					spacing = 0
				}
				definitions = append(definitions, NativeColumnV1{ID: nativeColumnID(id, index), Ordinal: nativeInt(index), WidthTwips: nativeInt64(width), SpaceAfterTwips: nativeInt64(spacing)})
			}
			section.Page.ColumnLayout = "explicit"
			section.Page.Columns = nativeInt(count)
			section.Page.ColumnSpacingTwips = nativeInt64(0)
			section.Page.ColumnDefinitions = definitions
		case "headerReference", "footerReference":
			if !nativeExactLeaf(child, xml.Name{Space: extractor.wordNS, Local: "type"}, xml.Name{Space: extractor.relNS, Local: "id"}) {
				extractor.addUnsupported("UNMODELED_SECTION_PROPERTY", "sections", id, extractor.mainPart, child, "Header/footer reference markup has attributes or children outside the exact v1 subset")
				continue
			}
			kind, okKind := nativeAttr(child, extractor.wordNS, "type")
			if !okKind {
				kind = "default"
			}
			if kind != "default" && kind != "first" && kind != "even" {
				return NativeSectionV1{}, fmt.Errorf("docxpatch: native extract: section %s has invalid %s kind %q", node.Path, child.Name.Local, kind)
			}
			relID, okRel := nativeAttr(child, extractor.relNS, "id")
			if !okRel || relID == "" {
				return NativeSectionV1{}, fmt.Errorf("docxpatch: native extract: section %s has spoofed or missing relationship id", node.Path)
			}
			refKey := child.Name.Local + "\x00" + kind
			if seenRefs[refKey] {
				return NativeSectionV1{}, fmt.Errorf("docxpatch: native extract: section %s has duplicate %s %q reference", node.Path, child.Name.Local, kind)
			}
			seenRefs[refKey] = true
			storyID := extractor.storyByRel[relID]
			wantKind := strings.TrimSuffix(child.Name.Local, "Reference")
			if storyID == "" || !extractor.mainRelationshipMatches(relID, wantKind) {
				return NativeSectionV1{}, fmt.Errorf("docxpatch: native extract: section relationship %q does not resolve to a modeled %s", relID, wantKind)
			}
			ref := NativeHeaderFooterReferenceV1{Kind: kind, StoryID: storyID, RelationshipID: relID}
			if wantKind == "header" {
				section.HeaderRefs = append(section.HeaderRefs, ref)
			} else {
				section.FooterRefs = append(section.FooterRefs, ref)
			}
		default:
			extractor.addUnsupported("UNMODELED_SECTION_PROPERTY", "sections", id, extractor.mainPart, child, "This section property is preserved verbatim")
		}
	}
	if !seenSingleton["pgSz"] {
		extractor.addUnsupported("MISSING_PAGE_SIZE", "sections", id, extractor.mainPart, node, "A present section-properties element requires explicit page size for exact pagination geometry")
	}
	if !seenSingleton["pgMar"] {
		extractor.addUnsupported("MISSING_PAGE_MARGINS", "sections", id, extractor.mainPart, node, "A present section-properties element requires explicit page margins for exact pagination geometry")
	}
	extractor.attestSectionColumnGeometry(&section, node)
	return section, nil
}

func (extractor *nativeExtractor) attestSectionColumnGeometry(section *NativeSectionV1, node *nativeXMLNode) {
	page := &section.Page
	bodyWidth := nativeInt64Value(page.WidthTwips) - nativeInt64Value(page.Margins.LeftTwips) - nativeInt64Value(page.Margins.RightTwips) - nativeInt64Value(page.Margins.GutterTwips)
	if bodyWidth <= 0 {
		extractor.addUnsupported("AMBIGUOUS_SECTION_COLUMNS", "sections", section.ID, extractor.mainPart, node, "Section margins leave no exact positive column body width")
		return
	}
	if page.ColumnLayout == "equal-width" {
		count := nativeIntValue(page.Columns)
		available := bodyWidth - int64(count-1)*nativeInt64Value(page.ColumnSpacingTwips)
		if count < 1 || available <= 0 || available%int64(count) != 0 {
			extractor.addUnsupported("AMBIGUOUS_SECTION_COLUMNS", "sections", section.ID, extractor.mainPart, node, "Equal-width columns are not exactly divisible in native twips")
		}
		return
	}
	total := int64(0)
	firstWidth := int64(-1)
	unequal := false
	for _, column := range page.ColumnDefinitions {
		width := nativeInt64Value(column.WidthTwips)
		space := nativeInt64Value(column.SpaceAfterTwips)
		if firstWidth < 0 {
			firstWidth = width
		} else if width != firstWidth {
			unequal = true
		}
		if width <= 0 || space < 0 || total > 9007199254740991-width-space {
			extractor.addUnsupported("AMBIGUOUS_SECTION_COLUMNS", "sections", section.ID, extractor.mainPart, node, "Explicit column arithmetic is invalid or exceeds the safe integer bound")
			return
		}
		total += width + space
	}
	if unequal {
		extractor.addUnsupported("UNEQUAL_SECTION_COLUMNS", "sections", section.ID, extractor.mainPart, node, "Unequal column widths require per-column shaping outside the exact v1 slice")
	}
	if total != bodyWidth {
		extractor.addUnsupported("AMBIGUOUS_SECTION_COLUMNS", "sections", section.ID, extractor.mainPart, node, "Explicit column widths and gaps do not exactly cover the section body width")
	}
}

func (extractor *nativeExtractor) defaultSection(node *nativeXMLNode, startsAtBlockID string) NativeSectionV1 {
	id := extractor.objectID("section", extractor.mainPart, node, "")
	return NativeSectionV1{
		ID: id, Anchor: extractor.anchor(extractor.mainPart, node), StartsAtBlockID: startsAtBlockID, BreakType: "next-page", TitlePage: nativeBool(false),
		Page: NativePageGeometryV1{
			WidthTwips: nativeInt64(12240), HeightTwips: nativeInt64(15840), Orientation: "portrait",
			Margins: NativePageMarginsV1{TopTwips: nativeInt64(1440), RightTwips: nativeInt64(1440), BottomTwips: nativeInt64(1440), LeftTwips: nativeInt64(1440), HeaderTwips: nativeInt64(720), FooterTwips: nativeInt64(720), GutterTwips: nativeInt64(0)},
			Columns: nativeInt(1), ColumnSpacingTwips: nativeInt64(720), ColumnLayout: "equal-width", ColumnDefinitions: nativeColumnIdentities(id, 1),
		},
		HeaderRefs: []NativeHeaderFooterReferenceV1{}, FooterRefs: []NativeHeaderFooterReferenceV1{},
	}
}

func nativeColumnID(sectionID string, ordinal int) string {
	return sectionID + ":column:" + strconv.Itoa(ordinal)
}

func nativeColumnIdentities(sectionID string, count int) []NativeColumnV1 {
	columns := make([]NativeColumnV1, count)
	for index := range columns {
		columns[index] = NativeColumnV1{ID: nativeColumnID(sectionID, index), Ordinal: nativeInt(index)}
	}
	return columns
}

func (extractor *nativeExtractor) mainRelationshipMatches(id, kind string) bool {
	base := relBaseTransitional
	if extractor.wordNS == wordMLStrict {
		base = relBaseStrict
	}
	for _, rel := range extractor.pkg.rels[extractor.mainPart] {
		if rel.ID == id {
			return !rel.External && rel.Type == base+kind
		}
	}
	return false
}
