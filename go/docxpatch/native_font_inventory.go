package docxpatch

import (
	"bytes"
	"crypto/sha256"
	"encoding/base64"
	"encoding/hex"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"regexp"
	"sort"
	"strings"
)

const (
	NativeDOCXFontInventoryProtocol = "injoffice.docx.font-inventory"
	NativeDOCXFontInventoryVersion  = 1
	NativeDOCXMaxFontFaces          = 4_096
	nativeDOCXObfuscatedFontType    = "application/vnd.openxmlformats-officedocument.obfuscatedFont"
)

var nativeDOCXFontKey = regexp.MustCompile(`^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$`)

const nativeDOCXZeroFontKey = "{00000000-0000-0000-0000-000000000000}"

// NativeDOCXFontInventoryV1 is a read-only, package-revision-bound inventory.
// It is intentionally separate from NativeDocumentV1 so font parts remain
// preserve-verbatim package assets. NativeTextManifest is the narrow projection
// consumed by the existing shaper/page-paint contract; it contains only exact,
// document-backed, content-addressed faces and never contains fallback chains.
type NativeDOCXFontInventoryV1 struct {
	Protocol                 string                        `json:"protocol"`
	Version                  int                           `json:"version"`
	DocumentID               string                        `json:"document_id"`
	Revision                 string                        `json:"revision"`
	PackageSHA256            string                        `json:"package_sha256"`
	MainPart                 string                        `json:"main_part"`
	MainSHA256               string                        `json:"main_sha256"`
	InventorySHA256          string                        `json:"inventory_sha256"`
	FontTable                *NativeDOCXFontTableBindingV1 `json:"font_table,omitempty"`
	Families                 []NativeDOCXFontFamilyV1      `json:"families"`
	References               []NativeDOCXFontReferenceV1   `json:"references"`
	NativeTextManifest       *NativeDOCXTextFontManifestV1 `json:"native_text_manifest,omitempty"`
	NativeTextManifestSHA256 *string                       `json:"native_text_manifest_sha256,omitempty"`
}

type NativeDOCXFontTableBindingV1 struct {
	PartName                string  `json:"part_name"`
	SHA256                  string  `json:"sha256"`
	MainRelationshipsPart   string  `json:"main_relationships_part"`
	MainRelationshipsSHA256 string  `json:"main_relationships_sha256"`
	RelationshipID          string  `json:"relationship_id"`
	RelationshipType        string  `json:"relationship_type"`
	RelationshipTarget      string  `json:"relationship_target"`
	FontRelationshipsPart   *string `json:"font_relationships_part,omitempty"`
	FontRelationshipsSHA256 *string `json:"font_relationships_sha256,omitempty"`
}

type NativeDOCXFontFamilyV1 struct {
	FamilyID string                 `json:"family_id"`
	Name     string                 `json:"name"`
	AltName  *string                `json:"alt_name,omitempty"`
	Faces    []NativeDOCXFontFaceV1 `json:"faces"`
}

type NativeDOCXFontReferenceV1 struct {
	Family   string   `json:"family"`
	Weight   int      `json:"weight"`
	Style    string   `json:"style"`
	ScopeIDs []string `json:"scope_ids"`
}

type NativeDOCXFontFaceV1 struct {
	FaceID  string                         `json:"face_id"`
	Family  string                         `json:"family"`
	AltName *string                        `json:"alt_name,omitempty"`
	Weight  int                            `json:"weight"`
	Style   string                         `json:"style"`
	Stretch int                            `json:"stretch"`
	Source  NativeDOCXEmbeddedFontSourceV1 `json:"source"`
}

type NativeDOCXEmbeddedFontSourceV1 struct {
	Kind                string                      `json:"kind"`
	FaceSlot            string                      `json:"face_slot"`
	FontTablePart       string                      `json:"font_table_part"`
	FontTablePath       string                      `json:"font_table_path"`
	RelationshipsPart   string                      `json:"relationships_part"`
	RelationshipsSHA256 string                      `json:"relationships_sha256"`
	RelationshipID      string                      `json:"relationship_id"`
	RelationshipType    string                      `json:"relationship_type"`
	RelationshipTarget  string                      `json:"relationship_target"`
	AssetPart           string                      `json:"asset_part"`
	AssetContentType    string                      `json:"asset_content_type"`
	StoredByteLength    int                         `json:"stored_byte_length"`
	StoredSHA256        string                      `json:"stored_sha256"`
	ContentSHA256       string                      `json:"content_sha256"`
	ResourceID          string                      `json:"resource_id"`
	CollectionIndex     *int                        `json:"collection_index,omitempty"`
	Obfuscation         NativeDOCXFontObfuscationV1 `json:"obfuscation"`
	Licensing           NativeDOCXFontLicensingV1   `json:"licensing"`
}

type NativeDOCXFontObfuscationV1 struct {
	Algorithm string `json:"algorithm"`
	FontKey   string `json:"font_key"`
	Subsetted *bool  `json:"subsetted,omitempty"`
}

// Licensing is extracted from the exact decoded SFNT OS/2 fsType table. The
// native worker independently rechecks it against supplied bytes before shaping.
type NativeDOCXFontLicensingV1 struct {
	EmbeddingOrigin string `json:"embedding_origin"`
	RightsSource    string `json:"rights_source"`
	RightsStatus    string `json:"rights_status"`
	EmbeddingRights string `json:"embedding_rights"`
	NoSubsetting    bool   `json:"no_subsetting"`
	AllowedScope    string `json:"allowed_scope"`
}

// These structs exactly mirror the existing NativeFontManifest wire shape.
// Keeping the projection here avoids introducing a second shaping contract.
type NativeDOCXTextFontManifestV1 struct {
	Version        int                        `json:"version"`
	ManifestID     string                     `json:"manifestId"`
	Revision       string                     `json:"revision"`
	Faces          []NativeDOCXTextFontFaceV1 `json:"faces"`
	FallbackChains []NativeDOCXFontFallbackV1 `json:"fallbackChains"`
}

type NativeDOCXTextFontFaceV1 struct {
	FaceID  string                     `json:"faceId"`
	Family  string                     `json:"family"`
	Aliases []string                   `json:"aliases,omitempty"`
	Weight  int                        `json:"weight"`
	Style   string                     `json:"style"`
	Stretch int                        `json:"stretch"`
	Source  NativeDOCXTextFontSourceV1 `json:"source"`
}

type NativeDOCXTextFontSourceV1 struct {
	Kind            string `json:"kind"`
	ResourceID      string `json:"resourceId"`
	ContentDigest   string `json:"contentDigest"`
	CollectionIndex *int   `json:"collectionIndex,omitempty"`
}

type NativeDOCXFontFallbackV1 struct {
	ChainID string   `json:"chainId"`
	FaceIDs []string `json:"faceIds"`
}

type NativeDOCXFontResolutionRequestV1 struct {
	Family  string `json:"family"`
	Weight  int    `json:"weight"`
	Style   string `json:"style"`
	Stretch int    `json:"stretch"`
}

type NativeDOCXResolvedFontAssetV1 struct {
	Face      NativeDOCXFontFaceV1
	Bytes     []byte
	Licensing NativeDOCXFontLicensingV1
}

// NativeDOCXPagePaintFontAssetV1 is the exact byte transport consumed by the
// Node page-paint worker. Its identities are copied only from a validated,
// package-reconstructed inventory; callers do not supply a replacement
// manifest, resource id, face slot, digest, or collection index.
type NativeDOCXPagePaintFontAssetV1 struct {
	FaceID          string `json:"face_id"`
	FaceSlot        string `json:"face_slot"`
	ResourceID      string `json:"resource_id"`
	ContentDigest   string `json:"content_digest"`
	CollectionIndex *int   `json:"collection_index"`
	BytesBase64     string `json:"bytes_base64"`
}

type NativeDOCXFontResolutionError struct {
	Code    string
	Message string
}

func (e *NativeDOCXFontResolutionError) Error() string {
	return "docxpatch: native font resolver: " + e.Code + ": " + e.Message
}

// ExtractNativeDOCXFontInventoryV1 binds font identities and supported embedded
// resources to the exact package bytes. It never searches the host or system.
func ExtractNativeDOCXFontInventoryV1(data []byte) (*NativeDOCXFontInventoryV1, error) {
	return ExtractNativeDOCXFontInventoryV1WithOptions(data, NativeExtractionOptions{})
}

func ExtractNativeDOCXFontInventoryV1WithOptions(data []byte, options NativeExtractionOptions) (*NativeDOCXFontInventoryV1, error) {
	doc, err := ExtractNativeDocumentV1WithOptions(data, options)
	if err != nil {
		return nil, err
	}
	resolved, err := ResolveNativeDocumentLayoutV1WithOptions(data, options)
	if err != nil {
		return nil, err
	}
	pkg, err := openNativeDOCXPackage(data)
	if err != nil {
		return nil, err
	}
	mainPart, strict, err := pkg.officeDocumentPart()
	if err != nil {
		return nil, err
	}
	wordNS, relNS, relBase := wordMLTransitional, relNSTransitional, relBaseTransitional
	if strict {
		wordNS, relNS, relBase = wordMLStrict, relNSStrict, relBaseStrict
	}
	inventory := &NativeDOCXFontInventoryV1{
		Protocol: NativeDOCXFontInventoryProtocol, Version: NativeDOCXFontInventoryVersion,
		DocumentID: doc.DocumentID, Revision: doc.Revision, PackageSHA256: doc.Source.PackageSHA256,
		MainPart: mainPart, MainSHA256: nativeSHA(pkg.files[mainPart]), Families: []NativeDOCXFontFamilyV1{}, References: nativeDOCXFontReferences(resolved),
	}
	fontTableRel, err := nativeDOCXSingletonFontTable(pkg, mainPart, relBase)
	if err != nil {
		return nil, err
	}
	if fontTableRel != nil {
		fontTablePart := fontTableRel.PartName
		if !nativeASCIIEqual(pkg.contentTypes[fontTablePart], "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml") {
			return nil, fmt.Errorf("docxpatch: native font inventory: font table %q has content type %q", fontTablePart, pkg.contentTypes[fontTablePart])
		}
		mainRelsPart := pkg.relsPart[mainPart]
		if mainRelsPart == "" {
			return nil, fmt.Errorf("docxpatch: native font inventory: font table has no owning main relationship part")
		}
		binding := &NativeDOCXFontTableBindingV1{
			PartName: fontTablePart, SHA256: nativeSHA(pkg.files[fontTablePart]),
			MainRelationshipsPart: mainRelsPart, MainRelationshipsSHA256: nativeSHA(pkg.files[mainRelsPart]),
			RelationshipID: fontTableRel.ID, RelationshipType: fontTableRel.Type, RelationshipTarget: fontTableRel.Target,
		}
		if relsPart := pkg.relsPart[fontTablePart]; relsPart != "" {
			binding.FontRelationshipsPart = nativeString(relsPart)
			binding.FontRelationshipsSHA256 = nativeString(nativeSHA(pkg.files[relsPart]))
		}
		inventory.FontTable = binding
		families, err := nativeDOCXInventoryFamilies(pkg, fontTablePart, wordNS, relNS, relBase, binding)
		if err != nil {
			return nil, err
		}
		inventory.Families = families
	}
	nativeDOCXSortFontInventory(inventory)
	if faces := nativeDOCXTextFaces(inventory.Families); len(faces) > 0 {
		inventory.NativeTextManifest = &NativeDOCXTextFontManifestV1{
			Version: 1, ManifestID: "docx.fonts." + nativeStableToken(doc.DocumentID), Revision: doc.Revision,
			Faces: faces, FallbackChains: []NativeDOCXFontFallbackV1{},
		}
		manifestDigest := nativeDOCXCanonicalWireSHA256(inventory.NativeTextManifest)
		inventory.NativeTextManifestSHA256 = &manifestDigest
	}
	inventory.InventorySHA256 = nativeDOCXFontInventoryDigest(inventory)
	if err := ValidateNativeDOCXFontInventoryV1(inventory); err != nil {
		return nil, fmt.Errorf("docxpatch: native font inventory output is invalid: %w", err)
	}
	return inventory, nil
}

func nativeDOCXSingletonFontTable(pkg *nativePackage, mainPart, relBase string) (*nativeRelationship, error) {
	var selected *nativeRelationship
	other := relBaseStrict
	if relBase == relBaseStrict {
		other = relBaseTransitional
	}
	for _, rel := range pkg.rels[mainPart] {
		if rel.Type == other+"fontTable" {
			return nil, fmt.Errorf("docxpatch: native font inventory: fontTable relationship %q uses the wrong Strict/Transitional namespace", rel.ID)
		}
		if rel.Type != relBase+"fontTable" {
			continue
		}
		if rel.External || rel.PartName == "" {
			return nil, fmt.Errorf("docxpatch: native font inventory: fontTable relationship %q must be internal", rel.ID)
		}
		if selected != nil {
			return nil, fmt.Errorf("docxpatch: native font inventory: multiple fontTable relationships")
		}
		copyRel := rel
		selected = &copyRel
	}
	return selected, nil
}

func nativeDOCXInventoryFamilies(pkg *nativePackage, partName, wordNS, relNS, relBase string, binding *NativeDOCXFontTableBindingV1) ([]NativeDOCXFontFamilyV1, error) {
	root, err := parseNativeXML(partName, pkg.files[partName])
	if err != nil {
		return nil, err
	}
	if root.Name != (xml.Name{Space: wordNS, Local: "fonts"}) {
		return nil, fmt.Errorf("docxpatch: native font inventory: font table %q has spoofed or invalid root", partName)
	}
	known := map[string]bool{"fonts": true, "font": true, "altName": true, "notTrueType": true, "embedRegular": true, "embedBold": true, "embedItalic": true, "embedBoldItalic": true}
	if err := rejectNativeKnownLocalSpoofing(root, wordNS, known); err != nil {
		return nil, fmt.Errorf("docxpatch: native font inventory: font table %q: %w", partName, err)
	}
	if len(root.Children) > NativeDOCXMaxCollectionItems {
		return nil, fmt.Errorf("docxpatch: native font inventory: font table exceeds %d entries", NativeDOCXMaxCollectionItems)
	}
	seen := map[string]bool{}
	families := []NativeDOCXFontFamilyV1{}
	for _, node := range root.Children {
		if node.Name != (xml.Name{Space: wordNS, Local: "font"}) {
			continue
		}
		name, ok := nativeAttr(node, wordNS, "name")
		if !ok || !nativeBoundedResolvedString(name, 256) {
			return nil, fmt.Errorf("docxpatch: native font inventory: invalid font identity at %s", node.Path)
		}
		key := nativeASCIIFold(name)
		if seen[key] {
			return nil, fmt.Errorf("docxpatch: native font inventory: duplicate font identity %q", name)
		}
		seen[key] = true
		family := NativeDOCXFontFamilyV1{FamilyID: "font-family:" + nativeStableToken(name), Name: name, Faces: []NativeDOCXFontFaceV1{}}
		altNames := directNativeChildren(node, wordNS, "altName")
		if len(altNames) > 1 {
			return nil, fmt.Errorf("docxpatch: native font inventory: duplicate altName for %q", name)
		}
		if len(altNames) == 1 {
			alt, present := nativeAttr(altNames[0], wordNS, "val")
			if !present || !nativeBoundedResolvedString(alt, 256) || nativeASCIIFold(alt) == key {
				return nil, fmt.Errorf("docxpatch: native font inventory: invalid altName for %q", name)
			}
			family.AltName = nativeString(alt)
		}
		if values := directNativeChildren(node, wordNS, "notTrueType"); len(values) > 1 {
			return nil, fmt.Errorf("docxpatch: native font inventory: duplicate notTrueType for %q", name)
		} else if len(values) == 1 {
			value := true
			if raw, present := nativeAttr(values[0], wordNS, "val"); present {
				var valid bool
				value, valid = nativeLexicalOnOff(raw)
				if !valid {
					return nil, fmt.Errorf("docxpatch: native font inventory: invalid notTrueType for %q", name)
				}
			}
			if value && nativeDOCXHasEmbeddedFace(node, wordNS) {
				return nil, fmt.Errorf("docxpatch: native font inventory: embedded non-TrueType font %q is unsupported", name)
			}
		}
		slots := []struct {
			local, style string
			weight       int
		}{{"embedRegular", "normal", 400}, {"embedBold", "normal", 700}, {"embedItalic", "italic", 400}, {"embedBoldItalic", "italic", 700}}
		for _, slot := range slots {
			nodes := directNativeChildren(node, wordNS, slot.local)
			if len(nodes) > 1 {
				return nil, fmt.Errorf("docxpatch: native font inventory: duplicate %s for %q", slot.local, name)
			}
			if len(nodes) == 0 {
				continue
			}
			face, err := nativeDOCXEmbeddedFace(pkg, partName, wordNS, relNS, relBase, binding, nodes[0], family, slot.local, slot.weight, slot.style)
			if err != nil {
				return nil, err
			}
			family.Faces = append(family.Faces, face)
		}
		families = append(families, family)
	}
	return families, nil
}

func nativeDOCXHasEmbeddedFace(node *nativeXMLNode, wordNS string) bool {
	for _, local := range []string{"embedRegular", "embedBold", "embedItalic", "embedBoldItalic"} {
		if len(directNativeChildren(node, wordNS, local)) > 0 {
			return true
		}
	}
	return false
}

func nativeDOCXEmbeddedFace(pkg *nativePackage, fontTablePart, wordNS, relNS, relBase string, binding *NativeDOCXFontTableBindingV1, node *nativeXMLNode, family NativeDOCXFontFamilyV1, faceSlot string, weight int, style string) (NativeDOCXFontFaceV1, error) {
	if binding.FontRelationshipsPart == nil || binding.FontRelationshipsSHA256 == nil {
		return NativeDOCXFontFaceV1{}, fmt.Errorf("docxpatch: native font inventory: embedded face at %s has no relationship part", node.Path)
	}
	relID, ok := nativeAttr(node, relNS, "id")
	if !ok || !nativeIDPattern.MatchString(relID) {
		return NativeDOCXFontFaceV1{}, fmt.Errorf("docxpatch: native font inventory: embedded face at %s has an invalid relationship id", node.Path)
	}
	fontKey, ok := nativeAttr(node, wordNS, "fontKey")
	if !ok || !nativeDOCXFontKey.MatchString(fontKey) || fontKey == nativeDOCXZeroFontKey {
		return NativeDOCXFontFaceV1{}, fmt.Errorf("docxpatch: native font inventory: embedded face at %s has an invalid canonical fontKey", node.Path)
	}
	var subsetted *bool
	if raw, present := nativeAttr(node, wordNS, "subsetted"); present {
		value, valid := nativeLexicalOnOff(raw)
		if !valid {
			return NativeDOCXFontFaceV1{}, fmt.Errorf("docxpatch: native font inventory: embedded face at %s has an invalid subsetted value", node.Path)
		}
		subsetted = nativeBool(value)
	}
	var matched *nativeRelationship
	for index := range pkg.rels[fontTablePart] {
		rel := &pkg.rels[fontTablePart][index]
		if rel.ID == relID {
			matched = rel
			break
		}
	}
	if matched == nil || matched.External || matched.PartName == "" || matched.Type != relBase+"font" {
		return NativeDOCXFontFaceV1{}, fmt.Errorf("docxpatch: native font inventory: relationship %q for embedded face at %s is not an exact internal %sfont relationship", relID, node.Path, relBase)
	}
	if !nativeASCIIEqual(pkg.contentTypes[matched.PartName], nativeDOCXObfuscatedFontType) {
		return NativeDOCXFontFaceV1{}, fmt.Errorf("docxpatch: native font inventory: embedded asset %q has unsupported content type %q", matched.PartName, pkg.contentTypes[matched.PartName])
	}
	stored := pkg.files[matched.PartName]
	content, err := nativeDOCXDeobfuscateFont(stored, fontKey)
	if err != nil {
		return NativeDOCXFontFaceV1{}, fmt.Errorf("docxpatch: native font inventory: embedded asset %q: %w", matched.PartName, err)
	}
	if err := nativeDOCXValidateSFNTEnvelope(content); err != nil {
		return NativeDOCXFontFaceV1{}, fmt.Errorf("docxpatch: native font inventory: embedded asset %q: %w", matched.PartName, err)
	}
	licensing, err := nativeDOCXFontLicensing(content, subsetted)
	if err != nil {
		return NativeDOCXFontFaceV1{}, fmt.Errorf("docxpatch: native font inventory: embedded asset %q: %w", matched.PartName, err)
	}
	contentDigest := nativeSHA(content)
	faceID := "font-face:" + nativeStableToken(family.Name+"\x00"+fmt.Sprint(weight)+"\x00"+style+"\x00"+contentDigest)
	return NativeDOCXFontFaceV1{
		FaceID: faceID, Family: family.Name, AltName: family.AltName, Weight: weight, Style: style, Stretch: 100,
		Source: NativeDOCXEmbeddedFontSourceV1{
			Kind: "document", FaceSlot: faceSlot, FontTablePart: fontTablePart, FontTablePath: node.Path,
			RelationshipsPart: *binding.FontRelationshipsPart, RelationshipsSHA256: *binding.FontRelationshipsSHA256,
			RelationshipID: relID, RelationshipType: matched.Type, RelationshipTarget: matched.Target, AssetPart: matched.PartName,
			AssetContentType: pkg.contentTypes[matched.PartName], StoredByteLength: len(stored), StoredSHA256: nativeSHA(stored),
			ContentSHA256: contentDigest, ResourceID: "font:" + contentDigest, CollectionIndex: nil,
			Obfuscation: NativeDOCXFontObfuscationV1{Algorithm: "ecma-376-font-obfuscation", FontKey: fontKey, Subsetted: subsetted},
			Licensing:   licensing,
		},
	}, nil
}

func nativeDOCXDeobfuscateFont(stored []byte, fontKey string) ([]byte, error) {
	if len(stored) < 32 || !nativeDOCXFontKey.MatchString(fontKey) || fontKey == nativeDOCXZeroFontKey {
		return nil, fmt.Errorf("obfuscated font requires at least 32 bytes and a canonical uppercase fontKey")
	}
	hexKey := strings.ReplaceAll(strings.Trim(fontKey, "{}"), "-", "")
	key, err := hex.DecodeString(hexKey)
	if err != nil || len(key) != 16 {
		return nil, fmt.Errorf("invalid fontKey")
	}
	content := append([]byte(nil), stored...)
	for index := 0; index < 32; index++ {
		content[index] ^= key[15-index%16]
	}
	return content, nil
}

// This validates only the bounded sfnt transport envelope. It deliberately does
// not parse outlines, names, metrics, cmap, or collections.
func nativeDOCXValidateSFNTEnvelope(data []byte) error {
	if len(data) < 12 {
		return fmt.Errorf("malformed sfnt envelope")
	}
	if !bytes.Equal(data[:4], []byte{0, 1, 0, 0}) && string(data[:4]) != "OTTO" && string(data[:4]) != "true" {
		return fmt.Errorf("unsupported or malformed standalone sfnt signature")
	}
	numTables := int(data[4])<<8 | int(data[5])
	if numTables < 1 || numTables > 4096 || 12+16*numTables > len(data) {
		return fmt.Errorf("malformed sfnt table directory")
	}
	seen := map[string]bool{}
	for index := 0; index < numTables; index++ {
		base := 12 + 16*index
		tag := string(data[base : base+4])
		if seen[tag] {
			return fmt.Errorf("malformed sfnt duplicate table %q", tag)
		}
		seen[tag] = true
		offset := uint64(data[base+8])<<24 | uint64(data[base+9])<<16 | uint64(data[base+10])<<8 | uint64(data[base+11])
		length := uint64(data[base+12])<<24 | uint64(data[base+13])<<16 | uint64(data[base+14])<<8 | uint64(data[base+15])
		if offset > uint64(len(data)) || length > uint64(len(data))-offset {
			return fmt.Errorf("malformed sfnt table %q range", tag)
		}
	}
	return nil
}

func nativeDOCXFontLicensing(data []byte, subsetted *bool) (NativeDOCXFontLicensingV1, error) {
	numTables := int(data[4])<<8 | int(data[5])
	var os2 []byte
	for index := 0; index < numTables; index++ {
		base := 12 + 16*index
		if string(data[base:base+4]) != "OS/2" {
			continue
		}
		offset := uint64(data[base+8])<<24 | uint64(data[base+9])<<16 | uint64(data[base+10])<<8 | uint64(data[base+11])
		length := uint64(data[base+12])<<24 | uint64(data[base+13])<<16 | uint64(data[base+14])<<8 | uint64(data[base+15])
		if length < 10 {
			return NativeDOCXFontLicensingV1{}, fmt.Errorf("malformed OS/2 licensing table")
		}
		os2 = data[int(offset):int(offset+length)]
		break
	}
	if os2 == nil {
		return NativeDOCXFontLicensingV1{}, fmt.Errorf("missing OS/2 licensing table")
	}
	fsType := uint16(os2[8])<<8 | uint16(os2[9])
	if fsType&0xfcf1 != 0 || fsType&0x000e != 0 && fsType&0x000e != 0x0002 && fsType&0x000e != 0x0004 && fsType&0x000e != 0x0008 {
		return NativeDOCXFontLicensingV1{}, fmt.Errorf("reserved or conflicting OS/2 fsType 0x%04x", fsType)
	}
	if fsType&0x0002 != 0 {
		return NativeDOCXFontLicensingV1{}, fmt.Errorf("restricted-license embedding is refused")
	}
	if fsType&0x0200 != 0 {
		return NativeDOCXFontLicensingV1{}, fmt.Errorf("bitmap-only embedding is unsupported for outlines")
	}
	noSubsetting := fsType&0x0100 != 0
	if noSubsetting && subsetted != nil && *subsetted {
		return NativeDOCXFontLicensingV1{}, fmt.Errorf("subsetted asset conflicts with OS/2 no-subsetting rights")
	}
	rights := "installable"
	if fsType&0x0004 != 0 {
		rights = "preview-print"
	} else if fsType&0x0008 != 0 {
		rights = "editable"
	}
	return NativeDOCXFontLicensingV1{
		EmbeddingOrigin: "document-package", RightsSource: "sfnt-os2-fstype", RightsStatus: "verified",
		EmbeddingRights: rights, NoSubsetting: noSubsetting, AllowedScope: "document-only",
	}, nil
}

func nativeDOCXFontReferences(resolved *NativeResolvedLayoutInputV1) []NativeDOCXFontReferenceV1 {
	type key struct {
		family, style string
		weight        int
	}
	refs := map[key]map[string]bool{}
	add := func(family *string, bold, italic *bool, scope string) {
		if family == nil || *family == "" {
			return
		}
		weight, style := 400, "normal"
		if bold != nil && *bold {
			weight = 700
		}
		if italic != nil && *italic {
			style = "italic"
		}
		entry := key{family: *family, weight: weight, style: style}
		if refs[entry] == nil {
			refs[entry] = map[string]bool{}
		}
		refs[entry][scope] = true
	}
	for _, run := range resolved.Runs {
		add(run.Properties.FontFamily, run.Properties.Bold, run.Properties.Italic, run.RunID)
	}
	for _, paragraph := range resolved.Paragraphs {
		properties := paragraph.ParagraphMarkProperties
		add(properties.FontFamily, properties.Bold, properties.Italic, paragraph.ParagraphID)
		if paragraph.Numbering != nil {
			marker := paragraph.Numbering.Marker
			add(marker.FontFamily, marker.Bold, marker.Italic, paragraph.ParagraphID)
		}
	}
	result := make([]NativeDOCXFontReferenceV1, 0, len(refs))
	for entry, scopes := range refs {
		scopeIDs := make([]string, 0, len(scopes))
		for scope := range scopes {
			scopeIDs = append(scopeIDs, scope)
		}
		sort.Strings(scopeIDs)
		result = append(result, NativeDOCXFontReferenceV1{Family: entry.family, Weight: entry.weight, Style: entry.style, ScopeIDs: scopeIDs})
	}
	sort.Slice(result, func(i, j int) bool {
		left, right := result[i], result[j]
		return nativeASCIIFold(left.Family) < nativeASCIIFold(right.Family) || nativeASCIIFold(left.Family) == nativeASCIIFold(right.Family) && (left.Weight < right.Weight || left.Weight == right.Weight && left.Style < right.Style)
	})
	return result
}

func nativeDOCXSortFontInventory(inventory *NativeDOCXFontInventoryV1) {
	for index := range inventory.Families {
		sort.Slice(inventory.Families[index].Faces, func(i, j int) bool {
			left, right := inventory.Families[index].Faces[i], inventory.Families[index].Faces[j]
			return left.Weight < right.Weight || left.Weight == right.Weight && left.Style < right.Style
		})
	}
	sort.Slice(inventory.Families, func(i, j int) bool {
		return nativeASCIIFold(inventory.Families[i].Name) < nativeASCIIFold(inventory.Families[j].Name)
	})
}

func nativeDOCXTextFaces(families []NativeDOCXFontFamilyV1) []NativeDOCXTextFontFaceV1 {
	faces := []NativeDOCXTextFontFaceV1{}
	for _, family := range families {
		for _, face := range family.Faces {
			aliases := []string(nil)
			if face.AltName != nil {
				aliases = []string{*face.AltName}
			}
			faces = append(faces, NativeDOCXTextFontFaceV1{
				FaceID: face.FaceID, Family: face.Family, Aliases: aliases, Weight: face.Weight, Style: face.Style, Stretch: face.Stretch,
				Source: NativeDOCXTextFontSourceV1{Kind: "document", ResourceID: face.Source.ResourceID, ContentDigest: face.Source.ContentSHA256},
			})
		}
	}
	sort.Slice(faces, func(i, j int) bool { return faces[i].FaceID < faces[j].FaceID })
	return faces
}

func nativeStableToken(value string) string {
	return strings.TrimPrefix(nativeSHA([]byte(value)), "sha256:")[:24]
}

func nativeDOCXFontInventoryDigest(inventory *NativeDOCXFontInventoryV1) string {
	copyInventory := *inventory
	copyInventory.InventorySHA256 = ""
	return nativeDOCXCanonicalWireSHA256(&copyInventory)
}

// nativeDOCXCanonicalWireSHA256 is shared conceptually with the TypeScript
// native-wire encoder: object keys sort by Unicode code point, arrays retain
// order, JSON numbers remain integral, and HTML characters are not rewritten.
// encoding/json sorts map keys deterministically, so round-tripping through an
// interface removes Go struct declaration order from the digest contract.
func nativeDOCXCanonicalWireSHA256(value any) string {
	encoded, _ := json.Marshal(value)
	var wire any
	decoder := json.NewDecoder(bytes.NewReader(encoded))
	decoder.UseNumber()
	_ = decoder.Decode(&wire)
	var canonical bytes.Buffer
	encoder := json.NewEncoder(&canonical)
	encoder.SetEscapeHTML(false)
	_ = encoder.Encode(wire)
	data := bytes.TrimSuffix(canonical.Bytes(), []byte{'\n'})
	digest := sha256.Sum256(data)
	return "sha256:" + hex.EncodeToString(digest[:])
}

func ValidateNativeDOCXFontInventoryV1(inventory *NativeDOCXFontInventoryV1) error {
	if inventory == nil {
		return fmt.Errorf("font inventory is nil")
	}
	if inventory.Protocol != NativeDOCXFontInventoryProtocol || inventory.Version != NativeDOCXFontInventoryVersion {
		return fmt.Errorf("unsupported font inventory protocol/version")
	}
	if !nativeIDPattern.MatchString(inventory.DocumentID) || !nativeIDPattern.MatchString(inventory.Revision) {
		return fmt.Errorf("invalid document identity")
	}
	if !nativeSHA256.MatchString(inventory.PackageSHA256) || !nativeSHA256.MatchString(inventory.MainSHA256) || !nativeSHA256.MatchString(inventory.InventorySHA256) {
		return fmt.Errorf("invalid inventory digest")
	}
	if inventory.Revision != "rev:"+strings.TrimPrefix(inventory.PackageSHA256, "sha256:")[:32] {
		return fmt.Errorf("revision does not bind the full package digest")
	}
	if err := validateNativePartName(inventory.MainPart); err != nil {
		return fmt.Errorf("invalid main part: %w", err)
	}
	if inventory.InventorySHA256 != nativeDOCXFontInventoryDigest(inventory) {
		return fmt.Errorf("inventory_sha256 does not attest the complete inventory")
	}
	if len(inventory.Families) > NativeDOCXMaxCollectionItems || len(inventory.References) > NativeDOCXMaxCollectionItems {
		return fmt.Errorf("font inventory exceeds collection limits")
	}
	if inventory.FontTable == nil && len(inventory.Families) != 0 {
		return fmt.Errorf("font families require a font-table binding")
	}
	var fontRelationshipType string
	if inventory.FontTable != nil {
		binding := inventory.FontTable
		for _, part := range []string{binding.PartName, binding.MainRelationshipsPart} {
			if err := validateNativePartName(part); err != nil {
				return fmt.Errorf("invalid font-table binding: %w", err)
			}
		}
		if owner, ok := nativeRelationshipOwner(binding.MainRelationshipsPart); !ok || owner != inventory.MainPart {
			return fmt.Errorf("main relationship part does not exactly belong to the main part")
		}
		if !nativeSHA256.MatchString(binding.SHA256) || !nativeSHA256.MatchString(binding.MainRelationshipsSHA256) || !nativeIDPattern.MatchString(binding.RelationshipID) || binding.RelationshipType == "" || binding.RelationshipTarget == "" {
			return fmt.Errorf("incomplete font-table relationship binding")
		}
		switch binding.RelationshipType {
		case relBaseTransitional + "fontTable":
			fontRelationshipType = relBaseTransitional + "font"
		case relBaseStrict + "fontTable":
			fontRelationshipType = relBaseStrict + "font"
		default:
			return fmt.Errorf("invalid font-table relationship type")
		}
		if (binding.FontRelationshipsPart == nil) != (binding.FontRelationshipsSHA256 == nil) {
			return fmt.Errorf("incomplete font relationship-part binding")
		}
		if binding.FontRelationshipsPart != nil {
			if err := validateNativePartName(*binding.FontRelationshipsPart); err != nil || !nativeSHA256.MatchString(*binding.FontRelationshipsSHA256) {
				return fmt.Errorf("invalid font relationship-part binding")
			}
			if owner, ok := nativeRelationshipOwner(*binding.FontRelationshipsPart); !ok || owner != binding.PartName {
				return fmt.Errorf("font relationship part does not exactly belong to the font table")
			}
		}
		resolvedTarget, err := resolveNativeRelationshipTarget(inventory.MainPart, binding.RelationshipTarget)
		if err != nil || resolvedTarget != binding.PartName {
			return fmt.Errorf("font-table relationship target does not exactly bind the font table part")
		}
	}
	faceIDs, resources, assetParts, relationshipIDs := map[string]NativeDOCXFontFaceV1{}, map[string]bool{}, map[string]bool{}, map[string]bool{}
	familyNames, familyIDs := map[string]bool{}, map[string]bool{}
	totalFaces := 0
	for _, family := range inventory.Families {
		totalFaces += len(family.Faces)
		if totalFaces > NativeDOCXMaxFontFaces {
			return fmt.Errorf("font inventory exceeds the native text v1 face limit")
		}
	}
	for familyIndex, family := range inventory.Families {
		if !nativeIDPattern.MatchString(family.FamilyID) || !nativeBoundedResolvedString(family.Name, 256) {
			return fmt.Errorf("invalid font family")
		}
		familyKey := nativeASCIIFold(family.Name)
		if familyNames[familyKey] || familyIDs[family.FamilyID] {
			return fmt.Errorf("duplicate font family")
		}
		familyNames[familyKey], familyIDs[family.FamilyID] = true, true
		if familyIndex > 0 && nativeASCIIFold(inventory.Families[familyIndex-1].Name) >= familyKey {
			return fmt.Errorf("font families are not in canonical order")
		}
		if family.AltName != nil && (!nativeBoundedResolvedString(*family.AltName, 256) || nativeASCIIEqual(*family.AltName, family.Name)) {
			return fmt.Errorf("invalid font family alias")
		}
		for faceIndex, face := range family.Faces {
			aliasMismatch := (face.AltName == nil) != (family.AltName == nil) || face.AltName != nil && *face.AltName != *family.AltName
			if face.Family != family.Name || aliasMismatch || !nativeIDPattern.MatchString(face.FaceID) || faceIDs[face.FaceID].FaceID != "" {
				return fmt.Errorf("invalid or duplicate font face")
			}
			faceIDs[face.FaceID] = face
			if face.Weight != 400 && face.Weight != 700 || face.Style != "normal" && face.Style != "italic" || face.Stretch != 100 {
				return fmt.Errorf("unsupported font face metadata")
			}
			if faceIndex > 0 {
				previous := family.Faces[faceIndex-1]
				if previous.Weight > face.Weight || previous.Weight == face.Weight && previous.Style >= face.Style {
					return fmt.Errorf("font faces are not in canonical order")
				}
			}
			source := face.Source
			expectedSlot := map[string]string{"400\x00normal": "embedRegular", "700\x00normal": "embedBold", "400\x00italic": "embedItalic", "700\x00italic": "embedBoldItalic"}[fmt.Sprintf("%d\x00%s", face.Weight, face.Style)]
			if source.Kind != "document" || source.FaceSlot != expectedSlot || source.CollectionIndex != nil || !nativeSHA256.MatchString(source.StoredSHA256) || !nativeSHA256.MatchString(source.ContentSHA256) || source.ResourceID != "font:"+source.ContentSHA256 {
				return fmt.Errorf("font face is not content addressed")
			}
			if inventory.FontTable == nil || source.FontTablePart != inventory.FontTable.PartName || inventory.FontTable.FontRelationshipsPart == nil || source.RelationshipsPart != *inventory.FontTable.FontRelationshipsPart || source.RelationshipsSHA256 != *inventory.FontTable.FontRelationshipsSHA256 || source.RelationshipType != fontRelationshipType || !nativeASCIIEqual(source.AssetContentType, nativeDOCXObfuscatedFontType) {
				return fmt.Errorf("font face relationship is not exactly inventory backed")
			}
			resolvedTarget, err := resolveNativeRelationshipTarget(source.FontTablePart, source.RelationshipTarget)
			if err != nil || resolvedTarget != source.AssetPart {
				return fmt.Errorf("font relationship target does not exactly bind the asset part")
			}
			for _, part := range []string{source.FontTablePart, source.RelationshipsPart, source.AssetPart} {
				if err := validateNativePartName(part); err != nil {
					return fmt.Errorf("invalid font source part: %w", err)
				}
			}
			if source.StoredByteLength < 32 || source.RelationshipID == "" || source.RelationshipType == "" || source.RelationshipTarget == "" || source.AssetContentType == "" || source.FontTablePath == "" {
				return fmt.Errorf("incomplete font source binding")
			}
			if source.Obfuscation.Algorithm != "ecma-376-font-obfuscation" || !nativeDOCXFontKey.MatchString(source.Obfuscation.FontKey) || source.Obfuscation.FontKey == nativeDOCXZeroFontKey {
				return fmt.Errorf("invalid font obfuscation binding")
			}
			license := source.Licensing
			if license.EmbeddingOrigin != "document-package" || license.RightsSource != "sfnt-os2-fstype" || license.RightsStatus != "verified" || license.AllowedScope != "document-only" || license.EmbeddingRights != "installable" && license.EmbeddingRights != "preview-print" && license.EmbeddingRights != "editable" || license.NoSubsetting && source.Obfuscation.Subsetted != nil && *source.Obfuscation.Subsetted {
				return fmt.Errorf("invalid font licensing semantics")
			}
			if resources[source.ResourceID] || assetParts[nativeASCIIFold(source.AssetPart)] || relationshipIDs[source.RelationshipID] {
				return fmt.Errorf("duplicate font resource binding")
			}
			resources[source.ResourceID], assetParts[nativeASCIIFold(source.AssetPart)], relationshipIDs[source.RelationshipID] = true, true, true
		}
	}
	seenReferences := map[string]bool{}
	for referenceIndex, reference := range inventory.References {
		if !nativeBoundedResolvedString(reference.Family, 256) || reference.Weight != 400 && reference.Weight != 700 || reference.Style != "normal" && reference.Style != "italic" || len(reference.ScopeIDs) == 0 || len(reference.ScopeIDs) > NativeDOCXMaxCollectionItems {
			return fmt.Errorf("invalid font reference")
		}
		key := nativeASCIIFold(reference.Family) + "\x00" + fmt.Sprint(reference.Weight) + "\x00" + reference.Style
		if seenReferences[key] {
			return fmt.Errorf("duplicate font reference")
		}
		seenReferences[key] = true
		if referenceIndex > 0 {
			previous := inventory.References[referenceIndex-1]
			previousKey := nativeASCIIFold(previous.Family) + "\x00" + fmt.Sprintf("%04d", previous.Weight) + "\x00" + previous.Style
			currentKey := nativeASCIIFold(reference.Family) + "\x00" + fmt.Sprintf("%04d", reference.Weight) + "\x00" + reference.Style
			if previousKey >= currentKey {
				return fmt.Errorf("font references are not in canonical order")
			}
		}
		seenScopes := map[string]bool{}
		for scopeIndex, scope := range reference.ScopeIDs {
			if !nativeIDPattern.MatchString(scope) || seenScopes[scope] || scopeIndex > 0 && reference.ScopeIDs[scopeIndex-1] >= scope {
				return fmt.Errorf("invalid font reference scope")
			}
			seenScopes[scope] = true
		}
	}
	if inventory.NativeTextManifest == nil {
		if len(faceIDs) != 0 || inventory.NativeTextManifestSHA256 != nil {
			return fmt.Errorf("native text manifest is required for embedded faces")
		}
	} else {
		manifest := inventory.NativeTextManifest
		if inventory.NativeTextManifestSHA256 == nil || !nativeSHA256.MatchString(*inventory.NativeTextManifestSHA256) || *inventory.NativeTextManifestSHA256 != nativeDOCXCanonicalWireSHA256(manifest) || manifest.Version != 1 || manifest.ManifestID != "docx.fonts."+nativeStableToken(inventory.DocumentID) || manifest.Revision != inventory.Revision || len(manifest.FallbackChains) != 0 || len(manifest.Faces) != len(faceIDs) || len(manifest.Faces) > NativeDOCXMaxFontFaces {
			return fmt.Errorf("invalid native text manifest binding")
		}
		seenManifestFaces := map[string]bool{}
		for faceIndex, face := range manifest.Faces {
			backing := faceIDs[face.FaceID]
			expectedAliases := []string(nil)
			if backing.AltName != nil {
				expectedAliases = []string{*backing.AltName}
			}
			if backing.FaceID == "" || seenManifestFaces[face.FaceID] || faceIndex > 0 && manifest.Faces[faceIndex-1].FaceID >= face.FaceID || face.Family != backing.Family || !equalNativeStrings(face.Aliases, expectedAliases) || face.Weight != backing.Weight || face.Style != backing.Style || face.Stretch != backing.Stretch || face.Source.Kind != "document" || face.Source.CollectionIndex != backing.Source.CollectionIndex || !nativeSHA256.MatchString(face.Source.ContentDigest) || !resources[face.Source.ResourceID] || face.Source.ResourceID != "font:"+face.Source.ContentDigest || face.Source.ContentDigest != backing.Source.ContentSHA256 {
				return fmt.Errorf("native text face is not exactly inventory backed")
			}
			seenManifestFaces[face.FaceID] = true
		}
	}
	return nil
}

func equalNativeStrings(left, right []string) bool {
	if len(left) != len(right) {
		return false
	}
	for index := range left {
		if left[index] != right[index] {
			return false
		}
	}
	return true
}

func EncodeNativeDOCXFontInventoryV1(inventory *NativeDOCXFontInventoryV1) ([]byte, error) {
	if err := ValidateNativeDOCXFontInventoryV1(inventory); err != nil {
		return nil, err
	}
	data, err := json.Marshal(inventory)
	if err != nil {
		return nil, err
	}
	if len(data) > NativeDOCXMaxJSONBytes {
		return nil, fmt.Errorf("font inventory JSON exceeds %d bytes", NativeDOCXMaxJSONBytes)
	}
	return data, nil
}

func DecodeNativeDOCXFontInventoryV1(data []byte) (*NativeDOCXFontInventoryV1, error) {
	if len(data) > NativeDOCXMaxJSONBytes {
		return nil, fmt.Errorf("font inventory JSON exceeds %d bytes", NativeDOCXMaxJSONBytes)
	}
	if issues, err := preflightNativeJSON(data); err != nil {
		return nil, err
	} else if len(issues) > 0 {
		return nil, &NativeValidationError{Issues: issues}
	}
	decoder := json.NewDecoder(bytes.NewReader(data))
	decoder.DisallowUnknownFields()
	var inventory NativeDOCXFontInventoryV1
	if err := decoder.Decode(&inventory); err != nil {
		return nil, err
	}
	if err := requireJSONEOF(decoder); err != nil {
		return nil, err
	}
	if err := ValidateNativeDOCXFontInventoryV1(&inventory); err != nil {
		return nil, err
	}
	canonical, err := EncodeNativeDOCXFontInventoryV1(&inventory)
	if err != nil {
		return nil, err
	}
	if !bytes.Equal(data, canonical) {
		return nil, fmt.Errorf("font inventory JSON is not the exact canonical Go v1 encoding")
	}
	return &inventory, nil
}

// ResolveNativeDOCXFontAssetV1 performs only exact document resolution. It
// reconstructs the canonical inventory from the supplied package before
// releasing deobfuscated bytes, closing stale revision, forged path/digest,
// system-font, substitution, and fallback routes.
func ResolveNativeDOCXFontAssetV1(data []byte, inventory *NativeDOCXFontInventoryV1, request NativeDOCXFontResolutionRequestV1) (*NativeDOCXResolvedFontAssetV1, error) {
	if err := ValidateNativeDOCXFontInventoryV1(inventory); err != nil {
		return nil, &NativeDOCXFontResolutionError{Code: "INVALID_MANIFEST", Message: err.Error()}
	}
	if request.Family == "" || request.Stretch != 100 || request.Weight != 400 && request.Weight != 700 || request.Style != "normal" && request.Style != "italic" {
		return nil, &NativeDOCXFontResolutionError{Code: "SUBSTITUTION_REFUSED", Message: "request must name one exact document family and supported authored face"}
	}
	rebuilt, err := ExtractNativeDOCXFontInventoryV1WithOptions(data, NativeExtractionOptions{DocumentID: inventory.DocumentID})
	if err != nil {
		return nil, &NativeDOCXFontResolutionError{Code: "PACKAGE_REFUSED", Message: err.Error()}
	}
	expected, _ := EncodeNativeDOCXFontInventoryV1(inventory)
	actual, _ := EncodeNativeDOCXFontInventoryV1(rebuilt)
	if !bytes.Equal(expected, actual) {
		return nil, &NativeDOCXFontResolutionError{Code: "REVISION_MISMATCH", Message: "manifest does not exactly attest the supplied package"}
	}
	var selected *NativeDOCXFontFaceV1
	for familyIndex := range rebuilt.Families {
		family := &rebuilt.Families[familyIndex]
		matches := nativeASCIIEqual(family.Name, request.Family) || family.AltName != nil && nativeASCIIEqual(*family.AltName, request.Family)
		if !matches {
			continue
		}
		for faceIndex := range family.Faces {
			face := &family.Faces[faceIndex]
			if face.Weight == request.Weight && face.Style == request.Style && face.Stretch == request.Stretch {
				if selected != nil {
					return nil, &NativeDOCXFontResolutionError{Code: "AMBIGUOUS_FONT", Message: "multiple exact document faces match the request"}
				}
				selected = face
			}
		}
	}
	if selected == nil {
		return nil, &NativeDOCXFontResolutionError{Code: "FONT_UNAVAILABLE", Message: "the exact requested document-embedded face is unavailable"}
	}
	pkg, err := openNativeDOCXPackage(data)
	if err != nil {
		return nil, &NativeDOCXFontResolutionError{Code: "PACKAGE_REFUSED", Message: err.Error()}
	}
	return nativeDOCXResolvedFontAssetFromPackage(pkg, selected)
}

// ResolveNativeDOCXPagePaintFontAssetsV1 reconstructs the inventory once and
// emits exact, canonical worker assets in manifest-face order. This is the Go
// bridge seam for host applications: no caller-authored manifest or asset
// identity is accepted.
func ResolveNativeDOCXPagePaintFontAssetsV1(data []byte, inventory *NativeDOCXFontInventoryV1) ([]NativeDOCXPagePaintFontAssetV1, error) {
	if err := ValidateNativeDOCXFontInventoryV1(inventory); err != nil {
		return nil, &NativeDOCXFontResolutionError{Code: "INVALID_MANIFEST", Message: err.Error()}
	}
	rebuilt, err := ExtractNativeDOCXFontInventoryV1WithOptions(data, NativeExtractionOptions{DocumentID: inventory.DocumentID})
	if err != nil {
		return nil, &NativeDOCXFontResolutionError{Code: "PACKAGE_REFUSED", Message: err.Error()}
	}
	expected, _ := EncodeNativeDOCXFontInventoryV1(inventory)
	actual, _ := EncodeNativeDOCXFontInventoryV1(rebuilt)
	if !bytes.Equal(expected, actual) {
		return nil, &NativeDOCXFontResolutionError{Code: "REVISION_MISMATCH", Message: "manifest does not exactly attest the supplied package"}
	}
	pkg, err := openNativeDOCXPackage(data)
	if err != nil {
		return nil, &NativeDOCXFontResolutionError{Code: "PACKAGE_REFUSED", Message: err.Error()}
	}
	byFace := make(map[string]*NativeDOCXFontFaceV1)
	for familyIndex := range rebuilt.Families {
		for faceIndex := range rebuilt.Families[familyIndex].Faces {
			face := &rebuilt.Families[familyIndex].Faces[faceIndex]
			byFace[face.FaceID] = face
		}
	}
	assets := make([]NativeDOCXPagePaintFontAssetV1, 0, len(byFace))
	if rebuilt.NativeTextManifest == nil {
		return assets, nil
	}
	for _, manifestFace := range rebuilt.NativeTextManifest.Faces {
		face := byFace[manifestFace.FaceID]
		resolved, err := nativeDOCXResolvedFontAssetFromPackage(pkg, face)
		if err != nil {
			return nil, err
		}
		assets = append(assets, NativeDOCXPagePaintFontAssetV1{
			FaceID: face.FaceID, FaceSlot: face.Source.FaceSlot, ResourceID: face.Source.ResourceID,
			ContentDigest: face.Source.ContentSHA256, CollectionIndex: face.Source.CollectionIndex,
			BytesBase64: base64.StdEncoding.EncodeToString(resolved.Bytes),
		})
	}
	return assets, nil
}

func nativeDOCXResolvedFontAssetFromPackage(pkg *nativePackage, selected *NativeDOCXFontFaceV1) (*NativeDOCXResolvedFontAssetV1, error) {
	if selected == nil {
		return nil, &NativeDOCXFontResolutionError{Code: "FONT_UNAVAILABLE", Message: "the exact requested document-embedded face is unavailable"}
	}
	stored := pkg.files[selected.Source.AssetPart]
	if len(stored) != selected.Source.StoredByteLength || nativeSHA(stored) != selected.Source.StoredSHA256 {
		return nil, &NativeDOCXFontResolutionError{Code: "DIGEST_MISMATCH", Message: "stored font asset does not match the manifest"}
	}
	content, err := nativeDOCXDeobfuscateFont(stored, selected.Source.Obfuscation.FontKey)
	if err != nil || nativeSHA(content) != selected.Source.ContentSHA256 {
		return nil, &NativeDOCXFontResolutionError{Code: "DIGEST_MISMATCH", Message: "decoded font asset does not match the manifest"}
	}
	if err := nativeDOCXValidateSFNTEnvelope(content); err != nil {
		return nil, &NativeDOCXFontResolutionError{Code: "MALFORMED_FONT", Message: err.Error()}
	}
	return &NativeDOCXResolvedFontAssetV1{Face: *selected, Bytes: content, Licensing: selected.Source.Licensing}, nil
}
