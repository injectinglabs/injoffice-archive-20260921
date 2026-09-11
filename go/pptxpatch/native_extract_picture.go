package pptxpatch

import (
	"encoding/xml"
	"fmt"
	"mime"
	"strconv"
	"strings"
)

type nativePictureGap struct {
	code    string
	message string
}

type nativePictureGapSet struct {
	values []nativePictureGap
	seen   map[string]bool
}

// nativePicturePartInspection is package-scoped work state, not projected
// output state. It deliberately survives a refused group's rollback so many
// groups referencing one large media part cannot force the extractor to hash
// the same bounded package bytes repeatedly.
type nativePicturePartInspection struct {
	part       string
	digest     string
	byteLength int64
}

func (gaps *nativePictureGapSet) add(code, message string) {
	if gaps.seen == nil {
		gaps.seen = map[string]bool{}
	}
	if gaps.seen[code] {
		return
	}
	gaps.seen[code] = true
	gaps.values = append(gaps.values, nativePictureGap{code: code, message: message})
}

func (extractor *nativeExtractor) extractPicture(node *nativeXMLNode, slidePart, slideID string, relationships []nativeExtractRelationship, dialect nativeExtractDialect) (NativeElement, error) {
	if node == nil || node.Name != (xml.Name{Space: dialect.presentation, Local: "pic"}) {
		return NativeElement{}, fmt.Errorf("pptxpatch: native extract: invalid picture root")
	}
	gaps := nativePictureGapSet{}
	if err := requireOnlyNativeAttrs(node); err != nil {
		gaps.add("pptx.picture-markup-unavailable", "picture root attributes are not representable in native PPTX v1")
	}
	if err := requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.presentation, Local: "nvPicPr"},
		xml.Name{Space: dialect.presentation, Local: "blipFill"},
		xml.Name{Space: dialect.presentation, Local: "spPr"}); err != nil {
		gaps.add("pptx.picture-markup-unavailable", "picture contains unmodeled children or direct text")
	}

	nonVisual, err := nativeSingleton(node, dialect.presentation, "nvPicPr", true)
	if err != nil {
		return NativeElement{}, err
	}
	blipFill, err := nativeSingleton(node, dialect.presentation, "blipFill", true)
	if err != nil {
		return NativeElement{}, err
	}
	shapeProperties, err := nativeSingleton(node, dialect.presentation, "spPr", true)
	if err != nil {
		return NativeElement{}, err
	}

	objectID, name, err := validateNativePictureNonVisual(nonVisual, dialect, &gaps)
	if err != nil {
		return NativeElement{}, err
	}
	var crop *NativePictureCrop
	relationshipID, linkRelationshipID, err := validateNativePictureBlipFill(blipFill, dialect, &gaps, &crop)
	if err != nil {
		return NativeElement{}, err
	}
	transform, err := validateNativePictureShapeProperties(shapeProperties, dialect, &gaps)
	if err != nil {
		return NativeElement{}, err
	}

	relationship, err := exactNativePictureRelationship(relationships, relationshipID, dialect)
	if err != nil {
		return NativeElement{}, err
	}
	if linkRelationshipID != "" {
		if err := exactNativePictureExternalRelationship(relationships, linkRelationshipID, dialect); err != nil {
			return NativeElement{}, err
		}
	}
	contentType, supportedMIME, err := nativePictureContentType(extractor.pkg.contentTypes.forPart(relationship.Part))
	if err != nil {
		return NativeElement{}, err
	}
	if !supportedMIME {
		gaps.add("pptx.picture-content-type-unavailable", "picture image MIME is preserved but is outside the native PNG/JPEG preview subset")
	}
	assetID, err := extractor.nativePictureAsset(relationship.Part, contentType)
	if err != nil {
		return NativeElement{}, err
	}

	raw, err := rawNativeNode(extractor.pkg.parts[slidePart], node)
	if err != nil {
		return NativeElement{}, err
	}
	fingerprint := nativeSHA256(raw)
	elementID := extractor.identities.elements[nativeIdentityKey(slidePart, objectID)]
	if elementID == "" {
		elementID = stableNativeID("element", extractor.documentID+"\x00"+slidePart, objectID)
	}
	relID := relationshipID
	element := NativeElement{
		Kind: NativeElementKindPicture, ID: elementID, Provenance: NativeProvenanceParsed,
		Transform: transform, AssetID: &assetID, Crop: crop, Passthrough: []NativePassthroughRef{}, Children: nil,
		Source:        &NativeSourceAnchor{PartName: slidePart, ObjectID: objectID, RelationshipID: &relID, FingerprintSHA256: fingerprint},
		Compatibility: NativeCompatibility{Status: NativeCompatibilityStatusEditable, Diagnostics: []NativeDiagnostic{}},
	}
	if name != "" {
		element.Name = stringPointer(name)
	}
	if len(gaps.values) == 0 {
		return element, nil
	}
	passthrough, err := extractor.issuePassthrough(slidePart, objectID, fingerprint, raw, "pptx.picture-preserve-only")
	if err != nil {
		return NativeElement{}, err
	}
	if err := extractor.reserveNativePassthroughReference(); err != nil {
		return NativeElement{}, err
	}
	element.Passthrough = append(element.Passthrough, passthrough)
	element.Compatibility.Status = NativeCompatibilityStatusPreserveOnly
	partName := slidePart
	for _, gap := range gaps.values {
		element.Compatibility.Diagnostics = append(element.Compatibility.Diagnostics, NativeDiagnostic{
			Severity: NativeDiagnosticSeverityWarning,
			Code:     gap.code,
			Message:  gap.message,
			Scope:    &NativeDiagnosticScope{SlideID: &slideID, ElementID: &elementID, PartName: &partName},
		})
	}
	return element, nil
}

func validateNativePictureNonVisual(node *nativeXMLNode, dialect nativeExtractDialect, gaps *nativePictureGapSet) (string, string, error) {
	if err := requireOnlyNativeAttrs(node); err != nil {
		gaps.add("pptx.picture-nonvisual-unavailable", "picture nonvisual metadata is not fully modeled in native PPTX v1")
	}
	if err := requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.presentation, Local: "cNvPr"},
		xml.Name{Space: dialect.presentation, Local: "cNvPicPr"},
		xml.Name{Space: dialect.presentation, Local: "nvPr"}); err != nil {
		gaps.add("pptx.picture-nonvisual-unavailable", "picture nonvisual metadata is not fully modeled in native PPTX v1")
	}
	cNvPr, err := nativeSingleton(node, dialect.presentation, "cNvPr", true)
	if err != nil {
		return "", "", err
	}
	cNvPicPr, err := nativeSingleton(node, dialect.presentation, "cNvPicPr", true)
	if err != nil {
		return "", "", err
	}
	nvPr, err := nativeSingleton(node, dialect.presentation, "nvPr", true)
	if err != nil {
		return "", "", err
	}
	if err := requireOnlyNativeAttrs(cNvPr, xml.Name{Local: "id"}, xml.Name{Local: "name"}); err != nil || len(cNvPr.Children) != 0 || !onlyNativeXMLSpace(cNvPr.Text) {
		gaps.add("pptx.picture-nonvisual-unavailable", "picture accessibility, hyperlink, or extension metadata is not modeled in native PPTX v1")
	}
	nativeID, err := canonicalNativeUnsignedID(cNvPr, "", "id", 1)
	if err != nil {
		return "", "", err
	}
	name, _ := exactNativeAttr(cNvPr, "", "name")
	if utf16CodeUnitLengthBounded(name, 1025) > 1024 {
		return "", "", fmt.Errorf("pptxpatch: native extract: picture name exceeds contract bound")
	}
	if err := requireEmptyNativeElement(cNvPicPr); err != nil {
		gaps.add("pptx.picture-nonvisual-unavailable", "picture locks and nonvisual properties are not modeled in native PPTX v1")
	}
	if err := requireEmptyNativeElement(nvPr); err != nil {
		gaps.add("pptx.picture-inheritance-unavailable", "picture placeholder or inherited nonvisual properties are not modeled in native PPTX v1")
	}
	return "cNvPr-" + nativeID, name, nil
}

func validateNativePictureBlipFill(node *nativeXMLNode, dialect nativeExtractDialect, gaps *nativePictureGapSet, crop **NativePictureCrop) (string, string, error) {
	if err := requireOnlyNativeAttrs(node); err != nil {
		gaps.add("pptx.picture-fill-unavailable", "picture fill attributes are not modeled in native PPTX v1")
	}
	if err := requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.drawing, Local: "blip"},
		xml.Name{Space: dialect.drawing, Local: "srcRect"},
		xml.Name{Space: dialect.drawing, Local: "stretch"},
		xml.Name{Space: dialect.drawing, Local: "tile"}); err != nil {
		gaps.add("pptx.picture-fill-unavailable", "picture fill contains unmodeled markup")
	}
	blip, err := nativeSingleton(node, dialect.drawing, "blip", true)
	if err != nil {
		return "", "", err
	}
	sourceRect, err := nativeSingleton(node, dialect.drawing, "srcRect", false)
	if err != nil {
		return "", "", err
	}
	stretch, err := nativeSingleton(node, dialect.drawing, "stretch", false)
	if err != nil {
		return "", "", err
	}
	tile, err := nativeSingleton(node, dialect.drawing, "tile", false)
	if err != nil {
		return "", "", err
	}
	if stretch != nil && tile != nil {
		return "", "", fmt.Errorf("pptxpatch: native extract: picture fill cannot be both stretched and tiled")
	}
	if err := requireOnlyNativeAttrs(blip,
		xml.Name{Space: dialect.rels, Local: "embed"},
		xml.Name{Space: dialect.rels, Local: "link"}); err != nil {
		gaps.add("pptx.picture-effects-unavailable", "picture blip attributes are not modeled in native PPTX v1")
	}
	if len(blip.Children) != 0 || !onlyNativeXMLSpace(blip.Text) {
		gaps.add("pptx.picture-effects-unavailable", "picture effects or blip extensions are not modeled in native PPTX v1")
	}
	embed, embedOK := exactNativeAttr(blip, dialect.rels, "embed")
	link, linkOK := exactNativeAttr(blip, dialect.rels, "link")
	if !embedOK || embed == "" || !nativeIDPattern.MatchString(embed) {
		return "", "", fmt.Errorf("pptxpatch: native extract: picture requires one valid embedded image relationship")
	}
	if linkOK {
		if link == "" || !nativeIDPattern.MatchString(link) {
			return "", "", fmt.Errorf("pptxpatch: native extract: picture has invalid external link relationship id")
		}
		gaps.add("pptx.picture-external-link-unavailable", "linked picture semantics are preserved but not modeled in native PPTX v1")
	}
	if sourceRect != nil {
		insets, rectErr := validateNativePictureSourceRect(sourceRect)
		if rectErr != nil {
			return "", "", rectErr
		}
		if insets[0] < 0 || insets[1] < 0 || insets[2] < 0 || insets[3] < 0 || insets[0]+insets[2] >= 100_000 || insets[1]+insets[3] >= 100_000 {
			gaps.add("pptx.picture-crop-unavailable", "outset or degenerate picture crops are preserved but not modeled")
		} else if insets != [4]int64{} {
			*crop = &NativePictureCrop{Left: &insets[0], Top: &insets[1], Right: &insets[2], Bottom: &insets[3]}
		}
	}
	if tile != nil {
		gaps.add("pptx.picture-fill-unavailable", "tiled picture fill is preserved but cannot be represented in native PPTX v1")
	}
	if stretch == nil {
		gaps.add("pptx.picture-fill-unavailable", "picture fill mode is not explicitly representable in native PPTX v1")
	} else {
		if err := requireOnlyNativeAttrs(stretch); err != nil {
			gaps.add("pptx.picture-fill-unavailable", "picture stretch metadata is not modeled in native PPTX v1")
		}
		// DrawingML CT_StretchInfoProperties permits an omitted fillRect;
		// it denotes the complete destination rectangle, just like <fillRect/>.
		fillRect, fillErr := nativeSingleton(stretch, dialect.drawing, "fillRect", false)
		if fillErr != nil {
			return "", "", fillErr
		}
		if err := requireOnlyNativeChildren(stretch, xml.Name{Space: dialect.drawing, Local: "fillRect"}); err != nil || fillRect != nil && requireEmptyNativeElement(fillRect) != nil || !onlyNativeXMLSpace(stretch.Text) {
			gaps.add("pptx.picture-fill-unavailable", "picture stretch rectangle metadata is not modeled in native PPTX v1")
		}
	}
	return embed, link, nil
}

func validateNativePictureSourceRect(node *nativeXMLNode) ([4]int64, error) {
	var insets [4]int64
	if err := requireOnlyNativeAttrs(node, xml.Name{Local: "l"}, xml.Name{Local: "t"}, xml.Name{Local: "r"}, xml.Name{Local: "b"}); err != nil {
		return insets, fmt.Errorf("pptxpatch: native extract: unsupported picture crop metadata: %w", err)
	}
	if len(node.Children) != 0 || !onlyNativeXMLSpace(node.Text) {
		return insets, fmt.Errorf("pptxpatch: native extract: picture source rectangle must be empty")
	}
	for index, name := range []string{"l", "t", "r", "b"} {
		value, ok := exactNativeAttr(node, "", name)
		if !ok {
			continue
		}
		parsed, err := parseCanonicalNativeInt(value, -100_000, 100_000)
		if err != nil {
			return insets, fmt.Errorf("pptxpatch: native extract: invalid picture crop %s", name)
		}
		insets[index] = parsed
	}
	return insets, nil
}

func validateNativePictureShapeProperties(node *nativeXMLNode, dialect nativeExtractDialect, gaps *nativePictureGapSet) (NativeTransform, error) {
	if err := requireOnlyNativeAttrs(node); err != nil {
		gaps.add("pptx.picture-shape-unavailable", "picture shape properties contain unmodeled attributes")
	}
	if err := requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.drawing, Local: "xfrm"},
		xml.Name{Space: dialect.drawing, Local: "prstGeom"}); err != nil {
		gaps.add("pptx.picture-effects-unavailable", "picture shape effects, outline, or extensions are not modeled in native PPTX v1")
	}
	xfrm, err := nativeSingleton(node, dialect.drawing, "xfrm", true)
	if err != nil {
		return NativeTransform{}, err
	}
	geometry, err := nativeSingleton(node, dialect.drawing, "prstGeom", false)
	if err != nil {
		return NativeTransform{}, err
	}
	if geometry == nil {
		gaps.add("pptx.picture-geometry-unavailable", "picture rectangle geometry is not explicit in native PPTX v1")
	} else {
		preset, presetOK := exactNativeAttr(geometry, "", "prst")
		if err := requireOnlyNativeAttrs(geometry, xml.Name{Local: "prst"}); err != nil || !presetOK || preset != "rect" {
			gaps.add("pptx.picture-geometry-unavailable", "non-rectangle picture geometry is preserved but cannot be represented in native PPTX v1")
		}
		adjustments, adjustmentsErr := nativeSingleton(geometry, dialect.drawing, "avLst", true)
		if adjustmentsErr != nil {
			return NativeTransform{}, adjustmentsErr
		}
		if err := requireOnlyNativeChildren(geometry, xml.Name{Space: dialect.drawing, Local: "avLst"}); err != nil || requireEmptyNativeElement(adjustments) != nil {
			gaps.add("pptx.picture-geometry-unavailable", "picture geometry adjustments are not modeled in native PPTX v1")
		}
	}
	return validateNativePictureTransform(xfrm, dialect, gaps)
}

func validateNativePictureTransform(node *nativeXMLNode, dialect nativeExtractDialect, gaps *nativePictureGapSet) (NativeTransform, error) {
	if err := requireOnlyNativeAttrs(node, xml.Name{Local: "rot"}, xml.Name{Local: "flipH"}, xml.Name{Local: "flipV"}); err != nil {
		gaps.add("pptx.picture-transform-unavailable", "picture transform contains unmodeled attributes")
	}
	if err := requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.drawing, Local: "off"},
		xml.Name{Space: dialect.drawing, Local: "ext"}); err != nil {
		return NativeTransform{}, fmt.Errorf("pptxpatch: native extract: invalid picture transform children: %w", err)
	}
	off, err := nativeSingleton(node, dialect.drawing, "off", true)
	if err != nil {
		return NativeTransform{}, err
	}
	ext, err := nativeSingleton(node, dialect.drawing, "ext", true)
	if err != nil {
		return NativeTransform{}, err
	}
	if err := requireOnlyNativeAttrs(off, xml.Name{Local: "x"}, xml.Name{Local: "y"}); err != nil || requireOnlyNativeChildren(off) != nil {
		return NativeTransform{}, fmt.Errorf("pptxpatch: native extract: invalid picture offset")
	}
	if err := requireOnlyNativeAttrs(ext, xml.Name{Local: "cx"}, xml.Name{Local: "cy"}); err != nil || requireOnlyNativeChildren(ext) != nil {
		return NativeTransform{}, fmt.Errorf("pptxpatch: native extract: invalid picture extent")
	}
	x, err := requiredNativeInt64(off, "", "x")
	if err != nil {
		return NativeTransform{}, err
	}
	y, err := requiredNativeInt64(off, "", "y")
	if err != nil {
		return NativeTransform{}, err
	}
	cx, err := requiredNativePositiveInt64(ext, "", "cx")
	if err != nil {
		return NativeTransform{}, err
	}
	cy, err := requiredNativePositiveInt64(ext, "", "cy")
	if err != nil {
		return NativeTransform{}, err
	}
	if value, ok := exactNativeAttr(node, "", "rot"); ok {
		rotation, rotationErr := parseCanonicalNativeInt(value, -nativeMaxSafeInteger, nativeMaxSafeInteger)
		if rotationErr != nil {
			return NativeTransform{}, fmt.Errorf("pptxpatch: native extract: invalid picture rotation")
		}
		if rotation != 0 {
			gaps.add("pptx.picture-transform-unavailable", "nonzero picture rotation is preserved but cannot be represented in native PPTX v1")
		}
	}
	for _, name := range []string{"flipH", "flipV"} {
		if value, ok := exactNativeAttr(node, "", name); ok {
			flip, flipErr := nativeBool(value)
			if flipErr != nil {
				return NativeTransform{}, flipErr
			}
			if flip {
				gaps.add("pptx.picture-transform-unavailable", "picture flips are preserved but cannot be represented in native PPTX v1")
			}
		}
	}
	return NativeTransform{X: int64Pointer(x), Y: int64Pointer(y), Cx: int64Pointer(cx), Cy: int64Pointer(cy)}, nil
}

func exactNativePictureRelationship(relationships []nativeExtractRelationship, relationshipID string, dialect nativeExtractDialect) (nativeExtractRelationship, error) {
	var selected *nativeExtractRelationship
	for index := range relationships {
		if relationships[index].ID != relationshipID {
			continue
		}
		if selected != nil {
			return nativeExtractRelationship{}, fmt.Errorf("pptxpatch: native extract OPC: ambiguous duplicate picture relationship id")
		}
		selected = &relationships[index]
	}
	if selected == nil || selected.Type != dialect.relImage || !selected.internal() || selected.Part == "" {
		return nativeExtractRelationship{}, fmt.Errorf("pptxpatch: native extract OPC: picture relationship is missing, external, or not the exact image type")
	}
	return *selected, nil
}

func exactNativePictureExternalRelationship(relationships []nativeExtractRelationship, relationshipID string, dialect nativeExtractDialect) error {
	var selected *nativeExtractRelationship
	for index := range relationships {
		if relationships[index].ID != relationshipID {
			continue
		}
		if selected != nil {
			return fmt.Errorf("pptxpatch: native extract OPC: ambiguous duplicate linked-picture relationship id")
		}
		selected = &relationships[index]
	}
	if selected == nil || selected.Type != dialect.relImage || selected.TargetMode != "External" || selected.Part != "" || selected.Target == "" {
		return fmt.Errorf("pptxpatch: native extract OPC: linked picture relationship is missing, internal, or not the exact image type")
	}
	return nil
}

func nativePictureContentType(value string) (string, bool, error) {
	mediaType, parameters, err := mime.ParseMediaType(value)
	mediaType = asciiLowerNative(mediaType)
	if err != nil || len(parameters) != 0 || !strings.HasPrefix(mediaType, "image/") || len(mediaType) > 256 {
		return "", false, fmt.Errorf("pptxpatch: native extract OPC: picture target lacks a parameter-free image MIME")
	}
	supported := mediaType == "image/png" || mediaType == "image/jpeg"
	return mediaType, supported, nil
}

func (extractor *nativeExtractor) nativePictureAsset(part, contentType string) (string, error) {
	alias, err := nativePartAlias(part)
	if err != nil {
		return "", fmt.Errorf("pptxpatch: native extract OPC: invalid picture asset part: %w", err)
	}
	payload, ok := extractor.pkg.parts[part]
	if !ok {
		return "", fmt.Errorf("pptxpatch: native extract OPC: missing picture asset part")
	}
	inspection, exists := extractor.picturePartInspections[alias]
	if !exists {
		length := int64(len(payload))
		if length < 0 || length > nativeExtractMaxTotalMediaBytes-extractor.mediaBytesInspected {
			return "", fmt.Errorf("pptxpatch: native extract: cumulative picture inspection byte budget exceeded")
		}
		inspection = nativePicturePartInspection{part: part, digest: nativeSHA256(payload), byteLength: length}
		extractor.picturePartInspections[alias] = inspection
		extractor.mediaBytesInspected += length
	} else if inspection.part != part || inspection.byteLength != int64(len(payload)) {
		return "", fmt.Errorf("pptxpatch: native extract: picture part inspection identity is inconsistent")
	}
	if index, exists := extractor.assetByAlias[alias]; exists {
		asset := extractor.assets[index]
		if asset.Source == nil || nativeIdentityKey(asset.Source.PartName, asset.Source.ObjectID) != nativeIdentityKey(part, "asset-part") || asset.ContentType != contentType || asset.SHA256 != inspection.digest || asset.ByteLength == nil || *asset.ByteLength != inspection.byteLength {
			return "", fmt.Errorf("pptxpatch: native extract: shared picture asset identity is inconsistent")
		}
		return asset.ID, nil
	}
	if len(extractor.assets) >= nativeMaxAssets {
		return "", fmt.Errorf("pptxpatch: native extract: asset count exceeds %d", nativeMaxAssets)
	}
	if err := extractor.reserveNativeAssetBudget(int64(len(payload)), 0); err != nil {
		return "", err
	}
	objectID := "asset-part"
	assetID := extractor.identities.assets[nativeIdentityKey(part, objectID)]
	if assetID == "" {
		assetID = stableNativeID("asset", extractor.documentID+"\x00"+alias, objectID)
	}
	if owner, exists := extractor.assetIDOwners[assetID]; exists && owner != alias {
		return "", fmt.Errorf("pptxpatch: native extract: distinct picture parts collide on one asset id")
	}
	if _, exists := extractor.assetIDOwners[assetID]; !exists {
		extractor.assetIDOwnerJournal = append(extractor.assetIDOwnerJournal, assetID)
	}
	extractor.assetIDOwners[assetID] = alias
	digest := inspection.digest
	length := inspection.byteLength
	capability, err := extractor.issuePassthrough(part, objectID, digest, payload, "pptx.picture-asset-source")
	if err != nil {
		return "", err
	}
	if err := extractor.reserveNativePassthroughReference(); err != nil {
		return "", err
	}
	asset := NativeAsset{
		ID: assetID, Provenance: NativeProvenanceParsed, ContentType: contentType,
		SHA256: digest, ByteLength: &length, Passthrough: []NativePassthroughRef{capability},
		Source: &NativeSourceAnchor{PartName: part, ObjectID: objectID, FingerprintSHA256: digest},
	}
	if _, exists := extractor.assetByAlias[alias]; !exists {
		extractor.assetAliasJournal = append(extractor.assetAliasJournal, alias)
	}
	extractor.assetByAlias[alias] = len(extractor.assets)
	extractor.assets = append(extractor.assets, asset)
	return assetID, nil
}

func (extractor *nativeExtractor) reserveNativeAssetBudget(mediaBytes, base64CodeUnits int64) error {
	if mediaBytes < 0 || mediaBytes > nativeMaxAssetBytes || mediaBytes > nativeExtractMaxTotalMediaBytes-extractor.mediaBytesEmitted {
		return fmt.Errorf("pptxpatch: native extract: cumulative picture media byte budget exceeded")
	}
	if base64CodeUnits < 0 || base64CodeUnits > nativeMaxInlineAssetBase64CodeUnits || base64CodeUnits > nativeMaxTotalInlineAssetBase64CodeUnits-extractor.assetBase64Emitted {
		return fmt.Errorf("pptxpatch: native extract: cumulative picture base64 budget exceeded")
	}
	extractor.mediaBytesEmitted += mediaBytes
	extractor.assetBase64Emitted += base64CodeUnits
	return nil
}

func parseCanonicalNativeInt(value string, minimum, maximum int64) (int64, error) {
	if value == "" || strings.HasPrefix(value, "+") || (len(value) > 1 && value[0] == '0') || strings.HasPrefix(value, "-0") {
		return 0, fmt.Errorf("non-canonical integer")
	}
	parsed, err := strconv.ParseInt(value, 10, 64)
	if err != nil || parsed < minimum || parsed > maximum || strconv.FormatInt(parsed, 10) != value {
		return 0, fmt.Errorf("integer outside canonical range")
	}
	return parsed, nil
}
