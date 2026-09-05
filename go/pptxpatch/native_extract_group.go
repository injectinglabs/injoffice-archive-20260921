package pptxpatch

import (
	"encoding/xml"
	"errors"
	"fmt"
	"math/big"
	"strings"
)

// nativeGroupProjectionRefusal identifies a structurally valid group whose
// visual semantics cannot be represented exactly by pptx-native/v1. Callers
// preserve the complete p:grpSp subtree and must not emit any of its children.
type nativeGroupProjectionRefusal struct {
	code    string
	message string
}

func (refusal nativeGroupProjectionRefusal) Error() string {
	return "pptxpatch: native extract: " + refusal.message
}

func refuseNativeGroup(code, message string) error {
	return nativeGroupProjectionRefusal{code: code, message: message}
}

type nativeGroupTransform struct {
	x, y    int64
	cx, cy  int64
	childX  int64
	childY  int64
	childCx int64
	childCy int64
}

type nativeGroupExtractResult struct {
	element                  NativeElement
	usedPictureRelationships map[string]bool
}

type nativeGroupProjectionSnapshot struct {
	assetCount              int
	mediaBytesEmitted       int64
	assetBase64Emitted      int64
	textCodeUnitsEmitted    int64
	tableCellsEmitted       int
	outputNodesEmitted      int
	passthroughRefsEmitted  int
	passthroughUsed         int
	passthroughCacheJournal int
	tokenOwnerJournal       int
	assetAliasJournal       int
	assetIDOwnerJournal     int
	stagedRequestCount      int
	stagedPayloadBytes      int64
}

func (extractor *nativeExtractor) snapshotNativeGroupProjection() nativeGroupProjectionSnapshot {
	return nativeGroupProjectionSnapshot{
		assetCount: len(extractor.assets), mediaBytesEmitted: extractor.mediaBytesEmitted,
		assetBase64Emitted: extractor.assetBase64Emitted, textCodeUnitsEmitted: extractor.textCodeUnitsEmitted,
		tableCellsEmitted:  extractor.tableCellsEmitted,
		outputNodesEmitted: extractor.outputNodesEmitted, passthroughRefsEmitted: extractor.passthroughRefsEmitted,
		passthroughUsed:         extractor.passthroughUsed,
		passthroughCacheJournal: len(extractor.passthroughCacheJournal), tokenOwnerJournal: len(extractor.tokenOwnerJournal),
		assetAliasJournal: len(extractor.assetAliasJournal), assetIDOwnerJournal: len(extractor.assetIDOwnerJournal),
		stagedRequestCount: len(extractor.tokenStager.requests), stagedPayloadBytes: extractor.tokenStager.payloadBytes,
	}
}

// rollbackNativeGroupProjection removes every projected-output side effect of a
// group attempt. Package-scoped media inspection work/cache intentionally stays
// monotonic across rollback. Map journals keep this operation proportional to
// the attempted subtree rather than cloning every previously extracted object.
func (extractor *nativeExtractor) rollbackNativeGroupProjection(snapshot nativeGroupProjectionSnapshot) {
	extractor.assets = extractor.assets[:snapshot.assetCount]
	extractor.mediaBytesEmitted = snapshot.mediaBytesEmitted
	extractor.assetBase64Emitted = snapshot.assetBase64Emitted
	extractor.textCodeUnitsEmitted = snapshot.textCodeUnitsEmitted
	extractor.tableCellsEmitted = snapshot.tableCellsEmitted
	extractor.outputNodesEmitted = snapshot.outputNodesEmitted
	extractor.passthroughRefsEmitted = snapshot.passthroughRefsEmitted
	extractor.passthroughUsed = snapshot.passthroughUsed
	rollbackNativeStringMap(extractor.passthroughCache, &extractor.passthroughCacheJournal, snapshot.passthroughCacheJournal)
	rollbackNativeStringMap(extractor.tokenOwners, &extractor.tokenOwnerJournal, snapshot.tokenOwnerJournal)
	rollbackNativeStringMap(extractor.assetByAlias, &extractor.assetAliasJournal, snapshot.assetAliasJournal)
	rollbackNativeStringMap(extractor.assetIDOwners, &extractor.assetIDOwnerJournal, snapshot.assetIDOwnerJournal)
	extractor.tokenStager.rollback(snapshot.stagedRequestCount, snapshot.stagedPayloadBytes)
}

func rollbackNativeStringMap[V any](values map[string]V, journal *[]string, keep int) {
	for index := len(*journal) - 1; index >= keep; index-- {
		delete(values, (*journal)[index])
	}
	clear((*journal)[keep:])
	*journal = (*journal)[:keep]
}

// extractNativeGroupAtomically stages one complete subtree exactly once. A
// refusal rolls back every staged token, asset, identity alias, and output
// budget, while retaining package-scoped work accounting. No external
// capability is issued until the fully validated deck commits its
// NativePassthroughTokenTransaction.
func (extractor *nativeExtractor) extractNativeGroupAtomically(
	node *nativeXMLNode,
	slidePart string,
	slideID string,
	slideFingerprint string,
	dialect nativeExtractDialect,
	relationships []nativeExtractRelationship,
	diagnosticBudget int,
) (nativeGroupExtractResult, error) {
	extractor.groupProjectionSeen = true
	snapshot := extractor.snapshotNativeGroupProjection()
	result, err := extractor.extractNativeGroup(node, slidePart, slideID, slideFingerprint, dialect, relationships, 1)
	if err != nil {
		extractor.rollbackNativeGroupProjection(snapshot)
		return nativeGroupExtractResult{}, err
	}
	if diagnosticBudget < 0 || len(result.element.Compatibility.Diagnostics) > diagnosticBudget {
		extractor.rollbackNativeGroupProjection(snapshot)
		return nativeGroupExtractResult{}, refuseNativeGroup("pptx.group-diagnostic-budget-unavailable", "group diagnostics exceed the remaining bounded slide diagnostic budget")
	}
	return result, nil
}

func (extractor *nativeExtractor) extractNativeGroup(
	node *nativeXMLNode,
	slidePart string,
	slideID string,
	slideFingerprint string,
	dialect nativeExtractDialect,
	relationships []nativeExtractRelationship,
	depth int,
) (nativeGroupExtractResult, error) {
	if depth > nativeMaxDepth {
		return nativeGroupExtractResult{}, fmt.Errorf("pptxpatch: native extract: group depth exceeds %d", nativeMaxDepth)
	}
	if node == nil || node.Name != (xml.Name{Space: dialect.presentation, Local: "grpSp"}) {
		return nativeGroupExtractResult{}, fmt.Errorf("pptxpatch: native extract: invalid group root")
	}
	if err := rejectNativeGroupDialectMix(node, dialect); err != nil {
		return nativeGroupExtractResult{}, err
	}
	if err := requireOnlyNativeAttrs(node); err != nil {
		return nativeGroupExtractResult{}, refuseNativeGroup("pptx.group-markup-unavailable", "group root attributes are outside the exact native subset")
	}
	if err := requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.presentation, Local: "nvGrpSpPr"},
		xml.Name{Space: dialect.presentation, Local: "grpSpPr"},
		xml.Name{Space: dialect.presentation, Local: "sp"},
		xml.Name{Space: dialect.presentation, Local: "pic"},
		xml.Name{Space: dialect.presentation, Local: "cxnSp"},
		xml.Name{Space: dialect.presentation, Local: "graphicFrame"},
		xml.Name{Space: dialect.presentation, Local: "grpSp"}); err != nil {
		return nativeGroupExtractResult{}, refuseNativeGroup("pptx.group-child-unavailable", "group contains unsupported children, attributes, or direct text")
	}

	nonVisual, err := nativeSingleton(node, dialect.presentation, "nvGrpSpPr", true)
	if err != nil {
		return nativeGroupExtractResult{}, err
	}
	properties, err := nativeSingleton(node, dialect.presentation, "grpSpPr", true)
	if err != nil {
		return nativeGroupExtractResult{}, err
	}
	if len(node.Children) < 3 || node.Children[0] != nonVisual || node.Children[1] != properties {
		return nativeGroupExtractResult{}, fmt.Errorf("pptxpatch: native extract: malformed group child order")
	}
	objectID, name, err := validateNativeGroupNonVisual(nonVisual, dialect)
	if err != nil {
		var refusal nativeGroupProjectionRefusal
		if errors.As(err, &refusal) {
			return nativeGroupExtractResult{}, err
		}
		return nativeGroupExtractResult{}, fmt.Errorf("pptxpatch: native extract: invalid group nonvisual properties: %w", err)
	}
	transform, err := validateNativeGroupTransform(properties, dialect)
	if err != nil {
		return nativeGroupExtractResult{}, err
	}

	drawableChildren := 0
	for _, child := range node.Children {
		if child != nonVisual && child != properties {
			drawableChildren++
		}
	}
	if drawableChildren == 0 {
		return nativeGroupExtractResult{}, refuseNativeGroup("pptx.group-empty-unavailable", "empty groups are preserved but not projected by native PPTX v1")
	}
	if drawableChildren > nativeMaxElementsPerContainer {
		return nativeGroupExtractResult{}, fmt.Errorf("pptxpatch: native extract: group child count exceeds %d", nativeMaxElementsPerContainer)
	}

	children := make([]NativeElement, 0, drawableChildren)
	usedPictures := map[string]bool{}
	worst := NativeCompatibilityStatusEditable
	diagnostics := []NativeDiagnostic{}
	for _, child := range node.Children {
		if child == nonVisual || child == properties {
			continue
		}
		hidden, hiddenErr := nativeGroupChildHidden(child, dialect)
		if hiddenErr != nil {
			return nativeGroupExtractResult{}, hiddenErr
		}
		if hidden {
			return nativeGroupExtractResult{}, refuseNativeGroup("pptx.group-hidden-child-unavailable", "group contains a hidden descendant and is preserved rather than partially rendered")
		}
		var element NativeElement
		switch child.Name {
		case xml.Name{Space: dialect.presentation, Local: "sp"}:
			textBox, textBoxErr := nativeShapeIsTextBox(child, dialect)
			if textBoxErr != nil {
				return nativeGroupExtractResult{}, textBoxErr
			}
			if textBox {
				element, err = extractor.extractTextShape(child, slidePart, slideID, slideFingerprint, len(children), dialect)
				if err != nil {
					var duplicate nativeDuplicateSingletonError
					if errors.As(err, &duplicate) {
						return nativeGroupExtractResult{}, err
					}
					return nativeGroupExtractResult{}, refuseNativeGroup("pptx.group-text-unavailable", "group text child is outside the exact native subset")
				}
			} else {
				element, err = extractor.extractAutoShape(child, slidePart, slideID, dialect)
				if err != nil {
					return nativeGroupExtractResult{}, err
				}
			}
		case xml.Name{Space: dialect.presentation, Local: "pic"}:
			element, err = extractor.extractPicture(child, slidePart, slideID, relationships, dialect)
			if err != nil {
				return nativeGroupExtractResult{}, err
			}
			if element.Source != nil && element.Source.RelationshipID != nil {
				usedPictures[*element.Source.RelationshipID] = true
			}
		case xml.Name{Space: dialect.presentation, Local: "cxnSp"}:
			element, err = extractor.extractConnector(child, slidePart, slideID, dialect)
			if err != nil {
				return nativeGroupExtractResult{}, err
			}
		case xml.Name{Space: dialect.presentation, Local: "graphicFrame"}:
			element, err = extractor.extractNativeGraphicFrame(child, slidePart, slideID, relationships, dialect)
			if err != nil {
				var refusal nativeGraphicFrameProjectionRefusal
				if errors.As(err, &refusal) {
					code := "pptx.group-table-unavailable"
					if strings.HasPrefix(refusal.code, "pptx.chart-") {
						code = "pptx.group-chart-unavailable"
					}
					return nativeGroupExtractResult{}, refuseNativeGroup(code, refusal.message)
				}
				return nativeGroupExtractResult{}, err
			}
		case xml.Name{Space: dialect.presentation, Local: "grpSp"}:
			nested, nestedErr := extractor.extractNativeGroup(child, slidePart, slideID, slideFingerprint, dialect, relationships, depth+1)
			if nestedErr != nil {
				return nativeGroupExtractResult{}, nestedErr
			}
			element = nested.element
			for relationshipID := range nested.usedPictureRelationships {
				usedPictures[relationshipID] = true
			}
		default:
			return nativeGroupExtractResult{}, refuseNativeGroup("pptx.group-child-unavailable", "group contains an unsupported shape-tree child")
		}
		if element.Compatibility.Status == NativeCompatibilityStatusRefused {
			return nativeGroupExtractResult{}, refuseNativeGroup("pptx.group-child-refused-unavailable", "group contains a refused descendant and is preserved as one exact opaque subtree")
		}
		children = append(children, element)
		worst = worseNativeStatus(worst, element.Compatibility.Status)
		if element.Compatibility.Status != NativeCompatibilityStatusEditable {
			if len(diagnostics) > nativeMaxDiagnosticsPerScope-len(element.Compatibility.Diagnostics) {
				return nativeGroupExtractResult{}, refuseNativeGroup("pptx.group-diagnostic-budget-unavailable", "group descendant diagnostics exceed the bounded native contract")
			}
			diagnostics = append(diagnostics, element.Compatibility.Diagnostics...)
		}
	}

	raw, err := rawNativeNode(extractor.pkg.parts[slidePart], node)
	if err != nil {
		return nativeGroupExtractResult{}, err
	}
	fingerprint := nativeSHA256(raw)
	elementID := extractor.identities.elements[nativeIdentityKey(slidePart, objectID)]
	if elementID == "" {
		elementID = stableNativeID("element", extractor.documentID+"\x00"+slidePart, objectID)
	}
	element := NativeElement{
		Kind: NativeElementKindGroup, ID: elementID, Provenance: NativeProvenanceParsed,
		Transform:      NativeTransform{X: int64Pointer(transform.x), Y: int64Pointer(transform.y), Cx: int64Pointer(transform.cx), Cy: int64Pointer(transform.cy)},
		ChildTransform: &NativeTransform{X: int64Pointer(transform.childX), Y: int64Pointer(transform.childY), Cx: int64Pointer(transform.childCx), Cy: int64Pointer(transform.childCy)},
		Children:       children, Passthrough: []NativePassthroughRef{},
		Source:        &NativeSourceAnchor{PartName: slidePart, ObjectID: objectID, FingerprintSHA256: fingerprint},
		Compatibility: NativeCompatibility{Status: worst, Diagnostics: diagnostics},
	}
	if name != "" {
		element.Name = stringPointer(name)
	}
	return nativeGroupExtractResult{element: element, usedPictureRelationships: usedPictures}, nil
}

func nativeGroupChildHidden(node *nativeXMLNode, dialect nativeExtractDialect) (bool, error) {
	containerName := ""
	switch node.Name {
	case xml.Name{Space: dialect.presentation, Local: "sp"}:
		containerName = "nvSpPr"
	case xml.Name{Space: dialect.presentation, Local: "pic"}:
		containerName = "nvPicPr"
	case xml.Name{Space: dialect.presentation, Local: "cxnSp"}:
		containerName = "nvCxnSpPr"
	case xml.Name{Space: dialect.presentation, Local: "graphicFrame"}:
		containerName = "nvGraphicFramePr"
	case xml.Name{Space: dialect.presentation, Local: "grpSp"}:
		containerName = "nvGrpSpPr"
	default:
		return false, nil
	}
	nonVisual, err := nativeSingleton(node, dialect.presentation, containerName, true)
	if err != nil {
		return false, err
	}
	properties, err := nativeSingleton(nonVisual, dialect.presentation, "cNvPr", true)
	if err != nil {
		return false, err
	}
	value, exists := exactNativeAttr(properties, "", "hidden")
	if !exists {
		return false, nil
	}
	hidden, err := nativeBool(value)
	if err != nil {
		return false, fmt.Errorf("pptxpatch: native extract: invalid hidden flag on grouped descendant")
	}
	return hidden, nil
}

func validateNativeGroupNonVisual(node *nativeXMLNode, dialect nativeExtractDialect) (string, string, error) {
	if err := requireOnlyNativeAttrs(node); err != nil {
		return "", "", refuseNativeGroup("pptx.group-nonvisual-unavailable", "group nonvisual properties contain unmodeled attributes")
	}
	if err := requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.presentation, Local: "cNvPr"},
		xml.Name{Space: dialect.presentation, Local: "cNvGrpSpPr"},
		xml.Name{Space: dialect.presentation, Local: "nvPr"}); err != nil {
		return "", "", refuseNativeGroup("pptx.group-nonvisual-unavailable", "group nonvisual properties are outside the exact native subset")
	}
	cNvPr, err := nativeSingleton(node, dialect.presentation, "cNvPr", true)
	if err != nil {
		return "", "", err
	}
	cNvGrpSpPr, err := nativeSingleton(node, dialect.presentation, "cNvGrpSpPr", true)
	if err != nil {
		return "", "", err
	}
	nvPr, err := nativeSingleton(node, dialect.presentation, "nvPr", true)
	if err != nil {
		return "", "", err
	}
	if len(node.Children) != 3 || node.Children[0] != cNvPr || node.Children[1] != cNvGrpSpPr || node.Children[2] != nvPr {
		return "", "", fmt.Errorf("pptxpatch: native extract: malformed group nonvisual child order")
	}
	if err := requireOnlyNativeAttrs(cNvPr, xml.Name{Local: "id"}, xml.Name{Local: "name"}); err != nil || requireOnlyNativeChildren(cNvPr) != nil {
		return "", "", refuseNativeGroup("pptx.group-nonvisual-unavailable", "group identity metadata is outside the exact native subset")
	}
	if err := requireEmptyNativeElement(cNvGrpSpPr); err != nil {
		return "", "", refuseNativeGroup("pptx.group-locks-unavailable", "group locks are preserved but not projected")
	}
	if err := requireEmptyNativeElement(nvPr); err != nil {
		return "", "", refuseNativeGroup("pptx.group-inheritance-unavailable", "group placeholder or inheritance metadata is not resolved")
	}
	id, err := canonicalNativeUnsignedID(cNvPr, "", "id", 1)
	if err != nil {
		return "", "", err
	}
	name, _ := exactNativeAttr(cNvPr, "", "name")
	if utf16CodeUnitLengthBounded(name, 1025) > 1024 {
		return "", "", refuseNativeGroup("pptx.group-name-unavailable", "group name exceeds the native contract bound")
	}
	return "cNvPr-" + id, name, nil
}

func validateNativeGroupTransform(node *nativeXMLNode, dialect nativeExtractDialect) (nativeGroupTransform, error) {
	if err := requireOnlyNativeAttrs(node); err != nil {
		return nativeGroupTransform{}, refuseNativeGroup("pptx.group-properties-unavailable", "group properties contain unsupported attributes")
	}
	if err := requireOnlyNativeChildren(node, xml.Name{Space: dialect.drawing, Local: "xfrm"}); err != nil {
		return nativeGroupTransform{}, refuseNativeGroup("pptx.group-properties-unavailable", "group fill, effects, or extension markup is outside the exact native subset")
	}
	xfrm, err := nativeSingleton(node, dialect.drawing, "xfrm", false)
	if err != nil {
		return nativeGroupTransform{}, err
	}
	if xfrm == nil {
		return nativeGroupTransform{}, refuseNativeGroup("pptx.group-transform-unavailable", "group has no explicit transform and is preserved rather than assigned inferred defaults")
	}
	if err := requireOnlyNativeAttrs(xfrm, xml.Name{Local: "rot"}, xml.Name{Local: "flipH"}, xml.Name{Local: "flipV"}); err != nil {
		return nativeGroupTransform{}, refuseNativeGroup("pptx.group-transform-unavailable", "group transform contains unsupported attributes")
	}
	if err := requireOnlyNativeChildren(xfrm,
		xml.Name{Space: dialect.drawing, Local: "off"},
		xml.Name{Space: dialect.drawing, Local: "ext"},
		xml.Name{Space: dialect.drawing, Local: "chOff"},
		xml.Name{Space: dialect.drawing, Local: "chExt"}); err != nil {
		return nativeGroupTransform{}, fmt.Errorf("pptxpatch: native extract: malformed group transform: %w", err)
	}
	if value, ok := exactNativeAttr(xfrm, "", "rot"); ok {
		rotation, parseErr := parseCanonicalNativeInt(value, -nativeMaxSafeInteger, nativeMaxSafeInteger)
		if parseErr != nil {
			return nativeGroupTransform{}, fmt.Errorf("pptxpatch: native extract: invalid group rotation")
		}
		if rotation != 0 {
			return nativeGroupTransform{}, refuseNativeGroup("pptx.group-rotation-unavailable", "rotated groups are preserved rather than approximated")
		}
	}
	for _, name := range []string{"flipH", "flipV"} {
		if value, ok := exactNativeAttr(xfrm, "", name); ok {
			flip, parseErr := nativeBool(value)
			if parseErr != nil {
				return nativeGroupTransform{}, fmt.Errorf("pptxpatch: native extract: invalid group %s", name)
			}
			if flip {
				return nativeGroupTransform{}, refuseNativeGroup("pptx.group-flip-unavailable", "flipped groups are preserved rather than approximated")
			}
		}
	}
	off, err := nativeSingleton(xfrm, dialect.drawing, "off", false)
	if err != nil {
		return nativeGroupTransform{}, err
	}
	ext, err := nativeSingleton(xfrm, dialect.drawing, "ext", false)
	if err != nil {
		return nativeGroupTransform{}, err
	}
	childOff, err := nativeSingleton(xfrm, dialect.drawing, "chOff", false)
	if err != nil {
		return nativeGroupTransform{}, err
	}
	childExt, err := nativeSingleton(xfrm, dialect.drawing, "chExt", false)
	if err != nil {
		return nativeGroupTransform{}, err
	}
	if off == nil || ext == nil || childOff == nil || childExt == nil {
		return nativeGroupTransform{}, refuseNativeGroup("pptx.group-transform-unavailable", "group transform omits explicit coordinate-space metadata and is preserved rather than assigned inferred defaults")
	}
	if len(xfrm.Children) != 4 ||
		xfrm.Children[0] != off ||
		xfrm.Children[1] != ext ||
		xfrm.Children[2] != childOff ||
		xfrm.Children[3] != childExt {
		return nativeGroupTransform{}, fmt.Errorf("pptxpatch: native extract: malformed group transform child order")
	}
	x, y, err := nativeGroupPoint(off, false)
	if err != nil {
		return nativeGroupTransform{}, err
	}
	cx, cy, err := nativeGroupPoint(ext, true)
	if err != nil {
		return nativeGroupTransform{}, err
	}
	childX, childY, err := nativeGroupPoint(childOff, false)
	if err != nil {
		return nativeGroupTransform{}, err
	}
	childCx, childCy, err := nativeGroupPoint(childExt, true)
	if err != nil {
		return nativeGroupTransform{}, err
	}
	result := nativeGroupTransform{x: x, y: y, cx: cx, cy: cy, childX: childX, childY: childY, childCx: childCx, childCy: childCy}
	if _, _, _, _, err := nativeGroupAffineComponents(
		NativeTransform{X: &result.x, Y: &result.y, Cx: &result.cx, Cy: &result.cy},
		NativeTransform{X: &result.childX, Y: &result.childY, Cx: &result.childCx, Cy: &result.childCy},
	); err != nil {
		return nativeGroupTransform{}, refuseNativeGroup("pptx.group-fractional-transform-unavailable", "group transform cannot be represented by the exact renderer-neutral affine contract")
	}
	return result, nil
}

func nativeGroupPoint(node *nativeXMLNode, positive bool) (int64, int64, error) {
	first, second := "x", "y"
	if positive {
		first, second = "cx", "cy"
	}
	if err := requireOnlyNativeAttrs(node, xml.Name{Local: first}, xml.Name{Local: second}); err != nil || requireOnlyNativeChildren(node) != nil {
		return 0, 0, fmt.Errorf("pptxpatch: native extract: malformed group coordinate %s/%s", first, second)
	}
	left, err := requiredCanonicalNativeShapeInt(node, first, -nativeMaxSafeInteger, nativeMaxSafeInteger)
	if positive {
		left, err = requiredCanonicalNativeShapeInt(node, first, 1, nativeMaxSafeInteger)
	}
	if err != nil {
		return 0, 0, err
	}
	right, err := requiredCanonicalNativeShapeInt(node, second, -nativeMaxSafeInteger, nativeMaxSafeInteger)
	if positive {
		right, err = requiredCanonicalNativeShapeInt(node, second, 1, nativeMaxSafeInteger)
	}
	if err != nil {
		return 0, 0, err
	}
	return left, right, nil
}

func rejectNativeGroupDialectMix(node *nativeXMLNode, dialect nativeExtractDialect) error {
	oppositePresentation, oppositeDrawing, oppositeRelationships := nsPresentationStrict, nsDrawingStrict, nsOfficeRelsStrict
	if dialect.presentation == nsPresentationStrict {
		oppositePresentation, oppositeDrawing, oppositeRelationships = nsPresentationTransitional, nsDrawingTransitional, nsOfficeRelsTransitional
	}
	var walk func(*nativeXMLNode) error
	walk = func(current *nativeXMLNode) error {
		if current.Name.Space == oppositePresentation || current.Name.Space == oppositeDrawing || current.Name.Space == oppositeRelationships {
			return fmt.Errorf("pptxpatch: native extract: group mixes Strict and Transitional namespaces")
		}
		for _, attr := range current.Attrs {
			if attr.Name.Space == oppositePresentation || attr.Name.Space == oppositeDrawing || attr.Name.Space == oppositeRelationships {
				return fmt.Errorf("pptxpatch: native extract: group mixes Strict and Transitional namespaces")
			}
		}
		for _, child := range current.Children {
			if err := walk(child); err != nil {
				return err
			}
		}
		return nil
	}
	return walk(node)
}

const nativeGroupTransformPPM = int64(1_000_000)

// nativeGroupAffineComponents retains the complete DrawingML group transform
// as an integer-PPM affine operation. This preserves scaling of strokes, text,
// images, and descendants instead of flattening only child bounding boxes.
func nativeGroupAffineComponents(transform, childTransform NativeTransform) (scaleX, scaleY, translateX, translateY int64, err error) {
	if transform.X == nil || transform.Y == nil || transform.Cx == nil || transform.Cy == nil ||
		childTransform.X == nil || childTransform.Y == nil || childTransform.Cx == nil || childTransform.Cy == nil ||
		*transform.Cx <= 0 || *transform.Cy <= 0 || *childTransform.Cx <= 0 || *childTransform.Cy <= 0 {
		return 0, 0, 0, 0, fmt.Errorf("incomplete group affine transform")
	}
	scale := func(extent, childExtent int64) (int64, error) {
		numerator := new(big.Int).Mul(big.NewInt(extent), big.NewInt(nativeGroupTransformPPM))
		quotient, remainder := new(big.Int), new(big.Int)
		quotient.QuoRem(numerator, big.NewInt(childExtent), remainder)
		if remainder.Sign() != 0 || !quotient.IsInt64() || quotient.Sign() <= 0 || quotient.Int64() > nativeMaxSafeInteger {
			return 0, fmt.Errorf("group scale is not an exact safe integer PPM value")
		}
		return quotient.Int64(), nil
	}
	translation := func(offset, childOffset, ppm int64) (int64, error) {
		product := new(big.Int).Mul(big.NewInt(childOffset), big.NewInt(ppm))
		quotient, remainder := new(big.Int), new(big.Int)
		quotient.QuoRem(product, big.NewInt(nativeGroupTransformPPM), remainder)
		if remainder.Sign() != 0 {
			return 0, fmt.Errorf("group translation is not an exact integer EMU value")
		}
		result := new(big.Int).Sub(big.NewInt(offset), quotient)
		if !result.IsInt64() || result.Int64() < -nativeMaxSafeInteger || result.Int64() > nativeMaxSafeInteger {
			return 0, fmt.Errorf("group translation exceeds the safe integer range")
		}
		return result.Int64(), nil
	}
	if scaleX, err = scale(*transform.Cx, *childTransform.Cx); err != nil {
		return 0, 0, 0, 0, err
	}
	if scaleY, err = scale(*transform.Cy, *childTransform.Cy); err != nil {
		return 0, 0, 0, 0, err
	}
	if translateX, err = translation(*transform.X, *childTransform.X, scaleX); err != nil {
		return 0, 0, 0, 0, err
	}
	if translateY, err = translation(*transform.Y, *childTransform.Y, scaleY); err != nil {
		return 0, 0, 0, 0, err
	}
	return scaleX, scaleY, translateX, translateY, nil
}
