package pptxpatch

import (
	"encoding/xml"
	"errors"
	"fmt"
)

// extractNativeChartGraphicFrame projects an opaque chart frame. Semantic chart
// series are never interpreted. When the chart part has exactly one PNG/JPEG
// image relationship, that picture is attached as previewAssetId and painted
// at the frame's exact EMU. Missing or ambiguous previews stay preserve-only
// without inventing a chart renderer.
func (extractor *nativeExtractor) extractNativeChartGraphicFrame(node *nativeXMLNode, slidePart, slideID string, relationships []nativeExtractRelationship, dialect nativeExtractDialect) (NativeElement, error) {
	if node == nil || node.Name != (xml.Name{Space: dialect.presentation, Local: "graphicFrame"}) {
		return NativeElement{}, fmt.Errorf("pptxpatch: native extract: invalid chart graphic frame root")
	}
	if err := requireOnlyNativeAttrs(node); err != nil {
		return NativeElement{}, refuseNativeGraphicFrame("pptx.chart-markup-unavailable", "chart graphic frame root attributes are outside the exact native subset")
	}
	if err := requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.presentation, Local: "nvGraphicFramePr"},
		xml.Name{Space: dialect.presentation, Local: "xfrm"},
		xml.Name{Space: dialect.drawing, Local: "graphic"}); err != nil {
		return NativeElement{}, refuseNativeGraphicFrame("pptx.chart-markup-unavailable", "chart graphic frame contains unmodeled markup or direct text")
	}
	nonVisual, err := nativeSingleton(node, dialect.presentation, "nvGraphicFramePr", true)
	if err != nil {
		return NativeElement{}, err
	}
	transformNode, err := nativeSingleton(node, dialect.presentation, "xfrm", true)
	if err != nil {
		return NativeElement{}, err
	}
	graphic, err := nativeSingleton(node, dialect.drawing, "graphic", true)
	if err != nil {
		return NativeElement{}, err
	}
	if len(node.Children) != 3 || node.Children[0] != nonVisual || node.Children[1] != transformNode || node.Children[2] != graphic {
		return NativeElement{}, fmt.Errorf("pptxpatch: native extract: malformed chart graphic frame child order")
	}

	objectID, name, err := validateNativeChartNonVisual(nonVisual, dialect)
	if err != nil {
		return NativeElement{}, err
	}
	transform, err := validateNativeChartTransform(transformNode, dialect)
	if err != nil {
		return NativeElement{}, err
	}
	relationshipID, err := extractNativeChartRelationshipID(graphic, dialect)
	if err != nil {
		return NativeElement{}, err
	}
	chartRel, err := exactNativeChartRelationship(relationships, relationshipID, dialect)
	if err != nil {
		return NativeElement{}, err
	}
	if contentType := extractor.pkg.contentTypes.forPart(chartRel.Part); contentType != contentTypeChart {
		return NativeElement{}, fmt.Errorf("pptxpatch: native extract OPC: chart part %q has content type %q", chartRel.Part, contentType)
	}
	chartPayload, ok := extractor.pkg.parts[chartRel.Part]
	if !ok || len(chartPayload) == 0 {
		return NativeElement{}, fmt.Errorf("pptxpatch: native extract OPC: missing chart part %q", chartRel.Part)
	}
	if err := validateNativeChartSpace(chartPayload, chartRel.Part, dialect); err != nil {
		return NativeElement{}, refuseNativeGraphicFrame("pptx.chart-markup-unavailable", err.Error())
	}

	chartFingerprint := nativeSHA256(chartPayload)
	opaqueRef, err := extractor.issuePassthrough(chartRel.Part, "chart-space", chartFingerprint, chartPayload, "pptx.chart-opaque")
	if err != nil {
		return NativeElement{}, err
	}
	if err := extractor.reserveNativePassthroughReference(); err != nil {
		return NativeElement{}, err
	}

	previewAssetID, previewErr := extractor.extractNativeChartPreviewAsset(chartRel.Part, dialect)
	if previewErr != nil {
		return NativeElement{}, previewErr
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
	chart := NativeOpaqueChart{
		ChartPart: chartRel.Part, RelationshipID: relationshipID, OpaqueRef: opaqueRef, PreviewAssetID: previewAssetID,
	}
	element := NativeElement{
		Kind: NativeElementKindChart, ID: elementID, Provenance: NativeProvenanceParsed,
		Transform: transform, Chart: &chart, Passthrough: []NativePassthroughRef{}, Children: nil,
		Source:        &NativeSourceAnchor{PartName: slidePart, ObjectID: objectID, RelationshipID: &relID, FingerprintSHA256: fingerprint},
		Compatibility: NativeCompatibility{Status: NativeCompatibilityStatusPreserveOnly, Diagnostics: []NativeDiagnostic{}},
	}
	if name != "" {
		element.Name = stringPointer(name)
	}
	partName := slidePart
	element.Compatibility.Diagnostics = append(element.Compatibility.Diagnostics, NativeDiagnostic{
		Severity: NativeDiagnosticSeverityWarning, Code: "chart.opaque",
		Message: "Chart is preserved as an opaque relationship graph.",
		Scope:   &NativeDiagnosticScope{SlideID: &slideID, ElementID: &elementID, PartName: &partName},
	})
	return element, nil
}

func validateNativeChartNonVisual(node *nativeXMLNode, dialect nativeExtractDialect) (string, string, error) {
	if err := requireOnlyNativeAttrs(node); err != nil {
		return "", "", refuseNativeGraphicFrame("pptx.chart-nonvisual-unavailable", "chart nonvisual properties contain unmodeled attributes")
	}
	if err := requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.presentation, Local: "cNvPr"},
		xml.Name{Space: dialect.presentation, Local: "cNvGraphicFramePr"},
		xml.Name{Space: dialect.presentation, Local: "nvPr"}); err != nil {
		return "", "", refuseNativeGraphicFrame("pptx.chart-nonvisual-unavailable", "chart nonvisual properties are outside the exact subset")
	}
	cNvPr, err := nativeSingleton(node, dialect.presentation, "cNvPr", true)
	if err != nil {
		return "", "", err
	}
	cNvGraphicFramePr, err := nativeSingleton(node, dialect.presentation, "cNvGraphicFramePr", true)
	if err != nil {
		return "", "", err
	}
	nvPr, err := nativeSingleton(node, dialect.presentation, "nvPr", true)
	if err != nil {
		return "", "", err
	}
	if requireOnlyNativeAttrs(cNvPr, xml.Name{Local: "id"}, xml.Name{Local: "name"}) != nil || len(cNvPr.Children) != 0 || !onlyNativeXMLSpace(cNvPr.Text) {
		return "", "", refuseNativeGraphicFrame("pptx.chart-nonvisual-unavailable", "chart visibility, hyperlink, or extension metadata is not modeled")
	}
	nativeID, err := canonicalNativeUnsignedID(cNvPr, "", "id", 1)
	if err != nil {
		return "", "", err
	}
	name, _ := exactNativeAttr(cNvPr, "", "name")
	if utf16CodeUnitLengthBounded(name, 1025) > 1024 {
		return "", "", fmt.Errorf("pptxpatch: native extract: chart name exceeds contract bound")
	}
	if requireEmptyNativeElement(cNvGraphicFramePr) != nil || requireEmptyNativeElement(nvPr) != nil {
		return "", "", refuseNativeGraphicFrame("pptx.chart-inheritance-unavailable", "chart locks, placeholder, or inherited nonvisual properties are not resolved")
	}
	return "cNvPr-" + nativeID, name, nil
}

func validateNativeChartTransform(node *nativeXMLNode, dialect nativeExtractDialect) (NativeTransform, error) {
	transform, err := validateNativeTableTransform(node, dialect)
	if err != nil {
		var refusal nativeGraphicFrameProjectionRefusal
		if errors.As(err, &refusal) {
			return NativeTransform{}, refuseNativeGraphicFrame("pptx.chart-transform-unavailable", "chart transform is rotated, flipped, or outside the exact native subset")
		}
		return NativeTransform{}, err
	}
	return transform, nil
}

func extractNativeChartRelationshipID(graphic *nativeXMLNode, dialect nativeExtractDialect) (string, error) {
	if requireOnlyNativeAttrs(graphic) != nil || requireOnlyNativeChildren(graphic, xml.Name{Space: dialect.drawing, Local: "graphicData"}) != nil {
		return "", refuseNativeGraphicFrame("pptx.chart-markup-unavailable", "chart graphic container is outside the exact native subset")
	}
	data, err := nativeSingleton(graphic, dialect.drawing, "graphicData", true)
	if err != nil {
		return "", err
	}
	if requireOnlyNativeAttrs(data, xml.Name{Local: "uri"}) != nil {
		return "", refuseNativeGraphicFrame("pptx.chart-markup-unavailable", "chart graphic data has unmodeled type metadata")
	}
	uri, ok := exactNativeAttr(data, "", "uri")
	if !ok || uri != dialect.chart {
		return "", refuseNativeGraphicFrame("pptx.chart-markup-unavailable", "chart graphic data is not the exact DrawingML chart URI")
	}
	if requireOnlyNativeChildren(data, xml.Name{Space: dialect.chart, Local: "chart"}) != nil {
		return "", refuseNativeGraphicFrame("pptx.chart-markup-unavailable", "chart graphic data contains unknown markup")
	}
	chart, err := nativeSingleton(data, dialect.chart, "chart", true)
	if err != nil {
		return "", err
	}
	if requireOnlyNativeAttrs(chart, xml.Name{Space: dialect.rels, Local: "id"}) != nil || requireOnlyNativeChildren(chart) != nil {
		return "", refuseNativeGraphicFrame("pptx.chart-markup-unavailable", "chart reference markup is outside the exact native subset")
	}
	id, ok := exactNativeAttr(chart, dialect.rels, "id")
	if !ok || id == "" || !nativeIDPattern.MatchString(id) {
		return "", fmt.Errorf("pptxpatch: native extract: chart requires one valid chart relationship")
	}
	return id, nil
}

func exactNativeChartRelationship(relationships []nativeExtractRelationship, relationshipID string, dialect nativeExtractDialect) (nativeExtractRelationship, error) {
	var selected *nativeExtractRelationship
	for index := range relationships {
		if relationships[index].ID != relationshipID {
			continue
		}
		if selected != nil {
			return nativeExtractRelationship{}, fmt.Errorf("pptxpatch: native extract OPC: ambiguous duplicate chart relationship id")
		}
		selected = &relationships[index]
	}
	if selected == nil || selected.Type != dialect.relChart || !selected.internal() || selected.Part == "" {
		return nativeExtractRelationship{}, fmt.Errorf("pptxpatch: native extract OPC: chart relationship is missing, external, or not the exact chart type")
	}
	return *selected, nil
}

func validateNativeChartSpace(payload []byte, part string, dialect nativeExtractDialect) error {
	root, err := parseNativeXML(payload, part)
	if err != nil {
		return fmt.Errorf("chart part is not well-formed XML")
	}
	if root.Name != (xml.Name{Space: dialect.chart, Local: "chartSpace"}) {
		return fmt.Errorf("chart part root is not chartSpace")
	}
	return nil
}

func (extractor *nativeExtractor) extractNativeChartPreviewAsset(chartPart string, dialect nativeExtractDialect) (*string, error) {
	relationships, hasRels, err := extractor.optionalRelationships(chartPart)
	if err != nil {
		return nil, err
	}
	if !hasRels {
		return nil, nil
	}
	var images []nativeExtractRelationship
	for index := range relationships {
		rel := relationships[index]
		if rel.Type == dialect.relImage && rel.internal() && rel.Part != "" {
			images = append(images, rel)
		}
	}
	if len(images) != 1 {
		return nil, nil
	}
	contentType, supported, err := nativePictureContentType(extractor.pkg.contentTypes.forPart(images[0].Part))
	if err != nil || !supported {
		return nil, nil
	}
	assetID, err := extractor.nativePictureAsset(images[0].Part, contentType)
	if err != nil {
		return nil, err
	}
	return &assetID, nil
}
