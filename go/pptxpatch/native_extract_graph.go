package pptxpatch

import (
	"encoding/xml"
	"fmt"
	"strings"
)

type nativeSlideDependencyGraph struct {
	relationships []nativeExtractRelationship
	layoutPart    string
	layoutRoot    *nativeXMLNode
	layoutPayload []byte
	masterPart    string
	masterRoot    *nativeXMLNode
	masterPayload []byte
	themePart     string
	themeRoot     *nativeXMLNode
	themePayload  []byte
	unsupported   []nativeUnsupportedSource
}

func (extractor *nativeExtractor) resolveSlideDependencyGraph(slidePart string, dialect nativeExtractDialect) (nativeSlideDependencyGraph, error) {
	graph := nativeSlideDependencyGraph{}
	slideRelationships, err := extractor.parseRelationships(slidePart)
	if err != nil {
		return graph, err
	}
	if err := extractor.validateRelationshipSet(slideRelationships, dialect, slidePart); err != nil {
		return graph, err
	}
	graph.relationships = slideRelationships
	if err := extractor.appendRelationshipClosure(&graph, slidePart, slideRelationships, map[string]bool{dialect.relSlideLayout: true, dialect.relImage: true}, "slide"); err != nil {
		return graph, err
	}
	layoutRelationship, err := uniqueNativeInternalRelationship(slideRelationships, dialect.relSlideLayout, "slide layout")
	if err != nil {
		return graph, err
	}
	if !asciiEqualFoldNative(extractor.pkg.contentTypes.forPart(layoutRelationship.Part), contentTypeSlideLayout) {
		return graph, fmt.Errorf("pptxpatch: native extract OPC: slide layout has wrong effective content type")
	}
	graph.layoutPart = layoutRelationship.Part
	graph.layoutPayload = extractor.pkg.parts[graph.layoutPart]
	graph.layoutRoot, err = parseNativeXML(graph.layoutPayload, graph.layoutPart)
	if err != nil {
		return graph, err
	}
	if graph.layoutRoot.Name != (xml.Name{Space: dialect.presentation, Local: "sldLayout"}) {
		return graph, fmt.Errorf("pptxpatch: native extract: slide layout has wrong namespace/root")
	}

	layoutRelationships, err := extractor.parseRelationships(graph.layoutPart)
	if err != nil {
		return graph, err
	}
	if err := extractor.validateRelationshipSet(layoutRelationships, dialect, graph.layoutPart); err != nil {
		return graph, err
	}
	if err := extractor.appendRelationshipClosure(&graph, graph.layoutPart, layoutRelationships, map[string]bool{dialect.relSlideMaster: true}, "layout"); err != nil {
		return graph, err
	}
	masterRelationship, err := uniqueNativeInternalRelationship(layoutRelationships, dialect.relSlideMaster, "slide master")
	if err != nil {
		return graph, err
	}
	if !asciiEqualFoldNative(extractor.pkg.contentTypes.forPart(masterRelationship.Part), contentTypeSlideMaster) {
		return graph, fmt.Errorf("pptxpatch: native extract OPC: slide master has wrong effective content type")
	}
	graph.masterPart = masterRelationship.Part
	graph.masterPayload = extractor.pkg.parts[graph.masterPart]
	graph.masterRoot, err = parseNativeXML(graph.masterPayload, graph.masterPart)
	if err != nil {
		return graph, err
	}
	if graph.masterRoot.Name != (xml.Name{Space: dialect.presentation, Local: "sldMaster"}) {
		return graph, fmt.Errorf("pptxpatch: native extract: slide master has wrong namespace/root")
	}

	masterRelationships, err := extractor.parseRelationships(graph.masterPart)
	if err != nil {
		return graph, err
	}
	if err := extractor.validateRelationshipSet(masterRelationships, dialect, graph.masterPart); err != nil {
		return graph, err
	}
	if err := extractor.appendRelationshipClosure(&graph, graph.masterPart, masterRelationships, map[string]bool{dialect.relTheme: true}, "master"); err != nil {
		return graph, err
	}
	themeRelationship, err := uniqueNativeInternalRelationship(masterRelationships, dialect.relTheme, "theme")
	if err != nil {
		return graph, err
	}
	if !asciiEqualFoldNative(extractor.pkg.contentTypes.forPart(themeRelationship.Part), contentTypeTheme) {
		return graph, fmt.Errorf("pptxpatch: native extract OPC: theme has wrong effective content type")
	}
	graph.themePart = themeRelationship.Part
	graph.themePayload = extractor.pkg.parts[graph.themePart]
	graph.themeRoot, err = parseNativeXML(graph.themePayload, graph.themePart)
	if err != nil {
		return graph, err
	}
	if graph.themeRoot.Name != (xml.Name{Space: dialect.drawing, Local: "theme"}) {
		return graph, fmt.Errorf("pptxpatch: native extract: theme has wrong namespace/root")
	}
	themeRelationships, hasThemeRelationships, err := extractor.optionalRelationships(graph.themePart)
	if err != nil {
		return graph, err
	}
	if hasThemeRelationships {
		if err := extractor.validateRelationshipSet(themeRelationships, dialect, graph.themePart); err != nil {
			return graph, err
		}
		if err := extractor.appendRelationshipClosure(&graph, graph.themePart, themeRelationships, map[string]bool{}, "theme"); err != nil {
			return graph, err
		}
	}

	for _, source := range []struct {
		part    string
		payload []byte
		root    *nativeXMLNode
		kind    string
	}{
		{part: graph.layoutPart, payload: graph.layoutPayload, root: graph.layoutRoot, kind: "layout"},
		{part: graph.masterPart, payload: graph.masterPayload, root: graph.masterRoot, kind: "master"},
		{part: graph.themePart, payload: graph.themePayload, root: graph.themeRoot, kind: "theme"},
	} {
		unsupported, collectErr := collectNativeDependencyUnsupported(source.part, source.payload, source.root, source.kind, dialect)
		if collectErr != nil {
			return graph, collectErr
		}
		if len(graph.unsupported) > nativeExtractMaxPassthrough-len(unsupported) {
			return graph, fmt.Errorf("pptxpatch: native extract: dependency passthrough budget exceeded")
		}
		graph.unsupported = append(graph.unsupported, unsupported...)
	}
	return graph, nil
}

func (extractor *nativeExtractor) appendRelationshipClosure(graph *nativeSlideDependencyGraph, ownerPart string, relationships []nativeExtractRelationship, coreTypes map[string]bool, kind string) error {
	if len(graph.unsupported) >= nativeExtractMaxPassthrough {
		return fmt.Errorf("pptxpatch: native extract: relationship closure budget exceeded")
	}
	relsPart, err := extractor.actualRelationshipsPart(ownerPart)
	if err != nil {
		return err
	}
	relsPayload := extractor.pkg.parts[relsPart]
	graph.unsupported = append(graph.unsupported, makeNativeUnsupportedPayload(relsPayload, relsPart, kind+"-relationships", "pptx.relationship-map-preserve", kind+" relationship map is capability-bound to the source revision"))
	for _, relationship := range relationships {
		if coreTypes[relationship.Type] || !relationship.internal() {
			continue
		}
		if len(graph.unsupported) >= nativeExtractMaxPassthrough {
			return fmt.Errorf("pptxpatch: native extract: relationship closure budget exceeded")
		}
		payload := extractor.pkg.parts[relationship.Part]
		objectID := kind + "-rel-" + nativeSHA256([]byte(relationship.ID))[:24]
		graph.unsupported = append(graph.unsupported, makeNativeUnsupportedPayload(payload, relationship.Part, objectID, "pptx.unsupported-"+kind+"-relationship", kind+" relationship target is preserved as an opaque closure"))
	}
	return nil
}

func (extractor *nativeExtractor) validateRelationshipSet(relationships []nativeExtractRelationship, dialect nativeExtractDialect, ownerPart string) error {
	oppositeOfficeNamespace := nsOfficeRelsStrict
	if dialect.rels == nsOfficeRelsStrict {
		oppositeOfficeNamespace = nsOfficeRelsTransitional
	}
	for _, relationship := range relationships {
		if relationship.Namespace != dialect.packageRels {
			return fmt.Errorf("pptxpatch: native extract OPC: relationship dialect mismatch for %q", ownerPart)
		}
		if strings.HasPrefix(relationship.Type, oppositeOfficeNamespace+"/") {
			return fmt.Errorf("pptxpatch: native extract OPC: opposing Transitional/Strict relationship type in %q", ownerPart)
		}
		if relationship.internal() && strings.TrimSpace(extractor.pkg.contentTypes.forPart(relationship.Part)) == "" {
			return fmt.Errorf("pptxpatch: native extract OPC: relationship %s in %q targets a part with no effective content type", relationship.ID, ownerPart)
		}
	}
	return nil
}

func uniqueNativeInternalRelationship(relationships []nativeExtractRelationship, relationshipType, label string) (nativeExtractRelationship, error) {
	var result *nativeExtractRelationship
	for index := range relationships {
		relationship := &relationships[index]
		if relationship.Type != relationshipType {
			continue
		}
		if result != nil || !relationship.internal() {
			return nativeExtractRelationship{}, fmt.Errorf("pptxpatch: native extract OPC: %s relationship must be unique and internal", label)
		}
		result = relationship
	}
	if result == nil {
		return nativeExtractRelationship{}, fmt.Errorf("pptxpatch: native extract OPC: missing exact %s relationship", label)
	}
	return *result, nil
}

func collectNativeDependencyUnsupported(part string, payload []byte, root *nativeXMLNode, kind string, dialect nativeExtractDialect) ([]nativeUnsupportedSource, error) {
	result := []nativeUnsupportedSource{}
	appendNode := func(node *nativeXMLNode, suffix, code, message string) error {
		if len(result) >= nativeExtractMaxPassthrough {
			return fmt.Errorf("pptxpatch: native extract: dependency passthrough budget exceeded")
		}
		unsupported, err := makeNativeUnsupportedSource(payload, node, part, kind+"-"+suffix, code, message)
		if err != nil {
			return err
		}
		result = append(result, unsupported)
		return nil
	}
	if hasNativeSemanticAttrs(root) || !onlyNativeXMLSpace(root.Text) {
		if err := appendNode(root, "root", "pptx.unsupported-"+kind+"-markup", kind+" root contains unmodeled markup"); err != nil {
			return nil, err
		}
	}
	if kind == "theme" {
		for index, child := range root.Children {
			if err := appendNode(child, fmt.Sprintf("child-%d-%d", index, child.RawStart), "pptx.unsupported-theme-dependency", "theme semantics are relationship-routed but not resolved by native PPTX v1"); err != nil {
				return nil, err
			}
		}
		return result, nil
	}
	cSld, err := nativeSingleton(root, dialect.presentation, "cSld", true)
	if err != nil {
		return nil, err
	}
	for index, child := range root.Children {
		if child != cSld {
			if err := appendNode(child, fmt.Sprintf("child-%d-%d", index, child.RawStart), "pptx.unsupported-"+kind+"-dependency", kind+" contains unmodeled inheritance metadata"); err != nil {
				return nil, err
			}
		}
	}
	if hasNativeSemanticAttrs(cSld) || !onlyNativeXMLSpace(cSld.Text) {
		if err := appendNode(cSld, "common", "pptx.unsupported-"+kind+"-common-data", kind+" common slide data contains unmodeled metadata"); err != nil {
			return nil, err
		}
	}
	spTree, err := nativeSingleton(cSld, dialect.presentation, "spTree", true)
	if err != nil {
		return nil, err
	}
	objectIDs, err := collectNativeShapeTreeObjectIDs(spTree, dialect)
	if err != nil {
		return nil, err
	}
	rootID, err := validateNativeRootGroupScaffold(spTree, dialect)
	if err != nil {
		return nil, err
	}
	for index, child := range cSld.Children {
		if child != spTree {
			if err := appendNode(child, fmt.Sprintf("common-child-%d-%d", index, child.RawStart), "pptx.unsupported-"+kind+"-content", kind+" background or common content is not resolved by native PPTX v1"); err != nil {
				return nil, err
			}
		}
	}
	groupProperties, err := nativeSingleton(spTree, dialect.presentation, "grpSpPr", true)
	if err != nil {
		return nil, err
	}
	if hasNativeSemanticAttrs(groupProperties) || len(groupProperties.Children) != 0 || !onlyNativeXMLSpace(groupProperties.Text) {
		if err := appendNode(groupProperties, rootID+"-transform", "pptx.unsupported-"+kind+"-group-transform", kind+" root group transform is not representable in native PPTX v1"); err != nil {
			return nil, err
		}
	}
	for index, child := range spTree.Children {
		if child.Name == (xml.Name{Space: dialect.presentation, Local: "nvGrpSpPr"}) || child.Name == (xml.Name{Space: dialect.presentation, Local: "grpSpPr"}) {
			continue
		}
		objectID := objectIDs[child]
		if objectID == "" {
			objectID = fmt.Sprintf("shape-%d-%d", index, child.RawStart)
		}
		if err := appendNode(child, objectID, "pptx.unsupported-"+kind+"-shape", kind+" contributes inherited slide content not modeled by native PPTX v1"); err != nil {
			return nil, err
		}
	}
	return result, nil
}
