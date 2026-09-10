package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"fmt"
	"sort"
	"strings"
)

// ChartIdentity names one native chart without relying on its display name,
// position, or a recyclable relationship id. The chart part plus the owning
// drawing's workbook-issued cNvPr id stays stable across supported updates.
type ChartIdentity struct {
	Part        string `json:"part"`
	DrawingPart string `json:"drawingPart"`
	ObjectID    uint32 `json:"objectId"`
}

// ReadChartIdentities returns identities for unambiguous top-level chart
// anchors. Orphan chart parts remain visible through ReadCharts but cannot be
// edited until a host repairs their ownership graph. Lifecycle mutations are
// intentionally narrower and accept only two-cell anchors.
func ReadChartIdentities(data []byte) ([]ChartIdentity, error) {
	parts, err := readPivotArchive(data)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: read chart identities: %w", err)
	}
	var drawings []string
	for part := range parts {
		if strings.HasPrefix(part, "xl/drawings/") && strings.HasSuffix(part, ".xml") && !strings.Contains(part, "/_rels/") {
			drawings = append(drawings, part)
		}
	}
	sort.Strings(drawings)
	var identities []ChartIdentity
	seenPart := map[string]ChartIdentity{}
	for _, drawingPart := range drawings {
		index, err := indexShapeDrawing(parts[drawingPart])
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: read chart identities: drawing %q: %w", drawingPart, err)
		}
		relsPart := relsPartFor(drawingPart)
		relsData, hasRels := parts[relsPart]
		for _, anchor := range index.anchors {
			if anchor.objectType != "graphicFrame" || anchor.chartRelID == "" {
				continue
			}
			if anchor.directObjects != 1 || anchor.objectID == 0 {
				return nil, fmt.Errorf("xlsxpatch: read chart identities: chart object %d has %d top-level drawing objects", anchor.objectID, anchor.directObjects)
			}
			if !hasRels {
				return nil, fmt.Errorf("xlsxpatch: read chart identities: drawing relationships %q are missing", relsPart)
			}
			chartPart, err := exactChartRelationshipTarget(relsData, drawingPart, anchor.chartRelID)
			if err != nil {
				return nil, fmt.Errorf("xlsxpatch: read chart identities: %w", err)
			}
			if _, ok := parts[chartPart]; !ok {
				return nil, fmt.Errorf("xlsxpatch: read chart identities: chart part %q is missing", chartPart)
			}
			identity := ChartIdentity{Part: chartPart, DrawingPart: drawingPart, ObjectID: anchor.objectID}
			if previous, duplicate := seenPart[chartPart]; duplicate {
				return nil, fmt.Errorf("xlsxpatch: read chart identities: chart part %q has multiple drawing owners (%+v and %+v)", chartPart, previous, identity)
			}
			seenPart[chartPart] = identity
			identities = append(identities, identity)
		}
	}
	return identities, nil
}

// RemoveChart removes exactly one hydrated chart anchor, its drawing
// relationship, chart part, and content-type override. The drawing itself and
// every sibling anchor/relationship remain byte-for-byte intact.
func RemoveChart(orig []byte, identity ChartIdentity) ([]byte, error) {
	graph, err := resolveChartLifecycleGraph(orig, identity)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: remove chart: %w", err)
	}
	updatedDrawing := removeElementSpans(graph.parts[identity.DrawingPart], []xmlSpan{graph.target.span})
	updatedRels, err := relationshipDocumentWithout(graph.parts[graph.drawingRelsPart], graph.target.chartRelID, relTypeChart, identity.DrawingPart, identity.Part)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: remove chart: %w", err)
	}
	updatedContentTypes, err := contentTypesWithoutParts(graph.parts["[Content_Types].xml"], []string{identity.Part})
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: remove chart: %w", err)
	}
	return Apply(orig, Patch{
		Replace: map[string][]byte{
			identity.DrawingPart:  updatedDrawing,
			graph.drawingRelsPart: updatedRels,
			"[Content_Types].xml": updatedContentTypes,
		},
		Delete: map[string]bool{identity.Part: true},
	})
}

// UpdateChart surgically replaces one supported native chart while retaining
// its chart part, drawing part, cNvPr object id, and relationship id.
func UpdateChart(orig []byte, identity ChartIdentity, spec ChartWriteSpec) ([]byte, error) {
	if err := validateChartWriteSpec(spec); err != nil {
		return nil, fmt.Errorf("xlsxpatch: update chart: %w", err)
	}
	graph, err := resolveChartLifecycleGraph(orig, identity)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: update chart: %w", err)
	}
	if spec.SheetName != graph.sheetName {
		return nil, fmt.Errorf("xlsxpatch: update chart: identity belongs to worksheet %q, not %q", graph.sheetName, spec.SheetName)
	}
	spec = chartSpecWithCaches(func(name string) (string, bool) {
		data, ok := graph.parts[name]
		return string(data), ok
	}, spec)
	replacement := []byte(anchorXMLWithObjectID(spec, graph.target.chartRelID, identity.ObjectID))
	drawing := graph.parts[identity.DrawingPart]
	updatedDrawing := make([]byte, 0, len(drawing)-graph.target.span.end+graph.target.span.start+len(replacement))
	updatedDrawing = append(updatedDrawing, drawing[:graph.target.span.start]...)
	updatedDrawing = append(updatedDrawing, replacement...)
	updatedDrawing = append(updatedDrawing, drawing[graph.target.span.end:]...)
	out, err := Apply(orig, Patch{Replace: map[string][]byte{
		identity.DrawingPart: updatedDrawing,
		identity.Part:        []byte(chartSpaceXML(spec)),
	}})
	if err != nil {
		return nil, err
	}
	identities, err := ReadChartIdentities(out)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: update chart: verify identity: %w", err)
	}
	for _, current := range identities {
		if current == identity {
			return out, nil
		}
	}
	return nil, fmt.Errorf("xlsxpatch: update chart: replacement did not retain identity %+v", identity)
}

type chartLifecycleGraph struct {
	parts           map[string][]byte
	target          shapeAnchorNode
	drawingRelsPart string
	sheetName       string
}

func resolveChartLifecycleGraph(data []byte, identity ChartIdentity) (chartLifecycleGraph, error) {
	graph := chartLifecycleGraph{}
	if identity.Part == "" || !strings.HasPrefix(identity.Part, "xl/charts/") || !strings.HasSuffix(identity.Part, ".xml") || strings.Contains(identity.Part, "\\") {
		return graph, fmt.Errorf("invalid chart identity %+v", identity)
	}
	if identity.DrawingPart == "" || !strings.HasPrefix(identity.DrawingPart, "xl/drawings/") || !strings.HasSuffix(identity.DrawingPart, ".xml") || strings.Contains(identity.DrawingPart, "\\") || strings.Contains(identity.DrawingPart, "/_rels/") || identity.ObjectID == 0 {
		return graph, fmt.Errorf("invalid chart identity %+v", identity)
	}
	for _, part := range []string{identity.Part, identity.DrawingPart} {
		if _, err := canonicalNativeContractPartKey(part); err != nil {
			return graph, fmt.Errorf("invalid chart identity %+v: %w", identity, err)
		}
	}
	parts, err := readPivotArchive(data)
	if err != nil {
		return graph, err
	}
	graph.parts = parts
	for _, required := range []string{identity.Part, identity.DrawingPart, "[Content_Types].xml"} {
		if _, ok := parts[required]; !ok {
			return graph, fmt.Errorf("stale or missing chart part %q", required)
		}
	}
	graph.drawingRelsPart = relsPartFor(identity.DrawingPart)
	if _, ok := parts[graph.drawingRelsPart]; !ok {
		return graph, fmt.Errorf("drawing relationships %q are missing", graph.drawingRelsPart)
	}
	if _, hasChartRels := parts[relsPartFor(identity.Part)]; hasChartRels {
		return graph, fmt.Errorf("chart %q has an unsupported dependent relationship graph", identity.Part)
	}
	index, err := indexShapeDrawing(parts[identity.DrawingPart])
	if err != nil {
		return graph, fmt.Errorf("drawing part %q: %w", identity.DrawingPart, err)
	}
	for _, anchor := range index.anchors {
		if anchor.objectID == identity.ObjectID && anchor.objectType == "graphicFrame" && anchor.chartRelID != "" {
			if graph.target.objectID != 0 {
				return graph, fmt.Errorf("chart object id %d is ambiguous", identity.ObjectID)
			}
			graph.target = anchor
		}
	}
	if graph.target.objectID == 0 {
		return graph, fmt.Errorf("stale or missing chart object id %d", identity.ObjectID)
	}
	if graph.target.anchorType != "twoCellAnchor" || graph.target.directObjects != 1 {
		return graph, fmt.Errorf("chart object id %d is in unsupported %s graph", identity.ObjectID, graph.target.anchorType)
	}
	target, err := exactChartRelationshipTarget(parts[graph.drawingRelsPart], identity.DrawingPart, graph.target.chartRelID)
	if err != nil {
		return graph, err
	}
	if target != identity.Part {
		return graph, fmt.Errorf("chart object %d targets %q, not identity part %q", identity.ObjectID, target, identity.Part)
	}
	if err := requireChartContentType(parts["[Content_Types].xml"], identity.Part); err != nil {
		return graph, err
	}
	refCount, err := packageRelationshipReferenceCount(parts, identity.Part)
	if err != nil {
		return graph, err
	}
	if refCount != 1 {
		return graph, fmt.Errorf("chart part %q has %d package relationship owners, expected 1", identity.Part, refCount)
	}
	owners, err := ReadChartIdentities(data)
	if err != nil {
		return graph, err
	}
	partOwners := 0
	for _, owner := range owners {
		if owner.Part == identity.Part {
			partOwners++
		}
	}
	if partOwners != 1 {
		return graph, fmt.Errorf("chart part %q has %d drawing owners, expected 1", identity.Part, partOwners)
	}
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return graph, err
	}
	sheetNames, err := sheetNamesByDrawingPart(zr)
	if err != nil {
		return graph, err
	}
	graph.sheetName = sheetNames[identity.DrawingPart]
	if graph.sheetName == "" {
		return graph, fmt.Errorf("drawing part %q has no worksheet owner", identity.DrawingPart)
	}
	return graph, nil
}

func requireChartContentType(data []byte, part string) error {
	overrides, _, err := directChildElements(data, "Types", "Override")
	if err != nil {
		return fmt.Errorf("parse content types: %w", err)
	}
	wanted, err := canonicalOPCPartKey(part)
	if err != nil {
		return err
	}
	matches := 0
	for _, override := range overrides {
		name := attribute(override.start, "PartName")
		key, err := canonicalOPCPartKey(name)
		if err != nil {
			return fmt.Errorf("invalid content-type PartName %q: %w", name, err)
		}
		if key != wanted {
			continue
		}
		matches++
		if contentType := attribute(override.start, "ContentType"); contentType != ctChart {
			return fmt.Errorf("chart part %q has content type %q, expected %q", part, contentType, ctChart)
		}
	}
	if matches != 1 {
		return fmt.Errorf("chart part %q has %d content-type overrides, expected 1", part, matches)
	}
	return nil
}

func packageRelationshipReferenceCount(parts map[string][]byte, targetPart string) (int, error) {
	count := 0
	for relsPart, data := range parts {
		if !strings.HasSuffix(relsPart, ".rels") || (!strings.Contains(relsPart, "/_rels/") && relsPart != "_rels/.rels") {
			continue
		}
		_, base := relationshipOwner(relsPart)
		rels, err := parsePackageRelationships(string(data))
		if err != nil {
			return 0, fmt.Errorf("malformed package relationships %q: %w", relsPart, err)
		}
		for _, rel := range rels {
			external, err := packageRelationshipIsExternal(rel)
			if err != nil {
				return 0, fmt.Errorf("relationship %q in %q: %w", rel.ID, relsPart, err)
			}
			if external {
				continue
			}
			resolved, err := resolveRelPath(base, rel.Target)
			if err != nil {
				return 0, fmt.Errorf("relationship %q in %q: %w", rel.ID, relsPart, err)
			}
			if resolved == targetPart {
				count++
			}
		}
	}
	return count, nil
}

func exactChartRelationshipTarget(data []byte, drawingPart, relationshipID string) (string, error) {
	rels, err := parsePackageRelationships(string(data))
	if err != nil {
		return "", fmt.Errorf("malformed drawing relationships %q: %w", relsPartFor(drawingPart), err)
	}
	var target string
	matches := 0
	for _, rel := range rels {
		if rel.ID != relationshipID {
			continue
		}
		matches++
		if rel.Type != relTypeChart && !strings.HasSuffix(rel.Type, "/chart") {
			return "", fmt.Errorf("chart relationship %q has type %q", relationshipID, rel.Type)
		}
		external, err := packageRelationshipIsExternal(rel)
		if err != nil || external {
			return "", fmt.Errorf("chart relationship %q is not internal", relationshipID)
		}
		base := drawingPart[:strings.LastIndex(drawingPart, "/")]
		target, err = resolveRelPath(base, rel.Target)
		if err != nil {
			return "", fmt.Errorf("chart relationship %q target: %w", relationshipID, err)
		}
	}
	if matches != 1 {
		return "", fmt.Errorf("chart relationship %q has %d entries, expected 1", relationshipID, matches)
	}
	return target, nil
}
