package xlsxpatch

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"strconv"
	"strings"
)

// RemoveShape removes exactly one top-level DrawingML shape anchor. The
// drawing part and its worksheet relationship remain in place, which keeps
// every chart, image, group, relationship, extension, and sibling anchor
// byte-for-byte intact. A shape with inbound connector bindings is refused so
// the operation cannot leave dangling connection-site references.
func RemoveShape(orig []byte, identity ShapeIdentity) ([]byte, error) {
	graph, err := resolveShapeLifecycleGraph(orig, identity)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: remove shape: %w", err)
	}
	if graph.index.inbound[identity.ObjectID] > 0 {
		return nil, fmt.Errorf("xlsxpatch: remove shape: object %d has %d inbound connector bindings", identity.ObjectID, graph.index.inbound[identity.ObjectID])
	}
	updated := removeElementSpans(graph.parts[identity.DrawingPart], []xmlSpan{graph.target.span})
	return Apply(orig, Patch{Replace: map[string][]byte{identity.DrawingPart: updated}})
}

// UpdateShape replaces one hydrated top-level shape while retaining its
// DrawingML object id. Existing connectors that point at the shape therefore
// remain valid. When the target itself is a bound connector, its native
// cNvCxnSpPr is copied exactly into a line replacement; changing such a
// connector into a non-line shape fails closed.
func UpdateShape(orig []byte, identity ShapeIdentity, spec ShapeWriteSpec) ([]byte, error) {
	if err := validateShapeWriteSpec(spec); err != nil {
		return nil, fmt.Errorf("xlsxpatch: update shape: %w", err)
	}
	graph, err := resolveShapeLifecycleGraph(orig, identity)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: update shape: %w", err)
	}
	if spec.SheetName != graph.sheetName {
		return nil, fmt.Errorf("xlsxpatch: update shape: identity belongs to worksheet %q, not %q", graph.sheetName, spec.SheetName)
	}
	connectorProperties := graph.target.connectorProperties
	if graph.target.boundConnector && spec.Kind != "line" {
		return nil, fmt.Errorf("xlsxpatch: update shape: bound connector %d cannot be changed into %q", identity.ObjectID, spec.Kind)
	}
	if spec.Kind != "line" {
		connectorProperties = nil
	}
	replacement := []byte(shapeAnchorXMLWithIdentity(spec, identity.ObjectID, connectorProperties))
	data := graph.parts[identity.DrawingPart]
	updated := make([]byte, 0, len(data)-graph.target.span.end+graph.target.span.start+len(replacement))
	updated = append(updated, data[:graph.target.span.start]...)
	updated = append(updated, replacement...)
	updated = append(updated, data[graph.target.span.end:]...)
	out, err := Apply(orig, Patch{Replace: map[string][]byte{identity.DrawingPart: updated}})
	if err != nil {
		return nil, err
	}
	shapes, err := ReadShapes(out)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: update shape: verify identity: %w", err)
	}
	for _, shape := range shapes {
		if shape.Identity == identity {
			return out, nil
		}
	}
	return nil, fmt.Errorf("xlsxpatch: update shape: replacement did not retain identity %+v", identity)
}

func validateShapeWriteSpec(spec ShapeWriteSpec) error {
	if !shapeWritable(spec.Kind) {
		return fmt.Errorf("shape kind %q is not in the supported preset catalogue", spec.Kind)
	}
	if spec.Anchor.FromCol < 0 || spec.Anchor.FromRow < 0 || spec.Anchor.ToCol <= spec.Anchor.FromCol || spec.Anchor.ToRow <= spec.Anchor.FromRow {
		return fmt.Errorf("degenerate shape anchor")
	}
	return nil
}

type shapeLifecycleGraph struct {
	parts     map[string][]byte
	sheetName string
	index     shapeDrawingIndex
	target    shapeAnchorNode
}

func resolveShapeLifecycleGraph(data []byte, identity ShapeIdentity) (shapeLifecycleGraph, error) {
	graph := shapeLifecycleGraph{}
	if identity.DrawingPart == "" || !strings.HasPrefix(identity.DrawingPart, "xl/drawings/") || !strings.HasSuffix(identity.DrawingPart, ".xml") || strings.Contains(identity.DrawingPart, "\\") || strings.Contains(identity.DrawingPart, "/_rels/") || identity.ObjectID == 0 {
		return graph, fmt.Errorf("invalid shape identity %+v", identity)
	}
	if _, err := canonicalNativeContractPartKey(identity.DrawingPart); err != nil {
		return graph, fmt.Errorf("invalid shape identity %+v: %w", identity, err)
	}
	parts, err := readPivotArchive(data)
	if err != nil {
		return graph, err
	}
	graph.parts = parts
	drawing, ok := parts[identity.DrawingPart]
	if !ok {
		return graph, fmt.Errorf("stale or missing drawing part %q", identity.DrawingPart)
	}
	index, err := indexShapeDrawing(drawing)
	if err != nil {
		return graph, fmt.Errorf("drawing part %q: %w", identity.DrawingPart, err)
	}
	graph.index = index
	for _, anchor := range index.anchors {
		if anchor.objectID == identity.ObjectID && (anchor.objectType == "sp" || anchor.objectType == "cxnSp") {
			if graph.target.objectID != 0 {
				return graph, fmt.Errorf("object id %d is ambiguous", identity.ObjectID)
			}
			graph.target = anchor
		}
	}
	if graph.target.objectID == 0 {
		return graph, fmt.Errorf("stale or missing shape object id %d", identity.ObjectID)
	}
	if graph.target.anchorType != "twoCellAnchor" || graph.target.directObjects != 1 {
		return graph, fmt.Errorf("shape object id %d is in unsupported %s graph with %d top-level drawing objects", identity.ObjectID, graph.target.anchorType, graph.target.directObjects)
	}

	read := func(name string) (string, bool) { value, found := parts[name]; return string(value), found }
	sheets, err := readWorkbookSheets(read)
	if err != nil {
		return graph, err
	}
	owners := 0
	for _, sheet := range sheets {
		sheetXML, ok := parts[sheet.Part]
		if !ok {
			return graph, fmt.Errorf("missing worksheet part %q", sheet.Part)
		}
		drawingRelID := drawingRelIDInSheet(string(sheetXML))
		if drawingRelID == "" {
			continue
		}
		relsPart := relsPartFor(sheet.Part)
		relsData, ok := parts[relsPart]
		if !ok {
			return graph, fmt.Errorf("worksheet %q references a drawing but has no relationships part", sheet.Name)
		}
		rels, err := parsePackageRelationships(string(relsData))
		if err != nil {
			return graph, fmt.Errorf("worksheet relationships %q: %w", relsPart, err)
		}
		matches := 0
		for _, rel := range rels {
			if rel.ID != drawingRelID {
				continue
			}
			matches++
			if rel.Type != relTypeDrawing && !strings.HasSuffix(rel.Type, "/drawing") {
				return graph, fmt.Errorf("worksheet drawing relationship %q has type %q", rel.ID, rel.Type)
			}
			external, err := packageRelationshipIsExternal(rel)
			if err != nil || external {
				return graph, fmt.Errorf("worksheet drawing relationship %q is not internal", rel.ID)
			}
			base := sheet.Part[:strings.LastIndex(sheet.Part, "/")]
			target, err := resolveRelPath(base, rel.Target)
			if err != nil {
				return graph, fmt.Errorf("worksheet drawing relationship %q: %w", rel.ID, err)
			}
			if target == identity.DrawingPart {
				owners++
				graph.sheetName = sheet.Name
			}
		}
		if matches != 1 {
			return graph, fmt.Errorf("worksheet %q has %d relationships for drawing id %q, expected 1", sheet.Name, matches, drawingRelID)
		}
	}
	if owners != 1 {
		return graph, fmt.Errorf("drawing part %q has %d worksheet owners, expected 1", identity.DrawingPart, owners)
	}
	return graph, nil
}

type shapeAnchorNode struct {
	span                xmlSpan
	anchorType          string
	objectType          string
	objectID            uint32
	chartRelID          string
	directObjects       int
	connectorProperties []byte
	boundConnector      bool
}

type shapeDrawingIndex struct {
	anchors     []shapeAnchorNode
	maxObjectID uint32
	inbound     map[uint32]int
}

// indexShapeDrawing indexes direct worksheet-drawing anchors without
// reserializing XML. It also validates global cNvPr uniqueness, including
// charts, pictures, and shapes nested in groups, before any mutation occurs.
func indexShapeDrawing(data []byte) (shapeDrawingIndex, error) {
	result := shapeDrawingIndex{inbound: map[uint32]int{}}
	decoder := xml.NewDecoder(bytes.NewReader(data))
	depth := 0
	rootNS := ""
	seenIDs := map[uint32]bool{}
	var active *shapeAnchorNode
	var connectorSpan *xmlSpan
	for {
		before := int(decoder.InputOffset())
		token, err := decoder.Token()
		after := int(decoder.InputOffset())
		if err == io.EOF {
			break
		}
		if err != nil {
			return result, err
		}
		switch token := token.(type) {
		case xml.StartElement:
			depth++
			if depth == 1 {
				if token.Name.Local != "wsDr" {
					return result, fmt.Errorf("root element is %q, expected wsDr", token.Name.Local)
				}
				rootNS = token.Name.Space
				continue
			}
			if depth == 2 && token.Name.Space == rootNS && (token.Name.Local == "twoCellAnchor" || token.Name.Local == "oneCellAnchor" || token.Name.Local == "absoluteAnchor") {
				if active != nil {
					return result, fmt.Errorf("nested drawing anchors")
				}
				active = &shapeAnchorNode{span: xmlSpan{qname: rawQName(data, before, after), start: before, startTagEnd: after}, anchorType: token.Name.Local}
				continue
			}
			if active != nil && depth == 3 && token.Name.Space == rootNS && isDrawingObject(token.Name.Local) {
				active.directObjects++
				if active.directObjects == 1 {
					active.objectType = token.Name.Local
				}
			}
			if token.Name.Local == "cNvPr" {
				raw := attrVal(token, "id")
				value, parseErr := strconv.ParseUint(raw, 10, 32)
				if parseErr != nil || value == 0 {
					return result, fmt.Errorf("invalid cNvPr id %q", raw)
				}
				id := uint32(value)
				if seenIDs[id] {
					return result, fmt.Errorf("duplicate cNvPr id %d", id)
				}
				seenIDs[id] = true
				if id > result.maxObjectID {
					result.maxObjectID = id
				}
				if active != nil && depth == 5 && isDrawingObject(active.objectType) {
					active.objectID = id
				}
			}
			if active != nil && active.objectType == "graphicFrame" && token.Name.Local == "chart" {
				if id := attrVal(token, "id"); id != "" {
					if active.chartRelID != "" && active.chartRelID != id {
						return result, fmt.Errorf("graphicFrame contains multiple chart relationships")
					}
					active.chartRelID = id
				}
			}
			if active != nil && active.objectType == "cxnSp" && token.Name.Local == "cNvCxnSpPr" && depth == 5 {
				span := xmlSpan{qname: rawQName(data, before, after), start: before, startTagEnd: after}
				connectorSpan = &span
			}
			if token.Name.Local == "stCxn" || token.Name.Local == "endCxn" {
				raw := attrVal(token, "id")
				value, parseErr := strconv.ParseUint(raw, 10, 32)
				if parseErr != nil || value == 0 {
					return result, fmt.Errorf("invalid %s target id %q", token.Name.Local, raw)
				}
				result.inbound[uint32(value)]++
				if active != nil && active.objectType == "cxnSp" {
					active.boundConnector = true
				}
			}
		case xml.EndElement:
			if connectorSpan != nil && token.Name.Local == "cNvCxnSpPr" && depth == 5 {
				connectorSpan.endStart, connectorSpan.end = before, after
				active.connectorProperties = bytes.Clone(data[connectorSpan.start:connectorSpan.end])
				connectorSpan = nil
			}
			if active != nil && depth == 2 && token.Name.Space == rootNS && token.Name.Local == active.anchorType {
				active.span.endStart, active.span.end = before, after
				if (active.objectType == "sp" || active.objectType == "cxnSp") && active.objectID == 0 {
					return result, fmt.Errorf("top-level %s has no cNvPr identity", active.objectType)
				}
				result.anchors = append(result.anchors, *active)
				active = nil
				connectorSpan = nil
			}
			depth--
		}
	}
	if depth != 0 || rootNS == "" {
		return result, fmt.Errorf("incomplete wsDr document")
	}
	return result, nil
}

func isDrawingObject(local string) bool {
	switch local {
	case "sp", "cxnSp", "grpSp", "graphicFrame", "pic", "contentPart":
		return true
	default:
		return false
	}
}
