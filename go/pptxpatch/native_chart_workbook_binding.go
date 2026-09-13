package pptxpatch

import "encoding/xml"

const nativeChartWorkbookMaxBytes = 8 * 1024 * 1024
const nativeChartWorkbookContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

type nativeChartWorkbookBinding struct {
	RelationshipID string
	Part           string
	SHA256         string
	ByteLength     int64
	AutoUpdate     *bool
}

// Only the package-routed embedded workbook can authorize source cell data.
// External links, OLE and macro-enabled objects remain opaque. The XLSX engine
// separately validates the bytes; this boundary does not parse SpreadsheetML.
func (extractor *nativeExtractor) extractNativeChartWorkbookBinding(node *nativeXMLNode, chartPart string, d nativeExtractDialect) (*nativeChartWorkbookBinding, bool, error) {
	if node == nil || node.Name != (xml.Name{Space: d.chart, Local: "externalData"}) || requireOnlyNativeAttrs(node, xml.Name{Space: d.rels, Local: "id"}) != nil || !onlyNativeXMLSpace(node.Text) {
		return nil, false, nil
	}
	id, ok := exactNativeAttr(node, d.rels, "id")
	if !ok || id == "" || len(id) > 1024 {
		return nil, false, nil
	}
	var autoUpdate *bool
	if len(node.Children) > 1 {
		return nil, false, nil
	}
	if len(node.Children) == 1 {
		child := node.Children[0]
		if child.Name != (xml.Name{Space: d.chart, Local: "autoUpdate"}) {
			return nil, false, nil
		}
		raw, ok := nativeChartAttribute(child)
		if !ok {
			return nil, false, nil
		}
		value, e := nativeBool(raw)
		if e != nil {
			return nil, false, nil
		}
		autoUpdate = &value
	}
	relationships, exists, e := extractor.optionalRelationships(chartPart)
	if e != nil {
		return nil, false, e
	}
	if !exists {
		return nil, false, nil
	}
	wantedType := d.rels + "/package"
	for _, rel := range relationships {
		if rel.ID != id {
			continue
		}
		if rel.Type != wantedType || rel.Namespace != d.packageRels || rel.TargetMode == "External" || rel.Part == "" || extractor.pkg.contentTypes.forPart(rel.Part) != nativeChartWorkbookContentType {
			return nil, false, nil
		}
		bytes, present := extractor.pkg.parts[rel.Part]
		if !present || len(bytes) == 0 || len(bytes) > nativeChartWorkbookMaxBytes {
			return nil, false, nil
		}
		return &nativeChartWorkbookBinding{RelationshipID: id, Part: rel.Part, SHA256: nativeSHA256(bytes), ByteLength: int64(len(bytes)), AutoUpdate: autoUpdate}, true, nil
	}
	return nil, false, nil
}
