package docxpatch

import "encoding/xml"

const nativePartialNestedTableLimit = 64

type NativePartialNestedTableV1 struct {
	PackageSHA256 string               `json:"package_sha256"`
	PartSHA256    string               `json:"part_sha256"`
	TableID       string               `json:"table_id"`
	CellID        string               `json:"cell_id"`
	DiagnosticID  string               `json:"diagnostic_id"`
	Anchor        NativeSourceAnchorV1 `json:"anchor"`
}
type NativePartialNestedTablesV1 struct {
	Items        []NativePartialNestedTableV1 `json:"items"`
	OmittedCount int                          `json:"omitted_count"`
}

// Containment evidence only: the nested subtree stays opaque, diagnostic-bearing
// and read-only. No nested text, style or geometry is projected.
func inspectNativePartialNestedTables(data []byte, doc *NativeDocumentV1) (*NativePartialNestedTablesV1, error) {
	pkg, err := openNativeDOCXPackage(data)
	if err != nil {
		return nil, err
	}
	_, strict, err := pkg.officeDocumentPart()
	if err != nil {
		return nil, err
	}
	ns := wordMLTransitional
	if strict {
		ns = wordMLStrict
	}
	raw := pkg.files[doc.Source.MainPart]
	partSHA := nativeSHA(raw)
	root, err := parseNativeXML(doc.Source.MainPart, raw)
	if err != nil {
		return nil, err
	}
	nodes := map[string]*nativeXMLNode{}
	var visit func(*nativeXMLNode)
	visit = func(n *nativeXMLNode) {
		nodes[n.Path] = n
		for _, child := range n.Children {
			visit(child)
		}
	}
	visit(root)
	diagnostics := map[string][]NativeUnsupportedCapabilityV1{}
	for _, d := range doc.Unsupported {
		if d.Code == "NESTED_TABLE_OR_CELL_MARKUP" && d.Anchor != nil {
			diagnostics[d.ScopeID+"\x00"+d.Anchor.Path] = append(diagnostics[d.ScopeID+"\x00"+d.Anchor.Path], d)
		}
	}
	out := &NativePartialNestedTablesV1{Items: []NativePartialNestedTableV1{}}
	for _, block := range doc.Body.Blocks {
		if block.Table == nil {
			continue
		}
		table := block.Table
		for _, row := range table.Rows {
			for _, cell := range row.Cells {
				owner := nodes[cell.Anchor.Path]
				if owner == nil || owner.Name != (xml.Name{Space: ns, Local: "tc"}) || !nativeExactContainer(owner) || len(directNativeChildren(owner, ns, "tcPr")) > 1 || owner.parent == nil || owner.parent.Path != row.Anchor.Path || !nativeExactRevisionContainer(owner.parent, ns, "rsidR", "rsidRPr", "rsidTr") || owner.parent.parent == nil || owner.parent.parent.Path != table.Anchor.Path || !nativeExactContainer(owner.parent.parent) {
					continue
				}
				for _, child := range owner.Children {
					if child.Name != (xml.Name{Space: ns, Local: "tbl"}) {
						continue
					}
					anchor := NativeSourceAnchorV1{PartName: doc.Source.MainPart, Path: child.Path, StartByte: nativeInt64(child.Start), EndByte: nativeInt64(child.End), XMLSHA256: nativeSHA(raw[child.Start:child.End])}
					var match string
					count := 0
					for _, d := range diagnostics[table.ID+"\x00"+child.Path] {
						if d.Code == "NESTED_TABLE_OR_CELL_MARKUP" && d.ScopeID == table.ID && d.Anchor != nil && d.Anchor.Path == anchor.Path && d.Anchor.XMLSHA256 == anchor.XMLSHA256 && d.Anchor.PartName == anchor.PartName {
							match = d.ID
							count++
						}
					}
					if count != 1 {
						continue
					}
					if len(out.Items) >= nativePartialNestedTableLimit {
						out.OmittedCount++
						continue
					}
					out.Items = append(out.Items, NativePartialNestedTableV1{doc.Source.PackageSHA256, partSHA, table.ID, cell.ID, match, anchor})
				}
			}
		}
	}
	if len(out.Items) == 0 && out.OmittedCount == 0 {
		return nil, nil
	}
	return out, nil
}
