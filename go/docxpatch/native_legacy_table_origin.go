package docxpatch

import (
	"crypto/sha256"
	"encoding/hex"
)

// Source evidence only: this does not apply a shift or claim legacy Word parity.
type NativeDocxTableOriginSourceV1 struct {
	PartName string `json:"part_name"`
	Path     string `json:"path"`
	SHA256   string `json:"sha256"`
}
type NativeDocxLegacyTableOriginV1 struct {
	TableID         string                        `json:"table_id"`
	PackageSHA256   string                        `json:"package_sha256"`
	IndentTwips     int64                         `json:"indent_twips"`
	LeftMarginTwips int64                         `json:"left_margin_twips"`
	SourceIndent    NativeDocxTableOriginSourceV1 `json:"source_indent"`
	SourceMargin    NativeDocxTableOriginSourceV1 `json:"source_margin"`
}

func nativeLegacyTableOrigins(data []byte) ([]NativeDocxLegacyTableOriginV1, error) {
	r, err := newNativeLayoutResolver(data, NativeExtractionOptions{})
	if err != nil {
		return nil, err
	}
	if _, err = r.resolve(); err != nil {
		return nil, err
	}
	if r.parts.StylesPart == nil {
		return nil, nil
	}
	var facts []NativeDocxLegacyTableOriginV1
	for _, block := range r.doc.Body.Blocks {
		table := block.Table
		if table == nil || table.TableStyleID == nil {
			continue
		}
		geometry := r.resolveTableGeometry(table)
		if geometry == nil || geometry.Alignment != "left" || geometry.CellMargins.LeftTwips == 0 {
			continue
		}
		chain := r.styleChain("table", *table.TableStyleID, table.ID)
		styleSafe := true
		for _, style := range chain {
			if len(directNativeChildren(style.node, r.wordNS, "tcPr")) != 0 {
				styleSafe = false
			}
		}
		if !styleSafe {
			continue
		}
		var indent, margin *nativeXMLNode
		indentPart, marginPart := "", ""
		visit := func(owner *nativeXMLNode, part string) {
			if owner == nil {
				return
			}
			if node := firstDirectNativeChild(owner, r.wordNS, "tblInd"); node != nil {
				indent = node
				indentPart = part
			}
			if cellMargins := firstDirectNativeChild(owner, r.wordNS, "tblCellMar"); cellMargins != nil {
				if node := firstDirectNativeChild(cellMargins, r.wordNS, "left"); node != nil {
					margin = node
					marginPart = part
				}
			}
		}
		for _, style := range chain {
			visit(firstDirectNativeChild(style.node, r.wordNS, "tblPr"), *r.parts.StylesPart)
		}
		owner := r.nodeForAnchor(table.Anchor)
		if owner == nil {
			continue
		}
		visit(firstDirectNativeChild(owner, r.wordNS, "tblPr"), table.Anchor.PartName)
		if indent == nil || margin == nil {
			continue
		}
		if kind, ok := nativeAttr(margin, r.wordNS, "type"); !ok || kind != "dxa" {
			continue
		}
		clean := true
		for _, row := range table.Rows {
			rowNode := r.nodeForAnchor(row.Anchor)
			if rowNode == nil || len(directNativeChildren(rowNode, r.wordNS, "tblPrEx")) > 0 {
				clean = false
				break
			}
			for _, cell := range row.Cells {
				if cell.GridSpan == nil || *cell.GridSpan != 1 || cell.VerticalMerge != "none" {
					clean = false
					break
				}
				node := r.nodeForAnchor(cell.Anchor)
				if node == nil {
					clean = false
					break
				}
				if properties := firstDirectNativeChild(node, r.wordNS, "tcPr"); properties != nil {
					for _, child := range properties.Children {
						if child.Name.Space != r.wordNS || child.Name.Local != "tcW" {
							clean = false
						}
					}
				}
			}
		}
		if !clean {
			continue
		}
		source := func(node *nativeXMLNode, part string) NativeDocxTableOriginSourceV1 {
			digest := sha256.Sum256(r.pkg.files[part])
			return NativeDocxTableOriginSourceV1{PartName: part, Path: node.Path, SHA256: "sha256:" + hex.EncodeToString(digest[:])}
		}
		facts = append(facts, NativeDocxLegacyTableOriginV1{TableID: table.ID, PackageSHA256: r.doc.Source.PackageSHA256, IndentTwips: geometry.IndentTwips, LeftMarginTwips: geometry.CellMargins.LeftTwips, SourceIndent: source(indent, indentPart), SourceMargin: source(margin, marginPart)})
		if len(facts) > 1000 {
			return nil, nil
		}
	}
	return facts, nil
}
