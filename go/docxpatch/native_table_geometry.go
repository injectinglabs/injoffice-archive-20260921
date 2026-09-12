package docxpatch

import "encoding/xml"

// Effective geometry is separate from source table markup. In particular auto
// width explicitly resets an inherited preferred width without changing bytes.
type NativeResolvedTableGeometryV1 struct {
	Layout      string                   `json:"layout"`
	Alignment   string                   `json:"alignment"`
	IndentTwips int64                    `json:"indent_twips"`
	WidthType   string                   `json:"width_type"`
	WidthValue  int64                    `json:"width_value"`
	CellMargins NativeTableCellMarginsV1 `json:"cell_margins"`
}

func nativeValidResolvedTableGeometry(g *NativeResolvedTableGeometryV1) bool {
	if g.Layout != "fixed" && g.Layout != "autofit" || g.Alignment != "left" {
		return false
	}
	for _, n := range []int64{g.IndentTwips, g.CellMargins.TopTwips, g.CellMargins.BottomTwips, g.CellMargins.LeftTwips, g.CellMargins.RightTwips} {
		if n < 0 || n > nativeMaxTwipsForMilliPoints {
			return false
		}
	}
	switch g.WidthType {
	case "auto":
		return g.WidthValue == 0
	case "dxa":
		return g.WidthValue > 0 && g.WidthValue <= nativeMaxTwipsForMilliPoints
	case "pct":
		return g.WidthValue > 0 && g.WidthValue <= 5000
	}
	return false
}

func (resolver *nativeLayoutResolver) resolveTableGeometry(table *NativeTableV1) *NativeResolvedTableGeometryV1 {
	// This first tier requires an explicit actual style reference: do not guess
	// a built-in style name or silently choose among default table definitions.
	if table.TableStyleID == nil {
		return nil
	}
	chain := resolver.styleChain("table", *table.TableStyleID, table.ID)
	if len(chain) == 0 || chain[0].basedOn != "" {
		return nil
	}
	g := &NativeResolvedTableGeometryV1{Layout: "autofit", Alignment: "left", WidthType: "auto", CellMargins: NativeTableCellMarginsV1{LeftTwips: 115, RightTwips: 115}}
	// ECMA-376 §§17.4.28, .42, .50, .52, .63 and margin child defaults:
	// left alignment, zero indent, auto layout/width, 0/115/0/115 margins.
	for _, style := range chain {
		if !nativeExactContainer(style.node, xml.Name{Space: resolver.wordNS, Local: "type"}, xml.Name{Space: resolver.wordNS, Local: "styleId"}, xml.Name{Space: resolver.wordNS, Local: "default"}, xml.Name{Space: resolver.wordNS, Local: "customStyle"}) {
			return nil
		}
		parents := directNativeChildren(style.node, resolver.wordNS, "basedOn")
		if len(parents) > 1 || len(parents) == 1 && !nativeExactLeaf(parents[0], xml.Name{Space: resolver.wordNS, Local: "val"}) {
			return nil
		}
		if len(parents) == 1 {
			if value, ok := nativeAttr(parents[0], resolver.wordNS, "val"); !ok || !nativeIDPattern.MatchString(value) {
				return nil
			}
		}
		for _, child := range style.node.Children {
			if child.Name.Space != resolver.wordNS {
				return nil
			}
		}
		if len(directNativeChildren(style.node, resolver.wordNS, "tblStylePr")) != 0 || len(directNativeChildren(style.node, resolver.wordNS, "tblPr")) > 1 {
			return nil
		}
		if !resolver.applyTableGeometry(g, firstDirectNativeChild(style.node, resolver.wordNS, "tblPr")) {
			return nil
		}
	}
	node := resolver.nodeForAnchor(table.Anchor)
	if node == nil || len(directNativeChildren(node, resolver.wordNS, "tblPr")) > 1 {
		return nil
	}
	if !resolver.applyTableGeometry(g, firstDirectNativeChild(node, resolver.wordNS, "tblPr")) || !nativeValidResolvedTableGeometry(g) {
		return nil
	}
	return g
}

func (resolver *nativeLayoutResolver) applyTableGeometry(g *NativeResolvedTableGeometryV1, node *nativeXMLNode) bool {
	if node == nil {
		return true
	}
	if !nativeExactContainer(node) {
		return false
	}
	seen := map[string]bool{}
	for _, child := range node.Children {
		if child.Name.Space != resolver.wordNS || seen[child.Name.Local] {
			return false
		}
		seen[child.Name.Local] = true
		switch child.Name.Local {
		case "tblLayout":
			v, ok := nativeAttr(child, resolver.wordNS, "type")
			if !ok || v != "fixed" && v != "autofit" || !nativeExactLeaf(child, xml.Name{Space: resolver.wordNS, Local: "type"}) {
				return false
			}
			g.Layout = v
		case "jc":
			v, ok := nativeAttr(child, resolver.wordNS, "val")
			if !ok || v != "left" || !nativeExactLeaf(child, xml.Name{Space: resolver.wordNS, Local: "val"}) {
				return false
			}
			g.Alignment = v
		case "tblW", "tblInd":
			v, ok := nativeNonnegativeInt64Attr(child, resolver.wordNS, "w")
			kind, has := nativeAttr(child, resolver.wordNS, "type")
			if !ok || !has || !nativeExactLeaf(child, xml.Name{Space: resolver.wordNS, Local: "w"}, xml.Name{Space: resolver.wordNS, Local: "type"}) {
				return false
			}
			if child.Name.Local == "tblInd" {
				if kind != "dxa" {
					return false
				}
				g.IndentTwips = v
			} else {
				if kind == "auto" && v != 0 {
					return false
				}
				g.WidthType = kind
				g.WidthValue = v
			}
		case "tblCellMar":
			if !nativeExactContainer(child) {
				return false
			}
			sides := map[string]bool{}
			for _, side := range child.Children {
				if side.Name.Space != resolver.wordNS || sides[side.Name.Local] {
					return false
				}
				sides[side.Name.Local] = true
				v, ok := nativeNonnegativeInt64Attr(side, resolver.wordNS, "w")
				kind, has := nativeAttr(side, resolver.wordNS, "type")
				if !ok || has && kind != "dxa" || !nativeExactLeaf(side, xml.Name{Space: resolver.wordNS, Local: "w"}, xml.Name{Space: resolver.wordNS, Local: "type"}) {
					return false
				}
				switch side.Name.Local {
				case "top":
					g.CellMargins.TopTwips = v
				case "left":
					g.CellMargins.LeftTwips = v
				case "bottom":
					g.CellMargins.BottomTwips = v
				case "right":
					g.CellMargins.RightTwips = v
				default:
					return false
				}
			}
		case "tblStyle":
			v, ok := nativeAttr(child, resolver.wordNS, "val")
			if !ok || !nativeIDPattern.MatchString(v) || !nativeExactLeaf(child, xml.Name{Space: resolver.wordNS, Local: "val"}) {
				return false
			}
		case "tblLook":
			extractor := nativeExtractor{pkg: resolver.pkg, mainPart: resolver.mainPart, wordNS: resolver.wordNS, relNS: resolver.relBase[:len(resolver.relBase)-1]}
			if !extractor.inactiveTableLook(node, child) {
				return false
			}
		case "tblBorders":
			// Independently validated by source extraction and style/paint gates.
		default:
			return false
		}
	}
	return nativeValidResolvedTableGeometry(g)
}
