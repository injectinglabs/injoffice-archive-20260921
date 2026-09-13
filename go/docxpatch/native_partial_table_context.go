package docxpatch

import (
	"encoding/xml"
	"strconv"
)

type NativePartialTableTextStyleV1 struct {
	StyleID string               `json:"style_id"`
	Anchor  NativeSourceAnchorV1 `json:"anchor"`
}
type NativePartialTableTextContextV1 struct {
	PackageSHA256       string                                   `json:"package_sha256"`
	TableID             string                                   `json:"table_id"`
	LookDiagnosticID    string                                   `json:"look_diagnostic_id"`
	LookAnchor          NativeSourceAnchorV1                     `json:"look_anchor"`
	StylesPart          string                                   `json:"styles_part"`
	StylesSHA256        string                                   `json:"styles_sha256"`
	StyleChain          []NativePartialTableTextStyleV1          `json:"style_chain"`
	ResolvedDiagnostics []NativeAutomaticTableBorderDiagnosticV1 `json:"resolved_diagnostics"`
}

func nativePartialTextLook(n *nativeXMLNode, ns string) bool {
	attrs := []xml.Name{{Space: ns, Local: "val"}}
	names := []string{"firstRow", "lastRow", "firstColumn", "lastColumn", "noHBand", "noVBand"}
	for _, name := range names {
		attrs = append(attrs, xml.Name{Space: ns, Local: name})
	}
	if !nativeExactLeaf(n, attrs...) {
		return false
	}
	value, ok := nativeAttr(n, ns, "val")
	if !ok || len(value) != 4 {
		return false
	}
	bits, err := strconv.ParseUint(value, 16, 16)
	if err != nil || bits&^uint64(0x7e0) != 0 {
		return false
	}
	for i, name := range names {
		if value, has := nativeAttr(n, ns, name); has {
			if value != "0" && value != "1" {
				return false
			}
			if (value == "1") != (bits&(uint64(0x20)<<i) != 0) {
				return false
			}
		}
	}
	return true
}
func nativePartialTextWidth(n *nativeXMLNode, ns string) bool {
	if !nativeExactLeaf(n, xml.Name{Space: ns, Local: "w"}, xml.Name{Space: ns, Local: "type"}) {
		return false
	}
	kind, _ := nativeAttr(n, ns, "type")
	width, ok := nativeNonnegativeInt64Attr(n, ns, "w")
	return kind == "dxa" && ok && width <= 31680
}
func nativePartialTextTableProperties(n *nativeXMLNode, ns string) bool {
	if !nativeExactContainer(n) {
		return false
	}
	seen := map[string]bool{}
	for _, c := range n.Children {
		if c.Name.Space != ns || seen[c.Name.Local] {
			return false
		}
		seen[c.Name.Local] = true
		switch c.Name.Local {
		case "tblInd":
			if !nativePartialTextWidth(c, ns) {
				return false
			}
		case "tblCellMar":
			if !nativeExactContainer(c) {
				return false
			}
			edges := map[string]bool{}
			for _, e := range c.Children {
				if e.Name.Space != ns || edges[e.Name.Local] || (e.Name.Local != "top" && e.Name.Local != "bottom" && e.Name.Local != "left" && e.Name.Local != "right") || !nativePartialTextWidth(e, ns) {
					return false
				}
				edges[e.Name.Local] = true
			}
		case "tblBorders":
			if !nativeExactContainer(c) {
				return false
			}
			edges := map[string]bool{}
			for _, e := range c.Children {
				if e.Name.Space != ns || edges[e.Name.Local] {
					return false
				}
				edges[e.Name.Local] = true
				switch e.Name.Local {
				case "top", "bottom", "left", "right", "insideH", "insideV":
				default:
					return false
				}
				if !nativeExactLeaf(e, xml.Name{Space: ns, Local: "val"}, xml.Name{Space: ns, Local: "sz"}, xml.Name{Space: ns, Local: "space"}, xml.Name{Space: ns, Local: "color"}) {
					return false
				}
				kind, _ := nativeAttr(e, ns, "val")
				if kind != "single" && kind != "nil" && kind != "none" {
					return false
				}
				if value, has := nativeAttr(e, ns, "color"); has && value != "auto" {
					if _, ok := nativeExactRGB(value); !ok {
						return false
					}
				}
				for _, name := range []string{"sz", "space"} {
					if _, has := nativeAttr(e, ns, name); has {
						n, ok := nativeNonnegativeInt64Attr(e, ns, name)
						if !ok || name == "sz" && n > 768 || name == "space" && n > 31 {
							return false
						}
					}
				}
			}
		default:
			return false
		}
	}
	return true
}
func nativePartialTextTableStyle(n *nativeXMLNode, ns string) bool {
	if !nativeExactContainer(n, xml.Name{Space: ns, Local: "type"}, xml.Name{Space: ns, Local: "styleId"}, xml.Name{Space: ns, Local: "default"}, xml.Name{Space: ns, Local: "customStyle"}) {
		return false
	}
	kind, _ := nativeAttr(n, ns, "type")
	if kind != "table" {
		return false
	}
	for _, name := range []string{"default", "customStyle"} {
		if value, has := nativeAttr(n, ns, name); has && value != "0" && value != "1" {
			return false
		}
	}
	seen := map[string]bool{}
	for _, c := range n.Children {
		if c.Name.Space != ns || seen[c.Name.Local] {
			return false
		}
		seen[c.Name.Local] = true
		switch c.Name.Local {
		case "name", "basedOn", "uiPriority", "rsid":
			if !nativeExactLeaf(c, xml.Name{Space: ns, Local: "val"}) {
				return false
			}
			value, present := nativeAttr(c, ns, "val")
			if !present || value == "" || len(value) > 256 {
				return false
			}
			if c.Name.Local == "basedOn" && !nativeIDPattern.MatchString(value) {
				return false
			}
		case "semiHidden", "unhideWhenUsed", "qFormat":
			if !nativeExactLeaf(c, xml.Name{Space: ns, Local: "val"}) {
				return false
			}
			if _, ok := nativeOnOff(c, ns); !ok {
				return false
			}
		case "tblPr":
			if !nativePartialTextTableProperties(c, ns) {
				return false
			}
		case "pPr":
			if !nativeExactContainer(c) || len(c.Children) > 1 {
				return false
			}
			for _, spacing := range c.Children {
				if spacing.Name != (xml.Name{Space: ns, Local: "spacing"}) || !nativeExactLeaf(spacing, xml.Name{Space: ns, Local: "before"}, xml.Name{Space: ns, Local: "after"}, xml.Name{Space: ns, Local: "line"}, xml.Name{Space: ns, Local: "lineRule"}) {
					return false
				}
				for _, name := range []string{"before", "after", "line"} {
					if _, has := nativeAttr(spacing, ns, name); has {
						n, ok := nativeNonnegativeInt64Attr(spacing, ns, name)
						if !ok || n > 31680 {
							return false
						}
					}
				}
				if rule, has := nativeAttr(spacing, ns, "lineRule"); has && rule != "auto" && rule != "exact" && rule != "atLeast" {
					return false
				}
			}
		default:
			return false // Includes run properties, conditional styles and revisions.
		}
	}
	return true
}

// Geometry-only style evidence is only for the explicit plain-text policy.
// Original diagnostics, strict paint and mutation authority are unchanged.
func inspectNativePartialTableTextContexts(data []byte, doc *NativeDocumentV1, layout *NativeResolvedLayoutInputV1) ([]NativePartialTableTextContextV1, error) {
	candidate := false
	for _, d := range doc.Unsupported {
		if d.Code == "UNMODELED_TABLE_PROPERTY" && d.Anchor != nil {
			candidate = true
			break
		}
	}
	if !candidate {
		return nil, nil
	}
	resolver, err := newNativeLayoutResolver(data, NativeExtractionOptions{})
	if err != nil {
		return nil, err
	}
	if resolver.parts.StylesPart == nil {
		return nil, nil
	}
	part := *resolver.parts.StylesPart
	out := []NativePartialTableTextContextV1{}
	for _, b := range doc.Body.Blocks {
		t := b.Table
		if t == nil || t.TableStyleID == nil || len(out) >= 64 {
			continue
		}
		node := resolver.nodeForAnchor(t.Anchor)
		if node == nil {
			continue
		}
		props := directNativeChildren(node, resolver.wordNS, "tblPr")
		if len(props) != 1 || !nativeExactContainer(props[0]) {
			continue
		}
		looks := directNativeChildren(props[0], resolver.wordNS, "tblLook")
		styles := directNativeChildren(props[0], resolver.wordNS, "tblStyle")
		if len(looks) != 1 || !nativePartialTextLook(looks[0], resolver.wordNS) || len(styles) != 1 || !nativeExactLeaf(styles[0], xml.Name{Space: resolver.wordNS, Local: "val"}) {
			continue
		}
		styleID, _ := nativeAttr(styles[0], resolver.wordNS, "val")
		if styleID != *t.TableStyleID {
			continue
		}
		chain := resolver.styleChain("table", styleID, t.ID)
		if len(chain) == 0 || len(chain) > 16 || chain[0].basedOn != "" {
			continue
		}
		evidence := NativePartialTableTextContextV1{PackageSHA256: doc.Source.PackageSHA256, TableID: t.ID, StylesPart: part, StylesSHA256: nativeSHA(resolver.pkg.files[part]), StyleChain: []NativePartialTableTextStyleV1{}, ResolvedDiagnostics: []NativeAutomaticTableBorderDiagnosticV1{}}
		exact := true
		paths := map[string]bool{}
		for _, layer := range chain {
			if layer.partName != part || !nativePartialTextTableStyle(layer.node, resolver.wordNS) {
				exact = false
				break
			}
			n := layer.node
			id, _ := nativeAttr(n, resolver.wordNS, "styleId")
			evidence.StyleChain = append(evidence.StyleChain, NativePartialTableTextStyleV1{id, NativeSourceAnchorV1{PartName: part, Path: n.Path, StartByte: nativeInt64(n.Start), EndByte: nativeInt64(n.End), XMLSHA256: nativeSHA(resolver.pkg.files[part][n.Start:n.End])}})
			paths[n.Path+"/w:tblPr[1]"] = true
		}
		if !exact {
			continue
		}
		for _, d := range doc.Unsupported {
			if d.Code == "UNMODELED_TABLE_PROPERTY" && d.ScopeID == t.ID && d.Anchor != nil && d.Anchor.Path == looks[0].Path && d.Anchor.PartName == doc.Source.MainPart {
				if evidence.LookDiagnosticID != "" {
					exact = false
					break
				}
				evidence.LookDiagnosticID = d.ID
				evidence.LookAnchor = *d.Anchor
			}
		}
		if !exact || evidence.LookDiagnosticID == "" {
			continue
		}
		diagnosticPaths := map[string]bool{}
		for _, d := range layout.Diagnostics {
			if d.Code == "TABLE_STYLE_EFFECTS_PRESERVED" && d.ScopeID == t.ID && d.PartName != nil && *d.PartName == part && d.Path != nil && paths[*d.Path] {
				if diagnosticPaths[*d.Path] {
					exact = false
					break
				}
				diagnosticPaths[*d.Path] = true
				evidence.ResolvedDiagnostics = append(evidence.ResolvedDiagnostics, NativeAutomaticTableBorderDiagnosticV1{d.Code, d.ScopeID, *d.PartName, *d.Path})
			}
		}
		if !exact || len(evidence.ResolvedDiagnostics) > 16 {
			continue
		}
		out = append(out, evidence)
	}
	if len(out) == 0 {
		return nil, nil
	}
	return out, nil
}
