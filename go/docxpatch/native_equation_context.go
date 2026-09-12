package docxpatch

import (
	"encoding/xml"
	"fmt"
)

type NativeEquationContextTabV1 struct {
	Kind          string `json:"kind"`
	PositionTwips int64  `json:"position_twips"`
	Leader        string `json:"leader"`
}
type NativeEquationContextNoticeV1 struct {
	Kind             string                       `json:"kind"`
	PackageSHA256    string                       `json:"package_sha256"`
	PartSHA256       string                       `json:"part_sha256"`
	Anchor           NativeSourceAnchorV1         `json:"anchor"`
	DiagnosticOrigin string                       `json:"diagnostic_origin"`
	DiagnosticID     string                       `json:"diagnostic_id,omitempty"`
	Code             string                       `json:"code"`
	ScopeID          string                       `json:"scope_id"`
	Value            string                       `json:"value,omitempty"`
	CharacterSet     string                       `json:"character_set,omitempty"`
	TabStops         []NativeEquationContextTabV1 `json:"tab_stops,omitempty"`
}

// These notices authorize only an explicitly unpaginated browser MathML view.
// They neither qualify fonts nor change extraction/resolution diagnostics.
func inspectNativeEquationContext(data []byte, doc *NativeDocumentV1, resolved *NativeResolvedLayoutInputV1, equations []NativePartialEquationV1) ([]NativeEquationContextNoticeV1, error) {
	out := []NativeEquationContextNoticeV1{}
	if len(equations) == 0 {
		return out, nil
	}
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
	equationParagraphs := map[string]bool{}
	for _, e := range equations {
		equationParagraphs[e.ParagraphID] = true
	}
	roots := map[string]map[string]*nativeXMLNode{}
	lookup := func(part, path string) (*nativeXMLNode, error) {
		if roots[part] == nil {
			raw, ok := pkg.files[part]
			if !ok {
				return nil, nil
			}
			root, e := parseNativeXML(part, raw)
			if e != nil {
				return nil, e
			}
			nodes := map[string]*nativeXMLNode{}
			var visit func(*nativeXMLNode)
			visit = func(n *nativeXMLNode) {
				nodes[n.Path] = n
				for _, c := range n.Children {
					visit(c)
				}
			}
			visit(root)
			roots[part] = nodes
		}
		return roots[part][path], nil
	}
	add := func(origin, id, code, scope, part, path string) error {
		n, e := lookup(part, path)
		if e != nil {
			return e
		}
		if n == nil || n.parent == nil || n.Name.Space != ns || len(directNativeChildren(n.parent, ns, n.Name.Local)) != 1 {
			return nil
		}
		notice := NativeEquationContextNoticeV1{PackageSHA256: doc.Source.PackageSHA256, PartSHA256: nativeSHA(pkg.files[part]), Anchor: NativeSourceAnchorV1{PartName: part, Path: n.Path, StartByte: nativeInt64(n.Start), EndByte: nativeInt64(n.End), XMLSHA256: nativeSHA(pkg.files[part][n.Start:n.End])}, DiagnosticOrigin: origin, DiagnosticID: id, Code: code, ScopeID: scope}
		val := xml.Name{Space: ns, Local: "val"}
		switch {
		case origin == "document" && code == "UNMODELED_SECTION_PROPERTY" && n.Name.Local == "textDirection" && part == doc.Source.MainPart:
			if n.parent.Name != (xml.Name{Space: ns, Local: "sectPr"}) || !nativeExactContainer(n.parent) || !nativeExactLeaf(n, val) {
				return nil
			}
			v, ok := nativeAttr(n, ns, "val")
			if !ok || v != "lrTb" {
				return nil
			}
			notice.Kind = "horizontal-section"
			notice.Value = v
		case origin == "resolved" && code == "UNMODELED_FONT_METADATA" && scope == doc.DocumentID && resolved.SourceParts.FontTablePart != nil && part == *resolved.SourceParts.FontTablePart && n.Name.Local == "charset":
			font := n.parent
			if font.Name != (xml.Name{Space: ns, Local: "font"}) || font.parent == nil || font.parent.Name != (xml.Name{Space: ns, Local: "fonts"}) || font.parent.parent != nil || !nativeExactContainer(font.parent) || !nativeExactContainer(font, xml.Name{Space: ns, Local: "name"}) || !nativeExactLeaf(n, val, xml.Name{Space: ns, Local: "characterSet"}) {
				return nil
			}
			v, ok := nativeAttr(n, ns, "val")
			c, cok := nativeAttr(n, ns, "characterSet")
			if !ok || !cok || !(v == "00" && c == "windows-1252" || v == "80" && c == "utf-8") {
				return nil
			}
			notice.Kind = "ignored-font-matching"
			notice.Value = v
			notice.CharacterSet = c
		case origin == "resolved" && code == "UNMODELED_PARAGRAPH_PROPERTY" && equationParagraphs[scope] && resolved.SourceParts.StylesPart != nil && part == *resolved.SourceParts.StylesPart:
			owner := n.parent
			style := owner.parent
			if owner.Name != (xml.Name{Space: ns, Local: "pPr"}) || !nativeExactContainer(owner) || style == nil || style.Name != (xml.Name{Space: ns, Local: "style"}) || style.parent == nil || style.parent.Name != (xml.Name{Space: ns, Local: "styles"}) || style.parent.parent != nil || len(directNativeChildren(style, ns, "pPr")) != 1 || !nativeExactContainer(style, xml.Name{Space: ns, Local: "type"}, xml.Name{Space: ns, Local: "styleId"}, xml.Name{Space: ns, Local: "default"}, xml.Name{Space: ns, Local: "customStyle"}) {
				return nil
			}
			switch n.Name.Local {
			case "suppressAutoHyphens":
				v, ok := nativeAttr(n, ns, "val")
				if !ok || v != "true" || !nativeExactLeaf(n, val) {
					return nil
				}
				notice.Kind = "disabled-paragraph-hyphenation"
				notice.Value = v
			case "tabs":
				if !nativeExactInactiveTabCandidates(n, ns) {
					return nil
				}
				notice.Kind = "unused-paragraph-tab-stops"
				for _, t := range n.Children {
					kind, _ := nativeAttr(t, ns, "val")
					position, _ := nativeNonnegativeInt64Attr(t, ns, "pos")
					notice.TabStops = append(notice.TabStops, NativeEquationContextTabV1{Kind: kind, PositionTwips: position, Leader: "none"})
				}
			default:
				return nil
			}
		default:
			return nil
		}
		out = append(out, notice)
		if len(out) > 1000 {
			return fmt.Errorf("equation context exceeds 1000 notices")
		}
		return nil
	}
	for _, d := range doc.Unsupported {
		if d.Anchor != nil && d.Preservation == "refuse-mutation" {
			if e := add("document", d.ID, d.Code, d.ScopeID, d.Anchor.PartName, d.Anchor.Path); e != nil {
				return nil, e
			}
		}
	}
	for _, d := range resolved.Diagnostics {
		if d.PartName != nil && d.Path != nil && d.Severity == "unsupported" && d.Preservation == "preserve-verbatim" {
			if e := add("resolved", "", d.Code, d.ScopeID, *d.PartName, *d.Path); e != nil {
				return nil, e
			}
		}
	}
	return out, nil
}
