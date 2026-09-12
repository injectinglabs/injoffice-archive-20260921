package docxpatch

import (
	"encoding/xml"
	"fmt"
)

func nativePartialMathNamespace(wordNS string) string {
	if wordNS == wordMLStrict {
		return "http://purl.oclc.org/ooxml/officeDocument/math"
	}
	return "http://schemas.openxmlformats.org/officeDocument/2006/math"
}
func nativeDirectMathRoot(n *nativeXMLNode, wordNS string) bool {
	return n.parent != nil && n.parent.Name == (xml.Name{Space: wordNS, Local: "p"}) && n.Name.Space == nativePartialMathNamespace(wordNS) && (n.Name.Local == "oMath" || n.Name.Local == "oMathPara")
}

type NativePartialMathNodeV1 struct {
	Kind     string                    `json:"kind"`
	Text     string                    `json:"text,omitempty"`
	Children []NativePartialMathNodeV1 `json:"children,omitempty"`
}
type NativePartialEquationV1 struct {
	PackageSHA256 string                   `json:"package_sha256"`
	ParagraphID   string                   `json:"paragraph_id"`
	Anchor        NativeSourceAnchorV1     `json:"anchor"`
	DiagnosticID  string                   `json:"diagnostic_id"`
	Status        string                   `json:"status"`
	Tree          *NativePartialMathNodeV1 `json:"tree,omitempty"`
	Reason        string                   `json:"reason,omitempty"`
}

// No math properties are discarded to manufacture support. Unknown shapes,
// formatting, tracked changes and foreign markup omit the entire equation.
func nativePartialMathTree(n *nativeXMLNode, ns string, depth int, nodes, units *int) (*NativePartialMathNodeV1, bool) {
	*nodes++
	if depth > 32 || *nodes > 256 || n.Name.Space != ns {
		return nil, false
	}
	if n.Name.Local == "t" {
		if len(n.Children) != 0 || n.Text == "" {
			return nil, false
		}
		for _, a := range n.Attrs {
			if a.Name.Space == "xmlns" || a.Name.Local == "xmlns" && a.Name.Space == "" {
				continue
			}
			if a.Name != (xml.Name{Space: "http://www.w3.org/XML/1998/namespace", Local: "space"}) || a.Value != "preserve" {
				return nil, false
			}
		}
		*units += len(n.Text)
		if *units > 32768 {
			return nil, false
		}
		return &NativePartialMathNodeV1{Kind: "text", Text: n.Text}, true
	}
	if !nativeExactContainer(n) {
		return nil, false
	}
	result := &NativePartialMathNodeV1{Kind: "row"}
	children := n.Children
	switch n.Name.Local {
	case "oMath", "oMathPara", "r", "e", "num", "den", "sup", "sub", "deg":
		if len(children) == 0 {
			return nil, false
		}
		if n.Name.Local == "r" {
			for _, c := range children {
				if c.Name != (xml.Name{Space: ns, Local: "t"}) {
					return nil, false
				}
			}
		}
		if n.Name.Local == "oMathPara" {
			for _, c := range children {
				if c.Name != (xml.Name{Space: ns, Local: "oMath"}) {
					return nil, false
				}
			}
		}
		if n.Name.Local != "r" && n.Name.Local != "oMathPara" {
			for _, c := range children {
				if c.Name.Space != ns {
					return nil, false
				}
				switch c.Name.Local {
				case "r", "f", "sSup", "sSub", "rad":
				default:
					return nil, false
				}
			}
		}
	case "f", "sSup", "sSub":
		want := []string{"num", "den"}
		result.Kind = "fraction"
		if n.Name.Local == "sSup" {
			want = []string{"e", "sup"}
			result.Kind = "superscript"
		}
		if n.Name.Local == "sSub" {
			want = []string{"e", "sub"}
			result.Kind = "subscript"
		}
		if len(children) != 2 {
			return nil, false
		}
		for i, c := range children {
			if c.Name != (xml.Name{Space: ns, Local: want[i]}) {
				return nil, false
			}
		}
	case "rad":
		// Property-free indexed radicals preserve both nonempty operands. MathML
		// mroot takes the radicand before the index, opposite the OMML order.
		if len(children) == 2 && children[0].Name == (xml.Name{Space: ns, Local: "deg"}) && children[1].Name == (xml.Name{Space: ns, Local: "e"}) {
			children = []*nativeXMLNode{children[1], children[0]}
			result.Kind = "indexed-radical"
			break
		}
		if len(children) != 3 || children[0].Name != (xml.Name{Space: ns, Local: "radPr"}) || children[1].Name != (xml.Name{Space: ns, Local: "deg"}) || children[2].Name != (xml.Name{Space: ns, Local: "e"}) {
			return nil, false
		}
		pr, deg := children[0], children[1]
		if !nativeExactContainer(pr) || len(pr.Children) != 1 || !nativeExactLeaf(deg) {
			return nil, false
		}
		hide := pr.Children[0]
		v, ok := nativeAttr(hide, ns, "val")
		if hide.Name != (xml.Name{Space: ns, Local: "degHide"}) || !nativeExactLeaf(hide, xml.Name{Space: ns, Local: "val"}) || !ok || (v != "1" && v != "true") {
			return nil, false
		}
		children = children[2:]
		result.Kind = "radical"
	default:
		return nil, false
	}
	for _, c := range children {
		item, ok := nativePartialMathTree(c, ns, depth+1, nodes, units)
		if !ok {
			return nil, false
		}
		result.Children = append(result.Children, *item)
	}
	return result, true
}

func inspectNativePartialEquations(data []byte, document *NativeDocumentV1) ([]NativePartialEquationV1, error) {
	pkg, err := openNativeDOCXPackage(data)
	if err != nil {
		return nil, err
	}
	main, strict, err := pkg.officeDocumentPart()
	if err != nil {
		return nil, err
	}
	root, err := parseNativeXML(main, pkg.files[main])
	if err != nil {
		return nil, err
	}
	wordNS := wordMLTransitional
	if strict {
		wordNS = wordMLStrict
	}
	nodesByPath := map[string]*nativeXMLNode{}
	var visit func(*nativeXMLNode)
	visit = func(n *nativeXMLNode) {
		if nativeDirectMathRoot(n, wordNS) {
			nodesByPath[n.Path] = n
			return
		}
		for _, c := range n.Children {
			visit(c)
		}
	}
	visit(root)
	out := []NativePartialEquationV1{}
	for _, d := range document.Unsupported {
		if d.Code != "UNMODELED_PARAGRAPH_CONTENT" || d.Anchor == nil || d.Anchor.PartName != main {
			continue
		}
		n := nodesByPath[d.Anchor.Path]
		if n == nil {
			continue
		}
		fact := NativePartialEquationV1{PackageSHA256: document.Source.PackageSHA256, ParagraphID: d.ScopeID, Anchor: *d.Anchor, DiagnosticID: d.ID, Status: "omitted", Reason: "Unsupported equation structure or preview budget"}
		count, units := 0, 0
		if len(out) < 100 {
			if tree, ok := nativePartialMathTree(n, nativePartialMathNamespace(wordNS), 0, &count, &units); ok {
				fact.Status = "supported"
				fact.Tree = tree
				fact.Reason = ""
			}
		}
		out = append(out, fact)
		if len(out) > 1000 {
			return nil, fmt.Errorf("partial equation inventory exceeds 1000 entries")
		}
	}
	return out, nil
}
