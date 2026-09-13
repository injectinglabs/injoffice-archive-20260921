package docxpatch

import (
	"encoding/xml"
	"fmt"
	"strings"
	"unicode/utf16"
)

const nativeTextboxWPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
const nativeTextboxVML = "urn:schemas-microsoft-com:vml"

type NativePartialTextboxV1 struct {
	PackageSHA256 string               `json:"package_sha256"`
	PartSHA256    string               `json:"part_sha256"`
	ParagraphID   string               `json:"paragraph_id"`
	DiagnosticID  string               `json:"diagnostic_id"`
	Anchor        NativeSourceAnchorV1 `json:"anchor"`
	Kind          string               `json:"kind"`
	Status        string               `json:"status"`
	Paragraphs    []string             `json:"paragraphs"`
	Reason        string               `json:"reason"`
}
type NativePartialTextboxesV1 struct {
	Items        []NativePartialTextboxV1 `json:"items"`
	OmittedCount int                      `json:"omitted_count"`
}

// This sidecar never changes native extraction, editing authority or pagination.
// Only an unambiguous direct DrawingML/VML textbox in a body paragraph is read.
// TextBoxContent can contain revisions, tables and other stories: those are not
// flattened. See Microsoft Open XML SDK Wordprocessing.TextBoxContent and
// Office2010.Word.DrawingShape.TextBoxInfo2 / Vml.TextBox documentation.
func inspectNativePartialTextboxes(data []byte, doc *NativeDocumentV1) (*NativePartialTextboxesV1, error) {
	resolver, err := newNativeLayoutResolver(data, NativeExtractionOptions{})
	if err != nil {
		return nil, err
	}
	ns, main := resolver.wordNS, doc.Source.MainPart
	raw := resolver.pkg.files[main]
	stylesOK := nativePartialTextboxStyles(resolver)
	partSHA := nativeSHA(raw)
	drawings := map[string][]NativeUnsupportedCapabilityV1{}
	contextQualified := nativeExactContainer(resolver.mainRoot)
	for _, diagnostic := range doc.Unsupported {
		if diagnostic.Capability == "drawings" {
			drawings[diagnostic.ScopeID] = append(drawings[diagnostic.ScopeID], diagnostic)
		} else if diagnostic.Code != "DEFAULT_SECTION_INFERRED" {
			contextQualified = false
		}
	}
	out := &NativePartialTextboxesV1{Items: []NativePartialTextboxV1{}}
	nodes := map[string]*nativeXMLNode{}
	var visit func(*nativeXMLNode)
	visit = func(n *nativeXMLNode) {
		nodes[n.Path] = n
		for _, c := range n.Children {
			visit(c)
		}
	}
	visit(resolver.mainRoot)
	total := 0
	for _, b := range doc.Body.Blocks {
		if b.Paragraph == nil {
			continue
		}
		p := b.Paragraph
		owner := nodes[p.Anchor.Path]
		if owner == nil {
			continue
		}
		for _, d := range drawings[p.ID] {
			if d.Capability != "drawings" || d.ScopeID != p.ID || d.Anchor == nil || d.Anchor.PartName != main {
				continue
			}
			n := nodes[d.Anchor.Path]
			if n == nil || nativeSHA(raw[n.Start:n.End]) != d.Anchor.XMLSHA256 {
				continue
			}
			root := n
			for depth := 0; root != nil && root.Name != (xml.Name{Space: ns, Local: "drawing"}) && root.Name != (xml.Name{Space: ns, Local: "pict"}) && depth < 4; depth++ {
				root = root.parent
			}
			if root == nil || root.Name.Space != ns || (root.Name.Local != "drawing" && root.Name.Local != "pict") || root.parent == nil || root.parent.Name != (xml.Name{Space: ns, Local: "r"}) || root.parent.parent != owner {
				continue
			}
			if !nativePartialTextboxDiagnosticAnchor(root, n, ns) {
				continue
			}
			kind := "drawingml"
			if root.Name.Local == "pict" {
				kind = "vml"
			}
			// Presence detection never constitutes admission. In particular Choice and
			// Fallback descendants cannot become two authoritatively selected stories.
			if len(nativeDescendants(root, ns, "txbxContent")) == 0 {
				continue
			}
			if len(out.Items) >= 64 {
				out.OmittedCount++
				continue
			}
			item := NativePartialTextboxV1{doc.Source.PackageSHA256, partSHA, p.ID, d.ID, *d.Anchor, kind, "omitted", []string{}, "unsupported-textbox-structure"}
			content := nativePartialTextboxContent(root, ns)
			contextOK := contextQualified && owner.parent != nil && nativeExactContainer(owner.parent)
			if !stylesOK {
				item.Reason = "unqualified-text-visibility"
			}
			if content != nil && contextOK && stylesOK && nativePartialTextboxOwner(owner, root.parent, ns) {
				paragraphs, reason := nativePartialTextboxParagraphs(resolver, content, root, p)
				item.Reason = reason
				units := 0
				for _, s := range paragraphs {
					units += len(utf16.Encode([]rune(s)))
				}
				if reason == "" && total+units <= 100000 {
					item.Status = "supported"
					item.Paragraphs = paragraphs
					total += units
				} else if reason == "" {
					item.Reason = "text-limit"
				}
			}
			out.Items = append(out.Items, item)
		}
	}
	if len(out.Items) == 0 && out.OmittedCount == 0 {
		return nil, nil
	}
	return out, nil
}

func nativePartialTextboxOwner(p, r *nativeXMLNode, ns string) bool {
	if !nativeExactRevisionContainer(p, ns, "rsidR", "rsidRDefault", "rsidP", "rsidRPr") || !nativeExactRevisionContainer(r, ns, "rsidR", "rsidRPr") {
		return false
	}
	if len(directNativeChildren(p, ns, "pPr")) > 1 || len(directNativeChildren(r, ns, "rPr")) > 1 {
		return false
	}
	for _, c := range r.Children {
		if c.Name == (xml.Name{Space: ns, Local: "rPr"}) && !nativePartialReviewRunProperties(c, ns) {
			return false
		}
	}
	for _, c := range p.Children {
		if c.Name == (xml.Name{Space: ns, Local: "r"}) {
			continue
		}
		if c.Name != (xml.Name{Space: ns, Local: "pPr"}) || !nativeExactContainer(c) {
			return false
		}
		for _, v := range c.Children {
			if v.Name != (xml.Name{Space: ns, Local: "pStyle"}) || !nativeExactLeaf(v, xml.Name{Space: ns, Local: "val"}) {
				return false
			}
		}
	}
	return true
}
func nativePartialTextboxContent(root *nativeXMLNode, ns string) *nativeXMLNode {
	if !nativeExactContainer(root) {
		return nil
	}
	var box *nativeXMLNode
	if root.Name.Local == "pict" {
		if len(root.Children) != 1 {
			return nil
		}
		shape := root.Children[0]
		if shape.Name != (xml.Name{Space: nativeTextboxVML, Local: "shape"}) && shape.Name != (xml.Name{Space: nativeTextboxVML, Local: "rect"}) {
			return nil
		}
		if !nativeExactContainer(shape, xml.Name{Local: "id"}, xml.Name{Local: "style"}) {
			return nil
		}
		if style, ok := nativeAttr(shape, "", "style"); ok {
			for _, pair := range strings.Split(style, ";") {
				if strings.TrimSpace(pair) == "" {
					continue
				}
				kv := strings.Split(pair, ":")
				if len(kv) != 2 {
					return nil
				}
				switch strings.TrimSpace(kv[0]) {
				case "position", "left", "top", "width", "height", "margin-left", "margin-top", "z-index":
				default:
					return nil
				}
			}
		}
		if len(shape.Children) != 1 {
			return nil
		}
		box = shape.Children[0]
		if box.Name != (xml.Name{Space: nativeTextboxVML, Local: "textbox"}) || !nativeExactContainer(box, xml.Name{Local: "inset"}) {
			return nil
		}
	} else {
		wp, a := wordDrawingTransitional, drawingMLTransitional
		if ns == wordMLStrict {
			wp, a = wordDrawingStrict, drawingMLStrict
		}
		if len(root.Children) != 1 {
			return nil
		}
		container := root.Children[0]
		if container.Name.Space != wp || (container.Name.Local != "inline" && container.Name.Local != "anchor") {
			return nil
		}
		// Outer geometry is not rendered. Unknown/hidden containers stay opaque.
		allowed := []xml.Name{}
		for _, s := range []string{"distT", "distB", "distL", "distR", "simplePos", "relativeHeight", "behindDoc", "locked", "layoutInCell", "allowOverlap"} {
			allowed = append(allowed, xml.Name{Local: s})
		}
		if !nativeExactContainer(container, allowed...) {
			return nil
		}
		var graphic *nativeXMLNode
		for _, c := range container.Children {
			if c.Name == (xml.Name{Space: a, Local: "graphic"}) {
				if graphic != nil {
					return nil
				}
				graphic = c
				continue
			}
			if c.Name.Space != wp {
				return nil
			}
			switch c.Name.Local {
			case "extent", "effectExtent", "docPr", "cNvGraphicFramePr", "simplePos", "positionH", "positionV", "wrapNone", "wrapSquare", "wrapTopAndBottom":
			default:
				return nil
			}
			if !nativeTextboxGeometryNode(c, wp, a) {
				return nil
			}
		}
		if graphic == nil || !nativeExactContainer(graphic) || len(graphic.Children) != 1 {
			return nil
		}
		gd := graphic.Children[0]
		uri, _ := nativeAttr(gd, "", "uri")
		if gd.Name != (xml.Name{Space: a, Local: "graphicData"}) || !nativeExactContainer(gd, xml.Name{Local: "uri"}) || uri != nativeTextboxWPS || len(gd.Children) != 1 {
			return nil
		}
		shape := gd.Children[0]
		if shape.Name != (xml.Name{Space: nativeTextboxWPS, Local: "wsp"}) || !nativeExactContainer(shape) {
			return nil
		}
		seen := map[string]bool{}
		for _, c := range shape.Children {
			if c.Name.Space != nativeTextboxWPS || seen[c.Name.Local] {
				return nil
			}
			seen[c.Name.Local] = true
			switch c.Name.Local {
			case "txbx":
				box = c
			case "cNvPr":
				if !nativeExactLeaf(c, xml.Name{Local: "id"}, xml.Name{Local: "name"}, xml.Name{Local: "descr"}, xml.Name{Local: "title"}) {
					return nil
				}
			case "cNvSpPr":
				if !nativeExactLeaf(c, xml.Name{Local: "txBox"}) {
					return nil
				}
				if value, ok := nativeAttr(c, "", "txBox"); ok && value != "1" && value != "true" {
					return nil
				}
			case "spPr":
				if !nativeExactContainer(c) {
					return nil
				}
				for _, g := range c.Children {
					if g.Name != (xml.Name{Space: a, Local: "prstGeom"}) || !nativeExactContainer(g, xml.Name{Local: "prst"}) {
						return nil
					}
					value, _ := nativeAttr(g, "", "prst")
					if value != "rect" {
						return nil
					}
					if len(g.Children) > 1 {
						return nil
					}
					for _, v := range g.Children {
						if v.Name != (xml.Name{Space: a, Local: "avLst"}) || !nativeExactLeaf(v) {
							return nil
						}
					}
				}
				if len(c.Children) > 1 {
					return nil
				}
			case "bodyPr", "style":
				if !nativeExactLeaf(c) {
					return nil
				}
			default:
				return nil
			}
		}
		if box == nil || !nativeExactContainer(box) {
			return nil
		}
	}
	if len(box.Children) != 1 {
		return nil
	}
	content := box.Children[0]
	if content.Name != (xml.Name{Space: ns, Local: "txbxContent"}) || !nativeExactContainer(content) || len(nativeDescendants(root, ns, "txbxContent")) != 1 {
		return nil
	}
	return content
}
func nativeTextboxGeometryNode(n *nativeXMLNode, wp, a string) bool {
	if n.Name.Space != wp {
		return false
	}
	attrs := []xml.Name{}
	switch n.Name.Local {
	case "docPr":
		for _, k := range []string{"id", "name", "descr", "title"} {
			attrs = append(attrs, xml.Name{Local: k})
		}
	case "extent":
		attrs = []xml.Name{{Local: "cx"}, {Local: "cy"}}
	case "effectExtent":
		attrs = []xml.Name{{Local: "l"}, {Local: "t"}, {Local: "r"}, {Local: "b"}}
	case "simplePos":
		attrs = []xml.Name{{Local: "x"}, {Local: "y"}}
	case "wrapSquare":
		attrs = []xml.Name{{Local: "wrapText"}, {Local: "distT"}, {Local: "distB"}, {Local: "distL"}, {Local: "distR"}}
	case "wrapTopAndBottom":
		attrs = []xml.Name{{Local: "distT"}, {Local: "distB"}}
	case "wrapNone", "cNvGraphicFramePr":
	case "positionH", "positionV":
		if !nativeExactContainer(n, xml.Name{Local: "relativeFrom"}) || len(n.Children) != 1 {
			return false
		}
		c := n.Children[0]
		return c.Name.Space == wp && (c.Name.Local == "posOffset" || c.Name.Local == "align") && len(c.Children) == 0 && len(c.Attrs) == 0
	default:
		return false
	}
	return nativeExactLeaf(n, attrs...)
}

func nativePartialTextboxParagraphs(resolver *nativeLayoutResolver, content, root *nativeXMLNode, owner *NativeParagraphV1) ([]string, string) {
	localResolver := *resolver
	localResolver.diagnostics = append([]NativeResolutionDiagnosticV1{}, resolver.diagnostics...)
	localResolver.diagnosticSet = map[string]bool{}
	for key, value := range resolver.diagnosticSet {
		localResolver.diagnosticSet[key] = value
	}
	resolver = &localResolver
	ns, main := resolver.wordNS, resolver.mainPart
	raw := resolver.pkg.files[main]
	anchor := func(n *nativeXMLNode) NativeSourceAnchorV1 {
		return NativeSourceAnchorV1{PartName: main, Path: n.Path, StartByte: nativeInt64(n.Start), EndByte: nativeInt64(n.End), XMLSHA256: nativeSHA(raw[n.Start:n.End])}
	}
	if len(content.Children) == 0 || len(content.Children) > 64 {
		return nil, "paragraph-limit"
	}
	// Resolve the owning drawing run too: inherited hidden formatting on that run
	// must not leak an otherwise ordinary textbox story.
	checks := []NativeParagraphV1{{ID: owner.ID, Anchor: owner.Anchor, Runs: []NativeRunV1{{ID: "textbox-owner", Anchor: anchor(root), Kind: "text", Text: nativeString("")}}}}
	texts := []string{}
	for pi, p := range content.Children {
		if p.Name != (xml.Name{Space: ns, Local: "p"}) || !nativeExactRevisionContainer(p, ns, "rsidR", "rsidRDefault", "rsidP", "rsidRPr") {
			return nil, "unsupported-textbox-content"
		}
		if len(p.Children) > 129 {
			return nil, "text-limit"
		}
		paragraph := NativeParagraphV1{ID: fmt.Sprintf("textbox-p-%d-%d", p.Start, pi), Anchor: anchor(p)}
		text := ""
		for ri, r := range p.Children {
			if r.Name == (xml.Name{Space: ns, Local: "pPr"}) && ri == 0 {
				if !nativeExactContainer(r) {
					return nil, "unqualified-text-visibility"
				}
				for _, c := range r.Children {
					if c.Name != (xml.Name{Space: ns, Local: "pStyle"}) || !nativeExactLeaf(c, xml.Name{Space: ns, Local: "val"}) {
						return nil, "unqualified-text-visibility"
					}
				}
				continue
			}
			if r.Name != (xml.Name{Space: ns, Local: "r"}) || !nativeExactRevisionContainer(r, ns, "rsidR", "rsidRPr") {
				return nil, "unsupported-textbox-content"
			}
			if len(r.Children) > 129 {
				return nil, "text-limit"
			}
			for ti, t := range r.Children {
				if t.Name == (xml.Name{Space: ns, Local: "rPr"}) && ti == 0 {
					if !nativePartialReviewRunProperties(t, ns) {
						return nil, "unqualified-text-visibility"
					}
					continue
				}
				if t.Name != (xml.Name{Space: ns, Local: "t"}) || len(t.Children) != 0 {
					return nil, "unsupported-textbox-content"
				}
				for _, a := range t.Attrs {
					if nativeSettingsNamespaceDeclaration(a) {
						continue
					}
					if a.Name != (xml.Name{Space: "http://www.w3.org/XML/1998/namespace", Local: "space"}) || (a.Value != "preserve" && a.Value != "default") {
						return nil, "unsupported-textbox-content"
					}
				}
				text += t.Text
				paragraph.Runs = append(paragraph.Runs, NativeRunV1{ID: fmt.Sprintf("textbox-t-%d", t.Start), Anchor: anchor(t), Kind: "text", Text: nativeString(t.Text)})
			}
		}
		if len(utf16.Encode([]rune(text))) > 4096 {
			return nil, "text-limit"
		}
		texts = append(texts, text)
		checks = append(checks, paragraph)
	}
	// Conservative: any source/style-resolution diagnostic refuses this separate
	// text inventory. Rendering-neutral exceptions can be added with own evidence.
	if len(resolver.diagnostics) > 0 {
		return nil, "unqualified-text-visibility"
	}
	result := &NativeResolvedLayoutInputV1{}
	for i := range checks {
		resolver.resolveParagraph(&checks[i], result, newNativeNumberingState(), nil)
	}
	if len(resolver.diagnostics) > 0 || resolver.diagnosticOverflow {
		return nil, "unqualified-text-visibility"
	}
	for _, r := range result.Runs {
		if r.Properties.Hidden != nil && *r.Properties.Hidden {
			return nil, "hidden-text"
		}
	}
	return texts, ""
}

// The ordinary style loader intentionally preserves unmodeled source. Absence
// of its diagnostics is not proof that every inheritance container was read.
// Textbox visibility additionally requires this closed, source-level grammar.
func nativePartialTextboxStyles(resolver *nativeLayoutResolver) bool {
	if resolver.parts.StylesPart == nil {
		return true
	}
	root, err := parseNativeXML(*resolver.parts.StylesPart, resolver.pkg.files[*resolver.parts.StylesPart])
	if err != nil {
		return false
	}
	ns := resolver.wordNS
	if root.Name != (xml.Name{Space: ns, Local: "styles"}) || !nativeExactContainer(root) || len(root.Children) > NativeDOCXMaxCollectionItems {
		return false
	}
	defaultsSeen := false
	runProperties := func(n *nativeXMLNode) bool {
		if !nativePartialReviewRunProperties(n, ns) {
			return false
		}
		// rStyle is a reference in an ordinary run, not an alternate inheritance
		// layer which the default/style rPr loader resolves.
		return len(directNativeChildren(n, ns, "rStyle")) == 0
	}
	for _, child := range root.Children {
		if child.Name.Space != ns {
			return false
		}
		switch child.Name.Local {
		case "docDefaults":
			if defaultsSeen || !nativeExactContainer(child) {
				return false
			}
			defaultsSeen = true
			seen := map[string]bool{}
			for _, container := range child.Children {
				if container.Name.Space != ns || seen[container.Name.Local] || !nativeExactContainer(container) || len(container.Children) > 1 {
					return false
				}
				seen[container.Name.Local] = true
				expected := ""
				switch container.Name.Local {
				case "rPrDefault":
					expected = "rPr"
				case "pPrDefault":
					expected = "pPr"
				default:
					return false
				}
				for _, properties := range container.Children {
					if properties.Name != (xml.Name{Space: ns, Local: expected}) {
						return false
					}
					if expected == "rPr" {
						if !runProperties(properties) {
							return false
						}
					} else if !nativeExactLeaf(properties) {
						return false
					}
				}
			}
		case "style":
			if !nativeExactContainer(child, xml.Name{Space: ns, Local: "type"}, xml.Name{Space: ns, Local: "styleId"}, xml.Name{Space: ns, Local: "default"}, xml.Name{Space: ns, Local: "customStyle"}) {
				return false
			}
			kind, _ := nativeAttr(child, ns, "type")
			if kind != "paragraph" && kind != "character" {
				return false
			}
			if flag, present := nativeAttr(child, ns, "customStyle"); present {
				if _, valid := nativeLexicalOnOff(flag); !valid {
					return false
				}
			}
			seen := map[string]bool{}
			for _, properties := range child.Children {
				if properties.Name.Space != ns || seen[properties.Name.Local] {
					return false
				}
				seen[properties.Name.Local] = true
				switch properties.Name.Local {
				case "name", "basedOn":
					if !nativeExactLeaf(properties, xml.Name{Space: ns, Local: "val"}) {
						return false
					}
					if value, present := nativeAttr(properties, ns, "val"); !present || (properties.Name.Local == "basedOn" && !nativeIDPattern.MatchString(value)) {
						return false
					}
				case "rPr":
					if !runProperties(properties) {
						return false
					}
				case "pPr":
					if !nativeExactLeaf(properties) {
						return false
					}
				default:
					return false
				}
			}
		default:
			return false
		}
	}
	return true
}

// Match the consumer's closed diagnostic boundary grammar. Refusals anchored
// on individual extent/property leaves stay in the native omission inventory;
// they must not create a sidecar the consumer cannot source-join.
func nativePartialTextboxDiagnosticAnchor(root, node *nativeXMLNode, ns string) bool {
	if node == root {
		return true
	}
	if root.Name != (xml.Name{Space: ns, Local: "drawing"}) {
		return false
	}
	wp, a := wordDrawingTransitional, drawingMLTransitional
	if ns == wordMLStrict {
		wp, a = wordDrawingStrict, drawingMLStrict
	}
	container := node
	if node.Name == (xml.Name{Space: a, Local: "graphicData"}) {
		container = node.parent
		if container == nil || container.Name != (xml.Name{Space: a, Local: "graphic"}) {
			return false
		}
		container = container.parent
	} else if node.Name == (xml.Name{Space: a, Local: "graphic"}) {
		container = node.parent
	}
	return container != nil && container.parent == root && container.Name.Space == wp && (container.Name.Local == "inline" || container.Name.Local == "anchor")
}
