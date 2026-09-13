package docxpatch

import (
	"encoding/xml"
	"regexp"
)

const nativeTextboxWrapPolicy = "ascii-space-greedy-v1"

var nativeTextboxWrapWords = regexp.MustCompile(`^[A-Za-z0-9]+( [A-Za-z0-9]+)*$`)

const nativeTextboxPunctuationPolicy = "ascii-punctuation-space-greedy-v1"

var nativeTextboxSentenceWords = regexp.MustCompile(`^\(?[A-Za-z0-9]+([.'/-][A-Za-z0-9]+)*\)?[,.!?;:]?( \(?[A-Za-z0-9]+([.'/-][A-Za-z0-9]+)*\)?[,.!?;:]?)*$`)

func nativeTextboxTextWrapPolicy(text string) string {
	if len(text) == 0 || len(text) > 4096 {
		return ""
	}
	if nativeTextboxWrapWords.MatchString(text) {
		return nativeTextboxWrapPolicy
	}
	if !nativeTextboxSentenceWords.MatchString(text) {
		return ""
	}
	depth := 0
	for _, c := range text {
		if c == '(' {
			depth++
			if depth > 1 {
				return ""
			}
		}
		if c == ')' {
			depth--
			if depth < 0 {
				return ""
			}
		}
	}
	if depth != 0 {
		return ""
	}
	return nativeTextboxPunctuationPolicy
}

type NativeTextboxWrapLayoutV1 struct {
	Policy               string               `json:"policy"`
	BodyPropertiesAnchor NativeSourceAnchorV1 `json:"body_properties_anchor"`
	ParagraphAnchor      NativeSourceAnchorV1 `json:"paragraph_anchor"`
	RunAnchor            NativeSourceAnchorV1 `json:"run_anchor"`
	TextAnchor           NativeSourceAnchorV1 `json:"text_anchor"`
	SpacingAnchor        NativeSourceAnchorV1 `json:"spacing_anchor"`
	LineStepTwips        int64                `json:"line_step_twips"`
}

func nativeTextboxWrapText(p *nativeXMLNode, ns string) (string, int64, bool) {
	r, spacing := p.Children[1], p.Children[0].Children[0]
	if !nativeGeometryChildren(r, ns, "rPr", "t") || !nativeExactLeaf(spacing, xml.Name{Space: ns, Local: "before"}, xml.Name{Space: ns, Local: "after"}, xml.Name{Space: ns, Local: "line"}, xml.Name{Space: ns, Local: "lineRule"}) {
		return "", 0, false
	}
	for key, want := range map[string]string{"before": "0", "after": "0", "lineRule": "exact"} {
		value, ok := nativeAttr(spacing, ns, key)
		if !ok || value != want {
			return "", 0, false
		}
	}
	step, ok := nativePositiveInt64Attr(spacing, ns, "line")
	if !ok || step > 25600 {
		return "", 0, false
	}
	t := r.Children[1]
	if len(t.Children) != 0 || len(t.Text) == 0 || len(t.Text) > 4096 || nativeTextboxTextWrapPolicy(t.Text) == "" {
		return "", 0, false
	}
	space, ok := nativeAttr(t, "http://www.w3.org/XML/1998/namespace", "space")
	if !ok || space != "preserve" {
		return "", 0, false
	}
	for _, attr := range t.Attrs {
		if !nativeSettingsNamespaceDeclaration(attr) && (attr.Name != (xml.Name{Space: "http://www.w3.org/XML/1998/namespace", Local: "space"}) || attr.Value != "preserve") {
			return "", 0, false
		}
	}
	return t.Text, step, true
}
func nativeTextboxWrapEvidence(drawing *nativeXMLNode, ns, part string, raw []byte) *NativeTextboxWrapLayoutV1 {
	p := nativeDescendants(drawing, ns, "txbxContent")[0].Children[0]
	body := nativeDescendants(drawing, nativeTextboxWPS, "bodyPr")[0]
	wrap, _ := nativeAttr(body, "", "wrap")
	if wrap != "square" {
		return nil
	}
	_, step, ok := nativeTextboxWrapText(p, ns)
	if !ok {
		return nil
	}
	anchor := func(n *nativeXMLNode) NativeSourceAnchorV1 {
		start, end := n.Start, n.End
		return NativeSourceAnchorV1{part, n.Path, &start, &end, nativeSHA(raw[start:end])}
	}
	r := p.Children[1]
	return &NativeTextboxWrapLayoutV1{nativeTextboxTextWrapPolicy(r.Children[1].Text), anchor(body), anchor(p), anchor(r), anchor(r.Children[1]), anchor(p.Children[0].Children[0]), step}
}
