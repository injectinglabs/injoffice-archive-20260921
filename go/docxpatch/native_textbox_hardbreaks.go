package docxpatch

import (
	"encoding/xml"
	"strings"
)

type NativeTextboxHardBreakLineV1 struct {
	Ordinal           int                   `json:"ordinal"`
	TextAnchor        NativeSourceAnchorV1  `json:"text_anchor"`
	BreakBeforeAnchor *NativeSourceAnchorV1 `json:"break_before_anchor"`
	StartUTF16        int                   `json:"start_utf16"`
	EndUTF16          int                   `json:"end_utf16"`
}
type NativeTextboxHardBreakLayoutV1 struct {
	ParagraphAnchor NativeSourceAnchorV1           `json:"paragraph_anchor"`
	RunAnchor       NativeSourceAnchorV1           `json:"run_anchor"`
	SpacingAnchor   NativeSourceAnchorV1           `json:"spacing_anchor"`
	LineStepTwips   int64                          `json:"line_step_twips"`
	Lines           []NativeTextboxHardBreakLineV1 `json:"lines"`
}

// Existing geometry grammar already qualifies the one paragraph and run.
func nativeTextboxHardBreakText(p *nativeXMLNode, ns string) (string, int64, bool) {
	r, spacing := p.Children[1], p.Children[0].Children[0]
	if len(r.Children) < 2 || len(r.Children) > 32 || len(r.Children)%2 != 0 {
		return "", 0, false
	}
	step := int64(0)
	if len(r.Children) == 2 {
		if !nativeGeometryAttrs(spacing, ns, map[string]string{"before": "0", "after": "0", "line": "240", "lineRule": "auto"}) || len(spacing.Children) != 0 {
			return "", 0, false
		}
	} else {
		if !nativeExactLeaf(spacing, xml.Name{Space: ns, Local: "before"}, xml.Name{Space: ns, Local: "after"}, xml.Name{Space: ns, Local: "line"}, xml.Name{Space: ns, Local: "lineRule"}) {
			return "", 0, false
		}
		for k, v := range map[string]string{"before": "0", "after": "0", "lineRule": "exact"} {
			got, ok := nativeAttr(spacing, ns, k)
			if !ok || got != v {
				return "", 0, false
			}
		}
		var ok bool
		step, ok = nativePositiveInt64Attr(spacing, ns, "line")
		if !ok || step > 25600 {
			return "", 0, false
		}
	}
	texts := []string{}
	for i, n := range r.Children[1:] {
		if i%2 == 1 {
			if n.Name != (xml.Name{Space: ns, Local: "br"}) || !nativeGeometryAttrs(n, ns, map[string]string{"type": "textWrapping", "clear": "none"}) || len(n.Children) != 0 {
				return "", 0, false
			}
			continue
		}
		if n.Name != (xml.Name{Space: ns, Local: "t"}) || len(n.Children) != 0 {
			return "", 0, false
		}
		space, ok := nativeAttr(n, "http://www.w3.org/XML/1998/namespace", "space")
		if !ok || space != "preserve" {
			return "", 0, false
		}
		for _, attr := range n.Attrs {
			if !nativeSettingsNamespaceDeclaration(attr) && (attr.Name != (xml.Name{Space: "http://www.w3.org/XML/1998/namespace", Local: "space"}) || attr.Value != "preserve") {
				return "", 0, false
			}
		}
		if len(n.Text) == 0 || len(n.Text) > 4096 {
			return "", 0, false
		}
		for _, c := range n.Text {
			if c < 32 || c > 126 {
				return "", 0, false
			}
		}
		texts = append(texts, n.Text)
	}
	text := strings.Join(texts, "\n")
	if len(text) > 4096 {
		return "", 0, false
	}
	return text, step, true
}
func nativeTextboxHardBreakEvidence(drawing *nativeXMLNode, ns, part string, raw []byte) *NativeTextboxHardBreakLayoutV1 {
	p := nativeDescendants(drawing, ns, "txbxContent")[0].Children[0]
	_, step, ok := nativeTextboxHardBreakText(p, ns)
	if !ok || step == 0 {
		return nil
	}
	anchor := func(n *nativeXMLNode) NativeSourceAnchorV1 {
		start, end := n.Start, n.End
		return NativeSourceAnchorV1{part, n.Path, &start, &end, nativeSHA(raw[start:end])}
	}
	r := p.Children[1]
	out := &NativeTextboxHardBreakLayoutV1{ParagraphAnchor: anchor(p), RunAnchor: anchor(r), SpacingAnchor: anchor(p.Children[0].Children[0]), LineStepTwips: step, Lines: []NativeTextboxHardBreakLineV1{}}
	offset := 0
	for i := 1; i < len(r.Children); i += 2 {
		n := r.Children[i]
		var before *NativeSourceAnchorV1
		if i > 1 {
			a := anchor(r.Children[i-1])
			before = &a
			offset++
		}
		out.Lines = append(out.Lines, NativeTextboxHardBreakLineV1{len(out.Lines), anchor(n), before, offset, offset + len(n.Text)})
		offset += len(n.Text)
	}
	return out
}
