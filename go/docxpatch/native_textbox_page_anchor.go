package docxpatch

import (
	"encoding/xml"
	"strconv"
)

// Source coordinates only: selecting the owning physical page still requires
// body pagination. This evidence never discharges the drawing diagnostic.
type NativeTextboxStackingV1 struct {
	BehindDoc      bool   `json:"behind_doc"`
	RelativeHeight uint32 `json:"relative_height"`
}

type NativeTextboxPageAnchorV1 struct {
	Stacking           *NativeTextboxStackingV1 `json:"stacking,omitempty"`
	Policy             string                   `json:"policy"`
	SourceAnchor       NativeSourceAnchorV1     `json:"source_anchor"`
	HorizontalAnchor   NativeSourceAnchorV1     `json:"horizontal_anchor"`
	VerticalAnchor     NativeSourceAnchorV1     `json:"vertical_anchor"`
	XEMU               int64                    `json:"x_emu"`
	YEMU               int64                    `json:"y_emu"`
	HorizontalRelative string                   `json:"horizontal_relative,omitempty"`
	VerticalRelative   string                   `json:"vertical_relative,omitempty"`
	HorizontalAlign    string                   `json:"horizontal_align,omitempty"`
	VerticalAlign      string                   `json:"vertical_align,omitempty"`
}

func nativeTextboxPageOffsets(n *nativeXMLNode, wp, a string) (int64, int64, bool) {
	values := map[string]string{}
	for _, name := range []string{"distT", "distB", "distL", "distR", "simplePos", "relativeHeight", "behindDoc", "locked"} {
		values[name] = "0"
	}
	values["layoutInCell"], values["allowOverlap"] = "1", "1"
	layer, ok := nativeTextboxStacking(n)
	if !ok {
		return 0, 0, false
	}
	values["relativeHeight"] = strconv.FormatUint(uint64(layer.RelativeHeight), 10)
	values["behindDoc"], _ = nativeUnqualifiedAttr(n, "behindDoc")
	if n.Name != (xml.Name{Space: wp, Local: "anchor"}) || !nativeGeometryAttrs(n, "", values) || len(n.Children) != 8 {
		return 0, 0, false
	}
	for i, name := range []string{"simplePos", "positionH", "positionV", "extent", "effectExtent", "wrapNone", "docPr"} {
		if n.Children[i].Name != (xml.Name{Space: wp, Local: name}) {
			return 0, 0, false
		}
	}
	if n.Children[7].Name != (xml.Name{Space: a, Local: "graphic"}) || !nativeGeometryAttrs(n.Children[0], "", map[string]string{"x": "0", "y": "0"}) || len(n.Children[0].Children) != 0 || !nativeExactLeaf(n.Children[5]) {
		return 0, 0, false
	}
	var offsets [2]int64
	for i, p := range n.Children[1:3] {
		v, _, _, ok := nativeTextboxPositionAxis(p, wp, i == 0)
		if !ok {
			return 0, 0, false
		}
		offsets[i] = v
	}
	return offsets[0], offsets[1], true
}

func nativeTextboxPositionAxis(p *nativeXMLNode, wp string, horizontal bool) (int64, string, string, bool) {
	relative, _ := nativeUnqualifiedAttr(p, "relativeFrom")
	switch relative {
	case "page", "margin", "insideMargin", "outsideMargin":
	case "column", "character", "leftMargin", "rightMargin":
		if !horizontal {
			return 0, "", "", false
		}
	case "paragraph", "line", "topMargin", "bottomMargin":
		if horizontal {
			return 0, "", "", false
		}
	default:
		return 0, "", "", false
	}

	if v, _, ok := nativeDrawingPosition(p, wp); ok {
		if v < -127000000 || v > 127000000 || v%127 != 0 || strconv.FormatInt(v, 10) != p.Children[0].Text {
			return 0, "", "", false
		}
		return v, relative, "", true
	}
	if len(p.Children) != 1 || !nativeExactContainer(p, xml.Name{Local: "relativeFrom"}) {
		return 0, "", "", false
	}
	c := p.Children[0]
	if c.Name != (xml.Name{Space: wp, Local: "align"}) || len(c.Attrs) != 0 || len(c.Children) != 0 {
		return 0, "", "", false
	}
	if relative == "character" || relative == "paragraph" || relative == "line" {
		return 0, "", "", false
	}
	switch c.Text {
	case "center", "inside", "outside":
	case "left", "right":
		if !horizontal {
			return 0, "", "", false
		}
	case "top", "bottom":
		if horizontal {
			return 0, "", "", false
		}
	default:
		return 0, "", "", false
	}

	return 0, relative, c.Text, true
}

func nativeTextboxPageAnchorEvidence(drawing *nativeXMLNode, ns, part string, raw []byte) *NativeTextboxPageAnchorV1 {
	wp, a := wordDrawingTransitional, drawingMLTransitional
	if ns == wordMLStrict {
		wp, a = wordDrawingStrict, drawingMLStrict
	}
	n := drawing.Children[0]
	x, y, ok := nativeTextboxPageOffsets(n, wp, a)
	if !ok {
		return nil
	}
	anchor := func(n *nativeXMLNode) NativeSourceAnchorV1 { return nativeTextboxSourceAnchor(n, part, raw) }
	result := &NativeTextboxPageAnchorV1{Policy: "page-offset-no-wrap-v1", SourceAnchor: anchor(n), HorizontalAnchor: anchor(n.Children[1].Children[0]), VerticalAnchor: anchor(n.Children[2].Children[0]), XEMU: x, YEMU: y}
	_, h, ha, _ := nativeTextboxPositionAxis(n.Children[1], wp, true)
	_, v, va, _ := nativeTextboxPositionAxis(n.Children[2], wp, false)
	layer, _ := nativeTextboxStacking(n)
	if layer.BehindDoc || layer.RelativeHeight != 0 {
		result.Stacking = &layer
	}
	if h != "page" || v != "page" || ha != "" || va != "" || x < 0 || y < 0 || result.Stacking != nil {
		result.Policy = "relative-position-no-wrap-v2"
		result.HorizontalRelative = h
		result.VerticalRelative = v
		result.HorizontalAlign = ha
		result.VerticalAlign = va
	}
	return result
}

func nativeTextboxStacking(n *nativeXMLNode) (NativeTextboxStackingV1, bool) {
	height, _ := nativeUnqualifiedAttr(n, "relativeHeight")
	behind, _ := nativeUnqualifiedAttr(n, "behindDoc")
	rank, err := strconv.ParseUint(height, 10, 32)
	if err != nil || strconv.FormatUint(rank, 10) != height || (behind != "0" && behind != "1" && behind != "false" && behind != "true") {
		return NativeTextboxStackingV1{}, false
	}
	return NativeTextboxStackingV1{behind == "1" || behind == "true", uint32(rank)}, true
}
