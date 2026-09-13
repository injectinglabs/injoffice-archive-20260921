package docxpatch

import (
	"encoding/xml"
	"strconv"
)

// Source coordinates only: selecting the owning physical page still requires
// body pagination. This evidence never discharges the drawing diagnostic.
type NativeTextboxPageAnchorV1 struct {
	Policy           string               `json:"policy"`
	SourceAnchor     NativeSourceAnchorV1 `json:"source_anchor"`
	HorizontalAnchor NativeSourceAnchorV1 `json:"horizontal_anchor"`
	VerticalAnchor   NativeSourceAnchorV1 `json:"vertical_anchor"`
	XEMU             int64                `json:"x_emu"`
	YEMU             int64                `json:"y_emu"`
}

func nativeTextboxPageOffsets(n *nativeXMLNode, wp, a string) (int64, int64, bool) {
	values := map[string]string{}
	for _, name := range []string{"distT", "distB", "distL", "distR", "simplePos", "relativeHeight", "behindDoc", "locked"} {
		values[name] = "0"
	}
	values["layoutInCell"], values["allowOverlap"] = "1", "1"
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
		v, relative, ok := nativeDrawingPosition(p, wp)
		if !ok || relative != "page" || v < 0 || v > 127000000 || v%127 != 0 || strconv.FormatInt(v, 10) != p.Children[0].Text {
			return 0, 0, false
		}
		offsets[i] = v
	}
	return offsets[0], offsets[1], true
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
	return &NativeTextboxPageAnchorV1{"page-offset-no-wrap-v1", anchor(n), anchor(n.Children[1].Children[0]), anchor(n.Children[2].Children[0]), x, y}
}
