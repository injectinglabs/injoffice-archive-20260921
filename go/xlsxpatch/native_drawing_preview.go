package xlsxpatch

import (
	"encoding/xml"
	"fmt"
	"path"
	"regexp"
	"strconv"
	"strings"
)

var previewCreationGUID = regexp.MustCompile(`^\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}$`)

func previewDrawingCreationID(n *previewXML, a string) bool {
	if !previewAttrs(n) || len(n.children) != 1 {
		return false
	}
	ext := previewChildNS(n, a, "ext")
	if !previewAttrs(ext, xml.Name{Local: "uri"}) || ext.attr("uri") != "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" || len(ext.children) != 1 {
		return false
	}
	id := previewChildNS(ext, "http://schemas.microsoft.com/office/drawing/2014/main", "creationId")
	return previewAttrs(id, xml.Name{Local: "id"}) && len(id.children) == 0 && previewCreationGUID.MatchString(id.attr("id"))
}

type NativeDrawingMarkerV1 struct {
	Column       int   `json:"column"`
	Row          int   `json:"row"`
	ColumnOffset int64 `json:"column_offset_emu"`
	RowOffset    int64 `json:"row_offset_emu"`
}
type NativeDrawingAnchorV1 struct {
	Kind   string                 `json:"kind"`
	From   NativeDrawingMarkerV1  `json:"from"`
	To     *NativeDrawingMarkerV1 `json:"to,omitempty"`
	Width  int64                  `json:"width_emu,omitempty"`
	Height int64                  `json:"height_emu,omitempty"`
}
type NativeDrawingObjectV1 struct {
	SheetID     string                 `json:"sheet_id"`
	SheetPart   string                 `json:"sheet_part"`
	DrawingPart string                 `json:"drawing_part"`
	Ordinal     int                    `json:"ordinal"`
	Kind        string                 `json:"kind"`
	ChartPart   string                 `json:"chart_part,omitempty"`
	Anchor      *NativeDrawingAnchorV1 `json:"anchor,omitempty"`
	Warnings    []string               `json:"warnings"`
}

func previewChildNS(n *previewXML, ns, name string) *previewXML {
	if n == nil {
		return nil
	}
	var found *previewXML
	for _, c := range n.children {
		if c.name == (xml.Name{Space: ns, Local: name}) {
			if found != nil {
				return nil
			}
			found = c
		}
	}
	return found
}
func previewAttrs(n *previewXML, allowed ...xml.Name) bool {
	if n == nil || strings.TrimSpace(n.text) != "" {
		return false
	}
	for _, a := range n.attrs {
		if isPreviewNamespaceDeclaration(a) {
			continue
		}
		ok := false
		for _, k := range allowed {
			if a.Name == k {
				ok = true
			}
		}
		if !ok {
			return false
		}
	}
	return true
}
func previewQualifiedAttr(n *previewXML, ns, name string) string {
	if n == nil {
		return ""
	}
	for _, a := range n.attrs {
		if a.Name == (xml.Name{Space: ns, Local: name}) {
			return a.Value
		}
	}
	return ""
}
func previewCanonicalInt(v string, max int64) (int64, bool) {
	n, e := strconv.ParseInt(v, 10, 64)
	return n, e == nil && n >= 0 && n <= max && strconv.FormatInt(n, 10) == v
}
func previewDrawingMarker(n *previewXML, ns string) (NativeDrawingMarkerV1, bool) {
	var m NativeDrawingMarkerV1
	if !previewAttrs(n) || len(n.children) != 4 {
		return m, false
	}
	values := map[string]int64{}
	for _, key := range []string{"col", "row", "colOff", "rowOff"} {
		c := previewChildNS(n, ns, key)
		if c == nil || len(c.children) != 0 {
			return m, false
		}
		for _, a := range c.attrs {
			if !isPreviewNamespaceDeclaration(a) {
				return m, false
			}
		}
		max := int64(2147483647)
		if key == "col" {
			max = 16383
		}
		if key == "row" {
			max = 1048575
		}
		v, ok := previewCanonicalInt(strings.TrimSpace(c.text), max)
		if !ok {
			return m, false
		}
		values[key] = v
	}
	column, row := values["col"], values["row"]
	// Keep the narrowing-conversion bounds local as well as in the parser.
	// These source limits fit int on both native and 32-bit WASM targets.
	if column < 0 || column > 16383 || row < 0 || row > 1048575 {
		return m, false
	}
	return NativeDrawingMarkerV1{int(column), int(row), values["colOff"], values["rowOff"]}, true
}
func previewDrawingAnchor(n *previewXML, ns string) *NativeDrawingAnchorV1 {
	if n == nil || n.name.Space != ns || (n.name.Local != "twoCellAnchor" && n.name.Local != "oneCellAnchor") {
		return nil
	}
	if !previewAttrs(n, xml.Name{Local: "editAs"}) {
		return nil
	}
	if v := n.attr("editAs"); v != "" && (n.name.Local != "twoCellAnchor" || (v != "twoCell" && v != "oneCell" && v != "absolute")) {
		return nil
	}
	from, ok := previewDrawingMarker(previewChildNS(n, ns, "from"), ns)
	if !ok {
		return nil
	}
	result := &NativeDrawingAnchorV1{Kind: n.name.Local, From: from}
	if n.name.Local == "twoCellAnchor" {
		to, ok := previewDrawingMarker(previewChildNS(n, ns, "to"), ns)
		if !ok {
			return nil
		}
		result.To = &to
	} else {
		ext := previewChildNS(n, ns, "ext")
		if !previewAttrs(ext, xml.Name{Local: "cx"}, xml.Name{Local: "cy"}) || len(ext.children) != 0 {
			return nil
		}
		w, ok := previewCanonicalInt(ext.attr("cx"), 2147483647)
		h, ok2 := previewCanonicalInt(ext.attr("cy"), 2147483647)
		if !ok || !ok2 || w == 0 || h == 0 {
			return nil
		}
		result.Width, result.Height = w, h
	}
	return result
}
func previewInternalTarget(pkg *nativeWorkbookPackage, owner, id, kind string) (string, error) {
	raw, ok := pkg.files[path.Join(path.Dir(owner), "_rels", path.Base(owner)+".rels")]
	if !ok {
		return "", nil
	}
	rels, err := parseRoutingRelationships(raw)
	if err != nil {
		return "", err
	}
	for _, r := range rels {
		if r.id != id {
			continue
		}
		if r.targetMode == "External" || (r.relType != officeRelNamespaceTransitional+"/"+kind && r.relType != officeRelNamespaceStrict+"/"+kind) {
			return "", nil
		}
		target, e := resolveNativeRelationshipTarget(owner, path.Dir(owner), r.target)
		if e != nil {
			return "", nil
		}
		actual, _, found := pkg.index.lookupResolved(target)
		if found {
			return actual, nil
		}
	}
	return "", nil
}
func previewChartFrame(frame *previewXML, xdr, a, c, r string) string {
	if !previewAttrs(frame, xml.Name{Local: "macro"}) || frame.attr("macro") != "" || len(frame.children) != 3 {
		return ""
	}
	xfrm := previewChildNS(frame, xdr, "xfrm")
	if !previewAttrs(xfrm) || len(xfrm.children) != 2 {
		return ""
	}
	for _, p := range []struct {
		name string
		keys []string
	}{{"off", []string{"x", "y"}}, {"ext", []string{"cx", "cy"}}} {
		n := previewChildNS(xfrm, a, p.name)
		if !previewAttrs(n, xml.Name{Local: p.keys[0]}, xml.Name{Local: p.keys[1]}) || len(n.children) != 0 || n.attr(p.keys[0]) != "0" || n.attr(p.keys[1]) != "0" {
			return ""
		}
	}
	nv := previewChildNS(frame, xdr, "nvGraphicFramePr")
	if !previewAttrs(nv) || len(nv.children) != 2 {
		return ""
	}
	props := previewChildNS(nv, xdr, "cNvPr")
	if !previewAttrs(props, xml.Name{Local: "id"}, xml.Name{Local: "name"}) {
		return ""
	}
	for _, p := range props.children {
		if p.name.Space != a || (p.name.Local != "hlinkClick" && p.name.Local != "hlinkHover" && p.name.Local != "extLst") {
			return ""
		}
		if p.name.Local == "extLst" && !previewDrawingCreationID(p, a) {
			return ""
		}
		if p.name.Local != "extLst" && (!previewAttrs(p, xml.Name{Space: r, Local: "id"}) || len(p.children) != 0 || previewQualifiedAttr(p, r, "id") == "") {
			return ""
		}
	}
	locks := previewChildNS(nv, xdr, "cNvGraphicFramePr")
	if !previewAttrs(locks) || len(locks.children) != 0 {
		return ""
	}
	graphic := previewChildNS(frame, a, "graphic")
	if !previewAttrs(graphic) || len(graphic.children) != 1 {
		return ""
	}
	data := previewChildNS(graphic, a, "graphicData")
	if !previewAttrs(data, xml.Name{Local: "uri"}) || data.attr("uri") != c || len(data.children) != 1 {
		return ""
	}
	chart := previewChildNS(data, c, "chart")
	if !previewAttrs(chart, xml.Name{Space: r, Local: "id"}) || len(chart.children) != 0 {
		return ""
	}
	return previewQualifiedAttr(chart, r, "id")
}

func previewNativeDrawings(pkg *nativeWorkbookPackage, sheets []NativeWorkbookSheetV2, charts []NativeChartPreviewV1) ([]NativeDrawingObjectV1, error) {
	result := []NativeDrawingObjectV1{}
	knownCharts := map[string]bool{}
	for _, c := range charts {
		knownCharts[c.Part] = true
	}
	for _, sheet := range sheets {
		root, err := parsePreviewXML(pkg.files[sheet.PartName])
		if err != nil {
			return nil, err
		}
		drawingCount := 0
		for _, child := range root.children {
			if child.name.Local == "drawing" {
				drawingCount++
			}
		}
		if drawingCount > 1 {
			return nil, fmt.Errorf("drawing preview refuses ambiguous worksheet drawing declarations")
		}
		for _, drawing := range root.children {
			if drawing.name.Local != "drawing" {
				continue
			}
			base := NativeDrawingObjectV1{SheetID: sheet.ID, SheetPart: sheet.PartName, Kind: "unsupported", Warnings: []string{"Drawing placement unavailable: unsupported or unresolved source drawing. No position is guessed."}}
			r := officeRelNamespaceTransitional
			xdr := "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
			a := "http://schemas.openxmlformats.org/drawingml/2006/main"
			c := "http://schemas.openxmlformats.org/drawingml/2006/chart"
			if root.name.Space == spreadsheetMLStrict {
				r = officeRelNamespaceStrict
				xdr = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"
				a = "http://purl.oclc.org/ooxml/drawingml/main"
				c = "http://purl.oclc.org/ooxml/drawingml/chart"
			}
			part := ""
			if drawing.name.Space == root.name.Space && previewAttrs(drawing, xml.Name{Space: r, Local: "id"}) && len(drawing.children) == 0 {
				part, err = previewInternalTarget(pkg, sheet.PartName, previewQualifiedAttr(drawing, r, "id"), "drawing")
				if err != nil {
					return nil, err
				}
			}
			base.DrawingPart = part
			var dr *previewXML
			if part != "" {
				dr, err = parsePreviewXML(pkg.files[part])
				if err != nil {
					return nil, err
				}
			}
			if dr == nil || dr.name != (xml.Name{Space: xdr, Local: "wsDr"}) || !previewAttrs(dr) {
				result = append(result, base)
			} else {
				for i, anchor := range dr.children {
					item := base
					item.Ordinal = i + 1
					item.Anchor = previewDrawingAnchor(anchor, xdr)
					if item.Anchor != nil && len(anchor.children) == 4 {
						frame := previewChildNS(anchor, xdr, "graphicFrame")
						client := previewChildNS(anchor, xdr, "clientData")
						if previewAttrs(client) && len(client.children) == 0 {
							id := previewChartFrame(frame, xdr, a, c, r)
							target := ""
							if id != "" {
								target, err = previewInternalTarget(pkg, part, id, "chart")
								if err != nil {
									return nil, err
								}
							}
							if knownCharts[target] {
								item.Kind = "chart"
								item.ChartPart = target
								item.Warnings = []string{"Source anchor geometry only. Cached chart painting remains approximate; hyperlinks are never activated and nonvisual interaction metadata is not rendered."}
							}
						}
					}
					result = append(result, item)
					if len(result) > 256 {
						return nil, fmt.Errorf("drawing preview exceeds 256 source objects")
					}
				}
			}
			if len(result) > 256 {
				return nil, fmt.Errorf("drawing preview exceeds 256 source objects")
			}
		}
	}
	return result, nil
}
