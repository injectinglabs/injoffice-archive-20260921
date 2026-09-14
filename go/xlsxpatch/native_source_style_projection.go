package xlsxpatch

import (
	"encoding/xml"
	"fmt"
	"strings"
)

func sourceStyleXFSupported(data []byte, e styleTableEntry, ns string) bool {
	if !styleAttributesOnly(e.start, "numFmtId", "fontId", "fillId", "borderId", "xfId", "applyNumberFormat", "applyFont", "applyFill", "applyBorder", "applyAlignment", "applyProtection") {
		return false
	}
	if v, ok, err := unqualifiedXMLAttribute(e.start, "applyProtection"); err != nil || ok && v != "true" && v != "false" && v != "0" && v != "1" {
		return false
	}
	seen := map[string]bool{}
	for _, c := range e.children {
		if c.start.Name.Space != ns || requireEmptyStyleNode(data, c) != nil || seen[c.start.Name.Local] {
			return false
		}
		seen[c.start.Name.Local] = true
		switch c.start.Name.Local {
		case "alignment":
			if !styleAttributesOnly(c.start, "horizontal", "vertical", "wrapText", "textRotation", "indent", "shrinkToFit") {
				return false
			}
			for _, k := range []string{"textRotation", "indent"} {
				if v, ok, _ := unqualifiedXMLAttribute(c.start, k); ok && v != "0" {
					return false
				}
			}
			if v, ok, _ := unqualifiedXMLAttribute(c.start, "shrinkToFit"); ok && v != "false" && v != "0" {
				return false
			}
		case "protection":
			if !styleAttributesOnly(c.start, "locked", "hidden") {
				return false
			}
			if v, ok, _ := unqualifiedXMLAttribute(c.start, "hidden"); ok && v != "false" && v != "0" {
				return false
			}
			if v, ok, _ := unqualifiedXMLAttribute(c.start, "locked"); ok && v != "true" && v != "false" && v != "0" && v != "1" {
				return false
			}
		default:
			return false
		}
	}
	return true
}
func projectSourceStyle(r *styleRegistry, id int) (NativeSourceStyleV1, error) {
	out := NativeSourceStyleV1{ID: id, Borders: map[string]string{}, Warnings: []string{}}
	fail := func(m string) (NativeSourceStyleV1, error) {
		return out, fmt.Errorf("xlsxpatch: source-style preview: style %d: %s", id, m)
	}
	xf := r.cellXfs[id]
	base := effectiveCellStyleXF(r.styleXfs[xf.xfID])
	ns := r.index.namespace
	if !sourceStyleXFSupported(r.data, r.index.cellXfs.entries[id], ns) || !sourceStyleXFSupported(r.data, r.index.cellStyleXfs.entries[xf.xfID], ns) {
		return fail("unqualified XF or parent properties")
	}
	fontID := effectiveStyleComponent(xf.fontID, base.fontID, xf.applyFont)
	fillID := effectiveStyleComponent(xf.fillID, base.fillID, xf.applyFill)
	borderID := effectiveStyleComponent(xf.borderID, base.borderID, xf.applyBorder)
	numFmtID := effectiveStyleComponent(xf.numFmtID, base.numFmtID, xf.applyNumberFormat)
	// Display selection only; neither source records nor apply flags are modified.
	if xf.applyFill == nil {
		fillID = xf.fillID
	}
	if xf.applyNumberFormat == nil {
		numFmtID = xf.numFmtID
	}
	f := r.fonts[fontID]
	fe := r.index.fonts.entries[fontID]
	if !styleAttributesOnly(fe.start) || f.name == nil || !nativeRichFont.MatchString(*f.name) || f.size == nil || *f.size < 1 || *f.size > 409 {
		return fail("unsupported font name or size")
	}
	seen := map[string]bool{}
	for _, c := range fe.children {
		if c.start.Name.Space != ns || requireEmptyStyleNode(r.data, c) != nil || seen[c.start.Name.Local] {
			return fail("foreign/duplicate font property")
		}
		seen[c.start.Name.Local] = true
		switch c.start.Name.Local {
		case "name", "sz", "b", "i", "family", "charset":
			if !styleAttributesOnly(c.start, "val") {
				return fail("unknown font attributes")
			}
		case "color":
			if !styleAttributesOnly(c.start, "rgb", "theme", "tint") {
				return fail("unsupported font color")
			}
		default:
			return fail("unmodeled font effect")
		}
	}
	out.ParentID = xf.xfID
	out.ParentSHA256 = nativeWorkbookDigest(r.styleXfs[xf.xfID].raw)
	out.FontID, out.FillID, out.BorderID = fontID, fillID, borderID
	out.RawSHA256 = nativeWorkbookDigest(xf.raw)
	out.FontSHA256 = nativeWorkbookDigest(f.raw)
	out.FillSHA256 = nativeWorkbookDigest(r.fills[fillID].raw)
	out.BorderSHA256 = nativeWorkbookDigest(r.borders[borderID].raw)
	out.FontName = *f.name
	out.FontSizePoints = *f.size
	out.Bold = f.bold
	out.Italic = f.italic
	out.FontColor = "#000000"
	if f.colorSafe && f.color != nil {
		out.FontColor = *f.color
	} else {
		out.Warnings = append(out.Warnings, "Unresolved source font color uses explicit host black.")
	}
	out.NumberFormatID = numFmtID
	var ok bool
	out.NumberFormat, ok = r.numberFormatCode(numFmtID)
	if !ok || len(out.NumberFormat) > 1024 {
		return fail("unsupported number-format record")
	}
	fill := r.index.fills.entries[fillID]
	if !styleAttributesOnly(fill.start) || len(fill.children) != 1 {
		return fail("unsupported fill")
	}
	pattern := fill.children[0]
	if pattern.start.Name.Space != ns || pattern.start.Name.Local != "patternFill" || !styleAttributesOnly(pattern.start, "patternType") {
		return fail("unsupported fill pattern")
	}
	kind, _, _ := unqualifiedXMLAttribute(pattern.start, "patternType")
	switch kind {
	case "none":
		if len(pattern.children) != 0 {
			return fail("nonempty no-fill")
		}
	case "solid":
		colors := map[string]bool{}
		for _, c := range pattern.children {
			if c.start.Name.Space != ns || requireEmptyStyleNode(r.data, c) != nil || colors[c.start.Name.Local] {
				return fail("ambiguous fill colors")
			}
			colors[c.start.Name.Local] = true
			if c.start.Name.Local != "fgColor" && c.start.Name.Local != "bgColor" {
				return fail("unknown fill color")
			}
			color, safe := supportedRGBStyleColor(c.start)
			if !safe || color == nil {
				return fail("indirect fill color")
			}
			if c.start.Name.Local == "fgColor" {
				out.FillColor = *color
			}
		}
		if out.FillColor == "" {
			return fail("solid fill has no foreground")
		}
	default:
		return fail("unsupported fill pattern")
	}
	border := r.index.borders.entries[borderID]
	if !styleAttributesOnly(border.start, "diagonalUp", "diagonalDown") {
		return fail("unknown border attributes")
	}
	for _, k := range []string{"diagonalUp", "diagonalDown"} {
		if v, ok, _ := unqualifiedXMLAttribute(border.start, k); ok && v != "0" && v != "false" {
			return fail("active diagonal border")
		}
	}
	sides := map[string]bool{}
	for _, c := range border.children {
		n := c.start.Name.Local
		if c.start.Name.Space != ns || sides[n] {
			return fail("foreign/duplicate border side")
		}
		sides[n] = true
		if n == "diagonal" {
			if !styleAttributesOnly(c.start) || requireEmptyStyleNode(r.data, c) != nil {
				return fail("diagonal content")
			}
			continue
		}
		if n != "left" && n != "right" && n != "top" && n != "bottom" {
			return fail("unknown border side")
		}
		side, supported, err := parseStyleBorderSide(r.data, c, ns)
		if err != nil || !supported {
			return fail("unsupported border side")
		}
		if side != nil {
			if side.style != "thin" {
				return fail("only thin borders are supported")
			}
			out.Borders[n] = side.color
		}
	}
	a := effectiveStyleAlignment(xf, base)
	out.Horizontal = effectiveHorizontal(a.horizontal, nil)
	out.Vertical = effectiveVertical(a.vertical, nil)
	out.Wrap = effectiveWrap(a.wrap, nil)
	if out.Horizontal != "general" && out.Horizontal != "left" && out.Horizontal != "center" && out.Horizontal != "right" {
		return fail("unsupported horizontal alignment")
	}
	if out.Vertical != "top" && out.Vertical != "center" && out.Vertical != "bottom" {
		return fail("unsupported vertical alignment")
	}
	return out, nil
}

// The host grid ignores print settings. Content and stored geometry ownership
// are separately closed; unknown visual or structural families are refused.
func qualifySourceStyleSheet(root *previewXML, ns string) error {
	return qualifySourceStyleSheetWithNodes(root, ns, nil)
}

// admitted contains only fully qualified nodes from the original worksheet tree.
func qualifySourceStyleSheetWithNodes(root *previewXML, ns string, admitted map[*previewXML]bool) error {
	fail := func() error {
		return fmt.Errorf("xlsxpatch: source-style preview: unqualified worksheet structure or geometry")
	}
	if !sourceStyleNode(root, "worksheet") {
		return fail()
	}
	allowed := map[string][]string{
		"worksheet": {}, "sheetPr": {"filterMode"}, "pageSetUpPr": {"fitToPage", "autoPageBreaks"}, "dimension": {"ref"}, "sheetViews": {},
		"sheetView": {"showFormulas", "showGridLines", "showRowColHeaders", "showZeros", "rightToLeft", "tabSelected", "showOutlineSymbols", "defaultGridColor", "view", "topLeftCell", "colorId", "zoomScale", "zoomScaleNormal", "zoomScalePageLayoutView", "workbookViewId"}, "selection": {"pane", "activeCell", "activeCellId", "sqref"},
		"sheetFormatPr": {"defaultColWidth", "baseColWidth", "defaultRowHeight", "customHeight", "zeroHeight", "outlineLevelRow", "outlineLevelCol"}, "cols": {}, "col": {"collapsed", "customWidth", "hidden", "outlineLevel", "max", "min", "style", "width"}, "sheetData": {},
		"row": {"r", "customFormat", "ht", "hidden", "customHeight", "outlineLevel", "collapsed"}, "c": {"r", "s", "t"}, "v": {}, "f": {"aca", "t"}, "is": {}, "t": {}, "mergeCells": {"count"}, "mergeCell": {"ref"},
		"printOptions": {"headings", "gridLines", "gridLinesSet", "horizontalCentered", "verticalCentered"}, "pageMargins": {"left", "right", "top", "bottom", "header", "footer"}, "pageSetup": {"paperSize", "scale", "fitToWidth", "fitToHeight", "pageOrder", "orientation", "blackAndWhite", "draft", "cellComments", "horizontalDpi", "verticalDpi", "copies"}, "headerFooter": {"differentFirst", "differentOddEven"}, "oddHeader": {}, "oddFooter": {},
	}
	children := map[string]string{"worksheet": "sheetPr dimension sheetViews sheetFormatPr cols sheetData mergeCells printOptions pageMargins pageSetup headerFooter", "sheetPr": "pageSetUpPr", "sheetViews": "sheetView", "sheetView": "selection", "cols": "col", "sheetData": "row", "row": "c", "c": "v f is", "is": "t", "mergeCells": "mergeCell", "headerFooter": "oddHeader oddFooter"}
	var walk func(*previewXML) error
	walk = func(n *previewXML) error {
		attrs, ok := allowed[n.name.Local]
		if !ok || n.name.Space != ns {
			return fail()
		}
		if n.name.Local == "t" {
			if !nativeRichNode(n, ns, "t", xml.Name{Space: "http://www.w3.org/XML/1998/namespace", Local: "space"}) {
				return fail()
			}
		} else {
			copy := *n
			if n.name.Local == "v" || n.name.Local == "f" || n.name.Local == "oddHeader" || n.name.Local == "oddFooter" {
				copy.text = ""
			}
			if !sourceStyleNode(&copy, n.name.Local, attrs...) {
				return fail()
			}
		}
		for _, a := range n.attrs {
			if isNamespaceDeclaration(a) {
				continue
			}
			switch a.Name.Local {
			case "hidden", "collapsed", "zeroHeight", "customFormat", "rightToLeft", "showFormulas", "filterMode", "aca":
				if a.Value != "0" && a.Value != "false" {
					return fail()
				}
			case "outlineLevel", "outlineLevelRow", "outlineLevelCol":
				if a.Value != "0" {
					return fail()
				}
			case "style":
				if a.Value != "0" {
					return fail()
				}
			case "t":
				if n.name.Local == "f" && a.Value != "normal" {
					return fail()
				}
			}
		}
		seen := map[string]bool{}
		for _, c := range n.children {
			if admitted[c] {
				continue
			}
			if !strings.Contains(" "+children[n.name.Local]+" ", " "+c.name.Local+" ") {
				return fail()
			}
			if n.name.Local != "sheetData" && n.name.Local != "row" && n.name.Local != "cols" && n.name.Local != "mergeCells" && seen[c.name.Local] {
				return fail()
			}
			seen[c.name.Local] = true
			if err := walk(c); err != nil {
				return err
			}
		}
		return nil
	}
	return walk(root)
}
