package pptxpatch

import (
	"encoding/xml"
	"fmt"
	"regexp"
	"strings"
)

var inspectionUUID = regexp.MustCompile(`^\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}$`)
var inspectionLanguage = regexp.MustCompile(`^[A-Za-z]{2,8}(?:-[A-Za-z0-9]{1,8}){0,3}$`)
var inspectionRGB = regexp.MustCompile(`^[0-9A-Fa-f]{6}$`)

// This grammar describes readable source, not a style cascade. Every accepted
// formatting property is explicitly omitted from the inspection's plain text.
// Missing layout/paint values never become Office defaults.
var inspectionChildren = map[string]string{
	"tbl": "tblPr tblGrid tr*", "tblPr": "tableStyleId", "tblGrid": "gridCol*", "gridCol": "",
	"tr": "tc*", "tc": "txBody tcPr", "txBody": "bodyPr lstStyle p*", "bodyPr": "", "lstStyle": "",
	"p": "pPr r* endParaRPr", "r": "rPr t", "rPr": "", "endParaRPr": "", "t": "", "tableStyleId": "",
	"pPr":   "lnSpc spcBef spcAft buClrTx buSzTx buFontTx buNone tabLst defRPr",
	"lnSpc": "spcPct spcPts", "spcBef": "spcPct spcPts", "spcAft": "spcPct spcPts",
	"spcPct": "", "spcPts": "", "buClrTx": "", "buSzTx": "", "buFontTx": "", "buNone": "", "tabLst": "", "defRPr": "",
	"tcPr": "lnL lnR lnT lnB lnTlToBr lnBlToTr noFill solidFill",
	"lnL":  "noFill solidFill prstDash round headEnd tailEnd", "lnR": "noFill solidFill prstDash round headEnd tailEnd",
	"lnT": "noFill solidFill prstDash round headEnd tailEnd", "lnB": "noFill solidFill prstDash round headEnd tailEnd",
	"lnTlToBr": "noFill prstDash", "lnBlToTr": "noFill prstDash", "noFill": "", "solidFill": "schemeClr srgbClr",
	"schemeClr": "", "srgbClr": "", "prstDash": "", "round": "", "headEnd": "", "tailEnd": "",
}

func inspectionAttr(node, name, value string) bool {
	integer := func(min, max int64) bool { _, err := parseCanonicalNativeInt(value, min, max); return err == nil }
	member := func(values string) bool { return strings.Contains(" "+values+" ", " "+value+" ") }
	switch node {
	case "tblPr":
		return (name == "firstRow" || name == "bandRow") && (value == "0" || value == "1")
	case "gridCol":
		return name == "w" && integer(1, 1_000_000_000)
	case "tr":
		return name == "h" && integer(1, 1_000_000_000)
	case "rPr", "endParaRPr":
		switch name {
		case "lang":
			return inspectionLanguage.MatchString(value)
		case "dirty", "smtClean", "err":
			return value == "0" || value == "1"
		case "baseline":
			return value == "0"
		}
	case "pPr":
		switch name {
		case "marL", "marR", "indent":
			return integer(-100_000_000, 100_000_000)
		case "defTabSz":
			return integer(1, 100_000_000)
		case "rtl", "eaLnBrk", "latinLnBrk", "hangingPunct":
			return value == "0" || value == "1"
		case "algn":
			return member("l ctr r just dist thaiDist justLow")
		case "fontAlgn":
			return member("auto t ctr base b")
		}
	case "lnL", "lnR", "lnT", "lnB", "lnTlToBr", "lnBlToTr":
		switch name {
		case "w":
			return integer(1, nativeMaxLineWidthEmu)
		case "cap":
			return member("flat rnd sq")
		case "cmpd":
			return member("sng dbl thickThin thinThick tri")
		case "algn":
			return member("ctr in")
		}
	case "schemeClr":
		return name == "val" && member("bg1 tx1 bg2 tx2 accent1 accent2 accent3 accent4 accent5 accent6 hlink folHlink dk1 lt1 dk2 lt2")
	case "srgbClr":
		return name == "val" && inspectionRGB.MatchString(value)
	case "prstDash":
		return name == "val" && member("solid dot dash lgDash dashDot lgDashDot lgDashDotDot sysDash sysDot sysDashDot sysDashDotDot")
	case "headEnd", "tailEnd":
		return name == "type" && value == "none" || (name == "w" || name == "len") && value == "med"
	case "spcPct":
		return name == "val" && integer(0, 1_000_000)
	case "spcPts":
		return name == "val" && integer(0, 158400)
	}
	return false
}

func inspectTableGrammar(node *nativeXMLNode, d nativeExtractDialect) error {
	grammar, known := inspectionChildren[node.Name.Local]
	if !known || node.Name.Space != d.drawing {
		return fmt.Errorf("table inspection: unknown source markup")
	}
	for _, name := range strings.Fields(map[string]string{"gridCol": "w", "tr": "h", "schemeClr": "val", "srgbClr": "val", "prstDash": "val", "spcPct": "val", "spcPts": "val", "headEnd": "type w len", "tailEnd": "type w len"}[node.Name.Local]) {
		if _, ok := exactNativeAttr(node, "", name); !ok {
			return fmt.Errorf("table inspection: missing required source attribute")
		}
	}
	for _, a := range node.Attrs {
		if a.Name.Space == "xmlns" || a.Name.Local == "xmlns" {
			continue
		}
		if a.Name.Space != "" || !inspectionAttr(node.Name.Local, a.Name.Local, a.Value) {
			return fmt.Errorf("table inspection: unqualified source attribute on %s", node.Name.Local)
		}
	}
	if node.Name.Local == "t" {
		if len(node.Children) != 0 || utf16CodeUnitLengthBounded(node.Text, nativeTableInspectionMaxText+1) > nativeTableInspectionMaxText {
			return fmt.Errorf("table inspection: invalid or oversized text")
		}
	} else if node.Name.Local == "tableStyleId" {
		if !inspectionUUID.MatchString(node.Text) {
			return fmt.Errorf("table inspection: malformed style identifier")
		}
	} else if !onlyNativeXMLSpace(node.Text) {
		return fmt.Errorf("table inspection: unexpected source text")
	}
	positions := map[string]int{}
	repeats := map[string]bool{}
	for i, name := range strings.Fields(grammar) {
		clean := strings.TrimSuffix(name, "*")
		positions[clean] = i
		repeats[clean] = name != clean
	}
	last := -1
	seen := map[string]bool{}
	for _, child := range node.Children {
		position, ok := positions[child.Name.Local]
		if !ok || position < last || seen[child.Name.Local] && !repeats[child.Name.Local] {
			return fmt.Errorf("table inspection: duplicate or out-of-order source child")
		}
		last = position
		seen[child.Name.Local] = true
		if err := inspectTableGrammar(child, d); err != nil {
			return err
		}
	}
	if node.Name.Local == "solidFill" || node.Name.Local == "lnSpc" || node.Name.Local == "spcBef" || node.Name.Local == "spcAft" {
		if len(node.Children) != 1 {
			return fmt.Errorf("table inspection: ambiguous paint/spacing source")
		}
	}
	if seen["noFill"] && seen["solidFill"] {
		return fmt.Errorf("table inspection: conflicting paint source")
	}
	for _, name := range strings.Fields(map[string]string{"tbl": "tblPr tblGrid tr", "tblGrid": "gridCol", "tr": "tc", "tc": "txBody tcPr", "txBody": "bodyPr lstStyle p", "r": "rPr t"}[node.Name.Local]) {
		if !seen[name] {
			return fmt.Errorf("table inspection: missing required source child")
		}
	}
	return nil
}

func inspectNativeTableSource(node *nativeXMLNode, d nativeExtractDialect, budget *nativeInspectionBudget) (*NativePPTXInspectedTable, string, error) {
	fail := func(err error) (*NativePPTXInspectedTable, string, error) { return nil, "unavailable", err }
	// Identify a well-formed non-table payload before applying table-specific
	// nonvisual qualifications (charts use different metadata).
	if graphic, e := nativeSingleton(node, d.drawing, "graphic", true); e == nil {
		if gd, e := nativeSingleton(graphic, d.drawing, "graphicData", true); e == nil {
			if uri, ok := exactNativeAttr(gd, "", "uri"); ok && uri != "" && uri != nativeDrawingTableURI && requireOnlyNativeAttrs(gd, xml.Name{Local: "uri"}) == nil {
				return nil, "", nil
			}
		}
	}
	if err := budget.scan(node, d); err != nil {
		return fail(err)
	}
	if requireOnlyNativeAttrs(node) != nil || requireOnlyNativeChildren(node, xml.Name{Space: d.presentation, Local: "nvGraphicFramePr"}, xml.Name{Space: d.presentation, Local: "xfrm"}, xml.Name{Space: d.drawing, Local: "graphic"}) != nil {
		return fail(fmt.Errorf("table inspection: unqualified frame"))
	}
	nv, err := nativeSingleton(node, d.presentation, "nvGraphicFramePr", true)
	if err != nil {
		return fail(err)
	}
	xfrm, err := nativeSingleton(node, d.presentation, "xfrm", true)
	if err != nil {
		return fail(err)
	}
	graphic, err := nativeSingleton(node, d.drawing, "graphic", true)
	if err != nil {
		return fail(err)
	}
	if len(node.Children) != 3 || node.Children[0] != nv || node.Children[1] != xfrm || node.Children[2] != graphic {
		return fail(fmt.Errorf("table inspection: malformed frame order"))
	}
	id, err := inspectTableNonVisual(nv, d)
	if err != nil {
		return fail(err)
	}
	transform, err := validateNativeTableTransform(xfrm, d)
	if err != nil {
		return nil, id, err
	}
	if *transform.X < -1_000_000_000 || *transform.X > 1_000_000_000 || *transform.Y < -1_000_000_000 || *transform.Y > 1_000_000_000 || *transform.Cx > 1_000_000_000 || *transform.Cy > 1_000_000_000 {
		return fail(fmt.Errorf("table inspection: frame exceeds preview coordinate budget"))
	}
	if requireOnlyNativeAttrs(graphic) != nil || requireOnlyNativeChildren(graphic, xml.Name{Space: d.drawing, Local: "graphicData"}) != nil {
		return fail(fmt.Errorf("table inspection: invalid graphic"))
	}
	gd, err := nativeSingleton(graphic, d.drawing, "graphicData", true)
	if err != nil {
		return fail(err)
	}
	uri, _ := exactNativeAttr(gd, "", "uri")
	if uri != nativeDrawingTableURI {
		return nil, id, nil
	}
	if requireOnlyNativeAttrs(gd, xml.Name{Local: "uri"}) != nil || requireOnlyNativeChildren(gd, xml.Name{Space: d.drawing, Local: "tbl"}) != nil {
		return fail(fmt.Errorf("table inspection: unknown graphic data"))
	}
	tbl, err := nativeSingleton(gd, d.drawing, "tbl", true)
	if err != nil {
		return fail(err)
	}
	if err := inspectTableGrammar(tbl, d); err != nil {
		return nil, id, err
	}
	grid, err := nativeSingleton(tbl, d.drawing, "tblGrid", true)
	if err != nil {
		return fail(err)
	}
	columns, total, err := extractNativeTableGrid(grid, d)
	if err != nil {
		return fail(err)
	}
	rows := nativeChildren(tbl, d.drawing, "tr")
	if len(rows) == 0 || len(rows) > 32 || len(columns) > 32 || len(rows)*len(columns) > nativeTableInspectionMaxCells || total != *transform.Cx {
		return fail(fmt.Errorf("table inspection: unsupported grid geometry"))
	}
	result := &NativePPTXInspectedTable{ObjectID: id, Rect: NativePPTXInspectionRect{*transform.X, *transform.Y, *transform.Cx, *transform.Cy}, Cells: []NativePPTXInspectedCell{}, Warnings: []string{"Plain source text and stored cell geometry only; this is not a rendered Office table.", "Table styles, borders and fills are omitted, including any authored dash patterns.", "Source text styling, margins, alignment, spacing and inherited font metrics are not applied."}}
	y := int64(0)
	for ri, row := range rows {
		height, err := requiredCanonicalNativeTableInt(row, "h", 1, 1_000_000_000)
		if err != nil {
			return fail(err)
		}
		cells := nativeChildren(row, d.drawing, "tc")
		if len(cells) != len(columns) {
			return fail(fmt.Errorf("table inspection: ragged or merged cells"))
		}
		x := int64(0)
		for ci, cell := range cells {
			body, err := nativeSingleton(cell, d.drawing, "txBody", true)
			if err != nil {
				return fail(err)
			}
			paragraphs := []string{}
			for _, p := range nativeChildren(body, d.drawing, "p") {
				var text strings.Builder
				for _, r := range nativeChildren(p, d.drawing, "r") {
					t, err := nativeSingleton(r, d.drawing, "t", true)
					if err != nil {
						return fail(err)
					}
					text.WriteString(t.Text)
				}
				paragraphs = append(paragraphs, text.String())
			}
			if len(paragraphs) == 0 {
				return fail(fmt.Errorf("table inspection: missing paragraph"))
			}
			result.Cells = append(result.Cells, NativePPTXInspectedCell{ri, ci, NativePPTXInspectionRect{x, y, columns[ci], height}, paragraphs})
			x += columns[ci]
		}
		y += height
	}
	if y != *transform.Cy {
		return fail(fmt.Errorf("table inspection: frame/row height mismatch"))
	}
	return result, id, nil
}
