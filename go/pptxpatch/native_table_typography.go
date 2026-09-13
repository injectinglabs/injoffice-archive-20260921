package pptxpatch

import "encoding/xml"

// This is an exact, agreement-only source slice. It does not call the inherited
// preview sanitizer, discard active kerning, or infer conflicting precedence.
func (e *nativeExtractor) exactSourceNoBorderTable(frame *nativeXMLNode, part string, d nativeExtractDialect) (nativeExactTable, bool) {
	fail := func() (nativeExactTable, bool) { return nativeExactTable{}, false }
	inspected, _, err := inspectNativeTableSource(frame, d, &nativeInspectionBudget{})
	if err != nil || inspected == nil || len(inspected.Cells) != 1 {
		return fail()
	}
	// Avoid dependency parsing for visible-border tables outside this slice.
	graphicNode := nativeChild(frame, d.drawing, "graphic")
	dataNode := nativeChild(graphicNode, d.drawing, "graphicData")
	tableNode := nativeChild(dataNode, d.drawing, "tbl")
	rowNode := nativeChild(tableNode, d.drawing, "tr")
	cellNode := nativeChild(rowNode, d.drawing, "tc")
	properties := nativeChild(cellNode, d.drawing, "tcPr")
	for _, edge := range []string{"lnL", "lnR", "lnT", "lnB"} {
		line := nativeChild(properties, d.drawing, edge)
		if line == nil || requireOnlyNativeAttrs(line) != nil || !nativePaintOnlyChild(line, d, "noFill") || requireEmptyNativeElement(line.Children[0]) != nil {
			return fail()
		}
	}
	root, err := parseNativeXML(e.pkg.parts[part], part)
	if err != nil {
		return fail()
	}
	common, err := nativeSingleton(root, d.presentation, "cSld", true)
	if err != nil || hasNativeSemanticAttrs(common) {
		return fail()
	}
	tree, err := nativeSingleton(common, d.presentation, "spTree", true)
	if err != nil || hasNativeSemanticAttrs(tree) {
		return fail()
	}
	group, err := nativeSingleton(tree, d.presentation, "grpSpPr", true)
	if err != nil || !inspectTableRootGroup(group, d) {
		return fail()
	}
	// Refuse grouped/placeholder frames even if their local text looks compatible.
	member := false
	for _, child := range tree.Children {
		if child.RawStart == frame.RawStart && child.RawEnd == frame.RawEnd {
			member = true
		}
	}
	if !member {
		return fail()
	}
	c := e.tablePaintContext(part, root, d)
	paint := inspectNativeTablePaint(frame, c, d)
	if paint == nil || paint.Border != nil {
		return fail()
	}
	family, err := c.theme.resolveTypeface("+mn-lt")
	if err != nil {
		return fail()
	}
	colorNode := &nativeXMLNode{Children: []*nativeXMLNode{{Name: xml.Name{Space: d.drawing, Local: "schemeClr"}, Attrs: []xml.Attr{{Name: xml.Name{Local: "val"}, Value: "tx1"}}}}}
	color, err := exactNativeSolidColor(colorNode, d, c.theme)
	if err != nil || color != "000000" {
		return fail()
	}
	rels, err := e.parseRelationships("")
	if err != nil {
		return fail()
	}
	office, err := uniqueNativeInternalRelationship(rels, d.rels+"/officeDocument", "presentation")
	if err != nil {
		return fail()
	}
	presentation, err := parseNativeXML(e.pkg.parts[office.Part], office.Part)
	if err != nil {
		return fail()
	}
	for _, child := range presentation.Children {
		if child.Name.Local == "defaultTextStyle" && child.Name.Space != d.presentation {
			return fail()
		}
	}
	defaults, err := nativeSingleton(presentation, d.presentation, "defaultTextStyle", true)
	if err != nil {
		return fail()
	}
	graph, err := e.resolveSlideDependencyGraph(part, d)
	if err != nil {
		return fail()
	}
	txStyles, err := nativeSingleton(graph.masterRoot, d.presentation, "txStyles", true)
	if err != nil {
		return fail()
	}
	if requireOnlyNativeAttrs(txStyles) != nil || requireOnlyNativeChildren(txStyles, xml.Name{Space: d.presentation, Local: "titleStyle"}, xml.Name{Space: d.presentation, Local: "bodyStyle"}, xml.Name{Space: d.presentation, Local: "otherStyle"}) != nil {
		return fail()
	}
	other, err := nativeSingleton(txStyles, d.presentation, "otherStyle", true)
	if err != nil {
		return fail()
	}
	size, kern, ok := exactTableLevelOne(defaults, d, c.theme, family, color)
	if !ok {
		return fail()
	}
	otherSize, otherKern, ok := exactTableLevelOne(other, d, c.theme, family, color)
	if !ok || size != otherSize || kern != otherKern {
		return fail()
	}
	graphic := nativeChild(frame, d.drawing, "graphic")
	data := nativeChild(graphic, d.drawing, "graphicData")
	table := nativeChild(data, d.drawing, "tbl")
	rows := nativeChildren(table, d.drawing, "tr")
	if len(rows) != 1 {
		return fail()
	}
	cell := nativeChild(rows[0], d.drawing, "tc")
	body := nativeChild(cell, d.drawing, "txBody")
	if body == nil || len(body.Children) != 3 || requireEmptyNativeElement(body.Children[0]) != nil || requireEmptyNativeElement(body.Children[1]) != nil {
		return fail()
	}
	p := body.Children[2]
	if p.Name != (xml.Name{Space: d.drawing, Local: "p"}) || requireOnlyNativeAttrs(p) != nil {
		return fail()
	}
	runs := []NativeTextRun{}
	var text string
	var units int64
	for _, child := range p.Children {
		if child.Name == (xml.Name{Space: d.drawing, Local: "endParaRPr"}) {
			if !exactTableRunMetadata(child, false) {
				return fail()
			}
			continue
		}
		if child.Name != (xml.Name{Space: d.drawing, Local: "r"}) || len(child.Children) != 2 || !exactTableRunMetadata(child.Children[0], true) {
			return fail()
		}
		textNode := child.Children[1]
		if textNode.Name != (xml.Name{Space: d.drawing, Local: "t"}) {
			return fail()
		}
		for _, cp := range textNode.Text {
			if cp < 32 || cp > 126 {
				return fail()
			}
		}
		lang, _ := exactNativeAttr(child.Children[0], "", "lang")
		if !validNativeLanguage(lang) {
			return fail()
		}
		if len(runs) > 0 && *runs[0].Language != lang {
			return fail()
		}
		units += int64(len(textNode.Text))
		text += textNode.Text
		runs = append(runs, NativeTextRun{Text: stringPointer(textNode.Text), Bold: boolPointer(false), Italic: boolPointer(false), FontFamily: stringPointer(family), Color: stringPointer(color), FontSizeHundredthPt: int64Pointer(size), KerningThresholdHundredthPt: int64Pointer(kern), Language: stringPointer(lang)})
	}
	if len(runs) == 0 {
		return fail()
	}
	// Equivalent metadata-only source runs share one shaping span, so kerning
	// is not lost at arbitrary source serialization boundaries.
	runs[0].Text = &text
	runs = runs[:1]
	width, height := inspected.Rect.Width, inspected.Rect.Height
	// ECMA-376 Part1 CT_TableCellProperties: omitted margins and horizontal clip.
	if width <= 182880 || height <= 91440 {
		return fail()
	}
	paragraphs := []NativeParagraph{{Align: nativeTableAlignLeft(), Level: int64Pointer(0), Bullet: boolPointer(false), Runs: runs}}
	nativeCell := NativeTableCell{Text: &text, Paragraphs: &paragraphs, TextBody: &NativeTextBodyLayout{LeftInsetEMU: int64Pointer(91440), RightInsetEMU: int64Pointer(91440), TopInsetEMU: int64Pointer(45720), BottomInsetEMU: int64Pointer(45720), Wrap: NativeTextWrapSquare, VerticalAnchor: NativeTextVerticalAnchorTop, AutoFit: "none", HorizontalOverflow: "clip", VerticalOverflow: "overflow"}}
	usage, err := preflightNativeTableOutput(rows, 1, d)
	if err != nil {
		return fail()
	}
	// Account for the source language and kerning threshold on every run.
	return nativeExactTable{table: NativeTable{ColumnWidths: []int64{width}, RowHeights: []int64{height}, Rows: [][]NativeTableCell{{nativeCell}}}, outputNodes: usage.outputNodes + 2*len(runs), textCodeUnits: units}, true
}
func nativeTableAlignLeft() *NativeTextAlign { v := NativeTextAlignLeft; return &v }

func exactTableRunMetadata(node *nativeXMLNode, baseline bool) bool {
	attrs := []xml.Name{{Local: "lang"}, {Local: "dirty"}, {Local: "smtClean"}, {Local: "err"}}
	if baseline {
		attrs = append(attrs, xml.Name{Local: "baseline"})
	}
	if requireOnlyNativeAttrs(node, attrs...) != nil || requireOnlyNativeChildren(node) != nil {
		return false
	}
	for _, a := range node.Attrs {
		switch a.Name.Local {
		case "lang":
			if !validNativeLanguage(a.Value) {
				return false
			}
		case "baseline":
			if a.Value != "0" {
				return false
			}
		default:
			if _, err := nativeBool(a.Value); err != nil {
				return false
			}
		}
	}
	return true
}

func exactTableLevelOne(layer *nativeXMLNode, d nativeExtractDialect, theme nativeResolvedTheme, family, color string) (int64, int64, bool) {
	fail := func() (int64, int64, bool) { return 0, 0, false }
	if layer == nil || requireOnlyNativeAttrs(layer) != nil || !onlyNativeXMLSpace(layer.Text) {
		return fail()
	}
	seen := map[string]bool{}
	var level *nativeXMLNode
	for _, child := range layer.Children {
		name := child.Name.Local
		if child.Name.Space != d.drawing || seen[name] {
			return fail()
		}
		seen[name] = true
		if name == "defPPr" {
			if requireOnlyNativeAttrs(child) != nil || !nativePaintOnlyChild(child, d, "defRPr") || !exactTableRunMetadata(child.Children[0], false) {
				return fail()
			}
			continue
		}
		if len(name) != 7 || name[:3] != "lvl" || name[3] < '1' || name[3] > '9' || name[4:] != "pPr" {
			return fail()
		}
		if name == "lvl1pPr" {
			level = child
		}
	}
	if level == nil || requireOnlyNativeAttrs(level, xml.Name{Local: "marL"}, xml.Name{Local: "algn"}, xml.Name{Local: "defTabSz"}, xml.Name{Local: "rtl"}, xml.Name{Local: "eaLnBrk"}, xml.Name{Local: "latinLnBrk"}, xml.Name{Local: "hangingPunct"}) != nil || !nativePaintOnlyChild(level, d, "defRPr") {
		return fail()
	}
	for name, want := range map[string]string{"marL": "0", "algn": "l", "defTabSz": "914400", "rtl": "0", "eaLnBrk": "1", "latinLnBrk": "0", "hangingPunct": "1"} {
		value, ok := exactNativeAttr(level, "", name)
		if !ok || value != want {
			return fail()
		}
	}
	run := level.Children[0]
	if !onlyNativeXMLSpace(run.Text) || requireOnlyNativeAttrs(run, xml.Name{Local: "sz"}, xml.Name{Local: "kern"}) != nil || len(run.Children) != 4 {
		return fail()
	}
	size, err := requiredCanonicalNativeTableInt(run, "sz", 1, 400000)
	if err != nil {
		return fail()
	}
	kern, err := requiredCanonicalNativeTableInt(run, "kern", 0, 400000)
	if err != nil {
		return fail()
	}
	for i, name := range []string{"solidFill", "latin", "ea", "cs"} {
		if run.Children[i].Name != (xml.Name{Space: d.drawing, Local: name}) {
			return fail()
		}
	}
	actual, err := exactNativeSolidColor(run.Children[0], d, theme)
	if err != nil || actual != color {
		return fail()
	}
	for i, token := range []string{"+mn-lt", "+mn-ea", "+mn-cs"} {
		node := run.Children[i+1]
		value, ok := exactNativeAttr(node, "", "typeface")
		if !ok || value != token || requireOnlyNativeAttrs(node, xml.Name{Local: "typeface"}) != nil || requireOnlyNativeChildren(node) != nil {
			return fail()
		}
	}
	actualFamily, err := theme.resolveTypeface("+mn-lt")
	if err != nil || actualFamily != family {
		return fail()
	}
	return size, kern, true
}
