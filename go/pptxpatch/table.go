package pptxpatch

import (
	"bytes"
	"encoding/xml"
	"strconv"
	"strings"
)

// tableNode is a deliberately tiny XML tree used only while decoding a
// graphicFrame. Unlike a permissive struct-tag decode, it lets the table
// reader prove that every child and attribute is in its supported subset.
// That is important here: silently ignoring a table property changes the
// visible document, which is worse than omitting an unsupported table.
type tableNode struct {
	start    xml.StartElement
	children []*tableNode
	text     string
}

func readTableNode(dec *xml.Decoder, start xml.StartElement) (*tableNode, error) {
	n := &tableNode{start: start}
	for {
		tok, err := dec.Token()
		if err != nil {
			return nil, err
		}
		switch v := tok.(type) {
		case xml.StartElement:
			child, err := readTableNode(dec, v)
			if err != nil {
				return nil, err
			}
			n.children = append(n.children, child)
		case xml.CharData:
			n.text += string(v)
		case xml.EndElement:
			if v.Name == start.Name {
				return n, nil
			}
			return nil, errMalformedTable
		}
	}
}

// parseTableGraphicFrame reads exactly the writer's conservative subset. A
// false result means the source used something we cannot write back without
// loss (for example, a merge, a themed border, or a rotated frame), so the
// caller leaves that object out rather than inventing a different table.
func parseTableGraphicFrame(raw string) (Shape, bool) {
	dec := xml.NewDecoder(strings.NewReader(raw))
	tok, err := dec.Token()
	if err != nil {
		return Shape{}, false
	}
	start, ok := tok.(xml.StartElement)
	if !ok || start.Name.Local != "graphicFrame" {
		return Shape{}, false
	}
	n, err := readTableNode(dec, start)
	if err != nil {
		return Shape{}, false
	}
	return tableShapeFromNode(n)
}

var errMalformedTable = &tableError{"malformed table XML"}

type tableError struct{ message string }

func (e *tableError) Error() string { return e.message }

func tableShapeFromNode(root *tableNode) (Shape, bool) {
	if !tableElement(root, "graphicFrame") || !tableAttrs(root) || !tableOnlyWhitespace(root) ||
		!tableChildren(root, "nvGraphicFramePr", "xfrm", "graphic") {
		return Shape{}, false
	}
	nv, xfrm, graphic := root.children[0], root.children[1], root.children[2]
	name, ok := tableGraphicName(nv)
	if !ok {
		return Shape{}, false
	}
	x, y, cx, cy, ok := tableGraphicXfrm(xfrm)
	if !ok {
		return Shape{}, false
	}
	table, ok := tableGraphicData(graphic)
	if !ok || !table.Valid() {
		return Shape{}, false
	}
	if sumInts(table.Columns) != cx || sumInts(table.RowHeights) != cy {
		// A graphicFrame's transform and its grid are both authoritative. If
		// they disagree, preserving only one would visibly rescale the table.
		return Shape{}, false
	}
	return Shape{Name: name, X: x, Y: y, Cx: cx, Cy: cy, Table: &table}, true
}

func tableGraphicName(n *tableNode) (string, bool) {
	if !tableElement(n, "nvGraphicFramePr") || !tableAttrs(n) || !tableOnlyWhitespace(n) ||
		!tableChildren(n, "cNvPr", "cNvGraphicFramePr", "nvPr") {
		return "", false
	}
	cNvPr := n.children[0]
	if !tableElement(cNvPr, "cNvPr") || !tableOnlyWhitespace(cNvPr) || len(cNvPr.children) != 0 ||
		!tableAttrs(cNvPr, "id", "name") || attrVal(cNvPr.start, "id") == "" {
		return "", false
	}
	if _, err := strconv.Atoi(attrVal(cNvPr.start, "id")); err != nil {
		return "", false
	}
	for _, child := range n.children[1:] {
		if !tableAttrs(child) || !tableOnlyWhitespace(child) || len(child.children) != 0 {
			return "", false
		}
	}
	return attrVal(cNvPr.start, "name"), true
}

func tableGraphicXfrm(n *tableNode) (x, y, cx, cy int, ok bool) {
	if !tableElement(n, "xfrm") || !tableAttrs(n) || !tableOnlyWhitespace(n) || !tableChildren(n, "off", "ext") {
		return 0, 0, 0, 0, false
	}
	off, ext := n.children[0], n.children[1]
	if !tableElement(off, "off") || !tableAttrs(off, "x", "y") || len(off.children) != 0 || !tableOnlyWhitespace(off) ||
		!tableElement(ext, "ext") || !tableAttrs(ext, "cx", "cy") || len(ext.children) != 0 || !tableOnlyWhitespace(ext) {
		return 0, 0, 0, 0, false
	}
	var err error
	if x, err = strconv.Atoi(attrVal(off.start, "x")); err != nil {
		return 0, 0, 0, 0, false
	}
	if y, err = strconv.Atoi(attrVal(off.start, "y")); err != nil {
		return 0, 0, 0, 0, false
	}
	if cx, err = strconv.Atoi(attrVal(ext.start, "cx")); err != nil || cx <= 0 {
		return 0, 0, 0, 0, false
	}
	if cy, err = strconv.Atoi(attrVal(ext.start, "cy")); err != nil || cy <= 0 {
		return 0, 0, 0, 0, false
	}
	return x, y, cx, cy, true
}

func tableGraphicData(n *tableNode) (Table, bool) {
	if !tableElement(n, "graphic") || !tableAttrs(n) || !tableOnlyWhitespace(n) || !tableChildren(n, "graphicData") {
		return Table{}, false
	}
	data := n.children[0]
	const tableURI = "http://schemas.openxmlformats.org/drawingml/2006/table"
	if !tableAttrs(data, "uri") || attrVal(data.start, "uri") != tableURI || !tableOnlyWhitespace(data) || !tableChildren(data, "tbl") {
		return Table{}, false
	}
	return tableFromNode(data.children[0])
}

func tableFromNode(n *tableNode) (Table, bool) {
	if !tableElement(n, "tbl") || !tableAttrs(n) || !tableOnlyWhitespace(n) || len(n.children) < 3 ||
		n.children[0].start.Name.Local != "tblPr" || n.children[1].start.Name.Local != "tblGrid" {
		return Table{}, false
	}
	if !tableElement(n.children[0], "tblPr") || !tableAttrs(n.children[0]) || !tableOnlyWhitespace(n.children[0]) || len(n.children[0].children) != 0 {
		return Table{}, false
	}
	grid := n.children[1]
	if !tableElement(grid, "tblGrid") || !tableAttrs(grid) || !tableOnlyWhitespace(grid) || len(grid.children) == 0 {
		return Table{}, false
	}
	t := Table{}
	for _, col := range grid.children {
		if !tableElement(col, "gridCol") || !tableAttrs(col, "w") || !tableOnlyWhitespace(col) || len(col.children) != 0 {
			return Table{}, false
		}
		w, err := strconv.Atoi(attrVal(col.start, "w"))
		if err != nil || w <= 0 {
			return Table{}, false
		}
		t.Columns = append(t.Columns, w)
	}
	for _, row := range n.children[2:] {
		cells, h, ok := tableRowFromNode(row, len(t.Columns))
		if !ok {
			return Table{}, false
		}
		t.Rows = append(t.Rows, cells)
		t.RowHeights = append(t.RowHeights, h)
	}
	return t, t.Valid()
}

func tableRowFromNode(n *tableNode, columns int) ([]TableCell, int, bool) {
	if !tableElement(n, "tr") || !tableAttrs(n, "h") || !tableOnlyWhitespace(n) || len(n.children) != columns {
		return nil, 0, false
	}
	h, err := strconv.Atoi(attrVal(n.start, "h"))
	if err != nil || h <= 0 {
		return nil, 0, false
	}
	cells := make([]TableCell, 0, columns)
	for _, cell := range n.children {
		c, ok := tableCellFromNode(cell)
		if !ok {
			return nil, 0, false
		}
		cells = append(cells, c)
	}
	return cells, h, true
}

func tableCellFromNode(n *tableNode) (TableCell, bool) {
	if !tableElement(n, "tc") || !tableAttrs(n) || !tableOnlyWhitespace(n) || !tableChildren(n, "txBody", "tcPr") {
		return TableCell{}, false
	}
	text, align, ok := tableCellText(n.children[0])
	if !ok {
		return TableCell{}, false
	}
	fill, border, ok := tableCellProperties(n.children[1])
	if !ok {
		return TableCell{}, false
	}
	return TableCell{Text: text, Align: align, Fill: fill, Border: border}, true
}

func tableCellText(n *tableNode) (string, TextAlign, bool) {
	if !tableElement(n, "txBody") || !tableAttrs(n) || !tableOnlyWhitespace(n) || len(n.children) < 3 ||
		n.children[0].start.Name.Local != "bodyPr" || n.children[1].start.Name.Local != "lstStyle" {
		return "", "", false
	}
	for _, base := range n.children[:2] {
		if !tableAttrs(base) || !tableOnlyWhitespace(base) || len(base.children) != 0 {
			return "", "", false
		}
	}
	var paragraphs []string
	var align TextAlign
	for _, p := range n.children[2:] {
		value, pAlign, ok := tableParagraphText(p)
		if !ok || (align != "" && pAlign != align) {
			return "", "", false
		}
		align = pAlign
		paragraphs = append(paragraphs, value)
	}
	if align == "" {
		align = AlignLeft
	}
	return strings.Join(paragraphs, "\n"), align, true
}

func tableParagraphText(n *tableNode) (string, TextAlign, bool) {
	if !tableElement(n, "p") || !tableAttrs(n) || !tableOnlyWhitespace(n) || len(n.children) == 0 || n.children[0].start.Name.Local != "pPr" {
		return "", "", false
	}
	pPr := n.children[0]
	if !tableAttrs(pPr, "algn") || !tableOnlyWhitespace(pPr) || !tableChildren(pPr, "buNone") ||
		!tableAttrs(pPr.children[0]) || !tableOnlyWhitespace(pPr.children[0]) || len(pPr.children[0].children) != 0 {
		return "", "", false
	}
	align := TextAlign(attrVal(pPr.start, "algn"))
	if align == "" {
		align = AlignLeft
	}
	if align != AlignLeft && align != AlignCenter && align != AlignRight {
		return "", "", false
	}
	var text strings.Builder
	for _, run := range n.children[1:] {
		if !tableElement(run, "r") || !tableAttrs(run) || !tableOnlyWhitespace(run) || !tableChildren(run, "t") {
			return "", "", false
		}
		t := run.children[0]
		if !tableElement(t, "t") || !tableTextAttrs(t) || len(t.children) != 0 {
			return "", "", false
		}
		text.WriteString(t.text)
	}
	return text.String(), align, true
}

func tableCellProperties(n *tableNode) (string, TableBorder, bool) {
	if !tableElement(n, "tcPr") || !tableAttrs(n) || !tableOnlyWhitespace(n) || len(n.children) != 5 {
		return "", TableBorder{}, false
	}
	fill, ok := tableFill(n.children[0])
	if !ok {
		return "", TableBorder{}, false
	}
	var border TableBorder
	for i, edge := range []string{"lnL", "lnR", "lnT", "lnB"} {
		if !tableElement(n.children[i+1], edge) {
			return "", TableBorder{}, false
		}
		got, ok := tableBorderFromNode(n.children[i+1])
		if !ok || (i > 0 && got != border) {
			return "", TableBorder{}, false
		}
		border = got
	}
	return fill, border, true
}

func tableFill(n *tableNode) (string, bool) {
	if !tableOnlyWhitespace(n) || !tableAttrs(n) {
		return "", false
	}
	switch n.start.Name.Local {
	case "noFill":
		return "", len(n.children) == 0
	case "solidFill":
		if !tableChildren(n, "srgbClr") {
			return "", false
		}
		return tableSRGB(n.children[0])
	default:
		return "", false
	}
}

func tableBorderFromNode(n *tableNode) (TableBorder, bool) {
	if !tableOnlyWhitespace(n) {
		return TableBorder{}, false
	}
	if len(n.children) != 1 {
		return TableBorder{}, false
	}
	if n.children[0].start.Name.Local == "noFill" {
		return TableBorder{}, tableAttrs(n) && tableAttrs(n.children[0]) && len(n.children[0].children) == 0 && tableOnlyWhitespace(n.children[0])
	}
	if !tableAttrs(n, "w") || n.children[0].start.Name.Local != "solidFill" {
		return TableBorder{}, false
	}
	w, err := strconv.Atoi(attrVal(n.start, "w"))
	if err != nil || w <= 0 {
		return TableBorder{}, false
	}
	color, ok := tableFill(n.children[0])
	if !ok || color == "" {
		return TableBorder{}, false
	}
	return TableBorder{Color: color, WidthPt: float64(w) / EMUPerPt}, true
}

func tableSRGB(n *tableNode) (string, bool) {
	if !tableElement(n, "srgbClr") || !tableAttrs(n, "val") || !tableOnlyWhitespace(n) || len(n.children) != 0 {
		return "", false
	}
	v := attrVal(n.start, "val")
	if len(v) != 6 {
		return "", false
	}
	for _, r := range v {
		if !((r >= '0' && r <= '9') || (r >= 'a' && r <= 'f') || (r >= 'A' && r <= 'F')) {
			return "", false
		}
	}
	return "#" + strings.ToLower(v), true
}

func tableElement(n *tableNode, local string) bool { return n.start.Name.Local == local }

func tableOnlyWhitespace(n *tableNode) bool { return strings.TrimSpace(n.text) == "" }

func tableChildren(n *tableNode, names ...string) bool {
	if len(n.children) != len(names) {
		return false
	}
	for i, name := range names {
		if n.children[i].start.Name.Local != name {
			return false
		}
	}
	return true
}

func tableAttrs(n *tableNode, names ...string) bool {
	attrs := tableNonNamespaceAttrs(n.start.Attr)
	if len(attrs) != len(names) {
		return false
	}
	for _, wanted := range names {
		if attrVal(n.start, wanted) == "" {
			return false
		}
	}
	return true
}

func tableTextAttrs(n *tableNode) bool {
	attrs := tableNonNamespaceAttrs(n.start.Attr)
	if len(attrs) == 0 {
		return true
	}
	return len(attrs) == 1 && attrs[0].Name.Space == "http://www.w3.org/XML/1998/namespace" && attrs[0].Name.Local == "space" && attrs[0].Value == "preserve"
}

// encode/xml's token re-encoder emits namespace bindings on individual
// elements when graphicFrameXML extracts a fragment. Those declarations are
// not table styling, so strict attribute checks must ignore them while still
// rejecting every ordinary unexpected OOXML attribute.
func tableNonNamespaceAttrs(attrs []xml.Attr) []xml.Attr {
	out := make([]xml.Attr, 0, len(attrs))
	for _, attr := range attrs {
		if attr.Name.Space == "xmlns" || (attr.Name.Space == "" && attr.Name.Local == "xmlns") {
			continue
		}
		out = append(out, attr)
	}
	return out
}

func sumInts(values []int) int {
	total := 0
	for _, value := range values {
		total += value
	}
	return total
}

// graphicFrameXML consumes the current start tag and returns a standalone XML
// representation for the strict table parser. encoding/xml recreates the
// required namespace bindings as it writes the token stream.
func graphicFrameXML(dec *xml.Decoder, start xml.StartElement) (string, error) {
	var b bytes.Buffer
	enc := xml.NewEncoder(&b)
	if err := enc.EncodeToken(start); err != nil {
		return "", err
	}
	depth := 1
	for depth > 0 {
		tok, err := dec.Token()
		if err != nil {
			return "", err
		}
		if err := enc.EncodeToken(tok); err != nil {
			return "", err
		}
		switch tok.(type) {
		case xml.StartElement:
			depth++
		case xml.EndElement:
			depth--
		}
	}
	if err := enc.Flush(); err != nil {
		return "", err
	}
	return b.String(), nil
}
