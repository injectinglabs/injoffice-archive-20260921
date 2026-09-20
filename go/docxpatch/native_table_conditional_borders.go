package docxpatch

import (
	"encoding/xml"
	"strings"
)

// Borders of a conditionally styled table, resolved per cell.
//
// A w:tblStylePr region's w:tcBorders (ECMA-376 17.7.6.6, 17.4.66) describe
// the region as one box: top/bottom/left/right are its outer edges and
// insideH/insideV the edges between its cells. Regions apply in
// nativeTableRegionOrder, and an edge no region states falls back to the
// table's own w:tblBorders. Two cells sharing an edge then resolve it the way
// Word does (17.4.39): the heavier border wins, a stated "nil" weighing
// nothing, and equal weights go to the darker colour.
//
// Word paints an automatic border colour black over every fill, dark or
// light: its raster of conditionalstyles-tbllook.docx puts the 2.25 pt auto
// top and bottom rules in 000000 over 833C0B, FFC000, FF0000 and 385623
// alike. A theme colour with a tint or shade is read from the w:color Word
// writes beside it, exactly as the fills are.

// NativeResolvedTableCellBordersV1 is one cell's four painted edges after the
// conditional cascade and the shared-edge resolution.
type NativeResolvedTableCellBordersV1 struct {
	CellID  string               `json:"cell_id"`
	Borders NativeTableBordersV1 `json:"borders"`
}

// nativeRegionBorders holds a region's stated edges; a nil entry is unstated.
type nativeRegionBorders struct {
	top, left, bottom, right, insideH, insideV *NativeTableBorderV1
}

func (borders *nativeRegionBorders) set(edge string, border *NativeTableBorderV1) bool {
	switch edge {
	case "top":
		borders.top = border
	case "left":
		borders.left = border
	case "bottom":
		borders.bottom = border
	case "right":
		borders.right = border
	case "insideH":
		borders.insideH = border
	case "insideV":
		borders.insideV = border
	default:
		return false
	}
	return true
}

// nativeAuthoredTableBorder reads a border edge as Word paints it: single or
// absent, auto colour black, a theme tint or shade taken from the authored
// w:color. Other border styles are outside this tier and return ok=false.
func nativeAuthoredTableBorder(node *nativeXMLNode, wordNS string, resolveTheme func(string) (string, bool)) (*NativeTableBorderV1, bool) {
	if border, ok := nativeExtractTableBorder(node, wordNS, resolveTheme); ok {
		return border, true
	}
	if !nativeExactLeaf(node,
		xml.Name{Space: wordNS, Local: "val"},
		xml.Name{Space: wordNS, Local: "sz"},
		xml.Name{Space: wordNS, Local: "color"},
		xml.Name{Space: wordNS, Local: "space"},
		xml.Name{Space: wordNS, Local: "themeColor"},
		xml.Name{Space: wordNS, Local: "themeTint"},
		xml.Name{Space: wordNS, Local: "themeShade"},
	) || !nativeZeroOrAbsentTwipAttr(node, wordNS, "space") {
		return nil, false
	}
	value, ok := nativeAttr(node, wordNS, "val")
	if !ok {
		return nil, false
	}
	if value == "none" || value == "nil" {
		return &NativeTableBorderV1{Style: "none", SizeEighthPoints: 0}, true
	}
	if value != "single" {
		return nil, false
	}
	size, sizeOK := nativeNonnegativeInt64Attr(node, wordNS, "sz")
	if !sizeOK || size <= 0 || size > 768 {
		return nil, false
	}
	color, hasColor := nativeAttr(node, wordNS, "color")
	rgb := "000000"
	if hasColor && !strings.EqualFold(color, "auto") {
		exact, valid := nativeExactRGB(color)
		if !valid {
			return nil, false
		}
		rgb = exact
	}
	return &NativeTableBorderV1{Style: "single", SizeEighthPoints: size, ColorRGB: nativeString(rgb)}, true
}

func (resolver *nativeLayoutResolver) parseRegionBorders(node *nativeXMLNode) (*nativeRegionBorders, bool) {
	if !nativeExactContainer(node) {
		return nil, false
	}
	out := &nativeRegionBorders{}
	seen := map[string]bool{}
	for _, edge := range node.Children {
		if edge.Name.Space != resolver.wordNS || seen[edge.Name.Local] {
			return nil, false
		}
		seen[edge.Name.Local] = true
		border, ok := nativeAuthoredTableBorder(edge, resolver.wordNS, resolver.resolveThemeSrgb)
		if !ok {
			return nil, false
		}
		switch edge.Name.Local {
		case "tl2br", "tr2bl":
			// Diagonals are not painted; they do not change the four edges.
			continue
		}
		if !out.set(edge.Name.Local, border) {
			return nil, false
		}
	}
	return out, true
}

// mergeRegionBorders overlays later-stated edges on earlier ones.
func mergeRegionBorders(base, overlay *nativeRegionBorders) *nativeRegionBorders {
	if overlay == nil {
		return base
	}
	if base == nil {
		copied := *overlay
		return &copied
	}
	out := *base
	for _, pair := range []struct {
		target **NativeTableBorderV1
		value  *NativeTableBorderV1
	}{{&out.top, overlay.top}, {&out.left, overlay.left}, {&out.bottom, overlay.bottom}, {&out.right, overlay.right}, {&out.insideH, overlay.insideH}, {&out.insideV, overlay.insideV}} {
		if pair.value != nil {
			*pair.target = pair.value
		}
	}
	return &out
}

// tableBordersFromContract reads the table's own resolved w:tblBorders (the
// document's, else the chain's) into region form.
func regionBordersFromTable(borders *NativeTableBordersV1) *nativeRegionBorders {
	if borders == nil {
		return &nativeRegionBorders{}
	}
	return &nativeRegionBorders{top: borders.Top, left: borders.Left, bottom: borders.Bottom, right: borders.Right, insideH: borders.InsideHorizontal, insideV: borders.InsideVertical}
}

// regionExtent is the box a region occupies, in rows and grid columns, for
// the cell at (row, column): row regions span every column, column regions
// span every row, corners are the cell itself, and a band is its contiguous
// group of rowBandSize rows or colBandSize columns.
func (conditional *nativeConditionalTableStyle) regionExtent(kind string, row, rowCount, column, span, columnCount int) (top, bottom, left, right int) {
	look := conditional.look
	top, bottom, left, right = 0, rowCount-1, 0, columnCount-1
	switch kind {
	case "firstRow":
		top, bottom = 0, 0
	case "lastRow":
		top, bottom = rowCount-1, rowCount-1
	case "firstCol":
		left, right = 0, 0
	case "lastCol":
		left, right = columnCount-1, columnCount-1
	case "band1Horz", "band2Horz":
		start, end := 0, rowCount-1
		if look.firstRow {
			start = 1
		}
		if look.lastRow {
			end = rowCount - 2
		}
		group := (row - start) / conditional.rowBandSize
		top = start + group*conditional.rowBandSize
		bottom = top + conditional.rowBandSize - 1
		if bottom > end {
			bottom = end
		}
	case "band1Vert", "band2Vert":
		start, end := 0, columnCount-1
		if look.firstColumn {
			start = 1
		}
		if look.lastColumn {
			end = columnCount - 2
		}
		group := (column - start) / conditional.colBandSize
		left = start + group*conditional.colBandSize
		right = left + conditional.colBandSize - 1
		if right > end {
			right = end
		}
	case "nwCell", "neCell", "swCell", "seCell":
		top, bottom, left, right = row, row, column, column+span-1
	}
	return top, bottom, left, right
}

// applyRegionEdges overlays a region's stated edges on a cell's four edges
// according to where the cell sits in the region's box.
func applyRegionEdges(cell *nativeRegionBorders, region *nativeRegionBorders, onTop, onBottom, onLeft, onRight bool) {
	pick := func(outer, inside *NativeTableBorderV1, edgeOfRegion bool) *NativeTableBorderV1 {
		if edgeOfRegion {
			return outer
		}
		return inside
	}
	if border := pick(region.top, region.insideH, onTop); border != nil {
		cell.top = border
	}
	if border := pick(region.bottom, region.insideH, onBottom); border != nil {
		cell.bottom = border
	}
	if border := pick(region.left, region.insideV, onLeft); border != nil {
		cell.left = border
	}
	if border := pick(region.right, region.insideV, onRight); border != nil {
		cell.right = border
	}
}

// heavierTableBorder resolves one shared edge: the wider border wins, a
// stated none weighs nothing, and equal widths go to the darker colour.
func heavierTableBorder(a, b *NativeTableBorderV1) *NativeTableBorderV1 {
	weight := func(border *NativeTableBorderV1) int64 {
		if border == nil || border.Style != "single" {
			return 0
		}
		return border.SizeEighthPoints
	}
	darkness := func(border *NativeTableBorderV1) int {
		if border == nil || border.ColorRGB == nil {
			return 0
		}
		total := 0
		for i := 0; i+1 < len(*border.ColorRGB); i += 2 {
			var channel int
			for _, digit := range (*border.ColorRGB)[i : i+2] {
				channel = channel*16 + hexDigitValue(digit)
			}
			total += channel
		}
		return -total
	}
	wa, wb := weight(a), weight(b)
	if wa != wb {
		if wa > wb {
			return a
		}
		return b
	}
	if wa == 0 {
		if a != nil {
			return a
		}
		return b
	}
	if darkness(a) >= darkness(b) {
		return a
	}
	return b
}

func hexDigitValue(digit rune) int {
	switch {
	case digit >= '0' && digit <= '9':
		return int(digit - '0')
	case digit >= 'A' && digit <= 'F':
		return int(digit-'A') + 10
	case digit >= 'a' && digit <= 'f':
		return int(digit-'a') + 10
	}
	return 0
}

func paintedTableBorder(border *NativeTableBorderV1) *NativeTableBorderV1 {
	if border == nil {
		return nil
	}
	copied := *border
	if copied.Style != "single" {
		copied.Style, copied.SizeEighthPoints, copied.ColorRGB = "none", 0, nil
	}
	return &copied
}

// cellBorders resolves every cell's four edges. It returns nil when no region
// states a border and the table's own borders were already exact, when a cell
// is merged, or when a cell states its own w:tcBorders (cell conflict
// resolution against direct borders is not modeled).
func (conditional *nativeConditionalTableStyle) cellBorders(table *NativeTableV1, tableBorders *NativeTableBordersV1, tableBordersExact bool) []NativeResolvedTableCellBordersV1 {
	regionStated := false
	for _, kind := range nativeTableRegionOrder {
		if region := conditional.regions[kind]; region != nil && region.borders != nil {
			regionStated = true
		}
	}
	if whole := conditional.regions["wholeTable"]; whole != nil && whole.borders != nil {
		regionStated = true
	}
	if !regionStated && tableBordersExact || tableBorders == nil && !regionStated {
		return nil
	}
	rowCount := len(table.Rows)
	columnCount := len(table.GridWidthsTwips)
	type placed struct {
		row, column, span int
		cell              *NativeTableCellV1
		edges             nativeRegionBorders
	}
	cells := []*placed{}
	grid := map[[2]int]*placed{}
	for rowIndex := range table.Rows {
		column := 0
		for cellIndex := range table.Rows[rowIndex].Cells {
			cell := &table.Rows[rowIndex].Cells[cellIndex]
			if cell.VerticalMerge != "none" || cell.Borders != nil {
				return nil
			}
			span := 1
			if cell.GridSpan != nil {
				span = *cell.GridSpan
			}
			entry := &placed{row: rowIndex, column: column, span: span, cell: cell}
			cells = append(cells, entry)
			for offset := 0; offset < span; offset++ {
				grid[[2]int{rowIndex, column + offset}] = entry
			}
			column += span
		}
		if column > columnCount {
			columnCount = column
		}
	}
	if len(cells) == 0 || len(cells) > nativeMaxConditionalTableCells {
		return nil
	}
	base := regionBordersFromTable(tableBorders)
	for _, entry := range cells {
		edges := nativeRegionBorders{}
		// The table's own borders: outer edges on the outside, inside edges
		// between cells.
		applyRegionEdges(&edges, base, entry.row == 0, entry.row == rowCount-1, entry.column == 0, entry.column+entry.span == columnCount)
		if whole := conditional.regions["wholeTable"]; whole != nil && whole.borders != nil {
			applyRegionEdges(&edges, whole.borders, entry.row == 0, entry.row == rowCount-1, entry.column == 0, entry.column+entry.span == columnCount)
		}
		for _, kind := range conditional.cellRegions(entry.row, rowCount, entry.column, entry.span, columnCount) {
			region := conditional.regions[kind]
			if region.borders == nil {
				continue
			}
			top, bottom, left, right := conditional.regionExtent(kind, entry.row, rowCount, entry.column, entry.span, columnCount)
			applyRegionEdges(&edges, region.borders, entry.row == top, entry.row == bottom, entry.column == left, entry.column+entry.span-1 == right)
		}
		entry.edges = edges
	}
	// Shared edges: the right neighbour's left and the lower neighbour's top.
	for _, entry := range cells {
		if right := grid[[2]int{entry.row, entry.column + entry.span}]; right != nil {
			winner := heavierTableBorder(entry.edges.right, right.edges.left)
			entry.edges.right, right.edges.left = winner, winner
		}
		if below := grid[[2]int{entry.row + 1, entry.column}]; below != nil && below.column == entry.column && below.span == entry.span {
			winner := heavierTableBorder(entry.edges.bottom, below.edges.top)
			entry.edges.bottom, below.edges.top = winner, winner
		}
	}
	out := make([]NativeResolvedTableCellBordersV1, 0, len(cells))
	for _, entry := range cells {
		out = append(out, NativeResolvedTableCellBordersV1{CellID: entry.cell.ID, Borders: NativeTableBordersV1{Top: paintedTableBorder(entry.edges.top), Right: paintedTableBorder(entry.edges.right), Bottom: paintedTableBorder(entry.edges.bottom), Left: paintedTableBorder(entry.edges.left)}})
	}
	return out
}
