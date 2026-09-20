package docxpatch

import (
	"encoding/xml"
	"strconv"
	"strings"
)

// Conditional table-style regions (ECMA-376 17.7.6.6 w:tblStylePr), selected
// per table by w:tblLook (17.4.56) and resolved per cell.
//
// nativeTableRegionOrder is the precedence in which the regions a cell belongs
// to are applied, lowest first, so a later region's stated property overrides
// an earlier region's. The order was read off Microsoft Word 16.112.4's own
// raster of conditionalstyles-tbllook.docx, whose "Medium Shading 2" style gives
// all twelve regions a distinct fill: the vertical bands cover the horizontal
// bands (only band1Vert/band2Vert colours appear in the interior), the first
// and last column cover both bands, the first and last row cover the columns,
// and each corner cell covers everything under it. LibreOffice paints the same
// order. wholeTable is not in the list: it applies to every cell before any of
// these.
var nativeTableRegionOrder = []string{"band1Horz", "band2Horz", "band1Vert", "band2Vert", "firstCol", "lastCol", "firstRow", "lastRow", "nwCell", "neCell", "swCell", "seCell"}

type nativeTableLook struct {
	firstRow, lastRow, firstColumn, lastColumn, noHBand, noVBand bool
}

// One merged layer per region type: the chain's tblStylePr elements of that
// type merge base-first into a single style layer, so a toggle property the
// region states applies once (17.7.3), and the fill is the last one stated.
type nativeTableRegionLayer struct {
	definition *nativeStyleDefinition
	fill       *string
	borders    *nativeRegionBorders
}

type nativeConditionalTableStyle struct {
	look        nativeTableLook
	rowBandSize int
	colBandSize int
	regions     map[string]*nativeTableRegionLayer
	// The chain's w:tblBorders read as Word paints them (auto colour black,
	// theme tints from the authored colour); nil when a layer's borders could
	// not be read that way either.
	tableBorders           *NativeTableBordersV1
	tableBordersUnresolved bool
}

// NativeResolvedTableCellShadingV1 is the fill the conditional cascade resolves
// for one cell, after the table-level w:tcPr/w:shd and every region the cell
// belongs to have been applied in precedence order.
type NativeResolvedTableCellShadingV1 struct {
	CellID     string `json:"cell_id"`
	ShadingRGB string `json:"shading_rgb"`
}

const nativeMaxConditionalTableCells = 4096

// tableLook reads the table's own w:tblLook. Word writes the six switches
// twice, packed into w:val and as named attributes; when both forms are present
// they must agree. An absent or malformed element leaves the selection
// unresolved rather than assuming Word's authoring default.
func (resolver *nativeLayoutResolver) tableLook(table *NativeTableV1) (nativeTableLook, bool) {
	var look nativeTableLook
	node := resolver.nodeForAnchor(table.Anchor)
	if node == nil {
		return look, false
	}
	properties := directNativeChildren(node, resolver.wordNS, "tblPr")
	if len(properties) != 1 {
		return look, false
	}
	looks := directNativeChildren(properties[0], resolver.wordNS, "tblLook")
	if len(looks) != 1 {
		return look, false
	}
	element := looks[0]
	allowed := []xml.Name{{Space: resolver.wordNS, Local: "val"}}
	for _, switchBit := range nativeTableLookBits {
		allowed = append(allowed, xml.Name{Space: resolver.wordNS, Local: switchBit.attr})
	}
	if !nativeExactLeaf(element, allowed...) {
		return look, false
	}
	var mask uint64
	value, hasValue := nativeAttr(element, resolver.wordNS, "val")
	if hasValue {
		if len(value) == 0 || len(value) > 4 {
			return look, false
		}
		parsed, err := strconv.ParseUint(value, 16, 16)
		if err != nil || parsed & ^uint64(0x07e0) != 0 {
			return look, false
		}
		mask = parsed
	}
	named := false
	for _, switchBit := range nativeTableLookBits {
		if _, present := nativeAttr(element, resolver.wordNS, switchBit.attr); !present {
			continue
		}
		on, valid := nativeOnOffAttr(element, resolver.wordNS, switchBit.attr, false)
		if !valid || hasValue && on != (mask&switchBit.bit != 0) {
			return look, false
		}
		named = true
		if on {
			mask |= switchBit.bit
		}
	}
	if !hasValue && !named {
		return look, false
	}
	look.firstRow = mask&0x0020 != 0
	look.lastRow = mask&0x0040 != 0
	look.firstColumn = mask&0x0080 != 0
	look.lastColumn = mask&0x0100 != 0
	look.noHBand = mask&0x0200 != 0
	look.noVBand = mask&0x0400 != 0
	return look, true
}

var nativeTableRegionTypes = map[string]bool{"wholeTable": true, "band1Horz": true, "band2Horz": true, "band1Vert": true, "band2Vert": true, "firstCol": true, "lastCol": true, "firstRow": true, "lastRow": true, "nwCell": true, "neCell": true, "swCell": true, "seCell": true}

// resolveConditionalTableStyle resolves every w:tblStylePr in the chain into
// per-region layers. It returns ok=false when the selection or a region's
// structure cannot be read exactly, in which case the caller keeps the
// pre-existing refusal. A region property this tier does not model (cell
// borders, row properties, other cell properties) does not fail the
// resolution: it is disclosed at its own element under
// CONDITIONAL_TABLE_STYLE_PRESERVED so strict paint still refuses the table,
// while the fills and the paragraph/run cascade that were read exactly apply.
func (resolver *nativeLayoutResolver) resolveConditionalTableStyle(table *NativeTableV1, chain []*nativeStyleDefinition) (*nativeConditionalTableStyle, bool) {
	look, ok := resolver.tableLook(table)
	if !ok {
		return nil, false
	}
	conditional := &nativeConditionalTableStyle{look: look, rowBandSize: 1, colBandSize: 1, regions: map[string]*nativeTableRegionLayer{}}
	for _, layer := range chain {
		if tblPr := firstDirectNativeChild(layer.node, resolver.wordNS, "tblPr"); tblPr != nil {
			for _, borders := range directNativeChildren(tblPr, resolver.wordNS, "tblBorders") {
				parsed, ok := resolver.parseRegionBorders(borders)
				if !ok {
					conditional.tableBordersUnresolved = true
					continue
				}
				conditional.tableBorders = mergeNativeTableBorders(conditional.tableBorders, &NativeTableBordersV1{Top: parsed.top, Right: parsed.right, Bottom: parsed.bottom, Left: parsed.left, InsideHorizontal: parsed.insideH, InsideVertical: parsed.insideV})
			}
			for _, name := range []string{"tblStyleRowBandSize", "tblStyleColBandSize"} {
				sizes := directNativeChildren(tblPr, resolver.wordNS, name)
				if len(sizes) > 1 {
					return nil, false
				}
				if len(sizes) == 0 {
					continue
				}
				size, valid := nativePositiveIntAttr(sizes[0], resolver.wordNS, "val")
				if !valid || size > 1000 || !nativeExactLeaf(sizes[0], xml.Name{Space: resolver.wordNS, Local: "val"}) {
					return nil, false
				}
				if name == "tblStyleRowBandSize" {
					conditional.rowBandSize = size
				} else {
					conditional.colBandSize = size
				}
			}
		}
		seen := map[string]bool{}
		for _, region := range directNativeChildren(layer.node, resolver.wordNS, "tblStylePr") {
			if !nativeExactContainer(region, xml.Name{Space: resolver.wordNS, Local: "type"}) {
				return nil, false
			}
			kind, has := nativeAttr(region, resolver.wordNS, "type")
			if !has || !nativeTableRegionTypes[kind] || seen[kind] {
				return nil, false
			}
			seen[kind] = true
			if !resolver.mergeConditionalTableRegion(table, layer, kind, region, conditional) {
				return nil, false
			}
		}
	}
	return conditional, true
}

func (resolver *nativeLayoutResolver) mergeConditionalTableRegion(table *NativeTableV1, layer *nativeStyleDefinition, kind string, region *nativeXMLNode, conditional *nativeConditionalTableStyle) bool {
	merged := conditional.regions[kind]
	if merged == nil {
		merged = &nativeTableRegionLayer{definition: &nativeStyleDefinition{id: layer.id + ":" + kind, kind: "table", partName: layer.partName, node: region}}
		conditional.regions[kind] = merged
	}
	counts := nativeDirectWordChildCounts(region, resolver.wordNS)
	for _, child := range region.Children {
		if child.Name.Space != resolver.wordNS || counts[child.Name.Local] > 1 {
			return false
		}
		switch child.Name.Local {
		case "pPr":
			applyNativeParagraphProperties(&merged.definition.p, resolver.parseParagraphProperties(layer.partName, child, table.ID))
		case "rPr":
			applyNativeRunProperties(&merged.definition.r, resolver.parseRunProperties(layer.partName, child, table.ID), false)
		case "tblPr":
			// Word writes an empty w:tblPr in every region. Table properties a
			// region does state are not modeled for that region.
			if !nativeExactContainer(child) {
				return false
			}
			if len(child.Children) > 0 {
				resolver.addDiagnostic("CONDITIONAL_TABLE_STYLE_PRESERVED", table.ID, layer.partName, child, "Conditional table-style table properties are preserved and not applied to the region")
			}
		case "trPr":
			resolver.addDiagnostic("CONDITIONAL_TABLE_STYLE_PRESERVED", table.ID, layer.partName, child, "Conditional table-style row properties are preserved and not applied to the region")
		case "tcPr":
			if !nativeExactContainer(child) {
				return false
			}
			cellCounts := nativeDirectWordChildCounts(child, resolver.wordNS)
			for _, property := range child.Children {
				if property.Name.Space != resolver.wordNS || cellCounts[property.Name.Local] > 1 {
					return false
				}
				switch property.Name.Local {
				case "shd":
					fill, ok := nativeExtractCellShading(property, resolver.wordNS, resolver.resolveThemeSrgb)
					if !ok {
						// Word's built-in styles state region fills as a theme
						// colour with a tint or shade and also write the colour
						// that tint or shade produced into w:fill (17.3.5: the
						// theme attributes supersede it when a consumer re-themes).
						// The authored value is what Word painted, so it is
						// applied and the theme derivation is disclosed as not
						// recomputed.
						fill, ok = nativeAuthoredClearFill(property, resolver.wordNS)
						if !ok {
							return false
						}
						resolver.addDiagnostic("CONDITIONAL_TABLE_STYLE_PRESERVED", table.ID, layer.partName, property, "Conditional table-style theme fill tint or shade is not recomputed; the authored w:fill value is applied")
					}
					merged.fill = fill
				case "tcBorders":
					// Region borders are resolved per cell for the approximate
					// lane (native_table_conditional_borders.go); the exact tier
					// keeps refusing until shared-edge resolution is proven.
					parsed, ok := resolver.parseRegionBorders(property)
					if !ok {
						// A region the look never selects cannot reach a cell,
						// so its unreadable borders do not block the others.
						if conditional.regionSelectable(kind) {
							conditional.tableBordersUnresolved = true
						}
						resolver.addDiagnostic("CONDITIONAL_TABLE_STYLE_PRESERVED", table.ID, layer.partName, property, "Conditional table-style cell borders are outside the single/none subset and are not applied")
						continue
					}
					merged.borders = mergeRegionBorders(merged.borders, parsed)
					resolver.addDiagnostic("CONDITIONAL_TABLE_STYLE_PRESERVED", table.ID, layer.partName, property, "Conditional table-style cell borders are applied per cell by the approximate preview; exact shared-edge resolution is not proven")
				default:
					resolver.addDiagnostic("CONDITIONAL_TABLE_STYLE_PRESERVED", table.ID, layer.partName, property, "Conditional table-style cell property "+property.Name.Local+" is preserved and not applied to the region")
				}
			}
		default:
			return false
		}
	}
	return true
}

// cellRegions lists the regions a cell belongs to, in application order. Row
// banding starts after the first row when the first-row switch is on and stops
// before the last row when the last-row switch is on; column banding does the
// same with the first and last columns. Both offsets are visible in Word's
// raster of conditionalstyles-tbllook.docx (its second table's band1 fills sit
// on rows 1 and 3 and columns 1 and 3 with row 0 and column 0 excluded).
func (conditional *nativeConditionalTableStyle) cellRegions(rowIndex, rowCount, column, span, columnCount int) []string {
	look := conditional.look
	firstRow := look.firstRow && rowIndex == 0
	lastRow := look.lastRow && rowIndex == rowCount-1
	firstColumn := look.firstColumn && column == 0
	lastColumn := look.lastColumn && column+span == columnCount
	member := map[string]bool{}
	if !look.noHBand {
		start, end := 0, rowCount-1
		if look.firstRow {
			start = 1
		}
		if look.lastRow {
			end = rowCount - 2
		}
		if rowIndex >= start && rowIndex <= end {
			if ((rowIndex-start)/conditional.rowBandSize)%2 == 0 {
				member["band1Horz"] = true
			} else {
				member["band2Horz"] = true
			}
		}
	}
	if !look.noVBand {
		start, end := 0, columnCount-1
		if look.firstColumn {
			start = 1
		}
		if look.lastColumn {
			end = columnCount - 2
		}
		if column >= start && column <= end {
			if ((column-start)/conditional.colBandSize)%2 == 0 {
				member["band1Vert"] = true
			} else {
				member["band2Vert"] = true
			}
		}
	}
	member["firstCol"] = firstColumn
	member["lastCol"] = lastColumn
	member["firstRow"] = firstRow
	member["lastRow"] = lastRow
	member["nwCell"] = firstRow && firstColumn
	member["neCell"] = firstRow && lastColumn
	member["swCell"] = lastRow && firstColumn
	member["seCell"] = lastRow && lastColumn
	regions := []string{}
	for _, kind := range nativeTableRegionOrder {
		if member[kind] && conditional.regions[kind] != nil {
			regions = append(regions, kind)
		}
	}
	return regions
}

// cellCascade returns, for every cell, the style layers that apply to its
// paragraphs (the table style chain followed by the cell's regions) and the
// cell fills the cascade resolves. baseFill is the table-level w:tcPr/w:shd.
// It returns ok=false when the table's grid cannot be read.
func (conditional *nativeConditionalTableStyle) cellCascade(table *NativeTableV1, chain []*nativeStyleDefinition, baseFill *string) (map[string][]*nativeStyleDefinition, []NativeResolvedTableCellShadingV1, bool) {
	columnCount := len(table.GridWidthsTwips)
	for _, row := range table.Rows {
		width := 0
		for _, cell := range row.Cells {
			span := 1
			if cell.GridSpan != nil {
				span = *cell.GridSpan
			}
			if span < 1 {
				return nil, nil, false
			}
			width += span
		}
		if width > columnCount {
			columnCount = width
		}
	}
	if columnCount == 0 || len(table.Rows) == 0 {
		return nil, nil, false
	}
	whole := conditional.regions["wholeTable"]
	layers := map[string][]*nativeStyleDefinition{}
	fills := []NativeResolvedTableCellShadingV1{}
	total := 0
	for rowIndex, row := range table.Rows {
		column := 0
		for _, cell := range row.Cells {
			total++
			if total > nativeMaxConditionalTableCells {
				return nil, nil, false
			}
			span := 1
			if cell.GridSpan != nil {
				span = *cell.GridSpan
			}
			cellLayers := append([]*nativeStyleDefinition{}, chain...)
			fill := baseFill
			if whole != nil {
				cellLayers = append(cellLayers, whole.definition)
				if whole.fill != nil {
					fill = whole.fill
				}
			}
			for _, kind := range conditional.cellRegions(rowIndex, len(table.Rows), column, span, columnCount) {
				region := conditional.regions[kind]
				cellLayers = append(cellLayers, region.definition)
				if region.fill != nil {
					fill = region.fill
				}
			}
			layers[cell.ID] = cellLayers
			if fill != nil {
				fills = append(fills, NativeResolvedTableCellShadingV1{CellID: cell.ID, ShadingRGB: *fill})
			}
			column += span
		}
	}
	return layers, fills, true
}

// nativeAuthoredClearFill reads the explicit RGB of a clear w:shd whose theme
// attributes nativeExtractCellShading would not resolve.
func nativeAuthoredClearFill(node *nativeXMLNode, wordNS string) (*string, bool) {
	if !nativeExactLeaf(node,
		xml.Name{Space: wordNS, Local: "val"},
		xml.Name{Space: wordNS, Local: "color"},
		xml.Name{Space: wordNS, Local: "fill"},
		xml.Name{Space: wordNS, Local: "themeColor"},
		xml.Name{Space: wordNS, Local: "themeTint"},
		xml.Name{Space: wordNS, Local: "themeShade"},
		xml.Name{Space: wordNS, Local: "themeFill"},
		xml.Name{Space: wordNS, Local: "themeFillTint"},
		xml.Name{Space: wordNS, Local: "themeFillShade"},
	) {
		return nil, false
	}
	if value, ok := nativeAttr(node, wordNS, "val"); !ok || value != "clear" {
		return nil, false
	}
	if color, present := nativeAttr(node, wordNS, "color"); present && !strings.EqualFold(color, "auto") {
		return nil, false
	}
	fill, present := nativeAttr(node, wordNS, "fill")
	rgb, ok := nativeExactRGB(fill)
	if !present || !ok {
		return nil, false
	}
	return nativeString(rgb), true
}

// regionSelectable reports whether the table's look can select a region at all.
func (conditional *nativeConditionalTableStyle) regionSelectable(kind string) bool {
	look := conditional.look
	switch kind {
	case "firstRow":
		return look.firstRow
	case "lastRow":
		return look.lastRow
	case "firstCol":
		return look.firstColumn
	case "lastCol":
		return look.lastColumn
	case "band1Horz", "band2Horz":
		return !look.noHBand
	case "band1Vert", "band2Vert":
		return !look.noVBand
	case "nwCell":
		return look.firstRow && look.firstColumn
	case "neCell":
		return look.firstRow && look.lastColumn
	case "swCell":
		return look.lastRow && look.firstColumn
	case "seCell":
		return look.lastRow && look.lastColumn
	}
	return true
}
