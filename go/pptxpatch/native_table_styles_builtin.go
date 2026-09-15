package pptxpatch

import (
	"encoding/xml"
	"fmt"
	"math"
	"strings"
)

// Built-in PowerPoint table styles are referenced by GUID from a:tableStyleId
// and are usually absent from ppt/tableStyles.xml. This file resolves a small,
// bounded catalog of those definitions into cell fills and borders as an
// explicitly labeled read-only preview. Nothing here is qualified Office
// fidelity: text styling (tcTxStyle), per-edge border overrides, effects and
// every GUID outside the catalog still refuse.
//
// Sources for the catalog:
//   - ECMA-376 Part 1 (5th ed.) §20.1.4.2.27 tblStyleLst names
//     {5C22544A-7EE6-4342-B048-85BDC9FD1C3A} "Medium Style 2 - Accent 1" as the
//     default table style. §20.1.4.2.1 band1H (accent1 tint 40%), §20.1.4.2.11
//     firstCol, §20.1.4.2.12 firstRow, §20.1.4.2.16 lastCol and §20.1.4.2.17
//     lastRow (solid accent1 fills) and §20.1.4.2.6/.14/.15/.18/.22 (single
//     12700 EMU cell borders) document the region definitions of this family.
//   - The whole-table fill (accent tint 20%, reused by band2H/band2V) and the
//     lt1 border color follow the serialization PowerPoint writes for this
//     style into ppt/tableStyles.xml; the no-accent variant
//     {073A0DAA-6AF3-43AB-8588-CEC1D06C72B9} "Medium Style 2" substitutes dk1
//     for accent1.
//   - ECMA-376 Part 1 §21.1.3.15 tblPr defines the firstRow/firstCol/lastRow/
//     lastCol/bandRow/bandCol switches (default false) and §20.1.4.2.34 wholeTbl
//     as the formatting used when no switch applies.
//   - ECMA-376 Part 1 §20.1.2.3.34 defines tint as a mix with white but not the
//     color space. This preview mixes in linear sRGB with standard sRGB
//     encoding (the same policy as linear-srgb-path-tone-20-40-v1) and declares
//     it through nativeBuiltinTableStylePolicy.
//
// No table style definitions were copied from other implementations.

const nativeBuiltinTableStylePreviewCode = "pptx.table-builtin-style-preview"
const nativeBuiltinTableStyleTextCode = "pptx.table-style-text-unavailable"
const nativeBuiltinTableStylePolicy = "builtin-table-style-catalog-linear-srgb-tint-v1"

// ECMA-376 Part 1 §20.1.4.2.6/.14/.15/.18/.22: single 1pt cell borders.
const nativeBuiltinTableStyleBorderWidthEMU = int64(12700)

// nativeMaxTableStylesPartBytes bounds the untrusted ppt/tableStyles.xml scan.
const nativeMaxTableStylesPartBytes = 128 * 1024

type nativeBuiltinTableStyleFill struct {
	slot string
	// tint is an ST_PositiveFixedPercentage in 1/100000 units; 0 means none.
	tint int64
}

type nativeBuiltinTableStyle struct {
	id         string
	name       string
	borderSlot string
	wholeTbl   nativeBuiltinTableStyleFill
	band1H     nativeBuiltinTableStyleFill
	band2H     nativeBuiltinTableStyleFill
	band1V     nativeBuiltinTableStyleFill
	band2V     nativeBuiltinTableStyleFill
	firstRow   nativeBuiltinTableStyleFill
	lastRow    nativeBuiltinTableStyleFill
	firstCol   nativeBuiltinTableStyleFill
	lastCol    nativeBuiltinTableStyleFill
}

// nativeMediumStyle2 builds one member of the "Medium Style 2" family with the
// given accent scheme slot.
func nativeMediumStyle2(id, name, accent string) nativeBuiltinTableStyle {
	solid := nativeBuiltinTableStyleFill{slot: accent}
	light := nativeBuiltinTableStyleFill{slot: accent, tint: 20000}
	medium := nativeBuiltinTableStyleFill{slot: accent, tint: 40000}
	return nativeBuiltinTableStyle{
		id: id, name: name, borderSlot: "lt1",
		wholeTbl: light, band1H: medium, band2H: light, band1V: medium, band2V: light,
		firstRow: solid, lastRow: solid, firstCol: solid, lastCol: solid,
	}
}

// nativeBuiltinTableStyles is keyed by upper-case GUID.
var nativeBuiltinTableStyles = map[string]nativeBuiltinTableStyle{
	"{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}": nativeMediumStyle2("{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}", "Medium Style 2", "dk1"),
	"{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}": nativeMediumStyle2("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}", "Medium Style 2 - Accent 1", "accent1"),
}

type nativeTableStyleFlags struct {
	firstRow, lastRow, firstCol, lastCol, bandRow, bandCol bool
}

type nativeResolvedBuiltinTableStyle struct {
	style  nativeBuiltinTableStyle
	flags  nativeTableStyleFlags
	border NativeTableBorder
	fills  map[nativeBuiltinTableStyleFill]string
}

func refuseNativeTableStyle(message string) error {
	return refuseNativeGraphicFrame("pptx.table-style-unavailable", message)
}

// parseNativeTableStyleFlags accepts only the ECMA-376 §21.1.3.15 conditional
// formatting switches plus one a:tableStyleId child. Table-level fills,
// effects, inline a:tableStyle definitions, extensions and rtl tables refuse.
func parseNativeTableStyleFlags(node *nativeXMLNode, dialect nativeExtractDialect) (nativeTableStyleFlags, string, error) {
	flags := nativeTableStyleFlags{}
	if requireOnlyNativeAttrs(node,
		xml.Name{Local: "rtl"}, xml.Name{Local: "firstRow"}, xml.Name{Local: "firstCol"},
		xml.Name{Local: "lastRow"}, xml.Name{Local: "lastCol"}, xml.Name{Local: "bandRow"}, xml.Name{Local: "bandCol"}) != nil {
		return flags, "", refuseNativeTableStyle("table properties carry attributes outside the table-style switches")
	}
	for _, attr := range node.Attrs {
		if attr.Name.Space == "xmlns" || attr.Name.Local == "xmlns" {
			continue
		}
		value, err := nativeBool(attr.Value)
		if err != nil {
			return flags, "", refuseNativeTableStyle("table-style switch is not a canonical boolean")
		}
		switch attr.Name.Local {
		case "rtl":
			if value {
				return flags, "", refuseNativeTableStyle("right-to-left tables are preserved but not projected")
			}
		case "firstRow":
			flags.firstRow = value
		case "firstCol":
			flags.firstCol = value
		case "lastRow":
			flags.lastRow = value
		case "lastCol":
			flags.lastCol = value
		case "bandRow":
			flags.bandRow = value
		case "bandCol":
			flags.bandCol = value
		}
	}
	if requireOnlyNativeChildren(node, xml.Name{Space: dialect.drawing, Local: "tableStyleId"}) != nil {
		return flags, "", refuseNativeTableStyle("table styles, banding, inheritance, fills, and effects are preserved but not resolved")
	}
	styleNode, err := nativeSingleton(node, dialect.drawing, "tableStyleId", true)
	if err != nil {
		return flags, "", refuseNativeTableStyle("table style identifier is missing or duplicated")
	}
	if hasNativeSemanticAttrs(styleNode) || len(styleNode.Children) != 0 || !inspectionUUID.MatchString(styleNode.Text) {
		return flags, "", refuseNativeTableStyle("table style identifier is not one exact GUID")
	}
	return flags, strings.ToUpper(styleNode.Text), nil
}

// nativeTableStylesPartDefines reports whether the presentation's optional
// tableStyles part carries an explicit definition for styleID. A malformed,
// oversized or ambiguous part is an error so callers fail closed.
func (extractor *nativeExtractor) nativeTableStylesPartDefines(styleID string, dialect nativeExtractDialect) (bool, error) {
	roots, err := extractor.parseRelationships("")
	if err != nil {
		return false, err
	}
	office, err := uniqueNativeInternalRelationship(roots, dialect.rels+"/officeDocument", "presentation")
	if err != nil {
		return false, err
	}
	relationships, err := extractor.parseRelationships(office.Part)
	if err != nil {
		return false, err
	}
	var found *nativeExtractRelationship
	for index := range relationships {
		relationship := &relationships[index]
		if relationship.Type != dialect.rels+"/tableStyles" {
			continue
		}
		if found != nil || !relationship.internal() {
			return false, fmt.Errorf("table styles relationship must be unique and internal")
		}
		found = relationship
	}
	if found == nil {
		return false, nil
	}
	payload, ok := extractor.pkg.parts[found.Part]
	if !ok {
		return false, fmt.Errorf("table styles part %q is missing", found.Part)
	}
	if len(payload) > nativeMaxTableStylesPartBytes {
		return false, fmt.Errorf("table styles part exceeds the bounded scan size")
	}
	root, err := parseNativeXML(payload, found.Part)
	if err != nil {
		return false, err
	}
	if root.Name != (xml.Name{Space: dialect.drawing, Local: "tblStyleLst"}) || requireOnlyNativeChildren(root, xml.Name{Space: dialect.drawing, Local: "tblStyle"}) != nil {
		return false, fmt.Errorf("table styles part is not one exact style list")
	}
	for _, style := range root.Children {
		id, ok := exactNativeAttr(style, "", "styleId")
		if !ok || !inspectionUUID.MatchString(id) {
			return false, fmt.Errorf("table style definition lacks an exact GUID")
		}
		if strings.ToUpper(id) == styleID {
			return true, nil
		}
	}
	return false, nil
}

// resolveNativeBuiltinTableStyle turns a non-empty a:tblPr into a resolved
// catalog style or a projection refusal. Unknown GUIDs and GUIDs the package
// defines explicitly refuse: this preview never guesses a definition.
func (extractor *nativeExtractor) resolveNativeBuiltinTableStyle(node *nativeXMLNode, dialect nativeExtractDialect) (*nativeResolvedBuiltinTableStyle, error) {
	flags, styleID, err := parseNativeTableStyleFlags(node, dialect)
	if err != nil {
		return nil, err
	}
	style, ok := nativeBuiltinTableStyles[styleID]
	if !ok {
		return nil, refuseNativeTableStyle("table style " + styleID + " is outside the bounded built-in catalog and is preserved but not resolved")
	}
	defined, err := extractor.nativeTableStylesPartDefines(styleID, dialect)
	if err != nil {
		return nil, refuseNativeTableStyle("table styles part could not be qualified: " + err.Error())
	}
	if defined {
		return nil, refuseNativeTableStyle("table style " + styleID + " is defined explicitly by the package; authored style definitions are preserved but not resolved")
	}
	resolved := &nativeResolvedBuiltinTableStyle{style: style, flags: flags, fills: map[nativeBuiltinTableStyleFill]string{}}
	borderColor, err := resolved.resolveFill(nativeBuiltinTableStyleFill{slot: style.borderSlot}, extractor.theme)
	if err != nil {
		return nil, refuseNativeTableStyle("table style border color has no exact theme snapshot")
	}
	resolved.border = NativeTableBorder{Color: borderColor, WidthEMU: int64Pointer(nativeBuiltinTableStyleBorderWidthEMU)}
	for _, fill := range []nativeBuiltinTableStyleFill{style.wholeTbl, style.band1H, style.band2H, style.band1V, style.band2V, style.firstRow, style.lastRow, style.firstCol, style.lastCol} {
		if _, err := resolved.resolveFill(fill, extractor.theme); err != nil {
			return nil, refuseNativeTableStyle("table style fill color has no exact theme snapshot")
		}
	}
	return resolved, nil
}

func (resolved *nativeResolvedBuiltinTableStyle) resolveFill(fill nativeBuiltinTableStyleFill, theme nativeResolvedTheme) (string, error) {
	if cached, ok := resolved.fills[fill]; ok {
		return cached, nil
	}
	base, err := theme.resolveSchemeColor(fill.slot)
	if err != nil {
		return "", err
	}
	color, err := parseNativeSRGBHex(base)
	if err != nil {
		return "", err
	}
	if fill.tint != 0 {
		color = applyNativeLinearSRGBTint(color, fill.tint)
	}
	resolved.fills[fill] = color.hex()
	return color.hex(), nil
}

// regionFill selects the ECMA-376 §20.1.4.2 table part for one cell. Whole
// table first, then horizontal and vertical bands over body cells, then the
// first/last column and finally the first/last row, matching PowerPoint's
// row-over-column precedence. Corner parts (nwCell etc.) are not defined by
// this family, so no corner override exists.
func (resolved *nativeResolvedBuiltinTableStyle) regionFill(row, column, rows, columns int) nativeBuiltinTableStyleFill {
	style, flags := resolved.style, resolved.flags
	fill := style.wholeTbl
	bodyRow := 0
	if flags.firstRow {
		bodyRow = 1
	}
	if flags.bandRow && row >= bodyRow {
		if (row-bodyRow)%2 == 0 {
			fill = style.band1H
		} else {
			fill = style.band2H
		}
	}
	bodyColumn := 0
	if flags.firstCol {
		bodyColumn = 1
	}
	if flags.bandCol && column >= bodyColumn {
		if (column-bodyColumn)%2 == 0 {
			fill = style.band1V
		} else {
			fill = style.band2V
		}
	}
	if flags.firstCol && column == 0 {
		fill = style.firstCol
	}
	if flags.lastCol && column == columns-1 {
		fill = style.lastCol
	}
	if flags.lastRow && row == rows-1 {
		fill = style.lastRow
	}
	if flags.firstRow && row == 0 {
		fill = style.firstRow
	}
	return fill
}

func (resolved *nativeResolvedBuiltinTableStyle) cellFill(row, column, rows, columns int) (string, bool) {
	color, ok := resolved.fills[resolved.regionFill(row, column, rows, columns)]
	return color, ok
}

func (resolved *nativeResolvedBuiltinTableStyle) diagnostic() NativeDiagnostic {
	return NativeDiagnostic{
		Severity: NativeDiagnosticSeverityWarning, Code: nativeBuiltinTableStylePreviewCode,
		Message: "built-in table style " + resolved.style.id + " (" + resolved.style.name + ") is resolved from the bounded catalog with policy " + nativeBuiltinTableStylePolicy + "; cell fills and uniform 1pt borders are read-only preview, table-style text and per-edge border overrides are not projected",
	}
}

// applyNativeLinearSRGBTint mixes the color with white in linear sRGB
// (IEC 61966-2-1 transfer) and re-encodes with final 8-bit rounding. It is the
// declared preview policy for catalog fills, not a claim about Office's
// unspecified tint color space.
func applyNativeLinearSRGBTint(color nativeSRGBColor, tint int64) nativeSRGBColor {
	amount := float64(nativeClampInt64(tint, 0, nativePositiveFixedPct)) / float64(nativePositiveFixedPct)
	mix := func(channel int64) int64 {
		linear := nativeSRGBDecode(float64(nativeClampInt64(channel, 0, 255)) / 255)
		toned := linear*amount + (1 - amount)
		return nativeClampInt64(int64(math.Round(nativeSRGBEncode(toned)*255)), 0, 255)
	}
	return nativeSRGBColor{r: mix(color.r), g: mix(color.g), b: mix(color.b)}
}

func nativeSRGBDecode(value float64) float64 {
	if value <= 0.04045 {
		return value / 12.92
	}
	return math.Pow((value+0.055)/1.055, 2.4)
}

func nativeSRGBEncode(value float64) float64 {
	if value <= 0.0031308 {
		return 12.92 * value
	}
	return 1.055*math.Pow(value, 1/2.4) - 0.055
}

// extractNativeBuiltinStyledTable projects the rows of a table whose paint
// comes from a resolved catalog style. Cells must be text-free: table-style
// text inherits tcTxStyle and theme typography that this preview does not
// model, so any run, break or field refuses instead of guessing metrics.
func (extractor *nativeExtractor) extractNativeBuiltinStyledTable(rows []*nativeXMLNode, columns []int64, resolved *nativeResolvedBuiltinTableStyle, usage nativeTableOutputUsage, dialect nativeExtractDialect) (nativeExactTable, error) {
	table := NativeTable{ColumnWidths: columns, RowHeights: make([]int64, 0, len(rows)), Rows: make([][]NativeTableCell, 0, len(rows))}
	var rowTotal int64
	for rowIndex, rowNode := range rows {
		if requireOnlyNativeAttrs(rowNode, xml.Name{Local: "h"}) != nil || requireOnlyNativeChildren(rowNode, xml.Name{Space: dialect.drawing, Local: "tc"}) != nil {
			return nativeExactTable{}, refuseNativeGraphicFrame("pptx.table-row-unavailable", "table row contains unknown markup")
		}
		height, err := requiredCanonicalNativeTableInt(rowNode, "h", 1, nativeMaxSafeInteger)
		if err != nil {
			return nativeExactTable{}, err
		}
		cells := nativeChildren(rowNode, dialect.drawing, "tc")
		if len(cells) != len(columns) {
			return nativeExactTable{}, refuseNativeGraphicFrame("pptx.table-merge-unavailable", "merged or non-rectangular table rows are preserved but not projected")
		}
		row := make([]NativeTableCell, 0, len(columns))
		for columnIndex, cellNode := range cells {
			cell, cellErr := extractNativeBuiltinStyledTableCell(cellNode, rowIndex, columnIndex, len(rows), len(columns), resolved, dialect)
			if cellErr != nil {
				return nativeExactTable{}, cellErr
			}
			row = append(row, cell)
		}
		if rowTotal > nativeMaxSafeInteger-height {
			return nativeExactTable{}, fmt.Errorf("pptxpatch: native extract: table row geometry exceeds safe integer precision")
		}
		rowTotal += height
		table.RowHeights = append(table.RowHeights, height)
		table.Rows = append(table.Rows, row)
	}
	diagnostic := resolved.diagnostic()
	return nativeExactTable{table: table, outputNodes: usage.outputNodes, textCodeUnits: 0, legacyCells: true, styleDiagnostic: &diagnostic}, nil
}

func extractNativeBuiltinStyledTableCell(node *nativeXMLNode, rowIndex, columnIndex, rows, columns int, resolved *nativeResolvedBuiltinTableStyle, dialect nativeExtractDialect) (NativeTableCell, error) {
	if hasNativeSemanticAttrs(node) {
		for _, attr := range node.Attrs {
			switch attr.Name.Local {
			case "gridSpan", "rowSpan", "hMerge", "vMerge":
				return NativeTableCell{}, refuseNativeGraphicFrame("pptx.table-merge-unavailable", "merged table cells are preserved but not projected")
			}
		}
		return NativeTableCell{}, refuseNativeGraphicFrame("pptx.table-cell-unavailable", "table cell attributes are outside the exact native subset")
	}
	if requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.drawing, Local: "txBody"},
		xml.Name{Space: dialect.drawing, Local: "tcPr"}) != nil {
		return NativeTableCell{}, refuseNativeGraphicFrame("pptx.table-cell-unavailable", "table cell contains extension or unknown markup")
	}
	textBody, err := requiredNativeTableSubsetChild(node, dialect.drawing, "txBody", "pptx.table-text-unavailable", "table cell text body is omitted and requires inherited content")
	if err != nil {
		return NativeTableCell{}, err
	}
	properties, err := nativeSingleton(node, dialect.drawing, "tcPr", false)
	if err != nil {
		return NativeTableCell{}, err
	}
	if node.Children[0] != textBody || (properties != nil && (len(node.Children) != 2 || node.Children[1] != properties)) || (properties == nil && len(node.Children) != 1) {
		return NativeTableCell{}, fmt.Errorf("pptxpatch: native extract: malformed table cell child order")
	}
	if err := requireNativeStyledTableCellTextFree(textBody, rowIndex, columnIndex, dialect); err != nil {
		return NativeTableCell{}, err
	}
	if properties != nil {
		// ECMA-376 Part 1 §21.1.3.17 / CT_TableCellProperties: layout attributes
		// have schema defaults and cannot change the paint of a text-free cell.
		if requireOnlyNativeAttrs(properties,
			xml.Name{Local: "marL"}, xml.Name{Local: "marR"}, xml.Name{Local: "marT"}, xml.Name{Local: "marB"},
			xml.Name{Local: "anchor"}, xml.Name{Local: "anchorCtr"}, xml.Name{Local: "horzOverflow"}, xml.Name{Local: "vert"}) != nil {
			return NativeTableCell{}, refuseNativeGraphicFrame("pptx.table-cell-layout-unavailable", "table cell properties carry attributes outside the exact native subset")
		}
		if requireOnlyNativeChildren(properties) != nil {
			return NativeTableCell{}, refuseNativeTableStyle("explicit cell borders, fills or effects layered over a table style are preserved but not resolved")
		}
	}
	fill, ok := resolved.cellFill(rowIndex, columnIndex, rows, columns)
	if !ok {
		return NativeTableCell{}, fmt.Errorf("pptxpatch: native extract: table style fill was not resolved")
	}
	border := resolved.border
	return NativeTableCell{Text: stringPointer(""), Fill: stringPointer(fill), Border: &NativeTableBorder{Color: border.Color, WidthEMU: int64Pointer(*border.WidthEMU)}}, nil
}

// requireNativeStyledTableCellTextFree accepts only bodyPr, lstStyle and
// paragraphs that carry no text content. Paragraph and end-paragraph
// properties are tolerated because they cannot paint without runs.
func requireNativeStyledTableCellTextFree(textBody *nativeXMLNode, rowIndex, columnIndex int, dialect nativeExtractDialect) error {
	if hasNativeSemanticAttrs(textBody) || requireOnlyNativeChildren(textBody,
		xml.Name{Space: dialect.drawing, Local: "bodyPr"},
		xml.Name{Space: dialect.drawing, Local: "lstStyle"},
		xml.Name{Space: dialect.drawing, Local: "p"}) != nil {
		return refuseNativeGraphicFrame("pptx.table-cell-layout-unavailable", "table cell text body contains unmodeled markup")
	}
	bodyPr, err := nativeSingleton(textBody, dialect.drawing, "bodyPr", true)
	if err != nil {
		return err
	}
	listStyle, err := nativeSingleton(textBody, dialect.drawing, "lstStyle", true)
	if err != nil {
		return err
	}
	paragraphs := nativeChildren(textBody, dialect.drawing, "p")
	if len(paragraphs) == 0 || len(textBody.Children) != len(paragraphs)+2 || textBody.Children[0] != bodyPr || textBody.Children[1] != listStyle {
		return fmt.Errorf("pptxpatch: native extract: malformed table text-body child order")
	}
	if requireEmptyNativeElement(bodyPr) != nil || requireEmptyNativeElement(listStyle) != nil {
		return refuseNativeGraphicFrame("pptx.table-cell-layout-unavailable", "table cell body properties or list styles are not resolved under a table style")
	}
	if len(paragraphs) > nativeMaxParagraphsPerElement {
		return fmt.Errorf("pptxpatch: native extract: table paragraph budget exceeded")
	}
	for _, paragraph := range paragraphs {
		if hasNativeSemanticAttrs(paragraph) || requireOnlyNativeChildren(paragraph,
			xml.Name{Space: dialect.drawing, Local: "pPr"},
			xml.Name{Space: dialect.drawing, Local: "endParaRPr"}) != nil {
			return refuseNativeGraphicFrame(nativeBuiltinTableStyleTextCode, fmt.Sprintf("table cell %d,%d text inherits table-style typography that is not projected; styled tables are previewed only when every cell is text-free", rowIndex, columnIndex))
		}
	}
	return nil
}
