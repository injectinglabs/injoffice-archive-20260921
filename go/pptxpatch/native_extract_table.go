package pptxpatch

import (
	"encoding/xml"
	"fmt"
	"strings"
)

const nativeDrawingTableURI = "http://schemas.openxmlformats.org/drawingml/2006/table"

// nativeGraphicFrameProjectionRefusal identifies a structurally readable
// graphic frame whose complete visual semantics are outside the exact native
// table subset. Callers preserve the whole p:graphicFrame subtree and never
// emit a partial table projection.
type nativeGraphicFrameProjectionRefusal struct {
	code    string
	message string
}

func (refusal nativeGraphicFrameProjectionRefusal) Error() string {
	return "pptxpatch: native extract: " + refusal.message
}

func refuseNativeGraphicFrame(code, message string) error {
	return nativeGraphicFrameProjectionRefusal{code: code, message: message}
}

type nativeExactTable struct {
	table         NativeTable
	outputNodes   int
	textCodeUnits int64
}

type nativeTableExtractionBudget struct {
	outputNodes   int
	tableCells    int
	textCodeUnits int64
}

type nativeTableOutputUsage struct {
	outputNodes   int
	tableCells    int
	textCodeUnits int64
}

// extractNativeTableGraphicFrame projects only a self-contained DrawingML
// table: explicit frame/grid geometry, explicit cell margins and top anchoring,
// self-contained text runs, explicit sRGB/no-fill cell paint, and explicit
// no-fill borders. Merges, table styles, inherited/theme paint, borders,
// effects, charts, and embedded-object frames are preserved as one opaque
// subtree instead of being approximated.
func (extractor *nativeExtractor) extractNativeGraphicFrame(node *nativeXMLNode, slidePart, slideID string, relationships []nativeExtractRelationship, dialect nativeExtractDialect) (NativeElement, error) {
	if node == nil || node.Name != (xml.Name{Space: dialect.presentation, Local: "graphicFrame"}) {
		return NativeElement{}, fmt.Errorf("pptxpatch: native extract: invalid graphic frame root")
	}
	if err := rejectNativeGraphicFrameDialectMix(node, dialect); err != nil {
		return NativeElement{}, err
	}
	graphic, err := nativeSingleton(node, dialect.drawing, "graphic", false)
	if err != nil {
		return NativeElement{}, err
	}
	if graphic != nil {
		if data, dataErr := nativeSingleton(graphic, dialect.drawing, "graphicData", false); dataErr != nil {
			return NativeElement{}, dataErr
		} else if data != nil {
			if uri, ok := exactNativeAttr(data, "", "uri"); ok && uri == dialect.chart {
				return extractor.extractNativeChartGraphicFrame(node, slidePart, slideID, relationships, dialect)
			}
		}
	}
	return extractor.extractNativeTableGraphicFrame(node, slidePart, dialect)
}

func (extractor *nativeExtractor) extractNativeTableGraphicFrame(node *nativeXMLNode, slidePart string, dialect nativeExtractDialect) (NativeElement, error) {
	if node == nil || node.Name != (xml.Name{Space: dialect.presentation, Local: "graphicFrame"}) {
		return NativeElement{}, fmt.Errorf("pptxpatch: native extract: invalid graphic frame root")
	}
	if err := rejectNativeGraphicFrameDialectMix(node, dialect); err != nil {
		return NativeElement{}, err
	}
	if err := requireOnlyNativeAttrs(node); err != nil {
		return NativeElement{}, refuseNativeGraphicFrame("pptx.table-markup-unavailable", "graphic frame root attributes are outside the exact native table subset")
	}
	if err := requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.presentation, Local: "nvGraphicFramePr"},
		xml.Name{Space: dialect.presentation, Local: "xfrm"},
		xml.Name{Space: dialect.drawing, Local: "graphic"}); err != nil {
		return NativeElement{}, refuseNativeGraphicFrame("pptx.table-markup-unavailable", "graphic frame contains unmodeled markup or direct text")
	}
	nonVisual, err := nativeSingleton(node, dialect.presentation, "nvGraphicFramePr", true)
	if err != nil {
		return NativeElement{}, err
	}
	transformNode, err := nativeSingleton(node, dialect.presentation, "xfrm", true)
	if err != nil {
		return NativeElement{}, err
	}
	graphic, err := nativeSingleton(node, dialect.drawing, "graphic", true)
	if err != nil {
		return NativeElement{}, err
	}
	if len(node.Children) != 3 || node.Children[0] != nonVisual || node.Children[1] != transformNode || node.Children[2] != graphic {
		return NativeElement{}, fmt.Errorf("pptxpatch: native extract: malformed graphic frame child order")
	}

	objectID, name, err := validateNativeTableNonVisual(nonVisual, dialect)
	if err != nil {
		return NativeElement{}, err
	}
	transform, err := validateNativeTableTransform(transformNode, dialect)
	if err != nil {
		return NativeElement{}, err
	}
	exact, err := extractor.extractNativeExactTable(graphic, transform, dialect, nativeTableExtractionBudget{
		outputNodes: nativeMaxNodes - extractor.outputNodesEmitted, tableCells: nativeMaxTableCells - extractor.tableCellsEmitted,
		textCodeUnits: int64(nativeMaxTotalTextCodeUnits) - extractor.textCodeUnitsEmitted,
	})
	if err != nil {
		return NativeElement{}, err
	}
	if err := extractor.reserveNativeTableOutput(exact); err != nil {
		return NativeElement{}, err
	}

	raw, err := rawNativeNode(extractor.pkg.parts[slidePart], node)
	if err != nil {
		return NativeElement{}, err
	}
	fingerprint := nativeSHA256(raw)
	elementID := extractor.identities.elements[nativeIdentityKey(slidePart, objectID)]
	if elementID == "" {
		elementID = stableNativeID("element", extractor.documentID+"\x00"+slidePart, objectID)
	}
	element := NativeElement{
		Kind: NativeElementKindTable, ID: elementID, Provenance: NativeProvenanceParsed,
		Transform: transform, Table: &exact.table, Passthrough: []NativePassthroughRef{}, Children: nil,
		Source:        &NativeSourceAnchor{PartName: slidePart, ObjectID: objectID, FingerprintSHA256: fingerprint},
		Compatibility: NativeCompatibility{Status: NativeCompatibilityStatusEditable, Diagnostics: []NativeDiagnostic{}},
	}
	if name != "" {
		element.Name = stringPointer(name)
	}
	return element, nil
}

func validateNativeTableNonVisual(node *nativeXMLNode, dialect nativeExtractDialect) (string, string, error) {
	if err := requireOnlyNativeAttrs(node); err != nil {
		return "", "", refuseNativeGraphicFrame("pptx.table-nonvisual-unavailable", "table nonvisual properties contain unmodeled attributes")
	}
	if err := requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.presentation, Local: "cNvPr"},
		xml.Name{Space: dialect.presentation, Local: "cNvGraphicFramePr"},
		xml.Name{Space: dialect.presentation, Local: "nvPr"}); err != nil {
		return "", "", refuseNativeGraphicFrame("pptx.table-nonvisual-unavailable", "table nonvisual properties are outside the exact subset")
	}
	cNvPr, err := nativeSingleton(node, dialect.presentation, "cNvPr", true)
	if err != nil {
		return "", "", err
	}
	cNvGraphicFramePr, err := nativeSingleton(node, dialect.presentation, "cNvGraphicFramePr", true)
	if err != nil {
		return "", "", err
	}
	nvPr, err := nativeSingleton(node, dialect.presentation, "nvPr", true)
	if err != nil {
		return "", "", err
	}
	if len(node.Children) != 3 || node.Children[0] != cNvPr || node.Children[1] != cNvGraphicFramePr || node.Children[2] != nvPr {
		return "", "", fmt.Errorf("pptxpatch: native extract: malformed table nonvisual child order")
	}
	if err := requireOnlyNativeAttrs(cNvPr, xml.Name{Local: "id"}, xml.Name{Local: "name"}); err != nil || len(cNvPr.Children) != 0 || !onlyNativeXMLSpace(cNvPr.Text) {
		return "", "", refuseNativeGraphicFrame("pptx.table-nonvisual-unavailable", "table visibility, hyperlink, accessibility, or extension metadata is not modeled")
	}
	nativeID, err := canonicalNativeUnsignedID(cNvPr, "", "id", 1)
	if err != nil {
		return "", "", err
	}
	name, _ := exactNativeAttr(cNvPr, "", "name")
	if utf16CodeUnitLengthBounded(name, 1025) > 1024 {
		return "", "", fmt.Errorf("pptxpatch: native extract: table name exceeds contract bound")
	}
	if requireEmptyNativeElement(cNvGraphicFramePr) != nil || requireEmptyNativeElement(nvPr) != nil {
		return "", "", refuseNativeGraphicFrame("pptx.table-inheritance-unavailable", "table locks, placeholder, or inherited nonvisual properties are not resolved")
	}
	return "cNvPr-" + nativeID, name, nil
}

func validateNativeTableTransform(node *nativeXMLNode, dialect nativeExtractDialect) (NativeTransform, error) {
	if err := requireOnlyNativeAttrs(node, xml.Name{Local: "rot"}, xml.Name{Local: "flipH"}, xml.Name{Local: "flipV"}); err != nil {
		return NativeTransform{}, refuseNativeGraphicFrame("pptx.table-transform-unavailable", "table transform contains unmodeled attributes")
	}
	if err := requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.drawing, Local: "off"},
		xml.Name{Space: dialect.drawing, Local: "ext"}); err != nil {
		return NativeTransform{}, fmt.Errorf("pptxpatch: native extract: invalid table transform children: %w", err)
	}
	off, err := requiredNativeTableSubsetChild(node, dialect.drawing, "off", "pptx.table-transform-unavailable", "table offset is inherited or omitted")
	if err != nil {
		return NativeTransform{}, err
	}
	ext, err := requiredNativeTableSubsetChild(node, dialect.drawing, "ext", "pptx.table-transform-unavailable", "table extent is inherited or omitted")
	if err != nil {
		return NativeTransform{}, err
	}
	if len(node.Children) != 2 || node.Children[0] != off || node.Children[1] != ext {
		return NativeTransform{}, fmt.Errorf("pptxpatch: native extract: malformed table transform child order")
	}
	if requireOnlyNativeAttrs(off, xml.Name{Local: "x"}, xml.Name{Local: "y"}) != nil || requireOnlyNativeChildren(off) != nil ||
		requireOnlyNativeAttrs(ext, xml.Name{Local: "cx"}, xml.Name{Local: "cy"}) != nil || requireOnlyNativeChildren(ext) != nil {
		return NativeTransform{}, fmt.Errorf("pptxpatch: native extract: invalid table transform geometry")
	}
	x, err := requiredCanonicalNativeTableInt(off, "x", -nativeMaxSafeInteger, nativeMaxSafeInteger)
	if err != nil {
		return NativeTransform{}, err
	}
	y, err := requiredCanonicalNativeTableInt(off, "y", -nativeMaxSafeInteger, nativeMaxSafeInteger)
	if err != nil {
		return NativeTransform{}, err
	}
	cx, err := requiredCanonicalNativeTableInt(ext, "cx", 1, nativeMaxSafeInteger)
	if err != nil {
		return NativeTransform{}, err
	}
	cy, err := requiredCanonicalNativeTableInt(ext, "cy", 1, nativeMaxSafeInteger)
	if err != nil {
		return NativeTransform{}, err
	}
	if raw, ok := exactNativeAttr(node, "", "rot"); ok {
		rotation, parseErr := parseCanonicalNativeInt(raw, -nativeMaxSafeInteger, nativeMaxSafeInteger)
		if parseErr != nil {
			return NativeTransform{}, fmt.Errorf("pptxpatch: native extract: invalid table rotation")
		}
		if rotation != 0 {
			return NativeTransform{}, refuseNativeGraphicFrame("pptx.table-transform-unavailable", "rotated tables are preserved but not approximated")
		}
	}
	for _, name := range []string{"flipH", "flipV"} {
		if raw, ok := exactNativeAttr(node, "", name); ok {
			flipped, boolErr := nativeBool(raw)
			if boolErr != nil {
				return NativeTransform{}, fmt.Errorf("pptxpatch: native extract: invalid table %s", name)
			}
			if flipped {
				return NativeTransform{}, refuseNativeGraphicFrame("pptx.table-transform-unavailable", "flipped tables are preserved but not approximated")
			}
		}
	}
	return NativeTransform{X: int64Pointer(x), Y: int64Pointer(y), Cx: int64Pointer(cx), Cy: int64Pointer(cy)}, nil
}

func requiredCanonicalNativeTableInt(node *nativeXMLNode, local string, minimum, maximum int64) (int64, error) {
	value, ok := exactNativeAttr(node, "", local)
	if !ok {
		return 0, fmt.Errorf("pptxpatch: native extract: missing table %s", local)
	}
	parsed, err := parseCanonicalNativeInt(value, minimum, maximum)
	if err != nil {
		return 0, fmt.Errorf("pptxpatch: native extract: invalid table %s", local)
	}
	return parsed, nil
}

func requiredNativeTableSubsetChild(node *nativeXMLNode, space, local, code, message string) (*nativeXMLNode, error) {
	child, err := nativeSingleton(node, space, local, false)
	if err != nil {
		return nil, err
	}
	if child == nil {
		return nil, refuseNativeGraphicFrame(code, message)
	}
	return child, nil
}

func (extractor *nativeExtractor) extractNativeExactTable(graphic *nativeXMLNode, transform NativeTransform, dialect nativeExtractDialect, budget nativeTableExtractionBudget) (nativeExactTable, error) {
	if requireOnlyNativeAttrs(graphic) != nil || requireOnlyNativeChildren(graphic, xml.Name{Space: dialect.drawing, Local: "graphicData"}) != nil {
		return nativeExactTable{}, refuseNativeGraphicFrame("pptx.graphic-frame-markup-unavailable", "graphic container is outside the exact native table subset")
	}
	data, err := nativeSingleton(graphic, dialect.drawing, "graphicData", true)
	if err != nil {
		return nativeExactTable{}, err
	}
	if requireOnlyNativeAttrs(data, xml.Name{Local: "uri"}) != nil {
		return nativeExactTable{}, refuseNativeGraphicFrame("pptx.graphic-frame-kind-unavailable", "graphic frame data has unmodeled type metadata")
	}
	uri, ok := exactNativeAttr(data, "", "uri")
	if !ok || uri != nativeDrawingTableURI {
		return nativeExactTable{}, refuseNativeGraphicFrame("pptx.graphic-frame-kind-unavailable", "charts and embedded-object graphic frames remain exact opaque content")
	}
	if err := requireOnlyNativeChildren(data, xml.Name{Space: dialect.drawing, Local: "tbl"}); err != nil {
		return nativeExactTable{}, refuseNativeGraphicFrame("pptx.table-markup-unavailable", "table graphic data contains unknown markup")
	}
	tableNode, err := nativeSingleton(data, dialect.drawing, "tbl", true)
	if err != nil {
		return nativeExactTable{}, err
	}
	return extractor.extractNativeTableNode(tableNode, transform, dialect, budget)
}

func (extractor *nativeExtractor) extractNativeTableNode(node *nativeXMLNode, transform NativeTransform, dialect nativeExtractDialect, budget nativeTableExtractionBudget) (nativeExactTable, error) {
	if requireOnlyNativeAttrs(node) != nil || requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.drawing, Local: "tblPr"},
		xml.Name{Space: dialect.drawing, Local: "tblGrid"},
		xml.Name{Space: dialect.drawing, Local: "tr"}) != nil {
		return nativeExactTable{}, refuseNativeGraphicFrame("pptx.table-markup-unavailable", "table contains merges, extensions, or unknown markup")
	}
	tableProperties, err := requiredNativeTableSubsetChild(node, dialect.drawing, "tblPr", "pptx.table-style-unavailable", "table properties are omitted and require inherited defaults")
	if err != nil {
		return nativeExactTable{}, err
	}
	grid, err := nativeSingleton(node, dialect.drawing, "tblGrid", true)
	if err != nil {
		return nativeExactTable{}, err
	}
	if len(node.Children) < 2 || node.Children[0] != tableProperties || node.Children[1] != grid {
		return nativeExactTable{}, fmt.Errorf("pptxpatch: native extract: malformed table child order")
	}
	if requireEmptyNativeElement(tableProperties) != nil {
		return nativeExactTable{}, refuseNativeGraphicFrame("pptx.table-style-unavailable", "table styles, banding, inheritance, fills, and effects are preserved but not resolved")
	}
	columns, columnTotal, err := extractNativeTableGrid(grid, dialect)
	if err != nil {
		return nativeExactTable{}, err
	}
	rows := nativeChildren(node, dialect.drawing, "tr")
	if len(rows) == 0 {
		return nativeExactTable{}, refuseNativeGraphicFrame("pptx.table-row-unavailable", "empty DrawingML tables are preserved but not projected")
	}
	if len(rows) > nativeMaxTableRows {
		return nativeExactTable{}, fmt.Errorf("pptxpatch: native extract: table row count exceeds the contract bound")
	}
	if len(columns) > nativeMaxTableColumns || len(columns) > nativeMaxTableCells/len(rows) {
		return nativeExactTable{}, fmt.Errorf("pptxpatch: native extract: table cell count exceeds the contract bound")
	}
	usage, err := preflightNativeTableOutput(rows, len(columns), dialect)
	if err != nil {
		return nativeExactTable{}, err
	}
	if usage.outputNodes > budget.outputNodes || usage.tableCells > budget.tableCells || usage.textCodeUnits > budget.textCodeUnits {
		return nativeExactTable{}, fmt.Errorf("pptxpatch: native extract: cumulative table output budget exceeded")
	}
	table := NativeTable{ColumnWidths: columns, RowHeights: make([]int64, 0, len(rows)), Rows: make([][]NativeTableCell, 0, len(rows))}
	var rowTotal int64
	for rowIndex, rowNode := range rows {
		row, height, rowErr := extractor.extractNativeTableRow(rowNode, len(columns), columns, rowIndex, dialect)
		if rowErr != nil {
			return nativeExactTable{}, rowErr
		}
		if rowTotal > nativeMaxSafeInteger-height {
			return nativeExactTable{}, fmt.Errorf("pptxpatch: native extract: table row geometry exceeds safe integer precision")
		}
		rowTotal += height
		table.RowHeights = append(table.RowHeights, height)
		table.Rows = append(table.Rows, row)
	}
	if transform.Cx == nil || transform.Cy == nil || columnTotal != *transform.Cx || rowTotal != *transform.Cy {
		return nativeExactTable{}, refuseNativeGraphicFrame("pptx.table-geometry-unavailable", "table frame extents must equal the exact row and column track totals")
	}
	return nativeExactTable{table: table, outputNodes: usage.outputNodes, textCodeUnits: usage.textCodeUnits}, nil
}

func extractNativeTableGrid(node *nativeXMLNode, dialect nativeExtractDialect) ([]int64, int64, error) {
	if requireOnlyNativeAttrs(node) != nil || requireOnlyNativeChildren(node, xml.Name{Space: dialect.drawing, Local: "gridCol"}) != nil {
		return nil, 0, refuseNativeGraphicFrame("pptx.table-grid-unavailable", "table grid contains unknown or inherited column semantics")
	}
	gridColumns := nativeChildren(node, dialect.drawing, "gridCol")
	if len(gridColumns) == 0 {
		return nil, 0, refuseNativeGraphicFrame("pptx.table-grid-unavailable", "empty DrawingML table grids are preserved but not projected")
	}
	if len(gridColumns) > nativeMaxTableColumns {
		return nil, 0, fmt.Errorf("pptxpatch: native extract: table column count exceeds the contract bound")
	}
	columns := make([]int64, 0, len(gridColumns))
	var total int64
	for _, column := range gridColumns {
		if requireOnlyNativeAttrs(column, xml.Name{Local: "w"}) != nil || requireOnlyNativeChildren(column) != nil {
			return nil, 0, refuseNativeGraphicFrame("pptx.table-grid-unavailable", "table grid column contains unmodeled markup")
		}
		width, err := requiredCanonicalNativeTableInt(column, "w", 1, nativeMaxSafeInteger)
		if err != nil {
			return nil, 0, err
		}
		if total > nativeMaxSafeInteger-width {
			return nil, 0, fmt.Errorf("pptxpatch: native extract: table column geometry exceeds safe integer precision")
		}
		total += width
		columns = append(columns, width)
	}
	return columns, total, nil
}

// preflightNativeTableOutput computes a conservative upper bound matching
// the JSON value-budget shape before NativeTableCell/paragraph/run objects are
// materialized. Exact cells always carry derived text, authoritative paragraphs,
// a complete text body, and at most one explicit fill.
func preflightNativeTableOutput(rows []*nativeXMLNode, columnCount int, dialect nativeExtractDialect) (nativeTableOutputUsage, error) {
	usage := nativeTableOutputUsage{}
	// NativeTable + its three slices + column/height scalars + row slices.
	nodes := 4 + columnCount + 2*len(rows)
	add := func(count int) error {
		if count < 0 || nodes > nativeMaxNodes-count {
			return fmt.Errorf("pptxpatch: native extract: table output node budget exceeded")
		}
		nodes += count
		return nil
	}
	for _, row := range rows {
		for _, cell := range nativeChildren(row, dialect.drawing, "tc") {
			usage.tableCells++
			// Cell + text + paragraphs slice + complete text body + optional fill.
			if err := add(14); err != nil {
				return nativeTableOutputUsage{}, err
			}
			textBody := nativeChild(cell, dialect.drawing, "txBody")
			if textBody == nil {
				continue
			}
			paragraphs := nativeChildren(textBody, dialect.drawing, "p")
			if len(paragraphs) > nativeMaxParagraphsPerElement {
				return nativeTableOutputUsage{}, fmt.Errorf("pptxpatch: native extract: table paragraph budget exceeded")
			}
			var cellTextUnits int64
			for paragraphIndex, paragraph := range paragraphs {
				if paragraphIndex != 0 {
					cellTextUnits++
				}
				// Paragraph + runs slice + align/level/bullet.
				if err := add(5); err != nil {
					return nativeTableOutputUsage{}, err
				}
				// Run + text/bold/italic/family/size/color.
				runs := nativeChildren(paragraph, dialect.drawing, "r")
				if len(runs) > nativeMaxRunsPerParagraph {
					return nativeTableOutputUsage{}, fmt.Errorf("pptxpatch: native extract: table run budget exceeded")
				}
				if err := add(7 * len(runs)); err != nil {
					return nativeTableOutputUsage{}, err
				}
				for _, run := range runs {
					text := nativeChild(run, dialect.drawing, "t")
					if text == nil {
						continue
					}
					units := int64(utf16CodeUnitLengthBounded(text.Text, nativeMaxTextCodeUnits+1))
					if units > nativeMaxTextCodeUnits || usage.textCodeUnits > int64(nativeMaxTotalTextCodeUnits)-units {
						return nativeTableOutputUsage{}, fmt.Errorf("pptxpatch: native extract: table text budget exceeded")
					}
					if cellTextUnits > int64(nativeMaxTextCodeUnits)-units {
						return nativeTableOutputUsage{}, fmt.Errorf("pptxpatch: native extract: flattened table cell text budget exceeded")
					}
					cellTextUnits += units
					usage.textCodeUnits += units
				}
			}
		}
	}
	usage.outputNodes = nodes
	return usage, nil
}

func validateNativeTableTextOrder(textBody *nativeXMLNode, dialect nativeExtractDialect) error {
	for _, child := range textBody.Children {
		if child.Name != (xml.Name{Space: dialect.drawing, Local: "bodyPr"}) && child.Name != (xml.Name{Space: dialect.drawing, Local: "lstStyle"}) && child.Name != (xml.Name{Space: dialect.drawing, Local: "p"}) {
			return nil
		}
	}
	bodyPr, err := nativeSingleton(textBody, dialect.drawing, "bodyPr", true)
	if err != nil {
		return err
	}
	lstStyle, err := nativeSingleton(textBody, dialect.drawing, "lstStyle", true)
	if err != nil {
		return err
	}
	paragraphs := nativeChildren(textBody, dialect.drawing, "p")
	if len(paragraphs) == 0 || len(textBody.Children) != len(paragraphs)+2 || textBody.Children[0] != bodyPr || textBody.Children[1] != lstStyle {
		return fmt.Errorf("pptxpatch: native extract: malformed table text-body child order")
	}
	for paragraphIndex, paragraph := range paragraphs {
		if textBody.Children[paragraphIndex+2] != paragraph {
			return fmt.Errorf("pptxpatch: native extract: malformed table paragraph order")
		}
		properties, propertyErr := requiredNativeTableSubsetChild(paragraph, dialect.drawing, "pPr", "pptx.table-text-unavailable", "table paragraph properties are omitted and require inheritance")
		if propertyErr != nil {
			return propertyErr
		}
		runs := nativeChildren(paragraph, dialect.drawing, "r")
		if len(runs) == 0 {
			return nil
		}
		if len(paragraph.Children) != len(runs)+1 || paragraph.Children[0] != properties {
			// Valid a:br/fld/endParaRPr semantics are handled as opaque refusals by
			// the shared text extractor; only a pure but reordered pPr/r sequence
			// is malformed here.
			for _, child := range paragraph.Children {
				if child.Name != (xml.Name{Space: dialect.drawing, Local: "pPr"}) && child.Name != (xml.Name{Space: dialect.drawing, Local: "r"}) {
					return nil
				}
			}
			return fmt.Errorf("pptxpatch: native extract: malformed table paragraph child order")
		}
		for runIndex, run := range runs {
			if paragraph.Children[runIndex+1] != run {
				return fmt.Errorf("pptxpatch: native extract: malformed table run order")
			}
			runProperties, runPropertyErr := requiredNativeTableSubsetChild(run, dialect.drawing, "rPr", "pptx.table-text-unavailable", "table run properties are omitted and require inheritance")
			if runPropertyErr != nil {
				return runPropertyErr
			}
			for _, child := range run.Children {
				if child.Name != (xml.Name{Space: dialect.drawing, Local: "rPr"}) && child.Name != (xml.Name{Space: dialect.drawing, Local: "t"}) {
					return nil
				}
			}
			text, textErr := nativeSingleton(run, dialect.drawing, "t", true)
			if textErr != nil {
				return textErr
			}
			if len(run.Children) != 2 || run.Children[0] != runProperties || run.Children[1] != text {
				return fmt.Errorf("pptxpatch: native extract: malformed table run child order")
			}
			for _, child := range runProperties.Children {
				if child.Name != (xml.Name{Space: dialect.drawing, Local: "solidFill"}) && child.Name != (xml.Name{Space: dialect.drawing, Local: "latin"}) {
					return nil
				}
			}
			fill, fillErr := requiredNativeTableSubsetChild(runProperties, dialect.drawing, "solidFill", "pptx.table-text-unavailable", "table run color is omitted and requires inheritance")
			if fillErr != nil {
				return fillErr
			}
			latin, latinErr := requiredNativeTableSubsetChild(runProperties, dialect.drawing, "latin", "pptx.table-text-unavailable", "table run typeface is omitted and requires inheritance")
			if latinErr != nil {
				return latinErr
			}
			if len(runProperties.Children) != 2 || runProperties.Children[0] != fill || runProperties.Children[1] != latin {
				return fmt.Errorf("pptxpatch: native extract: malformed table run-property child order")
			}
		}
	}
	return nil
}

func (extractor *nativeExtractor) extractNativeTableRow(node *nativeXMLNode, columnCount int, columns []int64, rowIndex int, dialect nativeExtractDialect) ([]NativeTableCell, int64, error) {
	if requireOnlyNativeAttrs(node, xml.Name{Local: "h"}) != nil || requireOnlyNativeChildren(node, xml.Name{Space: dialect.drawing, Local: "tc"}) != nil {
		return nil, 0, refuseNativeGraphicFrame("pptx.table-row-unavailable", "table row contains unknown markup")
	}
	height, err := requiredCanonicalNativeTableInt(node, "h", 1, nativeMaxSafeInteger)
	if err != nil {
		return nil, 0, err
	}
	cells := nativeChildren(node, dialect.drawing, "tc")
	if len(cells) != columnCount {
		return nil, 0, refuseNativeGraphicFrame("pptx.table-merge-unavailable", "merged or non-rectangular table rows are preserved but not projected")
	}
	result := make([]NativeTableCell, 0, columnCount)
	for columnIndex, cellNode := range cells {
		cell, cellErr := extractor.extractNativeTableCell(cellNode, columns[columnIndex], height, rowIndex, columnIndex, dialect)
		if cellErr != nil {
			return nil, 0, cellErr
		}
		result = append(result, cell)
	}
	return result, height, nil
}

func (extractor *nativeExtractor) extractNativeTableCell(node *nativeXMLNode, width, height int64, rowIndex, columnIndex int, dialect nativeExtractDialect) (NativeTableCell, error) {
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
	properties, err := requiredNativeTableSubsetChild(node, dialect.drawing, "tcPr", "pptx.table-cell-layout-unavailable", "table cell properties are omitted and require inherited layout and paint")
	if err != nil {
		return NativeTableCell{}, err
	}
	if len(node.Children) != 2 || node.Children[0] != textBody || node.Children[1] != properties {
		return NativeTableCell{}, fmt.Errorf("pptxpatch: native extract: malformed table cell child order")
	}
	if err := validateNativeTableTextOrder(textBody, dialect); err != nil {
		return NativeTableCell{}, err
	}
	paragraphs, err := extractor.extractNativeParagraphs(textBody, dialect)
	if err != nil {
		var duplicate nativeDuplicateSingletonError
		if isNativeDuplicateSingleton(err, &duplicate) {
			return NativeTableCell{}, err
		}
		return NativeTableCell{}, refuseNativeGraphicFrame("pptx.table-text-unavailable", fmt.Sprintf("table cell %d,%d text is not self-contained: %v", rowIndex, columnIndex, err))
	}
	layout, fill, err := extractNativeTableCellProperties(properties, textBody, width, height, dialect, extractor.theme)
	if err != nil {
		return NativeTableCell{}, err
	}
	text, err := flattenNativeTableParagraphs(paragraphs)
	if err != nil {
		return NativeTableCell{}, err
	}
	cell := NativeTableCell{Text: &text, Paragraphs: &paragraphs, TextBody: layout, Fill: fill}
	return cell, nil
}

func extractNativeTableCellProperties(node, textBody *nativeXMLNode, width, height int64, dialect nativeExtractDialect, theme nativeResolvedTheme) (*NativeTextBodyLayout, *string, error) {
	if requireOnlyNativeAttrs(node,
		xml.Name{Local: "marL"}, xml.Name{Local: "marR"}, xml.Name{Local: "marT"}, xml.Name{Local: "marB"},
		xml.Name{Local: "anchor"}, xml.Name{Local: "anchorCtr"}, xml.Name{Local: "horzOverflow"}, xml.Name{Local: "vert"}) != nil {
		return nil, nil, refuseNativeGraphicFrame("pptx.table-cell-layout-unavailable", "table cell margins, anchoring, or text flow are not fully explicit")
	}
	margins := map[string]int64{}
	for _, name := range []string{"marL", "marR", "marT", "marB"} {
		value, err := requiredCanonicalNativeTableInt(node, name, 0, nativeMaxTextInsetEMU)
		if err != nil {
			return nil, nil, refuseNativeGraphicFrame("pptx.table-cell-layout-unavailable", "table cell margins must be explicit nonnegative signed 32-bit EMU values")
		}
		margins[name] = value
	}
	anchor, anchorOK := exactNativeAttr(node, "", "anchor")
	anchorCenter, anchorCenterOK := exactNativeAttr(node, "", "anchorCtr")
	overflow, overflowOK := exactNativeAttr(node, "", "horzOverflow")
	vertical, verticalOK := exactNativeAttr(node, "", "vert")
	centered, boolErr := nativeBool(anchorCenter)
	if !anchorOK || anchor != "t" || !anchorCenterOK || boolErr != nil || centered || !overflowOK || overflow != "overflow" || !verticalOK || vertical != "horz" {
		return nil, nil, refuseNativeGraphicFrame("pptx.table-cell-layout-unavailable", "only explicit top-anchored horizontal overflow cell text is represented exactly")
	}
	if requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.drawing, Local: "lnL"},
		xml.Name{Space: dialect.drawing, Local: "lnR"},
		xml.Name{Space: dialect.drawing, Local: "lnT"},
		xml.Name{Space: dialect.drawing, Local: "lnB"},
		xml.Name{Space: dialect.drawing, Local: "noFill"},
		xml.Name{Space: dialect.drawing, Local: "solidFill"}) != nil {
		return nil, nil, refuseNativeGraphicFrame("pptx.table-cell-paint-unavailable", "table cell contains unsupported borders, fills, effects, or extension markup")
	}
	for _, name := range []string{"lnL", "lnR", "lnT", "lnB"} {
		line, err := requiredNativeTableSubsetChild(node, dialect.drawing, name, "pptx.table-border-unavailable", "table borders are omitted and require inherited line semantics")
		if err != nil {
			return nil, nil, err
		}
		if requireOnlyNativeAttrs(line) != nil || requireOnlyNativeChildren(line, xml.Name{Space: dialect.drawing, Local: "noFill"}) != nil {
			return nil, nil, refuseNativeGraphicFrame("pptx.table-border-unavailable", "visible, inherited, dashed, or nonuniform table borders remain opaque")
		}
		noFill, err := requiredNativeTableSubsetChild(line, dialect.drawing, "noFill", "pptx.table-border-unavailable", "table border paint is omitted and requires inherited line semantics")
		if err != nil {
			return nil, nil, err
		}
		if requireEmptyNativeElement(noFill) != nil {
			return nil, nil, refuseNativeGraphicFrame("pptx.table-border-unavailable", "table no-border markup is outside the exact subset")
		}
	}
	noFill, noFillErr := nativeSingleton(node, dialect.drawing, "noFill", false)
	solidFill, solidFillErr := nativeSingleton(node, dialect.drawing, "solidFill", false)
	if noFillErr != nil {
		return nil, nil, noFillErr
	}
	if solidFillErr != nil {
		return nil, nil, solidFillErr
	}
	if (noFill == nil) == (solidFill == nil) {
		return nil, nil, refuseNativeGraphicFrame("pptx.table-fill-unavailable", "table cell fill must be explicit no-fill or exact sRGB solid fill")
	}
	fillNode := noFill
	if fillNode == nil {
		fillNode = solidFill
	}
	if len(node.Children) != 5 || node.Children[0].Name.Local != "lnL" || node.Children[1].Name.Local != "lnR" || node.Children[2].Name.Local != "lnT" || node.Children[3].Name.Local != "lnB" || node.Children[4] != fillNode {
		return nil, nil, fmt.Errorf("pptxpatch: native extract: malformed table cell-property child order")
	}
	var fill *string
	if noFill != nil {
		if requireEmptyNativeElement(noFill) != nil {
			return nil, nil, refuseNativeGraphicFrame("pptx.table-fill-unavailable", "table cell no-fill markup is outside the exact subset")
		}
	} else {
		color, colorErr := exactNativeSolidColor(solidFill, dialect, theme)
		if colorErr != nil {
			return nil, nil, refuseNativeGraphicFrame("pptx.table-fill-unavailable", "transformed, gradient, pattern, picture, or inherited cell fills remain opaque")
		}
		fill = &color
	}
	bodyPr, err := nativeSingleton(textBody, dialect.drawing, "bodyPr", true)
	if err != nil {
		return nil, nil, err
	}
	if requireEmptyNativeElement(bodyPr) != nil {
		return nil, nil, refuseNativeGraphicFrame("pptx.table-cell-layout-unavailable", "table cell body properties contain unsupported wrap, rotation, autofit, 3D, or extension semantics")
	}
	if width-margins["marL"]-margins["marR"] <= 0 || height-margins["marT"]-margins["marB"] <= 0 {
		return nil, nil, refuseNativeGraphicFrame("pptx.table-cell-layout-unavailable", "table cell margins leave no positive text layout bounds")
	}
	layout := &NativeTextBodyLayout{
		LeftInsetEMU: int64Pointer(margins["marL"]), RightInsetEMU: int64Pointer(margins["marR"]),
		TopInsetEMU: int64Pointer(margins["marT"]), BottomInsetEMU: int64Pointer(margins["marB"]),
		Wrap: NativeTextWrapSquare, VerticalAnchor: NativeTextVerticalAnchorTop, AutoFit: "none",
		HorizontalOverflow: "overflow", VerticalOverflow: "overflow",
	}
	return layout, fill, nil
}

func flattenNativeTableParagraphs(paragraphs []NativeParagraph) (string, error) {
	var builder strings.Builder
	for paragraphIndex, paragraph := range paragraphs {
		if paragraphIndex != 0 {
			builder.WriteByte('\n')
		}
		for _, run := range paragraph.Runs {
			if run.Text == nil {
				return "", fmt.Errorf("pptxpatch: native extract: table run text is incomplete")
			}
			builder.WriteString(*run.Text)
			if builder.Len() > nativeMaxTextCodeUnits*4 {
				return "", fmt.Errorf("pptxpatch: native extract: flattened table text exceeds the bounded contract")
			}
		}
	}
	text := builder.String()
	if utf16CodeUnitLengthBounded(text, nativeMaxTextCodeUnits+1) > nativeMaxTextCodeUnits {
		return "", fmt.Errorf("pptxpatch: native extract: flattened table text exceeds the bounded contract")
	}
	return text, nil
}

func (extractor *nativeExtractor) reserveNativeTableOutput(exact nativeExactTable) error {
	rows, columns := len(exact.table.Rows), len(exact.table.ColumnWidths)
	if rows == 0 || columns == 0 || rows > nativeMaxTableRows || columns > nativeMaxTableColumns || columns > nativeMaxTableCells/rows {
		return fmt.Errorf("pptxpatch: native extract: table dimensions exceed the bounded native contract")
	}
	cells := rows * columns
	if cells > nativeMaxTableCells-extractor.tableCellsEmitted {
		return fmt.Errorf("pptxpatch: native extract: cumulative table cell budget exceeded")
	}
	nodes := exact.outputNodes
	if nodes <= 0 {
		return fmt.Errorf("pptxpatch: native extract: table output preflight is incomplete")
	}
	if nodes > nativeMaxNodes-extractor.outputNodesEmitted {
		return fmt.Errorf("pptxpatch: native extract: table output node budget exceeded")
	}
	var textUnits int64
	for _, row := range exact.table.Rows {
		for _, cell := range row {
			if cell.Paragraphs == nil {
				return fmt.Errorf("pptxpatch: native extract: authoritative table text is incomplete")
			}
			for _, paragraph := range *cell.Paragraphs {
				for _, run := range paragraph.Runs {
					if run.Text == nil {
						return fmt.Errorf("pptxpatch: native extract: authoritative table run text is incomplete")
					}
					units := int64(utf16CodeUnitLengthBounded(*run.Text, nativeMaxTextCodeUnits+1))
					if units > nativeMaxTextCodeUnits || textUnits > int64(nativeMaxTotalTextCodeUnits)-units {
						return fmt.Errorf("pptxpatch: native extract: table text budget exceeded")
					}
					textUnits += units
				}
			}
		}
	}
	if textUnits > int64(nativeMaxTotalTextCodeUnits)-extractor.textCodeUnitsEmitted {
		return fmt.Errorf("pptxpatch: native extract: cumulative table text budget exceeded")
	}
	if textUnits != exact.textCodeUnits {
		return fmt.Errorf("pptxpatch: native extract: table text preflight changed during projection")
	}
	extractor.tableCellsEmitted += cells
	extractor.outputNodesEmitted += nodes
	extractor.textCodeUnitsEmitted += textUnits
	return nil
}

func rejectNativeGraphicFrameDialectMix(node *nativeXMLNode, dialect nativeExtractDialect) error {
	opposite := map[string]bool{}
	if dialect.presentation == nsPresentationTransitional {
		opposite[nsPresentationStrict] = true
		opposite[nsDrawingStrict] = true
		opposite[nsOfficeRelsStrict] = true
		opposite[nsChartStrict] = true
	} else {
		opposite[nsPresentationTransitional] = true
		opposite[nsDrawingTransitional] = true
		opposite[nsOfficeRelsTransitional] = true
		opposite[nsChartTransitional] = true
	}
	var walk func(*nativeXMLNode) error
	walk = func(current *nativeXMLNode) error {
		if opposite[current.Name.Space] {
			return fmt.Errorf("pptxpatch: native extract: graphic frame mixes Strict and Transitional namespaces")
		}
		for _, attr := range current.Attrs {
			if attr.Name.Space != "xmlns" && attr.Name.Local != "xmlns" && opposite[attr.Name.Space] {
				return fmt.Errorf("pptxpatch: native extract: graphic frame mixes Strict and Transitional namespaces")
			}
		}
		for _, child := range current.Children {
			if err := walk(child); err != nil {
				return err
			}
		}
		return nil
	}
	return walk(node)
}
