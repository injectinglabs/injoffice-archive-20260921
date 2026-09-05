package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
)

// Shape WRITING: turn an InjOffice ShapeSpec into a real OOXML drawing
// object inside the workbook, so shapes made in the browser exist in the
// .xlsx itself and open in Excel / Google Sheets / LibreOffice.
//
// Simpler than a chart: a shape is inline drawingML (<xdr:sp> / <xdr:cxnSp>)
// living directly in the drawing part — no separate chart part, no extra
// relationship, no content-type override beyond the drawing part itself.
// AddShape reuses the SAME drawing-part plumbing chartwrite.go already
// built (worksheetPartFor, drawingRelIDInSheet, appendAnchorToDrawing,
// insertDrawingElement, ...) — one drawing part per worksheet is the
// schema's rule, and a shape extends it exactly like a chart would.

// ShapeAnchor places the shape over the grid in cell coordinates — same
// twoCellAnchor semantics as ChartAnchor.
type ShapeAnchor struct {
	FromCol, FromRow, ToCol, ToRow int
}

// ShapeWriteSpec is what AddShape needs. Kind follows @injoffice/shapes'
// ShapeKind vocabulary (rect, ellipse, arrow, line, text, callout). Colors
// are #rrggbb or "" (transparent/none).
type ShapeWriteSpec struct {
	SheetName   string
	Kind        string
	Text        string
	Fill        string
	Stroke      string
	StrokeWidth float64 // EMU line width computed as StrokeWidth * 12700 (1pt); 0 = default (1pt)
	TextColor   string
	FontSize    float64 // points; 0 = default (14)
	Anchor      ShapeAnchor
}

// shapePresets is the curated shape catalogue: the OOXML/ECMA-376 preset
// geometry vocabulary is 182 entries total (verified against python-pptx's
// MSO_AUTO_SHAPE_TYPE, which mirrors it exactly) — most are obscure or
// duplicative for annotation purposes. S6.2 shipped the initial 80-shape
// subset; S6.3 (2026-08-25) added 39 more presets (rect-corner variants,
// misc basic-shape extras, 13 more flowchart symbols) — bringing this to
// 119. Grouped for the UI. Our shape KIND *is* the OOXML prst string
// directly (no translation table) for everything except the two structural
// specials: "line" (an xdr:cxnSp connector, not an xdr:sp) and "text" (an
// xdr:sp prst="rect" marked with the real OOXML txBox="1" idiom).
var shapePresets = map[string]bool{
	// Basic
	"rect": true, "roundRect": true, "ellipse": true, "triangle": true, "rtTriangle": true,
	"parallelogram": true, "trapezoid": true, "diamond": true, "pentagon": true, "hexagon": true,
	"heptagon": true, "octagon": true, "decagon": true, "dodecagon": true, "plus": true,
	"pie": true, "chord": true, "donut": true, "teardrop": true, "cube": true, "can": true,
	"heart": true, "lightningBolt": true, "cloud": true, "smileyFace": true,
	// Basic — rect-corner variants (S6.3)
	"round1Rect": true, "round2SameRect": true, "round2DiagRect": true,
	"snip1Rect": true, "snip2SameRect": true, "snip2DiagRect": true, "snipRoundRect": true,
	// Basic — extras (S6.3)
	"mathPlus": true, "frame": true, "halfFrame": true, "corner": true, "diagStripe": true,
	"noSmoking": true, "blockArc": true, "foldedCorner": true, "bevel": true, "sun": true,
	"moon": true, "arc": true, "plaque": true, "leftBracket": true, "rightBracket": true,
	"leftBrace": true, "rightBrace": true, "homePlate": true,
	// Arrows
	"rightArrow": true, "leftArrow": true, "upArrow": true, "downArrow": true,
	"leftRightArrow": true, "upDownArrow": true, "quadArrow": true, "bentArrow": true,
	"bentUpArrow": true, "uturnArrow": true, "leftUpArrow": true, "curvedRightArrow": true,
	"curvedLeftArrow": true, "curvedUpArrow": true, "curvedDownArrow": true,
	"stripedRightArrow": true, "notchedRightArrow": true, "chevron": true, "circularArrow": true,
	// Callouts
	"wedgeRectCallout": true, "wedgeRoundRectCallout": true, "wedgeEllipseCallout": true,
	"cloudCallout": true, "borderCallout1": true, "borderCallout2": true, "callout1": true,
	"callout2": true,
	// Stars & banners
	"star4": true, "star5": true, "star6": true, "star7": true, "star8": true, "star10": true,
	"star12": true, "star16": true, "star24": true, "star32": true, "ribbon": true,
	"ribbon2": true, "ellipseRibbon": true, "wave": true, "doubleWave": true,
	"irregularSeal1": true, "irregularSeal2": true,
	// Flowchart
	"flowChartProcess": true, "flowChartAlternateProcess": true, "flowChartDecision": true,
	"flowChartInputOutput": true, "flowChartPredefinedProcess": true,
	"flowChartInternalStorage": true, "flowChartDocument": true, "flowChartTerminator": true,
	"flowChartPreparation": true, "flowChartManualInput": true, "flowChartConnector": true,
	"flowChartSummingJunction": true,
	// Flowchart — S6.3
	"flowChartMultidocument": true, "flowChartManualOperation": true,
	"flowChartOffpageConnector": true, "flowChartMagneticDisk": true,
	"flowChartMagneticDrum": true, "flowChartDisplay": true, "flowChartDelay": true,
	"flowChartOr": true, "flowChartCollate": true, "flowChartSort": true,
	"flowChartExtract": true, "flowChartMerge": true, "flowChartPunchedTape": true,
}

// arrowFillKinds are the block-arrow family: rendered filled-with-stroke-
// color (no separate outline), matching how the browser draws them — a
// solid arrow glyph, not an outlined one.
var arrowFillKinds = map[string]bool{
	"rightArrow": true, "leftArrow": true, "upArrow": true, "downArrow": true,
	"leftRightArrow": true, "upDownArrow": true, "quadArrow": true, "bentArrow": true,
	"bentUpArrow": true, "uturnArrow": true, "leftUpArrow": true, "curvedRightArrow": true,
	"curvedLeftArrow": true, "curvedUpArrow": true, "curvedDownArrow": true,
	"stripedRightArrow": true, "notchedRightArrow": true, "chevron": true, "circularArrow": true,
}

func shapeWritable(kind string) bool {
	return kind == "line" || kind == "text" || shapePresets[kind]
}

// AddShape returns new workbook bytes containing the shape. The original is
// untouched on any error.
func AddShape(orig []byte, spec ShapeWriteSpec) ([]byte, error) {
	if err := validateShapeWriteSpec(spec); err != nil {
		return nil, fmt.Errorf("xlsxpatch: add shape: %w", err)
	}

	zr, err := zip.NewReader(bytes.NewReader(orig), int64(len(orig)))
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: add shape: %w", err)
	}
	read := func(name string) (string, bool) {
		for _, f := range zr.File {
			if f.Name == name {
				rc, err := f.Open()
				if err != nil {
					return "", false
				}
				defer rc.Close()
				b, err := io.ReadAll(rc)
				if err != nil {
					return "", false
				}
				return string(b), true
			}
		}
		return "", false
	}

	sheetPart, err := worksheetPartFor(read, spec.SheetName)
	if err != nil {
		return nil, err
	}
	sheetXML, ok := read(sheetPart)
	if !ok {
		return nil, fmt.Errorf("xlsxpatch: worksheet part %q unreadable", sheetPart)
	}

	patch := Patch{Replace: map[string][]byte{}, Add: map[string][]byte{}}
	objectID := uint32(1)
	anchor := ""

	sheetRelsPart := relsPartFor(sheetPart)
	sheetRels, hasSheetRels := read(sheetRelsPart)

	// Existing drawing on this sheet (a chart, an image, an earlier shape):
	// extend it. No new relationship needed — a shape is inline geometry,
	// not a reference to another part.
	if existingDrawingRelID := drawingRelIDInSheet(sheetXML); existingDrawingRelID != "" && hasSheetRels {
		drawingPart, err := relTarget(sheetRels, existingDrawingRelID, "xl/worksheets")
		if err != nil {
			return nil, err
		}
		drawingXML, ok := read(drawingPart)
		if !ok {
			return nil, fmt.Errorf("xlsxpatch: drawing part %q unreadable", drawingPart)
		}
		index, err := indexShapeDrawing([]byte(drawingXML))
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: add shape: drawing part %q: %w", drawingPart, err)
		}
		if index.maxObjectID == ^uint32(0) {
			return nil, fmt.Errorf("xlsxpatch: add shape: drawing part %q has exhausted DrawingML object ids", drawingPart)
		}
		objectID = index.maxObjectID + 1
		anchor = shapeAnchorXMLWithIdentity(spec, objectID, nil)
		newDrawingXML, err := appendAnchorToDrawing(drawingXML, anchor)
		if err != nil {
			return nil, err
		}
		patch.Replace[drawingPart] = []byte(newDrawingXML)
		return Apply(orig, patch)
	}

	// Fresh drawing path — identical shape to AddChart's, minus the
	// chart-part/chart-rels/chart-content-type pieces a bare shape needs
	// none of.
	drawingPart := nextFreePart(zr, "xl/drawings/drawing", ".xml")
	anchor = shapeAnchorXMLWithIdentity(spec, objectID, nil)
	patch.Add[drawingPart] = []byte(drawingXMLDoc(anchor))

	if !hasSheetRels {
		sheetRels = emptyRelsXML
	}
	drawingRelID := nextFreeRelID(sheetRels)
	newSheetRels, err := appendRelationship(sheetRels, drawingRelID, relTypeDrawing, partRelTargetFrom(sheetPart, drawingPart))
	if err != nil {
		return nil, err
	}
	if hasSheetRels {
		patch.Replace[sheetRelsPart] = []byte(newSheetRels)
	} else {
		patch.Add[sheetRelsPart] = []byte(newSheetRels)
	}

	newSheetXML, err := insertDrawingElement(sheetXML, drawingRelID)
	if err != nil {
		return nil, err
	}
	patch.Replace[sheetPart] = []byte(newSheetXML)

	ct, ok := read("[Content_Types].xml")
	if !ok {
		return nil, fmt.Errorf("xlsxpatch: missing [Content_Types].xml")
	}
	newCT, err := contentTypesWith(ct, map[string]string{"/" + drawingPart: ctDrawing})
	if err != nil {
		return nil, err
	}
	patch.Replace["[Content_Types].xml"] = []byte(newCT)

	return Apply(orig, patch)
}

// ---- XML generation ----

func emuLineWidth(pt float64) int {
	if pt <= 0 {
		pt = 1.5
	}
	return int(pt * 12700) // 1pt = 12700 EMU
}

func fillXML(hex string) string {
	if hex == "" {
		return `<a:noFill/>`
	}
	return fmt.Sprintf(`<a:solidFill><a:srgbClr val=%q/></a:solidFill>`, hexClr(hex))
}

func lineXML(strokeHex string, widthPt float64) string {
	if strokeHex == "" {
		return `<a:ln><a:noFill/></a:ln>`
	}
	return fmt.Sprintf(`<a:ln w="%d"><a:solidFill><a:srgbClr val=%q/></a:solidFill></a:ln>`, emuLineWidth(widthPt), hexClr(strokeHex))
}

// hexClr strips a leading # and upcases, since srgbClr@val takes bare hex.
func hexClr(s string) string {
	if len(s) > 0 && s[0] == '#' {
		s = s[1:]
	}
	out := make([]byte, 0, len(s))
	for _, c := range []byte(s) {
		if c >= 'a' && c <= 'f' {
			c -= 'a' - 'A'
		}
		out = append(out, c)
	}
	return string(out)
}

func shapeTxBodyXML(spec ShapeWriteSpec) string {
	if spec.Text == "" {
		return `<xdr:txBody><a:bodyPr wrap="square"/><a:p/></xdr:txBody>`
	}
	sz := spec.FontSize
	if sz <= 0 {
		sz = 14
	}
	color := spec.TextColor
	if color == "" {
		color = "1D2427"
	}
	return fmt.Sprintf(
		`<xdr:txBody><a:bodyPr wrap="square" anchor="ctr"/><a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="%d"><a:solidFill><a:srgbClr val=%q/></a:solidFill></a:rPr><a:t>%s</a:t></a:r></a:p></xdr:txBody>`,
		int(sz*100), hexClr(color), esc(spec.Text))
}

// shapeAnchorXML builds the twoCellAnchor wrapping one xdr:sp or xdr:cxnSp,
// mirroring chartwrite.go's anchorXML but with inline shape geometry
// instead of a graphicFrame referencing a separate chart part.
func shapeAnchorXMLWithIdentity(spec ShapeWriteSpec, objectID uint32, connectorProperties []byte) string {
	a := spec.Anchor
	name := esc(shapeDisplayName(spec))

	var body string
	if spec.Kind == "line" {
		cNvCxnSpPr := `<xdr:cNvCxnSpPr/>`
		if len(connectorProperties) > 0 {
			cNvCxnSpPr = string(connectorProperties)
		}
		body = fmt.Sprintf(
			`<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="%d" name=%q/>%s</xdr:nvCxnSpPr>`+
				`<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom>%s</xdr:spPr>`+
				`</xdr:cxnSp>`,
			objectID, name, cNvCxnSpPr, lineXML(spec.Stroke, spec.StrokeWidth))
	} else {
		// The shape's kind IS its OOXML prst value directly — "text" is
		// the one exception (always prst="rect", distinguished by the
		// txBox="1" marker below, not by its own preset name).
		prst := spec.Kind
		if spec.Kind == "text" {
			prst = "rect"
		}
		fill := spec.Fill
		if arrowFillKinds[spec.Kind] {
			// The browser renders the block-arrow family as filled-with-
			// stroke-color glyphs, not outlined shapes — match that.
			fill = spec.Stroke
		}
		// txBox="1" is the standard OOXML marker for "this rect is a plain
		// text box" (what Word/Excel/PowerPoint's own Insert Text Box
		// writes) — the unambiguous way to tell our "text" kind apart from
		// "rect" on read, since both use prst="rect".
		cNvSpPr := `<xdr:cNvSpPr/>`
		if spec.Kind == "text" {
			cNvSpPr = `<xdr:cNvSpPr txBox="1"/>`
		}
		body = fmt.Sprintf(
			`<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="%d" name=%q/>%s</xdr:nvSpPr>`+
				`<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst=%q><a:avLst/></a:prstGeom>%s%s</xdr:spPr>`+
				`%s</xdr:sp>`,
			objectID, name, cNvSpPr, prst, fillXML(fill), shapeLineXMLFor(spec), shapeTxBodyXML(spec))
	}

	return fmt.Sprintf(`<xdr:twoCellAnchor xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`+
		`<xdr:from><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>`+
		`<xdr:to><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>`+
		`%s<xdr:clientData/></xdr:twoCellAnchor>`,
		a.FromCol, a.FromRow, a.ToCol, a.ToRow, body)
}

// shapeLineXMLFor omits the outline for kinds the browser renders without a
// separate stroke pass on top of the fill (the arrow family uses
// fill-as-stroke above).
func shapeLineXMLFor(spec ShapeWriteSpec) string {
	if arrowFillKinds[spec.Kind] {
		return `<a:ln><a:noFill/></a:ln>`
	}
	return lineXML(spec.Stroke, spec.StrokeWidth)
}

func shapeDisplayName(spec ShapeWriteSpec) string {
	if spec.Text != "" {
		return spec.Text
	}
	return spec.Kind
}
