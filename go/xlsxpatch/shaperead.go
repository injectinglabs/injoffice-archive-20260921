package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"sort"
	"strconv"
	"strings"
)

// Shape READING: the Go-side mirror of @injoffice/shapes' ShapeSpec — enough
// to render an agent-generated (openpyxl, Excel, ...) shape in the browser
// layer, and to round-trip a browser-created one back after Save+reload.
//
// Unlike charts, a shape has no part of its own: it is inline drawingML
// (<xdr:sp>/<xdr:cxnSp>) living directly in a drawing part, so ShapeInfo
// carries its sheet name and anchor together rather than needing a separate
// anchor lookup.
//
// Kind IS the OOXML prst value (rect, chevron, star5, ...) except the two
// structural specials: "line" (an xdr:cxnSp) and "text" (prst="rect" with
// the txBox="1" marker — see shapewrite.go's shapePresets/arrowFillKinds).

// ShapeInfo is one shape found in the workbook.
type ShapeInfo struct {
	Identity    ShapeIdentity `json:"identity"`
	SheetName   string        `json:"sheetName"`
	Kind        string        `json:"kind"`
	Text        string        `json:"text,omitempty"`
	Fill        string        `json:"fill,omitempty"`
	Stroke      string        `json:"stroke,omitempty"`
	StrokeWidth float64       `json:"strokeWidth,omitempty"`
	TextColor   string        `json:"textColor,omitempty"`
	FontSize    float64       `json:"fontSize,omitempty"`
	Anchor      ShapeAnchor   `json:"anchor"`
}

// ShapeIdentity names one inline DrawingML object without relying on its
// display name or its position among sibling anchors. DrawingPart plus the
// workbook-issued cNvPr id remains stable when the object is moved or edited.
type ShapeIdentity struct {
	DrawingPart string `json:"drawingPart"`
	ObjectID    uint32 `json:"objectId"`
}

// ReadShapes parses every drawing part's top-level inline shapes/connectors.
// Malformed or duplicate DrawingML identities fail the read: returning an
// ambiguous identity would make a later update or delete unsafe. Grouped
// shapes remain opaque and are not presented as independently editable.
func ReadShapes(data []byte) ([]ShapeInfo, error) {
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: read shapes: %w", err)
	}
	ins, err := Inspect(data)
	if err != nil {
		return nil, err
	}
	if len(ins.DrawingParts) == 0 {
		return nil, nil
	}
	sheetNameByDrawingPart, err := sheetNamesByDrawingPart(zr)
	if err != nil {
		return nil, err
	}

	parts := append([]string(nil), ins.DrawingParts...)
	sort.Strings(parts)

	var out []ShapeInfo
	for _, part := range parts {
		sheetName := sheetNameByDrawingPart[part]
		if sheetName == "" {
			continue // orphaned drawing part (not hooked to any worksheet)
		}
		var file *zip.File
		for _, f := range zr.File {
			if f.Name == part {
				file = f
				break
			}
		}
		if file == nil {
			return nil, fmt.Errorf("xlsxpatch: drawing part %q is missing", part)
		}
		rc, err := file.Open()
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: open drawing part %q: %w", part, err)
		}
		drawing, readErr := io.ReadAll(rc)
		closeErr := rc.Close()
		if readErr != nil {
			return nil, fmt.Errorf("xlsxpatch: read drawing part %q: %w", part, readErr)
		}
		if closeErr != nil {
			return nil, fmt.Errorf("xlsxpatch: close drawing part %q: %w", part, closeErr)
		}
		if _, indexErr := indexShapeDrawing(drawing); indexErr != nil {
			return nil, fmt.Errorf("xlsxpatch: read shapes from %q: %w", part, indexErr)
		}
		shapes, perr := parseDrawingShapes(bytes.NewReader(drawing), sheetName, part)
		if perr != nil {
			return nil, fmt.Errorf("xlsxpatch: read shapes from %q: %w", part, perr)
		}
		out = append(out, shapes...)
	}
	return out, nil
}

// sheetNamesByDrawingPart resolves every worksheet's <drawing r:id> to the
// drawing part it points at, keyed by that part's path.
func sheetNamesByDrawingPart(zr *zip.Reader) (map[string]string, error) {
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
	wb, ok := read("xl/workbook.xml")
	if !ok {
		return nil, fmt.Errorf("xlsxpatch: missing xl/workbook.xml")
	}
	wbRels, ok := read("xl/_rels/workbook.xml.rels")
	if !ok {
		return nil, fmt.Errorf("xlsxpatch: missing workbook rels")
	}
	out := map[string]string{}
	dec := xml.NewDecoder(strings.NewReader(wb))
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: parse workbook.xml: %w", err)
		}
		se, isStart := tok.(xml.StartElement)
		if !isStart || se.Name.Local != "sheet" {
			continue
		}
		name, rid := "", ""
		for _, a := range se.Attr {
			switch a.Name.Local {
			case "name":
				name = a.Value
			case "id":
				rid = a.Value
			}
		}
		if name == "" || rid == "" {
			continue
		}
		sheetPart, err := relTarget(wbRels, rid, "xl")
		if err != nil {
			continue
		}
		sheetXML, ok := read(sheetPart)
		if !ok {
			continue
		}
		drawingRelID := drawingRelIDInSheet(sheetXML)
		if drawingRelID == "" {
			continue
		}
		sheetRels, ok := read(relsPartFor(sheetPart))
		if !ok {
			continue
		}
		drawingPart, err := relTarget(sheetRels, drawingRelID, "xl/worksheets")
		if err != nil {
			continue
		}
		if prior := out[drawingPart]; prior != "" && prior != name {
			return nil, fmt.Errorf("xlsxpatch: drawing part %q is owned by multiple worksheets", drawingPart)
		}
		out[drawingPart] = name
	}
	return out, nil
}

// parseDrawingShapes walks one drawing part's xdr:twoCellAnchor elements,
// emitting one ShapeInfo per xdr:sp/xdr:cxnSp found (graphicFrame anchors —
// charts — and pic anchors — images — are skipped, not ours to read here).
func parseDrawingShapes(r io.Reader, sheetName, drawingPart string) ([]ShapeInfo, error) {
	dec := xml.NewDecoder(r)
	var out []ShapeInfo
	var cur *ShapeInfo
	var stack []string
	anchorDepth := 0
	directShapeDepth := 0
	fromCol, fromRow, toCol, toRow := 0, 0, 0, 0
	var curCoord *int
	seenIDs := map[uint32]bool{}

	inside := func(name string) bool {
		for _, s := range stack {
			if s == name {
				return true
			}
		}
		return false
	}

	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return out, fmt.Errorf("xlsxpatch: parse drawing: %w", err)
		}
		switch t := tok.(type) {
		case xml.StartElement:
			local := t.Name.Local
			stack = append(stack, local)
			switch local {
			case "twoCellAnchor":
				anchorDepth++
				cur = nil
				fromCol, fromRow, toCol, toRow = 0, 0, 0, 0
			case "from":
				if anchorDepth > 0 {
					curCoord = nil
				}
			case "to":
				// handled via CharData below using inside()
			case "col":
				if inside("from") {
					curCoord = &fromCol
				} else if inside("to") {
					curCoord = &toCol
				}
			case "row":
				if inside("from") {
					curCoord = &fromRow
				} else if inside("to") {
					curCoord = &toRow
				}
			case "cxnSp":
				if len(stack) >= 2 && stack[len(stack)-2] == "twoCellAnchor" {
					cur = &ShapeInfo{SheetName: sheetName, Kind: "line", Identity: ShapeIdentity{DrawingPart: drawingPart}}
					directShapeDepth = len(stack)
				}
			case "sp":
				if len(stack) >= 2 && stack[len(stack)-2] == "twoCellAnchor" {
					cur = &ShapeInfo{SheetName: sheetName, Identity: ShapeIdentity{DrawingPart: drawingPart}}
					directShapeDepth = len(stack)
				}
			case "cNvPr":
				if cur != nil && directShapeDepth > 0 && len(stack) == directShapeDepth+2 && cur.Identity.ObjectID == 0 {
					raw := attrVal(t, "id")
					value, err := strconv.ParseUint(raw, 10, 32)
					if err != nil || value == 0 {
						return out, fmt.Errorf("shape has invalid cNvPr id %q", raw)
					}
					cur.Identity.ObjectID = uint32(value)
				}
			case "cNvSpPr":
				if cur != nil && attrVal(t, "txBox") == "1" {
					cur.Kind = "text"
				}
			case "prstGeom":
				// The shape's kind IS its prst value, straight through
				// (see shapewrite.go's shapePresets) — "rect" is the one
				// ambiguous case (also how "text" is written), and the
				// cNvSpPr case above already resolved that one first when
				// applicable, so cur.Kind == "" here means "genuinely not
				// yet known". A prst outside our curated catalogue is
				// still surfaced (as itself) rather than dropped — the
				// browser falls back to a generic labeled box for a kind
				// it doesn't have real geometry for, same as an unknown
				// chart type stays visible as "unknown" rather than
				// vanishing.
				if cur != nil && cur.Kind == "" {
					cur.Kind = attrVal(t, "prst")
				}
			case "solidFill":
				// handled on the following srgbClr
			case "srgbClr":
				if cur == nil {
					break
				}
				v := "#" + strings.ToLower(attrVal(t, "val"))
				switch {
				case inside("rPr"):
					cur.TextColor = v
				case inside("ln"):
					cur.Stroke = v
				case inside("spPr"):
					cur.Fill = v
				}
			case "ln":
				if cur != nil {
					if w := attrVal(t, "w"); w != "" {
						if n, err := strconv.Atoi(w); err == nil {
							cur.StrokeWidth = float64(n) / 12700
						}
					}
				}
			case "rPr":
				if cur != nil {
					if sz := attrVal(t, "sz"); sz != "" {
						if n, err := strconv.Atoi(sz); err == nil {
							cur.FontSize = float64(n) / 100
						}
					}
				}
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "sp", "cxnSp":
				if cur != nil && len(stack) == directShapeDepth {
					cur.Anchor = ShapeAnchor{FromCol: fromCol, FromRow: fromRow, ToCol: toCol, ToRow: toRow}
					if cur.Identity.ObjectID == 0 {
						return out, fmt.Errorf("shape has no cNvPr identity")
					}
					if seenIDs[cur.Identity.ObjectID] {
						return out, fmt.Errorf("duplicate shape cNvPr id %d", cur.Identity.ObjectID)
					}
					seenIDs[cur.Identity.ObjectID] = true
					if cur.Kind != "" {
						out = append(out, *cur)
					}
					cur = nil
					directShapeDepth = 0
				}
			case "twoCellAnchor":
				anchorDepth--
			}
			if len(stack) > 0 {
				stack = stack[:len(stack)-1]
			}
		case xml.CharData:
			text := strings.TrimSpace(string(t))
			if text == "" {
				continue
			}
			if curCoord != nil && (stack[len(stack)-1] == "col" || stack[len(stack)-1] == "row") {
				if n, err := strconv.Atoi(text); err == nil {
					*curCoord = n
				}
				curCoord = nil
			}
			if cur != nil && inside("txBody") && stack[len(stack)-1] == "t" {
				cur.Text += text
			}
		}
	}
	return out, nil
}
