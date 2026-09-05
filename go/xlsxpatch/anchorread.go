package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"strconv"
	"strings"
)

// Anchor reading: where does each chart sit on the grid? Needed by the
// transplant path (re-creating a chart in a rebuilt workbook should keep its
// position) and by the browser bridge (rendering an agent-made chart at the
// same place it occupies in the file).
//
// The mapping goes drawing part → (twoCellAnchor from/to, chart rel id) →
// drawing rels → chart part name.

// ReadChartAnchors returns the grid anchor for each chart part that has one.
// Charts without a resolvable anchor are simply absent from the map — the
// caller falls back to a default placement.
func ReadChartAnchors(data []byte) (map[string]ChartAnchor, error) {
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: read anchors: %w", err)
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

	ins, err := Inspect(data)
	if err != nil {
		return nil, err
	}
	out := map[string]ChartAnchor{}
	for _, drawingPart := range ins.DrawingParts {
		drawingXML, ok := read(drawingPart)
		if !ok {
			continue
		}
		rels, ok := read(relsPartFor(drawingPart))
		if !ok {
			continue
		}
		for relID, anchor := range anchorsByChartRel(drawingXML) {
			chartPart, err := relTarget(rels, relID, drawingPart[:strings.LastIndex(drawingPart, "/")])
			if err != nil {
				continue
			}
			out[chartPart] = anchor
		}
	}
	return out, nil
}

// anchorsByChartRel parses one drawing document: for every two-cell anchor
// containing a chart graphicFrame, yields relID → anchor.
func anchorsByChartRel(drawingXML string) map[string]ChartAnchor {
	dec := xml.NewDecoder(strings.NewReader(drawingXML))
	out := map[string]ChartAnchor{}
	var stack []string
	var cur ChartAnchor
	var curRel string
	inFrom, inTo := false, false
	var pendingText *int

	setInt := func(dst *int, s string) {
		if n, err := strconv.Atoi(strings.TrimSpace(s)); err == nil {
			*dst = n
		}
	}
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return out
		}
		switch t := tok.(type) {
		case xml.StartElement:
			local := t.Name.Local
			stack = append(stack, local)
			switch local {
			case "twoCellAnchor", "oneCellAnchor", "absoluteAnchor":
				cur = ChartAnchor{}
				curRel = ""
			case "from":
				inFrom, inTo = true, false
			case "to":
				inFrom, inTo = false, true
			case "col":
				if inFrom {
					pendingText = &cur.FromCol
				} else if inTo {
					pendingText = &cur.ToCol
				}
			case "row":
				if inFrom {
					pendingText = &cur.FromRow
				} else if inTo {
					pendingText = &cur.ToRow
				}
			case "chart":
				for _, a := range t.Attr {
					if a.Name.Local == "id" {
						curRel = a.Value
					}
				}
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "from", "to":
				inFrom, inTo = false, false
			case "col", "row":
				pendingText = nil
			case "twoCellAnchor", "oneCellAnchor", "absoluteAnchor":
				if curRel != "" && cur.ToCol > cur.FromCol && cur.ToRow > cur.FromRow {
					out[curRel] = cur
				}
				curRel = ""
			}
			if len(stack) > 0 {
				stack = stack[:len(stack)-1]
			}
		case xml.CharData:
			if pendingText != nil {
				setInt(pendingText, string(t))
			}
		}
	}
	return out
}
