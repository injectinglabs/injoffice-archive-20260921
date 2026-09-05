package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"fmt"
	"strings"
)

// Inspection is a cheap structural summary of an xlsx: what's in it that an
// editor might not model. Used by the gateway's overwrite guard ("does the
// file being replaced contain charts the replacement lost?") without parsing
// any chart XML.
type Inspection struct {
	EntryCount int
	// ChartParts are the chart XML part names (xl/charts/chart1.xml, ...).
	ChartParts []string
	// PivotParts are pivot table/cache part names.
	PivotParts []string
	// DrawingParts are drawing part names (anchors for charts/images/shapes).
	DrawingParts []string
	// MediaParts are embedded workbook media entries.
	MediaParts []string
	// OtherDrawingParts are VML, controls, ActiveX, and embedded-object parts
	// that may be referenced by an otherwise unknown drawing object.
	OtherDrawingParts []string
	HasImages         bool
}

// HasCharts reports whether the workbook contains chart parts.
func (i Inspection) HasCharts() bool { return len(i.ChartParts) > 0 }

// HasPivots reports whether the workbook contains pivot parts.
func (i Inspection) HasPivots() bool { return len(i.PivotParts) > 0 }

// Inspect summarizes an xlsx's structure. Errors only on unreadable zips —
// a valid zip that isn't a spreadsheet just yields an empty summary.
func Inspect(data []byte) (Inspection, error) {
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return Inspection{}, fmt.Errorf("xlsxpatch: inspect: %w", err)
	}
	var ins Inspection
	ins.EntryCount = len(zr.File)
	for _, f := range zr.File {
		name := f.Name
		switch {
		case strings.HasPrefix(name, "xl/charts/chart") && strings.HasSuffix(name, ".xml") && !strings.Contains(name, "Colors") && !strings.Contains(name, "Style"):
			ins.ChartParts = append(ins.ChartParts, name)
		case strings.HasPrefix(name, "xl/pivotTables/") || strings.HasPrefix(name, "xl/pivotCache/"):
			ins.PivotParts = append(ins.PivotParts, name)
		case strings.HasPrefix(name, "xl/drawings/") && strings.HasSuffix(name, ".xml"):
			ins.DrawingParts = append(ins.DrawingParts, name)
		case strings.HasPrefix(name, "xl/media/"):
			ins.HasImages = true
			ins.MediaParts = append(ins.MediaParts, name)
		case strings.HasPrefix(name, "xl/vmlDrawings/") || strings.HasPrefix(name, "xl/ctrlProps/") || strings.HasPrefix(name, "xl/activeX/") || strings.HasPrefix(name, "xl/embeddings/"):
			ins.OtherDrawingParts = append(ins.OtherDrawingParts, name)
		}
	}
	return ins, nil
}
