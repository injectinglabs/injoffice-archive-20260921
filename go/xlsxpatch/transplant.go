package xlsxpatch

import (
	"fmt"
	"strings"
)

// Chart transplant: the save-path bridge between an editor that rebuilds
// workbooks (Univer's client-side export) and files that contain charts the
// editor doesn't model. TransplantCharts re-creates every chart from the
// PREVIOUS file inside the REBUILT file — type, title, series bindings, and
// grid position — so "user edits cells and saves" no longer costs the charts.
//
// Fail-closed contract: if ANY chart cannot be faithfully transplanted (a
// type outside the writable family, an unparsable part, a series bound to a
// sheet the rebuilt file no longer has), the whole operation errors and the
// caller falls back to refusing the save. Partial preservation would be
// silent data loss with extra steps.
//
// Known v1 fidelity limits (documented, not silent): custom chart styling
// (explicit series colors, fonts) and stacked grouping are not yet carried
// over — the chart survives with default styling. Tracked in the roadmap.

// TransplantCharts returns dst with src's charts re-created inside it.
// src is the previous file (charts present); dst is the rebuilt file
// (charts absent). Neither input is modified.
func TransplantCharts(src, dst []byte) ([]byte, error) {
	charts, err := ReadCharts(src)
	if err != nil {
		return nil, err
	}
	if len(charts) == 0 {
		return dst, nil
	}
	anchors, err := ReadChartAnchors(src)
	if err != nil {
		return nil, err
	}

	out := dst
	for i, c := range charts {
		if !writableTypes[c.Type] {
			return nil, fmt.Errorf("xlsxpatch: transplant: chart %s has type %q outside the writable family", c.Part, c.Type)
		}
		if len(c.Series) == 0 {
			return nil, fmt.Errorf("xlsxpatch: transplant: chart %s has no series", c.Part)
		}
		sheetName, err := sheetNameFromRef(firstRef(c))
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: transplant: chart %s: %w", c.Part, err)
		}
		anchor, ok := anchors[c.Part]
		if !ok {
			// No resolvable anchor — cascade placement like a fresh insert.
			offset := i * 2
			anchor = ChartAnchor{FromCol: 5 + offset, FromRow: 2 + offset, ToCol: 13 + offset, ToRow: 18 + offset}
		}
		series := make([]WriteSeries, len(c.Series))
		for j, s := range c.Series {
			series[j] = WriteSeries{
				Name:          s.Name,
				NameRef:       s.NameRef,
				CategoriesRef: s.CategoriesRef,
				ValuesRef:     s.ValuesRef,
			}
		}
		next, err := AddChart(out, ChartWriteSpec{
			SheetName: sheetName,
			Type:      c.Type,
			Title:     c.Title,
			Series:    series,
			Anchor:    anchor,
		})
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: transplant: chart %s: %w", c.Part, err)
		}
		out = next
	}
	return out, nil
}

// firstRef picks the reference that identifies the chart's home sheet.
func firstRef(c ChartInfo) string {
	for _, s := range c.Series {
		if s.ValuesRef != "" {
			return s.ValuesRef
		}
		if s.CategoriesRef != "" {
			return s.CategoriesRef
		}
		if s.NameRef != "" {
			return s.NameRef
		}
	}
	return ""
}

// sheetNameFromRef extracts the sheet name from "Sheet!$A$1:$B$5" or
// "'My Sheet'!$A$1" forms.
func sheetNameFromRef(ref string) (string, error) {
	idx := strings.LastIndex(ref, "!")
	if idx <= 0 {
		return "", fmt.Errorf("series reference %q has no sheet qualifier", ref)
	}
	name := ref[:idx]
	if strings.HasPrefix(name, "'") && strings.HasSuffix(name, "'") && len(name) >= 2 {
		name = strings.ReplaceAll(name[1:len(name)-1], "''", "'")
	}
	if name == "" {
		return "", fmt.Errorf("series reference %q has empty sheet name", ref)
	}
	return name, nil
}
