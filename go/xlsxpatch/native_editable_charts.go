package xlsxpatch

import "strings"

type NativeEditableChartV1 struct {
	Identity          ChartIdentity         `json:"identity"`
	SheetID           string                `json:"sheet_id"`
	FingerprintSHA256 string                `json:"fingerprint_sha256"`
	ChartType         string                `json:"chart_type"`
	Title             string                `json:"title"`
	Range             StyleRange            `json:"range"`
	Anchor            NativeChartAnchorV1   `json:"anchor"`
	Editable          bool                  `json:"editable"`
	Refusal           string                `json:"refusal,omitempty"`
	Categories        []string              `json:"categories"`
	Series            []NativeChartSeriesV1 `json:"series"`
}

type NativeChartSeriesV1 struct {
	Name   string   `json:"name"`
	Values []string `json:"values"`
}

func projectNativeEditableChartsV1(data []byte) []NativeEditableChartV1 {
	charts, err := ReadCharts(data)
	if err != nil {
		return nil
	}
	pkg, err := openNativeWorkbookPackage(data)
	if err != nil {
		return nil
	}
	anchors, _ := ReadChartAnchors(data)
	workbook, err := ExtractNativeWorkbookV1(data)
	if err != nil {
		return nil
	}
	var out []NativeEditableChartV1
	for _, chart := range charts {
		if chart.Identity == nil {
			continue
		}
		e := NativeEditableChartV1{
			Identity: *chart.Identity, FingerprintSHA256: nativeWorkbookDigest(pkg.files[chart.Identity.Part]),
			ChartType: chart.Type, Title: chart.Title, Categories: []string{}, Series: []NativeChartSeriesV1{},
		}
		if chart.Type != "column" && chart.Type != "bar" && chart.Type != "line" && chart.Type != "pie" {
			e.ChartType, e.Refusal = "unsupported", "Chart type is preserved in the workbook but cannot be edited here."
		}
		if a, ok := anchors[chart.Identity.Part]; ok {
			e.Anchor = NativeChartAnchorV1{FromRow: a.FromRow, FromColumn: a.FromCol, ToRow: a.ToRow, ToColumn: a.ToCol}
		}
		sheet := nativeChartSheetForDrawing(workbook, pkg, chart.Identity.DrawingPart)
		if sheet == nil {
			e.Refusal = "Chart drawing has no worksheet owner."
			out = append(out, e)
			continue
		}
		e.SheetID = sheet.ID
		src, cats, series, reason := nativeChartQualify(sheet, chart)
		if reason != "" {
			if e.Refusal == "" {
				e.Refusal = reason
			}
			out = append(out, e)
			continue
		}
		e.Range, e.Categories, e.Series = src, cats, series
		e.Editable = e.ChartType != "unsupported" && e.Anchor != (NativeChartAnchorV1{})
		if !e.Editable && e.Refusal == "" {
			e.Refusal = "Chart is preserved in the workbook but cannot be edited here."
		}
		out = append(out, e)
	}
	return out
}

func nativeChartSheetForDrawing(workbook *NativeWorkbookV1, pkg *nativeWorkbookPackage, drawingPart string) *NativeWorkbookSheetV1 {
	for i := range workbook.Sheets {
		sheet := &workbook.Sheets[i]
		rels, ok := pkg.files[relsPartFor(sheet.PartName)]
		if !ok {
			continue
		}
		parsed, err := parsePackageRelationships(string(rels))
		if err != nil {
			continue
		}
		base := sheet.PartName[:strings.LastIndex(sheet.PartName, "/")]
		for _, rel := range parsed {
			if rel.Type != relTypeDrawing && !strings.HasSuffix(rel.Type, "/drawing") {
				continue
			}
			if target, err := resolveRelPath(base, rel.Target); err == nil && target == drawingPart {
				return sheet
			}
		}
	}
	return nil
}

func nativeChartQualify(sheet *NativeWorkbookSheetV1, chart ChartInfo) (StyleRange, []string, []NativeChartSeriesV1, string) {
	if len(chart.Series) == 0 {
		return StyleRange{}, nil, nil, "Chart has no series references."
	}
	prefix := sheet.Name + "!"
	r1, c1, r2, c2, err := nativeChartParseRef(chart.Series[0].CategoriesRef, prefix)
	if err != nil || c1 != c2 || r1 < 1 {
		return StyleRange{}, nil, nil, "Chart categories are not a single-column literal range on this sheet."
	}
	endCol := c1
	for i, s := range chart.Series {
		nr, nc, nr2, nc2, nerr := nativeChartParseRef(s.NameRef, prefix)
		vr, vc, vr2, vc2, verr := nativeChartParseRef(s.ValuesRef, prefix)
		if nerr != nil || verr != nil || nr != r1-1 || nr2 != r1-1 || nc != nc2 || vc != vc2 || vc != nc || vr != r1 || vr2 != r2 || nc != c1+1+i {
			return StyleRange{}, nil, nil, "Chart series are not a contiguous header row on this sheet."
		}
		if nc > endCol {
			endCol = nc
		}
	}
	src := StyleRange{Row: r1 - 1, Column: c1, EndRow: r2, EndColumn: endCol}
	built, err := nativeChartSeries(sheet, src, chart.Type)
	if err != nil {
		return src, nil, nil, err.Error()
	}
	cells := map[cellKey]*NativeWorkbookCellV1{}
	for i := range sheet.Cells {
		c := &sheet.Cells[i]
		cells[cellKey{row: c.Row, column: c.Column}] = c
	}
	var cats []string
	for row := src.Row + 1; row <= src.EndRow; row++ {
		c := cells[cellKey{row: row, column: src.Column}]
		text := ""
		if c != nil && c.Value != nil {
			if c.Value.Text != nil {
				text = *c.Value.Text
			} else if c.Value.Lexical != nil {
				text = *c.Value.Lexical
			}
		}
		cats = append(cats, text)
	}
	out := make([]NativeChartSeriesV1, len(built))
	for i, s := range built {
		vals := make([]string, len(cats))
		col := src.Column + 1 + i
		for row := src.Row + 1; row <= src.EndRow; row++ {
			c := cells[cellKey{row: row, column: col}]
			if c != nil && c.Value != nil && c.Value.Lexical != nil {
				vals[row-src.Row-1] = *c.Value.Lexical
			}
		}
		out[i] = NativeChartSeriesV1{Name: s.Name, Values: vals}
	}
	return src, cats, out, ""
}

func nativeChartParseRef(ref, prefix string) (int, int, int, int, error) {
	if !strings.HasPrefix(ref, prefix) {
		return 0, 0, 0, 0, errChartRef{}
	}
	return parseDimensionReference(strings.TrimPrefix(ref, prefix))
}

type errChartRef struct{}

func (errChartRef) Error() string { return "chart formula is not a local A1 range" }
