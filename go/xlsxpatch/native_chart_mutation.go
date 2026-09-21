package xlsxpatch

import (
	"fmt"
	"strconv"
	"strings"
)

const (
	chartInsert     = "chart.insert"
	chartUpdate     = "chart.update"
	chartDelete     = "chart.delete"
	maxNativeCharts = 32
)

type NativeChartAnchorV1 struct {
	FromRow    int `json:"from_row"`
	FromColumn int `json:"from_column"`
	ToRow      int `json:"to_row"`
	ToColumn   int `json:"to_column"`
}

type ChartMutation struct {
	OperationID               string              `json:"operation_id"`
	SheetID                   string              `json:"sheet_id"`
	Kind                      string              `json:"kind"`
	ChartType                 string              `json:"chart_type,omitempty"`
	Title                     string              `json:"title"`
	Range                     StyleRange          `json:"range"`
	Anchor                    NativeChartAnchorV1 `json:"anchor"`
	Identity                  *ChartIdentity      `json:"identity,omitempty"`
	ExpectedFingerprintSHA256 string              `json:"expected_fingerprint_sha256,omitempty"`
}

func chartOperationIDs(items []ChartMutation) []string {
	ids := make([]string, len(items))
	for i := range items {
		ids[i] = items[i].OperationID
	}
	return ids
}

func validateChartMutations(items []ChartMutation) error {
	seen := map[string]bool{}
	for i, m := range items {
		if m.OperationID == "" || len(m.OperationID) > maxOperationIDLen || !restrictedIDPattern.MatchString(m.OperationID) || seen[m.OperationID] || m.SheetID == "" {
			return fmt.Errorf("xlsxpatch: native mutation: charts[%d]", i)
		}
		seen[m.OperationID] = true
		okType := m.ChartType == "column" || m.ChartType == "bar" || m.ChartType == "line" || m.ChartType == "pie"
		okRange := m.Range.EndRow > m.Range.Row && m.Range.EndColumn > m.Range.Column && m.Range.EndRow-m.Range.Row <= 1000 && m.Range.EndColumn-m.Range.Column <= 8
		okAnchor := m.Anchor.ToRow > m.Anchor.FromRow && m.Anchor.ToColumn > m.Anchor.FromColumn
		switch m.Kind {
		case chartInsert:
			if m.Identity != nil || !okType || !okRange || !okAnchor {
				return fmt.Errorf("xlsxpatch: native mutation: charts[%d] insert", i)
			}
		case chartUpdate:
			if m.Identity == nil || !nativeWorkbookSHA.MatchString(m.ExpectedFingerprintSHA256) || !okType || !okRange || !okAnchor {
				return fmt.Errorf("xlsxpatch: native mutation: charts[%d] update", i)
			}
		case chartDelete:
			if m.Identity == nil || !nativeWorkbookSHA.MatchString(m.ExpectedFingerprintSHA256) {
				return fmt.Errorf("xlsxpatch: native mutation: charts[%d] delete", i)
			}
		default:
			return fmt.Errorf("xlsxpatch: native mutation: charts[%d] kind", i)
		}
	}
	return nil
}

func preflightNativeChartMutationTargets(workbook *NativeWorkbookV1, mutations []ChartMutation) error {
	sheets := indexNativeMutationSheets(workbook)
	for _, m := range mutations {
		s, ok := sheets[m.SheetID]
		if !ok || !s.sheet.Editable {
			return fmt.Errorf("xlsxpatch: native mutation: operation %q sheet", m.OperationID)
		}
	}
	return nil
}

func ApplyNativeChartMutations(orig []byte, workbook *NativeWorkbookV1, mutations []ChartMutation) ([]byte, error) {
	out := orig
	for _, m := range mutations {
		var err error
		var next []byte
		switch m.Kind {
		case chartDelete:
			if err = nativeChartFP(out, *m.Identity, m.ExpectedFingerprintSHA256); err == nil {
				next, err = RemoveChart(out, *m.Identity)
			}
		case chartUpdate:
			var spec ChartWriteSpec
			if err = nativeChartFP(out, *m.Identity, m.ExpectedFingerprintSHA256); err == nil {
				spec, err = nativeChartSpec(workbook, m)
			}
			if err == nil {
				next, err = UpdateChart(out, *m.Identity, spec)
			}
		default:
			var spec ChartWriteSpec
			spec, err = nativeChartSpec(workbook, m)
			if err == nil {
				next, err = AddChart(out, spec)
			}
		}
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: native mutation: %s %q: %w", m.Kind, m.OperationID, err)
		}
		out = next
	}
	return out, nil
}

func nativeChartFP(data []byte, id ChartIdentity, want string) error {
	pkg, err := openNativeWorkbookPackage(data)
	if err != nil {
		return err
	}
	part := pkg.files[id.Part]
	if got := nativeWorkbookDigest(part); got != want {
		return fmt.Errorf("stale chart fingerprint: expected %q, exact part is %q", want, got)
	}
	return nil
}

func nativeChartSpec(workbook *NativeWorkbookV1, m ChartMutation) (ChartWriteSpec, error) {
	var sheet *NativeWorkbookSheetV1
	for i := range workbook.Sheets {
		if workbook.Sheets[i].ID == m.SheetID {
			sheet = &workbook.Sheets[i]
			break
		}
	}
	if sheet == nil {
		return ChartWriteSpec{}, fmt.Errorf("missing sheet")
	}
	series, err := nativeChartSeries(sheet, m.Range, m.ChartType)
	if err != nil {
		return ChartWriteSpec{}, err
	}
	return ChartWriteSpec{SheetName: sheet.Name, Type: m.ChartType, Title: m.Title, Series: series,
		Anchor: ChartAnchor{FromCol: m.Anchor.FromColumn, FromRow: m.Anchor.FromRow, ToCol: m.Anchor.ToColumn, ToRow: m.Anchor.ToRow}}, nil
}

func nativeChartSeries(sheet *NativeWorkbookSheetV1, src StyleRange, typ string) ([]WriteSeries, error) {
	cells := map[cellKey]*NativeWorkbookCellV1{}
	for i := range sheet.Cells {
		c := &sheet.Cells[i]
		cells[cellKey{row: c.Row, column: c.Column}] = c
	}
	abs := func(r1, c1, r2, c2 int) string {
		a := sheet.Name + "!$" + columnName(c1) + "$" + strconv.Itoa(r1+1)
		if r1 == r2 && c1 == c2 {
			return a
		}
		return a + ":$" + columnName(c2) + "$" + strconv.Itoa(r2+1)
	}
	var series []WriteSeries
	for col := src.Column + 1; col <= src.EndColumn; col++ {
		name := fmt.Sprintf("Series %d", len(series)+1)
		if h := cells[cellKey{row: src.Row, column: col}]; h != nil && h.Formula == nil && h.Value != nil {
			if h.Value.Text != nil && *h.Value.Text != "" {
				name = *h.Value.Text
			} else if h.Value.Lexical != nil && *h.Value.Lexical != "" {
				name = *h.Value.Lexical
			}
		}
		for row := src.Row + 1; row <= src.EndRow; row++ {
			v := cells[cellKey{row: row, column: col}]
			c := cells[cellKey{row: row, column: src.Column}]
			if v == nil || v.Formula != nil || v.Value == nil || v.Value.Kind != "number" || v.Value.Lexical == nil ||
				c == nil || c.Formula != nil || c.Value == nil {
				return nil, fmt.Errorf("chart source must be literal numbers with category labels")
			}
			if typ == "pie" {
				n, convErr := strconv.ParseFloat(*v.Value.Lexical, 64)
				if convErr != nil || n < 0 {
					return nil, fmt.Errorf("pie charts require non-negative literal numbers")
				}
			}
		}
		series = append(series, WriteSeries{Name: name, NameRef: abs(src.Row, col, src.Row, col), CategoriesRef: abs(src.Row+1, src.Column, src.EndRow, src.Column), ValuesRef: abs(src.Row+1, col, src.EndRow, col)})
	}
	if len(series) == 0 {
		return nil, fmt.Errorf("chart needs at least one series")
	}
	return series, nil
}

func nativeChartOwned(item NativeWorkbookUnsupportedV1, workbook *NativeWorkbookV1, transaction NativeWorkbookMutationTransactionV1) bool {
	if item.Code == "CHART_CONTENT" || item.Code == "DRAWING_OR_MEDIA_CONTENT" || item.Code == "DRAWING_REFERENCE" || item.Capability == "charts" || item.Capability == "drawings" {
		return true
	}
	if item.PartName == nil {
		return false
	}
	n := strings.ToLower(*item.PartName)
	if strings.Contains(n, "/charts/") || strings.Contains(n, "/drawings/") {
		return true
	}
	for _, m := range transaction.Charts {
		if p := workbookSheetPart(workbook, m.SheetID); p != "" && *item.PartName == relsPartFor(p) {
			return true
		}
	}
	return false
}

func nativeExpectedUnsupportedAfterCharts(before, after *NativeWorkbookV1, transaction NativeWorkbookMutationTransactionV1, expected []NativeWorkbookUnsupportedV1) []NativeWorkbookUnsupportedV1 {
	if len(transaction.Charts) == 0 {
		return expected
	}
	out := make([]NativeWorkbookUnsupportedV1, 0, len(after.Unsupported))
	for _, item := range expected {
		if !nativeChartOwned(item, before, transaction) {
			out = append(out, item)
		}
	}
	for _, item := range after.Unsupported {
		if nativeChartOwned(item, before, transaction) {
			out = append(out, item)
		}
	}
	return out
}

func nativeChartMutationAllowedParts(original, produced []byte, workbook *NativeWorkbookV1, transaction NativeWorkbookMutationTransactionV1) (map[string]bool, error) {
	allowed := map[string]bool{"[Content_Types].xml": true}
	for _, m := range transaction.Charts {
		if p := workbookSheetPart(workbook, m.SheetID); p != "" {
			allowed[p], allowed[relsPartFor(p)] = true, true
		}
	}
	before, err := openNativeWorkbookPackage(original)
	if err != nil {
		return nil, err
	}
	after, err := openNativeWorkbookPackage(produced)
	if err != nil {
		return nil, err
	}
	for _, files := range []map[string][]byte{before.files, after.files} {
		for name := range files {
			n := strings.ToLower(name)
			if strings.Contains(n, "/charts/") || strings.Contains(n, "/drawings/") {
				allowed[name] = true
			}
		}
	}
	return allowed, nil
}
