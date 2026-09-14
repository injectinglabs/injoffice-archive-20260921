package xlsxpatch

import (
	"encoding/xml"
	"fmt"
	"math"
	"strings"
)

// NativeSourceStylePreviewV1 is a read-only compatibility envelope. It has no
// native workbook revision, editable model or mutation capability.
type NativeSourceStylePreviewV1 struct {
	Protocol      string                        `json:"protocol"`
	Version       int                           `json:"version"`
	ReadOnly      bool                          `json:"read_only"`
	Fidelity      string                        `json:"fidelity"`
	PackageSHA256 string                        `json:"package_sha256"`
	WorkbookPart  string                        `json:"workbook_part"`
	StylesPart    string                        `json:"styles_part"`
	Date1904      bool                          `json:"date1904"`
	StrictError   string                        `json:"strict_error"`
	Warnings      []string                      `json:"warnings"`
	Conflicts     []NativeSourceStyleConflictV1 `json:"conflicts"`
	Styles        []NativeSourceStyleV1         `json:"styles"`
	Sheets        []NativeSourceStyleSheetV1    `json:"sheets"`
}
type NativeSourceStyleConflictV1 struct {
	StyleID           int    `json:"style_id"`
	ParentID          int    `json:"parent_id"`
	Component         string `json:"component"`
	ApplyFlag         string `json:"apply_flag"`
	ParentComponentID int    `json:"parent_component_id"`
	DirectComponentID int    `json:"direct_component_id"`
}
type NativeSourceStyleV1 struct {
	ID             int               `json:"id"`
	ParentID       int               `json:"parent_id"`
	ParentSHA256   string            `json:"parent_sha256"`
	FontID         int               `json:"font_id"`
	FillID         int               `json:"fill_id"`
	BorderID       int               `json:"border_id"`
	RawSHA256      string            `json:"raw_sha256"`
	FontSHA256     string            `json:"font_sha256"`
	FillSHA256     string            `json:"fill_sha256"`
	BorderSHA256   string            `json:"border_sha256"`
	NumberFormatID int               `json:"number_format_id"`
	NumberFormat   string            `json:"number_format"`
	FontName       string            `json:"font_name"`
	FontSizePoints float64           `json:"font_size_points"`
	FontColor      string            `json:"font_color"`
	Bold           bool              `json:"bold"`
	Italic         bool              `json:"italic"`
	FillColor      string            `json:"fill_color"`
	Horizontal     string            `json:"horizontal"`
	Vertical       string            `json:"vertical"`
	Wrap           bool              `json:"wrap"`
	Borders        map[string]string `json:"borders"`
	Warnings       []string          `json:"warnings"`
}
type NativeSourceStyleCellV1 struct {
	Ref     string `json:"ref"`
	Row     int    `json:"row"`
	Column  int    `json:"column"`
	StyleID uint32 `json:"style_id"`
	Kind    string `json:"kind"`
	Text    string `json:"text"`
	Lexical string `json:"lexical"`
	Formula string `json:"formula"`
	Cached  bool   `json:"cached"`
}
type NativeSourceStyleMergeV1 struct {
	Row       int `json:"row"`
	Column    int `json:"column"`
	EndRow    int `json:"end_row"`
	EndColumn int `json:"end_column"`
}
type NativeSourceStyleSheetV1 struct {
	ID           string                     `json:"id"`
	Name         string                     `json:"name"`
	Part         string                     `json:"part"`
	RowHeights   []float64                  `json:"row_heights"`
	ColumnWidths []float64                  `json:"column_widths"`
	Cells        []NativeSourceStyleCellV1  `json:"cells"`
	Merges       []NativeSourceStyleMergeV1 `json:"merges"`
}

// PreviewNativeSourceStylesV1 reads original bytes through validated OPC/XML
// routes. Only absent applyFill/applyNumberFormat conflicts are tolerated, and
// each is reported. The strict registry and native resolver remain unchanged.
func PreviewNativeSourceStylesV1(data []byte) (*NativeSourceStylePreviewV1, error) {
	return previewNativeSourceStyles(data, func(root *previewXML, _, _ []byte) error {
		return qualifySourceStyleSheet(root, spreadsheetMLTransitional)
	})
}

// The qualifier admits only original worksheet nodes validated by its profile.
// It never rewrites a source part or changes native extraction authority.
func previewNativeSourceStyles(data []byte, qualify func(*previewXML, []byte, []byte) error) (*NativeSourceStylePreviewV1, error) {
	fail := func(reason string) (*NativeSourceStylePreviewV1, error) {
		return nil, fmt.Errorf("xlsxpatch: source-style preview: %s", reason)
	}
	pkg, err := openNativeWorkbookPackage(data)
	if err != nil {
		return nil, err
	}
	location, err := locateWorkbookPartBytes(pkg.index, func(n string) ([]byte, bool) { b, ok := pkg.files[n]; return b, ok })
	if err != nil {
		return nil, err
	}
	if location.strict {
		return fail("Strict sources are outside this compatibility profile")
	}
	if err = requireNativeContentType(pkg, location.part, nativeWorkbookContentType); err != nil {
		return nil, err
	}
	ex := &nativeWorkbookExtractor{pkg: pkg, workbook: location, namespace: spreadsheetMLTransitional, relNamespace: officeRelNamespaceTransitional, modeled: map[string]bool{}, unsupportedKeys: map[string]bool{}, claimedXML: map[string]bool{}}
	for _, p := range []string{pkg.contentTypesPart, location.part, location.relsPart, "_rels/.rels"} {
		if err = ex.claimCoreXML(p); err != nil {
			return nil, err
		}
	}
	if err = ex.validateAllRelationships(); err != nil {
		return nil, err
	}
	routes, err := ex.extractWorkbookRoutes(pkg.files[location.part])
	if err != nil {
		return nil, err
	}
	if len(routes) != 1 || routes[0].state != "visible" {
		return fail("requires exactly one visible worksheet")
	}
	stylesPart, err := ex.relatedCorePart(relTypeStylesTransitional, relTypeStylesStrict, stylesPartContentType, "styles", true)
	if err != nil {
		return nil, err
	}
	if err = ex.claimCoreXML(stylesPart); err != nil {
		return nil, err
	}
	if _, err = parsePreviewXML(pkg.files[stylesPart]); err != nil {
		return nil, err
	}
	if _, err = newStyleRegistry(pkg.files[stylesPart]); err == nil {
		return fail("source styles do not require compatibility recovery")
	}
	strictError := err.Error()
	registry, err := readStyleRegistryRecords(pkg.files[stylesPart])
	if err != nil {
		return nil, err
	}
	if registry.index.namespace != ex.namespace {
		return fail("opposing styles dialect")
	}
	if len(registry.cellXfs) > 128 || len(registry.styleXfs) > 128 || len(registry.fonts) > 128 || len(registry.fills) > 128 || len(registry.borders) > 128 {
		return fail("style tables exceed 128 records")
	}
	if registry.index.numFmts != nil {
		if len(registry.index.numFmts.entries) > 128 {
			return fail("number-format table exceeds 128 records")
		}
		for _, entry := range registry.index.numFmts.entries {
			if !styleAttributesOnly(entry.start, "numFmtId", "formatCode") {
				return fail("unqualified number-format attributes")
			}
		}
	}
	result := &NativeSourceStylePreviewV1{Protocol: "injoffice.xlsx.source-style-preview", Version: 1, ReadOnly: true, Fidelity: "approximate", PackageSHA256: nativeWorkbookDigest(data), WorkbookPart: location.part, StylesPart: stylesPart, StrictError: strictError, Warnings: []string{"Read-only source grid. Absent fill/number-format apply flags use the directly recorded IDs; conflicts remain explicit. No source bytes are repaired.", "Browser font matching, column widths, wrapping and border metrics are approximate. Print settings and workbook calculation are not applied; saved formula caches may be stale."}, Conflicts: []NativeSourceStyleConflictV1{}, Styles: []NativeSourceStyleV1{}, Sheets: []NativeSourceStyleSheetV1{}}
	for _, xf := range registry.styleXfs {
		if err = registry.validateXFReferenceIDs(xf, false); err != nil {
			return nil, err
		}
	}
	for i, xf := range registry.cellXfs {
		if err = registry.validateXFReferenceIDs(xf, true); err != nil {
			return nil, err
		}
		base := effectiveCellStyleXF(registry.styleXfs[xf.xfID])
		for _, c := range []struct {
			name, flag     string
			direct, parent int
			apply          *bool
			allow          bool
		}{{"font", "applyFont", xf.fontID, base.fontID, xf.applyFont, false}, {"fill", "applyFill", xf.fillID, base.fillID, xf.applyFill, true}, {"border", "applyBorder", xf.borderID, base.borderID, xf.applyBorder, false}, {"number-format", "applyNumberFormat", xf.numFmtID, base.numFmtID, xf.applyNumberFormat, true}} {
			if c.direct != c.parent && (c.apply == nil || !*c.apply) {
				if !c.allow || c.apply != nil {
					return fail("explicit false or unsupported inheritance conflict")
				}
				result.Conflicts = append(result.Conflicts, NativeSourceStyleConflictV1{i, xf.xfID, c.name, c.flag, c.parent, c.direct})
			}
		}
		if (xf.applyAlignment == nil || !*xf.applyAlignment) && xf.alignment.present && !styleAlignmentsEquivalent(xf.alignment, base.alignment) {
			return fail("alignment inheritance conflict")
		}
	}
	if len(result.Conflicts) == 0 {
		return fail("no supported absent-flag conflict")
	}
	shared, _, err := ex.extractSharedStrings()
	if err != nil {
		return nil, err
	}
	for _, item := range ex.unsupported {
		if strings.HasPrefix(item.Code, "SHARED_STRING_") {
			return fail("unqualified shared-string table")
		}
	}
	route := routes[0]
	if err = ex.claimCoreXML(route.part); err != nil {
		return nil, err
	}
	sheet, err := ex.extractWorksheet(route, 0, pkg.files[route.part], shared, len(registry.cellXfs), true)
	if err != nil {
		return nil, err
	}
	// Read-only source geometry gets its own closed qualification, independent of
	// the native sheet's editability/refusal fields. No NativeWorkbook is built.
	raw, err := parsePreviewXML(pkg.files[route.part])
	if err != nil {
		return nil, err
	}
	if err = qualify(raw, pkg.files[stylesPart], pkg.files[route.part]); err != nil {
		return nil, err
	}
	if len(sheet.Cells) > 4096 || len(sheet.MergedRanges) > 128 {
		return fail("stored cells or merges exceed preview bounds")
	}
	rows, cols := 1, 1
	for _, c := range sheet.Cells {
		if c.Row >= 128 || c.Column >= 32 {
			return fail("source grid exceeds 128 rows or 32 columns")
		}
		rows = max(rows, c.Row+1)
		cols = max(cols, c.Column+1)
	}
	for _, m := range sheet.MergedRanges {
		if m.EndRow >= 128 || m.EndColumn >= 32 {
			return fail("merge exceeds grid bound")
		}
		rows = max(rows, m.EndRow+1)
		cols = max(cols, m.EndColumn+1)
	}
	out := NativeSourceStyleSheetV1{ID: route.id, Name: route.name, Part: route.part, RowHeights: make([]float64, rows), ColumnWidths: make([]float64, cols), Cells: []NativeSourceStyleCellV1{}, Merges: []NativeSourceStyleMergeV1{}}
	height, width := 15.0, 8.43
	if sheet.SheetFormat != nil {
		height = sheet.SheetFormat.DefaultRowHeightPoints
		if sheet.SheetFormat.DefaultColumnWidth != nil {
			width = *sheet.SheetFormat.DefaultColumnWidth
		}
	}
	for i := range out.RowHeights {
		out.RowHeights[i] = height
	}
	for i := range out.ColumnWidths {
		out.ColumnWidths[i] = width
	}
	for _, r := range sheet.Rows {
		if r.Row >= rows {
			continue
		}
		if r.Hidden || r.StyleID != nil {
			return fail("hidden or row-style geometry unsupported")
		}
		if r.HeightPoints != nil {
			out.RowHeights[r.Row] = *r.HeightPoints
		}
	}
	for _, c := range sheet.Columns {
		if c.Hidden || c.BestFit || c.StyleID != nil && *c.StyleID != 0 {
			return fail("hidden/best-fit or column-style geometry unsupported")
		}
		if c.Width != nil {
			for i := c.Column; i <= c.EndColumn && i < cols; i++ {
				out.ColumnWidths[i] = *c.Width
			}
		}
	}
	for _, v := range out.RowHeights {
		if math.IsNaN(v) || math.IsInf(v, 0) || v <= 0 || v > 409 {
			return fail("unsupported row height")
		}
	}
	for _, v := range out.ColumnWidths {
		if math.IsNaN(v) || math.IsInf(v, 0) || v <= 0 || v > 255 {
			return fail("unsupported column width")
		}
	}
	for _, m := range sheet.MergedRanges {
		out.Merges = append(out.Merges, NativeSourceStyleMergeV1{m.Row, m.Column, m.EndRow, m.EndColumn})
	}
	used := map[int]bool{0: true}
	totalText := 0
	for _, c := range sheet.Cells {
		entry := NativeSourceStyleCellV1{Ref: c.Ref, Row: c.Row, Column: c.Column, StyleID: c.StyleID, Kind: "blank"}
		v := c.Value
		if c.Formula != nil {
			if c.Formula.Type != "normal" || c.Formula.Ref != nil || c.Formula.SharedIndex != nil {
				return fail("formula groups are unsupported")
			}
			entry.Formula = c.Formula.Text
			v = c.Formula.Cached
			entry.Cached = v != nil
			if v == nil {
				return fail("formula has no saved result")
			}
		}
		if v != nil {
			if v.Rich {
				return fail("rich strings are outside this source-grid slice")
			}
			entry.Kind = v.Kind
			if v.Text != nil {
				entry.Text = *v.Text
			}
			if v.Lexical != nil {
				entry.Lexical = *v.Lexical
			}
		}
		switch entry.Kind {
		case "blank", "number", "string", "boolean", "error", "date":
		default:
			return fail("opaque cell values are unsupported")
		}
		totalText += nativeRichUnits(entry.Text) + nativeRichUnits(entry.Lexical) + nativeRichUnits(entry.Formula)
		if totalText > 65536 {
			return fail("aggregate cell text exceeds preview bound")
		}
		if len(entry.Text) > 4096 || len(entry.Lexical) > 256 || len(entry.Formula) > 2048 {
			return fail("cell text/lexical/formula exceeds preview bound")
		}
		for _, m := range out.Merges {
			if c.Row >= m.Row && c.Row <= m.EndRow && c.Column >= m.Column && c.Column <= m.EndColumn && (c.Row != m.Row || c.Column != m.Column) && (entry.Kind != "blank" || entry.Formula != "") {
				return fail("covered merge cell has source content")
			}
		}
		used[int(c.StyleID)] = true
		out.Cells = append(out.Cells, entry)
	}
	for i := range registry.cellXfs {
		if used[i] {
			style, err := projectSourceStyle(registry, i)
			if err != nil {
				return nil, err
			}
			result.Styles = append(result.Styles, style)
		}
	}
	result.Date1904, err = parseWorkbookDate1904(pkg.files[location.part], spreadsheetMLTransitional)
	if err != nil {
		return nil, err
	}
	result.Sheets = append(result.Sheets, out)
	return result, nil
}

func sourceStyleNode(n *previewXML, name string, attrs ...string) bool {
	keys := []xml.Name{}
	for _, a := range attrs {
		keys = append(keys, xml.Name{Local: a})
	}
	return nativeRichNode(n, spreadsheetMLTransitional, name, keys...) && strings.TrimSpace(n.text) == ""
}
