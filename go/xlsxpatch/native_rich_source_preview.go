package xlsxpatch

import (
	"encoding/xml"
	"fmt"
	"math"
	"strings"
)

// NativeRichSourcePreviewV1 is a read-only content rectangle from original
// source records. It is not an editable workbook or a print layout.
type NativeRichSourcePreviewV1 struct {
	Protocol        string                        `json:"protocol"`
	Version         int                           `json:"version"`
	ReadOnly        bool                          `json:"read_only"`
	Fidelity        string                        `json:"fidelity"`
	PackageSHA256   string                        `json:"package_sha256"`
	WorkbookPart    string                        `json:"workbook_part"`
	StylesPart      string                        `json:"styles_part"`
	StylesSHA256    string                        `json:"styles_sha256"`
	WorksheetSHA256 string                        `json:"worksheet_sha256"`
	StrictError     string                        `json:"strict_error"`
	ParentCount     NativeRichSourceCountV1       `json:"parent_count"`
	Styles          []NativeSourceStyleV1         `json:"styles"`
	UnusedStyleIDs  []int                         `json:"unused_style_ids"`
	Sheet           NativeSourceStyleSheetV1      `json:"sheet"`
	RichCells       []NativeRichSourceCellV1      `json:"rich_cells"`
	OmittedRows     []NativeRichSourceDimensionV1 `json:"omitted_rows"`
	OmittedColumns  []NativeRichSourceDimensionV1 `json:"omitted_columns"`
	Warnings        []string                      `json:"warnings"`
}
type NativeRichSourceCountV1 struct {
	Table       string `json:"table"`
	Declaration string `json:"declaration"`
	Observed    int    `json:"observed"`
}
type NativeRichSourceDimensionV1 struct {
	Index int     `json:"index"`
	Size  float64 `json:"size"`
}
type NativeRichSourceCellV1 struct {
	Ref     string                  `json:"ref"`
	StyleID uint32                  `json:"style_id"`
	Text    string                  `json:"text"`
	Runs    []NativeRichSourceRunV1 `json:"runs"`
}
type NativeRichSourceRunV1 struct {
	Start int                 `json:"start"`
	End   int                 `json:"end"`
	Style NativeRichTextRunV1 `json:"style"`
}

func richSourceError(reason string) error {
	return fmt.Errorf("xlsxpatch: rich-source preview: %s", reason)
}

// PreviewNativeRichSourceV1 admits only an absent parent-XF count, inferred
// from bounded original records. Every displayed cell must explicitly own a
// qualified nondefault style in a complete A1-based rectangle. No bytes change.
func PreviewNativeRichSourceV1(data []byte) (*NativeRichSourcePreviewV1, error) {
	pkg, err := openNativeWorkbookPackage(data)
	if err != nil {
		return nil, err
	}
	location, err := locateWorkbookPartBytes(pkg.index, func(n string) ([]byte, bool) { b, ok := pkg.files[n]; return b, ok })
	if err != nil {
		return nil, err
	}
	if location.strict {
		return nil, richSourceError("requires Transitional source")
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
		return nil, richSourceError("requires one visible worksheet")
	}
	sp, err := ex.relatedCorePart(relTypeStylesTransitional, relTypeStylesStrict, stylesPartContentType, "styles", true)
	if err != nil {
		return nil, err
	}
	if err = ex.claimCoreXML(sp); err != nil {
		return nil, err
	}
	_, strictErr := newStyleRegistry(pkg.files[sp])
	if strictErr == nil {
		return nil, richSourceError("no missing-count recovery required")
	}
	r, count, err := readRichSourceRegistry(pkg.files[sp])
	if err != nil {
		return nil, err
	}
	for _, xf := range r.styleXfs {
		if err = r.validateXFReferences(xf, false); err != nil {
			return nil, err
		}
	}
	for _, xf := range r.cellXfs {
		if err = r.validateXFReferences(xf, true); err != nil {
			return nil, err
		}
	}
	shared, sharedPart, err := ex.extractSharedStrings()
	if err != nil {
		return nil, err
	}
	// Shared strings are outside this inline-only profile, including unused tables.
	if sharedPart != "" {
		return nil, richSourceError("shared-string table outside inline profile")
	}
	for _, u := range ex.unsupported {
		if strings.HasPrefix(u.Code, "SHARED_STRING_") {
			return nil, richSourceError("unqualified shared strings")
		}
	}
	route := routes[0]
	if err = ex.claimCoreXML(route.part); err != nil {
		return nil, err
	}
	sheet, err := ex.extractWorksheet(route, 0, pkg.files[route.part], shared, len(r.cellXfs), true)
	if err != nil {
		return nil, err
	}
	raw, err := parsePreviewXML(pkg.files[route.part])
	if err != nil {
		return nil, err
	}
	rich, err := qualifyRichSourceSheet(raw)
	if err != nil {
		return nil, err
	}
	if len(sheet.Cells) == 0 || len(sheet.Cells) > 4096 || len(sheet.MergedRanges) != 0 {
		return nil, richSourceError("requires unmerged bounded content")
	}
	rows, cols := 0, 0
	for _, c := range sheet.Cells {
		if c.Row >= 128 || c.Column >= 32 || c.StyleID == 0 {
			return nil, richSourceError("content exceeds bounds or uses default style")
		}
		rows = max(rows, c.Row+1)
		cols = max(cols, c.Column+1)
	}
	if len(sheet.Cells) != rows*cols {
		return nil, richSourceError("sparse content rectangle")
	}
	out := &NativeRichSourcePreviewV1{Protocol: "injoffice.xlsx.rich-source-preview", Version: 1, ReadOnly: true, Fidelity: "approximate", PackageSHA256: nativeWorkbookDigest(data), WorkbookPart: location.part, StylesPart: sp, StylesSHA256: nativeWorkbookDigest(pkg.files[sp]), WorksheetSHA256: nativeWorkbookDigest(pkg.files[route.part]), StrictError: strictErr.Error(), ParentCount: NativeRichSourceCountV1{"cellStyleXfs", "absent", count}, Styles: []NativeSourceStyleV1{}, UnusedStyleIDs: []int{}, RichCells: []NativeRichSourceCellV1{}, OmittedRows: []NativeRichSourceDimensionV1{}, OmittedColumns: []NativeRichSourceDimensionV1{}, Warnings: []string{"Read-only content rectangle. The missing parent-style count is recorded from original direct records; no source bytes are repaired.", "Font matching, rich wrapping and grid metrics are approximate. Print settings, viewport settings and geometry outside the content rectangle are not applied.", "General numbers retain their exact stored text. No formulas are evaluated. Unused styles are not projected; all displayed cells explicitly select qualified styles.", "Absent run properties use the displayed cell font. Disabled strike and underline are no-op declarations; font-family and baseline hints remain approximate."}}
	out.Sheet = NativeSourceStyleSheetV1{ID: route.id, Name: route.name, Part: route.part, RowHeights: make([]float64, rows), ColumnWidths: make([]float64, cols), Cells: []NativeSourceStyleCellV1{}, Merges: []NativeSourceStyleMergeV1{}}
	// Require source dimensions for every displayed band; never invent the Normal font or widths.
	for _, v := range sheet.Rows {
		if v.Row >= 128 || v.Hidden || v.StyleID != nil || v.HeightPoints == nil || !richSourceSize(*v.HeightPoints, 409) {
			return nil, richSourceError("unqualified row geometry")
		}
		if v.Row < rows {
			out.Sheet.RowHeights[v.Row] = *v.HeightPoints
		} else {
			out.OmittedRows = append(out.OmittedRows, NativeRichSourceDimensionV1{v.Row, *v.HeightPoints})
		}
	}
	for _, v := range sheet.Columns {
		if v.EndColumn >= 32 || v.Hidden || v.BestFit || v.StyleID != nil || v.Width == nil || !richSourceSize(*v.Width, 255) {
			return nil, richSourceError("unqualified column geometry")
		}
		for i := v.Column; i <= v.EndColumn; i++ {
			if i < cols {
				out.Sheet.ColumnWidths[i] = *v.Width
			} else {
				out.OmittedColumns = append(out.OmittedColumns, NativeRichSourceDimensionV1{i, *v.Width})
			}
		}
	}
	for _, v := range out.Sheet.RowHeights {
		if v == 0 {
			return nil, richSourceError("missing content row height")
		}
	}
	for _, v := range out.Sheet.ColumnWidths {
		if v == 0 {
			return nil, richSourceError("missing content column width")
		}
	}
	used := map[int]bool{}
	seen := map[string]bool{}
	total, runCount := 0, 0
	for _, c := range sheet.Cells {
		if seen[c.Ref] || c.Ref != cellReference(c.Row, c.Column) || c.Formula != nil || c.Value == nil {
			return nil, richSourceError("ambiguous cell ownership or formula")
		}
		seen[c.Ref] = true
		v := c.Value
		entry := NativeSourceStyleCellV1{Ref: c.Ref, Row: c.Row, Column: c.Column, StyleID: c.StyleID, Kind: v.Kind}
		if v.Text != nil {
			entry.Text = *v.Text
		}
		if v.Lexical != nil {
			entry.Lexical = *v.Lexical
		}
		if v.Kind != "number" && v.Kind != "string" || v.Kind == "string" && v.Storage != "inline" || len(entry.Text) > 4096 || len(entry.Lexical) > 256 {
			return nil, richSourceError("unsupported cell value")
		}
		total += nativeRichUnits(entry.Text) + nativeRichUnits(entry.Lexical)
		if total > 65536 {
			return nil, richSourceError("aggregate text bound")
		}
		runs, ok := rich[c.Ref]
		if ok != v.Rich {
			return nil, richSourceError("rich source/extraction mismatch")
		}
		if ok {
			if len(runs) != len(v.runs) {
				return nil, richSourceError("rich run count mismatch")
			}
			rc := NativeRichSourceCellV1{Ref: c.Ref, StyleID: c.StyleID, Text: entry.Text, Runs: []NativeRichSourceRunV1{}}
			joined := ""
			offset := 0
			for i, run := range runs {
				if run.Text != v.runs[i].Text {
					return nil, richSourceError("rich run text mismatch")
				}
				end := offset + nativeRichUnits(run.Text)
				rc.Runs = append(rc.Runs, NativeRichSourceRunV1{offset, end, run})
				offset = end
				joined += run.Text
			}
			if joined != entry.Text {
				return nil, richSourceError("rich text join mismatch")
			}
			runCount += len(runs)
			if runCount > 1024 || len(out.RichCells) >= 256 {
				return nil, richSourceError("rich aggregate bound")
			}
			out.RichCells = append(out.RichCells, rc)
		}
		used[int(c.StyleID)] = true
		out.Sheet.Cells = append(out.Sheet.Cells, entry)
	}
	if len(out.RichCells) == 0 {
		return nil, richSourceError("no qualified rich cells")
	}
	for i, xf := range r.cellXfs {
		if !used[i] {
			out.UnusedStyleIDs = append(out.UnusedStyleIDs, i)
			continue
		}
		if xf.applyFont == nil || !*xf.applyFont || xf.applyFill == nil || !*xf.applyFill || xf.applyBorder == nil || !*xf.applyBorder || xf.applyNumberFormat == nil || !*xf.applyNumberFormat || xf.applyAlignment == nil || !*xf.applyAlignment {
			return nil, richSourceError("used style lacks explicit component ownership")
		}
		s, err := projectSourceStyleWithQualifiers(r, i, richSourceXFSupported, richSourceDisabledFont)
		if err != nil {
			return nil, err
		}
		if len(s.Warnings) != 0 || s.NumberFormat != "General" {
			return nil, richSourceError("requires direct font color and General format")
		}
		out.Styles = append(out.Styles, s)
	}
	return out, nil
}
func richSourceSize(v, max float64) bool {
	return !math.IsNaN(v) && !math.IsInf(v, 0) && v > 0 && v <= max
}
func richSourceXFSupported(data []byte, e styleTableEntry, ns string) bool {
	return sourceStyleXFWithAlignment(data, e, ns, func(a xml.StartElement) bool {
		if !styleAttributesOnly(a, "horizontal", "vertical", "wrapText", "textRotation", "indent", "shrinkToFit", "readingOrder") {
			return false
		}
		v, ok, err := unqualifiedXMLAttribute(a, "readingOrder")
		return err == nil && (!ok || v == "1")
	})
}
func richSourceDisabledFont(a xml.StartElement) bool {
	if !styleAttributesOnly(a, "val") {
		return false
	}
	v, ok, err := unqualifiedXMLAttribute(a, "val")
	return err == nil && ok && (a.Name.Local == "strike" && (v == "0" || v == "false") || a.Name.Local == "u" && v == "none")
}

func readRichSourceRegistry(data []byte) (*styleRegistry, int, error) {
	root, err := parsePreviewXML(data)
	if err != nil {
		return nil, 0, err
	}
	if !sourceStyleNode(root, "styleSheet") {
		return nil, 0, richSourceError("unqualified styles root")
	}
	seen := map[string]bool{}
	count := 0
	for _, n := range root.children {
		if seen[n.name.Local] || n.name.Space != spreadsheetMLTransitional {
			return nil, 0, richSourceError("duplicate or foreign style table")
		}
		seen[n.name.Local] = true
		switch n.name.Local {
		case "cellStyleXfs":
			if !sourceStyleNode(n, "cellStyleXfs") || len(n.children) == 0 || len(n.children) > 128 {
				return nil, 0, richSourceError("requires absent parent count and 1–128 records")
			}
			count = len(n.children)
		case "fonts", "fills", "borders", "cellXfs", "numFmts":
			if !sourceStyleNode(n, n.name.Local, "count") || len(n.children) > 128 {
				return nil, 0, richSourceError("unqualified bounded style table")
			}
		case "cellStyles":
			if !sourceStyleNode(n, "cellStyles", "count") || n.attr("count") != "1" || len(n.children) != 1 {
				return nil, 0, richSourceError("unqualified named style table")
			}
			c := n.children[0]
			if !sourceStyleNode(c, "cellStyle", "xfId", "name") || c.attr("xfId") != "0" || c.attr("name") != "Normal" || len(c.children) != 0 {
				return nil, 0, richSourceError("unqualified Normal style")
			}
		case "dxfs":
			if !sourceStyleNode(n, "dxfs", "count") || n.attr("count") != "0" || len(n.children) != 0 {
				return nil, 0, richSourceError("differential styles outside profile")
			}
		case "tableStyles":
			if !sourceStyleNode(n, "tableStyles", "count", "defaultTableStyle", "defaultPivotStyle") || n.attr("count") != "0" || len(n.children) != 0 {
				return nil, 0, richSourceError("custom table styles outside profile")
			}
		case "colors":
			if !sourceStyleNode(n, "colors") || len(n.children) != 1 {
				return nil, 0, richSourceError("unqualified color table")
			}
			colors := n.children[0]
			if !sourceStyleNode(colors, "indexedColors") || len(colors.children) > 64 {
				return nil, 0, richSourceError("unqualified indexed colors")
			}
			for _, c := range colors.children {
				v := c.attr("rgb")
				if !sourceStyleNode(c, "rgbColor", "rgb") || len(c.children) != 0 || len(v) != 8 || strings.Trim(v, "0123456789abcdefABCDEF") != "" {
					return nil, 0, richSourceError("invalid indexed color")
				}
			}
		default:
			return nil, 0, richSourceError("unknown style metadata")
		}
	}
	if count == 0 {
		return nil, 0, richSourceError("missing parent table")
	}
	index, err := parseStyleTableWithCount(data, func(a xml.StartElement) (int, error) {
		if a.Name.Space == spreadsheetMLTransitional && a.Name.Local == "cellStyleXfs" {
			return count, nil
		}
		return styleTableCount(a)
	})
	if err != nil {
		return nil, 0, err
	}
	r, err := decodeStyleRegistryRecords(data, index)
	if err != nil {
		return nil, 0, err
	}
	if r.index.numFmts != nil {
		for _, e := range r.index.numFmts.entries {
			if !styleAttributesOnly(e.start, "numFmtId", "formatCode") {
				return nil, 0, richSourceError("unknown number-format metadata")
			}
		}
	}
	return r, count, nil
}

func qualifyRichSourceSheet(root *previewXML) (map[string][]NativeRichTextRunV1, error) {
	admitted := map[*previewXML]bool{}
	rich := map[string][]NativeRichTextRunV1{}
	seen := map[string]bool{}
	for _, n := range root.children {
		if seen[n.name.Local] {
			return nil, richSourceError("duplicate worksheet section")
		}
		seen[n.name.Local] = true
		switch n.name.Local {
		case "sheetPr":
			if !sourceStyleNode(n, "sheetPr") || len(n.children) != 1 {
				return nil, richSourceError("unqualified sheet properties")
			}
			c := n.children[0]
			if !sourceStyleNode(c, "outlinePr") || len(c.children) != 0 {
				return nil, richSourceError("active outline properties")
			}
			admitted[c] = true
		case "headerFooter":
			if !sourceStyleNode(n, "headerFooter", "alignWithMargins") || n.attr("alignWithMargins") != "0" || len(n.children) != 0 {
				return nil, richSourceError("active header/footer")
			}
			admitted[n] = true
		case "sheetData":
			for _, row := range n.children {
				for _, c := range row.children {
					for _, v := range c.children {
						if v.name.Local == "is" && len(v.children) > 0 && v.children[0].name.Local == "r" {
							if c.attr("t") != "inlineStr" || len(c.children) != 1 || rich[c.attr("r")] != nil {
								return nil, richSourceError("ambiguous inline rich ownership")
							}
							runs, reason := nativeRichRuns(v, spreadsheetMLTransitional)
							if reason != "" {
								return nil, richSourceError(reason)
							}
							for _, r := range runs {
								if r.FontScheme != "" || r.Underline != "none" && r.Underline != "" {
									return nil, richSourceError("active or theme-dependent run effects")
								}
							}
							rich[c.attr("r")] = runs
							admitted[v] = true
						}
					}
				}
			}
		case "sheetViews", "cols", "pageMargins", "pageSetup":
			// Qualified recursively below; viewport and print state are disclosed omissions.
		default:
			return nil, richSourceError("worksheet family outside rich profile")
		}
	}
	if err := qualifySourceStyleSheetWithNodes(root, spreadsheetMLTransitional, admitted); err != nil {
		return nil, err
	}
	return rich, nil
}
