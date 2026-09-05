package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"fmt"
	"sort"
	"strings"
)

const (
	nativeGetCorpusSSNS  = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
	nativeGetCorpusRelNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	nativeGetCorpusPkgNS = "http://schemas.openxmlformats.org/package/2006/relationships"
	nativeGetCorpusCTNS  = "http://schemas.openxmlformats.org/package/2006/content-types"

	nativeGetCorpusGETOK              = 200
	nativeGetCorpusGETUnprocessable   = 422
	nativeGetCorpusGETCompileFailed   = 503
	nativeGetCorpusRefusalUnsupported = "xlsx.native.unsupported"
	nativeGetCorpusRefusalFont        = "xlsx.native.font-unavailable"
)

// NativeGetCorpusCase is one committed native GET fixture: the same bytes
// used by the playground demo and by extract/compile/GET tests.
type NativeGetCorpusCase struct {
	ID                   string   `json:"id"`
	File                 string   `json:"file"`
	Title                string   `json:"title"`
	Extract              string   `json:"extract"` // "ok" or "error"
	ExtractError         string   `json:"extract_error,omitempty"`
	RequiredUnsupported  []string `json:"required_unsupported"`
	ForbiddenUnsupported []string `json:"forbidden_unsupported"`
	DrawingCoverage      string   `json:"drawing_coverage"` // "none", "charts", or "refuse"
	GETStatus            int      `json:"get_status"`
	GETRefusal           string   `json:"get_refusal,omitempty"`
	Charts               int      `json:"charts"`
	Demo                 string   `json:"demo"` // "pass" or "refuse"
}

// NativeGetCorpusCases is the GET-blocker chain that 422/503'd Injecting:
// font, apply flags, Excel-default gridlines/xf, freeze, quotePrefix, charts, shapes.
func NativeGetCorpusCases() []NativeGetCorpusCase {
	forbiddenExact := []string{"SHEET_VIEW_GEOMETRY", "STYLE_RECORD_ATTRIBUTES", "DRAWING_REFERENCE"}
	forbiddenNoViewStyle := []string{"SHEET_VIEW_GEOMETRY", "STYLE_RECORD_ATTRIBUTES"}
	return []NativeGetCorpusCase{
		{
			ID: "pass-agent-dejavu", File: "pass-agent-dejavu.xlsx",
			Title: "Agent DejaVu table (no Excel view/xf noise)",
			Extract: "ok", ForbiddenUnsupported: forbiddenExact,
			DrawingCoverage: "none", GETStatus: nativeGetCorpusGETOK, Demo: "pass",
		},
		{
			ID: "pass-empty-inline-str", File: "pass-empty-inline-str.xlsx",
			Title: "Blank t=inlineStr cell with no <is> (Excel empty cell)",
			Extract: "ok", ForbiddenUnsupported: forbiddenExact,
			DrawingCoverage: "none", GETStatus: nativeGetCorpusGETOK, Demo: "pass",
		},
		{
			ID: "pass-excel-defaults", File: "pass-excel-defaults.xlsx",
			Title: "Excel-default gridlines-on + xf pivotButton/quotePrefix=0",
			Extract: "ok", ForbiddenUnsupported: forbiddenExact,
			DrawingCoverage: "none", GETStatus: nativeGetCorpusGETOK, Demo: "pass",
		},
		{
			ID: "pass-excel-defaults-chart", File: "pass-excel-defaults-chart.xlsx",
			Title: "Excel defaults plus a DrawingML column chart (the Injecting GET chain)",
			Extract: "ok", RequiredUnsupported: []string{"DRAWING_REFERENCE"},
			ForbiddenUnsupported: forbiddenNoViewStyle, DrawingCoverage: "charts",
			GETStatus: nativeGetCorpusGETOK, Charts: 1, Demo: "pass",
		},
		{
			ID: "refuse-calibri", File: "refuse-calibri.xlsx",
			Title: "Calibri Normal (host font-unavailable)",
			Extract: "ok", ForbiddenUnsupported: forbiddenExact,
			DrawingCoverage: "none", GETStatus: nativeGetCorpusGETUnprocessable,
			GETRefusal: nativeGetCorpusRefusalFont, Demo: "refuse",
		},
		{
			ID: "refuse-apply-flags", File: "refuse-apply-flags.xlsx",
			Title: "cellXf ids differ without apply flags",
			Extract: "error", ExtractError: "apply flag is false or absent",
			DrawingCoverage: "none", GETStatus: nativeGetCorpusGETUnprocessable,
			GETRefusal: nativeGetCorpusRefusalUnsupported, Demo: "refuse",
		},
		{
			ID: "refuse-gridlines-off", File: "refuse-gridlines-off.xlsx",
			Title: "showGridLines=0 remains view geometry",
			Extract: "ok", RequiredUnsupported: []string{"SHEET_VIEW_GEOMETRY"},
			ForbiddenUnsupported: []string{"STYLE_RECORD_ATTRIBUTES"},
			DrawingCoverage: "none", GETStatus: nativeGetCorpusGETCompileFailed, Demo: "refuse",
		},
		{
			ID: "refuse-freeze-pane", File: "refuse-freeze-pane.xlsx",
			Title: "Frozen pane remains view geometry",
			Extract: "ok", RequiredUnsupported: []string{"SHEET_VIEW_GEOMETRY"},
			ForbiddenUnsupported: []string{"STYLE_RECORD_ATTRIBUTES"},
			DrawingCoverage: "none", GETStatus: nativeGetCorpusGETCompileFailed, Demo: "refuse",
		},
		{
			ID: "refuse-quote-prefix", File: "refuse-quote-prefix.xlsx",
			Title: "quotePrefix=1 remains style-record authority",
			Extract: "ok", RequiredUnsupported: []string{"STYLE_RECORD_ATTRIBUTES"},
			ForbiddenUnsupported: []string{"SHEET_VIEW_GEOMETRY"},
			DrawingCoverage: "none", GETStatus: nativeGetCorpusGETCompileFailed, Demo: "refuse",
		},
		{
			ID: "refuse-shape-drawing", File: "refuse-shape-drawing.xlsx",
			Title: "Shape drawing is not a chart-covered GET overlay",
			Extract: "ok", RequiredUnsupported: []string{"DRAWING_REFERENCE"},
			ForbiddenUnsupported: forbiddenNoViewStyle, DrawingCoverage: "refuse",
			GETStatus: nativeGetCorpusGETUnprocessable, GETRefusal: nativeGetCorpusRefusalUnsupported, Demo: "refuse",
		},
	}
}

// NativeGetCorpusCaseByID returns the manifest row for id.
func NativeGetCorpusCaseByID(id string) (NativeGetCorpusCase, bool) {
	for _, item := range NativeGetCorpusCases() {
		if item.ID == id {
			return item, true
		}
	}
	return NativeGetCorpusCase{}, false
}

// BuildNativeGetCorpusFile returns the .xlsx bytes for a corpus id.
func BuildNativeGetCorpusFile(id string) ([]byte, NativeGetCorpusCase, error) {
	spec, ok := NativeGetCorpusCaseByID(id)
	if !ok {
		return nil, NativeGetCorpusCase{}, fmt.Errorf("xlsxpatch: unknown native GET corpus id %q", id)
	}
	font := "DejaVu Sans"
	if id == "refuse-calibri" {
		font = "Calibri"
	}
	sheetViews := ""
	xfExtra := ""
	stylesOverride := ""
	switch id {
	case "pass-excel-defaults", "pass-excel-defaults-chart":
		sheetViews = `<sheetViews><sheetView showGridLines="1" workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>`
		xfExtra = ` pivotButton="0" quotePrefix="0"`
	case "refuse-gridlines-off":
		sheetViews = `<sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>`
	case "refuse-freeze-pane":
		sheetViews = `<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="B2" sqref="B2"/></sheetView></sheetViews>`
	case "refuse-quote-prefix":
		xfExtra = ` quotePrefix="1"`
	case "refuse-apply-flags":
		stylesOverride = nativeGetCorpusMissingApplyStyles(font)
	}
	entries := nativeGetCorpusWorkbook(font, sheetViews, xfExtra, stylesOverride)
	if id == "pass-empty-inline-str" {
		sheet := string(entries["xl/worksheets/sheet1.xml"])
		sheet = strings.Replace(sheet, `<dimension ref="A1:C5"/>`, `<dimension ref="A1:C19"/>`, 1)
		sheet = strings.Replace(sheet, `</sheetData>`, `<row r="19"><c r="B19" t="inlineStr"/></row></sheetData>`, 1)
		entries["xl/worksheets/sheet1.xml"] = []byte(sheet)
	}
	raw, err := nativeGetCorpusZip(entries)
	if err != nil {
		return nil, spec, err
	}
	switch id {
	case "pass-excel-defaults-chart":
		raw, err = AddChart(raw, ChartWriteSpec{
			SheetName: "Data",
			Type:      "column",
			Title:     "Revenue by Quarter",
			Series: []WriteSeries{
				{Name: "Revenue", NameRef: "Data!$B$1", CategoriesRef: "Data!$A$2:$A$5", ValuesRef: "Data!$B$2:$B$5"},
			},
			Anchor: ChartAnchor{FromCol: 4, FromRow: 1, ToCol: 10, ToRow: 16},
		})
	case "refuse-shape-drawing":
		raw, err = AddShape(raw, ShapeWriteSpec{
			SheetName: "Data",
			Kind:      "rect",
			Text:      "Note",
			Fill:      "#1F4E78",
			Stroke:    "#000000",
			Anchor:    ShapeAnchor{FromCol: 4, FromRow: 1, ToCol: 8, ToRow: 8},
		})
	}
	if err != nil {
		return nil, spec, err
	}
	return raw, spec, nil
}

func nativeGetCorpusWorkbook(fontName, sheetViews, xfExtra, stylesOverride string) map[string][]byte {
	styles := stylesOverride
	if styles == "" {
		styles = `<styleSheet xmlns="` + nativeGetCorpusSSNS + `">` +
			`<fonts count="1"><font><name val="` + fontName + `"/><sz val="11"/><color rgb="FF000000"/></font></fonts>` +
			`<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
			`<borders count="1"><border/></borders>` +
			`<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
			`<cellXfs count="2">` +
			`<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"` + xfExtra + `/>` +
			`<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"` + xfExtra + `><alignment horizontal="center" vertical="bottom"/></xf>` +
			`</cellXfs>` +
			`<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
			`</styleSheet>`
	}
	sheet := `<worksheet xmlns="` + nativeGetCorpusSSNS + `">` +
		`<dimension ref="A1:C5"/>` +
		sheetViews +
		`<sheetFormatPr baseColWidth="8" defaultColWidth="12" defaultRowHeight="15" customHeight="0" zeroHeight="0"/>` +
		`<cols><col min="1" max="1" width="14" customWidth="1"/><col min="2" max="3" width="12" customWidth="1"/></cols>` +
		`<sheetData>` +
		`<row r="1" ht="18" customHeight="1">` +
		`<c r="A1" s="1" t="inlineStr"><is><t>Quarter</t></is></c>` +
		`<c r="B1" s="1" t="inlineStr"><is><t>Revenue</t></is></c>` +
		`<c r="C1" s="1" t="inlineStr"><is><t>Costs</t></is></c>` +
		`</row>` +
		`<row r="2"><c r="A2" t="inlineStr"><is><t>Q1</t></is></c><c r="B2"><v>10</v></c><c r="C2"><v>4</v></c></row>` +
		`<row r="3"><c r="A3" t="inlineStr"><is><t>Q2</t></is></c><c r="B3"><v>20</v></c><c r="C3"><v>8</v></c></row>` +
		`<row r="4"><c r="A4" t="inlineStr"><is><t>Q3</t></is></c><c r="B4"><v>15</v></c><c r="C4"><v>6</v></c></row>` +
		`<row r="5"><c r="A5" t="inlineStr"><is><t>Q4</t></is></c><c r="B5"><v>25</v></c><c r="C5"><v>10</v></c></row>` +
		`</sheetData></worksheet>`
	return map[string][]byte{
		"[Content_Types].xml": []byte(`<Types xmlns="` + nativeGetCorpusCTNS + `">` +
			`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
			`<Default Extension="xml" ContentType="application/xml"/>` +
			`<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
			`<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
			`<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
			`</Types>`),
		"_rels/.rels": []byte(`<Relationships xmlns="` + nativeGetCorpusPkgNS + `"><Relationship Id="rOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
		"xl/workbook.xml": []byte(`<workbook xmlns="` + nativeGetCorpusSSNS + `" xmlns:r="` + nativeGetCorpusRelNS + `"><bookViews><workbookView/></bookViews><sheets>` +
			`<sheet name="Data" sheetId="1" r:id="rSheet"/></sheets></workbook>`),
		"xl/_rels/workbook.xml.rels": []byte(`<Relationships xmlns="` + nativeGetCorpusPkgNS + `">` +
			`<Relationship Id="rSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>` +
			`<Relationship Id="rStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
			`</Relationships>`),
		"xl/styles.xml":           []byte(styles),
		"xl/worksheets/sheet1.xml": []byte(sheet),
	}
}

func nativeGetCorpusMissingApplyStyles(fontName string) string {
	return `<styleSheet xmlns="` + nativeGetCorpusSSNS + `">` +
		`<numFmts count="1"><numFmt numFmtId="164" formatCode="$#,##0"/></numFmts>` +
		`<fonts count="2">` +
		`<font><name val="` + fontName + `"/><sz val="11"/><color rgb="FF000000"/></font>` +
		`<font><name val="` + fontName + `"/><b val="1"/><color rgb="FFFFFFFF"/><sz val="11"/></font>` +
		`</fonts>` +
		`<fills count="2"><fill><patternFill patternType="none"/></fill>` +
		`<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/></patternFill></fill></fills>` +
		`<borders count="2"><border/><border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/></border></borders>` +
		`<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
		`<cellXfs count="4">` +
		`<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
		`<xf numFmtId="0" fontId="1" fillId="1" borderId="0" applyAlignment="1" xfId="0"><alignment horizontal="center"/></xf>` +
		`<xf numFmtId="0" fontId="0" fillId="0" borderId="1" applyAlignment="1" xfId="0"><alignment horizontal="center"/></xf>` +
		`<xf numFmtId="164" fontId="0" fillId="0" borderId="1" applyAlignment="1" xfId="0"><alignment horizontal="right"/></xf>` +
		`</cellXfs>` +
		`<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
		`</styleSheet>`
}

func nativeGetCorpusZip(entries map[string][]byte) ([]byte, error) {
	names := make([]string, 0, len(entries))
	for name := range entries {
		names = append(names, name)
	}
	sort.Strings(names)
	var buf bytes.Buffer
	writer := zip.NewWriter(&buf)
	for _, name := range names {
		part, err := writer.CreateHeader(&zip.FileHeader{Name: name, Method: zip.Deflate})
		if err != nil {
			return nil, err
		}
		if _, err := part.Write(entries[name]); err != nil {
			return nil, err
		}
	}
	if err := writer.Close(); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

func nativeGetCorpusHasCode(items []NativeWorkbookUnsupportedV2, code string) bool {
	for _, item := range items {
		if item.Code == code {
			return true
		}
	}
	return false
}

func nativeGetCorpusCodes(items []NativeWorkbookUnsupportedV2) []string {
	seen := map[string]bool{}
	var codes []string
	for _, item := range items {
		if item.Code == "" || seen[item.Code] {
			continue
		}
		seen[item.Code] = true
		codes = append(codes, item.Code)
	}
	sort.Strings(codes)
	return codes
}

func nativeGetCorpusCheckUnsupported(items []NativeWorkbookUnsupportedV2, spec NativeGetCorpusCase) error {
	for _, code := range spec.RequiredUnsupported {
		if !nativeGetCorpusHasCode(items, code) {
			return fmt.Errorf("%s missing required unsupported %s (got %s)", spec.ID, code, strings.Join(nativeGetCorpusCodes(items), ","))
		}
	}
	for _, code := range spec.ForbiddenUnsupported {
		if nativeGetCorpusHasCode(items, code) {
			return fmt.Errorf("%s inventoried forbidden unsupported %s (got %s)", spec.ID, code, strings.Join(nativeGetCorpusCodes(items), ","))
		}
	}
	return nil
}
