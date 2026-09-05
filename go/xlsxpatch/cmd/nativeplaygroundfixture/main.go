// Command nativeplaygroundfixture writes the repository-owned InjOffice
// launch-readiness workbook plus its native extraction and chart sidecars.
// It is a deterministic generator, not a hosted extractor.
package main

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"

	"github.com/injectinglabs/injoffice/go/xlsxpatch"
)

const (
	ssNS     = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
	relNS    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	pkgRelNS = "http://schemas.openxmlformats.org/package/2006/relationships"
	ctNS     = "http://schemas.openxmlformats.org/package/2006/content-types"
)

func main() {
	if err := run(); err != nil {
		fmt.Fprintln(os.Stderr, "nativeplaygroundfixture:", err)
		os.Exit(1)
	}
}

func run() error {
	root, err := os.Getwd()
	if err != nil {
		return err
	}
	outDir := filepath.Join(root, "apps", "playground", "public", "native-fixture")
	if len(os.Args) > 1 {
		outDir = os.Args[1]
	}
	if err := os.MkdirAll(outDir, 0o755); err != nil {
		return err
	}

	charted, err := buildPlaygroundFixture()
	if err != nil {
		return err
	}

	workbook, err := xlsxpatch.ExtractNativeWorkbookV2(charted)
	if err != nil {
		return err
	}
	workbookJSON, err := xlsxpatch.EncodeNativeWorkbookV2(workbook)
	if err != nil {
		return err
	}
	charts, err := xlsxpatch.ReadCharts(charted)
	if err != nil {
		return err
	}
	anchors, err := xlsxpatch.ReadChartAnchors(charted)
	if err != nil {
		return err
	}
	envelope, err := json.MarshalIndent(struct {
		Charts  []xlsxpatch.ChartInfo            `json:"charts"`
		Anchors map[string]xlsxpatch.ChartAnchor `json:"anchors"`
	}{Charts: charts, Anchors: anchors}, "", "  ")
	if err != nil {
		return err
	}

	if err := os.WriteFile(filepath.Join(outDir, "workbook.json"), append(workbookJSON, '\n'), 0o644); err != nil {
		return err
	}
	if err := os.WriteFile(filepath.Join(outDir, "launch-readiness-plan.xlsx"), charted, 0o644); err != nil {
		return err
	}
	if err := os.WriteFile(filepath.Join(outDir, "charts.json"), append(envelope, '\n'), 0o644); err != nil {
		return err
	}
	fmt.Fprintf(os.Stderr, "wrote %s\n", outDir)
	return nil
}

func buildPlaygroundFixture() ([]byte, error) {
	original, err := buildZip(playgroundWorkbook())
	if err != nil {
		return nil, err
	}
	return xlsxpatch.AddChart(original, xlsxpatch.ChartWriteSpec{
		SheetName: "Launch Readiness",
		Type:      "column",
		Title:     "Budget and spend by workstream",
		Series: []xlsxpatch.WriteSeries{
			{Name: "Budget", NameRef: "'Launch Readiness'!$D$1", CategoriesRef: "'Launch Readiness'!$A$2:$A$7", ValuesRef: "'Launch Readiness'!$D$2:$D$7"},
			{Name: "Spend", NameRef: "'Launch Readiness'!$E$1", CategoriesRef: "'Launch Readiness'!$A$2:$A$7", ValuesRef: "'Launch Readiness'!$E$2:$E$7"},
		},
		Anchor: xlsxpatch.ChartAnchor{FromCol: 7, FromRow: 1, ToCol: 14, ToRow: 18},
	})
}

func playgroundWorkbook() map[string][]byte {
	return map[string][]byte{
		"[Content_Types].xml": []byte(`<Types xmlns="` + ctNS + `">` +
			`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
			`<Default Extension="xml" ContentType="application/xml"/>` +
			`<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
			`<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
			`<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
			`</Types>`),
		"_rels/.rels": []byte(`<Relationships xmlns="` + pkgRelNS + `"><Relationship Id="rOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
		"xl/workbook.xml": []byte(`<workbook xmlns="` + ssNS + `" xmlns:r="` + relNS + `"><bookViews><workbookView/></bookViews><sheets>` +
			`<sheet name="Launch Readiness" sheetId="1" r:id="rSheet"/></sheets></workbook>`),
		"xl/_rels/workbook.xml.rels": []byte(`<Relationships xmlns="` + pkgRelNS + `">` +
			`<Relationship Id="rSheet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>` +
			`<Relationship Id="rStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
			`</Relationships>`),
		"xl/styles.xml": []byte(`<styleSheet xmlns="` + ssNS + `">` +
			`<numFmts count="1"><numFmt numFmtId="164" formatCode="$#,##0"/></numFmts>` +
			`<fonts count="3">` +
			`<font><name val="DejaVu Sans"/><sz val="11"/><color rgb="FF172033"/></font>` +
			`<font><name val="DejaVu Sans"/><sz val="11"/><b/><color rgb="FFFFFFFF"/></font>` +
			`<font><name val="DejaVu Sans"/><sz val="11"/><b/><color rgb="FF172033"/></font>` +
			`</fonts>` +
			`<fills count="4">` +
			`<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>` +
			`<fill><patternFill patternType="solid"><fgColor rgb="FF234F78"/></patternFill></fill>` +
			`<fill><patternFill patternType="solid"><fgColor rgb="FFEAF1F8"/></patternFill></fill>` +
			`</fills>` +
			`<borders count="2"><border/><border><left style="thin"><color rgb="FFD7DEE8"/></left><right style="thin"><color rgb="FFD7DEE8"/></right><top style="thin"><color rgb="FFD7DEE8"/></top><bottom style="thin"><color rgb="FFD7DEE8"/></bottom></border></borders>` +
			`<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
			`<cellXfs count="6">` +
			`<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
			`<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>` +
			`<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>` +
			`<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>` +
			`<xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>` +
			`<xf numFmtId="164" fontId="2" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>` +
			`</cellXfs>` +
			`<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
			`</styleSheet>`),
		"xl/worksheets/sheet1.xml": []byte(`<worksheet xmlns="` + ssNS + `">` +
			`<dimension ref="A1:F8"/>` +
			`<sheetFormatPr baseColWidth="8" defaultColWidth="12" defaultRowHeight="15" customHeight="0" zeroHeight="0"/>` +
			`<cols><col min="1" max="1" width="18" customWidth="1"/><col min="2" max="2" width="16" customWidth="1"/><col min="3" max="3" width="13" customWidth="1"/><col min="4" max="5" width="14" customWidth="1"/><col min="6" max="6" width="24" customWidth="1"/></cols>` +
			`<sheetData>` +
			`<row r="1" ht="28" customHeight="1">` +
			`<c r="A1" s="1" t="inlineStr"><is><t>Workstream</t></is></c>` +
			`<c r="B1" s="1" t="inlineStr"><is><t>Owner</t></is></c>` +
			`<c r="C1" s="1" t="inlineStr"><is><t>Status</t></is></c>` +
			`<c r="D1" s="1" t="inlineStr"><is><t>Budget</t></is></c>` +
			`<c r="E1" s="1" t="inlineStr"><is><t>Spend</t></is></c>` +
			`<c r="F1" s="1" t="inlineStr"><is><t>Next milestone</t></is></c>` +
			`</row>` +
			businessRow(2, "Platform", "Maya Chen", "On track", "185000", "142500", "Load test complete") +
			businessRow(3, "Mobile", "Leo Martin", "At risk", "120000", "103800", "Accessibility audit") +
			businessRow(4, "Analytics", "Priya Shah", "On track", "95000", "67250", "Executive dashboard") +
			businessRow(5, "Security", "Omar Reed", "Review", "80000", "61800", "Pen-test remediation") +
			businessRow(6, "Enablement", "Sofia King", "On track", "55000", "38400", "Partner playbook") +
			businessRow(7, "Support", "Noah Brooks", "Ready", "45000", "29750", "Runbook rehearsal") +
			`<row r="8" ht="20" customHeight="1"><c r="A8" s="4" t="inlineStr"><is><t>Overall</t></is></c><c r="B8" s="4" t="inlineStr"><is><t>Program office</t></is></c><c r="C8" s="4" t="inlineStr"><is><t>On track</t></is></c><c r="D8" s="5"><v>580000</v></c><c r="E8" s="5"><v>443500</v></c><c r="F8" s="4" t="inlineStr"><is><t>Beta launch - Oct 14</t></is></c></row>` +
			`</sheetData></worksheet>`),
	}
}

func businessRow(row int, workstream, owner, status, budget, spend, milestone string) string {
	return fmt.Sprintf(`<row r="%[1]d" ht="20" customHeight="1"><c r="A%[1]d" s="2" t="inlineStr"><is><t>%[2]s</t></is></c><c r="B%[1]d" s="2" t="inlineStr"><is><t>%[3]s</t></is></c><c r="C%[1]d" s="2" t="inlineStr"><is><t>%[4]s</t></is></c><c r="D%[1]d" s="3"><v>%[5]s</v></c><c r="E%[1]d" s="3"><v>%[6]s</v></c><c r="F%[1]d" s="2" t="inlineStr"><is><t>%[7]s</t></is></c></row>`, row, workstream, owner, status, budget, spend, milestone)
}

func buildZip(entries map[string][]byte) ([]byte, error) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	names := make([]string, 0, len(entries))
	for name := range entries {
		names = append(names, name)
	}
	for i := 0; i < len(names); i++ {
		for j := i + 1; j < len(names); j++ {
			if names[j] < names[i] {
				names[i], names[j] = names[j], names[i]
			}
		}
	}
	for _, name := range names {
		w, err := zw.Create(name)
		if err != nil {
			return nil, err
		}
		if _, err := w.Write(entries[name]); err != nil {
			return nil, err
		}
	}
	if err := zw.Close(); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}
