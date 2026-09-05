package xlsxpatch

import (
	"strings"
	"testing"
)

func TestSetPrintSetup_WritesTrioAndPrintArea(t *testing.T) {
	orig := buildZip(t, fixtureWorkbook(false))
	out, err := SetPrintSetup(orig, PrintSetup{
		SheetName:    "Data",
		PrintAreaRef: "A1:F20",
		Orientation:  "landscape",
		FitToWidth:   1,
		HeaderCenter: "Quarterly Report",
		FooterCenter: "Page &P of &N",
	})
	if err != nil {
		t.Fatal(err)
	}

	sheet := readEntry(t, out, "xl/worksheets/sheet1.xml")
	for _, want := range []string{
		`<pageMargins left="0.7"`,
		`<pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/>`,
		`<oddHeader>&amp;CQuarterly Report</oddHeader>`,
		`<oddFooter>&amp;CPage &amp;P of &amp;N</oddFooter>`,
		`<pageSetUpPr fitToPage="1"/>`,
	} {
		if !strings.Contains(sheet, want) {
			t.Errorf("worksheet missing %s\n%s", want, sheet)
		}
	}
	// Order: margins before setup before headerFooter.
	if !(strings.Index(sheet, "<pageMargins") < strings.Index(sheet, "<pageSetup") &&
		strings.Index(sheet, "<pageSetup") < strings.Index(sheet, "<headerFooter")) {
		t.Errorf("print trio out of schema order: %s", sheet)
	}

	wb := readEntry(t, out, "xl/workbook.xml")
	if !strings.Contains(wb, `<definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$F$20</definedName>`) {
		t.Errorf("workbook missing print area: %s", wb)
	}
}

func TestSetPrintSetup_ReplacesExistingAndClearsArea(t *testing.T) {
	orig := buildZip(t, fixtureWorkbook(false))
	one, err := SetPrintSetup(orig, PrintSetup{SheetName: "Data", PrintAreaRef: "A1:B2", Orientation: "landscape"})
	if err != nil {
		t.Fatal(err)
	}
	// Second apply: portrait, custom margins, NO print area — must replace,
	// not accumulate.
	two, err := SetPrintSetup(one, PrintSetup{
		SheetName:   "Data",
		Orientation: "portrait",
		Margins:     &PrintMargins{Left: 1, Right: 1, Top: 1, Bottom: 1, Header: 0.5, Footer: 0.5},
	})
	if err != nil {
		t.Fatal(err)
	}
	sheet := readEntry(t, two, "xl/worksheets/sheet1.xml")
	if strings.Count(sheet, "<pageSetup") != 1 || strings.Count(sheet, "<pageMargins") != 1 {
		t.Errorf("print elements accumulated: %s", sheet)
	}
	if !strings.Contains(sheet, `orientation="portrait"`) || !strings.Contains(sheet, `left="1"`) {
		t.Errorf("second setup not applied: %s", sheet)
	}
	wb := readEntry(t, two, "xl/workbook.xml")
	if strings.Contains(wb, "_xlnm.Print_Area") {
		t.Errorf("cleared print area still present: %s", wb)
	}
}

func TestSetPrintSetup_ChartAndDrawingOrderPreserved(t *testing.T) {
	// Print trio must land BEFORE <drawing> per schema.
	withChart, err := AddChart(buildZip(t, fixtureWorkbook(false)), writeSpec("column"))
	if err != nil {
		t.Fatal(err)
	}
	out, err := SetPrintSetup(withChart, PrintSetup{SheetName: "Data", Orientation: "landscape"})
	if err != nil {
		t.Fatal(err)
	}
	sheet := readEntry(t, out, "xl/worksheets/sheet1.xml")
	if !(strings.Index(sheet, "<pageSetup") < strings.Index(sheet, "<drawing ")) {
		t.Errorf("pageSetup must precede drawing: %s", sheet)
	}
	// Chart untouched (Apply verified, but assert reader still sees it).
	charts, err := ReadCharts(out)
	if err != nil || len(charts) != 1 {
		t.Fatalf("chart lost after print setup: %v", err)
	}
}

func TestSetPrintSetup_PaperScaleAndPrintOptions(t *testing.T) {
	orig := buildZip(t, fixtureWorkbook(false))
	out, err := SetPrintSetup(orig, PrintSetup{
		SheetName: "Data", PaperSize: "A4", Scale: 125,
		HorizontalCentered: true, VerticalCentered: true,
		PrintGridlines: true, PrintHeadings: true,
	})
	if err != nil {
		t.Fatal(err)
	}
	sheet := readEntry(t, out, "xl/worksheets/sheet1.xml")
	for _, want := range []string{
		`<printOptions horizontalCentered="1" verticalCentered="1" gridLines="1" headings="1"/>`,
		`<pageSetup orientation="portrait" scale="125" paperSize="9"/>`,
	} {
		if !strings.Contains(sheet, want) {
			t.Errorf("worksheet missing %s\n%s", want, sheet)
		}
	}
	if !(strings.Index(sheet, "<printOptions") < strings.Index(sheet, "<pageMargins")) {
		t.Errorf("printOptions must precede pageMargins: %s", sheet)
	}

	cleared, err := SetPrintSetup(out, PrintSetup{SheetName: "Data", PaperSize: "Letter"})
	if err != nil {
		t.Fatal(err)
	}
	sheet = readEntry(t, cleared, "xl/worksheets/sheet1.xml")
	if strings.Contains(sheet, "<printOptions") || strings.Contains(sheet, `scale="125"`) {
		t.Errorf("stale print options or scale survived replacement: %s", sheet)
	}
	if !strings.Contains(sheet, `paperSize="1"`) {
		t.Errorf("letter paper code missing: %s", sheet)
	}
}

func TestSetPrintSetup_Rejections(t *testing.T) {
	orig := buildZip(t, fixtureWorkbook(false))
	if _, err := SetPrintSetup(orig, PrintSetup{SheetName: "Data", Orientation: "diagonal"}); err == nil {
		t.Error("bad orientation must be rejected")
	}
	if _, err := SetPrintSetup(orig, PrintSetup{SheetName: "Data", PrintAreaRef: "junk"}); err == nil {
		t.Error("bad print area must be rejected")
	}
	if _, err := SetPrintSetup(orig, PrintSetup{SheetName: "Ghost"}); err == nil {
		t.Error("unknown sheet must be rejected")
	}
	if _, err := SetPrintSetup(orig, PrintSetup{SheetName: "Data", PaperSize: "A0"}); err == nil {
		t.Error("unknown paper size must be rejected")
	}
	if _, err := SetPrintSetup(orig, PrintSetup{SheetName: "Data", Scale: 9}); err == nil {
		t.Error("out-of-range scale must be rejected")
	}
	if _, err := SetPrintSetup(orig, PrintSetup{SheetName: "Data", Scale: 100, FitToWidth: 1}); err == nil {
		t.Error("scale plus fit must be rejected")
	}
}
