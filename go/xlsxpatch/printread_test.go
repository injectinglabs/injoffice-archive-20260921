package xlsxpatch

import (
	"encoding/json"
	"strings"
	"testing"
)

func TestReadPrintSetups_HydratesWrittenSubset(t *testing.T) {
	original := buildZip(t, fixtureWorkbook(false))
	written, err := SetPrintSetup(original, PrintSetup{
		SheetName: "Data", PrintAreaRef: "A1:F20", Orientation: "landscape",
		PaperSize: "A4", Scale: 125,
		Margins:            &PrintMargins{Left: 1, Right: 1, Top: 1, Bottom: 1, Header: 0.5, Footer: 0.5},
		HeaderCenter:       "Quarterly &P",
		FooterCenter:       "&F",
		HorizontalCentered: true,
		VerticalCentered:   true,
		PrintGridlines:     true,
		PrintHeadings:      true,
	})
	if err != nil {
		t.Fatal(err)
	}

	setups, err := ReadPrintSetups(written)
	if err != nil {
		t.Fatal(err)
	}
	if len(setups) != 1 {
		t.Fatalf("setups = %d, want 1", len(setups))
	}
	got := setups[0]
	if got.SheetName != "Data" || got.PrintAreaRef != "$A$1:$F$20" || got.Orientation != "landscape" || got.PaperSize != "A4" {
		t.Fatalf("unexpected identity/setup: %+v", got)
	}
	if got.Scale == nil || *got.Scale != 125 || got.Margins == nil || got.Margins.Left != 1 {
		t.Fatalf("scale/margins not hydrated: %+v", got)
	}
	if got.OddHeader != "&CQuarterly &P" || got.OddFooter != "&C&F" {
		t.Fatalf("headers not hydrated: %+v", got)
	}
	if !got.HorizontalCentered || !got.VerticalCentered || !got.PrintGridlines || !got.PrintHeadings {
		t.Fatalf("print options not hydrated: %+v", got)
	}
	if len(got.Warnings) != 0 {
		t.Fatalf("unexpected warnings: %v", got.Warnings)
	}
	encoded, err := json.Marshal(got)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(encoded), `"margins":{"left":1,"right":1,"top":1,"bottom":1,"header":0.5,"footer":0.5}`) {
		t.Fatalf("print hydration JSON does not match the TypeScript wire shape: %s", encoded)
	}
}

func TestReadPrintSetups_RepeatedTitlesAndUnsupportedWarnings(t *testing.T) {
	entries := fixtureWorkbook(false)
	entries["xl/workbook.xml"] = `<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$C$9</definedName><definedName name="_xlnm.Print_Titles" localSheetId="0">Data!$1:$2,Data!$A:$B</definedName></definedNames></workbook>`
	entries["xl/worksheets/sheet1.xml"] = `<worksheet><sheetData/><printOptions horizontalCentered="true" gridLines="1" draft="1"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup orientation="landscape" paperSize="99" fitToWidth="2" fitToHeight="0" blackAndWhite="1"/><headerFooter differentOddEven="1"><oddHeader>&amp;LPrivate&amp;CVisible</oddHeader></headerFooter></worksheet>`

	setups, err := ReadPrintSetups(buildZip(t, entries))
	if err != nil {
		t.Fatal(err)
	}
	got := setups[0]
	if got.PrintAreaRef != "$A$1:$C$9" || got.RepeatRowsRef != "$1:$2" || got.RepeatColumnsRef != "$A:$B" {
		t.Fatalf("defined names not hydrated: %+v", got)
	}
	if got.FitToWidth == nil || *got.FitToWidth != 2 || got.FitToHeight == nil || *got.FitToHeight != 0 {
		t.Fatalf("fit settings not hydrated: %+v", got)
	}
	joined := strings.Join(got.Warnings, " ")
	for _, expected := range []string{"draft", "paper size", "blackAndWhite", "differentOddEven"} {
		if !strings.Contains(joined, expected) {
			t.Errorf("warnings %q missing %q", joined, expected)
		}
	}
}

func TestReadPrintSetups_RejectsUnsafeOrMalformedXML(t *testing.T) {
	entries := fixtureWorkbook(false)
	entries["xl/worksheets/sheet1.xml"] = `<!DOCTYPE worksheet><worksheet><sheetData/></worksheet>`
	if _, err := ReadPrintSetups(buildZip(t, entries)); err == nil || !strings.Contains(err.Error(), "DOCTYPE") {
		t.Fatalf("DOCTYPE must be rejected, got %v", err)
	}

	entries = fixtureWorkbook(false)
	entries["xl/worksheets/sheet1.xml"] = `<worksheet><sheetData/><pageSetup/><pageSetup/></worksheet>`
	if _, err := ReadPrintSetups(buildZip(t, entries)); err == nil || !strings.Contains(err.Error(), "duplicate pageSetup") {
		t.Fatalf("duplicate print elements must be rejected, got %v", err)
	}
}
