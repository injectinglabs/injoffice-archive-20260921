package xlsxpatch

import (
	"sort"
	"strings"
	"testing"
)

const neutralityWorksheet = `<worksheet xmlns="%NS%" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="%AC%" mc:Ignorable="x14ac">` +
	`<sheetViews><sheetView showGridLines="0" tabSelected="1" zoomScale="220" workbookViewId="0"><pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/><selection activeCell="B3" sqref="B3"/></sheetView></sheetViews>` +
	`<sheetFormatPr defaultRowHeight="14.4" customHeight="1" outlineLevelRow="2" thickBottom="1" x14ac:dyDescent="0.3"/>` +
	`<cols><col min="1" max="1" width="12.5" customWidth="1" collapsed="0" outlineLevel="1"/></cols>` +
	`<sheetData><row r="1" spans="1:3" thickTop="1" thickBot="1" customFormat="1" s="2" outlineLevel="1" x14ac:dyDescent="0.3"/></sheetData>` +
	`<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`

func neutralitySource(ns string) string {
	return strings.NewReplacer("%NS%", ns, "%AC%", nativeRowDescentNamespace).Replace(neutralityWorksheet)
}

func neutralityCodes(t *testing.T, source string) []string {
	t.Helper()
	result := previewNativeDimensionNeutrality([]byte(source), "xl/worksheets/sheet1.xml")
	if result.SheetPart != "xl/worksheets/sheet1.xml" || result.Policy != nativeDimensionNeutralityPolicy {
		t.Fatalf("unexpected identity: %+v", result)
	}
	if len(result.Warnings) != 1 || result.Warnings[0] == "" {
		t.Fatalf("missing disclosure: %+v", result)
	}
	if !sort.StringsAreSorted(result.Codes) {
		t.Fatalf("codes are not canonically ordered: %+v", result.Codes)
	}
	return result.Codes
}

func hasNeutralCode(codes []string, code string) bool {
	for _, item := range codes {
		if item == code {
			return true
		}
	}
	return false
}

// Display state, outline bookkeeping, border flags, descent metadata and
// markup-compatibility attributes clear every code they raise.
func TestDimensionNeutralityQualifiesNonDimensionalMarkup(t *testing.T) {
	for _, ns := range []string{spreadsheetMLTransitional, spreadsheetMLStrict} {
		codes := neutralityCodes(t, neutralitySource(ns))
		for _, want := range []string{"COLS_ATTRIBUTES", "COLUMN_DIMENSION_EXTRAS", "ROW_DIMENSION_EXTRAS", "SHEET_FORMAT_EXTRAS", "SHEET_VIEW_GEOMETRY", "WORKSHEET_ATTRIBUTES"} {
			if !hasNeutralCode(codes, want) {
				t.Fatalf("%s was not cleared for %s: %v", want, ns, codes)
			}
		}
		if len(codes) != 6 {
			t.Fatalf("unexpected cleared codes for %s: %v", ns, codes)
		}
	}
}

// Every refusal that protects the dimension projection survives.
func TestDimensionNeutralityKeepsDimensionBearingRefusals(t *testing.T) {
	ns := spreadsheetMLTransitional
	for _, testCase := range []struct {
		name    string
		replace [2]string
		keeps   string
	}{
		// Layout direction is not display-only: it mirrors the sheet.
		{"right-to-left", [2]string{`showGridLines="0"`, `showGridLines="0" rightToLeft="1"`}, "SHEET_VIEW_GEOMETRY"},
		// An unknown sheetView option may be anything at all.
		{"unknown-view-attribute", [2]string{`zoomScale="220"`, `zoomScale="220" futureView="1"`}, "SHEET_VIEW_GEOMETRY"},
		// Extension lists and pivot selections are not enumerated window state.
		{"view-extension", [2]string{`<selection activeCell="B3" sqref="B3"/>`, `<extLst/>`}, "SHEET_VIEW_GEOMETRY"},
		{"foreign-view-child", [2]string{`<selection activeCell="B3" sqref="B3"/>`, `<selection xmlns="urn:foreign"/>`}, "SHEET_VIEW_GEOMETRY"},
		{"second-sheet-view", [2]string{`</sheetViews>`, `<sheetView workbookViewId="1"/></sheetViews>`}, "SHEET_VIEW_GEOMETRY"},
		{"sheet-views-attribute", [2]string{`<sheetViews>`, `<sheetViews future="1">`}, "SHEET_VIEW_GEOMETRY"},
		// Unknown sheetFormatPr markup may carry a default metric.
		{"unknown-sheet-format", [2]string{`outlineLevelRow="2"`, `outlineLevelRow="2" futureWidth="9"`}, "SHEET_FORMAT_EXTRAS"},
		{"foreign-sheet-format-attribute", [2]string{`x14ac:dyDescent="0.3"/>`, `xmlns:z="urn:foreign" z:height="9"/>`}, "SHEET_FORMAT_EXTRAS"},
		{"sheet-format-children", [2]string{`x14ac:dyDescent="0.3"/>`, `x14ac:dyDescent="0.3"><ext/></sheetFormatPr>`}, "SHEET_FORMAT_EXTRAS"},
		// Phonetic guides participate in automatic sizing, so they keep refusing.
		{"column-phonetic", [2]string{`collapsed="0"`, `collapsed="0" phonetic="1"`}, "COLUMN_DIMENSION_EXTRAS"},
		{"row-phonetic", [2]string{`outlineLevel="1" x14ac:dyDescent`, `outlineLevel="1" ph="1" x14ac:dyDescent`}, "ROW_DIMENSION_EXTRAS"},
		{"unknown-column-attribute", [2]string{`width="12.5"`, `width="12.5" futureWidth="3"`}, "COLUMN_DIMENSION_EXTRAS"},
		{"unknown-row-attribute", [2]string{`spans="1:3"`, `spans="1:3" futureHeight="3"`}, "ROW_DIMENSION_EXTRAS"},
		{"foreign-column-child", [2]string{`<cols>`, `<cols><futureCol/>`}, "COLUMN_DIMENSION_EXTRAS"},
		{"foreign-row-child", [2]string{`<sheetData>`, `<sheetData><futureRow/>`}, "ROW_DIMENSION_EXTRAS"},
	} {
		source := strings.Replace(neutralitySource(ns), testCase.replace[0], testCase.replace[1], 1)
		if source == neutralitySource(ns) {
			t.Fatalf("%s: fixture substitution did not apply", testCase.name)
		}
		if codes := neutralityCodes(t, source); hasNeutralCode(codes, testCase.keeps) {
			t.Fatalf("%s falsely cleared %s: %v", testCase.name, testCase.keeps, codes)
		}
	}
}

// An unreadable, oversized or foreign part clears nothing at all.
func TestDimensionNeutralityRefusesUnreadableParts(t *testing.T) {
	for _, source := range []string{
		"",
		"not xml",
		`<worksheet xmlns="urn:foreign"/>`,
		`<chartsheet xmlns="` + spreadsheetMLTransitional + `"/>`,
		`<worksheet xmlns="` + spreadsheetMLTransitional + `">` + strings.Repeat(`<row/>`, 20001) + `</worksheet>`,
	} {
		result := previewNativeDimensionNeutrality([]byte(source), "xl/worksheets/sheet1.xml")
		if len(result.Codes) != 0 || len(result.Warnings) != 1 || result.Warnings[0] != nativeDimensionNeutralityUnavailable {
			t.Fatalf("unreadable part was qualified: %+v", result)
		}
	}
}

// The evidence joins the sheet parts of the workbook it was read from.
func TestDimensionNeutralityJoinsWorksheetParts(t *testing.T) {
	parts := nativeWorkbookFixture(false)
	parts["Charts/chart1.xml"] = previewChartFixture()
	objects, err := InspectNativeWorkbookObjectsV1(buildZip(t, parts))
	if err != nil {
		t.Fatalf("objects preview failed: %v", err)
	}
	if len(objects.DimensionNeutrality) == 0 {
		t.Fatal("no dimension neutrality evidence was produced")
	}
	seen := map[string]bool{}
	for _, entry := range objects.DimensionNeutrality {
		if entry.Policy != nativeDimensionNeutralityPolicy || entry.SheetPart == "" || seen[entry.SheetPart] {
			t.Fatalf("unusable evidence entry: %+v", entry)
		}
		seen[entry.SheetPart] = true
		for _, code := range entry.Codes {
			switch code {
			case "COLS_ATTRIBUTES", "COLUMN_DIMENSION_EXTRAS", "ROW_DIMENSION_EXTRAS", "SHEET_FORMAT_EXTRAS", "SHEET_VIEW_GEOMETRY", "WORKSHEET_ATTRIBUTES":
			default:
				t.Fatalf("evidence cleared a code outside the policy: %s", code)
			}
		}
	}
}
