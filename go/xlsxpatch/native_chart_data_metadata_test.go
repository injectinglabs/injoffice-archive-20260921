package xlsxpatch

import (
	"bytes"
	"encoding/xml"
	"os"
	"strings"
	"testing"
)

func TestNativeChartDataMetadataClosedSubtrees(t *testing.T) {
	for _, namespace := range []string{spreadsheetMLTransitional, spreadsheetMLStrict} {
		for _, test := range []struct {
			source string
			want   bool
		}{
			{`<bookViews><workbookView/></bookViews>`, true},
			{`<bookViews><workbookView visibility="hidden" minimized="true" showHorizontalScroll="0" showVerticalScroll="1" showSheetTabs="false" xWindow="-2147483648" yWindow="2147483647" windowWidth="4294967295" windowHeight="0" tabRatio="600" firstSheet="0" activeTab="2" autoFilterDateGrouping="0"/><workbookView/></bookViews>`, true},
			{`<dimension ref="A1:XFD1048576"/>`, true}, {`<dimension ref="B2"/>`, true},
			{`<bookViews/>`, false},
			{"<bookViews> \t\r\n<workbookView> \t\n</workbookView></bookViews>", true},
			{"<bookViews>\u00a0<workbookView/></bookViews>", false},
			{"<bookViews><workbookView>\u2000</workbookView></bookViews>", false}, {`<bookViews extra="1"><workbookView/></bookViews>`, false},
			{`<bookViews><workbookView><extLst/></workbookView></bookViews>`, false},
			{`<bookViews><workbookView unknown="1"/></bookViews>`, false},
			{`<bookViews><workbookView activeTab="0" activeTab="1"/></bookViews>`, false},
			{`<bookViews><workbookView xmlns:f="urn:foreign" f:activeTab="0"/></bookViews>`, false},
			{`<bookViews><workbookView xmlns="urn:foreign"/></bookViews>`, false},
			{`<bookViews><workbookView visibility="invalid"/></bookViews>`, false},
			{`<bookViews><workbookView xWindow="2147483648"/></bookViews>`, false},
			{`<bookViews><workbookView activeTab="-1"/></bookViews>`, false},
			{`<bookViews><workbookView minimized="yes"/></bookViews>`, false},
			{`<bookViews>text<workbookView/></bookViews>`, false},
			{`<dimension/>`, false}, {`<dimension ref="A1" extra="x"/>`, false},
			{`<dimension ref="A1" ref="B2"/>`, false}, {`<dimension ref="B2:A1"/>`, false},
			{`<dimension ref="A1"><unknown/></dimension>`, false}, {`<dimension ref="XFE1"/>`, false},
			{`<dimension ref="A1" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x"/>`, false},
		} {
			index := strings.IndexAny(test.source, " />")
			source := test.source[:index] + ` xmlns="` + namespace + `"` + test.source[index:]
			decoder := xml.NewDecoder(strings.NewReader(source))
			token, err := decoder.Token()
			if err != nil {
				t.Fatal(err)
			}
			got, err := consumeNativeChartDataMetadata(decoder, token.(xml.StartElement), namespace)
			if err != nil || got != test.want {
				t.Fatalf("%s: got %v, %v; want %v", source, got, err, test.want)
			}
		}
	}
}

func TestNativeChartDataMetadataRetainsUnknownNeighborsAndSource(t *testing.T) {
	for _, strict := range []bool{false, true} {
		parts := nativeWorkbookFixture(strict)
		parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `<sheetData>`, `<unknown/><sheetData>`, 1)
		data := buildZip(t, parts)
		original := append([]byte(nil), data...)
		workbook, err := ExtractNativeWorkbookV1(data)
		if err != nil {
			t.Fatal(err)
		}
		for _, code := range []string{"WORKBOOK_VIEW_METADATA", "WORKSHEET_DIMENSION_METADATA", "UNMODELED_WORKBOOK_FEATURE", "UNMODELED_WORKSHEET_FEATURE"} {
			if !hasNativeWorkbookUnsupported(workbook, code) {
				t.Fatalf("missing %s", code)
			}
		}
		if !bytes.Equal(data, original) {
			t.Fatal("source modified")
		}
		v2, err := ExtractNativeWorkbookV2(data)
		if err != nil {
			t.Fatal(err)
		}
		codes := map[string]bool{}
		for _, item := range v2.Unsupported {
			codes[item.Code] = true
		}
		for _, code := range []string{"WORKBOOK_VIEW_METADATA", "WORKSHEET_DIMENSION_METADATA", "UNMODELED_WORKBOOK_FEATURE", "UNMODELED_WORKSHEET_FEATURE"} {
			if !codes[code] {
				t.Fatalf("V2 missing %s", code)
			}
		}
	}
}

func TestNativeChartDataMetadataOrdinaryWorkbook(t *testing.T) {
	data, err := os.ReadFile("testdata/native-get-corpus/pass-excel-defaults.xlsx")
	if err != nil {
		t.Fatal(err)
	}
	workbook, err := ExtractNativeWorkbookV2(data)
	if err != nil {
		t.Fatal(err)
	}
	codes := map[string]bool{}
	for _, item := range workbook.Unsupported {
		codes[item.Code] = true
	}
	if !codes["WORKBOOK_VIEW_METADATA"] || !codes["WORKSHEET_DIMENSION_METADATA"] || codes["UNMODELED_WORKBOOK_FEATURE"] || codes["UNMODELED_WORKSHEET_FEATURE"] {
		t.Fatalf("unexpected codes %v", codes)
	}
}

func TestNativeChartDataMetadataPreservesCountaExactSourceGate(t *testing.T) {
	parts := nativePrintCountaFixture(false, `<c r="A1"><v>1</v></c>`, "")
	wb, _ := nativePrintCountaExtract(t, parts)
	found := false
	for _, item := range wb.Unsupported {
		if item.Code == "WORKBOOK_VIEW_METADATA" {
			found = true
		}
	}
	if !found {
		t.Fatal("fixture did not exercise new classification")
	}
	for _, replacement := range []string{`<bookViews><workbookView activeTab="1"/></bookViews>`, `<bookViews unknown="1"><workbookView/></bookViews>`, `<bookViews><workbookView><unknown/></workbookView></bookViews>`} {
		raw := strings.Replace(parts["Book/Workbook.xml"], `<bookViews><workbookView/></bookViews>`, replacement, 1)
		if newNativePrintCountaSourceContext(wb, []byte(raw), wb.Source.PackageSHA256) != nil {
			t.Fatalf("existing source gate widened: %s", replacement)
		}
	}
}
