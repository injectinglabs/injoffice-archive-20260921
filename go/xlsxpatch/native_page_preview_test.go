package xlsxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func TestNativeFitPageSettingsRoutedReadOnly(t *testing.T) {
	for _, strict := range []bool{false, true} {
		parts := nativeWorkbookFixture(strict)
		parts["Charts/chart1.xml"] = previewChartFixture()
		raw := parts["Sheets/s1.xml"]
		end := strings.Index(raw, ">") + 1
		raw = raw[:end] + `<sheetPr><pageSetUpPr fitToPage="true"/></sheetPr>` + raw[end:]
		parts["Sheets/s1.xml"] = strings.Replace(raw, `</worksheet>`, `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="portrait" fitToWidth="1" fitToHeight="0"/></worksheet>`, 1)
		source := buildZip(t, parts)
		original := bytes.Clone(source)
		before, err := ExtractNativeWorkbookV2(source)
		if err != nil {
			t.Fatal(err)
		}
		beforeJSON, err := EncodeNativeWorkbookV2(before)
		if err != nil {
			t.Fatal(err)
		}
		got, err := InspectNativeWorkbookObjectsV1(source)
		if err != nil {
			t.Fatal(err)
		}
		if len(got.PageSettings) != 2 || got.PageSettings[0].SheetID != "7" || got.PageSettings[0].SheetPart != "Sheets/s1.xml" || got.PageSettings[0].Status != "available" || got.PageSettings[0].Settings.FitToPage == nil || got.PageSettings[1].Status != "unavailable" {
			t.Fatalf("routed fit lost: %+v", got.PageSettings)
		}
		after, err := ExtractNativeWorkbookV2(source)
		if err != nil {
			t.Fatal(err)
		}
		afterJSON, err := EncodeNativeWorkbookV2(after)
		if err != nil || !bytes.Equal(source, original) || !bytes.Equal(beforeJSON, afterJSON) || bytes.Contains(afterJSON, []byte(`"fit_to_page"`)) {
			t.Fatal("preview changed native mutation authority")
		}
	}
}

func TestNativePageSettingsPageOrder(t *testing.T) {
	for _, ns := range []string{spreadsheetMLTransitional, spreadsheetMLStrict} {
		for _, order := range []string{"downThenOver", "overThenDown"} {
			raw := `<worksheet xmlns="` + ns + `"><pageSetup paperSize="1" orientation="landscape" scale="75" pageOrder="` + order + `"/><pageMargins left="0.4" right="0.6" top="0.8" bottom="1" header="0.2" footer="0.2"/></worksheet>`
			got := previewNativePageSettings([]byte(raw), "sheet.xml", "1")
			if got.Status != "available" || got.Settings.PageOrder != order || got.Settings.Left != 0.4 || got.Settings.Right != 0.6 {
				t.Fatalf("lost explicit page order/margins: %+v", got)
			}
			for _, replacement := range []string{`pageOrder=""`, `pageOrder="diagonal"`, `xmlns:f="urn:foreign" f:pageOrder="overThenDown"`, `pageOrder="overThenDown" pageOrder="downThenOver"`, `pageOrder="overThenDown" fitToWidth="1"`} {
				bad := strings.Replace(raw, `pageOrder="`+order+`"`, replacement, 1)
				if previewNativePageSettings([]byte(bad), "sheet.xml", "1").Status != "unavailable" {
					t.Fatalf("accepted unqualified page setup %s", replacement)
				}
			}
		}
	}
}

func TestNativePageSettingsRejectDuplicateAndForeignOwnersBothDialects(t *testing.T) {
	for _, ns := range []string{spreadsheetMLTransitional, spreadsheetMLStrict} {
		setup := `<pageSetup paperSize="9" orientation="portrait" scale="100"/>`
		margins := `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>`
		for _, extra := range []string{setup, margins, strings.Replace(setup, "<pageSetup", `<pageSetup xmlns="urn:foreign"`, 1), strings.Replace(margins, "<pageMargins", `<pageMargins xmlns="urn:foreign"`, 1)} {
			raw := `<worksheet xmlns="` + ns + `">` + setup + margins + extra + `</worksheet>`
			if previewNativePageSettings([]byte(raw), "sheet.xml", "1").Status != "unavailable" {
				t.Fatal("ambiguous source settings accepted")
			}
		}
	}
}

func TestNativePageSettingsExplicitAndFailClosed(t *testing.T) {
	raw := `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><pageSetup paperSize="9" orientation="portrait" scale="100"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`
	got := previewNativePageSettings([]byte(raw), "xl/worksheets/sheet1.xml", "1")
	if got.Status != "available" || got.Settings.Paper != "A4" || got.Settings.Scale != 100 {
		t.Fatalf("unexpected %+v", got)
	}
	for _, edit := range [][2]string{{` scale="100"`, ""}, {`scale="100"`, `scale="0x64"`}, {`left="0.7"`, `left="NaN"`}, {`left="0.7"`, `left="0x1p2"`}, {`paperSize="9"`, `paperSize="999"`}, {`scale="100"`, `scale="100" horizontalDpi="300"`}, {`</worksheet>`, `<rowBreaks/></worksheet>`}, {`</worksheet>`, `<headerFooter/></worksheet>`}, {`</worksheet>`, `<pageSetup paperSize="9" orientation="portrait" scale="100"/></worksheet>`}, {`<pageSetup`, `<pageSetup xmlns="urn:foreign"`}} {
		if previewNativePageSettings([]byte(strings.Replace(raw, edit[0], edit[1], 1)), "sheet.xml", "1").Status != "unavailable" {
			t.Fatalf("accepted %v", edit)
		}
	}
}

func TestNativePageSettingsExplicitFit(t *testing.T) {
	for _, ns := range []string{spreadsheetMLTransitional, spreadsheetMLStrict} {
		raw := `<worksheet xmlns="` + ns + `"><sheetPr><pageSetUpPr fitToPage="true"/></sheetPr><pageSetup paperSize="9" orientation="portrait" fitToWidth="1" fitToHeight="0"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`
		for _, tc := range []struct {
			from, to             string
			width, height, scale int
		}{
			{"", "", 1, 0, 100},
			{`fitToPage="true"`, `fitToPage="1"`, 1, 0, 100},
			{`fitToWidth="1" fitToHeight="0"`, `fitToWidth="0" fitToHeight="2" scale="75"`, 0, 2, 75},
			{`fitToWidth="1" fitToHeight="0"`, `fitToWidth="100" fitToHeight="100" pageOrder="overThenDown"`, 100, 100, 100},
		} {
			input := raw
			if tc.from != "" {
				input = strings.Replace(raw, tc.from, tc.to, 1)
			}
			got := previewNativePageSettings([]byte(input), "routed/sheet.xml", "17")
			if got.Status != "available" || got.SheetPart != "routed/sheet.xml" || got.SheetID != "17" || got.Settings.FitToPage == nil || got.Settings.FitToPage.Width != tc.width || got.Settings.FitToPage.Height != tc.height || got.Settings.Scale != tc.scale || len(got.Warnings) != 2 {
				t.Fatalf("fit settings lost: %+v", got)
			}
		}
		for _, edit := range [][2]string{
			{`<sheetPr><pageSetUpPr fitToPage="true"/></sheetPr>`, ``},
			{`fitToPage="true"`, `fitToPage="false"`}, {`fitToPage="true"`, `fitToPage="TRUE"`},
			{`fitToPage="true"`, `fitToPage="true" autoPageBreaks="true"`},
			{`fitToPage="true"`, `xmlns:f="urn:foreign" f:fitToPage="true"`},
			{`fitToPage="true"`, `fitToPage="true" fitToPage="true"`},
			{`<pageSetUpPr`, `<pageSetUpPr xmlns="urn:foreign"`},
			{`<sheetPr>`, `<sheetPr xmlns="urn:foreign">`},
			{`<sheetPr>`, `<sheetPr published="true">`},
			{`</sheetPr>`, `<pageSetUpPr fitToPage="true"/></sheetPr>`},
			{`</sheetPr>`, `</sheetPr><sheetPr/>`},
			{`</sheetPr>`, `</sheetPr><pageSetUpPr fitToPage="true"/>`},
			{`</sheetPr>`, `</sheetPr><other><pageSetUpPr fitToPage="true"/></other>`},
			{`<sheetPr><pageSetUpPr fitToPage="true"/></sheetPr>`, `<sheetPr><other><pageSetUpPr fitToPage="true"/></other></sheetPr>`},
			{`fitToWidth="1"`, `fitToWidth="0"`}, {`fitToWidth="1"`, `fitToWidth="101"`},
			{`fitToWidth="1"`, `fitToWidth="01"`}, {`fitToWidth="1"`, `fitToWidth="1.0"`},
			{`fitToHeight="0"`, `fitToHeight="-1"`}, {`fitToHeight="0"`, `fitToHeight="+1"`},
			{`fitToHeight="0"`, `fitToHeight="0x1"`}, {`fitToHeight="0"`, ``},
			{`fitToHeight="0"`, `fitToHeight="0" scale=""`},
			{`fitToHeight="0"`, `fitToHeight="0" scale="9"`},
			{`fitToHeight="0"`, `fitToHeight="0" scale="NaN"`},
			{`fitToHeight="0"`, `fitToHeight="0" horizontalDpi="300"`},
			{`</worksheet>`, `<rowBreaks/></worksheet>`},
			{`</worksheet>`, `<printOptions/></worksheet>`},
			{`</worksheet>`, `<headerFooter/></worksheet>`},
		} {
			got := previewNativePageSettings([]byte(strings.Replace(raw, edit[0], edit[1], 1)), "sheet.xml", "1")
			if got.Status != "unavailable" || got.Settings != nil {
				t.Fatalf("accepted unsupported fit %v: %+v", edit, got)
			}
		}
	}
}
