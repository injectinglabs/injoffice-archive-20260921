package xlsxpatch

import (
	"strings"
	"testing"
)

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
