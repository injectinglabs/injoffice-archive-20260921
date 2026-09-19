package xlsxpatch

import (
	"bytes"
	"encoding/json"
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
			for _, replacement := range []string{`pageOrder=""`, `pageOrder="diagonal"`, `xmlns:f="urn:foreign" f:pageOrder="overThenDown"`, `pageOrder="overThenDown" pageOrder="downThenOver"`, `pageOrder="overThenDown" paperWidth="8.5in"`} {
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
	if previewNativePageSettings([]byte(strings.Replace(raw, ` scale="100"`, "", 1)), "sheet.xml", "1").Status != "available" {
		t.Fatal("omitted scale must default to 100")
	}
	sentinel := strings.Replace(raw, `scale="100"`, `horizontalDpi="4294967293" verticalDpi="0" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"`, 1)
	if got := previewNativePageSettings([]byte(sentinel), "sheet.xml", "1"); got.Status != "available" || got.Settings.Scale != 100 {
		t.Fatalf("sentinel printer DPI / r:id must be ignored: %+v", got)
	}
	for _, edit := range [][2]string{{`scale="100"`, `scale="0x64"`}, {`left="0.7"`, `left="NaN"`}, {`left="0.7"`, `left="0x1p2"`}, {`paperSize="9"`, `paperSize="999"`}, {`scale="100"`, `scale="100" paperWidth="8.5in"`}, {`scale="100"`, `scale="100" cellComments="atEnd"`}, {`scale="100"`, `scale="100" errors="blank"`}, {`scale="100"`, `scale="100" useFirstPageNumber="true" firstPageNumber="4"`}, {`</worksheet>`, `<rowBreaks/></worksheet>`}, {`</worksheet>`, `<colBreaks/></worksheet>`}, {`</worksheet>`, `<printOptions headings="true"/></worksheet>`}, {`</worksheet>`, `<printOptions horizontalCentered="true"/></worksheet>`}, {`</worksheet>`, `<printOptions verticalCentered="true"/></worksheet>`}, {`</worksheet>`, `<printOptions gridLines="yes"/></worksheet>`}, {`</worksheet>`, `<printOptions><headings/></printOptions></worksheet>`}, {`</worksheet>`, `<headerFooter/><headerFooter/></worksheet>`}, {`</worksheet>`, `<pageSetup paperSize="9" orientation="portrait" scale="100"/></worksheet>`}, {`<pageSetup`, `<pageSetup xmlns="urn:foreign"`}} {
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
			{`fitToPage="true"`, `fitToPage="TRUE"`},
			{`fitToPage="true"`, `fitToPage="true" autoPageBreaks="true"`},
			{`fitToPage="true"`, `xmlns:f="urn:foreign" f:fitToPage="true"`},
			{`fitToPage="true"`, `fitToPage="true" fitToPage="true"`},
			{`<pageSetUpPr`, `<pageSetUpPr xmlns="urn:foreign"`},
			{`<sheetPr>`, `<sheetPr xmlns="urn:foreign">`},
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
			{`fitToHeight="0"`, `fitToHeight="0" paperWidth="8.5in"`},
			{`</worksheet>`, `<rowBreaks/></worksheet>`},
			{`</worksheet>`, `<colBreaks/></worksheet>`},
			{`</worksheet>`, `<printOptions headings="true"/></worksheet>`},
		} {
			got := previewNativePageSettings([]byte(strings.Replace(raw, edit[0], edit[1], 1)), "sheet.xml", "1")
			if got.Status != "unavailable" || got.Settings != nil {
				t.Fatalf("accepted unsupported fit %v: %+v", edit, got)
			}
		}
	}
}

// A worksheet that authors margins and declares no pageSetup has no authored
// paper to contradict, which a host may default. A pageSetup that exists but is
// ambiguous or unsupported is a different fact and must stay unavailable.
func TestNativePageSettingsMarginsOnlyWithoutPageSetup(t *testing.T) {
	margins := `<pageMargins left="0.7" right="0.75" top="0.8" bottom="0.85" header="0.3" footer="0.3"/>`
	for _, ns := range []string{spreadsheetMLTransitional, spreadsheetMLStrict} {
		raw := `<worksheet xmlns="` + ns + `">` + margins + `</worksheet>`
		got := previewNativePageSettings([]byte(raw), "xl/worksheets/sheet1.xml", "1")
		if got.Status != "margins-only" || got.Settings != nil || got.Margins == nil {
			t.Fatalf("unexpected %+v", got)
		}
		if got.Margins.Left != 0.7 || got.Margins.Right != 0.75 || got.Margins.Top != 0.8 || got.Margins.Bottom != 0.85 {
			t.Fatalf("authored margins not reported exactly: %+v", got.Margins)
		}
		if len(got.Warnings) != 1 || !strings.Contains(got.Warnings[0], "no pageSetup element") {
			t.Fatalf("missing disclosure: %+v", got.Warnings)
		}
	}
	base := `<worksheet xmlns="` + spreadsheetMLTransitional + `">` + margins
	for _, tc := range []struct{ name, body string }{
		{"duplicate pageSetup reads as ambiguous, not absent", `<pageSetup paperSize="9" orientation="portrait"/><pageSetup paperSize="1" orientation="landscape"/>`},
		{"fit-to-page activation without its dimensions", `<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>`},
		{"row breaks still refuse", `<rowBreaks/>`},
		{"column breaks still refuse", `<colBreaks/>`},
		{"printed headings still refuse", `<printOptions headings="true"/>`},
		{"foreign margins", ``},
	} {
		body := base + tc.body + `</worksheet>`
		if tc.name == "foreign margins" {
			body = `<worksheet xmlns="` + spreadsheetMLTransitional + `">` + strings.Replace(margins, "<pageMargins", `<pageMargins xmlns="urn:foreign"`, 1) + `</worksheet>`
		}
		if got := previewNativePageSettings([]byte(body), "sheet.xml", "1"); got.Status != "unavailable" {
			t.Fatalf("%s: got %q", tc.name, got.Status)
		}
	}
	// Margins must still be exact and bounded to be reported.
	for _, bad := range []string{`left="NaN"`, `left="-1"`, `left="21"`} {
		body := `<worksheet xmlns="` + spreadsheetMLTransitional + `">` + strings.Replace(margins, `left="0.7"`, bad, 1) + `</worksheet>`
		if got := previewNativePageSettings([]byte(body), "sheet.xml", "1"); got.Status != "unavailable" {
			t.Fatalf("accepted %s: %q", bad, got.Status)
		}
	}
	// No margins at all remains unavailable: nothing authored to build a page from.
	if got := previewNativePageSettings([]byte(`<worksheet xmlns="`+spreadsheetMLTransitional+`"/>`), "sheet.xml", "1"); got.Status != "unavailable" {
		t.Fatalf("bare worksheet: %q", got.Status)
	}
}

// ECMA-376 §18.3.1.63 gives paperSize, orientation and scale attribute
// defaults, so a pageSetup that omits them still states a page. The shapes here
// are the ones the hard-v2 corpus actually ships and that this tier used to
// refuse outright: an Excel pageSetup carrying only orientation, a LibreOffice
// worksheet whose printOptions, headerFooter, sheetPr and printer attributes
// surround a fully authored pageSetup, and margins with headerFooter and no
// pageSetup at all.
func TestNativePageSettingsECMADefaultsAndNonGeometryPrintFacts(t *testing.T) {
	margins := `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>`
	rels := `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`
	sheet := func(body string) []byte {
		return []byte(`<worksheet xmlns="` + spreadsheetMLTransitional + `" ` + rels + `>` + body + `</worksheet>`)
	}
	for _, tc := range []struct {
		name        string
		body        string
		paper       string
		orientation string
		scale       int
		defaults    []string
		disclose    []string
		header      string
		footer      string
	}{
		{
			name:  "authored orientation only defaults Letter and 100%",
			body:  margins + `<pageSetup orientation="portrait" r:id="rId1"/>`,
			paper: "Letter", orientation: "portrait", scale: 100,
			defaults: []string{"paper", "scale"},
			disclose: []string{"ECMA-376 CT_PageSetup attribute default"},
		},
		{
			name:  "an empty pageSetup is Letter portrait at 100%",
			body:  margins + `<pageSetup/>`,
			paper: "Letter", orientation: "portrait", scale: 100,
			defaults: []string{"paper", "orientation", "scale"},
			disclose: []string{"ECMA-376 CT_PageSetup attribute default"},
		},
		{
			name:  `orientation="default" is the printer default, which is portrait`,
			body:  margins + `<pageSetup paperSize="9" orientation="default" scale="75"/>`,
			paper: "A4", orientation: "portrait", scale: 75,
			defaults: []string{"orientation"},
			disclose: []string{"ECMA-376 CT_PageSetup attribute default"},
		},
		{
			name:  "authored paper with header and footer content still paginates",
			body:  margins + `<pageSetup paperSize="9" orientation="portrait" r:id="rId1"/><headerFooter><oddHeader>&amp;C&amp;A</oddHeader><oddFooter>&amp;CPage &amp;P</oddFooter></headerFooter>`,
			paper: "A4", orientation: "portrait", scale: 100,
			defaults: []string{"scale"},
			disclose: []string{"odd-page header and footer that every printed page carries"},
			header:   "&C&A", footer: "&CPage &P",
		},
		{
			name: "a fully authored LibreOffice page survives its print options and printer attributes",
			body: `<sheetPr filterMode="false"><pageSetUpPr fitToPage="false"/></sheetPr>` +
				`<printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"/>` +
				margins +
				`<pageSetup paperSize="1" scale="90" firstPageNumber="1" fitToWidth="1" fitToHeight="1" pageOrder="downThenOver" orientation="portrait" usePrinterDefaults="false" blackAndWhite="false" draft="false" cellComments="none" useFirstPageNumber="true" horizontalDpi="300" verticalDpi="300" copies="1"/>` +
				`<headerFooter differentFirst="false" differentOddEven="false"><oddHeader>&amp;C&amp;A</oddHeader></headerFooter>`,
			paper: "Letter", orientation: "portrait", scale: 90,
			defaults: nil,
			disclose: []string{"odd-page header and footer that every printed page carries", "Printed gridlines are not painted", "Printer-directed page setup attributes are not resolved"},
			header:   "&C&A",
		},
		{
			name:  "stored fit dimensions without the fit switch stay a percentage page",
			body:  margins + `<pageSetup paperSize="1" orientation="landscape" scale="100" fitToWidth="1" fitToHeight="1"/>`,
			paper: "Letter", orientation: "landscape", scale: 100,
			defaults: nil,
		},
	} {
		got := previewNativePageSettings(sheet(tc.body), "xl/worksheets/sheet1.xml", "1")
		if got.Status != "available" || got.Settings == nil {
			t.Fatalf("%s: %+v", tc.name, got)
		}
		if got.Settings.Paper != tc.paper || got.Settings.Orientation != tc.orientation || got.Settings.Scale != tc.scale || got.Settings.FitToPage != nil {
			t.Fatalf("%s: %+v", tc.name, got.Settings)
		}
		if got.Settings.Left != 0.7 || got.Settings.Top != 0.75 {
			t.Fatalf("%s: authored margins changed: %+v", tc.name, got.Settings)
		}
		if strings.Join(got.Defaults, ",") != strings.Join(tc.defaults, ",") {
			t.Fatalf("%s: defaulted facts %v, want %v", tc.name, got.Defaults, tc.defaults)
		}
		if tc.header == "" && tc.footer == "" {
			if got.HeaderFooter != nil {
				t.Fatalf("%s: unauthored header/footer reported: %+v", tc.name, got.HeaderFooter)
			}
		} else if got.HeaderFooter == nil || got.HeaderFooter.OddHeader != tc.header || got.HeaderFooter.OddFooter != tc.footer {
			t.Fatalf("%s: header/footer %+v, want %q / %q", tc.name, got.HeaderFooter, tc.header, tc.footer)
		}
		joined := strings.Join(got.Warnings, "\n")
		for _, want := range tc.disclose {
			if !strings.Contains(joined, want) {
				t.Fatalf("%s: undisclosed %q in %v", tc.name, want, got.Warnings)
			}
		}
		if len(tc.defaults) == 0 && strings.Contains(joined, "ECMA-376 CT_PageSetup attribute default") {
			t.Fatalf("%s: disclosed a default it did not apply", tc.name)
		}
		if len(got.Warnings) < 1 || len(got.Warnings) > 8 {
			t.Fatalf("%s: %d warnings exceed the decoded bound", tc.name, len(got.Warnings))
		}
	}
	// No pageSetup at all keeps reporting authored margins for the host to
	// choose paper against; surrounding print markup no longer hides that.
	noSetup := previewNativePageSettings(sheet(margins+`<headerFooter alignWithMargins="0"/>`), "xl/worksheets/sheet1.xml", "1")
	if noSetup.Status != "margins-only" || noSetup.Margins == nil || noSetup.Defaults != nil {
		t.Fatalf("margins with header/footer and no pageSetup: %+v", noSetup)
	}
	if !strings.Contains(strings.Join(noSetup.Warnings, "\n"), "Header and footer text is not painted") {
		t.Fatalf("undisclosed header/footer: %v", noSetup.Warnings)
	}
	// One odd pair that every printed page carries is the only shape whose text
	// is reported. Everything else keeps saying the header is not painted rather
	// than painting a pair some pages do not print, or one placed by a rule this
	// tier does not apply.
	for _, tc := range []struct{ name, header string }{
		{"a first page of its own", `<headerFooter differentFirst="true"><oddHeader>&amp;C&amp;A</oddHeader><firstHeader>&amp;CFirst</firstHeader></headerFooter>`},
		{"different odd and even pages", `<headerFooter differentOddEven="true"><oddHeader>&amp;C&amp;A</oddHeader><evenHeader>&amp;CEven</evenHeader></headerFooter>`},
		{"an even page this tier never sees", `<headerFooter><oddHeader>&amp;C&amp;A</oddHeader><evenHeader>&amp;CEven</evenHeader></headerFooter>`},
		{"a header that does not scale with the document", `<headerFooter scaleWithDoc="false"><oddHeader>&amp;C&amp;A</oddHeader></headerFooter>`},
		{"a header aligned to the printer, not the margins", `<headerFooter alignWithMargins="false"><oddHeader>&amp;C&amp;A</oddHeader></headerFooter>`},
		{"an unreadable flag", `<headerFooter differentFirst="yes"><oddHeader>&amp;C&amp;A</oddHeader></headerFooter>`},
		{"an unknown attribute", `<headerFooter zoom="2"><oddHeader>&amp;C&amp;A</oddHeader></headerFooter>`},
		{"a duplicated odd header", `<headerFooter><oddHeader>&amp;C&amp;A</oddHeader><oddHeader>&amp;C&amp;A</oddHeader></headerFooter>`},
		{"foreign markup inside the header", `<headerFooter><oddHeader xmlns="urn:foreign">&amp;C&amp;A</oddHeader></headerFooter>`},
		{"an empty pair", `<headerFooter differentFirst="false"/>`},
		{"unbounded header text", `<headerFooter><oddHeader>` + strings.Repeat("x", 1025) + `</oddHeader></headerFooter>`},
	} {
		got := previewNativePageSettings(sheet(margins+`<pageSetup paperSize="1" orientation="portrait" scale="100"/>`+tc.header), "xl/worksheets/sheet1.xml", "1")
		if got.Status != "available" || got.HeaderFooter != nil {
			t.Fatalf("%s: %+v %+v", tc.name, got.Status, got.HeaderFooter)
		}
		if !strings.Contains(strings.Join(got.Warnings, "\n"), "Header and footer text is not painted") {
			t.Fatalf("%s: undisclosed refusal: %v", tc.name, got.Warnings)
		}
	}
	// Facts with no defensible default still refuse rather than inventing one.
	for _, tc := range []struct{ name, body string }{
		{"manual row breaks", margins + `<pageSetup paperSize="1" orientation="portrait"/><rowBreaks count="1"><brk id="4" man="1"/></rowBreaks>`},
		{"manual column breaks", margins + `<pageSetup paperSize="1" orientation="portrait"/><colBreaks count="1"><brk id="2" man="1"/></colBreaks>`},
		{"printed row and column headings", margins + `<pageSetup paperSize="1" orientation="portrait"/><printOptions headings="true"/>`},
		{"centered printing", margins + `<pageSetup paperSize="1" orientation="portrait"/><printOptions horizontalCentered="true"/>`},
		{"unsupported paper", margins + `<pageSetup paperSize="5" orientation="portrait"/>`},
		{"custom paper dimensions", margins + `<pageSetup paperHeight="11in" paperWidth="8.5in" orientation="portrait"/>`},
		{"comments printed at the end add pages", margins + `<pageSetup paperSize="1" orientation="portrait" cellComments="atEnd"/>`},
		{"a stated first page number", margins + `<pageSetup paperSize="1" orientation="portrait" useFirstPageNumber="true" firstPageNumber="7"/>`},
		{"suppressed automatic page breaks", `<sheetPr><pageSetUpPr fitToPage="false" autoPageBreaks="false"/></sheetPr>` + margins + `<pageSetup paperSize="1" orientation="portrait"/>`},
		{"foreign sheetPr attributes", `<sheetPr xmlns:f="urn:foreign" f:filterMode="false"><pageSetUpPr fitToPage="false"/></sheetPr>` + margins + `<pageSetup paperSize="1" orientation="portrait"/>`},
	} {
		if got := previewNativePageSettings(sheet(tc.body), "xl/worksheets/sheet1.xml", "1"); got.Status != "unavailable" || got.Settings != nil {
			t.Fatalf("%s: defaulted a fact with no default: %+v", tc.name, got)
		}
	}
}

// ECMA-376 Part 1 §18.3.1.62 measures every pageMargins attribute from the
// paper edge, so header and footer are not additions to top and bottom: Excel
// prints the body between max(top, header) and max(bottom, footer). The
// attributes were parsed and then dropped, which silently gave every worksheet
// with a header margin deeper than its top margin a taller body than Excel
// prints. These are the authored margins of hard-v2 cell-anchored-hidden-shapes.xlsx.
func TestNativePageSettingsProjectHeaderAndFooterMargins(t *testing.T) {
	for _, ns := range []string{spreadsheetMLTransitional, spreadsheetMLStrict} {
		margins := `<pageMargins left="0.7" right="0.7" top="0.63" bottom="0" header="0.79" footer="0.19685"/>`
		got := previewNativePageSettings([]byte(`<worksheet xmlns="`+ns+`">`+margins+`<pageSetup paperSize="9" orientation="portrait" scale="80"/></worksheet>`), "xl/worksheets/sheet1.xml", "1")
		if got.Status != "available" || got.Settings == nil || got.Settings.Header != 0.79 || got.Settings.Footer != 0.19685 || got.Settings.Top != 0.63 || got.Settings.Bottom != 0 {
			t.Fatalf("authored header/footer margins lost: %+v", got.Settings)
		}
		// A worksheet with no pageSetup reports them too: its host must reserve
		// the same bands once it chooses paper.
		bare := previewNativePageSettings([]byte(`<worksheet xmlns="`+ns+`">`+margins+`</worksheet>`), "xl/worksheets/sheet1.xml", "1")
		if bare.Status != "margins-only" || bare.Margins == nil || bare.Margins.Header != 0.79 || bare.Margins.Footer != 0.19685 {
			t.Fatalf("authored header/footer margins lost without pageSetup: %+v", bare.Margins)
		}
		raw, err := json.Marshal(got)
		if err != nil {
			t.Fatal(err)
		}
		for _, key := range []string{`"header_inches":0.79`, `"footer_inches":0.19685`} {
			if !strings.Contains(string(raw), key) {
				t.Fatalf("page settings do not carry %s: %s", key, raw)
			}
		}
	}
}
