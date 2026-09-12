package xlsxpatch

import (
	"bytes"
	"encoding/json"
	"strings"
	"testing"
)

func TestNativePrintTitlesClosedOwnership(t *testing.T) {
	sheets := []NativeWorkbookSheetV2{{ID: "7", Name: "Data", Order: 0, PartName: "s1.xml"}, {ID: "99", Name: "Other", Order: 1, PartName: "s2.xml"}}
	def := `<definedName name="_xlnm.Print_Titles" localSheetId="0">Data!$1:$2</definedName>`
	for _, ns := range []string{spreadsheetMLTransitional, spreadsheetMLStrict} {
		wrap := func(s string) []byte { return []byte(`<workbook xmlns="` + ns + `">` + s + `</workbook>`) }
		for _, content := range []string{
			`<definedNames unexpected="1">` + def + `</definedNames>`,
			`<definedNames>` + def + `<unknown/></definedNames>`,
			`<definedNames>` + def + `</definedNames><definedNames/>`,
			`<wrapper><definedNames>` + def + `</definedNames></wrapper>`,
			`<definedNames xmlns="urn:foreign">` + def + `</definedNames>`,
			`<definedNames>` + strings.Replace(def, `localSheetId="0"`, `localSheetId="00"`, 1) + `</definedNames>`,
			`<definedNames>` + strings.Replace(def, `localSheetId="0"`, `localSheetId="7"`, 1) + `</definedNames>`,
			`<definedNames>` + strings.Replace(def, `localSheetId="0"`, `localSheetId="0" localSheetId="1"`, 1) + `</definedNames>`,
			`<definedNames>` + strings.Replace(def, `_xlnm.Print_Titles`, `_xlnm.print_titles`, 1) + `</definedNames>`,
			`<definedNames>` + strings.Replace(def, ` name=`, ` xmlns:f="urn:foreign" f:name=`, 1) + `</definedNames>`,
		} {
			got := previewNativePrintTitles(wrap(content), sheets)
			for _, entry := range got {
				encoded, _ := json.Marshal(entry)
				if entry.Status != "unavailable" || bytes.Contains(encoded, []byte(`"rows"`)) || bytes.Contains(encoded, []byte(`"columns"`)) {
					t.Fatalf("%s: %s", content, encoded)
				}
			}
		}
		got := previewNativePrintTitles(wrap(`<definedNames>`+def+`</definedNames>`), []NativeWorkbookSheetV2{sheets[1], sheets[0]})
		if got[0].Status != "unavailable" || got[1].Status != "available" || got[1].SheetID != "7" {
			t.Fatal("workbook ordinal routing", got)
		}
	}
	if got := previewNativePrintTitles(bytes.Repeat([]byte(" "), 2*1024*1024+1), sheets); got[0].Status != "unavailable" {
		t.Fatal(got)
	}
}

func TestNativePrintTitlesParser(t *testing.T) {
	for _, tc := range []struct {
		ref   string
		valid bool
	}{
		{`'O''Brien!, Data'!$1:$2`, true},
		{`'O''Brien!, Data'!$B:$D`, true},
		{`'O''Brien!, Data'!$1:$2,'O''Brien!, Data'!$A:$B`, true},
		{`'O''Brien!, Data'!$XFD:$XFD,'O''Brien!, Data'!$1048576:$1048576`, true},
		{`'O''Brien!, Data'!$2:$1`, false},
		{`'O''Brien!, Data'!$B:$A`, false},
		{`'O''Brien!, Data'!$A:$XFE`, false},
		{`'O''Brien!, Data'!$1:$1048577`, false},
		{`'O''Brien!, Data'!$0:$1`, false},
		{`'O''Brien!, Data'!$01:$2`, false},
		{`'O''Brien!, Data'!1:2`, false},
		{`'O''Brien!, Data'!$A$1:$B$2`, false},
		{`'O''Brien!, Data'!$1:$2,'O''Brien!, Data'!$3:$4`, false},
		{`'O''Brien!, Data'!$A:$B,'O''Brien!, Data'!$C:$D`, false},
		{`'O''Brien!, Data'!$1:$2,Other!$A:$B`, false},
		{`'O''Brien!, Data'!$1:$2,'O''Brien!, Data'!$A:$B,'O''Brien!, Data'!$C:$D`, false},
		{`'O''Brien!, Data'!$1:$2,`, false},
		{`='O''Brien!, Data'!$1:$2`, false},
		{`'[book.xlsx]O''Brien!, Data'!$1:$2`, false},
		{`'O''Brien!, Data:Other'!$1:$2`, false},
		{strings.Repeat("x", 4097), false},
	} {
		r, c := parseNativePrintTitles(tc.ref, "O'Brien!, Data")
		if (r != nil || c != nil) != tc.valid {
			t.Errorf("%q => %+v %+v", tc.ref, r, c)
		}
	}
	r, c := parseNativePrintTitles(`Sheet1!$B:$C,Sheet1!$2:$4`, "Sheet1")
	if *r != (NativePrintTitleSpanV1{1, 3}) || *c != (NativePrintTitleSpanV1{1, 2}) {
		t.Fatal(r, c)
	}
}

func TestNativePrintTitlesScopeAndReadOnly(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, tc := range []struct {
			label, title  string
			first, second bool
		}{
			{"valid", `<definedName name="_xlnm.Print_Titles" localSheetId="0">'Data Set'!$1:$2,'Data Set'!$A:$B</definedName>`, true, true},
			{"duplicate", `<definedName name="_xlnm.Print_Titles" localSheetId="0">'Data Set'!$1:$2</definedName><definedName name="_xlnm.Print_Titles" localSheetId="0">'Data Set'!$1:$2</definedName>`, false, true},
			{"invalid", `<definedName name="_xlnm.Print_Titles" localSheetId="0">'Data Set'!$A$1</definedName>`, false, true},
			{"global", `<definedName name="_xlnm.Print_Titles">'Data Set'!$1:$2</definedName>`, false, false},
			{"foreign-attr", `<definedName name="_xlnm.Print_Titles" xmlns:f="urn:foreign" f:localSheetId="0">'Data Set'!$1:$2</definedName>`, false, false},
			{"extra-attr", `<definedName name="_xlnm.Print_Titles" localSheetId="0" hidden="1">'Data Set'!$1:$2</definedName>`, false, true},
			{"nested", `<definedName name="_xlnm.Print_Titles" localSheetId="0">'Data Set'!$1:$2<other/></definedName>`, false, true},
		} {
			t.Run(tc.label+map[bool]string{true: "strict", false: "transitional"}[strict], func(t *testing.T) {
				parts := nativeWorkbookFixture(strict)
				parts["Charts/chart1.xml"] = previewChartFixture()
				// Fixture uses routed parts and sparse IDs 7/99. localSheetId is ordinal.
				second := `<definedName name="_xlnm.Print_Titles" localSheetId="1">Second!$A:$A</definedName>`
				// Derive the second sheet's actual name from extraction.
				base, err := ExtractNativeWorkbookV2(buildZip(t, parts))
				if err != nil {
					t.Fatal(err)
				}
				second = strings.Replace(second, "Second", base.Sheets[1].Name, 1)
				parts["Book/Workbook.xml"] = strings.Replace(parts["Book/Workbook.xml"], `</sheets>`, `</sheets><definedNames>`+tc.title+second+`<definedName name="_xlnm.Print_Area" localSheetId="0">'Data Set'!$B$2:$D$8</definedName></definedNames>`, 1)
				src := buildZip(t, parts)
				before := bytes.Clone(src)
				nativeBefore, err := ExtractNativeWorkbookV2(src)
				if err != nil {
					t.Fatal(err)
				}
				encodedBefore, _ := EncodeNativeWorkbookV2(nativeBefore)
				got, err := InspectNativeWorkbookObjectsV1(src)
				if err != nil {
					t.Fatal(err)
				}
				for i, want := range []bool{tc.first, tc.second} {
					title := got.PrintTitles[i]
					if (title.Status == "available") != want || (title.Rows != nil || title.Columns != nil) != want || len(title.Warnings) != 1 || title.SheetID != base.Sheets[i].ID || title.SheetPart != base.Sheets[i].PartName {
						t.Fatalf("%+v", got.PrintTitles)
					}
				}
				if (got.PrintAreas[0].Status == "available") != tc.first {
					t.Fatalf("area/title policy %+v", got.PrintAreas)
				}
				nativeAfter, err := ExtractNativeWorkbookV2(src)
				if err != nil {
					t.Fatal(err)
				}
				encodedAfter, _ := EncodeNativeWorkbookV2(nativeAfter)
				if !bytes.Equal(before, src) || !bytes.Equal(encodedBefore, encodedAfter) || bytes.Contains(encodedAfter, []byte("print_titles")) {
					t.Fatal("mutation authority changed")
				}
			})
		}
	}
	if got := previewNativePrintTitles(nil, make([]NativeWorkbookSheetV2, 65)); len(got) != 64 {
		t.Fatal(len(got))
	}
}
