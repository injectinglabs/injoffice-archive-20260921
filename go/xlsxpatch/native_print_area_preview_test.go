package xlsxpatch

import (
	"bytes"
	"encoding/json"
	"strings"
	"testing"
)

func TestNativePrintAreaRect(t *testing.T) {
	for _, tc := range []struct {
		name, ref string
		valid     bool
	}{
		{"Sheet1", `Sheet1!$B$2:$D$8`, true},
		{"Data Set", `'Data Set'!$A$1`, true},
		{"O'Brien!", `'O''Brien!'!$XFD$1048576`, true},
		{"Sheet1", `Sheet1!$A$1:$XFD$1048576`, true},
		{"Sheet1", `Sheet1!A1:B2`, false},
		{"Sheet1", `Sheet1!$A1:$B$2`, false},
		{"Sheet1", `Other!$A$1`, false},
		{"Sheet1", `'[file.xlsx]Sheet1'!$A$1`, false},
		{"Sheet1", `'Sheet1:Sheet2'!$A$1`, false},
		{"Sheet1", `Sheet1!$A$1,Sheet1!$B$2`, false},
		{"Sheet1", `=Sheet1!$A$1`, false},
		{"Sheet1", `OFFSET(Sheet1!$A$1,0,0)`, false},
		{"Sheet1", `Sheet1!NamedRange`, false},
		{"Sheet1", `Sheet1!$A:$D`, false},
		{"Sheet1", `Sheet1!$1:$8`, false},
		{"Sheet1", `Sheet1!$B$2:$A$1`, false},
		{"Sheet1", `Sheet1!$A$0`, false},
		{"Sheet1", `Sheet1!$A$01`, false},
		{"Sheet1", `Sheet1!$XFE$1`, false},
		{"Sheet1", `Sheet1!$A$1048577`, false},
		{"Sheet1", `'Sheet1!$A$1`, false},
		{"Sheet1", `'Sheet1'x!$A$1`, false},
		{"Sheet1", ` Sheet1!$A$1`, false},
		{"Sheet1", strings.Repeat("x", 4097), false},
	} {
		if got := parseNativePrintAreaRect(tc.ref, tc.name); (got != nil) != tc.valid {
			t.Errorf("%q => %+v", tc.ref, got)
		}
	}
}

func TestNativePrintAreasNamespacesScopeAndAmbiguity(t *testing.T) {
	sheets := []NativeWorkbookSheetV2{{ID: "7", Name: "Data Set", Order: 0, PartName: "s1.xml"}, {ID: "99", Name: "Second", Order: 1, PartName: "s2.xml"}}
	def := `<definedName name="_xlnm.Print_Area" localSheetId="0">'Data Set'!$B$2:$D$8</definedName>`
	second := `<definedName name="_xlnm.Print_Area" localSheetId="1">Second!$A$1</definedName>`
	for _, ns := range []string{spreadsheetMLTransitional, spreadsheetMLStrict} {
		wrap := func(s string) string { return `<workbook xmlns="` + ns + `">` + s + `</workbook>` }
		for _, tc := range []struct {
			label, content string
			first, second  bool
		}{
			{"valid", `<definedNames>` + def + second + `</definedNames>`, true, true},
			{"none", ``, false, false},
			{"owner-attr", `<definedNames unexpected="1">` + def + second + `</definedNames>`, false, false},
			{"owner-unknown-child", `<definedNames>` + def + second + `<other/></definedNames>`, false, false},
			{"owner-foreign-child", `<definedNames>` + def + second + `<definedName xmlns="urn:foreign" name="Other"/></definedNames>`, false, false},
			{"foreign-name-attr", `<definedNames>` + def + `<definedName xmlns:f="urn:foreign" f:name="_xlnm.Print_Titles" localSheetId="0">x</definedName></definedNames>`, false, false},
			{"duplicate", `<definedNames>` + def + def + second + `</definedNames>`, false, true},
			{"case-ambiguous", `<definedNames>` + def + strings.Replace(def, "_xlnm.Print_Area", "_xlnm.print_area", 1) + second + `</definedNames>`, false, false},
			{"titles", `<definedNames>` + def + second + `<definedName name="_xlnm.Print_Titles" localSheetId="0">'Data Set'!$1:$2</definedName></definedNames>`, true, true},
			{"titles-invalid", `<definedNames>` + def + second + `<definedName name="_xlnm.Print_Titles" localSheetId="0">'Data Set'!$A$1:$B$2</definedName></definedNames>`, false, true},
			{"global", `<definedNames>` + def + strings.Replace(second, ` localSheetId="1"`, ``, 1) + `</definedNames>`, false, false},
			{"sheet-id-not-ordinal", `<definedNames>` + strings.Replace(def, `localSheetId="0"`, `localSheetId="7"`, 1) + second + `</definedNames>`, false, false},
			{"malformed-scope", `<definedNames>` + strings.Replace(def, `localSheetId="0"`, `localSheetId="00"`, 1) + second + `</definedNames>`, false, false},
			{"foreign", `<definedNames>` + def + strings.Replace(second, `<definedName`, `<definedName xmlns="urn:foreign"`, 1) + `</definedNames>`, false, false},
			{"nested", `<definedNames><other>` + def + `</other>` + second + `</definedNames>`, false, false},
			{"nested-owner", `<other><definedNames>` + def + `</definedNames></other>`, false, false},
			{"duplicate-owner", `<definedNames>` + def + `</definedNames><definedNames>` + second + `</definedNames>`, false, false},
			{"foreign-owner", `<definedNames>` + def + `</definedNames><definedNames xmlns="urn:foreign">` + second + `</definedNames>`, false, false},
			{"foreign-attr", `<definedNames>` + strings.Replace(def, `localSheetId="0"`, `xmlns:f="urn:foreign" f:localSheetId="0"`, 1) + second + `</definedNames>`, false, false},
			{"extra-attr", `<definedNames>` + strings.Replace(def, `localSheetId="0"`, `localSheetId="0" hidden="1"`, 1) + second + `</definedNames>`, false, true},
			{"child", `<definedNames>` + strings.Replace(def, `</definedName>`, `<other/></definedName>`, 1) + second + `</definedNames>`, false, true},
			{"duplicate-attr", `<definedNames>` + strings.Replace(def, `localSheetId="0"`, `localSheetId="0" localSheetId="1"`, 1) + second + `</definedNames>`, false, false},
		} {
			t.Run(tc.label+ns, func(t *testing.T) {
				got := previewNativePrintAreas([]byte(wrap(tc.content)), sheets)
				sets := previewNativePrintAreaSets([]byte(wrap(tc.content)), sheets)
				for i, available := range []bool{tc.first, tc.second} {
					if (sets[i].Status == "available") != available || (len(sets[i].Areas) == 1) != available || sets[i].SheetID != sheets[i].ID {
						t.Fatalf("unexpected set %+v", sets)
					}
					if (got[i].Status == "available") != available || (got[i].Area != nil) != available || len(got[i].Warnings) != 1 || got[i].SheetID != sheets[i].ID {
						t.Fatalf("unexpected %+v", got)
					}
				}
			})
		}
		got := previewNativePrintAreas([]byte(wrap(`<definedNames>`+def+`</definedNames>`)), sheets)
		if *got[0].Area != (NativePrintAreaRectV1{Row: 1, Column: 1, EndRow: 7, EndColumn: 3}) {
			t.Fatal(got)
		}
		reordered := []NativeWorkbookSheetV2{sheets[1], sheets[0]}
		got = previewNativePrintAreas([]byte(wrap(`<definedNames>`+def+`</definedNames>`)), reordered)
		if got[0].Status != "unavailable" || got[1].Status != "available" || got[1].SheetID != "7" {
			t.Fatal("did not use workbook ordinal", got)
		}
		sets := previewNativePrintAreaSets([]byte(wrap(`<definedNames>`+def+`</definedNames>`)), reordered)
		if sets[0].Status != "unavailable" || sets[1].Status != "available" || sets[1].SheetID != "7" {
			t.Fatal("set did not use workbook ordinal", sets)
		}
	}
	if got := previewNativePrintAreas(bytes.Repeat([]byte(" "), 2*1024*1024+1), sheets); got[0].Status != "unavailable" {
		t.Fatal(got)
	}
	if got := previewNativePrintAreaSets(bytes.Repeat([]byte(" "), 2*1024*1024+1), sheets); got[0].Status != "unavailable" {
		t.Fatal(got)
	}
	if got := previewNativePrintAreas(nil, make([]NativeWorkbookSheetV2, 65)); len(got) != 64 {
		t.Fatal(len(got))
	}
	if got := previewNativePrintAreaSets(nil, make([]NativeWorkbookSheetV2, 65)); len(got) != 64 {
		t.Fatal(len(got))
	}
}

func TestNativePrintAreaRoutedPackageReadOnly(t *testing.T) {
	for _, strict := range []bool{false, true} {
		parts := nativeWorkbookFixture(strict)
		parts["Charts/chart1.xml"] = previewChartFixture()
		parts["Book/Workbook.xml"] = strings.Replace(parts["Book/Workbook.xml"], `</sheets>`, `</sheets><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">'Data Set'!$B$2:$D$8</definedName></definedNames>`, 1)
		source := buildZip(t, parts)
		before := bytes.Clone(source)
		nativeBefore, err := ExtractNativeWorkbookV2(source)
		if err != nil {
			t.Fatal(err)
		}
		nativeJSONBefore, err := EncodeNativeWorkbookV2(nativeBefore)
		if err != nil {
			t.Fatal(err)
		}
		got, err := InspectNativeWorkbookObjectsV1(source)
		if err != nil {
			t.Fatal(err)
		}
		if len(got.PrintAreas) != 2 || got.PrintAreas[0].Status != "available" || got.PrintAreas[0].SheetID != "7" || got.PrintAreas[0].SheetPart != "Sheets/s1.xml" || got.PrintAreas[1].Status != "unavailable" {
			t.Fatalf("%+v", got.PrintAreas)
		}
		if !bytes.Equal(before, source) {
			t.Fatal("source bytes changed")
		}
		nativeAfter, err := ExtractNativeWorkbookV2(source)
		if err != nil {
			t.Fatal(err)
		}
		nativeJSONAfter, err := EncodeNativeWorkbookV2(nativeAfter)
		if err != nil || !bytes.Equal(nativeJSONBefore, nativeJSONAfter) || bytes.Contains(nativeJSONAfter, []byte(`"print_areas"`)) {
			t.Fatal("native authority changed")
		}
		encoded, err := json.Marshal(got.PrintAreas[1])
		if err != nil || bytes.Contains(encoded, []byte(`"area"`)) {
			t.Fatalf("unavailable area serialized: %s", encoded)
		}
	}
}
