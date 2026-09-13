package xlsxpatch

import (
	"bytes"
	"reflect"
	"strings"
	"testing"
)

func TestNativePrintOffset(t *testing.T) {
	for _, tc := range []struct {
		formula string
		want    NativePrintAreaRectV1
	}{
		{`OFFSET('Data Set'!$D$3,3,-2,1,1)`, NativePrintAreaRectV1{5, 1, 5, 1}},
		{`OFFSET('Data Set'!$D$3:$F$5,3,-2)`, NativePrintAreaRectV1{5, 1, 7, 3}},
		{`OFFSET('Data Set'!$D$3:$F$5, 3, -2, 2)`, NativePrintAreaRectV1{5, 1, 6, 3}},
		{`OFFSET('Data Set'!$B$2,-1,-1,1048576,16384)`, NativePrintAreaRectV1{0, 0, 1048575, 16383}},
	} {
		if got := parseNativePrintOffset(tc.formula, "Data Set"); got == nil || *got != tc.want {
			t.Errorf("%s: %+v", tc.formula, got)
		}
	}
	if got := parseNativePrintOffset(`OFFSET('O''Brien,(!'!$A$1,0,0)`, "O'Brien,(!"); got == nil || got.Row != 0 {
		t.Fatal(got)
	}
	for _, formula := range []string{
		`OFFSET('Data Set'!$A$1,-1,0)`, `OFFSET('Data Set'!$A$1,0,-1)`,
		`OFFSET('Data Set'!$XFD$1048576,1,0)`, `OFFSET('Data Set'!$XFD$1048576,0,1)`,
		`OFFSET('Data Set'!$A$1,0,0,0,1)`, `OFFSET('Data Set'!$A$1,0,0,-1,1)`,
		`OFFSET('Data Set'!$A$1,0,0,1,0)`, `OFFSET('Data Set'!$A$1,0,0,1,16385)`,
		`OFFSET('Data Set'!$A$1,999999999999999999999,0)`,
		`OFFSET('Data Set'!$A$1,0.5,0)`, `OFFSET('Data Set'!$A$1,1e2,0)`,
		`OFFSET('Data Set'!$A$1,0,0,COUNTA(A:A),1)`, `OFFSET('Data Set'!$A$1,A1,0)`,
		`OFFSET(Name,0,0)`, `OFFSET(Other!$A$1,0,0)`, `OFFSET('[book]Data Set'!$A$1,0,0)`,
		`OFFSET('Data Set'!A1,0,0)`, `OFFSET('Data Set'!$A$1,0)`,
		`OFFSET('Data Set'!$A$1,0,0,,2)`, `OFFSET('Data Set'!$A$1,0,0,1,1,1)`,
		`OFFSET('Data Set'!$A$1,0,0),'Data Set'!$C$3`,
		`OFFSET(OFFSET('Data Set'!$A$1,0,0),0,0)`,
		`=OFFSET('Data Set'!$A$1,0,0)`, `offset('Data Set'!$A$1,0,0)`,
		"OFFSET('Data Set'!$A$1,\n0,0)", strings.Repeat("x", 2049),
	} {
		if got := parseNativePrintOffset(formula, "Data Set"); got != nil {
			t.Errorf("accepted %q: %+v", formula, got)
		}
	}
}

func TestNativePrintOffsetRoutedOwnership(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, invalid := range []string{"", "duplicate", "global", "foreign", "titles", "dynamic"} {
			parts := nativeWorkbookFixture(strict)
			parts["Charts/chart1.xml"] = previewChartFixture()
			formula := `OFFSET('Data Set'!$D$3:$F$5,3,-2)`
			name := `<definedName name="_xlnm.Print_Area" localSheetId="0">` + formula + `</definedName>`
			switch invalid {
			case "duplicate":
				name += name
			case "global":
				name = strings.Replace(name, ` localSheetId="0"`, "", 1)
			case "foreign":
				name = strings.Replace(name, `<definedName `, `<definedName xmlns="urn:foreign" `, 1)
			case "titles":
				name += `<definedName name="_xlnm.Print_Titles" localSheetId="0">BAD</definedName>`
			case "dynamic":
				name = strings.Replace(name, ",3,-2", ",COUNTA(A:A),-2", 1)
			}
			parts["Book/Workbook.xml"] = strings.Replace(parts["Book/Workbook.xml"], `</sheets>`, `</sheets><definedNames>`+name+`</definedNames>`, 1)
			source := buildZip(t, parts)
			before := bytes.Clone(source)
			nativeBefore, err := ExtractNativeWorkbookV2(source)
			if err != nil {
				t.Fatal(err)
			}
			got, err := InspectNativeWorkbookObjectsV1(source)
			if err != nil {
				t.Fatal(err)
			}
			entry := got.PrintAreaSets[0]
			if invalid == "" {
				if entry.Status != "available" || len(entry.Areas) != 1 || entry.Areas[0] != (NativePrintAreaRectV1{5, 1, 7, 3}) || entry.SheetID != "7" || entry.SheetPart != "Sheets/s1.xml" || !strings.Contains(strings.Join(entry.Warnings, " "), formula) {
					t.Fatalf("%+v", entry)
				}
			} else if entry.Status != "unavailable" || len(entry.Areas) != 0 {
				t.Fatalf("%s: %+v", invalid, entry)
			}
			if got.PrintAreas[0].Status != "unavailable" {
				t.Fatal("legacy literal selector changed")
			}
			nativeAfter, err := ExtractNativeWorkbookV2(source)
			if err != nil || !reflect.DeepEqual(nativeBefore, nativeAfter) || !bytes.Equal(source, before) {
				t.Fatal("source or mutation authority changed")
			}
		}
	}
}
