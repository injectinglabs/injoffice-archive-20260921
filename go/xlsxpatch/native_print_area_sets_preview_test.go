package xlsxpatch

import (
	"bytes"
	"encoding/json"
	"fmt"
	"strings"
	"testing"
)

func TestNativePrintAreaSetsGrammar(t *testing.T) {
	for _, tc := range []struct {
		name, ref string
		count     int
	}{
		{"Sheet1", `Sheet1!$F$6:$G$8,Sheet1!$B$2:$D$4`, 2},
		{"O'Brien,!", `'O''Brien,!'!$B$2,'O''Brien,!'!$C$2`, 2},
		{"Sheet1", `Sheet1!$A$1:$B$2,Sheet1!$C$1:$D$2`, 2},
		{"Sheet1", `Sheet1!$A$1:$B$2,Sheet1!$A$3:$B$4`, 2},
		{"Sheet1", `Sheet1!$A$1`, 1},
		{"Sheet1", `Sheet1!$A$1,Sheet1!$A$1`, 0},
		{"Sheet1", `Sheet1!$A$1:$B$2,Sheet1!$B$2:$C$3`, 0},
		{"Sheet1", `Sheet1!$A$1:$D$4,Sheet1!$B$2:$C$3`, 0},
		{"Sheet1", `Sheet1!$A$1,Other!$B$2`, 0},
		{"Sheet1", `Sheet1!$A$1,Sheet1!$A:$B`, 0},
		{"Sheet1", `Sheet1!$A$1,OFFSET(Sheet1!$B$2,0,0)`, 0},
		{"Sheet1", `Sheet1!$A$1,'[other]Sheet1'!$B$2`, 0},
		{"Sheet1", `Sheet1!$A$1,Sheet1!B2`, 0},
		{"Sheet1", `Sheet1!$A$1,'Sheet1:Other'!$B$2`, 0},
		{"Sheet1", `Sheet1!$A$1,`, 0},
		{"Sheet1", `,Sheet1!$A$1`, 0},
		{"Sheet1", `Sheet1!$A$1,,Sheet1!$B$2`, 0},
		{"Sheet1", `'Sheet1!$A$1,Sheet1!$B$2`, 0},
		{"Sheet1", strings.Repeat("x", 4097), 0},
	} {
		if got := parseNativePrintAreaSet(tc.ref, tc.name); len(got) != tc.count {
			t.Errorf("%q: %+v", tc.ref, got)
		}
	}
	refs := []string{}
	for i := 1; i <= 17; i++ {
		refs = append(refs, fmt.Sprintf("Sheet1!$A$%d", i))
	}
	if len(parseNativePrintAreaSet(strings.Join(refs[:16], ","), "Sheet1")) != 16 || parseNativePrintAreaSet(strings.Join(refs, ","), "Sheet1") != nil {
		t.Fatal("rectangle count bound")
	}
	got := parseNativePrintAreaSet(`Sheet1!$F$6:$G$8,Sheet1!$B$2:$D$4`, "Sheet1")
	if got[0].Row != 5 || got[1].Row != 1 {
		t.Fatal("source order lost", got)
	}
}

func TestNativePrintAreaSetsRoutedReadOnly(t *testing.T) {
	for _, strict := range []bool{false, true} {
		parts := nativeWorkbookFixture(strict)
		parts["Charts/chart1.xml"] = previewChartFixture()
		parts["Book/Workbook.xml"] = strings.Replace(parts["Book/Workbook.xml"], `</sheets>`, `</sheets><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">'Data Set'!$F$6:$G$8,'Data Set'!$B$2:$D$4</definedName></definedNames>`, 1)
		source := buildZip(t, parts)
		before := bytes.Clone(source)
		native, err := ExtractNativeWorkbookV2(source)
		if err != nil {
			t.Fatal(err)
		}
		beforeJSON, err := EncodeNativeWorkbookV2(native)
		if err != nil {
			t.Fatal(err)
		}
		got, err := InspectNativeWorkbookObjectsV1(source)
		if err != nil {
			t.Fatal(err)
		}
		if len(got.PrintAreaSets) != 2 || got.PrintAreaSets[0].SheetID != "7" || got.PrintAreaSets[0].SheetPart != "Sheets/s1.xml" || len(got.PrintAreaSets[0].Areas) != 2 || got.PrintAreaSets[1].Status != "unavailable" || got.PrintAreas[0].Status != "unavailable" {
			t.Fatalf("%+v", got)
		}
		native, err = ExtractNativeWorkbookV2(source)
		if err != nil {
			t.Fatal(err)
		}
		afterJSON, err := EncodeNativeWorkbookV2(native)
		if err != nil || !bytes.Equal(before, source) || !bytes.Equal(beforeJSON, afterJSON) || bytes.Contains(afterJSON, []byte(`"print_area_sets"`)) {
			t.Fatal("mutation authority or bytes changed")
		}
		encoded, err := json.Marshal(got.PrintAreaSets[1])
		if err != nil || bytes.Contains(encoded, []byte(`"areas"`)) {
			t.Fatal("unavailable includes areas")
		}
	}
}
