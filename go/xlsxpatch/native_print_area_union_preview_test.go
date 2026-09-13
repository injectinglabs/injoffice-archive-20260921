package xlsxpatch

import (
	"bytes"
	"fmt"
	"reflect"
	"strings"
	"testing"
)

func TestNativePrintFormulaUnionRouted(t *testing.T) {
	formula := `'Data Set'!$F$6:$G$8, OFFSET('Data Set'!$D$3,3,-2), OFFSET('Data Set'!$A$10,0,0,'Data Set'!$H$1,2)`
	want := []NativePrintAreaRectV1{{5, 5, 7, 6}, {5, 1, 5, 1}, {9, 0, 10, 1}}
	for _, strict := range []bool{false, true} {
		source := buildZip(t, nativePrintSourceCellFixture(strict, formula, `<c r="H1"><v>02</v></c>`, ""))
		before := bytes.Clone(source)
		workbook, err := ExtractNativeWorkbookV2(source)
		if err != nil {
			t.Fatal(err)
		}
		got, err := InspectNativeWorkbookObjectsV1(source)
		if err != nil {
			t.Fatal(err)
		}
		entry := got.PrintAreaSets[0]
		if entry.Status != "available" || !reflect.DeepEqual(entry.Areas, want) || entry.SheetID != "7" || entry.SheetPart != "Sheets/s1.xml" || got.PackageSHA256 != workbook.Source.PackageSHA256 {
			t.Fatalf("source range/order/join changed: %+v", entry)
		}
		if len(entry.Warnings) != 3 || entry.Warnings[1] != "Source _xlnm.Print_Area formula: "+formula || !strings.Contains(entry.Warnings[2], "'Data Set'!$H$1 = 02") || !strings.Contains(entry.Warnings[0], "source order") {
			t.Fatalf("missing provenance: %+v", entry.Warnings)
		}
		if got.PrintAreas[0].Status != "unavailable" {
			t.Fatal("legacy literal selector changed")
		}
		after, err := ExtractNativeWorkbookV2(source)
		if err != nil || !reflect.DeepEqual(workbook, after) || !bytes.Equal(source, before) {
			t.Fatal("source or mutation authority changed")
		}
	}
}

func TestNativePrintFormulaUnionBoundsAndPrecedence(t *testing.T) {
	sheet := NativeWorkbookSheetV2{Name: "Data Set", Editable: true}
	for _, count := range []int{16, 17} {
		parts := []string{`OFFSET('Data Set'!$A$1,0,0)`}
		for row := 2; row <= count; row++ {
			parts = append(parts, fmt.Sprintf("'Data Set'!$A$%d", row))
		}
		got, warnings := parseNativePrintFormulaUnion(strings.Join(parts, ","), sheet)
		if count == 16 && (len(got) != 16 || len(warnings) != 0) || count == 17 && got != nil {
			t.Fatalf("component bound %d: %+v", count, got)
		}
		objects, err := InspectNativeWorkbookObjectsV1(buildZip(t, nativePrintSourceCellFixture(false, strings.Join(parts, ","), "", "")))
		if err != nil || count == 16 && len(objects.PrintAreaSets[0].Areas) != 16 || count == 17 && objects.PrintAreaSets[0].Status != "unavailable" {
			t.Fatalf("routed component bound %d: %+v, %v", count, objects, err)
		}
	}
	base := `OFFSET('Data Set'!$A$1,0,0),'Data Set'!$A$2`
	for _, size := range []int{2048, 2049} {
		text := base + strings.Repeat(" ", size-len(base))
		got, _ := parseNativePrintFormulaUnion(text, sheet)
		if size == 2048 && len(got) != 2 || size == 2049 && got != nil {
			t.Fatalf("text bound %d: %+v", size, got)
		}
	}
	// Exercise the existing literal parser's independent text budget directly.
	longName := strings.Repeat("S", 140)
	literals := []string{}
	for row := 1; row <= 16; row++ {
		literals = append(literals, fmt.Sprintf("'%s'!$A$%d", longName, row))
	}
	text := strings.Join(literals, ",")
	if len(text) <= 2048 || len(parseNativePrintAreaSet(text, longName)) != 16 {
		t.Fatal("old literal parser's text budget changed")
	}
	for _, old := range []string{`'Data Set'!$A$1,'Data Set'!$A$2`, `OFFSET('Data Set'!$A$1,0,0)`, `'Data Set'!$A$1, 'Data Set'!$A$2`} {
		if got, _ := parseNativePrintFormulaUnion(old, sheet); got != nil {
			t.Fatalf("fallback took literal-only or single-component ownership: %q", old)
		}
	}
}

func TestNativePrintFormulaUnionSourceBudget(t *testing.T) {
	ref := `'Data Set'!$H$1`
	four := `OFFSET('Data Set'!$A$1,0,0,` + ref + `,` + ref + `),OFFSET('Data Set'!$A$10,0,0,` + ref + `,` + ref + `)`
	five := four + `,OFFSET('Data Set'!$A$20,0,0,` + ref + `,1)`
	for _, formula := range []string{four, five} {
		got, err := InspectNativeWorkbookObjectsV1(buildZip(t, nativePrintSourceCellFixture(false, formula, `<c r="H1"><v>2</v></c>`, "")))
		if err != nil {
			t.Fatal(err)
		}
		entry := got.PrintAreaSets[0]
		if formula == four {
			if entry.Status != "available" || len(entry.Areas) != 2 || len(entry.Warnings) != 6 {
				t.Fatalf("four occurrences refused: %+v", entry)
			}
			for _, warning := range entry.Warnings {
				if len(warning) > 4096 || strings.ContainsAny(warning, "\r\n\t") {
					t.Fatal("warning transport overflow")
				}
			}
		} else if entry.Status != "unavailable" || len(entry.Areas) != 0 || len(entry.Warnings) != 1 {
			t.Fatalf("fifth occurrence or partial dependency leak: %+v", entry)
		}
	}
}

func TestNativePrintFormulaUnionQuotedSheetTokens(t *testing.T) {
	for _, name := range []string{"O'Brien,(!", "Data\u00a0Set", "Sheet)(,Name"} {
		quoted := "'" + strings.ReplaceAll(name, "'", "''") + "'!"
		text := "OFFSET(" + quoted + "$A$1,0,0), " + quoted + "$C$3"
		got, _ := parseNativePrintFormulaUnion(text, NativeWorkbookSheetV2{Name: name})
		if !reflect.DeepEqual(got, []NativePrintAreaRectV1{{0, 0, 0, 0}, {2, 2, 2, 2}}) {
			t.Fatalf("quoted token altered: %q: %+v", name, got)
		}
	}
}

func TestNativePrintFormulaUnionAtomicRefusals(t *testing.T) {
	prefix := `'Data Set'!$F$6:$G$8,`
	for _, invalid := range []string{
		`OFFSET('Data Set'!$D$3,3,2)`, // Resolves to F6, overlapping the first area.
		`OFFSET('Data Set'!$D$3,3,-2),'Data Set'!$B$6`,
		`OFFSET('Data Set'!$A$1,-1,0)`,
		`OFFSET('Data Set'!$A$1,0,0,'Data Set'!$J$1,1)`, // Missing source input.
		`OFFSET('Data Set'!$A$1,0,0,'Data Set'!$H$1,1)`, // Formula cache, not literal.
		`OFFSET('Data Set'!$A$1,0,0,COUNTA(A:A),1)`,
		`OFFSET('Data Set'!$A$1,0,0,1,1`,
		`OFFSET('Data Set'!$A$1,0,0))`,
		`OFFSET('Data Set'!$A$1,0,0),`,
		` ,OFFSET('Data Set'!$A$1,0,0)`,
		`OFFSET(OFFSET('Data Set'!$A$1,0,0),0,0)`,
		`OFFSET('Data Set'!$A$1,0,0)+1`,
		`OFFSET('Data Set'!$A$1,0,0),'Data Set'!A2`,
		`OFFSET('Data Set'!$A$1,0,0),Hidden!$A$2`,
		`OFFSET('Data Set'!$A$1,0,0),'[external]Data Set'!$A$2`,
		`OFFSET('Data Set'!$A$1,0,0),'Data Set!$A$2`,
		`OFFSET('Data Set'!$A$1,0,0),'Data Set'!$A$2'bad'`,
		`OFFSET('Data Set'!$A$1,0,0),'Data Set'!$A$2!$C$3`,
		"OFFSET('Data Set'!$A$1,0,0),\u00a0'Data Set'!$A$2",
		"OFFSET('Data Set'!$A$1,0,0)\n",
	} {
		parts := nativePrintSourceCellFixture(false, prefix+invalid, `<c r="H1"><f>1+1</f><v>2</v></c>`, "")
		got, err := InspectNativeWorkbookObjectsV1(buildZip(t, parts))
		if err != nil {
			t.Fatalf("%q: %v", invalid, err)
		}
		if entry := got.PrintAreaSets[0]; entry.Status != "unavailable" || len(entry.Areas) != 0 || len(entry.Warnings) != 1 {
			t.Fatalf("partial set or unsupported expression accepted: %q: %+v", invalid, entry)
		}
	}
	valid := prefix + `OFFSET('Data Set'!$A$1,0,0)`
	for _, formula := range []string{"=" + valid, "(" + valid + ")"} {
		if got, _ := parseNativePrintFormulaUnion(formula, NativeWorkbookSheetV2{Name: "Data Set"}); got != nil {
			t.Fatalf("general expression accepted: %q", formula)
		}
	}
	for _, invalidName := range []string{
		`<definedName name="_xlnm.Print_Area" localSheetId="0">` + valid + `</definedName>`,
		`<definedName name="_xlnm.Print_Titles" localSheetId="0">BAD</definedName>`,
	} {
		parts := nativePrintSourceCellFixture(false, valid, "", "")
		parts["Book/Workbook.xml"] = strings.Replace(parts["Book/Workbook.xml"], `</definedNames>`, invalidName+`</definedNames>`, 1)
		got, err := InspectNativeWorkbookObjectsV1(buildZip(t, parts))
		if err != nil || got.PrintAreaSets[0].Status != "unavailable" {
			t.Fatalf("name/title ownership bypass: %+v, %v", got, err)
		}
	}
}
