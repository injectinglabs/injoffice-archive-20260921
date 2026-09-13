package xlsxpatch

import (
	"reflect"
	"strconv"
	"strings"
	"testing"
)

func TestNativePrintCountaOffsetArguments(t *testing.T) {
	wb, ctx := nativePrintCountaExtract(t, nativePrintCountaFixture(false, `<c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="H1"><v>3</v></c>`, ""))
	for _, tc := range []struct {
		formula string
		want    NativePrintAreaRectV1
		deps    int
	}{
		{`OFFSET('Data Set'!$A$1,0,0,COUNTA('Data Set'!$A$1:$B$1),4)`, NativePrintAreaRectV1{0, 0, 1, 3}, 1},
		{`OFFSET('Data Set'!$A$1,COUNTA('Data Set'!$A$1:$B$1),0,1,1)`, NativePrintAreaRectV1{2, 0, 2, 0}, 1},
		{`OFFSET('Data Set'!$A$1,COUNTA('Data Set'!$C$1:$D$1),0,1,'Data Set'!$H$1)`, NativePrintAreaRectV1{0, 0, 0, 2}, 2},
		{`OFFSET('Data Set'!$A$1,0,0, COUNTA( 'Data Set'!$A$1:$B$1 ) ,'Data Set'!$H$1)`, NativePrintAreaRectV1{0, 0, 1, 2}, 2},
	} {
		b := nativePrintCountaBudget{}
		area, w := parseNativePrintCountaOffset(tc.formula, &wb.Sheets[0], ctx, &b)
		if area == nil || *area != tc.want || len(w) != tc.deps || b.dependencies != tc.deps {
			t.Fatalf("%s: %+v %+v %+v", tc.formula, area, w, b)
		}
	}
}

func TestNativePrintCountaBudgetsAndAtomicRefusal(t *testing.T) {
	wb, ctx := nativePrintCountaExtract(t, nativePrintCountaFixture(false, `<c r="A1"><v>1</v></c>`, ""))
	valid := `OFFSET('Data Set'!$A$1,0,0,COUNTA('Data Set'!$A$1:$A$100000),1)`
	b := nativePrintCountaBudget{}
	if area, _ := parseNativePrintCountaOffset(valid, &wb.Sheets[0], ctx, &b); area == nil || b.coordinates != 100000 {
		t.Fatal("exact coordinate budget")
	}
	for _, formula := range []string{
		strings.Replace(valid, "100000", "100001", 1),
		`OFFSET('Data Set'!$A$1,0,0,COUNTA('Data Set'!$B$1:$B$2),1)`,
		`OFFSET('Data Set'!$A$1,0,0,COUNTA('Data Set'!$A:$A),1)`,
		`OFFSET('Data Set'!$A$1,0,0,COUNTA('Data Set'!A1:A2),1)`,
		`OFFSET('Data Set'!$A$1,0,0,COUNTA('Other'!$A$1:$A$2),1)`,
		`OFFSET('Data Set'!$A$1,0,0,COUNTA(COUNTA('Data Set'!$A$1)),1)`,
		`OFFSET('Data Set'!$A$1,0,0,COUNTA('Data Set'!$A$1,'Data Set'!$B$1),1)`,
		`OFFSET('Data Set'!$A$1,0,0,COUNTA('Data Set'!$A$1)-1,1)`,
		`OFFSET('Data Set'!$A$1,0,0,COUNTA('Data Set'!$A$1),'Data Set'!$H$1)`,
	} {
		budget := nativePrintCountaBudget{}
		before := budget
		area, w := parseNativePrintCountaOffset(formula, &wb.Sheets[0], ctx, &budget)
		if area != nil || w != nil || !reflect.DeepEqual(before, budget) {
			t.Fatalf("non-atomic refusal %s: %+v %+v %+v", formula, area, w, budget)
		}
	}
	small := `OFFSET('Data Set'!$A$1,0,0,COUNTA('Data Set'!$A$1),1)`
	budget := nativePrintCountaBudget{}
	for i := 0; i < 4; i++ {
		if a, _ := parseNativePrintCountaOffset(small, &wb.Sheets[0], ctx, &budget); a == nil {
			t.Fatal("four occurrences must fit")
		}
	}
	before := budget
	if a, w := parseNativePrintCountaOffset(small, &wb.Sheets[0], ctx, &budget); a != nil || w != nil || budget != before {
		t.Fatal("fifth occurrence accepted")
	}
}

func TestNativePrintCountaAreasGlobalBudgets(t *testing.T) {
	wb, ctx := nativePrintCountaExtract(t, nativePrintCountaFixture(false, `<c r="A1"><v>1</v></c><c r="H1"><v>1</v></c>`, ""))
	dynamic := `OFFSET('Data Set'!$A$1,0,0,COUNTA('Data Set'!$A$1),1)`
	text := `'Data Set'!$C$1,` + dynamic + `,OFFSET('Data Set'!$E$1,0,0,'Data Set'!$H$1,1)`
	areas, w := parseNativePrintCountaAreas(text, &wb.Sheets[0], ctx)
	if len(areas) != 3 || areas[0].Column != 2 || areas[1].Column != 0 || areas[2].Column != 4 || len(w) != 2 {
		t.Fatalf("source order or dependencies: %+v %+v", areas, w)
	}
	for _, bad := range []string{text + `,'Data Set'!$A$1`, text + `,UNKNOWN('Data Set'!$G$1)`, text + `,OFFSET('Data Set'!$G$1,0,0,COUNTA('Data Set'!$B$1),1)`, `(` + text + `)`, dynamic + `,OFFSET('Data Set'!$G$1,0,0,COUNTA(OFFSET('Data Set'!$A$1,0,0)),1)`} {
		if a, w := parseNativePrintCountaAreas(bad, &wb.Sheets[0], ctx); a != nil || w != nil {
			t.Fatalf("partial set: %s %+v %+v", bad, a, w)
		}
	}
	exact := dynamic
	for _, col := range []string{"C", "E", "G"} {
		exact += `,OFFSET('Data Set'!$` + col + `$1,0,0,'Data Set'!$H$1,1)`
	}
	if a, w := parseNativePrintCountaAreas(exact, &wb.Sheets[0], ctx); len(a) != 4 || len(w) != 4 {
		t.Fatal("four mixed inputs must fit")
	}
	if a, w := parseNativePrintCountaAreas(exact+`,OFFSET('Data Set'!$I$1,0,0,'Data Set'!$H$1,1)`, &wb.Sheets[0], ctx); a != nil || w != nil {
		t.Fatal("fifth global input")
	}
	// Non-overlapping output areas but cumulative input coordinate work exceeds
	// the limit: input overlap does not cancel work or dependencies.
	repeated := `OFFSET('Data Set'!$A$1,0,0,COUNTA('Data Set'!$A$1:$A$50001),1),OFFSET('Data Set'!$C$1,0,0,COUNTA('Data Set'!$A$1:$A$50000),1)`
	if a, w := parseNativePrintCountaAreas(repeated, &wb.Sheets[0], ctx); a != nil || w != nil {
		t.Fatal("global coordinate limit")
	}
	for i := 0; i < 15; i++ {
		dynamic += `,'Data Set'!$C$` + strconv.Itoa(i+1)
	}
	if a, _ := parseNativePrintCountaAreas(dynamic, &wb.Sheets[0], ctx); len(a) != 16 {
		t.Fatal("exact sixteen")
	}
	if a, w := parseNativePrintCountaAreas(dynamic+`,'Data Set'!$C$16`, &wb.Sheets[0], ctx); a != nil || w != nil {
		t.Fatal("seventeen")
	}
}

func TestNativePrintCountaQuotedSheetTokens(t *testing.T) {
	for _, name := range []string{"O'Brien,(Q)", "表"} {
		parts := nativePrintCountaFixture(false, `<c r="A1"><v>1</v></c>`, "")
		escaped := strings.ReplaceAll(name, "'", "''")
		parts["Book/Workbook.xml"] = strings.ReplaceAll(parts["Book/Workbook.xml"], "Data_x0020_Set", name)
		parts["Book/Workbook.xml"] = strings.ReplaceAll(parts["Book/Workbook.xml"], "Data Set", escaped)
		wb, ctx := nativePrintCountaExtract(t, parts)
		formula := `OFFSET('` + escaped + `'!$A$1,0,0,COUNTA('` + escaped + `'!$A$1),1)`
		if a, _ := parseNativePrintCountaAreas(formula, &wb.Sheets[0], ctx); len(a) != 1 {
			t.Fatalf("quoted name %s", name)
		}
	}
}

func TestNativePrintCountaSavedNameRouting(t *testing.T) {
	formula := `'Data Set'!$C$1,OFFSET('Data Set'!$A$1,0,0,COUNTA('Data Set'!$A$1:$B$1),1)`
	for _, strict := range []bool{false, true} {
		parts := nativePrintCountaFixture(strict, `<c r="A1"><v>1</v></c><c r="B1" t="b"><v>0</v></c>`, "")
		parts["Book/Workbook.xml"] = strings.Replace(parts["Book/Workbook.xml"], `OFFSET('Data Set'!$A$1,0,0,COUNTA('Data Set'!$A$1:$F$1),4)`, formula, 1)
		wb, ctx := nativePrintCountaExtract(t, parts)
		public, err := InspectNativeWorkbookObjectsV1(buildZip(t, parts))
		if err != nil {
			t.Fatal(err)
		}
		got := public.PrintAreaSets[0]
		if public.PackageSHA256 != wb.Source.PackageSHA256 || public.PrintAreas[0].Status != "unavailable" {
			t.Fatal("public source join or legacy selector changed")
		}
		if !reflect.DeepEqual(got, previewNativePrintAreaSets([]byte(parts["Book/Workbook.xml"]), wb.Sheets, ctx)[0]) {
			t.Fatal("public/private projection differs")
		}
		if got.Status != "available" || len(got.Areas) != 2 || got.Areas[0].Column != 2 || got.Areas[1].EndRow != 1 || got.SheetID != "7" || got.SheetPart != "Sheets/s1.xml" || len(got.Warnings) != 3 || got.Warnings[1] != "Source _xlnm.Print_Area formula: "+formula || !strings.Contains(got.Warnings[2], " = 2.") {
			t.Fatalf("%+v", got)
		}
		// Omitting the private actual-source context never enables source counting.
		if noContext := previewNativePrintAreaSets([]byte(parts["Book/Workbook.xml"]), wb.Sheets)[0]; noContext.Status != "unavailable" {
			t.Fatal("counted without certified source context")
		}
		for _, bad := range []string{
			strings.Replace(parts["Book/Workbook.xml"], `</definedNames>`, `<definedName name="_xlnm.Print_Titles" localSheetId="0">COUNTA('Data Set'!$A$1)</definedName></definedNames>`, 1),
			strings.Replace(parts["Book/Workbook.xml"], `</definedNames>`, `<definedName name="_xlnm.Print_Area" localSheetId="0">'Data Set'!$A$1</definedName></definedNames>`, 1),
			strings.Replace(parts["Book/Workbook.xml"], formula, formula+`,'Data Set'!$A$1`, 1),
		} {
			got := previewNativePrintAreaSets([]byte(bad), wb.Sheets, ctx)[0]
			if got.Status != "unavailable" || got.Areas != nil || len(got.Warnings) != 1 {
				t.Fatal("non-atomic title/name/overlap refusal")
			}
		}
	}
}

func TestNativePrintCountaFormulaLengthAndSyntaxBounds(t *testing.T) {
	wb, ctx := nativePrintCountaExtract(t, nativePrintCountaFixture(false, `<c r="A1"><v>1</v></c>`, ""))
	formula := `OFFSET('Data Set'!$A$1,0,0,COUNTA('Data Set'!$A$1),1)`
	exact := strings.Replace(formula, "OFFSET(", "OFFSET("+strings.Repeat(" ", 2048-len(formula)), 1)
	if a, w := parseNativePrintCountaAreas(exact, &wb.Sheets[0], ctx); len(a) != 1 || len(w) != 1 {
		t.Fatal("exact 2048 bytes")
	}
	for _, bad := range []string{" " + exact, "=" + formula, strings.Replace(formula, "COUNTA(", "counta(", 1), strings.Replace(formula, "COUNTA(", "COUNTA(\u00a0", 1), strings.Replace(formula, "COUNTA(", "COUNTA(\t", 1), strings.Replace(formula, "$A$1),1)", "$A$1),0)", 1), strings.Replace(formula, "OFFSET('Data Set'!$A$1", "OFFSET('Data Set'!$XFD$1048576", 1) + ",'Data Set'!$XFD$1048576"} {
		if a, w := parseNativePrintCountaAreas(bad, &wb.Sheets[0], ctx); a != nil || w != nil {
			t.Fatalf("syntax/bound refusal: %q", bad)
		}
	}
}
