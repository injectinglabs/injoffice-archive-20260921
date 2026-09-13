package xlsxpatch

import (
	"bytes"
	"reflect"
	"strings"
	"testing"
)

func nativePrintSourceCellFixture(strict bool, formula, cells, extras string) map[string]string {
	parts := nativeWorkbookFixture(strict)
	parts["Charts/chart1.xml"] = previewChartFixture()
	ns := spreadsheetMLTransitional
	if strict {
		ns = spreadsheetMLStrict
	}
	parts["Sheets/s1.xml"] = `<worksheet xmlns="` + ns + `"><sheetData><row r="1">` + cells + `</row></sheetData>` + extras + `</worksheet>`
	name := `<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">` + formula + `</definedName></definedNames>`
	parts["Book/Workbook.xml"] = strings.Replace(parts["Book/Workbook.xml"], `</sheets>`, `</sheets>`+name, 1)
	return parts
}

func TestNativePrintSourceCellOffsetRouted(t *testing.T) {
	cells := `<c r="H1"><v>+3</v></c><c r="I1" t="n"><v>-2</v></c><c r="J1"><v>02</v></c><c r="K1"><v>3</v></c>`
	for _, strict := range []bool{false, true} {
		for _, tc := range []struct {
			formula string
			want    NativePrintAreaRectV1
			inputs  int
		}{
			{`OFFSET('Data Set'!$D$3:$F$5,'Data Set'!$H$1,'Data Set'!$I$1)`, NativePrintAreaRectV1{5, 1, 7, 3}, 2},
			{`OFFSET('Data Set'!$D$3:$F$5,'Data Set'!$H$1,'Data Set'!$I$1,'Data Set'!$J$1,'Data Set'!$K$1)`, NativePrintAreaRectV1{5, 1, 6, 3}, 4},
			{`OFFSET('Data Set'!$A$1,0,0, 'Data Set'!$J$1 ,3)`, NativePrintAreaRectV1{0, 0, 1, 2}, 1},
			{`OFFSET('Data Set'!$D$3,'Data Set'!$H$1,-2)`, NativePrintAreaRectV1{5, 1, 5, 1}, 1},
			{`OFFSET('Data Set'!$A$1,0,0,'Data Set'!$J$1,'Data Set'!$J$1)`, NativePrintAreaRectV1{0, 0, 1, 1}, 2},
		} {
			t.Run(tc.formula, func(t *testing.T) {
				source := buildZip(t, nativePrintSourceCellFixture(strict, tc.formula, cells, ""))
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
				if entry.Status != "available" || len(entry.Areas) != 1 || entry.Areas[0] != tc.want || entry.SheetID != "7" || entry.SheetPart != "Sheets/s1.xml" || len(entry.Warnings) != 2+tc.inputs {
					t.Fatalf("%+v", entry)
				}
				if !strings.Contains(entry.Warnings[0], "saved numeric source cells") || entry.Warnings[1] != "Source _xlnm.Print_Area formula: "+tc.formula || !strings.Contains(entry.Warnings[2], "Reinspect") {
					t.Fatalf("missing provenance: %+v", entry.Warnings)
				}
				for _, warning := range entry.Warnings {
					if len(warning) > 4096 || strings.ContainsAny(warning, "\r\n\t") {
						t.Fatalf("unsafe warning transport: %q", warning)
					}
				}
				if got.PackageSHA256 != workbook.Source.PackageSHA256 || got.PrintAreas[0].Status != "unavailable" {
					t.Fatal("source join or legacy literal projection changed")
				}
				after, err := ExtractNativeWorkbookV2(source)
				if err != nil || !reflect.DeepEqual(workbook, after) || !bytes.Equal(source, before) {
					t.Fatal("source or mutation authority changed")
				}
			})
		}
	}
}

func TestNativePrintSourceCellOffsetRefusals(t *testing.T) {
	formula := `OFFSET('Data Set'!$D$3,0,0,'Data Set'!$H$1,1)`
	for _, strict := range []bool{false, true} {
		for _, tc := range []struct{ name, cell, extras string }{
			{"missing", "", ""},
			{"blank", `<c r="H1"/>`, ""},
			{"text", `<c r="H1" t="inlineStr"><is><t>2</t></is></c>`, ""},
			{"boolean", `<c r="H1" t="b"><v>1</v></c>`, ""},
			{"error", `<c r="H1" t="e"><v>#N/A</v></c>`, ""},
			{"date", `<c r="H1" t="d"><v>2026-09-12</v></c>`, ""},
			{"fraction", `<c r="H1"><v>2.5</v></c>`, ""},
			{"integral-decimal-outside-grammar", `<c r="H1"><v>2.0</v></c>`, ""},
			{"exponent-outside-grammar", `<c r="H1"><v>2e0</v></c>`, ""},
			{"too-large", `<c r="H1"><v>1048577</v></c>`, ""},
			{"too-many-digits", `<c r="H1"><v>00000002</v></c>`, ""},
			{"zero-height", `<c r="H1"><v>0</v></c>`, ""},
			{"negative-height", `<c r="H1"><v>-2</v></c>`, ""},
			{"grid-overflow", `<c r="H1"><v>1048576</v></c>`, ""},
			{"formula-cache", `<c r="H1"><f>1+1</f><v>2</v></c>`, ""},
			{"array-cache", `<c r="H1"><f t="array" ref="H1:I1">1+1</f><v>2</v></c>`, ""},
			{"covered-array-cache", `<c r="G1"><f t="array" ref="G1:H1">1+1</f><v>2</v></c><c r="H1"><v>2</v></c>`, ""},
			{"shared-cache", `<c r="G1"><f t="shared" si="0" ref="G1:H1">1+1</f><v>2</v></c><c r="H1"><f t="shared" si="0"/><v>2</v></c>`, ""},
			{"covered-data-table-cache", `<c r="G1"><f t="dataTable" ref="G1:H1"/><v>2</v></c><c r="H1"><v>2</v></c>`, ""},
			{"metadata", `<c r="H1" cm="1"><v>2</v></c>`, ""},
			{"unknown-attribute", `<c r="H1" opaque="1"><v>2</v></c>`, ""},
			{"foreign-child", `<c r="H1"><v>2</v><x:opaque xmlns:x="urn:test"/></c>`, ""},
			{"merge-owner", `<c r="H1"><v>2</v></c>`, `<mergeCells count="1"><mergeCell ref="H1:I1"/></mergeCells>`},
			{"protected-sheet", `<c r="H1"><v>2</v></c>`, `<sheetProtection sheet="1"/>`},
		} {
			t.Run(tc.name, func(t *testing.T) {
				got, err := InspectNativeWorkbookObjectsV1(buildZip(t, nativePrintSourceCellFixture(strict, formula, tc.cell, tc.extras)))
				if err != nil {
					t.Fatal(err)
				}
				if entry := got.PrintAreaSets[0]; entry.Status != "unavailable" || len(entry.Areas) != 0 {
					t.Fatalf("accepted %s: %+v", tc.name, entry)
				}
			})
		}
	}
	for _, arg := range []string{`H1`, `$H$1`, `'Data Set'!H1`, `'Data Set'!$H$1:$H$1`, `'Data Set'!$H$1:$I$1`, `Hidden!$H$1`, `'[book]Data Set'!$H$1`, `Name`, `SUM('Data Set'!$H$1)`, `'Data Set'!$H$1+1`, "\u00a0'Data Set'!$H$1", "'Data Set'!$H$1\u2003"} {
		parts := nativePrintSourceCellFixture(false, strings.Replace(formula, `'Data Set'!$H$1`, arg, 1), `<c r="H1"><v>2</v></c>`, "")
		got, err := InspectNativeWorkbookObjectsV1(buildZip(t, parts))
		if err != nil || got.PrintAreaSets[0].Status != "unavailable" {
			t.Fatalf("%q: %+v, %v", arg, got, err)
		}
	}
}

func TestNativePrintSourceCellOffsetOwnershipAndReinspection(t *testing.T) {
	formula := `OFFSET('Data Set'!$A$1,0,0,'Data Set'!$H$1,1)`
	parts := nativePrintSourceCellFixture(false, formula, `<c r="H1"><v>2</v></c>`, "")
	first, err := InspectNativeWorkbookObjectsV1(buildZip(t, parts))
	if err != nil {
		t.Fatal(err)
	}
	parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `<v>2</v>`, `<v>4</v>`, 1)
	second, err := InspectNativeWorkbookObjectsV1(buildZip(t, parts))
	if err != nil || first.PackageSHA256 == second.PackageSHA256 || first.PrintAreaSets[0].Areas[0].EndRow != 1 || second.PrintAreaSets[0].Areas[0].EndRow != 3 || !strings.Contains(second.PrintAreaSets[0].Warnings[2], " = 4.") {
		t.Fatalf("did not re-read current source cell: %+v, %v", second, err)
	}
	for _, transform := range []func(string) string{
		func(s string) string { return strings.Replace(s, ` localSheetId="0"`, "", 1) },
		func(s string) string { return strings.Replace(s, `localSheetId="0"`, `localSheetId="1"`, 1) },
		func(s string) string {
			return strings.Replace(s, `<definedName `, `<definedName xmlns="urn:foreign" `, 1)
		},
		func(s string) string {
			return strings.Replace(s, `</definedNames>`, `<definedName name="_xlnm.Print_Area" localSheetId="0">`+formula+`</definedName></definedNames>`, 1)
		},
		func(s string) string {
			return strings.Replace(s, `</definedNames>`, `<definedName name="_xlnm.Print_Titles" localSheetId="0">BAD</definedName></definedNames>`, 1)
		},
	} {
		invalid := nativePrintSourceCellFixture(false, formula, `<c r="H1"><v>2</v></c>`, "")
		invalid["Book/Workbook.xml"] = transform(invalid["Book/Workbook.xml"])
		got, err := InspectNativeWorkbookObjectsV1(buildZip(t, invalid))
		if err != nil || got.PrintAreaSets[0].Status != "unavailable" {
			t.Fatalf("source ownership bypass: %+v, %v", got, err)
		}
	}
	workbook, err := ExtractNativeWorkbookV2(buildZip(t, parts))
	if err != nil {
		t.Fatal(err)
	}
	sheet := workbook.Sheets[0]
	sheet.Cells = append(sheet.Cells, sheet.Cells[0])
	if got, warnings := parseNativePrintSourceCellOffset(formula, sheet); got != nil || len(warnings) != 0 {
		t.Fatal("ambiguous source cell accepted")
	}
}
