package xlsxpatch

import (
	"bytes"
	"reflect"
	"strings"
	"testing"
)

func nativeConditionalFixture(strict bool) map[string]string {
	parts := nativeWorkbookFixture(strict)
	parts["Charts/chart1.xml"] = previewChartFixture()
	ns := spreadsheetMLTransitional
	if strict {
		ns = spreadsheetMLStrict
	}
	parts["Sheets/s1.xml"] = `<worksheet xmlns="` + ns + `"><dimension ref="A1:A3"/><sheetData><row r="1"><c r="A1"><v>-2</v></c></row><row r="2"><c r="A2"><v>0</v></c></row><row r="3"><c r="A3"><f>1+2</f><v>3</v></c></row></sheetData><conditionalFormatting sqref="A1:A3"><cfRule type="cellIs" operator="greaterThan" priority="7" stopIfTrue="1" dxfId="0"><formula>0</formula></cfRule></conditionalFormatting></worksheet>`
	parts["Meta/Styles.style"] = strings.Replace(parts["Meta/Styles.style"], `</styleSheet>`, `<dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FF12aB34"/></patternFill></fill></dxf></dxfs></styleSheet>`, 1)
	return parts
}

func TestNativeConditionalFillSourceAndComparison(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, tc := range []struct {
			operator string
			want     []bool
		}{
			{"equal", []bool{false, true, false}}, {"notEqual", []bool{true, false, true}},
			{"lessThan", []bool{true, false, false}}, {"lessThanOrEqual", []bool{true, true, false}},
			{"greaterThan", []bool{false, false, true}}, {"greaterThanOrEqual", []bool{false, true, true}},
		} {
			parts := nativeConditionalFixture(strict)
			parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `operator="greaterThan"`, `operator="`+tc.operator+`"`, 1)
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
			if len(got.ConditionalFills) != 1 {
				t.Fatalf("%+v", got.ConditionalFills)
			}
			entry := got.ConditionalFills[0]
			if entry.Status != "available" || entry.SheetID != "7" || entry.SheetPart != "Sheets/s1.xml" || entry.Rule.Fill != "#12AB34" || entry.Rule.Priority != 7 || !entry.Rule.StopIfTrue || entry.Rule.DxfID != 0 || entry.Rule.Operand != "0" || len(entry.Cells) != 3 {
				t.Fatalf("%+v", entry)
			}
			for i, want := range tc.want {
				if entry.Cells[i].Matches != want || entry.Cells[i].Row != i || entry.Cells[i].Column != 0 || entry.Cells[i].Cached != (i == 2) {
					t.Fatalf("%s %+v", tc.operator, entry.Cells)
				}
			}
			if !strings.Contains(strings.Join(entry.Warnings, " "), "freshness is unknown") {
				t.Fatal(entry.Warnings)
			}
			nativeAfter, err := ExtractNativeWorkbookV2(source)
			if err != nil || !reflect.DeepEqual(nativeBefore, nativeAfter) || !bytes.Equal(before, source) {
				t.Fatal("source or native mutation authority changed")
			}
			found := false
			for _, u := range nativeAfter.Unsupported {
				if u.Code == "CONDITIONAL_FORMATTING" {
					found = true
				}
			}
			if !found {
				t.Fatal("conditional preservation inventory weakened")
			}
		}
	}
}

func TestNativeConditionalFillAtomicRefusals(t *testing.T) {
	cases := []struct{ name, part, old, replacement string }{
		{"metadata", "Sheets/s1.xml", `<c r="A2">`, `<c r="A2" cm="1">`},
		{"opaque-cell", "Sheets/s1.xml", `<c r="A2"><v>0</v></c>`, `<c r="A2"><v>0</v><extra/></c>`},
		{"formula-attributes", "Sheets/s1.xml", `<f>1+2</f>`, `<f ca="1">1+2</f>`},
		{"misplaced-dxfs", "Meta/Styles.style", `</styleSheet>`, `<unknown><dxfs/></unknown></styleSheet>`},
		{"second-rule", "Sheets/s1.xml", `</conditionalFormatting>`, `<cfRule type="cellIs" priority="8" dxfId="0" operator="equal"><formula>0</formula></cfRule></conditionalFormatting>`},
		{"second-range", "Sheets/s1.xml", `</worksheet>`, `<conditionalFormatting sqref="Z1"><cfRule type="expression" priority="8"><formula>1</formula></cfRule></conditionalFormatting></worksheet>`},
		{"misplaced-rule", "Sheets/s1.xml", `</worksheet>`, `<unknown><cfRule type="cellIs" priority="9"/></unknown></worksheet>`},
		{"foreign-rule", "Sheets/s1.xml", `<cfRule `, `<cfRule xmlns="urn:foreign" `},
		{"ext", "Sheets/s1.xml", `</worksheet>`, `<extLst/></worksheet>`},
		{"unknown-option", "Sheets/s1.xml", `type="cellIs"`, `type="cellIs" rank="1"`},
		{"pivot", "Sheets/s1.xml", `sqref="A1:A3"`, `sqref="A1:A3" pivot="1"`},
		{"union", "Sheets/s1.xml", `sqref="A1:A3"`, `sqref="A1:A3 C1"`},
		{"oversized", "Sheets/s1.xml", `sqref="A1:A3"`, `sqref="A1:A4097"`},
		{"merge", "Sheets/s1.xml", `</worksheet>`, `<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells></worksheet>`},
		{"table", "Sheets/s1.xml", `</worksheet>`, `<tableParts count="0"/></worksheet>`},
		{"expression", "Sheets/s1.xml", `type="cellIs"`, `type="expression"`},
		{"cell-operand", "Sheets/s1.xml", `<formula>0</formula>`, `<formula>A1</formula>`},
		{"decimal-operand", "Sheets/s1.xml", `<formula>0</formula>`, `<formula>0.5</formula>`},
		{"between", "Sheets/s1.xml", `operator="greaterThan"`, `operator="between"`},
		{"stop", "Sheets/s1.xml", `stopIfTrue="1"`, `stopIfTrue="yes"`},
		{"priority", "Sheets/s1.xml", `priority="7"`, `priority="0"`},
		{"dxf-missing", "Sheets/s1.xml", `dxfId="0"`, `dxfId="1"`},
		{"missing-cache", "Sheets/s1.xml", `<f>1+2</f><v>3</v>`, `<f>1+2</f>`},
		{"decimal-cell", "Sheets/s1.xml", `<v>3</v>`, `<v>3.5</v>`},
		{"leading-zero", "Sheets/s1.xml", `<v>3</v>`, `<v>03</v>`},
		{"too-many-digits", "Sheets/s1.xml", `<v>3</v>`, `<v>1000000000000000</v>`},
		{"blank", "Sheets/s1.xml", `<c r="A2"><v>0</v></c>`, `<c r="A2"/>`},
		{"text", "Sheets/s1.xml", `<c r="A2"><v>0</v></c>`, `<c r="A2" t="inlineStr"><is><t>0</t></is></c>`},
		{"error", "Sheets/s1.xml", `<c r="A2"><v>0</v></c>`, `<c r="A2" t="e"><v>#VALUE!</v></c>`},
		{"other-dxf-style", "Meta/Styles.style", `<dxf><fill>`, `<dxf><font><b/></font><fill>`},
		{"pattern", "Meta/Styles.style", `patternType="solid"><fgColor rgb="FF12aB34"`, `patternType="darkGrid"><fgColor rgb="FF12aB34"`},
		{"theme", "Meta/Styles.style", `rgb="FF12aB34"`, `theme="5"`},
		{"alpha", "Meta/Styles.style", `rgb="FF12aB34"`, `rgb="0012AB34"`},
		{"foreign-dxf", "Meta/Styles.style", `<dxf>`, `<dxf xmlns="urn:foreign">`},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			parts := nativeConditionalFixture(false)
			parts[tc.part] = strings.Replace(parts[tc.part], tc.old, tc.replacement, 1)
			source := buildZip(t, parts)
			got, err := InspectNativeWorkbookObjectsV1(source)
			if err != nil {
				t.Fatal(err)
			}
			if len(got.ConditionalFills) != 1 || got.ConditionalFills[0].Status != "unavailable" || got.ConditionalFills[0].Rule != nil || len(got.ConditionalFills[0].Cells) != 0 {
				t.Fatalf("partial or unqualified conditional overlay: %+v", got.ConditionalFills)
			}
		})
	}
}

func TestNativeConditionalFillIntegerEdgesAndAggregateBound(t *testing.T) {
	for _, value := range []string{"-0", "999999999999999", "-999999999999999"} {
		parts := nativeConditionalFixture(false)
		parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `<v>0</v>`, `<v>`+value+`</v>`, 1)
		pkg, err := openNativeWorkbookPackage(buildZip(t, parts))
		if err != nil {
			t.Fatal(err)
		}
		workbook, err := ExtractNativeWorkbookV2(buildZip(t, parts))
		if err != nil {
			t.Fatal(err)
		}
		got := previewNativeConditionalFill(pkg, workbook.Sheets[0], 3)
		if got.Status != "available" || got.Cells[1].Lexical != value {
			t.Fatalf("%+v", got)
		}
		if limited := previewNativeConditionalFill(pkg, workbook.Sheets[0], 2); limited.Status != "unavailable" || len(limited.Cells) != 0 {
			t.Fatal(limited)
		}
	}
}

func TestNativeConditionalFillRefusesArrayFollowersAndDuplicateProjection(t *testing.T) {
	parts := nativeConditionalFixture(false)
	parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `<c r="A1"><v>-2</v></c>`, `<c r="A1"><f t="array" ref="A1:A3">ROW(A1:A3)</f><v>-2</v></c>`, 1)
	parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `<f>1+2</f>`, "", 1)
	parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `sqref="A1:A3"`, `sqref="A2:A3"`, 1)
	source := buildZip(t, parts)
	got, err := InspectNativeWorkbookObjectsV1(source)
	if err != nil {
		t.Fatal(err)
	}
	if len(got.ConditionalFills) != 1 || got.ConditionalFills[0].Status != "unavailable" {
		t.Fatal(got.ConditionalFills)
	}
	parts = nativeConditionalFixture(false)
	source = buildZip(t, parts)
	pkg, err := openNativeWorkbookPackage(source)
	if err != nil {
		t.Fatal(err)
	}
	workbook, err := ExtractNativeWorkbookV2(source)
	if err != nil {
		t.Fatal(err)
	}
	sheet := workbook.Sheets[0]
	sheet.Cells = append(sheet.Cells, sheet.Cells[0])
	if entry := previewNativeConditionalFill(pkg, sheet, 4096); entry.Status != "unavailable" {
		t.Fatal(entry)
	}
	sheet = workbook.Sheets[0]
	sheet.Cells[0].Ref = "B1"
	if entry := previewNativeConditionalFill(pkg, sheet, 4096); entry.Status != "unavailable" {
		t.Fatal(entry)
	}
}
