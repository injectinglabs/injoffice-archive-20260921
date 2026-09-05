package xlsxpatch

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"strings"
	"testing"
)

func nativeMergedWorkbook(t *testing.T, strict bool, worksheet string) []byte {
	t.Helper()
	parts := nativeWorkbookFixture(strict)
	parts["Sheets/s1.xml"] = worksheet
	return buildZip(t, parts)
}

func nativeMergedWorksheet(namespace, content string) string {
	return `<worksheet xmlns="` + namespace + `">` + content + `</worksheet>`
}

func TestExtractNativeWorkbookV1MergedRangesTransitionalAndStrict(t *testing.T) {
	for _, strict := range []bool{false, true} {
		strict := strict
		name := "transitional"
		namespace := spreadsheetMLTransitional
		if strict {
			name, namespace = "strict", spreadsheetMLStrict
		}
		t.Run(name, func(t *testing.T) {
			worksheet := nativeMergedWorksheet(namespace,
				`<cols><col min="2" max="2" hidden="1" style="1"/></cols>`+
					`<sheetData><row r="1" hidden="1"><c r="A1"><v>1</v></c><c r="B1" s="1"/></row>`+
					`<row r="2"><c r="A2" s="1"/><c r="B2" s="1"/></row>`+
					`<row r="3"><c r="C3"><v>2</v></c></row></sheetData>`+
					`<mergeCells count="2"><mergeCell ref="D4:E4"/><mergeCell ref="$a$1:$b$2"/></mergeCells>`)
			workbook, err := ExtractNativeWorkbookV1(nativeMergedWorkbook(t, strict, worksheet))
			if err != nil {
				t.Fatal(err)
			}
			sheet := workbook.Sheets[0]
			want := []NativeWorkbookMergedRangeV1{
				{Ref: "A1:B2", Row: 0, Column: 0, EndRow: 1, EndColumn: 1, Editable: false},
				{Ref: "D4:E4", Row: 3, Column: 3, EndRow: 3, EndColumn: 4, Editable: false},
			}
			if fmt.Sprint(sheet.MergedRanges) != fmt.Sprint(want) {
				t.Fatalf("merged ranges = %#v, want %#v", sheet.MergedRanges, want)
			}
			if sheet.PartName != "Sheets/s1.xml" {
				t.Fatalf("actual worksheet part spelling = %q", sheet.PartName)
			}
			if !sheet.Editable || sheet.RefusalCode != nil {
				t.Fatalf("unrelated cells should remain sheet-editable: %#v", sheet)
			}
			for _, ref := range []string{"A1", "B1", "A2", "B2"} {
				if findNativeCell(t, workbook, "7", ref).Editable {
					t.Errorf("merged cell %s remained editable", ref)
				}
			}
			if !findNativeCell(t, workbook, "7", "C3").Editable {
				t.Error("unrelated cell C3 became non-editable")
			}
			if anchor := findNativeCell(t, workbook, "7", "A1"); anchor.Value == nil || anchor.Value.Lexical == nil || *anchor.Value.Lexical != "1" {
				t.Fatalf("top-left value authority was lost: %#v", anchor)
			}
			if findNativeCell(t, workbook, "7", "B1").StyleID != 1 || findNativeCell(t, workbook, "7", "B2").StyleID != 1 {
				t.Error("covered blank-cell styles were not retained")
			}
			if len(sheet.Rows) != 1 || !sheet.Rows[0].Hidden || len(sheet.Columns) != 1 || !sheet.Columns[0].Hidden {
				t.Fatalf("hidden row/column projection changed: rows=%#v columns=%#v", sheet.Rows, sheet.Columns)
			}
			if !hasNativeWorkbookUnsupported(workbook, "MERGED_CELLS") {
				t.Error("raw merged-cell source authority was not inventoried")
			}
			mergeSources := 0
			for _, item := range workbook.Unsupported {
				if item.Code != "MERGED_CELLS" {
					continue
				}
				mergeSources++
				if item.PartName == nil || *item.PartName != sheet.PartName || item.Preservation != "preserve-exact" {
					t.Errorf("merged source authority lost exact part spelling: %#v", item)
				}
			}
			if mergeSources != 1 {
				t.Errorf("MERGED_CELLS source diagnostics = %d, want 1", mergeSources)
			}
		})
	}
}

func TestExtractNativeWorkbookV1RejectsMalformedOrAmbiguousMergedRanges(t *testing.T) {
	namespace := spreadsheetMLTransitional
	opposing := spreadsheetMLStrict
	tests := []struct {
		name, content, want string
	}{
		{name: "before sheetData", content: `<mergeCells><mergeCell ref="A1:B1"/></mergeCells><sheetData/>`, want: "outside canonical schema order"},
		{name: "duplicate containers", content: `<sheetData/><mergeCells><mergeCell ref="A1:B1"/></mergeCells><mergeCells><mergeCell ref="C1:D1"/></mergeCells>`, want: "duplicated"},
		{name: "root mergeCell", content: `<sheetData/><mergeCell ref="A1:B1"/>`, want: "outside mergeCells"},
		{name: "opposing container", content: `<sheetData/><s:mergeCells xmlns:s="` + opposing + `"><s:mergeCell ref="A1:B1"/></s:mergeCells>`, want: "opposing Strict/Transitional"},
		{name: "opposing child", content: `<sheetData/><mergeCells><s:mergeCell xmlns:s="` + opposing + `" ref="A1:B1"/></mergeCells>`, want: "unsupported direct child"},
		{name: "missing ref", content: `<sheetData/><mergeCells><mergeCell/></mergeCells>`, want: "requires a non-empty ref"},
		{name: "empty ref", content: `<sheetData/><mergeCells><mergeCell ref=""/></mergeCells>`, want: "requires a non-empty ref"},
		{name: "trailing absolute marker", content: `<sheetData/><mergeCells><mergeCell ref="A$1$:B1"/></mergeCells>`, want: "reference grammar"},
		{name: "duplicate absolute marker", content: `<sheetData/><mergeCells><mergeCell ref="$$A1:B1"/></mergeCells>`, want: "reference grammar"},
		{name: "absolute marker before colon", content: `<sheetData/><mergeCells><mergeCell ref="A1$:B1"/></mergeCells>`, want: "reference grammar"},
		{name: "leading zero row", content: `<sheetData/><mergeCells><mergeCell ref="A01:B1"/></mergeCells>`, want: "reference grammar"},
		{name: "reversed", content: `<sheetData/><mergeCells><mergeCell ref="B2:A1"/></mergeCells>`, want: "bounds are reversed"},
		{name: "row out of bounds", content: `<sheetData/><mergeCells><mergeCell ref="A1:A1048577"/></mergeCells>`, want: "invalid merged range"},
		{name: "column out of bounds", content: `<sheetData/><mergeCells><mergeCell ref="XFD1:XFE1"/></mergeCells>`, want: "invalid merged range"},
		{name: "single cell", content: `<sheetData/><mergeCells><mergeCell ref="A1"/></mergeCells>`, want: "must span multiple cells"},
		{name: "canonical duplicate", content: `<sheetData/><mergeCells><mergeCell ref="A1:B2"/><mergeCell ref="$a$1:$b$2"/></mergeCells>`, want: "duplicate merged range"},
		{name: "partial overlap", content: `<sheetData/><mergeCells><mergeCell ref="A1:C2"/><mergeCell ref="C2:D3"/></mergeCells>`, want: "overlaps another"},
		{name: "nested overlap", content: `<sheetData/><mergeCells><mergeCell ref="A1:D4"/><mergeCell ref="B2:C3"/></mergeCells>`, want: "overlaps another"},
		{name: "count mismatch", content: `<sheetData/><mergeCells count="2"><mergeCell ref="A1:B1"/></mergeCells>`, want: "does not match"},
		{name: "count overflow", content: `<sheetData/><mergeCells count="100001"><mergeCell ref="A1:B1"/></mergeCells>`, want: "exceeds 100000"},
		{name: "empty container", content: `<sheetData/><mergeCells count="0"/>`, want: "at least one"},
		{name: "container attribute", content: `<sheetData/><mergeCells count="1" future="x"><mergeCell ref="A1:B1"/></mergeCells>`, want: "unexpected semantic attribute"},
		{name: "duplicate count", content: `<sheetData/><mergeCells count="1" count="1"><mergeCell ref="A1:B1"/></mergeCells>`, want: "duplicate attribute"},
		{name: "child attribute", content: `<sheetData/><mergeCells><mergeCell ref="A1:B1" future="x"/></mergeCells>`, want: "unexpected semantic attribute"},
		{name: "duplicate ref", content: `<sheetData/><mergeCells><mergeCell ref="A1:B1" ref="A1:B1"/></mergeCells>`, want: "duplicate attribute"},
		{name: "nonempty child", content: `<sheetData/><mergeCells><mergeCell ref="A1:B1">x</mergeCell></mergeCells>`, want: "must be empty"},
		{name: "direct text", content: `<sheetData/><mergeCells>x<mergeCell ref="A1:B1"/></mergeCells>`, want: "unsupported direct text"},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			workbook, err := ExtractNativeWorkbookV1(nativeMergedWorkbook(t, false, nativeMergedWorksheet(namespace, test.content)))
			if err == nil || workbook != nil || !strings.Contains(err.Error(), test.want) {
				t.Fatalf("workbook=%#v err=%v, want error containing %q", workbook, err, test.want)
			}
		})
	}
}

func TestExtractNativeWorkbookV1RejectsCoveredNonAnchorValuesAndFormulas(t *testing.T) {
	for name, covered := range map[string]string{
		"value":   `<c r="B1"><v>2</v></c>`,
		"formula": `<c r="B1"><f>1+1</f><v>2</v></c>`,
	} {
		t.Run(name, func(t *testing.T) {
			worksheet := nativeMergedWorksheet(spreadsheetMLTransitional,
				`<sheetData><row r="1"><c r="A1"><v>1</v></c>`+covered+`</row></sheetData>`+
					`<mergeCells><mergeCell ref="A1:B1"/></mergeCells>`)
			workbook, err := ExtractNativeWorkbookV1(nativeMergedWorkbook(t, false, worksheet))
			if err == nil || workbook != nil || !strings.Contains(err.Error(), "covered non-anchor cell B1 contains a value or formula") {
				t.Fatalf("workbook=%#v err=%v", workbook, err)
			}
		})
	}
}

func TestParseNativeMergedRangesRefusesActualCardinalityPastLimit(t *testing.T) {
	var xmlText strings.Builder
	xmlText.WriteString(`<mergeCells xmlns="` + spreadsheetMLTransitional + `">`)
	for row := 1; row <= NativeXLSXMaxMergedRanges+1; row++ {
		fmt.Fprintf(&xmlText, `<mergeCell ref="A%d:B%d"/>`, row, row)
	}
	xmlText.WriteString(`</mergeCells>`)
	decoder := xml.NewDecoder(bytes.NewBufferString(xmlText.String()))
	root, err := nativeReadXMLRoot(decoder)
	if err != nil {
		t.Fatal(err)
	}
	extractor := nativeWorkbookExtractor{namespace: spreadsheetMLTransitional}
	ranges, err := extractor.parseNativeMergedRanges(decoder, root)
	if err == nil || ranges != nil || !strings.Contains(err.Error(), "exceeds 100000 ranges") {
		t.Fatalf("ranges=%d err=%v", len(ranges), err)
	}
}
