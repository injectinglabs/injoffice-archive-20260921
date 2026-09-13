package xlsxpatch

import (
	"strings"
	"testing"
)

func nativePrintCountaFixture(strict bool, cells, extras string) map[string]string {
	parts := nativePrintSourceCellFixture(strict, `OFFSET('Data Set'!$A$1,0,0,COUNTA('Data Set'!$A$1:$F$1),4)`, cells, extras)

	return parts
}

func nativePrintCountaExtract(t *testing.T, parts map[string]string) (*NativeWorkbookV2, *nativePrintCountaSourceContext) {
	t.Helper()
	wb, err := ExtractNativeWorkbookV2(buildZip(t, parts))
	if err != nil {
		t.Fatal(err)
	}
	ctx := newNativePrintCountaSourceContext(wb, []byte(parts["Book/Workbook.xml"]), wb.Source.PackageSHA256)
	return wb, ctx
}

func TestNativePrintCountaLiteralOccupancy(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, tc := range []struct {
			name, cells string
			want        int
			ok          bool
		}{
			{"Microsoft-five-values", `<c r="A1"><v>39790</v></c><c r="B1"><v>19</v></c><c r="C1"><v>22.24</v></c><c r="D1" t="b"><v>1</v></c><c r="E1" t="e"><v>#DIV/0!</v></c>`, 5, true},
			{"absent", "", 0, true},
			{"styled-blank", `<c r="A1" s="0"/>`, 0, true},
			{"numeric-empty", `<c r="A1"><v/></c>`, 0, true},
			{"zero-false-exponent", `<c r="A1"><v>0</v></c><c r="B1" t="b"><v>0</v></c><c r="C1"><v>1e-2</v></c>`, 3, true},
			{"whitespace", `<c r="A1" t="inlineStr"><is><t xml:space="preserve"> </t></is></c>`, 1, true},
			{"empty-inline-explicit", `<c r="A1" t="inlineStr"><is><t/></is></c>`, 0, false},
			{"empty-inline-omitted", `<c r="A1" t="inlineStr"/>`, 0, false},
			{"plain-shared", `<c r="A1" t="s"><v>1</v></c>`, 1, true},
			{"rich-shared", `<c r="A1" t="s"><v>0</v></c>`, 0, false},
			{"date", `<c r="A1" t="d"><v>2026-09-12</v></c>`, 1, true},
			{"formula-cache", `<c r="A1"><f>1+1</f><v>2</v></c>`, 0, false},
			{"formula-no-cache", `<c r="A1"><f>1+1</f></c>`, 0, false},
			{"formula-empty-text", `<c r="A1" t="str"><f>IF(1,"","")</f></c>`, 0, false},
			{"formula-string-literal", `<c r="A1" t="str"><v>text</v></c>`, 0, false},
		} {
			t.Run(tc.name, func(t *testing.T) {
				wb, ctx := nativePrintCountaExtract(t, nativePrintCountaFixture(strict, tc.cells, ""))
				n, ok := ctx.count(&wb.Sheets[0], NativePrintAreaRectV1{0, 0, 0, 5})
				if ok != tc.ok || ok && n != tc.want {
					t.Fatalf("got %d,%v want %d,%v", n, ok, tc.want, tc.ok)
				}
			})
		}
	}
}

func TestNativePrintCountaCompletenessRefusals(t *testing.T) {
	for _, tc := range []struct{ name, cells, extras string }{
		{"opaque-cell", `<c r="A1"><v>1</v><foreign xmlns="urn:test"/></c>`, ""},
		{"metadata", `<c r="A1" cm="1"><v>1</v></c>`, ""},
		{"group-absent-follower", `<c r="A1"><f t="array" ref="A1:F1">1</f><v>1</v></c>`, ""},
		{"group-cache-follower", `<c r="A1"><f t="dataTable" ref="A1:F1"/><v>1</v></c><c r="B1"><v>2</v></c>`, ""},
		{"foreign-sheet", `<c r="A1"><v>1</v></c>`, `<x:opaque xmlns:x="urn:test"><x:c r="B1">hidden</x:c></x:opaque>`},
		{"alternate-content", "", `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Fallback><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></mc:Fallback></mc:AlternateContent>`},
		{"merge", `<c r="A1"><v>1</v></c>`, `<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>`},
	} {
		t.Run(tc.name, func(t *testing.T) {
			wb, ctx := nativePrintCountaExtract(t, nativePrintCountaFixture(false, tc.cells, tc.extras))
			if _, ok := ctx.count(&wb.Sheets[0], NativePrintAreaRectV1{0, 1, 0, 5}); ok {
				t.Fatal("incomplete source accepted")
			}
		})
	}
}

func TestNativePrintCountaContextOwnership(t *testing.T) {
	parts := nativePrintCountaFixture(false, `<c r="A1"><v>1</v></c>`, "")
	wb, ctx := nativePrintCountaExtract(t, parts)
	if ctx == nil || !ctx.certified[&wb.Sheets[0]] {
		t.Fatal("positive inventory must be allowed")
	}
	copySheet := wb.Sheets[0]
	if _, ok := ctx.count(&copySheet, NativePrintAreaRectV1{0, 0, 0, 0}); ok {
		t.Fatal("detached caller view accepted")
	}
	if newNativePrintCountaSourceContext(wb, []byte(parts["Book/Workbook.xml"]), strings.Repeat("0", 64)) != nil {
		t.Fatal("stale source hash")
	}
	for _, extra := range []string{`<unknown/>`, `<x:definedNames xmlns:x="urn:test"/>`, `<calcPr calcMode="auto"/>`, `<bookViews><workbookView activeTab="1"/></bookViews>`, `<bookViews unknown="1"><workbookView/></bookViews>`, `<calcPr calcId="191029"><unknown/></calcPr>`} {
		raw := strings.Replace(parts["Book/Workbook.xml"], `</workbook>`, extra+`</workbook>`, 1)
		if newNativePrintCountaSourceContext(wb, []byte(raw), wb.Source.PackageSHA256) != nil {
			t.Fatalf("unknown workbook content accepted: %s", extra)
		}
	}
	wb.Sheets[0].Cells = append(wb.Sheets[0].Cells, wb.Sheets[0].Cells[0])
	if nativePrintCountaCertifySheet(wb, &wb.Sheets[0]) {
		t.Fatal("duplicate cells accepted")
	}
}

func TestNativePrintCountaEmptySharedAndRejectedEmptyCache(t *testing.T) {
	parts := nativePrintCountaFixture(false, `<c r="A1" t="s"><v>1</v></c>`, "")
	parts["Meta/Strings.xml"] = strings.Replace(parts["Meta/Strings.xml"], `<t>Plain_x0020_Text</t>`, `<t/>`, 1)
	wb, ctx := nativePrintCountaExtract(t, parts)
	if n, ok := ctx.count(&wb.Sheets[0], NativePrintAreaRectV1{0, 0, 0, 0}); !ok || n != 1 {
		t.Fatal("explicit empty shared text counts once")
	}
	parts = nativePrintCountaFixture(false, `<c r="A1" t="str"><f>IF(1,"","")</f><v/></c>`, "")
	if _, err := ExtractNativeWorkbookV2(buildZip(t, parts)); err == nil {
		t.Fatal("current extractor's invalid empty cache lexical unexpectedly admitted")
	}
}

func TestNativePrintCountaUnsupportedInventoryGuards(t *testing.T) {
	base := nativePrintCountaFixture(false, `<c r="A1"><v>1</v></c>`, "")
	for _, tc := range []struct {
		name string
		edit func(*NativeWorkbookV2)
	}{
		{"unknown-global", func(w *NativeWorkbookV2) {
			w.Unsupported = append(w.Unsupported, NativeWorkbookUnsupportedV2{Code: "FUTURE_UNKNOWN", Capability: "styles", ScopeID: "workbook", PartName: nativeWorkbookString("Meta/Styles.style")})
		}},
		{"known-code-wrong-capability", func(w *NativeWorkbookV2) { w.Unsupported[0].Capability = "drawings" }},
		{"known-code-core-part", func(w *NativeWorkbookV2) { w.Unsupported[0].PartName = nativeWorkbookString(w.Source.WorkbookPart) }},
		{"known-code-selected-sheet", func(w *NativeWorkbookV2) { w.Unsupported[0].PartName = nativeWorkbookString(w.Sheets[0].PartName) }},
		{"missing-part", func(w *NativeWorkbookV2) { w.Unsupported[0].PartName = nil }},
		{"unknown-selected-sheet", func(w *NativeWorkbookV2) {
			w.Unsupported = append(w.Unsupported, NativeWorkbookUnsupportedV2{Code: "FUTURE_UNKNOWN", Capability: "extensions", ScopeID: "sheet:7", PartName: nativeWorkbookString("Sheets/s1.xml")})
		}},
		{"opaque-SST", func(w *NativeWorkbookV2) {
			w.Unsupported = append(w.Unsupported, NativeWorkbookUnsupportedV2{Code: "SHARED_STRING_TABLE_OPAQUE_CONTENT", Capability: "rich-text", ScopeID: "workbook", PartName: nativeWorkbookString("Meta/Strings.xml")})
		}},
		{"inconsistent-other-sheet-scope", func(w *NativeWorkbookV2) {
			w.Unsupported = append(w.Unsupported, NativeWorkbookUnsupportedV2{Code: "FUTURE_UNKNOWN", Capability: "extensions", ScopeID: "sheet:9", PartName: nativeWorkbookString("Sheets/s1.xml")})
		}},
	} {
		t.Run(tc.name, func(t *testing.T) {
			wb, _ := nativePrintCountaExtract(t, base)
			tc.edit(wb)
			ctx := newNativePrintCountaSourceContext(wb, []byte(base["Book/Workbook.xml"]), wb.Source.PackageSHA256)
			if _, ok := ctx.count(&wb.Sheets[0], NativePrintAreaRectV1{0, 0, 0, 0}); ok {
				t.Fatal("unsafe certificate")
			}
		})
	}
	wb, _ := nativePrintCountaExtract(t, base)
	wb.Unsupported = append(wb.Unsupported, NativeWorkbookUnsupportedV2{Code: "FUTURE_UNKNOWN", Capability: "extensions", ScopeID: "sheet:9", PartName: nativeWorkbookString("Sheets/s2.xml")})
	ctx := newNativePrintCountaSourceContext(wb, []byte(base["Book/Workbook.xml"]), wb.Source.PackageSHA256)
	if _, ok := ctx.count(&wb.Sheets[0], NativePrintAreaRectV1{0, 0, 0, 0}); !ok {
		t.Fatal("explicitly disjoint other sheet must not spoil occupancy")
	}
}

func TestNativePrintCountaWorkbookGateShapes(t *testing.T) {
	parts := nativePrintCountaFixture(false, `<c r="A1"><v>1</v></c>`, "")
	wb, _ := nativePrintCountaExtract(t, parts)
	for _, replacement := range []string{`<calcPr calcId="1" unknown="0"/>`, `<calcPr calcId="-1"/>`, `<calcPr calcId="1.0"/>`, `<calcPr calcId="4294967296"/>`, `<calcPr calcId="1"><unknown/></calcPr>`, `<calcPr xmlns="urn:foreign" calcId="1"/>`} {
		raw := strings.Replace(parts["Book/Workbook.xml"], `<calcPr calcId="191029"/>`, replacement, 1)
		if newNativePrintCountaSourceContext(wb, []byte(raw), wb.Source.PackageSHA256) != nil {
			t.Fatalf("bad calcPr accepted: %s", replacement)
		}
	}
	for _, replacement := range []string{`<bookViews><workbookView activeTab="1"/></bookViews>`, `<bookViews opaque="1"><workbookView/></bookViews>`, `<bookViews><workbookView><unknown/></workbookView></bookViews>`} {
		raw := strings.Replace(parts["Book/Workbook.xml"], `<bookViews><workbookView/></bookViews>`, replacement, 1)
		if newNativePrintCountaSourceContext(wb, []byte(raw), wb.Source.PackageSHA256) != nil {
			t.Fatalf("bad bookViews accepted: %s", replacement)
		}
	}
}

func TestNativePrintCountaScanBudget(t *testing.T) {
	parts := nativePrintCountaFixture(false, "", "")
	wb, _ := nativePrintCountaExtract(t, parts)
	// Private invariant test: exact complete inventory bound, independent of
	// coordinate area bound, without producing an oversized fixture archive.
	wb.Sheets[0].Cells = make([]NativeWorkbookCellV2, nativePrintCountaLimit)
	for i := range wb.Sheets[0].Cells {
		wb.Sheets[0].Cells[i] = NativeWorkbookCellV2{Row: i, Column: 0, Ref: cellReference(i, 0), Editable: true}
	}
	if !nativePrintCountaCertifySheet(wb, &wb.Sheets[0]) {
		t.Fatal("exact scan budget")
	}
	wb.Sheets[0].Cells = append(wb.Sheets[0].Cells, NativeWorkbookCellV2{Row: nativePrintCountaLimit, Column: 0, Ref: cellReference(nativePrintCountaLimit, 0), Editable: true})
	if nativePrintCountaCertifySheet(wb, &wb.Sheets[0]) {
		t.Fatal("scan budget exceeded")
	}
}
