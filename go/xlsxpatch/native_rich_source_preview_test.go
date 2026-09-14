package xlsxpatch

import (
	"bytes"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func richSourceFixture() map[string]string {
	p := sourceStyleFixture()
	delete(p, "Meta/Strings.xml")
	p["[Content_Types].xml"] = strings.Replace(p["[Content_Types].xml"], `<Override PartName="/Meta/Strings.XML" ContentType="`+nativeSharedStringsType+`"/>`, "", 1)
	p["Book/_rels/Workbook.xml.rels"] = strings.Replace(p["Book/_rels/Workbook.xml.rels"], `<Relationship Id="rStrings" Type="`+relTypeSharedStringsTransitional+`" Target="../Meta/Strings.XML"/>`, "", 1)
	p["Meta/Styles.style"] = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="General"/></numFmts><fonts count="2"><font/><font><name val="Arial"/><sz val="10"/><color rgb="FF000000"/><b/><i val="0"/><strike val="0"/><u val="none"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border><left style="thin"><color rgb="FFD3D3D3"/></left><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="164" fontId="1" fillId="0" borderId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="general" vertical="top" wrapText="1" textRotation="0" readingOrder="1"/></xf></cellXfs></styleSheet>`
	p["Sheets/s1.xml"] = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><outlinePr/></sheetPr><sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews><cols><col min="1" max="2" customWidth="1" width="26"/><col min="3" max="3" customWidth="1" width="0.0234375"/></cols><sheetData><row r="1" ht="17" customHeight="0"><c r="A1" s="1" t="inlineStr"><is><r><rPr><b/><rFont val="Arial"/><sz val="10"/><color rgb="FF000000"/><strike val="0"/><u val="none"/></rPr><t>Year</t></r><r><rPr><b val="0"/><color rgb="FFFF0000"/></rPr><t xml:space="preserve"> 🧮 3</t></r></is></c><c r="B1" s="1" t="inlineStr"><is><t xml:space="preserve">Budget </t></is></c></row><row r="2" ht="17.05" customHeight="0"><c r="A2" s="1"><v>1</v></c><c r="B2" s="1"><v>42250557.5799999</v></c></row><row r="3" ht="409" customHeight="1"/></sheetData><headerFooter alignWithMargins="0"/></worksheet>`
	return p
}
func TestRichSourcePreview(t *testing.T) {
	b := buildZip(t, richSourceFixture())
	before := bytes.Clone(b)
	v, e := PreviewNativeRichSourceV1(b)
	if e != nil {
		t.Fatal(e)
	}
	if v.Protocol != "injoffice.xlsx.rich-source-preview" || v.Version != 1 || !v.ReadOnly || v.Fidelity != "approximate" || v.PackageSHA256 != nativeWorkbookDigest(b) || v.ParentCount.Observed != 1 || v.ParentCount.Declaration != "absent" || v.StrictError == "" {
		t.Fatalf("authority: %#v", v)
	}
	if len(v.Styles) != 1 || v.Styles[0].ID != 1 || !v.Styles[0].Wrap || v.Styles[0].FontName != "Arial" || v.Styles[0].Borders["left"] != "#D3D3D3" || len(v.UnusedStyleIDs) != 1 || v.UnusedStyleIDs[0] != 0 {
		t.Fatalf("style ownership: %#v", v)
	}
	if len(v.Sheet.Cells) != 4 || v.Sheet.Cells[3].Lexical != "42250557.5799999" || v.Sheet.Cells[1].Text != "Budget " || len(v.RichCells) != 1 || v.RichCells[0].Text != "Year 🧮 3" || v.RichCells[0].Runs[1].Start != 4 || v.RichCells[0].Runs[1].End != 9 || *v.RichCells[0].Runs[1].Style.FontColor != "#FF0000" {
		t.Fatalf("source joins: %#v", v)
	}
	if len(v.OmittedRows) != 1 || v.OmittedRows[0].Index != 2 || v.OmittedRows[0].Size != 409 || len(v.OmittedColumns) != 1 || v.OmittedColumns[0].Index != 2 || v.OmittedColumns[0].Size != 0.0234375 {
		t.Fatalf("outside geometry: %#v", v)
	}
	if _, e = ExtractNativeWorkbookV2(b); e == nil {
		t.Fatal("strict extraction accepted missing count")
	}
	if _, e = PreviewNativeSourceStylesV1(b); e == nil {
		t.Fatal("V1 accepted missing count")
	}
	if _, e = PreviewNativeSourceStylesV2(b); e == nil {
		t.Fatal("V2 accepted missing count")
	}
	tx := NativeWorkbookMutationTransactionV1{ExpectedRevision: "rev:" + strings.TrimPrefix(nativeWorkbookDigest(b), "sha256:"), Cells: []CellMutation{{OperationID: "x", SheetID: "7", Kind: CellSetValue, Cell: CellRef{Row: 1, Column: 1}, Value: "7"}}}
	if _, e = ApplyNativeWorkbookMutationTransactionV1(b, tx); e == nil {
		t.Fatal("mutation accepted missing count")
	}
	if !bytes.Equal(b, before) {
		t.Fatal("source changed")
	}
	j, e := json.Marshal(v)
	if e != nil {
		t.Fatal(e)
	}
	if bytes.Contains(j, []byte(`"editable"`)) || bytes.Contains(j, []byte(`"revision"`)) {
		t.Fatal("mutation authority leaked")
	}
	if dir := os.Getenv("XLSX_RICH_SOURCE_EVIDENCE_DIR"); dir != "" {
		if e = os.MkdirAll(dir, 0755); e != nil {
			t.Fatal(e)
		}
		for n, b := range map[string][]byte{"source.xlsx": b, "preview.json": j} {
			if e = os.WriteFile(filepath.Join(dir, n), b, 0644); e != nil {
				t.Fatal(e)
			}
		}
	}
}
func TestRichSourceRefusals(t *testing.T) {
	for _, tc := range []struct{ name, part, old, new string }{
		{"present-count", "Meta/Styles.style", "<cellStyleXfs>", `<cellStyleXfs count="1">`},
		{"malformed-count", "Meta/Styles.style", "<cellStyleXfs>", `<cellStyleXfs count="bad">`},
		{"foreign-count", "Meta/Styles.style", "<cellStyleXfs>", `<cellStyleXfs xmlns:q="urn:q" q:count="1">`},
		{"missing-font-count", "Meta/Styles.style", `fonts count="2"`, `fonts`},
		{"wrong-font-count", "Meta/Styles.style", `fonts count="2"`, `fonts count="3"`},
		{"unknown-style-root", "Meta/Styles.style", "</styleSheet>", "<extLst/></styleSheet>"},
		{"duplicate-parent-table", "Meta/Styles.style", "</cellStyleXfs>", "</cellStyleXfs><cellStyleXfs><xf/></cellStyleXfs>"},
		{"bad-parent-ref", "Meta/Styles.style", `fontId="1" fillId="0"`, `xfId="20" fontId="1" fillId="0"`},
		{"inactive-font-owner", "Meta/Styles.style", `applyFont="1"`, `applyFont="0"`},
		{"missing-font-owner", "Meta/Styles.style", `applyFont="1"`, ``},
		{"non-General", "Meta/Styles.style", `formatCode="General"`, `formatCode="0.00"`},
		{"active-strike", "Meta/Styles.style", `strike val="0"`, `strike val="1"`},
		{"active-underline", "Meta/Styles.style", `u val="none"`, `u val="single"`},
		{"reading-order", "Meta/Styles.style", `readingOrder="1"`, `readingOrder="2"`},
		{"font-theme", "Meta/Styles.style", `color rgb="FF000000"`, `color theme="1"`},
		{"default-style", "Sheets/s1.xml", `r="A2" s="1"`, `r="A2" s="0"`},
		{"sparse", "Sheets/s1.xml", `<c r="A2" s="1"><v>1</v></c>`, ``},
		{"formula", "Sheets/s1.xml", `<v>1</v>`, `<f>1</f><v>1</v>`},
		{"merge", "Sheets/s1.xml", `</worksheet>`, `<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells></worksheet>`},
		{"hidden-row", "Sheets/s1.xml", `r="2" ht=`, `r="2" hidden="1" ht=`},
		{"missing-height", "Sheets/s1.xml", `ht="17.05"`, ``},
		{"missing-width", "Sheets/s1.xml", `width="26"`, ``},
		{"outside-hidden", "Sheets/s1.xml", `r="3" ht=`, `r="3" hidden="1" ht=`},
		{"outside-size", "Sheets/s1.xml", `ht="409"`, `ht="410"`},
		{"outline", "Sheets/s1.xml", `<outlinePr/>`, `<outlinePr summaryBelow="0"/>`},
		{"header", "Sheets/s1.xml", `<headerFooter alignWithMargins="0"/>`, `<headerFooter alignWithMargins="0"><oddHeader>test</oddHeader></headerFooter>`},
		{"duplicate-header", "Sheets/s1.xml", `<headerFooter alignWithMargins="0"/>`, `<headerFooter alignWithMargins="0"/><headerFooter alignWithMargins="0"/>`},
		{"run-effect", "Sheets/s1.xml", `strike val="0"`, `strike val="1"`},
		{"run-theme", "Sheets/s1.xml", `<b/>`, `<b/><scheme val="minor"/>`},
		{"run-unknown", "Sheets/s1.xml", `<b/>`, `<b/><shadow/>`},
		{"run-duplicate", "Sheets/s1.xml", `<b/>`, `<b/><b/>`},
		{"run-phonetic", "Sheets/s1.xml", `</is></c><c r="B1"`, `<rPh sb="0" eb="1"><t>x</t></rPh></is></c><c r="B1"`},
		{"cell-metadata", "Sheets/s1.xml", `r="A1" s=`, `r="A1" cm="1" s=`},
		{"unknown-sheet", "Sheets/s1.xml", `</worksheet>`, `<extLst/></worksheet>`},
	} {
		t.Run(tc.name, func(t *testing.T) {
			p := richSourceFixture()
			if !strings.Contains(p[tc.part], tc.old) {
				t.Fatal("fixture edit missing")
			}
			p[tc.part] = strings.Replace(p[tc.part], tc.old, tc.new, 1)
			v, e := PreviewNativeRichSourceV1(buildZip(t, p))
			if e == nil || v != nil {
				t.Fatalf("accepted %s: %#v", tc.name, v)
			}
		})
	}
}

func TestRichSourceObservedParentCount(t *testing.T) {
	for _, n := range []int{0, 2, 128, 129} {
		t.Run(fmt.Sprint(n), func(t *testing.T) {
			p := richSourceFixture()
			original := `<cellStyleXfs><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>`
			replacement := `<cellStyleXfs>` + strings.Repeat(`<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>`, n) + `</cellStyleXfs>`
			p["Meta/Styles.style"] = strings.Replace(p["Meta/Styles.style"], original, replacement, 1)
			if _, e := parseStyleTable([]byte(p["Meta/Styles.style"])); e == nil {
				t.Fatal("strict count parser changed")
			}
			v, e := PreviewNativeRichSourceV1(buildZip(t, p))
			if n == 0 || n > 128 {
				if e == nil || v != nil {
					t.Fatal("unbounded count accepted")
				}
			} else if e != nil || v.ParentCount.Observed != n {
				t.Fatalf("observed count %d: %v", n, e)
			}
		})
	}
}
