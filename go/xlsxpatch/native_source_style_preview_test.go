package xlsxpatch

import (
	"bytes"
	"encoding/json"
	"strings"
	"testing"
)

func sourceStyleFixture() map[string]string {
	p := nativeMutationFixture(false)
	p["Meta/Styles.style"] = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00&quot; €&quot;"/></numFmts><fonts count="1"><font><name val="Arial"/><sz val="10"/><color rgb="FF000000"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1E2761"/><bgColor rgb="FF333333"/></patternFill></fill></fills><borders count="1"><border diagonalUp="false" diagonalDown="false"><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="1" borderId="0" xfId="0"/></cellXfs></styleSheet>`
	p["Sheets/s1.xml"] = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="2" width="12" customWidth="true"/></cols><sheetData><row r="1" ht="25.5" customHeight="true"><c r="A1" t="inlineStr" s="1"><is><t>Expenses</t></is></c><c r="B1" s="1"/></row><row r="2"><c r="A2" s="1"><f aca="false">SUM(1,2)</f><v>3</v></c><c r="B2"><v>4.5</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells></worksheet>`
	// The generic fixture contains another hidden sheet; this profile is one-sheet.
	p["Book/Workbook.xml"] = strings.Replace(p["Book/Workbook.xml"], `<sheet name="Hidden" sheetId="9" state="hidden" r:id="rSheet2"/>`, "", 1)
	return p
}
func TestSourceStylePreviewKeepsStrictAuthority(t *testing.T) {
	p := sourceStyleFixture()
	b := buildZip(t, p)
	before := bytes.Clone(b)
	if _, err := ExtractNativeWorkbookV2(b); err == nil {
		t.Fatal("strict extraction accepted conflicting source")
	}
	got, err := PreviewNativeSourceStylesV1(b)
	if err != nil {
		t.Fatal(err)
	}
	if got.Protocol != "injoffice.xlsx.source-style-preview" || !got.ReadOnly || got.Fidelity != "approximate" || got.PackageSHA256 != nativeWorkbookDigest(b) || len(got.Conflicts) != 2 || got.StrictError == "" {
		t.Fatalf("source authority lost: %#v", got)
	}
	if len(got.Sheets) != 1 || len(got.Sheets[0].Merges) != 1 || len(got.Sheets[0].Cells) != 4 || got.Sheets[0].Cells[2].Lexical != "3" || !got.Sheets[0].Cells[2].Cached || got.Styles[1].FillColor != "#1E2761" || got.Styles[1].NumberFormatID != 164 {
		t.Fatalf("source content/styles lost: %#v", got)
	}
	encoded, err := json.Marshal(got)
	if err != nil {
		t.Fatal(err)
	}
	if bytes.Contains(encoded, []byte(`"editable"`)) || bytes.Contains(encoded, []byte(`"revision"`)) {
		t.Fatal("native mutation authority leaked")
	}
	tx := NativeWorkbookMutationTransactionV1{ExpectedRevision: "rev:" + strings.TrimPrefix(nativeWorkbookDigest(b), "sha256:"), Cells: []CellMutation{{OperationID: "x", SheetID: "7", Kind: CellSetValue, Cell: CellRef{Row: 1, Column: 1}, Value: "7"}}}
	if _, err = ApplyNativeWorkbookMutationTransactionV1(b, tx); err == nil {
		t.Fatal("mutation accepted conflicting source")
	}
	if !bytes.Equal(before, b) {
		t.Fatal("source bytes changed")
	}
}
func TestSourceStylePreviewRefusesUnsafeSources(t *testing.T) {
	for name, edit := range map[string]func(map[string]string){
		"explicit-false-fill": func(p map[string]string) {
			p["Meta/Styles.style"] = strings.Replace(p["Meta/Styles.style"], `numFmtId="164" fontId="0" fillId="1"`, `applyFill="false" numFmtId="164" fontId="0" fillId="1"`, 1)
		},
		"explicit-false-format": func(p map[string]string) {
			p["Meta/Styles.style"] = strings.Replace(p["Meta/Styles.style"], `numFmtId="164" fontId="0" fillId="1"`, `applyNumberFormat="0" numFmtId="164" fontId="0" fillId="1"`, 1)
		},
		"number-format-metadata": func(p map[string]string) {
			p["Meta/Styles.style"] = strings.Replace(p["Meta/Styles.style"], `<numFmt numFmtId="164"`, `<numFmt unknown="1" numFmtId="164"`, 1)
		},
		"invalid-id": func(p map[string]string) {
			p["Meta/Styles.style"] = strings.Replace(p["Meta/Styles.style"], `fillId="1"`, `fillId="99"`, 1)
		},
		"foreign-style": func(p map[string]string) {
			p["Meta/Styles.style"] = strings.Replace(p["Meta/Styles.style"], `<sz val="10"/>`, `<sz xmlns="urn:foreign" val="10"/>`, 1)
		},
		"nonempty-font": func(p map[string]string) {
			p["Meta/Styles.style"] = strings.Replace(p["Meta/Styles.style"], `<name val="Arial"/>`, `<name val="Arial">unexpected</name>`, 1)
		},
		"nonempty-diagonal": func(p map[string]string) {
			p["Meta/Styles.style"] = strings.Replace(p["Meta/Styles.style"], `<diagonal/>`, `<diagonal>unexpected</diagonal>`, 1)
		},
		"strike": func(p map[string]string) {
			p["Meta/Styles.style"] = strings.Replace(p["Meta/Styles.style"], `</font>`, `<strike/></font>`, 1)
		},
		"hidden-row": func(p map[string]string) {
			p["Sheets/s1.xml"] = strings.Replace(p["Sheets/s1.xml"], `<row r="2">`, `<row r="2" hidden="true">`, 1)
		},
		"formula-group": func(p map[string]string) {
			p["Sheets/s1.xml"] = strings.Replace(p["Sheets/s1.xml"], `<f aca="false">`, `<f t="array" ref="A2:A3">`, 1)
		},
		"metadata": func(p map[string]string) {
			p["Sheets/s1.xml"] = strings.Replace(p["Sheets/s1.xml"], `r="A2"`, `r="A2" cm="1"`, 1)
		},
		"covered-content": func(p map[string]string) {
			p["Sheets/s1.xml"] = strings.Replace(p["Sheets/s1.xml"], `<c r="B1" s="1"/>`, `<c r="B1" s="1"><v>1</v></c>`, 1)
		},
		"drawing": func(p map[string]string) {
			p["Sheets/s1.xml"] = strings.Replace(p["Sheets/s1.xml"], `</worksheet>`, `<drawing/></worksheet>`, 1)
		},
		"foreign-cell": func(p map[string]string) {
			p["Sheets/s1.xml"] = strings.Replace(p["Sheets/s1.xml"], `<c r="B2">`, `<c xmlns="urn:foreign" r="B2">`, 1)
		},
		"bound": func(p map[string]string) {
			p["Sheets/s1.xml"] = strings.Replace(p["Sheets/s1.xml"], `r="B2"`, `r="AG2"`, 1)
		},
		"duplicate-attribute": func(p map[string]string) {
			p["Meta/Styles.style"] = strings.Replace(p["Meta/Styles.style"], `fillId="1"`, `fillId="1" fillId="0"`, 1)
		},
	} {
		t.Run(name, func(t *testing.T) {
			p := sourceStyleFixture()
			edit(p)
			if got, err := PreviewNativeSourceStylesV1(buildZip(t, p)); err == nil || got != nil {
				t.Fatalf("unsafe source accepted: %#v", got)
			}
		})
	}
}
