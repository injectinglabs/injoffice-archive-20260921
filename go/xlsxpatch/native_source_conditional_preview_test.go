package xlsxpatch

import (
	"bytes"
	"encoding/json"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func sourceConditionalFixture() map[string]string {
	p := sourceStyleFixture()
	p["Meta/Styles.style"] = strings.Replace(p["Meta/Styles.style"], `</styleSheet>`, `<dxfs count="1"><dxf><font><b val="1"/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs></styleSheet>`, 1)
	s := p["Sheets/s1.xml"]
	s = strings.Replace(s, `<c r="A2" s="1"><f aca="false">SUM(1,2)</f><v>3</v></c><c r="B2"><v>4.5</v></c>`, `<c r="A2"><v>327</v></c><c r="B2" t="str"><f>IF(A2&lt;400,"REORDER","OK")</f><v>REORDER</v></c>`, 1)
	s = strings.Replace(s, `</worksheet>`, `<conditionalFormatting sqref="B2"><cfRule type="cellIs" priority="2" operator="equal" aboveAverage="0" equalAverage="0" bottom="0" percent="0" rank="0" text="" dxfId="0"><formula>"REORDER"</formula></cfRule></conditionalFormatting><conditionalFormatting sqref="A2"><cfRule type="dataBar" priority="3"><dataBar showValue="1" minLength="10" maxLength="90"><cfvo type="num" val="0"/><cfvo type="num" val="500"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"><x14:id xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">{5F90FECE-9E24-4951-B2E0-4C33D355D6E9}</x14:id></ext></extLst></cfRule></conditionalFormatting><extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:conditionalFormatting><x14:cfRule type="dataBar" id="{5F90FECE-9E24-4951-B2E0-4C33D355D6E9}"><x14:dataBar minLength="10" maxLength="90" axisPosition="none" gradient="true"><x14:cfvo type="num"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="num"><xm:f>500</xm:f></x14:cfvo><x14:negativeFillColor rgb="FF638EC6"/><x14:axisColor rgb="FF000000"/></x14:dataBar></x14:cfRule><xm:sqref>A2</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst></worksheet>`, 1)
	p["Sheets/s1.xml"] = s
	return p
}
func TestSourceConditionalPreview(t *testing.T) {
	p := sourceConditionalFixture()
	data := buildZip(t, p)
	before := bytes.Clone(data)
	got, err := PreviewNativeSourceStylesV2(data)
	if err != nil {
		t.Fatal(err)
	}
	if got.Version != 2 || !got.ReadOnly || got.Fidelity != "approximate" || got.Grid.PackageSHA256 != nativeWorkbookDigest(data) || got.StylesSHA256 != nativeWorkbookDigest([]byte(p["Meta/Styles.style"])) || got.WorksheetSHA256 != nativeWorkbookDigest([]byte(p["Sheets/s1.xml"])) {
		t.Fatalf("lost source identity: %+v", got)
	}
	r, b := got.TextRule, got.DataBar
	if r == nil || !r.Bold || r.FontColor != "#9C0006" || r.BackgroundColor != "#FFC7CE" || len(r.Cells) != 1 || !r.Cells[0].Cached || !r.Cells[0].Matched || r.Cells[0].Ref != "B2" || r.Cells[0].Value != "REORDER" {
		t.Fatalf("lost differential evidence: %+v", r)
	}
	if b == nil || b.Minimum != 0 || b.Maximum != 500 || !b.Gradient || b.Color != "#638EC6" || len(b.Cells) != 1 || b.Cells[0].Ref != "A2" || b.Cells[0].Value != 327 || b.Cells[0].LengthPercent != 62.32 {
		t.Fatalf("lost data bar: %+v", b)
	}
	if !bytes.Equal(before, data) {
		t.Fatal("changed source")
	}
	if got, err := PreviewNativeSourceStylesV1(data); got != nil || err == nil {
		t.Fatal("V1 authority expanded")
	}
	if _, err := ExtractNativeWorkbookV2(data); err == nil {
		t.Fatal("strict extraction authority expanded")
	}
	tx := NativeWorkbookMutationTransactionV1{ExpectedRevision: "rev:" + strings.TrimPrefix(nativeWorkbookDigest(data), "sha256:"), Cells: []CellMutation{{OperationID: "x", SheetID: "7", Kind: CellSetValue, Cell: CellRef{Row: 1, Column: 1}, Value: "7"}}}
	if _, err := ApplyNativeWorkbookMutationTransactionV1(data, tx); err == nil {
		t.Fatal("mutation authority expanded")
	}
	if dir := os.Getenv("XLSX_SOURCE_CONDITIONAL_EVIDENCE_DIR"); dir != "" {
		encoded, err := json.Marshal(got)
		if err != nil {
			t.Fatal(err)
		}
		if err = os.MkdirAll(dir, 0755); err != nil {
			t.Fatal(err)
		}
		for name, b := range map[string][]byte{"source.xlsx": data, "preview.json": encoded} {
			if err = os.WriteFile(filepath.Join(dir, name), b, 0644); err != nil {
				t.Fatal(err)
			}
		}
	}
}
func TestSourceConditionalRefusals(t *testing.T) {
	for _, tc := range []struct{ name, part, old, new string }{
		{"orphan-link", "Sheets/s1.xml", `id="{5F90FECE-9E24-4951-B2E0-4C33D355D6E9}"`, `id="{5F90FECE-9E24-4951-B2E0-4C33D355D6E8}"`},
		{"unknown-rule-attr", "Sheets/s1.xml", `priority="2"`, `priority="2" stopIfTrue="1"`},
		{"competing-priority", "Sheets/s1.xml", `priority="3"`, `priority="2"`},
		{"overlap", "Sheets/s1.xml", `sqref="B2"`, `sqref="A2"`},
		{"merged-range", "Sheets/s1.xml", `sqref="B2"`, `sqref="B1"`},
		{"range-holes", "Sheets/s1.xml", `sqref="B2"`, `sqref="B2:B3"`},
		{"rule-expression", "Sheets/s1.xml", `<formula>"REORDER"</formula>`, `<formula>LOWER("REORDER")</formula>`},
		{"literal-case", "Sheets/s1.xml", `<formula>"REORDER"</formula>`, `<formula>"Reorder"</formula>`},
		{"cache-case", "Sheets/s1.xml", `<v>REORDER</v>`, `<v>Reorder</v>`},
		{"cache-missing", "Sheets/s1.xml", `<v>REORDER</v>`, ``},
		{"cache-number", "Sheets/s1.xml", `t="str"`, `t="n"`},
		{"bar-formula", "Sheets/s1.xml", `<v>327</v>`, `<f>300+27</f><v>327</v>`},
		{"bar-negative", "Sheets/s1.xml", `<v>327</v>`, `<v>-1</v>`},
		{"bar-over-max", "Sheets/s1.xml", `<v>327</v>`, `<v>501</v>`},
		{"bar-decimal", "Sheets/s1.xml", `<v>327</v>`, `<v>327.0</v>`},
		{"bar-background", "Sheets/s1.xml", `<c r="A2">`, `<c r="A2" s="1">`},
		{"bar-dynamic", "Sheets/s1.xml", `type="num" val="0"`, `type="min" val="0"`},
		{"x14-bound-conflict", "Sheets/s1.xml", `<xm:f>500</xm:f>`, `<xm:f>501</xm:f>`},
		{"x14-axis", "Sheets/s1.xml", `axisPosition="none"`, `axisPosition="automatic"`},
		{"x14-range", "Sheets/s1.xml", `<xm:sqref>A2</xm:sqref>`, `<xm:sqref>A3</xm:sqref>`},
		{"x14-priority", "Sheets/s1.xml", `<x14:cfRule type="dataBar"`, `<x14:cfRule priority="3" type="dataBar"`},
		{"foreign-x14", "Sheets/s1.xml", `xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"`, `xmlns:x14="urn:foreign"`},
		{"nested-x14", "Sheets/s1.xml", `<xm:f>500</xm:f>`, `<xm:f><unknown/>500</xm:f>`},
		{"dxf-metadata", "Meta/Styles.style", `<dxf>`, `<dxf unknown="1">`},
		{"dxf-background-form", "Meta/Styles.style", `<patternFill><bgColor`, `<patternFill patternType="solid"><bgColor`},
		{"dxf-alpha", "Meta/Styles.style", `FFFFC7CE`, `00FFC7CE`},
		{"dxf-count", "Meta/Styles.style", `<dxfs count="1">`, `<dxfs count="2">`},
		{"dxf-bold-off", "Meta/Styles.style", `<b val="1"/>`, `<b val="0"/>`},
		{"base-unknown", "Sheets/s1.xml", `</worksheet>`, `<unknown/></worksheet>`},
	} {
		t.Run(tc.name, func(t *testing.T) {
			p := sourceConditionalFixture()
			if !strings.Contains(p[tc.part], tc.old) {
				t.Fatal("bad fixture replacement")
			}
			p[tc.part] = strings.Replace(p[tc.part], tc.old, tc.new, 1)
			if got, err := PreviewNativeSourceStylesV2(buildZip(t, p)); got != nil || err == nil {
				t.Fatalf("accepted unsafe input: %+v", got)
			}
		})
	}
}
func TestSourceConditionalStoredInputs(t *testing.T) {
	for _, tc := range []struct {
		value  string
		length float64
	}{{"0", 10}, {"500", 90}, {"250", 50}} {
		p := sourceConditionalFixture()
		p["Sheets/s1.xml"] = strings.Replace(p["Sheets/s1.xml"], `<v>327</v>`, `<v>`+tc.value+`</v>`, 1)
		// A deliberately inconsistent formula cache proves the preview never evaluates IF.
		p["Sheets/s1.xml"] = strings.Replace(p["Sheets/s1.xml"], `<v>REORDER</v>`, `<v>OK</v>`, 1)
		got, err := PreviewNativeSourceStylesV2(buildZip(t, p))
		if err != nil {
			t.Fatal(err)
		}
		if got.DataBar.Cells[0].LengthPercent != tc.length || got.TextRule.Cells[0].Matched || got.TextRule.Cells[0].Value != "OK" || !got.TextRule.Cells[0].Cached {
			t.Fatal("lost stored-value semantics")
		}
	}
}

func TestSourceConditionalFrozenView(t *testing.T) {
	view := `<sheetViews><sheetView workbookViewId="0"><pane xSplit="0" ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="topLeft" activeCell="A1" activeCellId="0" sqref="A1"/><selection pane="bottomLeft" activeCell="A1" activeCellId="0" sqref="A1"/></sheetView></sheetViews>`
	for _, tc := range []struct {
		name, old, new string
		good           bool
	}{
		{name: "qualified", good: true},
		{"split-columns", `xSplit="0"`, `xSplit="1"`, false},
		{"wrong-origin", `topLeftCell="A2"`, `topLeftCell="B2"`, false},
		{"nonempty-pane", `state="frozen"/>`, `state="frozen"><unknown/></pane>`, false},
		{"unknown-selection", `pane="bottomLeft" activeCell="A1"`, `pane="bottomLeft" activeCell="B1"`, false},
		{"unknown-view", `workbookViewId="0"`, `workbookViewId="0" unknown="1"`, false},
		{"outside-grid", `ySplit="1" topLeftCell="A2"`, `ySplit="2" topLeftCell="A3"`, false},
	} {
		t.Run(tc.name, func(t *testing.T) {
			p := sourceConditionalFixture()
			v := view
			if tc.old != "" {
				v = strings.Replace(v, tc.old, tc.new, 1)
			}
			p["Sheets/s1.xml"] = strings.Replace(p["Sheets/s1.xml"], `<sheetFormatPr`, v+`<sheetFormatPr`, 1)
			got, err := PreviewNativeSourceStylesV2(buildZip(t, p))
			if tc.good {
				if err != nil {
					t.Fatal(err)
				}
				if got.FrozenView == nil || got.FrozenView.FrozenRows != 1 || got.FrozenView.TopLeftCell != "A2" || len(got.Warnings) != 3 {
					t.Fatal("lost pane evidence")
				}
			} else if got != nil || err == nil {
				t.Fatal("accepted unqualified pane")
			}
		})
	}
}
func TestSourceConditionalSingleFamilies(t *testing.T) {
	remove := func(s, start, end string) string {
		a := strings.Index(s, start)
		b := strings.Index(s[a:], end) + a + len(end)
		return s[:a] + s[b:]
	}
	for _, family := range []string{"text", "bar"} {
		t.Run(family, func(t *testing.T) {
			p := sourceConditionalFixture()
			s := p["Sheets/s1.xml"]
			if family == "text" {
				s = remove(s, `<conditionalFormatting sqref="A2">`, `</conditionalFormatting>`)
				s = remove(s, `<extLst><ext uri="{78C0`, `</extLst>`)
			} else {
				s = remove(s, `<conditionalFormatting sqref="B2">`, `</conditionalFormatting>`)
				p["Meta/Styles.style"] = remove(p["Meta/Styles.style"], `<dxfs`, `</dxfs>`)
			}
			p["Sheets/s1.xml"] = s
			got, err := PreviewNativeSourceStylesV2(buildZip(t, p))
			if err != nil {
				t.Fatal(err)
			}
			if (got.TextRule != nil) != (family == "text") || (got.DataBar != nil) != (family == "bar") {
				t.Fatal("wrong family")
			}
		})
	}
	if got, err := PreviewNativeSourceStylesV2(buildZip(t, sourceStyleFixture())); got != nil || err == nil {
		t.Fatal("V2 accepted a source without conditional declarations")
	}
}
