package xlsxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func tableNumberFixture() map[string]string {
	parts := nativeWorkbookFixture(false)
	parts["Charts/chart1.xml"] = previewChartFixture()
	parts["Tables/table.xml"] = `<table xmlns="` + spreadsheetMLTransitional + `" displayName="Original" ref="A1:C4" totalsRowCount="1"><tableColumns count="3"><tableColumn id="1" name="Label"/><tableColumn id="2" name="Count"/><tableColumn id="3" name="Amount" totalsRowDxfId="0"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>`
	parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/Tables/table.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>`, 1)
	parts["Sheets/s1.xml"] = `<worksheet xmlns="` + spreadsheetMLTransitional + `" xmlns:r="` + officeRelNamespaceTransitional + `"><sheetData><row r="4"><c r="C4"><f>AVERAGE(C2:C3)</f><v>2.75</v></c></row></sheetData><tableParts count="1"><tablePart r:id="table1"/></tableParts></worksheet>`
	parts["Sheets/_rels/s1.xml.rels"] = strings.Replace(parts["Sheets/_rels/s1.xml.rels"], `</Relationships>`, `<Relationship Id="table1" Type="`+officeRelNamespaceTransitional+`/table" Target="../Tables/table.xml"/></Relationships>`, 1)
	parts["Meta/Styles.style"] = strings.Replace(parts["Meta/Styles.style"], `</styleSheet>`, `<dxfs count="1"><dxf><numFmt numFmtId="164" formatCode="0.00&amp;quot; X&amp;quot;"/></dxf></dxfs></styleSheet>`, 1)
	parts["Meta/Styles.style"] = strings.ReplaceAll(parts["Meta/Styles.style"], "&amp;quot;", "&quot;")
	return parts
}
func TestTableDifferentialFormatsRetainSavedValueAndQualifiedStyles(t *testing.T) {
	for _, explicit := range []bool{false, true} {
		parts := tableNumberFixture()
		if explicit {
			parts["Meta/Styles.style"] = strings.Replace(parts["Meta/Styles.style"], `xfId="0"/>`, `xfId="0" applyNumberFormat="1"/>`, 1)
		}
		source := buildZip(t, parts)
		before := bytes.Clone(source)
		got, err := InspectNativeWorkbookObjectsV1(source)
		if err != nil {
			t.Fatal(err)
		}
		if !bytes.Equal(source, before) {
			t.Fatal("inspection mutated source")
		}
		formats := got.Tables[0].NumberFormats
		if len(formats) != 1 || formats[0].Ref != "C4:C4" || formats[0].NumberFormat != `0.00" X"` {
			t.Fatalf("bad source DXF: %+v", formats)
		}
		if (len(formats[0].StyleIDs) > 0) == explicit {
			t.Fatalf("explicit General override lost: %+v", formats[0])
		}
	}
}
func TestTableDifferentialFormatIndicesAndRegions(t *testing.T) {
	for _, change := range []string{"index", "columns", "foreign", "duplicate", "format-id"} {
		parts := tableNumberFixture()
		switch change {
		case "index":
			parts["Tables/table.xml"] = strings.Replace(parts["Tables/table.xml"], `totalsRowDxfId="0"`, `totalsRowDxfId="9"`, 1)
		case "columns":
			parts["Tables/table.xml"] = strings.Replace(parts["Tables/table.xml"], `ref="A1:C4"`, `ref="A1:D4"`, 1)
		case "foreign":
			parts["Meta/Styles.style"] = strings.Replace(parts["Meta/Styles.style"], `<dxf>`, `<dxf xmlns="urn:foreign">`, 1)
		case "duplicate":
			parts["Meta/Styles.style"] = strings.Replace(parts["Meta/Styles.style"], `</dxf>`, `<numFmt numFmtId="165" formatCode="0.0"/></dxf>`, 1)
		case "format-id":
			parts["Meta/Styles.style"] = strings.Replace(parts["Meta/Styles.style"], `numFmtId="164"`, `numFmtId="oops"`, 1)
		}
		if _, err := InspectNativeWorkbookObjectsV1(buildZip(t, parts)); err == nil {
			t.Fatalf("accepted %s", change)
		}
	}
}

func TestTableDifferentialFormatProjectionBudgets(t *testing.T) {
	table := NativeTablePreviewV1{NumberFormats: []NativeTableNumberFormatV1{{StyleIDs: make([]int, 4096)}}, BorderPreview: &NativeTableBorderPreviewV1{StyleIDs: make([]int, 4096)}}
	if !nativeTableStyleIDsWithinBudget([]NativeTablePreviewV1{table, table}) || nativeTableStyleIDsWithinBudget([]NativeTablePreviewV1{table, table, table}) {
		t.Fatal("DXF and border styles must share cumulative style budget")
	}
	table = NativeTablePreviewV1{NumberFormats: make([]NativeTableNumberFormatV1, 1024)}
	if !nativeTableStyleIDsWithinBudget([]NativeTablePreviewV1{table}) || nativeTableStyleIDsWithinBudget([]NativeTablePreviewV1{table, {NumberFormats: []NativeTableNumberFormatV1{{}}}}) {
		t.Fatal("format regions must be capped at 1024")
	}
}
