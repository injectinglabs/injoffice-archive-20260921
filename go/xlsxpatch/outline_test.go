package xlsxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func outlineFixture(sheet string) map[string]string {
	entries := cellMutationFixture()
	entries["xl/worksheets/sheet1.xml"] = sheet
	return entries
}

func outlineGroup(id string, axis XLSXOutlineAxis, start, end int, collapsed bool) XLSXOutlineGroup {
	return XLSXOutlineGroup{ID: id, SheetID: "7", Axis: axis, Start: start, End: end, Collapsed: collapsed}
}

func TestReadWorksheetOutline_ReconstructsNestedDisjointGroupsAndRawMetadata(t *testing.T) {
	sheet := `<x:worksheet xmlns:x="` + spreadsheetMLTransitional + `"><x:sheetPr><x:outlinePr summaryBelow="1" summaryRight="0"/></x:sheetPr>` +
		`<x:cols><x:col min="1" max="1" collapsed="1"/><x:col min="2" max="4" outlineLevel="1" hidden="1"/><x:col min="7" max="8" outlineLevel="1"/></x:cols>` +
		`<x:sheetData><x:row r="1" outlineLevel="1" hidden="1"/><x:row r="2" outlineLevel="2" hidden="1"/><x:row r="3" outlineLevel="2" hidden="1"/><x:row r="4" outlineLevel="1" hidden="1" collapsed="1"><x:c r="A4"><x:v>keep</x:v></x:c></x:row><x:row r="5" collapsed="1"/><x:row r="7" outlineLevel="1"/><x:row r="8" outlineLevel="1"/></x:sheetData><x:extLst><x:ext uri="keep"/></x:extLst></x:worksheet>`
	snapshot, err := ReadWorksheetOutline(buildZip(t, outlineFixture(sheet)), "7")
	if err != nil {
		t.Fatal(err)
	}
	if !snapshot.SummaryBelow || snapshot.SummaryRight {
		t.Fatalf("summary directions = below:%v right:%v", snapshot.SummaryBelow, snapshot.SummaryRight)
	}
	if len(snapshot.Rows) != 5 {
		t.Fatalf("row metadata bands = %d, want 5: %+v", len(snapshot.Rows), snapshot.Rows)
	}
	if len(snapshot.Columns) != 3 {
		t.Fatalf("column metadata bands = %d, want 3: %+v", len(snapshot.Columns), snapshot.Columns)
	}
	if len(snapshot.Groups) != 5 {
		t.Fatalf("groups = %d, want 5: %+v", len(snapshot.Groups), snapshot.Groups)
	}
	want := map[string]bool{
		"row:0:3:true":     false,
		"row:1:2:true":     false,
		"row:6:7:false":    false,
		"column:1:3:true":  false,
		"column:6:7:false": false,
	}
	for _, group := range snapshot.Groups {
		key := string(group.Axis) + ":" + strconvItoa(group.Start) + ":" + strconvItoa(group.End) + ":" + strconvBool(group.Collapsed)
		if _, ok := want[key]; !ok {
			t.Errorf("unexpected group %+v", group)
			continue
		}
		if want[key] {
			t.Errorf("duplicate group %s", key)
		}
		want[key] = true
		if group.ID == "" || !strings.HasPrefix(group.ID, "xlsx-") {
			t.Errorf("group has non-deterministic-looking id %q", group.ID)
		}
	}
	for key, found := range want {
		if !found {
			t.Errorf("missing group %s", key)
		}
	}
	again, err := ReadWorksheetOutline(buildZip(t, outlineFixture(sheet)), "7")
	if err != nil {
		t.Fatal(err)
	}
	for index := range snapshot.Groups {
		if snapshot.Groups[index].ID != again.Groups[index].ID {
			t.Fatal("outline ids are not deterministic")
		}
	}
}

func TestApplyWorksheetOutline_RoundTripsAndPreservesUnrelatedBytes(t *testing.T) {
	sheet := `<?xml version="1.0"?><x:worksheet xmlns:x="` + spreadsheetMLTransitional + `"><x:sheetPr codeName="Keep"><x:pageSetUpPr fitToPage="1"/></x:sheetPr><x:dimension ref="A1:C10"/><x:cols><x:col xmlns:z="urn:opaque" min="1" max="8" width="11" customWidth="1" style="3" z:keep="yes"/></x:cols><x:sheetData><x:row r="1" custom="keep"><x:c r="A1"><x:v>alpha</x:v></x:c></x:row><x:row r="10" hidden="1"><x:c r="B10"><x:v>manual-hidden</x:v></x:c></x:row></x:sheetData><x:mergeCells count="1"><x:mergeCell ref="A1:B1"/></x:mergeCells><x:extLst><x:ext uri="keep"/></x:extLst></x:worksheet>`
	entries := outlineFixture(sheet)
	original := buildZip(t, entries)
	write := WorksheetOutlineWrite{
		SheetID: "7", SummaryBelow: true, SummaryRight: false,
		Groups: []XLSXOutlineGroup{
			outlineGroup("rows-outer", XLSXOutlineRows, 1, 4, true),
			outlineGroup("rows-inner", XLSXOutlineRows, 2, 3, true),
			outlineGroup("rows-disjoint", XLSXOutlineRows, 7, 8, false),
			outlineGroup("columns", XLSXOutlineColumns, 1, 3, true),
		},
	}
	output, err := ApplyWorksheetOutline(original, write)
	if err != nil {
		t.Fatal(err)
	}
	snapshot, err := ReadWorksheetOutline(output, "7")
	if err != nil {
		t.Fatal(err)
	}
	if !snapshot.SummaryBelow || snapshot.SummaryRight || len(snapshot.Groups) != len(write.Groups) {
		t.Fatalf("unexpected round trip: %+v", snapshot)
	}
	updated := readEntry(t, output, "xl/worksheets/sheet1.xml")
	for _, want := range []string{
		`<x:sheetPr codeName="Keep"><x:outlinePr summaryBelow="1" summaryRight="0"/><x:pageSetUpPr fitToPage="1"/></x:sheetPr>`,
		`<x:c r="A1"><x:v>alpha</x:v></x:c>`,
		`<x:row r="10" hidden="1"><x:c r="B10"><x:v>manual-hidden</x:v></x:c></x:row>`,
		`width="11" customWidth="1" style="3" z:keep="yes"`,
		`<x:mergeCells count="1"><x:mergeCell ref="A1:B1"/></x:mergeCells><x:extLst><x:ext uri="keep"/></x:extLst>`,
	} {
		if !strings.Contains(updated, want) {
			t.Errorf("worksheet lost %q\n%s", want, updated)
		}
	}
	for _, name := range []string{"_rels/.rels", "xl/workbook.xml", "xl/_rels/workbook.xml.rels", "xl/charts/chart1.xml", "customXml/item1.xml"} {
		if got := readEntry(t, output, name); got != entries[name] {
			t.Errorf("untouched OPC part %q changed", name)
		}
	}
}

func TestApplyWorksheetOutline_ClearsPriorOutlineMetadataWithoutUnhidingLevelZeroItems(t *testing.T) {
	sheet := `<worksheet xmlns="` + spreadsheetMLTransitional + `"><sheetPr><outlinePr summaryBelow="1" summaryRight="1"/></sheetPr><cols><col min="1" max="2" outlineLevel="1" hidden="1" collapsed="1"/><col min="4" max="4" hidden="1"/></cols><sheetData><row r="1" outlineLevel="1" hidden="1"/><row r="2" collapsed="1"/><row r="3" hidden="1"/></sheetData></worksheet>`
	original := buildZip(t, outlineFixture(sheet))
	output, err := ApplyWorksheetOutline(original, WorksheetOutlineWrite{SheetID: "7", SummaryBelow: true, SummaryRight: true})
	if err != nil {
		t.Fatal(err)
	}
	updated := readEntry(t, output, "xl/worksheets/sheet1.xml")
	if strings.Contains(updated, "outlineLevel=") || strings.Contains(updated, "collapsed=") {
		t.Fatalf("old outline metadata remains: %s", updated)
	}
	for _, want := range []string{`<row r="3" hidden="1"/>`, `<col min="4" max="4" hidden="1"/>`} {
		if !strings.Contains(updated, want) {
			t.Errorf("manual hidden metadata was not preserved: %q in %s", want, updated)
		}
	}
	if snapshot, err := ReadWorksheetOutline(output, "7"); err != nil || len(snapshot.Groups) != 0 {
		t.Fatalf("cleared snapshot = %+v, %v", snapshot, err)
	}
}

func TestApplyWorksheetOutline_IsByteNoopWhenProjectionMatches(t *testing.T) {
	sheet := `<worksheet xmlns="` + spreadsheetMLTransitional + `"><sheetPr><outlinePr summaryBelow="1" summaryRight="1"/></sheetPr><sheetData><row r="1" outlineLevel="1"/><row r="2" outlineLevel="1"/></sheetData></worksheet>`
	original := buildZip(t, outlineFixture(sheet))
	write := WorksheetOutlineWrite{SheetID: "7", SummaryBelow: true, SummaryRight: true, Groups: []XLSXOutlineGroup{outlineGroup("same", XLSXOutlineRows, 0, 1, false)}}
	output, err := ApplyWorksheetOutline(original, write)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(output, original) {
		t.Fatal("matching outline projection was not a byte-level no-op")
	}
}

func TestApplyWorksheetOutline_WritesSummaryAboveAndRight(t *testing.T) {
	original := buildZip(t, outlineFixture(`<x:worksheet xmlns:x="`+spreadsheetMLTransitional+`"><x:sheetData/></x:worksheet>`))
	write := WorksheetOutlineWrite{
		SheetID: "7", SummaryBelow: false, SummaryRight: true,
		Groups: []XLSXOutlineGroup{
			outlineGroup("rows-above", XLSXOutlineRows, 1, 2, true),
			outlineGroup("columns-right", XLSXOutlineColumns, 0, 1, true),
		},
	}
	output, err := ApplyWorksheetOutline(original, write)
	if err != nil {
		t.Fatal(err)
	}
	sheet := readEntry(t, output, "xl/worksheets/sheet1.xml")
	for _, want := range []string{
		`<x:outlinePr summaryBelow="0" summaryRight="1"/>`,
		`<x:row r="1" collapsed="1"/>`,
		`<x:row r="2" outlineLevel="1" hidden="1"/>`,
		`<x:row r="3" outlineLevel="1" hidden="1"/>`,
		`<x:col min="3" max="3" collapsed="1"/>`,
	} {
		if !strings.Contains(sheet, want) {
			t.Errorf("missing %q in %s", want, sheet)
		}
	}
	snapshot, err := ReadWorksheetOutline(output, "7")
	if err != nil {
		t.Fatal(err)
	}
	if snapshot.SummaryBelow || !snapshot.SummaryRight || len(snapshot.Groups) != 2 {
		t.Fatalf("unexpected reopened summary projection: %+v", snapshot)
	}
}

func TestReadWorksheetOutline_RefusesMalformedOrAmbiguousMetadata(t *testing.T) {
	tests := map[string]string{
		"level overflow":                   `<sheetData><row r="1" outlineLevel="8"/></sheetData>`,
		"noncanonical level":               `<sheetData><row r="1" outlineLevel="+1"/></sheetData>`,
		"invalid hidden":                   `<sheetData><row r="1" outlineLevel="1" hidden="yes"/></sheetData>`,
		"duplicate ranges from depth jump": `<sheetData><row r="1" outlineLevel="2"/></sheetData>`,
		"visible collapsed detail":         `<sheetData><row r="1" outlineLevel="1"/><row r="2" collapsed="1"/></sheetData>`,
		"orphan collapsed":                 `<sheetData><row r="1" collapsed="1"/></sheetData>`,
		"outline properties child":         `<sheetPr><outlinePr><ext/></outlinePr></sheetPr><sheetData/>`,
	}
	for name, body := range tests {
		t.Run(name, func(t *testing.T) {
			sheet := `<worksheet xmlns="` + spreadsheetMLTransitional + `">` + body + `</worksheet>`
			if snapshot, err := ReadWorksheetOutline(buildZip(t, outlineFixture(sheet)), "7"); err == nil {
				t.Fatalf("expected refusal, got %+v", snapshot)
			}
		})
	}
}

func TestApplyWorksheetOutline_RefusesInvalidSnapshotsAtomically(t *testing.T) {
	original := buildZip(t, outlineFixture(`<worksheet xmlns="`+spreadsheetMLTransitional+`"><sheetData/></worksheet>`))
	valid := outlineGroup("valid", XLSXOutlineRows, 1, 4, false)
	cases := map[string]WorksheetOutlineWrite{
		"bad sheet":              {SheetID: " missing", Groups: []XLSXOutlineGroup{valid}},
		"bad id":                 {SheetID: "7", Groups: []XLSXOutlineGroup{func() XLSXOutlineGroup { group := valid; group.ID = "bad id"; return group }()}},
		"sheet mismatch":         {SheetID: "7", Groups: []XLSXOutlineGroup{func() XLSXOutlineGroup { group := valid; group.SheetID = "9"; return group }()}},
		"range overflow":         {SheetID: "7", Groups: []XLSXOutlineGroup{outlineGroup("overflow", XLSXOutlineColumns, 0, excelMaxColumns, false)}},
		"crossing":               {SheetID: "7", Groups: []XLSXOutlineGroup{outlineGroup("a", XLSXOutlineRows, 1, 5, false), outlineGroup("b", XLSXOutlineRows, 3, 7, false)}},
		"duplicate range":        {SheetID: "7", Groups: []XLSXOutlineGroup{outlineGroup("a", XLSXOutlineRows, 1, 5, false), outlineGroup("b", XLSXOutlineRows, 1, 5, false)}},
		"collapsed before first": {SheetID: "7", SummaryBelow: false, Groups: []XLSXOutlineGroup{outlineGroup("first", XLSXOutlineRows, 0, 2, true)}},
	}
	deep := WorksheetOutlineWrite{SheetID: "7"}
	for level := 0; level < XLSXMaxOutlineDepth+1; level++ {
		deep.Groups = append(deep.Groups, outlineGroup("depth-"+strconvItoa(level), XLSXOutlineRows, level, 20-level, false))
	}
	cases["depth eight"] = deep
	for name, write := range cases {
		t.Run(name, func(t *testing.T) {
			output, err := ApplyWorksheetOutline(original, write)
			if err == nil || output != nil {
				t.Fatalf("expected atomic refusal, output=%d err=%v", len(output), err)
			}
		})
	}
}

func strconvItoa(value int) string {
	if value == 0 {
		return "0"
	}
	var digits [20]byte
	position := len(digits)
	for value > 0 {
		position--
		digits[position] = byte('0' + value%10)
		value /= 10
	}
	return string(digits[position:])
}

func strconvBool(value bool) string {
	if value {
		return "true"
	}
	return "false"
}
