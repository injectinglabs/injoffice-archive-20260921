package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"io"
	"strings"
	"testing"
)

func sparklineFixture(t *testing.T, sheetXML string) []byte {
	t.Helper()
	entries := fixtureWorkbook(false)
	entries["xl/worksheets/sheet1.xml"] = sheetXML
	entries["xl/media/untouched.bin"] = "untouched-native-payload"
	return buildZip(t, entries)
}

func TestSetAndReadSparklines_AllKindsGroupsAndRoundTrip(t *testing.T) {
	orig := sparklineFixture(t, `<worksheet><sheetData/><extLst><ext uri="keep-me"><future:feature xmlns:future="urn:future">same</future:feature></ext></extLst></worksheet>`)
	min, max := -5.0, 20.0
	groups := []SparklineWriteGroup{
		{ID: "grouped-lines", TargetSheetName: "Data", Type: "line", Options: SparklineOptions{EmptyCells: "connect", ShowMarkers: true, ShowHigh: true, LineWeight: 1.5, Min: &min, Max: &max, Colors: SparklineColors{Series: "#112233", High: "#00AA00"}}, Sparklines: []SparklineWriteItem{
			{SourceSheetName: "Data", SourceRef: "A1:E1", TargetCellRef: "F1"},
			{SourceSheetName: "Data", SourceRef: "A2:E2", TargetCellRef: "F2"},
		}},
		{ID: "column", TargetSheetName: "Data", Type: "column", Options: SparklineOptions{EmptyCells: "zero", ShowNegative: true}, Sparklines: []SparklineWriteItem{{SourceSheetName: "Data", SourceRef: "A3:E3", TargetCellRef: "F3"}}},
		{ID: "wins", TargetSheetName: "Data", Type: "win-loss", Options: SparklineOptions{RightToLeft: true}, Sparklines: []SparklineWriteItem{{SourceSheetName: "Data", SourceRef: "A4:E4", TargetCellRef: "F4"}}},
	}
	out, err := SetSparklines(orig, groups)
	if err != nil {
		t.Fatal(err)
	}
	worksheet := readEntry(t, out, "xl/worksheets/sheet1.xml")
	for _, want := range []string{`uri="keep-me"`, `<future:feature xmlns:future="urn:future">same</future:feature>`, `type="line"`, `type="column"`, `type="stacked"`, `<xm:sqref>F4</xm:sqref>`} {
		if !strings.Contains(worksheet, want) {
			t.Errorf("worksheet missing %s", want)
		}
	}
	if got := readEntry(t, out, "xl/media/untouched.bin"); got != "untouched-native-payload" {
		t.Fatalf("untouched part changed: %q", got)
	}
	infos, err := ReadSparklines(out)
	if err != nil {
		t.Fatal(err)
	}
	if len(infos) != 4 {
		t.Fatalf("got %d sparklines, want 4: %#v", len(infos), infos)
	}
	if infos[0].Type != "line" || infos[1].Type != "line" || infos[2].Type != "column" || infos[3].Type != "win-loss" {
		t.Fatalf("types did not round trip: %#v", infos)
	}
	if infos[0].GroupID == "" || infos[0].GroupID != infos[1].GroupID {
		t.Fatalf("group composition did not round trip: %#v", infos[:2])
	}
	if infos[0].Options.EmptyCells != "connect" || !infos[0].Options.ShowMarkers || infos[0].Options.Min == nil || *infos[0].Options.Min != min || infos[0].Options.Colors.Series != "#112233" {
		t.Fatalf("options did not round trip: %#v", infos[0].Options)
	}
	if len(infos[2].Warnings) != 0 || len(infos[3].Warnings) != 0 {
		t.Fatalf("unexpected diagnostics: %#v", infos)
	}
	second, err := SetSparklines(out, groups)
	if err != nil {
		t.Fatal(err)
	}
	secondInfos, err := ReadSparklines(second)
	if err != nil || len(secondInfos) != 4 {
		t.Fatalf("second round trip: %v %#v", err, secondInfos)
	}
}

func TestReadSparklines_UnsupportedStateHasDiagnosticsAndBlocksWrite(t *testing.T) {
	xml := `<worksheet><sheetData/><extLst><ext uri="` + sparklineExtensionURI + `" xmlns:x14="` + x14Namespace + `" xmlns:xm="` + xmNamespace + `"><x14:sparklineGroups><x14:sparklineGroup type="line" dateAxis="1"><x14:colorSeries theme="4"/><x14:sparklines><x14:sparkline><xm:f>[Other.xlsx]Data!A1:A3</xm:f><xm:sqref>D1:D2</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst></worksheet>`
	orig := sparklineFixture(t, xml)
	infos, err := ReadSparklines(orig)
	if err != nil {
		t.Fatal(err)
	}
	if len(infos) != 1 || len(infos[0].Warnings) < 4 {
		t.Fatalf("expected explicit diagnostics, got %#v", infos)
	}
	_, err = SetSparklines(orig, []SparklineWriteGroup{{TargetSheetName: "Data", Type: "line", Sparklines: []SparklineWriteItem{{SourceSheetName: "Data", SourceRef: "A1:A3", TargetCellRef: "D1"}}}})
	if err == nil || !strings.Contains(err.Error(), "refused") {
		t.Fatalf("unsupported source must block destructive replacement, got %v", err)
	}
}

func TestReadSparklines_MalformedXMLFailsClosed(t *testing.T) {
	orig := sparklineFixture(t, `<worksheet><sheetData/><extLst><ext uri="`+sparklineExtensionURI+`"><x14:sparklineGroups>`)
	if _, err := ReadSparklines(orig); err == nil {
		t.Fatal("malformed XML must fail")
	}
	emptyGroup := sparklineFixture(t, `<worksheet><sheetData/><extLst><ext uri="`+sparklineExtensionURI+`"><sparklineGroups><sparklineGroup><sparklines/></sparklineGroup></sparklineGroups></ext></extLst></worksheet>`)
	if _, err := ReadSparklines(emptyGroup); err == nil {
		t.Fatal("empty native group must fail")
	}
}

func TestSetSparklines_ValidationAndRemovalPreserveOtherExtensions(t *testing.T) {
	orig := sparklineFixture(t, `<worksheet><sheetData/><extLst><ext uri="keep-me"><value>same</value></ext></extLst></worksheet>`)
	cases := []SparklineWriteGroup{
		{TargetSheetName: "Data", Type: "heatmap", Sparklines: []SparklineWriteItem{{SourceSheetName: "Data", SourceRef: "A1:A3", TargetCellRef: "D1"}}},
		{TargetSheetName: "Data", Type: "line", Sparklines: []SparklineWriteItem{{SourceSheetName: "Data", SourceRef: "A1:B2", TargetCellRef: "D1"}}},
		{TargetSheetName: "Data", Type: "line", Options: SparklineOptions{Colors: SparklineColors{Series: "red"}}, Sparklines: []SparklineWriteItem{{SourceSheetName: "Data", SourceRef: "A1:A3", TargetCellRef: "D1"}}},
	}
	for _, bad := range cases {
		if _, err := SetSparklines(orig, []SparklineWriteGroup{bad}); err == nil {
			t.Fatalf("expected validation error for %#v", bad)
		}
	}
	one, err := SetSparklines(orig, []SparklineWriteGroup{{TargetSheetName: "Data", Type: "line", Sparklines: []SparklineWriteItem{{SourceSheetName: "Data", SourceRef: "A1:A3", TargetCellRef: "D1"}}}})
	if err != nil {
		t.Fatal(err)
	}
	cleared, err := SetSparklines(one, nil)
	if err != nil {
		t.Fatal(err)
	}
	got := readEntry(t, cleared, "xl/worksheets/sheet1.xml")
	if strings.Contains(got, sparklineExtensionURI) || !strings.Contains(got, `uri="keep-me"`) {
		t.Fatalf("wrong extension removal: %s", got)
	}
}

func TestSetSparklines_UntouchedEntryCompressedBytesRemainIdentical(t *testing.T) {
	orig := sparklineFixture(t, `<worksheet><sheetData/></worksheet>`)
	out, err := SetSparklines(orig, []SparklineWriteGroup{{TargetSheetName: "Data", Type: "line", Sparklines: []SparklineWriteItem{{SourceSheetName: "Data", SourceRef: "A1:A3", TargetCellRef: "D1"}}}})
	if err != nil {
		t.Fatal(err)
	}
	raw := func(data []byte, name string) []byte {
		zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
		if err != nil {
			t.Fatal(err)
		}
		for _, file := range zr.File {
			if file.Name == name {
				reader, err := file.OpenRaw()
				if err != nil {
					t.Fatal(err)
				}
				value, err := io.ReadAll(reader)
				if err != nil {
					t.Fatal(err)
				}
				return value
			}
		}
		t.Fatalf("missing %s", name)
		return nil
	}
	if !bytes.Equal(raw(orig, "xl/media/untouched.bin"), raw(out, "xl/media/untouched.bin")) {
		t.Fatal("untouched compressed bytes changed")
	}
}
