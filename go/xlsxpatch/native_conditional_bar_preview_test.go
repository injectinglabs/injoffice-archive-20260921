package xlsxpatch

import (
	"bytes"
	"reflect"
	"strings"
	"testing"
)

const nativeBarX14Namespace = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"

// One legacy rule plus its x14 twin, which is the pair Excel always writes.
// A1:A4 hold 1, 2, 3 and -2 so the automatic bounds straddle zero and both
// sides of the axis are exercised.
func nativeBarWorksheet(t *testing.T, conditional string) []byte {
	t.Helper()
	parts := nativeWorkbookFixture(false)
	parts["Charts/chart1.xml"] = previewChartFixture()
	parts["Book/_rels/Workbook.xml.rels"] = strings.Replace(parts["Book/_rels/Workbook.xml.rels"], `</Relationships>`,
		`<Relationship Id="rTheme" Type="`+relTypeThemeTransitional+`" Target="../theme/theme1.xml"/></Relationships>`, 1)
	parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`,
		`<Override PartName="/theme/theme1.xml" ContentType="`+themePartContentType+`"/></Types>`, 1)
	parts["theme/theme1.xml"] = nativeScaleThemePart
	parts["Sheets/s1.xml"] = `<worksheet xmlns="` + spreadsheetMLTransitional + `"><dimension ref="A1:A4"/><sheetData>` +
		`<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row>` +
		`<row r="3"><c r="A3"><v>3</v></c></row><row r="4"><c r="A4"><v>-2</v></c></row>` +
		`</sheetData>` + conditional + `</worksheet>`
	return buildZip(t, parts)
}

func nativeBarFixture(t *testing.T, legacy, extension string) []byte {
	t.Helper()
	return nativeBarWorksheet(t,
		`<conditionalFormatting sqref="A1:A4"><cfRule type="dataBar" priority="3"><dataBar>`+legacy+
			`</dataBar><extLst><ext uri="`+nativeBarDataBarExtURI+`" xmlns:x14="`+nativeBarX14Namespace+`"><x14:id>{BAR-0001}</x14:id></ext></extLst></cfRule></conditionalFormatting>`+
			`<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="`+nativeBarX14Namespace+`"><x14:conditionalFormattings><x14:conditionalFormatting>`+
			`<x14:cfRule type="dataBar" id="{BAR-0001}"><x14:dataBar `+extension+`</x14:dataBar></x14:cfRule>`+
			`<xm:sqref xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">A1:A4</xm:sqref>`+
			`</x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>`)
}

func nativeBarEntry(t *testing.T, source []byte) NativeConditionalBarFillPreviewV1 {
	t.Helper()
	objects, err := InspectNativeWorkbookObjectsV1(source)
	if err != nil {
		t.Fatal(err)
	}
	if len(objects.ConditionalBarFills) != 1 {
		t.Fatalf("expected one sheet overlay, got %+v", objects.ConditionalBarFills)
	}
	return objects.ConditionalBarFills[0]
}

const nativeBarPlainExtension = `minLength="0" maxLength="100" gradient="0"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/>`

func TestNativeConditionalBarFillMeasuresBarsFromTheBoundSpan(t *testing.T) {
	// Automatic bounds run from min(0, lowest) to max(0, highest), so the axis
	// sits two fifths in and each bar is its own share of that span.
	source := nativeBarFixture(t, `<cfvo type="min"/><cfvo type="max"/><color rgb="FF0000FF"/>`, nativeBarPlainExtension)
	before := bytes.Clone(source)
	nativeBefore, err := ExtractNativeWorkbookV2(source)
	if err != nil {
		t.Fatal(err)
	}
	entry := nativeBarEntry(t, source)
	if entry.Status != "available" || entry.SheetID != "7" || entry.SheetPart != "Sheets/s1.xml" {
		t.Fatalf("%+v", entry)
	}
	want := []NativeConditionalBarFillCellV1{
		{Row: 0, Column: 0, Start: 400, End: 600, Axis: -1, Color: "#0000FF"},
		{Row: 1, Column: 0, Start: 400, End: 800, Axis: -1, Color: "#0000FF"},
		{Row: 2, Column: 0, Start: 400, End: 1000, Axis: -1, Color: "#0000FF"},
		{Row: 3, Column: 0, Start: 0, End: 400, Axis: -1, Color: "#0000FF"},
	}
	if !reflect.DeepEqual(entry.Cells, want) {
		t.Fatalf("bars = %+v; want %+v", entry.Cells, want)
	}
	nativeAfter, err := ExtractNativeWorkbookV2(source)
	if err != nil || !reflect.DeepEqual(nativeBefore, nativeAfter) || !bytes.Equal(before, source) {
		t.Fatal("source or native mutation authority changed")
	}
	preserved := false
	for _, unsupported := range nativeAfter.Unsupported {
		if unsupported.Code == "CONDITIONAL_FORMATTING" {
			preserved = true
		}
	}
	if !preserved {
		t.Fatal("data-bar preview weakened the conditional preservation inventory")
	}
}

// A negative value puts zero inside the span, and Excel then grows its bar
// leftwards from the axis in the negative colour. Measuring it rightwards
// would paint the bar on the wrong side of the axis.
func TestNativeConditionalBarFillGrowsNegativeValuesFromTheAxis(t *testing.T) {
	source := nativeBarFixture(t,
		`<cfvo type="min"/><cfvo type="max"/><color rgb="FF0000FF"/>`,
		`minLength="0" maxLength="100" gradient="0" axisPosition="automatic"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:negativeFillColor rgb="FFFF0000"/><x14:axisColor rgb="FF00FF00"/>`)
	entry := nativeBarEntry(t, source)
	want := []NativeConditionalBarFillCellV1{
		{Row: 0, Column: 0, Start: 400, End: 600, Axis: 400, Color: "#0000FF", AxisColor: "#00FF00"},
		{Row: 1, Column: 0, Start: 400, End: 800, Axis: 400, Color: "#0000FF", AxisColor: "#00FF00"},
		{Row: 2, Column: 0, Start: 400, End: 1000, Axis: 400, Color: "#0000FF", AxisColor: "#00FF00"},
		{Row: 3, Column: 0, Start: 0, End: 400, Axis: 400, Color: "#FF0000", AxisColor: "#00FF00"},
	}
	if !reflect.DeepEqual(entry.Cells, want) {
		t.Fatalf("bars = %+v; want %+v", entry.Cells, want)
	}
}

// The x14 twin carries the geometry, so the downlevel bounds must not be the
// ones measured when the two disagree.
func TestNativeConditionalBarFillPrefersTheExtensionBounds(t *testing.T) {
	source := nativeBarFixture(t,
		`<cfvo type="num" val="0"/><cfvo type="num" val="100"/><color rgb="FF0000FF"/>`,
		`minLength="0" maxLength="100" gradient="0"><x14:cfvo type="num" value="0"/><x14:cfvo type="num" value="4"/>`)
	entry := nativeBarEntry(t, source)
	// The downlevel 0..100 span would make every bar a sliver; 0..4 does not.
	// A4 clamps to the zero bound, and a bar of no length with no axis to draw
	// is not emitted at all.
	if len(entry.Cells) != 3 || entry.Cells[0].End != 250 || entry.Cells[2].End != 750 {
		t.Fatalf("%+v", entry.Cells)
	}
}

func TestNativeConditionalBarFillRefusesUnqualifiedSources(t *testing.T) {
	for _, testCase := range []struct{ name, legacy, extension string }{
		{
			// A gradient bar's painted colour varies along its own length and
			// is not the flat fill this tier emits.
			name: "gradient bar", legacy: `<cfvo type="min"/><cfvo type="max"/><color rgb="FF0000FF"/>`,
			extension: `minLength="0" maxLength="100"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/>`,
		},
		{
			// Clipped lengths change where every bar starts and ends.
			name: "clipped minimum length", legacy: `<cfvo type="min"/><cfvo type="max"/><color rgb="FF0000FF"/>`,
			extension: `minLength="10" maxLength="90" gradient="0"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/>`,
		},
		{
			name: "middle axis", legacy: `<cfvo type="min"/><cfvo type="max"/><color rgb="FF0000FF"/>`,
			extension: `minLength="0" maxLength="100" gradient="0" axisPosition="middle"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/>`,
		},
		{
			name: "formula bound", legacy: `<cfvo type="min"/><cfvo type="max"/><color rgb="FF0000FF"/>`,
			extension: `minLength="0" maxLength="100" gradient="0"><x14:cfvo type="autoMin"/><x14:cfvo type="formula" value="2*A1+3"/>`,
		},
		{
			name: "indexed bar colour", legacy: `<cfvo type="min"/><cfvo type="max"/><color indexed="9"/>`,
			extension: nativeBarPlainExtension,
		},
		{
			// A border the rule asks for but never names would be dropped, so
			// the bar would paint a size Excel does not.
			name: "declared border with no colour", legacy: `<cfvo type="min"/><cfvo type="max"/><color rgb="FF0000FF"/>`,
			extension: `minLength="0" maxLength="100" gradient="0" border="1"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/>`,
		},
	} {
		t.Run(testCase.name, func(t *testing.T) {
			entry := nativeBarEntry(t, nativeBarFixture(t, testCase.legacy, testCase.extension))
			if entry.Status != "unavailable" || len(entry.Cells) != 0 || len(entry.Ranges) != 0 {
				t.Fatalf("%+v", entry)
			}
			if len(entry.Warnings) != 1 || !strings.Contains(entry.Warnings[0], "remain preserved") {
				t.Fatalf("%v", entry.Warnings)
			}
		})
	}
}

// A downlevel rule with no x14 twin carries no axis, no negative colour and no
// gradient flag, so its geometry is not established.
func TestNativeConditionalBarFillRefusesADownlevelOnlyRule(t *testing.T) {
	source := nativeBarWorksheet(t, `<conditionalFormatting sqref="A1:A4"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF0000FF"/></dataBar></cfRule></conditionalFormatting>`)
	entry := nativeBarEntry(t, source)
	if entry.Status != "unavailable" || len(entry.Cells) != 0 {
		t.Fatalf("%+v", entry)
	}
}

func TestNativeConditionalBarFillIsAbsentWithoutADataBar(t *testing.T) {
	objects, err := InspectNativeWorkbookObjectsV1(buildZip(t, nativeConditionalFixture(false)))
	if err != nil {
		t.Fatal(err)
	}
	if len(objects.ConditionalBarFills) != 0 {
		t.Fatalf("%+v", objects.ConditionalBarFills)
	}
}
