package xlsxpatch

import (
	"bytes"
	"reflect"
	"strings"
	"testing"
)

const nativeScaleThemePart = `<a:theme xmlns:a="` + drawingMLNamespace + `" name="Office"><a:themeElements><a:clrScheme name="Office">` +
	`<a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>` +
	`<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>` +
	`<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>` +
	`<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>` +
	`<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>` +
	`<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>` +
	`</a:clrScheme></a:themeElements></a:theme>`

// Four ascending values in A1:A4 plus a bound cell in C1, so every cfvo type
// this tier resolves has a range of its own to be read against.
func nativeScaleFixture(t *testing.T, conditional string) []byte {
	t.Helper()
	parts := nativeWorkbookFixture(false)
	parts["Charts/chart1.xml"] = previewChartFixture()
	parts["Book/_rels/Workbook.xml.rels"] = strings.Replace(parts["Book/_rels/Workbook.xml.rels"], `</Relationships>`,
		`<Relationship Id="rTheme" Type="`+relTypeThemeTransitional+`" Target="../theme/theme1.xml"/></Relationships>`, 1)
	parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`,
		`<Override PartName="/theme/theme1.xml" ContentType="`+themePartContentType+`"/></Types>`, 1)
	parts["theme/theme1.xml"] = nativeScaleThemePart
	parts["Sheets/s1.xml"] = `<worksheet xmlns="` + spreadsheetMLTransitional + `"><dimension ref="A1:C4"/><sheetData>` +
		`<row r="1"><c r="A1"><v>1</v></c><c r="C1"><v>4</v></c></row>` +
		`<row r="2"><c r="A2"><v>2</v></c></row>` +
		`<row r="3"><c r="A3"><v>3</v></c></row>` +
		`<row r="4"><c r="A4"><f>2+2</f><v>4</v></c></row>` +
		`</sheetData>` + conditional + `</worksheet>`
	return buildZip(t, parts)
}

func nativeScaleEntry(t *testing.T, source []byte) NativeConditionalScaleFillPreviewV1 {
	t.Helper()
	objects, err := InspectNativeWorkbookObjectsV1(source)
	if err != nil {
		t.Fatal(err)
	}
	if len(objects.ConditionalScaleFills) != 1 {
		t.Fatalf("expected one sheet overlay, got %+v", objects.ConditionalScaleFills)
	}
	return objects.ConditionalScaleFills[0]
}

func nativeScaleColors(entry NativeConditionalScaleFillPreviewV1) []string {
	out := make([]string, 0, len(entry.Cells))
	for _, cell := range entry.Cells {
		out = append(out, cell.Color)
	}
	return out
}

func TestNativeConditionalScaleFillInterpolatesEveryResolvedBound(t *testing.T) {
	for _, testCase := range []struct {
		name string
		rule string
		want []string
	}{
		{
			// Two explicit stops over the range's own extremes: the middle two
			// cells are the exact thirds Excel paints, not the end colours.
			name: "two stop min max",
			rule: `<cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF0000FF"/>`,
			want: []string{"#FF0000", "#AA0055", "#5500AA", "#0000FF"},
		},
		{
			// A `num` bound may be written as a reference to one cell; reading
			// it as a literal would leave the whole rule unpainted.
			name: "num bound from a cell reference",
			rule: `<cfvo type="num" val="1"/><cfvo type="num" val="$C$1"/><color rgb="FFFF0000"/><color rgb="FF0000FF"/>`,
			want: []string{"#FF0000", "#AA0055", "#5500AA", "#0000FF"},
		},
		{
			// percent walks the min..max span; percentile walks the sorted
			// values. On 1,2,3,4 both land on the same two bounds.
			name: "percent bounds",
			rule: `<cfvo type="percent" val="0"/><cfvo type="percent" val="100"/><color rgb="FFFF0000"/><color rgb="FF0000FF"/>`,
			want: []string{"#FF0000", "#AA0055", "#5500AA", "#0000FF"},
		},
		{
			name: "percentile bounds clamp outside the inner span",
			rule: `<cfvo type="percentile" val="33.33"/><cfvo type="percentile" val="66.67"/><color rgb="FF000000"/><color rgb="FFFFFFFF"/>`,
			want: []string{"#000000", "#000000", "#FFFFFF", "#FFFFFF"},
		},
		{
			name: "three stop scale interpolates each half separately",
			rule: `<cfvo type="min"/><cfvo type="num" val="2"/><cfvo type="max"/><color rgb="FF000000"/><color rgb="FFFF0000"/><color rgb="FFFFFFFF"/>`,
			want: []string{"#000000", "#FF0000", "#FF8080", "#FFFFFF"},
		},
		{
			// A theme stop carries no RGB of its own; an unresolved slot would
			// silently paint the wrong colour or nothing at all.
			name: "tinted theme stops",
			rule: `<cfvo type="min"/><cfvo type="max"/><color theme="4" tint="0.4"/><color theme="4"/>`,
			want: []string{"#8FAADC", "#7697D4", "#5D85CC", "#4472C4"},
		},
	} {
		t.Run(testCase.name, func(t *testing.T) {
			source := nativeScaleFixture(t, `<conditionalFormatting sqref="A1:A4"><cfRule type="colorScale" priority="3"><colorScale>`+testCase.rule+`</colorScale></cfRule></conditionalFormatting>`)
			before := bytes.Clone(source)
			nativeBefore, err := ExtractNativeWorkbookV2(source)
			if err != nil {
				t.Fatal(err)
			}
			entry := nativeScaleEntry(t, source)
			if entry.Status != "available" || entry.SheetID != "7" || entry.SheetPart != "Sheets/s1.xml" {
				t.Fatalf("%+v", entry)
			}
			if got := nativeScaleColors(entry); !reflect.DeepEqual(got, testCase.want) {
				t.Fatalf("colors = %v; want %v", got, testCase.want)
			}
			for index, cell := range entry.Cells {
				if cell.Row != index || cell.Column != 0 {
					t.Fatalf("cells are not row-major over the source range: %+v", entry.Cells)
				}
			}
			if len(entry.Ranges) != 1 || entry.Ranges[0].Ref != "A1:A4" || entry.Ranges[0].Priority != 3 {
				t.Fatalf("%+v", entry.Ranges)
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
				t.Fatal("colour-scale preview weakened the conditional preservation inventory")
			}
		})
	}
}

// A bound this tier cannot resolve must cost only its own rule. Dropping the
// whole sheet would leave a workbook whose other scales are exactly known
// painting nothing at all.
func TestNativeConditionalScaleFillPaintsEveryRuleItResolves(t *testing.T) {
	source := nativeScaleFixture(t,
		`<conditionalFormatting sqref="A1:A4"><cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF0000FF"/></colorScale></cfRule></conditionalFormatting>`+
			`<conditionalFormatting sqref="C1:C1"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="formula" val="2*A1+3"/><color rgb="FFFF0000"/><color rgb="FF0000FF"/></colorScale></cfRule></conditionalFormatting>`)
	entry := nativeScaleEntry(t, source)
	if entry.Status != "available" || len(entry.Cells) != 4 || len(entry.Ranges) != 1 || entry.Ranges[0].Ref != "A1:A4" {
		t.Fatalf("%+v", entry)
	}
	if !strings.Contains(strings.Join(entry.Warnings, " "), "1 colour-scale rule(s) were not painted") {
		t.Fatalf("the unpainted rule was not disclosed: %v", entry.Warnings)
	}
}

func TestNativeConditionalScaleFillRefusesUnqualifiedSources(t *testing.T) {
	for _, testCase := range []struct{ name, conditional string }{
		{
			// Two rules over one cell resolve by priority and stopIfTrue, which
			// this tier does not evaluate; painting either one would be a guess.
			name: "overlapping rules",
			conditional: `<conditionalFormatting sqref="A1:A4"><cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF0000FF"/></colorScale></cfRule></conditionalFormatting>` +
				`<conditionalFormatting sqref="A2:A3"><cfRule type="cellIs" operator="greaterThan" priority="1" dxfId="0"><formula>1</formula></cfRule></conditionalFormatting>`,
		},
		{
			// An x14 rule claims its range in an extension; a scale that meets
			// one is painted only once that rule is read too.
			name: "range claimed by a worksheet extension",
			conditional: `<conditionalFormatting sqref="A1:A4"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF0000FF"/></colorScale></cfRule></conditionalFormatting>` +
				`<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="dataBar" id="{00000000-0000-0000-0000-000000000000}"/><xm:sqref xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">A2:A3</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>`,
		},
		{
			name:        "descending bounds",
			conditional: `<conditionalFormatting sqref="A1:A4"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="num" val="9"/><cfvo type="num" val="0"/><color rgb="FFFF0000"/><color rgb="FF0000FF"/></colorScale></cfRule></conditionalFormatting>`,
		},
		{
			name:        "indexed stop colour",
			conditional: `<conditionalFormatting sqref="A1:A4"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color indexed="9"/><color rgb="FF0000FF"/></colorScale></cfRule></conditionalFormatting>`,
		},
		{
			name:        "translucent stop colour",
			conditional: `<conditionalFormatting sqref="A1:A4"><cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="80FF0000"/><color rgb="FF0000FF"/></colorScale></cfRule></conditionalFormatting>`,
		},
	} {
		t.Run(testCase.name, func(t *testing.T) {
			entry := nativeScaleEntry(t, nativeScaleFixture(t, testCase.conditional))
			if entry.Status != "unavailable" || len(entry.Cells) != 0 || len(entry.Ranges) != 0 {
				t.Fatalf("%+v", entry)
			}
			if len(entry.Warnings) != 1 || !strings.Contains(entry.Warnings[0], "remain preserved") {
				t.Fatalf("%v", entry.Warnings)
			}
		})
	}
}

// A workbook with no colour scale at all must not gain an overlay entry: the
// field is the signal the render path keys on.
func TestNativeConditionalScaleFillIsAbsentWithoutAColourScale(t *testing.T) {
	objects, err := InspectNativeWorkbookObjectsV1(buildZip(t, nativeConditionalFixture(false)))
	if err != nil {
		t.Fatal(err)
	}
	if len(objects.ConditionalScaleFills) != 0 {
		t.Fatalf("%+v", objects.ConditionalScaleFills)
	}
}
