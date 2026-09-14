package pptxpatch

import (
	"bytes"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestNativeAreaMinimumBaselineFixtures(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, mode := range []string{"standard", "stacked", "percent", "zero", "zero-min", "reversed", "labels", "ticks", "grouped", "large"} {
			name := fmt.Sprintf("area-min-%v-%s", strict, mode)
			source := nativeAreaBrowserXML(mode)
			if mode == "zero-min" {
				source = strings.ReplaceAll(nativeAreaBrowserXML("zero"), `<c:min val="-2"/>`, `<c:min val="0"/>`)
			}
			if mode == "reversed" {
				source = strings.ReplaceAll(nativeAreaBrowserXML("standard"), `orientation val="minMax"`, `orientation val="maxMin"`)
			}
			if mode == "ticks" {
				source = nativeAreaBrowserXML("labels")
				source = strings.Replace(source, `<c:max val="2"/>`, `<c:max val="0"/>`, 1)
				source = strings.ReplaceAll(source, `<c:majorTickMark val="none"/>`, `<c:majorTickMark val="out"/>`)
			}
			if strict {
				d, _ := nativeDialectForPresentation(xmlNamePresentation(true))
				normal, _ := nativeDialectForPresentation(xmlNamePresentation(false))
				source = strings.NewReplacer(normal.chart, d.chart, normal.drawing, d.drawing, `lblOffset val="0"`, `lblOffset val="0%"`).Replace(source)
			}
			frame := strings.NewReplacer(`y="2000000"`, `y="700000"`, `cx="3000000" cy="2000000"`, `cx="9000000" cy="5000000"`).Replace(nativeChartGraphicFrameXML(strict, 3, name, ""))
			if mode == "grouped" {
				frame = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="9" name="Area group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot="1800000" flipV="1"><a:off x="1000000" y="1000000"/><a:ext cx="6000001" cy="4000001"/><a:chOff x="0" y="0"/><a:chExt cx="12000000" cy="8000000"/></a:xfrm></p:grpSpPr>` + frame + `</p:grpSp>`
			}
			input := nativeSeriesOrderChartOnly(t, nativeChartFixture(t, nativeChartFixtureOptions{strict: strict, chartXML: source, frameXML: frame, omitPreview: true}), false, false)
			before := bytes.Clone(input)
			deck, e := ExtractNativePPTX(input, nativeAtomicTestExtractOptions())
			if e != nil {
				t.Fatal(e)
			}
			chart := nativeRadarFindChart(deck.Slides[0].Elements)
			if chart == nil || chart.Chart.LiteralArea == nil || chart.Chart.LiteralArea.SourceBaseline == nil {
				t.Fatalf("missing baseline %s", name)
			}
			area := chart.Chart.LiteralArea
			if area.SourceBaseline.Crossing != "min" || area.SourceBaseline.Value != *area.YAxis.Min || !validNativeLiteralArea(area) || chart.Chart.OpaqueRef.FingerprintSHA256 != nativeSHA256([]byte(source)) {
				t.Fatal("source baseline identity drift")
			}
			encoded, e := json.Marshal(deck)
			if e != nil {
				t.Fatal(e)
			}
			if _, e = DecodeNativePPTXJSON(encoded); e != nil {
				t.Fatal(e)
			}
			for _, bad := range []NativeAreaSourceBaseline{{Crossing: "max", Value: *area.YAxis.Min}, {Crossing: "min", Value: "999"}} {
				saved := area.SourceBaseline
				area.SourceBaseline = &bad
				if validNativeLiteralArea(area) {
					t.Fatal("forged baseline accepted")
				}
				area.SourceBaseline = saved
			}
			if !bytes.Equal(input, before) {
				t.Fatal("source changed")
			}
			if out := os.Getenv("INJOFFICE_PPTX_AREA_MIN_FIXTURES"); out != "" {
				if e = os.MkdirAll(out, 0700); e != nil {
					t.Fatal(e)
				}
				for ext, b := range map[string][]byte{".pptx": input, "-deck.json": encoded} {
					if e = os.WriteFile(filepath.Join(out, name+ext), b, 0600); e != nil {
						t.Fatal(e)
					}
				}
			}
		}
	}
	for _, name := range []string{"autoZero", "max", "negative-stacked", "negative-percent", "duplicate"} {
		source := nativeAreaXML(false, "standard")
		switch name {
		case "autoZero", "max":
			source = strings.Replace(source, `crosses val="min"`, `crosses val="`+name+`"`, 1)
		case "duplicate":
			source = strings.Replace(source, `<c:crosses val="min"/>`, `<c:crosses val="min"/><c:crosses val="min"/>`, 1)
		default:
			grouping := "stacked"
			if name == "negative-percent" {
				grouping = "percentStacked"
			}
			source = strings.Replace(source, `grouping val="standard"`, `grouping val="`+grouping+`"`, 1)
		}
		input := nativeSeriesOrderChartOnly(t, nativeChartFixture(t, nativeChartFixtureOptions{chartXML: source, omitPreview: true}), false, false)
		deck, e := ExtractNativePPTX(input, nativeAtomicTestExtractOptions())
		if e != nil {
			t.Fatal(e)
		}
		chart := nativeRadarFindChart(deck.Slides[0].Elements)
		if chart == nil || chart.Chart.LiteralArea != nil || chart.Chart.OpaqueRef.FingerprintSHA256 != nativeSHA256([]byte(source)) {
			t.Fatalf("source refusal changed %s", name)
		}
		if out := os.Getenv("INJOFFICE_PPTX_AREA_MIN_FIXTURES"); out != "" {
			encoded, _ := json.Marshal(deck)
			for ext, b := range map[string][]byte{".pptx": input, "-deck.json": encoded} {
				if e = os.WriteFile(filepath.Join(out, "area-min-negative-"+name+ext), b, 0600); e != nil {
					t.Fatal(e)
				}
			}
		}
	}
}
