package pptxpatch

import (
	"bytes"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"
)

func nativeLiteralDoughnutXML(strict bool) string {
	return nativeLiteralDoughnutWithPaint(strict, literalPiePaint("0000FF")+literalPiePoint("1", "FF0000"), 3)
}
func nativeLiteralDoughnutWithPaint(strict bool, paint string, count int) string {
	source := literalPiePaintSource(strict, paint, count)
	return strings.ReplaceAll(strings.Replace(source, `<c:firstSliceAng val="90"/>`, `<c:firstSliceAng val="90"/><c:holeSize val="50"/>`, 1), "pieChart", "doughnutChart")
}
func TestNativeLiteralDoughnutSource(t *testing.T) {
	for _, strict := range []bool{false, true} {
		d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
		source := nativeLiteralDoughnutXML(strict)
		ring := extractNativeLiteralDoughnut([]byte(source), "chart.xml", d)
		if ring == nil || ring.HoleSize != 50 || ring.FirstSliceAngle != 90 || !reflect.DeepEqual(ring.Values, []int64{1, 2, 3}) || !reflect.DeepEqual(ring.Colors, []string{"#0000FF", "#FF0000", "#0000FF"}) {
			t.Fatalf("wrong source doughnut: %#v", ring)
		}
		if extractNativeLiteralPie([]byte(source), "chart.xml", d) != nil || extractNativeLiteralDoughnut([]byte(nativeLiteralPieXML(strict)), "chart.xml", d) != nil {
			t.Fatal("source families conflated")
		}
		data := nativeChartFixture(t, nativeChartFixtureOptions{strict: strict, omitPreview: true, chartXML: source})
		before := append([]byte(nil), data...)
		deck, err := ExtractNativePPTX(data, nativeTestExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		chart := nativeFixtureChart(t, deck.Slides[0])
		if chart.Chart.LiteralDoughnut == nil || chart.Chart.LiteralPie != nil || chart.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
			t.Fatal("missing read-only source projection")
		}
		if !bytes.Equal(before, data) || chart.Chart.OpaqueRef.FingerprintSHA256 != nativeSHA256([]byte(source)) {
			t.Fatal("source bytes/fingerprint changed")
		}
		encoded, err := MarshalNativePPTXJSON(deck)
		if err != nil {
			t.Fatal(err)
		}
		if _, err = DecodeNativePPTXJSON(encoded); err != nil {
			t.Fatal(err)
		}
		chart.Chart.LiteralPie = &NativeLiteralPie{Profile: "literal-pie-v1", Values: []int64{1}, Colors: []string{"#000000"}}
		if _, err = MarshalNativePPTXJSON(deck); err == nil {
			t.Fatal("competing families accepted")
		}
	}
}
func TestNativeLiteralDoughnutSourceGuards(t *testing.T) {
	for _, strict := range []bool{false, true} {
		d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
		source := nativeLiteralDoughnutXML(strict)
		for name, change := range map[string][2]string{
			"missing hole": {`<c:holeSize val="50"/>`, ""}, "hole9": {`holeSize val="50"`, `holeSize val="9"`}, "hole91": {`holeSize val="50"`, `holeSize val="91"`}, "fraction": {`holeSize val="50"`, `holeSize val="50.5"`}, "leading zero": {`holeSize val="50"`, `holeSize val="050"`}, "foreign hole": {`c:holeSize`, `a:holeSize`}, "hole attribute": {`holeSize val="50"`, `holeSize val="50" unknown="1"`},
			"missing angle": {`<c:firstSliceAng val="90"/>`, ""}, "angle361": {`firstSliceAng val="90"`, `firstSliceAng val="361"`}, "cache": {"numLit", "numCache"}, "zero": {`<c:v>1</c:v>`, `<c:v>0</c:v>`}, "count": {`ptCount val="3"`, `ptCount val="2"`}, "labels": {`</c:ser>`, `<c:dLbls/></c:ser>`}, "second ring": {`</c:ser>`, `</c:ser><c:ser/>`}, "vary": {`varyColors val="0"`, `varyColors val="1"`}, "effects": {`</a:solidFill>`, `</a:solidFill><a:effectLst/>`}, "unknown tail": {`</c:chartSpace>`, `<c:extLst/></c:chartSpace>`},
		} {
			t.Run(name, func(t *testing.T) {
				if extractNativeLiteralDoughnut([]byte(strings.ReplaceAll(source, change[0], change[1])), "chart.xml", d) != nil {
					t.Fatal("unsupported source projected")
				}
			})
		}
		for _, hole := range []string{"10", "90"} {
			for _, angle := range []string{"0", "360"} {
				qualified := strings.ReplaceAll(strings.ReplaceAll(source, `holeSize val="50"`, `holeSize val="`+hole+`"`), `firstSliceAng val="90"`, `firstSliceAng val="`+angle+`"`)
				if extractNativeLiteralDoughnut([]byte(qualified), "chart.xml", d) == nil {
					t.Fatal("valid boundary refused")
				}
			}
		}
	}
}
func TestNativeLiteralDoughnutCounts(t *testing.T) {
	for _, strict := range []bool{false, true} {
		d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
		for _, count := range []int{1, 64, 65} {
			source := nativeLiteralDoughnutWithPaint(strict, literalPiePaint("0000FF"), count)
			ring := extractNativeLiteralDoughnut([]byte(source), "chart.xml", d)
			if count <= 64 {
				if ring == nil || len(ring.Values) != count {
					t.Fatal("valid count refused")
				}
			} else if ring != nil {
				t.Fatal("point budget exceeded")
			}
		}
		if extractNativeLiteralDoughnut([]byte(nativeLiteralDoughnutWithPaint(strict, literalPiePoint("1", "FF0000"), 3)), "chart.xml", d) != nil {
			t.Fatal("missing effective paint admitted")
		}
	}
}

func TestNativeLiteralDoughnutBrowserFixture(t *testing.T) {
	dir := os.Getenv("INJOFFICE_LITERAL_DOUGHNUT_EVIDENCE")
	if dir == "" {
		t.Skip("optional external browser fixture")
	}
	for name, source := range map[string]string{"full": nativeLiteralDoughnutWithPaint(false, literalPiePaint("0000FF"), 1), "doughnut": nativeLiteralDoughnutXML(false), "refused": strings.ReplaceAll(nativeLiteralDoughnutXML(false), `<c:holeSize val="50"/>`, "")} {
		data := nativeChartFixture(t, nativeChartFixtureOptions{omitPreview: true, chartXML: source})
		if err := os.WriteFile(filepath.Join(dir, name+".pptx"), data, 0600); err != nil {
			t.Fatal(err)
		}
	}
}
