package pptxpatch

import (
	"bytes"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"testing"
)

func TestNativeSignedAreaMinimumFixtures(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, grouping := range []string{"stacked", "percentStacked"} {
			for _, mode := range []string{"mixed", "negative", "zero-category", "reversed", "labels", "large", "positive-scale", "grouped"} {
				if mode == "large" && grouping != "stacked" {
					continue
				}
				name := fmt.Sprintf("signed-area-%v-%s-%s", strict, grouping, mode)
				source := nativeAreaXML(false, grouping)
				if mode == "labels" {
					source = strings.Replace(nativeAreaBrowserXML("labels"), `grouping val="standard"`, `grouping val="`+grouping+`"`, 1)
				}
				min, max := "-10", "10"
				if grouping == "percentStacked" || mode == "large" {
					min, max = "-1", "1"
				}
				if mode == "positive-scale" {
					min = "0"
				}
				source = regexp.MustCompile(`<c:min val="[^"]+"/>`).ReplaceAllString(source, `<c:min val="`+min+`"/>`)
				source = regexp.MustCompile(`<c:max val="[^"]+"/>`).ReplaceAllString(source, `<c:max val="`+max+`"/>`)
				count := 3
				rows := [][]string{{"4", "4", "4"}, {"-2", "-2", "-2"}, {"3", "3", "3"}, {"-1", "-1", "-1"}}
				if mode == "negative" {
					rows = [][]string{{"-2", "-2", "-2"}, {"-3", "-3", "-3"}, {"-1", "-1", "-1"}, {"-4", "-4", "-4"}}
				}
				if mode == "zero-category" {
					rows = [][]string{{"2", "0", "-2"}, {"-2", "-0", "2"}, {"3", "0.0", "-3"}, {"-3", "0", "3"}}
				}
				if mode == "large" {
					count = 256
					rows = [][]string{{}, {}}
					for i := 0; i < count; i++ {
						v, d := "-4", "8"
						if i%2 == 1 {
							v, d = "4", "-8"
						}
						rows[0] = append(rows[0], v)
						rows[1] = append(rows[1], d)
					}
				}
				re := regexp.MustCompile(`(?s)<c:ser>.*?</c:ser>`)
				first := re.FindString(source)
				series := ""
				for i, row := range rows {
					item := strings.NewReplacer(`idx val="5"`, fmt.Sprintf(`idx val="%d"`, 12+i), `order val="0"`, fmt.Sprintf(`order val="%d"`, len(rows)-1-i), "123456", []string{"1E88E5", "E53935", "43A047", "8E24AA"}[i]).Replace(first)
					cat, val := fmt.Sprintf(`<c:cat><c:strLit><c:ptCount val="%d"/>`, count), fmt.Sprintf(`<c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="%d"/>`, count)
					for j, v := range row {
						cat += fmt.Sprintf(`<c:pt idx="%d"><c:v>%c</c:v></c:pt>`, j, 'A'+j%26)
						val += fmt.Sprintf(`<c:pt idx="%d"><c:v>%s</c:v></c:pt>`, j, v)
					}
					item = regexp.MustCompile(`(?s)<c:cat>.*?</c:val>`).ReplaceAllString(item, cat+`</c:strLit></c:cat>`+val+`</c:numLit></c:val>`)
					series += item
				}
				source = re.ReplaceAllString(source, series)
				if mode == "reversed" {
					source = strings.ReplaceAll(source, `orientation val="minMax"`, `orientation val="maxMin"`)
				}
				if strict {
					a, _ := nativeDialectForPresentation(xmlNamePresentation(false))
					b, _ := nativeDialectForPresentation(xmlNamePresentation(true))
					source = strings.NewReplacer(a.chart, b.chart, a.drawing, b.drawing, `lblOffset val="0"`, `lblOffset val="0%"`).Replace(source)
				}
				frame := strings.NewReplacer(`y="2000000"`, `y="700000"`, `cx="3000000" cy="2000000"`, `cx="9000000" cy="5000000"`).Replace(nativeChartGraphicFrameXML(strict, 3, name, ""))
				if mode == "grouped" {
					frame = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="9" name="Signed area group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot="1800000" flipV="1"><a:off x="1000000" y="1000000"/><a:ext cx="6000001" cy="4000001"/><a:chOff x="0" y="0"/><a:chExt cx="12000000" cy="8000000"/></a:xfrm></p:grpSpPr>` + frame + `</p:grpSp>`
				}
				input := nativeSeriesOrderChartOnly(t, nativeChartFixture(t, nativeChartFixtureOptions{strict: strict, chartXML: source, frameXML: frame, omitPreview: true}), false, false)
				before := bytes.Clone(input)
				deck, e := ExtractNativePPTX(input, nativeAtomicTestExtractOptions())
				if e != nil {
					t.Fatal(e)
				}
				chart := nativeRadarFindChart(deck.Slides[0].Elements)
				if chart == nil || chart.Chart.LiteralArea == nil || !validNativeLiteralArea(chart.Chart.LiteralArea) {
					t.Fatalf("source refused %s", name)
				}
				area := chart.Chart.LiteralArea
				if area.SourceBaseline.Value != min || area.Series[0].Order != int64(len(rows)-1) || chart.Chart.OpaqueRef.FingerprintSHA256 != nativeSHA256([]byte(source)) {
					t.Fatal("source identity changed")
				}
				saved := area.SourceBaseline
				area.SourceBaseline = nil
				if validNativeLiteralArea(area) {
					t.Fatal("legacy signed source admitted")
				}
				area.SourceBaseline = saved
				encoded, e := json.Marshal(deck)
				if e != nil {
					t.Fatal(e)
				}
				if _, e = DecodeNativePPTXJSON(encoded); e != nil {
					t.Fatal(e)
				}
				if !bytes.Equal(input, before) {
					t.Fatal("source mutated")
				}
				if out := os.Getenv("INJOFFICE_PPTX_SIGNED_AREA_FIXTURES"); out != "" {
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
	}
}
