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

func nativeRadarFindChart(elements []NativeElement) *NativeElement {
	for i := range elements {
		if elements[i].Kind == NativeElementKindChart {
			return &elements[i]
		}
		if c := nativeRadarFindChart(elements[i].Children); c != nil {
			return c
		}
	}
	return nil
}
func TestNativeLiteralRadarSourceIntegration(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, filled := range []bool{false, true} {
			for _, mode := range []string{"plain", "rotated", "grouped", "reversed"} {
				name := fmt.Sprintf("radar-%v-%v-%s", strict, filled, mode)
				t.Run(name, func(t *testing.T) {
					source := nativeSeriesOrderXML(nativeRadarXML(strict, filled))
					at := 0
					source = regexp.MustCompile(`(?s)<c:ser>.*?</c:ser>`).ReplaceAllStringFunc(source, func(row string) string {
						color := []string{"1E88E5", "E53935", "43A047"}[at]
						raw := []string{"-5.00", "5e0", "15.0"}[at]
						row = regexp.MustCompile(`(?s)<c:val>.*?</c:val>`).ReplaceAllStringFunc(row, func(values string) string {
							return regexp.MustCompile(`<c:v>[^<]*</c:v>`).ReplaceAllString(values, "<c:v>"+raw+"</c:v>")
						})
						at++
						return strings.ReplaceAll(strings.ReplaceAll(row, "E53935", color), "123456", color)
					})
					if mode == "reversed" {
						source = strings.ReplaceAll(source, `orientation val="minMax"`, `orientation val="maxMin"`)
					}
					frame := nativeChartGraphicFrameXML(strict, 3, name, "")
					if mode == "rotated" {
						frame = strings.Replace(frame, `<p:xfrm>`, `<p:xfrm rot="1800000" flipH="1">`, 1)
					}
					if mode == "grouped" {
						frame = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="9" name="Radar group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot="1800000" flipV="1"><a:off x="1000000" y="1000000"/><a:ext cx="6000001" cy="4000001"/><a:chOff x="0" y="0"/><a:chExt cx="12000000" cy="8000000"/></a:xfrm></p:grpSpPr>` + frame + `</p:grpSp>`
					}
					original := nativeChartFixture(t, nativeChartFixtureOptions{strict: strict, chartXML: source, omitPreview: true, frameXML: frame})
					input := nativeSeriesOrderChartOnly(t, original, false, false)
					before := bytes.Clone(input)
					deck, err := ExtractNativePPTX(input, nativeAtomicTestExtractOptions())
					if err != nil {
						t.Fatal(err)
					}
					c := nativeRadarFindChart(deck.Slides[0].Elements)
					if c == nil || c.Chart.LiteralRadar == nil || ((mode == "rotated" || mode == "grouped") && c.GraphicFrameLayout == nil) || c.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
						t.Fatalf("missing source radar mode=%s", mode)
					}
					r := c.Chart.LiteralRadar
					if !validNativeLiteralRadar(r) || len(r.Series) != 3 || r.Series[0].Order != 2 || r.Series[0].Index != 12 || r.Series[0].Color != "#1E88E5" {
						t.Fatal("radar identity changed", r)
					}
					if c.Chart.OpaqueRef.FingerprintSHA256 != nativeSHA256([]byte(source)) || c.Source == nil || c.Source.RelationshipID == nil || *c.Source.RelationshipID != c.Chart.RelationshipID {
						t.Fatal("lost source binding")
					}
					encoded, err := json.Marshal(deck)
					if err != nil {
						t.Fatal(err)
					}
					if _, err = DecodeNativePPTXJSON(encoded); err != nil {
						t.Fatal(err)
					}
					r.Style = "marker"
					if len(ValidateNativePPTX(deck)) == 0 {
						t.Fatal("unsupported public style admitted")
					}
					r.Style = map[bool]string{false: "standard", true: "filled"}[filled]
					saved := r.Series[0].Values[0]
					r.Series[0].Values[0] = "21"
					if len(ValidateNativePPTX(deck)) == 0 {
						t.Fatal("out-of-scale public value admitted")
					}
					r.Series[0].Values[0] = saved
					c.Chart.LiteralPie = &NativeLiteralPie{Profile: "literal-pie-v1", Values: []int64{1}, Colors: []string{"#000000"}}
					if len(ValidateNativePPTX(deck)) == 0 {
						t.Fatal("competing families admitted")
					}
					c.Chart.LiteralPie = nil
					if !bytes.Equal(input, before) {
						t.Fatal("source package mutated")
					}
					if dir := os.Getenv("INJOFFICE_PPTX_RADAR_FIXTURES"); dir != "" {
						if err = os.MkdirAll(dir, 0700); err != nil {
							t.Fatal(err)
						}
						for ext, data := range map[string][]byte{".pptx": input, "-deck.json": encoded} {
							if err = os.WriteFile(filepath.Join(dir, name+ext), data, 0600); err != nil {
								t.Fatal(err)
							}
						}
					}
				})
			}
		}
	}
}
func TestNativeLiteralRadarUnsupportedSourceStaysOpaque(t *testing.T) {
	for name, mutate := range map[string]func(string) string{
		"marker": func(s string) string {
			return strings.Replace(s, `radarStyle val="standard"`, `radarStyle val="marker"`, 1)
		},
		"outside":   func(s string) string { return strings.Replace(s, `<c:v>-1</c:v>`, `<c:v>-11</c:v>`, 1) },
		"extension": func(s string) string { return strings.Replace(s, `</c:radarChart>`, `<c:extLst/></c:radarChart>`, 1) },
		"labels": func(s string) string {
			return strings.Replace(s, `tickLblPos val="none"`, `tickLblPos val="nextTo"`, 1)
		},
		"reference": func(s string) string {
			return strings.Replace(strings.Replace(s, `<c:val><c:numLit>`, `<c:val><c:numRef><c:f>Data!A1:A3</c:f><c:numCache>`, 1), `</c:numLit></c:val>`, `</c:numCache></c:numRef></c:val>`, 1)
		},
	} {
		t.Run(name, func(t *testing.T) {
			source := mutate(nativeRadarXML(false, false))
			d, _ := nativeDialectForPresentation(xmlNamePresentation(false))
			if extractNativeLiteralRadar([]byte(source), "chart.xml", d) != nil {
				t.Fatal("unsupported source admitted")
			}
			input := nativeChartFixture(t, nativeChartFixtureOptions{chartXML: source, omitPreview: true})
			before := bytes.Clone(input)
			deck, err := ExtractNativePPTX(input, nativeAtomicTestExtractOptions())
			if err != nil {
				t.Fatal(err)
			}
			c := nativeRadarFindChart(deck.Slides[0].Elements)
			if c == nil || c.Chart.LiteralRadar != nil || c.Chart.OpaqueRef.FingerprintSHA256 != nativeSHA256([]byte(source)) || c.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
				t.Fatal("unsupported radar source did not retain exact opaque authority")
			}
			if !bytes.Equal(input, before) {
				t.Fatal("unsupported source package mutated")
			}

		})
	}
}
