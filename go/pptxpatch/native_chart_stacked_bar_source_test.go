package pptxpatch

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func nativeStackedBarXML(strict bool, grouping string) string {
	source := strings.Replace(nativeBarXML(strict, true), `grouping val="clustered"`, `grouping val="`+grouping+`"`, 1)
	if strict {
		return strings.Replace(source, `overlap val="0%"`, `overlap val="100%"`, 1)
	}
	return strings.Replace(source, `overlap val="0"`, `overlap val="100"`, 1)
}
func TestNativeChartStackedBarSource(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, grouping := range []string{"stacked", "percentStacked"} {
			d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
			source := []byte(nativeStackedBarXML(strict, grouping))
			before := append([]byte(nil), source...)
			chart := extractNativeChartStackedBar(source, "chart.xml", d)
			if chart == nil || chart.Grouping != grouping || chart.Series[0].Values[0] != "-0.5" || !bytes.Equal(source, before) {
				t.Fatalf("lost stacked source: %#v", chart)
			}
			if extractNativeChartBar(source, "chart.xml", d) != nil {
				t.Fatal("stacked chart admitted by clustered profile")
			}
			horizontal := strings.NewReplacer(`barDir val="col"`, `barDir val="bar"`, `axPos val="b"`, `axPos val="l"`, `axPos val="l"`, `axPos val="b"`, `orientation val="minMax"`, `orientation val="maxMin"`).Replace(string(source))
			if chart = extractNativeChartStackedBar([]byte(horizontal), "chart.xml", d); chart == nil || chart.Direction != "bar" || chart.ValueAxis.Orientation != "maxMin" {
				t.Fatal("lost horizontal/reversed source")
			}
		}
	}
}
func TestNativeChartStackedBarRefusals(t *testing.T) {
	d, _ := nativeDialectForPresentation(xmlNamePresentation(false))
	source := nativeStackedBarXML(false, "stacked")
	for name, pair := range map[string][2]string{"partial overlap": {`overlap val="100"`, `overlap val="99"`}, "wrong family": {`grouping val="stacked"`, `grouping val="clustered"`}, "cache": {`numLit`, `numCache`}, "axis": {`crossAx val="20"`, `crossAx val="90"`}, "extra family": {`</c:barChart>`, `</c:barChart><c:lineChart/>`}, "unknown": {`</c:barChart>`, `<c:extLst/></c:barChart>`}} {
		t.Run(name, func(t *testing.T) {
			if extractNativeChartStackedBar([]byte(strings.ReplaceAll(source, pair[0], pair[1])), "chart.xml", d) != nil {
				t.Fatal("unqualified source admitted")
			}
		})
	}
}

func TestNativeChartStackedBarXMLOrder(t *testing.T) {
	for _, strict := range []bool{false, true} {
		d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
		first := nativeBarSeriesXML(strict)
		second := strings.NewReplacer(`idx val="7"`, `idx val="9"`, `order val="0"`, `order val="1"`).Replace(first)
		payload := []byte(strings.Replace(nativeStackedBarXML(strict, "stacked"), first, second+first, 1))
		before := append([]byte(nil), payload...)
		chart := extractNativeChartStackedBar(payload, "chart.xml", d)
		if chart == nil || len(chart.Series) != 2 || chart.Series[0].Index != 9 || chart.Series[0].Order != 1 || chart.Series[1].Index != 7 || chart.Series[1].Order != 0 || !bytes.Equal(payload, before) {
			t.Fatal("XML order or original order metadata changed")
		}
	}
}

// These source fixtures become the actual WASM/browser inputs at attachment.
// Admission is checked directly until the shared native contract is connected.
func TestNativeStackedBrowserFixtures(t *testing.T) {
	for _, family := range []string{"bar", "line"} {
		for _, grouping := range []string{"stacked", "percentStacked"} {
			for _, mode := range []string{"mixed", "zero", "permuted", "reversed", "horizontal", "labels"} {
				if family == "line" && mode == "horizontal" {
					continue
				}
				name := family + "-" + grouping + "-" + mode
				t.Run(name, func(t *testing.T) {
					source := nativeStackedBarXML(false, grouping)
					if family == "line" {
						source = strings.Replace(nativeConnectedXML(false, false), `grouping val="standard"`, `grouping val="`+grouping+`"`, 1)
					}
					start, end := strings.Index(source, "<c:ser"), strings.Index(source, "</c:ser>")+len("</c:ser>")
					first := source[start:end]
					catStart, valEnd := strings.Index(first, "<c:cat>"), strings.Index(first, "</c:val>")+len("</c:val>")
					series := make([]string, 4)
					for i, value := range []string{"4", "-2", "3", "-1"} {
						if mode == "zero" {
							value = "0"
						}
						points := ""
						for j := 0; j < 3; j++ {
							points += fmt.Sprintf(`<c:pt idx="%d"><c:v>%s</c:v></c:pt>`, j, value)
						}
						categories := `<c:cat><c:strLit><c:ptCount val="3"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strLit></c:cat>`
						values := `<c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="3"/>` + points + `</c:numLit></c:val>`
						item := first[:catStart] + categories + values + first[valEnd:]
						oldIndex := "7"
						if family == "line" {
							oldIndex = "5"
						}
						series[i] = strings.NewReplacer(`idx val="`+oldIndex+`"`, fmt.Sprintf(`idx val="%d"`, 10+i), `order val="0"`, fmt.Sprintf(`order val="%d"`, i), "123456", []string{"E53935", "43A047", "1E88E5", "8E24AA"}[i]).Replace(item)
					}
					if mode == "permuted" {
						series[0], series[3] = series[3], series[0]
						series[1], series[2] = series[2], series[1]
					}
					source = source[:start] + strings.Join(series, "") + source[end:]
					source = strings.ReplaceAll(source, `<c:max val="20"/>`, `<c:max val="10"/>`)
					if grouping == "percentStacked" {
						source = strings.NewReplacer(`<c:max val="10"/>`, `<c:max val="1"/>`, `<c:min val="-10"/>`, `<c:min val="-1"/>`).Replace(source)
					}
					if mode == "reversed" {
						source = strings.ReplaceAll(source, `orientation val="minMax"`, `orientation val="maxMin"`)
					}
					if mode == "horizontal" {
						source = strings.NewReplacer(`barDir val="col"`, `barDir val="bar"`, `axPos val="b"`, `axPos val="l"`, `axPos val="l"`, `axPos val="b"`).Replace(source)
					}
					if mode == "labels" {
						source = strings.ReplaceAll(source, `<c:tickLblPos val="none"/>`, `<c:tickLblPos val="low"/>`)
						source = strings.ReplaceAll(source, `</c:spPr><c:crossAx`, `</c:spPr>`+nativeAxisTextXML(false)+`<c:crossAx`)
						source = strings.Replace(source, `<c:axPos val="l"/>`, `<c:axPos val="l"/><c:numFmt formatCode="0.0" sourceLinked="0"/>`, 1)
						unit := "5"
						if grouping == "percentStacked" {
							unit = ".5"
						}
						source = strings.Replace(source, `</c:valAx>`, `<c:majorUnit val="`+unit+`"/></c:valAx>`, 1)
						source = strings.Replace(source, `</c:catAx>`, `<c:auto val="0"/><c:lblAlgn val="ctr"/><c:lblOffset val="0"/><c:tickLblSkip val="1"/><c:tickMarkSkip val="1"/><c:noMultiLvlLbl val="1"/></c:catAx>`, 1)
					}
					d, _ := nativeDialectForPresentation(xmlNamePresentation(false))
					if family == "bar" && extractNativeChartStackedBar([]byte(source), "chart.xml", d) == nil {
						t.Fatal("bar fixture refused")
					}
					if family == "line" && extractNativeChartStackedLine([]byte(source), "chart.xml", d) == nil {
						t.Fatal("line fixture refused")
					}
					frame := strings.NewReplacer(`y="2000000"`, `y="700000"`, `cx="3000000" cy="2000000"`, `cx="9000000" cy="5000000"`).Replace(nativeChartGraphicFrameXML(false, 3, name, ""))
					input := nativeChartFixture(t, nativeChartFixtureOptions{omitPreview: true, chartXML: source, frameXML: frame})
					archive, err := zip.NewReader(bytes.NewReader(input), int64(len(input)))
					if err != nil {
						t.Fatal(err)
					}
					parts := []nativeExtractZipPart{}
					for _, entry := range archive.File {
						r, err := entry.Open()
						if err != nil {
							t.Fatal(err)
						}
						payload, err := io.ReadAll(r)
						r.Close()
						if err != nil {
							t.Fatal(err)
						}
						data := string(payload)
						if entry.Name == "relocated/slides/slide-a.xml" {
							for {
								start := strings.Index(data, "<p:sp>")
								if start < 0 {
									break
								}
								end := start + strings.Index(data[start:], "</p:sp>") + len("</p:sp>")
								data = data[:start] + data[end:]
							}
						}
						parts = append(parts, nativeExtractZipPart{name: entry.Name, data: data})
					}
					input = writeNativeExtractZip(t, parts)
					deck, err := ExtractNativePPTX(input, nativeTestExtractOptions())
					if err != nil {
						t.Fatal(err)
					}
					element := nativeFixtureChart(t, deck.Slides[0])
					if family == "bar" && element.Chart.LiteralStackedBar == nil || family == "line" && element.Chart.LiteralStackedLine == nil {
						t.Fatal("stacked attachment missing")
					}
					if element.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
						t.Fatal("source ownership changed")
					}
					encoded, err := MarshalNativePPTXJSON(deck)
					if err != nil {
						t.Fatal(err)
					}
					if _, err = DecodeNativePPTXJSON(encoded); err != nil {
						t.Fatal(err)
					}
					element.Chart.LiteralBubble = &NativeLiteralBubble{}
					if len(ValidateNativePPTX(deck)) == 0 {
						t.Fatal("competing chart profiles accepted")
					}

					if output := os.Getenv("INJOFFICE_PPTX_STACKED_FIXTURES"); output != "" {
						if err := os.MkdirAll(output, 0700); err != nil {
							t.Fatal(err)
						}
						if err := os.WriteFile(filepath.Join(output, name+".pptx"), input, 0600); err != nil {
							t.Fatal(err)
						}
					}
				})
			}
		}
	}
}
