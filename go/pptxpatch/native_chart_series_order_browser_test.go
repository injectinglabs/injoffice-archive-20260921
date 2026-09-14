package pptxpatch

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"testing"
)

func TestNativeSeriesOrderBrowserFixtures(t *testing.T) {
	for _, name := range []string{"bar", "line", "scatter", "area", "bubble", "workbook-bar", "workbook-line", "workbook-scatter", "workbook-bubble"} {
		t.Run(name, func(t *testing.T) {
			workbook := strings.HasPrefix(name, "workbook-")
			family := strings.TrimPrefix(name, "workbook-")
			var source string
			if workbook {
				source = nativeWorkbookChartXML(false, family)
				if family == "bubble" {
					source = nativeBubbleXML(false, true)
				}
			} else {
				switch family {
				case "bar":
					source = nativeBarXML(false, true)
				case "line", "scatter":
					source = nativeConnectedXML(false, family == "scatter")
				case "area":
					source = nativeAreaXML(false, "standard")
				case "bubble":
					source = nativeBubbleXML(false, false)
				}
			}
			source = nativeSeriesOrderXML(source)
			// Give each series an independently identifiable paint and workbook range.
			count := 0
			source = regexp.MustCompile(`(?s)<c:ser(?:\s[^>]*)?>.*?</c:ser>`).ReplaceAllStringFunc(source, func(row string) string {
				order := []int{2, 0, 1}[count]
				count++
				row = strings.ReplaceAll(row, "123456", []string{"E53935", "43A047", "1E88E5"}[order])
				if workbook {
					row = regexp.MustCompile(`<c:f>.*?</c:f>`).ReplaceAllStringFunc(row, func(ref string) string {
						if strings.Contains(ref, "$A$") && family != "scatter" {
							return `<c:f>Data!$A$2:$A$3</c:f>`
						}
						return fmt.Sprintf(`<c:f>Data!$%c$2:$%c$3</c:f>`, 'B'+order, 'B'+order)
					})
					// Bubble sizes use positive column D independently of Y values.
					if family == "bubble" {
						row = regexp.MustCompile(`(?s)<c:bubbleSize>.*?</c:bubbleSize>`).ReplaceAllStringFunc(row, func(string) string {
							return `<c:bubbleSize><c:numRef><c:f>Data!$D$2:$D$3</c:f></c:numRef></c:bubbleSize>`
						})
					}
					// Referenced point count is two, including color overrides.
					row = strings.ReplaceAll(row, `pt idx="2"`, `pt idx="1"`)
				}
				return row
			})
			var input []byte
			if workbook {
				input = nativeWorkbookInspectionFixture(t, false, "column", func(parts map[string]string) {
					parts["relocated/charts/source.xml"] = source
					parts["relocated/embeddings/Data.xlsx"] = string(nativeStackedWorkbookBytes(t, "mixed"))
					parts["relocated/slides/slide-a.xml"] = regexp.MustCompile(`(?s)<p:sp>.*?</p:sp>`).ReplaceAllString(parts["relocated/slides/slide-a.xml"], "")
				})
				inspection, err := InspectNativePPTXChartWorkbooks(input)
				if err != nil || len(inspection.Charts) != 1 {
					t.Fatalf("inspection: %v %+v", err, inspection)
				}
				for i, order := range []int64{2, 0, 1} {
					s := inspection.Charts[0].Source.Series[i]
					if s.Order != order || s.Index != 10+order || !strings.Contains(s.ValueReference.Formula, fmt.Sprintf("$%c$", 'B'+order)) {
						t.Fatal("reference identity changed")
					}
				}
			} else {
				input = nativeChartFixture(t, nativeChartFixtureOptions{omitPreview: true, chartXML: source})
			}
			deck, err := ExtractNativePPTX(input, nativeTestExtractOptions())
			if err != nil || len(deck.Slides) == 0 {
				t.Fatal(err)
			}
			if out := os.Getenv("INJOFFICE_PPTX_SERIES_ORDER_FIXTURES"); out != "" {
				if err = os.MkdirAll(out, 0700); err != nil {
					t.Fatal(err)
				}
				deckJSON, _ := json.Marshal(deck)
				if err = os.WriteFile(filepath.Join(out, name+"-deck.json"), deckJSON, 0600); err != nil {
					t.Fatal(err)
				}
				if workbook {
					inspection, e := InspectNativePPTXChartWorkbooks(input)
					if e != nil {
						t.Fatal(e)
					}
					raw, _ := json.Marshal(inspection)
					if err = os.WriteFile(filepath.Join(out, name+"-inspection.json"), raw, 0600); err != nil {
						t.Fatal(err)
					}
				}
				if err = os.WriteFile(filepath.Join(out, name+".pptx"), input, 0600); err != nil {
					t.Fatal(err)
				}
			}
		})
	}
}
