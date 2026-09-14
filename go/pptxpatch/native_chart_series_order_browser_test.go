package pptxpatch

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"testing"
)

// Derive separate chart-only packages without rewriting the original fixture.
// Label variants are authored in XML before extraction, never in private models.
func nativeSeriesOrderChartOnly(t *testing.T, original []byte, labels, invalidLabels bool) []byte {
	archive, err := zip.NewReader(bytes.NewReader(original), int64(len(original)))
	if err != nil {
		t.Fatal(err)
	}
	parts := []nativeExtractZipPart{}
	for _, entry := range archive.File {
		r, e := entry.Open()
		if e != nil {
			t.Fatal(e)
		}
		raw, e := io.ReadAll(r)
		r.Close()
		if e != nil {
			t.Fatal(e)
		}
		data := string(raw)
		if entry.Name == "relocated/slides/slide-a.xml" {
			data = regexp.MustCompile(`(?s)<p:sp>.*?</p:sp>`).ReplaceAllString(data, "")
		}

		// Original workbook-scatter samples repeat identical XY points.
		// Only this separately named chart-only derivative selects two different
		// saved X cells, retaining per-series Y references and all stale caches.
		if strings.Contains(data, "<c:scatterChart>") && strings.Contains(data, "<c:externalData") {
			data = regexp.MustCompile(`(?s)<c:xVal>.*?</c:xVal>`).ReplaceAllStringFunc(data, func(x string) string {
				return regexp.MustCompile(`<c:f>.*?</c:f>`).ReplaceAllStringFunc(x, func(string) string { return `<c:f>Data!$B$2:$C$2</c:f>` })
			})
		}
		if labels && strings.HasPrefix(entry.Name, "relocated/charts/") && strings.HasSuffix(entry.Name, ".xml") {
			data = strings.ReplaceAll(data, `<c:tickLblPos val="none"/>`, `<c:tickLblPos val="low"/>`)
			data = strings.ReplaceAll(data, `</c:spPr><c:crossAx`, `</c:spPr>`+nativeAxisTextXML(false)+`<c:crossAx`)
			data = regexp.MustCompile(`(?s)<c:valAx>.*?</c:valAx>`).ReplaceAllStringFunc(data, func(axis string) string {
				axis = strings.Replace(axis, `<c:majorTickMark`, `<c:numFmt formatCode="0.0" sourceLinked="0"/><c:majorTickMark`, 1)
				return strings.Replace(axis, `</c:valAx>`, `<c:majorUnit val="5"/></c:valAx>`, 1)
			})
			data = strings.ReplaceAll(data, `</c:catAx>`, `<c:auto val="0"/><c:lblAlgn val="ctr"/><c:lblOffset val="0"/><c:tickLblSkip val="1"/><c:tickMarkSkip val="1"/><c:noMultiLvlLbl val="1"/></c:catAx>`)
		}
		if invalidLabels {
			data = strings.ReplaceAll(data, `tickLblPos val="low"`, `tickLblPos val="nextTo"`)
		}
		parts = append(parts, nativeExtractZipPart{name: entry.Name, data: data})
	}
	return writeNativeExtractZip(t, parts)
}

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
			original := bytes.Clone(input)
			for _, variant := range []string{"", "-chart-only", "-chart-only-labels", "-chart-only-invalid-labels"} {
				candidate := input
				if variant != "" {
					candidate = nativeSeriesOrderChartOnly(t, input, strings.HasSuffix(variant, "labels"), strings.HasSuffix(variant, "invalid-labels"))
				}
				invalidLabels := strings.HasSuffix(variant, "invalid-labels")
				name := name + variant
				deck, err := ExtractNativePPTX(candidate, nativeTestExtractOptions())
				if err != nil || len(deck.Slides) == 0 {
					t.Fatal(err)
				}
				if variant != "" && len(deck.Slides[0].Elements) != 1 {
					t.Fatal("derived source is not chart-only")
				}
				if !workbook {
					raw, _ := json.Marshal(nativeFixtureChart(t, deck.Slides[0]).Chart)
					if strings.Contains(string(raw), `"profile":"literal-`) == invalidLabels {
						t.Fatalf("source label qualification drift: %s", raw)
					}
				}
				if workbook {
					inspection, e := InspectNativePPTXChartWorkbooks(candidate)
					if e != nil || (len(inspection.Charts) == 1) == invalidLabels {
						t.Fatal("workbook source label qualification drift", e)
					}
				}
				if out := os.Getenv("INJOFFICE_PPTX_SERIES_ORDER_FIXTURES"); out != "" {
					if err = os.MkdirAll(out, 0700); err != nil {
						t.Fatal(err)
					}
					identity, _ := json.Marshal(map[string]string{"originalPackageSHA256": nativeSHA256(original), "packageSHA256": nativeSHA256(candidate), "variant": variant})
					if err = os.WriteFile(filepath.Join(out, name+"-source-identity.json"), identity, 0600); err != nil {
						t.Fatal(err)
					}
					deckJSON, _ := json.Marshal(deck)
					if err = os.WriteFile(filepath.Join(out, name+"-deck.json"), deckJSON, 0600); err != nil {
						t.Fatal(err)
					}
					if workbook {
						inspection, e := InspectNativePPTXChartWorkbooks(candidate)
						if e != nil {
							t.Fatal(e)
						}
						raw, _ := json.Marshal(inspection)
						if err = os.WriteFile(filepath.Join(out, name+"-inspection.json"), raw, 0600); err != nil {
							t.Fatal(err)
						}
					}
					if err = os.WriteFile(filepath.Join(out, name+".pptx"), candidate, 0600); err != nil {
						t.Fatal(err)
					}
				}
			}
			if !bytes.Equal(input, original) {
				t.Fatal("original fixture mutated")
			}
		})
	}
}
