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

func nativeRadarZipParts(t *testing.T, input []byte) map[string]string {
	t.Helper()
	z, e := zip.NewReader(bytes.NewReader(input), int64(len(input)))
	if e != nil {
		t.Fatal(e)
	}
	out := map[string]string{}
	for _, f := range z.File {
		r, e := f.Open()
		if e != nil {
			t.Fatal(e)
		}
		b, e := io.ReadAll(r)
		r.Close()
		if e != nil {
			t.Fatal(e)
		}
		out[f.Name] = string(b)
	}
	return out
}
func TestNativeWorkbookRadarFixtures(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("INJOFFICE_PPTX_RADAR_FIXTURES", dir)
	TestNativeLiteralRadarSourceIntegration(t)
	for _, strict := range []bool{false, true} {
		for _, filled := range []bool{false, true} {
			for _, mode := range []string{"plain", "rotated", "grouped", "reversed"} {
				name := fmt.Sprintf("radar-%v-%v-%s", strict, filled, mode)
				original, e := os.ReadFile(filepath.Join(dir, name+".pptx"))
				if e != nil {
					t.Fatal(e)
				}
				parts := nativeRadarZipParts(t, original)
				source := parts["relocated/charts/chart1.xml"]
				at := 0
				source = regexp.MustCompile(`(?s)<c:ser>.*?</c:ser>`).ReplaceAllStringFunc(source, func(s string) string {
					s = regexp.MustCompile(`(?s)<c:cat>.*?</c:cat>`).ReplaceAllString(s, `<c:cat><c:strRef><c:f>Data!A2:A4</c:f><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>POISON</c:v></c:pt><c:pt idx="1"><c:v>POISON</c:v></c:pt><c:pt idx="2"><c:v>POISON</c:v></c:pt></c:strCache></c:strRef></c:cat>`)
					col := string(rune('B' + at))
					at++
					return regexp.MustCompile(`(?s)<c:val>.*?</c:val>`).ReplaceAllString(s, `<c:val><c:numRef><c:f>Data!`+col+`2:`+col+`4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>999</c:v></c:pt><c:pt idx="1"><c:v>999</c:v></c:pt><c:pt idx="2"><c:v>999</c:v></c:pt></c:numCache></c:numRef></c:val>`)
				})
				d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
				source = strings.Replace(source, `</c:chartSpace>`, `<c:externalData xmlns:r="`+d.rels+`" r:id="workbook"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>`, 1)
				wbparts := nativeRadarZipParts(t, nativeStackedWorkbookBytes(t, "mixed"))
				wbparts["xl/worksheets/sheet1.xml"] = strings.NewReplacer(`<v>4</v>`, `<v>-5.00</v>`, `<v>-2</v>`, `<v>5e0</v>`, `<v>3</v>`, `<v>15.0</v>`).Replace(wbparts["xl/worksheets/sheet1.xml"])
				for _, negative := range []string{"", "formula", "outside", "missing"} {
					if negative != "" && (strict || filled || mode != "plain") {
						continue
					}
					wbp := []nativeExtractZipPart{}
					for key, value := range wbparts {
						if key == "xl/worksheets/sheet1.xml" {
							switch negative {
							case "formula":
								value = strings.Replace(value, `<v>-5.00</v>`, `<f>1+1</f><v>-5.00</v>`, 1)
							case "outside":
								value = strings.Replace(value, `<v>-5.00</v>`, `<v>21</v>`, 1)
							case "missing":
								value = strings.Replace(value, `<c r="B2"><v>-5.00</v></c>`, ``, 1)
							}
						}
						wbp = append(wbp, nativeExtractZipPart{name: key, data: value})
					}
					wb := writeNativeExtractZip(t, wbp)
					input := nativeWorkbookInspectionFixture(t, strict, "line", func(p map[string]string) {
						p["relocated/charts/source.xml"] = source
						p["relocated/embeddings/Data.xlsx"] = string(wb)
						p["relocated/slides/slide-a.xml"] = parts["relocated/slides/slide-a.xml"]
					})
					before := bytes.Clone(input)
					inspection, e := InspectNativePPTXChartWorkbooks(input)
					if e != nil || len(inspection.Charts) != 1 || len(inspection.Omissions) != 0 {
						t.Fatalf("radar inspection refused %s: %v %#v", name, e, inspection.Omissions)
					}
					deck, e := ExtractNativePPTX(input, nativeAtomicTestExtractOptions())
					if e != nil {
						t.Fatal(e)
					}
					chart := nativeRadarFindChart(deck.Slides[0].Elements)
					s := inspection.Charts[0].Source
					if chart == nil || chart.Chart.LiteralRadar != nil || s.Family != "radar" || len(s.Series) != 3 || s.Series[0].Order != 2 || s.Series[0].Index != 12 || s.Series[0].ValueReference.Formula != "Data!B2:B4" || !s.Series[0].ValueReference.CachePresent || inspection.Charts[0].ChartSHA256 != nativeSHA256([]byte(source)) {
						t.Fatal("radar authority lost")
					}
					if !bytes.Equal(input, before) {
						t.Fatal("source changed")
					}
					if out := os.Getenv("INJOFFICE_PPTX_WORKBOOK_RADAR_FIXTURES"); out != "" {
						if e = os.MkdirAll(out, 0700); e != nil {
							t.Fatal(e)
						}
						suffix := ""
						if negative != "" {
							suffix = "-" + negative
						}
						dj, _ := json.Marshal(deck)
						ij, _ := json.Marshal(inspection)
						for ext, b := range map[string][]byte{".pptx": input, "-deck.json": dj, "-inspection.json": ij} {
							if e = os.WriteFile(filepath.Join(out, name+suffix+ext), b, 0600); e != nil {
								t.Fatal(e)
							}
						}
					}
				}
			}
		}
	}
}
