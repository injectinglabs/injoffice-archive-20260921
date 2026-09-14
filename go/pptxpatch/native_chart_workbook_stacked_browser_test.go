package pptxpatch

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"testing"
)

// Preserve the ordinary workbook metadata fixture while replacing only its
// authored sheetData/dimension with explicit signed test cells.
func nativeStackedWorkbookBytes(t *testing.T, mode string) []byte {
	t.Helper()
	original, err := os.ReadFile("../xlsxpatch/testdata/native-get-corpus/pass-excel-defaults.xlsx")
	if err != nil {
		t.Fatal(err)
	}
	archive, err := zip.NewReader(bytes.NewReader(original), int64(len(original)))
	if err != nil {
		t.Fatal(err)
	}
	values := []string{"4", "-2", "3", "-1"}
	if mode == "negative" {
		values = []string{"-2", "-3", "-1", "-4"}
	}
	if mode == "cancelling" {
		values = []string{"2", "-2", "3", "-3"}
	}
	if mode == "zero" {
		values = []string{"0", "0", "0", "0"}
	}
	sheet := "<sheetData>"
	for row := 2; row <= 4; row++ {
		sheet += fmt.Sprintf(`<row r="%d"><c r="A%d" t="inlineStr"><is><t>%c</t></is></c>`, row, row, 'A'+row-2)
		for i, value := range values {
			sheet += fmt.Sprintf(`<c r="%c%d"><v>%s</v></c>`, 'B'+i, row, value)
		}
		sheet += "</row>"
	}
	sheet += "</sheetData>"
	parts := []nativeExtractZipPart{}
	for _, entry := range archive.File {
		r, err := entry.Open()
		if err != nil {
			t.Fatal(err)
		}
		raw, err := io.ReadAll(r)
		r.Close()
		if err != nil {
			t.Fatal(err)
		}
		data := string(raw)
		if entry.Name == "xl/worksheets/sheet1.xml" {
			data = strings.Replace(data, `ref="A1:C5"`, `ref="A2:E4"`, 1)
			data = regexp.MustCompile(`(?s)<sheetData>.*?</sheetData>`).ReplaceAllString(data, sheet)
		}
		parts = append(parts, nativeExtractZipPart{name: entry.Name, data: data})
	}
	return writeNativeExtractZip(t, parts)
}
func TestNativeWorkbookStackedBrowserFixtures(t *testing.T) {
	for _, family := range []string{"column", "bar", "line"} {
		for _, grouping := range []string{"stacked", "percentStacked"} {
			for _, mode := range []string{"mixed", "negative", "cancelling", "zero", "permuted", "labels"} {
				name := family + "-" + grouping + "-" + mode
				t.Run(name, func(t *testing.T) {
					source := nativeWorkbookStackedXML(false, family, grouping)
					first := nativeWorkbookSeriesXML(false, family == "line", false)
					series := []string{}
					for i := 0; i < 4; i++ {
						oldIndex := "7"
						if family == "line" {
							oldIndex = "5"
						}
						item := strings.NewReplacer(`idx val="`+oldIndex+`"`, fmt.Sprintf(`idx val="%d"`, 10+i), `order val="0"`, fmt.Sprintf(`order val="%d"`, i), `Sheet1!$A$2:$A$3`, `Data!$A$2:$A$4`, `Sheet1!$A$2:$A$4`, `Data!$A$2:$A$4`, `Sheet1!$B$2:$B$3`, fmt.Sprintf(`Data!$%c$2:$%c$4`, 'B'+i, 'B'+i), `Sheet1!$B$2:$B$4`, fmt.Sprintf(`Data!$%c$2:$%c$4`, 'B'+i, 'B'+i), "123456", []string{"E53935", "43A047", "1E88E5", "8E24AA"}[i]).Replace(first)
						series = append(series, item)
					}
					if mode == "permuted" {
						series[0], series[3] = series[3], series[0]
						series[1], series[2] = series[2], series[1]
					}
					source = strings.Replace(source, first, strings.Join(series, ""), 1)
					source = strings.ReplaceAll(source, `<c:max val="20"/>`, `<c:max val="10"/>`)
					if grouping == "percentStacked" {
						source = strings.NewReplacer(`<c:max val="10"/>`, `<c:max val="1"/>`, `<c:min val="-10"/>`, `<c:min val="-1"/>`).Replace(source)
					}
					if mode == "labels" {
						source = strings.ReplaceAll(source, `<c:tickLblPos val="none"/>`, `<c:tickLblPos val="low"/>`)
						source = strings.ReplaceAll(source, `</c:spPr><c:crossAx`, `</c:spPr>`+nativeAxisTextXML(false)+`<c:crossAx`)
						position := "l"
						if family == "bar" {
							position = "b"
						}
						source = strings.Replace(source, `<c:axPos val="`+position+`"/>`, `<c:axPos val="`+position+`"/><c:numFmt formatCode="0.0" sourceLinked="0"/>`, 1)
						unit := "5"
						if grouping == "percentStacked" {
							unit = ".5"
						}
						source = strings.Replace(source, `</c:valAx>`, `<c:majorUnit val="`+unit+`"/></c:valAx>`, 1)
						source = strings.Replace(source, `</c:catAx>`, `<c:auto val="0"/><c:lblAlgn val="ctr"/><c:lblOffset val="0"/><c:tickLblSkip val="1"/><c:tickMarkSkip val="1"/><c:noMultiLvlLbl val="1"/></c:catAx>`, 1)
					}
					workbook := nativeStackedWorkbookBytes(t, mode)
					input := nativeWorkbookInspectionFixture(t, false, family, func(parts map[string]string) {
						parts["relocated/charts/source.xml"] = source
						parts["relocated/embeddings/Data.xlsx"] = string(workbook)
						slide := regexp.MustCompile(`(?s)<p:sp>.*?</p:sp>`).ReplaceAllString(parts["relocated/slides/slide-a.xml"], "")
						parts["relocated/slides/slide-a.xml"] = strings.NewReplacer(`y="2000000"`, `y="700000"`, `cx="3000000" cy="2000000"`, `cx="9000000" cy="5000000"`).Replace(slide)
					})
					before := append([]byte(nil), input...)
					inspection, err := InspectNativePPTXChartWorkbooks(input)
					if err != nil || len(inspection.Charts) != 1 || len(inspection.Workbooks) != 1 || len(inspection.Omissions) != 0 {
						t.Fatalf("fixture inspection refused: %v %#v", err, inspection.Omissions)
					}
					if !bytes.Equal(before, input) || inspection.Charts[0].Source.Grouping != grouping || len(inspection.Charts[0].Source.Series) != 4 {
						t.Fatal("source grouping/bytes drift")
					}
					if out := os.Getenv("INJOFFICE_PPTX_WORKBOOK_STACKED_FIXTURES"); out != "" {
						if err = os.MkdirAll(out, 0700); err != nil {
							t.Fatal(err)
						}
						if err = os.WriteFile(filepath.Join(out, name+".pptx"), input, 0600); err != nil {
							t.Fatal(err)
						}
					}
				})
			}
		}
	}
}
