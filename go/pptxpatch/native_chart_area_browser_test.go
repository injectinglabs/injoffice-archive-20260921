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

func nativeAreaBrowserXML(name string) string {
	source := nativeAreaXML(false, "standard")
	source = strings.NewReplacer(`<c:max val="20"/>`, `<c:max val="2"/>`, `<c:min val="-10"/>`, `<c:min val="-2"/>`).Replace(source)
	switch name {
	case "percent":
		return strings.NewReplacer(`<c:max val="20"/>`, `<c:max val="1"/>`, `<c:min val="-10"/>`, `<c:min val="0"/>`).Replace(nativeAreaPercentLexemeXML(false))
	case "zero":
		source = strings.NewReplacer(`grouping val="standard"`, `grouping val="percentStacked"`, "<c:v>-1</c:v>", "<c:v>-0</c:v>", "<c:v>2.5</c:v>", "<c:v>0.0</c:v>", "<c:v>1e0</c:v>", "<c:v>0</c:v>").Replace(source)
	case "negative-stack":
		source = strings.Replace(source, `grouping val="standard"`, `grouping val="stacked"`, 1)
	case "stacked":
		source = strings.NewReplacer(`grouping val="standard"`, `grouping val="stacked"`, "<c:v>-1</c:v>", "<c:v>1</c:v>").Replace(source)
		fallthrough
	case "standard":
		start, end := strings.Index(source, "<c:ser>"), strings.Index(source, "</c:ser>")+len("</c:ser>")
		first := source[start:end]
		second := strings.NewReplacer(`idx val="5"`, `idx val="9"`, `order val="0"`, `order val="1"`, "123456", "CC5500", "<c:v>2.5</c:v>", "<c:v>1</c:v>").Replace(first)
		source = source[:start] + second + first + source[end:]
	case "large":
		start, end := strings.Index(source, "<c:cat>"), strings.Index(source, "</c:val>")+len("</c:val>")
		cat, val := `<c:cat><c:strLit><c:ptCount val="256"/>`, `<c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="256"/>`
		for i := 0; i < 256; i++ {
			cat += fmt.Sprintf(`<c:pt idx="%d"><c:v>%d</c:v></c:pt>`, i, i)
			value := "-1"
			if i%2 == 1 {
				value = "1"
			}
			val += fmt.Sprintf(`<c:pt idx="%d"><c:v>%s</c:v></c:pt>`, i, value)
		}
		source = source[:start] + cat + `</c:strLit></c:cat>` + val + `</c:numLit></c:val>` + source[end:]
		source = strings.NewReplacer(`<c:max val="2"/>`, `<c:max val=".5"/>`, `<c:min val="-2"/>`, `<c:min val="-.5"/>`).Replace(source)
	case "labels":
		source = strings.ReplaceAll(source, `<c:tickLblPos val="none"/>`, `<c:tickLblPos val="low"/>`)
		source = strings.ReplaceAll(source, `</c:spPr><c:crossAx`, `</c:spPr>`+nativeAxisTextXML(false)+`<c:crossAx`)
		source = strings.Replace(source, `<c:axPos val="l"/>`, `<c:axPos val="l"/><c:numFmt formatCode="0.0" sourceLinked="0"/>`, 1)
		source = strings.Replace(source, `</c:valAx>`, `<c:majorUnit val="1"/></c:valAx>`, 1)
		source = strings.Replace(source, `</c:catAx>`, `<c:auto val="0"/><c:lblAlgn val="ctr"/><c:lblOffset val="0"/><c:tickLblSkip val="1"/><c:tickMarkSkip val="1"/><c:noMultiLvlLbl val="1"/></c:catAx>`, 1)
	}
	return source
}

func TestNativeAreaBrowserFixtures(t *testing.T) {
	for _, name := range []string{"standard", "stacked", "percent", "zero", "large", "labels", "negative-stack"} {
		t.Run(name, func(t *testing.T) {
			frame := strings.NewReplacer(`y="2000000"`, `y="700000"`, `cx="3000000" cy="2000000"`, `cx="9000000" cy="5000000"`).Replace(nativeChartGraphicFrameXML(false, 3, "Area "+name, ""))
			input := nativeChartFixture(t, nativeChartFixtureOptions{omitPreview: true, chartXML: nativeAreaBrowserXML(name), frameXML: frame})
			archive, err := zip.NewReader(bytes.NewReader(input), int64(len(input)))
			if err != nil {
				t.Fatal(err)
			}
			parts := []nativeExtractZipPart{}
			for _, entry := range archive.File {
				reader, err := entry.Open()
				if err != nil {
					t.Fatal(err)
				}
				payload, err := io.ReadAll(reader)
				reader.Close()
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
			chart := nativeFixtureChart(t, deck.Slides[0])
			if chart.Chart.LiteralArea == nil {
				t.Fatal("fixture qualification changed")
			}
			if output := os.Getenv("INJOFFICE_PPTX_AREA_FIXTURES"); output != "" {
				if err = os.MkdirAll(output, 0700); err != nil {
					t.Fatal(err)
				}
				if err = os.WriteFile(filepath.Join(output, name+".pptx"), input, 0600); err != nil {
					t.Fatal(err)
				}
			}
		})
	}
}
