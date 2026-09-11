package pptxpatch

import (
	"bytes"
	"fmt"
	"image"
	"image/color"
	"image/png"
	"os"
	"strings"
	"testing"
)

func TestNativeMeasuredPreviewBrowserFixture(t *testing.T) {
	output := os.Getenv("INJOFFICE_PPTX_PREVIEW_FIXTURE")
	if output == "" {
		t.Skip("optional measured preview source export")
	}
	for _, withBullets := range []bool{false, true} {
		input := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
			part := "relocated/slides/slide-a.xml"
			slide := parts[part]
			start, end := strings.Index(slide, "<p:sp>"), strings.Index(slide, "</p:sp>")+len("</p:sp>")
			shape := slide[start:end]
			shape = strings.ReplaceAll(shape, `typeface="Aptos"`, `typeface="DejaVu Sans"`)
			shape = strings.ReplaceAll(shape, `b="1"`, `b="0"`)
			shape = strings.Replace(shape, `sz="3200"`, `sz="1200"`, 1)
			shape = strings.Replace(shape, `sz="3200"`, `sz="2400"`, 1)
			shape = strings.Replace(shape, "Hello ", "Small ", 1)
			shape = strings.Replace(shape, "world", "large", 1)
			shapes := ""
			for i, anchor := range []string{"t", "ctr", "b"} {
				item := strings.Replace(shape, `id="2"`, fmt.Sprintf(`id="%d"`, i+2), 1)
				item = strings.Replace(item, `x="914400" y="457200"`, fmt.Sprintf(`x="%d" y="1000000"`, 500000+i*3900000), 1)
				item = strings.Replace(item, `cx="4572000" cy="914400"`, `cx="3300000" cy="3000000"`, 1)
				item = strings.Replace(item, `<a:bodyPr/>`, `<a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" wrap="none" anchor="`+anchor+`"/>`, 1)
				if withBullets {
					item = strings.Replace(item, `<a:pPr algn="ctr" lvl="0"><a:buNone/></a:pPr>`, `<a:pPr algn="l" lvl="2" marL="400000" indent="-300000"><a:buChar char="▪"/></a:pPr>`, 1)
					item = strings.Replace(item, `wrap="none"`, `wrap="square"`, 1)
					item = strings.Replace(item, `<a:t>large</a:t>`, `<a:t>large repeated text wraps here</a:t>`, 1)
				}
				shapes += item
			}
			parts[part] = slide[:start] + shapes + slide[end:]
		}})
		if _, err := ExtractNativePPTX(input, nativeTestExtractOptions()); err != nil {
			t.Fatal(err)
		}
		name := output
		if withBullets {
			name = strings.TrimSuffix(output, ".pptx") + "-bullets.pptx"
		}
		if err := os.WriteFile(name, input, 0600); err != nil {
			t.Fatal(err)
		}
	}
	quadrants := image.NewRGBA(image.Rect(0, 0, 80, 80))
	colors := []color.RGBA{{255, 0, 0, 255}, {0, 255, 0, 255}, {0, 0, 255, 255}, {255, 255, 0, 255}}
	for y := 0; y < 80; y++ {
		for x := 0; x < 80; x++ {
			quadrants.SetRGBA(x, y, colors[(y/40)*2+x/40])
		}
	}
	var encoded bytes.Buffer
	if err := png.Encode(&encoded, quadrants); err != nil {
		t.Fatal(err)
	}
	picture := nativePictureFixture(t, nativePictureFixtureOptions{imageData: encoded.String(), sourceRect: `<a:srcRect l="50000" b="50000"/>`})
	part := "relocated/slides/slide-a.xml"
	slide := string(chartZipEntry(t, picture, part))
	start, end := strings.Index(slide, "<p:sp>"), strings.Index(slide, "</p:sp>")+len("</p:sp>")
	picture = replaceChartZipEntry(t, picture, part, []byte(slide[:start]+slide[end:]))
	if err := os.WriteFile(strings.TrimSuffix(output, ".pptx")+"-crop.pptx", picture, 0600); err != nil {
		t.Fatal(err)
	}
	arrows := ""
	for i, kind := range []string{"triangle", "stealth", "diamond", "oval", "arrow"} {
		line := nativeAutoShapeSolidLine("120000", "flat", `<a:round/>`, "2255aa")
		line = strings.Replace(line, `</a:ln>`, `<a:tailEnd type="`+kind+`" w="lg" len="lg"/></a:ln>`, 1)
		connector := nativeConnectorXML(i+10, kind, `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>`, line, "", "", "", "")
		connector = strings.Replace(connector, fmt.Sprintf(`<a:off x="%d" y="%d"/>`, (i+10)*100000, (i+10)*50000), fmt.Sprintf(`<a:off x="1000000" y="%d"/>`, 800000+i*1100000), 1)
		connector = strings.Replace(connector, `cx="1000000" cy="500000"`, `cx="3000000" cy="200000"`, 1)
		arrows += connector
	}
	arrowSource := nativeConnectorFixture(t, false, arrows)
	slide = string(chartZipEntry(t, arrowSource, part))
	start, end = strings.Index(slide, "<p:sp>"), strings.Index(slide, "</p:sp>")+len("</p:sp>")
	arrowSource = replaceChartZipEntry(t, arrowSource, part, []byte(slide[:start]+slide[end:]))
	if err := os.WriteFile(strings.TrimSuffix(output, ".pptx")+"-arrows.pptx", arrowSource, 0600); err != nil {
		t.Fatal(err)
	}
	group := `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Chart preview group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="6858000"/><a:chOff x="0" y="0"/><a:chExt cx="12192000" cy="6858000"/></a:xfrm></p:grpSpPr>` + nativeChartGraphicFrameXML(false, 4, "Cached chart preview", "") + `</p:grpSp>`
	chart := nativeChartFixture(t, nativeChartFixtureOptions{previewData: encoded.String(), frameXML: group})
	slide = string(chartZipEntry(t, chart, part))
	start, end = strings.Index(slide, "<p:sp>"), strings.Index(slide, "</p:sp>")+len("</p:sp>")
	chart = replaceChartZipEntry(t, chart, part, []byte(slide[:start]+slide[end:]))
	deck, err := ExtractNativePPTX(chart, nativeAtomicTestExtractOptions())
	if err != nil || len(deck.Slides) != 1 || len(deck.Slides[0].Elements) != 1 || deck.Slides[0].Elements[0].Kind != NativeElementKindGroup || len(deck.Slides[0].Elements[0].Children) != 1 || deck.Slides[0].Elements[0].Children[0].Kind != NativeElementKindChart {
		t.Fatalf("chart-only fixture drift: %v", err)
	}
	if err := os.WriteFile(strings.TrimSuffix(output, ".pptx")+"-chart.pptx", chart, 0600); err != nil {
		t.Fatal(err)
	}
}
