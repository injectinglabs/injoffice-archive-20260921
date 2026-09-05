package pptxpatch

import (
	"fmt"
	"strings"
	"testing"
)

func TestExtractNativePPTXPaintsExactChartPreviewPicture(t *testing.T) {
	t.Parallel()
	for _, strict := range []bool{false, true} {
		strict := strict
		t.Run(map[bool]string{false: "transitional", true: "strict"}[strict], func(t *testing.T) {
			t.Parallel()
			preview := "\x89PNG\r\n\x1a\nchart-preview"
			deck, err := ExtractNativePPTX(nativeChartFixture(t, nativeChartFixtureOptions{strict: strict, previewData: preview}), nativeTestExtractOptions())
			if err != nil {
				t.Fatalf("extract chart preview: %v", err)
			}
			chart := nativeFixtureChart(t, deck.Slides[0])
			if chart.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || chart.Chart == nil || chart.Chart.PreviewAssetID == nil || chart.Source == nil || chart.Source.RelationshipID == nil || *chart.Source.RelationshipID != "rIdChart" {
				t.Fatalf("chart preview was not projected: %#v", chart)
			}
			if chart.Transform.X == nil || *chart.Transform.X != 1_000_000 || chart.Transform.Y == nil || *chart.Transform.Y != 2_000_000 || chart.Transform.Cx == nil || *chart.Transform.Cx != 3_000_000 || chart.Transform.Cy == nil || *chart.Transform.Cy != 2_000_000 {
				t.Fatalf("chart EMU changed: %#v", chart.Transform)
			}
			if chart.Chart.RelationshipID != "rIdChart" || chart.Chart.ChartPart != "relocated/charts/chart1.xml" || chart.Chart.OpaqueRef.OwnerPart != chart.Chart.ChartPart {
				t.Fatalf("opaque chart graph was not bound: %#v", chart.Chart)
			}
			if len(deck.Assets) != 1 || deck.Assets[0].ID != *chart.Chart.PreviewAssetID || deck.Assets[0].ContentType != "image/png" || deck.Assets[0].SHA256 != nativeSHA256([]byte(preview)) || deck.Assets[0].ByteLength == nil || *deck.Assets[0].ByteLength != int64(len(preview)) {
				t.Fatalf("preview bytes were not the exact package picture: %#v", deck.Assets)
			}
			if issues := ValidateNativePPTX(deck); len(issues) != 0 {
				t.Fatalf("invalid chart deck: %#v", issues)
			}
		})
	}
}

func TestExtractNativePPTXChartWithoutPreviewStaysOpaque(t *testing.T) {
	t.Parallel()
	deck, err := ExtractNativePPTX(nativeChartFixture(t, nativeChartFixtureOptions{omitPreview: true}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract chart without preview: %v", err)
	}
	chart := nativeFixtureChart(t, deck.Slides[0])
	if chart.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || chart.Chart == nil || chart.Chart.PreviewAssetID != nil || len(deck.Assets) != 0 {
		t.Fatalf("missing preview invented a chart renderer: %#v assets=%#v", chart, deck.Assets)
	}
}

func TestExtractNativePPTXAmbiguousChartImagesDoNotGuessAPreview(t *testing.T) {
	t.Parallel()
	deck, err := ExtractNativePPTX(nativeChartFixture(t, nativeChartFixtureOptions{secondPreview: true}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract ambiguous chart images: %v", err)
	}
	chart := nativeFixtureChart(t, deck.Slides[0])
	if chart.Chart == nil || chart.Chart.PreviewAssetID != nil {
		t.Fatalf("multiple chart images were guessed as a preview: %#v", chart)
	}
}

func TestExtractNativePPTXUnknownChartArrowMarkupRemainsRefused(t *testing.T) {
	t.Parallel()
	raw := nativeChartGraphicFrameXML(false, 3, "Broken chart", ` foo="bar"`)
	deck, err := ExtractNativePPTX(nativeChartFixture(t, nativeChartFixtureOptions{frameXML: raw}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract malformed chart: %v", err)
	}
	for _, element := range deck.Slides[0].Elements {
		if element.Kind == NativeElementKindChart {
			t.Fatalf("malformed chart leaked a projection: %#v", element)
		}
	}
	if deck.Slides[0].Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
		t.Fatalf("malformed chart did not preserve the slide: %#v", deck.Slides[0].Compatibility)
	}
}

func TestExtractNativePPTXGroupedChartPreviewRetainsFrameEMU(t *testing.T) {
	t.Parallel()
	group := `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Chart group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="400" cy="500"/><a:chOff x="0" y="0"/><a:chExt cx="400" cy="500"/></a:xfrm></p:grpSpPr>` +
		nativeChartGraphicFrameXML(false, 4, "Grouped chart", "") + `</p:grpSp>`
	deck, err := ExtractNativePPTX(nativeChartFixture(t, nativeChartFixtureOptions{frameXML: group}), nativeAtomicTestExtractOptions())
	if err != nil {
		t.Fatalf("extract grouped chart: %v", err)
	}
	var projected *NativeElement
	for index := range deck.Slides[0].Elements {
		if deck.Slides[0].Elements[index].Kind == NativeElementKindGroup {
			projected = &deck.Slides[0].Elements[index]
		}
	}
	if projected == nil || len(projected.Children) != 1 || projected.Children[0].Kind != NativeElementKindChart || projected.Children[0].Chart == nil || projected.Children[0].Chart.PreviewAssetID == nil {
		t.Fatalf("grouped chart preview was not projected: %#v", projected)
	}
	child := projected.Children[0]
	if *child.Transform.X != 1_000_000 || *child.Transform.Cx != 3_000_000 || child.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
		t.Fatalf("grouped chart EMU/status changed: %#v", child)
	}
}

type nativeChartFixtureOptions struct {
	strict        bool
	previewData   string
	omitPreview   bool
	secondPreview bool
	frameXML      string
}

func nativeChartFixture(t *testing.T, options nativeChartFixtureOptions) []byte {
	t.Helper()
	chartType := relChartTransitional
	imageType := relImageTransitional
	chartNS := nsChartTransitional
	if options.strict {
		chartType = relChartStrict
		imageType = relImageStrict
		chartNS = nsChartStrict
	}
	preview := options.previewData
	if preview == "" {
		preview = "\x89PNG\r\n\x1a\npreview"
	}
	chartPart := "relocated/charts/chart1.xml"
	previewPart := "relocated/media/chart-preview.png"
	secondPart := "relocated/media/chart-extra.png"
	extra := []nativeExtractZipPart{
		{name: chartPart, data: `<c:chartSpace xmlns:c="` + chartNS + `"><c:chart><c:plotArea/></c:chart></c:chartSpace>`},
	}
	chartRels := `<Relationships xmlns="` + nsPackageRels + `">`
	if !options.omitPreview {
		extra = append(extra, nativeExtractZipPart{name: previewPart, data: preview})
		chartRels += fmt.Sprintf(`<Relationship Id="rIdPreview" Type="%s" Target="../media/chart-preview.png"/>`, imageType)
	}
	if options.secondPreview {
		extra = append(extra, nativeExtractZipPart{name: secondPart, data: preview + "-extra"})
		chartRels += fmt.Sprintf(`<Relationship Id="rIdExtra" Type="%s" Target="../media/chart-extra.png"/>`, imageType)
	}
	chartRels += `</Relationships>`
	if !options.omitPreview || options.secondPreview {
		extra = append(extra, nativeExtractZipPart{name: "relocated/charts/_rels/chart1.xml.rels", data: chartRels})
	}
	frame := options.frameXML
	if frame == "" {
		frame = nativeChartGraphicFrameXML(options.strict, 3, "Chart 1", "")
	}
	return nativeExtractFixture(t, nativeExtractFixtureOptions{
		strict: options.strict, extraParts: extra,
		mutate: func(parts map[string]string) {
			parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`,
				fmt.Sprintf(`<Override PartName="/%s" ContentType="%s"/>`, chartPart, contentTypeChart)+func() string {
					suffix := ""
					if !options.omitPreview {
						suffix += fmt.Sprintf(`<Override PartName="/%s" ContentType="image/png"/>`, previewPart)
					}
					if options.secondPreview {
						suffix += fmt.Sprintf(`<Override PartName="/%s" ContentType="image/png"/>`, secondPart)
					}
					return suffix
				}()+`</Types>`, 1)
			parts["relocated/slides/_rels/slide-a.xml.rels"] = strings.Replace(parts["relocated/slides/_rels/slide-a.xml.rels"], `</Relationships>`,
				fmt.Sprintf(`<Relationship Id="rIdChart" Type="%s" Target="../charts/chart1.xml"/></Relationships>`, chartType), 1)
			parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, frame+`</p:spTree>`, 1)
		},
	})
}

func nativeChartGraphicFrameXML(strict bool, id int, name, extraAttrs string) string {
	chartNS := nsChartTransitional
	relsNS := nsOfficeRelsTransitional
	if strict {
		chartNS = nsChartStrict
		relsNS = nsOfficeRelsStrict
	}
	return fmt.Sprintf(`<p:graphicFrame%s><p:nvGraphicFramePr><p:cNvPr id="%d" name="%s"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="1000000" y="2000000"/><a:ext cx="3000000" cy="2000000"/></p:xfrm><a:graphic><a:graphicData uri="%s"><c:chart xmlns:c="%s" xmlns:r="%s" r:id="rIdChart"/></a:graphicData></a:graphic></p:graphicFrame>`, extraAttrs, id, name, chartNS, chartNS, relsNS)
}

func nativeFixtureChart(t *testing.T, slide NativeSlide) NativeElement {
	t.Helper()
	for _, element := range slide.Elements {
		if element.Kind == NativeElementKindChart {
			return element
		}
	}
	t.Fatalf("slide has no native chart: %#v", slide)
	return NativeElement{}
}
