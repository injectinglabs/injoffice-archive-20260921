package docxpatch

import (
	"bytes"
	"encoding/json"
	"strconv"
	"strings"
	"testing"
)

const nativeApproximateChartTestTheme = `<a:theme xmlns:a="` + drawingMLTransitional + `" name="Office"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme></a:themeElements></a:theme>`

// nativeApproximateChartSeries mirrors the c:ser markup Word writes for a
// clustered column chart with cached string categories and numeric values.
func nativeApproximateChartSeries(index int, title, fill, line string, values []string) string {
	points := ""
	for i, value := range values {
		points += `<c:pt idx="` + itoa(i) + `"><c:v>` + value + `</c:v></c:pt>`
	}
	return `<c:ser><c:idx val="` + itoa(index) + `"/><c:order val="` + itoa(index) + `"/><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>` + title + `</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
		`<c:spPr>` + fill + line + `<a:effectLst/></c:spPr><c:invertIfNegative val="0"/>` +
		`<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$5</c:f><c:strCache><c:ptCount val="4"/><c:pt idx="0"><c:v>Kategória 1</c:v></c:pt><c:pt idx="1"><c:v>Kategória 2</c:v></c:pt><c:pt idx="2"><c:v>Kategória 3</c:v></c:pt><c:pt idx="3"><c:v>Kategória 4</c:v></c:pt></c:strCache></c:strRef></c:cat>` +
		`<c:val><c:numRef><c:f>Sheet1!$B$2:$B$5</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="` + itoa(len(values)) + `"/>` + points + `</c:numCache></c:numRef></c:val></c:ser>`
}

func itoa(value int) string { return strconv.Itoa(value) }

const nativeApproximateChartGreyText = `<c:txPr><a:bodyPr rot="-60000000" vert="horz"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900" b="0" i="0"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:pPr><a:endParaRPr lang="hu-HU"/></a:p></c:txPr>`
const nativeApproximateChartGreyLine = `<c:spPr><a:noFill/><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr>`

func nativeApproximateChartPart(title, plot, legend string) string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="` + nativeChartNSTransitional + `" xmlns:a="` + drawingMLTransitional + `" xmlns:r="` + relNSTransitional + `"><c:date1904 val="0"/><c:lang val="hu-HU"/><c:roundedCorners val="0"/><c:chart>` + title + `<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>` + plot +
		`<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>` + nativeApproximateChartGreyLine + nativeApproximateChartGreyText + `<c:crossAx val="2"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>` +
		`<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines>` + nativeApproximateChartGreyLine + `</c:majorGridlines><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>` + nativeApproximateChartGreyText + `<c:crossAx val="1"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>` +
		`<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr></c:plotArea>` + legend + `<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang="hu-HU"/></a:p></c:txPr><c:externalData r:id="rId3"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>`
}

func nativeApproximateChartBarPlot(series ...string) string {
	return `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>` + strings.Join(series, "") + `<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls><c:gapWidth val="219"/><c:overlap val="-27"/><c:axId val="1"/><c:axId val="2"/></c:barChart>`
}

const nativeApproximateChartAutoTitle = `<c:title><c:overlay val="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr><c:txPr><a:bodyPr rot="0" vert="horz"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0" i="0"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:pPr><a:endParaRPr lang="hu-HU"/></a:p></c:txPr></c:title>`
const nativeApproximateChartBottomLegend = `<c:legend><c:legendPos val="b"/><c:overlay val="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>` + nativeApproximateChartGreyText + `</c:legend>`

func nativeApproximateChartFixturePart() string {
	return nativeApproximateChartPart(nativeApproximateChartAutoTitle, nativeApproximateChartBarPlot(
		nativeApproximateChartSeries(0, "1. adatsor", `<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>`, `<a:ln w="76200"><a:solidFill><a:schemeClr val="accent6"/></a:solidFill><a:prstDash val="sysDot"/></a:ln>`, []string{"4.3", "2.5", "3.5", "4.5"}),
		nativeApproximateChartSeries(1, "2. adatsor", `<a:solidFill><a:schemeClr val="accent2"/></a:solidFill>`, `<a:ln w="76200"><a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill><a:prstDash val="sysDash"/></a:ln>`, []string{"2.4", "4.4000000000000004", "1.8", "2.8"}),
		nativeApproximateChartSeries(2, "3. adatsor", `<a:solidFill><a:schemeClr val="accent3"/></a:solidFill>`, `<a:ln w="76200"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill><a:prstDash val="dash"/></a:ln>`, []string{"2", "2", "3", "5"}),
	), nativeApproximateChartBottomLegend)
}

const nativeApproximateChartInline = `<w:drawing xmlns:wp="` + wordDrawingTransitional + `"><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="5486400" cy="3200400"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Diagram 1"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="` + drawingMLTransitional + `"><a:graphicData uri="` + nativeChartNSTransitional + `"><c:chart xmlns:c="` + nativeChartNSTransitional + `" xmlns:r="` + relNSTransitional + `" r:id="rIdChart"/></a:graphicData></a:graphic></wp:inline></w:drawing>`

// nativeApproximateChartSource builds a package with one chart part, its
// relationship, content type overrides and the Office theme unless mutated.
func nativeApproximateChartSource(t *testing.T, body, chartXML string, mutate func(map[string]string)) []byte {
	parts := nativeMutationParts(nativeMutationMain(body))
	parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/word/charts/chart1.xml" ContentType="`+nativeChartContentType+`"/></Types>`, 1)
	parts["word/_rels/document.xml.rels"] = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="theme" Type="` + relBaseTransitional + `theme" Target="theme/theme1.xml"/><Relationship Id="rIdChart" Type="` + relBaseTransitional + `chart" Target="charts/chart1.xml"/></Relationships>`
	parts["word/theme/theme1.xml"] = nativeApproximateChartTestTheme
	parts["word/charts/chart1.xml"] = chartXML
	if mutate != nil {
		mutate(parts)
	}
	return buildNativeDOCX(t, nativeEntries(parts))
}

func TestApproximateDrawingChartsClusteredColumn(t *testing.T) {
	source := nativeApproximateChartSource(t, `<w:p><w:r><w:rPr><w:noProof/></w:rPr>`+nativeApproximateChartInline+`</w:r><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/></w:p>`, nativeApproximateChartFixturePart(), nil)
	before := append([]byte(nil), source...)
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	if !hasUnsupportedCode(doc, "PICTURE_GRAPHIC_REQUIRED") {
		t.Fatalf("strict extraction must keep refusing the chart: %#v", doc.Unsupported)
	}
	out, err := InspectNativeApproximateDrawingChartsV1(source)
	if err != nil {
		t.Fatal(err)
	}
	if out == nil || len(out.Items) != 1 || out.OmittedCount != 0 || out.Protocol != NativeApproximateDrawingChartsProtocol || out.Policy != NativeApproximateDrawingChartPolicy || out.PackageSHA256 != doc.Source.PackageSHA256 {
		t.Fatalf("unexpected sidecar: %#v", out)
	}
	item := out.Items[0]
	if item.Status != "supported" || item.Placement != "inline" || item.PageAnchor != nil || item.WidthEMU != 5486400 || item.HeightEMU != 3200400 || item.ChartPart != "word/charts/chart1.xml" || !strings.HasPrefix(item.ChartPartSHA256, "sha256:") || item.Chart == nil {
		t.Fatalf("unexpected chart item: %#v", item)
	}
	paragraph := doc.Body.Blocks[0].Paragraph
	if item.ParagraphID != paragraph.ID || len(item.DiagnosticIDs) != 1 || len(paragraph.Runs) != 0 {
		t.Fatalf("chart must join the retained source refusal: %#v", item)
	}
	joined := false
	for _, d := range doc.Unsupported {
		if d.ID == item.DiagnosticIDs[0] && d.ScopeID == paragraph.ID && d.Anchor != nil && *d.Anchor.StartByte >= *item.RunAnchor.StartByte && *d.Anchor.EndByte <= *item.RunAnchor.EndByte {
			joined = true
		}
	}
	if !joined || *item.Anchor.StartByte < *item.RunAnchor.StartByte || *item.Anchor.EndByte > *item.RunAnchor.EndByte {
		t.Fatalf("diagnostic/run anchors do not nest: %#v", item)
	}
	chart := item.Chart
	if chart.Kind != "bar" || chart.BarDirection != "column" || chart.Grouping != "clustered" || chart.GapWidthPercent != 219 || chart.OverlapPercent != -27 {
		t.Fatalf("unexpected bar model: %#v", chart)
	}
	if strings.Join(chart.Categories, "|") != "Kategória 1|Kategória 2|Kategória 3|Kategória 4" || len(chart.Series) != 3 {
		t.Fatalf("unexpected categories/series: %#v", chart)
	}
	first, second, third := chart.Series[0], chart.Series[1], chart.Series[2]
	if first.Title == nil || *first.Title != "1. adatsor" || strings.Join(first.Values, ",") != "4.3,2.5,3.5,4.5" || first.FillRGB != "4472C4" || first.Line == nil || first.Line.RGB != "70AD47" || first.Line.WidthEMU != 76200 || first.Line.Dash != "sysDot" {
		t.Fatalf("first series: %#v %#v", first, first.Line)
	}
	if second.FillRGB != "ED7D31" || second.Line == nil || second.Line.RGB != "FFFF00" || second.Line.Dash != "sysDash" || strings.Join(second.Values, ",") != "2.4,4.4000000000000004,1.8,2.8" {
		t.Fatalf("second series: %#v", second)
	}
	if third.FillRGB != "A5A5A5" || third.Line == nil || third.Line.RGB != "C00000" || third.Line.Dash != "dash" || third.Order != 2 {
		t.Fatalf("third series: %#v", third)
	}
	if chart.CategoryAxis.Deleted || chart.CategoryAxis.Line == nil || chart.CategoryAxis.Line.RGB != "D9D9D9" || chart.CategoryAxis.Line.WidthEMU != 9525 || chart.CategoryAxis.Labels == nil || chart.CategoryAxis.Labels.SizeHundredthPt != 900 || chart.CategoryAxis.Labels.Family != "Calibri" || chart.CategoryAxis.Labels.RGB != "595959" || chart.CategoryAxis.MajorGridlines != nil {
		t.Fatalf("category axis: %#v %#v", chart.CategoryAxis, chart.CategoryAxis.Labels)
	}
	if chart.ValueAxis.Deleted || chart.ValueAxis.Line != nil || chart.ValueAxis.MajorGridlines == nil || chart.ValueAxis.MajorGridlines.RGB != "D9D9D9" || chart.ValueAxis.Labels == nil || chart.ValueAxis.Min != nil || chart.ValueAxis.Max != nil || chart.ValueAxis.MajorUnit != nil || chart.ValueAxis.NumberFormat != "General" {
		t.Fatalf("value axis: %#v", chart.ValueAxis)
	}
	if chart.Title == nil || chart.Title.Text != nil || chart.Title.Font.SizeHundredthPt != 1400 || chart.Title.Font.Family != "Calibri" || chart.Title.Font.RGB != "595959" || chart.Title.Overlay {
		t.Fatalf("title: %#v", chart.Title)
	}
	if chart.Legend == nil || chart.Legend.Position != "b" || chart.Legend.Font.SizeHundredthPt != 900 || chart.Legend.Overlay {
		t.Fatalf("legend: %#v", chart.Legend)
	}
	if chart.AreaFillRGB == nil || *chart.AreaFillRGB != "FFFFFF" || chart.AreaLine == nil || chart.AreaLine.RGB != "D9D9D9" || chart.PlotFillRGB != nil || chart.PlotLine != nil {
		t.Fatalf("chart area paint: %#v", chart)
	}
	notes := strings.Join(item.Notes, "|")
	for _, want := range []string{"automatic title text", "value axis scale is not authored", "major unit is not authored", "theme minor font", "luminance transform approximated"} {
		if !strings.Contains(notes, want) {
			t.Fatalf("missing disclosure %q in %v", want, item.Notes)
		}
	}
	if !bytes.Equal(source, before) {
		t.Fatal("source bytes changed")
	}
	encoded, err := json.Marshal(out)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(encoded), `"chart_part":"word/charts/chart1.xml"`) {
		t.Fatalf("json: %s", encoded)
	}
}

func TestApproximateDrawingChartsExplicitTitleScaleAndAnchor(t *testing.T) {
	title := `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1600" b="1"/></a:pPr><a:r><a:rPr lang="en-US" sz="1600" b="1"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:rPr><a:t>Sales </a:t></a:r><a:r><a:rPr lang="en-US"/><a:t>2024</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`
	part := nativeApproximateChartPart(title, nativeApproximateChartBarPlot(nativeApproximateChartSeries(0, "Only", `<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>`, ``, []string{"1", "-2", "3.5", "0"})), `<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>`)
	part = strings.Replace(part, `<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling>`, `<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="maxMin"/><c:max val="10"/><c:min val="-5"/></c:scaling>`, 1)
	part = strings.Replace(part, `<c:crossBetween val="between"/></c:valAx>`, `<c:crossBetween val="between"/><c:majorUnit val="2.5"/></c:valAx>`, 1)
	anchored := `<w:drawing xmlns:wp="` + wordDrawingTransitional + `"><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>914400</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>457200</wp:posOffset></wp:positionV><wp:extent cx="2743200" cy="1828800"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="2" name="Chart 2"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="` + drawingMLTransitional + `"><a:graphicData uri="` + nativeChartNSTransitional + `"><c:chart xmlns:c="` + nativeChartNSTransitional + `" xmlns:r="` + relNSTransitional + `" r:id="rIdChart"/></a:graphicData></a:graphic></wp:anchor></w:drawing>`
	source := nativeApproximateChartSource(t, `<w:p><w:r><w:t>before</w:t></w:r><w:r>`+anchored+`</w:r><w:r><w:t>after</w:t></w:r></w:p>`, part, nil)
	out, err := InspectNativeApproximateDrawingChartsV1(source)
	if err != nil {
		t.Fatal(err)
	}
	if out == nil || len(out.Items) != 1 {
		t.Fatalf("missing chart: %#v", out)
	}
	item := out.Items[0]
	if item.Status != "supported" || item.Placement != "anchored" || item.Wrap != "square" || item.PageAnchor == nil || item.PageAnchor.HorizontalRelative != "margin" || item.PageAnchor.XEMU != 914400 || item.PageAnchor.VerticalRelative != "paragraph" || item.PageAnchor.YEMU != 457200 || item.WidthEMU != 2743200 {
		t.Fatalf("anchored chart: %#v %#v", item, item.PageAnchor)
	}
	chart := item.Chart
	if chart.Title == nil || chart.Title.Text == nil || *chart.Title.Text != "Sales 2024" || chart.Title.Font.SizeHundredthPt != 1600 || !chart.Title.Font.Bold || chart.Title.Font.RGB != "112233" {
		t.Fatalf("explicit title: %#v", chart.Title)
	}
	if chart.ValueAxis.Orientation != "maxMin" || chart.ValueAxis.Min == nil || *chart.ValueAxis.Min != "-5" || chart.ValueAxis.Max == nil || *chart.ValueAxis.Max != "10" || chart.ValueAxis.MajorUnit == nil || *chart.ValueAxis.MajorUnit != "2.5" {
		t.Fatalf("explicit scale: %#v", chart.ValueAxis)
	}
	if len(chart.Series) != 1 || chart.Series[0].FillRGB != "00FF00" || chart.Series[0].Line != nil || strings.Join(chart.Series[0].Values, ",") != "1,-2,3.5,0" || chart.Legend == nil || chart.Legend.Position != "r" {
		t.Fatalf("series/legend: %#v %#v", chart.Series, chart.Legend)
	}
	notes := strings.Join(item.Notes, "|")
	if strings.Contains(notes, "scale is not authored") || !strings.Contains(notes, "wrapping around the chart") {
		t.Fatalf("notes: %v", item.Notes)
	}
	// The default chart font falls back to the theme minor font when the legend has no txPr.
	if chart.Legend.Font.Family != "Calibri" || chart.Legend.Font.SizeHundredthPt != 1000 || chart.Legend.Font.RGB != "000000" {
		t.Fatalf("default legend font: %#v", chart.Legend.Font)
	}
}

func TestApproximateDrawingChartsOmissions(t *testing.T) {
	base := nativeApproximateChartFixturePart()
	body := `<w:p><w:r>` + nativeApproximateChartInline + `</w:r></w:p>`
	for _, test := range []struct{ name, from, to, reason string }{
		{"pie", `<c:barChart>`, `<c:pieChart>`, "unsupported-chart-type:pieChart"},
		{"stacked", `<c:grouping val="clustered"/>`, `<c:grouping val="stacked"/>`, "unsupported-grouping:stacked"},
		{"missing value cache", `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="4"/><c:pt idx="0"><c:v>4.3</c:v></c:pt><c:pt idx="1"><c:v>2.5</c:v></c:pt><c:pt idx="2"><c:v>3.5</c:v></c:pt><c:pt idx="3"><c:v>4.5</c:v></c:pt></c:numCache>`, ``, "missing-value-cache"},
		{"non numeric cached value", `<c:v>4.3</c:v>`, `<c:v>4,3</c:v>`, "missing-value-cache"},
		{"combination", `<c:catAx>`, `<c:lineChart><c:axId val="1"/><c:axId val="2"/></c:lineChart><c:catAx>`, "combination-chart"},
		{"missing value axis", `<c:valAx>`, `<c:valAx><c:axId val="9"/>`, "missing-axis"},
		{"overlap out of range", `<c:overlap val="-27"/>`, `<c:overlap val="-150"/>`, "invalid-overlap"},
		{"log scale", `<c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>`, `<c:scaling><c:logBase val="10"/><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>`, "unsupported-axis-paint"},
		{"not a chart space", `<c:chartSpace `, `<c:userShapes `, "not-chart-space"},
	} {
		t.Run(test.name, func(t *testing.T) {
			part := strings.Replace(base, test.from, test.to, 1)
			if test.name == "pie" {
				part = strings.Replace(part, `</c:barChart>`, `</c:pieChart>`, 1)
			}
			if test.name == "not a chart space" {
				part = strings.Replace(part, `</c:chartSpace>`, `</c:userShapes>`, 1)
			}
			if part == base {
				t.Fatal("fixture replacement did not apply")
			}
			out, err := InspectNativeApproximateDrawingChartsV1(nativeApproximateChartSource(t, body, part, nil))
			if err != nil {
				t.Fatal(err)
			}
			if out == nil || len(out.Items) != 1 || out.Items[0].Status != "omitted" || out.Items[0].Reason != test.reason || out.Items[0].Chart != nil {
				t.Fatalf("expected omission %q: %#v", test.reason, out)
			}
		})
	}
	t.Run("package refusals", func(t *testing.T) {
		for _, test := range []struct {
			name, reason string
			mutate       func(map[string]string)
		}{
			{"missing relationship", "missing-chart-relationship", func(parts map[string]string) {
				parts["word/_rels/document.xml.rels"] = strings.Replace(parts["word/_rels/document.xml.rels"], `Id="rIdChart"`, `Id="rIdOther"`, 1)
			}},
			{"content type", "chart-content-type-mismatch", func(parts map[string]string) {
				parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], nativeChartContentType, "application/xml", 1)
			}},
			{"external target", "external-chart-relationship", func(parts map[string]string) {
				parts["word/_rels/document.xml.rels"] = strings.Replace(parts["word/_rels/document.xml.rels"], `Target="charts/chart1.xml"`, `Target="https://example.com/chart1.xml" TargetMode="External"`, 1)
			}},
			{"malformed chart xml", "invalid-chart-xml", func(parts map[string]string) {
				parts["word/charts/chart1.xml"] = strings.TrimSuffix(parts["word/charts/chart1.xml"], `</c:chartSpace>`)
			}},
		} {
			out, err := InspectNativeApproximateDrawingChartsV1(nativeApproximateChartSource(t, body, base, test.mutate))
			if err != nil {
				t.Fatal(test.name, err)
			}
			if out == nil || len(out.Items) != 1 || out.Items[0].Status != "omitted" || out.Items[0].Reason != test.reason {
				t.Fatalf("%s: expected %q: %#v", test.name, test.reason, out)
			}
		}
	})
	t.Run("no theme refuses scheme colours", func(t *testing.T) {
		out, err := InspectNativeApproximateDrawingChartsV1(nativeApproximateChartSource(t, body, base, func(parts map[string]string) {
			delete(parts, "word/theme/theme1.xml")
			parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`, ``, 1)
			parts["word/_rels/document.xml.rels"] = strings.Replace(parts["word/_rels/document.xml.rels"], `<Relationship Id="theme" Type="`+relBaseTransitional+`theme" Target="theme/theme1.xml"/>`, ``, 1)
		}))
		if err != nil {
			t.Fatal(err)
		}
		if out == nil || len(out.Items) != 1 || out.Items[0].Status != "omitted" || out.Items[0].Reason != "unsupported-series-fill" {
			t.Fatalf("scheme colours without a theme must refuse: %#v", out)
		}
	})
	t.Run("shared run is omitted", func(t *testing.T) {
		out, err := InspectNativeApproximateDrawingChartsV1(nativeApproximateChartSource(t, `<w:p><w:r><w:t>text</w:t>`+nativeApproximateChartInline+`</w:r></w:p>`, base, nil))
		if err != nil {
			t.Fatal(err)
		}
		if out == nil || len(out.Items) != 1 || out.Items[0].Status != "omitted" || out.Items[0].Reason != "shared-run" {
			t.Fatalf("a run that also carries text must be omitted: %#v", out)
		}
	})
	t.Run("shapes and text are not described", func(t *testing.T) {
		shape := `<w:drawing xmlns:wp="` + wordDrawingTransitional + `" xmlns:a="` + drawingMLTransitional + `" xmlns:wps="` + nativeTextboxWPS + `"><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="4" name="Rectangle 4"/><a:graphic><a:graphicData uri="` + nativeTextboxWPS + `"><wps:wsp><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>`
		for _, paragraph := range []string{`<w:p><w:r>` + shape + `</w:r></w:p>`, `<w:p><w:r><w:t>plain</w:t></w:r></w:p>`} {
			out, err := InspectNativeApproximateDrawingChartsV1(nativeApproximateChartSource(t, paragraph, base, nil))
			if err != nil || out != nil {
				t.Fatalf("expected nil sidecar: %#v %v", out, err)
			}
		}
	})
	t.Run("chart budget", func(t *testing.T) {
		run := `<w:r>` + nativeApproximateChartInline + `</w:r>`
		out, err := InspectNativeApproximateDrawingChartsV1(nativeApproximateChartSource(t, `<w:p>`+strings.Repeat(run, nativeApproximateDrawingChartLimit+2)+`</w:p>`, base, nil))
		if err != nil {
			t.Fatal(err)
		}
		if out == nil || len(out.Items) != nativeApproximateDrawingChartLimit || out.OmittedCount != 2 {
			t.Fatalf("chart budget: %d items, %d omitted", len(out.Items), out.OmittedCount)
		}
	})
}

func TestApproximateDrawingChartsDefaultsAndGaps(t *testing.T) {
	// No explicit series paint, a missing point, a deleted category axis and no
	// legend: defaults are disclosed and gaps stay empty.
	series := `<c:ser><c:idx val="3"/><c:order val="0"/><c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numLit></c:val></c:ser>`
	part := nativeApproximateChartPart(``, nativeApproximateChartBarPlot(series), ``)
	part = strings.Replace(part, `<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>`, `<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/>`, 1)
	out, err := InspectNativeApproximateDrawingChartsV1(nativeApproximateChartSource(t, `<w:p><w:r>`+nativeApproximateChartInline+`</w:r></w:p>`, part, nil))
	if err != nil {
		t.Fatal(err)
	}
	if out == nil || len(out.Items) != 1 || out.Items[0].Status != "supported" {
		t.Fatalf("expected supported chart: %#v", out)
	}
	chart := out.Items[0].Chart
	if len(chart.Series) != 1 || chart.Series[0].Index != 3 || chart.Series[0].Title != nil || strings.Join(chart.Series[0].Values, ",") != "1,,3" || chart.Series[0].FillRGB != "FFC000" || chart.Series[0].Line != nil {
		t.Fatalf("defaulted series: %#v", chart.Series[0])
	}
	if strings.Join(chart.Categories, "|") != "||" || !chart.CategoryAxis.Deleted || chart.CategoryAxis.Labels != nil || chart.CategoryAxis.Line != nil || chart.Title != nil || chart.Legend != nil {
		t.Fatalf("defaults: %#v", chart)
	}
	notes := strings.Join(out.Items[0].Notes, "|")
	if !strings.Contains(notes, "theme accent cycle") || !strings.Contains(notes, "category labels are not cached") {
		t.Fatalf("defaults must be disclosed: %v", out.Items[0].Notes)
	}
}
