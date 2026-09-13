package pptxpatch

import (
	"strings"
	"testing"
)

func nativeAxisTextXML(strict bool) string {
	c, a := nsChartTransitional, nsDrawingTransitional
	if strict {
		c, a = nsChartStrict, nsDrawingStrict
	}
	return `<c:txPr xmlns:c="` + c + `" xmlns:a="` + a + `"><a:bodyPr rot="0" vert="horz" wrap="none" lIns="0" rIns="0" tIns="0" bIns="0" anchor="t" anchorCtr="0" horzOverflow="overflow" vertOverflow="overflow"><a:noAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="ctr"><a:defRPr b="0" i="0" sz="1200" lang="en-US" kern="0"><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:latin typeface="DejaVu Sans"/><a:ea typeface="DejaVu Sans"/><a:cs typeface="DejaVu Sans"/></a:defRPr></a:pPr></a:p></c:txPr>`
}
func TestNativeChartAxisCompleteTextStyle(t *testing.T) {
	for _, strict := range []bool{false, true} {
		d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
		node, e := parseNativeXML([]byte(nativeAxisTextXML(strict)), "axis.xml")
		if e != nil {
			t.Fatal(e)
		}
		style, ok := extractNativeChartAxisLabelStyle(node, d)
		if !ok || style.FontFamily != "DejaVu Sans" || style.FontSize != 1200 || style.Color != "123456" || style.Language != "en-US" || style.Bold || style.Italic {
			t.Fatalf("missing complete style: %+v", style)
		}
		for _, pair := range [][2]string{{` kern="0"`, ``}, {` sz="1200"`, ``}, {`b="0"`, `b="yes"`}, {`wrap="none"`, `wrap="square"`}, {`rot="0"`, `rot="60000"`}, {`<a:noAutofit/>`, `<a:spAutoFit/>`}, {`<a:ea typeface="DejaVu Sans"/>`, ``}, {`<a:cs typeface="DejaVu Sans"/>`, `<a:cs typeface="Other"/>`}, {`typeface="DejaVu Sans"`, `typeface="+mn-lt"`}, {`srgbClr val="123456"`, `schemeClr val="accent1"`}, {`</a:p>`, `<a:endParaRPr sz="2000"/></a:p>`}, {`algn="ctr"`, `algn="r"`}, {`lang="en-US"`, `lang="en_US"`}} {
			node, e := parseNativeXML([]byte(strings.ReplaceAll(nativeAxisTextXML(strict), pair[0], pair[1])), "bad.xml")
			if e != nil {
				t.Fatal(e)
			}
			if _, ok := extractNativeChartAxisLabelStyle(node, d); ok {
				t.Fatalf("unqualified source style accepted: %v", pair)
			}
		}
	}
}
