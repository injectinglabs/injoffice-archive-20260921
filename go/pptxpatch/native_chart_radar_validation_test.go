package pptxpatch

import (
	"strings"
	"testing"
)

func TestNativeRadarGeometryQualification(t *testing.T) {
	d, _ := nativeDialectForPresentation(xmlNamePresentation(false))
	input := nativeRadarXML(false, false)
	c := extractNativeChartRadarSource([]byte(input), "chart.xml", d, false)
	if c == nil || !validNativeRadarGeometrySource(c) {
		t.Fatal("ordinary radar rejected")
	}
	outside := strings.Replace(input, `<c:v>-1</c:v>`, `<c:v>-11</c:v>`, 1)
	c = extractNativeChartRadarSource([]byte(outside), "chart.xml", d, false)
	if outside == input {
		t.Fatal("fixture value missing")
	}
	if c == nil || validNativeRadarGeometrySource(c) {
		t.Fatal("outside values must retain source but refuse geometry")
	}
}
