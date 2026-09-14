package pptxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func TestNativeChartStackedLineSource(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, grouping := range []string{"stacked", "percentStacked"} {
			d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
			source := strings.Replace(nativeConnectedXML(strict, false), `grouping val="standard"`, `grouping val="`+grouping+`"`, 1)
			first := nativeConnectedSeriesXML(false)
			second := strings.NewReplacer(`idx val="5"`, `idx val="9"`, `order val="0"`, `order val="1"`, `123456`, `ABCDEF`).Replace(first)
			source = strings.Replace(source, first, second+first, 1)
			payload := []byte(source)
			before := append([]byte(nil), payload...)
			chart := extractNativeChartStackedLine(payload, "chart.xml", d)
			if chart == nil || chart.Grouping != grouping || len(chart.Series) != 2 || chart.Series[0].Index != 9 || chart.Series[1].Index != 5 || chart.Series[0].Order != 1 || chart.Series[1].Order != 0 || strings.Join(chart.Series[0].Values, ",") != "-1,2.5,1e0" || !bytes.Equal(payload, before) {
				t.Fatalf("lost source order/lexemes: %#v", chart)
			}
			if extractNativeChartLine(payload, "chart.xml", d) != nil {
				t.Fatal("stacked chart admitted by standard line profile")
			}
		}
	}
}
func TestNativeChartStackedLineRefusals(t *testing.T) {
	d, _ := nativeDialectForPresentation(xmlNamePresentation(false))
	source := strings.Replace(nativeConnectedXML(false, false), `grouping val="standard"`, `grouping val="percentStacked"`, 1)
	for name, pair := range map[string][2]string{"standard": {`percentStacked`, `standard`}, "smooth": {`smooth val="0"`, `smooth val="1"`}, "marker": {`symbol val="none"`, `symbol val="circle"`}, "sparse": {`ptCount val="3"`, `ptCount val="4"`}, "duplicate": {`pt idx="2"`, `pt idx="1"`}, "cache": {`numLit`, `numCache`}, "extra family": {`</c:lineChart>`, `</c:lineChart><c:barChart/>`}} {
		t.Run(name, func(t *testing.T) {
			if extractNativeChartStackedLine([]byte(strings.ReplaceAll(source, pair[0], pair[1])), "chart.xml", d) != nil {
				t.Fatal("unqualified source admitted")
			}
		})
	}
}
