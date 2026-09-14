package officehttp

import (
	"net/http/httptest"
	"testing"
)

func TestPPTXPreviewChartQueryExplicit(t *testing.T) {
	for _, query := range []string{"", "slide=0", "charts=source-literal", "slide=1&charts=source-literal&fonts=operator-substitution"} {
		if _, err := parsePPTXPreviewSlide(httptest.NewRequest("POST", PPTXPreviewPath+"?"+query, nil)); err != nil {
			t.Fatalf("valid query %s: %v", query, err)
		}
	}
	for _, query := range []string{"charts=", "charts=true", "charts=source-literal&charts=source-literal", "charts=source-literal;fonts=operator-substitution", "Charts=source-literal", "font_manifest_path=/tmp/fonts.json"} {
		if _, err := parsePPTXPreviewSlide(httptest.NewRequest("POST", PPTXPreviewPath+"?"+query, nil)); err == nil {
			t.Fatalf("invalid query %s accepted", query)
		}
	}
}

func TestPPTXPreviewChartResponseBinding(t *testing.T) {
	if !validPPTXChartPreviewMode([]byte(`{}`), false) || !validPPTXChartPreviewMode([]byte(`{"source_chart_preview":true,"chart_axis_layout_policy":"supplied-outline-margins-v1"}`), true) {
		t.Fatal("exact chart mode refused")
	}
	for _, raw := range []string{`{}`, `{"source_chart_preview":false}`, `{"source_chart_preview":"true"}`, `{"source_chart_preview":true}`, `{"source_chart_preview":true,"chart_axis_layout_policy":"other"}`, `{"Source_Chart_Preview":true,"chart_axis_layout_policy":"supplied-outline-margins-v1"}`} {
		if validPPTXChartPreviewMode([]byte(raw), true) {
			t.Fatalf("enabled mismatch accepted %s", raw)
		}
	}
	for _, raw := range []string{`{"source_chart_preview":false}`, `{"source_chart_preview":null}`, `{"chart_axis_layout_policy":null}`, `{"source_chart_preview":true,"chart_axis_layout_policy":"supplied-outline-margins-v1"}`} {
		if validPPTXChartPreviewMode([]byte(raw), false) {
			t.Fatalf("disabled mismatch accepted %s", raw)
		}
	}
}
