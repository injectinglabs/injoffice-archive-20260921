package officehttp

import (
	"bytes"
	"context"
	"encoding/json"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"
)

func TestPPTXWorkbookChartModeSourceAndEcho(t *testing.T) {
	data := readPinned(t, commonPPTXSHA, "officecompat", "corpus", "generated", "packages", "pptx-transitional-common.pptx")
	original := append([]byte(nil), data...)
	off, err := pptxPreviewInput(context.Background(), data, 0, PPTXPreviewOptions{})
	if err != nil {
		t.Fatal(err)
	}
	if off["workbook_chart_preview"] != nil || off["workbook_chart_data"] != nil {
		t.Fatal("default enabled workbook mode")
	}
	on, err := pptxPreviewInput(context.Background(), data, 0, PPTXPreviewOptions{WorkbookChartPreview: true})
	if err != nil {
		t.Fatal(err)
	}
	if on["workbook_chart_preview"] != true || on["source_chart_preview"] != false || !bytes.Equal(data, original) {
		t.Fatal("mode/source changed")
	}
	payload := on["workbook_chart_data"].(pptxWorkbookPayload)
	var inspection map[string]any
	if json.Unmarshal([]byte(payload.InspectionJSON), &inspection) != nil || inspection["package_sha256"] != on["package_sha256"] {
		t.Fatal("inspection lost source binding")
	}
	for _, query := range []string{"charts=source-workbook", "slide=0&charts=source-workbook"} {
		if _, err := parsePPTXPreviewSlide(httptest.NewRequest(http.MethodPost, PPTXPreviewPath+"?"+query, nil)); err != nil {
			t.Fatal(err)
		}
	}
	for _, query := range []string{"charts=workbook", "charts=source-workbook&charts=source-literal", "charts=source-workbook&path=/tmp/book.xlsx"} {
		if _, err := parsePPTXPreviewSlide(httptest.NewRequest(http.MethodPost, PPTXPreviewPath+"?"+query, nil)); err == nil {
			t.Fatal("unqualified workbook query accepted")
		}
	}
	good := []byte(`{"workbook_chart_preview":true,"chart_axis_layout_policy":"supplied-outline-margins-v1"}`)
	if !validPPTXWorkbookPreviewMode(good, true, false) || validPPTXWorkbookPreviewMode(good, false, false) || validPPTXWorkbookPreviewMode(good, true, true) {
		t.Fatal("workbook echo mismatch")
	}
	for _, source := range []string{`{}`, `{"workbook_chart_preview":false}`, `{"workbook_chart_preview":true}`, `{"workbook_chart_preview":"true","chart_axis_layout_policy":"supplied-outline-margins-v1"}`, `{"workbook_chart_preview":true,"source_chart_preview":true,"chart_axis_layout_policy":"supplied-outline-margins-v1"}`} {
		if validPPTXWorkbookPreviewMode([]byte(source), true, false) {
			t.Fatal("bad workbook echo accepted")
		}
	}
}

func TestPPTXWorkbookAggregateUTF8Budget(t *testing.T) {
	payload := pptxWorkbookPayload{InspectionJSON: "{}", Workbooks: []pptxWorkbookContract{}}
	encoded, err := json.Marshal(payload)
	if err != nil {
		t.Fatal(err)
	}
	payload.InspectionJSON = strings.Repeat("a", pptxWorkbookPayloadLimit-len(encoded)+2)
	if err := checkPPTXWorkbookPayload(payload); err != nil {
		t.Fatal("exact budget refused", err)
	}
	payload.InspectionJSON += "a"
	if checkPPTXWorkbookPayload(payload) == nil {
		t.Fatal("budget overflow accepted")
	}
	payload = pptxWorkbookPayload{InspectionJSON: strings.Repeat("é", 3*1024*1024), Workbooks: []pptxWorkbookContract{{ContractJSON: strings.Repeat("é", 2*1024*1024)}}}
	if checkPPTXWorkbookPayload(payload) == nil {
		t.Fatal("aggregate UTF-8 treated as character count")
	}
}
