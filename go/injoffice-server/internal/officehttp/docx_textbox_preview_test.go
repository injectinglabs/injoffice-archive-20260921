package officehttp

import (
	"bytes"
	"encoding/json"
	"net/http"
	"net/http/httptest"
	"os"
	"strings"
	"testing"
)

func TestDOCXTextboxPreviewGate(t *testing.T) {
	for _, tc := range []struct {
		method string
		status int
	}{{http.MethodGet, 405}, {http.MethodPost, 503}} {
		out := httptest.NewRecorder()
		NewHandler(nil).ServeHTTP(out, httptest.NewRequest(tc.method, DOCXTextboxPreviewPath, strings.NewReader("invalid")))
		if out.Code != tc.status {
			t.Fatalf("status %d want %d", out.Code, tc.status)
		}
	}
	gate := make(chan struct{}, 1)
	gate <- struct{}{}
	for _, query := range []string{"", "?font=/caller.ttf"} {
		probe := &previewReadProbe{}
		out := httptest.NewRecorder()
		handleDOCXPreview(out, httptest.NewRequest(http.MethodPost, DOCXTextboxPreviewPath+query, probe), DOCXPreviewOptions{WorkerPath: "/operator.js"}, gate)
		want := 503
		if query != "" {
			want = 400
		}
		if out.Code != want || probe.reads != 0 || len(gate) != 1 {
			t.Fatal("textbox request bypassed gate or read invalid query")
		}
	}
}

func TestDOCXTextboxPreviewRealWorker(t *testing.T) {
	worker, fixture := os.Getenv("INJOFFICE_TEST_DOCX_PREVIEW_WORKER"), os.Getenv("INJOFFICE_TEST_TEXTBOX_PAGE_FIXTURE")
	if worker == "" || fixture == "" {
		t.Skip("set real worker and textbox fixture paths")
	}
	data, err := os.ReadFile(fixture)
	if err != nil {
		t.Fatal(err)
	}
	before := bytes.Clone(data)
	request := httptest.NewRequest(http.MethodPost, DOCXTextboxPreviewPath, bytes.NewReader(data))
	request.Header.Set("Content-Type", DOCXContentType)
	out := httptest.NewRecorder()
	handleDOCXPreview(out, request, DOCXPreviewOptions{WorkerPath: worker}, make(chan struct{}, 1))
	if out.Code != 200 {
		t.Fatalf("textbox response %d: %s", out.Code, out.Body)
	}
	var response struct {
		Preview struct {
			Protocol          string `json:"protocol"`
			ReadOnly          bool   `json:"read_only"`
			SourceDiagnostics []any  `json:"source_diagnostics"`
			Textbox           struct {
				X int `json:"x_millipoints"`
				Y int `json:"y_millipoints"`
			} `json:"textbox"`
			Body struct {
				Pages []any `json:"pages"`
			} `json:"body_paint"`
		} `json:"preview"`
	}
	if err := json.Unmarshal(out.Body.Bytes(), &response); err != nil {
		t.Fatal(err)
	}
	p := response.Preview
	if p.Protocol != "injoffice.docx.textbox-page-preview" || !p.ReadOnly || len(p.SourceDiagnostics) != 1 || len(p.Body.Pages) != 1 || p.Textbox.X != 72000 || p.Textbox.Y != 144000 || !bytes.Equal(data, before) {
		t.Fatalf("invalid textbox source composition: %s", out.Body)
	}
	if out.Header().Get("Cache-Control") != "no-store" {
		t.Fatal("missing no-store response")
	}
}
