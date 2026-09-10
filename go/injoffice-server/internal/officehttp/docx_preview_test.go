package officehttp

import (
	"bytes"
	"context"
	"crypto/sha256"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"net/http/httptest"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"testing"
	"time"
)

func TestDOCXPreviewDisabledAndReadOnly(t *testing.T) {
	handler := NewHandler(nil)
	for _, test := range []struct {
		method string
		status int
	}{{http.MethodGet, 405}, {http.MethodPost, 503}} {
		response := httptest.NewRecorder()
		handler.ServeHTTP(response, httptest.NewRequest(test.method, DOCXPreviewPath, bytes.NewReader([]byte("invalid"))))
		if response.Code != test.status {
			t.Fatalf("status %d want %d", response.Code, test.status)
		}
	}
}

func TestDOCXPreviewInputRejectsInvalidPackage(t *testing.T) {
	if _, err := docxPreviewInput(context.Background(), []byte("not a zip")); err == nil {
		t.Fatal("accepted malformed package")
	}
}

type previewReadProbe struct{ reads int }

func (probe *previewReadProbe) Read([]byte) (int, error) {
	probe.reads++
	return 0, io.EOF
}

func TestDOCXPreviewBusyDoesNotReadRequestAndFailureReleasesGate(t *testing.T) {
	gate := make(chan struct{}, 1)
	gate <- struct{}{}
	probe := &previewReadProbe{}
	response := httptest.NewRecorder()
	options := DOCXPreviewOptions{WorkerPath: "/operator-configured-worker.js"}
	handleDOCXPreview(response, httptest.NewRequest(http.MethodPost, DOCXPreviewPath, probe), options, gate)
	if response.Code != http.StatusServiceUnavailable || probe.reads != 0 || len(gate) != 1 {
		t.Fatalf("busy request consumed input or slot: status=%d reads=%d slots=%d", response.Code, probe.reads, len(gate))
	}
	<-gate
	for attempt := 0; attempt < 2; attempt++ {
		response = httptest.NewRecorder()
		request := httptest.NewRequest(http.MethodPost, DOCXPreviewPath, strings.NewReader("not a DOCX"))
		request.Header.Set("Content-Type", DOCXContentType)
		handleDOCXPreview(response, request, options, gate)
		if response.Code != http.StatusUnprocessableEntity || len(gate) != 0 {
			t.Fatalf("invalid document did not release slot: status=%d slots=%d", response.Code, len(gate))
		}
	}
}

func TestDOCXPreviewOversizeAndCancellation(t *testing.T) {
	tooLarge := make([]byte, 8*1024*1024+1)
	if _, err := docxPreviewInput(context.Background(), tooLarge); err == nil || !strings.Contains(err.Error(), "8 MiB") {
		t.Fatalf("oversize input was not refused before extraction: %v", err)
	}
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	if _, err := docxPreviewInput(ctx, []byte("invalid")); !errors.Is(err, context.Canceled) {
		t.Fatalf("cancelled input was parsed instead of refused: %v", err)
	}
	gate := make(chan struct{}, 1)
	response := httptest.NewRecorder()
	request := httptest.NewRequest(http.MethodPost, DOCXPreviewPath, bytes.NewReader(tooLarge))
	request.Header.Set("Content-Type", DOCXContentType)
	handleDOCXPreview(response, request, DOCXPreviewOptions{WorkerPath: "/operator-configured-worker.js"}, gate)
	if response.Code != http.StatusRequestEntityTooLarge || len(gate) != 0 {
		t.Fatalf("oversize HTTP request did not release slot: status=%d slots=%d", response.Code, len(gate))
	}
}

// Workers here deliberately exercise only framing/lifecycle failures. They do
// not produce document paint or replace the separate real-compiler test.
func previewFixtureWorker(t *testing.T, javascript string) DOCXPreviewOptions {
	t.Helper()
	if _, err := exec.LookPath("node"); err != nil {
		t.Skip("Node is required for preview subprocess lifecycle tests")
	}
	path := filepath.Join(t.TempDir(), "worker.cjs")
	if err := os.WriteFile(path, []byte("process.stdin.resume(); process.stdin.on('end', () => {"+javascript+"});"), 0600); err != nil {
		t.Fatal(err)
	}
	return DOCXPreviewOptions{WorkerPath: path}
}

func TestDOCXPreviewWorkerRejectsMalformedFramesAndEnvelopes(t *testing.T) {
	for _, test := range []struct{ name, code, errorText string }{
		{"empty", "", "invalid frame"},
		{"short-header", "process.stdout.write(Buffer.from([0,0,0]));", "invalid frame"},
		{"truncated-payload", "process.stdout.write(Buffer.from([0,0,0,10,123]));", "invalid frame"},
		{"trailing-frame", "process.stdout.write(Buffer.from([0,0,0,0,0]));", "invalid frame"},
		{"invalid-json", `const data=Buffer.from('{'); const header=Buffer.alloc(4); header.writeUInt32BE(data.length); process.stdout.write(Buffer.concat([header,data]));`, "invalid envelope"},
		{"wrong-identity", `const data=Buffer.from(JSON.stringify({protocol:'injoffice.docx.page-paint-worker',version:1,id:'other-request',ok:true,result:{}})); const header=Buffer.alloc(4); header.writeUInt32BE(data.length); process.stdout.write(Buffer.concat([header,data]));`, "invalid envelope"},
		{"nonzero-exit", "process.exitCode=1;", "invalid frame"},
		{"worker-refusal", `const data=Buffer.from(JSON.stringify({protocol:'injoffice.docx.page-paint-worker',version:1,id:'preview',ok:false,error:{message:'qualified fixture refusal'}})); const header=Buffer.alloc(4); header.writeUInt32BE(data.length); process.stdout.write(Buffer.concat([header,data]));`, "qualified fixture refusal"},
	} {
		t.Run(test.name, func(t *testing.T) {
			result, err := compileDOCXPreview(context.Background(), previewFixtureWorker(t, test.code), map[string]any{})
			if err == nil || !strings.Contains(err.Error(), test.errorText) || result != nil {
				t.Fatalf("bad worker output accepted: result=%s error=%v", result, err)
			}
		})
	}
}

func TestDOCXPreviewWorkerParentDeadlineStopsOpenPipe(t *testing.T) {
	worker := previewFixtureWorker(t, "setInterval(() => {}, 1000);")
	ctx, cancel := context.WithTimeout(context.Background(), time.Second)
	defer cancel()
	started := time.Now()
	result, err := compileDOCXPreview(ctx, worker, map[string]any{})
	if err == nil || !strings.Contains(err.Error(), "time budget") || result != nil || !errors.Is(ctx.Err(), context.DeadlineExceeded) {
		t.Fatalf("hung worker did not honor parent deadline: result=%s error=%v context=%v", result, err, ctx.Err())
	}
	if elapsed := time.Since(started); elapsed > 5*time.Second {
		t.Fatalf("worker stdout read outlived parent deadline: %s", elapsed)
	}
}

// Opt-in integration test uses the real built compiler, never fabricated paths.
func TestDOCXPreviewRealWorker(t *testing.T) {
	worker := os.Getenv("INJOFFICE_TEST_DOCX_PREVIEW_WORKER")
	fixture := os.Getenv("INJOFFICE_TEST_DOCX_PREVIEW_FIXTURE")
	if worker == "" || fixture == "" {
		t.Skip("set worker and qualified fixture paths to exercise real native compilation")
	}
	data, err := os.ReadFile(fixture)
	if err != nil {
		t.Fatal(err)
	}
	handler := NewHandlerWithDOCXPreview(nil, DOCXPreviewOptions{WorkerPath: worker})
	request := httptest.NewRequest(http.MethodPost, DOCXPreviewPath, bytes.NewReader(data))
	request.Header.Set("Content-Type", DOCXContentType)
	response := httptest.NewRecorder()
	handler.ServeHTTP(response, request)
	var result map[string]json.RawMessage
	if err := json.Unmarshal(response.Body.Bytes(), &result); err != nil {
		t.Fatal(err)
	}
	if response.Code != 200 || !bytes.Contains(result["page_paint_output"], []byte(`"status":"painted"`)) {
		var request struct {
			PaginationRequest struct {
				Document struct {
					UnsupportedCapabilities any `json:"unsupported_capabilities"`
				} `json:"document"`
			} `json:"pagination_request"`
			PaginatedLayout struct {
				Diagnostics any `json:"diagnostics"`
			} `json:"paginated_layout"`
		}
		_ = json.Unmarshal(result["page_paint_request"], &request)
		var output struct {
			Diagnostics any `json:"diagnostics"`
		}
		_ = json.Unmarshal(result["page_paint_output"], &output)
		t.Fatalf("native preview status=%d errors=%s diagnostics=%+v paint=%+v", response.Code, result["error"], request.PaginatedLayout.Diagnostics, output.Diagnostics)
	}
	var output struct {
		Provenance struct {
			PackageSHA256 string `json:"package_sha256"`
		} `json:"provenance"`
		Pages []struct {
			Commands []struct {
				Kind string `json:"kind"`
				Path []any  `json:"path"`
			} `json:"commands"`
		} `json:"pages"`
	}
	if err := json.Unmarshal(result["page_paint_output"], &output); err != nil {
		t.Fatal(err)
	}
	if len(output.Pages) != 2 {
		t.Fatalf("native pages=%d want 2", len(output.Pages))
	}
	if output.Provenance.PackageSHA256 != fmt.Sprintf("sha256:%x", sha256.Sum256(data)) {
		t.Fatal("paint does not bind exact input bytes")
	}
	for _, page := range output.Pages {
		glyphs := 0
		for _, command := range page.Commands {
			if command.Kind == "fill_glyph_path" && len(command.Path) > 0 {
				glyphs++
			}
		}
		if glyphs == 0 {
			t.Fatal("native page has no real glyph contours")
		}
	}
}
