package officehttp

import (
	"bytes"
	"context"
	"crypto/sha256"
	"encoding/json"
	"errors"
	"fmt"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"
	"time"
)

func TestPPTXPreviewQueryIsClosed(t *testing.T) {
	for _, query := range []string{"text=", "text=true", "text=source-inherited&text=source-inherited", "text=source-inherited&unknown=1"} {
		request := httptest.NewRequest(http.MethodPost, PPTXPreviewPath+"?"+query, nil)
		if _, err := parsePPTXPreviewSlide(request); err == nil {
			t.Fatalf("accepted inherited query %s", query)
		}
	}
	for _, query := range []string{"text=source-inherited", "slide=0&text=source-inherited&autofit=source-frame"} {
		request := httptest.NewRequest(http.MethodPost, PPTXPreviewPath+"?"+query, nil)
		if _, err := parsePPTXPreviewSlide(request); err != nil {
			t.Fatal(err)
		}
	}
	for _, query := range []string{"slide=-1", "slide=01", "slide=+1", "slide=", "slide=1&slide=2", "font_manifest_path=/tmp/font.json", "slide=1;foo=2", "slide=%zz", "autofit=true", "autofit=", "autofit=source-frame&autofit=source-frame", "autofit=source-frame&extra=1"} {
		request := httptest.NewRequest(http.MethodPost, PPTXPreviewPath, nil)
		request.URL.RawQuery = query
		if _, err := parsePPTXPreviewSlide(request); err == nil {
			t.Fatalf("accepted %q", query)
		}
	}
	for _, query := range []string{"", "slide=0", "slide=12", "autofit=source-frame", "slide=0&autofit=source-frame"} {
		request := httptest.NewRequest(http.MethodPost, PPTXPreviewPath+"?"+query, nil)
		if _, err := parsePPTXPreviewSlide(request); err != nil {
			t.Fatalf("refused %q: %v", query, err)
		}
	}
}

func TestPPTXPreviewDisabledBusyAndRequestBudgets(t *testing.T) {
	options := PPTXPreviewOptions{WorkerPath: "/operator/worker.js", FontManifestPath: "/operator/fonts.json"}
	for _, test := range []struct {
		method  string
		options PPTXPreviewOptions
		busy    bool
		code    int
	}{
		{http.MethodGet, options, false, http.StatusMethodNotAllowed},
		{http.MethodPost, PPTXPreviewOptions{}, false, http.StatusServiceUnavailable},
		{http.MethodPost, PPTXPreviewOptions{WorkerPath: "relative.js", FontManifestPath: "/fonts.json"}, false, http.StatusServiceUnavailable},
		{http.MethodPost, options, true, http.StatusServiceUnavailable},
	} {
		gate := make(chan struct{}, 1)
		if test.busy {
			gate <- struct{}{}
		}
		response := httptest.NewRecorder()
		handlePPTXPreview(response, httptest.NewRequest(test.method, PPTXPreviewPath, strings.NewReader("invalid")), test.options, gate)
		if response.Code != test.code || response.Header().Get("Cache-Control") != "no-store" {
			t.Fatalf("response=%d %s", response.Code, response.Body.String())
		}
	}
	gate := make(chan struct{}, 1)
	request := httptest.NewRequest(http.MethodPost, PPTXPreviewPath, bytes.NewReader(make([]byte, 8*1024*1024+1)))
	request.Header.Set("Content-Type", PPTXContentType)
	response := httptest.NewRecorder()
	handlePPTXPreview(response, request, options, gate)
	if response.Code != http.StatusRequestEntityTooLarge || len(gate) != 0 {
		t.Fatalf("oversize/gate: %d %d", response.Code, len(gate))
	}
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	if _, err := pptxPreviewInput(ctx, []byte("invalid"), 0, options); !errors.Is(err, context.Canceled) {
		t.Fatal(err)
	}
}

func TestPPTXPreviewInputBindsOriginalPackageWithoutStoring(t *testing.T) {
	data := readPinned(t, commonPPTXSHA, "officecompat", "corpus", "generated", "packages", "pptx-transitional-common.pptx")
	before := append([]byte(nil), data...)
	input, err := pptxPreviewInput(context.Background(), data, 0, PPTXPreviewOptions{FontManifestPath: "/operator/fonts.json"})
	if err != nil {
		t.Fatal(err)
	}
	if input["package_sha256"] != fmt.Sprintf("%x", sha256.Sum256(data)) || input["font_manifest_path"] != "/operator/fonts.json" || !bytes.Equal(data, before) {
		t.Fatal("source/config identity drift")
	}
	if _, err := pptxPreviewInput(context.Background(), data, 999999, PPTXPreviewOptions{}); err == nil {
		t.Fatal("accepted missing slide")
	}
}

func TestPPTXPreviewWorkerProtocolBudgetAndDeadline(t *testing.T) {
	worker := previewFixtureWorker(t, `const payload=Buffer.from(JSON.stringify({protocol:'injoffice.docx.page-paint-worker',version:1,id:'preview',ok:true,result:{}})); const h=Buffer.alloc(4); h.writeUInt32BE(payload.length); process.stdout.write(Buffer.concat([h,payload]));`)
	if _, err := compilePreviewWorker(context.Background(), worker.WorkerPath, "injoffice.pptx.preview-worker", map[string]any{}, 1024, 1024); err == nil {
		t.Fatal("accepted wrong worker protocol")
	}
	if _, err := compilePreviewWorker(context.Background(), worker.WorkerPath, "injoffice.pptx.preview-worker", map[string]any{}, 1, 1024); err == nil {
		t.Fatal("accepted oversized input frame")
	}
	if _, err := compilePreviewWorker(context.Background(), worker.WorkerPath, "injoffice.docx.page-paint-worker", map[string]any{}, 1024, 20); err == nil {
		t.Fatal("accepted oversized output frame")
	}
	hung := previewFixtureWorker(t, "setInterval(()=>{},1000);")
	ctx, cancel := context.WithTimeout(context.Background(), 100*time.Millisecond)
	defer cancel()
	if _, err := compilePreviewWorker(ctx, hung.WorkerPath, "injoffice.pptx.preview-worker", map[string]any{}, 1024, 1024); err == nil || ctx.Err() == nil {
		t.Fatal("worker ignored deadline")
	}
}

func TestPPTXPreviewRejectsStaleOrMissingWorkerIdentity(t *testing.T) {
	data := readPinned(t, commonPPTXSHA, "officecompat", "corpus", "generated", "packages", "pptx-transitional-common.pptx")
	for _, test := range []struct {
		name  string
		patch map[string]any
		want  int
	}{
		{"matching", nil, http.StatusOK},
		{"wrong-package", map[string]any{"package_sha256": strings.Repeat("0", 64)}, http.StatusUnprocessableEntity},
		{"wrong-slide", map[string]any{"slide_index": 1}, http.StatusUnprocessableEntity},
		{"missing-slide", map[string]any{"slide_index": nil}, http.StatusUnprocessableEntity},
		{"wrong-count", map[string]any{"slide_count": 999}, http.StatusUnprocessableEntity},
		{"wrong-approximation-count", map[string]any{"source_frame_autofit_count": 1}, http.StatusUnprocessableEntity},
		{"unsolicited-inherited-count", map[string]any{"inherited_text_preview_count": 1, "inherited_text_policy": "source-latin-inheritance-approximate-v1"}, http.StatusUnprocessableEntity},
		{"orphan-inherited-policy", map[string]any{"inherited_text_policy": "source-latin-inheritance-approximate-v1"}, http.StatusUnprocessableEntity},
	} {
		t.Run(test.name, func(t *testing.T) {
			input, err := pptxPreviewInput(context.Background(), data, 0, PPTXPreviewOptions{})
			if err != nil {
				t.Fatal(err)
			}
			// The corpus fixture has one slide, independently checked by extraction.
			deckJSON, _ := json.Marshal(input["deck"])
			var deck struct {
				Slides []json.RawMessage `json:"slides"`
			}
			if err := json.Unmarshal(deckJSON, &deck); err != nil {
				t.Fatal(err)
			}
			result := map[string]any{"version": 1, "package_sha256": input["package_sha256"], "slide_index": 0, "slide_count": len(deck.Slides)}
			for key, value := range test.patch {
				if value == nil {
					delete(result, key)
				} else {
					result[key] = value
				}
			}
			payload, _ := json.Marshal(map[string]any{"protocol": "injoffice.pptx.preview-worker", "version": 1, "id": "preview", "ok": true, "result": result})
			worker := previewFixtureWorker(t, fmt.Sprintf("const payload=Buffer.from(%q);const h=Buffer.alloc(4);h.writeUInt32BE(payload.length);process.stdout.write(Buffer.concat([h,payload]));", string(payload)))
			gate := make(chan struct{}, 1)
			request := httptest.NewRequest(http.MethodPost, PPTXPreviewPath, bytes.NewReader(data))
			request.Header.Set("Content-Type", PPTXContentType)
			response := httptest.NewRecorder()
			handlePPTXPreview(response, request, PPTXPreviewOptions{WorkerPath: worker.WorkerPath, FontManifestPath: "/operator/fonts.json"}, gate)
			if response.Code != test.want || len(gate) != 0 || response.Header().Get("Cache-Control") != "no-store" {
				t.Fatalf("response=%d gate=%d body=%s", response.Code, len(gate), response.Body.String())
			}
		})
	}
}

func TestPPTXInheritedPreviewWorkerPolicyJoin(t *testing.T) {
	data := readPinned(t, commonPPTXSHA, "officecompat", "corpus", "generated", "packages", "pptx-transitional-common.pptx")
	input, err := pptxPreviewInput(context.Background(), data, 0, PPTXPreviewOptions{InheritedTextPreview: true})
	if err != nil {
		t.Fatal(err)
	}
	count := input["inherited_text_preview_count"].(int)
	if count == 0 {
		t.Fatal("synthetic source has no inherited preview elements")
	}
	for _, tc := range []struct {
		name   string
		count  any
		policy any
		want   int
	}{{"matching", count, "source-latin-inheritance-approximate-v1", 200}, {"wrong count", count + 1, "source-latin-inheritance-approximate-v1", 422}, {"missing policy", count, nil, 422}, {"wrong policy", count, "future", 422}, {"fractional count", 1.5, "source-latin-inheritance-approximate-v1", 422}} {
		t.Run(tc.name, func(t *testing.T) {
			result := map[string]any{"version": 1, "package_sha256": input["package_sha256"], "slide_index": 0, "slide_count": 1, "inherited_text_preview_count": tc.count}
			if tc.policy != nil {
				result["inherited_text_policy"] = tc.policy
			}
			payload, _ := json.Marshal(map[string]any{"protocol": "injoffice.pptx.preview-worker", "version": 1, "id": "preview", "ok": true, "result": result})
			worker := previewFixtureWorker(t, fmt.Sprintf("const p=Buffer.from(%q);const h=Buffer.alloc(4);h.writeUInt32BE(p.length);process.stdout.write(Buffer.concat([h,p]));", string(payload)))
			request := httptest.NewRequest(http.MethodPost, PPTXPreviewPath+"?text=source-inherited", bytes.NewReader(data))
			request.Header.Set("Content-Type", PPTXContentType)
			response := httptest.NewRecorder()
			handlePPTXPreview(response, request, PPTXPreviewOptions{WorkerPath: worker.WorkerPath, FontManifestPath: "/operator/fonts.json"}, make(chan struct{}, 1))
			if response.Code != tc.want {
				t.Fatalf("status %d body %s", response.Code, response.Body.String())
			}
		})
	}
}
