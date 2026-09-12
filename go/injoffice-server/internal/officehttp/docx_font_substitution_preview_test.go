package officehttp

import (
	"encoding/json"
	"github.com/injectinglabs/injoffice/go/docxpatch"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestDOCXFontSubstitutionBoundary(t *testing.T) {
	for _, tc := range []struct {
		method string
		status int
	}{{http.MethodGet, 405}, {http.MethodPost, 503}} {
		response := httptest.NewRecorder()
		NewHandler(nil).ServeHTTP(response, httptest.NewRequest(tc.method, DOCXFontSubstitutionPreviewPath, strings.NewReader("invalid")))
		if response.Code != tc.status {
			t.Fatalf("status %d expected %d", response.Code, tc.status)
		}
	}
	probe := &previewReadProbe{}
	response := httptest.NewRecorder()
	handleDOCXPreview(response, httptest.NewRequest(http.MethodPost, DOCXFontSubstitutionPreviewPath+"?fonts=unknown", probe), DOCXPreviewOptions{WorkerPath: "/worker.js", FontManifestPath: "/operator.json"}, make(chan struct{}, 1))
	if response.Code != 400 || probe.reads != 0 {
		t.Fatal("font preview did not reject query before reading source")
	}
}

func TestDOCXFontSubstitutionEvidenceJoins(t *testing.T) {
	digest := "sha256:" + strings.Repeat("a", 64)
	config := map[string]any{"version": 1, "substitutions": map[string]any{"version": 1, "mappings": []any{map[string]any{"sourceFamily": "Missing", "targetFamily": "Selected", "weight": 400, "style": "normal"}}}, "faces": []any{map[string]any{"family": "Selected", "weight": 400, "style": "normal", "sha256": digest}}}
	path := filepath.Join(t.TempDir(), "operator.json")
	raw, _ := json.Marshal(config)
	if err := os.WriteFile(path, raw, 0600); err != nil {
		t.Fatal(err)
	}
	var policy docxOperatorFontPolicy
	policyBytes, _ := json.Marshal(config["substitutions"])
	_ = json.Unmarshal(policyBytes, &policy)
	family := "Missing"
	layout := &docxpatch.NativeResolvedLayoutInputV1{DocumentID: "doc", Revision: "rev", Runs: []docxpatch.NativeResolvedRunV1{{RunID: "r1", Properties: docxpatch.NativeResolvedRunPropertiesV1{FontFamily: &family}}}}
	input := map[string]any{"resolved_layout": layout, "document": map[string]any{"document_id": "doc", "revision": "rev", "source": map[string]any{"package_sha256": digest}}}
	faceID := "host-font-0-" + strings.Repeat("a", 16)
	manifest := map[string]any{"version": 1, "manifestId": "fonts", "revision": "rev", "faces": []any{map[string]any{"faceId": faceID, "family": "Selected", "weight": 400, "style": "normal", "stretch": 100, "source": map[string]any{"kind": "host", "contentDigest": digest, "resourceId": faceID}}}, "fallbackChains": []any{}}
	base := map[string]any{"protocol": "injoffice.docx.font-substitution-preview", "version": 1, "read_only": true, "fidelity": "approximate", "status": "painted", "source": map[string]any{"document_id": "doc", "revision": "rev", "package_sha256": digest}, "policy": "explicit-whole-run-font-substitution-v1", "operator_policy": policy, "policy_sha256": fontPreviewJSONDigest(policy), "selected_font_manifest": manifest, "rendering_provenance": map[string]any{"document_id": "doc", "revision": "rev", "package_sha256": digest, "font_manifest": map[string]any{"manifest_id": "fonts", "revision": "rev", "sha256": fontPreviewJSONDigest(manifest)}}, "substitutions": []any{map[string]any{"source_id": "r1", "source_role": "run", "source_family": "Missing", "selected_family": "Selected", "face_id": faceID, "font_digest": digest, "weight": 400, "style": "normal"}}}
	original, _ := json.Marshal(base)
	if err := validateDOCXFontSubstitutionPreview(original, input, path); err != nil {
		t.Fatal(err)
	}
	for _, mutate := range []func(map[string]any){
		func(v map[string]any) { v["source"].(map[string]any)["revision"] = "stale" },
		func(v map[string]any) { v["policy_sha256"] = "sha256:" + strings.Repeat("b", 64) },
		func(v map[string]any) { v["read_only"] = false },
		func(v map[string]any) { v["substitutions"].([]any)[0].(map[string]any)["source_id"] = "unknown" },
		func(v map[string]any) {
			v["substitutions"].([]any)[0].(map[string]any)["font_digest"] = "sha256:" + strings.Repeat("b", 64)
		},
		func(v map[string]any) { v["substitutions"].([]any)[0].(map[string]any)["face_id"] = "unknown" },
		func(v map[string]any) {
			v["substitutions"] = append(v["substitutions"].([]any), v["substitutions"].([]any)[0])
		},
	} {
		var v map[string]any
		_ = json.Unmarshal(original, &v)
		mutate(v)
		raw, _ := json.Marshal(v)
		if validateDOCXFontSubstitutionPreview(raw, input, path) == nil {
			t.Fatal("forged evidence accepted")
		}
	}
}
