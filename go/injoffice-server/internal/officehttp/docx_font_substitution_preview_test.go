package officehttp

import (
	"archive/zip"
	"bytes"
	"context"
	"encoding/json"
	"github.com/injectinglabs/injoffice/go/docxpatch"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestDOCXFontCompositionConstruction(t *testing.T) {
	for _, mode := range []string{"15", "12"} {
		var buf bytes.Buffer
		writer := zip.NewWriter(&buf)
		parts := map[string]string{
			"[Content_Types].xml":          `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>`,
			"_rels/.rels":                  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
			"word/document.xml":            `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr><w:t>A</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
			"word/_rels/document.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="settings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>`,
			"word/settings.xml":            `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="` + mode + `"/></w:compat></w:settings>`,
		}
		parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], "</Types>", `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`, 1)
		parts["word/_rels/document.xml.rels"] = strings.Replace(parts["word/_rels/document.xml.rels"], "</Relationships>", `<Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`, 1)
		parts["word/styles.xml"] = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"/></w:styles>`
		for name, content := range parts {
			entry, err := writer.Create(name)
			if err != nil {
				t.Fatal(err)
			}
			if _, err = entry.Write([]byte(content)); err != nil {
				t.Fatal(err)
			}
		}
		if err := writer.Close(); err != nil {
			t.Fatal(err)
		}
		data := buf.Bytes()
		input, err := docxPreviewInput(context.Background(), data)
		if err != nil {
			t.Fatal(err)
		}
		composition, err := docxFontPreviewComposition(input, data)
		if err != nil {
			t.Fatal(err)
		}
		if composition["source_document"] != input["document"] || composition["source_resolved_layout"] != input["resolved_layout"] || composition["source_pagination_settings"] != input["pagination_settings"] || composition["source_font_inventory_json"] != input["font_inventory_json"] {
			t.Fatal("composition did not retain originals")
		}
		if (composition["legacy_eligibility"] != nil) != (mode == "12") || (composition["font_size_policy"] != nil) != (mode == "12") {
			t.Fatalf("unexpected closed policy composition %#v", composition)
		}
		if composition["automatic_borders"] != nil {
			t.Fatal("automatic borders invented without source table")
		}
		input["resolved_layout"].(*docxpatch.NativeResolvedLayoutInputV1).Tables = []docxpatch.NativeResolvedTableV1{{AutomaticBorderPreview: &docxpatch.NativeAutomaticTableBorderPreviewV1{Policy: docxpatch.NativeAutomaticTableBorderPolicyV1}}}
		withBorders, err := docxFontPreviewComposition(input, data)
		if err != nil || withBorders["automatic_borders"] != true {
			t.Fatal("qualified source border flag not routed")
		}
		delete(input, "pagination_settings")
		if _, err = docxFontPreviewComposition(input, data); err == nil {
			t.Fatal("missing source settings accepted")
		}
	}
}

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
	// The HTTP boundary joins the complete worker echo to independently
	// extracted request composition, including fields unknown to paint itself.
	composition := map[string]any{"source_document": input["document"], "source_resolved_layout": layout, "source_font_inventory_json": "canonical-source-inventory"}
	composed := map[string]any{}
	_ = json.Unmarshal(original, &composed)
	composed["composition"] = composition
	compositionBytes, _ := json.Marshal(composition)
	var canonicalComposition map[string]any
	_ = json.Unmarshal(compositionBytes, &canonicalComposition)
	composed["composition_sha256"] = fontPreviewJSONDigest(canonicalComposition)
	composedBytes, _ := json.Marshal(composed)
	if err := validateDOCXFontSubstitutionPreview(composedBytes, input, path, composition); err != nil {
		t.Fatal(err)
	}
	if validateDOCXFontSubstitutionPreview(composedBytes, input, path) == nil {
		t.Fatal("unexpected composition accepted without request authority")
	}
	if validateDOCXFontSubstitutionPreview(original, input, path, composition) == nil {
		t.Fatal("missing composition accepted")
	}
	for _, mutate := range []func(map[string]any){
		func(v map[string]any) { delete(v, "composition_sha256") },
		func(v map[string]any) { v["composition_sha256"] = "sha256:" + strings.Repeat("b", 64) },
		func(v map[string]any) {
			v["composition"].(map[string]any)["source_font_inventory_json"] = "changed"
			v["composition_sha256"] = fontPreviewJSONDigest(v["composition"])
		},
		func(v map[string]any) {
			v["composition"].(map[string]any)["extra_policy"] = true
			v["composition_sha256"] = fontPreviewJSONDigest(v["composition"])
		},
	} {
		var changed map[string]any
		_ = json.Unmarshal(composedBytes, &changed)
		mutate(changed)
		encoded, _ := json.Marshal(changed)
		if validateDOCXFontSubstitutionPreview(encoded, input, path, composition) == nil {
			t.Fatal("forged composition accepted")
		}
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
