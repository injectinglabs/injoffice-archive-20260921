package officehttp

import (
	"crypto/sha256"
	"encoding/json"
	"fmt"
	"github.com/injectinglabs/injoffice/go/pptxpatch"
	"net/http/httptest"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestPPTXFontSubstitutionEvidence(t *testing.T) {
	source := "Authored"
	paragraphs := []pptxpatch.NativeParagraph{{Runs: []pptxpatch.NativeTextRun{{FontFamily: &source}}}}
	elements := []pptxpatch.NativeElement{{ID: "shape:test", Kind: "text", Paragraphs: &paragraphs}}
	policy := `{"version":1,"mappings":[{"sourceFamily":"Authored","weight":400,"style":"normal","targetFamily":"Selected"}]}`
	digest := "sha256:" + strings.Repeat("a", 64)
	manifest := filepath.Join(t.TempDir(), "fonts.json")
	if err := os.WriteFile(manifest, []byte(`{"version":1,"faces":[{"family":"Selected","weight":400,"style":"normal","sha256":"`+digest+`"}],"substitutions":`+policy+`}`), 0600); err != nil {
		t.Fatal(err)
	}
	entry := map[string]any{"source_id": "shape:test", "paragraph_index": 0, "run_index": 0, "source_family": source, "selected_family": "Selected", "face_id": "font-0", "font_digest": digest}
	base := map[string]any{"font_substitutions": []any{entry}, "font_substitution_policy": "explicit-whole-run-font-substitution-v1", "font_substitution_policy_sha256": fmt.Sprintf("sha256:%x", sha256.Sum256([]byte(policy))), "font_digests": []string{digest}}
	encode := func(v any) json.RawMessage {
		b, e := json.Marshal(v)
		if e != nil {
			t.Fatal(e)
		}
		return b
	}
	if err := validatePPTXFontSubstitutions(encode(base), elements, true, manifest); err != nil {
		t.Fatal(err)
	}
	if validatePPTXFontSubstitutions(encode(base), elements, false, manifest) == nil {
		t.Fatal("accepted unrequested substitution")
	}
	for _, change := range []map[string]any{{"source_id": "other"}, {"source_family": "other"}, {"selected_family": "other"}, {"face_id": "font-1"}, {"font_digest": "sha256:" + strings.Repeat("b", 64)}, {"run_index": 1}, {"run_index": -1}, {"paragraph_index": nil}} {
		mutated := map[string]any{}
		for k, v := range entry {
			mutated[k] = v
		}
		for k, v := range change {
			mutated[k] = v
		}
		envelope := map[string]any{}
		for k, v := range base {
			envelope[k] = v
		}
		envelope["font_substitutions"] = []any{mutated}
		if validatePPTXFontSubstitutions(encode(envelope), elements, true, manifest) == nil {
			t.Fatalf("accepted forged %v", change)
		}
	}
	for _, change := range []map[string]any{{"font_substitutions": []any{}}, {"font_substitutions": []any{entry, entry}}, {"font_substitution_policy": "unknown"}, {"font_substitution_policy_sha256": "sha256:" + strings.Repeat("b", 64)}, {"font_digests": []string{}}} {
		envelope := map[string]any{}
		for k, v := range base {
			envelope[k] = v
		}
		for k, v := range change {
			envelope[k] = v
		}
		if validatePPTXFontSubstitutions(encode(envelope), elements, true, manifest) == nil {
			t.Fatalf("accepted forged envelope %v", change)
		}
	}
	if validatePPTXFontSubstitutions(json.RawMessage(`{}`), elements, false, "") != nil {
		t.Fatal("changed legacy exact behavior")
	}
}
func TestPPTXFontSubstitutionQuery(t *testing.T) {
	for _, query := range []string{"fonts=", "fonts=true", "fonts=operator-substitution&fonts=operator-substitution", "fonts=operator-substitution&other=x"} {
		if _, err := parsePPTXPreviewSlide(httptest.NewRequest("POST", PPTXPreviewPath+"?"+query, nil)); err == nil {
			t.Fatal(query)
		}
	}
	if _, err := parsePPTXPreviewSlide(httptest.NewRequest("POST", PPTXPreviewPath+"?fonts=operator-substitution&text=source-inherited&autofit=source-frame", nil)); err != nil {
		t.Fatal(err)
	}
}
