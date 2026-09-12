package officehttp

import (
	"bytes"
	"crypto/sha256"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"strings"

	"github.com/injectinglabs/injoffice/go/docxpatch"
)

type docxOperatorFontPolicy struct {
	Version  int `json:"version"`
	Mappings []struct {
		SourceFamily string `json:"sourceFamily"`
		Weight       int    `json:"weight"`
		Style        string `json:"style"`
		TargetFamily string `json:"targetFamily"`
	} `json:"mappings"`
}

func fontPreviewJSONDigest(value any) string {
	var out bytes.Buffer
	enc := json.NewEncoder(&out)
	enc.SetEscapeHTML(false)
	if enc.Encode(value) != nil {
		return ""
	}
	return fmt.Sprintf("sha256:%x", sha256.Sum256(bytes.TrimSuffix(out.Bytes(), []byte{'\n'})))
}

// The worker independently qualifies source text, shapes and paint. This boundary
// additionally joins its evidence to this request and the operator-only policy.
func validateDOCXFontSubstitutionPreview(data json.RawMessage, input map[string]any, path string, compositions ...map[string]any) error {
	fail := func() error { return errors.New("DOCX substitution preview source or operator evidence mismatch") }
	var response struct {
		Protocol          string         `json:"protocol"`
		Version           int            `json:"version"`
		ReadOnly          bool           `json:"read_only"`
		Fidelity          string         `json:"fidelity"`
		Status            string         `json:"status"`
		Composition       map[string]any `json:"composition"`
		CompositionSHA256 string         `json:"composition_sha256"`
		Source            struct {
			DocumentID string `json:"document_id"`
			Revision   string `json:"revision"`
			Digest     string `json:"package_sha256"`
		} `json:"source"`
		Policy         string                 `json:"policy"`
		PolicyDigest   string                 `json:"policy_sha256"`
		OperatorPolicy docxOperatorFontPolicy `json:"operator_policy"`
		Manifest       map[string]any         `json:"selected_font_manifest"`
		Provenance     struct {
			DocumentID string `json:"document_id"`
			Revision   string `json:"revision"`
			Digest     string `json:"package_sha256"`
			Fonts      struct {
				ID       string `json:"manifest_id"`
				Revision string `json:"revision"`
				Digest   string `json:"sha256"`
			} `json:"font_manifest"`
		} `json:"rendering_provenance"`
		Entries []struct {
			SourceID       string `json:"source_id"`
			Role           string `json:"source_role"`
			SourceFamily   string `json:"source_family"`
			SelectedFamily string `json:"selected_family"`
			FaceID         string `json:"face_id"`
			Digest         string `json:"font_digest"`
			Weight         int    `json:"weight"`
			Style          string `json:"style"`
		} `json:"substitutions"`
	}
	if json.Unmarshal(data, &response) != nil || response.Protocol != "injoffice.docx.font-substitution-preview" || response.Version != 1 || !response.ReadOnly || response.Fidelity != "approximate" || response.Policy != "explicit-whole-run-font-substitution-v1" || len(response.Entries) > 10000 {
		return fail()
	}
	if len(compositions) > 0 {
		raw, e := json.Marshal(compositions[0])
		if e != nil || len(raw) > 8*1024*1024 {
			return fail()
		}
		var original map[string]any
		if json.Unmarshal(raw, &original) != nil || response.Composition == nil || fontPreviewJSONDigest(original) != fontPreviewJSONDigest(response.Composition) || response.CompositionSHA256 != fontPreviewJSONDigest(original) {
			return fail()
		}
	} else if response.Composition != nil || response.CompositionSHA256 != "" {
		return fail()
	}
	layout, ok := input["resolved_layout"].(*docxpatch.NativeResolvedLayoutInputV1)
	if !ok {
		return fail()
	}
	docBytes, err := json.Marshal(input["document"])
	if err != nil {
		return fail()
	}
	var source struct {
		DocumentID string `json:"document_id"`
		Revision   string `json:"revision"`
		Source     struct {
			Digest string `json:"package_sha256"`
		} `json:"source"`
	}
	if json.Unmarshal(docBytes, &source) != nil || source.DocumentID != response.Source.DocumentID || source.Revision != response.Source.Revision || source.Source.Digest != response.Source.Digest || response.Provenance.DocumentID != source.DocumentID || response.Provenance.Revision != source.Revision || response.Provenance.Digest != source.Source.Digest {
		return fail()
	}
	if response.Manifest == nil || response.Manifest["manifestId"] != response.Provenance.Fonts.ID || response.Manifest["revision"] != response.Provenance.Fonts.Revision || fontPreviewJSONDigest(response.Manifest) != response.Provenance.Fonts.Digest {
		return fail()
	}
	stat, err := os.Stat(path)
	if err != nil || !stat.Mode().IsRegular() || stat.Size() < 1 || stat.Size() > 65536 {
		return fail()
	}
	raw, err := os.ReadFile(path)
	if err != nil || len(raw) > 65536 {
		return fail()
	}
	var config struct {
		Version int                    `json:"version"`
		Policy  docxOperatorFontPolicy `json:"substitutions"`
		Faces   []struct {
			Family string `json:"family"`
			Weight int    `json:"weight"`
			Style  string `json:"style"`
			Digest string `json:"sha256"`
		} `json:"faces"`
	}
	if json.Unmarshal(raw, &config) != nil || config.Version != 1 || config.Policy.Version != 1 || len(config.Policy.Mappings) > 32 || len(config.Faces) > 32 || fontPreviewJSONDigest(config.Policy) != response.PolicyDigest || fontPreviewJSONDigest(response.OperatorPolicy) != response.PolicyDigest {
		return fail()
	}
	props := map[string]docxpatch.NativeResolvedRunPropertiesV1{}
	for _, run := range layout.Runs {
		props["run\x00"+run.RunID] = run.Properties
	}
	for _, p := range layout.Paragraphs {
		props["paragraph-mark\x00"+p.ParagraphID] = p.ParagraphMarkProperties
		if p.Numbering != nil && p.Numbering.Format != "bullet" {
			props["list-marker\x00"+p.ParagraphID] = p.Numbering.Marker
		}
	}
	seen := map[string]bool{}
	for _, e := range response.Entries {
		key := e.Role + "\x00" + e.SourceID
		p, ok := props[key]
		if !ok || seen[key] || p.FontFamily == nil || *p.FontFamily != e.SourceFamily || !fontPreviewDigest.MatchString(e.Digest) {
			return fail()
		}
		seen[key] = true
		weight, style := 400, "normal"
		if p.Bold != nil && *p.Bold {
			weight = 700
		}
		if p.Italic != nil && *p.Italic {
			style = "italic"
		}
		if e.Weight != weight || e.Style != style {
			return fail()
		}
		mapped := false
		for _, m := range config.Policy.Mappings {
			if strings.EqualFold(m.SourceFamily, e.SourceFamily) && m.TargetFamily == e.SelectedFamily && m.Weight == weight && m.Style == style {
				mapped = true
			}
		}
		selected := false
		for index, f := range config.Faces {
			if strings.EqualFold(f.Family, e.SourceFamily) && f.Weight == weight && f.Style == style {
				return fail()
			}
			if f.Family == e.SelectedFamily && f.Weight == weight && f.Style == style && f.Digest == e.Digest && e.FaceID == fmt.Sprintf("host-font-%d-%s", index, e.Digest[7:23]) {
				selected = true
			}
		}
		joined := false
		faces, _ := response.Manifest["faces"].([]any)
		for _, item := range faces {
			face, ok := item.(map[string]any)
			if !ok {
				continue
			}
			s, _ := face["source"].(map[string]any)
			if face["faceId"] == e.FaceID && face["family"] == e.SelectedFamily && face["weight"] == float64(weight) && face["style"] == style && s["kind"] == "host" && s["contentDigest"] == e.Digest {
				joined = true
			}
		}
		if !mapped || !selected || !joined {
			return fail()
		}
	}
	return nil
}
