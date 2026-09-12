package officehttp

import (
	"bytes"
	"crypto/sha256"
	"encoding/json"
	"errors"
	"fmt"
	"github.com/injectinglabs/injoffice/go/pptxpatch"
	"os"
	"regexp"
	"strings"
)

var fontPreviewDigest = regexp.MustCompile(`^sha256:[a-f0-9]{64}$`)

func validatePPTXFontSubstitutions(data json.RawMessage, elements []pptxpatch.NativeElement, allowed bool, manifestPath string) error {
	var envelope struct {
		Entries []struct {
			SourceID       string `json:"source_id"`
			Paragraph      *int   `json:"paragraph_index"`
			Run            *int   `json:"run_index"`
			SourceFamily   string `json:"source_family"`
			SelectedFamily string `json:"selected_family"`
			FaceID         string `json:"face_id"`
			Digest         string `json:"font_digest"`
		} `json:"font_substitutions"`
		Policy      string   `json:"font_substitution_policy"`
		Digest      string   `json:"font_substitution_policy_sha256"`
		FontDigests []string `json:"font_digests"`
	}
	fail := func() error { return errors.New("font substitution evidence does not match the opted-in source slide") }
	if json.Unmarshal(data, &envelope) != nil {
		return fail()
	}
	if len(envelope.Entries) == 0 {
		var fields map[string]json.RawMessage
		_ = json.Unmarshal(data, &fields)
		if fields["font_substitutions"] != nil || fields["font_substitution_policy"] != nil || fields["font_substitution_policy_sha256"] != nil {
			return fail()
		}
		return nil
	}
	if !allowed || len(envelope.Entries) > 20000 || envelope.Policy != "explicit-whole-run-font-substitution-v1" || !fontPreviewDigest.MatchString(envelope.Digest) {
		return fail()
	}
	var config struct {
		Version int `json:"version"`
		Faces   []struct {
			Family string `json:"family"`
			Weight int    `json:"weight"`
			Style  string `json:"style"`
			Digest string `json:"sha256"`
		} `json:"faces"`
		Policy struct {
			Version  int `json:"version"`
			Mappings []struct {
				SourceFamily string `json:"sourceFamily"`
				Weight       int    `json:"weight"`
				Style        string `json:"style"`
				TargetFamily string `json:"targetFamily"`
			} `json:"mappings"`
		} `json:"substitutions"`
	}
	stat, err := os.Stat(manifestPath)
	if err != nil || !stat.Mode().IsRegular() || stat.Size() < 1 || stat.Size() > 65536 {
		return fail()
	}
	configured, err := os.ReadFile(manifestPath)
	if err != nil || len(configured) > 65536 || json.Unmarshal(configured, &config) != nil || config.Version != 1 || config.Policy.Version != 1 || len(config.Policy.Mappings) > 32 || len(config.Faces) > 32 {
		return fail()
	}
	var canonical bytes.Buffer
	encoder := json.NewEncoder(&canonical)
	encoder.SetEscapeHTML(false)
	if encoder.Encode(config.Policy) != nil {
		return fail()
	}
	policyDigest := fmt.Sprintf("sha256:%x", sha256.Sum256(bytes.TrimSuffix(canonical.Bytes(), []byte{'\n'})))
	if policyDigest != envelope.Digest {
		return fail()
	}
	source := map[string][]pptxpatch.NativeParagraph{}
	var collect func([]pptxpatch.NativeElement)
	collect = func(es []pptxpatch.NativeElement) {
		for _, e := range es {
			if (e.Kind == "text" || e.Kind == "shape") && e.Paragraphs != nil {
				source[e.ID] = *e.Paragraphs
			}
			collect(e.Children)
		}
	}
	collect(elements)
	seen := map[string]bool{}
	for _, e := range envelope.Entries {
		paragraphs, ok := source[e.SourceID]
		if !ok || e.Paragraph == nil || e.Run == nil || *e.Paragraph < 0 || *e.Paragraph >= len(paragraphs) || *e.Run < 0 || *e.Run >= len(paragraphs[*e.Paragraph].Runs) || e.FaceID == "" || len(e.FaceID) > 128 || e.SelectedFamily == "" || len(e.SelectedFamily) > 128 || !fontPreviewDigest.MatchString(e.Digest) {
			return fail()
		}
		run := paragraphs[*e.Paragraph].Runs[*e.Run]
		if run.FontFamily == nil || *run.FontFamily != e.SourceFamily || e.SourceFamily == e.SelectedFamily {
			return fail()
		}
		weight, style := 400, "normal"
		if run.Bold != nil && *run.Bold {
			weight = 700
		}
		if run.Italic != nil && *run.Italic {
			style = "italic"
		}
		mapped := false
		for _, m := range config.Policy.Mappings {
			if strings.EqualFold(m.SourceFamily, e.SourceFamily) && strings.EqualFold(m.TargetFamily, e.SelectedFamily) && m.Weight == weight && m.Style == style {
				mapped = true
			}
		}
		selected := false
		for index, f := range config.Faces {
			if strings.EqualFold(f.Family, e.SourceFamily) && f.Weight == weight && f.Style == style {
				return fail()
			}
			if f.Family == e.SelectedFamily && f.Weight == weight && f.Style == style && f.Digest == e.Digest && e.FaceID == fmt.Sprintf("font-%d", index) {
				selected = true
			}
		}
		if !mapped || !selected {
			return fail()
		}
		found := false
		for _, digest := range envelope.FontDigests {
			if digest == e.Digest {
				found = true
			}
		}
		key := fmt.Sprintf("%s/%d/%d", e.SourceID, *e.Paragraph, *e.Run)
		if !found || seen[key] {
			return fail()
		}
		seen[key] = true
	}
	return nil
}
