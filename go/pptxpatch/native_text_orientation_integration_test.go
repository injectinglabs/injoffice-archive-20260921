package pptxpatch

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestNativeTextOrientationSourceAuthority(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, attrs := range []string{`rot="-5400000"`, `upright="1"`, `rot="21600000" upright="false"`, `rot="-2147483648" upright="true"`} {
			input := nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, mutate: func(parts map[string]string) {
				parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `<a:bodyPr/>`, `<a:bodyPr `+attrs+`/>`, 1)
			}})
			before := string(input)
			deck, err := ExtractNativePPTX(input, nativeMutationExtractOptions())
			if err != nil {
				t.Fatal(err)
			}
			target := deck.Slides[0].Elements[0]
			if target.TextBody == nil || target.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || !nativeDiagnosticsContain(target.Compatibility.Diagnostics, "pptx.text-orientation-preview") {
				t.Fatalf("missing source orientation/authority: %#v", target)
			}
			if issues := ValidateNativePPTX(deck); len(issues) != 0 {
				t.Fatal(issues)
			}
			paragraphs := nativeMutationParagraphs("Must remain unchanged")
			output, err := ApplyNativePPTXMutations(input, NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{{OperationID: "replace", Kind: NativePPTXReplaceText, ElementID: target.ID, ExpectedFingerprintSHA256: target.Source.FingerprintSHA256, Paragraphs: &paragraphs}}})
			if err == nil || !strings.Contains(err.Error(), "preview-only") || len(output) != 0 || string(input) != before {
				t.Fatalf("orientation mutation authority leaked: %v", err)
			}
		}
	}
}

// Functional source fixtures, consumed by the actual WASM/font-worker/browser proof.
func TestNativeTextOrientationBrowserFixtures(t *testing.T) {
	dir := os.Getenv("INJOFFICE_PPTX_TEXT_ORIENTATION_FIXTURE_DIR")
	if dir == "" {
		t.Skip("external proof fixture generation")
	}
	if err := os.MkdirAll(dir, 0755); err != nil {
		t.Fatal(err)
	}
	for _, c := range []struct{ name, shape, body string }{
		{"body-counter-rotation", ` rot="5400000"`, `rot="-5400000"`},
		{"upright-anisotropic", ` rot="5400000"`, `upright="true" rot="1800000"`},
		{"upright-above-slide-clip", ` rot="5400000"`, `upright="true" rot="1800000"`},
		{"upright-reflected", ` rot="1800000" flipH="1"`, `upright="true"`},
	} {
		input := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
			raw := parts["relocated/slides/slide-a.xml"]
			raw = strings.ReplaceAll(raw, `typeface="Aptos"`, `typeface="DejaVu Sans"`)
			raw = strings.ReplaceAll(raw, `b="1"`, `b="0"`)
			raw = strings.ReplaceAll(raw, `sz="3200"`, `sz="1200"`)
			if c.name != "upright-above-slide-clip" {
				raw = strings.Replace(raw, `<a:off x="914400" y="457200"/>`, `<a:off x="914400" y="3200000"/>`, 1)
			}
			raw = strings.Replace(raw, `<a:xfrm>`, `<a:xfrm`+c.shape+`>`, 1)
			raw = strings.Replace(raw, `<a:bodyPr/>`, `<a:bodyPr `+c.body+` lIns="100000" rIns="200000" tIns="30000" bIns="50000"/>`, 1)
			if c.name == "upright-anisotropic" || c.name == "upright-above-slide-clip" {
				start, end := strings.Index(raw, "<p:sp>"), strings.Index(raw, "</p:sp>")+len("</p:sp>")
				group := `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="99" name="Anisotropic upright"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9000000" cy="4000000"/><a:chOff x="0" y="0"/><a:chExt cx="6000000" cy="4000000"/></a:xfrm></p:grpSpPr>` + raw[start:end] + `</p:grpSp>`
				raw = raw[:start] + group + raw[end:]
			}
			parts["relocated/slides/slide-a.xml"] = raw
		}})
		deck, err := ExtractNativePPTX(input, nativeAtomicTestExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		if issues := ValidateNativePPTX(deck); len(issues) != 0 {
			t.Fatal(issues)
		}
		encoded, err := json.MarshalIndent(deck, "", "  ")
		if err != nil {
			t.Fatal(err)
		}
		for ext, data := range map[string][]byte{".pptx": input, "-go.json": encoded} {
			if err := os.WriteFile(filepath.Join(dir, c.name+ext), data, 0644); err != nil {
				t.Fatal(err)
			}
		}
		fmt.Printf("%s: %s\n", c.name, deck.Slides[0].Elements[0].Compatibility.Status)
	}
}
