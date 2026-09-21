//go:build !js

package main

import (
	"bytes"
	"encoding/base64"
	"encoding/json"
	"image"
	"image/png"
	"os/exec"
	"testing"

	"github.com/injectinglabs/injoffice/go/docxpatch"
)

func TestWASMInsertionMatchesNativeBytes(t *testing.T) {
	wasm, runtime, script := requireNodeHarness(t)
	source := buildContractDOCX(t)
	doc, _ := extractNativeJSON(t, source)
	p := doc.Body.Blocks[0].Paragraph
	var picture bytes.Buffer
	if err := png.Encode(&picture, image.NewNRGBA(image.Rect(0, 0, 4, 3))); err != nil {
		t.Fatal(err)
	}
	for _, extra := range []map[string]any{
		{"operation": "page_break.insert", "split": map[string]any{"run_id": p.Runs[0].ID, "offset_utf16": 1}},
		{"operation": "block.insert_after", "image": map[string]any{"data_base64": base64.StdEncoding.EncodeToString(picture.Bytes()), "content_type": "image/png", "width_emu": 9525, "height_emu": 9525, "alt_text": "test"}},
	} {
		extra["target_kind"] = "paragraph"
		extra["target_id"] = p.ID
		extra["expected_xml_sha256"] = p.Anchor.XMLSHA256
		payload, err := json.Marshal(map[string]any{"mutations": []any{extra}})
		if err != nil {
			t.Fatal(err)
		}
		want, err := docxpatch.ApplyNativeMutationPayloadV1(source, payload, doc.Source.PackageSHA256)
		if err != nil {
			t.Fatal(err)
		}
		originalPath, payloadPath := writeContractInputs(t, source, payload)
		command := exec.Command("node", script, "apply", "--wasm", wasm, "--wasm-exec", runtime, "--original", originalPath, "--payload", payloadPath, "--expected-revision", doc.Source.PackageSHA256)
		got, err := command.Output()
		if err != nil {
			t.Fatalf("WASM insert failed: %v %s", err, stderrFrom(err))
		}
		if !bytes.Equal(got, want.Package) {
			t.Fatal("WASM insertion differs from native output")
		}
	}
}
