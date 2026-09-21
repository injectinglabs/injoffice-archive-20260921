package docxpatch

import (
	"bytes"
	"encoding/base64"
	"encoding/json"
	"image"
	"image/png"
	"strings"
	"testing"
)

func TestNativeInsertPageBreakCaret(t *testing.T) {
	source, doc := nativeFormatSource(t, `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>A😀 B</w:t></w:r></w:p>`)
	p := doc.Body.Blocks[0].Paragraph
	offset := 3
	m := nativeInsertMutation{TargetKind: "paragraph", TargetID: p.ID, SHA: p.Anchor.XMLSHA256, Operation: "page_break.insert", RunID: p.Runs[0].ID, Offset: &offset}
	payload, _ := marshalNativeInsertTest(m)
	result, err := ApplyNativeMutationPayloadV1(source, payload, doc.Source.PackageSHA256)
	if err != nil {
		t.Fatal(err)
	}
	xml := string(readNativeZipPart(t, result.Package, docPart))
	if !strings.Contains(xml, `<w:t>A😀</w:t><w:br xmlns:w="`+wordMLTransitional+`" w:type="page"/><w:t xml:space="preserve"> B</w:t>`) {
		t.Fatal(xml)
	}
	assertNativeRawPartPreserved(t, source, result.Package, "word/styles.xml")
	for _, bad := range []int{-1, 2, 99} {
		m.Offset = &bad
		payload, _ = marshalNativeInsertTest(m)
		if output, err := ApplyNativeMutationPayloadV1(source, payload, doc.Source.PackageSHA256); err == nil || output != nil {
			t.Fatalf("accepted offset %d", bad)
		}
	}
}

func TestNativeInsertImageAndEmptyParagraph(t *testing.T) {
	source, doc := nativeFormatSource(t, `<w:p><w:r><w:t>Keep</w:t></w:r></w:p>`)
	p := doc.Body.Blocks[0].Paragraph
	var data bytes.Buffer
	if err := png.Encode(&data, image.NewNRGBA(image.Rect(0, 0, 4, 3))); err != nil {
		t.Fatal(err)
	}
	m := nativeInsertMutation{TargetKind: "paragraph", TargetID: p.ID, SHA: p.Anchor.XMLSHA256, Operation: "block.insert_after", Image: &nativeInsertImage{Data: base64.StdEncoding.EncodeToString(data.Bytes()), ContentType: "image/png", Width: 38100, Height: 28575, Alt: "A & B"}}
	apply := func(m nativeInsertMutation) (*NativeDOCXMutationResultV1, error) {
		payload, _ := marshalNativeInsertTest(m)
		return ApplyNativeMutationPayloadV1(source, payload, doc.Source.PackageSHA256)
	}
	result, err := apply(m)
	if err != nil {
		t.Fatal(err)
	}
	if len(result.Document.Body.Blocks) != 3 {
		t.Fatal("image must include empty continuation paragraph")
	}
	d := result.Document.Body.Blocks[1].Paragraph.Runs[0].Drawing
	if d.AltText == nil || *d.AltText != "A & B" || !bytes.Equal(readNativeZipPart(t, result.Package, *d.MediaPart), data.Bytes()) {
		t.Fatal("image bytes/alt mismatch")
	}
	assertNativeRawPartPreserved(t, source, result.Package, "word/styles.xml")
	badImage := append([]byte(nil), data.Bytes()...)
	badImage[len(badImage)/2] ^= 1
	originalImage := m.Image.Data
	m.Image.Data = base64.StdEncoding.EncodeToString(badImage)
	if out, err := apply(m); err == nil || out != nil {
		t.Fatal("accepted corrupt image")
	}
	m.Image.Data = originalImage
	m.Image.ContentType = "image/jpeg"
	if out, err := apply(m); err == nil || out != nil {
		t.Fatal("accepted mismatched MIME")
	}
	m.Image = nil
	text := ""
	m.Text = &text
	result, err = apply(m)
	if err != nil {
		t.Fatal(err)
	}
	if len(result.Document.Body.Blocks) != 2 {
		t.Fatal("paragraph missing")
	}
	m.SHA = strings.Repeat("0", 64)
	if out, err := apply(m); err == nil || out != nil {
		t.Fatal("accepted stale target")
	}
}

func TestNativeInsertRejectsDuplicateAndUnknownFields(t *testing.T) {
	source, doc := nativeFormatSource(t, `<w:p><w:r><w:t>Keep</w:t></w:r></w:p>`)
	for _, payload := range []string{`{"mutations":[{"operation":"block.insert_after","operation":"block.insert_after"}]}`, `{"mutations":[{"operation":"block.insert_after","table":{}}]}`} {
		if out, err := ApplyNativeMutationPayloadV1(source, []byte(payload), doc.Source.PackageSHA256); err == nil || out != nil {
			t.Fatal("accepted invalid payload")
		}
	}
}

func marshalNativeInsertTest(m nativeInsertMutation) ([]byte, error) {
	raw := map[string]any{"target_kind": m.TargetKind, "target_id": m.TargetID, "expected_xml_sha256": m.SHA, "operation": m.Operation}
	if m.Operation == "page_break.insert" {
		raw["split"] = map[string]any{"run_id": m.RunID, "offset_utf16": m.Offset}
	} else if m.Image != nil {
		raw["image"] = m.Image
	} else {
		raw["text"] = m.Text
	}
	return json.Marshal(map[string]any{"mutations": []any{raw}})
}
