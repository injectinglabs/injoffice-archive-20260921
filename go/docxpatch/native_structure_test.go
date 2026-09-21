package docxpatch

import (
	"encoding/json"
	"strings"
	"testing"
)

func structurePayload(t *testing.T, p *NativeParagraphV1, fields map[string]any) []byte {
	t.Helper()
	fields["target_kind"], fields["target_id"], fields["expected_xml_sha256"] = "paragraph", p.ID, p.Anchor.XMLSHA256
	data, err := json.Marshal(map[string]any{"mutations": []any{fields}})
	if err != nil {
		t.Fatal(err)
	}
	return data
}

func TestNativeInsertParagraphRoundTrip(t *testing.T) {
	source, doc := nativeFormatSource(t, `<w:p><w:r><w:t>before</w:t></w:r></w:p><w:p><w:r><w:t>after</w:t></w:r></w:p>`)
	payload := structurePayload(t, doc.Body.Blocks[0].Paragraph, map[string]any{"operation": "block.insert_after", "text": ""})
	result, err := ApplyNativeMutationPayloadV1(source, payload, doc.Source.PackageSHA256)
	if err != nil {
		t.Fatal(err)
	}
	if len(result.Document.Body.Blocks) != 3 || nativeStructureText(result.Document.Body.Blocks[1].Paragraph) != "" {
		t.Fatal("missing empty paragraph")
	}
	main := string(readNativeZipPart(t, result.Package, "word/document.xml"))
	original := string(readNativeZipPart(t, source, "word/document.xml"))
	added := `<w:p xmlns:w="` + wordMLTransitional + `"><w:r><w:t></w:t></w:r></w:p>`
	if strings.Replace(main, added, "", 1) != original {
		t.Fatal("unrelated XML changed")
	}
	assertNativeRawPartPreserved(t, source, result.Package, "word/styles.xml")
	if _, err := ApplyNativeMutationPayloadV1(source, payload, result.Document.Source.PackageSHA256); err == nil {
		t.Fatal("accepted stale revision")
	}
}

func TestNativeParagraphSplitRoundTrip(t *testing.T) {
	source, doc := nativeFormatSource(t, `<w:p w14:paraId="01020304"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>A😀B</w:t></w:r><w:r><w:t> tail</w:t></w:r></w:p><w:p><w:r><w:t>neighbor</w:t></w:r></w:p>`)
	p := doc.Body.Blocks[0].Paragraph
	for _, offset := range []int{0, 1, 3, 4} {
		payload := structurePayload(t, p, map[string]any{"operation": "paragraph.split", "split": map[string]any{"run_id": p.Runs[0].ID, "offset_utf16": offset}})
		result, err := ApplyNativeMutationPayloadV1(source, payload, doc.Source.PackageSHA256)
		if err != nil {
			t.Fatal(err)
		}
		left, right := result.Document.Body.Blocks[0].Paragraph, result.Document.Body.Blocks[1].Paragraph
		if nativeStructureText(left)+nativeStructureText(right) != "A😀B tail" || *right.Properties.Alignment != "center" || !*right.Runs[0].Properties.Bold {
			t.Fatal("split lost text or formatting")
		}
		assertNativeRawPartPreserved(t, source, result.Package, "word/styles.xml")
	}
	payload := structurePayload(t, p, map[string]any{"operation": "paragraph.split", "split": map[string]any{"run_id": p.Runs[0].ID, "offset_utf16": 2}})
	if _, err := ApplyNativeMutationPayloadV1(source, payload, doc.Source.PackageSHA256); err == nil {
		t.Fatal("accepted a split surrogate")
	}
}
