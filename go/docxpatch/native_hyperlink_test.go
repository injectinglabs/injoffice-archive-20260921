package docxpatch

import (
	"encoding/json"
	"strings"
	"testing"
)

func hyperlinkPayload(t *testing.T, run NativeRunV1, address any, span *NativeDOCXTextRangeV1) []byte {
	t.Helper()
	link := map[string]any{"url": address}
	if run.Hyperlink != nil {
		link["expected_xml_sha256"] = run.Hyperlink.Anchor.XMLSHA256
	}
	mutation := map[string]any{"target_kind": "run", "target_id": run.ID, "expected_xml_sha256": run.Anchor.XMLSHA256, "operation": "hyperlink.set", "hyperlink": link}
	if span != nil {
		mutation["range"] = map[string]int{"start_utf16": span.StartUTF16, "end_utf16": span.EndUTF16}
	}
	raw, err := json.Marshal(map[string]any{"mutations": []any{mutation}})
	if err != nil {
		t.Fatal(err)
	}
	return raw
}

func TestNativeHyperlinkInsertUpdateRemoveRoundTrip(t *testing.T) {
	source, doc := nativeFormatSource(t, `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>one 😀 two</w:t></w:r></w:p>`)
	run := doc.Body.Blocks[0].Paragraph.Runs[0]
	payload := hyperlinkPayload(t, run, "https://example.com/?a=1&b=2", &NativeDOCXTextRangeV1{StartUTF16: 4, EndUTF16: 6})
	result, err := ApplyNativeMutationPayloadV1(source, payload, doc.Source.PackageSHA256)
	if err != nil {
		t.Fatal(err)
	}
	p := result.Document.Body.Blocks[0].Paragraph
	if len(p.Runs) != 3 || p.Runs[1].Hyperlink == nil || *p.Runs[1].Text != "😀" || !*p.Runs[1].Properties.Bold {
		t.Fatalf("wrong linked runs: %+v", p.Runs)
	}
	if p.Runs[0].Hyperlink != nil || p.Runs[2].Hyperlink != nil {
		t.Fatal("linked unselected text")
	}
	if !strings.Contains(string(readNativeZipPart(t, result.Package, "word/_rels/document.xml.rels")), "a=1&amp;b=2") {
		t.Fatal("relationship missing or not escaped")
	}
	assertNativeRawPartPreserved(t, source, result.Package, "word/styles.xml")
	payload = hyperlinkPayload(t, p.Runs[1], "mailto:editor@example.com", nil)
	updated, err := ApplyNativeMutationPayloadV1(result.Package, payload, result.Document.Source.PackageSHA256)
	if err != nil {
		t.Fatal(err)
	}
	linked := updated.Document.Body.Blocks[0].Paragraph.Runs[1]
	if linked.Hyperlink.URL != "mailto:editor@example.com" {
		t.Fatal("link update lost")
	}
	removed, err := ApplyNativeMutationPayloadV1(updated.Package, hyperlinkPayload(t, linked, nil, nil), updated.Document.Source.PackageSHA256)
	if err != nil {
		t.Fatal(err)
	}
	p = removed.Document.Body.Blocks[0].Paragraph
	if nativeStructureText(p) != "one 😀 two" || p.Runs[1].Hyperlink != nil || !*p.Runs[1].Properties.Bold {
		t.Fatal("remove changed text or formatting")
	}
	assertNativeRawPartPreserved(t, updated.Package, removed.Package, "word/_rels/document.xml.rels")
}

func TestNativeHyperlinkSharedRelationshipAndParagraphSelection(t *testing.T) {
	main := nativeMutationMain(`<w:p><w:hyperlink xmlns:r="` + relNSTransitional + `" r:id="shared"><w:r><w:t>first</w:t></w:r></w:hyperlink><w:r><w:t> rest</w:t></w:r></w:p><w:p><w:hyperlink xmlns:r="` + relNSTransitional + `" r:id="shared"><w:r><w:t>second</w:t></w:r></w:hyperlink></w:p>`)
	parts := nativeMutationParts(main)
	parts["word/_rels/document.xml.rels"] = `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="shared" Type="` + relNSTransitional + `/hyperlink" Target="https://old.example/" TargetMode="External"/></Relationships>`
	source := buildNativeDOCX(t, nativeEntries(parts))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	p := doc.Body.Blocks[0].Paragraph
	payload := structurePayload(t, p, map[string]any{"operation": "hyperlink.set", "hyperlink": map[string]any{"url": "https://new.example/"}, "range": map[string]int{"start_utf16": 2, "end_utf16": 8}})
	result, err := ApplyNativeMutationPayloadV1(source, payload, doc.Source.PackageSHA256)
	if err != nil {
		t.Fatal(err)
	}
	if result.Document.Body.Blocks[1].Paragraph.Runs[0].Hyperlink.URL != "https://old.example/" {
		t.Fatal("shared relationship was retargeted")
	}
	got := result.Document.Body.Blocks[0].Paragraph
	if nativeStructureText(got) != "first rest" || got.Runs[0].Hyperlink.URL != "https://old.example/" || got.Runs[1].Hyperlink.URL != "https://new.example/" || got.Runs[2].Hyperlink.URL != "https://new.example/" || got.Runs[3].Hyperlink != nil {
		t.Fatalf("selection did not survive: %+v", got.Runs)
	}
}

func TestNativeHyperlinkRefusesInvalidAndStalePayloads(t *testing.T) {
	source, doc := nativeFormatSource(t, `<w:p><w:r><w:t>A😀B</w:t></w:r></w:p>`)
	run := doc.Body.Blocks[0].Paragraph.Runs[0]
	for _, address := range []string{"javascript:alert(1)", "file:///tmp/file", "https://example.com/\n", "https://user:pass@example.com", "mailto:", "https://"} {
		if result, err := ApplyNativeMutationPayloadV1(source, hyperlinkPayload(t, run, address, nil), doc.Source.PackageSHA256); err == nil || result != nil {
			t.Fatalf("accepted unsafe URL %q", address)
		}
	}
	if result, err := ApplyNativeMutationPayloadV1(source, hyperlinkPayload(t, run, "https://example.com", &NativeDOCXTextRangeV1{StartUTF16: 1, EndUTF16: 2}), doc.Source.PackageSHA256); err == nil || result != nil {
		t.Fatal("accepted split surrogate")
	}
	payload := hyperlinkPayload(t, run, "https://example.com", nil)
	if _, err := ApplyNativeMutationPayloadV1(source, payload, "sha256:"+strings.Repeat("0", 64)); err == nil {
		t.Fatal("accepted stale revision")
	}
	run.Anchor.XMLSHA256 = "sha256:" + strings.Repeat("0", 64)
	if _, err := ApplyNativeMutationPayloadV1(source, hyperlinkPayload(t, run, "https://example.com", nil), doc.Source.PackageSHA256); err == nil {
		t.Fatal("accepted stale anchor")
	}
}

func TestNativeHyperlinkStrictSelfClosingRelationships(t *testing.T) {
	main := strings.ReplaceAll(nativeMutationMain(`<w:p><w:r><w:t>strict</w:t></w:r></w:p>`), wordMLTransitional, wordMLStrict)
	parts := nativeMutationParts(main)
	parts["_rels/.rels"] = strings.ReplaceAll(parts["_rels/.rels"], relNSTransitional, relNSStrict)
	parts["word/_rels/document.xml.rels"] = `<Relationships xmlns="` + opcRelationshipsNS + `" />`
	source := buildNativeDOCX(t, nativeEntries(parts))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	result, err := ApplyNativeMutationPayloadV1(source, hyperlinkPayload(t, doc.Body.Blocks[0].Paragraph.Runs[0], "https://example.com", nil), doc.Source.PackageSHA256)
	if err != nil {
		t.Fatal(err)
	}
	if result.Document.Body.Blocks[0].Paragraph.Runs[0].Hyperlink.URL != "https://example.com" {
		t.Fatal("strict link missing")
	}
}

func TestNativeHyperlinkRefusesComplexWrappersAndStaleWrapperAnchor(t *testing.T) {
	parts := nativeMutationParts(nativeMutationMain(`<w:p><w:hyperlink xmlns:r="` + relNSTransitional + `" r:id="link"><w:r><w:t>one</w:t></w:r><w:r><w:t>two</w:t></w:r></w:hyperlink></w:p>`))
	parts["word/_rels/document.xml.rels"] = `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="link" Type="` + relNSTransitional + `/hyperlink" Target="https://example.com" TargetMode="External"/></Relationships>`
	source := buildNativeDOCX(t, nativeEntries(parts))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	if doc.Body.Blocks[0].Paragraph.EditPolicy.Mode != "read-only" || doc.Body.Blocks[0].Paragraph.Runs[0].CanEditHyperlink {
		t.Fatal("complex wrapper advertised edits")
	}
	parts["word/document.xml"] = strings.Replace(parts["word/document.xml"], `<w:r><w:t>two</w:t></w:r>`, "", 1)
	source = buildNativeDOCX(t, nativeEntries(parts))
	doc, err = ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	run := doc.Body.Blocks[0].Paragraph.Runs[0]
	run.Hyperlink.Anchor.XMLSHA256 = "sha256:" + strings.Repeat("0", 64)
	if result, err := ApplyNativeMutationPayloadV1(source, hyperlinkPayload(t, run, nil, nil), doc.Source.PackageSHA256); err == nil || result != nil {
		t.Fatal("accepted stale wrapper anchor")
	}
}
