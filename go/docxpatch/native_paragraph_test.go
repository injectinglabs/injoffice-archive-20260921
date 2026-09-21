package docxpatch

import (
	"fmt"
	"strings"
	"testing"
)

func TestNativeParagraphLayoutRoundTrip(t *testing.T) {
	for _, prefix := range []string{"w", "word"} {
		t.Run(prefix, func(t *testing.T) {
			main := nativeMutationMain(`<w:p w14:paraId="01020304"><w:pPr><!--keep--><w:keepNext/><w:spacing w:before='20' w:after="40"/><w:ind w:left="100"/><w:jc w:val="center"/></w:pPr><w:r><w:t>unchanged</w:t></w:r></w:p>`)
			main = strings.ReplaceAll(strings.ReplaceAll(main, "w:", prefix+":"), "xmlns:w=", "xmlns:"+prefix+"=")
			source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(main)))
			doc, err := ExtractNativeDocumentV1(source)
			if err != nil {
				t.Fatal(err)
			}
			p := doc.Body.Blocks[0].Paragraph
			payload := []byte(fmt.Sprintf(`{"mutations":[{"target_kind":"paragraph","target_id":%q,"expected_xml_sha256":%q,"properties":{"spacing_before_twips":120,"spacing_after_twips":240,"indent_left_twips":720,"indent_right_twips":-120,"first_line_twips":360,"line_spacing":360,"line_rule":"auto"}}]}`, p.ID, p.Anchor.XMLSHA256))
			result, err := ApplyNativeMutationPayloadV1(source, payload, doc.Source.PackageSHA256)
			if err != nil {
				t.Fatal(err)
			}
			got := string(readNativeZipPart(t, result.Package, "word/document.xml"))
			for _, keep := range []string{"<!--keep-->", "<" + prefix + ":keepNext/>", "<" + prefix + `:jc ` + prefix + `:val="center"/>`, `unchanged`} {
				if !strings.Contains(got, keep) {
					t.Fatalf("lost %s: %s", keep, got)
				}
			}
			assertNativeRawPartPreserved(t, source, result.Package, "word/styles.xml")
			p = result.Document.Body.Blocks[0].Paragraph
			if p.EditPolicy.Mode != "read-write" {
				t.Fatalf("lost write policy: %#v", p.EditPolicy)
			}
			payload = []byte(fmt.Sprintf(`{"mutations":[{"target_kind":"paragraph","target_id":%q,"expected_xml_sha256":%q,"properties":{"spacing_before_twips":null,"first_line_twips":null,"hanging_twips":240,"line_spacing":480,"line_rule":"exact"}}]}`, p.ID, p.Anchor.XMLSHA256))
			cleared, err := ApplyNativeMutationPayloadV1(result.Package, payload, result.Document.Source.PackageSHA256)
			if err != nil {
				t.Fatal(err)
			}
			props := cleared.Document.Body.Blocks[0].Paragraph.Properties
			if props.SpacingBeforeTwips != nil || props.FirstLineTwips != nil || *props.HangingTwips != 240 || *props.SpacingAfterTwips != 240 {
				t.Fatalf("clear/merge failed: %#v", props)
			}
		})
	}
}

func TestNativeParagraphLayoutRefusals(t *testing.T) {
	for _, properties := range []string{`{"line_spacing":1.5}`, `{"indent_left_twips":-31681}`, `{"spacing_before_twips":-1}`, `{"line_rule":"wrong"}`, `{"hanging_twips":1,"first_line_twips":2}`, `{"spacing_before_twips":1,"bold":true}`, `{"spacing_before_twips":1,"spacing_before_twips":2}`} {
		payload := []byte(fmt.Sprintf(`{"mutations":[{"target_kind":"paragraph","target_id":"p1","expected_xml_sha256":"%s","properties":%s}]}`, strings.Repeat("a", 64), properties))
		if _, err := DecodeNativeDOCXFormatMutationPayloadV1(payload); err == nil {
			t.Fatalf("accepted %s", properties)
		}
	}
}

func TestNativeParagraphLayoutCreatesProperties(t *testing.T) {
	for _, body := range []string{`<w:p><w:r><w:t>text</w:t></w:r></w:p>`, `<w:p><w:pPr/><w:r><w:t>text</w:t></w:r></w:p>`} {
		source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(nativeMutationMain(body))))
		doc, err := ExtractNativeDocumentV1(source)
		if err != nil {
			t.Fatal(err)
		}
		p := doc.Body.Blocks[0].Paragraph
		payload := []byte(fmt.Sprintf(`{"mutations":[{"target_kind":"paragraph","target_id":%q,"expected_xml_sha256":%q,"properties":{"spacing_after_twips":120,"indent_left_twips":720}}]}`, p.ID, p.Anchor.XMLSHA256))
		if _, err := ApplyNativeMutationPayloadV1(source, payload, doc.Source.PackageSHA256); err != nil {
			t.Fatal(err)
		}
	}
}
