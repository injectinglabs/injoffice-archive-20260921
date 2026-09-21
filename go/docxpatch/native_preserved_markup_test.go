package docxpatch

import (
	"bytes"
	"strings"
	"testing"
)

// Exercise the public mutation boundary and compare every XML byte, not just
// tag counts: a text edit must not normalize or discard opaque source markup.
func assertNativeTextPreservationPolicy(t *testing.T, source []byte, doc *NativeDocumentV1, editable bool) {
	t.Helper()
	p := doc.Body.Blocks[0].Paragraph
	if got := p.EditPolicy.Mode == "read-write" && nativePolicyAllows(p.EditPolicy, "text.replace"); got != editable {
		t.Fatalf("text authority = %#v, want editable=%v", p.EditPolicy, editable)
	}
	var run *NativeRunV1
	for i := range p.Runs {
		if p.Runs[i].Kind == "text" && p.Runs[i].Text != nil {
			run = &p.Runs[i]
			break
		}
	}
	if run == nil {
		t.Fatal("fixture has no text run")
	}
	original := bytes.Clone(source)
	result, err := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXTextMutationV1{{TargetKind: "run", TargetID: run.ID, ExpectedXMLSHA256: run.Anchor.XMLSHA256, Text: "Edited"}})
	if !bytes.Equal(original, source) {
		t.Fatal("mutation modified its input")
	}
	if !editable {
		if err == nil || result != nil {
			t.Fatal("unsafe text mutation was not refused atomically")
		}
		return
	}
	if err != nil {
		t.Fatal(err)
	}
	part := string(readNativeZipPart(t, source, run.Anchor.PartName))
	start, end := *run.Anchor.StartByte, *run.Anchor.EndByte
	raw := part[start:end]
	contentStart, contentEnd := strings.Index(raw, ">")+1, strings.LastIndex(raw, "<")
	want := part[:start] + raw[:contentStart] + "Edited" + raw[contentEnd:] + part[end:]
	if got := string(readNativeZipPart(t, result.Package, run.Anchor.PartName)); got != want {
		t.Fatalf("text edit changed preserved markup:\ngot %s\nwant %s", got, want)
	}
	if len(result.Evidence.ChangedParts) != 1 {
		t.Fatalf("unexpected changed parts: %#v", result.Evidence)
	}
	reopened, err := ExtractNativeDocumentV1(result.Package)
	if err != nil {
		t.Fatal(err)
	}
	if *reopened.Body.Blocks[0].Paragraph.Runs[0].Text != "Edited" {
		t.Fatal("edited text did not survive re-extraction")
	}
}

func assertNativePropertyPatchRefused(t *testing.T, source []byte, doc *NativeDocumentV1) {
	t.Helper()
	p := doc.Body.Blocks[0].Paragraph
	if nativePolicyAllows(p.EditPolicy, "properties.patch") {
		t.Fatal("preserved property block admits properties.patch")
	}
	before := bytes.Clone(source)
	for _, mutation := range []NativeDOCXFormatMutationV1{
		{TargetKind: "paragraph", TargetID: p.ID, ExpectedXMLSHA256: p.Anchor.XMLSHA256, ParagraphProperties: &NativeDOCXParagraphPropertyPatchV1{Alignment: nativeString("center")}},
		{TargetKind: "run", TargetID: p.Runs[0].ID, ExpectedXMLSHA256: p.Runs[0].Anchor.XMLSHA256, Properties: NativeDOCXRunPropertyPatchV1{Bold: nativeBool(true)}},
	} {
		result, err := ApplyNativeFormatMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXFormatMutationV1{mutation})
		if err == nil || result != nil || !strings.Contains(err.Error(), "UNSUPPORTED_CONSTRUCT") {
			t.Fatalf("property rewrite not refused atomically: %v", err)
		}
	}
	if !bytes.Equal(before, source) {
		t.Fatal("refused patch changed source")
	}
}

func TestNativePreservedMarkupTextRoundTrips(t *testing.T) {
	for _, tc := range []struct {
		name, paragraph, run string
		patchRefused         bool
	}{
		{"mark textFill", `<w:rPr><w14:textFill><w14:solidFill><w14:srgbClr w14:val="13579B"/></w14:solidFill></w14:textFill></w:rPr>`, "", true},
		{"font hint", "", `<w:rFonts w:ascii="Arial" w:hint="eastAsia"/>`, false},
		{"mark font hint", `<w:rPr><w:rFonts w:ascii="Arial" w:hint="default"/></w:rPr>`, "", true},
		{"bare fonts", "", `<w:rFonts/>`, false},
		{"foreign property", "", `<x:decoration xmlns:x="urn:synthetic"><x:color value="keep"/></x:decoration>`, false},
		{"shading", "", `<w:shd w:val="clear" w:color="auto" w:fill="F1F2F3"/>`, false},
		{"kerning", "", `<w:kern w:val="0"/>`, false},
		{"unknown paragraph property", `<w:contextualSpacing w:val="1"/>`, "", true},
	} {
		t.Run(tc.name, func(t *testing.T) {
			for _, strict := range []bool{false, true} {
				main := nativeMutationMain(`<w:p><w:pPr>` + tc.paragraph + `</w:pPr><w:r><w:rPr>` + tc.run + `</w:rPr><w:t>Before</w:t></w:r></w:p>`)
				parts := nativeMutationParts(main)
				if strict {
					for name, value := range parts {
						parts[name] = strings.ReplaceAll(strings.ReplaceAll(value, wordMLTransitional, wordMLStrict), relBaseTransitional, relBaseStrict)
					}
				}
				source := buildNativeDOCX(t, nativeEntries(parts))
				doc, err := ExtractNativeDocumentV1(source)
				if err != nil {
					t.Fatal(err)
				}
				assertNativeTextPreservationPolicy(t, source, doc, true)
				if tc.patchRefused {
					assertNativePropertyPatchRefused(t, source, doc)
				}
				for _, operation := range []string{"text.replace", "block.insert_after", "paragraph.split", "page_break.insert", "hyperlink.set"} {
					if !nativePolicyAllows(doc.Body.Blocks[0].Paragraph.EditPolicy, operation) {
						t.Fatalf("lost %s", operation)
					}
				}
			}
		})
	}
}
