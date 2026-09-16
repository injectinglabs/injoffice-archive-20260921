package docxpatch

import (
	"bytes"
	"strings"
	"testing"
)

// Unmodeled paragraph-property neighbors that resolved layout does not round-trip.
// Extract preserves them verbatim and marks the owning paragraph read-only.
const nativeWritebackUnmodeledPPr = `<w:pPr>` +
	`<w:pBdr>` +
	`<w:top w:val="dashed" w:sz="24" w:space="1" w:color="C00000"/>` +
	`<w:bottom w:val="dashed" w:sz="18" w:space="4" w:color="003399"/>` +
	`</w:pBdr>` +
	`<w:shd w:val="pct10" w:color="auto" w:fill="F2F2F2"/>` +
	`<w:spacing w:before="240" w:after="0" w:beforeLines="100" w:afterLines="200" w:line="360" w:lineRule="auto"/>` +
	`<w:ind w:firstLineChars="200" w:firstLine="480"/>` +
	`<w:jc w:val="both"/>` +
	`</w:pPr>`

const nativeWritebackModeledPPr = `<w:pPr>` +
	`<w:pStyle w:val="Keep"/>` +
	`<w:keepNext/>` +
	`<w:keepLines/>` +
	`<w:pageBreakBefore/>` +
	`<w:widowControl/>` +
	`<w:jc w:val="both"/>` +
	`</w:pPr>`

const nativeWritebackRunRPr = `<w:rPr>` +
	`<w:b/>` +
	`<w:color w:val="004080"/>` +
	`<w:rFonts w:ascii="Calibri" w:eastAsia="宋体"/>` +
	`</w:rPr>`

func nativeWritebackFidelityParts(body string) map[string]string {
	parts := nativeMutationParts(nativeMutationMain(body))
	parts["word/opaque.xml"] = `<opaque xmlns="urn:injoffice:writeback-fidelity">preserve raw storage</opaque>`
	return parts
}

func TestApplyNativeTextMutationsV1PreservesUnmodeledParagraphPropertyNeighbors(t *testing.T) {
	rich := `<w:p w14:paraId="01020304">` + nativeWritebackUnmodeledPPr + `<w:r><w:t>Rich</w:t></w:r></w:p>`
	mutable := `<w:p w14:paraId="0A0B0C0D">` + nativeWritebackModeledPPr + `<w:r><w:t>Mutable</w:t></w:r></w:p>`
	parts := nativeWritebackFidelityParts(rich + mutable)
	source := buildNativeDOCX(t, nativeEntries(parts))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	if len(doc.Body.Blocks) != 2 || doc.Body.Blocks[0].Paragraph == nil || doc.Body.Blocks[1].Paragraph == nil {
		t.Fatalf("body blocks = %#v", doc.Body.Blocks)
	}
	richParagraph := doc.Body.Blocks[0].Paragraph
	mutableParagraph := doc.Body.Blocks[1].Paragraph
	if len(richParagraph.Runs) != 1 || richParagraph.Runs[0].Text == nil || *richParagraph.Runs[0].Text != "Rich" {
		t.Fatalf("rich paragraph runs = %#v", richParagraph.Runs)
	}

	sourceBefore := bytes.Clone(source)
	var produced []byte
	if nativeWritebackParagraphMutable(richParagraph) {
		result := mustApplyNativeWritebackText(t, source, doc, "paragraph", richParagraph.ID, richParagraph.Anchor.XMLSHA256, "Changed")
		assertNativeWritebackExactTextSplice(t, parts, source, result.Package, "Rich", "Changed")
		if got := *result.Document.Body.Blocks[0].Paragraph.Runs[0].Text; got != "Changed" {
			t.Fatalf("reopened rich text = %q", got)
		}
		produced = result.Package
	} else {
		assertNativeWritebackRefusesConstruct(t, source, doc, richParagraph)
		if !bytes.Equal(source, sourceBefore) {
			t.Fatal("refused unmodeled pPr mutation mutated caller-owned source bytes")
		}
		if got := string(readNativeZipPart(t, source, "word/document.xml")); got != parts["word/document.xml"] {
			t.Fatalf("refused mutation rewrote source XML:\n got %s\nwant %s", got, parts["word/document.xml"])
		}
		if !nativeWritebackParagraphMutable(mutableParagraph) {
			t.Fatalf("modeled pPr paragraph should remain writable: %#v", mutableParagraph.EditPolicy)
		}
		result := mustApplyNativeWritebackText(t, source, doc, "paragraph", mutableParagraph.ID, mutableParagraph.Anchor.XMLSHA256, "Changed")
		assertNativeWritebackExactTextSplice(t, parts, source, result.Package, "Mutable", "Changed")
		if got := *result.Document.Body.Blocks[1].Paragraph.Runs[0].Text; got != "Changed" {
			t.Fatalf("reopened mutable text = %q", got)
		}
		if got := *result.Document.Body.Blocks[0].Paragraph.Runs[0].Text; got != "Rich" {
			t.Fatalf("unmodeled pPr paragraph text changed = %q", got)
		}
		produced = result.Package
	}

	gotMain := string(readNativeZipPart(t, produced, "word/document.xml"))
	for _, fragment := range []string{
		nativeWritebackUnmodeledPPr,
		`<w:top w:val="dashed" w:sz="24" w:space="1" w:color="C00000"/>`,
		`<w:bottom w:val="dashed" w:sz="18" w:space="4" w:color="003399"/>`,
		`<w:shd w:val="pct10" w:color="auto" w:fill="F2F2F2"/>`,
		`w:before="240"`,
		`w:after="0"`,
		`w:beforeLines="100"`,
		`w:afterLines="200"`,
		`w:firstLineChars="200"`,
		`w:firstLine="480"`,
		`<w:jc w:val="both"/>`,
	} {
		if !strings.Contains(gotMain, fragment) {
			t.Fatalf("unmodeled pPr neighbor %q lost original spelling:\n%s", fragment, gotMain)
		}
	}
}

func TestApplyNativeTextMutationsV1PreservesRunPropertyWrappers(t *testing.T) {
	body := `<w:p w14:paraId="11223344"><w:r>` + nativeWritebackRunRPr + `<w:t>Styled</w:t></w:r></w:p>`
	parts := nativeWritebackFidelityParts(body)
	source := buildNativeDOCX(t, nativeEntries(parts))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	paragraph := doc.Body.Blocks[0].Paragraph
	if !nativeWritebackParagraphMutable(paragraph) {
		t.Fatalf("latin/eastAsia run should remain writable: %#v", paragraph.EditPolicy)
	}
	mutation := []NativeDOCXTextMutationV1{{TargetKind: "run", TargetID: paragraph.Runs[0].ID, ExpectedXMLSHA256: paragraph.Runs[0].Anchor.XMLSHA256, Text: "Updated"}}
	result, err := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, mutation)
	if err != nil {
		t.Fatal(err)
	}
	assertNativeWritebackExactTextSplice(t, parts, source, result.Package, "Styled", "Updated")
	assertNativeWritebackDeterministic(t, source, doc.Source.PackageSHA256, mutation, result.Package)
	gotMain := string(readNativeZipPart(t, result.Package, "word/document.xml"))
	if !strings.Contains(gotMain, nativeWritebackRunRPr) {
		t.Fatalf("run rPr wrapper changed: %s", gotMain)
	}
	if !strings.Contains(gotMain, `w:ascii="Calibri"`) || !strings.Contains(gotMain, `w:eastAsia="宋体"`) {
		t.Fatalf("dual latin/eastAsia rFonts were not both preserved: %s", gotMain)
	}
	if got := *result.Document.Body.Blocks[0].Paragraph.Runs[0].Text; got != "Updated" {
		t.Fatalf("reopened run text = %q", got)
	}
}

func TestApplyNativeTextMutationsV1PreservesXMLSpaceOnInteriorRunText(t *testing.T) {
	body := `<w:p w14:paraId="22222222"><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> keep  interior </w:t></w:r></w:p>`
	parts := nativeWritebackFidelityParts(body)
	source := buildNativeDOCX(t, nativeEntries(parts))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	paragraph := doc.Body.Blocks[0].Paragraph
	if !nativeWritebackParagraphMutable(paragraph) {
		t.Fatalf("preserved-space run should remain writable: %#v", paragraph.EditPolicy)
	}
	run := paragraph.Runs[0]
	if run.Text == nil || *run.Text != " keep  interior " {
		t.Fatalf("preserved-space projection = %#v", run.Text)
	}
	result := mustApplyNativeWritebackText(t, source, doc, "run", run.ID, run.Anchor.XMLSHA256, " keep  changed ")
	assertNativeWritebackExactTextSplice(t, parts, source, result.Package, " keep  interior ", " keep  changed ")
	gotMain := string(readNativeZipPart(t, result.Package, "word/document.xml"))
	if !strings.Contains(gotMain, `<w:t xml:space="preserve"> keep  changed </w:t>`) {
		t.Fatalf("xml:space=preserve was dropped or rewritten: %s", gotMain)
	}
	if strings.Contains(gotMain, `> keep  interior <`) {
		t.Fatal("interior preserved-space text was not replaced")
	}
}

func TestApplyNativeTextMutationsV1PreservesExplicitZeroAfterSpacing(t *testing.T) {
	spaced := `<w:p w14:paraId="33333333"><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t>Spaced</w:t></w:r></w:p>`
	other := `<w:p w14:paraId="44444444"><w:r><w:t>Other</w:t></w:r></w:p>`
	parts := nativeWritebackFidelityParts(spaced + other)
	source := buildNativeDOCX(t, nativeEntries(parts))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	spacedParagraph := doc.Body.Blocks[0].Paragraph
	otherParagraph := doc.Body.Blocks[1].Paragraph
	sourceBefore := bytes.Clone(source)
	var produced []byte
	if nativeWritebackParagraphMutable(spacedParagraph) {
		result := mustApplyNativeWritebackText(t, source, doc, "run", spacedParagraph.Runs[0].ID, spacedParagraph.Runs[0].Anchor.XMLSHA256, "Changed")
		assertNativeWritebackExactTextSplice(t, parts, source, result.Package, "Spaced", "Changed")
		produced = result.Package
	} else {
		assertNativeWritebackRefusesConstruct(t, source, doc, spacedParagraph)
		if !bytes.Equal(source, sourceBefore) {
			t.Fatal("refused after=0 mutation mutated caller-owned source bytes")
		}
		if !nativeWritebackParagraphMutable(otherParagraph) {
			t.Fatalf("sibling paragraph should remain writable: %#v", otherParagraph.EditPolicy)
		}
		result := mustApplyNativeWritebackText(t, source, doc, "run", otherParagraph.Runs[0].ID, otherParagraph.Runs[0].Anchor.XMLSHA256, "Changed")
		assertNativeWritebackExactTextSplice(t, parts, source, result.Package, "Other", "Changed")
		produced = result.Package
	}
	gotMain := string(readNativeZipPart(t, produced, "word/document.xml"))
	if !strings.Contains(gotMain, `<w:spacing w:after="0"/>`) {
		t.Fatalf("explicit w:after=0 was dropped: %s", gotMain)
	}
}

func TestApplyNativeTextMutationsV1PreservesUntouchedOPCParts(t *testing.T) {
	body := `<w:p w14:paraId="55555555"><w:r><w:t>Before</w:t></w:r></w:p>`
	parts := nativeWritebackFidelityParts(body)
	source := buildNativeDOCX(t, nativeEntries(parts))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	run := doc.Body.Blocks[0].Paragraph.Runs[0]
	result := mustApplyNativeWritebackText(t, source, doc, "run", run.ID, run.Anchor.XMLSHA256, "After")
	assertNativeWritebackExactTextSplice(t, parts, source, result.Package, "Before", "After")
	for _, name := range []string{"word/styles.xml", "_rels/.rels", "word/opaque.xml", "[Content_Types].xml"} {
		assertNativeRawPartPreserved(t, source, result.Package, name)
	}
	if len(result.Evidence.ChangedParts) != 1 || result.Evidence.ChangedParts[0].PartName != "word/document.xml" {
		t.Fatalf("changed-part evidence = %#v", result.Evidence.ChangedParts)
	}
}

func TestApplyNativeTextMutationsV1WritebackIsByteDeterministic(t *testing.T) {
	body := `<w:p w14:paraId="66666666"><w:r><w:t>Once</w:t></w:r></w:p>`
	parts := nativeWritebackFidelityParts(body)
	source := buildNativeDOCX(t, nativeEntries(parts))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	run := doc.Body.Blocks[0].Paragraph.Runs[0]
	mutation := []NativeDOCXTextMutationV1{{TargetKind: "run", TargetID: run.ID, ExpectedXMLSHA256: run.Anchor.XMLSHA256, Text: "Twice"}}
	first, err := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, mutation)
	if err != nil {
		t.Fatal(err)
	}
	assertNativeWritebackDeterministic(t, source, doc.Source.PackageSHA256, mutation, first.Package)
	assertNativeWritebackExactTextSplice(t, parts, source, first.Package, "Once", "Twice")
}

func TestApplyNativeTextMutationsV1PreservesModeledParagraphPropertiesOnSuccessfulSplice(t *testing.T) {
	body := `<w:p w14:paraId="77777777">` + nativeWritebackModeledPPr + `<w:r>` + nativeWritebackRunRPr + `<w:t>Before</w:t></w:r></w:p>`
	parts := nativeWritebackFidelityParts(body)
	source := buildNativeDOCX(t, nativeEntries(parts))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	paragraph := doc.Body.Blocks[0].Paragraph
	if !nativeWritebackParagraphMutable(paragraph) {
		t.Fatalf("modeled pPr + run rPr should remain writable: %#v", paragraph.EditPolicy)
	}
	if _, exists := indexNativeTextTargets(doc)["paragraph\x00"+paragraph.ID]; !exists {
		t.Fatal("single-run paragraph was not addressable without flattening")
	}
	result := mustApplyNativeWritebackText(t, source, doc, "paragraph", paragraph.ID, paragraph.Anchor.XMLSHA256, `After & <exact>`)
	assertNativeWritebackExactTextSplice(t, parts, source, result.Package, "Before", `After & <exact>`)
	gotMain := string(readNativeZipPart(t, result.Package, "word/document.xml"))
	if !strings.Contains(gotMain, nativeWritebackModeledPPr) {
		t.Fatalf("modeled pPr rewritten: %s", gotMain)
	}
	if !strings.Contains(gotMain, nativeWritebackRunRPr) {
		t.Fatalf("run rPr rewritten: %s", gotMain)
	}
}

func nativeWritebackParagraphMutable(paragraph *NativeParagraphV1) bool {
	return paragraph != nil && paragraph.EditPolicy.Mode == "read-write" && nativePolicyAllows(paragraph.EditPolicy, "text.replace")
}

func mustApplyNativeWritebackText(t *testing.T, source []byte, doc *NativeDocumentV1, kind, id, sha, text string) *NativeDOCXMutationResultV1 {
	t.Helper()
	result, err := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXTextMutationV1{{
		TargetKind: kind, TargetID: id, ExpectedXMLSHA256: sha, Text: text,
	}})
	if err != nil {
		t.Fatal(err)
	}
	return result
}

func assertNativeWritebackRefusesConstruct(t *testing.T, source []byte, doc *NativeDocumentV1, paragraph *NativeParagraphV1) {
	t.Helper()
	run := paragraph.Runs[0]
	for _, kind := range []string{"run", "paragraph"} {
		id, sha := run.ID, run.Anchor.XMLSHA256
		if kind == "paragraph" {
			id, sha = paragraph.ID, paragraph.Anchor.XMLSHA256
		}
		result, err := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXTextMutationV1{{
			TargetKind: kind, TargetID: id, ExpectedXMLSHA256: sha, Text: "Changed",
		}})
		if result != nil {
			t.Fatalf("%s mutation of preservation-only paragraph returned a candidate: %#v", kind, result)
		}
		assertNativeMutationCode(t, err, "UNSUPPORTED_CONSTRUCT")
	}
}

func assertNativeWritebackExactTextSplice(t *testing.T, parts map[string]string, source, produced []byte, oldText, newText string) {
	t.Helper()
	wantMain := strings.Replace(parts["word/document.xml"], ">"+oldText+"</w:t>", ">"+nativeMutationEscapeText(newText)+"</w:t>", 1)
	if wantMain == parts["word/document.xml"] {
		t.Fatalf("fixture old text %q was not found for splice assertion", oldText)
	}
	got := string(readNativeZipPart(t, produced, "word/document.xml"))
	if got != wantMain {
		t.Fatalf("main XML was not a single exact text splice:\n got %s\nwant %s", got, wantMain)
	}
	assertNativeRawPartPreserved(t, source, produced, "word/styles.xml")
	assertNativeRawPartPreserved(t, source, produced, "_rels/.rels")
	assertNativeRawPartPreserved(t, source, produced, "word/opaque.xml")
}

func assertNativeWritebackDeterministic(t *testing.T, source []byte, revision string, mutation []NativeDOCXTextMutationV1, first []byte) {
	t.Helper()
	repeated, err := ApplyNativeTextMutationsV1(source, revision, mutation)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(first, repeated.Package) {
		t.Fatal("identical native mutation was not byte deterministic")
	}
}
