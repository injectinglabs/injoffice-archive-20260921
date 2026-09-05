package docxpatch

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"errors"
	"io"
	"os"
	"strings"
	"testing"
	"time"
)

func nativeMutationParts(mainXML string) map[string]string {
	return map[string]string{
		"[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
		"_rels/.rels":         `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="office" Type="` + relBaseTransitional + `officeDocument" Target="word/document.xml"/></Relationships>`,
		"word/document.xml":   mainXML,
		"word/styles.xml":     `<w:styles xmlns:w="` + testW + `"><w:style w:type="paragraph" w:styleId="Keep"><w:name w:val="untouched"/></w:style></w:styles>`,
	}
}

func nativeMutationMain(body string) string {
	return `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="` + testW + `" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>` + body + `</w:body></w:document>`
}

func TestApplyNativeTextMutationsV1ExactParagraphPreservation(t *testing.T) {
	mainXML := nativeMutationMain(`<w:p w14:paraId="01020304"><w:r><w:rPr><w:b/></w:rPr><w:t>Before</w:t></w:r></w:p>`)
	parts := nativeMutationParts(mainXML)
	source := buildNativeDOCX(t, nativeEntries(parts))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	paragraph := doc.Body.Blocks[0].Paragraph
	if paragraph.EditPolicy.Mode != "read-write" || !nativePolicyAllows(paragraph.EditPolicy, "text.replace") {
		t.Fatalf("safe paragraph policy = %#v", paragraph.EditPolicy)
	}
	result, err := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXTextMutationV1{{TargetKind: "paragraph", TargetID: paragraph.ID, ExpectedXMLSHA256: paragraph.Anchor.XMLSHA256, Text: `After & <exact>`}})
	if err != nil {
		t.Fatal(err)
	}
	wantMain := strings.Replace(mainXML, `>Before</w:t>`, `>After &amp; &lt;exact&gt;</w:t>`, 1)
	if got := string(readNativeZipPart(t, result.Package, "word/document.xml")); got != wantMain {
		t.Fatalf("main XML was not a single exact text splice:\n got %s\nwant %s", got, wantMain)
	}
	if got := *result.Document.Body.Blocks[0].Paragraph.Runs[0].Text; got != `After & <exact>` {
		t.Fatalf("reopened text = %q", got)
	}
	if len(result.Evidence.ChangedParts) != 1 || result.Evidence.ChangedParts[0].PartName != "word/document.xml" || result.Evidence.ChangedParts[0].BeforeSHA256 == result.Evidence.ChangedParts[0].AfterSHA256 {
		t.Fatalf("changed-part evidence = %#v", result.Evidence)
	}
	if result.Evidence.UntouchedPartsVerified != len(parts)-1 || result.Evidence.SourceRevision != doc.Source.PackageSHA256 || result.Evidence.ResultRevision != result.Document.Source.PackageSHA256 {
		t.Fatalf("preservation evidence = %#v", result.Evidence)
	}
	if result.Evidence.SourceRevision != nativeSHA(source) || result.Evidence.ResultRevision != nativeSHA(result.Package) || result.Document.Revision == result.Evidence.ResultRevision {
		t.Fatalf("outer exact-byte CAS and native opaque revision were ambiguously joined: document=%q evidence=%#v", result.Document.Revision, result.Evidence)
	}
	assertNativeRawPartPreserved(t, source, result.Package, "word/styles.xml")
	assertNativeRawPartPreserved(t, source, result.Package, "_rels/.rels")
}

func TestApplyNativeTextMutationsV1PreservesUntouchedRawMetadataAndArchiveComment(t *testing.T) {
	mainXML := nativeMutationMain(`<w:p w14:paraId="01020304"><w:r><w:t>Before</w:t></w:r></w:p>`)
	parts := nativeMutationParts(mainXML)
	parts["word/opaque.xml"] = `<opaque xmlns="urn:injoffice:test">preserve raw storage</opaque>`
	source := buildNativeMutationMetadataDOCX(t, parts)
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	run := doc.Body.Blocks[0].Paragraph.Runs[0]
	result, err := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXTextMutationV1{{TargetKind: "run", TargetID: run.ID, ExpectedXMLSHA256: run.Anchor.XMLSHA256, Text: "After"}})
	if err != nil {
		t.Fatal(err)
	}
	beforeZip, err := zip.NewReader(bytes.NewReader(source), int64(len(source)))
	if err != nil {
		t.Fatal(err)
	}
	afterZip, err := zip.NewReader(bytes.NewReader(result.Package), int64(len(result.Package)))
	if err != nil {
		t.Fatal(err)
	}
	if beforeZip.Comment == "" || beforeZip.Comment != afterZip.Comment {
		t.Fatalf("archive comment changed: before=%q after=%q", beforeZip.Comment, afterZip.Comment)
	}
	for name := range parts {
		if name != "word/document.xml" {
			assertNativeRawPartPreserved(t, source, result.Package, name)
		}
	}
}

func TestApplyNativeTextMutationsV1ExactRunAndRoundTrip(t *testing.T) {
	mainXML := nativeMutationMain(`<w:p w14:paraId="11111111"><w:r><w:rPr><w:i/></w:rPr><w:t>Alpha</w:t></w:r><w:r><w:t xml:space="preserve"> Beta </w:t></w:r></w:p>`)
	source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(mainXML)))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	paragraph := doc.Body.Blocks[0].Paragraph
	if _, exists := indexNativeTextTargets(doc)["paragraph\x00"+paragraph.ID]; exists {
		t.Fatal("multi-run paragraph must refuse lossy paragraph flattening")
	}
	run := paragraph.Runs[1]
	mutation := []NativeDOCXTextMutationV1{{TargetKind: "run", TargetID: run.ID, ExpectedXMLSHA256: run.Anchor.XMLSHA256, Text: " Gamma "}}
	first, err := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, mutation)
	if err != nil {
		t.Fatal(err)
	}
	repeated, err := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, mutation)
	if err != nil || !bytes.Equal(first.Package, repeated.Package) {
		t.Fatalf("identical native mutation was not byte deterministic: err=%v", err)
	}
	gotMain := string(readNativeZipPart(t, first.Package, "word/document.xml"))
	if !strings.Contains(gotMain, `<w:r><w:rPr><w:i/></w:rPr><w:t>Alpha</w:t></w:r><w:r><w:t xml:space="preserve"> Gamma </w:t></w:r>`) {
		t.Fatalf("run formatting or lexical wrappers changed: %s", gotMain)
	}
	nextRun := first.Document.Body.Blocks[0].Paragraph.Runs[1]
	second, err := ApplyNativeTextMutationsV1(first.Package, first.Document.Source.PackageSHA256, []NativeDOCXTextMutationV1{{TargetKind: "run", TargetID: nextRun.ID, ExpectedXMLSHA256: nextRun.Anchor.XMLSHA256, Text: "Delta"}})
	if err != nil {
		t.Fatal(err)
	}
	if got := *second.Document.Body.Blocks[0].Paragraph.Runs[1].Text; got != "Delta" {
		t.Fatalf("second native round trip text = %q", got)
	}
	if _, err := ExtractNativeDocumentV1(second.Package); err != nil {
		t.Fatalf("second output did not reopen independently: %v", err)
	}
}

func TestApplyNativeTextMutationPayloadV1PreservesComplexPackageAcrossStories(t *testing.T) {
	parts := transitionalNativeParts()
	preserveOnly := `<w:sdt><w:sdtPr><w:tag w:val="opaque-control"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Controlled</w:t></w:r></w:p></w:sdtContent></w:sdt>` +
		`<w:p><w:bookmarkStart w:id="9" w:name="KeepBookmark"/><w:r><w:t>Bookmarked</w:t></w:r><w:bookmarkEnd w:id="9"/></w:p>` +
		`<w:p><w:ins w:id="11" w:author="Ada"><w:r><w:t>Tracked insertion</w:t></w:r></w:ins></w:p>` +
		`<w:p w14:paraId="11223344"><w:r><w:t>Body mutable</w:t></w:r></w:p>`
	finalSection := `<w:sectPr><w:type w:val="nextPage"`
	parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], finalSection, preserveOnly+finalSection, 1)
	parts["Custom/Opaque.XML"] = `<opaque xmlns="urn:injoffice:test"><keep bytes="exact">relationship-independent</keep></opaque>`
	source := buildNativeDOCX(t, nativeEntries(parts))
	sourceBefore := bytes.Clone(source)
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	if len(doc.Sections) != 2 || len(doc.Headers) != 1 || len(doc.Footers) != 1 || len(doc.Comments) != 1 || len(doc.CommentStories) != 1 {
		t.Fatalf("complex package topology was not modeled: sections=%d headers=%d footers=%d comments=%d comment-stories=%d", len(doc.Sections), len(doc.Headers), len(doc.Footers), len(doc.Comments), len(doc.CommentStories))
	}
	body := nativeMutationTargetWithText(t, doc, "Body mutable")
	header := nativeMutationTargetWithText(t, doc, "Native header")
	mutations := []NativeDOCXTextMutationV1{
		{TargetKind: body.kind, TargetID: nativeMutationTargetID(t, doc, body), ExpectedXMLSHA256: body.anchor.XMLSHA256, Text: `Body changed & exact`},
		{TargetKind: header.kind, TargetID: nativeMutationTargetID(t, doc, header), ExpectedXMLSHA256: header.anchor.XMLSHA256, Text: "Header changed"},
	}
	payload, err := json.Marshal(struct {
		Mutations []NativeDOCXTextMutationV1 `json:"mutations"`
	}{Mutations: mutations})
	if err != nil {
		t.Fatal(err)
	}
	result, err := ApplyNativeTextMutationPayloadV1(source, payload, doc.Source.PackageSHA256)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(source, sourceBefore) {
		t.Fatal("successful transaction mutated caller-owned source bytes")
	}
	wantMain := strings.Replace(parts["Custom/Main.XML"], `>Body mutable</w:t>`, `>Body changed &amp; exact</w:t>`, 1)
	if got := string(readNativeZipPart(t, result.Package, "Custom/Main.XML")); got != wantMain {
		t.Fatalf("complex main story was not changed by one exact splice:\n got %s\nwant %s", got, wantMain)
	}
	wantHeader := strings.Replace(parts["Custom/Stories/HeaderA.XML"], `>Native header</w:t>`, `>Header changed</w:t>`, 1)
	if got := string(readNativeZipPart(t, result.Package, "Custom/Stories/HeaderA.XML")); got != wantHeader {
		t.Fatalf("header story was not changed by one exact splice:\n got %s\nwant %s", got, wantHeader)
	}
	if len(result.Evidence.ChangedParts) != 2 || result.Evidence.ChangedParts[0].PartName != "Custom/Main.XML" || result.Evidence.ChangedParts[1].PartName != "Custom/Stories/HeaderA.XML" {
		t.Fatalf("cross-story changed-part evidence = %#v", result.Evidence.ChangedParts)
	}
	for name := range parts {
		if name == "Custom/Main.XML" || name == "Custom/Stories/HeaderA.XML" {
			continue
		}
		assertNativeRawPartPreserved(t, source, result.Package, name)
	}
	for _, lexical := range []string{"opaque-control", "KeepBookmark", "Tracked insertion", "w:instrText", "w:commentRangeStart", "w:footerReference", "w:sectPr"} {
		if !strings.Contains(wantMain, lexical) {
			t.Fatalf("preservation fixture lost required construct %q", lexical)
		}
	}
}

func TestApplyNativeTextMutationPayloadV1RefusesCrossPartBatchAtomically(t *testing.T) {
	parts := transitionalNativeParts()
	parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `<w:sectPr><w:type w:val="nextPage"`, `<w:p w14:paraId="11223344"><w:r><w:t>Body mutable</w:t></w:r></w:p><w:sectPr><w:type w:val="nextPage"`, 1)
	source := buildNativeDOCX(t, nativeEntries(parts))
	sourceBefore := bytes.Clone(source)
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	body := nativeMutationTargetWithText(t, doc, "Body mutable")
	header := nativeMutationTargetWithText(t, doc, "Native header")
	mutations := []NativeDOCXTextMutationV1{
		{TargetKind: body.kind, TargetID: nativeMutationTargetID(t, doc, body), ExpectedXMLSHA256: body.anchor.XMLSHA256, Text: "Body changed"},
		{TargetKind: header.kind, TargetID: nativeMutationTargetID(t, doc, header), ExpectedXMLSHA256: "sha256:" + strings.Repeat("0", 64), Text: "Header changed"},
	}
	payload, err := json.Marshal(struct {
		Mutations []NativeDOCXTextMutationV1 `json:"mutations"`
	}{Mutations: mutations})
	if err != nil {
		t.Fatal(err)
	}
	result, err := ApplyNativeTextMutationPayloadV1(source, payload, doc.Source.PackageSHA256)
	if result != nil {
		t.Fatalf("stale cross-part batch returned a candidate: %#v", result)
	}
	assertNativeMutationCode(t, err, "STALE_TARGET")
	if !bytes.Equal(source, sourceBefore) {
		t.Fatal("failed cross-part batch mutated caller-owned source bytes")
	}
	result, err = ApplyNativeTextMutationPayloadV1(source, []byte(`{not-json`), "sha256:"+strings.Repeat("0", 64))
	if result != nil {
		t.Fatalf("stale outer CAS returned a candidate: %#v", result)
	}
	assertNativeMutationCode(t, err, "STALE_REVISION")
}

func TestApplyNativeTextMutationsV1RefusesSemanticNoop(t *testing.T) {
	mainXML := nativeMutationMain(`<w:p><w:r><w:t>Same</w:t></w:r></w:p>`)
	source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(mainXML)))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	run := doc.Body.Blocks[0].Paragraph.Runs[0]
	result, err := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXTextMutationV1{{TargetKind: "run", TargetID: run.ID, ExpectedXMLSHA256: run.Anchor.XMLSHA256, Text: "Same"}})
	if result != nil {
		t.Fatalf("semantic no-op returned a candidate: %#v", result)
	}
	assertNativeMutationCode(t, err, "SEMANTIC_NO_OP")
}

func TestApplyNativeTextMutationsV1RequiresFullPackageSHA256CAS(t *testing.T) {
	source, doc, run := nativeMutationFixture(t, `<w:p><w:r><w:t>Before</w:t></w:r></w:p>`)
	mutation := []NativeDOCXTextMutationV1{{TargetKind: "run", TargetID: run.ID, ExpectedXMLSHA256: run.Anchor.XMLSHA256, Text: "After"}}
	if doc.Revision == doc.Source.PackageSHA256 || doc.Source.PackageSHA256 != nativeSHA(source) {
		t.Fatalf("native opaque revision and exact package CAS are not distinct: revision=%q source=%q", doc.Revision, doc.Source.PackageSHA256)
	}
	result, err := ApplyNativeTextMutationsV1(source, doc.Revision, mutation)
	if result != nil {
		t.Fatalf("opaque native revision returned a candidate: %#v", result)
	}
	assertNativeMutationCode(t, err, "INVALID_REVISION")
	result, err = ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, mutation)
	if err != nil || result == nil || result.Evidence.SourceRevision != doc.Source.PackageSHA256 || result.Evidence.ResultRevision != nativeSHA(result.Package) {
		t.Fatalf("full exact-byte CAS did not join to native save evidence: result=%#v err=%v", result, err)
	}
}

func TestApplyNativeTextMutationsV1RefusesStaleUnsupportedAndOverlapping(t *testing.T) {
	t.Run("stale revision", func(t *testing.T) {
		source, doc, run := nativeMutationFixture(t, `<w:p><w:r><w:t>Before</w:t></w:r></w:p>`)
		_, err := ApplyNativeTextMutationsV1(source, "sha256:"+strings.Repeat("0", 64), []NativeDOCXTextMutationV1{{TargetKind: "run", TargetID: run.ID, ExpectedXMLSHA256: run.Anchor.XMLSHA256, Text: "After"}})
		assertNativeMutationCode(t, err, "STALE_REVISION")
		_ = doc
	})
	t.Run("stale anchor", func(t *testing.T) {
		source, doc, run := nativeMutationFixture(t, `<w:p><w:r><w:t>Before</w:t></w:r></w:p>`)
		_, err := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXTextMutationV1{{TargetKind: "run", TargetID: run.ID, ExpectedXMLSHA256: "sha256:" + strings.Repeat("0", 64), Text: "After"}})
		assertNativeMutationCode(t, err, "STALE_TARGET")
	})
	t.Run("unsupported hyperlink", func(t *testing.T) {
		source, doc, run := nativeMutationFixture(t, `<w:p><w:hyperlink><w:r><w:t>Linked</w:t></w:r></w:hyperlink></w:p>`)
		if doc.Body.Blocks[0].Paragraph.EditPolicy.Mode != "read-only" {
			t.Fatal("hyperlink paragraph unexpectedly writable")
		}
		_, err := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXTextMutationV1{{TargetKind: "run", TargetID: run.ID, ExpectedXMLSHA256: run.Anchor.XMLSHA256, Text: "Flattened"}})
		assertNativeMutationCode(t, err, "UNSUPPORTED_CONSTRUCT")
	})
	t.Run("overlap", func(t *testing.T) {
		source, doc, run := nativeMutationFixture(t, `<w:p><w:r><w:t>Before</w:t></w:r></w:p>`)
		paragraph := doc.Body.Blocks[0].Paragraph
		_, err := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXTextMutationV1{
			{TargetKind: "paragraph", TargetID: paragraph.ID, ExpectedXMLSHA256: paragraph.Anchor.XMLSHA256, Text: "One"},
			{TargetKind: "run", TargetID: run.ID, ExpectedXMLSHA256: run.Anchor.XMLSHA256, Text: "Two"},
		})
		assertNativeMutationCode(t, err, "OVERLAPPING_TARGETS")
	})
	t.Run("whitespace semantics", func(t *testing.T) {
		source, doc, run := nativeMutationFixture(t, `<w:p><w:r><w:t>Before</w:t></w:r></w:p>`)
		_, err := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXTextMutationV1{{TargetKind: "run", TargetID: run.ID, ExpectedXMLSHA256: run.Anchor.XMLSHA256, Text: " leading"}})
		assertNativeMutationCode(t, err, "UNSUPPORTED_LEXICAL_FORM")
	})
	t.Run("attribute-value spoof does not grant whitespace preservation", func(t *testing.T) {
		source, doc, run := nativeMutationFixture(t, `<w:p><w:r><w:t data-note="xml:space='preserve'">Before</w:t></w:r></w:p>`)
		_, err := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXTextMutationV1{{TargetKind: "run", TargetID: run.ID, ExpectedXMLSHA256: run.Anchor.XMLSHA256, Text: " leading"}})
		assertNativeMutationCode(t, err, "UNSUPPORTED_LEXICAL_FORM")
	})
}

func TestApplyNativeTextMutationsV1RefusesCommentedTextWithoutPartialBatch(t *testing.T) {
	mainXML := nativeMutationMain(`<w:p><w:r><w:t>First</w:t></w:r><w:r><w:t>Be<!--preserve-->fore</w:t></w:r></w:p>`)
	source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(mainXML)))
	sourceBefore := bytes.Clone(source)
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	runs := doc.Body.Blocks[0].Paragraph.Runs
	result, err := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXTextMutationV1{
		{TargetKind: "run", TargetID: runs[0].ID, ExpectedXMLSHA256: runs[0].Anchor.XMLSHA256, Text: "Changed"},
		{TargetKind: "run", TargetID: runs[1].ID, ExpectedXMLSHA256: runs[1].Anchor.XMLSHA256, Text: "After"},
	})
	if result != nil {
		t.Fatalf("comment-destroying batch returned a candidate: %#v", result)
	}
	assertNativeMutationCode(t, err, "UNSUPPORTED_LEXICAL_FORM")
	if !bytes.Equal(source, sourceBefore) {
		t.Fatal("failed atomic batch mutated caller-owned source bytes")
	}
	if got := string(readNativeZipPart(t, source, "word/document.xml")); got != mainXML {
		t.Fatalf("failed atomic batch changed source XML:\n got %s\nwant %s", got, mainXML)
	}
}

func TestApplyNativeTextMutationsV1TreatsCDATAAsTextNotPreserveOnlyMarkup(t *testing.T) {
	mainXML := nativeMutationMain(`<w:p><w:r><w:t><![CDATA[Before <!-- literal]]></w:t></w:r></w:p>`)
	source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(mainXML)))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	run := doc.Body.Blocks[0].Paragraph.Runs[0]
	if run.Text == nil || *run.Text != "Before <!-- literal" {
		t.Fatalf("CDATA text projection = %#v", run.Text)
	}
	result, err := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXTextMutationV1{{TargetKind: "run", TargetID: run.ID, ExpectedXMLSHA256: run.Anchor.XMLSHA256, Text: "After"}})
	if err != nil {
		t.Fatal(err)
	}
	wantMain := strings.Replace(mainXML, `<![CDATA[Before <!-- literal]]>`, `After`, 1)
	if got := string(readNativeZipPart(t, result.Package, "word/document.xml")); got != wantMain {
		t.Fatalf("CDATA-backed text was not replaced exactly:\n got %s\nwant %s", got, wantMain)
	}
}

func TestApplyNativeTextMutationsV1ExpandsSelfClosedTextExactly(t *testing.T) {
	source, doc, run := nativeMutationFixture(t, `<w:p><w:r><w:t /></w:r></w:p>`)
	result, err := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXTextMutationV1{{TargetKind: "run", TargetID: run.ID, ExpectedXMLSHA256: run.Anchor.XMLSHA256, Text: "Now"}})
	if err != nil {
		t.Fatal(err)
	}
	if got := string(readNativeZipPart(t, result.Package, "word/document.xml")); !strings.Contains(got, `<w:t >Now</w:t>`) {
		t.Fatalf("self-closed native text was not expanded safely: %s", got)
	}
	if got := *result.Document.Body.Blocks[0].Paragraph.Runs[0].Text; got != "Now" {
		t.Fatalf("reopened self-closed mutation = %q", got)
	}
}

func TestApplyNativeTextMutationsV1ResourceBounds(t *testing.T) {
	source, doc, run := nativeMutationFixture(t, `<w:p><w:r><w:t>Before</w:t></w:r></w:p>`)
	tooMany := make([]NativeDOCXTextMutationV1, NativeDOCXMaxMutations+1)
	if _, err := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, tooMany); err == nil {
		t.Fatal("mutation-count bound was not enforced")
	} else {
		assertNativeMutationCode(t, err, "INVALID_MUTATION_COUNT")
	}
	_, err := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXTextMutationV1{{TargetKind: "run", TargetID: run.ID, ExpectedXMLSHA256: run.Anchor.XMLSHA256, Text: strings.Repeat("x", NativeDOCXMaxTextLength+1)}})
	assertNativeMutationCode(t, err, "INVALID_TEXT")
	_, err = ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXTextMutationV1{{TargetKind: "run", TargetID: run.ID, ExpectedXMLSHA256: run.Anchor.XMLSHA256, Text: strings.Repeat("😀", NativeDOCXMaxTextLength/2+1)}})
	assertNativeMutationCode(t, err, "INVALID_TEXT")
	_, err = ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXTextMutationV1{{TargetKind: "run", TargetID: run.ID, ExpectedXMLSHA256: run.Anchor.XMLSHA256, Text: "bad\x00text"}})
	assertNativeMutationCode(t, err, "INVALID_TEXT")
	for name, mutation := range map[string]NativeDOCXTextMutationV1{
		"empty target id":       {TargetKind: "run", TargetID: "", ExpectedXMLSHA256: run.Anchor.XMLSHA256, Text: "After"},
		"oversized target id":   {TargetKind: "run", TargetID: "run:" + strings.Repeat("x", 300), ExpectedXMLSHA256: run.Anchor.XMLSHA256, Text: "After"},
		"unicode target id":     {TargetKind: "run", TargetID: "run:é", ExpectedXMLSHA256: run.Anchor.XMLSHA256, Text: "After"},
		"short target hash":     {TargetKind: "run", TargetID: run.ID, ExpectedXMLSHA256: "sha256:abcd", Text: "After"},
		"uppercase target hash": {TargetKind: "run", TargetID: run.ID, ExpectedXMLSHA256: "sha256:" + strings.Repeat("A", 64), Text: "After"},
	} {
		t.Run(name, func(t *testing.T) {
			result, selectorErr := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXTextMutationV1{mutation})
			if result != nil {
				t.Fatalf("invalid selector returned a candidate: %#v", result)
			}
			assertNativeMutationCode(t, selectorErr, "INVALID_SELECTOR")
		})
	}
}

func TestApplyNativeTextMutationsV1RefusesDigitallySignedPackage(t *testing.T) {
	parts := nativeMutationParts(nativeMutationMain(`<w:p><w:r><w:t>Before</w:t></w:r></w:p>`))
	parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/_xmlsignatures/origin.sigs" ContentType="application/vnd.openxmlformats-package.digital-signature-origin"/></Types>`, 1)
	parts["_rels/.rels"] = strings.Replace(parts["_rels/.rels"], `</Relationships>`, `<Relationship Id="signature-origin" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin" Target="_xmlsignatures/origin.sigs"/></Relationships>`, 1)
	parts["_xmlsignatures/origin.sigs"] = `<SignatureOrigin xmlns="urn:injoffice:test"/>`
	source := buildNativeDOCX(t, nativeEntries(parts))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	run := doc.Body.Blocks[0].Paragraph.Runs[0]
	result, err := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXTextMutationV1{{TargetKind: "run", TargetID: run.ID, ExpectedXMLSHA256: run.Anchor.XMLSHA256, Text: "After"}})
	if result != nil {
		t.Fatalf("signed package mutation returned a candidate: %#v", result)
	}
	assertNativeMutationCode(t, err, "SIGNED_PACKAGE")
}

func TestDecodeNativeDOCXTextMutationPayloadV1StrictAndBounded(t *testing.T) {
	valid := `{"mutations":[{"target_kind":"run","target_id":"run:one","expected_xml_sha256":"sha256:` + strings.Repeat("a", 64) + `","text":"After"}]}`
	mutations, err := DecodeNativeDOCXTextMutationPayloadV1([]byte(valid))
	if err != nil || len(mutations) != 1 || mutations[0].Text != "After" {
		t.Fatalf("valid payload decode = %#v, %v", mutations, err)
	}
	validPair := strings.Replace(valid, `"After"`, `"\ud83d\ude00e\u0301"`, 1)
	mutations, err = DecodeNativeDOCXTextMutationPayloadV1([]byte(validPair))
	if err != nil || len(mutations) != 1 || mutations[0].Text != "😀é" {
		t.Fatalf("valid surrogate/combining payload decode = %#v, %v", mutations, err)
	}
	for name, encoded := range map[string]string{
		"unknown outer":      `{"mutations":[],"extra":true}`,
		"duplicate outer":    `{"mutations":[],"mutations":[]}`,
		"unknown mutation":   `{"mutations":[{"target_kind":"run","target_id":"x","expected_xml_sha256":"x","text":"x","extra":"x"}]}`,
		"duplicate mutation": `{"mutations":[{"target_kind":"run","target_kind":"run","target_id":"x","expected_xml_sha256":"x","text":"x"}]}`,
		"non-string scalar":  `{"mutations":[{"target_kind":"run","target_id":1,"expected_xml_sha256":"x","text":"x"}]}`,
		"missing field":      `{"mutations":[{"target_kind":"run","target_id":"x","text":"x"}]}`,
		"high surrogate":     `{"mutations":[{"target_kind":"run","target_id":"run:one","expected_xml_sha256":"sha256:` + strings.Repeat("a", 64) + `","text":"\ud800"}]}`,
		"low surrogate":      `{"mutations":[{"target_kind":"run","target_id":"run:one","expected_xml_sha256":"sha256:` + strings.Repeat("a", 64) + `","text":"\udc00"}]}`,
		"broken pair":        `{"mutations":[{"target_kind":"run","target_id":"run:one","expected_xml_sha256":"sha256:` + strings.Repeat("a", 64) + `","text":"\ud800A"}]}`,
		"empty batch":        `{"mutations":[]}`,
		"trailing JSON":      valid + `{}`,
		"array root":         `[]`,
	} {
		t.Run(name, func(t *testing.T) {
			decoded, decodeErr := DecodeNativeDOCXTextMutationPayloadV1([]byte(encoded))
			if decodeErr == nil || decoded != nil {
				t.Fatalf("invalid payload accepted: decoded=%#v err=%v", decoded, decodeErr)
			}
			assertNativeMutationCode(t, decodeErr, "INVALID_PAYLOAD")
		})
	}
	oversized := make([]byte, NativeDOCXMaxMutationPayloadBytes+1)
	if decoded, decodeErr := DecodeNativeDOCXTextMutationPayloadV1(oversized); decodeErr == nil || decoded != nil {
		t.Fatalf("oversized payload accepted: decoded=%#v err=%v", decoded, decodeErr)
	}
	invalidUTF8 := []byte(valid)
	invalidUTF8[strings.Index(valid, "After")] = 0xff
	if decoded, decodeErr := DecodeNativeDOCXTextMutationPayloadV1(invalidUTF8); decodeErr == nil || decoded != nil {
		t.Fatalf("invalid UTF-8 payload accepted: decoded=%#v err=%v", decoded, decodeErr)
	}
}

func TestNativeDOCXMutationVerifiersRejectInventoryAndTopologyDrift(t *testing.T) {
	anchor := func(path string) *NativeSourceAnchorV1 {
		start, end := int64(1), int64(2)
		return &NativeSourceAnchorV1{PartName: "word/document.xml", Path: path, StartByte: &start, EndByte: &end, XMLSHA256: "sha256:" + strings.Repeat("a", 64)}
	}
	first := NativeUnsupportedCapabilityV1{ID: "unsupported:first", Code: "OPAQUE", Capability: "test", ScopeID: "paragraph:first", Anchor: anchor("/w:document[1]/w:body[1]/w:p[1]"), Preservation: "refuse-mutation", Message: "preserve"}
	second := NativeUnsupportedCapabilityV1{ID: "unsupported:second", Code: "OPAQUE", Capability: "test", ScopeID: "paragraph:second", Anchor: anchor("/w:document[1]/w:body[1]/w:p[2]"), Preservation: "refuse-mutation", Message: "preserve"}
	if err := verifyNativeDOCXUnsupportedInventory([]NativeUnsupportedCapabilityV1{first}, []NativeUnsupportedCapabilityV1{first, second}); err == nil {
		t.Fatal("after-only unsupported item was accepted")
	}
	remapped := first
	remapped.ScopeID = "paragraph:other"
	if err := verifyNativeDOCXUnsupportedInventory([]NativeUnsupportedCapabilityV1{first}, []NativeUnsupportedCapabilityV1{remapped}); err == nil {
		t.Fatal("remapped unsupported item was accepted")
	}
	if err := verifyNativeDOCXUnsupportedInventory([]NativeUnsupportedCapabilityV1{first, second}, []NativeUnsupportedCapabilityV1{first, first}); err == nil {
		t.Fatal("duplicate result unsupported identity was accepted")
	}
	if err := verifyNativeMutationTargetInventory(map[string]string{"run\x00part\x00/before": "Before"}, map[string]string{"run\x00part\x00/before": "Before", "run\x00part\x00/after-only": "After"}); err == nil {
		t.Fatal("after-only native text target was accepted")
	}

	_, before, _ := nativeMutationFixture(t, `<w:p><w:r><w:t>Before</w:t></w:r></w:p>`)
	raw, err := EncodeNativeDocumentV1(before)
	if err != nil {
		t.Fatal(err)
	}
	after, err := DecodeNativeDocumentV1(raw)
	if err != nil {
		t.Fatal(err)
	}
	after.Body.Blocks = append(after.Body.Blocks, after.Body.Blocks[0])
	if err := verifyNativeDOCXTopology(before, after); err == nil {
		t.Fatal("after-only document block was accepted")
	}

	width, rowHeight, gridSpan := int64(7200), int64(400), 2
	repeatHeader, cantSplit, titlePage := true, true, true
	style, layout, alignment, shading, heightRule := "Grid", "fixed", "center", "AABBCC", "atLeast"
	tableAnchor := before.Body.Blocks[0].Paragraph.Anchor
	before.Body.Blocks = append(before.Body.Blocks, NativeBlockV1{Kind: "table", ID: "block:table-proof", Table: &NativeTableV1{
		ID: "table:proof", Anchor: tableAnchor, EditPolicy: before.Body.Blocks[0].Paragraph.EditPolicy,
		TableStyleID: &style, WidthTwips: &width, Layout: &layout, Alignment: &alignment,
		GridWidthsTwips: []int64{3600, 3600}, CellMargins: &NativeTableCellMarginsV1{TopTwips: 10, RightTwips: 20, BottomTwips: 30, LeftTwips: 40},
		Borders: &NativeTableBordersV1{Top: &NativeTableBorderV1{Style: "single", SizeEighthPoints: 8}},
		Rows: []NativeTableRowV1{{ID: "row:proof", Anchor: tableAnchor, HeightTwips: &rowHeight, HeightRule: &heightRule, RepeatHeader: &repeatHeader, CantSplit: &cantSplit,
			Cells: []NativeTableCellV1{{ID: "cell:proof", Anchor: tableAnchor, WidthTwips: &width, GridSpan: &gridSpan, VerticalMerge: "restart", ShadingRGB: &shading}}}},
	}})
	before.Sections[0].TitlePage = &titlePage
	cloneDocument := func() *NativeDocumentV1 {
		t.Helper()
		encoded, marshalErr := json.Marshal(before)
		if marshalErr != nil {
			t.Fatal(marshalErr)
		}
		var cloned NativeDocumentV1
		if unmarshalErr := json.Unmarshal(encoded, &cloned); unmarshalErr != nil {
			t.Fatal(unmarshalErr)
		}
		return &cloned
	}
	for name, mutate := range map[string]func(*NativeDocumentV1){
		"table geometry": func(value *NativeDocumentV1) {
			replacement := int64(7100)
			value.Body.Blocks[1].Table.WidthTwips = &replacement
		},
		"row cant-split": func(value *NativeDocumentV1) {
			replacement := false
			value.Body.Blocks[1].Table.Rows[0].CantSplit = &replacement
		},
		"cell shading": func(value *NativeDocumentV1) {
			replacement := "FFFFFF"
			value.Body.Blocks[1].Table.Rows[0].Cells[0].ShadingRGB = &replacement
		},
		"section title": func(value *NativeDocumentV1) { replacement := false; value.Sections[0].TitlePage = &replacement },
	} {
		t.Run(name, func(t *testing.T) {
			changed := cloneDocument()
			mutate(changed)
			if err := verifyNativeDOCXTopology(before, changed); err == nil {
				t.Fatal("merged native contract drift was accepted")
			}
		})
	}
}

func TestNativeMutationStaticNoLegacyAuthority(t *testing.T) {
	source, err := os.ReadFile("native_mutate.go")
	if err != nil {
		t.Fatal(err)
	}
	lower := strings.ToLower(string(source))
	for _, forbidden := range []string{"mammoth", "luckyexcel", "document.queryselector", "archive/zip", "encoding/xml", "text/html"} {
		if strings.Contains(lower, forbidden) {
			t.Fatalf("native mutation path contains forbidden legacy/browser authority %q", forbidden)
		}
	}
}

func nativeMutationFixture(t *testing.T, paragraphXML string) ([]byte, *NativeDocumentV1, NativeRunV1) {
	t.Helper()
	source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(nativeMutationMain(paragraphXML))))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	run := doc.Body.Blocks[0].Paragraph.Runs[0]
	return source, doc, run
}

func nativeMutationTargetWithText(t *testing.T, doc *NativeDocumentV1, text string) nativeTextTarget {
	t.Helper()
	var match nativeTextTarget
	count := 0
	for _, target := range indexNativeTextTargets(doc) {
		if target.kind == "run" && target.text == text {
			match = target
			count++
		}
	}
	if count != 1 {
		t.Fatalf("run target with text %q matched %d times", text, count)
	}
	return match
}

func nativeMutationTargetID(t *testing.T, doc *NativeDocumentV1, want nativeTextTarget) string {
	t.Helper()
	for key, candidate := range indexNativeTextTargets(doc) {
		if candidate.kind == want.kind && candidate.partName == want.partName && candidate.path == want.path {
			return strings.TrimPrefix(key, want.kind+"\x00")
		}
	}
	t.Fatalf("target %s %s %s has no stable selector", want.kind, want.partName, want.path)
	return ""
}

func assertNativeMutationCode(t *testing.T, err error, want string) {
	t.Helper()
	var mutationErr *NativeDOCXMutationErrorV1
	if !errors.As(err, &mutationErr) || mutationErr.Code != want {
		t.Fatalf("error = %v, want native mutation code %q", err, want)
	}
}

func assertNativeRawPartPreserved(t *testing.T, before, after []byte, name string) {
	t.Helper()
	readRaw := func(data []byte) (*zip.File, []byte) {
		zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
		if err != nil {
			t.Fatal(err)
		}
		for _, file := range zr.File {
			if file.Name != name {
				continue
			}
			reader, err := file.OpenRaw()
			if err != nil {
				t.Fatal(err)
			}
			raw, err := io.ReadAll(reader)
			if err != nil {
				t.Fatal(err)
			}
			return file, raw
		}
		t.Fatalf("missing raw part %q", name)
		return nil, nil
	}
	beforeFile, beforeRaw := readRaw(before)
	afterFile, afterRaw := readRaw(after)
	if !sameRawDOCXZipHeader(beforeFile, afterFile) {
		t.Fatalf("untouched ZIP metadata changed for %q", name)
	}
	if !bytes.Equal(beforeRaw, afterRaw) {
		t.Fatalf("untouched compressed bytes changed for %q", name)
	}
}

func buildNativeMutationMetadataDOCX(t *testing.T, parts map[string]string) []byte {
	t.Helper()
	var out bytes.Buffer
	writer := zip.NewWriter(&out)
	if err := writer.SetComment("injoffice native mutation archive comment"); err != nil {
		t.Fatal(err)
	}
	for index, entry := range nativeEntries(parts) {
		method := uint16(zip.Deflate)
		if index%2 == 0 {
			method = zip.Store
		}
		header := &zip.FileHeader{Name: entry.name, Method: method, Comment: "preserve:" + entry.name, Extra: []byte{0xfe, 0xca, 0, 0}}
		header.SetModTime(time.Date(2026, time.August, 28, 12, index%60, 0, 0, time.UTC))
		header.SetMode(0o640)
		partWriter, err := writer.CreateHeader(header)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := partWriter.Write(entry.data); err != nil {
			t.Fatal(err)
		}
	}
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	return out.Bytes()
}
