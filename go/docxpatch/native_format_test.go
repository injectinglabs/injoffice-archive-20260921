package docxpatch

import (
	"strings"
	"testing"
)

func nativeFormatSource(t *testing.T, body string) ([]byte, *NativeDocumentV1) {
	t.Helper()
	source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(nativeMutationMain(body))))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	return source, doc
}

func nativeFormatTestBool(value bool) *bool { return &value }

func TestNativeMutableParagraphAdmitsRunPropertyPatches(t *testing.T) {
	_, doc := nativeFormatSource(t, `<w:p w14:paraId="01020304"><w:r><w:t>Quarterly report</w:t></w:r></w:p>`)
	paragraph := doc.Body.Blocks[0].Paragraph
	if !nativePolicyAllows(paragraph.EditPolicy, "text.replace") || !nativePolicyAllows(paragraph.EditPolicy, "properties.patch") {
		t.Fatalf("mutable paragraph policy = %#v", paragraph.EditPolicy)
	}
}

func TestApplyNativeFormatMutationsV1BoldsExactlyTheSelectedWord(t *testing.T) {
	mainXML := nativeMutationMain(`<w:p w14:paraId="01020304"><w:r><w:rPr><w:i/></w:rPr><w:t>Quarterly report</w:t></w:r></w:p>`)
	parts := nativeMutationParts(mainXML)
	source := buildNativeDOCX(t, nativeEntries(parts))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	run := doc.Body.Blocks[0].Paragraph.Runs[0]
	result, err := ApplyNativeFormatMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXFormatMutationV1{{
		TargetKind: "run", TargetID: run.ID, ExpectedXMLSHA256: run.Anchor.XMLSHA256,
		Properties: NativeDOCXRunPropertyPatchV1{Bold: nativeFormatTestBool(true)},
		Range:      &NativeDOCXTextRangeV1{StartUTF16: 0, EndUTF16: 9},
	}})
	if err != nil {
		t.Fatal(err)
	}
	want := strings.Replace(mainXML,
		`<w:r><w:rPr><w:i/></w:rPr><w:t>Quarterly report</w:t></w:r>`,
		`<w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>Quarterly</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> report</w:t></w:r>`, 1)
	if got := string(readNativeZipPart(t, result.Package, "word/document.xml")); got != want {
		t.Fatalf("run split was not exact:\n got %s\nwant %s", got, want)
	}
	runs := result.Document.Body.Blocks[0].Paragraph.Runs
	if len(runs) != 2 || *runs[0].Text != "Quarterly" || *runs[1].Text != " report" {
		t.Fatalf("re-extracted runs = %#v", runs)
	}
	if runs[0].Properties.Bold == nil || !*runs[0].Properties.Bold || runs[0].Properties.Italic == nil || !*runs[0].Properties.Italic {
		t.Fatalf("formatted run properties = %#v", runs[0].Properties)
	}
	if runs[1].Properties.Bold != nil {
		t.Fatalf("unselected run gained bold: %#v", runs[1].Properties)
	}
	assertNativeRawPartPreserved(t, source, result.Package, "word/styles.xml")
}

func TestApplyNativeFormatMutationsV1SplitsBothBoundariesAndPreservesOtherProperties(t *testing.T) {
	mainXML := nativeMutationMain(`<w:p w14:paraId="01020304"><w:r><w:rPr><w:rStyle w:val="Keep"/><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="MS Mincho"/><w:color w:val="112233"/></w:rPr><w:t>alpha beta gamma</w:t></w:r></w:p>`)
	source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(mainXML)))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	paragraph := doc.Body.Blocks[0].Paragraph
	size := 28
	family := "Georgia"
	underline := "single"
	result, err := ApplyNativeFormatMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXFormatMutationV1{{
		TargetKind: "paragraph", TargetID: paragraph.ID, ExpectedXMLSHA256: paragraph.Anchor.XMLSHA256,
		Properties: NativeDOCXRunPropertyPatchV1{FontSizeHalfPoints: &size, FontFamily: &family, Underline: &underline},
		Range:      &NativeDOCXTextRangeV1{StartUTF16: 6, EndUTF16: 10},
	}})
	if err != nil {
		t.Fatal(err)
	}
	got := string(readNativeZipPart(t, result.Package, "word/document.xml"))
	middle := `<w:r><w:rPr><w:rStyle w:val="Keep"/><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia" w:eastAsia="MS Mincho"/><w:color w:val="112233"/><w:sz w:val="28"/><w:u w:val="single"/></w:rPr><w:t>beta</w:t></w:r>`
	if !strings.Contains(got, middle) {
		t.Fatalf("patched run properties were not merged in schema order:\n%s", got)
	}
	if !strings.Contains(got, `<w:t xml:space="preserve">alpha </w:t>`) || !strings.Contains(got, `<w:t xml:space="preserve"> gamma</w:t>`) {
		t.Fatalf("split pieces did not preserve edge whitespace:\n%s", got)
	}
	runs := result.Document.Body.Blocks[0].Paragraph.Runs
	if len(runs) != 3 || *runs[0].Text != "alpha " || *runs[1].Text != "beta" || *runs[2].Text != " gamma" {
		t.Fatalf("re-extracted runs = %#v", runs)
	}
	if runs[1].Properties.FontFamily == nil || *runs[1].Properties.FontFamily != "Georgia" || runs[1].Properties.FontSizeHalfPoint == nil || *runs[1].Properties.FontSizeHalfPoint != 28 {
		t.Fatalf("re-extracted middle run = %#v", runs[1].Properties)
	}
	if runs[0].Properties.FontFamily == nil || *runs[0].Properties.FontFamily != "Arial" || runs[0].Properties.Underline != nil {
		t.Fatalf("unselected run was formatted: %#v", runs[0].Properties)
	}
	if runs[1].Properties.CharacterStyleID == nil || *runs[1].Properties.CharacterStyleID != "Keep" || runs[1].Properties.Color == nil || *runs[1].Properties.Color != "112233" {
		t.Fatalf("existing run properties were not preserved: %#v", runs[1].Properties)
	}
}

func TestApplyNativeFormatMutationsV1FormatsAcrossRunsAndCreatesMissingProperties(t *testing.T) {
	mainXML := nativeMutationMain(`<w:p w14:paraId="01020304"><w:r><w:t>one </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>two</w:t></w:r></w:p>`)
	source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(mainXML)))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	paragraph := doc.Body.Blocks[0].Paragraph
	color := "ff0000"
	result, err := ApplyNativeFormatMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXFormatMutationV1{{
		TargetKind: "paragraph", TargetID: paragraph.ID, ExpectedXMLSHA256: paragraph.Anchor.XMLSHA256,
		Properties: NativeDOCXRunPropertyPatchV1{Color: &color},
	}})
	if err != nil {
		t.Fatal(err)
	}
	want := strings.Replace(mainXML,
		`<w:r><w:t>one </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>two</w:t></w:r>`,
		`<w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t xml:space="preserve">one </w:t></w:r><w:r><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr><w:t>two</w:t></w:r>`, 1)
	if got := string(readNativeZipPart(t, result.Package, "word/document.xml")); got != want {
		t.Fatalf("paragraph-wide colour patch was not exact:\n got %s\nwant %s", got, want)
	}
	for _, run := range result.Document.Body.Blocks[0].Paragraph.Runs {
		if run.Properties.Color == nil || *run.Properties.Color != "FF0000" {
			t.Fatalf("run %q = %#v", run.ID, run.Properties)
		}
	}
}

func TestApplyNativeFormatMutationsV1RefusesUnsafeRequests(t *testing.T) {
	source, doc := nativeFormatSource(t, `<w:p w14:paraId="01020304"><w:r><w:t>Quarterly report</w:t></w:r></w:p>`)
	paragraph := doc.Body.Blocks[0].Paragraph
	run := paragraph.Runs[0]
	valid := NativeDOCXFormatMutationV1{TargetKind: "run", TargetID: run.ID, ExpectedXMLSHA256: run.Anchor.XMLSHA256, Properties: NativeDOCXRunPropertyPatchV1{Bold: nativeFormatTestBool(true)}}
	stale := valid
	stale.ExpectedXMLSHA256 = "sha256:" + strings.Repeat("0", 64)
	missing := valid
	missing.TargetID = "run:absent"
	overrun := valid
	overrun.Range = &NativeDOCXTextRangeV1{StartUTF16: 0, EndUTF16: 99}
	empty := valid
	empty.Properties = NativeDOCXRunPropertyPatchV1{}
	badUnderline := "squiggly"
	underline := valid
	underline.Properties = NativeDOCXRunPropertyPatchV1{Underline: &badUnderline}
	for name, testCase := range map[string]struct {
		mutation NativeDOCXFormatMutationV1
		code     string
	}{
		"stale target":      {stale, "STALE_TARGET"},
		"absent target":     {missing, "TARGET_NOT_FOUND"},
		"range past text":   {overrun, "INVALID_RANGE"},
		"no property":       {empty, "INVALID_SELECTOR"},
		"unmodeled value":   {underline, "INVALID_SELECTOR"},
		"unknown kind":      {NativeDOCXFormatMutationV1{TargetKind: "table", TargetID: run.ID, ExpectedXMLSHA256: run.Anchor.XMLSHA256, Properties: valid.Properties}, "INVALID_SELECTOR"},
		"stale revision ok": {valid, ""},
	} {
		result, err := ApplyNativeFormatMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXFormatMutationV1{testCase.mutation})
		if testCase.code == "" {
			if err != nil {
				t.Fatalf("%s: %v", name, err)
			}
			continue
		}
		mutationError, ok := err.(*NativeDOCXMutationErrorV1)
		if !ok || mutationError.Code != testCase.code || result != nil {
			t.Fatalf("%s: err = %v, result = %v", name, err, result != nil)
		}
	}
	if _, err := ApplyNativeFormatMutationsV1(source, "sha256:"+strings.Repeat("a", 64), []NativeDOCXFormatMutationV1{valid}); err == nil {
		t.Fatal("a stale package revision was accepted")
	}
}

func TestApplyNativeFormatMutationsV1RefusesRunsItCannotSplitExactly(t *testing.T) {
	source, doc := nativeFormatSource(t, `<w:p w14:paraId="01020304"><w:r><w:t>tabbed</w:t><w:tab/></w:r></w:p>`)
	paragraph := doc.Body.Blocks[0].Paragraph
	if paragraph.EditPolicy.Mode != "read-write" {
		t.Fatalf("a text run beside a tab control should stay mutable: %#v", paragraph.EditPolicy)
	}
	_, err := ApplyNativeFormatMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXFormatMutationV1{{
		TargetKind: "paragraph", TargetID: paragraph.ID, ExpectedXMLSHA256: paragraph.Anchor.XMLSHA256,
		Properties: NativeDOCXRunPropertyPatchV1{Bold: nativeFormatTestBool(true)},
	}})
	mutationError, ok := err.(*NativeDOCXMutationErrorV1)
	if !ok || mutationError.Code != "UNSUPPORTED_CONSTRUCT" {
		t.Fatalf("multi-content run formatting error = %v", err)
	}
}

func TestApplyNativeFormatMutationsV1RefusesAReadOnlyParagraph(t *testing.T) {
	source, doc := nativeFormatSource(t, `<w:p w14:paraId="01020304"><w:ins w:id="1" w:author="a" w:date="2024-01-01T00:00:00Z"><w:r><w:t>tracked</w:t></w:r></w:ins></w:p>`)
	paragraph := doc.Body.Blocks[0].Paragraph
	if paragraph.EditPolicy.Mode != "read-only" {
		t.Fatalf("revision-tracked paragraph policy = %#v", paragraph.EditPolicy)
	}
	_, err := ApplyNativeFormatMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXFormatMutationV1{{
		TargetKind: "paragraph", TargetID: paragraph.ID, ExpectedXMLSHA256: paragraph.Anchor.XMLSHA256,
		Properties: NativeDOCXRunPropertyPatchV1{Bold: nativeFormatTestBool(true)},
	}})
	mutationError, ok := err.(*NativeDOCXMutationErrorV1)
	if !ok || mutationError.Code != "UNSUPPORTED_CONSTRUCT" {
		t.Fatalf("read-only paragraph error = %v", err)
	}
}

func TestApplyNativeMutationPayloadV1RoutesBothPayloadShapes(t *testing.T) {
	source, doc := nativeFormatSource(t, `<w:p w14:paraId="01020304"><w:r><w:t>Quarterly report</w:t></w:r></w:p>`)
	run := doc.Body.Blocks[0].Paragraph.Runs[0]
	formatted, err := ApplyNativeMutationPayloadV1(source, []byte(`{"mutations":[{"target_kind":"run","target_id":"`+run.ID+`","expected_xml_sha256":"`+run.Anchor.XMLSHA256+`","properties":{"bold":true},"range":{"start_utf16":0,"end_utf16":9}}]}`), doc.Source.PackageSHA256)
	if err != nil {
		t.Fatal(err)
	}
	if got := string(readNativeZipPart(t, formatted.Package, "word/document.xml")); !strings.Contains(got, `<w:rPr><w:b/></w:rPr><w:t>Quarterly</w:t>`) {
		t.Fatalf("formatting payload was not applied:\n%s", got)
	}
	replaced, err := ApplyNativeMutationPayloadV1(source, []byte(`{"mutations":[{"target_kind":"run","target_id":"`+run.ID+`","expected_xml_sha256":"`+run.Anchor.XMLSHA256+`","text":"Annual report"}]}`), doc.Source.PackageSHA256)
	if err != nil {
		t.Fatal(err)
	}
	if got := *replaced.Document.Body.Blocks[0].Paragraph.Runs[0].Text; got != "Annual report" {
		t.Fatalf("text payload = %q", got)
	}
	mixed := `{"mutations":[{"target_kind":"run","target_id":"` + run.ID + `","expected_xml_sha256":"` + run.Anchor.XMLSHA256 + `","text":"x"},{"target_kind":"run","target_id":"` + run.ID + `","expected_xml_sha256":"` + run.Anchor.XMLSHA256 + `","properties":{"bold":true}}]}`
	if _, err := ApplyNativeMutationPayloadV1(source, []byte(mixed), doc.Source.PackageSHA256); err == nil {
		t.Fatal("a mixed text and formatting transaction was accepted")
	}
	if _, err := ApplyNativeMutationPayloadV1(source, []byte(`{"mutations":[{"target_kind":"run","target_id":"`+run.ID+`","expected_xml_sha256":"`+run.Anchor.XMLSHA256+`","properties":{"bold":true},"weight":700}]}`), doc.Source.PackageSHA256); err == nil {
		t.Fatal("an unknown mutation field was accepted")
	}
}

func TestApplyNativeFormatMutationsV1AlignsAParagraph(t *testing.T) {
	for name, testCase := range map[string]struct{ body, want string }{
		"no paragraph properties": {
			`<w:p w14:paraId="01020304"><w:r><w:t>Quarterly report</w:t></w:r></w:p>`,
			`<w:p w14:paraId="01020304"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Quarterly report</w:t></w:r></w:p>`,
		},
		"existing properties keep their order": {
			`<w:p w14:paraId="01020304"><w:pPr><w:pStyle w:val="Keep"/><w:keepNext/></w:pPr><w:r><w:t>Quarterly report</w:t></w:r></w:p>`,
			`<w:p w14:paraId="01020304"><w:pPr><w:pStyle w:val="Keep"/><w:keepNext/><w:jc w:val="center"/></w:pPr><w:r><w:t>Quarterly report</w:t></w:r></w:p>`,
		},
		"existing alignment is replaced in place": {
			`<w:p w14:paraId="01020304"><w:pPr><w:jc w:val="right"/><w:keepLines/></w:pPr><w:r><w:t>Quarterly report</w:t></w:r></w:p>`,
			`<w:p w14:paraId="01020304"><w:pPr><w:jc w:val="center"/><w:keepLines/></w:pPr><w:r><w:t>Quarterly report</w:t></w:r></w:p>`,
		},
	} {
		t.Run(name, func(t *testing.T) {
			source, doc := nativeFormatSource(t, testCase.body)
			paragraph := doc.Body.Blocks[0].Paragraph
			center := "center"
			result, err := ApplyNativeFormatMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXFormatMutationV1{{
				TargetKind: "paragraph", TargetID: paragraph.ID, ExpectedXMLSHA256: paragraph.Anchor.XMLSHA256,
				ParagraphProperties: &NativeDOCXParagraphPropertyPatchV1{Alignment: &center},
			}})
			if err != nil {
				t.Fatal(err)
			}
			if got := string(readNativeZipPart(t, result.Package, "word/document.xml")); got != nativeMutationMain(testCase.want) {
				t.Fatalf("alignment patch was not exact:\n got %s\nwant %s", got, nativeMutationMain(testCase.want))
			}
			if after := result.Document.Body.Blocks[0].Paragraph.Properties; after.Alignment == nil || *after.Alignment != "center" {
				t.Fatalf("re-extracted paragraph properties = %#v", after)
			}
		})
	}
}

func TestApplyNativeMutationPayloadV1AlignsAndRefusesMisplacedAlignment(t *testing.T) {
	source, doc := nativeFormatSource(t, `<w:p w14:paraId="01020304"><w:r><w:t>Quarterly report</w:t></w:r></w:p>`)
	paragraph := doc.Body.Blocks[0].Paragraph
	run := paragraph.Runs[0]
	aligned, err := ApplyNativeMutationPayloadV1(source, []byte(`{"mutations":[{"target_kind":"paragraph","target_id":"`+paragraph.ID+`","expected_xml_sha256":"`+paragraph.Anchor.XMLSHA256+`","properties":{"alignment":"both"}}]}`), doc.Source.PackageSHA256)
	if err != nil {
		t.Fatal(err)
	}
	if got := aligned.Document.Body.Blocks[0].Paragraph.Properties.Alignment; got == nil || *got != "both" {
		t.Fatalf("aligned paragraph = %#v", aligned.Document.Body.Blocks[0].Paragraph.Properties)
	}
	for name, payload := range map[string]string{
		"run target":          `{"mutations":[{"target_kind":"run","target_id":"` + run.ID + `","expected_xml_sha256":"` + run.Anchor.XMLSHA256 + `","properties":{"alignment":"center"}}]}`,
		"mixed with run bold": `{"mutations":[{"target_kind":"paragraph","target_id":"` + paragraph.ID + `","expected_xml_sha256":"` + paragraph.Anchor.XMLSHA256 + `","properties":{"alignment":"center","bold":true}}]}`,
		"unmodeled value":     `{"mutations":[{"target_kind":"paragraph","target_id":"` + paragraph.ID + `","expected_xml_sha256":"` + paragraph.Anchor.XMLSHA256 + `","properties":{"alignment":"middle"}}]}`,
		"with a range":        `{"mutations":[{"target_kind":"paragraph","target_id":"` + paragraph.ID + `","expected_xml_sha256":"` + paragraph.Anchor.XMLSHA256 + `","properties":{"alignment":"center"},"range":{"start_utf16":0,"end_utf16":4}}]}`,
	} {
		if _, err := ApplyNativeMutationPayloadV1(source, []byte(payload), doc.Source.PackageSHA256); err == nil {
			t.Fatalf("%s: alignment request was accepted", name)
		}
	}
}

// The schema position matters for elements Word writes after w:jc, which this
// tier preserves but does not model, so exercise the writer on its own.
func TestNativeFormatParagraphPropertiesInsertsAtItsSchemaPosition(t *testing.T) {
	part := []byte(`<w:p xmlns:w="` + testW + `"><w:pPr><w:pStyle w:val="Keep"/><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>text</w:t></w:r></w:p>`)
	root, err := parseNativeXML("word/document.xml", part)
	if err != nil {
		t.Fatal(err)
	}
	splice, err := nativeFormatParagraphProperties(part, root, "right")
	if err != nil {
		t.Fatal(err)
	}
	want := `<w:pPr><w:pStyle w:val="Keep"/><w:jc w:val="right"/><w:outlineLvl w:val="0"/></w:pPr>`
	if got := string(splice.text); got != want {
		t.Fatalf("alignment insertion =\n got %s\nwant %s", got, want)
	}
}
