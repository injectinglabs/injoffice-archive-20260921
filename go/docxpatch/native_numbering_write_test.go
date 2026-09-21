package docxpatch

import (
	"strings"
	"testing"
)

func nativeNumberingTestParts(body, numbering string) map[string]string {
	parts := nativeMutationParts(nativeMutationMain(body))
	if numbering == "" {
		return parts
	}
	parts["word/numbering.xml"] = numbering
	parts["word/_rels/document.xml.rels"] = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rNum" Type="` + relBaseTransitional + `numbering" Target="numbering.xml"/></Relationships>`
	parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`,
		`<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>`, 1)
	return parts
}

func nativeSimpleNumberingXML(format string) string {
	text := "•"
	if format != "bullet" {
		text = "%1."
	}
	return `<w:numbering xmlns:w="` + wordMLTransitional + `"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>` +
		`<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="` + format + `"/><w:suff w:val="space"/><w:lvlText w:val="` + text + `"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>` +
		`<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="` + format + `"/><w:suff w:val="space"/><w:lvlText w:val="` + text + `"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>` +
		`</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`
}

func TestExtractNativeDocumentV1AdvertisesListNumbering(t *testing.T) {
	source := buildNativeDOCX(t, nativeEntries(nativeNumberingTestParts(
		`<w:p w14:paraId="01020304"><w:r><w:t>Item</w:t></w:r></w:p>`,
		nativeSimpleNumberingXML("bullet"),
	)))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	if len(doc.NumberingDefinitions) != 1 || doc.NumberingDefinitions[0].NumID != "1" || doc.NumberingDefinitions[0].Levels[0].Format != "bullet" {
		t.Fatalf("numbering catalog = %#v", doc.NumberingDefinitions)
	}
	if !nativePolicyAllows(doc.Body.Blocks[0].Paragraph.EditPolicy, "properties.patch") {
		t.Fatalf("mutable paragraph policy = %#v", doc.Body.Blocks[0].Paragraph.EditPolicy)
	}
}

func TestApplyNativeNumberingMutationsV1CreatesABulletListAndReExtractsIt(t *testing.T) {
	body := `<w:p w14:paraId="01020304"><w:r><w:t>Quarterly report</w:t></w:r></w:p>`
	source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(nativeMutationMain(body))))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	paragraph := doc.Body.Blocks[0].Paragraph
	kind := "bullet"
	result, err := ApplyNativeNumberingMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXFormatMutationV1{{
		TargetKind: "paragraph", TargetID: paragraph.ID, ExpectedXMLSHA256: paragraph.Anchor.XMLSHA256,
		ParagraphProperties: &NativeDOCXParagraphPropertyPatchV1{Kind: &kind},
	}})
	if err != nil {
		t.Fatal(err)
	}
	got := string(readNativeZipPart(t, result.Package, "word/document.xml"))
	if !strings.Contains(got, `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>`) {
		t.Fatalf("created numbering was not attached in schema order:\n%s", got)
	}
	numbering := string(readNativeZipPart(t, result.Package, "word/numbering.xml"))
	if !strings.Contains(numbering, `w:numFmt w:val="bullet"`) || !strings.Contains(numbering, `w:numId="1"`) {
		t.Fatalf("created numbering part = %s", numbering)
	}
	after := result.Document.Body.Blocks[0].Paragraph.Properties.Numbering
	if after == nil || after.NumID != "1" || nativeIntValue(after.Level) != 0 {
		t.Fatalf("re-extracted numbering = %#v", after)
	}
	if len(result.Document.NumberingDefinitions) == 0 {
		t.Fatal("created numbering was not catalogued")
	}
	assertNativeRawPartPreserved(t, source, result.Package, "word/styles.xml")
}

func TestApplyNativeNumberingMutationsV1ReusesAnExistingBulletAndRemovesIt(t *testing.T) {
	body := `<w:p w14:paraId="01020304"><w:r><w:t>Item</w:t></w:r></w:p>`
	source := buildNativeDOCX(t, nativeEntries(nativeNumberingTestParts(body, nativeSimpleNumberingXML("bullet"))))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	paragraph := doc.Body.Blocks[0].Paragraph
	attached, err := ApplyNativeMutationPayloadV1(source, []byte(`{"mutations":[{"target_kind":"paragraph","target_id":"`+paragraph.ID+`","expected_xml_sha256":"`+paragraph.Anchor.XMLSHA256+`","properties":{"numbering_kind":"bullet"}}]}`), doc.Source.PackageSHA256)
	if err != nil {
		t.Fatal(err)
	}
	got := string(readNativeZipPart(t, attached.Package, "word/document.xml"))
	want := strings.Replace(nativeMutationMain(body), `<w:p w14:paraId="01020304">`, `<w:p w14:paraId="01020304"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>`, 1)
	if got != want {
		t.Fatalf("reuse was not an exact pPr splice:\n got %s\nwant %s", got, want)
	}
	assertNativeRawPartPreserved(t, source, attached.Package, "word/numbering.xml")
	paragraph = attached.Document.Body.Blocks[0].Paragraph
	removed, err := ApplyNativeMutationPayloadV1(attached.Package, []byte(`{"mutations":[{"target_kind":"paragraph","target_id":"`+paragraph.ID+`","expected_xml_sha256":"`+paragraph.Anchor.XMLSHA256+`","properties":{"numbering_num_id":null,"numbering_level":null}}]}`), attached.Document.Source.PackageSHA256)
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(string(readNativeZipPart(t, removed.Package, "word/document.xml")), "numPr") {
		t.Fatalf("numbering was not removed:\n%s", readNativeZipPart(t, removed.Package, "word/document.xml"))
	}
	if removed.Document.Body.Blocks[0].Paragraph.Properties != nil && removed.Document.Body.Blocks[0].Paragraph.Properties.Numbering != nil {
		t.Fatalf("re-extracted numbering after remove = %#v", removed.Document.Body.Blocks[0].Paragraph.Properties.Numbering)
	}
}

func TestApplyNativeNumberingMutationsV1IndentsByChangingIlvl(t *testing.T) {
	body := `<w:p w14:paraId="01020304"><w:pPr><w:pStyle w:val="Keep"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>`
	source := buildNativeDOCX(t, nativeEntries(nativeNumberingTestParts(body, nativeSimpleNumberingXML("decimal"))))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	paragraph := doc.Body.Blocks[0].Paragraph
	level := 1
	numID := "1"
	result, err := ApplyNativeNumberingMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXFormatMutationV1{{
		TargetKind: "paragraph", TargetID: paragraph.ID, ExpectedXMLSHA256: paragraph.Anchor.XMLSHA256,
		ParagraphProperties: &NativeDOCXParagraphPropertyPatchV1{NumID: &numID, Level: &level},
	}})
	if err != nil {
		t.Fatal(err)
	}
	got := string(readNativeZipPart(t, result.Package, "word/document.xml"))
	if !strings.Contains(got, `<w:pStyle w:val="Keep"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr>`) {
		t.Fatalf("indent did not keep sibling pPr children or schema position:\n%s", got)
	}
	if nativeIntValue(result.Document.Body.Blocks[0].Paragraph.Properties.Numbering.Level) != 1 {
		t.Fatalf("re-extracted level = %#v", result.Document.Body.Blocks[0].Paragraph.Properties.Numbering)
	}
}

func TestApplyNativeNumberingMutationsV1AppendsAMissingKindWithoutRewritingExistingDefinitions(t *testing.T) {
	body := `<w:p w14:paraId="01020304"><w:r><w:t>Item</w:t></w:r></w:p>`
	existing := nativeSimpleNumberingXML("decimal")
	source := buildNativeDOCX(t, nativeEntries(nativeNumberingTestParts(body, existing)))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	paragraph := doc.Body.Blocks[0].Paragraph
	kind := "bullet"
	result, err := ApplyNativeNumberingMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXFormatMutationV1{{
		TargetKind: "paragraph", TargetID: paragraph.ID, ExpectedXMLSHA256: paragraph.Anchor.XMLSHA256,
		ParagraphProperties: &NativeDOCXParagraphPropertyPatchV1{Kind: &kind},
	}})
	if err != nil {
		t.Fatal(err)
	}
	numbering := string(readNativeZipPart(t, result.Package, "word/numbering.xml"))
	if !strings.Contains(numbering, existing[len(`<w:numbering xmlns:w="`+wordMLTransitional+`">`):len(existing)-len(`</w:numbering>`)]) {
		t.Fatalf("existing numbering children were rewritten:\n%s", numbering)
	}
	if !strings.Contains(numbering, `w:numFmt w:val="bullet"`) {
		t.Fatalf("bullet abstract was not appended:\n%s", numbering)
	}
	if result.Document.Body.Blocks[0].Paragraph.Properties.Numbering == nil {
		t.Fatal("appended bullet was not attached")
	}
}

func TestApplyNativeNumberingMutationsV1RefusesUnsafeRequests(t *testing.T) {
	body := `<w:p w14:paraId="01020304"><w:r><w:t>Item</w:t></w:r></w:p>`
	source := buildNativeDOCX(t, nativeEntries(nativeNumberingTestParts(body, nativeSimpleNumberingXML("bullet"))))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	paragraph := doc.Body.Blocks[0].Paragraph
	run := paragraph.Runs[0]
	for name, payload := range map[string]string{
		"run target":         `{"mutations":[{"target_kind":"run","target_id":"` + run.ID + `","expected_xml_sha256":"` + run.Anchor.XMLSHA256 + `","properties":{"numbering_kind":"bullet"}}]}`,
		"mixed with bold":    `{"mutations":[{"target_kind":"paragraph","target_id":"` + paragraph.ID + `","expected_xml_sha256":"` + paragraph.Anchor.XMLSHA256 + `","properties":{"numbering_kind":"bullet","bold":true}}]}`,
		"unknown kind":       `{"mutations":[{"target_kind":"paragraph","target_id":"` + paragraph.ID + `","expected_xml_sha256":"` + paragraph.Anchor.XMLSHA256 + `","properties":{"numbering_kind":"outline"}}]}`,
		"level out of range": `{"mutations":[{"target_kind":"paragraph","target_id":"` + paragraph.ID + `","expected_xml_sha256":"` + paragraph.Anchor.XMLSHA256 + `","properties":{"numbering_num_id":"1","numbering_level":9}}]}`,
		"with a range":       `{"mutations":[{"target_kind":"paragraph","target_id":"` + paragraph.ID + `","expected_xml_sha256":"` + paragraph.Anchor.XMLSHA256 + `","properties":{"numbering_kind":"bullet"},"range":{"start_utf16":0,"end_utf16":4}}]}`,
		"stale fingerprint":  `{"mutations":[{"target_kind":"paragraph","target_id":"` + paragraph.ID + `","expected_xml_sha256":"sha256:` + strings.Repeat("0", 64) + `","properties":{"numbering_kind":"bullet"}}]}`,
		"missing instance":   `{"mutations":[{"target_kind":"paragraph","target_id":"` + paragraph.ID + `","expected_xml_sha256":"` + paragraph.Anchor.XMLSHA256 + `","properties":{"numbering_num_id":"99","numbering_level":0}}]}`,
	} {
		if _, err := ApplyNativeMutationPayloadV1(source, []byte(payload), doc.Source.PackageSHA256); err == nil {
			t.Fatalf("%s: numbering request was accepted", name)
		}
	}
}

func TestNativeNumberingParagraphPropertiesInsertsAtItsSchemaPosition(t *testing.T) {
	part := []byte(`<w:p xmlns:w="` + testW + `"><w:pPr><w:pStyle w:val="Keep"/><w:jc w:val="left"/></w:pPr><w:r><w:t>text</w:t></w:r></w:p>`)
	root, err := parseNativeXML("word/document.xml", part)
	if err != nil {
		t.Fatal(err)
	}
	paragraph := &NativeParagraphV1{Anchor: NativeSourceAnchorV1{Path: root.Path, StartByte: nativeInt64(root.Start), EndByte: nativeInt64(root.End), XMLSHA256: nativeSHA(part[root.Start:root.End])}}
	splice, err := nativeNumberingParagraphSplice(part, root, paragraph, false, "1", 0)
	if err != nil {
		t.Fatal(err)
	}
	got := string(append(append([]byte(nil), part[:splice.start]...), append(splice.text, part[splice.end:]...)...))
	if !strings.Contains(got, `<w:pPr><w:pStyle w:val="Keep"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:jc w:val="left"/></w:pPr>`) {
		t.Fatalf("numPr was not inserted at its ECMA-376 position:\n%s", got)
	}
}
