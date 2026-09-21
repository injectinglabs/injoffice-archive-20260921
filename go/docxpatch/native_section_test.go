package docxpatch

import (
	"fmt"
	"strings"
	"testing"
)

const nativePageTestPatch = `{"width_twips":16838,"height_twips":11906,"orientation":"landscape","margin_top_twips":720,"margin_right_twips":1080,"margin_bottom_twips":720,"margin_left_twips":1080}`
const nativePageTestSize = `<w:pgSz w:w="12240" w:h="15840"/>`
const nativePageTestMargins = `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header='360' w:footer="480" w:gutter="120"/>`

func nativePageTestPayload(section NativeSectionV1, page string) []byte {
	return []byte(fmt.Sprintf(`{"mutations":[{"target_kind":"section","target_id":%q,"expected_xml_sha256":%q,"operation":"section.page.patch","page":%s}]}`, section.ID, section.Anchor.XMLSHA256, page))
}

func TestNativeSectionPageRoundTrip(t *testing.T) {
	for _, prefix := range []string{"w", "word"} {
		t.Run(prefix, func(t *testing.T) {
			body := `<w:p><w:r><w:t>Keep body bytes</w:t></w:r></w:p>`
			main := nativeMutationMain(body + `<w:sectPr w:rsidR="01020304"><!--section comment-->` + nativePageTestSize + nativePageTestMargins + `<w:pgNumType w:start="2"/><w:cols w:num="1" w:space="720"/></w:sectPr>`)
			main = strings.ReplaceAll(strings.ReplaceAll(main, "w:", prefix+":"), "xmlns:w=", "xmlns:"+prefix+"=")
			source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(main)))
			doc, err := ExtractNativeDocumentV1(source)
			if err != nil {
				t.Fatal(err)
			}
			if !nativePolicyAllows(*doc.Sections[0].EditPolicy, "section.page.patch") {
				t.Fatalf("section policy: %#v", doc.Sections[0].EditPolicy)
			}
			result, err := ApplyNativeMutationPayloadV1(source, nativePageTestPayload(doc.Sections[0], nativePageTestPatch), doc.Source.PackageSHA256)
			if err != nil {
				t.Fatal(err)
			}
			got := string(readNativeZipPart(t, result.Package, "word/document.xml"))
			for _, keep := range []string{strings.ReplaceAll(body, "w:", prefix+":"), `<!--section comment-->`, prefix + `:rsidR="01020304"`, prefix + `:header='360'`, prefix + `:footer="480"`, prefix + `:gutter="120"`, `<` + prefix + `:pgNumType ` + prefix + `:start="2"/>`} {
				if !strings.Contains(got, keep) {
					t.Fatalf("lost %s: %s", keep, got)
				}
			}
			assertNativeRawPartPreserved(t, source, result.Package, "word/styles.xml")
			page := result.Document.Sections[0].Page
			if *page.WidthTwips != 16838 || *page.HeightTwips != 11906 || page.Orientation != "landscape" || *page.Margins.LeftTwips != 1080 {
				t.Fatalf("wrong geometry: %#v", page)
			}
			// Reuse of either stale package or stale target yields no candidate bytes.
			if out, err := ApplyNativeMutationPayloadV1(result.Package, nativePageTestPayload(doc.Sections[0], nativePageTestPatch), doc.Source.PackageSHA256); err == nil || out != nil || !strings.Contains(err.Error(), "STALE_REVISION") {
				t.Fatalf("stale revision: %v", err)
			}
			if out, err := ApplyNativeMutationPayloadV1(result.Package, nativePageTestPayload(doc.Sections[0], nativePageTestPatch), result.Document.Source.PackageSHA256); err == nil || out != nil || !strings.Contains(err.Error(), "STALE_TARGET") {
				t.Fatalf("stale target: %v", err)
			}
		})
	}
}

func TestNativeSectionPageCreatesPropertiesInSchemaOrder(t *testing.T) {
	for _, section := range []string{`<w:sectPr/>`, `<w:sectPr><w:pgNumType w:start="3"/></w:sectPr>`} {
		source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(nativeMutationMain(`<w:p><w:r><w:t>Body</w:t></w:r></w:p>`+section))))
		doc, err := ExtractNativeDocumentV1(source)
		if err != nil {
			t.Fatal(err)
		}
		result, err := ApplyNativeMutationPayloadV1(source, nativePageTestPayload(doc.Sections[0], nativePageTestPatch), doc.Source.PackageSHA256)
		if err != nil {
			t.Fatal(err)
		}
		xml := string(readNativeZipPart(t, result.Package, "word/document.xml"))
		size, margins, number := strings.Index(xml, "<w:pgSz"), strings.Index(xml, "<w:pgMar"), strings.Index(xml, "<w:pgNumType")
		if size < 0 || margins < size || (number >= 0 && number < margins) {
			t.Fatalf("wrong section order: %s", xml)
		}
		if *result.Document.Sections[0].Page.Margins.HeaderTwips != 720 {
			t.Fatal("default header distance changed")
		}
	}
}

func TestNativeSectionPageRefusals(t *testing.T) {
	for _, section := range []string{"", `<w:sectPr>` + nativePageTestSize + nativePageTestSize + nativePageTestMargins + `</w:sectPr>`, `<w:sectPr>` + nativePageTestSize + nativePageTestMargins + `<w:cols w:num="2"/></w:sectPr>`, `<w:sectPr>` + nativePageTestSize + nativePageTestMargins + `<w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>`} {
		source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(nativeMutationMain(`<w:p><w:r><w:t>Body</w:t></w:r></w:p>`+section))))
		doc, err := ExtractNativeDocumentV1(source)
		if err != nil {
			t.Fatal(err)
		}
		if doc.Sections[0].EditPolicy.Mode != "read-only" || doc.Sections[0].EditPolicy.Refusal == nil {
			t.Fatalf("unsafe section advertised writes: %s", section)
		}
		if out, err := ApplyNativeMutationPayloadV1(source, nativePageTestPayload(doc.Sections[0], nativePageTestPatch), doc.Source.PackageSHA256); err == nil || out != nil {
			t.Fatalf("accepted unsupported section: %s", section)
		}
	}
	for _, page := range []string{strings.Replace(nativePageTestPatch, `16838`, `1.5`, 1), strings.Replace(nativePageTestPatch, `16838`, `1000`, 1), strings.Replace(nativePageTestPatch, `720`, `-1`, 1), strings.Replace(nativePageTestPatch, `"landscape"`, `null`, 1), strings.Replace(nativePageTestPatch, `"landscape"`, `"sideways"`, 1), strings.Replace(nativePageTestPatch, `"width_twips":16838`, `"width_twips":16838,"width_twips":16838`, 1)} {
		if _, err := decodeNativeSectionPage([]byte(page)); err == nil {
			t.Fatalf("accepted invalid page %s", page)
		}
	}
}

func TestNativeSectionPageRefusesMultipleSectionsAndSignatures(t *testing.T) {
	section := `<w:sectPr>` + nativePageTestSize + nativePageTestMargins + `</w:sectPr>`
	multi := nativeMutationMain(`<w:p><w:pPr>` + section + `</w:pPr><w:r><w:t>First</w:t></w:r></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p>` + section)
	signed := nativeMutationParts(nativeMutationMain(`<w:p><w:r><w:t>Signed</w:t></w:r></w:p>` + section))
	signed["[Content_Types].xml"] = strings.Replace(signed["[Content_Types].xml"], `</Types>`, `<Override PartName="/_xmlsignatures/origin.sigs" ContentType="application/vnd.openxmlformats-package.digital-signature-origin"/></Types>`, 1)
	signed["_rels/.rels"] = strings.Replace(signed["_rels/.rels"], `</Relationships>`, `<Relationship Id="signature-origin" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin" Target="_xmlsignatures/origin.sigs"/></Relationships>`, 1)
	signed["_xmlsignatures/origin.sigs"] = `<SignatureOrigin xmlns="urn:injoffice:test"/>`
	for _, test := range []struct {
		parts map[string]string
		code  string
	}{{nativeMutationParts(multi), "UNSUPPORTED_SECTION_STRUCTURE"}, {signed, "SIGNED_PACKAGE"}} {
		source := buildNativeDOCX(t, nativeEntries(test.parts))
		doc, err := ExtractNativeDocumentV1(source)
		if err != nil {
			t.Fatal(err)
		}
		out, err := ApplyNativeMutationPayloadV1(source, nativePageTestPayload(doc.Sections[0], nativePageTestPatch), doc.Source.PackageSHA256)
		if out != nil {
			t.Fatal("refusal returned candidate bytes")
		}
		assertNativeMutationCode(t, err, test.code)
	}
}

func TestNativeSectionPagePreservesAttributeNamespaceAliases(t *testing.T) {
	main := nativeMutationMain(`<w:p><w:r><w:t>Body</w:t></w:r></w:p><w:sectPr>` + strings.ReplaceAll(nativePageTestSize, "w:w=", "alias:w=") + nativePageTestMargins + `</w:sectPr>`)
	main = strings.Replace(main, `xmlns:w14=`, `xmlns:alias="`+testW+`" xmlns:w14=`, 1)
	source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(main)))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	result, err := ApplyNativeMutationPayloadV1(source, nativePageTestPayload(doc.Sections[0], nativePageTestPatch), doc.Source.PackageSHA256)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(readNativeZipPart(t, result.Package, "word/document.xml")), `alias:w="16838"`) {
		t.Fatal("lost the source attribute prefix")
	}
}
