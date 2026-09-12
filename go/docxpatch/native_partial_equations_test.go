package docxpatch

import (
	"encoding/json"
	"strings"
	"testing"
)

func TestPartialEquationSourceAndStrictRefusal(t *testing.T) {
	text := `<m:r><m:t>x&lt;y</m:t></m:r>`
	for _, body := range []string{text, `<m:f><m:num>` + text + `</m:num><m:den>` + text + `</m:den></m:f>`, `<m:sSup><m:e>` + text + `</m:e><m:sup>` + text + `</m:sup></m:sSup>`, `<m:sSub><m:e>` + text + `</m:e><m:sub>` + text + `</m:sub></m:sSub>`, `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>` + text + `</m:e></m:rad>`} {
		xml := nativeMutationMain(`<w:p><w:r><w:t>Before</w:t></w:r><m:oMath xmlns:m="` + nativePartialMathNamespace(testW) + `">` + body + `</m:oMath></w:p>`)
		source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(xml)))
		doc, err := ExtractNativeDocumentV1(source)
		if err != nil {
			t.Fatal(err)
		}
		if doc.Body.Blocks[0].Paragraph.EditPolicy.Mode != "read-only" {
			t.Fatal("equation paragraph became editable")
		}
		p := doc.Body.Blocks[0].Paragraph
		for _, target := range []NativeDOCXTextMutationV1{{TargetKind: "paragraph", TargetID: p.ID, ExpectedXMLSHA256: p.Anchor.XMLSHA256, Text: "Changed"}, {TargetKind: "run", TargetID: p.Runs[0].ID, ExpectedXMLSHA256: p.Runs[0].Anchor.XMLSHA256, Text: "Changed"}} {
			if _, err := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXTextMutationV1{target}); err == nil {
				t.Fatal("equation-adjacent text mutation admitted")
			}
		}
		data, err := InspectNativePartialSourceV1(source)
		if err != nil {
			t.Fatal(err)
		}
		var result struct {
			Equations []NativePartialEquationV1 `json:"equations"`
		}
		if err = json.Unmarshal(data, &result); err != nil {
			t.Fatal(err)
		}
		if len(result.Equations) != 1 || result.Equations[0].Status != "supported" || result.Equations[0].PackageSHA256 != doc.Source.PackageSHA256 || result.Equations[0].Tree == nil {
			t.Fatalf("missing source-bound equation: %s", data)
		}
		found := false
		for _, d := range doc.Unsupported {
			if d.ID == result.Equations[0].DiagnosticID && d.Code == "UNMODELED_PARAGRAPH_CONTENT" {
				found = true
			}
		}
		if !found {
			t.Fatal("strict diagnostic removed")
		}
	}
}
func TestPartialEquationUnknownContentRemainsOmitted(t *testing.T) {
	for _, body := range []string{`<m:r><m:rPr/><m:t>x</m:t></m:r>`, `<m:r><m:t href="evil">x</m:t></m:r>`, `<m:foo/>`, `<m:f><m:num><m:r><m:t>x</m:t></m:r></m:num></m:f>`, `<m:r xmlns:m="urn:evil"><m:t>x</m:t></m:r>`, strings.Repeat(`<m:e>`, 34) + `<m:r><m:t>x</m:t></m:r>` + strings.Repeat(`</m:e>`, 34)} {
		source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(nativeMutationMain(`<w:p><m:oMath xmlns:m="`+nativePartialMathNamespace(testW)+`">`+body+`</m:oMath></w:p>`))))
		data, err := InspectNativePartialSourceV1(source)
		if err != nil {
			t.Fatal(err)
		}
		var result struct {
			Equations []NativePartialEquationV1 `json:"equations"`
		}
		json.Unmarshal(data, &result)
		if len(result.Equations) != 1 || result.Equations[0].Status != "omitted" || result.Equations[0].Tree != nil {
			t.Fatalf("unsupported math exposed: %s", data)
		}
	}
}
func TestPartialEquationDoesNotAdmitSpoofedWordRunsElsewhere(t *testing.T) {
	for _, body := range []string{`<w:p><m:r xmlns:m="` + nativePartialMathNamespace(testW) + `"><m:t>x</m:t></m:r></w:p>`, `<w:p><m:oMath xmlns:m="urn:evil"><m:r><m:t>x</m:t></m:r></m:oMath></w:p>`} {
		source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(nativeMutationMain(body))))
		if _, err := ExtractNativeDocumentV1(source); err == nil {
			t.Fatal("namespace spoofing admitted")
		}
	}
}
func TestPartialEquationRejectsIllegalWrapperPlacement(t *testing.T) {
	for _, body := range []string{`<m:num><m:r><m:t>x</m:t></m:r></m:num>`, `<m:t>x</m:t>`, `<m:oMathPara><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara>`} {
		root, err := parseNativeXML("math.xml", []byte(`<m:oMath xmlns:m="`+nativePartialMathNamespace(testW)+`">`+body+`</m:oMath>`))
		if err != nil {
			t.Fatal(err)
		}
		count, units := 0, 0
		if _, ok := nativePartialMathTree(root, nativePartialMathNamespace(testW), 0, &count, &units); ok {
			t.Fatal("illegal wrapper produced a tree")
		}
	}
}
func TestPartialEquationStrictDialectAdmission(t *testing.T) {
	for _, ns := range []string{nativePartialMathNamespace(testWS), nativePartialMathNamespace(testW)} {
		main := strings.ReplaceAll(nativeMutationMain(`<w:p><m:oMath xmlns:m="`+ns+`"><m:r><m:t>x</m:t></m:r></m:oMath></w:p>`), testW, testWS)
		parts := nativeMutationParts(main)
		parts["_rels/.rels"] = strings.ReplaceAll(parts["_rels/.rels"], relBaseTransitional, relBaseStrict)
		source := buildNativeDOCX(t, nativeEntries(parts))
		_, err := ExtractNativeDocumentV1(source)
		if ns == nativePartialMathNamespace(testWS) && err != nil {
			t.Fatal(err)
		}
		if ns != nativePartialMathNamespace(testWS) && err == nil {
			t.Fatal("wrong math dialect admitted")
		}
	}
}
