package docxpatch

import (
	"encoding/json"
	"strings"
	"testing"
)

func equationContextParts() map[string]string {
	parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + testW + `"><w:style w:type="paragraph" w:styleId="S"><w:pPr><w:suppressAutoHyphens w:val="true"/><w:tabs><w:tab w:val="left" w:pos="709" w:leader="none"/></w:tabs></w:pPr></w:style></w:styles>`)
	parts["word/document.xml"] = `<w:document xmlns:w="` + testW + `"><w:body><w:p><w:pPr><w:pStyle w:val="S"/></w:pPr><m:oMath xmlns:m="` + nativePartialMathNamespace(testW) + `"><m:r><m:t>x</m:t></m:r></m:oMath></w:p><w:sectPr><w:textDirection w:val="lrTb"/></w:sectPr></w:body></w:document>`
	parts["word/fontTable.xml"] = `<w:fonts xmlns:w="` + testW + `"><w:font w:name="Arial"><w:charset w:val="00" w:characterSet="windows-1252"/></w:font></w:fonts>`
	parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/></Types>`, 1)
	parts["word/_rels/document.xml.rels"] = strings.Replace(parts["word/_rels/document.xml.rels"], `</Relationships>`, `<Relationship Id="fonts" Type="`+relBaseTransitional+`fontTable" Target="fontTable.xml"/></Relationships>`, 1)
	return parts
}
func TestEquationContextExactSourceNotices(t *testing.T) {
	parts := equationContextParts()
	source := buildNativeDOCX(t, nativeEntries(parts))
	doc, e := ExtractNativeDocumentV1(source)
	if e != nil {
		t.Fatal(e)
	}
	layout, e := ResolveNativeDocumentLayoutV1(source)
	if e != nil {
		t.Fatal(e)
	}
	eq, e := inspectNativePartialEquations(source, doc)
	if e != nil {
		t.Fatal(e)
	}
	before, _ := json.Marshal([]any{doc, layout})
	notices, e := inspectNativeEquationContext(source, doc, layout, eq)
	if e != nil {
		t.Fatal(e)
	}
	after, _ := json.Marshal([]any{doc, layout})
	if string(before) != string(after) {
		t.Fatal("source diagnostics changed")
	}
	if len(notices) != 4 {
		t.Fatalf("expected4 bounded notices, got %#v", notices)
	}
	for _, n := range notices {
		raw := []byte(parts[n.Anchor.PartName])
		if n.PackageSHA256 != doc.Source.PackageSHA256 || n.PartSHA256 != nativeSHA(raw) || n.Anchor.XMLSHA256 != nativeSHA(raw[*n.Anchor.StartByte:*n.Anchor.EndByte]) {
			t.Fatal("source hash mismatch")
		}
	}
}
func TestEquationContextRejectsUnqualifiedLeaves(t *testing.T) {
	cases := []struct{ part, old, new, kind string }{
		{"word/document.xml", `w:val="lrTb"`, `w:val="tbRl"`, "horizontal-section"},
		{"word/document.xml", `<w:textDirection w:val="lrTb"/>`, `<w:textDirection w:val="lrTb"/><w:textDirection w:val="lrTb"/>`, "horizontal-section"},
		{"word/document.xml", `<w:textDirection w:val="lrTb"/>`, `<x:textDirection xmlns:x="urn:foreign" w:val="lrTb"/>`, "horizontal-section"},
		{"word/styles.xml", `w:val="true"`, `w:val="false"`, "disabled-paragraph-hyphenation"},
		{"word/styles.xml", `<w:suppressAutoHyphens w:val="true"/>`, `<w:suppressAutoHyphens w:val="true"><w:vanish/></w:suppressAutoHyphens>`, "disabled-paragraph-hyphenation"},
		{"word/styles.xml", `<w:tab w:val="left" w:pos="709" w:leader="none"/>`, ``, "unused-paragraph-tab-stops"},
		{"word/styles.xml", `w:pos="709"`, `w:pos="-1"`, "unused-paragraph-tab-stops"},
		{"word/styles.xml", `w:leader="none"`, `w:leader="dot"`, "unused-paragraph-tab-stops"},
		{"word/styles.xml", `<w:tab w:val="left" w:pos="709" w:leader="none"/>`, `<w:tab w:val="left" w:pos="709"/><w:tab w:val="left" w:pos="709"/>`, "unused-paragraph-tab-stops"},
		{"word/fontTable.xml", `w:characterSet="windows-1252"`, `w:characterSet="unknown"`, "ignored-font-matching"},
		{"word/fontTable.xml", `<w:charset w:val="00" w:characterSet="windows-1252"/>`, `<w:charset w:val="00" w:characterSet="windows-1252"/><w:charset w:val="00"/>`, "ignored-font-matching"},
	}
	for _, c := range cases {
		parts := equationContextParts()
		parts[c.part] = strings.Replace(parts[c.part], c.old, c.new, 1)
		source := buildNativeDOCX(t, nativeEntries(parts))
		data, e := InspectNativePartialSourceV1(source)
		if e != nil {
			continue
		}
		var b struct {
			Notices []NativeEquationContextNoticeV1 `json:"equation_context_notices"`
		}
		if e = json.Unmarshal(data, &b); e != nil {
			t.Fatal(e)
		}
		for _, n := range b.Notices {
			if n.Kind == c.kind {
				t.Fatalf("unqualified %s produced notice: %s", c.new, data)
			}
		}
	}
}
