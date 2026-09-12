package docxpatch

import (
	"strings"
	"testing"
)

func TestNativeFontDescriptorQualification(t *testing.T) {
	valid := []string{`<w:panose1 w:val="020F0502020204030204"/>`, `<w:charset w:val="00"/>`, `<w:family w:val="swiss"/>`, `<w:pitch w:val="variable"/>`, `<w:sig w:usb0="E0002EFF" w:usb1="C000785B" w:usb2="00000009" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/>`}
	invalid := []string{`<w:charset w:val="02"/>`, `<w:panose1 w:val="ZZ0F0502020204030204"/>`, `<w:family w:val="unknown"/>`, `<w:pitch w:val="variable" extra="1"/>`, `<w:pitch w:val="variable">text</w:pitch>`, `<w:pitch w:val="variable"><w:b/></w:pitch>`, `<w:sig w:usb0="00000000"/>`, `<w:notTrueType/>`, `<x:pitch xmlns:x="urn:foreign" w:val="variable"/>`}
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		for _, markup := range append(append([]string{}, valid...), invalid...) {
			root, err := parseNativeXML("fonts.xml", []byte(`<w:font xmlns:w="`+ns+`" w:name="Arial">`+markup+`</w:font>`))
			if err != nil {
				t.Fatal(err)
			}
			want := false
			for _, candidate := range valid {
				if markup == candidate {
					want = true
				}
			}
			if got := nativeQualifiedFontDescriptor(root.Children[0], root, ns); got != want {
				t.Fatalf("%s: got %v want %v", markup, got, want)
			}
		}
		for _, markup := range valid {
			for _, owner := range []string{` extra="1">`, `>unexpected text`} {
				root, err := parseNativeXML("fonts.xml", []byte(`<w:font xmlns:w="`+ns+`" w:name="Arial"`+owner+markup+`</w:font>`))
				if err != nil {
					t.Fatal(err)
				}
				if nativeQualifiedFontDescriptor(root.Children[0], root, ns) {
					t.Fatalf("malformed font owner qualified: %s", owner)
				}
			}
			root, err := parseNativeXML("fonts.xml", []byte(`<w:font xmlns:w="`+ns+`" w:name="Arial">`+markup+markup+`</w:font>`))
			if err != nil {
				t.Fatal(err)
			}
			for _, child := range root.Children {
				if nativeQualifiedFontDescriptor(child, root, ns) {
					t.Fatalf("duplicate qualified: %s", markup)
				}
			}
		}
	}
	parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `"/>`)
	parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/word/fonts.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/></Types>`, 1)
	parts["word/_rels/document.xml.rels"] = strings.Replace(parts["word/_rels/document.xml.rels"], `</Relationships>`, `<Relationship Id="fonts" Type="`+relBaseTransitional+`fontTable" Target="fonts.xml"/></Relationships>`, 1)
	parts["word/fonts.xml"] = `<w:fonts xmlns:w="` + wordMLTransitional + `"><w:font w:name="Arial">` + strings.Join(valid, "") + `</w:font></w:fonts>`
	data := buildNativeDOCX(t, nativeEntries(parts))
	before := string(data)
	resolved, err := ResolveNativeDocumentLayoutV1(data)
	if err != nil {
		t.Fatal(err)
	}
	if !hasResolutionDiagnostic(resolved, "FONT_MATCHING_METADATA_PRESERVED") || hasResolutionDiagnostic(resolved, "UNMODELED_FONT_METADATA") {
		t.Fatalf("diagnostics: %#v", resolved.Diagnostics)
	}
	if string(data) != before {
		t.Fatal("source changed")
	}
	for _, owner := range []string{` extra="1">`, `>unexpected text`} {
		parts["word/fonts.xml"] = `<w:fonts xmlns:w="` + wordMLTransitional + `"` + owner + `<w:font w:name="Arial">` + strings.Join(valid, "") + `</w:font></w:fonts>`
		resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
		if err == nil && (!hasResolutionDiagnostic(resolved, "UNMODELED_FONT_METADATA") || hasResolutionDiagnostic(resolved, "FONT_MATCHING_METADATA_PRESERVED")) {
			t.Fatalf("malformed font table owner qualified: %s", owner)
		}
	}
}
