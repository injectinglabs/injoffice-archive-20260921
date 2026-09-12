package docxpatch

import (
	"encoding/xml"
	"strings"
	"testing"
)

func TestFontTableIgnorableDeclarationIsBounded(t *testing.T) {
	for _, test := range []struct {
		attrs, text string
		want        bool
	}{
		{``, ``, true},
		{` xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"`, ``, true},
		{` xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"`, ``, false},
		{` xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="urn:spoof" mc:Ignorable="w14"`, ``, false},
		{` extra="1"`, ``, false}, {``, `bad`, false},
	} {
		root, err := parseNativeXML("fonts.xml", []byte(`<w:fonts xmlns:w="`+wordMLTransitional+`"`+test.attrs+`>`+test.text+`</w:fonts>`))
		if err != nil {
			t.Fatal(err)
		}
		if nativeQualifiedFontTableOwner(root) != test.want {
			t.Fatalf("%s %s", test.attrs, test.text)
		}
	}
}

func TestUnusedFontDescriptorsRequireCompleteResolvedUsage(t *testing.T) {
	for _, usage := range []string{"unused", "run", "mark", "marker", "alias", "uncertain"} {
		node := &nativeXMLNode{Name: xml.Name{Space: wordMLTransitional, Local: "charset"}, Path: "/w:fonts[1]/w:font[2]/w:charset[1]"}
		resolver := &nativeLayoutResolver{doc: &NativeDocumentV1{DocumentID: "document:test"}, diagnosticSet: map[string]bool{}, unusedFontDescriptors: []nativeUnusedFontDescriptor{{name: "Symbol", alias: nativeString("Alias"), part: "fonts.xml", node: node}}}
		arial := NativeResolvedRunPropertiesV1{FontFamily: nativeString("Arial")}
		symbol := NativeResolvedRunPropertiesV1{FontFamily: nativeString("Symbol")}
		layout := &NativeResolvedLayoutInputV1{Runs: []NativeResolvedRunV1{{Properties: arial}}, Paragraphs: []NativeResolvedParagraphV1{{ParagraphMarkProperties: arial}}}
		switch usage {
		case "run":
			layout.Runs[0].Properties = symbol
		case "mark":
			layout.Paragraphs[0].ParagraphMarkProperties = symbol
		case "marker":
			layout.Paragraphs[0].Numbering = &NativeResolvedNumberingV1{Marker: symbol}
		case "alias":
			layout.Runs[0].Properties.FontFamily = nativeString("ALIAS")
		case "uncertain":
			layout.Runs[0].Properties.FontFamily = nil
		}
		resolver.resolveUnusedFontDescriptors(layout)
		want := "UNMODELED_FONT_METADATA"
		if usage == "unused" {
			want = "FONT_MATCHING_METADATA_PRESERVED"
		}
		if len(resolver.diagnostics) != 1 || resolver.diagnostics[0].Code != want {
			t.Fatalf("%s: %#v", usage, resolver.diagnostics)
		}
	}
}

func TestSourceUnusedLegacyDescriptorsAndActiveCharsetEE(t *testing.T) {
	for _, family := range []string{"Arial", "Symbol"} {
		parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="` + family + `"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`)
		parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:r><w:t>test</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
		parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/word/fonts.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/></Types>`, 1)
		parts["word/_rels/document.xml.rels"] = strings.Replace(parts["word/_rels/document.xml.rels"], `</Relationships>`, `<Relationship Id="fonts" Type="`+relBaseTransitional+`fontTable" Target="fonts.xml"/></Relationships>`, 1)
		parts["word/fonts.xml"] = `<w:fonts xmlns:w="` + wordMLTransitional + `"><w:font w:name="Arial"><w:charset w:val="EE"/></w:font><w:font w:name="Symbol"><w:charset w:val="02"/><w:notTrueType/></w:font></w:fonts>`
		layout, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
		if err != nil {
			t.Fatal(err)
		}
		if hasResolutionDiagnostic(layout, "UNMODELED_FONT_METADATA") != (family == "Symbol") {
			t.Fatalf("family %s: %#v", family, layout.Diagnostics)
		}
		if !hasResolutionDiagnostic(layout, "FONT_MATCHING_METADATA_PRESERVED") {
			t.Fatal("lost source descriptor evidence")
		}
	}
}
