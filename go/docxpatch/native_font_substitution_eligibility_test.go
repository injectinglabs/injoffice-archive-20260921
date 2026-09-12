package docxpatch

import (
	"reflect"
	"strings"
	"testing"
)

func fontSubstitutionTestPackage(t *testing.T, ns, descriptor, extra string) []byte {
	parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + ns + `"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`)
	parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/word/fonts.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/></Types>`, 1)
	parts["word/_rels/document.xml.rels"] = strings.Replace(parts["word/_rels/document.xml.rels"], `</Relationships>`, `<Relationship Id="fonts" Type="`+relBaseTransitional+`fontTable" Target="fonts.xml"/></Relationships>`, 1)
	parts["word/fonts.xml"] = `<w:fonts xmlns:w="` + ns + `"><w:font w:name="Arial">` + descriptor + `</w:font>` + extra + `</w:fonts>`
	if ns == wordMLStrict {
		for name, value := range parts {
			parts[name] = strings.ReplaceAll(strings.ReplaceAll(value, wordMLTransitional, wordMLStrict), relBaseTransitional, relBaseStrict)
		}
	}
	return buildNativeDOCX(t, nativeEntries(parts))
}
func TestNativeFontSubstitutionEligibility(t *testing.T) {
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		data := fontSubstitutionTestPackage(t, ns, `<w:charset w:val="EE"/><w:panose1 w:val="020F0502020204030204"/>`, `<w:font w:name="Unused Symbol"><w:charset w:val="02"/><w:notTrueType/></w:font>`)
		original := string(data)
		baseline, e := ResolveNativeDocumentLayoutV1(data)
		if e != nil {
			t.Fatal(e)
		}
		result, e := ExtractNativeDOCXFontSubstitutionEligibilityV1(data)
		if e != nil {
			t.Fatal(e)
		}
		if len(result.Facts) != 4 || result.Facts[0].Use != "latin-matching" || result.Facts[2].Use != "unused" || result.FontTable == nil || result.FontTable.PartName != "word/fonts.xml" {
			t.Fatalf("invalid facts %#v", result)
		}
		if baseline.Fonts[0].Name != "Arial" || baseline.Fonts[1].Name != "Unused Symbol" || result.Facts[0].Path != "/w:fonts[1]/w:font[1]/w:charset[1]" || result.Facts[2].Path != "/w:fonts[1]/w:font[2]/w:charset[1]" {
			t.Fatal("descriptor owners must retain original declaration order")
		}
		after, e := ResolveNativeDocumentLayoutV1(data)
		if e != nil || !reflect.DeepEqual(after, baseline) || string(data) != original {
			t.Fatal("source or strict resolver changed")
		}
	}
}
func TestNativeFontSubstitutionEligibilityRefusesAmbiguity(t *testing.T) {
	unusedAlias := fontSubstitutionTestPackage(t, wordMLTransitional, `<w:charset w:val="00"/>`, `<w:font w:name="Unused"><w:charset w:val="00"/></w:font><w:font w:name="Other"><w:altName w:val="Unused"/><w:charset w:val="00"/></w:font>`)
	if _, e := ExtractNativeDOCXFontSubstitutionEligibilityV1(unusedAlias); e != nil {
		t.Fatalf("unused alias overlap participates in no font selection: %v", e)
	}
	for _, descriptor := range []string{`<w:charset w:val="02"/>`, `<w:charset w:val="80"/>`, `<w:charset w:val="00"/><w:charset w:val="00"/>`, `<w:charset w:val="00" extra="1"/>`, `<w:charset w:val="00">text</w:charset>`, `<w:charset w:val="00"/><w:notTrueType/>`, `<w:charset w:val="00"/><w:panose1 w:val="05000000000000000000"/>`, `<w:charset w:val="00"/><x:pitch xmlns:x="urn:foreign" w:val="fixed"/>`} {
		if _, e := ExtractNativeDOCXFontSubstitutionEligibilityV1(fontSubstitutionTestPackage(t, wordMLTransitional, descriptor, "")); e == nil {
			t.Fatalf("unsafe descriptor accepted: %s", descriptor)
		}
	}
	data := fontSubstitutionTestPackage(t, wordMLTransitional, `<w:charset w:val="00"/>`, `<w:font w:name="Unused Symbol"><w:altName w:val="Arial"/><w:charset w:val="02"/></w:font>`)
	if _, e := ExtractNativeDOCXFontSubstitutionEligibilityV1(data); e == nil {
		t.Fatal("used symbol alias accepted")
	}
	data = fontSubstitutionTestPackage(t, wordMLTransitional, `<w:charset w:val="00"/>`, `<w:font w:name="Ambiguous"><w:altName w:val="Arial"/><w:charset w:val="00"/></w:font>`)
	if _, e := ExtractNativeDOCXFontSubstitutionEligibilityV1(data); e == nil {
		t.Fatal("ambiguous Latin alias accepted")
	}
}
