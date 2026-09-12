package docxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func nativeThemeCascadeParts(theme string) map[string]string {
	styles := `<w:styles xmlns:w="` + wordMLTransitional + `"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:color w:themeColor="accent1"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Heading"><w:rPr><w:rFonts w:asciiTheme="majorAscii" w:hAnsiTheme="majorHAnsi"/></w:rPr></w:style><w:style w:type="character" w:styleId="Minor"><w:rPr><w:rFonts w:asciiTheme="minorAscii" w:hAnsiTheme="minorHAnsi"/></w:rPr></w:style></w:styles>`
	parts := resolvedStylesTestParts(styles)
	parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/assets/theme.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>`, 1)
	parts["word/_rels/document.xml.rels"] = strings.Replace(parts["word/_rels/document.xml.rels"], `</Relationships>`, `<Relationship Id="theme" Type="`+relBaseTransitional+`theme" Target="../assets/THEME.xml"/></Relationships>`, 1)
	parts["Assets/Theme.XML"] = theme
	return parts
}

func TestNativeThemeResolvesBeforeDocumentAndStyleCascades(t *testing.T) {
	for _, strict := range []bool{false, true} {
		parts := nativeThemeCascadeParts(nativeThemeSrgbFixtureXML(drawingMLTransitional))
		parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:r><w:t>default</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading"/></w:pPr><w:r><w:t>heading</w:t></w:r><w:r><w:rPr><w:rStyle w:val="Minor"/></w:rPr><w:t>character</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Direct Face" w:hAnsi="Direct Face"/><w:color w:val="123456"/></w:rPr><w:t>direct</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
		if strict {
			for key, value := range parts {
				parts[key] = strings.ReplaceAll(strings.ReplaceAll(strings.ReplaceAll(value, wordMLTransitional, wordMLStrict), drawingMLTransitional, drawingMLStrict), relBaseTransitional, relBaseStrict)
			}
		}
		data := buildNativeDOCX(t, nativeEntries(parts))
		before := append([]byte(nil), data...)
		resolved, err := ResolveNativeDocumentLayoutV1(data)
		if err != nil {
			t.Fatal(err)
		}
		if len(resolved.Runs) != 4 {
			t.Fatalf("unexpected runs %#v", resolved.Runs)
		}
		for index, family := range []string{"Calibri", "Calibri Light", "Calibri", "Direct Face"} {
			run := resolved.Runs[index].Properties
			color := "4472C4"
			if index == 3 {
				color = "123456"
			}
			if run.FontFamily == nil || *run.FontFamily != family || run.Color == nil || *run.Color != color {
				t.Fatalf("strict=%v run%d did not resolve actual theme/cascade %#v", strict, index, run)
			}
		}
		if hasResolutionDiagnostic(resolved, "THEME_FONT_PRESERVED") || hasResolutionDiagnostic(resolved, "THEME_COLOR_PRESERVED") {
			t.Fatalf("related theme was unavailable during parse %#v", resolved.Diagnostics)
		}
		if !bytes.Equal(data, before) {
			t.Fatal("theme resolution changed source bytes")
		}
	}
}

func TestNativeThemeFontCascadeDoesNotChooseAmbiguousOrMissingLatinFaces(t *testing.T) {
	base := nativeThemeSrgbFixtureXML(drawingMLTransitional)
	for name, theme := range map[string]string{
		"missing Latin":      strings.Replace(base, `<a:latin typeface="Calibri"/>`, ``, 1),
		"duplicate Latin":    strings.Replace(base, `<a:latin typeface="Calibri"/>`, `<a:latin typeface="Calibri"/><a:latin typeface="Other"/>`, 1),
		"spoof Latin":        strings.Replace(base, `<a:latin typeface="Calibri"/>`, `<x:latin xmlns:x="urn:spoof" typeface="Other"/><a:latin typeface="Calibri"/>`, 1),
		"duplicate scheme":   strings.Replace(base, `</a:fontScheme>`, `</a:fontScheme><a:fontScheme name="Other"/>`, 1),
		"duplicate elements": strings.Replace(base, `</a:theme>`, `<a:themeElements/></a:theme>`, 1),
		"nested Latin":       strings.Replace(base, `<a:latin typeface="Calibri"/>`, `<a:latin typeface="Calibri"><a:latin typeface="Other"/></a:latin>`, 1),
	} {
		t.Run(name, func(t *testing.T) {
			resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(nativeThemeCascadeParts(theme))))
			if err != nil {
				t.Fatal(err)
			}
			if !hasResolutionDiagnostic(resolved, "THEME_FONT_PRESERVED") || resolved.Runs[0].Properties.FontFamily != nil {
				t.Fatalf("ambiguous theme became exact %#v", resolved)
			}
		})
	}
}

func TestNativeThemeResolvesBeforeNumberingMarkerCascade(t *testing.T) {
	parts := nativeThemeCascadeParts(nativeThemeSrgbFixtureXML(drawingMLTransitional))
	parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>`, 1)
	parts["word/_rels/document.xml.rels"] = strings.Replace(parts["word/_rels/document.xml.rels"], `</Relationships>`, `<Relationship Id="numbering" Type="`+relBaseTransitional+`numbering" Target="numbering.xml"/></Relationships>`, 1)
	parts["word/numbering.xml"] = `<w:numbering xmlns:w="` + wordMLTransitional + `"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorAscii" w:hAnsiTheme="majorHAnsi"/><w:color w:themeColor="accent1"/></w:rPr></w:lvl></w:abstractNum><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`
	parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>List item</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
	resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	numbering := resolved.Paragraphs[0].Numbering
	if numbering == nil || numbering.Marker.FontFamily == nil || *numbering.Marker.FontFamily != "Calibri Light" || numbering.Marker.Color == nil || *numbering.Marker.Color != "4472C4" {
		t.Fatalf("numbering marker did not resolve related theme %#v; diagnostics %#v", numbering, resolved.Diagnostics)
	}
	if hasResolutionDiagnostic(resolved, "THEME_FONT_PRESERVED") || hasResolutionDiagnostic(resolved, "THEME_COLOR_PRESERVED") {
		t.Fatalf("numbering parsed before theme %#v", resolved.Diagnostics)
	}
}
