package docxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func TestNativeWholeTableTextStyleCascade(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, test := range []struct {
			name, normal, paragraph, direct, character string
			fontSize                                   int
			alignment                                  string
		}{
			{"table overrides defaults", "", "", "", "", 28, "center"},
			{"default eleven point exception", `<w:pPr><w:jc w:val="left"/></w:pPr><w:rPr><w:sz w:val="22"/></w:rPr>`, "", "", "", 28, "center"},
			{"default twelve point exception", `<w:pPr><w:jc w:val="left"/></w:pPr><w:rPr><w:sz w:val="24"/></w:rPr>`, "", "", "", 28, "center"},
			{"other default size overrides", `<w:rPr><w:sz w:val="26"/></w:rPr>`, "", "", "", 26, "center"},
			{"inherited default exceptions", `<w:basedOn w:val="Ancestor"/>`, "", "", "", 28, "center"},
			{"default overrides inherited exception", `<w:basedOn w:val="Ancestor"/><w:rPr><w:sz w:val="26"/></w:rPr>`, "", "", "", 26, "center"},
			{"child paragraph overrides inherited default", `<w:basedOn w:val="Ancestor"/>`, `<w:pPr><w:jc w:val="left"/></w:pPr><w:rPr><w:sz w:val="22"/></w:rPr>`, "", "", 22, "left"},
			{"paragraph style overrides", "", `<w:pPr><w:jc w:val="right"/></w:pPr><w:rPr><w:sz w:val="32"/></w:rPr>`, "", "", 32, "right"},
			{"character style overrides", "", "", "", `<w:sz w:val="30"/>`, 30, "center"},
			{"direct run overrides", "", "", `<w:sz w:val="36"/>`, `<w:sz w:val="30"/>`, 36, "center"},
		} {
			t.Run(test.name+map[bool]string{false: " transitional", true: " strict"}[strict], func(t *testing.T) {
				styles := `<w:styles xmlns:w="` + wordMLTransitional + `"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="20"/><w:color w:val="000000"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="table" w:styleId="Base"><w:pPr><w:jc w:val="right"/><w:spacing w:after="120"/></w:pPr><w:rPr><w:sz w:val="26"/><w:b/></w:rPr></w:style><w:style w:type="table" w:styleId="Table"><w:basedOn w:val="Base"/><w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:sz w:val="28"/><w:color w:val="224466"/><w:b/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Normal" w:default="1">` + test.normal + `</w:style><w:style w:type="paragraph" w:styleId="Body"><w:basedOn w:val="Normal"/>` + test.paragraph + `</w:style><w:style w:type="character" w:styleId="Character"><w:rPr>` + test.character + `</w:rPr></w:style></w:styles>`
				styles = strings.Replace(styles, `</w:styles>`, `<w:style w:type="paragraph" w:styleId="Ancestor"><w:pPr><w:jc w:val="left"/></w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:style></w:styles>`, 1)
				parts := resolvedStylesTestParts(styles)
				parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:r><w:t>outside</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblStyle w:val="Table"/></w:tblPr><w:tr><w:tc><w:p><w:pPr><w:pStyle w:val="Body"/></w:pPr><w:r><w:rPr><w:rStyle w:val="Character"/>` + test.direct + `</w:rPr><w:t>inside</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`
				if strict {
					for key, value := range parts {
						parts[key] = strings.ReplaceAll(strings.ReplaceAll(value, wordMLTransitional, wordMLStrict), relBaseTransitional, relBaseStrict)
					}
				}
				data := buildNativeDOCX(t, nativeEntries(parts))
				before := append([]byte(nil), data...)
				resolved, err := ResolveNativeDocumentLayoutV1(data)
				if err != nil {
					t.Fatal(err)
				}
				if len(resolved.Runs) != 2 || len(resolved.Paragraphs) != 2 {
					t.Fatalf("unexpected content %#v", resolved)
				}
				p, r := resolved.Paragraphs[1].Properties, resolved.Runs[1].Properties
				if r.FontSizeHalfPoint == nil || *r.FontSizeHalfPoint != test.fontSize || p.Alignment == nil || *p.Alignment != test.alignment || p.SpacingAfterTwips == nil || *p.SpacingAfterTwips != 120 || r.Color == nil || *r.Color != "224466" || r.Bold == nil || *r.Bold {
					t.Fatalf("incorrect table cascade p=%#v r=%#v", p, r)
				}
				outside := resolved.Runs[0].Properties
				if outside.Color == nil || *outside.Color != "000000" || *resolved.Paragraphs[0].Properties.Alignment != "left" {
					t.Fatal("table formatting leaked into body")
				}
				if hasResolutionDiagnostic(resolved, "TABLE_STYLE_EFFECTS_PRESERVED") {
					t.Fatalf("supported whole-table text refused: %#v", resolved.Diagnostics)
				}
				if !bytes.Equal(data, before) {
					t.Fatal("source bytes changed")
				}
			})
		}
	}
}

func TestNativeWholeTableTextStyleKeepsUnsupportedEffectsExplicit(t *testing.T) {
	for markup, code := range map[string]string{
		`<w:pPr><w:tabs/></w:pPr>`:                                             "UNMODELED_PARAGRAPH_PROPERTY",
		`<w:rPr><w:emboss/></w:rPr>`:                                           "UNMODELED_RUN_PROPERTY",
		`<w:trPr><w:cantSplit/></w:trPr>`:                                      "TABLE_STYLE_EFFECTS_PRESERVED",
		`<w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr>`: "CONDITIONAL_TABLE_STYLE_PRESERVED",
	} {
		parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `"><w:style w:type="table" w:styleId="Table">` + markup + `</w:style></w:styles>`)
		parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:tbl><w:tblPr><w:tblStyle w:val="Table"/></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>inside</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`
		resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
		if err != nil {
			t.Fatal(err)
		}
		if !hasResolutionDiagnostic(resolved, code) {
			t.Fatalf("missing %s for %s: %#v", code, markup, resolved.Diagnostics)
		}
	}
}
