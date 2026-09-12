package docxpatch

import (
	"strings"
	"testing"
)

func TestPageFieldSourcePropertiesPreserveResultAuthority(t *testing.T) {
	props := `<w:rPr><w:rStyle w:val="PageNumber"/><w:sz w:val="16"/><w:szCs w:val="40"/></w:rPr>`
	resultProps := strings.Replace(props, `<w:sz `, `<w:noProof/><w:sz `, 1)
	run := func(content string) string {
		return `<w:r w:rsidR="00566C22" w:rsidRPr="00780270">` + props + content + `</w:r>`
	}
	field := run(`<w:fldChar w:fldCharType="begin"/>`) + run(`<w:instrText xml:space="preserve"> PAGE </w:instrText>`) + run(`<w:fldChar w:fldCharType="separate"/>`) + `<w:r w:rsidR="002D103B">` + resultProps + `<w:t>999 stale</w:t></w:r>` + run(`<w:fldChar w:fldCharType="end"/>`)
	for _, strict := range []bool{false, true} {
		for _, variant := range []string{"qualified", "rtl", "bad-revision", "unknown-run", "bad-size", "bad-proof", "duplicate-proof", "unknown-property", "unknown-text"} {
			markup := field
			switch variant {
			case "rtl":
				markup = strings.Replace(markup, resultProps, strings.Replace(resultProps, `<w:noProof/>`, `<w:noProof/><w:rtl/>`, 1), 1)
			case "bad-revision":
				markup = strings.Replace(markup, `w:rsidR="00566C22"`, `w:rsidR="invalid"`, 1)
			case "unknown-run":
				markup = strings.Replace(markup, `<w:r `, `<w:r extra="1" `, 1)
			case "bad-size":
				markup = strings.ReplaceAll(markup, `w:val="40"`, `w:val="bad"`)
			case "bad-proof":
				markup = strings.Replace(markup, `<w:noProof/>`, `<w:noProof w:val="bad"/>`, 1)
			case "duplicate-proof":
				markup = strings.Replace(markup, `<w:noProof/>`, `<w:noProof/><w:noProof/>`, 1)
			case "unknown-property":
				markup = strings.Replace(markup, `<w:rPr>`, `<w:rPr><w:future/>`, 1)
			case "unknown-text":
				markup = strings.Replace(markup, `<w:t>999`, `<w:t extra="1">999`, 1)
			}
			styles := `<w:styles xmlns:w="` + wordMLTransitional + `"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="character" w:styleId="PageNumber"><w:name w:val="Page Number"/></w:style></w:styles>`
			parts := resolvedStylesTestParts(styles)
			parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:pPr><w:rPr><w:sz w:val="16"/><w:szCs w:val="40"/></w:rPr></w:pPr>` + markup + `</w:p><w:sectPr/></w:body></w:document>`
			if strict {
				for k, v := range parts {
					parts[k] = strings.ReplaceAll(strings.ReplaceAll(v, wordMLTransitional, wordMLStrict), relBaseTransitional, relBaseStrict)
				}
			}
			data := buildNativeDOCX(t, nativeEntries(parts))
			before := string(data)
			doc, err := ExtractNativeDocumentV1(data)
			if err != nil {
				t.Fatal(err)
			}
			paragraph := doc.Body.Blocks[0].Paragraph
			count := 0
			for _, r := range paragraph.Runs {
				if r.PageField != "" {
					count++
					if r.Text == nil || *r.Text != "" || r.Properties == nil || r.Properties.CharacterStyleID == nil || *r.Properties.CharacterStyleID != "PageNumber" || r.Properties.FontSizeHalfPoint == nil || *r.Properties.FontSizeHalfPoint != 16 {
						t.Fatalf("result authority changed: %#v", r)
					}
				}
			}
			want := 0
			if variant == "qualified" || variant == "rtl" {
				want = 1
			}
			if count != want {
				t.Fatalf("strict=%v %s fields=%d", strict, variant, count)
			}
			if paragraph.EditPolicy.Mode != "read-only" || string(data) != before {
				t.Fatal("source/mutation authority changed")
			}
			if want == 1 {
				layout, err := ResolveNativeDocumentLayoutV1(data)
				if err != nil {
					t.Fatal(err)
				}
				if hasResolutionDiagnostic(layout, "COMPLEX_SCRIPT_SIZE_PRESERVED") != (variant == "rtl") {
					t.Fatalf("script guard wrong: %s %#v", variant, layout.Diagnostics)
				}
			}
		}
	}
}
