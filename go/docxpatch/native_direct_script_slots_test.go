package docxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func TestNativeDirectScriptSlotsRemainContextQualified(t *testing.T) {
	base := `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="CJK Face"/><w:sz w:val="24"/><w:szCs w:val="40"/><w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="hi-IN"/>`
	for _, strict := range []bool{false, true} {
		for _, tc := range []struct {
			name, text, extra, props string
			safe, blocked            bool
		}{
			{"latin", "Hello", "", base, true, false},
			{"cjk", "A漢", "", base, true, true},
			{"rtl", "Hello", `<w:bidi/>`, base, true, true},
			{"bad-size", "Hello", "", strings.Replace(base, `w:val="40"`, `w:val="0"`, 1), false, false},
			{"duplicate-size", "Hello", "", base + `<w:szCs w:val="40"/>`, false, false},
			{"size-child", "Hello", "", strings.Replace(base, `<w:szCs w:val="40"/>`, `<w:szCs w:val="40"><w:b/></w:szCs>`, 1), false, false},
			{"bad-language", "Hello", "", strings.Replace(base, `zh-CN`, `bad value`, 1), false, false},
			{"unknown-language", "Hello", "", strings.Replace(base, `w:eastAsia="zh-CN"`, `w:eastAsia="zh-CN" w:unknown="x"`, 1), false, false},
		} {
			t.Run(tc.name+map[bool]string{false: "-transitional", true: "-strict"}[strict], func(t *testing.T) {
				parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `"/>`)
				parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:pPr>` + tc.extra + `<w:rPr>` + tc.props + `</w:rPr></w:pPr><w:r><w:rPr>` + tc.props + `</w:rPr><w:t>` + tc.text + `</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
				if strict {
					for k, v := range parts {
						parts[k] = strings.ReplaceAll(strings.ReplaceAll(v, wordMLTransitional, wordMLStrict), relBaseTransitional, relBaseStrict)
					}
				}
				data := buildNativeDOCX(t, nativeEntries(parts))
				before := bytes.Clone(data)
				doc, err := ExtractNativeDocumentV1(data)
				if err != nil {
					t.Fatal(err)
				}
				encoded, err := EncodeNativeDocumentV1(doc)
				if err != nil {
					t.Fatal(err)
				}
				if strings.Contains(string(encoded), `"code":"UNMODELED_PARAGRAPH_MARK_PROPERTIES"`) == tc.safe {
					t.Fatalf("mark source guard: %s", encoded)
				}
				if tc.safe && strings.Contains(string(encoded), `"code":"PARTIAL_RUN_PROPERTIES"`) {
					t.Fatalf("inactive source slot still refused: %s", encoded)
				}
				if strings.Contains(string(encoded), `"code":"PARTIAL_PARAGRAPH_PROPERTIES"`) == tc.safe {
					t.Fatalf("paragraph mark confused preservation with invalidity: %s", encoded)
				}
				if doc.Body.Blocks[0].Paragraph.EditPolicy.Mode != "read-only" || !bytes.Equal(before, data) {
					t.Fatal("source authority changed")
				}
				layout, err := ResolveNativeDocumentLayoutV1(data)
				if err != nil {
					t.Fatal(err)
				}
				if tc.safe && hasResolutionDiagnostic(layout, "SCRIPT_FONT_PRESERVED") != tc.blocked {
					t.Fatalf("script context lost: %#v", layout.Diagnostics)
				}
				if tc.safe && hasResolutionDiagnostic(layout, "COMPLEX_SCRIPT_SIZE_PRESERVED") != tc.blocked {
					t.Fatalf("size context lost: %#v", layout.Diagnostics)
				}
			})
		}
	}
}
