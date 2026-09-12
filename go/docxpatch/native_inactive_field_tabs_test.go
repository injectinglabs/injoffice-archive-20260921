package docxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func TestInactiveTabsRequireExtractorQualifiedPageFieldConsumers(t *testing.T) {
	simple := func(instruction, result string) string {
		return `<w:fldSimple w:instr="` + instruction + `"><w:r>` + result + `</w:r></w:fldSimple>`
	}
	flat := func(instruction, result string) string {
		return `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>` + instruction + `</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r>` + result + `</w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`
	}
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		for _, tc := range []struct {
			name, content, extra string
			valid                bool
		}{
			{name: "simple PAGE", content: simple("PAGE", `<w:t>1</w:t>`), valid: true},
			{name: "simple NUMPAGES", content: simple("NUMPAGES", `<w:t>1</w:t>`), valid: true},
			{name: "flat PAGE", content: flat("PAGE", `<w:t>1</w:t>`), valid: true},
			{name: "flat NUMPAGES", content: flat("NUMPAGES", `<w:t>1</w:t>`), valid: true},
			{name: "multiple qualified", content: simple("PAGE", `<w:t>1</w:t>`) + flat("NUMPAGES", `<w:t>3</w:t>`), valid: true},
			{name: "generated replaces cached tab", content: simple("PAGE", "<w:t>old\tcache</w:t>"), valid: true},
			{name: "unknown", content: simple("DATE", `<w:t>1</w:t>`)},
			{name: "mixed unknown", content: flat("PAGE", `<w:t>1</w:t>`) + simple("DATE", `<w:t>1</w:t>`)},
			{name: "outside tab", content: flat("PAGE", `<w:t>1</w:t>`) + `<w:r><w:tab/></w:r>`},
			{name: "before tab", content: `<w:r><w:tab/></w:r>` + flat("PAGE", `<w:t>1</w:t>`)},
			{name: "simple result tab", content: simple("PAGE", `<w:tab/><w:t>1</w:t>`)},
			{name: "flat result tab", content: flat("PAGE", `<w:tab/><w:t>1</w:t>`)},
			{name: "between metadata tab", content: strings.Replace(flat("PAGE", `<w:t>1</w:t>`), `<w:instrText>`, `<w:tab/><w:instrText>`, 1)},
			{name: "inserted sequence tab", content: strings.Replace(flat("PAGE", `<w:t>1</w:t>`), `<w:r><w:instrText>`, `<w:r><w:tab/></w:r><w:r><w:instrText>`, 1)},
			{name: "missing end", content: strings.Replace(flat("PAGE", `<w:t>1</w:t>`), `<w:r><w:fldChar w:fldCharType="end"/></w:r>`, "", 1)},
			{name: "hyperlink", content: `<w:hyperlink>` + simple("PAGE", `<w:t>1</w:t>`) + `</w:hyperlink>`},
			{name: "numbering", content: simple("PAGE", `<w:t>1</w:t>`), extra: `<w:numPr><w:numId w:val="7"/></w:numPr>`},
		} {
			t.Run(ns+tc.name, func(t *testing.T) {
				parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `"><w:style w:type="paragraph" w:styleId="Header"><w:pPr><w:tabs><w:tab w:val="right" w:pos="9406"/></w:tabs></w:pPr></w:style></w:styles>`)
				parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:pPr><w:pStyle w:val="Header"/>` + tc.extra + `</w:pPr>` + tc.content + `</w:p></w:body></w:document>`
				if ns == wordMLStrict {
					for k, v := range parts {
						parts[k] = strings.ReplaceAll(strings.ReplaceAll(v, wordMLTransitional, wordMLStrict), relBaseTransitional, relBaseStrict)
					}
				}
				data := buildNativeDOCX(t, nativeEntries(parts))
				before := bytes.Clone(data)
				layout, err := ResolveNativeDocumentLayoutV1(data)
				if err != nil {
					t.Fatal(err)
				}
				blocked := false
				for _, d := range layout.Diagnostics {
					if d.Code == "UNMODELED_PARAGRAPH_PROPERTY" {
						blocked = true
					}
				}
				if blocked == tc.valid {
					t.Fatalf("blocked=%v valid=%v diagnostics%#v", blocked, tc.valid, layout.Diagnostics)
				}
				if !bytes.Equal(data, before) {
					t.Fatal("source changed")
				}
			})
		}
	}
}
