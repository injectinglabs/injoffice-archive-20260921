package docxpatch

import (
	"strings"
	"testing"
)

// ECMA-376 17.7.2: a style reference naming a style the package does not define
// is ignored. Nothing is dropped, so it is disclosed under its own code and the
// consumer cascades from the document defaults, as Word does.
// `floating-table-section-columns.docx` and `StyleRef-DE.docx` in hard-v2 both
// reference a paragraph style called `plain` that their styles part never
// defines.
func TestResolveNativeDocumentLayoutV1UndefinedStyleReference(t *testing.T) {
	for _, testCase := range []struct {
		name, styles, body, want string
	}{
		{
			"undefined paragraph style",
			`<w:style w:type="paragraph" w:styleId="Present"><w:pPr><w:jc w:val="right"/></w:pPr></w:style>`,
			`<w:p><w:pPr><w:pStyle w:val="plain"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
			"UNDEFINED_STYLE_REFERENCE",
		},
		{
			"undefined character style",
			`<w:style w:type="paragraph" w:styleId="Present"><w:pPr><w:jc w:val="right"/></w:pPr></w:style>`,
			`<w:p><w:r><w:rPr><w:rStyle w:val="plain"/></w:rPr><w:t>x</w:t></w:r></w:p>`,
			"UNDEFINED_STYLE_REFERENCE",
		},
		{
			// A style that does exist but whose basedOn ancestor does not: the
			// chain lost a layer, so it keeps the code it has always had.
			"missing basedOn ancestor",
			`<w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Absent"/><w:pPr><w:jc w:val="right"/></w:pPr></w:style>`,
			`<w:p><w:pPr><w:pStyle w:val="Child"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>`,
			"MISSING_STYLE_REFERENCE",
		},
		{
			// A table style head is decided by the table's own diagnostic, so
			// its pairing is unchanged.
			"undefined table style",
			`<w:style w:type="paragraph" w:styleId="Present"><w:pPr><w:jc w:val="right"/></w:pPr></w:style>`,
			`<w:tbl><w:tblPr><w:tblStyle w:val="plain"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="4000"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>`,
			"MISSING_STYLE_REFERENCE",
		},
	} {
		t.Run(testCase.name, func(t *testing.T) {
			parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `">` + testCase.styles + `</w:styles>`)
			parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body>` + testCase.body + `<w:sectPr/></w:body></w:document>`
			resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			other := map[string]string{"UNDEFINED_STYLE_REFERENCE": "MISSING_STYLE_REFERENCE", "MISSING_STYLE_REFERENCE": "UNDEFINED_STYLE_REFERENCE"}[testCase.want]
			if !hasResolutionDiagnostic(resolved, testCase.want) || hasResolutionDiagnostic(resolved, other) {
				t.Fatalf("want %s and not %s, got %#v", testCase.want, other, resolved.Diagnostics)
			}
			for _, diagnostic := range resolved.Diagnostics {
				if diagnostic.Code != "UNDEFINED_STYLE_REFERENCE" {
					continue
				}
				if diagnostic.Severity != "unsupported" || diagnostic.Preservation != "preserve-verbatim" || diagnostic.PartName == nil || *diagnostic.PartName != "word/styles.xml" || diagnostic.Path != nil || !strings.Contains(diagnostic.Message, "no style layer was dropped") {
					t.Fatalf("undefined style reference is not disclosed at the styles part: %#v", diagnostic)
				}
			}
		})
	}
}
