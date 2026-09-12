package docxpatch

import (
	"fmt"
	"strings"
	"testing"
)

func TestInactiveTableLookRequiresCompleteUnconditionalStyleChain(t *testing.T) {
	base := `<w:style w:type="table" w:styleId="Base"><w:name w:val="Base"/></w:style>`
	child := `<w:style w:type="table" w:styleId="Child"><w:basedOn w:val="Base"/></w:style>`
	deep := `<w:style w:type="table" w:styleId="Child"><w:basedOn w:val="S0"/></w:style>`
	for i := 0; i < 64; i++ {
		deep += fmt.Sprintf(`<w:style w:type="table" w:styleId="S%d"><w:basedOn w:val="S%d"/></w:style>`, i, i+1)
	}
	deep += `<w:style w:type="table" w:styleId="S64"/>`
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		for _, test := range []struct {
			name, styles, look string
			valid              bool
		}{
			{"inherited", base + child, `<w:tblLook w:val="04A0"/>`, true},
			{"zero", base + child, `<w:tblLook w:val="0000"/>`, true},
			{"relocated styles", base + child, `<w:tblLook w:val="04A0"/>`, true},
			{"over depth", deep, `<w:tblLook w:val="04A0"/>`, false},
			{"unknown owner", base + strings.Replace(child, `w:styleId="Child"`, `w:styleId="Child" w:extra="1"`, 1), `<w:tblLook w:val="04A0"/>`, false},
			{"unknown layer", base + strings.Replace(child, `</w:style>`, `<w:unknown/></w:style>`, 1), `<w:tblLook w:val="04A0"/>`, false},
			{"invalid default", base + strings.Replace(child, `w:styleId="Child"`, `w:styleId="Child" w:default="maybe"`, 1), `<w:tblLook w:val="04A0"/>`, false},
			{"conditional ancestor", strings.Replace(base, `<w:name w:val="Base"/>`, `<w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr>`, 1) + child, `<w:tblLook w:val="04A0"/>`, false},
			{"missing ancestor", child, `<w:tblLook w:val="04A0"/>`, false},
			{"cycle", strings.Replace(base, `<w:name w:val="Base"/>`, `<w:basedOn w:val="Child"/>`, 1) + child, `<w:tblLook w:val="04A0"/>`, false},
			{"duplicate style", base + child + child, `<w:tblLook w:val="04A0"/>`, false},
			{"duplicate parent", base + strings.Replace(child, `</w:style>`, `<w:basedOn w:val="Base"/></w:style>`, 1), `<w:tblLook w:val="04A0"/>`, false},
			{"malformed", base + child, `<w:tblLook w:val="oops"/>`, false},
			{"reserved bits", base + child, `<w:tblLook w:val="FFFF"/>`, false},
			{"foreign attr", base + child, `<w:tblLook xmlns:x="urn:foreign" x:val="04A0"/>`, false},
			{"unknown child", base + child, `<w:tblLook w:val="04A0"><w:unknown/></w:tblLook>`, false},
			{"duplicate look", base + child, `<w:tblLook w:val="04A0"/><w:tblLook w:val="04A0"/>`, false},
		} {
			t.Run(test.name+ns, func(t *testing.T) {
				parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + ns + `">` + test.styles + `</w:styles>`)
				parts["word/document.xml"] = `<w:document xmlns:w="` + ns + `"><w:body><w:tbl><w:tblPr><w:tblStyle w:val="Child"/>` + test.look + `</w:tblPr><w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:p/></w:body></w:document>`
				if test.name == "relocated styles" {
					parts["word/renamed-styles.xml"] = parts["word/styles.xml"]
					delete(parts, "word/styles.xml")
					parts["word/_rels/document.xml.rels"] = strings.ReplaceAll(parts["word/_rels/document.xml.rels"], "styles.xml", "renamed-styles.xml")
					parts["[Content_Types].xml"] = strings.ReplaceAll(parts["[Content_Types].xml"], "/word/styles.xml", "/word/renamed-styles.xml")
				}
				if ns == wordMLStrict {
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
				blocked := false
				for _, d := range doc.Unsupported {
					if d.Code == "UNMODELED_TABLE_PROPERTY" {
						blocked = true
					}
				}
				if blocked == test.valid {
					t.Fatalf("unexpected qualification: %#v", doc.Unsupported)
				}
				if doc.Body.Blocks[0].Table.EditPolicy.Mode != "read-only" {
					t.Fatal("look must not authorize mutation")
				}
				if string(data) != before {
					t.Fatal("source mutated")
				}
			})
		}
	}
}
