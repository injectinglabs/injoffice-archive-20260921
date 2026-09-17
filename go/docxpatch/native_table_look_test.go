package docxpatch

import (
	"fmt"
	"strings"
	"testing"
)

func TestInactiveTableLookRequiresCompleteUnconditionalStyleChain(t *testing.T) {
	base := `<w:style w:type="table" w:styleId="Base"><w:name w:val="Base"/></w:style>`
	child := `<w:style w:type="table" w:styleId="Child"><w:basedOn w:val="Base"/></w:style>`
	// An unrelated conditional table style keeps every chain guard meaningful:
	// the package-wide proof cannot apply, so each case exercises the walk.
	conditional := `<w:style w:type="table" w:styleId="Fancy"><w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr></w:style>`
	deep := `<w:style w:type="table" w:styleId="Child"><w:basedOn w:val="S0"/></w:style>`
	for i := 0; i < 64; i++ {
		deep += fmt.Sprintf(`<w:style w:type="table" w:styleId="S%d"><w:basedOn w:val="S%d"/></w:style>`, i, i+1)
	}
	deep += `<w:style w:type="table" w:styleId="S64"/>`
	switches := ` w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"`
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		for _, test := range []struct {
			name, styles, look string
			// props replaces the default table-style reference when set.
			props string
			valid bool
		}{
			{name: "inherited", styles: base + child, look: `<w:tblLook w:val="04A0"/>`, valid: true},
			{name: "zero", styles: base + child, look: `<w:tblLook w:val="0000"/>`, valid: true},
			{name: "relocated styles", styles: base + child, look: `<w:tblLook w:val="04A0"/>`, valid: true},
			{name: "switch attributes", styles: base + child, look: `<w:tblLook w:val="04A0"` + switches + `/>`, valid: true},
			{name: "switch attributes disagree", styles: base + child, look: `<w:tblLook w:val="04A0" w:firstRow="0"/>`, valid: false},
			{name: "switch attribute unset bit", styles: base + child, look: `<w:tblLook w:val="04A0" w:lastRow="1"/>`, valid: false},
			{name: "switch attribute malformed", styles: base + child, look: `<w:tblLook w:val="04A0" w:firstRow="maybe"/>`, valid: false},
			{name: "over depth", styles: deep, look: `<w:tblLook w:val="04A0"/>`, valid: false},
			{name: "unknown owner", styles: base + strings.Replace(child, `w:styleId="Child"`, `w:styleId="Child" w:extra="1"`, 1), look: `<w:tblLook w:val="04A0"/>`, valid: false},
			{name: "unknown layer", styles: base + strings.Replace(child, `</w:style>`, `<w:unknown/></w:style>`, 1), look: `<w:tblLook w:val="04A0"/>`, valid: false},
			{name: "invalid default", styles: base + strings.Replace(child, `w:styleId="Child"`, `w:styleId="Child" w:default="maybe"`, 1), look: `<w:tblLook w:val="04A0"/>`, valid: false},
			{name: "conditional ancestor", styles: strings.Replace(base, `<w:name w:val="Base"/>`, `<w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr>`, 1) + child, look: `<w:tblLook w:val="04A0"/>`, valid: false},
			{name: "missing ancestor", styles: child, look: `<w:tblLook w:val="04A0"/>`, valid: false},
			{name: "cycle", styles: strings.Replace(base, `<w:name w:val="Base"/>`, `<w:basedOn w:val="Child"/>`, 1) + child, look: `<w:tblLook w:val="04A0"/>`, valid: false},
			{name: "duplicate style", styles: base + child + child, look: `<w:tblLook w:val="04A0"/>`, valid: false},
			{name: "duplicate parent", styles: base + strings.Replace(child, `</w:style>`, `<w:basedOn w:val="Base"/></w:style>`, 1), look: `<w:tblLook w:val="04A0"/>`, valid: false},
			{name: "malformed", styles: base + child, look: `<w:tblLook w:val="oops"/>`, valid: false},
			{name: "reserved bits", styles: base + child, look: `<w:tblLook w:val="FFFF"/>`, valid: false},
			{name: "foreign attr", styles: base + child, look: `<w:tblLook xmlns:x="urn:foreign" x:val="04A0"/>`, valid: false},
			{name: "unknown child", styles: base + child, look: `<w:tblLook w:val="04A0"><w:unknown/></w:tblLook>`, valid: false},
			{name: "duplicate look", styles: base + child, look: `<w:tblLook w:val="04A0"/><w:tblLook w:val="04A0"/>`, valid: false},
			// A conditional style exists, so an applied style that does not
			// resolve leaves the switches unproven.
			{name: "undefined style with conditional package", styles: base + child, look: `<w:tblLook w:val="04A0"/>`, props: `<w:tblStyle w:val="Absent"/>`, valid: false},
			{name: "absent style reference with conditional package", styles: base + child, look: `<w:tblLook w:val="04A0"/>`, props: ` `, valid: false},
		} {
			t.Run(test.name+ns, func(t *testing.T) {
				props := test.props
				if props == "" {
					props = `<w:tblStyle w:val="Child"/>`
				}
				parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + ns + `">` + conditional + test.styles + `</w:styles>`)
				parts["word/document.xml"] = `<w:document xmlns:w="` + ns + `"><w:body><w:tbl><w:tblPr>` + props + test.look + `</w:tblPr><w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:p/></w:body></w:document>`
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

// A package whose style definitions declare no conditional table formatting
// offers nothing for a tblLook switch to select, however the applied style
// resolves. That proof is independent of the style chain, so it also covers an
// absent w:tblStyle and a reference no style defines.
func TestInactiveTableLookAcceptsPackageWithoutConditionalTableFormat(t *testing.T) {
	plain := `<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style>`
	conditional := `<w:style w:type="table" w:styleId="Fancy"><w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr></w:style>`
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		for _, test := range []struct {
			name, root, styles, props string
			valid                     bool
		}{
			{name: "absent reference", styles: plain, props: ``, valid: true},
			{name: "undefined reference", styles: plain, props: `<w:tblStyle w:val="Tabellengitternetz"/>`, valid: true},
			{name: "no table style at all", styles: `<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/></w:style>`, props: ``, valid: true},
			{name: "ignorable root", root: ` xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"`, styles: plain, props: ``, valid: true},
			{name: "conditional elsewhere absent reference", styles: plain + conditional, props: ``, valid: false},
			{name: "conditional elsewhere undefined reference", styles: plain + conditional, props: `<w:tblStyle w:val="Tabellengitternetz"/>`, valid: false},
			// A conditional format inside foreign markup still counts.
			{name: "conditional under foreign wrapper", styles: plain + `<w:style w:type="table" w:styleId="Alt"><w:foreign xmlns:w="urn:other"><w:tblStylePr/></w:foreign></w:style>`, props: ``, valid: false},
		} {
			t.Run(test.name+ns, func(t *testing.T) {
				parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + ns + `"` + test.root + `>` + test.styles + `</w:styles>`)
				parts["word/document.xml"] = `<w:document xmlns:w="` + ns + `"><w:body><w:tbl><w:tblPr>` + test.props + `<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:p/></w:body></w:document>`
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
					if d.Code == "UNMODELED_TABLE_PROPERTY" && strings.Contains(d.Message, "Table look") {
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
