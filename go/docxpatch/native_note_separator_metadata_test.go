package docxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func TestNoteSeparatorRevisionMetadataKeepsExactContentGuards(t *testing.T) {
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		for _, role := range []string{"separator", "continuation-separator"} {
			instruction := "separator"
			if role == "continuation-separator" {
				instruction = "continuationSeparator"
			}
			for _, tc := range []struct {
				name, attrs, extra, properties, instruction string
				valid                                       bool
			}{
				{name: "exact revision", attrs: `w:rsidR="00164462" w:rsidRDefault="00164462"`, valid: true},
				{name: "paragraph revision and spacing", attrs: `w:rsidR="00916EE9" w:rsidRDefault="00916EE9" w:rsidP="002D662D"`, properties: `<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>`, valid: true},
				{name: "malformed paragraph revision", attrs: `w:rsidP="wrong"`},
				{name: "lowercase hex", attrs: `w:rsidR="aabbccdd"`, valid: true},
				{name: "no metadata", valid: true},
				{name: "nonhex", attrs: `w:rsidR="nothexid"`},
				{name: "short", attrs: `w:rsidR="001"`},
				{name: "unknown attr", attrs: `w:rsidR="00164462" extra="1"`},
				{name: "foreign revision", attrs: `xmlns:x="urn:foreign" x:rsidR="00164462"`},
				{name: "visible text", extra: `<w:r><w:t>visible</w:t></w:r>`},
				{name: "reference", extra: `<w:r><w:footnoteReference w:id="1"/></w:r>`},
				{name: "wrong instruction", instruction: "footnoteRef"},
				{name: "nonempty instruction", instruction: instruction + `>text</w:` + instruction + `><w:` + instruction},
				{name: "spacing", properties: `<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>`, valid: true},
				{name: "spacing owner attrs", properties: `<w:pPr extra="1"><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>`},
				{name: "spacing owner text", properties: `<w:pPr>bad<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>`},
			} {
				t.Run(ns+role+tc.name, func(t *testing.T) {
					leaf := instruction
					if tc.instruction != "" {
						leaf = tc.instruction
					}
					xml := `<w:footnote xmlns:w="` + ns + `"><w:p ` + tc.attrs + `>` + tc.properties + `<w:r><w:` + leaf + `/></w:r>` + tc.extra + `</w:p></w:footnote>`
					node, err := parseNativeXML("word/footnotes.xml", []byte(xml))
					if err != nil {
						t.Fatal(err)
					}
					if got := nativeExactNoteSentinel(node, ns, role); got != tc.valid {
						t.Fatalf("qualification=%v want%v", got, tc.valid)
					}
				})
			}
		}
	}
}

func TestExtractNoteSeparatorRevisionMetadataPreservesSource(t *testing.T) {
	parts := nativeNotePagePaintParts()
	for _, name := range []string{"Custom/Notes/Foot.XML", "Custom/Notes/End.XML"} {
		parts[name] = strings.ReplaceAll(parts[name], `<w:t>---</w:t>`, `<w:separator/>`)
		parts[name] = strings.ReplaceAll(parts[name], `<w:t>continued</w:t>`, `<w:continuationSeparator/>`)
		parts[name] = strings.ReplaceAll(parts[name], `<w:p>`, `<w:p w:rsidR="00164462" w:rsidRDefault="00164462">`)
	}
	data := buildNativeDOCX(t, nativeEntries(parts))
	before := bytes.Clone(data)
	doc, err := ExtractNativeDocumentV1(data)
	if err != nil {
		t.Fatal(err)
	}
	for _, d := range doc.Unsupported {
		if d.Code == "UNMODELED_NOTE_MARKUP" {
			t.Fatalf("valid separator refused: %#v", d)
		}
	}
	if len(doc.Notes) != 6 || !bytes.Equal(data, before) {
		t.Fatal("note content/source changed")
	}
}

// ECMA-376 17.11.14 makes a reserved separator story an ordinary story: its
// first paragraph carries the instruction, and Word paints whatever paragraphs
// follow it above the notes. Word writes a trailing empty paragraph itself, and
// an author can put visible text in one. The instruction paragraph still has to
// be exact, and a block that is not a paragraph keeps refusing.
func TestNoteSeparatorStoryCarriesFurtherParagraphs(t *testing.T) {
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		for _, role := range []string{"separator", "continuation-separator"} {
			instruction := "separator"
			if role == "continuation-separator" {
				instruction = "continuationSeparator"
			}
			for _, tc := range []struct {
				name, trailing string
				valid          bool
			}{
				{name: "instruction alone", valid: true},
				{name: "trailing empty paragraph", trailing: `<w:p/>`, valid: true},
				{name: "trailing text paragraph", trailing: `<w:p><w:r><w:t>Text in footnote separator</w:t></w:r></w:p>`, valid: true},
				{name: "several trailing paragraphs", trailing: `<w:p><w:r><w:t>one</w:t></w:r></w:p><w:p/>`, valid: true},
				{name: "trailing table", trailing: `<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`},
				{name: "trailing foreign block", trailing: `<w:sdt/>`},
				{name: "leading paragraph before the instruction", trailing: ``},
			} {
				t.Run(ns+role+tc.name, func(t *testing.T) {
					body := `<w:p><w:r><w:` + instruction + `/></w:r></w:p>` + tc.trailing
					if tc.name == "leading paragraph before the instruction" {
						body = `<w:p><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:` + instruction + `/></w:r></w:p>`
					}
					node, err := parseNativeXML("word/footnotes.xml", []byte(`<w:footnote xmlns:w="`+ns+`">`+body+`</w:footnote>`))
					if err != nil {
						t.Fatal(err)
					}
					if got := nativeExactNoteSentinel(node, ns, role); got != tc.valid {
						t.Fatalf("qualification=%v want %v", got, tc.valid)
					}
				})
			}
		}
	}
}
