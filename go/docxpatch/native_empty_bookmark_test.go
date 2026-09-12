package docxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func TestNativeEmptyBookmarkSourceQualification(t *testing.T) {
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		for _, variant := range []string{"exact", "name", "end", "negative", "leading-zero", "overflow", "unknown", "nested", "range", "foreign"} {
			markup := `<w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/>`
			switch variant {
			case "name":
				markup = strings.Replace(markup, ` w:name="_GoBack"`, "", 1)
			case "end":
				markup = strings.Replace(markup, `bookmarkEnd w:id="0"`, `bookmarkEnd w:id="1"`, 1)
			case "negative":
				markup = strings.ReplaceAll(markup, `w:id="0"`, `w:id="-1"`)
			case "leading-zero":
				markup = strings.ReplaceAll(markup, `w:id="0"`, `w:id="00"`)
			case "overflow":
				markup = strings.ReplaceAll(markup, `w:id="0"`, `w:id="99999999999999999"`)
			case "unknown":
				markup = strings.Replace(markup, `w:name=`, `extra="x" w:name=`, 1)
			case "nested":
				markup = strings.Replace(markup, `/>`, `><w:r/></w:bookmarkStart>`, 1)
			case "range":
				markup = strings.Replace(markup, `<w:bookmarkEnd`, `<w:r><w:t>visible</w:t></w:r><w:bookmarkEnd`, 1)
			case "foreign":
				markup = strings.Replace(markup, `<w:bookmarkEnd`, `<bookmarkEnd`, 1)
			}
			parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + ns + `"/>`)
			parts["word/document.xml"] = `<w:document xmlns:w="` + ns + `"><w:body><w:p><w:r><w:t>Text</w:t></w:r>` + markup + `</w:p><w:sectPr/></w:body></w:document>`
			if ns == wordMLStrict {
				for k, v := range parts {
					parts[k] = strings.ReplaceAll(v, relBaseTransitional, relBaseStrict)
				}
			}
			data := buildNativeDOCX(t, nativeEntries(parts))
			before := bytes.Clone(data)
			doc, err := ExtractNativeDocumentV1(data)
			if err != nil {
				t.Fatal(err)
			}
			found := false
			for _, d := range doc.Unsupported {
				if d.Code == "UNMODELED_PARAGRAPH_CONTENT" {
					found = true
				}
			}
			if found != (variant != "exact") {
				t.Fatalf("%s %s diagnostics=%#v", ns, variant, doc.Unsupported)
			}
			if !bytes.Equal(before, data) || doc.Body.Blocks[0].Paragraph.EditPolicy.Mode != "read-only" {
				t.Fatal("source or editing authority changed")
			}
		}
	}
}
