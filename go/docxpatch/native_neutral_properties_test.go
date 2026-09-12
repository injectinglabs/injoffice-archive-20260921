package docxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func TestNativeNeutralSourceProperties(t *testing.T) {
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		for _, name := range []string{"noProof", "autoSpaceDE", "autoSpaceDN"} {
			for _, value := range []string{"0", "false", "off", "1", "true", "on", "bad", ""} {
				attr := ` w:val="` + value + `"`
				if value == "" {
					attr = ""
				}
				markup := `<w:` + name + attr + `/>`
				want := name == "noProof" && value != "bad" || name != "noProof" && (value == "0" || value == "false" || value == "off")
				for _, variant := range []string{"plain", "duplicate", "attribute", "text", "nested", "owner"} {
					body, owner := markup, ""
					switch variant {
					case "duplicate":
						body += markup
					case "attribute":
						body = strings.Replace(markup, "/>", ` extra="x"/>`, 1)
					case "text":
						body = strings.Replace(markup, "/>", `>bad</w:`+name+`>`, 1)
					case "nested":
						body = strings.Replace(markup, "/>", `><w:b/></w:`+name+`>`, 1)
					case "owner":
						owner = ` extra="x"`
					}
					root, err := parseNativeXML("test.xml", []byte(`<w:rPr xmlns:w="`+ns+`"`+owner+`>`+body+`</w:rPr>`))
					if err != nil {
						t.Fatal(err)
					}
					if got := nativeNeutralSourceProperty(root.Children[0], root, ns); got != (want && variant == "plain") {
						t.Fatalf("%s %s %s: qualified=%v", name, value, variant, got)
					}
				}
			}
		}
	}
}

func TestNativeNeutralPropertiesPreserveMutationBoundary(t *testing.T) {
	for _, name := range []string{"noProof", "autoSpaceDE", "autoSpaceDN"} {
		for _, malformed := range []bool{false, true} {
			property := `<w:` + name + ` w:val="0"/>`
			if malformed {
				property = strings.Replace(property, `"0"`, `"invalid"`, 1)
			}
			pPr, rPr := property, ""
			if name == "noProof" {
				pPr, rPr = "", property
			}
			parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `"/>`)
			parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:pPr>` + pPr + `</w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="22"/>` + rPr + `</w:rPr><w:t>Proof</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
			data := buildNativeDOCX(t, nativeEntries(parts))
			before := bytes.Clone(data)
			doc, err := ExtractNativeDocumentV1(data)
			if err != nil {
				t.Fatal(err)
			}
			resolved, err := ResolveNativeDocumentLayoutV1(data)
			if err != nil {
				t.Fatal(err)
			}
			code := "UNMODELED_PARAGRAPH_PROPERTY"
			if name == "noProof" {
				code = "UNMODELED_RUN_PROPERTY"
			}
			if hasResolutionDiagnostic(resolved, code) != malformed {
				t.Fatalf("%s malformed=%v diagnostics=%#v", name, malformed, resolved.Diagnostics)
			}
			paragraph := doc.Body.Blocks[0].Paragraph
			if paragraph.EditPolicy.Mode != "read-only" || len(paragraph.EditPolicy.AllowedOperations) != 0 {
				t.Fatalf("preserved source became mutable: %#v", paragraph.EditPolicy)
			}
			if !bytes.Equal(data, before) {
				t.Fatal("source bytes changed")
			}
		}
	}
}
