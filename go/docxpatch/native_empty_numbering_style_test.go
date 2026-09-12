package docxpatch

import (
	"strings"
	"testing"
)

func TestEmptyDefaultNumberingStyleQualification(t *testing.T) {
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		base := `<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>`
		for _, variant := range []string{"valid", "formatting", "link", "duplicate", "unknown", "badbool", "badpriority", "foreign", "owner", "rootowner", "text"} {
			markup := base
			switch variant {
			case "formatting":
				markup = strings.Replace(markup, "</w:style>", "<w:pPr><w:numPr><w:numId w:val=\"1\"/></w:numPr></w:pPr></w:style>", 1)
			case "link":
				markup = strings.Replace(markup, "</w:style>", "<w:basedOn w:val=\"Other\"/></w:style>", 1)
			case "duplicate":
				markup += base
			case "unknown":
				markup = strings.Replace(markup, "<w:semiHidden/>", "<w:unknown/>", 1)
			case "badbool":
				markup = strings.Replace(markup, `w:default="1"`, `w:default="bad"`, 1)
			case "badpriority":
				markup = strings.Replace(markup, `w:val="99"`, `w:val="100"`, 1)
			case "foreign":
				markup = strings.Replace(markup, "<w:semiHidden/>", `<x:semiHidden xmlns:x="urn:foreign"/>`, 1)
			case "owner":
				markup = strings.Replace(markup, "<w:style ", `<w:style extra="1" `, 1)
			case "text":
				markup = strings.Replace(markup, "</w:style>", "bad</w:style>", 1)
			}
			rootAttrs := ""
			if variant == "rootowner" {
				rootAttrs = ` extra="1"`
			}
			root, err := parseNativeXML("styles.xml", []byte(`<w:styles xmlns:w="`+ns+`"`+rootAttrs+`>`+markup+`</w:styles>`))
			if err != nil {
				t.Fatal(err)
			}
			if got := nativeEmptyDefaultNumberingStyle(root.Children[0], root, ns); got != (variant == "valid") {
				t.Fatalf("%s: %v", variant, got)
			}
		}
	}
}

func TestEmptyDefaultNumberingStyleRetainsSourceDiagnostic(t *testing.T) {
	styles := `<w:styles xmlns:w="` + wordMLTransitional + `"><w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style></w:styles>`
	parts := resolvedStylesTestParts(styles)
	data := buildNativeDOCX(t, nativeEntries(parts))
	before := string(data)
	resolved, err := ResolveNativeDocumentLayoutV1(data)
	if err != nil {
		t.Fatal(err)
	}
	if !hasResolutionDiagnostic(resolved, "EMPTY_NUMBERING_STYLE_PRESERVED") || hasResolutionDiagnostic(resolved, "NUMBERING_STYLE_PRESERVED") {
		t.Fatalf("%#v", resolved.Diagnostics)
	}
	if string(data) != before {
		t.Fatal("source changed")
	}
}
