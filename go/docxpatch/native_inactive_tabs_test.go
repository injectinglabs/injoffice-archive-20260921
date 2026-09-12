package docxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func TestNativeUnusedStyleTabs(t *testing.T) {
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		for _, variant := range []string{"plain", "tab", "ptab", "literal", "field", "numbering", "bar", "leader", "duplicate-position", "duplicate-tabs", "missing-position", "negative", "foreign", "unknown", "nested"} {
			tabs := `<w:tabs><w:tab w:val="center" w:pos="4703"/><w:tab w:val="right" w:pos="9406"/></w:tabs>`
			content := `<w:r><w:t>Header</w:t></w:r>`
			extra := ""
			switch variant {
			case "tab":
				content = `<w:r><w:tab/><w:t>Header</w:t></w:r>`
			case "ptab":
				content = `<w:r><w:ptab w:alignment="right"/><w:t>Header</w:t></w:r>`
			case "literal":
				content = "<w:r><w:t>Head\ter</w:t></w:r>"
			case "field":
				content = `<w:fldSimple w:instr="DATE"><w:r><w:t>1</w:t></w:r></w:fldSimple>`
			case "numbering":
				extra = `<w:numPr><w:numId w:val="7"/></w:numPr>`
			case "bar":
				tabs = strings.Replace(tabs, `w:val="center"`, `w:val="bar"`, 1)
			case "leader":
				tabs = strings.Replace(tabs, `w:pos="4703"`, `w:pos="4703" w:leader="dot"`, 1)
			case "duplicate-position":
				tabs = strings.Replace(tabs, "9406", "4703", 1)
			case "duplicate-tabs":
				tabs += tabs
			case "missing-position":
				tabs = strings.Replace(tabs, ` w:pos="4703"`, "", 1)
			case "negative":
				tabs = strings.Replace(tabs, "4703", "-1", 1)
			case "foreign":
				tabs = strings.Replace(tabs, "<w:tab ", "<tab ", 1)
			case "unknown":
				tabs = strings.Replace(tabs, `w:pos="4703"`, `w:pos="4703" extra="1"`, 1)
			case "nested":
				tabs = strings.Replace(tabs, "/>", "><w:x/></w:tab>", 1)
			}
			parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + ns + `"><w:style w:type="paragraph" w:styleId="Header"><w:pPr>` + tabs + extra + `</w:pPr></w:style></w:styles>`)
			parts["word/document.xml"] = `<w:document xmlns:w="` + ns + `"><w:body><w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr>` + content + `</w:p></w:body></w:document>`
			if ns == wordMLStrict {
				for k, v := range parts {
					parts[k] = strings.ReplaceAll(strings.ReplaceAll(v, wordMLTransitional, wordMLStrict), relBaseTransitional, relBaseStrict)
				}
			}
			source := buildNativeDOCX(t, nativeEntries(parts))
			before := bytes.Clone(source)
			layout, err := ResolveNativeDocumentLayoutV1(source)
			if err != nil {
				if variant == "foreign" {
					continue
				}
				t.Fatal(err)
			}
			found := false
			for _, d := range layout.Diagnostics {
				if d.Code == "UNMODELED_PARAGRAPH_PROPERTY" || d.Code == "DUPLICATE_PARAGRAPH_PROPERTY" {
					found = true
				}
			}
			if found != (variant != "plain") {
				t.Fatalf("%s %s diagnostics %#v", ns, variant, layout.Diagnostics)
			}
			if !bytes.Equal(source, before) {
				t.Fatal("source changed")
			}
		}
	}
}

func TestNativeUnusedTabsConsumerIsolation(t *testing.T) {
	parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `"><w:style w:type="paragraph" w:styleId="Base"><w:pPr><w:tabs><w:tab w:val="right" w:pos="9406"/></w:tabs></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Base"/></w:style></w:styles>`)
	parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:pPr><w:pStyle w:val="Child"/></w:pPr><w:r><w:t>plain</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Child"/></w:pPr><w:r><w:tab/><w:t>active</w:t></w:r></w:p></w:body></w:document>`
	layout, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	found := false
	for _, d := range layout.Diagnostics {
		if d.Code == "UNMODELED_PARAGRAPH_PROPERTY" {
			found = true
			if d.ScopeID != layout.Paragraphs[1].ParagraphID {
				t.Fatal("inactive consumer polluted")
			}
		}
	}
	if !found {
		t.Fatal("active consumer lost refusal")
	}
}
