package docxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func TestNativeInactiveSectionProperties(t *testing.T) {
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		for _, tc := range []struct {
			xml string
			ok  bool
		}{
			{`<w:formProt w:val="0"/><w:noEndnote/>`, true},
			{`<w:formProt w:val="false"/><w:noEndnote w:val="0"/>`, true},
			{`<w:formProt/>`, false}, {`<w:formProt w:val="1"/>`, false},
			{`<w:formProt w:val="0" extra="x"/>`, false}, {`<w:formProt w:val="0">text</w:formProt>`, false},
			{`<w:formProt w:val="0"/><w:formProt w:val="0"/>`, false},
			{`<w:noEndnote/><w:noEndnote/>`, false}, {`<w:noEndnote w:val="bad"/>`, false},
			{`<w:noEndnote><w:b/></w:noEndnote>`, false},
		} {
			parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + ns + `"/>`)
			parts["word/document.xml"] = `<w:document xmlns:w="` + ns + `"><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:bottom="720" w:left="720" w:right="720" w:header="0" w:footer="0" w:gutter="0"/>` + tc.xml + `</w:sectPr></w:body></w:document>`
			if ns == wordMLStrict {
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
			blocked := false
			for _, d := range doc.Unsupported {
				if d.Capability == "sections" {
					blocked = true
				}
			}
			if blocked == tc.ok {
				t.Fatalf("%s: %#v", tc.xml, doc.Unsupported)
			}
			if !bytes.Equal(before, data) {
				t.Fatal("source changed")
			}
		}
	}
}

func TestNoContentEndnotesPackageProof(t *testing.T) {
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		relNS := relNSTransitional
		if ns == wordMLStrict {
			relNS = relNSStrict
		}
		fresh := func() *nativeExtractor {
			return &nativeExtractor{wordNS: ns, relNS: relNS, mainPart: "word/document.xml", pkg: &nativePackage{files: map[string][]byte{"word/document.xml": []byte(`<w:document xmlns:w="` + ns + `"><w:body/></w:document>`)}, contentTypes: map[string]string{}, rels: map[string][]nativeRelationship{}}}
		}
		addNotes := func(e *nativeExtractor) {
			e.pkg.files["word/endnotes.xml"] = []byte(`<w:endnotes xmlns:w="` + ns + `"><w:endnote w:id="-1" w:type="separator"><w:p w:rsidR="01234567"><w:r><w:separator/></w:r></w:p></w:endnote></w:endnotes>`)
			e.pkg.contentTypes["word/endnotes.xml"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
			e.pkg.rels[e.mainPart] = []nativeRelationship{{Type: relNS + "/endnotes", PartName: "word/endnotes.xml"}}
		}
		addSettings := func(e *nativeExtractor) {
			addNotes(e)
			e.pkg.files["word/settings.xml"] = []byte(`<w:settings xmlns:w="` + ns + `"><w:endnotePr><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr></w:settings>`)
			e.pkg.contentTypes["word/settings.xml"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
			e.pkg.rels[e.mainPart] = append(e.pkg.rels[e.mainPart], nativeRelationship{Type: relNS + "/settings", PartName: "word/settings.xml"})
		}
		for _, tc := range []struct {
			name   string
			change func(*nativeExtractor)
			ok     bool
		}{
			{"absent", func(e *nativeExtractor) {}, true}, {"reserved", addNotes, true},
			{"settings-reserved", addSettings, true},
			{"settings-content-id", func(e *nativeExtractor) {
				addSettings(e)
				e.pkg.files["word/settings.xml"] = bytes.Replace(e.pkg.files["word/settings.xml"], []byte(`w:id="0"`), []byte(`w:id="1"`), 1)
			}, false},
			{"settings-duplicate", func(e *nativeExtractor) {
				addSettings(e)
				e.pkg.files["word/settings.xml"] = bytes.Replace(e.pkg.files["word/settings.xml"], []byte(`w:id="0"`), []byte(`w:id="-1"`), 1)
			}, false},
			{"settings-orphan", func(e *nativeExtractor) { addSettings(e); e.pkg.rels[e.mainPart] = e.pkg.rels[e.mainPart][:1] }, false},
			{"content", func(e *nativeExtractor) {
				addNotes(e)
				e.pkg.files["word/endnotes.xml"] = []byte(`<w:endnotes xmlns:w="` + ns + `"><w:endnote w:id="1"><w:p/></w:endnote></w:endnotes>`)
			}, false},
			{"reference", func(e *nativeExtractor) {
				e.pkg.files["other.xml"] = []byte(`<w:endnoteReference xmlns:w="` + ns + `" w:id="1"/>`)
			}, false},
			{"foreign-reference", func(e *nativeExtractor) {
				e.pkg.files["other.xml"] = []byte(`<x:endnoteReference xmlns:x="urn:spoof"/>`)
			}, false},
			{"orphan", func(e *nativeExtractor) { addNotes(e); e.pkg.rels = nil }, false},
			{"external", func(e *nativeExtractor) { addNotes(e); e.pkg.rels[e.mainPart][0].External = true }, false},
			{"duplicate-rel", func(e *nativeExtractor) {
				addNotes(e)
				e.pkg.rels[e.mainPart] = append(e.pkg.rels[e.mainPart], e.pkg.rels[e.mainPart][0])
			}, false},
			{"wrong-uri", func(e *nativeExtractor) { addNotes(e); e.pkg.rels[e.mainPart][0].Type = "urn:spoof/endnotes" }, false},
			{"malformed", func(e *nativeExtractor) { e.pkg.files["other.xml"] = []byte(`<broken`) }, false},
			{"budget", func(e *nativeExtractor) {
				e.pkg.files["other.xml"] = []byte(`<x>` + strings.Repeat(" ", 8*1024*1024) + `</x>`)
			}, false},
			{"sentinel-text", func(e *nativeExtractor) {
				addNotes(e)
				e.pkg.files["word/endnotes.xml"] = bytes.Replace(e.pkg.files["word/endnotes.xml"], []byte(`<w:separator/>`), []byte(`<w:t>visible</w:t>`), 1)
			}, false},
		} {
			t.Run(tc.name+ns, func(t *testing.T) {
				e := fresh()
				tc.change(e)
				if got := e.proveNoContentEndnotes(); got != tc.ok {
					t.Fatalf("got %v want %v", got, tc.ok)
				}
				if e.proveNoContentEndnotes() != tc.ok {
					t.Fatal("cached result changed")
				}
			})
		}
	}
}
