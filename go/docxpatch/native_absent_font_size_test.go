package docxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func TestAbsentFontSizeReservedNotesOnly(t *testing.T) {
	for _, malformed := range []bool{false, true} {
		parts := nativeNotePagePaintParts()
		parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/Custom/Styles.XML" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`, 1)
		parts["Custom/_RELS/Main.XML.RELS"] = strings.Replace(parts["Custom/_RELS/Main.XML.RELS"], `</Relationships>`, `<Relationship Id="rStyles" Type="`+relBaseTransitional+`styles" Target="Styles.XML"/></Relationships>`, 1)
		parts["Custom/Styles.XML"] = `<w:styles xmlns:w="` + testW + `"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"/></w:styles>`
		for _, name := range []string{"Custom/Notes/Foot.XML", "Custom/Notes/End.XML"} {
			parts[name] = strings.ReplaceAll(parts[name], `<w:t>---</w:t>`, `<w:separator/>`)
			parts[name] = strings.ReplaceAll(parts[name], `<w:t>continued</w:t>`, `<w:continuationSeparator/>`)
			if malformed {
				parts[name] = strings.ReplaceAll(parts[name], `<w:separator/>`, `<w:separator extra="bad"/>`)
			}
		}
		data := buildNativeDOCX(t, nativeEntries(parts))
		before := bytes.Clone(data)
		facts, err := nativeAbsentFontSizes(data)
		if err != nil {
			t.Fatal(err)
		}
		count := 0
		for _, fact := range facts {
			if !strings.Contains(fact.PartName, "Notes/") {
				continue
			}
			count++
			if fact.ScopeKind != "paragraph-mark" || strings.Contains(fact.Path, "note[3]") {
				t.Fatalf("content note qualified: %#v", fact)
			}
		}
		want := 4
		if malformed {
			want = 0
		}
		if count != want {
			t.Fatalf("note facts=%d want%d: %#v", count, want, facts)
		}
		if !bytes.Equal(data, before) {
			t.Fatal("source mutated")
		}
	}
}

func TestAbsentFontSizeOwnerRejectsAmbiguousAndForeignLayers(t *testing.T) {
	for _, markup := range []string{
		`<w:rPr/><w:rPr/>`,
		`<w:rPr><x:sz xmlns:x="urn:foreign" w:val="22"/></w:rPr>`,
		`<w:rPr unknown="size"><w:i/></w:rPr>`,
		`<w:rPr><w:sz w:val="22" extra="bad"/></w:rPr>`,
		`<w:rPr><w:rPrChange><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrChange></w:rPr>`,
	} {
		node, err := parseNativeXML("word/document.xml", []byte(`<w:r xmlns:w="`+wordMLTransitional+`">`+markup+`</w:r>`))
		if err != nil {
			continue
		} // Namespace-spoof hard rejection is also safe.
		r := &nativeLayoutResolver{wordNS: wordMLTransitional}
		if r.absentOwnerRunSize(node) {
			t.Fatalf("false absence: %s", markup)
		}
	}
}

func TestAbsentFontSizeEvidenceRequiresRealSourceAbsence(t *testing.T) {
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		for _, tc := range []struct {
			name, defaults, style, mark, run, based string
			want                                    int
		}{
			{name: "absent", want: 2},
			{name: "default explicit", defaults: `<w:sz w:val="24"/>`},
			{name: "default invalid", defaults: `<w:sz w:val="oops"/>`},
			{name: "style explicit", style: `<w:sz w:val="24"/>`},
			{name: "style invalid", style: `<w:sz w:val="oops"/>`},
			{name: "style duplicate", style: `<w:sz w:val="22"/><w:sz w:val="24"/>`},
			{name: "style unknown", style: `<w:unknown/>`},
			{name: "style complex size", style: `<w:szCs w:val="22"/>`},
			{name: "mark explicit", mark: `<w:sz w:val="24"/>`, want: 1},
			{name: "mark invalid", mark: `<w:sz w:val="oops"/>`, want: 1},
			{name: "mark duplicate", mark: `<w:sz w:val="22"/><w:sz w:val="24"/>`, want: 1},
			{name: "run explicit", run: `<w:sz w:val="24"/>`, want: 1},
			{name: "run invalid", run: `<w:sz w:val="oops"/>`, want: 1},
			{name: "run duplicate", run: `<w:sz w:val="22"/><w:sz w:val="24"/>`, want: 1},
			{name: "missing ancestor", based: `<w:basedOn w:val="Missing"/>`},
			{name: "malformed ancestor", based: `<w:basedOn/>`},
			{name: "cycle", based: `<w:basedOn w:val="Normal"/>`},
		} {
			t.Run(tc.name+ns, func(t *testing.T) {
				styles := `<w:styles xmlns:w="` + wordMLTransitional + `"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>` + tc.defaults + `</w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal">` + tc.based + `<w:rPr>` + tc.style + `</w:rPr></w:style><w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"/></w:styles>`
				parts := resolvedStylesTestParts(styles)
				parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:pPr><w:rPr>` + tc.mark + `</w:rPr></w:pPr><w:r><w:rPr>` + tc.run + `</w:rPr><w:t>Source</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
				if ns == wordMLStrict {
					for k, v := range parts {
						parts[k] = strings.ReplaceAll(strings.ReplaceAll(v, wordMLTransitional, wordMLStrict), relBaseTransitional, relBaseStrict)
					}
				}
				data := buildNativeDOCX(t, nativeEntries(parts))
				before := bytes.Clone(data)
				eligibility, err := ExtractNativeDocxApproximationEligibilityV1(data)
				if err != nil {
					t.Fatal(err)
				}
				if len(eligibility.AbsentFontSizes) != tc.want {
					t.Fatalf("got%d want%d: %#v", len(eligibility.AbsentFontSizes), tc.want, eligibility.AbsentFontSizes)
				}
				for _, f := range eligibility.AbsentFontSizes {
					if f.PackageSHA256 != eligibility.PackageSHA256 || f.PartName != "word/document.xml" || f.ScopeID == "" || !strings.HasPrefix(f.Path, "/w:document[1]/w:body[1]/w:p[1]") {
						t.Fatalf("unboundfact %#v", f)
					}
				}
				layout, err := ResolveNativeDocumentLayoutV1(data)
				if err != nil {
					t.Fatal(err)
				}
				if tc.name == "absent" && (layout.Paragraphs[0].ParagraphMarkProperties.FontSizeHalfPoint != nil || layout.Runs[0].Properties.FontSizeHalfPoint != nil) {
					t.Fatal("strict size was invented")
				}
				if !bytes.Equal(data, before) {
					t.Fatal("source mutated")
				}
			})
		}
	}
}
