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

func TestAbsentFontSizeTableCellParagraphsRequireSizeFreeTableStyle(t *testing.T) {
	for _, tc := range []struct {
		name, tableStyle, tblStyleRef, wantMark, wantRun string
		want                                             int
	}{
		{name: "size-free table style chain", tableStyle: `<w:style w:type="table" w:styleId="Procedure"><w:basedOn w:val="TableNormal"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr><w:tblPr><w:tblInd w:w="360" w:type="dxa"/></w:tblPr></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:tblPr><w:tblCellMar><w:left w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>`, tblStyleRef: `<w:tblStyle w:val="Procedure"/>`, want: 2},
		{name: "table style size", tableStyle: `<w:style w:type="table" w:styleId="Procedure"><w:rPr><w:sz w:val="20"/></w:rPr></w:style>`, tblStyleRef: `<w:tblStyle w:val="Procedure"/>`},
		{name: "table style ancestor size", tableStyle: `<w:style w:type="table" w:styleId="Procedure"><w:basedOn w:val="TableNormal"/></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:rPr><w:sz w:val="20"/></w:rPr></w:style>`, tblStyleRef: `<w:tblStyle w:val="Procedure"/>`},
		{name: "conditional table style regions", tableStyle: `<w:style w:type="table" w:styleId="Procedure"><w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr></w:style>`, tblStyleRef: `<w:tblStyle w:val="Procedure"/>`},
		{name: "missing table style", tblStyleRef: `<w:tblStyle w:val="Procedure"/>`},
		{name: "no table style and no default table style", want: 2},
		{name: "no table style with size-free default table style", tableStyle: `<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:tblPr><w:tblInd w:w="0" w:type="dxa"/></w:tblPr></w:style>`, want: 2},
		{name: "no table style with sized default table style", tableStyle: `<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:rPr><w:sz w:val="20"/></w:rPr></w:style>`},
	} {
		t.Run(tc.name, func(t *testing.T) {
			styles := `<w:styles xmlns:w="` + wordMLTransitional + `" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"/><w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"/>` + tc.tableStyle + `</w:styles>`
			parts := resolvedStylesTestParts(styles)
			parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:tbl><w:tblPr>` + tc.tblStyleRef + `<w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:ind w:left="360"/></w:pPr><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>Sized</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
			data := buildNativeDOCX(t, nativeEntries(parts))
			before := bytes.Clone(data)
			eligibility, err := ExtractNativeDocxApproximationEligibilityV1(data)
			if err != nil {
				t.Fatal(err)
			}
			var cell []NativeDocxAbsentFontSizeV1
			for _, fact := range eligibility.AbsentFontSizes {
				if strings.Contains(fact.Path, "/w:tbl[1]/") {
					cell = append(cell, fact)
				}
			}
			if len(cell) != tc.want {
				t.Fatalf("cell facts=%d want %d: %#v", len(cell), tc.want, eligibility.AbsentFontSizes)
			}
			for _, fact := range cell {
				if fact.PackageSHA256 != eligibility.PackageSHA256 || fact.PartName != "word/document.xml" || !strings.HasPrefix(fact.Path, "/w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]/w:tc[1]/w:p[1]") || (fact.ScopeKind != "paragraph-mark" && fact.ScopeKind != "run") {
					t.Fatalf("unbound cell fact %#v", fact)
				}
			}
			layout, err := ResolveNativeDocumentLayoutV1(data)
			if err != nil {
				t.Fatal(err)
			}
			for _, run := range layout.Runs {
				if run.Properties.FontSizeHalfPoint != nil && *run.Properties.FontSizeHalfPoint != 24 && tc.want > 0 {
					t.Fatalf("strict size was invented: %#v", run.Properties)
				}
			}
			if !bytes.Equal(data, before) {
				t.Fatal("source mutated")
			}
		})
	}
}

func TestAbsentFontSizeAcceptsIgnorableStylesRootOnly(t *testing.T) {
	for _, tc := range []struct {
		name, attrs string
		want        int
	}{
		{name: "markup compatibility ignorable", attrs: ` xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"`, want: 2},
		{name: "foreign root attribute", attrs: ` unknown="value"`},
	} {
		t.Run(tc.name, func(t *testing.T) {
			styles := `<w:styles xmlns:w="` + wordMLTransitional + `"` + tc.attrs + `><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"/><w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"/></w:styles>`
			parts := resolvedStylesTestParts(styles)
			parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:r><w:t>Source</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
			eligibility, err := ExtractNativeDocxApproximationEligibilityV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			if len(eligibility.AbsentFontSizes) != tc.want {
				t.Fatalf("got %d want %d: %#v", len(eligibility.AbsentFontSizes), tc.want, eligibility.AbsentFontSizes)
			}
		})
	}
}

// TestAbsentFontSizeNoStylesPart pins the difference between a package that
// omits the styles part entirely and one that ships it. Without the part there
// is no w:docDefaults and no default paragraph/character style, so an empty
// applied-style chain is the complete chain and the absence is proven. With the
// part present the same empty chain may mean an unresolved or default style
// that carries a size, so it keeps refusing.
func TestAbsentFontSizeNoStylesPart(t *testing.T) {
	body := func(mark, run string) string {
		return `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:pPr><w:rPr>` + mark + `</w:rPr></w:pPr><w:r><w:rPr>` + run + `</w:rPr><w:t>Source</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
	}
	noStyles := func(document string) map[string]string {
		return map[string]string{
			"[Content_Types].xml":          `<Types xmlns="` + opcContentTypesNS + `"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
			"_rels/.rels":                  `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="office" Type="` + relBaseTransitional + `officeDocument" Target="word/document.xml"/></Relationships>`,
			"word/document.xml":            document,
			"word/_rels/document.xml.rels": `<Relationships xmlns="` + opcRelationshipsNS + `"></Relationships>`,
		}
	}
	for _, tc := range []struct {
		name, mark, run string
		want            int
	}{
		{name: "absent", want: 2},
		{name: "mark explicit", mark: `<w:sz w:val="24"/>`, want: 1},
		{name: "run explicit", run: `<w:sz w:val="24"/>`, want: 1},
		{name: "both explicit", mark: `<w:sz w:val="24"/>`, run: `<w:sz w:val="24"/>`},
		{name: "run invalid", run: `<w:sz w:val="oops"/>`, want: 1},
	} {
		t.Run(tc.name, func(t *testing.T) {
			data := buildNativeDOCX(t, nativeEntries(noStyles(body(tc.mark, tc.run))))
			before := bytes.Clone(data)
			eligibility, err := ExtractNativeDocxApproximationEligibilityV1(data)
			if err != nil {
				t.Fatal(err)
			}
			if len(eligibility.AbsentFontSizes) != tc.want {
				t.Fatalf("got%d want%d: %#v", len(eligibility.AbsentFontSizes), tc.want, eligibility.AbsentFontSizes)
			}
			for _, fact := range eligibility.AbsentFontSizes {
				if fact.PackageSHA256 != eligibility.PackageSHA256 || fact.PartName != "word/document.xml" || !strings.HasPrefix(fact.Path, "/w:document[1]/w:body[1]/w:p[1]") {
					t.Fatalf("unbound fact %#v", fact)
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
	// A styles part keeps the old rule: an empty applied-style chain is never
	// evidence, whatever the part happens to declare.
	for _, styles := range []string{
		`<w:styles xmlns:w="` + wordMLTransitional + `"/>`,
		`<w:styles xmlns:w="` + wordMLTransitional + `"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:rPr><w:sz w:val="24"/></w:rPr></w:style></w:styles>`,
	} {
		parts := resolvedStylesTestParts(styles)
		parts["word/document.xml"] = body("", "")
		data := buildNativeDOCX(t, nativeEntries(parts))
		eligibility, err := ExtractNativeDocxApproximationEligibilityV1(data)
		if err != nil {
			t.Fatal(err)
		}
		if len(eligibility.AbsentFontSizes) != 0 {
			t.Fatalf("a styles-bearing package guessed a missing style chain: %#v", eligibility.AbsentFontSizes)
		}
	}
}
