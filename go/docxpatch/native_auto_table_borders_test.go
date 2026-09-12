package docxpatch

import (
	"encoding/json"
	"strings"
	"testing"
)

func TestAutomaticBorderWhitePageUsesIndexedSourceOnce(t *testing.T) {
	for _, content := range []string{`<w:body><w:p/></w:body>`, `<w:background w:color="000000"/><w:body/>`} {
		root, err := parseNativeXML("word/document.xml", []byte(`<w:document xmlns:w="`+wordMLTransitional+`">`+content+`</w:document>`))
		if err != nil {
			t.Fatal(err)
		}
		resolver := &nativeLayoutResolver{doc: &NativeDocumentV1{}, wordNS: wordMLTransitional, mainRoot: root}
		want := !strings.Contains(content, "background")
		if got := resolver.automaticBorderWhitePage(); got != want || !resolver.autoBorderWhiteChecked {
			t.Fatalf("initial qualification = %v; want %v", got, want)
		}
		// A resolver's source is immutable in production. Removing this test's
		// pointer proves repeat table checks use the cached result, including refusal.
		resolver.mainRoot = nil
		if got := resolver.automaticBorderWhitePage(); got != want {
			t.Fatalf("cached qualification = %v; want %v", got, want)
		}
	}
}

func TestAutomaticTableBorderEvidencePreservesStrictSource(t *testing.T) {
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		for _, test := range []struct {
			name, border, cell, style, root string
			direct, valid                   bool
		}{
			{name: "inherited absent white", valid: true},
			{name: "direct absent white", direct: true, valid: true},
			{name: "white cell", cell: `<w:shd w:val="clear" w:fill="FFFFFF"/>`, valid: true},
			{name: "white style", style: `<w:tcPr><w:shd w:val="clear" w:fill="FFFFFF"/></w:tcPr>`, valid: true},
			{name: "dark cell", cell: `<w:shd w:val="clear" w:fill="000000"/>`},
			{name: "dark style", style: `<w:tcPr><w:shd w:val="clear" w:fill="000000"/></w:tcPr>`},
			{name: "unknown style fill", style: `<w:tcPr><w:shd w:val="clear" w:themeFill="accent1"/></w:tcPr>`},
			{name: "conditional", style: `<w:tblStylePr w:type="firstRow"/>`},
			{name: "background", root: `<w:background w:color="000000"/>`},
			{name: "white background unqualified", root: `<w:background w:color="FFFFFF"/>`},
			{name: "foreign root", root: `<x:unknown xmlns:x="urn:foreign"/>`},
			{name: "theme auto", border: ` w:themeColor="accent1"`},
			{name: "tint", border: ` w:themeTint="FF"`},
			{name: "foreign border", border: ` xmlns:x="urn:foreign" x:value="1"`},
			{name: "cell borders", cell: `<w:tcBorders><w:top w:val="single" w:sz="8" w:color="000000"/></w:tcBorders>`},
			{name: "merged", cell: `<w:gridSpan w:val="2"/>`},
			{name: "size overflow"},
			{name: "hex size"},
			{name: "nonzero space"},
			{name: "wrong case"},
			{name: "duplicate border"},
			{name: "multiple border owners"},
		} {
			t.Run(test.name+ns, func(t *testing.T) {
				borders := `<w:tblBorders>`
				for _, edge := range []string{"top", "right", "bottom", "left", "insideH", "insideV"} {
					borders += `<w:` + edge + ` w:val="single" w:sz="4" w:color="auto"` + test.border + `/>`
				}
				borders += `</w:tblBorders>`
				switch test.name {
				case "size overflow":
					borders = strings.ReplaceAll(borders, `w:sz="4"`, `w:sz="769"`)
				case "hex size":
					borders = strings.ReplaceAll(borders, `w:sz="4"`, `w:sz="0x4"`)
				case "nonzero space":
					borders = strings.ReplaceAll(borders, `w:sz="4"`, `w:sz="4" w:space="1"`)
				case "wrong case":
					borders = strings.ReplaceAll(borders, `w:color="auto"`, `w:color="AUTO"`)
				case "duplicate border":
					borders = strings.Replace(borders, `</w:tblBorders>`, `<w:top w:val="single" w:sz="4" w:color="auto"/></w:tblBorders>`, 1)
				}
				styleBorders, directBorders := borders, ""
				if test.direct {
					styleBorders = ""
					directBorders = borders
				}
				if test.name == "multiple border owners" {
					directBorders = borders
				}
				styles := `<w:styles xmlns:w="` + ns + `"><w:style w:type="table" w:styleId="Grid"><w:pPr><w:spacing w:after="0"/></w:pPr><w:tblPr>` + styleBorders + `</w:tblPr>` + test.style + `</w:style></w:styles>`
				parts := resolvedStylesTestParts(styles)
				parts["word/document.xml"] = `<w:document xmlns:w="` + ns + `">` + test.root + `<w:body><w:tbl><w:tblPr><w:tblStyle w:val="Grid"/>` + directBorders + `</w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="4000"/>` + test.cell + `</w:tcPr><w:p/></w:tc></w:tr></w:tbl><w:p/></w:body></w:document>`
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
				layout, err := ResolveNativeDocumentLayoutV1(data)
				if err != nil {
					t.Fatal(err)
				}
				fact := layout.Tables[0].AutomaticBorderPreview
				if !test.valid {
					if fact != nil {
						t.Fatalf("unexpected approximate eligibility %#v", fact)
					}
					return
				}
				if fact == nil {
					t.Fatalf("missing evidence: %#v", layout.Diagnostics)
				}
				if fact.Policy != NativeAutomaticTableBorderPolicyV1 || fact.PackageSHA256 != doc.Source.PackageSHA256 || !fact.ReadOnly || len(fact.AutomaticEdges) != 6 || len(fact.CellIDs) != 1 || len(fact.SourceDiagnostics) != 1 {
					t.Fatalf("invalid evidence %#v", fact)
				}
				if fact.SourceSHA256 != nativeSHA([]byte(parts[fact.SourcePart])) || *fact.Borders.Top.ColorRGB != "000000" || fact.Borders.Top.SizeEighthPoints != 4 {
					t.Fatal("wrong source or paint projection")
				}
				if doc.Body.Blocks[0].Table.Borders != nil || layout.Tables[0].Borders != nil {
					t.Fatal("strict source was recolored")
				}
				if layout.Paragraphs[0].Properties.SpacingAfterTwips == nil || *layout.Paragraphs[0].Properties.SpacingAfterTwips != 0 {
					t.Fatal("whole-table paragraph style lost")
				}
				if string(data) != before {
					t.Fatal("source bytes changed")
				}
				if test.direct && len(doc.Unsupported) == 0 {
					t.Fatal("strict direct border refusal lost")
				}
				for _, mutate := range []func(*NativeAutomaticTableBorderPreviewV1){
					func(f *NativeAutomaticTableBorderPreviewV1) { f.ReadOnly = false },
					func(f *NativeAutomaticTableBorderPreviewV1) { f.PackageSHA256 = "forged" },
					func(f *NativeAutomaticTableBorderPreviewV1) { f.CellIDs = append(f.CellIDs, f.CellIDs[0]) },
					func(f *NativeAutomaticTableBorderPreviewV1) {
						f.AutomaticEdges = append(f.AutomaticEdges, f.AutomaticEdges[0])
					},
					func(f *NativeAutomaticTableBorderPreviewV1) { f.BackgroundRGB = "000000" },
					func(f *NativeAutomaticTableBorderPreviewV1) { f.SourceDiagnostics = nil },
				} {
					encoded, _ := json.Marshal(layout)
					var changed NativeResolvedLayoutInputV1
					if err := json.Unmarshal(encoded, &changed); err != nil {
						t.Fatal(err)
					}
					mutate(changed.Tables[0].AutomaticBorderPreview)
					if _, err := EncodeNativeResolvedLayoutInputV1(&changed); err == nil {
						t.Fatal("accepted invalid evidence")
					}
				}
			})
		}
	}
}
