package pptxpatch

import (
	"bytes"
	"encoding/xml"
	"strings"
	"testing"
)

func nativeDiagnosticCodes(element NativeElement) map[string]int {
	codes := map[string]int{}
	for _, diagnostic := range element.Compatibility.Diagnostics {
		codes[diagnostic.Code]++
	}
	return codes
}

func TestNativeAuthoredNormAutofitStrictStillRefuses(t *testing.T) {
	for _, strict := range []bool{false, true} {
		deck, err := ExtractNativePPTX(nativeSourceFrameFixture(t, strict, `<a:bodyPr><a:normAutofit fontScale="62500" lnSpcReduction="20000"/></a:bodyPr>`), nativeMutationExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		element := deck.Slides[0].Elements[0]
		if element.Compatibility.Status != NativeCompatibilityStatusRefused || element.TextBody != nil {
			t.Fatalf("strict extraction painted normAutofit text: %+v", element)
		}
		for _, run := range (*element.Paragraphs)[0].Runs {
			if run.FontSizeHundredthPt != nil && *run.FontSizeHundredthPt != 3200 {
				t.Fatal("strict extraction scaled run sizes")
			}
		}
		if nativeDiagnosticCodes(element)[nativeAuthoredAutoFitCode] != 0 {
			t.Fatal("strict extraction emitted the approximation code")
		}
	}
}

func TestNativeAuthoredNormAutofitScalesRunsReadOnly(t *testing.T) {
	for _, strict := range []bool{false, true} {
		original := nativeSourceFrameFixture(t, strict, `<a:bodyPr><a:normAutofit fontScale="62500" lnSpcReduction="20000"/></a:bodyPr>`)
		before := bytes.Clone(original)
		options := nativeMutationExtractOptions()
		options.AllowSourceFrameAutoFitPreview = true
		deck, err := ExtractNativePPTX(original, options)
		if err != nil {
			t.Fatal(err)
		}
		element := deck.Slides[0].Elements[0]
		if element.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || element.TextBody == nil || element.TextBody.AutoFit != "none" || element.Paragraphs == nil || len(*element.Paragraphs) != 1 {
			t.Fatalf("authored autofit preview missing: %+v", element)
		}
		for _, run := range (*element.Paragraphs)[0].Runs {
			// The fixture authors sz="3200"; 3200 × 62.5% = 2000.
			if run.FontSizeHundredthPt == nil || *run.FontSizeHundredthPt != 2000 {
				t.Fatalf("run size was not scaled by the authored fontScale: %+v", run)
			}
		}
		codes := nativeDiagnosticCodes(element)
		if codes[nativeAuthoredAutoFitCode] != 1 || codes[nativeTextColumnsOmittedCode] != 0 {
			t.Fatalf("expected one authored autofit disclosure: %+v", element.Compatibility.Diagnostics)
		}
		for _, diagnostic := range element.Compatibility.Diagnostics {
			if diagnostic.Code == nativeAuthoredAutoFitCode && (!strings.Contains(diagnostic.Message, "fontScale=62.5%") || !strings.Contains(diagnostic.Message, "lnSpcReduction=20%")) {
				t.Fatalf("disclosure lacks the authored values: %s", diagnostic.Message)
			}
		}
		if issues := ValidateNativePPTX(deck); len(issues) > 0 {
			t.Fatalf("invalid approximate contract: %+v", issues)
		}
		if !bytes.Equal(original, before) {
			t.Fatal("preview changed source bytes")
		}
		paragraphs := nativeMutationParagraphs("No mutation authority")
		if _, err := ApplyNativePPTXMutations(original, NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{{OperationID: "replace", Kind: NativePPTXReplaceText, ElementID: element.ID, ExpectedFingerprintSHA256: element.Source.FingerprintSHA256, Paragraphs: &paragraphs}}}); err == nil {
			t.Fatal("authored autofit approximation granted editing permission")
		}
		deck.Slides[0].Elements[0].Compatibility.Status = NativeCompatibilityStatusEditable
		if len(ValidateNativePPTX(deck)) == 0 {
			t.Fatal("approximation validated as editable")
		}
	}
}

func TestNativeAuthoredNormAutofitWithoutValuesIsFullSize(t *testing.T) {
	options := nativeMutationExtractOptions()
	options.AllowSourceFrameAutoFitPreview = true
	deck, err := ExtractNativePPTX(nativeSourceFrameFixture(t, false, `<a:bodyPr><a:normAutofit/></a:bodyPr>`), options)
	if err != nil {
		t.Fatal(err)
	}
	element := deck.Slides[0].Elements[0]
	if element.Paragraphs == nil || len(*element.Paragraphs) != 1 || *(*element.Paragraphs)[0].Runs[0].FontSizeHundredthPt != 3200 {
		t.Fatalf("absent fontScale must mean 100%%: %+v", element.Paragraphs)
	}
	if element.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || nativeDiagnosticCodes(element)[nativeAuthoredAutoFitCode] != 1 {
		t.Fatalf("normAutofit without values must still be a disclosed read-only preview: %+v", element.Compatibility)
	}
}

func TestNativeAuthoredTextColumnsPaintSingleColumnWithDisclosure(t *testing.T) {
	options := nativeMutationExtractOptions()
	options.AllowSourceFrameAutoFitPreview = true
	deck, err := ExtractNativePPTX(nativeSourceFrameFixture(t, false, `<a:bodyPr numCol="3" spcCol="108000"/>`), options)
	if err != nil {
		t.Fatal(err)
	}
	element := deck.Slides[0].Elements[0]
	if element.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || element.Paragraphs == nil || len(*element.Paragraphs) != 1 {
		t.Fatalf("columns were not painted as a disclosed single column: %+v", element)
	}
	codes := nativeDiagnosticCodes(element)
	if codes[nativeTextColumnsOmittedCode] != 1 || codes[nativeAuthoredAutoFitCode] != 0 {
		t.Fatalf("expected one column disclosure: %+v", element.Compatibility.Diagnostics)
	}
	if issues := ValidateNativePPTX(deck); len(issues) > 0 {
		t.Fatalf("invalid approximate contract: %+v", issues)
	}
	strictDeck, err := ExtractNativePPTX(nativeSourceFrameFixture(t, false, `<a:bodyPr numCol="3" spcCol="108000"/>`), nativeMutationExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	if strictDeck.Slides[0].Elements[0].Compatibility.Status != NativeCompatibilityStatusRefused {
		t.Fatal("strict extraction painted multi-column text")
	}
}

func TestNativeApplyAuthoredFontScaleRounding(t *testing.T) {
	paragraphs := []NativeParagraph{{Runs: []NativeTextRun{{FontSizeHundredthPt: int64Pointer(3200)}, {FontSizeHundredthPt: int64Pointer(1)}, {}}}}
	nativeApplyAuthoredFontScale(paragraphs, &nativeAuthoredAutoFit{normAutofit: true, fontScale: 85000})
	if *paragraphs[0].Runs[0].FontSizeHundredthPt != 2720 || *paragraphs[0].Runs[1].FontSizeHundredthPt != 1 || paragraphs[0].Runs[2].FontSizeHundredthPt != nil {
		t.Fatalf("unexpected scaling: %+v", paragraphs[0].Runs)
	}
	nativeApplyAuthoredFontScale(paragraphs, nil)
	nativeApplyAuthoredFontScale(paragraphs, &nativeAuthoredAutoFit{normAutofit: true, fontScale: nativeAuthoredAutoFitFullSize})
	if *paragraphs[0].Runs[0].FontSizeHundredthPt != 2720 {
		t.Fatal("full-size or absent autofit must not change sizes")
	}
}

func nativePlaceholderPreviewFixture(t *testing.T, strict bool, slidePlaceholder, layoutPlaceholder string, removeSlideText bool) []byte {
	return nativePlaceholderFixture(t, strict, func(parts map[string]string) {
		slide := parts["relocated/slides/slide-a.xml"]
		slide = strings.Replace(slide, `<p:ph idx="7"/>`, slidePlaceholder, 1)
		if removeSlideText {
			start, end := strings.Index(slide, "<p:txBody>"), strings.Index(slide, "</p:txBody>")+len("</p:txBody>")
			slide = slide[:start] + slide[end:]
		}
		parts["relocated/slides/slide-a.xml"] = slide
		parts["relocated/layouts/layout.xml"] = strings.Replace(parts["relocated/layouts/layout.xml"], `<p:ph type="body" idx="7"/>`, layoutPlaceholder, 1)
	})
}

func TestNativeSubTitlePlaceholderInheritsOnlyInPreview(t *testing.T) {
	for _, strict := range []bool{false, true} {
		input := nativePlaceholderPreviewFixture(t, strict, `<p:ph type="subTitle" idx="7"/>`, `<p:ph type="subTitle" idx="7"/>`, false)
		before := bytes.Clone(input)
		strictDeck, err := ExtractNativePPTX(input, nativeTestExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		for _, element := range strictDeck.Slides[0].Elements {
			if element.Compatibility.Status != NativeCompatibilityStatusRefused {
				t.Fatalf("strict extraction projected a subTitle placeholder: %+v", element)
			}
		}
		options := nativeTestExtractOptions()
		options.AllowInheritedTextPreview = true
		deck, err := ExtractNativePPTX(input, options)
		if err != nil {
			t.Fatal(err)
		}
		if len(deck.Slides[0].Elements) != 1 {
			t.Fatalf("placeholder preview missing: %+v", deck.Slides[0].Compatibility)
		}
		element := deck.Slides[0].Elements[0]
		if element.Kind != NativeElementKindText || element.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || element.Placeholder == nil || *element.Placeholder != NativePlaceholderTypeSubTitle || element.Paragraphs == nil || len(*element.Paragraphs) == 0 {
			t.Fatalf("subTitle preview failed: %+v", element)
		}
		paragraph := (*element.Paragraphs)[0]
		if *element.Transform.X != 914400 || *element.Transform.Y != 457200 || paragraph.MarginLeftEmu == nil || *paragraph.MarginLeftEmu != 400000 || paragraph.Runs[0].FontFamily == nil || *paragraph.Runs[0].FontFamily != "Calibri" || paragraph.Runs[0].FontSizeHundredthPt == nil || *paragraph.Runs[0].FontSizeHundredthPt != 2400 {
			t.Fatalf("layout list style / master geometry cascade failed: %+v %+v", element.Transform, paragraph)
		}
		codes := nativeDiagnosticCodes(element)
		if codes[nativePlaceholderPreviewCode] != 1 || codes[nativeInheritedTextPreviewCode] != 1 || codes["pptx.inherited-placeholder-preview"] != 0 {
			t.Fatalf("placeholder preview must be disclosed as approximate inherited text: %+v", element.Compatibility.Diagnostics)
		}
		if issues := ValidateNativePPTX(deck); len(issues) > 0 {
			t.Fatalf("invalid native: %+v", issues)
		}
		if !bytes.Equal(input, before) {
			t.Fatal("source mutated")
		}
		replacement := nativeMutationParagraphs("changed")
		if _, err := resolveNativePPTXMutations(deck, []NativePPTXMutation{{OperationID: "edit", Kind: NativePPTXReplaceText, ElementID: element.ID, ExpectedFingerprintSHA256: element.Source.FingerprintSHA256, Paragraphs: &replacement}}); err == nil {
			t.Fatal("placeholder preview authorized mutation")
		}
	}
}

func TestNativeEmptyContentPlaceholderPaintsNoPromptText(t *testing.T) {
	input := nativePlaceholderPreviewFixture(t, false, `<p:ph type="obj" sz="quarter" idx="7"/>`, `<p:ph type="obj" sz="quarter" idx="7"/>`, true)
	options := nativeTestExtractOptions()
	options.AllowInheritedTextPreview = true
	deck, err := ExtractNativePPTX(input, options)
	if err != nil {
		t.Fatal(err)
	}
	if len(deck.Slides[0].Elements) != 1 {
		t.Fatalf("empty content placeholder was refused: %+v", deck.Slides[0].Compatibility)
	}
	element := deck.Slides[0].Elements[0]
	if element.Kind != NativeElementKindText || element.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || element.Paragraphs == nil || len(*element.Paragraphs) != 0 || element.TextBody == nil || element.Placeholder == nil || *element.Placeholder != NativePlaceholderTypeBody {
		t.Fatalf("empty placeholder projection invented content or lost its frame: %+v", element)
	}
	if issues := ValidateNativePPTX(deck); len(issues) > 0 {
		t.Fatalf("invalid native: %+v", issues)
	}
	found := false
	for _, diagnostic := range element.Compatibility.Diagnostics {
		found = found || (diagnostic.Code == nativePlaceholderPreviewCode && strings.Contains(diagnostic.Message, "no text body"))
	}
	if !found {
		t.Fatalf("missing empty-placeholder disclosure: %+v", element.Compatibility.Diagnostics)
	}
}

func TestNativePlaceholderPreviewPaintsAncestorFrameAndRefusesOtherKinds(t *testing.T) {
	input := nativePlaceholderFixture(t, false, func(parts map[string]string) {
		layout := parts["relocated/layouts/layout.xml"]
		layout = strings.Replace(layout, `<p:ph type="body" idx="7"/></p:nvPr></p:nvSpPr><p:spPr>`, `<p:ph type="body" idx="7"/></p:nvPr></p:nvSpPr><p:spPr><a:solidFill><a:srgbClr val="FBE4D5"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="C55A11"/></a:solidFill></a:ln>`, 1)
		layout = strings.Replace(layout, `<a:p/>`, `<a:p><a:r><a:rPr lang="en-US"/><a:t>Layout prompt</a:t></a:r></a:p>`, 1)
		parts["relocated/layouts/layout.xml"] = layout
		parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `<p:ph idx="7"/>`, `<p:ph type="obj" idx="7"/>`, 1)
	})
	options := nativeTestExtractOptions()
	options.AllowInheritedTextPreview = true
	deck, err := ExtractNativePPTX(input, options)
	if err != nil {
		t.Fatal(err)
	}
	if len(deck.Slides[0].Elements) != 1 {
		t.Fatalf("obj placeholder with painted layout ancestor was refused: %+v", deck.Slides[0].Compatibility)
	}
	element := deck.Slides[0].Elements[0]
	if element.Kind != NativeElementKindShape || element.Preset == nil || *element.Preset != NativeShapePresetRect || element.Fill == nil || *element.Fill != "FBE4D5" || element.Stroke == nil || element.Stroke.Color != "C55A11" {
		t.Fatalf("ancestor frame paint was not projected: %+v", element)
	}
	disclosed := false
	for _, diagnostic := range element.Compatibility.Diagnostics {
		if diagnostic.Code == nativePlaceholderPreviewCode {
			disclosed = strings.Contains(diagnostic.Message, "solid fill FBE4D5") && strings.Contains(diagnostic.Message, "solid outline C55A11") && strings.Contains(diagnostic.Message, "ancestor prompt paragraphs")
		}
	}
	if !disclosed {
		t.Fatalf("inherited frame paint was not disclosed: %+v", element.Compatibility.Diagnostics)
	}
	if issues := ValidateNativePPTX(deck); len(issues) > 0 {
		t.Fatalf("invalid native: %+v", issues)
	}
	for _, p := range *element.Paragraphs {
		for _, r := range p.Runs {
			if r.Text != nil && strings.Contains(*r.Text, "Layout prompt") {
				t.Fatal("ancestor prompt leaked into slide content")
			}
		}
	}
	for _, kind := range []string{"dt", "ftr", "sldNum", "pic", "chart", "tbl", "media", "dgm", "clipArt", "sldImg", "hdr"} {
		input := nativePlaceholderPreviewFixture(t, false, `<p:ph type="`+kind+`" idx="7"/>`, `<p:ph type="`+kind+`" idx="7"/>`, false)
		deck, err := ExtractNativePPTX(input, options)
		if err != nil {
			t.Fatal(err)
		}
		for _, element := range deck.Slides[0].Elements {
			if element.Compatibility.Status != NativeCompatibilityStatusRefused {
				t.Fatalf("%s placeholder was projected as inherited text: %+v", kind, element)
			}
		}
	}
	hidden := nativePlaceholderPreviewFixture(t, false, `<p:ph type="subTitle" idx="7"/>`, `<p:ph type="subTitle" idx="7"/>`, false)
	hidden = bytes.Replace(hidden, []byte(`<p:cNvPr id="2" name="Title"/>`), []byte(`<p:cNvPr id="2" name="Title" hidden="1"/>`), 1)
	if !bytes.Contains(hidden, []byte(`hidden="1"`)) {
		t.Fatal("fixture shape identity changed; hidden test is inert")
	}
	deck, err = ExtractNativePPTX(hidden, options)
	if err == nil {
		for _, element := range deck.Slides[0].Elements {
			if element.Compatibility.Status != NativeCompatibilityStatusRefused {
				t.Fatalf("hidden placeholder was painted: %+v", element)
			}
		}
	}
}

func TestNativeInheritedPreviewOmissionsAreDisclosedAndBounded(t *testing.T) {
	d := nativeExtractDialect{drawing: nsDrawingTransitional, presentation: nsPresentationTransitional}
	a := func(local, value string) xml.Attr { return xml.Attr{Name: xml.Name{Local: local}, Value: value} }
	el := func(local string, attrs []xml.Attr, children ...*nativeXMLNode) *nativeXMLNode {
		return &nativeXMLNode{Name: xml.Name{Space: d.drawing, Local: local}, Attrs: attrs, Children: children}
	}
	omit := &nativeInheritedTextOmissions{}
	paragraph := &nativeXMLNode{Attrs: []xml.Attr{a("fontAlgn", "base")}, Children: []*nativeXMLNode{
		el("lnSpc", nil, el("spcPct", []xml.Attr{a("val", "90000")})),
		el("spcBef", nil, el("spcPts", []xml.Attr{a("val", "1000")})),
		el("buClr", nil, el("srgbClr", []xml.Attr{a("val", "000000")})),
		el("buSzPct", []xml.Attr{a("val", "100000")}),
		el("buFont", []xml.Attr{a("typeface", "Arial"), a("panose", "020B0604020202020204")}),
		el("buChar", []xml.Attr{a("char", "•")}),
		el("tabLst", nil, el("tab", []xml.Attr{a("pos", "914400"), a("algn", "l")})),
	}}
	clean, err := sanitizeNativeInheritedPreviewProperties(paragraph, d, true, nativeResolvedTheme{}, omit)
	if err != nil {
		t.Fatal(err)
	}
	if len(clean.Attrs) != 0 || len(clean.Children) != 2 {
		t.Fatalf("unexpected paragraph projection: %+v", clean)
	}
	run := &nativeXMLNode{Attrs: []xml.Attr{a("sz", "3200"), a("strike", "noStrike"), a("spc", "-1"), a("noProof", "1"), a("altLang", "en-US"), a("kern", "1200")}}
	if _, err := sanitizeNativeInheritedPreviewProperties(run, d, false, nativeResolvedTheme{}, omit); err != nil {
		t.Fatal(err)
	}
	if got := strings.Join(omit.names(), ","); got != "a:buClr,a:buFont@panose,a:buSzPct,a:lnSpc,a:pPr@fontAlgn,a:rPr@spc,a:rPr@strike=noStrike,a:spcBef,a:tabLst" {
		t.Fatalf("unexpected omission disclosure: %s", got)
	}
	// Paint-active properties and malformed spacing still refuse.
	for _, node := range []*nativeXMLNode{
		{Attrs: []xml.Attr{a("strike", "sngStrike")}},
		{Attrs: []xml.Attr{a("u", "sng")}},
		{Attrs: []xml.Attr{a("baseline", "30000")}},
		{Attrs: []xml.Attr{a("spc", "abc")}},
		{Children: []*nativeXMLNode{el("highlight", nil, el("srgbClr", []xml.Attr{a("val", "FFFF00")}))}},
	} {
		if _, err := sanitizeNativeInheritedPreviewProperties(node, d, false, nativeResolvedTheme{}, &nativeInheritedTextOmissions{}); err == nil {
			t.Fatalf("paint-active run property hidden: %+v", node)
		}
	}
	for _, node := range []*nativeXMLNode{
		{Children: []*nativeXMLNode{el("lnSpc", nil, el("spcPct", []xml.Attr{a("val", "abc")}))}},
		{Children: []*nativeXMLNode{el("lnSpc", nil, el("spcPct", []xml.Attr{a("val", "1")}), el("spcPts", []xml.Attr{a("val", "1")}))}},
		{Children: []*nativeXMLNode{el("buAutoNum", []xml.Attr{a("type", "arabicPeriod")})}},
		{Children: []*nativeXMLNode{el("buClr", nil, el("prstClr", []xml.Attr{a("val", "black")}))}},
		{Children: []*nativeXMLNode{el("buFont", []xml.Attr{a("typeface", "Arial"), a("panose", "zz")})}},
		{Attrs: []xml.Attr{a("fontAlgn", "weird")}},
	} {
		if _, err := sanitizeNativeInheritedPreviewProperties(node, d, true, nativeResolvedTheme{}, &nativeInheritedTextOmissions{}); err == nil {
			t.Fatalf("malformed or paint-active paragraph property hidden: %+v", node)
		}
	}
	bounded := &nativeInheritedTextOmissions{}
	for i := 0; i < nativeMaxInheritedOmissionNames+10; i++ {
		bounded.add(strings.Repeat("x", i+1))
	}
	if len(bounded.names()) != nativeMaxInheritedOmissionNames+1 {
		t.Fatalf("omission names are unbounded: %d", len(bounded.names()))
	}
}

func TestNativeInheritedPreviewProjectsBreaksAndEmptyParagraphs(t *testing.T) {
	for _, strict := range []bool{false, true} {
		data := nativeShapeReferenceFixture(t, strict, nativeShapeStyleRefs, `lang="en-US" dirty="0">`, "", "Hi", nativeShapeReferenceFonts, func(parts map[string]string) {
			ns := nsDrawingTransitional
			if strict {
				ns = nsDrawingStrict
			}
			parts["relocated/deck.xml"] = strings.Replace(parts["relocated/deck.xml"], `</p:presentation>`, `<p:defaultTextStyle xmlns:a="`+ns+`"><a:defPPr><a:defRPr lang="en-US"/></a:defPPr><a:lvl1pPr algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:defRPr></a:lvl1pPr></p:defaultTextStyle></p:presentation>`, 1)
			part := "relocated/slides/slide-a.xml"
			slide := parts[part]
			start := strings.LastIndex(slide, "</a:p>") + len("</a:p>")
			slide = slide[:start] + `<a:p><a:pPr marL="342900" indent="-342900"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:buChar char="-"/></a:pPr><a:r><a:rPr lang="en-US" sz="1800" strike="noStrike" spc="-1"/><a:t>First</a:t></a:r><a:br><a:rPr lang="en-US"/></a:br><a:r><a:rPr lang="en-US" sz="1800"/><a:t> </a:t></a:r></a:p><a:p><a:endParaRPr lang="en-US" sz="1800"/></a:p>` + slide[start:]
			parts[part] = slide
		})
		options := nativeMutationExtractOptions()
		options.AllowInheritedTextPreview = true
		deck, err := ExtractNativePPTX(data, options)
		if err != nil {
			t.Fatal(err)
		}
		e := nativeFixtureAutoShapes(deck.Slides[0])[0]
		if e.Paragraphs == nil || len(*e.Paragraphs) != 4 {
			t.Fatalf("expected original, split, continuation and blank paragraphs: %+v %+v", e.Paragraphs, e.Compatibility.Diagnostics)
		}
		first, continuation, blank := (*e.Paragraphs)[1], (*e.Paragraphs)[2], (*e.Paragraphs)[3]
		if first.Bullet == nil || !*first.Bullet || *first.Runs[0].Text != "First" || *first.Runs[0].FontSizeHundredthPt != 1800 {
			t.Fatalf("bulleted first line lost: %+v", first)
		}
		if continuation.Bullet == nil || *continuation.Bullet || *continuation.IndentEmu != 0 || *continuation.MarginLeftEmu != 342900 || *continuation.Runs[0].Text != " " {
			t.Fatalf("break continuation not projected at the paragraph margin: %+v", continuation)
		}
		if len(blank.Runs) != 1 || *blank.Runs[0].Text != " " || *blank.Runs[0].FontSizeHundredthPt != 1800 {
			t.Fatalf("empty paragraph blank line lost its end-mark metrics: %+v", blank)
		}
		codes := nativeDiagnosticCodes(e)
		if codes[nativeInheritedTextOmissionsCode] != 1 {
			t.Fatalf("omissions were not disclosed: %+v", e.Compatibility.Diagnostics)
		}
		for _, diagnostic := range e.Compatibility.Diagnostics {
			if diagnostic.Code == nativeInheritedTextOmissionsCode {
				for _, name := range []string{"a:br→continuation-paragraph", "empty-paragraph→blank-line", "a:lnSpc", "a:rPr@spc", "a:rPr@strike=noStrike", "a:t@xml:space=preserve-assumed", "a:endParaRPr"} {
					if !strings.Contains(diagnostic.Message, name) {
						t.Fatalf("disclosure lacks %s: %s", name, diagnostic.Message)
					}
				}
			}
		}
		if issues := ValidateNativePPTX(deck); len(issues) > 0 {
			t.Fatalf("invalid approximate contract: %+v", issues)
		}
		strictDeck, err := ExtractNativePPTX(data, nativeMutationExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		if p := nativeFixtureAutoShapes(strictDeck.Slides[0])[0].Paragraphs; p != nil && len(*p) > 0 {
			t.Fatal("strict extraction painted break/empty paragraph projections")
		}
	}
}

func TestNativeShapeStyleZeroIndexAndUnpaintedOutlineAreExact(t *testing.T) {
	d := nativeExtractDialect{drawing: nsDrawingTransitional, presentation: nsPresentationTransitional}
	for _, line := range []string{`<a:ln><a:noFill/></a:ln>`, `<a:ln w="12700"><a:noFill/></a:ln>`, `<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:noFill/><a:prstDash val="dash"/><a:round/></a:ln>`} {
		node, err := parseNativeXML([]byte(`<a:spPr xmlns:a="`+d.drawing+`">`+line+`</a:spPr>`), "test.xml")
		if err != nil {
			t.Fatal(err)
		}
		gaps := nativeShapeGapSet{}
		stroke, err := validateNativeAutoShapeLine(node, d, nativeResolvedTheme{}, false, &gaps)
		if err != nil || stroke != nil || len(gaps.values) != 0 {
			t.Fatalf("no-fill outline must be exact without a width: %s %v %+v", line, err, gaps.values)
		}
	}
	for _, line := range []string{`<a:ln w="abc"><a:noFill/></a:ln>`, `<a:ln cap="bad"><a:noFill/></a:ln>`, `<a:ln><a:noFill/><a:round/><a:bevel/></a:ln>`, `<a:ln><a:noFill>x</a:noFill></a:ln>`, `<a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>`} {
		node, err := parseNativeXML([]byte(`<a:spPr xmlns:a="`+d.drawing+`">`+line+`</a:spPr>`), "test.xml")
		if err != nil {
			t.Fatal(err)
		}
		gaps := nativeShapeGapSet{}
		if stroke, err := validateNativeAutoShapeLine(node, d, nativeResolvedTheme{}, false, &gaps); err == nil && stroke == nil && len(gaps.values) == 0 {
			t.Fatalf("malformed or painted outline accepted as unpainted: %s", line)
		}
	}
}

func TestNativeBulletFollowTextMarkersCancelInheritedBulletFont(t *testing.T) {
	build := func(layoutLevel string) []byte {
		return nativePlaceholderFixture(t, false, func(parts map[string]string) {
			parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `<p:ph idx="7"/>`, `<p:ph type="subTitle" idx="7"/>`, 1)
			layout := parts["relocated/layouts/layout.xml"]
			layout = strings.Replace(layout, `<p:ph type="body" idx="7"/>`, `<p:ph type="subTitle" idx="7"/>`, 1)
			layout = strings.Replace(layout, `<a:lstStyle><a:lvl1pPr marL="400000"/></a:lstStyle>`, `<a:lstStyle>`+layoutLevel+`</a:lstStyle>`, 1)
			parts["relocated/layouts/layout.xml"] = layout
			parts["relocated/masters/master.xml"] = strings.Replace(parts["relocated/masters/master.xml"], `<a:buChar char="▪"/>`, `<a:buFont typeface="Arial"/><a:buChar char="▪"/>`, 1)
		})
	}
	options := nativeTestExtractOptions()
	options.AllowInheritedTextPreview = true
	control, err := ExtractNativePPTX(build(`<a:lvl1pPr marL="400000"/>`), options)
	if err != nil || len(control.Slides[0].Elements) != 1 || control.Slides[0].Elements[0].Paragraphs == nil {
		t.Fatalf("control placeholder preview failed: %v %+v", err, control.Slides[0].Compatibility)
	}
	if p := (*control.Slides[0].Elements[0].Paragraphs)[0]; p.BulletFontFamily == nil || *p.BulletFontFamily != "Arial" || p.BulletCharacter == nil || *p.BulletCharacter != "▪" {
		t.Fatalf("master bullet font did not cascade: %+v", p)
	}
	deck, err := ExtractNativePPTX(build(`<a:lvl1pPr marL="400000"><a:buFontTx/></a:lvl1pPr>`), options)
	if err != nil || len(deck.Slides[0].Elements) != 1 || deck.Slides[0].Elements[0].Paragraphs == nil {
		t.Fatalf("buFontTx placeholder preview failed: %v %+v", err, deck.Slides[0].Compatibility)
	}
	element := deck.Slides[0].Elements[0]
	if p := (*element.Paragraphs)[0]; p.BulletFontFamily != nil || p.BulletCharacter == nil || *p.BulletCharacter != "▪" {
		t.Fatalf("buFontTx did not cancel the inherited bullet font: %+v", p)
	}
	disclosed := false
	for _, diagnostic := range element.Compatibility.Diagnostics {
		disclosed = disclosed || (diagnostic.Code == nativeInheritedTextOmissionsCode && strings.Contains(diagnostic.Message, "a:buFontTx"))
	}
	if !disclosed {
		t.Fatalf("buFontTx was not disclosed: %+v", element.Compatibility.Diagnostics)
	}
	if issues := ValidateNativePPTX(deck); len(issues) > 0 {
		t.Fatalf("invalid native: %+v", issues)
	}
	d := nativeExtractDialect{drawing: nsDrawingTransitional, presentation: nsPresentationTransitional}
	el := func(local string) *nativeXMLNode {
		return &nativeXMLNode{Name: xml.Name{Space: d.drawing, Local: local}}
	}
	merged := mergeNativeStyleNodes(&nativeXMLNode{Children: []*nativeXMLNode{el("buFont"), el("buClr"), el("buSzPct")}}, &nativeXMLNode{Children: []*nativeXMLNode{el("buFontTx"), el("buClrTx"), el("buSzTx")}}, d)
	names := []string{}
	for _, child := range merged.Children {
		names = append(names, child.Name.Local)
	}
	if strings.Join(names, ",") != "buFontTx,buClrTx,buSzTx" {
		t.Fatalf("follow-text markers did not replace their slots: %v", names)
	}
	if got := nativeWithoutBulletTextMarkers(merged, d); len(got.Children) != 0 {
		t.Fatalf("markers leaked into the projection: %+v", got.Children)
	}
}

func TestNativeShapeStyleZeroIndexResolvesToNoPaint(t *testing.T) {
	d := nativeExtractDialect{drawing: nsDrawingTransitional, presentation: nsPresentationTransitional}
	themeRoot, err := parseNativeXML([]byte(`<a:theme xmlns:a="`+d.drawing+`"><a:themeElements><a:fmtScheme><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln></a:lnStyleLst></a:fmtScheme></a:themeElements></a:theme>`), "theme.xml")
	if err != nil {
		t.Fatal(err)
	}
	style, err := parseNativeXML([]byte(`<p:style xmlns:p="`+d.presentation+`" xmlns:a="`+d.drawing+`"><a:lnRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"/></p:style>`), "slide.xml")
	if err != nil {
		t.Fatal(err)
	}
	names := func(node *nativeXMLNode) string {
		out := []string{}
		for _, child := range node.Children {
			out = append(out, child.Name.Local)
		}
		return strings.Join(out, ",")
	}
	for _, tc := range []struct{ properties, expect string }{
		{`<a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm>`, "xfrm,noFill,ln"},
		{`<a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>`, "xfrm,solidFill,ln"},
		{`<a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm><a:noFill/><a:ln><a:noFill/></a:ln>`, "xfrm,noFill,ln"},
		{`<a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>`, "xfrm,ln,noFill"},
	} {
		properties, err := parseNativeXML([]byte(`<p:spPr xmlns:p="`+d.presentation+`" xmlns:a="`+d.drawing+`">`+tc.properties+`</p:spPr>`), "slide.xml")
		if err != nil {
			t.Fatal(err)
		}
		sourceChildren := len(properties.Children)
		resolved, err := resolveNativeShapeStyle(properties, style, themeRoot, d, nativeResolvedTheme{})
		if err != nil {
			t.Fatalf("idx=0 style refused: %v (%s)", err, tc.properties)
		}
		if got := names(resolved); got != tc.expect {
			t.Fatalf("unexpected resolved paint %q for %s", got, tc.properties)
		}
		gaps := nativeShapeGapSet{}
		if fill := validateNativeAutoShapeFill(resolved, d, nativeResolvedTheme{}, &gaps); len(gaps.values) != 0 {
			t.Fatalf("resolved fill is not exact: %+v %v", gaps.values, fill)
		}
		if _, err := validateNativeAutoShapeLine(resolved, d, nativeResolvedTheme{}, false, &gaps); err != nil || len(gaps.values) != 0 {
			t.Fatalf("resolved line is not exact: %v %+v", err, gaps.values)
		}
		if len(properties.Children) != sourceChildren {
			t.Fatal("style resolution mutated the source properties")
		}
	}
	// Non-zero indexes still need their placeholder color and a bounded entry.
	for _, refs := range []string{`<a:lnRef idx="1"/><a:fillRef idx="0"/><a:effectRef idx="0"/><a:fontRef idx="minor"/>`, `<a:lnRef idx="0"/><a:fillRef idx="1000"/><a:effectRef idx="0"/><a:fontRef idx="minor"/>`, `<a:lnRef idx="0"/><a:fillRef idx="0"/><a:effectRef idx="1"/><a:fontRef idx="minor"/>`} {
		style, err := parseNativeXML([]byte(`<p:style xmlns:p="`+d.presentation+`" xmlns:a="`+d.drawing+`">`+refs+`</p:style>`), "slide.xml")
		if err != nil {
			t.Fatal(err)
		}
		properties, _ := parseNativeXML([]byte(`<p:spPr xmlns:p="`+d.presentation+`" xmlns:a="`+d.drawing+`"><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm></p:spPr>`), "slide.xml")
		if _, err := resolveNativeShapeStyle(properties, style, themeRoot, d, nativeResolvedTheme{}); err == nil {
			t.Fatalf("unqualified style reference accepted: %s", refs)
		}
	}
}

func TestNativeAuthoredLnSpcReductionTravelsOnlyInTheApproximateTier(t *testing.T) {
	const body = `<a:bodyPr><a:normAutofit fontScale="85000" lnSpcReduction="20000"/></a:bodyPr>`
	for _, strict := range []bool{false, true} {
		// Strict extraction still refuses normAutofit, so the field cannot appear.
		strictDeck, err := ExtractNativePPTX(nativeSourceFrameFixture(t, strict, body), nativeMutationExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		if element := strictDeck.Slides[0].Elements[0]; element.TextBody != nil {
			t.Fatalf("strict extraction emitted a text body for normAutofit: %+v", element.TextBody)
		}

		options := nativeMutationExtractOptions()
		options.AllowSourceFrameAutoFitPreview = true
		deck, err := ExtractNativePPTX(nativeSourceFrameFixture(t, strict, body), options)
		if err != nil {
			t.Fatal(err)
		}
		element := deck.Slides[0].Elements[0]
		if element.TextBody == nil || element.TextBody.LineSpacingReductionPercent1000 == nil || *element.TextBody.LineSpacingReductionPercent1000 != 20000 {
			t.Fatalf("approximate extraction did not carry the authored lnSpcReduction: %+v", element.TextBody)
		}
		for _, diagnostic := range element.Compatibility.Diagnostics {
			if diagnostic.Code != nativeAuthoredAutoFitCode {
				continue
			}
			if !strings.Contains(diagnostic.Message, "reduces the line pitch by the authored lnSpcReduction=20%") {
				t.Fatalf("disclosure still claims the reduction is unapplied: %s", diagnostic.Message)
			}
			if strings.Contains(diagnostic.Message, "outside the native v1 layout contract") {
				t.Fatalf("disclosure still claims the reduction is unapplied: %s", diagnostic.Message)
			}
		}
		if issues := ValidateNativePPTX(deck); len(issues) > 0 {
			t.Fatalf("invalid approximate contract: %+v", issues)
		}
	}
}

func TestNativeAuthoredLnSpcReductionOmittedWhenAbsentOrZero(t *testing.T) {
	options := nativeMutationExtractOptions()
	options.AllowSourceFrameAutoFitPreview = true
	for _, body := range []string{
		`<a:bodyPr><a:normAutofit fontScale="85000"/></a:bodyPr>`,
		`<a:bodyPr><a:normAutofit fontScale="85000" lnSpcReduction="0"/></a:bodyPr>`,
		`<a:bodyPr numCol="3" spcCol="108000"/>`,
	} {
		deck, err := ExtractNativePPTX(nativeSourceFrameFixture(t, false, body), options)
		if err != nil {
			t.Fatalf("%s: %v", body, err)
		}
		element := deck.Slides[0].Elements[0]
		if element.TextBody == nil || element.TextBody.LineSpacingReductionPercent1000 != nil {
			t.Fatalf("%s emitted a line-spacing reduction: %+v", body, element.TextBody)
		}
		if issues := ValidateNativePPTX(deck); len(issues) > 0 {
			t.Fatalf("%s: invalid contract: %+v", body, issues)
		}
	}
}

func TestNativeAuthoredLnSpcReductionContractRules(t *testing.T) {
	options := nativeMutationExtractOptions()
	options.AllowSourceFrameAutoFitPreview = true
	extract := func(t *testing.T) NativePPTXDeck {
		t.Helper()
		deck, err := ExtractNativePPTX(nativeSourceFrameFixture(t, false, `<a:bodyPr><a:normAutofit fontScale="85000" lnSpcReduction="20000"/></a:bodyPr>`), options)
		if err != nil {
			t.Fatal(err)
		}
		return deck
	}
	for _, value := range []int64{1, 50000, 99999} {
		deck := extract(t)
		deck.Slides[0].Elements[0].TextBody.LineSpacingReductionPercent1000 = int64Pointer(value)
		if issues := ValidateNativePPTX(deck); len(issues) > 0 {
			t.Fatalf("reduction %d was rejected: %+v", value, issues)
		}
	}
	for _, value := range []int64{0, -1, 100000, 2147483647} {
		deck := extract(t)
		deck.Slides[0].Elements[0].TextBody.LineSpacingReductionPercent1000 = int64Pointer(value)
		if len(ValidateNativePPTX(deck)) == 0 {
			t.Fatalf("reduction %d validated", value)
		}
	}
	// The field is a read-only projection: it may not outlive its disclosure.
	deck := extract(t)
	deck.Slides[0].Elements[0].Compatibility.Diagnostics = nil
	deck.Slides[0].Elements[0].Compatibility.Status = NativeCompatibilityStatusEditable
	if len(ValidateNativePPTX(deck)) == 0 {
		t.Fatal("an editable element kept the authored line-spacing reduction")
	}
	deck = extract(t)
	for index := range deck.Slides[0].Elements[0].Compatibility.Diagnostics {
		if deck.Slides[0].Elements[0].Compatibility.Diagnostics[index].Code == nativeAuthoredAutoFitCode {
			deck.Slides[0].Elements[0].Compatibility.Diagnostics[index].Code = nativeTextColumnsOmittedCode
		}
	}
	if len(ValidateNativePPTX(deck)) == 0 {
		t.Fatal("the column disclosure authorized a line-pitch reduction")
	}
}
