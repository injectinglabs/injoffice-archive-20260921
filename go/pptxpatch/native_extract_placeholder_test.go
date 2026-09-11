package pptxpatch

import (
	"bytes"
	"fmt"
	"os"
	"strings"
	"testing"
)

func nativePlaceholderFixture(t *testing.T, strict bool, customize func(map[string]string)) []byte {
	return nativeStyledTextFixture(t, strict, "▪", func(parts map[string]string) {
		drawing, presentation := nsDrawingTransitional, nsPresentationTransitional
		if strict {
			drawing, presentation = nsDrawingStrict, nsPresentationStrict
		}
		slide := parts["relocated/slides/slide-a.xml"]
		start, end := strings.Index(slide, "<a:lstStyle>"), strings.Index(slide, "</a:lstStyle>")+len("</a:lstStyle>")
		list := slide[start:end]
		slide = slide[:start] + `<a:lstStyle/>` + slide[end:]
		slide = strings.Replace(slide, `<p:cNvSpPr txBox="1"/><p:nvPr/>`, `<p:cNvSpPr/><p:nvPr><p:ph idx="7"/></p:nvPr>`, 1)
		slide = strings.Replace(slide, `<p:spPr><a:xfrm><a:off x="914400" y="457200"/><a:ext cx="4572000" cy="914400"/></a:xfrm></p:spPr>`, `<p:spPr/>`, 1)
		parts["relocated/slides/slide-a.xml"] = slide
		shape := func(id, ph, properties, styles string) string {
			return `<p:sp><p:nvSpPr><p:cNvPr id="` + id + `" name="Inherited"/><p:cNvSpPr/><p:nvPr>` + ph + `</p:nvPr></p:nvSpPr><p:spPr>` + properties + `</p:spPr><p:txBody><a:bodyPr/>` + styles + `<a:p/></p:txBody></p:sp>`
		}
		layout := parts["relocated/layouts/layout.xml"]
		layout = strings.Replace(layout, `<p:sldLayout `, fmt.Sprintf(`<p:sldLayout xmlns:a="%s" `, drawing), 1)
		layout = strings.Replace(layout, `</p:spTree>`, shape("3", `<p:ph type="body" idx="7"/>`, "", `<a:lstStyle><a:lvl1pPr marL="400000"/></a:lstStyle>`)+`</p:spTree>`, 1)
		parts["relocated/layouts/layout.xml"] = layout
		master := nativeExactMasterWithColorMapXML(presentation)
		master = strings.Replace(master, `<p:sldMaster `, fmt.Sprintf(`<p:sldMaster xmlns:a="%s" `, drawing), 1)
		master = strings.Replace(master, `</p:spTree>`, shape("4", `<p:ph type="body" idx="99"/>`, `<a:xfrm><a:off x="914400" y="457200"/><a:ext cx="4572000" cy="914400"/></a:xfrm>`, `<a:lstStyle/>`)+`</p:spTree>`, 1)
		master = strings.Replace(master, `</p:sldMaster>`, `<p:txStyles>`+strings.Replace(strings.Replace(list, `a:lstStyle`, `p:bodyStyle`, 2), `algn="r"`, `algn="l"`, 1)+`</p:txStyles></p:sldMaster>`, 1)
		parts["relocated/masters/master.xml"] = master
		if customize != nil {
			customize(parts)
		}
	})
}

func TestNativePlaceholderRelationshipInheritance(t *testing.T) {
	for _, strict := range []bool{false, true} {
		input := nativePlaceholderFixture(t, strict, nil)
		before := append([]byte(nil), input...)
		deck, err := ExtractNativePPTX(input, nativeTestExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		if len(deck.Slides[0].Elements) != 1 {
			t.Fatalf("placeholder missing: %+v", deck.Slides[0].Compatibility)
		}
		element := deck.Slides[0].Elements[0]
		if element.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || element.Placeholder == nil || *element.Placeholder != "body" || element.Paragraphs == nil {
			t.Fatalf("inheritance failed: %+v", element)
		}
		paragraph := (*element.Paragraphs)[0]
		if *element.Transform.X != 914400 || *element.Transform.Y != 457200 || *paragraph.MarginLeftEmu != 400000 || *paragraph.BulletCharacter != "▪" || *paragraph.Runs[0].FontFamily != "Calibri" || *paragraph.Runs[0].FontSizeHundredthPt != 2400 {
			t.Fatalf("inheritance precedence failed: %+v %+v", element.Transform, paragraph)
		}
		if !bytes.Equal(input, before) {
			t.Fatal("source mutated")
		}
		if issues := ValidateNativePPTX(deck); len(issues) > 0 {
			t.Fatalf("invalid native: %+v", issues)
		}
	}
}

func TestNativeAncestorPromptTextIsNotSlideContent(t *testing.T) {
	for _, strict := range []bool{false, true} {
		input := nativePlaceholderFixture(t, strict, func(parts map[string]string) {
			for _, part := range []string{"relocated/layouts/layout.xml", "relocated/masters/master.xml"} {
				parts[part] = strings.Replace(parts[part], `<a:p/>`, `<a:p><a:pPr/><a:r><a:rPr/><a:t>Master authoring prompt</a:t></a:r><a:endParaRPr/></a:p>`, 1)
			}
		})
		before := append([]byte(nil), input...)
		deck, err := ExtractNativePPTX(input, nativeTestExtractOptions())
		if err != nil || len(deck.Slides[0].Elements) != 1 {
			t.Fatalf("unstyled prompt refused: %v %+v", err, deck.Slides[0].Compatibility)
		}
		element := deck.Slides[0].Elements[0]
		if element.Paragraphs == nil || *element.Transform.X != 914400 || *(*element.Paragraphs)[0].MarginLeftEmu != 400000 {
			t.Fatal("prompt changed inherited source geometry/style")
		}
		for _, p := range *element.Paragraphs {
			for _, r := range p.Runs {
				if r.Text != nil && strings.Contains(*r.Text, "Master authoring prompt") {
					t.Fatal("ancestor prompt leaked into slide content")
				}
			}
		}
		if !bytes.Equal(input, before) {
			t.Fatal("source package changed")
		}
	}
}

func TestNativeAncestorPromptUnknownStylesRemainRefused(t *testing.T) {
	for _, prompt := range []string{
		`<a:p><a:pPr algn="r"/><a:r><a:t>Prompt</a:t></a:r></a:p>`,
		`<a:p><a:r><a:rPr sz="1000"/><a:t>Prompt</a:t></a:r></a:p>`,
		`<a:p><a:r><a:t>Prompt</a:t><a:t>Duplicate</a:t></a:r></a:p>`,
		`<a:p><a:fld id="field"><a:t>Prompt</a:t></a:fld></a:p>`,
		`<a:p><a:endParaRPr b="1"/></a:p>`,
	} {
		input := nativePlaceholderFixture(t, false, func(parts map[string]string) {
			part := "relocated/layouts/layout.xml"
			parts[part] = strings.Replace(parts[part], `<a:p/>`, prompt, 1)
		})
		deck, err := ExtractNativePPTX(input, nativeTestExtractOptions())
		if err == nil && len(deck.Slides[0].Elements) != 0 {
			t.Fatalf("unknown prompt styling projected: %s", prompt)
		}
	}
}

func TestNativeAncestorPromptKeepsExplicitNestedListCascade(t *testing.T) {
	input := nativePlaceholderFixture(t, false, func(parts map[string]string) {
		for _, part := range []string{"relocated/layouts/layout.xml", "relocated/masters/master.xml"} {
			parts[part] = strings.ReplaceAll(parts[part], "lvl1pPr", "lvl3pPr")
			parts[part] = strings.Replace(parts[part], `<a:p/>`, `<a:p><a:r><a:t>Authoring prompt only</a:t></a:r></a:p>`, 1)
		}
		part := "relocated/slides/slide-a.xml"
		parts[part] = strings.Replace(parts[part], `lvl="0"`, `lvl="2"`, 1)
	})
	deck, err := ExtractNativePPTX(input, nativeTestExtractOptions())
	if err != nil || len(deck.Slides[0].Elements) != 1 {
		t.Fatalf("nested inheritance failed: %v", err)
	}
	p := (*deck.Slides[0].Elements[0].Paragraphs)[0]
	if *p.Level != 2 || *p.MarginLeftEmu != 400000 || *p.BulletCharacter != "▪" || *p.Runs[0].FontSizeHundredthPt != 2400 {
		t.Fatalf("wrong nested cascade: %+v", p)
	}
}

func TestNativeExplicitPromptParagraphDefaults(t *testing.T) {
	for _, strict := range []bool{false, true} {
		input := nativePlaceholderFixture(t, strict, func(parts map[string]string) {
			part := "relocated/masters/master.xml"
			// Move an existing exact level style into an explicitly indexed prompt.
			start := strings.Index(parts[part], `<a:lvl1pPr`)
			end := strings.Index(parts[part], `</a:lvl1pPr>`) + len(`</a:lvl1pPr>`)
			style := strings.ReplaceAll(parts[part][start:end], "lvl1pPr", "pPr")
			style = strings.Replace(style, `<a:pPr `, `<a:pPr lvl="0" `, 1)
			parts[part] = parts[part][:start] + parts[part][end:]
			parts[part] = strings.Replace(parts[part], `sz="1800"`, ``, 1)
			parts[part] = strings.Replace(parts[part], `<a:p/>`, `<a:p>`+style+`<a:r><a:t>Do not paint this prompt</a:t></a:r></a:p>`, 1)
		})
		deck, err := ExtractNativePPTX(input, nativeTestExtractOptions())
		if err != nil || len(deck.Slides[0].Elements) != 1 {
			t.Fatalf("explicit prompt style failed: %v", err)
		}
		p := (*deck.Slides[0].Elements[0].Paragraphs)[0]
		if *p.BulletCharacter != "▪" || *p.IndentEmu != -100000 || *p.MarginLeftEmu != 400000 || *p.Runs[0].FontSizeHundredthPt != 2400 {
			t.Fatalf("prompt style or local override lost: %+v", p)
		}
		for _, run := range p.Runs {
			if strings.Contains(*run.Text, "prompt") {
				t.Fatal("prompt leaked into content")
			}
		}
		if output := os.Getenv("INJOFFICE_PPTX_PREVIEW_FIXTURE"); output != "" && !strict {
			for _, part := range []string{"relocated/slides/slide-a.xml", "relocated/layouts/layout.xml", "relocated/masters/master.xml", "relocated/themes/theme.xml"} {
				xml := string(chartZipEntry(t, input, part))
				for _, change := range [][2]string{{"Calibri", "DejaVu Sans"}, {`b="1"`, `b="0"`}, {`i="1"`, `i="0"`}, {`marL="400000"`, `marL="600000"`}, {`marL="300000"`, `marL="600000"`}, {`indent="-100000"`, `indent="-500000"`}} {
					xml = strings.ReplaceAll(xml, change[0], change[1])
				}
				input = replaceChartZipEntry(t, input, part, []byte(xml))
			}
			if err := os.WriteFile(strings.TrimSuffix(output, ".pptx")+"-inherited.pptx", input, 0600); err != nil {
				t.Fatal(err)
			}
		}
	}
}

func TestNativeCompetingPromptDefaultsRefuse(t *testing.T) {
	for _, prompt := range []string{
		`<a:p><a:pPr lvl="0" marL="999"/></a:p>`,
		`<a:p><a:pPr lvl="0"><a:buNone/></a:pPr></a:p>`,
		`<a:p><a:pPr lvl="0"><a:defRPr sz="2222"/></a:pPr></a:p>`,
		`<a:p><a:pPr lvl="0"/></a:p><a:p><a:pPr lvl="0"/></a:p>`,
		`<a:p><a:pPr lvl="9"/></a:p>`,
	} {
		input := nativePlaceholderFixture(t, false, func(parts map[string]string) {
			part := "relocated/layouts/layout.xml"
			parts[part] = strings.Replace(parts[part], `<a:p/>`, prompt, 1)
		})
		deck, err := ExtractNativePPTX(input, nativeTestExtractOptions())
		if err == nil && len(deck.Slides[0].Elements) != 0 {
			t.Fatalf("competing prompt style accepted: %s", prompt)
		}
	}
}

func TestNativeTitlePlaceholderLocalOverrideAndSourceAnchor(t *testing.T) {
	var sourceShape string
	input := nativePlaceholderFixture(t, false, func(parts map[string]string) {
		for _, part := range []string{"relocated/layouts/layout.xml", "relocated/masters/master.xml"} {
			parts[part] = strings.ReplaceAll(strings.ReplaceAll(parts[part], `type="body"`, `type="title"`), "bodyStyle", "titleStyle")
		}
		part := "relocated/slides/slide-a.xml"
		slide := strings.Replace(parts[part], `<p:spPr/>`, `<p:spPr><a:xfrm><a:off x="1828800" y="914400"/><a:ext cx="4572000" cy="914400"/></a:xfrm></p:spPr>`, 1)
		parts[part] = slide
		sourceShape = slide[strings.Index(slide, "<p:sp>") : strings.Index(slide, "</p:sp>")+len("</p:sp>")]
	})
	deck, err := ExtractNativePPTX(input, nativeTestExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	if len(deck.Slides[0].Elements) != 1 {
		t.Fatal("title was not resolved")
	}
	element := deck.Slides[0].Elements[0]
	if *element.Placeholder != "title" || *element.Transform.X != 1828800 || *element.Transform.Y != 914400 {
		t.Fatalf("local transform did not override: %+v", element)
	}
	if element.Source == nil || element.Source.FingerprintSHA256 != nativeSHA256([]byte(sourceShape)) {
		t.Fatal("source anchor is not the unchanged original placeholder subtree")
	}
}

func TestNativeDuplicatePlaceholderTargetsStayClosed(t *testing.T) {
	for _, part := range []string{"relocated/layouts/layout.xml", "relocated/masters/master.xml"} {
		input := nativePlaceholderFixture(t, false, func(parts map[string]string) {
			source := parts[part]
			shape := source[strings.Index(source, "<p:sp>") : strings.Index(source, "</p:sp>")+len("</p:sp>")]
			shape = strings.Replace(shape, `id="3"`, `id="30"`, 1)
			shape = strings.Replace(shape, `id="4"`, `id="40"`, 1)
			parts[part] = strings.Replace(source, "</p:spTree>", shape+"</p:spTree>", 1)
		})
		deck, err := ExtractNativePPTX(input, nativeTestExtractOptions())
		if err == nil && len(deck.Slides[0].Elements) > 0 {
			t.Fatalf("ambiguous target in %s was projected", part)
		}
	}
}

func TestNativeOmittedSlidePlaceholderTypeCannotWidenInheritedSubset(t *testing.T) {
	for _, kind := range []string{"obj", "subTitle", "ctrTitle", ""} {
		input := nativePlaceholderFixture(t, false, func(parts map[string]string) {
			for _, part := range []string{"relocated/layouts/layout.xml", "relocated/masters/master.xml"} {
				parts[part] = strings.ReplaceAll(parts[part], `type="body"`, `type="`+kind+`"`)
			}
		})
		deck, err := ExtractNativePPTX(input, nativeTestExtractOptions())
		if err == nil && len(deck.Slides[0].Elements) > 0 {
			t.Fatalf("omitted slide type widened subset to %q", kind)
		}
	}
}

func TestNativePlaceholderAmbiguityAndUnknownMetadataStayClosed(t *testing.T) {
	for _, mutation := range []struct{ part, from, to string }{
		{"layout", `idx="7"`, `idx="8"`},
		{"layout", `type="body"`, `type="title"`},
		{"layout", `<p:ph type="body" idx="7"/>`, `<p:ph type="body" idx="7"/><p:ph type="body" idx="7"/>`},
		{"layout", `<a:bodyPr/>`, `<a:bodyPr unknown="1"/>`},
		{"layout", `<p:spPr></p:spPr>`, `<p:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr>`},
		{"master", `<a:lvl1pPr algn="l"`, `<a:lvl1pPr unknown="1" algn="l"`},
	} {
		input := nativePlaceholderFixture(t, false, func(parts map[string]string) {
			part := "relocated/" + mutation.part + "s/" + mutation.part + ".xml"
			source := parts[part]
			if !strings.Contains(source, mutation.from) {
				t.Fatalf("bad mutation %s", mutation.from)
			}
			parts[part] = strings.Replace(source, mutation.from, mutation.to, 1)
		})
		deck, err := ExtractNativePPTX(input, nativeTestExtractOptions())
		if err != nil {
			continue
		}
		for _, element := range deck.Slides[0].Elements {
			if element.Compatibility.Status != NativeCompatibilityStatusRefused {
				t.Fatalf("unresolved placeholder emitted: %+v", element)
			}
		}
	}
}
