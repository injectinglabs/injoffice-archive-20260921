package pptxpatch

import (
	"bytes"
	"fmt"
	"strings"
	"testing"
)

const nativeShapeReferenceFonts = `<a:fontScheme name="Synthetic"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>`

func nativeShapeReferenceFixture(t *testing.T, strict bool, refs, runProps, list, text, fonts string, changes ...func(map[string]string)) []byte {
	t.Helper()
	shape := nativeAutoShapeXML(3, "Theme text", "rect", `<a:noFill/>`, nativeAutoShapeNoLine("flat", `<a:round/>`), "")
	body := `<p:txBody><a:bodyPr><a:noAutofit/></a:bodyPr><a:lstStyle>` + list + `</a:lstStyle><a:p><a:pPr algn="l" lvl="0"><a:buNone/></a:pPr><a:r><a:rPr ` + runProps + `</a:rPr><a:t>` + text + `</a:t></a:r></a:p></p:txBody>`
	shape = strings.Replace(shape, `</p:sp>`, refs+body+`</p:sp>`, 1)
	return nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, mutate: func(parts map[string]string) {
		parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, shape+`</p:spTree>`, 1)
		drawing := nsDrawingTransitional
		if strict {
			drawing = nsDrawingStrict
		}
		parts["relocated/themes/theme.xml"] = fmt.Sprintf(`<a:theme xmlns:a="%s" name="Synthetic"><a:themeElements>%s<a:fmtScheme name="Synthetic"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst></a:fmtScheme></a:themeElements></a:theme>`, drawing, fonts)
		for _, change := range changes {
			change(parts)
		}
	}})
}

func TestNativeShapeFontReferenceDoesNotGuessExternalDefaults(t *testing.T) {
	for _, tc := range []struct {
		name, part, close, markup string
		blocked                   bool
	}{
		{"empty presentation", "relocated/deck.xml", "</p:presentation>", `<p:defaultTextStyle/>`, false},
		{"presentation defaults", "relocated/deck.xml", "</p:presentation>", `<p:defaultTextStyle><a:lvl1pPr xmlns:a="` + nsDrawingTransitional + `"><a:defRPr><a:latin typeface="Different"/></a:defRPr></a:lvl1pPr></p:defaultTextStyle>`, true},
		{"duplicate presentation", "relocated/deck.xml", "</p:presentation>", `<p:defaultTextStyle/><p:defaultTextStyle/>`, true},
		{"foreign presentation", "relocated/deck.xml", "</p:presentation>", `<x:defaultTextStyle xmlns:x="urn:foreign"/>`, true},
		{"empty master", "relocated/masters/master.xml", "</p:sldMaster>", `<p:txStyles><p:otherStyle/></p:txStyles>`, false},
		{"master defaults", "relocated/masters/master.xml", "</p:sldMaster>", `<p:txStyles><p:otherStyle><a:lvl1pPr xmlns:a="` + nsDrawingTransitional + `"><a:defRPr><a:latin typeface="Different"/></a:defRPr></a:lvl1pPr></p:otherStyle></p:txStyles>`, true},
		{"duplicate master", "relocated/masters/master.xml", "</p:sldMaster>", `<p:txStyles/><p:txStyles/>`, true},
		{"foreign master", "relocated/masters/master.xml", "</p:sldMaster>", `<x:txStyles xmlns:x="urn:foreign"/>`, true},
	} {
		for _, explicit := range []bool{false, true} {
			t.Run(fmt.Sprintf("%s/explicit=%v", tc.name, explicit), func(t *testing.T) {
				props := `b="0" i="0" sz="1800">`
				if explicit {
					props += `<a:latin typeface="Arial"/><a:solidFill><a:srgbClr val="112233"/></a:solidFill>`
				}
				original := nativeShapeReferenceFixture(t, false, nativeShapeStyleRefs, props, "", "Hi", nativeShapeReferenceFonts, func(parts map[string]string) {
					parts[tc.part] = strings.Replace(parts[tc.part], tc.close, tc.markup+tc.close, 1)
				})
				deck, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
				if err != nil {
					if strings.Contains(tc.name, "duplicate") || strings.Contains(tc.name, "foreign") {
						return
					}
					t.Fatal(err)
				}
				e := nativeFixtureAutoShapes(deck.Slides[0])[0]
				hasText := e.Paragraphs != nil && len(*e.Paragraphs) > 0
				if hasText != (explicit || !tc.blocked) {
					t.Fatalf("external defaults misclassified: %+v", e.Compatibility)
				}
			})
		}
	}
}

func TestNativeShapeFontReferenceOwnedPreview(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, tc := range []struct {
			name, idx, props, list, family, color string
			used                                  bool
		}{
			{"minor", "minor", `b="0" i="0" sz="1800">`, "", "Calibri", "FFFFFF", true},
			{"major", "major", `b="0" i="0" sz="1800">`, "", "Calibri Light", "FFFFFF", true},
			{"font override", "minor", `b="0" i="0" sz="1800"><a:latin typeface="Arial"/>`, "", "Arial", "FFFFFF", true},
			{"color override", "minor", `b="0" i="0" sz="1800"><a:solidFill><a:srgbClr val="112233"/></a:solidFill>`, "", "Calibri", "112233", true},
			{"both override", "minor", `b="0" i="0" sz="1800"><a:latin typeface="Arial"/><a:solidFill><a:srgbClr val="112233"/></a:solidFill>`, "", "Arial", "112233", false},
			{"list override", "minor", `b="0" i="0" sz="1800">`, `<a:lvl1pPr><a:defRPr><a:latin typeface="Arial"/><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:defRPr></a:lvl1pPr>`, "Arial", "112233", false},
		} {
			t.Run(fmt.Sprintf("%v/%s", strict, tc.name), func(t *testing.T) {
				refs := strings.Replace(nativeShapeStyleRefs, `fontRef idx="minor"`, `fontRef idx="`+tc.idx+`"`, 1)
				original := nativeShapeReferenceFixture(t, strict, refs, tc.props, tc.list, "Hi", nativeShapeReferenceFonts)
				before := bytes.Clone(original)
				deck, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
				if err != nil {
					t.Fatal(err)
				}
				e := nativeFixtureAutoShapes(deck.Slides[0])[0]
				if e.Paragraphs == nil || len(*e.Paragraphs) != 1 || len((*e.Paragraphs)[0].Runs) != 1 {
					t.Fatalf("text missing: %+v", e.Compatibility)
				}
				run := (*e.Paragraphs)[0].Runs[0]
				if run.FontFamily == nil || *run.FontFamily != tc.family || run.Color == nil || *run.Color != tc.color || run.FontSizeHundredthPt == nil || *run.FontSizeHundredthPt != 1800 || run.Bold == nil || *run.Bold || run.Italic == nil || *run.Italic {
					t.Fatalf("source properties lost: %+v", run)
				}
				used := false
				for _, d := range e.Compatibility.Diagnostics {
					used = used || d.Code == "pptx.shape-font-reference-preview"
				}
				if used != tc.used || e.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
					t.Fatalf("projection authority wrong: %+v", e.Compatibility)
				}
				if issues := ValidateNativePPTX(deck); len(issues) > 0 {
					t.Fatal(issues)
				}
				paragraphs := nativeMutationParagraphs("Changed")
				for _, op := range []NativePPTXMutation{
					{OperationID: "text", Kind: NativePPTXReplaceText, ElementID: e.ID, ExpectedFingerprintSHA256: e.Source.FingerprintSHA256, Paragraphs: &paragraphs},
					{OperationID: "shape", Kind: NativePPTXUpdateAutoShape, ElementID: e.ID, ExpectedFingerprintSHA256: e.Source.FingerprintSHA256, AutoShape: &NativePPTXAutoShapeMutation{Transform: e.Transform, Preset: NativeShapePresetRect, Fill: e.Fill, Stroke: e.Stroke}},
				} {
					_, err = ApplyNativePPTXMutations(original, NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{op}})
					if err == nil || !strings.Contains(err.Error(), "preview-only") {
						t.Fatalf("mutation allowed: %v", err)
					}
				}
				if !bytes.Equal(original, before) {
					t.Fatal("source rewritten")
				}
			})
		}
	}
}

func TestNativeShapeFontReferenceRetainsRefusalBoundaries(t *testing.T) {
	for _, tc := range []struct{ name, refs, props, text, fonts string }{
		{"missing size", nativeShapeStyleRefs, `b="0" i="0">`, "Hi", nativeShapeReferenceFonts},
		{"missing bold", nativeShapeStyleRefs, `i="0" sz="1800">`, "Hi", nativeShapeReferenceFonts},
		{"missing font scheme", nativeShapeStyleRefs, `b="0" i="0" sz="1800">`, "Hi", ""},
		{"non Latin", nativeShapeStyleRefs, `b="0" i="0" sz="1800">`, "مرحبا", nativeShapeReferenceFonts},
		{"bad index", strings.Replace(nativeShapeStyleRefs, `fontRef idx="minor"`, `fontRef idx="none"`, 1), `b="0" i="0" sz="1800">`, "Hi", nativeShapeReferenceFonts},
		{"duplicate reference", strings.Replace(nativeShapeStyleRefs, `</p:style>`, `<a:fontRef idx="minor"><a:srgbClr val="FFFFFF"/></a:fontRef></p:style>`, 1), `b="0" i="0" sz="1800">`, "Hi", nativeShapeReferenceFonts},
		{"bad color", strings.Replace(nativeShapeStyleRefs, `val="FFFFFF"`, `val="NOTRGB"`, 1), `b="0" i="0" sz="1800">`, "Hi", nativeShapeReferenceFonts},
	} {
		t.Run(tc.name, func(t *testing.T) {
			original := nativeShapeReferenceFixture(t, false, tc.refs, tc.props, "", tc.text, tc.fonts)
			deck, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
			if err != nil {
				return
			}
			e := nativeFixtureAutoShapes(deck.Slides[0])[0]
			if e.Compatibility.Status != NativeCompatibilityStatusRefused && e.Paragraphs != nil && len(*e.Paragraphs) > 0 {
				t.Fatalf("unsupported text became qualified: %+v", e)
			}
		})
	}
}
