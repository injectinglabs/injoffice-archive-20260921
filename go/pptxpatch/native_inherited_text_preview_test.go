package pptxpatch

import (
	"bytes"
	"encoding/xml"
	"strings"
	"testing"
)

func TestNativeInheritedTextPreviewSourceAndSafety(t *testing.T) {
	for _, strict := range []bool{false, true} {
		data := nativeShapeReferenceFixture(t, strict, nativeShapeStyleRefs, `lang="en-US" dirty="0">`, "", "Hi", nativeShapeReferenceFonts, func(parts map[string]string) {
			ns := nsDrawingTransitional
			if strict {
				ns = nsDrawingStrict
			}
			parts["relocated/deck.xml"] = strings.Replace(parts["relocated/deck.xml"], `</p:presentation>`, `<p:defaultTextStyle xmlns:a="`+ns+`"><a:defPPr><a:defRPr lang="en-US"/></a:defPPr><a:lvl1pPr algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:defRPr></a:lvl1pPr></p:defaultTextStyle></p:presentation>`, 1)
		})
		before := bytes.Clone(data)
		strictDeck, err := ExtractNativePPTX(data, nativeMutationExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		if p := nativeFixtureAutoShapes(strictDeck.Slides[0])[0].Paragraphs; p != nil && len(*p) > 0 {
			t.Fatal("strict text unexpectedly available")
		}
		options := nativeMutationExtractOptions()
		options.AllowInheritedTextPreview = true
		deck, err := ExtractNativePPTX(data, options)
		if err != nil {
			t.Fatal(err)
		}
		e := nativeFixtureAutoShapes(deck.Slides[0])[0]
		if e.Paragraphs == nil || len(*e.Paragraphs) != 1 {
			t.Fatalf("missing text: %+v", e.Compatibility)
		}
		run := (*e.Paragraphs)[0].Runs[0]
		if run.Text == nil || *run.Text != "Hi" || run.FontFamily == nil || *run.FontFamily != "Calibri" || run.Color == nil || *run.Color != "FFFFFF" || run.FontSizeHundredthPt == nil || *run.FontSizeHundredthPt != 1800 || run.Bold == nil || *run.Bold || run.Italic == nil || *run.Italic {
			t.Fatalf("wrong source policy %+v", run)
		}
		found := false
		for _, d := range e.Compatibility.Diagnostics {
			found = found || d.Code == nativeInheritedTextPreviewCode
		}
		if !found || e.Compatibility.Status == NativeCompatibilityStatusEditable {
			t.Fatal("missing read-only approximation")
		}
		if !bytes.Equal(data, before) {
			t.Fatal("source changed")
		}
		replacement := nativeMutationParagraphs("changed")
		_, err = resolveNativePPTXMutations(deck, []NativePPTXMutation{{OperationID: "edit", Kind: NativePPTXReplaceText, ElementID: e.ID, ExpectedFingerprintSHA256: e.Source.FingerprintSHA256, Paragraphs: &replacement}})
		if err == nil {
			t.Fatal("preview authorized mutation")
		}
	}
}

func TestNativeInheritedTextPreviewDroppedMetadataSafety(t *testing.T) {
	d := nativeExtractDialect{drawing: nsDrawingTransitional, presentation: nsPresentationTransitional}
	duplicateEnd := &nativeXMLNode{Children: []*nativeXMLNode{{Name: xml.Name{Space: d.drawing, Local: "lstStyle"}}, {Name: xml.Name{Space: d.drawing, Local: "p"}, Children: []*nativeXMLNode{{Name: xml.Name{Space: d.drawing, Local: "endParaRPr"}}, {Name: xml.Name{Space: d.drawing, Local: "endParaRPr"}}}}}}
	if _, err := (&nativeExtractor{}).inheritedTextPreview(duplicateEnd, nil, false, d); err == nil {
		t.Fatal("duplicate terminal elements hidden")
	}
	for _, field := range []string{"kern", "defTabSz", "rtl", "eaLnBrk", "latinLnBrk", "hangingPunct"} {
		node := &nativeXMLNode{Attrs: []xml.Attr{{Name: xml.Name{Local: field}, Value: "0"}, {Name: xml.Name{Local: field}, Value: "0"}}}
		if _, err := sanitizeNativeInheritedPreviewProperties(node, d, field != "kern", nativeResolvedTheme{}); err == nil {
			t.Fatalf("duplicate %s hidden", field)
		}
	}
	for _, name := range []string{"ea", "cs"} {
		for _, value := range []string{"bad text", ""} {
			node := &nativeXMLNode{Children: []*nativeXMLNode{{Name: xml.Name{Space: d.drawing, Local: name}, Attrs: []xml.Attr{{Name: xml.Name{Local: "typeface"}, Value: "+mn-ea"}}, Text: value}}}
			if value == "" {
				node.Children = append(node.Children, node.Children[0])
			}
			if _, err := sanitizeNativeInheritedPreviewProperties(node, d, false, nativeResolvedTheme{}); err == nil {
				t.Fatalf("unsafe %s hidden", name)
			}
		}
	}
	for _, node := range []*nativeXMLNode{{Text: "active"}, {Attrs: []xml.Attr{{Name: xml.Name{Local: "lang"}, Value: "bad_tag"}}}, {Attrs: []xml.Attr{{Name: xml.Name{Local: "sz"}, Value: "1200"}}}} {
		if err := validateNativeInheritedPreviewEnd(node, d); err == nil {
			t.Fatal("unsafe terminal metadata hidden")
		}
	}
}
