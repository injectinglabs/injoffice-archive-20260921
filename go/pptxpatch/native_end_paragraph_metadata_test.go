package pptxpatch

import (
	"bytes"
	"encoding/xml"
	"strings"
	"testing"
)

func TestNativeEndMarkNeedsFinalVisibleRun(t *testing.T) {
	mark := &nativeXMLNode{Name: xml.Name{Space: nsDrawingTransitional, Local: "endParaRPr"}}
	for _, p := range []NativeParagraph{{}, {Runs: []NativeTextRun{{Text: stringPointer("")}}}, {Runs: []NativeTextRun{{Text: stringPointer(" ")}}}, {Runs: []NativeTextRun{{Text: stringPointer("visible")}, {Text: stringPointer("")}}}} {
		if qualifyNativeEndParagraphMetadata(mark, p) == nil {
			t.Fatal("empty final run acquired end mark metrics")
		}
	}
}

func TestNativeEndParagraphMetadataPreservedWithoutMutation(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, tc := range []struct {
			end string
			ok  bool
		}{
			{`<a:endParaRPr lang="tr-TR" dirty="0"/>`, true},
			{`<a:endParaRPr lang="tr-TR"/>`, true},
			{`<a:endParaRPr dirty="true" smtClean="false"/>`, true},
			{`<a:endParaRPr/>`, true},
			{`<a:endParaRPr lang="en-US"/>`, false},
			{`<a:endParaRPr lang="tr_TR"/>`, false},
			{`<a:endParaRPr dirty="yes"/>`, false},
			{`<a:endParaRPr sz="2400"/>`, false},
			{`<a:endParaRPr b="0"/>`, false},
			{`<a:endParaRPr><a:latin typeface="Aptos"/></a:endParaRPr>`, false},
			{`<a:endParaRPr>text</a:endParaRPr>`, false},
			{`<a:endParaRPr/><a:endParaRPr/>`, false},
			{`<a:endParaRPr xmlns:x="urn:spoof" x:dirty="0"/>`, false},
		} {
			input := nativeStyledTextFixture(t, strict, "", func(parts map[string]string) {
				part := "relocated/slides/slide-a.xml"
				parts[part] = strings.Replace(parts[part], `<a:defRPr b="0"`, `<a:defRPr lang="tr-TR" b="0"`, 1)
				parts[part] = strings.Replace(parts[part], `</a:p>`, tc.end+`</a:p>`, 1)
			})
			before := bytes.Clone(input)
			deck, err := ExtractNativePPTX(input, nativeMutationExtractOptions())
			if err != nil {
				if !tc.ok {
					continue
				}
				t.Fatal(err)
			}
			found := false
			for _, element := range deck.Slides[0].Elements {
				if element.Paragraphs == nil {
					continue
				}
				found = true
				if !tc.ok {
					t.Fatalf("accepted active end mark %s", tc.end)
				}
				if element.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
					t.Fatal("end mark became editable")
				}
				if (*element.Paragraphs)[0].Runs[0].Language == nil || *(*element.Paragraphs)[0].Runs[0].Language != "tr-TR" {
					t.Fatal("language changed")
				}
				paragraphs := nativeMutationParagraphs("Replacement")
				_, err = ApplyNativePPTXMutations(input, NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{{OperationID: "replace", Kind: NativePPTXReplaceText, ElementID: element.ID, ExpectedFingerprintSHA256: element.Source.FingerprintSHA256, Paragraphs: &paragraphs}}})
				if err == nil {
					t.Fatal("end mark source could be removed by replacement")
				}
			}
			if found != tc.ok {
				t.Fatalf("projection=%v for %s", found, tc.end)
			}
			if !bytes.Equal(before, input) {
				t.Fatal("source mutated")
			}
		}
	}
}
