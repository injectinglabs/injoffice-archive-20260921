package pptxpatch

import (
	"bytes"
	"reflect"
	"strings"
	"testing"
)

// Wrapper write-back for text.replace: ApplyNativePPTXMutations replaces the
// <a:p> range and must keep <a:bodyPr>, <a:lstStyle>, p:nvSpPr/p:spPr,
// slide-local comments, and untouched OPC parts.
//
// encodeNativeParagraphs currently rebuilds <a:p> from native JSON (explicit
// algn, buNone, sRGB, single latin). These tests do not claim run-level
// schemeClr, dual latin/ea, u="dbl", or inherited bullets survive inside the
// rebuilt paragraphs.

const (
	nativeWritebackBodyPr    = `<a:bodyPr wrap="none" lIns="11111" rIns="22222" tIns="33333" bIns="44444" anchor="b"><a:noAutofit/></a:bodyPr>`
	nativeWritebackLstStyle  = `<a:lstStyle><a:lvl1pPr marL="228600"><a:buChar char="q"/><a:defRPr sz="1800"/></a:lvl1pPr></a:lstStyle>`
	nativeWritebackShapeName = "Writeback Wrapper Title"
	nativeWritebackNvSpPr    = `<p:nvSpPr><p:cNvPr id="2" name="Writeback Wrapper Title"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>`
	nativeWritebackXfrm      = `<a:xfrm><a:off x="123456" y="654321"/><a:ext cx="3333333" cy="2222222"/></a:xfrm>`
	nativeWritebackSpPr      = `<p:spPr>` + nativeWritebackXfrm + `</p:spPr>`
	nativeWritebackMarker    = `<!--opaque-slide-byte-marker-->`
	nativeWritebackText      = "Writeback replacement text"
)

func nativeWritebackFidelityFixture(t *testing.T, strict bool) (original []byte, siblingXML string) {
	t.Helper()
	siblingXML = nativeAutoShapeXML(3, "Untouched Sibling Shape", "rect", `<a:solidFill><a:srgbClr val="DDEEFF"/></a:solidFill>`, nativeAutoShapeSolidLine("12700", "flat", `<a:round/>`, "112233"), "")
	original = nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, mutate: func(parts map[string]string) {
		slide := parts["relocated/slides/slide-a.xml"]
		slide = strings.Replace(slide, `name="Title"`, `name="`+nativeWritebackShapeName+`"`, 1)
		slide = strings.Replace(slide, `<a:xfrm><a:off x="914400" y="457200"/><a:ext cx="4572000" cy="914400"/></a:xfrm>`, nativeWritebackXfrm, 1)
		slide = strings.Replace(slide, `<a:bodyPr/>`, nativeWritebackBodyPr, 1)
		slide = strings.Replace(slide, `<a:lstStyle/>`, nativeWritebackLstStyle, 1)
		slide = strings.Replace(slide, `</p:spTree>`, siblingXML+nativeWritebackMarker+`</p:spTree>`, 1)
		parts["relocated/slides/slide-a.xml"] = slide
	}})
	return original, siblingXML
}

func nativeWritebackTextReplaceRequest(t *testing.T, original []byte) (NativePPTXMutationRequest, NativeElement) {
	t.Helper()
	before, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
	if err != nil {
		t.Fatalf("extract write-back fixture: %v", err)
	}
	if len(before.Slides) != 1 || len(before.Slides[0].Elements) == 0 {
		t.Fatalf("write-back fixture topology: %#v", before.Slides)
	}
	text := before.Slides[0].Elements[0]
	if text.Kind != NativeElementKindText || text.Source == nil || text.Compatibility.Status != NativeCompatibilityStatusEditable {
		t.Fatalf("write-back text target is not exact/editable: %#v", text)
	}
	paragraphs := nativeMutationParagraphs(nativeWritebackText)
	return NativePPTXMutationRequest{ExpectedSourceRevision: *before.SourceRevision, Operations: []NativePPTXMutation{{
		OperationID: "writeback-replace", Kind: NativePPTXReplaceText, ElementID: text.ID,
		ExpectedFingerprintSHA256: text.Source.FingerprintSHA256, Paragraphs: &paragraphs,
	}}}, text
}

func TestApplyNativePPTXMutationsPreservesTextReplaceWrappers(t *testing.T) {
	t.Parallel()
	for _, strict := range []bool{false, true} {
		strict := strict
		t.Run(map[bool]string{false: "transitional", true: "strict"}[strict], func(t *testing.T) {
			t.Parallel()
			original, siblingXML := nativeWritebackFidelityFixture(t, strict)
			request, textElement := nativeWritebackTextReplaceRequest(t, original)
			produced, err := ApplyNativePPTXMutations(original, request)
			if err != nil {
				t.Fatalf("apply text.replace: %v", err)
			}
			if bytes.Equal(produced, original) {
				t.Fatal("text.replace returned the original package")
			}

			beforePackage, err := openNativeExtractPackage(original)
			if err != nil {
				t.Fatalf("open source OPC: %v", err)
			}
			afterPackage, err := openNativeExtractPackage(produced)
			if err != nil {
				t.Fatalf("open produced OPC: %v", err)
			}
			slidePart := textElement.Source.PartName
			for part, payload := range beforePackage.parts {
				if part == slidePart {
					continue
				}
				if !bytes.Equal(payload, afterPackage.parts[part]) {
					t.Fatalf("untouched OPC part %q changed", part)
				}
			}
			independentBefore := independentPPTXEntries(t, original)
			independentAfter := independentPPTXEntries(t, produced)
			for name, payload := range independentBefore {
				if name == slidePart {
					continue
				}
				if !bytes.Equal(payload, independentAfter[name]) {
					t.Fatalf("independent consumer: untouched OPC part %q changed", name)
				}
			}

			slideBytes := afterPackage.parts[slidePart]
			for _, fragment := range []string{
				nativeWritebackBodyPr,
				nativeWritebackLstStyle,
				nativeWritebackNvSpPr,
				nativeWritebackSpPr,
				nativeWritebackXfrm,
				nativeWritebackMarker,
				siblingXML,
				`<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>`,
			} {
				if !bytes.Contains(slideBytes, []byte(fragment)) {
					t.Fatalf("wrapper/neighbor %q was not preserved: %s", fragment, slideBytes)
				}
			}
			if !bytes.Contains(slideBytes, []byte(nativeWritebackText)) {
				t.Fatalf("replaced text is missing: %s", slideBytes)
			}
			if bytes.Contains(slideBytes, []byte("Hello ")) || bytes.Contains(slideBytes, []byte(">world<")) {
				t.Fatalf("original paragraph text was not replaced: %s", slideBytes)
			}

			afterOptions := nativeMutationExtractOptions()
			before, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
			if err != nil {
				t.Fatalf("re-extract source: %v", err)
			}
			afterOptions.Previous = &before
			after, err := ExtractNativePPTX(produced, afterOptions)
			if err != nil {
				t.Fatalf("reopen/extract produced package: %v", err)
			}
			afterText := after.Slides[0].Elements[0]
			if !nativeParagraphsEqual(*afterText.Paragraphs, *request.Operations[0].Paragraphs) {
				t.Fatalf("replaced paragraphs did not round trip: %#v", afterText.Paragraphs)
			}
			if !nativeTransformEqual(textElement.Transform, afterText.Transform) || !nativeStringPointerEqual(textElement.Name, afterText.Name) || textElement.Kind != afterText.Kind || textElement.Provenance != afterText.Provenance || !reflect.DeepEqual(textElement.Compatibility, afterText.Compatibility) || !nativePassthroughSemanticsEqual(textElement, afterText) {
				t.Fatalf("text.replace changed non-requested target state: before=%#v after=%#v", textElement, afterText)
			}
			if !reflect.DeepEqual(textElement.TextBody, afterText.TextBody) {
				t.Fatalf("text.replace changed bodyPr layout: before=%#v after=%#v", textElement.TextBody, afterText.TextBody)
			}
			if afterText.Name == nil || *afterText.Name != nativeWritebackShapeName {
				t.Fatalf("nvSpPr name did not survive: %#v", afterText.Name)
			}
		})
	}
}

func TestApplyNativePPTXMutationsTextReplaceIsDeterministic(t *testing.T) {
	t.Parallel()
	for _, strict := range []bool{false, true} {
		strict := strict
		t.Run(map[bool]string{false: "transitional", true: "strict"}[strict], func(t *testing.T) {
			t.Parallel()
			original, _ := nativeWritebackFidelityFixture(t, strict)
			request, _ := nativeWritebackTextReplaceRequest(t, original)
			first, err := ApplyNativePPTXMutations(original, request)
			if err != nil {
				t.Fatalf("first text.replace: %v", err)
			}
			second, err := ApplyNativePPTXMutations(original, request)
			if err != nil {
				t.Fatalf("second text.replace: %v", err)
			}
			if !bytes.Equal(first, second) {
				t.Fatalf("same text.replace from the same source produced %d vs %d distinct bytes", len(first), len(second))
			}
		})
	}
}
