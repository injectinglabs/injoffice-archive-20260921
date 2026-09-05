package pptxpatch

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"encoding/xml"
	"io"
	"os"
	"reflect"
	"strconv"
	"strings"
	"testing"
	"time"
)

func TestApplyNativePPTXMutationsRoundTripsTextAndAutoShape(t *testing.T) {
	t.Parallel()
	for _, strict := range []bool{false, true} {
		strict := strict
		name := map[bool]string{false: "transitional", true: "strict"}[strict]
		t.Run(name, func(t *testing.T) {
			t.Parallel()
			shapeXML := nativeAutoShapeXML(3, "Exact Shape", "rect", `<a:solidFill><a:srgbClr val="DDEEFF"/></a:solidFill>`, nativeAutoShapeSolidLine("12700", "flat", `<a:round/>`, "112233"), "")
			original := nativeAutoShapeFixture(t, strict, shapeXML+`<!--opaque-slide-byte-marker-->`)
			before, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
			if err != nil {
				t.Fatalf("extract source: %v", err)
			}
			textElement := before.Slides[0].Elements[0]
			shapeElement := nativeFixtureAutoShapes(before.Slides[0])[0]
			paragraphs := nativeMutationParagraphs("  Exact & native <text>  ")
			fill := "A1B2C3"
			width := int64(25400)
			capValue := NativeStrokeCapSquare
			joinValue := NativeStrokeJoinMiter
			dashValue := NativeStrokeDashSolid
			miter := int64(900000)
			x, y, cx, cy := int64(-42), int64(84), int64(2000000), int64(750000)
			shape := NativePPTXAutoShapeMutation{
				Transform: NativeTransform{X: &x, Y: &y, Cx: &cx, Cy: &cy}, Preset: NativeShapePresetDiamond, Fill: &fill,
				Stroke: &NativeStroke{Color: "445566", WidthEMU: &width, Cap: &capValue, Join: &joinValue, Dash: &dashValue, MiterLimit: &miter},
			}
			request := NativePPTXMutationRequest{ExpectedSourceRevision: *before.SourceRevision, Operations: []NativePPTXMutation{
				{OperationID: "replace-title", Kind: NativePPTXReplaceText, ElementID: textElement.ID, ExpectedFingerprintSHA256: textElement.Source.FingerprintSHA256, Paragraphs: &paragraphs},
				{OperationID: "update-shape", Kind: NativePPTXUpdateAutoShape, ElementID: shapeElement.ID, ExpectedFingerprintSHA256: shapeElement.Source.FingerprintSHA256, AutoShape: &shape},
			}}

			produced, err := ApplyNativePPTXMutations(original, request)
			if err != nil {
				t.Fatalf("apply native mutations: %v", err)
			}
			if bytes.Equal(produced, original) {
				t.Fatal("non-empty exact mutations returned the original package")
			}
			afterOptions := nativeMutationExtractOptions()
			afterOptions.Previous = &before
			after, err := ExtractNativePPTX(produced, afterOptions)
			if err != nil {
				t.Fatalf("reopen/extract: %v", err)
			}
			if after.SourceRevision == nil || *after.SourceRevision == *before.SourceRevision {
				t.Fatalf("source revision did not advance: before=%v after=%v", before.SourceRevision, after.SourceRevision)
			}
			afterText := after.Slides[0].Elements[0]
			afterShape := nativeFixtureAutoShapes(after.Slides[0])[0]
			if !nativeParagraphsEqual(*afterText.Paragraphs, paragraphs) {
				t.Fatalf("text did not round trip: %#v", afterText.Paragraphs)
			}
			if !nativeAutoShapeEquals(afterShape, shape) {
				t.Fatalf("AutoShape did not round trip: %#v", afterShape)
			}
			if !nativeTransformEqual(textElement.Transform, afterText.Transform) || !nativeStringPointerEqual(textElement.Name, afterText.Name) || textElement.Kind != afterText.Kind || textElement.Provenance != afterText.Provenance || !reflect.DeepEqual(textElement.Compatibility, afterText.Compatibility) || !nativePassthroughSemanticsEqual(textElement, afterText) {
				t.Fatalf("text mutation changed non-requested target state: before=%#v after=%#v", textElement, afterText)
			}
			if !nativeStringPointerEqual(shapeElement.Name, afterShape.Name) || !nativeParagraphsEqual(*shapeElement.Paragraphs, *afterShape.Paragraphs) || shapeElement.Kind != afterShape.Kind || shapeElement.Provenance != afterShape.Provenance || !reflect.DeepEqual(shapeElement.Compatibility, afterShape.Compatibility) || !nativePassthroughSemanticsEqual(shapeElement, afterShape) {
				t.Fatalf("AutoShape mutation changed non-requested target state: before=%#v after=%#v", shapeElement, afterShape)
			}
			beforePackage, _ := openNativeExtractPackage(original)
			afterPackage, _ := openNativeExtractPackage(produced)
			for part, payload := range beforePackage.parts {
				if part == textElement.Source.PartName {
					continue
				}
				if !bytes.Equal(payload, afterPackage.parts[part]) {
					t.Fatalf("untouched OPC part %q changed", part)
				}
			}
			slideBytes := afterPackage.parts[textElement.Source.PartName]
			if !bytes.Contains(slideBytes, []byte(`<!--opaque-slide-byte-marker-->`)) || !bytes.Contains(slideBytes, []byte(`<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>`)) {
				t.Fatalf("untouched slide-local source bytes were not preserved: %s", slideBytes)
			}
			assertIndependentPPTXMutationPackageConsumer(t, original, produced, textElement.Source.PartName)
		})
	}
}

func TestApplyNativePPTXMutationsRoundTripsGroupedTextAndShapeAtomically(t *testing.T) {
	t.Parallel()
	for _, strict := range []bool{false, true} {
		strict := strict
		name := map[bool]string{false: "transitional", true: "strict"}[strict]
		t.Run(name, func(t *testing.T) {
			t.Parallel()
			original := nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, mutate: func(parts map[string]string) {
				slide := parts["relocated/slides/slide-a.xml"]
				shapeStart := strings.Index(slide, `<p:sp><p:nvSpPr><p:cNvPr id="2"`)
				if shapeStart < 0 {
					t.Fatal("fixture text shape is unavailable")
				}
				shapeEnd := strings.Index(slide[shapeStart:], `</p:sp>`)
				if shapeEnd < 0 {
					t.Fatal("fixture text shape end is unavailable")
				}
				shapeEnd += shapeStart + len(`</p:sp>`)
				textChild := strings.Replace(slide[shapeStart:shapeEnd], `id="2" name="Title"`, `id="3" name="Grouped Text"`, 1)
				group := `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="2" name="Mutation Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10000000" cy="10000000"/><a:chOff x="0" y="0"/><a:chExt cx="10000000" cy="10000000"/></a:xfrm></p:grpSpPr>` + textChild + nativeGroupRectXML(4, "Grouped Shape", 1000, 2000, 3000, 4000) + `</p:grpSp><!--group-mutation-marker-->`
				parts["relocated/slides/slide-a.xml"] = slide[:shapeStart] + group + slide[shapeEnd:]
			}})
			original = rewriteNativePPTXArchiveCommentForMutationTest(t, original, "grouped native mutation comment")

			before, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
			if err != nil {
				t.Fatalf("extract grouped source through atomic mutation issuer: %v", err)
			}
			if len(before.Slides) != 1 || len(before.Slides[0].Elements) != 1 || before.Slides[0].Elements[0].Kind != NativeElementKindGroup {
				t.Fatalf("group source did not remain exact native authority: %#v", before.Slides)
			}
			groupBefore := before.Slides[0].Elements[0]
			if len(groupBefore.Children) != 2 || groupBefore.Children[0].Kind != NativeElementKindText || groupBefore.Children[1].Kind != NativeElementKindShape || groupBefore.Children[0].TextBody == nil {
				t.Fatalf("group descendants are not exact text/shape targets: %#v", groupBefore.Children)
			}
			textBefore, shapeBefore := groupBefore.Children[0], groupBefore.Children[1]
			paragraphs := nativeMutationParagraphs("grouped exact text")
			adversarial := cloneNativePPTXDeckForMutationTest(t, before)
			childX := adversarial.Slides[0].Elements[0].ChildTransform.X
			*childX = *childX + 1
			if err := verifyNativePPTXContractPreservation(before, adversarial, map[string]NativePPTXMutationKind{textBefore.ID: NativePPTXReplaceText}); err == nil || !strings.Contains(err.Error(), "untouched element") {
				t.Fatalf("descendant target authorization changed its group coordinate space: %v", err)
			}
			x, y, cx, cy := int64(5000), int64(6000), int64(7000), int64(8000)
			fill := "123456"
			shape := NativePPTXAutoShapeMutation{
				Transform: NativeTransform{X: &x, Y: &y, Cx: &cx, Cy: &cy},
				Preset:    NativeShapePresetEllipse,
				Fill:      &fill,
			}
			request := NativePPTXMutationRequest{ExpectedSourceRevision: *before.SourceRevision, Operations: []NativePPTXMutation{
				{OperationID: "group-text", Kind: NativePPTXReplaceText, ElementID: textBefore.ID, ExpectedFingerprintSHA256: textBefore.Source.FingerprintSHA256, Paragraphs: &paragraphs},
				{OperationID: "group-shape", Kind: NativePPTXUpdateAutoShape, ElementID: shapeBefore.ID, ExpectedFingerprintSHA256: shapeBefore.Source.FingerprintSHA256, AutoShape: &shape},
			}}
			produced, err := ApplyNativePPTXMutations(original, request)
			if err != nil {
				t.Fatalf("apply grouped native mutations: %v", err)
			}

			afterOptions := nativeMutationExtractOptions()
			afterOptions.Previous = &before
			after, err := ExtractNativePPTX(produced, afterOptions)
			if err != nil {
				t.Fatalf("extract grouped mutation result: %v", err)
			}
			groupAfter := after.Slides[0].Elements[0]
			if groupAfter.ID != groupBefore.ID || groupAfter.Source == nil || groupAfter.Source.FingerprintSHA256 == groupBefore.Source.FingerprintSHA256 || !reflect.DeepEqual(groupAfter.Transform, groupBefore.Transform) || !reflect.DeepEqual(groupAfter.ChildTransform, groupBefore.ChildTransform) || !reflect.DeepEqual(groupAfter.Compatibility, groupBefore.Compatibility) {
				t.Fatalf("group authority changed outside its descendant fingerprints: before=%#v after=%#v", groupBefore, groupAfter)
			}
			textAfter, shapeAfter := groupAfter.Children[0], groupAfter.Children[1]
			if textAfter.ID != textBefore.ID || !nativeParagraphsEqual(*textAfter.Paragraphs, paragraphs) || !reflect.DeepEqual(textAfter.TextBody, textBefore.TextBody) {
				t.Fatalf("grouped text/bodyPr authority drifted: before=%#v after=%#v", textBefore, textAfter)
			}
			if shapeAfter.ID != shapeBefore.ID || !nativeAutoShapeEquals(shapeAfter, shape) {
				t.Fatalf("grouped AutoShape did not round trip exactly: %#v", shapeAfter)
			}
			reader, err := zip.NewReader(bytes.NewReader(produced), int64(len(produced)))
			if err != nil {
				t.Fatal(err)
			}
			if reader.Comment != "grouped native mutation comment" {
				t.Fatalf("grouped mutation archive comment=%q", reader.Comment)
			}
			pkg, err := openNativeExtractPackage(produced)
			if err != nil {
				t.Fatal(err)
			}
			slide := pkg.parts["relocated/slides/slide-a.xml"]
			if !bytes.Contains(slide, []byte(`<!--group-mutation-marker-->`)) || !bytes.Contains(slide, []byte(`<a:chExt cx="10000000" cy="10000000"/>`)) {
				t.Fatalf("group-local passthrough bytes or child coordinate space changed: %s", slide)
			}

			groupOperation := NativePPTXMutation{OperationID: "group-direct", Kind: NativePPTXReplaceText, ElementID: groupBefore.ID, ExpectedFingerprintSHA256: groupBefore.Source.FingerprintSHA256, Paragraphs: &paragraphs}
			if rejected, err := ApplyNativePPTXMutations(original, NativePPTXMutationRequest{ExpectedSourceRevision: *before.SourceRevision, Operations: []NativePPTXMutation{groupOperation}}); err == nil || rejected != nil || !strings.Contains(err.Error(), "target is not text-bearing") {
				t.Fatalf("direct structural group mutation was not refused atomically: produced=%v err=%v", rejected != nil, err)
			}
		})
	}
}

func TestApplyNativePPTXMutationsUsesNamespacesNotSourcePrefixes(t *testing.T) {
	t.Parallel()
	original := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
		slide := parts["relocated/slides/slide-a.xml"]
		slide = strings.ReplaceAll(slide, "xmlns:p=", "xmlns:s=")
		slide = strings.ReplaceAll(slide, "xmlns:a=", "xmlns:d=")
		slide = strings.ReplaceAll(slide, "<p:", "<s:")
		slide = strings.ReplaceAll(slide, "</p:", "</s:")
		slide = strings.ReplaceAll(slide, "<a:", "<d:")
		slide = strings.ReplaceAll(slide, "</a:", "</d:")
		parts["relocated/slides/slide-a.xml"] = slide
	}})
	before, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
	if err != nil {
		t.Fatalf("extract arbitrary-prefix source: %v", err)
	}
	target := before.Slides[0].Elements[0]
	paragraphs := nativeMutationParagraphs("prefix-independent")
	operation := NativePPTXMutation{OperationID: "prefix-independent", Kind: NativePPTXReplaceText, ElementID: target.ID, ExpectedFingerprintSHA256: target.Source.FingerprintSHA256, Paragraphs: &paragraphs}
	produced, err := ApplyNativePPTXMutations(original, NativePPTXMutationRequest{ExpectedSourceRevision: *before.SourceRevision, Operations: []NativePPTXMutation{operation}})
	if err != nil {
		t.Fatalf("mutate arbitrary-prefix source: %v", err)
	}
	afterOptions := nativeMutationExtractOptions()
	afterOptions.Previous = &before
	after, err := ExtractNativePPTX(produced, afterOptions)
	if err != nil || !nativeParagraphsEqual(*after.Slides[0].Elements[0].Paragraphs, paragraphs) {
		t.Fatalf("arbitrary-prefix round trip: err=%v element=%#v", err, after.Slides)
	}
}

func TestApplyNativePPTXMutationsRejectsStaleAndUnsupportedInputsAtomically(t *testing.T) {
	t.Parallel()
	exactShape := nativeAutoShapeXML(3, "Exact", "rect", `<a:noFill/>`, nativeAutoShapeNoLine("flat", `<a:round/>`), "")
	original := nativeAutoShapeFixture(t, false, exactShape)
	deck, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	textElement := deck.Slides[0].Elements[0]
	shapeElement := nativeFixtureAutoShapes(deck.Slides[0])[0]
	paragraphs := nativeMutationParagraphs("safe")
	x, y, cx, cy := int64(1), int64(2), int64(3), int64(4)
	shape := NativePPTXAutoShapeMutation{Transform: NativeTransform{X: &x, Y: &y, Cx: &cx, Cy: &cy}, Preset: NativeShapePresetEllipse}
	validText := NativePPTXMutation{OperationID: "valid", Kind: NativePPTXReplaceText, ElementID: textElement.ID, ExpectedFingerprintSHA256: textElement.Source.FingerprintSHA256, Paragraphs: &paragraphs}
	validShape := NativePPTXMutation{OperationID: "shape", Kind: NativePPTXUpdateAutoShape, ElementID: shapeElement.ID, ExpectedFingerprintSHA256: shapeElement.Source.FingerprintSHA256, AutoShape: &shape}

	tests := []struct {
		name       string
		revision   string
		operations []NativePPTXMutation
		want       string
	}{
		{name: "stale package revision", revision: "rev-" + strings.Repeat("0", 64), operations: []NativePPTXMutation{validText}, want: "stale source revision"},
		{name: "stale element fingerprint", revision: *deck.SourceRevision, operations: []NativePPTXMutation{func() NativePPTXMutation {
			operation := validText
			operation.ExpectedFingerprintSHA256 = strings.Repeat("0", 64)
			return operation
		}()}, want: "stale element fingerprint"},
		{name: "duplicate target", revision: *deck.SourceRevision, operations: []NativePPTXMutation{validText, func() NativePPTXMutation { operation := validText; operation.OperationID = "again"; return operation }()}, want: "multiple operations target element"},
		{name: "atomic invalid tail", revision: *deck.SourceRevision, operations: []NativePPTXMutation{validText, func() NativePPTXMutation {
			operation := validShape
			operation.OperationID = "bad-tail"
			operation.AutoShape.Preset = NativeShapePresetStar5
			return operation
		}()}, want: "outside the exact native subset"},
	}
	for _, test := range tests {
		test := test
		t.Run(test.name, func(t *testing.T) {
			t.Parallel()
			produced, err := ApplyNativePPTXMutations(original, NativePPTXMutationRequest{ExpectedSourceRevision: test.revision, Operations: test.operations})
			if err == nil || !strings.Contains(err.Error(), test.want) || produced != nil {
				t.Fatalf("got produced=%v err=%v, want nil and %q", produced != nil, err, test.want)
			}
		})
	}
}

func TestApplyNativePPTXMutationsRefusesUnmodeledAutoShapeAndText(t *testing.T) {
	t.Parallel()
	refusedShapeXML := nativeAutoShapeXMLWithGeometry(3, `<a:custGeom><a:avLst/><a:pathLst/></a:custGeom>`, `<a:gradFill/>`, nativeAutoShapeNoLine("flat", `<a:round/>`), "")
	original := nativeAutoShapeFixture(t, false, refusedShapeXML)
	deck, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	refused := nativeFixtureAutoShapes(deck.Slides[0])[0]
	x, y, cx, cy := int64(1), int64(2), int64(3), int64(4)
	shape := NativePPTXAutoShapeMutation{Transform: NativeTransform{X: &x, Y: &y, Cx: &cx, Cy: &cy}, Preset: NativeShapePresetRect}
	operation := NativePPTXMutation{OperationID: "refused", Kind: NativePPTXUpdateAutoShape, ElementID: refused.ID, ExpectedFingerprintSHA256: refused.Source.FingerprintSHA256, AutoShape: &shape}
	if produced, err := ApplyNativePPTXMutations(original, NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{operation}}); err == nil || !strings.Contains(err.Error(), "is refused") || produced != nil {
		t.Fatalf("refused AutoShape mutation result: produced=%v err=%v", produced != nil, err)
	}

	paragraphs := nativeMutationParagraphs("bad")
	*paragraphs[0].Bullet = true
	text := deck.Slides[0].Elements[0]
	textOperation := NativePPTXMutation{OperationID: "unsupported-bullet", Kind: NativePPTXReplaceText, ElementID: text.ID, ExpectedFingerprintSHA256: text.Source.FingerprintSHA256, Paragraphs: &paragraphs}
	if produced, err := ApplyNativePPTXMutations(original, NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{textOperation}}); err == nil || !strings.Contains(err.Error(), "bullet semantics") || produced != nil {
		t.Fatalf("unsupported text mutation result: produced=%v err=%v", produced != nil, err)
	}
	emptyParagraphs := []NativeParagraph{}
	textOperation.OperationID = "empty-text"
	textOperation.Paragraphs = &emptyParagraphs
	if produced, err := ApplyNativePPTXMutations(original, NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{textOperation}}); err == nil || !strings.Contains(err.Error(), "at least one exact paragraph") || produced != nil {
		t.Fatalf("empty text mutation result: produced=%v err=%v", produced != nil, err)
	}
}

func TestApplyNativePPTXMutationsEnforcesResourceBounds(t *testing.T) {
	t.Parallel()
	tooMany := make([]NativePPTXMutation, nativeMaxMutations+1)
	if produced, err := ApplyNativePPTXMutations([]byte("not opened"), NativePPTXMutationRequest{ExpectedSourceRevision: "rev-" + strings.Repeat("0", 64), Operations: tooMany}); err == nil || !strings.Contains(err.Error(), "batch exceeds") || produced != nil {
		t.Fatalf("operation budget result: produced=%v err=%v", produced != nil, err)
	}

	original := nativeExtractFixture(t, nativeExtractFixtureOptions{})
	deck, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	text := deck.Slides[0].Elements[0]
	paragraphs := nativeMutationParagraphs(strings.Repeat("x", nativeMaxTextCodeUnits+1))
	operation := NativePPTXMutation{OperationID: "oversized-text", Kind: NativePPTXReplaceText, ElementID: text.ID, ExpectedFingerprintSHA256: text.Source.FingerprintSHA256, Paragraphs: &paragraphs}
	if produced, err := ApplyNativePPTXMutations(original, NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{operation}}); err == nil || !strings.Contains(err.Error(), "text resource budget") || produced != nil {
		t.Fatalf("text budget result: produced=%v err=%v", produced != nil, err)
	}

	budget := nativeMutationBudget{paragraphs: nativeMaxMutationParagraphs - 1, nodes: nativeMaxMutationNodes - 1}
	twoParagraphs := append(nativeMutationParagraphs("one"), nativeMutationParagraphs("two")...)
	if err := validateNativeMutationParagraphs(twoParagraphs, &budget); err == nil || !strings.Contains(err.Error(), "aggregate paragraph/node budget") {
		t.Fatalf("aggregate paragraph budget was accepted: %v", err)
	}
}

func TestNativePPTXMutationRefusesLiteralTabAndLineBreakTextBeforeWriting(t *testing.T) {
	t.Parallel()
	for _, value := range []string{"left\tright", "first\nsecond", "first\rsecond"} {
		paragraphs := nativeMutationParagraphs(value)
		if err := validateNativeMutationParagraphs(paragraphs, &nativeMutationBudget{}); err == nil || !strings.Contains(err.Error(), "literal tab or line-break") {
			t.Fatalf("control text %q was not visibly refused: %v", value, err)
		}
	}
}

func TestDecodeNativePPTXMutationRequestIsStrictBoundedAndCASJoined(t *testing.T) {
	t.Parallel()
	original := nativeExtractFixture(t, nativeExtractFixtureOptions{})
	deck, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	target := deck.Slides[0].Elements[0]
	paragraphs := nativeMutationParagraphs("strict payload")
	request := NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{{
		OperationID: "strict", Kind: NativePPTXReplaceText, ElementID: target.ID,
		ExpectedFingerprintSHA256: target.Source.FingerprintSHA256, Paragraphs: &paragraphs,
	}}}
	payload, err := json.Marshal(request)
	if err != nil {
		t.Fatal(err)
	}
	outer := "sha256:" + nativeSHA256(original)
	decoded, err := DecodeNativePPTXMutationRequest(payload, outer)
	if err != nil || decoded.ExpectedSourceRevision != *deck.SourceRevision || len(decoded.Operations) != 1 {
		t.Fatalf("valid strict payload decode: decoded=%#v err=%v", decoded, err)
	}

	operationID := `"operationId":"strict"`
	cases := []struct {
		name    string
		payload []byte
		outer   string
		want    string
	}{
		{name: "duplicate root", payload: []byte(strings.Replace(string(payload), `{"expectedSourceRevision":`, `{"expectedSourceRevision":"`+*deck.SourceRevision+`","expectedSourceRevision":`, 1)), outer: outer, want: "duplicate field"},
		{name: "duplicate nested", payload: []byte(strings.Replace(string(payload), operationID, operationID+`,"operationId":"again"`, 1)), outer: outer, want: "duplicate field"},
		{name: "unknown root", payload: []byte(strings.Replace(string(payload), `{"expectedSourceRevision":`, `{"legacyRevision":true,"expectedSourceRevision":`, 1)), outer: outer, want: "unknown field"},
		{name: "unknown nested", payload: []byte(strings.Replace(string(payload), operationID, operationID+`,"legacyPayload":true`, 1)), outer: outer, want: "unknown field"},
		{name: "case alias is unknown", payload: []byte(strings.Replace(string(payload), operationID, `"OperationId":"strict"`, 1)), outer: outer, want: "unknown field"},
		{name: "trailing JSON", payload: append(append([]byte(nil), payload...), []byte(` {}`)...), outer: outer, want: "trailing JSON"},
		{name: "invalid UTF-8", payload: append(append([]byte(nil), payload...), 0xff), outer: outer, want: "valid UTF-8"},
		{name: "invalid outer CAS", payload: payload, outer: "rev-" + strings.Repeat("0", 64), want: "outer expected revision"},
		{name: "mismatched native join", payload: payload, outer: "sha256:" + strings.Repeat("0", 64), want: "does not match outer CAS"},
	}
	for _, test := range cases {
		test := test
		t.Run(test.name, func(t *testing.T) {
			t.Parallel()
			if _, err := DecodeNativePPTXMutationRequest(test.payload, test.outer); err == nil || !strings.Contains(err.Error(), test.want) {
				t.Fatalf("decode error=%v, want %q", err, test.want)
			}
		})
	}
	oversized := bytes.Repeat([]byte{' '}, MaxNativePPTXMutationPayloadBytes+1)
	if _, err := DecodeNativePPTXMutationRequest(oversized, outer); err == nil || !strings.Contains(err.Error(), "1..") {
		t.Fatalf("oversized payload was accepted: %v", err)
	}
	if produced, err := ApplyNativePPTXMutationPayload(original, "sha256:"+strings.Repeat("0", 64), payload); err == nil || produced != nil || !strings.Contains(err.Error(), "stale outer") {
		t.Fatalf("stale outer CAS result: produced=%v err=%v", produced != nil, err)
	}
	if produced, err := ApplyNativePPTXMutationPayload(original, outer, payload); err != nil || len(produced) == 0 {
		t.Fatalf("strict payload apply: produced=%v err=%v", len(produced), err)
	}
}

func TestApplyNativePPTXMutationsRejectsSemanticNoOpsAtomically(t *testing.T) {
	t.Parallel()
	shapeXML := nativeAutoShapeXML(3, "Exact", "rect", `<a:noFill/>`, nativeAutoShapeNoLine("flat", `<a:round/>`), "")
	original := nativeAutoShapeFixture(t, false, shapeXML)
	deck, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	text := deck.Slides[0].Elements[0]
	shape := nativeFixtureAutoShapes(deck.Slides[0])[0]
	shapeNoop := NativePPTXAutoShapeMutation{Transform: shape.Transform, Preset: *shape.Preset, Fill: shape.Fill, Stroke: shape.Stroke}
	textNoop := NativePPTXMutation{OperationID: "text-noop", Kind: NativePPTXReplaceText, ElementID: text.ID, ExpectedFingerprintSHA256: text.Source.FingerprintSHA256, Paragraphs: text.Paragraphs}
	autoShapeNoop := NativePPTXMutation{OperationID: "shape-noop", Kind: NativePPTXUpdateAutoShape, ElementID: shape.ID, ExpectedFingerprintSHA256: shape.Source.FingerprintSHA256, AutoShape: &shapeNoop}
	changedParagraphs := nativeMutationParagraphs("would otherwise change")
	changedText := textNoop
	changedText.OperationID = "changed-text"
	changedText.Paragraphs = &changedParagraphs

	for _, operations := range [][]NativePPTXMutation{{textNoop}, {autoShapeNoop}, {changedText, autoShapeNoop}} {
		produced, err := ApplyNativePPTXMutations(original, NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: operations})
		if err == nil || produced != nil || !strings.Contains(err.Error(), "semantic no-op") {
			t.Fatalf("semantic no-op batch result: produced=%v err=%v", produced != nil, err)
		}
	}
}

func TestVerifyNativePPTXContractPreservationRejectsAfterOnlyAndTopologyChanges(t *testing.T) {
	t.Parallel()
	original := nativeExtractFixture(t, nativeExtractFixtureOptions{})
	before, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	target := before.Slides[0].Elements[0]
	targets := map[string]NativePPTXMutationKind{target.ID: NativePPTXReplaceText}

	tests := []struct {
		name string
		edit func(*NativePPTXDeck)
		want string
	}{
		{name: "after-only slide", edit: func(after *NativePPTXDeck) {
			added := cloneNativePPTXDeckForMutationTest(t, *after).Slides[0]
			added.ID = "slide-after-only"
			after.Slides = append(after.Slides, added)
		}, want: "slide topology"},
		{name: "after-only element", edit: func(after *NativePPTXDeck) {
			added := after.Slides[0].Elements[0]
			added.ID = "element-after-only"
			after.Slides[0].Elements = append(after.Slides[0].Elements, added)
		}, want: "element topology"},
		{name: "slide part remap", edit: func(after *NativePPTXDeck) {
			after.Slides[0].Source.PartName = "ppt/slides/remapped.xml"
		}, want: "source topology"},
		{name: "after-only target refusal", edit: func(after *NativePPTXDeck) {
			after.Slides[0].Elements[0].Compatibility.Status = NativeCompatibilityStatusRefused
		}, want: "compatibility inventory"},
		{name: "introduced passthrough", edit: func(after *NativePPTXDeck) {
			after.Slides[0].Passthrough = append(after.Slides[0].Passthrough, NativePassthroughRef{Token: "after-only", OwnerPart: after.Slides[0].Source.PartName, FingerprintSHA256: strings.Repeat("a", 64), Disposition: NativePassthroughDispositionPreserve})
		}, want: "unsupported inventory"},
		{name: "lost passthrough", edit: func(after *NativePPTXDeck) {
			if len(after.Slides[0].Passthrough) == 0 {
				after.Slides[0].Elements[0].Passthrough = nil
				return
			}
			after.Slides[0].Passthrough = after.Slides[0].Passthrough[:len(after.Slides[0].Passthrough)-1]
		}, want: "unsupported inventory"},
	}
	for _, test := range tests {
		test := test
		t.Run(test.name, func(t *testing.T) {
			t.Parallel()
			after := cloneNativePPTXDeckForMutationTest(t, before)
			test.edit(&after)
			if err := verifyNativePPTXContractPreservation(before, after, targets); err == nil || !strings.Contains(err.Error(), test.want) {
				t.Fatalf("preservation error=%v, want %q", err, test.want)
			}
		})
	}

	duplicate := cloneNativePPTXDeckForMutationTest(t, before)
	duplicate.Slides[0].Elements = append(duplicate.Slides[0].Elements, duplicate.Slides[0].Elements[0])
	if _, err := indexNativePPTXElements(duplicate); err == nil || !strings.Contains(err.Error(), "duplicate element id") {
		t.Fatalf("duplicate produced element id was accepted: %v", err)
	}
	duplicateSlide := cloneNativePPTXDeckForMutationTest(t, before)
	duplicateSlide.Slides = append(duplicateSlide.Slides, duplicateSlide.Slides[0])
	if _, err := indexNativePPTXElements(duplicateSlide); err == nil || !strings.Contains(err.Error(), "duplicate slide id") {
		t.Fatalf("duplicate produced slide id was accepted: %v", err)
	}
}

func TestVerifyNativePPTXRawEntriesRejectsUntouchedMetadataMutation(t *testing.T) {
	t.Parallel()
	original := nativeExtractFixture(t, nativeExtractFixtureOptions{})
	changed := rewriteNativePPTXEntryMetadataForMutationTest(t, original, "[Content_Types].xml")
	if err := verifyNativePPTXRawEntries(original, changed, map[string][]byte{}); err == nil || !strings.Contains(err.Error(), "metadata") {
		t.Fatalf("untouched metadata rewrite was accepted: %v", err)
	}
	changed = rewriteNativePPTXArchiveCommentForMutationTest(t, original, "unexpected archive comment")
	if err := verifyNativePPTXRawEntries(original, changed, map[string][]byte{}); err == nil || !strings.Contains(err.Error(), "archive comment") {
		t.Fatalf("archive comment rewrite was accepted: %v", err)
	}
}

func TestApplyNativePPTXMutationsPreservesArchiveComment(t *testing.T) {
	t.Parallel()
	original := rewriteNativePPTXArchiveCommentForMutationTest(t, nativeExtractFixture(t, nativeExtractFixtureOptions{}), "opaque package comment")
	deck, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	target := deck.Slides[0].Elements[0]
	paragraphs := nativeMutationParagraphs("preserve archive comment")
	operation := NativePPTXMutation{OperationID: "archive-comment", Kind: NativePPTXReplaceText, ElementID: target.ID, ExpectedFingerprintSHA256: target.Source.FingerprintSHA256, Paragraphs: &paragraphs}
	produced, err := ApplyNativePPTXMutations(original, NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{operation}})
	if err != nil {
		t.Fatalf("apply native mutation: %v", err)
	}
	reader, err := zip.NewReader(bytes.NewReader(produced), int64(len(produced)))
	if err != nil {
		t.Fatal(err)
	}
	if reader.Comment != "opaque package comment" {
		t.Fatalf("archive comment=%q, want preserved comment", reader.Comment)
	}
}

func TestApplyNativePPTXMutationsIndependentPackageConsumer(t *testing.T) {
	t.Parallel()
	for _, strict := range []bool{false, true} {
		strict := strict
		name := map[bool]string{false: "transitional", true: "strict"}[strict]
		t.Run(name, func(t *testing.T) {
			t.Parallel()
			shapeXML := nativeAutoShapeXML(3, "Exact Shape", "rect", `<a:solidFill><a:srgbClr val="DDEEFF"/></a:solidFill>`, nativeAutoShapeSolidLine("12700", "flat", `<a:round/>`, "112233"), "")
			original := nativeAutoShapeFixture(t, strict, shapeXML+`<!--opaque-slide-byte-marker-->`)
			before, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
			if err != nil {
				t.Fatalf("extract source: %v", err)
			}
			textElement := before.Slides[0].Elements[0]
			shapeElement := nativeFixtureAutoShapes(before.Slides[0])[0]
			paragraphs := nativeMutationParagraphs("  Exact & native <text>  ")
			fill := "A1B2C3"
			width := int64(25400)
			capValue := NativeStrokeCapSquare
			joinValue := NativeStrokeJoinMiter
			dashValue := NativeStrokeDashSolid
			miter := int64(900000)
			x, y, cx, cy := int64(-42), int64(84), int64(2000000), int64(750000)
			shape := NativePPTXAutoShapeMutation{
				Transform: NativeTransform{X: &x, Y: &y, Cx: &cx, Cy: &cy}, Preset: NativeShapePresetDiamond, Fill: &fill,
				Stroke: &NativeStroke{Color: "445566", WidthEMU: &width, Cap: &capValue, Join: &joinValue, Dash: &dashValue, MiterLimit: &miter},
			}
			produced, err := ApplyNativePPTXMutations(original, NativePPTXMutationRequest{ExpectedSourceRevision: *before.SourceRevision, Operations: []NativePPTXMutation{
				{OperationID: "replace-title", Kind: NativePPTXReplaceText, ElementID: textElement.ID, ExpectedFingerprintSHA256: textElement.Source.FingerprintSHA256, Paragraphs: &paragraphs},
				{OperationID: "update-shape", Kind: NativePPTXUpdateAutoShape, ElementID: shapeElement.ID, ExpectedFingerprintSHA256: shapeElement.Source.FingerprintSHA256, AutoShape: &shape},
			}})
			if err != nil {
				t.Fatalf("apply native mutations: %v", err)
			}
			assertIndependentPPTXMutationPackageConsumer(t, original, produced, textElement.Source.PartName)
		})
	}
}

func TestNativePPTXMutationPathHasNoLegacyReconstructionAuthority(t *testing.T) {
	t.Parallel()
	source, err := os.ReadFile("native_mutate.go")
	if err != nil {
		t.Fatal(err)
	}
	lowerSource := bytes.ToLower(source)
	for _, forbidden := range []string{"parsepptx", "buildpptx", "mammoth", "luckyexcel", "browser", "dom", "html"} {
		if bytes.Contains(lowerSource, []byte(forbidden)) {
			t.Fatalf("native mutation path contains forbidden legacy authority %q", forbidden)
		}
	}
}

func nativeMutationParagraphs(text string) []NativeParagraph {
	align := NativeTextAlignRight
	level := int64(2)
	bullet := false
	bold, italic := false, true
	size := int64(2800)
	color := "ABCDEF"
	family := "Aptos Display"
	return []NativeParagraph{{Align: &align, Level: &level, Bullet: &bullet, Runs: []NativeTextRun{{Text: &text, Bold: &bold, Italic: &italic, FontSizeHundredthPt: &size, Color: &color, FontFamily: &family}}}}
}

func cloneNativePPTXDeckForMutationTest(t *testing.T, source NativePPTXDeck) NativePPTXDeck {
	t.Helper()
	encoded, err := json.Marshal(source)
	if err != nil {
		t.Fatal(err)
	}
	var clone NativePPTXDeck
	if err := json.Unmarshal(encoded, &clone); err != nil {
		t.Fatal(err)
	}
	return clone
}

func rewriteNativePPTXEntryMetadataForMutationTest(t *testing.T, source []byte, target string) []byte {
	t.Helper()
	reader, err := zip.NewReader(bytes.NewReader(source), int64(len(source)))
	if err != nil {
		t.Fatal(err)
	}
	var output bytes.Buffer
	writer := zip.NewWriter(&output)
	for _, file := range reader.File {
		if file.Name != target {
			if err := writer.Copy(file); err != nil {
				t.Fatal(err)
			}
			continue
		}
		header := file.FileHeader
		header.Modified = time.Unix(1_700_000_000, 0).UTC()
		header.SetModTime(header.Modified)
		raw, err := file.OpenRaw()
		if err != nil {
			t.Fatal(err)
		}
		entry, err := writer.CreateRaw(&header)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := io.Copy(entry, raw); err != nil {
			t.Fatal(err)
		}
	}
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	return output.Bytes()
}

func rewriteNativePPTXArchiveCommentForMutationTest(t *testing.T, source []byte, comment string) []byte {
	t.Helper()
	reader, err := zip.NewReader(bytes.NewReader(source), int64(len(source)))
	if err != nil {
		t.Fatal(err)
	}
	var output bytes.Buffer
	writer := zip.NewWriter(&output)
	if err := writer.SetComment(comment); err != nil {
		t.Fatal(err)
	}
	for _, file := range reader.File {
		if err := writer.Copy(file); err != nil {
			t.Fatal(err)
		}
	}
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	return output.Bytes()
}

func assertIndependentPPTXMutationPackageConsumer(t *testing.T, original, produced []byte, slidePart string) {
	t.Helper()
	before := independentPPTXEntries(t, original)
	after := independentPPTXEntries(t, produced)
	assertIndependentPPTXContentTypes(t, after["[Content_Types].xml"])
	for name, payload := range before {
		if name == slidePart {
			continue
		}
		if !bytes.Equal(payload, after[name]) {
			t.Fatalf("independent consumer: untouched OPC part %q changed", name)
		}
	}
	slide, ok := after[slidePart]
	if !ok {
		t.Fatalf("independent consumer: mutated slide part %q is missing", slidePart)
	}
	if !bytes.Contains(slide, []byte(`<!--opaque-slide-byte-marker-->`)) {
		t.Fatal("independent consumer: opaque slide comment was not preserved")
	}
	texts, diamond := independentPPTXSlideMutations(t, slide)
	foundText := false
	for _, text := range texts {
		if text == "  Exact & native <text>  " {
			foundText = true
			break
		}
	}
	if !foundText {
		t.Fatalf("independent consumer: mutated text not present: %#v", texts)
	}
	if diamond.preset != string(NativeShapePresetDiamond) || diamond.fill != "A1B2C3" || diamond.x != -42 || diamond.y != 84 || diamond.cx != 2000000 || diamond.cy != 750000 {
		t.Fatalf("independent consumer: mutated AutoShape = %#v", diamond)
	}
}

func independentPPTXEntries(t *testing.T, data []byte) map[string][]byte {
	t.Helper()
	reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("independent consumer could not reopen OPC zip: %v", err)
	}
	entries := make(map[string][]byte, len(reader.File))
	for _, file := range reader.File {
		opened, openErr := file.Open()
		if openErr != nil {
			t.Fatalf("independent consumer could not open %q: %v", file.Name, openErr)
		}
		payload, readErr := io.ReadAll(opened)
		opened.Close()
		if readErr != nil {
			t.Fatalf("independent consumer could not read %q: %v", file.Name, readErr)
		}
		if _, exists := entries[file.Name]; exists {
			t.Fatalf("independent consumer found duplicate zip name %q", file.Name)
		}
		entries[file.Name] = payload
	}
	if _, ok := entries["[Content_Types].xml"]; !ok {
		t.Fatal("independent consumer missing [Content_Types].xml")
	}
	return entries
}

func assertIndependentPPTXContentTypes(t *testing.T, data []byte) {
	t.Helper()
	decoder := xml.NewDecoder(bytes.NewReader(data))
	root := false
	children := 0
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			t.Fatalf("independent consumer: [Content_Types].xml is not well-formed XML: %v", err)
		}
		start, ok := token.(xml.StartElement)
		if !ok {
			continue
		}
		if start.Name.Space != nsContentTypes {
			t.Fatalf("independent consumer: unexpected content-types namespace %q", start.Name.Space)
		}
		if !root {
			if start.Name.Local != "Types" {
				t.Fatalf("independent consumer: content-types root is %q", start.Name.Local)
			}
			root = true
			continue
		}
		if start.Name.Local != "Default" && start.Name.Local != "Override" {
			t.Fatalf("independent consumer: unexpected content-types child %q", start.Name.Local)
		}
		children++
	}
	if !root || children == 0 {
		t.Fatalf("independent consumer: [Content_Types].xml had root=%v children=%d", root, children)
	}
}

type independentPPTXAutoShape struct {
	preset string
	fill   string
	x, y   int64
	cx, cy int64
}

func independentPPTXSlideMutations(t *testing.T, data []byte) ([]string, independentPPTXAutoShape) {
	t.Helper()
	decoder := xml.NewDecoder(bytes.NewReader(data))
	var texts []string
	var diamond independentPPTXAutoShape
	var path []string
	var current independentPPTXAutoShape
	var text strings.Builder
	collectText := false
	inRunProps := 0
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			t.Fatalf("independent consumer: slide XML is not well-formed: %v", err)
		}
		switch token := token.(type) {
		case xml.StartElement:
			path = append(path, token.Name.Local)
			switch {
			case token.Name.Local == "sp" && isPresentationMLNamespace(token.Name.Space):
				current = independentPPTXAutoShape{}
			case token.Name.Local == "rPr" && isDrawingMLNamespace(token.Name.Space):
				inRunProps++
			case token.Name.Local == "t" && isDrawingMLNamespace(token.Name.Space):
				text.Reset()
				collectText = true
			case token.Name.Local == "prstGeom" && isDrawingMLNamespace(token.Name.Space):
				current.preset = unqualifiedXMLAttr(token, "prst")
			case token.Name.Local == "srgbClr" && isDrawingMLNamespace(token.Name.Space) && inRunProps == 0 && !independentPathHas(path, "ln"):
				current.fill = unqualifiedXMLAttr(token, "val")
			case token.Name.Local == "off" && isDrawingMLNamespace(token.Name.Space):
				current.x = independentXMLIntAttr(t, token, "x")
				current.y = independentXMLIntAttr(t, token, "y")
			case token.Name.Local == "ext" && isDrawingMLNamespace(token.Name.Space) && !independentPathHas(path, "chExt"):
				if independentPathHas(path, "xfrm") {
					current.cx = independentXMLIntAttr(t, token, "cx")
					current.cy = independentXMLIntAttr(t, token, "cy")
				}
			}
		case xml.CharData:
			if collectText {
				text.Write(token)
			}
		case xml.EndElement:
			if collectText && token.Name.Local == "t" {
				texts = append(texts, text.String())
				collectText = false
			}
			if token.Name.Local == "rPr" && isDrawingMLNamespace(token.Name.Space) && inRunProps > 0 {
				inRunProps--
			}
			if token.Name.Local == "sp" && isPresentationMLNamespace(token.Name.Space) && current.preset == string(NativeShapePresetDiamond) {
				diamond = current
			}
			if len(path) == 0 || path[len(path)-1] != token.Name.Local {
				t.Fatalf("independent consumer: slide element stack drifted at %q", token.Name.Local)
			}
			path = path[:len(path)-1]
		}
	}
	return texts, diamond
}

func isPresentationMLNamespace(namespace string) bool {
	return namespace == nsPresentationTransitional || namespace == nsPresentationStrict
}

func isDrawingMLNamespace(namespace string) bool {
	return namespace == nsDrawingTransitional || namespace == nsDrawingStrict
}

func independentPathHas(path []string, local string) bool {
	for _, name := range path {
		if name == local {
			return true
		}
	}
	return false
}

func independentXMLIntAttr(t *testing.T, start xml.StartElement, local string) int64 {
	t.Helper()
	value := unqualifiedXMLAttr(start, local)
	parsed, err := strconv.ParseInt(value, 10, 64)
	if err != nil {
		t.Fatalf("independent consumer: %s=%q is not an integer", local, value)
	}
	return parsed
}

func unqualifiedXMLAttr(start xml.StartElement, local string) string {
	for _, attr := range start.Attr {
		if attr.Name.Space == "" && attr.Name.Local == local {
			return attr.Value
		}
	}
	return ""
}
