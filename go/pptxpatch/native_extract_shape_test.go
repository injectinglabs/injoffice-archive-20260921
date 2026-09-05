package pptxpatch

import (
	"bytes"
	"fmt"
	"strings"
	"testing"
)

func TestExtractNativePPTXAutoShapesTransitionalAndStrict(t *testing.T) {
	t.Parallel()
	for _, strict := range []bool{false, true} {
		strict := strict
		t.Run(map[bool]string{false: "transitional", true: "strict"}[strict], func(t *testing.T) {
			t.Parallel()
			shapes := strings.Join([]string{
				nativeAutoShapeXML(3, "Rectangle", "rect", `<a:solidFill><a:srgbClr val="DDEEFF"/></a:solidFill>`, nativeAutoShapeSolidLine("12700", "flat", `<a:round/>`, "112233"), ""),
				nativeAutoShapeXML(4, "Ellipse", "ellipse", `<a:noFill/>`, nativeAutoShapeSolidLine("25400", "rnd", `<a:miter lim="800000"/>`, "445566"), ""),
				nativeAutoShapeXML(5, "Triangle", "triangle", `<a:solidFill><a:srgbClr val="Aa00fF"/></a:solidFill>`, nativeAutoShapeNoLine("sq", `<a:bevel/>`), ""),
				nativeAutoShapeXML(6, "Diamond", "diamond", `<a:solidFill><a:srgbClr val="00AA11"/></a:solidFill>`, nativeAutoShapeSolidLine("0", "sq", `<a:bevel/>`, "000000"), ` rot="0" flipH="false" flipV="0"`),
			}, "")
			deck, err := ExtractNativePPTX(nativeAutoShapeFixture(t, strict, shapes), nativeTestExtractOptions())
			if err != nil {
				t.Fatalf("extract AutoShapes: %v", err)
			}
			autoShapes := nativeFixtureAutoShapes(deck.Slides[0])
			if len(autoShapes) != 4 {
				t.Fatalf("expected four AutoShapes, got %#v", autoShapes)
			}
			wantPresets := []NativeShapePreset{NativeShapePresetRect, NativeShapePresetEllipse, NativeShapePresetTriangle, NativeShapePresetDiamond}
			for index, shape := range autoShapes {
				if shape.Preset == nil || *shape.Preset != wantPresets[index] || shape.Source == nil || shape.Source.ObjectID != fmt.Sprintf("cNvPr-%d", index+3) || shape.Paragraphs == nil || len(*shape.Paragraphs) != 0 || shape.Compatibility.Status != NativeCompatibilityStatusEditable || len(shape.Passthrough) != 0 {
					t.Fatalf("AutoShape %d was not exact/editable: %#v", index, shape)
				}
			}
			if autoShapes[0].Fill == nil || *autoShapes[0].Fill != "DDEEFF" || autoShapes[0].Stroke == nil || autoShapes[0].Stroke.Cap == nil || *autoShapes[0].Stroke.Cap != NativeStrokeCapFlat || autoShapes[0].Stroke.Join == nil || *autoShapes[0].Stroke.Join != NativeStrokeJoinRound || autoShapes[0].Stroke.Dash == nil || *autoShapes[0].Stroke.Dash != NativeStrokeDashSolid {
				t.Fatalf("solid fill/round stroke mismatch: %#v", autoShapes[0])
			}
			if autoShapes[1].Fill != nil || autoShapes[1].Stroke == nil || autoShapes[1].Stroke.Join == nil || *autoShapes[1].Stroke.Join != NativeStrokeJoinMiter || autoShapes[1].Stroke.MiterLimit == nil || *autoShapes[1].Stroke.MiterLimit != 800000 {
				t.Fatalf("no-fill/miter stroke mismatch: %#v", autoShapes[1])
			}
			if autoShapes[2].Fill == nil || *autoShapes[2].Fill != "AA00FF" || autoShapes[2].Stroke != nil {
				t.Fatalf("normalized fill/explicit no-line mismatch: %#v", autoShapes[2])
			}
			if issues := ValidateNativePPTX(deck); len(issues) != 0 {
				t.Fatalf("invalid extracted deck: %#v", issues)
			}
			encoded, err := MarshalNativePPTXJSON(deck)
			if err != nil {
				t.Fatalf("canonical shape contract: %v", err)
			}
			decoded, err := DecodeNativePPTXJSON(encoded)
			if err != nil || len(ValidateNativePPTX(decoded)) != 0 {
				t.Fatalf("shape contract round trip: err=%v issues=%#v", err, ValidateNativePPTX(decoded))
			}
		})
	}
}

func TestExtractNativePPTXAutoShapeUnsupportedRenderingIsRefusedNotApproximated(t *testing.T) {
	t.Parallel()
	shape := `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Opaque Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm rot="60000"><a:off x="100" y="200"/><a:ext cx="300000" cy="200000"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst/></a:custGeom><a:gradFill/><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:prstDash val="dash"/><a:round/></a:ln><a:effectLst><a:outerShdw/></a:effectLst></p:spPr><p:style/></p:sp>`
	var shapeRequest *NativePassthroughTokenRequest
	options := NativePPTXExtractOptions{TokenFactory: NativePassthroughTokenFactoryFunc(func(request NativePassthroughTokenRequest) (string, error) {
		if request.Reason == "pptx.autoshape-refused" {
			copy := request
			shapeRequest = &copy
		}
		return "token-" + nativeSHA256([]byte(request.OwnerPart + "\x00" + request.ObjectID + "\x00" + request.Reason))[:24], nil
	})}
	deck, err := ExtractNativePPTX(nativeAutoShapeFixture(t, false, shape), options)
	if err != nil {
		t.Fatalf("extract refused AutoShape: %v", err)
	}
	autoShapes := nativeFixtureAutoShapes(deck.Slides[0])
	if len(autoShapes) != 1 {
		t.Fatalf("unsupported AutoShape disappeared: %#v", deck.Slides[0].Elements)
	}
	element := autoShapes[0]
	if element.Compatibility.Status != NativeCompatibilityStatusRefused || element.Preset != nil || element.Fill != nil || element.Stroke != nil || len(element.Passthrough) != 1 || element.Source == nil {
		t.Fatalf("unsupported rendering was flattened or not capability-bound: %#v", element)
	}
	if shapeRequest == nil || deck.SourceRevision == nil || shapeRequest.SourceRevision != *deck.SourceRevision || shapeRequest.Reason != "pptx.autoshape-refused" || shapeRequest.OwnerPart != element.Source.PartName || shapeRequest.ObjectID != element.Source.ObjectID || shapeRequest.FingerprintSHA256 != element.Source.FingerprintSHA256 || shapeRequest.ByteLength != int64(len(shape)) || !bytes.Equal(shapeRequest.Payload, []byte(shape)) {
		t.Fatalf("refused AutoShape capability does not bind the exact source subtree: request=%#v source=%#v", shapeRequest, element.Source)
	}
	wantCodes := map[string]bool{
		"pptx.autoshape-geometry-unavailable":    false,
		"pptx.autoshape-fill-unavailable":        false,
		"pptx.autoshape-dash-unavailable":        false,
		"pptx.autoshape-effects-unavailable":     false,
		"pptx.autoshape-transform-unavailable":   false,
		"pptx.autoshape-theme-style-unavailable": false,
	}
	for _, diagnostic := range element.Compatibility.Diagnostics {
		if _, ok := wantCodes[diagnostic.Code]; ok {
			wantCodes[diagnostic.Code] = true
		}
		if diagnostic.Severity != NativeDiagnosticSeverityRefusal {
			t.Fatalf("rendering-critical gap was not a refusal: %#v", diagnostic)
		}
	}
	for code, found := range wantCodes {
		if !found {
			t.Fatalf("missing refusal diagnostic %q: %#v", code, element.Compatibility.Diagnostics)
		}
	}
	if issues := ValidateNativePPTX(deck); len(issues) != 0 {
		t.Fatalf("refused shape contract is invalid: %#v", issues)
	}
}

func TestExtractNativePPTXAutoShapeUnsupportedPresetAndAdjustmentsAreRefused(t *testing.T) {
	t.Parallel()
	for _, geometry := range []string{
		`<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>`,
		`<a:prstGeom prst="rect"><a:avLst><a:gd name="adj" fmla="val 10000"/></a:avLst></a:prstGeom>`,
	} {
		shape := nativeAutoShapeXMLWithGeometry(3, geometry, `<a:solidFill><a:srgbClr val="DDEEFF"/></a:solidFill>`, nativeAutoShapeSolidLine("12700", "flat", `<a:round/>`, "112233"), "")
		deck, err := ExtractNativePPTX(nativeAutoShapeFixture(t, false, shape), nativeTestExtractOptions())
		if err != nil {
			t.Fatalf("extract unsupported geometry: %v", err)
		}
		element := nativeFixtureAutoShapes(deck.Slides[0])[0]
		if element.Compatibility.Status != NativeCompatibilityStatusRefused || element.Preset != nil {
			t.Fatalf("unsupported geometry was approximated: %#v", element)
		}
	}
}

func TestExtractNativePPTXAutoShapeSourceFingerprintIsExactObjectSubtree(t *testing.T) {
	t.Parallel()
	firstShape := nativeAutoShapeXML(3, "Shape", "rect", `<a:solidFill><a:srgbClr val="112233"/></a:solidFill>`, nativeAutoShapeSolidLine("12700", "flat", `<a:round/>`, "445566"), "")
	secondShape := strings.Replace(firstShape, `val="112233"`, `val="112234"`, 1)
	first, err := ExtractNativePPTX(nativeAutoShapeFixture(t, false, firstShape), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("first extract: %v", err)
	}
	second, err := ExtractNativePPTX(nativeAutoShapeFixture(t, false, secondShape), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("second extract: %v", err)
	}
	firstElement := nativeFixtureAutoShapes(first.Slides[0])[0]
	secondElement := nativeFixtureAutoShapes(second.Slides[0])[0]
	if firstElement.ID != secondElement.ID || firstElement.Source.FingerprintSHA256 == secondElement.Source.FingerprintSHA256 {
		t.Fatalf("durable identity/fingerprint semantics mismatch: first=%#v second=%#v", firstElement.Source, secondElement.Source)
	}
}

func TestExtractNativePPTXAutoShapeReusesIdentityAcrossZOrder(t *testing.T) {
	t.Parallel()
	firstXML := nativeAutoShapeXML(3, "First", "rect", `<a:noFill/>`, nativeAutoShapeNoLine("flat", `<a:round/>`), "")
	secondXML := nativeAutoShapeXML(4, "Second", "ellipse", `<a:noFill/>`, nativeAutoShapeNoLine("flat", `<a:round/>`), "")
	payload := nativeAutoShapeFixture(t, false, firstXML+secondXML)
	previous, err := ExtractNativePPTX(payload, nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("initial extract: %v", err)
	}
	ids := map[string]string{}
	for index := range previous.Slides[0].Elements {
		element := &previous.Slides[0].Elements[index]
		if element.Kind != NativeElementKindShape {
			continue
		}
		element.ID = "shape-reused-" + element.Source.ObjectID
		ids[element.Source.ObjectID] = element.ID
	}
	if issues := ValidateNativePPTX(previous); len(issues) != 0 {
		t.Fatalf("test Previous invalid: %#v", issues)
	}
	options := nativeTestExtractOptions()
	options.Previous = &previous
	reordered, err := ExtractNativePPTX(nativeAutoShapeFixture(t, false, secondXML+firstXML), options)
	if err != nil {
		t.Fatalf("reordered extract: %v", err)
	}
	shapes := nativeFixtureAutoShapes(reordered.Slides[0])
	if len(shapes) != 2 || shapes[0].Source.ObjectID != "cNvPr-4" {
		t.Fatalf("z-order was not retained: %#v", shapes)
	}
	for _, shape := range shapes {
		if shape.ID != ids[shape.Source.ObjectID] {
			t.Fatalf("shape identity changed across z-order: %#v", shape)
		}
	}
}

func TestExtractNativePPTXAutoShapeRejectsDuplicateSingletonsAndHostileNumbers(t *testing.T) {
	t.Parallel()
	base := nativeAutoShapeXML(3, "Shape", "rect", `<a:solidFill><a:srgbClr val="112233"/></a:solidFill>`, nativeAutoShapeSolidLine("12700", "flat", `<a:round/>`, "445566"), "")
	tests := []string{
		strings.Replace(base, `</p:spPr>`, `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>`, 1),
		strings.Replace(base, `</p:spPr>`, `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></p:spPr>`, 1),
		strings.Replace(base, `</p:spPr>`, nativeAutoShapeNoLine("flat", `<a:round/>`)+`</p:spPr>`, 1),
		strings.Replace(base, `<a:xfrm>`, `<a:xfrm><a:off x="1" y="2"/>`, 1),
		strings.Replace(base, `<a:xfrm>`, `<a:badTransform>`, 1),
		strings.Replace(base, `w="12700"`, `w="12700" w="12701"`, 1),
		strings.Replace(base, `x="300000"`, `x="+300000"`, 1),
		strings.Replace(base, `x="300000"`, `x="0300000"`, 1),
		strings.Replace(base, `y="150000"`, `y="-0"`, 1),
		strings.Replace(base, `cx="1000000"`, `cx="01"`, 1),
	}
	for index, shape := range tests {
		if _, err := ExtractNativePPTX(nativeAutoShapeFixture(t, false, shape), nativeTestExtractOptions()); err == nil {
			t.Fatalf("hostile/ambiguous AutoShape case %d was accepted", index)
		}
	}
	for _, hostile := range []string{
		strings.Replace(base, `w="12700"`, `w="+12700"`, 1),
		strings.Replace(base, `w="12700"`, `w="20116801"`, 1),
		strings.Replace(base, `cap="flat"`, `cap="FLAT"`, 1),
		strings.Replace(base, `<a:round/>`, `<a:miter lim="2147483648"/>`, 1),
		strings.Replace(base, `<a:solidFill>`, `<a:solidFill unknown="1">`, 1),
		strings.Replace(base, `</p:spPr>`, `<u:opaque xmlns:u="urn:injoffice:unknown"/></p:spPr>`, 1),
	} {
		refused, err := ExtractNativePPTX(nativeAutoShapeFixture(t, false, hostile), nativeTestExtractOptions())
		if err != nil || nativeFixtureAutoShapes(refused.Slides[0])[0].Compatibility.Status != NativeCompatibilityStatusRefused {
			t.Fatalf("hostile line semantics were not preserved as a refusal: err=%v deck=%#v", err, refused)
		}
	}
}

func TestExtractNativePPTXAutoShapeRejectsOpposingDialectNamespaces(t *testing.T) {
	t.Parallel()
	for _, strict := range []bool{false, true} {
		opposingDrawing := nsDrawingStrict
		opposingRels := nsOfficeRelsStrict
		if strict {
			opposingDrawing = nsDrawingTransitional
			opposingRels = nsOfficeRelsTransitional
		}
		base := nativeAutoShapeXML(3, "Shape", "rect", `<a:noFill/>`, nativeAutoShapeNoLine("flat", `<a:round/>`), "")
		vectors := []string{
			strings.Replace(base, `</p:sp>`, fmt.Sprintf(`<p:style xmlns:x="%s"><x:fillRef idx="0"/></p:style></p:sp>`, opposingDrawing), 1),
			strings.Replace(base, `<p:cNvPr id="3"`, fmt.Sprintf(`<p:cNvPr xmlns:x="%s" x:id="rId9" id="3"`, opposingRels), 1),
		}
		for index, shape := range vectors {
			if _, err := ExtractNativePPTX(nativeAutoShapeFixture(t, strict, shape), nativeTestExtractOptions()); err == nil || !strings.Contains(err.Error(), "mixes Strict and Transitional namespaces") {
				t.Fatalf("strict=%v opposing-dialect vector %d was not rejected: %v", strict, index, err)
			}
		}
	}
}

func TestExtractNativePPTXDoesNotTrustUnrelatedPreviousDocumentIdentity(t *testing.T) {
	t.Parallel()
	payload := nativeAutoShapeFixture(t, false, nativeAutoShapeXML(3, "Shape", "rect", `<a:noFill/>`, nativeAutoShapeNoLine("flat", `<a:round/>`), ""))
	baseline, err := ExtractNativePPTX(payload, nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("baseline extract: %v", err)
	}
	encoded, err := MarshalNativePPTXJSON(baseline)
	if err != nil {
		t.Fatalf("encode baseline Previous: %v", err)
	}
	previous, err := DecodeNativePPTXJSON(encoded)
	if err != nil {
		t.Fatalf("decode baseline Previous: %v", err)
	}
	previous.DocumentID = "deck-unrelated-previous"
	previous.Slides[0].Source.PartName = "unrelated/slides/slide.xml"
	var changeSourcePart func([]NativeElement)
	changeSourcePart = func(elements []NativeElement) {
		for index := range elements {
			if elements[index].Source != nil {
				elements[index].Source.PartName = "unrelated/slides/slide.xml"
			}
			changeSourcePart(elements[index].Children)
		}
	}
	changeSourcePart(previous.Slides[0].Elements)
	if issues := ValidateNativePPTX(previous); len(issues) != 0 {
		t.Fatalf("unrelated Previous must remain contract-valid: %#v", issues)
	}
	options := nativeTestExtractOptions()
	options.Previous = &previous
	current, err := ExtractNativePPTX(payload, options)
	if err != nil {
		t.Fatalf("extract with unrelated Previous: %v", err)
	}
	if current.DocumentID != baseline.DocumentID || current.DocumentID == previous.DocumentID || current.Slides[0].ID != baseline.Slides[0].ID {
		t.Fatalf("unrelated Previous influenced current identities: baseline=%s/%s previous=%s current=%s/%s", baseline.DocumentID, baseline.Slides[0].ID, previous.DocumentID, current.DocumentID, current.Slides[0].ID)
	}
}

func TestExtractNativePPTXRejectsAutoShapeCapabilityTokenCollision(t *testing.T) {
	t.Parallel()
	first := nativeAutoShapeXMLWithGeometry(3, `<a:custGeom/>`, `<a:noFill/>`, nativeAutoShapeNoLine("flat", `<a:round/>`), "")
	second := nativeAutoShapeXMLWithGeometry(4, `<a:custGeom/>`, `<a:noFill/>`, nativeAutoShapeNoLine("flat", `<a:round/>`), "")
	options := nativeTestExtractOptions()
	options.TokenFactory = NativePassthroughTokenFactoryFunc(func(request NativePassthroughTokenRequest) (string, error) {
		if request.Reason == "pptx.autoshape-refused" {
			return "collision-autoshape", nil
		}
		return "token-" + nativeSHA256([]byte(request.OwnerPart + "\x00" + request.ObjectID + "\x00" + request.Reason))[:24], nil
	})
	if _, err := ExtractNativePPTX(nativeAutoShapeFixture(t, false, first+second), options); err == nil || !strings.Contains(err.Error(), "reused one capability") {
		t.Fatalf("distinct refused AutoShapes shared a capability token: %v", err)
	}
}

func nativeAutoShapeFixture(t *testing.T, strict bool, shapes string) []byte {
	t.Helper()
	return nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, mutate: func(parts map[string]string) {
		parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, shapes+`</p:spTree>`, 1)
	}})
}

func nativeAutoShapeXML(id int, name, preset, fill, line, xfrmAttrs string) string {
	return nativeAutoShapeXMLWithNameAndGeometry(id, name, fmt.Sprintf(`<a:prstGeom prst="%s"><a:avLst/></a:prstGeom>`, preset), fill, line, xfrmAttrs)
}

func nativeAutoShapeXMLWithGeometry(id int, geometry, fill, line, xfrmAttrs string) string {
	return nativeAutoShapeXMLWithNameAndGeometry(id, fmt.Sprintf("Shape %d", id), geometry, fill, line, xfrmAttrs)
}

func nativeAutoShapeXMLWithNameAndGeometry(id int, name, geometry, fill, line, xfrmAttrs string) string {
	return fmt.Sprintf(`<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm%s><a:off x="%d" y="%d"/><a:ext cx="1000000" cy="500000"/></a:xfrm>%s%s%s</p:spPr></p:sp>`, id, name, xfrmAttrs, id*100000, id*50000, geometry, fill, line)
}

func nativeAutoShapeSolidLine(width, cap, join, color string) string {
	return fmt.Sprintf(`<a:ln w="%s" cap="%s" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="%s"/></a:solidFill><a:prstDash val="solid"/>%s</a:ln>`, width, cap, color, join)
}

func nativeAutoShapeNoLine(cap, join string) string {
	return fmt.Sprintf(`<a:ln w="0" cap="%s" cmpd="sng" algn="ctr"><a:noFill/><a:prstDash val="solid"/>%s</a:ln>`, cap, join)
}

func nativeFixtureAutoShapes(slide NativeSlide) []NativeElement {
	result := []NativeElement{}
	for _, element := range slide.Elements {
		if element.Kind == NativeElementKindShape {
			result = append(result, element)
		}
	}
	return result
}
