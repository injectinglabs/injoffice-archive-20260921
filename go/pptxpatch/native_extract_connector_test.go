package pptxpatch

import (
	"bytes"
	"fmt"
	"strings"
	"testing"
)

func TestExtractNativePPTXStraightConnectorArrowsAreExactBooleanFlags(t *testing.T) {
	t.Parallel()
	line := nativeAutoShapeSolidLine("12700", "flat", `<a:round/>`, "123456")
	head := strings.Replace(line, `</a:ln>`, `<a:headEnd type="triangle" w="med" sz="med"/></a:ln>`, 1)
	tail := strings.Replace(line, `</a:ln>`, `<a:tailEnd type="stealth" w="sm"/></a:ln>`, 1)
	both := strings.Replace(line, `</a:ln>`, `<a:headEnd type="arrow" w="lg" sz="sm"/><a:tailEnd type="diamond" sz="lg"/></a:ln>`, 1)
	none := strings.Replace(line, `</a:ln>`, `<a:headEnd type="none" w="med" sz="med"/></a:ln>`, 1)
	children := strings.Join([]string{
		nativeConnectorXML(3, "Head", `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>`, head, "", "", "", ""),
		nativeConnectorXML(4, "Tail", `<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>`, tail, "", "", "", ""),
		nativeConnectorXML(5, "Both", `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>`, both, "", "", "", ""),
		nativeConnectorXML(6, "None", `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>`, none, "", "", "", ""),
	}, "")
	deck, err := ExtractNativePPTX(nativeConnectorFixture(t, false, children), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract connector arrows: %v", err)
	}
	connectors := nativeFixtureConnectors(deck.Slides[0])
	if len(connectors) != 4 {
		t.Fatalf("arrow connectors disappeared: %#v", deck.Slides[0].Elements)
	}
	want := []struct{ head, tail bool }{{true, false}, {false, true}, {true, true}, {false, false}}
	for index, connector := range connectors {
		if connector.Compatibility.Status != NativeCompatibilityStatusEditable || len(connector.Passthrough) != 0 {
			t.Fatalf("connector %d arrows were refused: %#v", index, connector)
		}
		gotHead := connector.HeadArrow != nil && *connector.HeadArrow
		gotTail := connector.TailArrow != nil && *connector.TailArrow
		if gotHead != want[index].head || gotTail != want[index].tail {
			t.Fatalf("connector %d arrow flags = head:%v tail:%v, want %#v: %#v", index, gotHead, gotTail, want[index], connector)
		}
	}
}

func TestExtractNativePPTXStraightConnectorsTransitionalAndStrict(t *testing.T) {
	t.Parallel()
	for _, strict := range []bool{false, true} {
		strict := strict
		t.Run(map[bool]string{false: "transitional", true: "strict"}[strict], func(t *testing.T) {
			t.Parallel()
			line := nativeAutoShapeSolidLine("12700", "flat", `<a:round/>`, "1234ab")
			children := strings.Join([]string{
				nativeConnectorXML(3, "Main diagonal", `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>`, line, "", "", "", ""),
				nativeConnectorXML(4, "Horizontal flip", `<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>`, line, ` flipH="1"`, "", "", ""),
				nativeConnectorXML(5, "Vertical flip", `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>`, line, ` flipV="true"`, "", "", ""),
				nativeConnectorXML(6, "Double flip", `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>`, line, ` flipH="1" flipV="1"`, "", "", ""),
				nativeAutoShapeXML(7, "Back-order sentinel", "rect", `<a:noFill/>`, nativeAutoShapeNoLine("flat", `<a:round/>`), ""),
			}, "")
			deck, err := ExtractNativePPTX(nativeConnectorFixture(t, strict, children), nativeTestExtractOptions())
			if err != nil {
				t.Fatalf("extract straight connectors: %v", err)
			}
			if len(deck.Slides) != 1 || len(deck.Slides[0].Elements) != 6 {
				t.Fatalf("connector z-order was not retained: %#v", deck.Slides)
			}
			wantKinds := []NativeElementKind{NativeElementKindText, NativeElementKindConnector, NativeElementKindConnector, NativeElementKindConnector, NativeElementKindConnector, NativeElementKindShape}
			for index, kind := range wantKinds {
				if deck.Slides[0].Elements[index].Kind != kind {
					t.Fatalf("element %d kind = %q, want %q", index, deck.Slides[0].Elements[index].Kind, kind)
				}
			}
			connectors := nativeFixtureConnectors(deck.Slides[0])
			if len(connectors) != 4 {
				t.Fatalf("connector count = %d, want 4: %#v", len(connectors), connectors)
			}
			wantAntiDiagonal := []bool{false, true, true, false}
			for index, connector := range connectors {
				if connector.Source == nil || connector.Source.ObjectID != fmt.Sprintf("cNvPr-%d", index+3) || connector.Compatibility.Status != NativeCompatibilityStatusEditable || len(connector.Passthrough) != 0 {
					t.Fatalf("connector %d was not exact/editable: %#v", index, connector)
				}
				antiDiagonal := connector.FlipH != nil && *connector.FlipH
				if antiDiagonal != wantAntiDiagonal[index] || connector.HeadArrow != nil || connector.TailArrow != nil {
					t.Fatalf("connector %d endpoint parity changed: %#v", index, connector)
				}
				if connector.Stroke == nil || connector.Stroke.Color != "1234AB" || connector.Stroke.WidthEMU == nil || *connector.Stroke.WidthEMU != 12700 || connector.Stroke.Cap == nil || *connector.Stroke.Cap != NativeStrokeCapFlat || connector.Stroke.Join == nil || *connector.Stroke.Join != NativeStrokeJoinRound || connector.Stroke.Dash == nil || *connector.Stroke.Dash != NativeStrokeDashSolid {
					t.Fatalf("connector %d stroke was not exact: %#v", index, connector.Stroke)
				}
			}
			if issues := ValidateNativePPTX(deck); len(issues) != 0 {
				t.Fatalf("invalid connector deck: %#v", issues)
			}
			encoded, err := MarshalNativePPTXJSON(deck)
			if err != nil {
				t.Fatalf("marshal connector deck: %v", err)
			}
			decoded, err := DecodeNativePPTXJSON(encoded)
			if err != nil || len(ValidateNativePPTX(decoded)) != 0 {
				t.Fatalf("connector contract round trip: err=%v issues=%#v", err, ValidateNativePPTX(decoded))
			}
		})
	}
}

func TestExtractNativePPTXConnectorRenderingGapsAreRefusedAndPreserved(t *testing.T) {
	t.Parallel()
	exactLine := nativeAutoShapeSolidLine("12700", "flat", `<a:round/>`, "123456")
	themeLine := strings.Replace(exactLine, `<a:srgbClr val="123456"/>`, `<a:schemeClr val="accent1"><a:lumMod val="50000"/></a:schemeClr>`, 1)
	dashedLine := strings.Replace(exactLine, `val="solid"`, `val="dash"`, 1)
	customGeometry := `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst/></a:custGeom>`
	cases := []struct {
		name      string
		geometry  string
		line      string
		xfrmAttrs string
		spExtra   string
		rootExtra string
		wantCode  string
	}{
		{name: "unknown-arrow-type", geometry: `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>`, line: strings.Replace(exactLine, `</a:ln>`, `<a:tailEnd type="star" w="med" sz="med"/></a:ln>`, 1), wantCode: "pptx.connector-line-unavailable"},
		{name: "unknown-arrow-size", geometry: `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>`, line: strings.Replace(exactLine, `</a:ln>`, `<a:tailEnd type="triangle" w="huge"/></a:ln>`, 1), wantCode: "pptx.connector-line-unavailable"},
		{name: "rotation", geometry: `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>`, line: exactLine, xfrmAttrs: ` rot="60000"`, wantCode: "pptx.connector-transform-unavailable"},
		{name: "bent", geometry: `<a:prstGeom prst="bentConnector3"><a:avLst/></a:prstGeom>`, line: exactLine, wantCode: "pptx.connector-geometry-unavailable"},
		{name: "custom", geometry: customGeometry, line: exactLine, wantCode: "pptx.connector-geometry-unavailable"},
		{name: "theme", geometry: `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>`, line: themeLine, wantCode: "pptx.connector-line-unavailable"},
		{name: "dash", geometry: `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>`, line: dashedLine, wantCode: "pptx.connector-dash-unavailable"},
		{name: "effect", geometry: `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>`, line: exactLine, spExtra: `<a:effectLst><a:outerShdw/></a:effectLst>`, wantCode: "pptx.connector-effects-unavailable"},
		{name: "style", geometry: `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>`, line: exactLine, rootExtra: `<p:style/>`, wantCode: "pptx.connector-theme-style-unavailable"},
	}
	for _, test := range cases {
		test := test
		t.Run(test.name, func(t *testing.T) {
			t.Parallel()
			raw := nativeConnectorXML(3, "Refused connector", test.geometry, test.line, test.xfrmAttrs, "", test.spExtra, test.rootExtra)
			var captured *NativePassthroughTokenRequest
			options := NativePPTXExtractOptions{TokenFactory: NativePassthroughTokenFactoryFunc(func(request NativePassthroughTokenRequest) (string, error) {
				if request.Reason == "pptx.connector-refused" {
					copy := request
					captured = &copy
				}
				return "token-" + nativeSHA256([]byte(request.OwnerPart + "\x00" + request.ObjectID + "\x00" + request.Reason))[:24], nil
			})}
			deck, err := ExtractNativePPTX(nativeConnectorFixture(t, false, raw), options)
			if err != nil {
				t.Fatalf("extract refused connector: %v", err)
			}
			connectors := nativeFixtureConnectors(deck.Slides[0])
			if len(connectors) != 1 {
				t.Fatalf("refused connector disappeared: %#v", deck.Slides[0].Elements)
			}
			connector := connectors[0]
			if connector.Compatibility.Status != NativeCompatibilityStatusRefused || len(connector.Passthrough) != 1 || connector.Source == nil {
				t.Fatalf("connector was approximated instead of refused: %#v", connector)
			}
			foundCode := false
			for _, diagnostic := range connector.Compatibility.Diagnostics {
				if diagnostic.Code == test.wantCode {
					foundCode = true
				}
				if diagnostic.Severity != NativeDiagnosticSeverityRefusal {
					t.Fatalf("rendering gap was not a refusal: %#v", diagnostic)
				}
			}
			if !foundCode {
				t.Fatalf("missing refusal %q: %#v", test.wantCode, connector.Compatibility.Diagnostics)
			}
			if captured == nil || captured.OwnerPart != connector.Source.PartName || captured.ObjectID != connector.Source.ObjectID || captured.FingerprintSHA256 != connector.Source.FingerprintSHA256 || captured.ByteLength != int64(len(raw)) || !bytes.Equal(captured.Payload, []byte(raw)) {
				t.Fatalf("refusal did not bind exact connector bytes: request=%#v source=%#v", captured, connector.Source)
			}
			if issues := ValidateNativePPTX(deck); len(issues) != 0 {
				t.Fatalf("invalid refused connector deck: %#v", issues)
			}
		})
	}
}

func TestExtractNativePPTXGroupedConnectorRetainsAffineAndOrder(t *testing.T) {
	t.Parallel()
	line := nativeAutoShapeSolidLine("25400", "rnd", `<a:bevel/>`, "ABCDEF")
	connector := nativeConnectorXML(4, "Grouped connector", `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>`, line, ` flipH="1"`, "", "", "")
	shape := nativeAutoShapeXML(5, "Grouped sentinel", "rect", `<a:noFill/>`, nativeAutoShapeNoLine("flat", `<a:round/>`), "")
	group := `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Connector group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="1000" y="2000"/><a:ext cx="600" cy="800"/><a:chOff x="100" y="200"/><a:chExt cx="300" cy="400"/></a:xfrm></p:grpSpPr>` + connector + shape + `</p:grpSp>`
	deck, err := ExtractNativePPTX(nativeConnectorFixture(t, false, group), nativeAtomicTestExtractOptions())
	if err != nil {
		t.Fatalf("extract grouped connector: %v", err)
	}
	var projected *NativeElement
	for index := range deck.Slides[0].Elements {
		if deck.Slides[0].Elements[index].Kind == NativeElementKindGroup {
			projected = &deck.Slides[0].Elements[index]
		}
	}
	if projected == nil || projected.ChildTransform == nil || len(projected.Children) != 2 || projected.Children[0].Kind != NativeElementKindConnector || projected.Children[1].Kind != NativeElementKindShape {
		t.Fatalf("group connector/order was not projected: %#v", projected)
	}
	if *projected.Transform.X != 1000 || *projected.Transform.Y != 2000 || *projected.Transform.Cx != 600 || *projected.Transform.Cy != 800 || *projected.ChildTransform.X != 100 || *projected.ChildTransform.Y != 200 || *projected.ChildTransform.Cx != 300 || *projected.ChildTransform.Cy != 400 {
		t.Fatalf("group affine changed: %#v", projected)
	}
	child := projected.Children[0]
	if *child.Transform.X != 400000 || *child.Transform.Y != 200000 || *child.Transform.Cx != 1000000 || *child.Transform.Cy != 500000 || child.FlipH == nil || !*child.FlipH || child.Stroke == nil || child.Compatibility.Status != NativeCompatibilityStatusEditable {
		t.Fatalf("group connector local geometry changed: %#v", child)
	}
	if issues := ValidateNativePPTX(deck); len(issues) != 0 {
		t.Fatalf("invalid grouped connector deck: %#v", issues)
	}
}

func TestExtractNativePPTXGroupedLateConnectorRefusalPublishesOnlyGroupCapability(t *testing.T) {
	t.Parallel()
	line := nativeAutoShapeSolidLine("12700", "flat", `<a:round/>`, "123456")
	attached := nativeConnectorXML(4, "Attached connector", `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>`, line, "", `<a:stCxn id="2" idx="0"/>`, "", "")
	refused := nativeConnectorXML(5, "Late bent", `<a:prstGeom prst="bentConnector3"><a:avLst/></a:prstGeom>`, line, "", "", "", "")
	group := `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Atomic connector group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></p:grpSpPr>` + attached + refused + `</p:grpSp>`
	requests := []NativePassthroughTokenRequest{}
	issuer := nativeAtomicTestTokenFactory(NativePassthroughTokenFactoryFunc(func(request NativePassthroughTokenRequest) (string, error) {
		requests = append(requests, request)
		return "token-" + nativeSHA256([]byte(request.OwnerPart + "\x00" + request.ObjectID + "\x00" + request.Reason))[:24], nil
	}))
	deck, err := ExtractNativePPTX(nativeConnectorFixture(t, false, group), NativePPTXExtractOptions{TokenFactory: issuer})
	if err != nil {
		t.Fatalf("late connector refusal: %v", err)
	}
	connectorOrGroup := []NativePassthroughTokenRequest{}
	for _, request := range requests {
		if request.OwnerPart == "relocated/slides/slide-a.xml" && strings.HasPrefix(request.ObjectID, "cNvPr-") {
			connectorOrGroup = append(connectorOrGroup, request)
		}
	}
	if len(connectorOrGroup) != 1 || connectorOrGroup[0].ObjectID != "cNvPr-3" || connectorOrGroup[0].Reason != "pptx.group-child-refused-unavailable" || !strings.HasPrefix(string(connectorOrGroup[0].Payload), `<p:grpSp`) {
		t.Fatalf("external issuer observed a nested orphan connector capability: %#v", connectorOrGroup)
	}
	for _, request := range requests {
		if strings.HasPrefix(request.Reason, "pptx.connector-") {
			t.Fatalf("connector capability escaped atomic group refusal: %#v", requests)
		}
	}
	for _, element := range deck.Slides[0].Elements {
		if element.Kind == NativeElementKindGroup || element.Kind == NativeElementKindConnector {
			t.Fatalf("opaque refused group leaked a partial projection: %#v", deck.Slides[0].Elements)
		}
	}
	if deck.Slides[0].Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
		t.Fatalf("opaque group did not preserve the slide: %#v", deck.Slides[0].Compatibility)
	}
}

func TestExtractNativePPTXConnectorMalformedOrMixedDialectFailsWithoutProjection(t *testing.T) {
	t.Parallel()
	line := nativeAutoShapeSolidLine("12700", "flat", `<a:round/>`, "123456")
	base := nativeConnectorXML(3, "Malformed connector", `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>`, line, "", "", "", "")
	duplicateTransform := strings.Replace(base, `<a:prstGeom`, `<a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom`, 1)
	zeroExtent := strings.Replace(base, `cx="1000000"`, `cx="0"`, 1)
	mixedDialect := strings.Replace(base, `<a:prstDash`, fmt.Sprintf(`<s:prstDash xmlns:s="%s"`, nsDrawingStrict), 1)
	for name, value := range map[string]string{
		"duplicate-transform": duplicateTransform,
		"zero-extent":         zeroExtent,
		"mixed-dialect":       mixedDialect,
	} {
		name, value := name, value
		t.Run(name, func(t *testing.T) {
			t.Parallel()
			deck, err := ExtractNativePPTX(nativeConnectorFixture(t, false, value), nativeTestExtractOptions())
			if err == nil || len(deck.Slides) != 0 {
				t.Fatalf("malformed connector produced a partial deck: err=%v deck=%#v", err, deck)
			}
		})
	}
}

func nativeConnectorFixture(t *testing.T, strict bool, children string) []byte {
	t.Helper()
	return nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, mutate: func(parts map[string]string) {
		parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, children+`</p:spTree>`, 1)
	}})
}

func nativeConnectorXML(id int, name, geometry, line, xfrmAttrs, connectionChildren, shapeExtra, rootExtra string) string {
	return fmt.Sprintf(`<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="%d" name="%s"/><p:cNvCxnSpPr>%s</p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm%s><a:off x="%d" y="%d"/><a:ext cx="1000000" cy="500000"/></a:xfrm>%s%s%s</p:spPr>%s</p:cxnSp>`, id, name, connectionChildren, xfrmAttrs, id*100000, id*50000, geometry, line, shapeExtra, rootExtra)
}

func nativeFixtureConnectors(slide NativeSlide) []NativeElement {
	result := []NativeElement{}
	for _, element := range slide.Elements {
		if element.Kind == NativeElementKindConnector {
			result = append(result, element)
		}
	}
	return result
}
