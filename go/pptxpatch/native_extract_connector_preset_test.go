package pptxpatch

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"strings"
	"testing"
)

const nativeConnectorStyleRefs = `<p:style><a:lnRef idx="%s"><a:srgbClr val="1F77B4"/></a:lnRef><a:fillRef idx="0"><a:srgbClr val="1F77B4"/></a:fillRef><a:effectRef idx="0"><a:srgbClr val="1F77B4"/></a:effectRef><a:fontRef idx="minor"><a:srgbClr val="000000"/></a:fontRef></p:style>`

// nativeConnectorPresetFixture appends connectors to the slide and installs an
// Office-like three-entry line style matrix so lnRef indices 1..3 resolve.
func nativeConnectorPresetFixture(t *testing.T, strict bool, children string) []byte {
	t.Helper()
	return nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, mutate: func(parts map[string]string) {
		parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, children+`</p:spTree>`, 1)
		drawing := nsDrawingTransitional
		if strict {
			drawing = nsDrawingStrict
		}
		matrixLine := func(width string) string {
			return fmt.Sprintf(`<a:ln w="%s" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`, width)
		}
		parts["relocated/themes/theme.xml"] = fmt.Sprintf(`<a:theme xmlns:a="%s" name="Synthetic"><a:themeElements><a:fmtScheme name="Synthetic"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst>%s%s%s</a:lnStyleLst></a:fmtScheme></a:themeElements></a:theme>`, drawing, matrixLine("6350"), matrixLine("12700"), matrixLine("19050"))
	}})
}

func nativeConnectorDiagnosticCodes(element NativeElement) map[string]NativeDiagnosticSeverity {
	codes := map[string]NativeDiagnosticSeverity{}
	for _, diagnostic := range element.Compatibility.Diagnostics {
		codes[diagnostic.Code] = diagnostic.Severity
	}
	return codes
}

func TestExtractNativePPTXConnectorPresetsEvaluateCatalogGeometryReadOnly(t *testing.T) {
	t.Parallel()
	exactLine := nativeAutoShapeSolidLine("12700", "flat", `<a:round/>`, "123456")
	for _, strict := range []bool{false, true} {
		strict := strict
		t.Run(map[bool]string{false: "transitional", true: "strict"}[strict], func(t *testing.T) {
			t.Parallel()
			children := strings.Join([]string{
				// Rotated, flipped elbow inheriting its outline from the theme matrix.
				nativeConnectorXML(3, "Elbow", `<a:prstGeom prst="bentConnector2"><a:avLst/></a:prstGeom>`, `<a:ln><a:tailEnd type="triangle"/></a:ln>`, ` rot="16200000" flipV="1"`, `<a:stCxn id="2" idx="0"/>`, "", fmt.Sprintf(nativeConnectorStyleRefs, "2")),
				// Literal adjustment on an explicit exact line.
				nativeConnectorXML(4, "Adjusted", `<a:prstGeom prst="bentConnector3"><a:avLst><a:gd name="adj1" fmla="val 25000"/></a:avLst></a:prstGeom>`, exactLine, "", "", "", ""),
				// Curved connector default definition.
				nativeConnectorXML(5, "Curved", `<a:prstGeom prst="curvedConnector3"><a:avLst/></a:prstGeom>`, exactLine, "", "", "", ""),
				// Rotated straight connector leaves the exact editable path.
				nativeConnectorXML(6, "Rotated straight", `<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>`, exactLine, ` rot="5400000"`, "", "", ""),
				// Local width overrides the inherited matrix outline.
				nativeConnectorXML(7, "Wide", `<a:prstGeom prst="bentConnector2"><a:avLst/></a:prstGeom>`, `<a:ln w="38100"><a:headEnd type="oval" w="lg"/></a:ln>`, "", "", "", fmt.Sprintf(nativeConnectorStyleRefs, "1")),
				// A schema-optional absent avLst is not the exact empty list: catalog preview.
				nativeConnectorXML(8, "No adjustment list", `<a:prstGeom prst="straightConnector1"/>`, exactLine, "", "", "", ""),
			}, "")
			input := nativeConnectorPresetFixture(t, strict, children)
			before := bytes.Clone(input)
			requests := []NativePassthroughTokenRequest{}
			options := NativePPTXExtractOptions{TokenFactory: NativePassthroughTokenFactoryFunc(func(request NativePassthroughTokenRequest) (string, error) {
				requests = append(requests, request)
				return "token-" + nativeSHA256([]byte(request.OwnerPart + "\x00" + request.ObjectID + "\x00" + request.Reason))[:24], nil
			})}
			deck, err := ExtractNativePPTX(input, options)
			if err != nil {
				t.Fatalf("extract preset connectors: %v", err)
			}
			if !bytes.Equal(input, before) {
				t.Fatal("projection changed source bytes")
			}
			connectors := nativeFixtureConnectors(deck.Slides[0])
			if len(connectors) != 6 {
				t.Fatalf("connector count = %d, want 6: %#v", len(connectors), deck.Slides[0].Elements)
			}
			for index, connector := range connectors {
				codes := nativeConnectorDiagnosticCodes(connector)
				if connector.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || len(connector.Passthrough) != 1 || connector.Geometry == nil || connector.FlipH != nil || connector.Preset != nil {
					t.Fatalf("connector %d was not a read-only evaluated preview: %#v", index, connector)
				}
				if codes[nativeConnectorPresetPreviewCode] != NativeDiagnosticSeverityWarning {
					t.Fatalf("connector %d lacks the declared preview policy: %#v", index, connector.Compatibility.Diagnostics)
				}
				for code, severity := range codes {
					if severity == NativeDiagnosticSeverityRefusal {
						t.Fatalf("connector %d carries refusal %q", index, code)
					}
				}
				if len(connector.Geometry.Paths) != 1 || connector.Geometry.Paths[0].FillMode != "none" || !connector.Geometry.Paths[0].Stroke || connector.Geometry.Paths[0].Commands[0].Kind != "moveTo" {
					t.Fatalf("connector %d geometry is not one open stroked path: %#v", index, connector.Geometry)
				}
				if connector.Stroke == nil || connector.Stroke.Cap == nil || connector.Stroke.Join == nil || connector.Stroke.Dash == nil {
					t.Fatalf("connector %d stroke incomplete: %#v", index, connector.Stroke)
				}
			}
			for _, request := range requests {
				if strings.HasPrefix(request.Reason, "pptx.connector-") && request.Reason != "pptx.connector-preserve-only" {
					t.Fatalf("unexpected connector capability reason: %#v", request)
				}
			}

			elbow := connectors[0]
			commands := elbow.Geometry.Paths[0].Commands
			if len(commands) != 3 || commands[1].Kind != "lineTo" || *commands[1].X != 1000000 || *commands[1].Y != 0 || commands[2].Kind != "lineTo" || *commands[2].X != 1000000 || *commands[2].Y != 500000 {
				t.Fatalf("bentConnector2 path changed: %#v", commands)
			}
			if elbow.Transform.RotationAngle == nil || *elbow.Transform.RotationAngle != 16200000 || elbow.Transform.FlipV == nil || !*elbow.Transform.FlipV || elbow.Transform.FlipH != nil || elbow.Transform.QuarterTurns != nil {
				t.Fatalf("elbow orientation was not projected as a source affine: %#v", elbow.Transform)
			}
			if elbow.Stroke.Color != "1F77B4" || *elbow.Stroke.WidthEMU != 12700 || *elbow.Stroke.Cap != NativeStrokeCapFlat || *elbow.Stroke.Join != NativeStrokeJoinMiter || elbow.Stroke.MiterLimit == nil || *elbow.Stroke.MiterLimit != 800000 {
				t.Fatalf("theme matrix outline was not inherited: %#v", elbow.Stroke)
			}
			if elbow.TailArrow == nil || !*elbow.TailArrow || elbow.TailEnd == nil || elbow.TailEnd.Type != "triangle" || elbow.HeadArrow != nil || elbow.HeadEnd != nil {
				t.Fatalf("elbow arrowheads lost: %#v", elbow)
			}
			elbowCodes := nativeConnectorDiagnosticCodes(elbow)
			for _, code := range []string{"pptx.source-affine-preview", "pptx.connector-theme-style-preview", "pptx.connector-connection-unavailable"} {
				if elbowCodes[code] != NativeDiagnosticSeverityWarning {
					t.Fatalf("elbow missing warning %q: %#v", code, elbow.Compatibility.Diagnostics)
				}
			}

			adjusted := connectors[1].Geometry.Paths[0].Commands
			if len(adjusted) != 4 || *adjusted[1].X != 250000 || *adjusted[1].Y != 0 || *adjusted[2].X != 250000 || *adjusted[2].Y != 500000 || *adjusted[3].X != 1000000 || *adjusted[3].Y != 500000 {
				t.Fatalf("bentConnector3 literal adjustment was not applied: %#v", adjusted)
			}
			if connectors[1].Transform.RotationAngle != nil || connectors[1].Transform.FlipH != nil || connectors[1].Transform.FlipV != nil {
				t.Fatalf("unrotated connector gained orientation: %#v", connectors[1].Transform)
			}
			curved := connectors[2].Geometry.Paths[0].Commands
			if len(curved) != 3 || curved[1].Kind != "cubicBezierTo" || curved[2].Kind != "cubicBezierTo" {
				t.Fatalf("curvedConnector3 path changed: %#v", curved)
			}
			rotated := connectors[3]
			if len(rotated.Geometry.Paths[0].Commands) != 2 || rotated.Transform.RotationAngle == nil || *rotated.Transform.RotationAngle != 5400000 || rotated.Transform.QuarterTurns != nil {
				t.Fatalf("rotated straight connector was not evaluated with a source affine: %#v", rotated)
			}
			wide := connectors[4]
			if wide.Stroke.Color != "1F77B4" || *wide.Stroke.WidthEMU != 38100 || *wide.Stroke.Join != NativeStrokeJoinMiter || wide.HeadEnd == nil || wide.HeadEnd.Type != "oval" || *wide.HeadEnd.W != "lg" {
				t.Fatalf("local outline override over the matrix entry lost: %#v", wide)
			}
			noList := connectors[5]
			if len(noList.Geometry.Paths[0].Commands) != 2 || noList.Transform.RotationAngle != nil || noList.FlipH != nil {
				t.Fatalf("straight connector without avLst was not a plain catalog preview: %#v", noList)
			}

			if issues := ValidateNativePPTX(deck); len(issues) != 0 {
				t.Fatalf("invalid preset connector deck: %#v", issues)
			}
			encoded, err := MarshalNativePPTXJSON(deck)
			if err != nil {
				t.Fatalf("marshal preset connector deck: %v", err)
			}
			decoded, err := DecodeNativePPTXJSON(encoded)
			if err != nil || len(ValidateNativePPTX(decoded)) != 0 {
				t.Fatalf("preset connector contract round trip: err=%v issues=%#v", err, ValidateNativePPTX(decoded))
			}
			// Evaluated connector geometry never grants editability.
			editable := deck
			editable.Slides[0].Elements = append([]NativeElement(nil), deck.Slides[0].Elements...)
			for index := range editable.Slides[0].Elements {
				if editable.Slides[0].Elements[index].Kind == NativeElementKindConnector {
					editable.Slides[0].Elements[index].Compatibility.Status = NativeCompatibilityStatusEditable
					break
				}
			}
			if issues := ValidateNativePPTX(editable); len(issues) == 0 {
				t.Fatal("editable connector with evaluated geometry was admitted")
			}
		})
	}
}

func TestExtractNativePPTXConnectorPresetGapsRemainRefused(t *testing.T) {
	t.Parallel()
	exactLine := nativeAutoShapeSolidLine("12700", "flat", `<a:round/>`, "123456")
	cases := []struct {
		name, geometry, line, xfrmAttrs, rootExtra, wantCode string
		cNvPrAttrs                                           string
	}{
		{name: "outside-connector-family", geometry: `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`, line: exactLine, wantCode: "pptx.connector-geometry-unavailable"},
		{name: "unknown-adjustment", geometry: `<a:prstGeom prst="bentConnector3"><a:avLst><a:gd name="adj9" fmla="val 1"/></a:avLst></a:prstGeom>`, line: exactLine, wantCode: "pptx.connector-geometry-unavailable"},
		{name: "non-literal-adjustment", geometry: `<a:prstGeom prst="bentConnector3"><a:avLst><a:gd name="adj1" fmla="*/ w 1 2"/></a:avLst></a:prstGeom>`, line: exactLine, wantCode: "pptx.connector-geometry-unavailable"},
		{name: "straight-adjustment", geometry: `<a:prstGeom prst="straightConnector1"><a:avLst><a:gd name="adj1" fmla="val 1"/></a:avLst></a:prstGeom>`, line: exactLine, wantCode: "pptx.connector-geometry-unavailable"},
		{name: "matrix-index", geometry: `<a:prstGeom prst="bentConnector2"><a:avLst/></a:prstGeom>`, line: `<a:ln/>`, rootExtra: fmt.Sprintf(nativeConnectorStyleRefs, "4"), wantCode: "pptx.connector-theme-style-unavailable"},
		{name: "effect-reference", geometry: `<a:prstGeom prst="bentConnector2"><a:avLst/></a:prstGeom>`, line: `<a:ln/>`, rootExtra: strings.Replace(fmt.Sprintf(nativeConnectorStyleRefs, "1"), `effectRef idx="0"`, `effectRef idx="1"`, 1), wantCode: "pptx.connector-theme-style-unavailable"},
		{name: "inherited-line-without-style", geometry: `<a:prstGeom prst="bentConnector2"><a:avLst/></a:prstGeom>`, line: `<a:ln><a:tailEnd type="triangle"/></a:ln>`, wantCode: "pptx.connector-line-unavailable"},
		{name: "local-dash-overrides-matrix", geometry: `<a:prstGeom prst="bentConnector2"><a:avLst/></a:prstGeom>`, line: `<a:ln><a:prstDash val="dash"/></a:ln>`, rootExtra: fmt.Sprintf(nativeConnectorStyleRefs, "1"), wantCode: "pptx.connector-dash-unavailable"},
		{name: "hidden", geometry: `<a:prstGeom prst="bentConnector2"><a:avLst/></a:prstGeom>`, line: exactLine, cNvPrAttrs: ` hidden="1"`, wantCode: "pptx.connector-nonvisual-unavailable"},
		{name: "rotated-custom", geometry: `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst/></a:custGeom>`, line: exactLine, xfrmAttrs: ` rot="5400000"`, wantCode: "pptx.connector-geometry-unavailable"},
		{name: "zero-length-first-segment", geometry: `<a:prstGeom prst="bentConnector3"><a:avLst><a:gd name="adj1" fmla="val 0"/></a:avLst></a:prstGeom>`, line: exactLine, wantCode: "pptx.connector-geometry-unavailable"},
		{name: "zero-length-last-segment", geometry: `<a:prstGeom prst="bentConnector3"><a:avLst><a:gd name="adj1" fmla="val 100000"/></a:avLst></a:prstGeom>`, line: exactLine, wantCode: "pptx.connector-geometry-unavailable"},
		{name: "unknown-transform-attribute", geometry: `<a:prstGeom prst="bentConnector2"><a:avLst/></a:prstGeom>`, line: exactLine, xfrmAttrs: ` rot="5400000" scale="2"`, wantCode: "pptx.connector-transform-unavailable"},
		{name: "invalid-rotation", geometry: `<a:prstGeom prst="bentConnector2"><a:avLst/></a:prstGeom>`, line: exactLine, xfrmAttrs: ` rot="ninety"`, wantCode: "pptx.connector-transform-unavailable"},
	}
	for _, test := range cases {
		test := test
		t.Run(test.name, func(t *testing.T) {
			t.Parallel()
			raw := nativeConnectorXML(3, "Refused preset connector", test.geometry, test.line, test.xfrmAttrs, "", "", test.rootExtra)
			if test.cNvPrAttrs != "" {
				raw = strings.Replace(raw, `name="Refused preset connector"`, `name="Refused preset connector"`+test.cNvPrAttrs, 1)
			}
			deck, err := ExtractNativePPTX(nativeConnectorPresetFixture(t, false, raw), nativeTestExtractOptions())
			if err != nil {
				t.Fatalf("extract refused preset connector: %v", err)
			}
			connectors := nativeFixtureConnectors(deck.Slides[0])
			if len(connectors) != 1 {
				t.Fatalf("refused connector disappeared: %#v", deck.Slides[0].Elements)
			}
			connector := connectors[0]
			if connector.Compatibility.Status != NativeCompatibilityStatusRefused || len(connector.Passthrough) != 1 || connector.Geometry != nil {
				t.Fatalf("connector was approximated instead of refused: %#v", connector)
			}
			if nativeConnectorDiagnosticCodes(connector)[test.wantCode] != NativeDiagnosticSeverityRefusal {
				t.Fatalf("missing refusal %q: %#v", test.wantCode, connector.Compatibility.Diagnostics)
			}
			if issues := ValidateNativePPTX(deck); len(issues) != 0 {
				t.Fatalf("invalid refused connector deck: %#v", issues)
			}
		})
	}
}

func TestNativeShapeStyleMergesLocalOutlineOverThemeMatrix(t *testing.T) {
	t.Parallel()
	fill := `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>`
	shape := nativeAutoShapeXML(3, "Thick themed rect", "rect", "", `<a:ln w="63500"/>`, "")
	shape = strings.Replace(shape, `</p:sp>`, nativeShapeStyleRefs+`</p:sp>`, 1)
	deck, err := ExtractNativePPTX(nativeShapeStyleFixture(t, false, shape, fill), nativeTestExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	element := nativeFixtureAutoShapes(deck.Slides[0])[0]
	if element.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || element.Stroke == nil || element.Stroke.Color != "123456" || *element.Stroke.WidthEMU != 63500 || *element.Stroke.Cap != NativeStrokeCapFlat || *element.Stroke.Join != NativeStrokeJoinMiter || element.Stroke.MiterLimit == nil || *element.Stroke.MiterLimit != 800000 || element.Fill == nil || *element.Fill != "ABCDEF" {
		t.Fatalf("local width did not inherit the matrix outline: %+v stroke=%+v", element, element.Stroke)
	}
	if issues := ValidateNativePPTX(deck); len(issues) != 0 {
		t.Fatal(issues)
	}
	// A local override that is itself outside the exact subset still refuses.
	dashed := strings.Replace(shape, `<a:ln w="63500"/>`, `<a:ln w="63500"><a:prstDash val="dash"/></a:ln>`, 1)
	deck, err = ExtractNativePPTX(nativeShapeStyleFixture(t, false, dashed, fill), nativeTestExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	element = nativeFixtureAutoShapes(deck.Slides[0])[0]
	if element.Compatibility.Status != NativeCompatibilityStatusRefused || element.Stroke != nil {
		t.Fatalf("dashed local override was approximated: %+v", element)
	}
}

func TestMergeNativeInheritedLineOverridesGroupsWithoutMutatingInputs(t *testing.T) {
	t.Parallel()
	ns := nsDrawingTransitional
	node := func(local string, attrs ...xml.Attr) *nativeXMLNode {
		return &nativeXMLNode{Name: xml.Name{Space: ns, Local: local}, Attrs: attrs}
	}
	attr := func(name, value string) xml.Attr { return xml.Attr{Name: xml.Name{Local: name}, Value: value} }
	base := node("ln", attr("w", "6350"), attr("cap", "flat"))
	base.Children = []*nativeXMLNode{node("solidFill"), node("prstDash", attr("val", "solid")), node("miter", attr("lim", "800000"))}
	local := node("ln", attr("w", "38100"))
	local.Children = []*nativeXMLNode{node("noFill"), node("round"), node("tailEnd", attr("type", "triangle")), node("unknown")}
	baseBefore, localBefore := fmt.Sprintf("%#v", *base), fmt.Sprintf("%#v", *local)
	merged := mergeNativeInheritedLine(base, local, ns)
	if fmt.Sprintf("%#v", *base) != baseBefore || fmt.Sprintf("%#v", *local) != localBefore {
		t.Fatal("merge mutated its inputs")
	}
	width, _ := exactNativeAttr(merged, "", "w")
	cap, _ := exactNativeAttr(merged, "", "cap")
	if width != "38100" || cap != "flat" || len(merged.Attrs) != 2 {
		t.Fatalf("attribute override wrong: %#v", merged.Attrs)
	}
	names := []string{}
	for _, child := range merged.Children {
		names = append(names, child.Name.Local)
	}
	if strings.Join(names, ",") != "prstDash,noFill,round,tailEnd,unknown" {
		t.Fatalf("choice groups were not replaced as a whole: %v", names)
	}
	if mergeNativeInheritedLine(nil, local, ns) != local || mergeNativeInheritedLine(base, nil, ns) != base {
		t.Fatal("nil inputs must pass the other side through")
	}
}
