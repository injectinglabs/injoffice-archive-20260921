package pptxpatch

import (
	"fmt"
	"reflect"
	"strings"
	"testing"
)

func TestExtractNativePPTXTextBodyDefaultsAndExplicitEquivalents(t *testing.T) {
	t.Parallel()
	explicit := `<a:bodyPr lIns="91440" rIns="91440" tIns="45720" bIns="45720" wrap="square" anchor="t" horzOverflow="overflow" vertOverflow="overflow" vert="horz" rot="0" numCol="1" rtlCol="0" fromWordArt="0" anchorCtr="0" forceAA="0" upright="0" compatLnSpc="0" spcFirstLastPara="0"><a:noAutofit/></a:bodyPr>`
	for _, strict := range []bool{false, true} {
		strict := strict
		t.Run(map[bool]string{false: "transitional", true: "strict"}[strict], func(t *testing.T) {
			t.Parallel()
			defaults, err := ExtractNativePPTX(nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict}), nativeTestExtractOptions())
			if err != nil {
				t.Fatalf("extract defaults: %v", err)
			}
			explicitDeck, err := ExtractNativePPTX(nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, mutate: func(parts map[string]string) {
				parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `<a:bodyPr/>`, explicit, 1)
			}}), nativeTestExtractOptions())
			if err != nil {
				t.Fatalf("extract explicit equivalents: %v", err)
			}
			left := defaults.Slides[0].Elements[0]
			right := explicitDeck.Slides[0].Elements[0]
			if left.Compatibility.Status != NativeCompatibilityStatusEditable || right.Compatibility.Status != NativeCompatibilityStatusEditable || len(left.Passthrough) != 0 || len(right.Passthrough) != 0 {
				t.Fatalf("exact body layouts were not editable: default=%#v explicit=%#v", left, right)
			}
			if left.TextBody == nil || !reflect.DeepEqual(left.TextBody, right.TextBody) {
				t.Fatalf("schema defaults did not materialize exactly: default=%#v explicit=%#v", left.TextBody, right.TextBody)
			}
			want := NativeTextBodyLayout{
				LeftInsetEMU: int64Pointer(91440), RightInsetEMU: int64Pointer(91440),
				TopInsetEMU: int64Pointer(45720), BottomInsetEMU: int64Pointer(45720),
				Wrap: NativeTextWrapSquare, VerticalAnchor: NativeTextVerticalAnchorTop,
				AutoFit: "none", HorizontalOverflow: "overflow", VerticalOverflow: "overflow",
			}
			if !reflect.DeepEqual(*left.TextBody, want) {
				t.Fatalf("unexpected materialized text body: %#v", left.TextBody)
			}
			if issues := ValidateNativePPTX(defaults); len(issues) != 0 {
				t.Fatalf("default-layout deck invalid: %#v", issues)
			}
		})
	}
}

func TestExtractNativePPTXTextBodyDirectInsetsNoWrapAndAnchors(t *testing.T) {
	t.Parallel()
	for _, anchor := range []struct {
		xml  string
		want NativeTextVerticalAnchor
	}{{"t", NativeTextVerticalAnchorTop}, {"ctr", NativeTextVerticalAnchorCenter}, {"b", NativeTextVerticalAnchorBottom}} {
		anchor := anchor
		t.Run(anchor.xml, func(t *testing.T) {
			t.Parallel()
			body := `<a:bodyPr lIns="100" rIns="200" tIns="300" bIns="400" wrap="none" anchor="` + anchor.xml + `"/>`
			deck, err := ExtractNativePPTX(nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
				parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `<a:bodyPr/>`, body, 1)
			}}), nativeTestExtractOptions())
			if err != nil {
				t.Fatalf("extract direct text body: %v", err)
			}
			layout := deck.Slides[0].Elements[0].TextBody
			if layout == nil || *layout.LeftInsetEMU != 100 || *layout.RightInsetEMU != 200 || *layout.TopInsetEMU != 300 || *layout.BottomInsetEMU != 400 || layout.Wrap != NativeTextWrapNone || layout.VerticalAnchor != anchor.want {
				t.Fatalf("direct body properties changed: %#v", layout)
			}
		})
	}
}

func TestExtractNativePPTXTextBodyUnsupportedSemanticsAreObjectRefusals(t *testing.T) {
	t.Parallel()
	vectors := []string{
		`<a:bodyPr lIns="-1"/>`,
		`<a:bodyPr lIns="2147483648"/>`,
		`<a:bodyPr lIns="3000000" rIns="3000000"/>`,
		`<a:bodyPr vert="vert"/>`,
		`<a:bodyPr vert="wordArtVert"/>`,
		`<a:bodyPr numCol="2"/>`,
		`<a:bodyPr anchorCtr="1"/>`,
		`<a:bodyPr anchor="just"/>`,
		`<a:bodyPr anchor="dist"/>`,
		`<a:bodyPr horzOverflow="clip"/>`,
		`<a:bodyPr vertOverflow="ellipsis"/>`,
		`<a:bodyPr><a:normAutofit fontScale="90000"/></a:bodyPr>`,
		`<a:bodyPr><a:prstTxWarp prst="textArchUp"><a:avLst/></a:prstTxWarp></a:bodyPr>`,
	}
	for _, body := range vectors {
		body := body
		t.Run(nativeSHA256([]byte(body))[:8], func(t *testing.T) {
			t.Parallel()
			var request *NativePassthroughTokenRequest
			options := nativeTestExtractOptions()
			options.TokenFactory = NativePassthroughTokenFactoryFunc(func(value NativePassthroughTokenRequest) (string, error) {
				if value.Reason == "pptx.text-layout-refused" {
					copy := value
					request = &copy
				}
				return "token-" + nativeSHA256([]byte(value.OwnerPart + "\x00" + value.ObjectID + "\x00" + value.Reason))[:24], nil
			})
			deck, err := ExtractNativePPTX(nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
				parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `<a:bodyPr/>`, body, 1)
			}}), options)
			if err != nil {
				t.Fatalf("unsupported layout must become an element refusal: %v", err)
			}
			element := deck.Slides[0].Elements[0]
			if element.Compatibility.Status != NativeCompatibilityStatusRefused || element.TextBody != nil || len(element.Passthrough) != 1 || request == nil || request.ObjectID != "cNvPr-2" || !strings.HasPrefix(string(request.Payload), `<p:sp>`) {
				t.Fatalf("unsupported body was not exact capability-backed refusal: element=%#v request=%#v", element, request)
			}
			if request.FingerprintSHA256 != element.Source.FingerprintSHA256 || request.FingerprintSHA256 != nativeSHA256(request.Payload) {
				t.Fatal("text-layout refusal does not bind the exact shape subtree")
			}
		})
	}
}

func TestExtractNativePPTXTextBodyRejectsNonXMLWhitespace(t *testing.T) {
	t.Parallel()
	for _, strict := range []bool{false, true} {
		strict := strict
		dialect := map[bool]string{false: "transitional", true: "strict"}[strict]
		t.Run(dialect+"-body", func(t *testing.T) {
			t.Parallel()
			deck, err := ExtractNativePPTX(nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, mutate: func(parts map[string]string) {
				parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `<a:bodyPr/>`, "<a:bodyPr>\u00a0</a:bodyPr>", 1)
			}}), nativeTestExtractOptions())
			if err != nil {
				t.Fatalf("NBSP text must become an object-local layout refusal: %v", err)
			}
			element := deck.Slides[0].Elements[0]
			if element.Compatibility.Status != NativeCompatibilityStatusRefused || element.TextBody != nil || len(element.Passthrough) != 1 {
				t.Fatalf("NBSP body text was treated as ignorable XML whitespace: %#v", element)
			}
		})
		t.Run(dialect+"-outside-root", func(t *testing.T) {
			t.Parallel()
			_, err := ExtractNativePPTX(nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, mutate: func(parts map[string]string) {
				parts["relocated/slides/slide-a.xml"] = "\u00a0" + parts["relocated/slides/slide-a.xml"]
			}}), nativeTestExtractOptions())
			if err == nil || !strings.Contains(err.Error(), "non-whitespace text outside root") {
				t.Fatalf("NBSP outside the root was not rejected as XML content: %v", err)
			}
		})
	}
}

func TestExtractNativePPTXTextRunPreservesVisibleUnicodeSpaceWithoutXMLSpace(t *testing.T) {
	t.Parallel()
	for _, strict := range []bool{false, true} {
		strict := strict
		for _, visibleSpace := range []string{"\u00a0", "\u3000"} {
			visibleSpace := visibleSpace
			name := map[bool]string{false: "transitional", true: "strict"}[strict] + "-" + nativeSHA256([]byte(visibleSpace))[:8]
			t.Run(name, func(t *testing.T) {
				t.Parallel()
				deck, err := ExtractNativePPTX(nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, mutate: func(parts map[string]string) {
					parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `>world</a:t>`, `>`+visibleSpace+`world</a:t>`, 1)
				}}), nativeTestExtractOptions())
				if err != nil {
					t.Fatalf("extract visible Unicode spacing: %v", err)
				}
				element := deck.Slides[0].Elements[0]
				if element.Compatibility.Status != NativeCompatibilityStatusEditable || element.Paragraphs == nil || len(*element.Paragraphs) != 1 || len((*element.Paragraphs)[0].Runs) != 2 {
					t.Fatalf("visible Unicode spacing was not retained as editable text: %#v", element)
				}
				text := (*element.Paragraphs)[0].Runs[1].Text
				if text == nil || *text != visibleSpace+"world" {
					t.Fatalf("visible Unicode spacing changed: %v", text)
				}
			})
		}
	}
}

func TestExtractNativePPTXTextBodyRejectsMalformedValuesWithBoundedErrors(t *testing.T) {
	t.Parallel()
	vectors := []string{
		`<a:bodyPr wrap="invalid"/>`,
		`<a:bodyPr anchor="invalid"/>`,
		`<a:bodyPr horzOverflow="invalid"/>`,
		`<a:bodyPr vertOverflow="invalid"/>`,
		`<a:bodyPr vert="invalid"/>`,
		`<a:bodyPr anchorCtr="on"/>`,
		`<a:bodyPr wrap="` + strings.Repeat("x", 4096) + `"/>`,
	}
	for index, body := range vectors {
		index := index
		body := body
		t.Run(nativeSHA256([]byte(body))[:8], func(t *testing.T) {
			t.Parallel()
			deck, err := ExtractNativePPTX(nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
				parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `<a:bodyPr/>`, body, 1)
			}}), nativeTestExtractOptions())
			if err != nil {
				if len(err.Error()) <= 512 {
					return
				}
				t.Fatalf("malformed attribute was reflected into an unbounded error (%d bytes)", len(err.Error()))
			}
			if deck.Compatibility.Status == NativeCompatibilityStatusEditable || deck.Slides[0].Compatibility.Status == NativeCompatibilityStatusEditable || len(deck.Slides[0].Passthrough) == 0 {
				t.Fatalf("malformed text-body vector %d was silently accepted: %#v", index, deck.Slides[0])
			}
			for _, diagnostic := range deck.Slides[0].Compatibility.Diagnostics {
				if len(diagnostic.Message) > 2048 {
					t.Fatalf("malformed attribute was reflected into an unbounded diagnostic (%d bytes)", len(diagnostic.Message))
				}
			}
		})
	}
}

func TestExtractNativePPTXTextBodyDoesNotEditEmptyOrMissingSourceParagraphs(t *testing.T) {
	t.Parallel()
	mutations := []func(string) string{
		func(value string) string {
			start := strings.Index(value, `<a:p>`)
			if start < 0 {
				return value
			}
			end := strings.Index(value[start:], `</a:p>`)
			if end < 0 {
				return value
			}
			return value[:start] + value[start+end+len(`</a:p>`):]
		},
		func(value string) string {
			for {
				start := strings.Index(value, `<a:r>`)
				if start < 0 {
					return value
				}
				end := strings.Index(value[start:], `</a:r>`)
				if end < 0 {
					return value
				}
				value = value[:start] + value[start+end+len(`</a:r>`):]
			}
		},
	}
	for index, mutate := range mutations {
		index := index
		mutate := mutate
		t.Run(fmt.Sprintf("vector-%d", index), func(t *testing.T) {
			t.Parallel()
			deck, err := ExtractNativePPTX(nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
				parts["relocated/slides/slide-a.xml"] = mutate(parts["relocated/slides/slide-a.xml"])
			}}), nativeTestExtractOptions())
			if err == nil && (deck.Compatibility.Status == NativeCompatibilityStatusEditable || deck.Slides[0].Compatibility.Status == NativeCompatibilityStatusEditable) {
				t.Fatalf("empty source paragraph vector %d was mislabeled editable: %#v", index, deck.Slides[0])
			}
		})
	}
}

func TestExtractNativePPTXTextBodyRefusesLiteralTabAndLineBreakTextAtObjectScope(t *testing.T) {
	t.Parallel()
	for _, strict := range []bool{false, true} {
		strict := strict
		for _, encoded := range []string{"&#x9;", "&#xA;", "&#xD;"} {
			encoded := encoded
			name := map[bool]string{false: "transitional", true: "strict"}[strict] + "-" + nativeSHA256([]byte(encoded))[:8]
			t.Run(name, func(t *testing.T) {
				t.Parallel()
				var request *NativePassthroughTokenRequest
				options := nativeTestExtractOptions()
				options.TokenFactory = NativePassthroughTokenFactoryFunc(func(value NativePassthroughTokenRequest) (string, error) {
					if value.Reason == "pptx.text-content-refused" {
						copy := value
						request = &copy
					}
					return "token-" + nativeSHA256([]byte(value.OwnerPart + "\x00" + value.ObjectID + "\x00" + value.Reason))[:24], nil
				})
				deck, err := ExtractNativePPTX(nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, mutate: func(parts map[string]string) {
					parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `world</a:t>`, `before`+encoded+`after</a:t>`, 1)
				}}), options)
				if err != nil {
					t.Fatalf("control text must become an object-local refusal: %v", err)
				}
				if len(deck.Slides[0].Elements) != 1 {
					t.Fatalf("control text was escalated beyond its owning object: %#v", deck.Slides[0])
				}
				element := deck.Slides[0].Elements[0]
				if element.Compatibility.Status != NativeCompatibilityStatusRefused || element.TextBody == nil || element.Paragraphs == nil || len(*element.Paragraphs) != 0 || len(element.Passthrough) != 1 {
					t.Fatalf("control text was not an explicit native object refusal: %#v", element)
				}
				if request == nil || request.ObjectID != "cNvPr-2" || request.OwnerPart != "relocated/slides/slide-a.xml" || request.FingerprintSHA256 != nativeSHA256(request.Payload) || !strings.HasPrefix(string(request.Payload), `<p:sp>`) {
					t.Fatalf("control-text capability did not bind the exact shape subtree: %#v", request)
				}
				found := false
				for _, diagnostic := range element.Compatibility.Diagnostics {
					if diagnostic.Code == "pptx.text-content-unavailable" && diagnostic.Severity == NativeDiagnosticSeverityRefusal {
						found = true
					}
				}
				if !found {
					t.Fatalf("missing stable control-text refusal diagnostic: %#v", element.Compatibility.Diagnostics)
				}
				if issues := ValidateNativePPTX(deck); len(issues) != 0 {
					t.Fatalf("control-text refusal deck invalid: %#v", issues)
				}
			})
		}
	}
}

func TestExtractNativePPTXTextBodyRefusesDrawingMLBreakWithExactObjectPassthrough(t *testing.T) {
	t.Parallel()
	for _, strict := range []bool{false, true} {
		strict := strict
		t.Run(map[bool]string{false: "transitional", true: "strict"}[strict], func(t *testing.T) {
			t.Parallel()
			var request *NativePassthroughTokenRequest
			options := nativeTestExtractOptions()
			options.TokenFactory = NativePassthroughTokenFactoryFunc(func(value NativePassthroughTokenRequest) (string, error) {
				if value.Reason == "pptx.text-content-refused" {
					copy := value
					request = &copy
				}
				return "token-" + nativeSHA256([]byte(value.OwnerPart + "\x00" + value.ObjectID + "\x00" + value.Reason))[:24], nil
			})
			deck, err := ExtractNativePPTX(nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, mutate: func(parts map[string]string) {
				parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</a:r><a:r>`, `</a:r><a:br/><a:r>`, 1)
			}}), options)
			if err != nil {
				t.Fatalf("DrawingML break must become an object-local refusal: %v", err)
			}
			if len(deck.Slides) != 1 || len(deck.Slides[0].Elements) != 1 {
				t.Fatalf("DrawingML break escaped its owning text object: %#v", deck.Slides)
			}
			element := deck.Slides[0].Elements[0]
			if element.Compatibility.Status != NativeCompatibilityStatusRefused || element.TextBody == nil || element.Paragraphs == nil || len(*element.Paragraphs) != 0 || len(element.Passthrough) != 1 {
				t.Fatalf("DrawingML break was not an explicit native object refusal: %#v", element)
			}
			if request == nil || request.ObjectID != "cNvPr-2" || request.OwnerPart != "relocated/slides/slide-a.xml" || request.FingerprintSHA256 != nativeSHA256(request.Payload) || !strings.Contains(string(request.Payload), `<a:br/>`) {
				t.Fatalf("DrawingML-break capability did not bind the exact shape subtree: %#v", request)
			}
			if element.Passthrough[0].Token != "token-"+nativeSHA256([]byte(request.OwnerPart + "\x00" + request.ObjectID + "\x00" + request.Reason))[:24] {
				t.Fatalf("DrawingML-break capability token drifted: %#v", element.Passthrough[0])
			}
			found := false
			for _, diagnostic := range element.Compatibility.Diagnostics {
				if diagnostic.Code == "pptx.text-content-unavailable" && diagnostic.Severity == NativeDiagnosticSeverityRefusal {
					found = true
				}
			}
			if !found {
				t.Fatalf("missing stable DrawingML-break refusal diagnostic: %#v", element.Compatibility.Diagnostics)
			}
			if issues := ValidateNativePPTX(deck); len(issues) != 0 {
				t.Fatalf("DrawingML-break refusal deck invalid: %#v", issues)
			}
		})
	}
}

func TestExtractNativePPTXAutoShapeTextCarriesExactBodyLayout(t *testing.T) {
	t.Parallel()
	paragraph := `<p:txBody><a:bodyPr lIns="1000" rIns="2000" tIns="3000" bIns="4000" wrap="square" anchor="b"/><a:lstStyle/><a:p><a:pPr algn="l" lvl="0"><a:buNone/></a:pPr><a:r><a:rPr b="0" i="0" sz="1200"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Aptos"/></a:rPr><a:t>shape text</a:t></a:r></a:p></p:txBody>`
	for _, strict := range []bool{false, true} {
		shape := strings.Replace(nativeAutoShapeXML(3, "Shape text", "rect", `<a:noFill/>`, nativeAutoShapeNoLine("flat", `<a:round/>`), ""), `</p:sp>`, paragraph+`</p:sp>`, 1)
		deck, err := ExtractNativePPTX(nativeAutoShapeFixture(t, strict, shape), nativeTestExtractOptions())
		if err != nil {
			t.Fatalf("strict=%v extract shape text: %v", strict, err)
		}
		element := nativeFixtureAutoShapes(deck.Slides[0])[0]
		if element.Compatibility.Status != NativeCompatibilityStatusEditable || element.TextBody == nil || element.TextBody.Wrap != NativeTextWrapSquare || element.TextBody.VerticalAnchor != NativeTextVerticalAnchorBottom || element.Paragraphs == nil || len(*element.Paragraphs) != 1 || len(element.Passthrough) != 0 {
			t.Fatalf("strict=%v exact shape text was not editable: %#v", strict, element)
		}
	}
}
