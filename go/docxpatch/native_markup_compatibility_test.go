package docxpatch

import (
	"strings"
	"testing"
)

const testWP14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"

// nativeMCAlternate wraps branches in an alternate that declares mc itself, the
// way Word writes one inside a run or a drawing container.
func nativeMCAlternate(branches string) string {
	return `<mc:AlternateContent xmlns:mc="` + testMC + `">` + branches + `</mc:AlternateContent>`
}

// nativeMCShapeDrawing builds the anchored rect fixture with caller-supplied
// horizontal and vertical placement markup, so an alternate can stand where
// Word writes wp14 relative positioning.
func nativeMCShapeDrawing(positionH, positionV string) string {
	return `<w:drawing xmlns:wp="` + wordDrawingTransitional + `" xmlns:a="` + drawingMLTransitional + `" xmlns:wps="` + nativeTextboxWPS + `">` +
		nativeApproximateAnchor(positionH, positionV) +
		`<a:graphic><a:graphicData uri="` + nativeTextboxWPS + `"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2998800" cy="2829600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="112233"/></a:solidFill></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>`
}

const nativeMCPlainPositionH = `<wp:positionH relativeFrom="page"><wp:posOffset>7333615</wp:posOffset></wp:positionH>`
const nativeMCPlainPositionV = `<wp:positionV relativeFrom="page"><wp:posOffset>106680</wp:posOffset></wp:positionV>`

// TestMarkupCompatibilityFallbackPlacement pins the ECMA-376 Part 3 rule on the
// shape Word actually writes: wp14 relative positioning sits in a mc:Choice this
// preview does not understand, so its mc:Fallback — the plain wp:posOffset twin
// — is what places the shape. Reading neither leaves an anchored shape with no
// position at all, which is how this used to be omitted.
func TestMarkupCompatibilityFallbackPlacement(t *testing.T) {
	positionH := nativeMCAlternate(`<mc:Choice xmlns:wp14="` + testWP14 + `" Requires="wp14"><wp:positionH relativeFrom="page"><wp14:pctPosHOffset>97000</wp14:pctPosHOffset></wp:positionH></mc:Choice><mc:Fallback>` + nativeMCPlainPositionH + `</mc:Fallback>`)
	positionV := nativeMCAlternate(`<mc:Choice xmlns:wp14="` + testWP14 + `" Requires="wp14"><wp:positionV relativeFrom="page"><wp14:pctPosVOffset>-1000</wp14:pctPosVOffset></wp:positionV></mc:Choice><mc:Fallback>` + nativeMCPlainPositionV + `</mc:Fallback>`)
	source := nativeApproximateShapeSource(t, `<w:p><w:r>`+nativeMCShapeDrawing(positionH, positionV)+`</w:r></w:p>`, nil)
	out, err := InspectNativeApproximateDrawingShapesV1(source)
	if err != nil {
		t.Fatal(err)
	}
	if out == nil || len(out.Items) != 1 {
		t.Fatalf("expected one described shape: %#v", out)
	}
	item := out.Items[0]
	if item.Status != "supported" || item.Reason != "" || item.PageAnchor == nil {
		t.Fatalf("fallback placement must be read: %#v", item)
	}
	anchor := item.PageAnchor
	if anchor.HorizontalRelative != "page" || anchor.XEMU != 7333615 || anchor.VerticalRelative != "page" || anchor.YEMU != 106680 {
		t.Fatalf("unexpected fallback anchor: %#v", anchor)
	}
}

// TestMarkupCompatibilityChoiceIsSelectedByNamespace pins that Requires names
// namespace PREFIXES: a prefix is only the namespace the xmlns declarations in
// scope give it, and every prefix a Choice requires has to be understood.
func TestMarkupCompatibilityChoiceIsSelectedByNamespace(t *testing.T) {
	// The understood Choice: wps bound to the wordprocessingShape namespace.
	understood := nativeMCAlternate(`<mc:Choice xmlns:wps="` + nativeTextboxWPS + `" Requires="wps">` + nativeMCShapeDrawing(nativeMCPlainPositionH, nativeMCPlainPositionV) + `</mc:Choice><mc:Fallback><w:pict><v:rect xmlns:v="` + nativeTextboxVML + `" style="position:absolute;width:100pt;height:50pt" fillcolor="#ff0000"/></w:pict></mc:Fallback>`)
	for _, test := range []struct {
		name, choice string
		described    bool
		reason       string
	}{
		{name: "prefix bound to the understood namespace", choice: understood, described: true},
		{
			// A package may bind wps to anything at all. Matching the literal
			// token would unwrap a branch this preview cannot read.
			name:      "prefix spoofed onto a foreign namespace",
			choice:    strings.Replace(understood, `<mc:Choice xmlns:wps="`+nativeTextboxWPS+`" Requires="wps">`, `<mc:Choice xmlns:wps="urn:spoof" Requires="wps">`, 1),
			described: false,
		},
		{
			name:      "prefix not bound at all",
			choice:    strings.Replace(understood, `<mc:Choice xmlns:wps="`+nativeTextboxWPS+`" Requires="wps">`, `<mc:Choice Requires="wps">`, 1),
			described: false,
		},
		{
			// Part 3 selects a Choice only when EVERY required namespace is
			// understood; a14 drawing extensions are not.
			name:      "one required prefix is not understood",
			choice:    strings.Replace(understood, ` Requires="wps">`, ` xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="wps a14">`, 1),
			described: false,
		},
	} {
		t.Run(test.name, func(t *testing.T) {
			if test.choice == understood && !test.described {
				t.Fatal("fixture replacement did not apply")
			}
			out, err := InspectNativeApproximateDrawingShapesV1(nativeApproximateShapeSource(t, `<w:p><w:r>`+test.choice+`</w:r></w:p>`, nil))
			if err != nil {
				t.Fatal(err)
			}
			if !test.described {
				// The VML fallback keeps the run's own source refusal; nothing
				// is described and nothing is painted.
				if out != nil {
					t.Fatalf("unselected choice must not be described: %#v", out)
				}
				return
			}
			if out == nil || len(out.Items) != 1 || out.Items[0].Status != "supported" {
				t.Fatalf("understood choice must be described: %#v", out)
			}
		})
	}
}

// TestMarkupCompatibilityEmptyBranchIsDisclosed pins the failure mode this rule
// is most dangerous in: a selected branch with nothing in it drops authored
// content, so it is reported as an omission rather than silently painted over.
func TestMarkupCompatibilityEmptyBranchIsDisclosed(t *testing.T) {
	a14 := `xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"`
	for _, test := range []struct{ name, branches, reason string }{
		{
			// The shape an XLSX package hit today: an unreadable Choice and an
			// empty Fallback. Painting nothing here is a page missing content.
			name:     "empty fallback",
			branches: `<mc:Choice ` + a14 + ` Requires="a14"><w:drawing/></mc:Choice><mc:Fallback/>`,
			reason:   nativeMCReasonEmptyBranch,
		},
		{
			name:     "no understood choice and no fallback",
			branches: `<mc:Choice ` + a14 + ` Requires="a14"><w:drawing/></mc:Choice>`,
			reason:   nativeMCReasonUnselected,
		},
		{
			name:     "choice without a Requires attribute",
			branches: `<mc:Choice><w:drawing/></mc:Choice><mc:Fallback><w:drawing/></mc:Fallback>`,
			reason:   nativeMCReasonInvalid,
		},
		{
			name:     "content outside the Part 3 model",
			branches: `<mc:Choice ` + a14 + ` Requires="a14"><w:drawing/></mc:Choice><w:drawing/>`,
			reason:   nativeMCReasonInvalid,
		},
		{
			name:     "a choice after the fallback",
			branches: `<mc:Choice ` + a14 + ` Requires="a14"><w:drawing/></mc:Choice><mc:Fallback/><mc:Choice ` + a14 + ` Requires="a14"><w:drawing/></mc:Choice>`,
			reason:   nativeMCReasonInvalid,
		},
	} {
		t.Run(test.name, func(t *testing.T) {
			source := nativeApproximateShapeSource(t, `<w:p><w:r>`+nativeMCAlternate(test.branches)+`</w:r></w:p>`, nil)
			doc, err := ExtractNativeDocumentV1(source)
			if err != nil {
				t.Fatal(err)
			}
			if !hasUnsupportedCode(doc, "UNMODELED_RUN_CONTENT") {
				t.Fatal("the alternate must keep its source refusal")
			}
			out, err := InspectNativeApproximateDrawingShapesV1(source)
			if err != nil {
				t.Fatal(err)
			}
			if out == nil || len(out.Items) != 1 || out.Items[0].Status != "omitted" || out.Items[0].Reason != test.reason {
				t.Fatalf("expected the omission %q to be disclosed: %#v", test.reason, out)
			}
		})
	}
}

// TestMarkupCompatibilityEmptyBranchInsideAnAnchor pins the same disclosure one
// level down: a shape whose placement resolves to nothing is omitted, not
// placed at whatever the reader happened to find.
func TestMarkupCompatibilityEmptyBranchInsideAnAnchor(t *testing.T) {
	positionH := nativeMCAlternate(`<mc:Choice xmlns:wp14="` + testWP14 + `" Requires="wp14"><wp:positionH relativeFrom="page"><wp14:pctPosHOffset>97000</wp14:pctPosHOffset></wp:positionH></mc:Choice><mc:Fallback/>`)
	source := nativeApproximateShapeSource(t, `<w:p><w:r>`+nativeMCShapeDrawing(positionH, nativeMCPlainPositionV)+`</w:r></w:p>`, nil)
	out, err := InspectNativeApproximateDrawingShapesV1(source)
	if err != nil {
		t.Fatal(err)
	}
	if out == nil || len(out.Items) != 1 || out.Items[0].Status != "omitted" || out.Items[0].Reason != nativeMCReasonEmptyBranch || out.Items[0].PageAnchor != nil {
		t.Fatalf("an empty placement branch must be disclosed: %#v", out)
	}
}
