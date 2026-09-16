package pptxpatch

import (
	"strings"
	"testing"
)

// A refused p:graphicFrame must leave the same kind of trace a refused p:sp
// leaves: an element that holds the authored box with no content in it. When
// the frame vanished from the element list, the slide handed downstream layout
// a shape set the source never had.
func TestExtractNativePPTXRefusedGraphicFrameKeepsAnEmptyRegion(t *testing.T) {
	t.Parallel()
	frame := nativeExactTableGraphicFrameXML(3, "Refused table", []int64{500000, 500000}, []int64{500000}, [][]string{{
		nativeExactTableCellXML("One", "l", "FFFFFF"), nativeExactTableCellXML("Two", "r", "EEEEEE"),
	}}, "")
	merged := strings.Replace(frame, `<a:tc>`, `<a:tc gridSpan="2">`, 1)
	if merged == frame {
		t.Fatal("table fixture drifted")
	}
	deck, err := ExtractNativePPTX(nativeTableFixture(t, false, merged), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract refused table frame: %v", err)
	}
	slide := deck.Slides[0]
	if issues := ValidateNativePPTX(deck); len(issues) != 0 {
		t.Fatalf("refused graphic frame region is not a valid native deck: %#v", issues)
	}
	// The region is the last element because the frame is the last shape-tree
	// child: a refusal may not renumber the elements that follow it.
	if len(slide.Elements) < 2 {
		t.Fatalf("refused graphic frame emitted no region: %#v", slide.Elements)
	}
	region := slide.Elements[len(slide.Elements)-1]
	if region.Kind != NativeElementKindShape || region.Compatibility.Status != NativeCompatibilityStatusRefused {
		t.Fatalf("refused graphic frame did not emit a refused shape region: %#v", region)
	}
	if region.Transform.X == nil || *region.Transform.X != 300000 || region.Transform.Y == nil || *region.Transform.Y != 150000 ||
		region.Transform.Cx == nil || *region.Transform.Cx != 1000000 || region.Transform.Cy == nil || *region.Transform.Cy != 500000 {
		t.Fatalf("refused graphic frame region lost the authored box: %#v", region.Transform)
	}
	// An empty region, never a filled box and never fabricated content (#259).
	if region.Preset != nil || region.Geometry != nil || region.Fill != nil || region.Stroke != nil || region.Table != nil || region.Chart != nil {
		t.Fatalf("refused graphic frame region claims paint it never rendered: %#v", region)
	}
	if region.Paragraphs == nil || len(*region.Paragraphs) != 0 {
		t.Fatalf("refused graphic frame region claims text: %#v", region.Paragraphs)
	}
	if region.Source == nil || region.Source.ObjectID != "cNvPr-3" {
		t.Fatalf("refused graphic frame region lost its source anchor: %#v", region.Source)
	}
	refusals := 0
	for _, diagnostic := range region.Compatibility.Diagnostics {
		if diagnostic.Severity == NativeDiagnosticSeverityRefusal && diagnostic.Code == "pptx.table-merge-unavailable" {
			refusals++
			if diagnostic.Scope == nil || diagnostic.Scope.ElementID == nil || *diagnostic.Scope.ElementID != region.ID {
				t.Fatalf("region refusal is not scoped to the region: %#v", diagnostic.Scope)
			}
		}
	}
	if refusals != 1 {
		t.Fatalf("refused graphic frame region does not name its refusal: %#v", region.Compatibility.Diagnostics)
	}
	// The exact bytes stay on the slide passthrough; the region owns none.
	if len(region.Passthrough) != 0 || len(slide.Passthrough) == 0 {
		t.Fatalf("refused graphic frame moved its preserved subtree: region=%#v slide=%d", region.Passthrough, len(slide.Passthrough))
	}
	// Siblings still paint: refusal stays per element.
	for _, element := range slide.Elements[:len(slide.Elements)-1] {
		if element.Compatibility.Status == NativeCompatibilityStatusRefused {
			t.Fatalf("frame refusal spread to a sibling: %#v", element)
		}
	}
}

// A refused chart frame takes the same route as a refused table frame: the
// refusal codes differ, the empty region does not.
func TestExtractNativePPTXRefusedChartGraphicFrameKeepsAnEmptyRegion(t *testing.T) {
	t.Parallel()
	frame := nativeChartGraphicFrameXML(false, 3, "Broken chart", ` foo="bar"`)
	deck, err := ExtractNativePPTX(nativeChartFixture(t, nativeChartFixtureOptions{frameXML: frame}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract refused chart frame: %v", err)
	}
	slide := deck.Slides[0]
	if issues := ValidateNativePPTX(deck); len(issues) != 0 {
		t.Fatalf("refused chart frame region is not a valid native deck: %#v", issues)
	}
	regions := 0
	for _, element := range slide.Elements {
		if element.Compatibility.Status != NativeCompatibilityStatusRefused {
			continue
		}
		regions++
		if element.Kind != NativeElementKindShape || element.Chart != nil || element.Preset != nil || element.Geometry != nil || element.Fill != nil {
			t.Fatalf("refused chart frame leaked a projection: %#v", element)
		}
		if element.Transform.Cx == nil || *element.Transform.Cx <= 0 || element.Transform.Cy == nil || *element.Transform.Cy <= 0 {
			t.Fatalf("refused chart frame region has no extent: %#v", element.Transform)
		}
	}
	if regions != 1 {
		t.Fatalf("refused chart frame did not emit exactly one region: %#v", slide.Elements)
	}
}

// The region reports the box the source states, so a frame that states no
// readable box emits nothing at all rather than a fabricated extent.
func TestExtractNativePPTXRefusedGraphicFrameWithoutAReadableBoxEmitsNoRegion(t *testing.T) {
	t.Parallel()
	frame := nativeExactTableGraphicFrameXML(3, "Refused table", []int64{500000, 500000}, []int64{500000}, [][]string{{
		nativeExactTableCellXML("One", "l", "FFFFFF"), nativeExactTableCellXML("Two", "r", "EEEEEE"),
	}}, "")
	// The unmodeled root attribute refuses the frame before its geometry is
	// read; the extent is then not something the source states exactly.
	broken := strings.Replace(frame, `<p:graphicFrame>`, `<p:graphicFrame foo="bar">`, 1)
	broken = strings.Replace(broken, `<a:ext cx="1000000" cy="500000"/>`, `<a:ext cy="500000"/>`, 1)
	if !strings.Contains(broken, `foo="bar"`) || strings.Contains(broken, `cx="1000000"`) {
		t.Fatal("table fixture drifted")
	}
	deck, err := ExtractNativePPTX(nativeTableFixture(t, false, broken), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract unreadable frame: %v", err)
	}
	slide := deck.Slides[0]
	for _, element := range slide.Elements {
		if element.Compatibility.Status == NativeCompatibilityStatusRefused {
			t.Fatalf("a frame with no stated box claimed an extent: %#v", element)
		}
	}
	if slide.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || len(slide.Passthrough) == 0 {
		t.Fatalf("unreadable frame was not preserved: %#v passthrough=%d", slide.Compatibility, len(slide.Passthrough))
	}
	if issues := ValidateNativePPTX(deck); len(issues) != 0 {
		t.Fatalf("invalid preserved deck: %#v", issues)
	}
}

// nativeRefusedFrameRegions collects the empty regions a refused graphic frame
// leaves behind, so the refusal fixtures can state that the frame keeps its box
// without restating the whole projection.
func nativeRefusedFrameRegions(t *testing.T, elements []NativeElement, code string) []NativeElement {
	t.Helper()
	regions := []NativeElement{}
	for _, element := range elements {
		if element.Compatibility.Status != NativeCompatibilityStatusRefused {
			continue
		}
		if element.Kind != NativeElementKindShape || element.Preset != nil || element.Geometry != nil || element.Fill != nil || element.Table != nil || element.Chart != nil || len(element.Children) != 0 {
			t.Fatalf("refused frame region claims content: %#v", element)
		}
		if element.Transform.Cx == nil || *element.Transform.Cx <= 0 || element.Transform.Cy == nil || *element.Transform.Cy <= 0 {
			t.Fatalf("refused frame region has no extent: %#v", element.Transform)
		}
		named := false
		for _, diagnostic := range element.Compatibility.Diagnostics {
			if diagnostic.Severity == NativeDiagnosticSeverityRefusal && diagnostic.Code == code {
				named = true
			}
		}
		if !named {
			t.Fatalf("refused frame region does not carry %q: %#v", code, element.Compatibility.Diagnostics)
		}
		regions = append(regions, element)
	}
	return regions
}
