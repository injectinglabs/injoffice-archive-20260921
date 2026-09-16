package pptxpatch

import (
	"strings"
	"testing"
)

// A refused diagram must say which layout shape refused and which construct
// did it. smartart-autofit-sync.pptx renders blank because its layout declares
// adjusted and rotated preset shapes; "diagram shape adjust values are not
// modeled" alone did not identify either.
func TestNativeDiagramLayoutRefusalNamesTheShapeAndTheConstruct(t *testing.T) {
	cases := []struct {
		name    string
		replace string
		want    []string
	}{
		{
			name:    "adjusted preset shape",
			replace: `<dgm:shape type="roundRect"><dgm:adjLst><dgm:adj idx="1" val="0.1"/></dgm:adjLst></dgm:shape>`,
			want: []string{
				"diagram layout shape type=roundRect declares dgm:adj idx=1 val=0.1",
				"preset adjust values are not modeled",
			},
		},
		{
			name:    "rotated layout shape",
			replace: `<dgm:shape type="round2SameRect" rot="180"><dgm:adjLst/></dgm:shape>`,
			want: []string{
				"diagram layout shape type=round2SameRect declares dgm:shape@rot=180",
				"rotated layout shapes are not modeled",
			},
		},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			layout := strings.Replace(nativeDiagramLayoutLayoutXML(nativeDiagramURITransitional), `<dgm:shape><dgm:adjLst/></dgm:shape>`, tc.replace, 1)
			deck, err := ExtractNativePPTX(nativeDiagramLayoutFixture(t, nativeDiagramLayoutFixtureOptions{layout: layout}), nativeDiagramLayoutApproximateOptions())
			if err != nil {
				t.Fatal(err)
			}
			message := ""
			for _, diagnostic := range deck.Slides[0].Compatibility.Diagnostics {
				if diagnostic.Code == nativeDiagramLayoutDefinitionCode {
					message = diagnostic.Message
				}
			}
			if message == "" {
				t.Fatalf("diagram was not refused with %s: %+v", nativeDiagramLayoutDefinitionCode, deck.Slides[0].Compatibility.Diagnostics)
			}
			// The slide-level passthrough names the refused frame as well.
			if !strings.HasPrefix(message, "p:graphicFrame ") {
				t.Fatalf("diagram refusal did not name the slide shape: %s", message)
			}
			for _, want := range tc.want {
				if !strings.Contains(message, want) {
					t.Fatalf("diagram refusal missing %q: %s", want, message)
				}
			}
			if issues := ValidateNativePPTX(deck); len(issues) != 0 {
				t.Fatalf("invalid refused diagram deck: %#v", issues)
			}
		})
	}
}
