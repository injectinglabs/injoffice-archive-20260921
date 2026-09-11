package docxpatch

import (
	"strings"
	"testing"
)

func TestNativePageRelativeFloatingImageSource(t *testing.T) {
	prefix := `<wp:anchor simplePos="0" relativeHeight="7" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1" distT="0" distB="0" distL="0" distR="0"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>1270000</wp:posOffset></wp:positionV><wp:wrapNone/>`
	for _, tc := range []struct {
		name, from, to string
		pass           bool
	}{
		{name: "qualified", pass: true},
		{"wrapping", "<wp:wrapNone/>", "<wp:wrapSquare/>", false},
		{"squareBothSides", "<wp:wrapNone/>", `<wp:wrapSquare wrapText="bothSides"/>`, true},
		{"squareLargest", "<wp:wrapNone/>", `<wp:wrapSquare wrapText="largest"/>`, false},
		{"duplicateWrap", "<wp:wrapNone/>", `<wp:wrapNone/><wp:wrapSquare wrapText="bothSides"/>`, false},
		{"relative", "relativeFrom=\"page\"", "relativeFrom=\"paragraph\"", false},
		{"collision", "allowOverlap=\"1\"", "allowOverlap=\"0\"", false},
		{"unknown", "locked=\"0\"", "locked=\"0\" hidden=\"1\"", false},
		{"duplicate", "relativeHeight=\"7\"", "relativeHeight=\"7\" relativeHeight=\"8\"", false},
		{"offsetAttribute", "<wp:posOffset>", "<wp:posOffset unknown=\"1\">", false},
		{"negative", ">914400<", ">-914400<", false},
		{"orderOverflow", "relativeHeight=\"7\"", "relativeHeight=\"4294967296\"", false},
		{"distanceTop", "distT=\"0\"", "distT=\"1\"", false},
		{"distanceBottom", "distB=\"0\"", "distB=\"1\"", false},
		{"distanceLeft", "distL=\"0\"", "distL=\"1\"", false},
		{"distanceRight", "distR=\"0\"", "distR=\"1\"", false},
	} {
		t.Run(tc.name, func(t *testing.T) {
			parts := transitionalNativeParts()
			xml := strings.Replace(parts["Custom/Main.XML"], `<wp:inline distT="0" distB="0" distL="0" distR="0">`, prefix, 1)
			xml = strings.Replace(xml, `</wp:inline>`, `</wp:anchor>`, 1)
			if tc.from != "" {
				xml = strings.Replace(xml, tc.from, tc.to, 1)
			}
			parts["Custom/Main.XML"] = xml
			doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				if !tc.pass && strings.Contains(err.Error(), "duplicate attribute") {
					return
				}
				t.Fatal(err)
			}
			var drawing *NativeDrawingV1
			for _, run := range doc.Body.Blocks[0].Paragraph.Runs {
				if run.Drawing != nil {
					drawing = run.Drawing
				}
			}
			if !tc.pass {
				if drawing != nil || !hasUnsupportedCode(doc, "FLOATING_DRAWING_SEMANTICS_PRESERVED") {
					t.Fatalf("unsafe anchor admitted: %#v", drawing)
				}
				return
			}
			if drawing == nil || drawing.FloatingLayer == nil || *drawing.FloatingLayer != "behind" || drawing.StackingOrder == nil || *drawing.StackingOrder != 7 || *drawing.XEMU != 914400 || *drawing.YEMU != 1270000 {
				t.Fatalf("floating source projection: %#v", drawing)
			}
			if drawing.EditPolicy.Mode != "read-only" {
				t.Fatal("drawing edit authority widened")
			}
		})
	}
}
