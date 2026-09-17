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
		// Zero keeps the fixture's own offsets; a case that rewrites one states it.
		wantX, wantY int64
	}{
		{name: "qualified", pass: true},
		{"wrapping", "<wp:wrapNone/>", "<wp:wrapSquare/>", false, 0, 0},
		{"squareBothSides", "<wp:wrapNone/>", `<wp:wrapSquare wrapText="bothSides"/>`, true, 0, 0},
		{"squareLargest", "<wp:wrapNone/>", `<wp:wrapSquare wrapText="largest"/>`, false, 0, 0},
		{"duplicateWrap", "<wp:wrapNone/>", `<wp:wrapNone/><wp:wrapSquare wrapText="bothSides"/>`, false, 0, 0},
		{"relative", "relativeFrom=\"page\"", "relativeFrom=\"paragraph\"", false, 0, 0},
		{"collision", "allowOverlap=\"1\"", "allowOverlap=\"0\"", false, 0, 0},
		{"unknown", "locked=\"0\"", "locked=\"0\" hidden=\"1\"", false, 0, 0},
		{"duplicate", "relativeHeight=\"7\"", "relativeHeight=\"7\" relativeHeight=\"8\"", false, 0, 0},
		{"offsetAttribute", "<wp:posOffset>", "<wp:posOffset unknown=\"1\">", false, 0, 0},
		{"negative", ">914400<", ">-914400<", false, 0, 0},
		{"orderOverflow", "relativeHeight=\"7\"", "relativeHeight=\"4294967296\"", false, 0, 0},
		{"distanceTop", "distT=\"0\"", "distT=\"1\"", false, 0, 0},
		{"distanceBottom", "distB=\"0\"", "distB=\"1\"", false, 0, 0},
		// distL/distR only widen the wrap region horizontally, which the square
		// wrap interval reproduces exactly, so they are projected rather than refused.
		{"distanceLeft", "distL=\"0\"", "distL=\"114300\"", true, 0, 0},
		{"distanceRight", "distR=\"0\"", "distR=\"114300\"", true, 0, 0},
		{"distanceLeftNegative", "distL=\"0\"", "distL=\"-1\"", false, 0, 0},
		{"distanceLeftUnbounded", "distL=\"0\"", "distL=\"91440001\"", false, 0, 0},
		{"distanceLeftLexical", "distL=\"0\"", "distL=\" 1 \"", false, 0, 0},
		{"columnOrigin", "<wp:positionH relativeFrom=\"page\">", "<wp:positionH relativeFrom=\"column\">", true, 0, 0},
		{"marginOrigin", "<wp:positionH relativeFrom=\"page\">", "<wp:positionH relativeFrom=\"margin\">", true, 0, 0},
		{"paragraphOrigin", "<wp:positionV relativeFrom=\"page\">", "<wp:positionV relativeFrom=\"paragraph\">", true, 0, 0},
		{"horizontalParagraphOrigin", "<wp:positionH relativeFrom=\"page\">", "<wp:positionH relativeFrom=\"paragraph\">", false, 0, 0},
		{"verticalColumnOrigin", "<wp:positionV relativeFrom=\"page\">", "<wp:positionV relativeFrom=\"column\">", false, 0, 0},
		{"unmodeledOrigin", "<wp:positionH relativeFrom=\"page\">", "<wp:positionH relativeFrom=\"leftMargin\">", false, 0, 0},
		{"columnOriginNegative", "<wp:positionH relativeFrom=\"page\"><wp:posOffset>914400<", "<wp:positionH relativeFrom=\"column\"><wp:posOffset>-4445<", true, -4445, 1270000},
		{"originOffsetUnbounded", "<wp:positionH relativeFrom=\"page\"><wp:posOffset>914400<", "<wp:positionH relativeFrom=\"column\"><wp:posOffset>91440001<", false, 0, 0},
		{"inertRelativeSize", "<wp:wrapNone/>", "<wp:wrapNone/><wp14:sizeRelH xmlns:wp14=\"" + wordDrawing2010 + "\" relativeFrom=\"page\"><wp14:pctWidth>0</wp14:pctWidth></wp14:sizeRelH>", true, 0, 0},
		{"activeRelativeSize", "<wp:wrapNone/>", "<wp:wrapNone/><wp14:sizeRelH xmlns:wp14=\"" + wordDrawing2010 + "\" relativeFrom=\"page\"><wp14:pctWidth>50000</wp14:pctWidth></wp14:sizeRelH>", false, 0, 0},
		{"duplicateRelativeSize", "<wp:wrapNone/>", "<wp:wrapNone/><wp14:sizeRelH xmlns:wp14=\"" + wordDrawing2010 + "\" relativeFrom=\"page\"><wp14:pctWidth>0</wp14:pctWidth></wp14:sizeRelH><wp14:sizeRelH xmlns:wp14=\"" + wordDrawing2010 + "\" relativeFrom=\"page\"><wp14:pctWidth>0</wp14:pctWidth></wp14:sizeRelH>", false, 0, 0},
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
			wantX, wantY := tc.wantX, tc.wantY
			if wantX == 0 && wantY == 0 {
				wantX, wantY = 914400, 1270000
			}
			if drawing == nil || drawing.FloatingLayer == nil || *drawing.FloatingLayer != "behind" || drawing.StackingOrder == nil || *drawing.StackingOrder != 7 || *drawing.XEMU != wantX || *drawing.YEMU != wantY {
				t.Fatalf("floating source projection: %#v", drawing)
			}
			if drawing.EditPolicy.Mode != "read-only" {
				t.Fatal("drawing edit authority widened")
			}
		})
	}
}
