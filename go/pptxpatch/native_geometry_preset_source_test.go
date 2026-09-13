package pptxpatch

import (
	"fmt"
	"testing"
)

func TestNativePresetSourceAdjustments(t *testing.T) {
	for _, ns := range []string{nativePresetDrawingNS, "http://purl.oclc.org/ooxml/drawingml/main"} {
		parse := func(body string) (*NativeEvaluatedGeometry, error) {
			node, err := parseNativeXML([]byte(fmt.Sprintf(`<a:prstGeom xmlns:a="%s" prst="triangle">%s</a:prstGeom>`, ns, body)), "preset.xml")
			if err != nil {
				return nil, err
			}
			return evaluateNativePresetSource(node, ns, 4000000, 3000000)
		}
		geometry, err := parse(`<a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst>`)
		if err != nil {
			t.Fatal(err)
		}
		if *geometry.Paths[0].Commands[1].X != 1000000 {
			t.Fatal("source adjustment not evaluated")
		}
		normal, err := parse("")
		if err != nil {
			t.Fatal(err)
		}
		if *normal.Paths[0].Commands[1].X != 2000000 {
			t.Fatal("omitted avLst did not use defaults")
		}
		for _, bad := range []string{`<a:avLst><a:gd name="unknown" fmla="val 1"/></a:avLst>`, `<a:avLst><a:gd name="adj" fmla="*/ 1 2 3"/></a:avLst>`, `<a:avLst><a:gd name="adj" fmla="val 1"/><a:gd name="adj" fmla="val 2"/></a:avLst>`, `<a:avLst extra="1"/>`, `<a:avLst/><a:avLst/>`, `<a:extLst/>`, `<a:avLst><a:gd name="adj" fmla="val 9007199254740992"/></a:avLst>`} {
			if geometry, err := parse(bad); err == nil || geometry != nil {
				t.Fatalf("invalid source override accepted: %s", bad)
			}
		}
	}
}
