package pptxpatch

import (
	"strings"
	"testing"
)

const nativeGeometryTestNS = "http://schemas.openxmlformats.org/drawingml/2006/main"

func nativeGeometryTestEvaluate(t *testing.T, body string) (*NativeEvaluatedGeometry, error) {
	t.Helper()
	node, err := parseNativeXML([]byte(`<a:custGeom xmlns:a="`+nativeGeometryTestNS+`">`+body+`</a:custGeom>`), "geometry.xml")
	if err != nil {
		return nil, err
	}
	return evaluateNativeCustomGeometry(node, nativeGeometryTestNS, 10000, 5000)
}

const nativeGeometryTestMove = `<a:moveTo><a:pt x="0" y="0"/></a:moveTo>`

func TestNativeGeometryCustomXML(t *testing.T) {
	geometry, err := nativeGeometryTestEvaluate(t, `<a:avLst><a:gd name="adj" fmla="val 1000"/></a:avLst><a:gdLst><a:gd name="inset" fmla="*/ adj 1 2"/></a:gdLst><a:ahLst/><a:cxnLst/><a:rect l="inset" t="0" r="w" b="h"/><a:pathLst><a:path w="100" h="100" fill="none" stroke="1">`+nativeGeometryTestMove+`<a:quadBezTo><a:pt x="25" y="100"/><a:pt x="50" y="50"/></a:quadBezTo><a:cubicBezTo><a:pt x="60" y="10"/><a:pt x="80" y="90"/><a:pt x="100" y="100"/></a:cubicBezTo><a:close/></a:path><a:path stroke="false"><a:moveTo><a:pt x="1000" y="0"/></a:moveTo><a:arcTo wR="1000" hR="500" stAng="0" swAng="21600000"/><a:close/></a:path></a:pathLst>`)
	if err != nil {
		t.Fatal(err)
	}
	if geometry.Profile != "drawingml-paths-v1" || geometry.TextRect.X != 500 || geometry.TextRect.CX != 9500 || len(geometry.Paths) != 2 {
		t.Fatalf("bad geometry: %+v", geometry)
	}
	p := geometry.Paths[0]
	if p.FillMode != "none" || !p.Stroke || *p.Commands[1].X1 != 2500 || *p.Commands[1].Y1 != 5000 || *p.Commands[2].X2 != 8000 {
		t.Fatal("path scaling/paint")
	}
	if geometry.Paths[1].Stroke || len(geometry.Paths[1].Commands) != 6 {
		t.Fatal("full-circle transport")
	}
}
func TestNativeGeometryCustomXMLRefusals(t *testing.T) {
	valid := `<a:pathLst><a:path>` + nativeGeometryTestMove + `<a:lnTo><a:pt x="w" y="h"/></a:lnTo></a:path></a:pathLst>`
	tests := []string{
		`<a:unknown/>` + valid, `<a:gdLst/><a:gdLst/>` + valid, `<a:gdLst/><a:avLst/>` + valid, `<a:ahLst><a:ahXY/></a:ahLst>` + valid,
		`<a:avLst><a:gd name="a1" fmla="*/ 1 2 3"/></a:avLst>` + valid,
		`<a:gdLst><a:gd name="1" fmla="val 42"/></a:gdLst>` + valid,
		`<a:rect l="0" t="0" r="0" b="h"/>` + valid,
		strings.Replace(valid, `<a:path>`, `<a:path fill="darken">`, 1), strings.Replace(valid, `<a:path>`, `<a:path extrusionOk="invalid">`, 1), strings.Replace(valid, `<a:path>`, `<a:path w="0">`, 1),
		strings.Replace(valid, nativeGeometryTestMove, "", 1), strings.Replace(valid, `<a:pt x="w" y="h"/>`, `<a:pt x="w" y="h"/><a:pt x="0" y="0"/>`, 1),
		strings.Replace(valid, `<a:pt x="w" y="h"/>`, `<a:pt x="w" y="h" extra="1"/>`, 1), strings.Replace(valid, `<a:lnTo>`, `<a:lnTo>bad`, 1),
		`<a:pathLst/>`,
	}
	for i, body := range tests {
		if g, err := nativeGeometryTestEvaluate(t, body); err == nil || g != nil {
			t.Fatalf("case %d partially accepted: %+v %v", i, g, err)
		}
	}
}
func TestNativeGeometryCombinedGuideBudget(t *testing.T) {
	var a, b strings.Builder
	for i := 0; i < 600; i++ {
		name := strings.Repeat("x", i+1)
		a.WriteString(`<a:gd name="a` + name + `" fmla="val 1"/>`)
		b.WriteString(`<a:gd name="b` + name + `" fmla="val 1"/>`)
	}
	body := `<a:avLst>` + a.String() + `</a:avLst><a:gdLst>` + b.String() + `</a:gdLst><a:pathLst><a:path>` + nativeGeometryTestMove + `</a:path></a:pathLst>`
	if _, err := nativeGeometryTestEvaluate(t, body); err == nil || !strings.Contains(err.Error(), "guide budget") {
		t.Fatalf("combined guide budget: %v", err)
	}
}
