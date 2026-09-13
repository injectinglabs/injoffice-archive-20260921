package pptxpatch

import (
	"strings"
	"testing"
)

func TestNativePresetCatalogInventory(t *testing.T) {
	names, err := nativePresetNames()
	if err != nil {
		t.Fatal(err)
	}
	if len(names) != 187 || names[0] != "accentBorderCallout1" || names[len(names)-1] != "wedgeRoundRectCallout" {
		t.Fatalf("catalog inventory %d", len(names))
	}
	for _, name := range names {
		if _, err := prepareNativePresetGeometry(name, nil); err != nil {
			t.Fatalf("%s: %v", name, err)
		}
	}
}
func TestNativePresetSequentialRename(t *testing.T) {
	source := `<a:custGeom xmlns:a="` + nativePresetDrawingNS + `"><a:gdLst><a:gd name="q" fmla="val 1"/><a:gd name="before" fmla="+- q 1 0"/><a:gd name="q" fmla="+- q 5 0"/><a:gd name="after" fmla="+- q before 0"/></a:gdLst><a:rect l="q" t="0" r="w" b="h"/><a:pathLst><a:path><a:moveTo><a:pt x="q" y="after"/></a:moveTo></a:path></a:pathLst></a:custGeom>`
	node, err := parseNativeXML([]byte(source), "test.xml")
	if err != nil {
		t.Fatal(err)
	}
	if err := normalizeNativePresetGuides(node); err != nil {
		t.Fatal(err)
	}
	g, err := evaluateNativeCustomGeometry(node, nativePresetDrawingNS, 100, 100)
	if err != nil {
		t.Fatal(err)
	}
	if g.TextRect.X != 6 || *g.Paths[0].Commands[0].X != 6 || *g.Paths[0].Commands[0].Y != 8 {
		t.Fatalf("sequential semantics changed: %+v", g)
	}
}
func TestNativePresetDefaultCoverageAudit(t *testing.T) {
	names, err := nativePresetNames()
	if err != nil {
		t.Fatal(err)
	}
	passed := 0
	failures := []string{}
	for _, name := range names {
		if _, err := evaluateNativePresetGeometry(name, nil, 4000000, 3000000); err != nil {
			failures = append(failures, name+": "+err.Error())
		} else {
			passed++
		}
	}
	t.Logf("default catalog coverage %d/%d; remaining:\n%s", passed, len(names), strings.Join(failures, "\n"))
	if passed != 187 {
		t.Fatal("incomplete default catalog geometry")
	}
}
func TestNativePresetAdjustmentNamesAndIsolation(t *testing.T) {
	if _, err := prepareNativePresetGeometry("unknown", nil); err == nil {
		t.Fatal("unknown preset accepted")
	}
	if _, err := prepareNativePresetGeometry("triangle", map[string]int64{"unknown": 1}); err == nil {
		t.Fatal("unknown adjustment accepted")
	}
	a, err := evaluateNativePresetGeometry("triangle", map[string]int64{"adj": 25000}, 4000000, 3000000)
	if err != nil {
		t.Fatal(err)
	}
	b, err := evaluateNativePresetGeometry("triangle", nil, 4000000, 3000000)
	if err != nil {
		t.Fatal(err)
	}
	if *a.Paths[0].Commands[1].X != 1000000 || *b.Paths[0].Commands[1].X != 2000000 {
		t.Fatal("override/default separation failed")
	}
}

func TestNativePresetExplicitErrata(t *testing.T) {
	for _, name := range []string{"circularArrow", "leftCircularArrow", "leftRightCircularArrow"} {
		node, err := prepareNativePresetGeometry(name, nil)
		if err != nil {
			t.Fatal(err)
		}
		guides, _ := newNativeGeometryGuides(400, 300)
		for _, list := range []string{"avLst", "gdLst"} {
			for _, gd := range nativeChildren(nativeChild(node, nativePresetDrawingNS, list), nativePresetDrawingNS, "gd") {
				key, _ := exactNativeAttr(gd, "", "name")
				formula, _ := exactNativeAttr(gd, "", "fmla")
				if err := guides.evaluate([]nativeGeometryGuide{{key, formula}}); err != nil {
					t.Fatal(err)
				}
			}
		}
		if guides.values["xB"] != guides.values["xH"]-guides.values["dxB"] {
			t.Fatal("catalog correction changed the defined three-operand subtraction")
		}
		if _, err := evaluateNativePresetGeometry(name, nil, 400, 300); err != nil {
			t.Fatalf("%s: %v", name, err)
		}
	}
	pie, err := evaluateNativePresetGeometry("pie", nil, 4000000, 3000000)
	if err != nil {
		t.Fatal(err)
	}
	// Independent inscribed-rectangle coordinates, not values copied from output.
	if pie.TextRect.X != 585786 || pie.TextRect.Y != 439340 || pie.TextRect.CX != 2828428 || pie.TextRect.CY != 2121320 {
		t.Fatalf("pie rectangle correction: %+v", pie.TextRect)
	}
	g, _ := newNativeGeometryGuides(400, 300)
	if _, err := g.formula("+- 10 0 3 0"); err == nil {
		t.Fatal("catalog repair leaked into source formula grammar")
	}
}

func TestNativePresetAllGeometryIndependentOfPendingToneModes(t *testing.T) {
	names, err := nativePresetNames()
	if err != nil {
		t.Fatal(err)
	}
	for _, name := range names {
		t.Run(name, func(t *testing.T) {
			node, err := prepareNativePresetGeometry(name, nil)
			if err != nil {
				t.Fatal(err)
			}
			// Numeric geometry acceptance is tested independently while shared fill-mode
			// contract integration is pending. Production never rewrites these modes.
			for _, path := range nativeChild(node, nativePresetDrawingNS, "pathLst").Children {
				for i, attr := range path.Attrs {
					if attr.Name.Local == "fill" && attr.Value != "none" {
						path.Attrs[i].Value = "norm"
					}
				}
			}
			g, err := evaluateNativePreparedPresetGeometry(node, 4000000, 3000000)
			if err != nil {
				t.Fatal(err)
			}
			if len(g.Paths) == 0 {
				t.Fatal("empty catalog geometry")
			}
		})
	}
}

func TestNativePresetCalloutExactZeroBranch(t *testing.T) {
	for _, name := range []string{"wedgeRectCallout", "wedgeRoundRectCallout"} {
		node, err := prepareNativePresetGeometry(name, map[string]int64{"adj1": 1, "adj2": 1})
		if err != nil {
			t.Fatal(err)
		}
		g, _ := newNativeGeometryGuides(999983, 777779)
		for _, list := range []string{"avLst", "gdLst"} {
			for _, gd := range nativeChild(node, nativePresetDrawingNS, list).Children {
				key, _ := exactNativeAttr(gd, "", "name")
				formula, _ := exactNativeAttr(gd, "", "fmla")
				if err := g.evaluateWithIntermediateLimit([]nativeGeometryGuide{{key, formula}}, nativeGeometryMaxIntermediate); err != nil {
					t.Fatal(err)
				}
			}
		}
		// dq=(w/100000)*h/w=h/100000=dyPos, so dz is exactly zero.
		if g.exact["dz"] == nil || g.exact["dz"].Sign() != 0 {
			t.Fatalf("%s nonzero dz: %v", name, g.exact["dz"])
		}
		if g.values["xr"] != 500001.49983 || g.values["yb"] != 777779 {
			t.Fatalf("%s wrong branch: xr=%v yb=%v", name, g.values["xr"], g.values["yb"])
		}
		geometry, err := evaluateNativePresetGeometry(name, map[string]int64{"adj1": 1, "adj2": 1}, 999983, 777779)
		if err != nil {
			t.Fatal(err)
		}
		right, bottom := 6, 10
		if name == "wedgeRoundRectCallout" {
			right, bottom = 8, 13
		}
		if *geometry.Paths[0].Commands[right].X != 500001 || *geometry.Paths[0].Commands[bottom].Y != 777779 {
			t.Fatalf("%s wrong final point", name)
		}
	}
}
