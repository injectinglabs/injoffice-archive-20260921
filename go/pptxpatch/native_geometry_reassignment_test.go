package pptxpatch

import (
	"math"
	"math/big"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"
)

func TestNativeGeometrySequentialReassignmentShadows(t *testing.T) {
	g, _ := newNativeGeometryGuides(10000, 5000)
	err := g.evaluate([]nativeGeometryGuide{{"a", "*/ 1 1 3"}, {"old", "val a"}, {"a", "*/ a 3 1"}, {"a", "+- a 2 0"}, {"result", "+- a 0 old"}})
	if err != nil {
		t.Fatal(err)
	}
	if g.exact["a"].Cmp(big.NewRat(3, 1)) != 0 || g.exact["old"].Cmp(big.NewRat(1, 3)) != 0 || g.exact["result"].Cmp(big.NewRat(8, 3)) != 0 {
		t.Fatal("assignment did not retain exact earlier values")
	}
	// cos(90+1/60000 degrees), then doubled; alias retains its old enclosure.
	if err = g.evaluate([]nativeGeometryGuide{{"wave", "cos 1000000 5400001"}, {"alias", "val wave"}}); err != nil {
		t.Fatal(err)
	}
	oldInterval, oldSymbol := g.intervals["wave"], g.symbols["wave"]
	if err = g.evaluate([]nativeGeometryGuide{{"wave", "*/ wave 2 1"}}); err != nil {
		t.Fatal(err)
	}
	reference := -0.5817764173314349876
	if g.exact["wave"] != nil || g.intervals["wave"].lo > reference || g.intervals["wave"].hi < reference || math.Abs(g.values["wave"]-reference) > 1e-8 {
		t.Fatal("self-reassignment lost its input interval")
	}
	if g.intervals["alias"] != oldInterval || g.symbols["alias"] != oldSymbol || g.symbols["wave"] == oldSymbol {
		t.Fatal("reassignment retroactively changed alias identity")
	}
	if err = g.evaluate([]nativeGeometryGuide{{"wave", "val 7"}}); err != nil || g.exact["wave"].Cmp(big.NewRat(7, 1)) != 0 || g.intervalOperand("wave") != (nativeGeometryInterval{7, 7}) {
		t.Fatal("exact overwrite retained obsolete uncertainty")
	}
}
func TestNativeGeometryReassignmentRefusalAndAtomicShadow(t *testing.T) {
	g, _ := newNativeGeometryGuides(10000, 5000)
	if err := g.evaluate([]nativeGeometryGuide{{"a", "sin 1000 60000"}}); err != nil {
		t.Fatal(err)
	}
	oldValue, oldExact, oldInterval, oldSymbol := g.values["a"], g.exact["a"], g.intervals["a"], g.symbols["a"]
	for _, formula := range []string{"*/ a 1 0", "val missing", "val 9007199254740992", "?: a missing 1"} {
		if err := g.evaluate([]nativeGeometryGuide{{"a", formula}}); err == nil {
			t.Fatalf("accepted %s", formula)
		}
		if g.values["a"] != oldValue || g.exact["a"] != oldExact || g.intervals["a"] != oldInterval || g.symbols["a"] != oldSymbol {
			t.Fatal("failed assignment changed a shadow")
		}
	}
	for _, name := range []string{"w", "hd3", "cd4", "1", "1e2", "+1", "NaN", ""} {
		if err := g.evaluate([]nativeGeometryGuide{{name, "val 1"}}); err == nil {
			t.Fatalf("accepted reserved/ambiguous name %q", name)
		}
	}
	if err := g.evaluate([]nativeGeometryGuide{{"future", "+- future 1 0"}}); err == nil {
		t.Fatal("initial self-reference accepted")
	}
}

const nativeReassignmentGuides = `<a:avLst><a:gd name="adj" fmla="val 125000"/><a:gd name="adj" fmla="val 250000"/></a:avLst><a:gdLst><a:gd name="edge" fmla="val adj"/><a:gd name="left" fmla="val edge"/><a:gd name="edge" fmla="+- edge 1500000 0"/><a:gd name="right" fmla="val edge"/><a:gd name="edge" fmla="+- edge 250000 0"/></a:gdLst>`
const nativeReassignmentPath = `<a:rect l="left" t="250000" r="right" b="edge"/><a:pathLst><a:path><a:moveTo><a:pt x="left" y="250000"/></a:moveTo><a:lnTo><a:pt x="right" y="250000"/></a:lnTo><a:lnTo><a:pt x="edge" y="edge"/></a:lnTo><a:lnTo><a:pt x="left" y="edge"/></a:lnTo><a:close/></a:path></a:pathLst>`

func nativeReassignmentFixture(t *testing.T, strict, negative bool) []byte {
	t.Helper()
	guides := nativeReassignmentGuides
	if negative {
		guides = strings.Replace(guides, `fmla="val adj"`, `fmla="val notYet"`, 1)
	}
	shape := nativeAutoShapeXMLWithNameAndGeometry(3, "Sequential custom guides", `<a:custGeom>`+guides+nativeReassignmentPath+`</a:custGeom>`, `<a:solidFill><a:srgbClr val="336699"/></a:solidFill>`, nativeAutoShapeNoLine("flat", `<a:round/>`), "")
	shape = strings.Replace(shape, `cx="1000000" cy="500000"`, `cx="4000000" cy="3000000"`, 1)
	return nativeAutoShapeFixture(t, strict, shape)
}
func TestNativeGeometryReassignmentSourceBothDialects(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, negative := range []bool{false, true} {
			input := nativeReassignmentFixture(t, strict, negative)
			before := append([]byte(nil), input...)
			deck, err := ExtractNativePPTX(input, nativeTestExtractOptions())
			if err != nil {
				t.Fatal(err)
			}
			shape := nativeFixtureAutoShapes(deck.Slides[0])[0]
			if negative {
				if shape.Geometry != nil || shape.Compatibility.Status != NativeCompatibilityStatusRefused {
					t.Fatal("forward reference source admitted")
				}
				continue
			}
			if shape.Geometry == nil || shape.Source == nil || shape.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
				t.Fatal("reassigned source geometry not read-only qualified")
			}
			commands := shape.Geometry.Paths[0].Commands
			if *commands[0].X != 250000 || *commands[1].X != 1750000 || *commands[2].X != 2000000 || *commands[2].Y != 2000000 || shape.Geometry.TextRect.CX != 1500000 {
				t.Fatalf("incorrect source-order coordinates: %+v", shape.Geometry)
			}
			if issues := ValidateNativePPTX(deck); len(issues) != 0 {
				t.Fatal(issues)
			}
			if !reflect.DeepEqual(before, input) {
				t.Fatal("source input changed")
			}
		}
	}
}
func TestNativeGeometryReassignmentCountsAllAssignments(t *testing.T) {
	guides := strings.Repeat(`<a:gd name="same" fmla="val 1"/>`, 600)
	body := `<a:avLst>` + guides + `</a:avLst><a:gdLst>` + guides + `</a:gdLst><a:pathLst><a:path>` + nativeGeometryTestMove + `</a:path></a:pathLst>`
	if _, err := nativeGeometryTestEvaluate(t, body); err == nil || !strings.Contains(err.Error(), "guide budget") {
		t.Fatalf("reused names bypassed operation budget: %v", err)
	}
}
func TestNativeGeometryReassignmentBrowserFixture(t *testing.T) {
	dir := os.Getenv("INJOFFICE_PPTX_GUIDE_FIXTURE_DIR")
	if dir == "" {
		t.Skip("optional browser fixture")
	}
	for _, item := range []struct {
		name     string
		negative bool
	}{{"guide-order", false}, {"guide-order-refused", true}} {
		input := nativeReassignmentFixture(t, false, item.negative)
		deck, err := ExtractNativePPTX(input, nativeTestExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		encoded, err := MarshalNativePPTXJSON(deck)
		if err != nil {
			t.Fatal(err)
		}
		if err = os.WriteFile(filepath.Join(dir, item.name+".pptx"), input, 0600); err != nil {
			t.Fatal(err)
		}
		if err = os.WriteFile(filepath.Join(dir, item.name+"-go.json"), encoded, 0600); err != nil {
			t.Fatal(err)
		}
	}
}

func TestNativeGeometryReassignmentAcrossAdjustmentAndGuideLists(t *testing.T) {
	geometry, err := nativeGeometryTestEvaluate(t, `<a:avLst><a:gd name="a" fmla="val 125"/></a:avLst><a:gdLst><a:gd name="before" fmla="val a"/><a:gd name="a" fmla="+- a 250 0"/></a:gdLst><a:pathLst><a:path><a:moveTo><a:pt x="before" y="0"/></a:moveTo><a:lnTo><a:pt x="a" y="h"/></a:lnTo></a:path></a:pathLst>`)
	if err != nil {
		t.Fatal(err)
	}
	if *geometry.Paths[0].Commands[0].X != 125 || *geometry.Paths[0].Commands[1].X != 375 {
		t.Fatal("cross-list replacement lost prior binding")
	}
}
