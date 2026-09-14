package pptxpatch

import (
	"fmt"
	"math"
	"testing"
)

func TestNativeGeometryGuideOperators(t *testing.T) {
	g, err := newNativeGeometryGuides(600, 400)
	if err != nil {
		t.Fatal(err)
	}
	tests := []struct {
		formula string
		want    float64
	}{
		{"*/ 12 5 4", 15}, {"+- 12 5 4", 13}, {"+/ 12 6 3", 6}, {"?: -1 7 9", 9}, {"?: 1 7 9", 7},
		{"abs -12", 12}, {"at2 -1 1", 8100000}, {"cat2 5 3 4", 3}, {"cos 6 cd2", -6},
		{"max -4 -2", -2}, {"min -4 -2", -4}, {"mod 2 3 6", 7}, {"pin 3 1 5", 3}, {"pin 3 4 5", 4}, {"pin 3 7 5", 5},
		{"sat2 5 3 4", 4}, {"sin 6 cd4", 6}, {"sqrt 81", 9}, {"tan 3 cd8", 3}, {"val wd3", 200}, {"val hd10", 40}, {"val wd12", 50}, {"val wd32", 18.75}, {"val cd3", 7200000},
	}
	for _, tt := range tests {
		t.Run(tt.formula, func(t *testing.T) {
			got, err := g.formula(tt.formula)
			if err != nil || math.Abs(got-tt.want) > 1e-8 {
				t.Fatalf("got %v/%v want %v", got, err, tt.want)
			}
		})
	}
}

func TestNativeGeometryGuideOrderAndRefusal(t *testing.T) {
	for _, formula := range []string{"val missing", "val 1.5", "val NaN", "val 9007199254740992", "val 1 extra", "nope 1", "*/ 1 2 0", "+/ 1 2 0", "sqrt -1", "tan 1 cd4", "at2 0 0", "val hd11"} {
		t.Run(formula, func(t *testing.T) {
			g, _ := newNativeGeometryGuides(600, 400)
			if _, err := g.formula(formula); err == nil {
				t.Fatal("accepted invalid formula")
			}
		})
	}
	g, _ := newNativeGeometryGuides(600, 400)
	if err := g.evaluate([]nativeGeometryGuide{{"a", "*/ w 1 3"}, {"b1", "+- a 5 0"}}); err != nil || g.values["b1"] != 205 {
		t.Fatalf("ordered evaluation: %v %v", g, err)
	}
	if err := g.evaluate([]nativeGeometryGuide{{"a", "+- a 1 0"}}); err != nil || g.values["a"] != 201 || g.values["b1"] != 205 {
		t.Fatal("sequential replacement did not preserve earlier results")
	}
	if err := g.evaluate([]nativeGeometryGuide{{"future", "val later"}, {"later", "val 1"}}); err == nil {
		t.Fatal("forward reference accepted")
	}
	if err := g.evaluate(make([]nativeGeometryGuide, nativeGeometryMaxGuides+1)); err == nil {
		t.Fatal("guide budget accepted")
	}
	for _, extent := range []float64{0, -1, math.NaN(), math.Inf(1), 1e16} {
		if _, err := newNativeGeometryGuides(extent, 10); err == nil {
			t.Fatalf("invalid extent %v", extent)
		}
	}
}

func TestNativeGeometryPeriodicGuideAngles(t *testing.T) {
	g, _ := newNativeGeometryGuides(600, 400)
	for _, op := range []string{"sin", "cos", "tan"} {
		for _, angle := range []int64{9007199254740991, -9007199254740991, 21600000, -21600000} {
			a, err := g.formula(fmt.Sprintf("%s 281474976710655 %d", op, angle))
			if err != nil {
				t.Fatal(err)
			}
			b, err := g.formula(fmt.Sprintf("%s 281474976710655 %d", op, angle%21600000))
			if err != nil || a != b {
				t.Fatalf("periodic angle %s %d: %v/%v %v", op, angle, a, b, err)
			}
		}
	}
}

func TestNativeGeometryIntermediateAndOutputBudgets(t *testing.T) {
	g, _ := newNativeGeometryGuides(4000000, 3000000)
	if err := g.evaluateWithIntermediateLimit([]nativeGeometryGuide{{"squared", "*/ w w 1"}, {"fourth", "*/ squared squared 1"}, {"restored", "*/ fourth 1 squared"}}, nativeGeometryMaxIntermediate); err != nil {
		t.Fatal(err)
	}
	if g.values["fourth"] <= nativeGeometryMaxMagnitude || g.values["restored"] != 16000000000000 {
		t.Fatal("bounded higher-order guides failed")
	}
	if _, err := nativeGeometryRound(g.values["fourth"]); err == nil {
		t.Fatal("intermediate leaked beyond output budget")
	}
	for i := 0; i < 12; i++ {
		name := fmt.Sprintf("grow%d", i)
		prior := "fourth"
		if i > 0 {
			prior = fmt.Sprintf("grow%d", i-1)
		}
		err := g.evaluateWithIntermediateLimit([]nativeGeometryGuide{{name, "*/ " + prior + " " + prior + " 1"}}, nativeGeometryMaxIntermediate)
		if err != nil {
			return
		}
	}
	t.Fatal("intermediate budget was not enforced")
}

func TestNativeGeometryCustomRejectsUnsafeCancellationIntermediate(t *testing.T) {
	g, _ := newNativeGeometryGuides(100000000, 100000000)
	// Exact arithmetic gives delta=1 and final=100000000. Binary64 would lose
	// the +1 at square=1e16 and silently produce final=0 without the guard.
	if err := g.evaluate([]nativeGeometryGuide{{"square", "*/ w w 1"}, {"increment", "+- square 1 0"}, {"delta", "+- increment 0 square"}, {"final", "*/ delta w 1"}}); err == nil {
		t.Fatal("unsafe arbitrary custom intermediate accepted")
	}
	if _, exists := g.values["square"]; exists {
		t.Fatal("unsafe result published")
	}
}
