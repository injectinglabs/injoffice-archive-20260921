package pptxpatch

import (
	"math/big"
	"testing"
)

func TestNativeGeometryRationalBranchesAndRoots(t *testing.T) {
	g, _ := newNativeGeometryGuides(999983, 777779)
	if err := g.evaluate([]nativeGeometryGuide{{"q", "*/ w h 3"}, {"back", "*/ q 3 h"}, {"zero", "+- back 0 w"}, {"choice", "?: zero 100 200"}, {"root", "sqrt 81"}, {"hypot", "mod 3 4 0"}, {"cardinal", "cos w cd4"}}); err != nil {
		t.Fatal(err)
	}
	for key, want := range map[string]int64{"zero": 0, "choice": 200, "root": 9, "hypot": 5, "cardinal": 0} {
		if g.exact[key] == nil || g.exact[key].Cmp(big.NewRat(want, 1)) != 0 {
			t.Fatalf("%s = %v", key, g.exact[key])
		}
	}
	if g.exact["wd3"].Cmp(big.NewRat(999983, 3)) != 0 {
		t.Fatal("builtin fraction rounded early")
	}
	if _, err := nativeGeometryRationalBound(new(big.Rat).SetInt(new(big.Int).Lsh(big.NewInt(1), nativeGeometryMaxRationalBits))); err == nil {
		t.Fatal("unbounded numerator")
	}
	if _, err := nativeGeometryRationalBound(new(big.Rat).SetFrac(big.NewInt(1), new(big.Int).Lsh(big.NewInt(1), nativeGeometryMaxRationalBits))); err == nil {
		t.Fatal("unbounded denominator")
	}
}
