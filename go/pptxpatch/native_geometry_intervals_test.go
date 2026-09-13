package pptxpatch

import (
	"math/big"
	"strings"
	"testing"
)

func TestNativeGeometryTrigIntervalContainsIndependentReference(t *testing.T) {
	g, _ := newNativeGeometryGuides(4000000, 3000000)
	if err := g.evaluate([]nativeGeometryGuide{{"nearZero", "cos 1 5400001"}}); err != nil {
		t.Fatal(err)
	}
	// cos((90+1/60000)degrees), independently evaluated with 80 decimal digits.
	reference, _ := new(big.Rat).SetString("-0.0000002908882086657174938")
	b := g.intervals["nearZero"]
	if new(big.Rat).SetFloat64(b.lo).Cmp(reference) > 0 || new(big.Rat).SetFloat64(b.hi).Cmp(reference) < 0 {
		t.Fatalf("reference outside %+v", b)
	}
}
func TestNativeGeometryUncertainBranchAndOutputRefusal(t *testing.T) {
	g, _ := newNativeGeometryGuides(4000000, 3000000)
	g.values["uncertain"] = 0
	g.intervals["uncertain"] = nativeGeometryInterval{-1e-12, 1e-12}
	if _, err := g.formula("?: uncertain 1 2"); err == nil || !strings.Contains(err.Error(), "uncertain") {
		t.Fatal("uncertain branch accepted")
	}
	if err := g.evaluateWithIntermediateLimit([]nativeGeometryGuide{{"large", "cos 281474976710655 5400001"}}, nativeGeometryMaxIntermediate); err != nil {
		t.Fatal(err)
	}
	if _, err := g.outputUncertainty("large", 1); err == nil {
		t.Fatal("wide final coordinate interval accepted")
	}
	if _, err := g.outputUncertainty("nearMissing", 1); err == nil {
		t.Fatal("unknown output accepted")
	}
}
func TestNativeGeometryEquivalentExpressionsKeepExactZero(t *testing.T) {
	g, _ := newNativeGeometryGuides(4000000, 3000000)
	if err := g.evaluate([]nativeGeometryGuide{{"irrational", "sqrt 2"}, {"a", "+- irrational 0 0"}, {"aliasB", "val irrational"}, {"zero", "+- a 0 aliasB"}, {"choice", "?: zero 1 2"}}); err != nil {
		t.Fatal(err)
	}
	if g.exact["zero"] == nil || g.exact["zero"].Sign() != 0 || g.values["choice"] != 2 {
		t.Fatal("equivalent algebra lost exact zero")
	}
}

func TestNativeGeometryAccumulatesArcConstructionUncertainty(t *testing.T) {
	arc := `<a:arcTo wR="10000000000" hR="10000000000" stAng="0" swAng="21600000"/>`
	evaluate := func(count int) error {
		xml := `<a:custGeom xmlns:a="` + nativeGeometryTestNS + `"><a:pathLst><a:path><a:moveTo><a:pt x="20000000000" y="10000000000"/></a:moveTo>` + strings.Repeat(arc, count) + `</a:path></a:pathLst></a:custGeom>`
		node, err := parseNativeXML([]byte(xml), "uncertainty.xml")
		if err != nil {
			return err
		}
		_, err = evaluateNativeCustomGeometry(node, nativeGeometryTestNS, 20000000000, 20000000000)
		return err
	}
	if err := evaluate(1); err != nil {
		t.Fatal(err)
	}
	if err := evaluate(60); err == nil || !strings.Contains(err.Error(), "accumulated path uncertainty") {
		t.Fatalf("arc chain failed to accumulate uncertainty: %v", err)
	}
}
