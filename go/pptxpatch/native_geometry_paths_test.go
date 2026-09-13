package pptxpatch

import (
	"math"
	"testing"
)

func TestNativeGeometryNonSquarePolarArc(t *testing.T) {
	// The 45-degree ray meets x²/100²+y²/50²=1 at (sqrt(2000),sqrt(2000)).
	q := nativeGeometryEllipsePoint(100, 50, math.Pi/4)
	if math.Abs(q.x-math.Sqrt(2000)) > 1e-10 || math.Abs(q.x-q.y) > 1e-10 {
		t.Fatalf("parametric angle used instead of polar: %+v", q)
	}
	p := nativeGeometryPathBuilder{sx: 100, sy: 100}
	if err := p.vertices("moveTo", nativeGeometryPoint{100, 0}); err != nil {
		t.Fatal(err)
	}
	if err := p.arc(100, 50, 0, 2700000); err != nil {
		t.Fatal(err)
	}
	a := p.commands[1]
	if *a.X != 4472 || *a.Y != 4472 || *a.RX != 10000 || *a.RY != 5000 || !*a.Clockwise || *a.LargeArc {
		t.Fatalf("wrong arc: %+v", a)
	}
}

func TestNativeGeometryFullArcAndSubpathState(t *testing.T) {
	p := nativeGeometryPathBuilder{sx: 1, sy: 1}
	_ = p.vertices("moveTo", nativeGeometryPoint{100, 0})
	if err := p.arc(100, 50, 0, -21600000); err != nil {
		t.Fatal(err)
	}
	if len(p.commands) != 3 || *p.commands[1].X != -100 || *p.commands[2].X != 100 || *p.commands[1].Clockwise {
		t.Fatalf("full arc winding: %+v", p.commands)
	}
	_ = p.vertices("lineTo", nativeGeometryPoint{40, 30})
	if err := p.close(); err != nil {
		t.Fatal(err)
	}
	if p.pen != p.start {
		t.Fatal("close does not restore subpath origin")
	}
	if err := p.arc(100, 50, 0, 5400000); err != nil {
		t.Fatal(err)
	}
	last := p.commands[len(p.commands)-1]
	if *last.X != 0 || *last.Y != 50 {
		t.Fatal("arc after close uses stale pen")
	}
}

func TestNativeGeometryBezierAndBounds(t *testing.T) {
	p := nativeGeometryPathBuilder{sx: 2, sy: 3}
	if err := p.vertices("lineTo", nativeGeometryPoint{}); err == nil {
		t.Fatal("line before move accepted")
	}
	if err := p.close(); err == nil {
		t.Fatal("close before move accepted")
	}
	_ = p.vertices("moveTo", nativeGeometryPoint{})
	if err := p.vertices("quadBezierTo", nativeGeometryPoint{1, 2}, nativeGeometryPoint{3, 4}); err != nil {
		t.Fatal(err)
	}
	if err := p.vertices("cubicBezierTo", nativeGeometryPoint{5, 6}, nativeGeometryPoint{7, 8}, nativeGeometryPoint{9, 10}); err != nil {
		t.Fatal(err)
	}
	q, c := p.commands[1], p.commands[2]
	if *q.X1 != 2 || *q.Y1 != 6 || *q.X != 6 || *q.Y != 12 || q.X2 != nil || *c.X2 != 14 || *c.Y2 != 24 || *c.X != 18 || *c.Y != 30 {
		t.Fatal("control point scaling")
	}
	if err := p.vertices("lineTo", nativeGeometryPoint{nativeGeometryMaxMagnitude, 0}); err == nil {
		t.Fatal("scaled overflow accepted")
	}
	if err := p.arc(0, 10, 0, 1); err == nil {
		t.Fatal("zero radius accepted")
	}
	if err := p.arc(10, 10, 0, 21600001); err == nil {
		t.Fatal("multi-revolution arc accepted")
	}
	for len(p.commands) < nativeGeometryMaxPathCommands {
		if err := p.close(); err != nil {
			t.Fatal(err)
		}
	}
	if err := p.close(); err == nil {
		t.Fatal("path budget accepted")
	}
}

func TestNativeGeometryArcRejectsImplicitRadiusCorrection(t *testing.T) {
	if nativeGeometryArcRadiiFit(0, 0, 3, 0, 1, 1) {
		t.Fatal("oversize chord accepted")
	}
	if !nativeGeometryArcRadiiFit(-9007199254740991, 0, 9007199254740991, 0, 9007199254740991, 1) {
		t.Fatal("exact large diameter refused")
	}
	if nativeGeometryArcRadiiFit(-9007199254740991, 0, 9007199254740991, 1, 9007199254740991, 1) {
		t.Fatal("extreme aspect oversize accepted")
	}
	p := nativeGeometryPathBuilder{sx: 1, sy: 1}
	_ = p.vertices("moveTo", nativeGeometryPoint{0, 0})
	if err := p.arc(1.4, 1.4, 0, 10800000); err == nil {
		t.Fatal("rounded arc changed radius silently")
	}
	p = nativeGeometryPathBuilder{sx: 0.01, sy: 1}
	_ = p.vertices("moveTo", nativeGeometryPoint{1, 0})
	if err := p.arc(1, 1, 0, 5400000); err == nil {
		t.Fatal("collapsed radius accepted")
	}
}
