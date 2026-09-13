package pptxpatch

import (
	"fmt"
	"math"
	"math/big"
)

const nativeGeometryMaxPaths = 128
const nativeGeometryMaxPathCommands = 512
const nativeGeometryMaxCommands = 8192

type nativeGeometryPoint struct{ x, y float64 }
type nativeGeometryPathBuilder struct {
	sx, sy     float64
	pen, start nativeGeometryPoint
	hasPen     bool
	commands   []NativeGeometryCommand
}

func nativeGeometryRound(v float64) (*int64, error) {
	if !nativeGeometryFinite(v) {
		return nil, fmt.Errorf("geometry coordinate exceeds numeric bounds")
	}
	rounded := int64(math.Round(v))
	return &rounded, nil
}
func (p *nativeGeometryPathBuilder) point(x, y float64) (*int64, *int64, error) {
	xx, err := nativeGeometryRound(x * p.sx)
	if err != nil {
		return nil, nil, err
	}
	yy, err := nativeGeometryRound(y * p.sy)
	return xx, yy, err
}
func (p *nativeGeometryPathBuilder) append(c NativeGeometryCommand) error {
	if len(p.commands) >= nativeGeometryMaxPathCommands {
		return fmt.Errorf("geometry path command budget exceeded")
	}
	p.commands = append(p.commands, c)
	return nil
}
func (p *nativeGeometryPathBuilder) vertices(kind string, points ...nativeGeometryPoint) error {
	required := map[string]int{"moveTo": 1, "lineTo": 1, "quadBezierTo": 2, "cubicBezierTo": 3}
	if n, ok := required[kind]; !ok || len(points) != n {
		return fmt.Errorf("invalid geometry command")
	}
	if kind != "moveTo" && !p.hasPen {
		return fmt.Errorf("geometry command before moveTo")
	}
	c := NativeGeometryCommand{Kind: kind}
	var err error
	c.X, c.Y, err = p.point(points[len(points)-1].x, points[len(points)-1].y)
	if err != nil {
		return err
	}
	if len(points) > 1 {
		c.X1, c.Y1, err = p.point(points[0].x, points[0].y)
		if err != nil {
			return err
		}
	}
	if len(points) > 2 {
		c.X2, c.Y2, err = p.point(points[1].x, points[1].y)
		if err != nil {
			return err
		}
	}
	if err = p.append(c); err != nil {
		return err
	}
	p.pen = points[len(points)-1]
	p.hasPen = true
	if kind == "moveTo" {
		p.start = p.pen
	}
	return nil
}
func (p *nativeGeometryPathBuilder) close() error {
	if !p.hasPen {
		return fmt.Errorf("geometry close before moveTo")
	}
	if err := p.append(NativeGeometryCommand{Kind: "close"}); err != nil {
		return err
	}
	p.pen = p.start
	return nil
}

// DrawingML angles are polar angles: a ray at 45 degrees meets a non-square
// ellipse on y=x. SVG's ellipse parameter is different. Convert before deriving
// endpoints; path-space scaling happens only afterwards (ECMA 20.1.9.4).
func nativeGeometryEllipsePoint(rx, ry, angle float64) nativeGeometryPoint {
	t := math.Atan2(rx*math.Sin(angle), ry*math.Cos(angle))
	return nativeGeometryPoint{rx * math.Cos(t), ry * math.Sin(t)}
}
func (p *nativeGeometryPathBuilder) arc(rx, ry, start, sweep float64) error {
	if !p.hasPen {
		return fmt.Errorf("geometry arc before moveTo")
	}
	if !nativeGeometryFinite(rx) || !nativeGeometryFinite(ry) || rx <= 0 || ry <= 0 || !nativeGeometryFinite(start) || !nativeGeometryFinite(sweep) || math.Abs(sweep) > 21600000 {
		return fmt.Errorf("invalid geometry arc")
	}
	if sweep == 0 {
		return nil
	}
	start = math.Mod(start, 21600000) * nativeGeometryAngleUnit
	sw := sweep * nativeGeometryAngleUnit
	offset := nativeGeometryEllipsePoint(rx, ry, start)
	center := nativeGeometryPoint{p.pen.x - offset.x, p.pen.y - offset.y}
	// At most half a revolution per segment avoids SVG's coincident-endpoint
	// full-circle omission and makes signed winding explicit.
	segments := int(math.Ceil(math.Abs(sweep) / 10800000))
	for i := 1; i <= segments; i++ {
		next := start + sw*float64(i)/float64(segments)
		offset = nativeGeometryEllipsePoint(rx, ry, next)
		end := nativeGeometryPoint{center.x + offset.x, center.y + offset.y}
		x, y, err := p.point(end.x, end.y)
		if err != nil {
			return err
		}
		rrx, rry, err := p.point(rx, ry)
		if err != nil {
			return err
		}
		if *rrx <= 0 || *rry <= 0 {
			return fmt.Errorf("geometry arc radius collapsed after EMU rounding")
		}
		px, py, err := p.point(p.pen.x, p.pen.y)
		if err != nil {
			return err
		}
		if !nativeGeometryArcRadiiFit(*px, *py, *x, *y, *rrx, *rry) {
			return fmt.Errorf("rounded geometry arc requires implicit radius correction")
		}
		large, clockwise := false, sw > 0
		if err = p.append(NativeGeometryCommand{Kind: "arcTo", X: x, Y: y, RX: rrx, RY: rry, LargeArc: &large, Clockwise: &clockwise}); err != nil {
			return err
		}
		p.pen = end
	}
	return nil
}

// SVG scales radii when the normalized endpoint chord exceeds a diameter.
// Test that predicate with integers so host floating-point epsilon cannot cause
// an unreported change in the evaluated geometry, including at extreme scales.
func nativeGeometryArcRadiiFit(x0, y0, x1, y1, rx, ry int64) bool {
	square := func(v *big.Int) *big.Int { return new(big.Int).Mul(v, v) }
	dx := new(big.Int).Sub(big.NewInt(x1), big.NewInt(x0))
	dy := new(big.Int).Sub(big.NewInt(y1), big.NewInt(y0))
	xx, yy := square(big.NewInt(rx)), square(big.NewInt(ry))
	lhs := new(big.Int).Add(new(big.Int).Mul(square(dx), yy), new(big.Int).Mul(square(dy), xx))
	rhs := new(big.Int).Mul(new(big.Int).Mul(xx, yy), big.NewInt(4))
	return lhs.Cmp(rhs) <= 0
}
