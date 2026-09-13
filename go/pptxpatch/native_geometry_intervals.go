package pptxpatch

import (
	"fmt"
	"math"
	"math/big"
)

// Non-rational guides retain a conservative binary64 enclosure. Basic algebra
// rounds outwards. Elementary functions add sixteen representable steps for
// the Go math approximation. This is a disclosed preview qualification policy,
// not a proof of arbitrary transcendental identities.
type nativeGeometryInterval struct{ lo, hi float64 }

func nativeGeometryIntervalRound(lo, hi float64, steps int) nativeGeometryInterval {
	for i := 0; i < steps; i++ {
		lo = math.Nextafter(lo, math.Inf(-1))
		hi = math.Nextafter(hi, math.Inf(1))
	}
	return nativeGeometryInterval{lo, hi}
}
func nativeGeometryRatInterval(r *big.Rat) nativeGeometryInterval {
	f, _ := r.Float64()
	c := new(big.Rat).SetFloat64(f).Cmp(r)
	if c < 0 {
		return nativeGeometryInterval{f, math.Nextafter(f, math.Inf(1))}
	}
	if c > 0 {
		return nativeGeometryInterval{math.Nextafter(f, math.Inf(-1)), f}
	}
	return nativeGeometryInterval{f, f}
}
func (g nativeGeometryGuides) intervalOperand(token string) nativeGeometryInterval {
	if exact := g.exactOperand(token); exact != nil {
		return nativeGeometryRatInterval(exact)
	}
	return g.intervals[token]
}
func nativeGeometryIntervalMul(a, b nativeGeometryInterval) nativeGeometryInterval {
	values := []float64{a.lo * b.lo, a.lo * b.hi, a.hi * b.lo, a.hi * b.hi}
	lo, hi := values[0], values[0]
	for _, v := range values[1:] {
		lo = math.Min(lo, v)
		hi = math.Max(hi, v)
	}
	return nativeGeometryIntervalRound(lo, hi, 1)
}
func nativeGeometryIntervalAbs(a nativeGeometryInterval) nativeGeometryInterval {
	if a.lo >= 0 {
		return a
	}
	if a.hi <= 0 {
		return nativeGeometryInterval{-a.hi, -a.lo}
	}
	return nativeGeometryInterval{0, math.Max(-a.lo, a.hi)}
}
func (g nativeGeometryGuides) intervalFormula(fields []string, exact *big.Rat) (nativeGeometryInterval, error) {
	if exact != nil {
		return nativeGeometryRatInterval(exact), nil
	}
	a := make([]nativeGeometryInterval, len(fields)-1)
	for i, token := range fields[1:] {
		a[i] = g.intervalOperand(token)
	}
	x := a[0]
	var y, z nativeGeometryInterval
	if len(a) > 1 {
		y = a[1]
	}
	if len(a) > 2 {
		z = a[2]
	}
	add := func(a, b nativeGeometryInterval) nativeGeometryInterval {
		return nativeGeometryIntervalRound(a.lo+b.lo, a.hi+b.hi, 1)
	}
	div := func(a, b nativeGeometryInterval) (nativeGeometryInterval, error) {
		if b.lo <= 0 && b.hi >= 0 {
			return nativeGeometryInterval{}, fmt.Errorf("uncertain geometry divisor")
		}
		return nativeGeometryIntervalMul(a, nativeGeometryIntervalRound(1/b.hi, 1/b.lo, 1)), nil
	}
	root := func(a nativeGeometryInterval) (nativeGeometryInterval, error) {
		if a.lo < 0 {
			return nativeGeometryInterval{}, fmt.Errorf("uncertain geometry square-root domain")
		}
		return nativeGeometryIntervalRound(math.Sqrt(a.lo), math.Sqrt(a.hi), 16), nil
	}
	switch fields[0] {
	case "val":
		return x, nil
	case "+-":
		return add(add(x, y), nativeGeometryInterval{-z.hi, -z.lo}), nil
	case "*/":
		return div(nativeGeometryIntervalMul(x, y), z)
	case "+/":
		return div(add(x, y), z)
	case "?:":
		if x.lo > 0 {
			return y, nil
		}
		if x.hi <= 0 {
			return z, nil
		}
		if fields[2] == fields[3] {
			return y, nil
		}
		return nativeGeometryInterval{}, fmt.Errorf("uncertain transcendental geometry branch %s [%g,%g]", fields[1], x.lo, x.hi)
	case "abs":
		return nativeGeometryIntervalAbs(x), nil
	case "min":
		return nativeGeometryInterval{math.Min(x.lo, y.lo), math.Min(x.hi, y.hi)}, nil
	case "max":
		return nativeGeometryInterval{math.Max(x.lo, y.lo), math.Max(x.hi, y.hi)}, nil
	case "pin":
		// General interval hull of every branch potentially selected.
		if y.lo >= x.hi && y.hi <= z.lo {
			return y, nil
		}
		if y.hi < x.lo {
			return x, nil
		}
		if y.lo >= x.hi && y.lo > z.hi {
			return z, nil
		}
		return nativeGeometryInterval{math.Min(x.lo, math.Min(y.lo, z.lo)), math.Max(x.hi, math.Max(y.hi, z.hi))}, nil
	case "sqrt":
		return root(x)
	case "mod":
		x = nativeGeometryIntervalAbs(x)
		y = nativeGeometryIntervalAbs(y)
		z = nativeGeometryIntervalAbs(z)
		return nativeGeometryIntervalRound(math.Hypot(math.Hypot(x.lo, y.lo), z.lo), math.Hypot(math.Hypot(x.hi, y.hi), z.hi), 16), nil
	case "cat2", "sat2":
		ay, az := nativeGeometryIntervalAbs(y), nativeGeometryIntervalAbs(z)
		radius := nativeGeometryIntervalRound(math.Hypot(ay.lo, az.lo), math.Hypot(ay.hi, az.hi), 16)
		component := y
		if fields[0] == "sat2" {
			component = z
		}
		ratio, err := div(component, radius)
		if err != nil {
			return nativeGeometryInterval{}, err
		}
		return nativeGeometryIntervalMul(x, ratio), nil
	case "sin", "cos", "tan":
		mid := y.lo + (y.hi-y.lo)/2
		angle := math.Mod(mid, 21600000) * nativeGeometryAngleUnit
		delta := (y.hi - y.lo) / 2 * nativeGeometryAngleUnit
		// Enclose midpoint/reduction and multiplication by the rounded pi constant.
		// Absolute radian error matters near trig zeros, where output ULPs shrink.
		delta += 16 * 2.220446049250313e-16 * (math.Abs(angle) + 1)
		if fields[0] == "tan" {
			// Tangent's pole must stay outside the whole angle interval.
			center := math.Remainder(angle, math.Pi)
			if delta >= math.Pi/2-math.Abs(center) {
				return nativeGeometryInterval{}, fmt.Errorf("uncertain geometry tangent domain")
			}
			return nativeGeometryIntervalMul(x, nativeGeometryIntervalRound(math.Tan(center-delta), math.Tan(center+delta), 16)), nil
		}
		value := math.Sin(angle)
		if fields[0] == "cos" {
			value = math.Cos(angle)
		}
		trig := nativeGeometryIntervalRound(math.Max(-1, value-delta), math.Min(1, value+delta), 16)
		return nativeGeometryIntervalMul(x, trig), nil
	case "at2":
		midx, midy := x.lo+(x.hi-x.lo)/2, y.lo+(y.hi-y.lo)/2
		uncertainty := math.Hypot((x.hi-x.lo)/2, (y.hi-y.lo)/2)
		radius := math.Hypot(midx, midy)
		if uncertainty >= radius {
			return nativeGeometryInterval{}, fmt.Errorf("uncertain geometry angle domain")
		}
		angle := math.Atan2(midy, midx) / nativeGeometryAngleUnit
		delta := math.Asin(uncertainty/radius) / nativeGeometryAngleUnit
		if angle-delta < -10800000 || angle+delta > 10800000 {
			return nativeGeometryInterval{-10800000, 10800000}, nil
		}
		return nativeGeometryIntervalRound(angle-delta, angle+delta, 16), nil
	}
	return nativeGeometryInterval{}, fmt.Errorf("unsupported interval geometry formula")
}

// Before the existing final integer rounding, uncertain source coordinates must
// be localized within one eighth EMU after path-space scaling. Quantization is
// separate; this does not claim exact transcendental coordinates.
const nativeGeometryMaxOutputUncertainty = 0.125

func (g nativeGeometryGuides) outputUncertainty(token string, scale float64) (float64, error) {
	value, err := g.resolve(token)
	if err != nil {
		return 0, err
	}
	bounds := g.intervalOperand(token)
	uncertainty := math.Max(math.Abs(value-bounds.lo), math.Abs(bounds.hi-value)) * math.Abs(scale)
	if math.IsNaN(uncertainty) || math.IsInf(uncertainty, 0) || uncertainty > nativeGeometryMaxOutputUncertainty {
		return 0, fmt.Errorf("geometry output uncertainty exceeds one eighth EMU")
	}
	return uncertainty, nil
}
func (g nativeGeometryGuides) qualifyOutputAttribute(node *nativeXMLNode, key string, scale float64) (float64, error) {
	token, ok := exactNativeAttr(node, "", key)
	if !ok {
		return 0, fmt.Errorf("missing geometry %s", key)
	}
	return g.outputUncertainty(token, scale)
}
