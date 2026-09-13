package pptxpatch

import (
	"fmt"
	"math"
	"strconv"
	"strings"
)

// Geometry evaluation stays in floating point until its public EMU boundary.
// Budgets are independent of XML size limits and include adjustment guides.
const nativeGeometryMaxGuides = 1024
const nativeGeometryMaxMagnitude = 9007199254740991.0
const nativeGeometryAngleUnit = math.Pi / 10800000

type nativeGeometryGuide struct{ Name, Formula string }
type nativeGeometryGuides map[string]float64

func newNativeGeometryGuides(width, height float64) (nativeGeometryGuides, error) {
	if !nativeGeometryFinite(width) || !nativeGeometryFinite(height) || width <= 0 || height <= 0 {
		return nil, fmt.Errorf("invalid geometry extent")
	}
	g := nativeGeometryGuides{"w": width, "h": height, "l": 0, "t": 0, "r": width, "b": height, "hc": width / 2, "vc": height / 2, "ss": math.Min(width, height), "ls": math.Max(width, height), "cd2": 10800000, "cd4": 5400000, "cd8": 2700000, "3cd4": 16200000, "3cd8": 8100000, "5cd8": 13500000, "7cd8": 18900000}
	for _, d := range []int{2, 3, 4, 5, 6, 8, 10} {
		g[fmt.Sprintf("wd%d", d)] = width / float64(d)
	}
	for _, d := range []int{2, 3, 4, 5, 6, 8} {
		g[fmt.Sprintf("hd%d", d)] = height / float64(d)
	}
	for _, d := range []int{2, 4, 6, 8, 16, 32} {
		g[fmt.Sprintf("ssd%d", d)] = g["ss"] / float64(d)
	}
	return g, nil
}

func nativeGeometryFinite(v float64) bool {
	return !math.IsNaN(v) && !math.IsInf(v, 0) && math.Abs(v) <= nativeGeometryMaxMagnitude
}
func (g nativeGeometryGuides) resolve(token string) (float64, error) {
	if v, ok := g[token]; ok {
		return v, nil
	}
	// DrawingML numeric operands are integer lexical tokens, not IEEE literals.
	if token == "" || strings.TrimSpace(token) != token {
		return 0, fmt.Errorf("invalid geometry operand")
	}
	v, err := strconv.ParseInt(token, 10, 64)
	if err != nil || !nativeGeometryFinite(float64(v)) {
		return 0, fmt.Errorf("unresolved geometry operand %q", token)
	}
	return float64(v), nil
}

func (g nativeGeometryGuides) evaluate(guides []nativeGeometryGuide) error {
	if len(guides) > nativeGeometryMaxGuides {
		return fmt.Errorf("geometry guide budget exceeded")
	}
	for _, guide := range guides {
		if guide.Name == "" || strings.ContainsAny(guide.Name, " \t\r\n") {
			return fmt.Errorf("invalid geometry guide name")
		}
		if _, numericErr := strconv.ParseFloat(guide.Name, 64); numericErr == nil {
			return fmt.Errorf("ambiguous numeric geometry guide name %q", guide.Name)
		}
		if _, exists := g[guide.Name]; exists {
			return fmt.Errorf("duplicate geometry guide %q", guide.Name)
		}
		value, err := g.formula(guide.Formula)
		if err != nil {
			return fmt.Errorf("guide %q: %w", guide.Name, err)
		}
		g[guide.Name] = value
	}
	return nil
}

func (g nativeGeometryGuides) formula(formula string) (float64, error) {
	fields := strings.Fields(formula)
	if len(fields) < 2 || len(fields) > 4 {
		return 0, fmt.Errorf("invalid geometry formula")
	}
	arity := map[string]int{"*/": 3, "+-": 3, "+/": 3, "?:": 3, "abs": 1, "at2": 2, "cat2": 3, "cos": 2, "max": 2, "min": 2, "mod": 3, "pin": 3, "sat2": 3, "sin": 2, "sqrt": 1, "tan": 2, "val": 1}
	n, ok := arity[fields[0]]
	if !ok || n != len(fields)-1 {
		return 0, fmt.Errorf("unsupported geometry formula %q", fields[0])
	}
	var a [3]float64
	for i := 0; i < n; i++ {
		v, err := g.resolve(fields[i+1])
		if err != nil {
			return 0, err
		}
		a[i] = v
	}
	x, y, z := a[0], a[1], a[2]
	var out float64
	switch fields[0] {
	case "*/":
		if z == 0 {
			return 0, fmt.Errorf("zero geometry divisor")
		}
		out = x * y / z
	case "+-":
		out = x + y - z
	case "+/":
		if z == 0 {
			return 0, fmt.Errorf("zero geometry divisor")
		}
		out = (x + y) / z
	case "?:":
		if x > 0 {
			out = y
		} else {
			out = z
		}
	case "abs":
		out = math.Abs(x)
	case "at2":
		if x == 0 && y == 0 {
			return 0, fmt.Errorf("undefined geometry angle")
		}
		out = math.Atan2(y, x) / nativeGeometryAngleUnit
	case "cat2":
		if y == 0 && z == 0 {
			return 0, fmt.Errorf("undefined geometry angle")
		}
		out = x * math.Cos(math.Atan2(z, y))
	case "cos":
		out = x * math.Cos(math.Mod(y, 21600000)*nativeGeometryAngleUnit)
	case "max":
		out = math.Max(x, y)
	case "min":
		out = math.Min(x, y)
	case "mod":
		out = math.Hypot(math.Hypot(x, y), z)
	case "pin":
		if y < x {
			out = x
		} else if y > z {
			out = z
		} else {
			out = y
		}
	case "sat2":
		if y == 0 && z == 0 {
			return 0, fmt.Errorf("undefined geometry angle")
		}
		out = x * math.Sin(math.Atan2(z, y))
	case "sin":
		out = x * math.Sin(math.Mod(y, 21600000)*nativeGeometryAngleUnit)
	case "sqrt":
		if x < 0 {
			return 0, fmt.Errorf("negative geometry square root")
		}
		out = math.Sqrt(x)
	case "tan":
		out = x * math.Tan(math.Mod(y, 21600000)*nativeGeometryAngleUnit)
	case "val":
		out = x
	}
	if !nativeGeometryFinite(out) {
		return 0, fmt.Errorf("geometry formula exceeds numeric bounds")
	}
	return out, nil
}
