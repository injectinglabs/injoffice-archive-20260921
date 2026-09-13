package pptxpatch

import "fmt"

func (v *nativeValidator) geometry(g NativeEvaluatedGeometry, p string) {
	if g.Profile != "drawingml-paths-v1" {
		v.add(p+".profile", "native.geometry", "unsupported geometry profile")
	}
	r := g.TextRect
	for key, value := range map[string]int64{"x": r.X, "y": r.Y, "cx": r.CX, "cy": r.CY} {
		v.integer(value, p+".textRect."+key)
	}
	if r.CX <= 0 || r.CY <= 0 || r.X > nativeMaxSafeInteger-r.CX || r.Y > nativeMaxSafeInteger-r.CY {
		v.add(p+".textRect", "native.geometry", "invalid geometry text rectangle")
	}
	if len(g.Paths) < 1 || len(g.Paths) > nativeGeometryMaxPaths {
		v.add(p+".paths", "native.geometryBudget", "invalid geometry path count")
		return
	}
	total := 0
	for i, path := range g.Paths {
		pp := fmt.Sprintf("%s.paths[%d]", p, i)
		if path.FillMode != "norm" && path.FillMode != "none" {
			v.add(pp+".fillMode", "native.geometry", "unsupported geometry fill")
		}
		total += len(path.Commands)
		if len(path.Commands) < 1 || len(path.Commands) > nativeGeometryMaxPathCommands || total > nativeGeometryMaxCommands {
			v.add(pp+".commands", "native.geometryBudget", "geometry command budget exceeded")
			continue
		}
		var x, y, sx, sy int64
		hasPen := false
		for j, c := range path.Commands {
			cp := fmt.Sprintf("%s.commands[%d]", pp, j)
			wanted := map[string][]string{"moveTo": {"x", "y"}, "lineTo": {"x", "y"}, "quadBezierTo": {"x", "y", "x1", "y1"}, "cubicBezierTo": {"x", "y", "x1", "y1", "x2", "y2"}, "arcTo": {"x", "y", "rx", "ry", "largeArc", "clockwise"}, "close": {}}
			fields, ok := wanted[c.Kind]
			if !ok {
				v.add(cp+".kind", "native.geometry", "unknown geometry command")
				continue
			}
			required := map[string]bool{}
			for _, name := range fields {
				required[name] = true
			}
			valid := true
			for name, value := range map[string]*int64{"x": c.X, "y": c.Y, "x1": c.X1, "y1": c.Y1, "x2": c.X2, "y2": c.Y2, "rx": c.RX, "ry": c.RY} {
				if required[name] != (value != nil) {
					v.add(cp+"."+name, "native.geometry", "command fields do not match kind")
					valid = false
				}
				if value != nil {
					v.integer(*value, cp+"."+name)
					if *value > nativeMaxSafeInteger || *value < -nativeMaxSafeInteger {
						valid = false
					}
				}
			}
			for name, value := range map[string]*bool{"largeArc": c.LargeArc, "clockwise": c.Clockwise} {
				if required[name] != (value != nil) {
					v.add(cp+"."+name, "native.geometry", "command fields do not match kind")
					valid = false
				}
			}
			if !valid {
				continue
			}
			if c.Kind != "moveTo" && !hasPen {
				v.add(cp, "native.geometry", "command precedes moveTo")
				continue
			}
			if c.Kind == "close" {
				x, y = sx, sy
				continue
			}
			if c.Kind == "arcTo" && (*c.RX <= 0 || *c.RY <= 0 || *c.LargeArc || !nativeGeometryArcRadiiFit(x, y, *c.X, *c.Y, *c.RX, *c.RY)) {
				v.add(cp, "native.geometry", "invalid arc radii or noncanonical long arc")
			}
			x, y = *c.X, *c.Y
			if c.Kind == "moveTo" {
				sx, sy = x, y
				hasPen = true
			}
		}
	}
}
