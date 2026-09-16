package pptxpatch

import (
	"encoding/xml"
	"fmt"
	"math"
	"strings"
)

// evaluateNativeCustomGeometry is the single source XML geometry evaluator.
// Unsupported clauses refuse the entire geometry, never a partial path list.
func evaluateNativeCustomGeometry(node *nativeXMLNode, ns string, width, height int64) (*NativeEvaluatedGeometry, error) {
	return evaluateNativeGeometryWithIntermediateLimit(node, ns, width, height, nativeGeometryMaxMagnitude)
}

// The wider limit is an internal catalog-only policy, never selected by XML.
func evaluateNativeGeometryWithIntermediateLimit(node *nativeXMLNode, ns string, width, height int64, limit float64) (*NativeEvaluatedGeometry, error) {
	if node == nil || node.Name != (xml.Name{Space: ns, Local: "custGeom"}) {
		return nil, fmt.Errorf("missing custom geometry")
	}
	if err := requireOnlyNativeAttrs(node); err != nil {
		return nil, err
	}
	allowed := []xml.Name{}
	for _, name := range []string{"avLst", "gdLst", "ahLst", "cxnLst", "rect", "pathLst"} {
		allowed = append(allowed, xml.Name{Space: ns, Local: name})
	}
	if err := requireOnlyNativeChildren(node, allowed...); err != nil {
		return nil, err
	}
	for _, name := range allowed {
		if len(nativeChildren(node, ns, name.Local)) > 1 {
			return nil, fmt.Errorf("duplicate geometry %s", name.Local)
		}
	}
	order := map[string]int{"avLst": 0, "gdLst": 1, "ahLst": 2, "cxnLst": 3, "rect": 4, "pathLst": 5}
	previous := -1
	for _, child := range node.Children {
		current := order[child.Name.Local]
		if current < previous {
			return nil, fmt.Errorf("invalid custom geometry child order")
		}
		previous = current
	}
	g, err := newNativeGeometryGuides(float64(width), float64(height))
	if err != nil {
		return nil, err
	}
	count := 0
	for _, name := range []string{"avLst", "gdLst"} {
		list := nativeChild(node, ns, name)
		if list == nil {
			continue
		}
		if err := requireOnlyNativeAttrs(list); err != nil {
			return nil, err
		}
		if err := requireOnlyNativeChildren(list, xml.Name{Space: ns, Local: "gd"}); err != nil {
			return nil, err
		}
		guides := []nativeGeometryGuide{}
		for _, gd := range list.Children {
			if err := requireOnlyNativeAttrs(gd, xml.Name{Local: "name"}, xml.Name{Local: "fmla"}); err != nil {
				return nil, err
			}
			if err := requireOnlyNativeChildren(gd); err != nil {
				return nil, err
			}
			key, _ := exactNativeAttr(gd, "", "name")
			formula, _ := exactNativeAttr(gd, "", "fmla")
			if name == "avLst" {
				fields := strings.Fields(formula)
				if len(fields) != 2 || fields[0] != "val" {
					return nil, fmt.Errorf("adjustment requires val formula")
				}
			}
			guides = append(guides, nativeGeometryGuide{key, formula})
		}
		count += len(guides)
		if count > nativeGeometryMaxGuides {
			return nil, fmt.Errorf("geometry guide budget exceeded")
		}
		if err := g.evaluateWithIntermediateLimit(guides, limit); err != nil {
			return nil, err
		}
	}
	// a:ahLst (ECMA-376 Part 1 §20.1.9.1) and a:cxnLst (§20.1.9.11) are authoring
	// affordances: adjust handles say where the interactive drag points sit and
	// which existing guides they write, and connection sites say where a
	// connector may attach. Neither contributes a segment to a:pathLst and
	// neither is consulted when the shape is painted, so a populated list paints
	// exactly what an empty one paints. Their contents are still structurally
	// qualified so unknown markup inside them keeps refusing.
	for _, name := range []string{"ahLst", "cxnLst"} {
		if list := nativeChild(node, ns, name); list != nil {
			if err := qualifyNativeGeometryUnpaintedList(list, ns, name); err != nil {
				return nil, err
			}
		}
	}
	result := &NativeEvaluatedGeometry{Profile: "drawingml-paths-v1", TextRect: NativeGeometryTextRect{CX: width, CY: height}, Paths: []NativeGeometryPath{}}
	if rect := nativeChild(node, ns, "rect"); rect != nil {
		if err := requireOnlyNativeAttrs(rect, xml.Name{Local: "l"}, xml.Name{Local: "t"}, xml.Name{Local: "r"}, xml.Name{Local: "b"}); err != nil {
			return nil, err
		}
		if err := requireOnlyNativeChildren(rect); err != nil {
			return nil, err
		}
		edges := []int64{}
		for _, key := range []string{"l", "t", "r", "b"} {
			v, err := nativeGeometryAttribute(g, rect, key)
			if err != nil {
				return nil, err
			}
			if _, err := g.qualifyOutputAttribute(rect, key, 1); err != nil {
				return nil, err
			}
			edge, err := nativeGeometryRound(v)
			if err != nil {
				return nil, err
			}
			edges = append(edges, *edge)
		}
		if edges[2] <= edges[0] || edges[3] <= edges[1] || !nativeGeometryFinite(float64(edges[2])-float64(edges[0])) || !nativeGeometryFinite(float64(edges[3])-float64(edges[1])) {
			return nil, fmt.Errorf("invalid geometry text rectangle")
		}
		result.TextRect = NativeGeometryTextRect{X: edges[0], Y: edges[1], CX: edges[2] - edges[0], CY: edges[3] - edges[1]}
	}
	list := nativeChild(node, ns, "pathLst")
	if err := requireOnlyNativeAttrs(list); err != nil {
		return nil, err
	}
	if err := requireOnlyNativeChildren(list, xml.Name{Space: ns, Local: "path"}); err != nil {
		return nil, err
	}
	if len(list.Children) == 0 || len(list.Children) > nativeGeometryMaxPaths {
		return nil, fmt.Errorf("geometry path budget exceeded")
	}
	commands := 0
	for _, path := range list.Children {
		evaluated, err := evaluateNativeGeometryPath(path, ns, g, float64(width), float64(height))
		if err != nil {
			return nil, err
		}
		commands += len(evaluated.Commands)
		if commands > nativeGeometryMaxCommands {
			return nil, fmt.Errorf("geometry total command budget exceeded")
		}
		result.Paths = append(result.Paths, *evaluated)
	}
	return result, nil
}

func nativeGeometryAttribute(g nativeGeometryGuides, node *nativeXMLNode, key string) (float64, error) {
	token, ok := exactNativeAttr(node, "", key)
	if !ok {
		return 0, fmt.Errorf("missing geometry %s", key)
	}
	return g.resolve(token)
}

func evaluateNativeGeometryPath(node *nativeXMLNode, ns string, g nativeGeometryGuides, width, height float64) (*NativeGeometryPath, error) {
	if err := requireOnlyNativeAttrs(node, xml.Name{Local: "w"}, xml.Name{Local: "h"}, xml.Name{Local: "fill"}, xml.Name{Local: "stroke"}, xml.Name{Local: "extrusionOk"}); err != nil {
		return nil, err
	}
	names := []xml.Name{}
	for _, name := range []string{"moveTo", "lnTo", "quadBezTo", "cubicBezTo", "arcTo", "close"} {
		names = append(names, xml.Name{Space: ns, Local: name})
	}
	if err := requireOnlyNativeChildren(node, names...); err != nil {
		return nil, err
	}
	if len(node.Children) == 0 || len(node.Children) > nativeGeometryMaxPathCommands {
		return nil, fmt.Errorf("geometry path command budget exceeded")
	}
	result := &NativeGeometryPath{FillMode: "norm", Stroke: true}
	if fill, ok := exactNativeAttr(node, "", "fill"); ok {
		if !nativeGeometryFillMode(fill) {
			return nil, fmt.Errorf("unsupported geometry path fill")
		}
		result.FillMode = fill
	}
	for _, name := range []string{"stroke", "extrusionOk"} {
		if value, ok := exactNativeAttr(node, "", name); ok {
			if value != "true" && value != "false" && value != "0" && value != "1" {
				return nil, fmt.Errorf("invalid geometry path boolean")
			}
			on := value == "true" || value == "1"
			if name == "stroke" {
				result.Stroke = on
			}
		}
	}
	p := nativeGeometryPathBuilder{sx: 1, sy: 1}
	exactUnitScale := true
	for _, axis := range []struct {
		key    string
		extent float64
		scale  *float64
	}{{"w", width, &p.sx}, {"h", height, &p.sy}} {
		if value, ok := exactNativeAttr(node, "", axis.key); ok {
			empty := nativeGeometryGuides{}
			v, err := empty.resolve(value)
			if err != nil || v <= 0 {
				return nil, fmt.Errorf("invalid geometry path coordinate extent")
			}
			*axis.scale = axis.extent / v
			exactUnitScale = exactUnitScale && v == axis.extent && nativeGeometryExactIntegerAttribute(empty, node, axis.key, v)
		}
	}
	exactPen, exactStart := false, false
	pathUncertainty := 0.0
	addUncertainty := func(value float64) error {
		pathUncertainty += value
		if math.IsNaN(pathUncertainty) || math.IsInf(pathUncertainty, 0) || pathUncertainty > nativeGeometryMaxOutputUncertainty {
			return fmt.Errorf("geometry accumulated path uncertainty exceeds one eighth EMU")
		}
		return nil
	}
	for _, command := range node.Children {
		kind := command.Name.Local
		switch kind {
		case "close":
			if err := requireEmptyNativeElement(command); err != nil {
				return nil, err
			}
			if err := p.close(); err != nil {
				return nil, err
			}
			exactPen = exactStart
		case "arcTo":
			if err := requireOnlyNativeAttrs(command, xml.Name{Local: "wR"}, xml.Name{Local: "hR"}, xml.Name{Local: "stAng"}, xml.Name{Local: "swAng"}); err != nil {
				return nil, err
			}
			if err := requireOnlyNativeChildren(command); err != nil {
				return nil, err
			}
			values := []float64{}
			for _, key := range []string{"wR", "hR", "stAng", "swAng"} {
				v, err := nativeGeometryAttribute(g, command, key)
				if err != nil {
					return nil, err
				}
				values = append(values, v)
			}
			exactArc := nativeGeometryExactCardinalArc(g, command, p, exactPen && exactUnitScale, values)
			// A polar ellipse can magnify angular uncertainty by rMax²/rMin.
			// Bound radius sensitivity conservatively by the cubed aspect ratio.
			maxR, minR := math.Max(values[0], values[1]), math.Min(values[0], values[1])
			if minR > 0 {
				aspect := maxR / minR
				scale := math.Max(math.Abs(p.sx), math.Abs(p.sy))
				uncertainty := 0.0
				for _, key := range []string{"wR", "hR"} {
					u, e := g.qualifyOutputAttribute(command, key, 2*scale*aspect*aspect*aspect)
					if e != nil {
						return nil, e
					}
					uncertainty += u
				}
				for _, key := range []string{"stAng", "swAng"} {
					u, e := g.qualifyOutputAttribute(command, key, 2*scale*maxR*aspect*nativeGeometryAngleUnit)
					if e != nil {
						return nil, e
					}
					uncertainty += u
				}
				// Every arc endpoint uses the prior pen and two polar offsets.
				// Conservatively sum the whole path, including all previous arcs.
				// The construction allowance includes up to four split segments,
				// libm conversion, ellipse conditioning and floating additions.
				magnitude := math.Max(maxR, math.Max(math.Abs(p.pen.x), math.Abs(p.pen.y)))
				if !exactArc {
					uncertainty += 1024 * 2.220446049250313e-16 * scale * magnitude * aspect * aspect * aspect
				}
				if err := addUncertainty(uncertainty); err != nil {
					return nil, err
				}
			}
			if err := p.arc(values[0], values[1], values[2], values[3]); err != nil {
				return nil, err
			}
			exactPen = exactArc
		default:
			if err := requireOnlyNativeAttrs(command); err != nil {
				return nil, err
			}
			if err := requireOnlyNativeChildren(command, xml.Name{Space: ns, Local: "pt"}); err != nil {
				return nil, err
			}
			required := map[string]int{"moveTo": 1, "lnTo": 1, "quadBezTo": 2, "cubicBezTo": 3}[kind]
			if len(command.Children) != required {
				return nil, fmt.Errorf("incorrect geometry point count")
			}
			points := []nativeGeometryPoint{}
			for _, point := range command.Children {
				if err := requireOnlyNativeAttrs(point, xml.Name{Local: "x"}, xml.Name{Local: "y"}); err != nil {
					return nil, err
				}
				if err := requireOnlyNativeChildren(point); err != nil {
					return nil, err
				}
				x, err := nativeGeometryAttribute(g, point, "x")
				if err != nil {
					return nil, err
				}
				y, err := nativeGeometryAttribute(g, point, "y")
				if err != nil {
					return nil, err
				}
				for _, axis := range []struct {
					key          string
					scale, value float64
				}{{"x", p.sx, x}, {"y", p.sy, y}} {
					u, err := g.qualifyOutputAttribute(point, axis.key, axis.scale)
					if err != nil {
						return nil, err
					}
					// Non-unit path-space division/multiplication is not exact.
					if axis.scale != 1 {
						u += 16 * 2.220446049250313e-16 * math.Abs(axis.value*axis.scale)
					}
					if err := addUncertainty(u); err != nil {
						return nil, err
					}
				}
				exactPen = exactUnitScale && p.sx == 1 && p.sy == 1 && nativeGeometryExactIntegerAttribute(g, point, "x", x) && nativeGeometryExactIntegerAttribute(g, point, "y", y)
				points = append(points, nativeGeometryPoint{x, y})
			}
			if kind == "moveTo" {
				exactStart = exactPen
			}
			mapped := map[string]string{"moveTo": "moveTo", "lnTo": "lineTo", "quadBezTo": "quadBezierTo", "cubicBezTo": "cubicBezierTo"}[kind]
			if err := p.vertices(mapped, points...); err != nil {
				return nil, err
			}
		}
	}
	result.Commands = p.commands
	return result, nil
}

// qualifyNativeGeometryUnpaintedList validates a custom-geometry clause that the
// evaluator deliberately ignores because it paints nothing. Only the DrawingML
// members of that clause and their documented attributes are tolerated; anything
// else refuses the geometry, exactly as an unknown child of a:custGeom does.
func qualifyNativeGeometryUnpaintedList(list *nativeXMLNode, ns, name string) error {
	if err := requireOnlyNativeAttrs(list); err != nil {
		return err
	}
	members := map[string][]string{
		"ahXY":    {"gdRefX", "minX", "maxX", "gdRefY", "minY", "maxY"},
		"ahPolar": {"gdRefAng", "minAng", "maxAng", "gdRefR", "minR", "maxR"},
		"cxn":     {"ang"},
	}
	allowed := []xml.Name{}
	for _, member := range map[string][]string{"ahLst": {"ahXY", "ahPolar"}, "cxnLst": {"cxn"}}[name] {
		allowed = append(allowed, xml.Name{Space: ns, Local: member})
	}
	if err := requireOnlyNativeChildren(list, allowed...); err != nil {
		return err
	}
	if len(list.Children) > nativeGeometryMaxGuides {
		return fmt.Errorf("geometry %s budget exceeded", name)
	}
	for _, child := range list.Children {
		attrs := []xml.Name{}
		for _, local := range members[child.Name.Local] {
			attrs = append(attrs, xml.Name{Local: local})
		}
		if err := requireOnlyNativeAttrs(child, attrs...); err != nil {
			return err
		}
		if err := requireOnlyNativeChildren(child, xml.Name{Space: ns, Local: "pos"}); err != nil {
			return err
		}
		position := nativeChild(child, ns, "pos")
		if position == nil || len(nativeChildren(child, ns, "pos")) != 1 {
			return fmt.Errorf("geometry %s entry requires one position", child.Name.Local)
		}
		if err := requireOnlyNativeAttrs(position, xml.Name{Local: "x"}, xml.Name{Local: "y"}); err != nil {
			return err
		}
		if err := requireOnlyNativeChildren(position); err != nil {
			return err
		}
	}
	return nil
}
