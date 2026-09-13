package pptxpatch

import (
	"encoding/xml"
	"fmt"
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
	// Nonempty interactive handles/connections are not yet qualified. Retain the
	// source as preserve-only instead of interpreting or dropping unknown clauses.
	for _, name := range []string{"ahLst", "cxnLst"} {
		if list := nativeChild(node, ns, name); list != nil {
			if err := requireEmptyNativeElement(list); err != nil {
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
		if fill != "norm" && fill != "none" {
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
		}
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
			if err := p.arc(values[0], values[1], values[2], values[3]); err != nil {
				return nil, err
			}
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
				points = append(points, nativeGeometryPoint{x, y})
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
