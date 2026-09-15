package pptxpatch

import (
	"encoding/xml"
	"math"
	"sort"
	"strconv"
	"strings"
)

// DrawingML diagram layout evaluation (ECMA-376 Part 1 §21.4). This file
// holds the two inputs of the layout engine:
//
//   - the data model (§21.4.3): dgm:pt points (doc/node/asst/parTrans/sibTrans/
//     pres) connected by dgm:cxn parOf relationships into one tree rooted at
//     the doc point;
//   - the layout definition (§21.4.2): a layoutNode tree whose forEach/choose/if
//     elements are evaluated over the data model axes (§21.4.7.6) to produce
//     the presentation tree, and whose constrLst entries (§21.4.2.8) are then
//     evaluated in document order.
//
// Everything is bounded: point count, layout depth, presentation node count,
// selection operations, and constraint evaluations. Anything outside the
// implemented subset refuses with a specific code; nothing is guessed.
const (
	nativeMaxDiagramLayoutPoints      = 256
	nativeMaxDiagramLayoutDepth       = 32
	nativeMaxDiagramLayoutPresNodes   = 2048
	nativeMaxDiagramLayoutSelections  = 65536
	nativeMaxDiagramLayoutConstraints = 4096

	nativeDiagramLayoutDataCode       = "pptx.diagram-layout-data-unavailable"
	nativeDiagramLayoutDefinitionCode = "pptx.diagram-layout-definition-unavailable"
	nativeDiagramLayoutBudgetCode     = "pptx.diagram-layout-budget-unavailable"
	nativeDiagramLayoutConstraintCode = "pptx.diagram-layout-constraint-unavailable"
)

type nativeDiagramPoint struct {
	id          string
	kind        string
	text        *nativeXMLNode
	placeholder bool
	parent      *nativeDiagramPoint
	children    []*nativeDiagramPoint
	parTrans    *nativeDiagramPoint
	sibTrans    *nativeDiagramPoint
	owner       *nativeDiagramPoint
	depth       int
	order       int
}

type nativeDiagramModel struct {
	doc      *nativeDiagramPoint
	points   map[string]*nativeDiagramPoint
	presVars map[string]map[string]string
	// presLabels holds the authored presStyleLbl of each presentation point,
	// keyed by presAssocID and presName, for layout nodes without styleLbl.
	presLabels map[string]string
}

// nativeDiagramLayoutVariableDefaults are the values a bare variable element
// (no val attribute) carries, per the CT_LayoutVariablePropertySet defaults.
var nativeDiagramLayoutVariableDefaults = map[string]string{"dir": "norm", "hierBranch": "std", "orgChart": "0", "bulletEnabled": "0", "chPref": "-1", "chMax": "-1", "animLvl": "none", "animOne": "none", "resizeHandles": "rel"}

func nativeDiagramLayoutRefuse(code, message string) error {
	return refuseNativeDiagram(code, message)
}

// parseNativeDiagramModel builds the data tree from dgm:dataModel (§21.4.3.1).
func parseNativeDiagramModel(root *nativeXMLNode, diagramNS string, dialect nativeExtractDialect) (*nativeDiagramModel, error) {
	refuse := func(message string) (*nativeDiagramModel, error) {
		return nil, nativeDiagramLayoutRefuse(nativeDiagramLayoutDataCode, message)
	}
	dgm := func(local string) xml.Name { return xml.Name{Space: diagramNS, Local: local} }
	if requireOnlyNativeChildren(root, dgm("ptLst"), dgm("cxnLst"), dgm("bg"), dgm("whole"), dgm("extLst")) != nil {
		return refuse("diagram data model contains unknown markup")
	}
	ptLst, err := nativeSingleton(root, diagramNS, "ptLst", true)
	if err != nil {
		return refuse("diagram data model requires one point list")
	}
	cxnLst, err := nativeSingleton(root, diagramNS, "cxnLst", false)
	if err != nil {
		return refuse("diagram data model has duplicate connection lists")
	}
	model := &nativeDiagramModel{points: map[string]*nativeDiagramPoint{}, presVars: map[string]map[string]string{}, presLabels: map[string]string{}}
	points := nativeChildren(ptLst, diagramNS, "pt")
	if len(points) > nativeMaxDiagramLayoutPoints {
		return nil, nativeDiagramLayoutRefuse(nativeDiagramLayoutBudgetCode, "diagram data model exceeds the bounded point budget")
	}
	for index, node := range points {
		if requireOnlyNativeAttrs(node, xml.Name{Local: "modelId"}, xml.Name{Local: "type"}, xml.Name{Local: "cxnId"}) != nil ||
			requireOnlyNativeChildren(node, dgm("prSet"), dgm("spPr"), dgm("t"), dgm("extLst")) != nil {
			return refuse("diagram data point contains unknown markup")
		}
		id, ok := exactNativeAttr(node, "", "modelId")
		if !ok || id == "" || model.points[id] != nil {
			return refuse("diagram data point requires a unique model id")
		}
		point := &nativeDiagramPoint{id: id, kind: "node", order: index}
		if kind, ok := exactNativeAttr(node, "", "type"); ok {
			point.kind = kind
		}
		switch point.kind {
		case "doc", "node", "asst", "parTrans", "sibTrans", "pres":
		default:
			return refuse("diagram data point type is unknown")
		}
		if spPr := nativeChild(node, diagramNS, "spPr"); spPr != nil && (len(spPr.Children) != 0 || len(spPr.Attrs) != 0) {
			return refuse("diagram data point carries authored shape property overrides; layout approximation does not apply them")
		}
		if prSet := nativeChild(node, diagramNS, "prSet"); prSet != nil {
			if err := model.parseNativeDiagramPropertySet(prSet, point, diagramNS); err != nil {
				return nil, err
			}
		}
		if text := nativeChild(node, diagramNS, "t"); text != nil {
			point.text = text
		}
		if point.kind == "doc" {
			if model.doc != nil {
				return refuse("diagram data model has more than one document point")
			}
			model.doc = point
		}
		model.points[id] = point
	}
	if model.doc == nil {
		return refuse("diagram data model has no document point")
	}
	type childLink struct {
		point *nativeDiagramPoint
		order int64
		par   *nativeDiagramPoint
		sib   *nativeDiagramPoint
	}
	links := map[*nativeDiagramPoint][]childLink{}
	if cxnLst != nil {
		for _, node := range nativeChildren(cxnLst, diagramNS, "cxn") {
			if requireOnlyNativeAttrs(node, xml.Name{Local: "modelId"}, xml.Name{Local: "type"}, xml.Name{Local: "srcId"}, xml.Name{Local: "destId"},
				xml.Name{Local: "srcOrd"}, xml.Name{Local: "destOrd"}, xml.Name{Local: "parTransId"}, xml.Name{Local: "sibTransId"}, xml.Name{Local: "presId"}) != nil ||
				requireOnlyNativeChildren(node, dgm("extLst")) != nil {
				return refuse("diagram connection contains unknown markup")
			}
			kind, _ := exactNativeAttr(node, "", "type")
			switch kind {
			case "", "parOf":
			case "presOf", "presParOf":
				// PowerPoint's cached presentation graph is not trusted; the
				// presentation tree is re-evaluated from the layout definition.
				continue
			default:
				return refuse("diagram connection type is unknown")
			}
			srcID, _ := exactNativeAttr(node, "", "srcId")
			destID, _ := exactNativeAttr(node, "", "destId")
			src, dest := model.points[srcID], model.points[destID]
			if src == nil || dest == nil || (dest.kind != "node" && dest.kind != "asst") || (src.kind != "doc" && src.kind != "node" && src.kind != "asst") {
				return refuse("diagram parent connection references a missing or non-node point")
			}
			if dest.parent != nil || dest == model.doc {
				return refuse("diagram parent connections form a cycle or a shared child")
			}
			dest.parent = src
			orderValue, _ := exactNativeAttr(node, "", "srcOrd")
			order, err := strconv.ParseInt(orderValue, 10, 32)
			if orderValue != "" && err != nil {
				return refuse("diagram connection order is non-canonical")
			}
			link := childLink{point: dest, order: order}
			if id, ok := exactNativeAttr(node, "", "parTransId"); ok && id != "" {
				if link.par = model.points[id]; link.par == nil || link.par.kind != "parTrans" || link.par.owner != nil {
					return refuse("diagram parent transition reference is invalid")
				}
				link.par.owner = dest
			}
			if id, ok := exactNativeAttr(node, "", "sibTransId"); ok && id != "" {
				if link.sib = model.points[id]; link.sib == nil || link.sib.kind != "sibTrans" || link.sib.owner != nil {
					return refuse("diagram sibling transition reference is invalid")
				}
				link.sib.owner = dest
			}
			links[src] = append(links[src], link)
		}
	}
	for parent, list := range links {
		sort.SliceStable(list, func(a, b int) bool { return list[a].order < list[b].order })
		for _, link := range list {
			link.point.parTrans, link.point.sibTrans = link.par, link.sib
			parent.children = append(parent.children, link.point)
		}
	}
	// Every node must hang below the document point (no orphans, no cycles).
	reached := 0
	stack := []*nativeDiagramPoint{model.doc}
	for len(stack) != 0 {
		point := stack[len(stack)-1]
		stack = stack[:len(stack)-1]
		reached++
		if reached > len(model.points) {
			return refuse("diagram parent connections form a cycle")
		}
		for _, child := range point.children {
			child.depth = point.depth + 1
			if child.parTrans != nil {
				child.parTrans.depth = child.depth
			}
			if child.sibTrans != nil {
				child.sibTrans.depth = child.depth
			}
			stack = append(stack, child)
		}
	}
	for _, point := range model.points {
		if (point.kind == "node" || point.kind == "asst") && point.parent == nil {
			return refuse("diagram node point is not connected to the document point")
		}
	}
	return model, nil
}

// parseNativeDiagramPropertySet reads dgm:prSet (§21.4.3.4). Presentation
// layout variables on pres points override the layout definition's varLst for
// the (presAssocID, presName) key. Geometry customizations are refused because
// the layout would silently differ from the authored result.
func (model *nativeDiagramModel) parseNativeDiagramPropertySet(prSet *nativeXMLNode, point *nativeDiagramPoint, diagramNS string) error {
	if requireOnlyNativeChildren(prSet, xml.Name{Space: diagramNS, Local: "presLayoutVars"}, xml.Name{Space: diagramNS, Local: "style"}, xml.Name{Space: diagramNS, Local: "extLst"}) != nil {
		return nativeDiagramLayoutRefuse(nativeDiagramLayoutDataCode, "diagram property set contains unknown markup")
	}
	if style := nativeChild(prSet, diagramNS, "style"); style != nil && len(style.Children) != 0 {
		return nativeDiagramLayoutRefuse(nativeDiagramLayoutDataCode, "diagram point carries an authored style override; layout approximation does not apply it")
	}
	presName := ""
	presAssoc := ""
	presLabel := ""
	for _, attr := range prSet.Attrs {
		if attr.Name.Space != "" {
			continue
		}
		switch attr.Name.Local {
		case "phldr":
			point.placeholder = attr.Value == "1" || attr.Value == "true"
		case "presName":
			presName = attr.Value
		case "presAssocID":
			presAssoc = attr.Value
		case "custAng", "custFlipVert", "custFlipHor", "custSzX", "custSzY", "custScaleX", "custScaleY", "custLinFactX", "custLinFactY", "custLinFactNeighborX", "custLinFactNeighborY", "custRadScaleRad", "custRadScaleInc":
			if attr.Value != "0" && attr.Value != "false" && attr.Value != "100000" {
				return nativeDiagramLayoutRefuse(nativeDiagramLayoutDataCode, "diagram point carries a geometry customization; layout approximation does not apply it")
			}
		case "presStyleLbl":
			presLabel = attr.Value
		case "custT", "phldrT", "loTypeId", "loCatId", "qsTypeId", "qsCatId", "csTypeId", "csCatId", "coherent3DOff", "presStyleIdx", "presStyleCnt":
		default:
			return nativeDiagramLayoutRefuse(nativeDiagramLayoutDataCode, "diagram property set attribute is unknown")
		}
	}
	if point.kind != "pres" || presName == "" || presAssoc == "" {
		return nil
	}
	if presLabel != "" {
		model.presLabels[presAssoc+"\x00"+presName] = presLabel
	}
	vars := nativeChild(prSet, diagramNS, "presLayoutVars")
	if vars == nil {
		return nil
	}
	values := map[string]string{}
	for _, variable := range vars.Children {
		if variable.Name.Space != diagramNS || len(variable.Children) != 0 || requireOnlyNativeAttrs(variable, xml.Name{Local: "val"}) != nil {
			return nativeDiagramLayoutRefuse(nativeDiagramLayoutDataCode, "diagram presentation layout variable is outside the modeled set")
		}
		if _, known := nativeDiagramLayoutVariableDefaults[variable.Name.Local]; !known {
			return nativeDiagramLayoutRefuse(nativeDiagramLayoutDataCode, "diagram presentation layout variable "+variable.Name.Local+" is not modeled")
		}
		value, ok := exactNativeAttr(variable, "", "val")
		if !ok {
			value = nativeDiagramLayoutVariableDefaults[variable.Name.Local]
		}
		values[variable.Name.Local] = value
	}
	model.presVars[presAssoc+"\x00"+presName] = values
	return nil
}

// nativeDiagramPointTypeMatches implements ST_ElementType (§21.4.7.27). A
// "node" is any data node including assistants (nonAsst excludes them).
func nativeDiagramPointTypeMatches(ptType string, point *nativeDiagramPoint) bool {
	switch ptType {
	case "", "all":
		return true
	case "node":
		return point.kind == "node" || point.kind == "asst"
	case "asst", "doc", "parTrans", "sibTrans", "pres":
		return point.kind == ptType
	case "nonAsst":
		return point.kind == "node" || point.kind == "doc"
	case "norm":
		return point.kind == "node" || point.kind == "asst" || point.kind == "doc"
	case "nonNorm":
		return point.kind == "parTrans" || point.kind == "sibTrans"
	}
	return false
}

// siblingSequence lists the parent's children with their transitions in the
// order [parTrans, node, sibTrans] so precedSib/followSib with ptType parTrans
// reach the connection points of the context node (§21.4.7.6).
func nativeDiagramSiblingSequence(parent *nativeDiagramPoint) []*nativeDiagramPoint {
	sequence := make([]*nativeDiagramPoint, 0, len(parent.children)*3)
	for _, child := range parent.children {
		if child.parTrans != nil {
			sequence = append(sequence, child.parTrans)
		}
		sequence = append(sequence, child)
		if child.sibTrans != nil {
			sequence = append(sequence, child.sibTrans)
		}
	}
	return sequence
}

func nativeDiagramDescendants(point *nativeDiagramPoint, out []*nativeDiagramPoint) []*nativeDiagramPoint {
	for _, child := range point.children {
		out = append(out, child)
		out = nativeDiagramDescendants(child, out)
	}
	return out
}

// nativeDiagramLayoutEvaluator expands the layout definition over the data.
type nativeDiagramLayoutEvaluator struct {
	model         *nativeDiagramModel
	ns            string
	forEachByName map[string]*nativeXMLNode
	presNodes     int
	selections    int
	constraints   int
}

type nativeDiagramConstraint struct {
	typ, forRel, forName, ptType, refType, refFor, refForName, refPtType, op string
	val                                                                      float64
	hasVal                                                                   bool
	fact                                                                     float64
}

type nativeDiagramRule struct {
	typ string
	val float64
}

// nativeDiagramPresNode is one evaluated layout node instance.
type nativeDiagramPresNode struct {
	name        string
	point       *nativeDiagramPoint
	parent      *nativeDiagramPresNode
	children    []*nativeDiagramPresNode
	vars        map[string]string
	alg         string
	params      map[string]string
	hasShape    bool
	shapeType   string
	hideGeom    bool
	zOrderOff   int64
	styleLbl    string
	presOf      []*nativeDiagramPoint
	constraints []nativeDiagramConstraint
	rules       []nativeDiagramRule
	vals        map[string]float64
	equalize    [][]*nativeDiagramPresNode

	// Layout results (unscaled units; see native_diagram_layout_hier.go).
	rect      nativeDiagramRect
	blockW    float64
	blockH    float64
	anchorX   float64
	rootLeft  float64
	rootRight float64
	laidOut   bool
	fontSize  int64
}

func (node *nativeDiagramPresNode) value(name string, fallback float64) float64 {
	if value, ok := node.vals[name]; ok {
		return value
	}
	return fallback
}

func (node *nativeDiagramPresNode) param(name, fallback string) string {
	if value, ok := node.params[name]; ok && value != "" {
		return value
	}
	return fallback
}

// lookupVariable resolves func="var" arguments: pres point overrides for the
// (point, layoutNode name) key win over the nearest ancestor varLst.
func (node *nativeDiagramPresNode) lookupVariable(model *nativeDiagramModel, name string) string {
	for current := node; current != nil; current = current.parent {
		if current.point != nil {
			if overrides, ok := model.presVars[current.point.id+"\x00"+current.name]; ok {
				if value, ok := overrides[name]; ok {
					return value
				}
			}
		}
		if value, ok := current.vars[name]; ok {
			return value
		}
	}
	return nativeDiagramLayoutVariableDefaults[name]
}

func newNativeDiagramLayoutEvaluator(model *nativeDiagramModel, layoutRoot *nativeXMLNode, diagramNS string) (*nativeDiagramLayoutEvaluator, *nativeXMLNode, error) {
	dgm := func(local string) xml.Name { return xml.Name{Space: diagramNS, Local: local} }
	if requireOnlyNativeChildren(layoutRoot, dgm("title"), dgm("desc"), dgm("catLst"), dgm("sampData"), dgm("styleData"), dgm("clrData"), dgm("layoutNode"), dgm("extLst")) != nil {
		return nil, nil, nativeDiagramLayoutRefuse(nativeDiagramLayoutDefinitionCode, "diagram layout definition contains unknown markup")
	}
	root, err := nativeSingleton(layoutRoot, diagramNS, "layoutNode", true)
	if err != nil {
		return nil, nil, nativeDiagramLayoutRefuse(nativeDiagramLayoutDefinitionCode, "diagram layout definition requires one root layout node")
	}
	evaluator := &nativeDiagramLayoutEvaluator{model: model, ns: diagramNS, forEachByName: map[string]*nativeXMLNode{}}
	if err := evaluator.indexForEach(root, 0); err != nil {
		return nil, nil, err
	}
	return evaluator, root, nil
}

func (evaluator *nativeDiagramLayoutEvaluator) indexForEach(node *nativeXMLNode, depth int) error {
	if depth > nativeMaxDiagramLayoutDepth*4 {
		return nativeDiagramLayoutRefuse(nativeDiagramLayoutBudgetCode, "diagram layout definition nesting exceeds the bounded depth")
	}
	for _, child := range node.Children {
		if child.Name == (xml.Name{Space: evaluator.ns, Local: "forEach"}) {
			if name, ok := exactNativeAttr(child, "", "name"); ok && name != "" {
				if _, seen := evaluator.forEachByName[name]; seen {
					return nativeDiagramLayoutRefuse(nativeDiagramLayoutDefinitionCode, "diagram layout definition declares a duplicate forEach name")
				}
				evaluator.forEachByName[name] = child
			}
		}
		if err := evaluator.indexForEach(child, depth+1); err != nil {
			return err
		}
	}
	return nil
}

func (evaluator *nativeDiagramLayoutEvaluator) selectionBudget(cost int) error {
	evaluator.selections += cost
	if evaluator.selections > nativeMaxDiagramLayoutSelections {
		return nativeDiagramLayoutRefuse(nativeDiagramLayoutBudgetCode, "diagram layout evaluation exceeds the bounded selection budget")
	}
	return nil
}

// selectPoints implements axis/ptType/st/cnt/step selection (§21.4.2.14,
// §21.4.7.6). Axis and ptType lists pair up position-wise; missing ptTypes
// select all.
func (evaluator *nativeDiagramLayoutEvaluator) selectPoints(context *nativeDiagramPoint, node *nativeXMLNode) ([]*nativeDiagramPoint, error) {
	axisAttr, _ := exactNativeAttr(node, "", "axis")
	ptTypeAttr, _ := exactNativeAttr(node, "", "ptType")
	axes := strings.Fields(axisAttr)
	if len(axes) == 0 {
		axes = []string{"self"}
	}
	ptTypes := strings.Fields(ptTypeAttr)
	current := []*nativeDiagramPoint{context}
	for index, axis := range axes {
		ptType := ""
		if index < len(ptTypes) {
			ptType = ptTypes[index]
		} else if len(ptTypes) == 1 && len(axes) == 1 {
			ptType = ptTypes[0]
		}
		next := []*nativeDiagramPoint{}
		for _, point := range current {
			items, err := evaluator.axisItems(point, axis)
			if err != nil {
				return nil, err
			}
			for _, item := range items {
				if nativeDiagramPointTypeMatches(ptType, item) {
					next = append(next, item)
				}
			}
		}
		if err := evaluator.selectionBudget(len(next) + 1); err != nil {
			return nil, err
		}
		current = next
	}
	start, err := nativeDiagramIntAttr(node, "st", 1)
	if err != nil {
		return nil, err
	}
	count, err := nativeDiagramIntAttr(node, "cnt", 0)
	if err != nil {
		return nil, err
	}
	step, err := nativeDiagramIntAttr(node, "step", 1)
	if err != nil {
		return nil, err
	}
	if step <= 0 {
		return nil, nativeDiagramLayoutRefuse(nativeDiagramLayoutDefinitionCode, "diagram selection step must be positive")
	}
	total := int64(len(current))
	if start < 0 {
		start = total + start + 1
	}
	if start < 1 {
		start = 1
	}
	selected := []*nativeDiagramPoint{}
	for index := start - 1; index < total; index += step {
		if count > 0 && int64(len(selected)) >= count {
			break
		}
		selected = append(selected, current[index])
	}
	return selected, nil
}

func nativeDiagramIntAttr(node *nativeXMLNode, local string, fallback int64) (int64, error) {
	value, ok := exactNativeAttr(node, "", local)
	if !ok || value == "" {
		return fallback, nil
	}
	parsed, err := strconv.ParseInt(value, 10, 32)
	if err != nil {
		return 0, nativeDiagramLayoutRefuse(nativeDiagramLayoutDefinitionCode, "diagram layout attribute "+local+" is non-canonical")
	}
	return parsed, nil
}

func (evaluator *nativeDiagramLayoutEvaluator) axisItems(point *nativeDiagramPoint, axis string) ([]*nativeDiagramPoint, error) {
	parent := point.parent
	if point.owner != nil {
		parent = point.owner.parent
	}
	switch axis {
	case "self":
		return []*nativeDiagramPoint{point}, nil
	case "ch":
		return point.children, nil
	case "des":
		return nativeDiagramDescendants(point, nil), nil
	case "desOrSelf":
		return nativeDiagramDescendants(point, []*nativeDiagramPoint{point}), nil
	case "par":
		if parent == nil {
			return nil, nil
		}
		return []*nativeDiagramPoint{parent}, nil
	case "root":
		return []*nativeDiagramPoint{evaluator.model.doc}, nil
	case "ancst", "ancstOrSelf":
		items := []*nativeDiagramPoint{}
		if axis == "ancstOrSelf" {
			items = append(items, point)
		}
		for current := parent; current != nil; current = current.parent {
			items = append(items, current)
		}
		return items, nil
	case "precedSib", "followSib", "preced", "follow":
		if parent == nil {
			return nil, nil
		}
		sequence := nativeDiagramSiblingSequence(parent)
		position := -1
		for index, item := range sequence {
			if item == point {
				position = index
			}
		}
		if position < 0 {
			return nil, nil
		}
		var peers []*nativeDiagramPoint
		if axis == "precedSib" || axis == "preced" {
			peers = sequence[:position]
		} else {
			peers = sequence[position+1:]
		}
		if axis == "precedSib" || axis == "followSib" {
			return peers, nil
		}
		items := []*nativeDiagramPoint{}
		for _, peer := range peers {
			items = append(items, peer)
			items = nativeDiagramDescendants(peer, items)
		}
		return items, nil
	case "none":
		return nil, nil
	}
	return nil, nativeDiagramLayoutRefuse(nativeDiagramLayoutDefinitionCode, "diagram layout axis "+axis+" is not modeled")
}

// evaluateCondition implements dgm:if (§21.4.2.15) for the var, cnt, depth and
// maxDepth functions. Position functions are outside the subset.
func (evaluator *nativeDiagramLayoutEvaluator) evaluateCondition(node *nativeXMLNode, pres *nativeDiagramPresNode, context *nativeDiagramPoint) (bool, error) {
	function, _ := exactNativeAttr(node, "", "func")
	operator, _ := exactNativeAttr(node, "", "op")
	expected, _ := exactNativeAttr(node, "", "val")
	if function == "var" {
		argument, _ := exactNativeAttr(node, "", "arg")
		actual := pres.lookupVariable(evaluator.model, argument)
		switch operator {
		case "equ":
			return actual == expected, nil
		case "neq":
			return actual != expected, nil
		}
		return false, nativeDiagramLayoutRefuse(nativeDiagramLayoutDefinitionCode, "diagram variable condition operator "+operator+" is not modeled")
	}
	selected, err := evaluator.selectPoints(context, node)
	if err != nil {
		return false, err
	}
	var actual int64
	switch function {
	case "cnt":
		actual = int64(len(selected))
	case "depth":
		if len(selected) != 0 {
			actual = int64(selected[0].depth)
		}
	case "maxDepth":
		for _, point := range selected {
			if relative := int64(point.depth - context.depth); relative > actual {
				actual = relative
			}
		}
	default:
		return false, nativeDiagramLayoutRefuse(nativeDiagramLayoutDefinitionCode, "diagram condition function "+function+" is not modeled")
	}
	want, err := strconv.ParseInt(expected, 10, 32)
	if err != nil {
		return false, nativeDiagramLayoutRefuse(nativeDiagramLayoutDefinitionCode, "diagram condition value is not an integer")
	}
	switch operator {
	case "equ":
		return actual == want, nil
	case "neq":
		return actual != want, nil
	case "gt":
		return actual > want, nil
	case "gte":
		return actual >= want, nil
	case "lt":
		return actual < want, nil
	case "lte":
		return actual <= want, nil
	}
	return false, nativeDiagramLayoutRefuse(nativeDiagramLayoutDefinitionCode, "diagram condition operator "+operator+" is not modeled")
}

// evaluateLayoutNode instantiates one layoutNode (§21.4.2.19) for a context
// point and appends it to parent.
func (evaluator *nativeDiagramLayoutEvaluator) evaluateLayoutNode(node *nativeXMLNode, context *nativeDiagramPoint, parent *nativeDiagramPresNode, depth int) (*nativeDiagramPresNode, error) {
	if depth > nativeMaxDiagramLayoutDepth {
		return nil, nativeDiagramLayoutRefuse(nativeDiagramLayoutBudgetCode, "diagram layout nesting exceeds the bounded depth")
	}
	evaluator.presNodes++
	if evaluator.presNodes > nativeMaxDiagramLayoutPresNodes {
		return nil, nativeDiagramLayoutRefuse(nativeDiagramLayoutBudgetCode, "diagram layout exceeds the bounded presentation node budget")
	}
	name, _ := exactNativeAttr(node, "", "name")
	styleLbl, _ := exactNativeAttr(node, "", "styleLbl")
	if styleLbl == "" && context != nil {
		// Layout nodes without styleLbl take the label PowerPoint recorded on
		// the matching presentation point; nothing is derived otherwise.
		styleLbl = evaluator.model.presLabels[context.id+"\x00"+name]
	}
	pres := &nativeDiagramPresNode{name: name, point: context, parent: parent, vars: map[string]string{}, params: map[string]string{}, styleLbl: styleLbl, vals: map[string]float64{}}
	if parent != nil {
		parent.children = append(parent.children, pres)
	}
	if err := evaluator.evaluateBody(node.Children, pres, context, depth, false); err != nil {
		return nil, err
	}
	return pres, nil
}

// evaluateBody walks layoutNode content: property elements attach to the
// current node; layoutNode/forEach/choose expand children. Inside a forEach
// body (loop=true) only expanding elements are valid.
func (evaluator *nativeDiagramLayoutEvaluator) evaluateBody(children []*nativeXMLNode, pres *nativeDiagramPresNode, context *nativeDiagramPoint, depth int, loop bool) error {
	for _, child := range children {
		if child.Name.Space != evaluator.ns {
			return nativeDiagramLayoutRefuse(nativeDiagramLayoutDefinitionCode, "diagram layout definition contains foreign markup")
		}
		switch child.Name.Local {
		case "layoutNode":
			if _, err := evaluator.evaluateLayoutNode(child, context, pres, depth+1); err != nil {
				return err
			}
			continue
		case "forEach":
			if err := evaluator.evaluateForEach(child, context, pres, depth+1); err != nil {
				return err
			}
			continue
		case "choose":
			branch, err := evaluator.chooseBranch(child, pres, context)
			if err != nil {
				return err
			}
			if branch != nil {
				if err := evaluator.evaluateBody(branch.Children, pres, context, depth, loop); err != nil {
					return err
				}
			}
			continue
		case "extLst":
			continue
		}
		if loop {
			return nativeDiagramLayoutRefuse(nativeDiagramLayoutDefinitionCode, "diagram forEach body contains a non-expanding element")
		}
		switch child.Name.Local {
		case "varLst":
			for _, variable := range child.Children {
				if variable.Name.Space != evaluator.ns || requireOnlyNativeAttrs(variable, xml.Name{Local: "val"}) != nil {
					return nativeDiagramLayoutRefuse(nativeDiagramLayoutDefinitionCode, "diagram layout variable is outside the modeled set")
				}
				value, ok := exactNativeAttr(variable, "", "val")
				if !ok {
					value = nativeDiagramLayoutVariableDefaults[variable.Name.Local]
				}
				pres.vars[variable.Name.Local] = value
			}
		case "alg":
			algType, _ := exactNativeAttr(child, "", "type")
			pres.alg = algType
			for _, param := range nativeChildren(child, evaluator.ns, "param") {
				key, _ := exactNativeAttr(param, "", "type")
				value, _ := exactNativeAttr(param, "", "val")
				pres.params[key] = value
			}
		case "shape":
			pres.hasShape = true
			pres.shapeType, _ = exactNativeAttr(child, "", "type")
			hide, _ := exactNativeAttr(child, "", "hideGeom")
			pres.hideGeom = hide == "1" || hide == "true"
			if rotation, ok := exactNativeAttr(child, "", "rot"); ok && rotation != "" && rotation != "0" {
				return nativeDiagramLayoutRefuse(nativeDiagramLayoutDefinitionCode, "diagram rotated layout shapes are not modeled")
			}
			// Blips are checked by local name: the relationship namespace differs
			// between transitional and strict packages.
			for _, attr := range child.Attrs {
				if attr.Name.Local == "blip" && attr.Value != "" {
					return nativeDiagramLayoutRefuse(nativeDiagramLayoutDefinitionCode, "diagram image shapes are not modeled")
				}
			}
			var err error
			if pres.zOrderOff, err = nativeDiagramIntAttr(child, "zOrderOff", 0); err != nil {
				return err
			}
			for _, adjLst := range nativeChildren(child, evaluator.ns, "adjLst") {
				if len(adjLst.Children) != 0 {
					return nativeDiagramLayoutRefuse(nativeDiagramLayoutDefinitionCode, "diagram shape adjust values are not modeled")
				}
			}
		case "presOf":
			selected, err := evaluator.selectPoints(context, child)
			if err != nil {
				return err
			}
			pres.presOf = selected
		case "constrLst":
			for _, constr := range child.Children {
				parsed, err := parseNativeDiagramConstraint(constr, evaluator.ns)
				if err != nil {
					return err
				}
				pres.constraints = append(pres.constraints, parsed)
			}
		case "ruleLst":
			for _, rule := range child.Children {
				parsed, err := parseNativeDiagramRule(rule, evaluator.ns)
				if err != nil {
					return err
				}
				pres.rules = append(pres.rules, parsed)
			}
		default:
			return nativeDiagramLayoutRefuse(nativeDiagramLayoutDefinitionCode, "diagram layout element "+child.Name.Local+" is not modeled")
		}
	}
	return nil
}

func (evaluator *nativeDiagramLayoutEvaluator) chooseBranch(choose *nativeXMLNode, pres *nativeDiagramPresNode, context *nativeDiagramPoint) (*nativeXMLNode, error) {
	for _, branch := range choose.Children {
		if branch.Name.Space != evaluator.ns {
			return nil, nativeDiagramLayoutRefuse(nativeDiagramLayoutDefinitionCode, "diagram choose contains foreign markup")
		}
		switch branch.Name.Local {
		case "if":
			matched, err := evaluator.evaluateCondition(branch, pres, context)
			if err != nil {
				return nil, err
			}
			if matched {
				return branch, nil
			}
		case "else":
			return branch, nil
		default:
			return nil, nativeDiagramLayoutRefuse(nativeDiagramLayoutDefinitionCode, "diagram choose contains an unknown branch")
		}
	}
	return nil, nil
}

func (evaluator *nativeDiagramLayoutEvaluator) evaluateForEach(node *nativeXMLNode, context *nativeDiagramPoint, parent *nativeDiagramPresNode, depth int) error {
	if depth > nativeMaxDiagramLayoutDepth {
		return nativeDiagramLayoutRefuse(nativeDiagramLayoutBudgetCode, "diagram layout nesting exceeds the bounded depth")
	}
	if ref, ok := exactNativeAttr(node, "", "ref"); ok && ref != "" {
		target := evaluator.forEachByName[ref]
		if target == nil {
			return nativeDiagramLayoutRefuse(nativeDiagramLayoutDefinitionCode, "diagram forEach references an unknown loop")
		}
		node = target
	}
	selected, err := evaluator.selectPoints(context, node)
	if err != nil {
		return err
	}
	for _, item := range selected {
		if err := evaluator.evaluateBody(node.Children, parent, item, depth, true); err != nil {
			return err
		}
	}
	return nil
}

func parseNativeDiagramConstraint(node *nativeXMLNode, diagramNS string) (nativeDiagramConstraint, error) {
	constraint := nativeDiagramConstraint{forRel: "self", refFor: "self", fact: 1}
	if node.Name != (xml.Name{Space: diagramNS, Local: "constr"}) || requireOnlyNativeChildren(node, xml.Name{Space: diagramNS, Local: "extLst"}) != nil {
		return constraint, nativeDiagramLayoutRefuse(nativeDiagramLayoutConstraintCode, "diagram constraint list contains unknown markup")
	}
	for _, attr := range node.Attrs {
		if attr.Name.Space != "" {
			return constraint, nativeDiagramLayoutRefuse(nativeDiagramLayoutConstraintCode, "diagram constraint carries a foreign attribute")
		}
		switch attr.Name.Local {
		case "type":
			constraint.typ = attr.Value
		case "for":
			constraint.forRel = attr.Value
		case "forName":
			constraint.forName = attr.Value
		case "ptType":
			constraint.ptType = attr.Value
		case "refType":
			constraint.refType = attr.Value
		case "refFor":
			constraint.refFor = attr.Value
		case "refForName":
			constraint.refForName = attr.Value
		case "refPtType":
			constraint.refPtType = attr.Value
		case "op":
			constraint.op = attr.Value
		case "val":
			value, err := nativeDiagramParseFloat(attr.Value)
			if err != nil {
				return constraint, err
			}
			constraint.val, constraint.hasVal = value, true
		case "fact":
			value, err := nativeDiagramParseFloat(attr.Value)
			if err != nil {
				return constraint, err
			}
			constraint.fact = value
		default:
			return constraint, nativeDiagramLayoutRefuse(nativeDiagramLayoutConstraintCode, "diagram constraint attribute "+attr.Name.Local+" is unknown")
		}
	}
	if constraint.typ == "" || constraint.typ == "none" {
		return constraint, nativeDiagramLayoutRefuse(nativeDiagramLayoutConstraintCode, "diagram constraint has no type")
	}
	switch constraint.op {
	case "", "none", "equ", "gte", "lte":
	default:
		return constraint, nativeDiagramLayoutRefuse(nativeDiagramLayoutConstraintCode, "diagram constraint operator "+constraint.op+" is unknown")
	}
	return constraint, nil
}

func nativeDiagramParseFloat(value string) (float64, error) {
	parsed, err := strconv.ParseFloat(strings.TrimSpace(value), 64)
	if err != nil || math.IsNaN(parsed) || math.IsInf(parsed, 0) || math.Abs(parsed) > 1e12 {
		return 0, nativeDiagramLayoutRefuse(nativeDiagramLayoutConstraintCode, "diagram constraint value is malformed or unbounded")
	}
	return parsed, nil
}

// parseNativeDiagramRule keeps the primFontSz lower bound (§21.4.2.24). Rules
// that grow or shrink geometry are outside the subset.
func parseNativeDiagramRule(node *nativeXMLNode, diagramNS string) (nativeDiagramRule, error) {
	rule := nativeDiagramRule{}
	if node.Name != (xml.Name{Space: diagramNS, Local: "rule"}) {
		return rule, nativeDiagramLayoutRefuse(nativeDiagramLayoutConstraintCode, "diagram rule list contains unknown markup")
	}
	rule.typ, _ = exactNativeAttr(node, "", "type")
	if rule.typ != "primFontSz" {
		return rule, nativeDiagramLayoutRefuse(nativeDiagramLayoutConstraintCode, "diagram rule type "+rule.typ+" is not modeled")
	}
	for _, attr := range node.Attrs {
		switch attr.Name.Local {
		case "val":
			value, err := nativeDiagramParseFloat(attr.Value)
			if err != nil {
				return rule, err
			}
			rule.val = value
		case "type", "for", "forName", "ptType":
		case "fact", "max":
			if attr.Value != "NaN" {
				return rule, nativeDiagramLayoutRefuse(nativeDiagramLayoutConstraintCode, "diagram rule factors are not modeled")
			}
		default:
			return rule, nativeDiagramLayoutRefuse(nativeDiagramLayoutConstraintCode, "diagram rule attribute "+attr.Name.Local+" is unknown")
		}
	}
	return rule, nil
}

// selectPres implements ST_ConstraintRelationship (§21.4.7.20) over the
// presentation tree.
func (node *nativeDiagramPresNode) selectPres(relationship, name, ptType string) []*nativeDiagramPresNode {
	matches := func(candidate *nativeDiagramPresNode) bool {
		return (name == "" || candidate.name == name) && (ptType == "" || (candidate.point != nil && nativeDiagramPointTypeMatches(ptType, candidate.point)))
	}
	switch relationship {
	case "", "self":
		return []*nativeDiagramPresNode{node}
	case "ch":
		out := []*nativeDiagramPresNode{}
		for _, child := range node.children {
			if matches(child) {
				out = append(out, child)
			}
		}
		return out
	case "des":
		out := []*nativeDiagramPresNode{}
		var walk func(current *nativeDiagramPresNode)
		walk = func(current *nativeDiagramPresNode) {
			for _, child := range current.children {
				if matches(child) {
					out = append(out, child)
				}
				walk(child)
			}
		}
		walk(node)
		return out
	}
	return nil
}

// evaluateConstraints applies every constrLst in presentation pre-order
// (§21.4.2.8). A constraint with op="equ" and neither val nor refType is an
// equalization directive; references read the current value of the referenced
// node and refuse when it was never set, so nothing is guessed.
func (evaluator *nativeDiagramLayoutEvaluator) evaluateConstraints(node *nativeDiagramPresNode) error {
	for _, constraint := range node.constraints {
		targets := node.selectPres(constraint.forRel, constraint.forName, constraint.ptType)
		if targets == nil {
			return nativeDiagramLayoutRefuse(nativeDiagramLayoutConstraintCode, "diagram constraint relationship "+constraint.forRel+" is not modeled")
		}
		evaluator.constraints += len(targets) + 1
		if evaluator.constraints > nativeMaxDiagramLayoutConstraints {
			return nativeDiagramLayoutRefuse(nativeDiagramLayoutBudgetCode, "diagram layout exceeds the bounded constraint budget")
		}
		if constraint.op == "equ" && !constraint.hasVal && constraint.refType == "" {
			if constraint.typ == "primFontSz" {
				node.equalize = append(node.equalize, targets)
			}
			continue
		}
		value := constraint.val
		if constraint.refType != "" {
			references := node.selectPres(constraint.refFor, constraint.refForName, constraint.refPtType)
			if len(references) == 0 {
				return nativeDiagramLayoutRefuse(nativeDiagramLayoutConstraintCode, "diagram constraint references a layout node that does not exist")
			}
			referenced, ok := references[0].vals[constraint.refType]
			if !ok {
				return nativeDiagramLayoutRefuse(nativeDiagramLayoutConstraintCode, "diagram constraint references "+constraint.refType+" before it is defined")
			}
			value = referenced * constraint.fact
		}
		for _, target := range targets {
			current, exists := target.vals[constraint.typ]
			switch constraint.op {
			case "gte":
				if exists && current > value {
					continue
				}
			case "lte":
				if exists && current < value {
					continue
				}
			}
			target.vals[constraint.typ] = value
		}
	}
	for _, child := range node.children {
		if err := evaluator.evaluateConstraints(child); err != nil {
			return err
		}
	}
	return nil
}
