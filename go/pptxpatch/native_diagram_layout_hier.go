package pptxpatch

import (
	"math"
	"strings"
)

// Layout algorithms (ECMA-376 Part 1 §21.4.7.1) for the modeled subset:
// composite, lin, hierRoot, hierChild, sp, tx and conn. The spec names the
// algorithms and their parameters (§21.4.7.49) but leaves the arithmetic to
// the implementation, so every rule below is a declared approximation:
//
//   - composite places children by their l/t/r/b/w/h constraints;
//   - hierRoot stacks its root shape, then assistant blocks, then regular
//     child blocks, separated by the sp constraint, aligned by hierAlign
//     (bCtrCh default; bCtrDes/bL/bR; alignOff as a fraction of the root
//     width). DEVIATION: §21.4.7.36 describes tL/tR as children placed above
//     the parent; orgChart1 selects tL with alignOff for shallow subtrees to
//     hang leaf children below their parent, so tL/tR are laid out BELOW the
//     root, left/right edge offset by alignOff x root width. This is declared
//     in the group diagnostic;
//   - lin packs its children end to end along linDir and shrinks the row
//     uniformly when it overruns the node's own extent;
//   - snake wraps those children into a grid of whole lines;
//   - cycle spaces them around an ellipse from stAng across spanAng;
//   - hierChild lays out its child subtrees along linDir separated by sibSp,
//     packing horizontal siblings against each other's painted contours
//     rather than their whole envelopes, or in two hanging columns around a
//     trunk line (secLinDir/secChAlign);
//   - the finished tree is scaled uniformly to fit the frame and centered;
//   - conn routes bCtr/tCtr/midL/midR sites as straight or right-angle bend
//     polylines, honoring bendPt, bendDist, begPad and endPad;
//   - tx fits the primary font size between primFontSz and its rule minimum
//     using an average-advance glyph model.
//
// Nothing outside this subset is laid out: pyra refuses.
const (
	nativeDiagramLayoutAlgorithmCode = "pptx.diagram-layout-algorithm-unavailable"
	nativeDiagramLayoutGeometryCode  = "pptx.diagram-layout-geometry-unavailable"

	// Text fitting model: average Latin glyph advance and line height as
	// fractions of the font size. Declared in the preview diagnostic.
	nativeDiagramTextAdvanceFactor     = 0.5
	nativeDiagramTextLineHeightFactor  = 1.2
	nativeDiagramDefaultPrimFontSizePt = 65.0
	nativeDiagramMinimumFontSizePt     = 5.0
	// nativeDiagramMaxFontSizePt bounds primFontSz and rule minimums; larger
	// values refuse before any fitting happens.
	nativeDiagramMaxFontSizePt = 400.0
	nativeDiagramPointEMU      = 12700.0
)

type nativeDiagramRect struct {
	x, y, w, h float64
}

func (rect nativeDiagramRect) right() float64  { return rect.x + rect.w }
func (rect nativeDiagramRect) bottom() float64 { return rect.y + rect.h }

// translate shifts a laid-out subtree.
func (node *nativeDiagramPresNode) translate(dx, dy float64) {
	node.rect.x += dx
	node.rect.y += dy
	node.anchorX += dx
	node.rootLeft += dx
	node.rootRight += dx
	for _, child := range node.children {
		child.translate(dx, dy)
	}
}

func (node *nativeDiagramPresNode) isConnector() bool { return node.alg == "conn" }

// layoutSubtree computes node.rect and its descendants relative to the block
// origin (0,0), plus blockW/blockH/anchorX/rootLeft/rootRight.
func (node *nativeDiagramPresNode) layoutSubtree(parentW, parentH float64) error {
	width := node.elasticExtent("w", node.value("w", parentW), parentW)
	height := node.elasticExtent("h", node.value("h", parentH), parentH)
	// A pure spacer may be constrained to a negative extent: that is how the
	// linear layouts overlap consecutive shapes (chevron1's space node asks
	// for w = -0.1 x the shape width). Spacers never paint, so the negative
	// extent only moves the packing cursor. Everything else refuses.
	if (width < 0 || height < 0) && node.alg != "sp" {
		return nativeDiagramLayoutRefuse(nativeDiagramLayoutConstraintCode, "diagram layout node has a negative size")
	}
	node.laidOut = true
	switch node.alg {
	case "", "sp", "tx":
		if len(node.children) != 0 {
			return nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram layout nodes nested under sp, tx or algorithm-less nodes are not laid out")
		}
		node.rect = nativeDiagramRect{0, 0, width, height}
		node.blockW, node.blockH, node.anchorX = width, height, width/2
		node.rootLeft, node.rootRight = 0, width
		return nil
	case "composite":
		return node.layoutComposite(width, height)
	case "hierRoot":
		return node.layoutHierRoot(width, height)
	case "hierChild":
		return node.layoutHierChild(width, height)
	case "lin":
		return node.layoutLin(width, height)
	case "snake":
		return node.layoutSnake(width, height)
	case "cycle":
		return node.layoutCycle(width, height)
	case "conn":
		return nil
	}
	return nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram layout algorithm "+node.alg+" is not implemented; only composite, lin, snake, cycle, hierRoot, hierChild, sp, tx and conn are approximated")
}

// layoutComposite positions children by explicit constraints (§21.4.7.1
// composite). Missing offsets default to 0; missing sizes to the composite.
// A child whose extent carries an INF extent rule then grows into whatever
// the composite has left over, and a composite whose own extent is elastic
// closes on the content it ends up holding.
func (node *nativeDiagramPresNode) layoutComposite(width, height float64) error {
	node.rect = nativeDiagramRect{0, 0, width, height}
	node.blockW, node.blockH, node.anchorX = width, height, width/2
	node.rootLeft, node.rootRight = 0, width
	// Growth from an earlier solve is rolled back first: a linear parent
	// re-solves its children against a relaxed extent, and growth that
	// compounded across those passes would not be the share the constraint
	// system states.
	if node.resetElasticChildren() && node.evaluator != nil {
		if err := node.evaluator.applyConstraints(node, false); err != nil {
			return err
		}
	}
	if err := node.placeCompositeChildren(width, height); err != nil {
		return err
	}
	grownW := node.growCompositeChildren("w", width)
	grownH := node.growCompositeChildren("h", height)
	if grownW || grownH {
		// The offsets a composite derives from a child extent -- the band
		// below the title sits at the title's own h -- have to follow the
		// extent the algorithm just changed, so this node's own constraint
		// list is solved again before the children are placed again.
		if node.evaluator != nil {
			if err := node.evaluator.applyConstraints(node, false); err != nil {
				return err
			}
		}
		// The re-solve is there to move the offsets, not to undo the growth,
		// so an extent this composite also assigns is restored afterwards.
		for _, child := range node.children {
			for typ, value := range child.grown {
				child.vals[typ] = value
			}
		}
		if err := node.placeCompositeChildren(width, height); err != nil {
			return err
		}
	}
	if node.isElastic("w") {
		node.blockW = node.compositeContent("w")
		node.rect.w, node.anchorX, node.rootRight = node.blockW, node.blockW/2, node.blockW
	}
	if node.isElastic("h") {
		node.blockH = node.compositeContent("h")
		node.rect.h = node.blockH
	}
	return nil
}

// placeCompositeChildren lays out and positions every child from the current
// constraint values. It is idempotent, so a composite can solve, grow and
// place again.
func (node *nativeDiagramPresNode) placeCompositeChildren(width, height float64) error {
	for _, child := range node.children {
		if child.isConnector() {
			return nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram connectors inside composite nodes are not modeled")
		}
		switch child.alg {
		case "", "sp", "tx", "composite", "lin":
		default:
			return nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram composite children must use composite, lin, sp or tx algorithms")
		}
		if err := child.layoutSubtree(width, height); err != nil {
			return err
		}
		childW, childH := child.blockW, child.blockH
		left, top := child.value("l", 0), child.value("t", 0)
		if right, ok := child.vals["r"]; ok {
			if _, hasLeft := child.vals["l"]; !hasLeft {
				left = right - childW
			}
		}
		if bottom, ok := child.vals["b"]; ok {
			if _, hasTop := child.vals["t"]; !hasTop {
				top = bottom - childH
			}
		}
		if center, ok := child.vals["ctrX"]; ok {
			left = center - childW/2
		}
		if center, ok := child.vals["ctrY"]; ok {
			top = center - childH/2
		}
		child.translate(left, top)
	}
	return nil
}

// compositeContent measures how far the placed children reach along one axis.
func (node *nativeDiagramPresNode) compositeContent(typ string) float64 {
	content := 0.0
	for _, child := range node.children {
		reach := child.rect.y + child.blockH
		if typ == "w" {
			reach = child.rect.x + child.blockW
		}
		if reach > content {
			content = reach
		}
	}
	return content
}

// growCompositeChildren shares the extent a composite has left over among the
// children whose own ruleLst declares that extent unbounded (§21.4.2.24
// val="INF"). The share is proportional to the extent each child asked for,
// because that is the only ordering the constraint system states, and no
// child passes an op="lte" bound of its own -- the title band of the list
// layouts is exactly such a capped child, and PowerPoint stops it at its
// bound and gives the rest to the body band. A capped child's unused share is
// NOT redistributed: the composite closes on the content instead.
func (node *nativeDiagramPresNode) growCompositeChildren(typ string, extent float64) bool {
	if extent <= 0 {
		return false
	}
	elastic := []*nativeDiagramPresNode{}
	base := 0.0
	for _, child := range node.children {
		if !child.isElastic(typ) || child.isConnector() {
			continue
		}
		elastic = append(elastic, child)
		base += child.recordElasticBase(typ)
	}
	if len(elastic) == 0 || base <= 0 {
		return false
	}
	slack := extent - node.compositeContent(typ)
	if slack <= 0 {
		return false
	}
	grown := false
	for _, child := range elastic {
		current := child.elasticBase[typ]
		want := current * (1 + slack/base)
		if bound, ok := child.upper[typ]; ok && want > bound {
			want = bound
		}
		if want > current+1 {
			if child.grown == nil {
				child.grown = map[string]float64{}
			}
			child.vals[typ], child.grown[typ] = want, want
			grown = true
		}
	}
	return grown
}

// resetElasticChildren restores every extent a previous solve grew.
func (node *nativeDiagramPresNode) resetElasticChildren() bool {
	reset := false
	for _, child := range node.children {
		for typ, recorded := range child.elasticBase {
			delete(child.grown, typ)
			if child.vals[typ] != recorded {
				child.vals[typ] = recorded
				reset = true
			}
		}
	}
	return reset
}

// recordElasticBase returns the extent a node asked for before any growth,
// remembering it the first time so later solves grow from the constraint.
func (node *nativeDiagramPresNode) recordElasticBase(typ string) float64 {
	if recorded, ok := node.elasticBase[typ]; ok {
		return recorded
	}
	current, ok := node.vals[typ]
	if !ok {
		if typ == "w" {
			current = node.blockW
		} else {
			current = node.blockH
		}
	}
	if node.elasticBase == nil {
		node.elasticBase = map[string]float64{}
	}
	node.elasticBase[typ] = current
	return current
}

// isElastic reports whether an extent rule made this node's w or h unbounded.
func (node *nativeDiagramPresNode) isElastic(typ string) bool {
	for _, rule := range node.appliedRules {
		if rule.typ == typ && rule.unbounded {
			return true
		}
	}
	return false
}

// elasticExtent clamps an unbounded extent to what the parent has to give.
// The linear list layouts write h = 1000 x w on a composite they mean to be
// content-sized and leave the INF rule to settle it; read literally that is a
// block a thousand frames tall, and the whole diagram scales to a hairline.
func (node *nativeDiagramPresNode) elasticExtent(typ string, value, available float64) float64 {
	if value > available && available > 0 && node.isElastic(typ) {
		return available
	}
	return value
}

// layoutHierRoot: root shape on top, then assistant blocks, then the other
// child blocks, each separated by sp (§21.4.7.1 hierRoot, §21.4.7.36).
func (node *nativeDiagramPresNode) layoutHierRoot(width, height float64) error {
	var root *nativeDiagramPresNode
	blocks := []*nativeDiagramPresNode{}
	for _, child := range node.children {
		switch {
		case child.isConnector():
			return nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram connectors directly under hierRoot are not modeled")
		case child.alg == "hierChild":
			blocks = append(blocks, child)
		case root == nil:
			root = child
		default:
			return nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram hierRoot requires one root shape node followed by hierChild blocks")
		}
	}
	if root == nil {
		return nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram hierRoot has no root shape node")
	}
	if err := root.layoutSubtree(width, height); err != nil {
		return err
	}
	rootRect := nativeDiagramRect{0, 0, root.blockW, root.blockH}
	spacing := node.value("sp", 0)
	alignOff := node.value("alignOff", 0)
	align := node.param("hierAlign", "bCtrCh")
	ordered := make([]*nativeDiagramPresNode, 0, len(blocks))
	for _, block := range blocks {
		if block.leadsWithAssistant() {
			ordered = append(ordered, block)
		}
	}
	for _, block := range blocks {
		if !block.leadsWithAssistant() {
			ordered = append(ordered, block)
		}
	}
	cursor := rootRect.h
	minX, maxX := 0.0, rootRect.w
	for _, block := range ordered {
		if err := block.layoutSubtree(width, height); err != nil {
			return err
		}
		if block.blockW <= 0 && block.blockH <= 0 {
			continue
		}
		top := cursor + spacing
		var left float64
		switch align {
		case "bCtrCh":
			left = rootRect.w/2 + alignOff*rootRect.w - block.anchorX
		case "bCtrDes":
			left = rootRect.w/2 + alignOff*rootRect.w - block.blockW/2
		case "tL":
			left = alignOff * rootRect.w
		case "tR":
			left = rootRect.w - alignOff*rootRect.w - block.blockW
		case "bL":
			left = 0
		case "bR":
			left = rootRect.w - block.blockW
		default:
			return nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram hierarchy alignment "+align+" is not modeled")
		}
		block.translate(left, top)
		cursor = top + block.blockH
		minX = math.Min(minX, left)
		maxX = math.Max(maxX, left+block.blockW)
	}
	if minX != 0 {
		root.translate(-minX, 0)
		for _, block := range ordered {
			block.translate(-minX, 0)
		}
	}
	node.rect = nativeDiagramRect{0, 0, maxX - minX, cursor}
	node.blockW, node.blockH = maxX-minX, cursor
	node.anchorX = rootRect.w/2 - minX
	node.rootLeft, node.rootRight = -minX, rootRect.w-minX
	return nil
}

// appendPaintedRects collects the rectangles this subtree actually paints, in
// the subtree's own coordinates. The predicate matches the one the element
// writer uses, so container blocks (hierRoot, hierChild), hidden geometry and
// connectors never contribute to sibling spacing.
func (node *nativeDiagramPresNode) appendPaintedRects(into []nativeDiagramRect) []nativeDiagramRect {
	if node.hasShape && node.shapeType != "" && !node.hideGeom && node.laidOut && !node.isConnector() && node.rect.w > 0 && node.rect.h > 0 {
		into = append(into, node.rect)
	}
	for _, child := range node.children {
		into = child.appendPaintedRects(into)
	}
	return into
}

// nativeDiagramContourLeft returns the smallest horizontal offset at or above
// floor at which none of shapes (given in subtree coordinates, to be drawn at
// vertical offset top) comes within sibSp of an already placed rectangle it
// overlaps vertically.
func nativeDiagramContourLeft(placed, shapes []nativeDiagramRect, top, sibSp, floor float64) float64 {
	left := floor
	for _, shape := range shapes {
		shapeTop, shapeBottom := shape.y+top, shape.y+top+shape.h
		for _, rect := range placed {
			if shapeBottom <= rect.y || rect.bottom() <= shapeTop {
				continue
			}
			if required := rect.right() + sibSp - shape.x; required > left {
				left = required
			}
		}
	}
	return left
}

func (node *nativeDiagramPresNode) leadsWithAssistant() bool {
	for _, child := range node.children {
		if !child.isConnector() {
			return child.point != nil && child.point.kind == "asst"
		}
	}
	return false
}

// layoutHierChild arranges child subtrees linearly (linDir, chAlign, sibSp)
// or in two hanging columns around a trunk line when secLinDir is set
// (§21.4.7.1 hierChild, §21.4.7.13, §21.4.7.42, §21.4.7.56, §21.4.7.57).
func (node *nativeDiagramPresNode) layoutHierChild(width, height float64) error {
	items := []*nativeDiagramPresNode{}
	for _, child := range node.children {
		if child.isConnector() {
			continue
		}
		if err := child.layoutSubtree(width, height); err != nil {
			return err
		}
		items = append(items, child)
	}
	node.rect = nativeDiagramRect{}
	if len(items) == 0 {
		return nil
	}
	linDir := node.param("linDir", "fromL")
	chAlign := node.param("chAlign", "")
	secLinDir := node.param("secLinDir", "none")
	sibSp := node.value("sibSp", 0)
	secSibSp := node.value("secSibSp", 0)
	minX, maxX, maxY := math.Inf(1), math.Inf(-1), 0.0
	extend := func(item *nativeDiagramPresNode, left, top float64) {
		item.translate(left, top)
		minX = math.Min(minX, left)
		maxX = math.Max(maxX, left+item.blockW)
		maxY = math.Max(maxY, top+item.blockH)
	}
	switch {
	case secLinDir != "none" && secLinDir != "":
		// Both-hanging: items alternate left/right of the trunk at x=0,
		// rows advance along secLinDir, rows are top-aligned by default.
		if secLinDir != "fromT" || (linDir != "fromL" && linDir != "fromR") {
			return nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram hanging hierarchy direction "+linDir+"/"+secLinDir+" is not modeled")
		}
		secChAlign := node.param("secChAlign", "t")
		if secChAlign != "t" && secChAlign != "b" {
			return nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram hanging hierarchy alignment "+secChAlign+" is not modeled")
		}
		top := 0.0
		for row := 0; row < len(items); row += 2 {
			pair := items[row:int(math.Min(float64(row+2), float64(len(items))))]
			rowHeight := 0.0
			for _, item := range pair {
				rowHeight = math.Max(rowHeight, item.blockH)
			}
			for index, item := range pair {
				leftSide := index == 0
				if linDir == "fromR" {
					leftSide = !leftSide
				}
				left := sibSp / 2
				if leftSide {
					left = -sibSp/2 - item.blockW
				}
				itemTop := top
				if secChAlign == "b" {
					itemTop = top + rowHeight - item.blockH
				}
				extend(item, left, itemTop)
			}
			top += rowHeight + secSibSp
		}
		node.anchorX = 0
	case linDir == "fromL" || linDir == "fromR":
		if chAlign == "" {
			chAlign = "t"
		}
		if chAlign != "t" && chAlign != "b" {
			return nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram horizontal hierarchy child alignment "+chAlign+" is not modeled")
		}
		rowHeight := 0.0
		for _, item := range items {
			rowHeight = math.Max(rowHeight, item.blockH)
		}
		ordered := items
		if linDir == "fromR" {
			ordered = make([]*nativeDiagramPresNode, 0, len(items))
			for index := len(items) - 1; index >= 0; index-- {
				ordered = append(ordered, items[index])
			}
		}
		left := 0.0
		placed := []nativeDiagramRect{}
		for index, item := range ordered {
			top := 0.0
			if chAlign == "b" {
				top = rowHeight - item.blockH
			}
			shapes := item.appendPaintedRects(nil)
			if index > 0 {
				// Contour packing: sibling subtrees only have to clear each
				// other where they actually meet, so a shallow subtree tucks
				// under a deeper sibling's overhang instead of being pushed
				// past its whole envelope. Never looser than envelope
				// packing, and every painted pair still clears by sibSp.
				left = math.Min(left, nativeDiagramContourLeft(placed, shapes, top, sibSp, ordered[index-1].rect.x))
			}
			extend(item, left, top)
			for _, rect := range shapes {
				placed = append(placed, nativeDiagramRect{rect.x + left, rect.y + top, rect.w, rect.h})
			}
			left += item.blockW + sibSp
		}
		// Ch-style alignment centers the row of child shapes, not the whole
		// descendant envelope (§21.4.7.36 bCtrCh).
		node.anchorX = (ordered[0].rootLeft + ordered[len(ordered)-1].rootRight) / 2
	case linDir == "fromT" || linDir == "fromB":
		if chAlign == "" {
			chAlign = "l"
		}
		if chAlign != "l" && chAlign != "r" {
			return nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram vertical hierarchy child alignment "+chAlign+" is not modeled")
		}
		columnWidth := 0.0
		for _, item := range items {
			columnWidth = math.Max(columnWidth, item.blockW)
		}
		ordered := items
		if linDir == "fromB" {
			ordered = make([]*nativeDiagramPresNode, 0, len(items))
			for index := len(items) - 1; index >= 0; index-- {
				ordered = append(ordered, items[index])
			}
		}
		top := 0.0
		for _, item := range ordered {
			left := 0.0
			if chAlign == "r" {
				left = columnWidth - item.blockW
			}
			extend(item, left, top)
			top += item.blockH + sibSp
		}
		node.anchorX = columnWidth / 2
	default:
		return nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram linear direction "+linDir+" is not modeled")
	}
	if minX != 0 {
		for _, item := range items {
			item.translate(-minX, 0)
		}
		node.anchorX -= minX
	}
	node.blockW, node.blockH = maxX-minX, maxY
	node.rect = nativeDiagramRect{0, 0, node.blockW, node.blockH}
	node.rootLeft, node.rootRight = items[0].rootLeft, items[len(items)-1].rootRight
	return nil
}

// nativeDiagramFitTransform maps the unscaled layout onto the frame: uniform
// scale to fit, centered in both directions.
type nativeDiagramFitTransform struct {
	scale          float64
	offsetX        float64
	offsetY        float64
	frameW, frameH float64
}

func nativeDiagramFit(root *nativeDiagramPresNode, frameW, frameH float64) (nativeDiagramFitTransform, error) {
	if root.blockW <= 0 || root.blockH <= 0 || frameW <= 0 || frameH <= 0 {
		return nativeDiagramFitTransform{}, nativeDiagramLayoutRefuse(nativeDiagramLayoutGeometryCode, "diagram layout produced no visible extent")
	}
	scale := math.Min(frameW/root.blockW, frameH/root.blockH)
	return nativeDiagramFitTransform{
		scale: scale, frameW: frameW, frameH: frameH,
		offsetX: (frameW - root.blockW*scale) / 2,
		offsetY: (frameH - root.blockH*scale) / 2,
	}, nil
}

func (fit nativeDiagramFitTransform) rect(rect nativeDiagramRect) nativeDiagramRect {
	return nativeDiagramRect{fit.offsetX + rect.x*fit.scale, fit.offsetY + rect.y*fit.scale, rect.w * fit.scale, rect.h * fit.scale}
}

func nativeDiagramRoundEMU(value float64) int64 {
	return int64(math.Round(value))
}

// nativeDiagramConnectorSite resolves a ST_ConnectorPoint (§21.4.7.18) on a
// frame-space rectangle.
func nativeDiagramConnectorSite(rect nativeDiagramRect, site string) (float64, float64, bool) {
	switch site {
	case "bCtr":
		return rect.x + rect.w/2, rect.bottom(), true
	case "tCtr":
		return rect.x + rect.w/2, rect.y, true
	case "midL":
		return rect.x, rect.y + rect.h/2, true
	case "midR":
		return rect.right(), rect.y + rect.h/2, true
	case "ctr":
		return rect.x + rect.w/2, rect.y + rect.h/2, true
	case "tL":
		return rect.x, rect.y, true
	case "tR":
		return rect.right(), rect.y, true
	case "bL":
		return rect.x, rect.bottom(), true
	case "bR":
		return rect.right(), rect.bottom(), true
	}
	return 0, 0, false
}

func nativeDiagramSiteIsVertical(site string) bool {
	return site == "bCtr" || site == "tCtr"
}

type nativeDiagramPoint2 struct{ x, y float64 }

// routeNativeDiagramConnector builds the polyline for a conn node (§21.4.7.19
// bend/stra, §21.4.7.8 bendPt) between the begin and end rectangles in frame
// space. bendDist and the pads are already scaled.
func routeNativeDiagramConnector(node *nativeDiagramPresNode, begin, end nativeDiagramRect, bendDist float64, hasBendDist bool, begPad, endPad float64) ([]nativeDiagramPoint2, error) {
	routing := node.param("connRout", "stra")
	if routing != "bend" && routing != "stra" {
		return nil, nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram connector routing "+routing+" is not modeled")
	}
	if dim := node.param("dim", "1D"); dim != "1D" {
		return nil, nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram two-dimensional connectors are not modeled")
	}
	for _, style := range []string{node.param("begSty", "noArr"), node.param("endSty", "noArr")} {
		if style != "noArr" {
			return nil, nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram connector arrowheads are not modeled")
		}
	}
	begSite := node.param("begPts", "auto")
	begX, begY, ok := nativeDiagramConnectorSite(begin, begSite)
	if !ok || strings.Contains(begSite, " ") {
		return nil, nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram connector begin site "+begSite+" is not modeled")
	}
	endSite := ""
	endX, endY := 0.0, 0.0
	bestDistance := math.Inf(1)
	for _, candidate := range strings.Fields(node.param("endPts", "auto")) {
		x, y, ok := nativeDiagramConnectorSite(end, candidate)
		if !ok {
			return nil, nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram connector end site "+candidate+" is not modeled")
		}
		// Among the offered sites the algorithm picks the one nearest to the
		// begin site (§21.4.7.18 lists the sites, not the choice rule).
		if distance := math.Abs(x-begX) + math.Abs(y-begY); distance < bestDistance {
			bestDistance, endSite, endX, endY = distance, candidate, x, y
		}
	}
	if endSite == "" {
		return nil, nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram connector end site is not modeled")
	}
	points := []nativeDiagramPoint2{{begX, begY}}
	if routing == "bend" && !(begX == endX || begY == endY) {
		begVertical, endVertical := nativeDiagramSiteIsVertical(begSite), nativeDiagramSiteIsVertical(endSite)
		switch {
		case begVertical && endVertical:
			bendY := (begY + endY) / 2
			if hasBendDist {
				switch node.param("bendPt", "def") {
				case "end":
					bendY = endY - math.Copysign(bendDist, endY-begY)
				case "beg", "def":
					bendY = begY + math.Copysign(bendDist, endY-begY)
				}
			}
			points = append(points, nativeDiagramPoint2{begX, bendY}, nativeDiagramPoint2{endX, bendY})
		case !begVertical && !endVertical:
			bendX := (begX + endX) / 2
			if hasBendDist {
				switch node.param("bendPt", "def") {
				case "end":
					bendX = endX - math.Copysign(bendDist, endX-begX)
				case "beg", "def":
					bendX = begX + math.Copysign(bendDist, endX-begX)
				}
			}
			points = append(points, nativeDiagramPoint2{bendX, begY}, nativeDiagramPoint2{bendX, endY})
		case begVertical:
			points = append(points, nativeDiagramPoint2{begX, endY})
		default:
			points = append(points, nativeDiagramPoint2{endX, begY})
		}
	}
	points = append(points, nativeDiagramPoint2{endX, endY})
	// begPad/endPad shorten the terminal segments (§21.4.7.21).
	points = nativeDiagramShortenSegment(points, begPad, false)
	points = nativeDiagramShortenSegment(points, endPad, true)
	return points, nil
}

func nativeDiagramShortenSegment(points []nativeDiagramPoint2, pad float64, atEnd bool) []nativeDiagramPoint2 {
	if pad <= 0 || len(points) < 2 {
		return points
	}
	from, to := 0, 1
	if atEnd {
		from, to = len(points)-1, len(points)-2
	}
	dx, dy := points[to].x-points[from].x, points[to].y-points[from].y
	length := math.Hypot(dx, dy)
	if length <= pad {
		return points
	}
	points[from].x += dx / length * pad
	points[from].y += dy / length * pad
	return points
}

// nativeDiagramTextLines counts wrapped lines for the declared glyph model.
// Words are split on spaces; a word wider than the box fails the fit.
func nativeDiagramTextFits(paragraphs [][]string, fontPt, width, height float64) bool {
	em := fontPt * nativeDiagramPointEMU
	advance := em * nativeDiagramTextAdvanceFactor
	lineHeight := em * nativeDiagramTextLineHeightFactor
	lines := 0
	for _, words := range paragraphs {
		lines++
		lineWidth := 0.0
		for _, word := range words {
			wordWidth := float64(len([]rune(word))) * advance
			if wordWidth > width {
				return false
			}
			if lineWidth == 0 {
				lineWidth = wordWidth
			} else if lineWidth+advance+wordWidth <= width {
				lineWidth += advance + wordWidth
			} else {
				lines++
				lineWidth = wordWidth
			}
		}
	}
	return float64(lines)*lineHeight <= height
}

// nativeDiagramFitFontSize returns the largest whole point size in
// [minimum, maximum] whose wrapped text fits the box, else the minimum. Fit
// is monotone in the size, so a binary search over the bounded range needs
// at most a few probes; callers refuse sizes above nativeDiagramMaxFontSizePt.
func nativeDiagramFitFontSize(paragraphs [][]string, maximum, minimum, width, height float64) int64 {
	low := int64(math.Max(1, math.Ceil(minimum)))
	high := int64(math.Floor(math.Min(maximum, nativeDiagramMaxFontSizePt)))
	if high < low {
		return low
	}
	best := low
	for probes := 0; low <= high && probes < 16; probes++ {
		mid := low + (high-low)/2
		if nativeDiagramTextFits(paragraphs, float64(mid), width, height) {
			best, low = mid, mid+1
		} else {
			high = mid - 1
		}
	}
	return best
}

// layoutLin packs children end to end along linDir, aligned across that axis
// by nodeVertAlign/nodeHorzAlign (§21.4.7.1 lin, §21.4.7.42, §21.4.7.45,
// §21.4.7.46). Spacer nodes contribute their extent along the packing axis
// only; a negative spacer overlaps its neighbors. When the packed row or
// column overruns the node's own extent every child is scaled by the one
// factor that makes it fit.
func (node *nativeDiagramPresNode) layoutLin(width, height float64) error {
	items := []*nativeDiagramPresNode{}
	for _, child := range node.children {
		if child.isConnector() {
			return nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram connectors inside linear nodes are not modeled")
		}
		if err := child.layoutSubtree(width, height); err != nil {
			return err
		}
		items = append(items, child)
	}
	node.rect = nativeDiagramRect{0, 0, width, height}
	node.blockW, node.blockH, node.anchorX = width, height, width/2
	node.rootLeft, node.rootRight = 0, width
	if len(items) == 0 {
		return nil
	}
	linDir := node.param("linDir", "fromL")
	horizontal := linDir == "fromL" || linDir == "fromR"
	if !horizontal && linDir != "fromT" && linDir != "fromB" {
		return nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram linear direction "+linDir+" is not modeled")
	}
	align := node.param("nodeVertAlign", "ctr")
	if !horizontal {
		align = node.param("nodeHorzAlign", "ctr")
	}
	switch {
	case horizontal && (align == "t" || align == "ctr" || align == "b"):
	case !horizontal && (align == "l" || align == "ctr" || align == "r"):
	default:
		return nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram linear node alignment "+align+" is not modeled")
	}
	along := func(item *nativeDiagramPresNode) float64 {
		if horizontal {
			return item.blockW
		}
		return item.blockH
	}
	cross := func(item *nativeDiagramPresNode) float64 {
		if horizontal {
			return item.blockH
		}
		return item.blockW
	}
	total := 0.0
	for _, item := range items {
		total += along(item)
	}
	available := width
	if !horizontal {
		available = height
	}
	if total > available && available > 0 && total > 0 {
		// Shrink to fit. ECMA-376 §21.4.2.24 lets the algorithm relax the
		// extent a child asked for down to its ruleLst minimum; every
		// constraint the child's own list derives FROM that extent then
		// re-solves, while what an ancestor assigned it across the packing
		// axis stays put. Re-running the child's own constraint list over
		// the relaxed extent is exactly that: a chevron's h = 0.4 x w
		// follows its narrowed width, a pillar's h = 0.63 x the frame does
		// not. The share is proportional, so equally constrained children
		// end up with equal shares.
		factor := available / total
		keys := nativeDiagramHorizontalConstraints
		if !horizontal {
			keys = nativeDiagramVerticalConstraints
		}
		for _, item := range items {
			if err := item.relaxAlong(keys, along(item)*factor, factor, width, height); err != nil {
				return err
			}
		}
		total = 0
		for _, item := range items {
			total += along(item)
		}
	}
	extent := 0.0
	for _, item := range items {
		// A spacer's cross extent is the inherited parent size, not a
		// measurement of anything it paints, so it never widens the row.
		if item.alg != "sp" && cross(item) > extent {
			extent = cross(item)
		}
	}
	ordered := items
	if linDir == "fromR" || linDir == "fromB" {
		ordered = make([]*nativeDiagramPresNode, 0, len(items))
		for index := len(items) - 1; index >= 0; index-- {
			ordered = append(ordered, items[index])
		}
	}
	cursor := 0.0
	for _, item := range ordered {
		offset := 0.0
		switch align {
		case "ctr":
			offset = (extent - cross(item)) / 2
		case "b", "r":
			offset = extent - cross(item)
		}
		if item.alg == "sp" {
			offset = 0
		}
		if horizontal {
			item.translate(cursor, offset)
		} else {
			item.translate(offset, cursor)
		}
		cursor += along(item)
	}
	if horizontal {
		node.blockW, node.blockH = math.Max(total, 0), extent
	} else {
		node.blockW, node.blockH = extent, math.Max(total, 0)
	}
	node.rect = nativeDiagramRect{0, 0, node.blockW, node.blockH}
	node.anchorX = node.blockW / 2
	node.rootLeft, node.rootRight = 0, node.blockW
	return nil
}

// nativeDiagramHorizontalConstraints and nativeDiagramVerticalConstraints
// name the modeled constraint types measured along one axis. A linear
// algorithm relaxes the extent it packs along, and every extent and offset a
// layout derived from it -- including the ones an ancestor assigned to a
// descendant -- is measured in the same relaxed units. The other axis is
// untouched.
var (
	nativeDiagramHorizontalConstraints = []string{"w", "l", "r", "ctrX"}
	nativeDiagramVerticalConstraints   = []string{"h", "t", "b", "ctrY"}
)

// relaxAlong re-solves one child of a linear node against a relaxed extent.
// ECMA-376 §21.4.2.24 lets the algorithm shrink a child to make the row fit,
// and PowerPoint re-solves the constraint system in the relaxed units. Here:
// every value in the subtree measured along the packing axis is scaled, the
// subtree's own constraint lists are evaluated again so relations inside it
// re-derive, and the allocation is restored -- the algorithm's share wins
// over the node's own preferred value. Equalization directives are cleared
// first so the second pass does not record the same group twice; rules aimed
// at the subtree from above it are kept, because only the subtree's own
// lists are re-evaluated.
func (node *nativeDiagramPresNode) relaxAlong(keys []string, extent, factor, parentW, parentH float64) error {
	node.scaleConstraints(keys, factor)
	node.clearEqualization()
	node.vals[keys[0]] = extent
	if node.evaluator != nil {
		if err := node.evaluator.evaluateConstraints(node); err != nil {
			return err
		}
	}
	node.vals[keys[0]] = extent
	return node.layoutSubtree(parentW, parentH)
}

func (node *nativeDiagramPresNode) scaleConstraints(keys []string, factor float64) {
	for _, key := range keys {
		if value, ok := node.vals[key]; ok {
			node.vals[key] = value * factor
		}
		if value, ok := node.elasticBase[key]; ok {
			node.elasticBase[key] = value * factor
		}
		if value, ok := node.upper[key]; ok {
			node.upper[key] = value * factor
		}
	}
	for _, child := range node.children {
		child.scaleConstraints(keys, factor)
	}
}

func (node *nativeDiagramPresNode) clearEqualization() {
	node.equalize = nil
	for _, child := range node.children {
		child.clearEqualization()
	}
}

// layoutSnake wraps the children into a grid of whole lines (§21.4.7.1
// snake, §21.4.7.35 grDir, §21.4.7.33 flowDir, §21.4.7.23 contDir). ECMA
// names the parameters and leaves the line breaking to the implementation.
// DEVIATION, declared in the group diagnostic: the break point is chosen by
// fit rather than by ST_BreakpointType, because the list layouts give every
// item the WHOLE diagram extent and rely on the algorithm to divide it. The
// line count that leaves the packed grid closest in aspect ratio to the node
// it fills is the one used, which is the grid PowerPoint draws for the block
// lists in the corpus (four equal items in a 3:2 frame become two rows of
// two, not four rows of one).
func (node *nativeDiagramPresNode) layoutSnake(width, height float64) error {
	items := []*nativeDiagramPresNode{}
	for _, child := range node.children {
		if child.isConnector() {
			return nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram connectors inside snake nodes are not modeled")
		}
		if err := child.layoutSubtree(width, height); err != nil {
			return err
		}
		items = append(items, child)
	}
	node.rect = nativeDiagramRect{0, 0, width, height}
	node.blockW, node.blockH, node.anchorX = width, height, width/2
	node.rootLeft, node.rootRight = 0, width
	if len(items) == 0 {
		return nil
	}
	growth := node.param("grDir", "tL")
	flow := node.param("flowDir", "row")
	continuation := node.param("contDir", "sameDir")
	if growth != "tL" && growth != "tR" && growth != "bL" && growth != "bR" {
		return nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram snake growth direction "+growth+" is not modeled")
	}
	if flow != "row" && flow != "col" {
		return nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram snake flow direction "+flow+" is not modeled")
	}
	if continuation != "sameDir" && continuation != "revDir" {
		return nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram snake continuation "+continuation+" is not modeled")
	}
	// Spacers separate items inside a line; the sp constraint separates the
	// lines themselves.
	cells, gap := []*nativeDiagramPresNode{}, 0.0
	for _, item := range items {
		if item.alg == "sp" {
			gap = math.Max(gap, item.blockW)
			continue
		}
		cells = append(cells, item)
	}
	if len(cells) == 0 {
		return nil
	}
	lineGap := node.value("sp", gap)
	cellW, cellH := 0.0, 0.0
	for _, cell := range cells {
		cellW, cellH = math.Max(cellW, cell.blockW), math.Max(cellH, cell.blockH)
	}
	if cellW <= 0 || cellH <= 0 {
		return nativeDiagramLayoutRefuse(nativeDiagramLayoutGeometryCode, "diagram snake children have no extent to wrap")
	}
	across := nativeDiagramSnakeLineLength(len(cells), cellW, cellH, gap, lineGap, width, height, flow == "col")
	columns, rows := across, (len(cells)+across-1)/across
	if flow == "col" {
		rows, columns = across, (len(cells)+across-1)/across
	}
	factor := 1.0
	if packed := float64(columns)*cellW + float64(columns-1)*gap; packed > width && width > 0 {
		factor = width / packed
	}
	if packed := float64(rows)*cellH + float64(rows-1)*lineGap; packed*factor > height && height > 0 {
		factor = math.Min(factor, height/packed)
	}
	if factor != 1 {
		for index, cell := range cells {
			if err := cell.relaxAlong(nativeDiagramHorizontalConstraints, cell.blockW*factor, factor, width, height); err != nil {
				return err
			}
			if err := cell.relaxAlong(nativeDiagramVerticalConstraints, cell.blockH*factor, factor, width, height); err != nil {
				return err
			}
			cells[index] = cell
		}
		cellW, cellH, gap, lineGap = cellW*factor, cellH*factor, gap*factor, lineGap*factor
	}
	for index, cell := range cells {
		line, offset := index/across, index%across
		if continuation == "revDir" && line%2 == 1 {
			offset = across - 1 - offset
			if last := len(cells) - line*across; line == (len(cells)-1)/across && last < across {
				offset -= across - last
			}
		}
		column, row := offset, line
		if flow == "col" {
			column, row = line, offset
		}
		if growth == "tR" || growth == "bR" {
			column = columns - 1 - column
		}
		if growth == "bL" || growth == "bR" {
			row = rows - 1 - row
		}
		// Lines are centered across the node, which is what off="ctr" asks
		// for and what a full line produces anyway.
		cell.translate(float64(column)*(cellW+gap)+(cellW-cell.blockW)/2, float64(row)*(cellH+lineGap)+(cellH-cell.blockH)/2)
	}
	node.blockW = float64(columns)*cellW + float64(columns-1)*gap
	node.blockH = float64(rows)*cellH + float64(rows-1)*lineGap
	node.rect = nativeDiagramRect{0, 0, node.blockW, node.blockH}
	node.anchorX = node.blockW / 2
	node.rootLeft, node.rootRight = 0, node.blockW
	return nil
}

// nativeDiagramSnakeLineLength picks how many cells share a line: the count
// whose packed grid comes closest in aspect ratio to the extent it has to
// fill. Ties go to the shorter line, so a single cell never starts a grid.
func nativeDiagramSnakeLineLength(count int, cellW, cellH, gap, lineGap, width, height float64, byColumn bool) int {
	if count <= 1 || width <= 0 || height <= 0 {
		return 1
	}
	target := width / height
	best, bestErr := 1, math.Inf(1)
	for across := 1; across <= count; across++ {
		down := (count + across - 1) / across
		columns, rows := across, down
		if byColumn {
			columns, rows = down, across
		}
		packedW := float64(columns)*cellW + float64(columns-1)*gap
		packedH := float64(rows)*cellH + float64(rows-1)*lineGap
		if packedW <= 0 || packedH <= 0 {
			continue
		}
		if distance := math.Abs(math.Log(packedW/packedH) - math.Log(target)); distance < bestErr-1e-9 {
			best, bestErr = across, distance
		}
	}
	return best
}

// layoutCycle spaces the children around an ellipse inscribed in the node
// (§21.4.7.1 cycle, §21.4.7.53 stAng, §21.4.7.52 spanAng, §21.4.7.22
// ctrShpMap, §21.4.7.50 rotPath). Angles are degrees clockwise from twelve
// o'clock; a negative span runs anticlockwise. With ctrShpMap="fNode" the
// first child is the hub and sits at the centre instead of on the ring.
// rotPath="alongPath" turns each shape to face along the ring.
func (node *nativeDiagramPresNode) layoutCycle(width, height float64) error {
	items := []*nativeDiagramPresNode{}
	for _, child := range node.children {
		if child.isConnector() {
			return nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram connectors inside cycle nodes are not modeled")
		}
		if err := child.layoutSubtree(width, height); err != nil {
			return err
		}
		if child.alg == "sp" {
			continue
		}
		items = append(items, child)
	}
	node.rect = nativeDiagramRect{0, 0, width, height}
	node.blockW, node.blockH, node.anchorX = width, height, width/2
	node.rootLeft, node.rootRight = 0, width
	if len(items) == 0 {
		return nil
	}
	start, err := nativeDiagramAngleParam(node, "stAng", 0)
	if err != nil {
		return err
	}
	span, err := nativeDiagramAngleParam(node, "spanAng", 360)
	if err != nil {
		return err
	}
	switch node.param("rotPath", "none") {
	case "none", "alongPath":
	default:
		return nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram cycle path rotation "+node.param("rotPath", "none")+" is not modeled")
	}
	alongPath := node.param("rotPath", "none") == "alongPath"
	var hub *nativeDiagramPresNode
	switch node.param("ctrShpMap", "none") {
	case "none":
	case "fNode":
		hub, items = items[0], items[1:]
	default:
		return nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram cycle centre mapping "+node.param("ctrShpMap", "none")+" is not modeled")
	}
	if hub != nil {
		hub.translate((width-hub.blockW)/2, (height-hub.blockH)/2)
	}
	if len(items) == 0 {
		return nil
	}
	cellW, cellH := 0.0, 0.0
	for _, item := range items {
		cellW, cellH = math.Max(cellW, item.blockW), math.Max(cellH, item.blockH)
	}
	// A full turn puts the last shape back on the first, so the step divides
	// the span by the count; a partial arc reaches its far end instead.
	steps := float64(len(items))
	if math.Abs(span) < 360 && len(items) > 1 {
		steps = float64(len(items) - 1)
	}
	// The cycle layouts give every shape the whole diagram extent and leave
	// the ring to the algorithm, the same way the linear ones do. Shrink
	// until neighbours clear each other on the ring: two shapes a step apart
	// on an ellipse of radius r are 2 r sin(step/2) apart, and the ring plus
	// one shape has to fit the node.
	factor := 1.0
	if half := math.Abs(span) * math.Pi / (360 * steps); half > 0 && half < math.Pi/2 {
		clearance := math.Sin(half)
		if cellW > 0 && width > 0 {
			factor = math.Min(factor, width*clearance/(cellW*(1+clearance)))
		}
		if cellH > 0 && height > 0 {
			factor = math.Min(factor, height*clearance/(cellH*(1+clearance)))
		}
	}
	if cellW > 0 && width > 0 {
		factor = math.Min(factor, width/cellW)
	}
	if cellH > 0 && height > 0 {
		factor = math.Min(factor, height/cellH)
	}
	if factor < 1 {
		for _, item := range items {
			if err := item.relaxAlong(nativeDiagramHorizontalConstraints, item.blockW*factor, factor, width, height); err != nil {
				return err
			}
			if err := item.relaxAlong(nativeDiagramVerticalConstraints, item.blockH*factor, factor, width, height); err != nil {
				return err
			}
		}
		cellW, cellH = cellW*factor, cellH*factor
	}
	// The ring is the largest ellipse that keeps every shape inside the node.
	radiusX, radiusY := math.Max((width-cellW)/2, 0), math.Max((height-cellH)/2, 0)
	for index, item := range items {
		degrees := start + span*float64(index)/steps
		radians := degrees * math.Pi / 180
		centerX := width/2 + radiusX*math.Sin(radians)
		centerY := height/2 - radiusY*math.Cos(radians)
		item.translate(centerX-item.blockW/2, centerY-item.blockH/2)
		if alongPath {
			item.setPathRotation(degrees)
		}
	}
	return nil
}

// setPathRotation turns a laid-out shape to face along the ring. Only the
// shape itself turns; its descendants keep the rotation they were authored
// with, because the contract carries one rotation per element.
func (node *nativeDiagramPresNode) setPathRotation(degrees float64) {
	units := int64(math.Round(degrees*60000)) % 21600000
	if units < 0 {
		units += 21600000
	}
	node.rotation60000 = units
	for _, child := range node.children {
		child.setPathRotation(degrees)
	}
}

// nativeDiagramAngleParam reads a degree-valued algorithm parameter.
func nativeDiagramAngleParam(node *nativeDiagramPresNode, name string, fallback float64) (float64, error) {
	raw := node.param(name, "")
	if raw == "" {
		return fallback, nil
	}
	degrees, err := nativeDiagramParseFloat(raw)
	if err != nil || math.Abs(degrees) > 3600 {
		return 0, nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram cycle angle "+name+"="+raw+" is outside the modeled range")
	}
	return degrees, nil
}
