package pptxpatch

import (
	"math"
	"strings"
)

// Layout algorithms (ECMA-376 Part 1 §21.4.7.1) for the hierarchy subset:
// composite, hierRoot, hierChild, sp, tx and conn. The spec names the
// algorithms and their parameters (§21.4.7.49) but leaves the arithmetic to
// the implementation, so every rule below is a declared approximation:
//
//   - composite places children by their l/t/r/b/w/h constraints;
//   - hierRoot stacks its root shape, then assistant blocks, then regular
//     child blocks, separated by the sp constraint, aligned by hierAlign
//     (bCtrCh default; tL/tR/bL/bR; alignOff as a fraction of the root width);
//   - hierChild lays out its child subtrees along linDir separated by sibSp,
//     or in two hanging columns around a trunk line (secLinDir/secChAlign);
//   - the finished tree is scaled uniformly to fit the frame and centered;
//   - conn routes bCtr/tCtr/midL/midR sites as straight or right-angle bend
//     polylines, honoring bendPt, bendDist, begPad and endPad;
//   - tx fits the primary font size between primFontSz and its rule minimum
//     using an average-advance glyph model.
//
// Nothing outside this subset is laid out: cycle, lin, pyra and snake refuse.
const (
	nativeDiagramLayoutAlgorithmCode = "pptx.diagram-layout-algorithm-unavailable"
	nativeDiagramLayoutGeometryCode  = "pptx.diagram-layout-geometry-unavailable"

	// Text fitting model: average Latin glyph advance and line height as
	// fractions of the font size. Declared in the preview diagnostic.
	nativeDiagramTextAdvanceFactor     = 0.5
	nativeDiagramTextLineHeightFactor  = 1.2
	nativeDiagramDefaultPrimFontSizePt = 65.0
	nativeDiagramMinimumFontSizePt     = 5.0
	nativeDiagramPointEMU              = 12700.0
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
	width := node.value("w", parentW)
	height := node.value("h", parentH)
	if width < 0 || height < 0 {
		return nativeDiagramLayoutRefuse(nativeDiagramLayoutConstraintCode, "diagram layout node has a negative size")
	}
	node.laidOut = true
	switch node.alg {
	case "", "sp", "tx":
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
	case "conn":
		return nil
	}
	return nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram layout algorithm "+node.alg+" is not implemented; only composite, hierRoot, hierChild, sp, tx and conn are approximated")
}

// layoutComposite positions children by explicit constraints (§21.4.7.1
// composite). Missing offsets default to 0; missing sizes to the composite.
func (node *nativeDiagramPresNode) layoutComposite(width, height float64) error {
	node.rect = nativeDiagramRect{0, 0, width, height}
	node.blockW, node.blockH, node.anchorX = width, height, width/2
	node.rootLeft, node.rootRight = 0, width
	for _, child := range node.children {
		if child.isConnector() {
			return nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram connectors inside composite nodes are not modeled")
		}
		switch child.alg {
		case "", "sp", "tx", "composite":
		default:
			return nativeDiagramLayoutRefuse(nativeDiagramLayoutAlgorithmCode, "diagram composite children must use composite, sp or tx algorithms")
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
		for _, item := range ordered {
			top := 0.0
			if chAlign == "b" {
				top = rowHeight - item.blockH
			}
			extend(item, left, top)
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
// [minimum, maximum] whose wrapped text fits the box, else the minimum.
func nativeDiagramFitFontSize(paragraphs [][]string, maximum, minimum, width, height float64) int64 {
	if maximum < minimum {
		maximum = minimum
	}
	for size := math.Floor(maximum); size >= minimum; size-- {
		if nativeDiagramTextFits(paragraphs, size, width, height) {
			return int64(size)
		}
	}
	return int64(math.Max(1, math.Ceil(minimum)))
}
