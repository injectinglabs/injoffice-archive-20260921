package pptxpatch

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"path"
	"strconv"
	"strings"
)

// Slide WRITING: build a real .pptx from scratch — a Deck compiles straight
// to bytes, no "patch an existing file" step (unlike xlsxpatch.AddShape,
// there is no existing presentation to extend yet; DeckSpec→Deck is new
// authoring, not an edit of a file the user already has). The minimal-parts
// skeleton (one slideMaster/slideLayout/theme, docProps, content types, root
// rels) follows ECMA-376 PresentationML and the Open Packaging Convention:
// a presentation part, one slide master, one blank slide layout, a theme,
// package relationships, and content types, written in Go for N slides.

const (
	xmlDecl = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n"
	nsA     = "http://schemas.openxmlformats.org/drawingml/2006/main"
	nsC     = "http://schemas.openxmlformats.org/drawingml/2006/chart"
	nsR     = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	nsP     = "http://schemas.openxmlformats.org/presentationml/2006/main"

	relTypeOfficeDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
	relTypeCoreProps      = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
	relTypeExtProps       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
	relTypeSlideMaster    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
	relTypeSlideLayout    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
	relTypeSlide          = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
	relTypeTheme          = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"

	ctPresentation = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
	ctSlideMaster  = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"
	ctSlideLayout  = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
	ctSlide        = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
	ctTheme        = "application/vnd.openxmlformats-officedocument.theme+xml"
	ctCoreProps    = "application/vnd.openxmlformats-package.core-properties+xml"
	ctExtProps     = "application/vnd.openxmlformats-officedocument.extended-properties+xml"
	ctChart        = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
)

func esc(s string) string {
	var b bytes.Buffer
	_ = xml.EscapeText(&b, []byte(s))
	return b.String()
}

// hexColor strips a leading # and upcases — srgbClr@val takes bare hex,
// same convention as xlsxpatch's hexClr.
func hexColor(s string) string {
	s = strings.TrimPrefix(s, "#")
	return strings.ToUpper(s)
}

func fillXML(hex string) string {
	if hex == "" {
		return `<a:noFill/>`
	}
	return fmt.Sprintf(`<a:solidFill><a:srgbClr val=%q/></a:solidFill>`, hexColor(hex))
}

func lineXML(hex string, widthPt float64) string {
	return lineXMLArrows(hex, widthPt, false, false)
}

// lineXMLArrows is lineXML plus optional head/tail triangle arrowheads —
// CT_LineProperties' child order is fill, then (unused by this package)
// dash/cap elements, then headEnd, then tailEnd, so the arrows must come
// after the solidFill, never before it.
func lineXMLArrows(hex string, widthPt float64, headArrow, tailArrow bool) string {
	if hex == "" {
		return `<a:ln><a:noFill/></a:ln>`
	}
	w := widthPt
	if w <= 0 {
		w = 1
	}
	var arrows strings.Builder
	if headArrow {
		arrows.WriteString(`<a:headEnd type="triangle"/>`)
	}
	if tailArrow {
		arrows.WriteString(`<a:tailEnd type="triangle"/>`)
	}
	return fmt.Sprintf(`<a:ln w="%d"><a:solidFill><a:srgbClr val=%q/></a:solidFill>%s</a:ln>`, int(w*EMUPerPt), hexColor(hex), arrows.String())
}

func runXML(r TextRun) string {
	var props strings.Builder
	if r.SizePt > 0 {
		fmt.Fprintf(&props, ` sz="%d"`, int(r.SizePt*100))
	}
	if r.Bold {
		props.WriteString(` b="1"`)
	}
	if r.Italic {
		props.WriteString(` i="1"`)
	}
	// Child element order follows CT_TextCharacterProperties' schema
	// sequence: fill before latin — PowerPoint/python-pptx both reject rPr
	// children out of order.
	var children strings.Builder
	if r.Color != "" {
		fmt.Fprintf(&children, `<a:solidFill><a:srgbClr val=%q/></a:solidFill>`, hexColor(r.Color))
	}
	if r.Font != "" {
		fmt.Fprintf(&children, `<a:latin typeface=%q/>`, esc(r.Font))
	}
	rPr := ""
	if props.Len() > 0 || children.Len() > 0 {
		if children.Len() == 0 {
			rPr = fmt.Sprintf(`<a:rPr%s/>`, props.String())
		} else {
			rPr = fmt.Sprintf(`<a:rPr%s>%s</a:rPr>`, props.String(), children.String())
		}
	}
	return fmt.Sprintf(`<a:r>%s<a:t>%s</a:t></a:r>`, rPr, esc(r.Text))
}

func paragraphXML(p Paragraph) string {
	var pPr strings.Builder
	pPr.WriteString(`<a:pPr`)
	if p.Align != "" {
		fmt.Fprintf(&pPr, ` algn="%s"`, p.Align)
	}
	if p.Level > 0 {
		fmt.Fprintf(&pPr, ` lvl="%d"`, p.Level)
	}
	if p.Bullet {
		marL := 228600 * (p.Level + 1)
		fmt.Fprintf(&pPr, ` marL="%d" indent="-228600"`, marL)
	}
	pPr.WriteString(`>`)
	if p.Bullet {
		pPr.WriteString(`<a:buChar char="&#8226;"/>`)
	} else {
		pPr.WriteString(`<a:buNone/>`)
	}
	pPr.WriteString(`</a:pPr>`)

	var runs strings.Builder
	for _, r := range p.Runs {
		runs.WriteString(runXML(r))
	}
	// A paragraph with no runs still needs to be well-formed; a:endParaRPr
	// isn't required for validity so a bare <a:p/> with just pPr is fine.
	if runs.Len() == 0 {
		return fmt.Sprintf(`<a:p>%s</a:p>`, pPr.String())
	}
	return fmt.Sprintf(`<a:p>%s%s</a:p>`, pPr.String(), runs.String())
}

// textBodyXML builds p:txBody. Returns "" when the shape carries no text at
// all and isn't a placeholder (a pure decorative shape needs no txBody).
func textBodyXML(s Shape) string {
	if len(s.Paragraphs) == 0 && s.Placeholder == "" {
		return ""
	}
	paras := s.Paragraphs
	if len(paras) == 0 {
		paras = []Paragraph{{}} // an empty placeholder still needs one paragraph
	}
	var b strings.Builder
	for _, p := range paras {
		b.WriteString(paragraphXML(p))
	}
	return fmt.Sprintf(`<p:txBody><a:bodyPr wrap="square" anchor="t"/><a:lstStyle/>%s</p:txBody>`, b.String())
}

func xfrmXML(s Shape) string {
	flip := ""
	if s.FlipH {
		flip = ` flipH="1"`
	}
	return fmt.Sprintf(`<a:xfrm%s><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>`, flip, s.X, s.Y, s.Cx, s.Cy)
}

func phXML(ph PlaceholderType) string {
	if ph == "" {
		return ""
	}
	if ph == PlaceholderBody {
		// "body" is the implicit/default placeholder type in OOXML — an
		// explicit type attribute is optional but writing it out plainly
		// keeps this generator's shapes self-describing (no inheritance to
		// resolve to know what a shape is), matching this slice's "always
		// explicit, never dependent on layout inheritance" design.
		return `<p:ph type="body"/>`
	}
	return fmt.Sprintf(`<p:ph type=%q/>`, string(ph))
}

// spXML builds one p:sp: a preset-geometry autoshape, or a text box/
// placeholder (prst="rect" + txBox="1" on the non-placeholder case — the
// same idiom xlsxpatch's shapewrite.go uses for its "text" kind).
func spXML(s Shape, id int) string {
	prst := string(s.Kind)
	cNvSpPr := `<p:cNvSpPr/>`
	if s.Kind == KindTextBox {
		prst = "rect"
		if s.Placeholder == "" {
			cNvSpPr = `<p:cNvSpPr txBox="1"/>`
		} else {
			cNvSpPr = `<p:cNvSpPr/>`
		}
	}
	name := s.Name
	if name == "" {
		name = string(s.Kind)
	}
	return fmt.Sprintf(
		`<p:sp><p:nvSpPr><p:cNvPr id="%d" name=%q/>%s<p:nvPr>%s</p:nvPr></p:nvSpPr>`+
			`<p:spPr>%s<a:prstGeom prst=%q><a:avLst/></a:prstGeom>%s%s</p:spPr>%s</p:sp>`,
		id, esc(name), cNvSpPr, phXML(s.Placeholder),
		xfrmXML(s), prst, fillXML(s.Fill), lineXML(s.Stroke, s.StrokeWidthPt),
		textBodyXML(s),
	)
}

// cxnSpXML builds one p:cxnSp — a straight line connector, matching
// xlsxpatch's xdr:cxnSp "line" kind but in pptx's absolute-xfrm coordinate
// space (no anchor grid to place it against).
func cxnSpXML(s Shape, id int) string {
	name := s.Name
	if name == "" {
		name = "line"
	}
	return fmt.Sprintf(
		`<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="%d" name=%q/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>`+
			`<p:spPr>%s<a:prstGeom prst="line"><a:avLst/></a:prstGeom>%s</p:spPr></p:cxnSp>`,
		id, esc(name), xfrmXML(s), lineXMLArrows(s.Stroke, s.StrokeWidthPt, s.HeadArrow, s.TailArrow),
	)
}

func picXML(s Shape, id int, rid string) (string, error) {
	if (s.ImageContentType != "image/png" && s.ImageContentType != "image/jpeg") || len(s.ImageData) == 0 {
		return "", fmt.Errorf("pptxpatch: image requires non-empty PNG or JPEG payload")
	}
	name := s.Name
	if name == "" {
		name = "image"
	}
	return fmt.Sprintf(`<p:pic><p:nvPicPr><p:cNvPr id="%d" name=%q/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed=%q/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>%s<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`, id, esc(name), rid, xfrmXML(s)), nil
}

func tableXML(s Shape, id int) (string, error) {
	if s.Table == nil || !s.Table.Valid() {
		return "", fmt.Errorf("pptxpatch: invalid table")
	}
	if s.Kind != "" || s.Placeholder != "" || len(s.Paragraphs) != 0 || s.Fill != "" || s.Stroke != "" || s.StrokeWidthPt != 0 || s.FlipH || s.HeadArrow || s.TailArrow || s.Cx <= 0 || s.Cy <= 0 {
		return "", fmt.Errorf("pptxpatch: table shape has unsupported non-table fields")
	}
	if s.Anim != (ShapeAnimation{}) {
		return "", fmt.Errorf("pptxpatch: table shape animations are unsupported")
	}
	if sumInts(s.Table.Columns) != s.Cx {
		return "", fmt.Errorf("pptxpatch: table column widths do not match frame width")
	}
	heights := s.Table.RowHeights
	if len(heights) == 0 {
		heights = evenTableHeights(s.Cy, len(s.Table.Rows))
	}
	if sumInts(heights) != s.Cy {
		return "", fmt.Errorf("pptxpatch: table row heights do not match frame height")
	}
	name := s.Name
	if name == "" {
		name = "Table"
	}
	var grid, rows strings.Builder
	for _, width := range s.Table.Columns {
		fmt.Fprintf(&grid, `<a:gridCol w="%d"/>`, width)
	}
	for rowIndex, row := range s.Table.Rows {
		fmt.Fprintf(&rows, `<a:tr h="%d">`, heights[rowIndex])
		for _, cell := range row {
			rows.WriteString(tableCellXML(cell))
		}
		rows.WriteString(`</a:tr>`)
	}
	return fmt.Sprintf(`<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="%d" name=%q/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid>%s</a:tblGrid>%s</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`, id, esc(name), s.X, s.Y, s.Cx, s.Cy, grid.String(), rows.String()), nil
}

func evenTableHeights(total, rows int) []int {
	heights := make([]int, rows)
	base, remainder := total/rows, total%rows
	for i := range heights {
		heights[i] = base
		if i < remainder {
			heights[i]++
		}
	}
	return heights
}
func tableCellXML(cell TableCell) string {
	var paragraphs strings.Builder
	for _, text := range strings.Split(cell.Text, "\n") {
		paragraphs.WriteString(tableCellParagraphXML(text, cell.Align))
	}
	return `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>` + paragraphs.String() + `</a:txBody><a:tcPr>` + fillXML(cell.Fill) + tableBordersXML(cell.Border) + `</a:tcPr></a:tc>`
}
func tableCellParagraphXML(text string, align TextAlign) string {
	if align == "" {
		align = AlignLeft
	}
	space := ""
	if strings.TrimSpace(text) != text {
		space = ` xml:space="preserve"`
	}
	return fmt.Sprintf(`<a:p><a:pPr algn=%q><a:buNone/></a:pPr><a:r><a:t%s>%s</a:t></a:r></a:p>`, align, space, esc(text))
}
func tableBordersXML(border TableBorder) string {
	var b strings.Builder
	for _, edge := range []string{"lnL", "lnR", "lnT", "lnB"} {
		if border.Color == "" && border.WidthPt == 0 {
			fmt.Fprintf(&b, `<a:%s><a:noFill/></a:%s>`, edge, edge)
		} else {
			fmt.Fprintf(&b, `<a:%s w="%d"><a:solidFill><a:srgbClr val=%q/></a:solidFill></a:%s>`, edge, int(border.WidthPt*EMUPerPt), hexColor(border.Color), edge)
		}
	}
	return b.String()
}

func chartGraphicFrameXML(s Shape, id int) (string, error) {
	if s.Chart == nil || !s.Chart.Valid() {
		return "", fmt.Errorf("pptxpatch: chart frame has no complete local chart graph")
	}
	if s.FlipH || s.X < 0 || s.Y < 0 || s.Cx <= 0 || s.Cy <= 0 {
		return "", fmt.Errorf("pptxpatch: chart frame has unsupported transform")
	}
	name := s.Name
	if name == "" {
		name = "Chart"
	}
	return fmt.Sprintf(
		`<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="%d" name=%q/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>`+
			`<p:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></p:xfrm>`+
			`<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id=%q/></a:graphicData></a:graphic></p:graphicFrame>`,
		id, esc(name), s.X, s.Y, s.Cx, s.Cy, esc(s.Chart.RelationshipID),
	), nil
}

func shapeXML(s Shape, id int, imageRID string) (string, error) {
	if s.Table != nil {
		return tableXML(s, id)
	}
	if !shapeKindSupported(s.Kind) {
		return "", fmt.Errorf("pptxpatch: shape kind %q is not in the supported catalogue", s.Kind)
	}
	if s.Kind == KindLine {
		return cxnSpXML(s, id), nil
	}
	if s.Kind == KindImage {
		return picXML(s, id, imageRID)
	}
	if s.Kind == KindChart {
		return chartGraphicFrameXML(s, id)
	}
	return spXML(s, id), nil
}

func bgXML(background string) string {
	if background == "" {
		return ""
	}
	return fmt.Sprintf(`<p:bg><p:bgPr>%s<a:effectLst/></p:bgPr></p:bg>`, fillXML(background))
}

// shapeID is the p:cNvPr id a shape at this slide-order index gets — 1 is
// reserved for the group shape's own p:cNvPr (see spTreeXML), so shapes
// start at 2. Shared by spTreeXML (which writes the ids) and timingXML
// (which needs to target the same ids via p:spTgt/@spid) so the two can
// never drift apart on numbering.
func shapeID(index int) int { return index + 2 }

func spTreeXML(slide Slide, imageRIDs map[int]string) (string, error) {
	var shapes strings.Builder
	for i, s := range slide.Shapes {
		x, err := shapeXML(s, shapeID(i), imageRIDs[i])
		if err != nil {
			return "", err
		}
		shapes.WriteString(x)
	}
	return `<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
		`<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
		shapes.String() + `</p:spTree>`, nil
}

// transitionDirAttr maps Direction onto OOXML's ST_TransitionSideDirectionType
// (l/r/u/d). "" (DirNone) means "omit the attribute, let the schema's own
// default ('l') apply" — not an error.
func transitionDirAttr(d Direction) (string, error) {
	switch d {
	case DirNone:
		return "", nil
	case DirLeft:
		return "l", nil
	case DirRight:
		return "r", nil
	case DirUp:
		return "u", nil
	case DirDown:
		return "d", nil
	default:
		return "", fmt.Errorf("pptxpatch: unsupported transition direction %q", d)
	}
}

// transitionXML builds a slide's optional p:transition (S11) — schema
// position is cSld, clrMapOvr, transition, timing, extLst (see slideXML),
// so this and timingXML are written as siblings after clrMapOvr. "" (no
// element at all) for TransitionNone, matching this package's "nothing
// extra for a deck that doesn't use the feature" convention.
func transitionXML(t SlideTransition) (string, error) {
	switch t.Type {
	case TransitionNone:
		return "", nil
	case TransitionFade:
		return `<p:transition spd="med"><p:fade/></p:transition>`, nil
	case TransitionPush, TransitionWipe:
		dir, err := transitionDirAttr(t.Direction)
		if err != nil {
			return "", err
		}
		tag := "p:push"
		if t.Type == TransitionWipe {
			tag = "p:wipe"
		}
		if dir == "" {
			return fmt.Sprintf(`<p:transition spd="med"><%s/></p:transition>`, tag), nil
		}
		return fmt.Sprintf(`<p:transition spd="med"><%s dir=%q/></p:transition>`, tag, dir), nil
	default:
		return "", fmt.Errorf("pptxpatch: unsupported slide transition type %q", t.Type)
	}
}

// validateAnim fails closed on any ShapeAnimation this package doesn't know
// how to write, rather than silently dropping it or emitting something a
// consumer can't make sense of.
func validateAnim(a ShapeAnimation) error {
	switch a.Effect {
	case AnimNone, AnimFade:
		return nil
	case AnimFlyIn:
		switch a.Direction {
		case DirNone, DirLeft, DirRight, DirUp, DirDown:
			return nil
		default:
			return fmt.Errorf("pptxpatch: unsupported fly-in direction %q", a.Direction)
		}
	default:
		return fmt.Errorf("pptxpatch: unsupported animation effect %q", a.Effect)
	}
}

// flyInAttrSign picks which normalized position attribute (ppt_x/ppt_y —
// OOXML's predefined "this shape's own left/top as a 0..1 fraction of slide
// width/height" animation targets) and offset sign an AnimFlyIn direction
// animates: "which edge does the shape fly in FROM" — e.g. DirLeft starts
// to the left of (a smaller x than) the shape's final position.
func flyInAttrSign(d Direction) (attr, sign string, err error) {
	switch d {
	case DirNone, DirDown:
		return "ppt_y", "+", nil
	case DirUp:
		return "ppt_y", "-", nil
	case DirLeft:
		return "ppt_x", "-", nil
	case DirRight:
		return "ppt_x", "+", nil
	default:
		return "", "", fmt.Errorf("pptxpatch: unsupported fly-in direction %q", d)
	}
}

// flyInSubtype is PowerPoint's own cosmetic "which Effect Options label/icon
// does the Animation Pane show" hint (presetSubtype) — best-effort common
// values, NOT independently verified against an authoritative enumeration
// (ECMA-376 doesn't publish one for this; PowerPoint's own is
// undocumented). Purely cosmetic: actual playback direction is fully and
// explicitly determined by flyInAttrSign's p:anim values regardless of what
// this returns.
func flyInSubtype(d Direction) string {
	switch d {
	case DirUp:
		return "1"
	case DirRight:
		return "2"
	case DirLeft:
		return "4"
	default: // DirNone, DirDown
		return "8"
	}
}

func formatFraction(f float64) string {
	return strconv.FormatFloat(f, 'f', -1, 64)
}

// timingXML builds a slide's optional p:timing (S11 entrance animations)
// plus its sibling p:bldLst — the latter is what keeps an animated shape
// hidden until its entrance effect actually fires (ECMA-376 §19.5's "build"
// mechanism). Returns "" (no p:timing at all) when no shape on this slide
// carries an animation, matching this package's convention throughout: a
// deck that doesn't use a feature gets byte-identical output to before that
// feature existed.
//
// Scope (S11, deliberately narrow — see AnimEffect's doc comment): two
// entrance effects only. Fade uses p:animEffect (transition="in"
// filter="fade") — a simple, schema-well-documented image-filter
// technique. FlyIn uses a p:anim on ppt_x/ppt_y with a relative
// self-referencing offset formula (e.g. "ppt_y+0.25" -> "ppt_y") — the same
// technique PowerPoint's own generated XML uses for its "Fly In" effect,
// deliberately avoiding this package computing any absolute/EMU position
// itself (the formula is resolved against the shape's OWN final position at
// render time, so it's correct regardless of what that position is).
//
// Every animation triggers automatically when the slide loads (stCondLst
// delay, no onClick condition) rather than waiting for a presenter click —
// the right default for an authored/exported deck rather than a live
// click-through presentation. DelayMs lets a caller stagger multiple
// shapes' entrances relative to each other; they're otherwise independent
// (concurrent) top-level nodes, not chained off one another, so getting one
// shape's stagger "wrong" can't cascade into every later shape's timing.
func timingXML(slide Slide) (string, error) {
	type animatedShape struct {
		spid int
		sh   Shape
	}
	var animated []animatedShape
	for i, s := range slide.Shapes {
		if s.Anim.Effect == AnimNone {
			continue
		}
		if err := validateAnim(s.Anim); err != nil {
			return "", err
		}
		animated = append(animated, animatedShape{spid: shapeID(i), sh: s})
	}
	if len(animated) == 0 {
		return "", nil
	}

	// cTn ids 1 (tmRoot) and 2 (mainSeq) are the fixed wrapper nodes below;
	// everything else counts up from 3. Ids only need to be unique within
	// this slide's own p:timing tree — nothing here cross-references them
	// by id (targets are always by shape spid via p:spTgt).
	nextID := 3
	newID := func() int {
		id := nextID
		nextID++
		return id
	}

	var pars, bld strings.Builder
	for _, a := range animated {
		delay := a.sh.Anim.DelayMs
		if delay < 0 {
			delay = 0
		}
		dur := a.sh.Anim.DurationMs
		if dur <= 0 {
			dur = 500
		}
		parID, setID, effID := newID(), newID(), newID()

		// Every entrance effect starts by revealing the shape (it's hidden
		// by default while it has an unfired build, per p:bldLst below) —
		// the same style.visibility set PowerPoint's own generated XML uses
		// for every entrance effect regardless of what the effect itself
		// animates.
		setXML := fmt.Sprintf(
			`<p:set><p:cBhvr><p:cTn id="%d" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>`+
				`<p:tgtEl><p:spTgt spid="%d"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>`+
				`<p:to><p:strVal val="visible"/></p:to></p:set>`,
			setID, a.spid,
		)

		var presetID, presetSubtype, effectXML string
		switch a.sh.Anim.Effect {
		case AnimFade:
			presetID, presetSubtype = "10", "0"
			effectXML = fmt.Sprintf(
				`<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="%d" dur="%d"/><p:tgtEl><p:spTgt spid="%d"/></p:tgtEl></p:cBhvr></p:animEffect>`,
				effID, dur, a.spid,
			)
		case AnimFlyIn:
			attr, sign, err := flyInAttrSign(a.sh.Anim.Direction)
			if err != nil {
				return "", err
			}
			dist := a.sh.Anim.Distance
			if dist <= 0 {
				dist = 0.25
			}
			presetID, presetSubtype = "2", flyInSubtype(a.sh.Anim.Direction)
			effectXML = fmt.Sprintf(
				`<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="%d" dur="%d" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>`+
					`<p:tgtEl><p:spTgt spid="%d"/></p:tgtEl><p:attrNameLst><p:attrName>%s</p:attrName></p:attrNameLst></p:cBhvr>`+
					`<p:tavLst><p:tav tm="0"><p:val><p:strVal val="%s%s%s"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="%s"/></p:val></p:tav></p:tavLst></p:anim>`,
				effID, dur, a.spid, attr, attr, sign, formatFraction(dist), attr,
			)
		}

		fmt.Fprintf(&pars,
			`<p:par><p:cTn id="%d" presetID="%s" presetClass="entr" presetSubtype="%s" fill="hold" nodeType="clickEffect">`+
				`<p:stCondLst><p:cond delay="%d"/></p:stCondLst><p:childTnLst>%s%s</p:childTnLst></p:cTn></p:par>`,
			parID, presetID, presetSubtype, delay, setXML, effectXML,
		)
		fmt.Fprintf(&bld, `<p:bldP spid="%d" grpId="0"/>`, a.spid)
	}

	return fmt.Sprintf(
		`<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>`+
			`<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>%s</p:childTnLst></p:cTn></p:seq>`+
			`</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst>%s</p:bldLst></p:timing>`,
		pars.String(), bld.String(),
	), nil
}

func slideXML(slide Slide, imageRIDOpts ...map[int]string) (string, error) {
	imageRIDs := map[int]string{}
	if len(imageRIDOpts) > 0 {
		imageRIDs = imageRIDOpts[0]
	}
	tree, err := spTreeXML(slide, imageRIDs)
	if err != nil {
		return "", err
	}
	transition, err := transitionXML(slide.Transition)
	if err != nil {
		return "", err
	}
	timing, err := timingXML(slide)
	if err != nil {
		return "", err
	}
	return xmlDecl +
		fmt.Sprintf(`<p:sld xmlns:a=%q xmlns:c=%q xmlns:r=%q xmlns:p=%q>`, nsA, nsC, nsR, nsP) +
		`<p:cSld>` + bgXML(slide.Background) + tree + `</p:cSld>` +
		`<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>` + transition + timing + `</p:sld>`, nil
}

// relativePartTarget makes an OPC Target relative to the part that owns the
// relationship. It only receives already validated ppt/ names, so no target
// can accidentally escape the presentation package.
func relativePartTarget(fromPart, toPart string) string {
	from := strings.Split(path.Dir(fromPart), "/")
	to := strings.Split(toPart, "/")
	i := 0
	for i < len(from) && i < len(to) && from[i] == to[i] {
		i++
	}
	parts := make([]string, 0, len(from)-i+len(to)-i)
	for j := i; j < len(from); j++ {
		parts = append(parts, "..")
	}
	parts = append(parts, to[i:]...)
	return strings.Join(parts, "/")
}

func chartSlideRelsXML(slide Slide, slidePart, imageRels string, imageCount int) (string, error) {
	seen := map[string]bool{"rId1": true} // slide layout owns rId1
	for i := 0; i < imageCount; i++ {
		seen[fmt.Sprintf("rId%d", i+2)] = true
	}
	var charts strings.Builder
	for _, s := range slide.Shapes {
		if s.Kind != KindChart {
			continue
		}
		if s.Chart == nil || !s.Chart.Valid() || seen[s.Chart.RelationshipID] {
			return "", fmt.Errorf("pptxpatch: duplicate or invalid slide chart relationship")
		}
		seen[s.Chart.RelationshipID] = true
		fmt.Fprintf(&charts, `<Relationship Id=%q Type=%q Target=%q/>`, s.Chart.RelationshipID, relTypeChart, relativePartTarget(slidePart, s.Chart.PartName))
	}
	return xmlDecl +
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
		fmt.Sprintf(`<Relationship Id="rId1" Type=%q Target="../slideLayouts/slideLayout1.xml"/>`, relTypeSlideLayout) +
		imageRels + charts.String() + `</Relationships>`, nil
}

func chartRelsXML(chart ChartPart) string {
	if len(chart.Relationships) == 0 {
		return ""
	}
	var rels strings.Builder
	for _, r := range chart.Relationships {
		fmt.Fprintf(&rels, `<Relationship Id=%q Type=%q Target=%q/>`, r.ID, r.Type, r.Target)
	}
	return xmlDecl + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` + rels.String() + `</Relationships>`
}

// collectChartParts validates and de-duplicates every opaque graph before
// BuildPPTX writes even one byte. Reusing the same chart part on more than
// one slide is safe only when every preserved byte and content type agrees.
func collectChartParts(deck Deck) (map[string]string, map[string]string, error) {
	parts := map[string]string{}
	contentTypes := map[string]string{}
	put := func(name, value, contentType string) error {
		if prev, ok := parts[name]; ok && prev != value {
			return fmt.Errorf("pptxpatch: conflicting opaque chart part %q", name)
		}
		if prev, ok := contentTypes[name]; ok && prev != contentType {
			return fmt.Errorf("pptxpatch: conflicting content type for %q", name)
		}
		parts[name], contentTypes[name] = value, contentType
		return nil
	}
	for _, slide := range deck.Slides {
		for _, s := range slide.Shapes {
			if s.Kind != KindChart {
				continue
			}
			if s.Chart == nil || !s.Chart.Valid() {
				return nil, nil, fmt.Errorf("pptxpatch: invalid opaque chart graph")
			}
			if err := put(s.Chart.PartName, string(s.Chart.XML), ctChart); err != nil {
				return nil, nil, err
			}
			if rels := chartRelsXML(*s.Chart); rels != "" {
				if err := put(relsPathFor(s.Chart.PartName), rels, "application/vnd.openxmlformats-package.relationships+xml"); err != nil {
					return nil, nil, err
				}
			}
			for _, p := range s.Chart.EmbeddedParts {
				if err := put(p.Name, string(p.Data), p.ContentType); err != nil {
					return nil, nil, err
				}
			}
		}
	}
	return parts, contentTypes, nil
}

const emptySpTree = `<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
	`<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>` +
	`<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree>`

func masterLayoutThemeXML() (master, masterRels, layout, layoutRels, theme string) {
	master = xmlDecl +
		fmt.Sprintf(`<p:sldMaster xmlns:a=%q xmlns:r=%q xmlns:p=%q>`, nsA, nsR, nsP) +
		`<p:cSld>` + emptySpTree + `</p:cSld>` +
		`<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
		`<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>` +
		`<p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>` +
		`</p:sldMaster>`
	masterRels = xmlDecl +
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
		fmt.Sprintf(`<Relationship Id="rId1" Type=%q Target="../slideLayouts/slideLayout1.xml"/>`, relTypeSlideLayout) +
		fmt.Sprintf(`<Relationship Id="rId2" Type=%q Target="../theme/theme1.xml"/>`, relTypeTheme) +
		`</Relationships>`
	layout = xmlDecl +
		fmt.Sprintf(`<p:sldLayout xmlns:a=%q xmlns:r=%q xmlns:p=%q type="blank">`, nsA, nsR, nsP) +
		`<p:cSld name="Blank">` + emptySpTree + `</p:cSld>` +
		`<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`
	layoutRels = xmlDecl +
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
		fmt.Sprintf(`<Relationship Id="rId1" Type=%q Target="../slideMasters/slideMaster1.xml"/>`, relTypeSlideMaster) +
		`</Relationships>`
	// Office default theme palette — well-known ECMA-376 DrawingML scheme
	// values (system colors plus the common Office sRGB accents).
	theme = xmlDecl +
		fmt.Sprintf(`<a:theme xmlns:a=%q name="InjOffice">`, nsA) +
		`<a:themeElements><a:clrScheme name="InjOffice">` +
		`<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>` +
		`<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>` +
		`<a:dk2><a:srgbClr val="1D2427"/></a:dk2>` +
		`<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>` +
		`<a:accent1><a:srgbClr val="2F6FED"/></a:accent1>` +
		`<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>` +
		`<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>` +
		`<a:accent4><a:srgbClr val="FFC000"/></a:accent4>` +
		`<a:accent5><a:srgbClr val="4472C4"/></a:accent5>` +
		`<a:accent6><a:srgbClr val="70AD47"/></a:accent6>` +
		`<a:hlink><a:srgbClr val="0563C1"/></a:hlink>` +
		`<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>` +
		`</a:clrScheme>` +
		`<a:fontScheme name="InjOffice">` +
		`<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>` +
		`<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>` +
		`</a:fontScheme>` +
		`<a:fmtScheme name="InjOffice">` +
		`<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>` +
		`<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>` +
		`<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>` +
		`<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>` +
		`</a:fmtScheme></a:themeElements></a:theme>`
	return
}

func docPropsXML(slideCount int) (core, app string) {
	core = xmlDecl +
		`<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
		`xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ` +
		`xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>InjOffice</dc:creator></cp:coreProperties>`
	app = xmlDecl +
		`<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ` +
		`xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
		`<Application>InjOffice</Application><Slides>` + strconv.Itoa(slideCount) + `</Slides></Properties>`
	return
}

// BuildPPTX serializes a Deck to real .pptx bytes: a complete, independently
// openable OOXML package (not a fragment to merge into something else).
func BuildPPTX(deck Deck) ([]byte, error) {
	cx, cy := deck.Cx, deck.Cy
	if cx <= 0 {
		cx = DefaultSlideCx
	}
	if cy <= 0 {
		cy = DefaultSlideCy
	}
	if len(deck.Slides) == 0 {
		return nil, fmt.Errorf("pptxpatch: deck has no slides")
	}
	chartParts, chartContentTypes, err := collectChartParts(deck)
	if err != nil {
		return nil, err
	}

	slideXMLs := make([]string, len(deck.Slides))
	slideImageRels := make([]string, len(deck.Slides))
	slideImageCounts := make([]int, len(deck.Slides))
	mediaEntries := map[string]string{}
	for i, s := range deck.Slides {
		rids := map[int]string{}
		var rels strings.Builder
		imageN := 0
		for j, shape := range s.Shapes {
			if shape.Kind != KindImage {
				continue
			}
			imageN++
			rid := fmt.Sprintf("rId%d", imageN+1)
			rids[j] = rid
			ext := "png"
			if shape.ImageContentType == "image/jpeg" {
				ext = "jpeg"
			}
			mediaName := fmt.Sprintf("ppt/media/slide%d-image%d.%s", i+1, imageN, ext)
			mediaEntries[mediaName] = string(shape.ImageData)
			rels.WriteString(fmt.Sprintf(`<Relationship Id=%q Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target=%q/>`, rid, "../media/"+mediaName[len("ppt/media/"):]))
		}
		slideImageRels[i] = rels.String()
		slideImageCounts[i] = imageN
		x, err := slideXML(s, rids)
		if err != nil {
			return nil, fmt.Errorf("pptxpatch: slide %d: %w", i+1, err)
		}
		slideXMLs[i] = x
	}

	master, masterRels, layout, layoutRels, theme := masterLayoutThemeXML()
	core, app := docPropsXML(len(deck.Slides))

	var sldIdLst, presRels strings.Builder
	presRels.WriteString(fmt.Sprintf(`<Relationship Id="rId1" Type=%q Target="slideMasters/slideMaster1.xml"/>`, relTypeSlideMaster))
	for i := range deck.Slides {
		sldID := 256 + i
		rID := i + 2 // rId1 is the master
		fmt.Fprintf(&sldIdLst, `<p:sldId id="%d" r:id="rId%d"/>`, sldID, rID)
		fmt.Fprintf(&presRels, `<Relationship Id="rId%d" Type=%q Target="slides/slide%d.xml"/>`, rID, relTypeSlide, i+1)
	}

	presentation := xmlDecl +
		fmt.Sprintf(`<p:presentation xmlns:a=%q xmlns:r=%q xmlns:p=%q>`, nsA, nsR, nsP) +
		`<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
		`<p:sldIdLst>` + sldIdLst.String() + `</p:sldIdLst>` +
		fmt.Sprintf(`<p:sldSz cx="%d" cy="%d"/><p:notesSz cx="6858000" cy="9144000"/>`, cx, cy) +
		`</p:presentation>`
	presentationRels := xmlDecl +
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` + presRels.String() + `</Relationships>`

	rootRels := xmlDecl +
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
		fmt.Sprintf(`<Relationship Id="rId1" Type=%q Target="ppt/presentation.xml"/>`, relTypeOfficeDocument) +
		fmt.Sprintf(`<Relationship Id="rId2" Type=%q Target="docProps/core.xml"/>`, relTypeCoreProps) +
		fmt.Sprintf(`<Relationship Id="rId3" Type=%q Target="docProps/app.xml"/>`, relTypeExtProps) +
		`</Relationships>`

	var ctOverrides strings.Builder
	fmt.Fprintf(&ctOverrides, `<Override PartName="/ppt/presentation.xml" ContentType=%q/>`, ctPresentation)
	fmt.Fprintf(&ctOverrides, `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType=%q/>`, ctSlideMaster)
	fmt.Fprintf(&ctOverrides, `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType=%q/>`, ctSlideLayout)
	fmt.Fprintf(&ctOverrides, `<Override PartName="/ppt/theme/theme1.xml" ContentType=%q/>`, ctTheme)
	fmt.Fprintf(&ctOverrides, `<Override PartName="/docProps/core.xml" ContentType=%q/>`, ctCoreProps)
	fmt.Fprintf(&ctOverrides, `<Override PartName="/docProps/app.xml" ContentType=%q/>`, ctExtProps)
	for i := range deck.Slides {
		fmt.Fprintf(&ctOverrides, `<Override PartName="/ppt/slides/slide%d.xml" ContentType=%q/>`, i+1, ctSlide)
	}
	for name, contentType := range chartContentTypes {
		// Relationship parts use the package-wide .rels default; emitting an
		// Override for them is redundant and some Office consumers dislike it.
		if strings.HasSuffix(name, ".rels") {
			continue
		}
		fmt.Fprintf(&ctOverrides, `<Override PartName=%q ContentType=%q/>`, "/"+name, contentType)
	}
	var imageDefaults strings.Builder
	for _, imageType := range []struct{ extension, contentType string }{{"png", "image/png"}, {"jpeg", "image/jpeg"}} {
		used := false
		for name := range mediaEntries {
			if strings.HasSuffix(name, "."+imageType.extension) {
				used = true
				break
			}
		}
		if used {
			fmt.Fprintf(&imageDefaults, `<Default Extension=%q ContentType=%q/>`, imageType.extension, imageType.contentType)
		}
	}
	contentTypes := xmlDecl +
		`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
		`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
		`<Default Extension="xml" ContentType="application/xml"/>` +
		imageDefaults.String() +
		ctOverrides.String() + `</Types>`

	entries := map[string]string{
		"[Content_Types].xml":                          contentTypes,
		"_rels/.rels":                                  rootRels,
		"docProps/core.xml":                            core,
		"docProps/app.xml":                             app,
		"ppt/presentation.xml":                         presentation,
		"ppt/_rels/presentation.xml.rels":              presentationRels,
		"ppt/slideMasters/slideMaster1.xml":            master,
		"ppt/slideMasters/_rels/slideMaster1.xml.rels": masterRels,
		"ppt/slideLayouts/slideLayout1.xml":            layout,
		"ppt/slideLayouts/_rels/slideLayout1.xml.rels": layoutRels,
		"ppt/theme/theme1.xml":                         theme,
	}
	for i, sx := range slideXMLs {
		entries[fmt.Sprintf("ppt/slides/slide%d.xml", i+1)] = sx
		rels, err := chartSlideRelsXML(deck.Slides[i], fmt.Sprintf("ppt/slides/slide%d.xml", i+1), slideImageRels[i], slideImageCounts[i])
		if err != nil {
			return nil, fmt.Errorf("pptxpatch: slide %d relationships: %w", i+1, err)
		}
		entries[fmt.Sprintf("ppt/slides/_rels/slide%d.xml.rels", i+1)] = rels
	}
	for name, data := range mediaEntries {
		if _, conflicts := chartParts[name]; conflicts {
			return nil, fmt.Errorf("pptxpatch: image media collides with opaque chart part %q", name)
		}
		entries[name] = data
	}
	for name, value := range chartParts {
		entries[name] = value
	}

	return zipFiles(entries)
}

// zipFiles writes a deterministic-order zip (sorted names) so BuildPPTX's
// output is byte-reproducible for identical input — useful for diffing in
// tests, not load-bearing for OOXML validity itself.
func zipFiles(entries map[string]string) ([]byte, error) {
	names := make([]string, 0, len(entries))
	for n := range entries {
		names = append(names, n)
	}
	for i := 0; i < len(names); i++ {
		for j := i + 1; j < len(names); j++ {
			if names[j] < names[i] {
				names[i], names[j] = names[j], names[i]
			}
		}
	}
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	for _, n := range names {
		w, err := zw.Create(n)
		if err != nil {
			return nil, fmt.Errorf("pptxpatch: zip create %s: %w", n, err)
		}
		if _, err := w.Write([]byte(entries[n])); err != nil {
			return nil, fmt.Errorf("pptxpatch: zip write %s: %w", n, err)
		}
	}
	if err := zw.Close(); err != nil {
		return nil, fmt.Errorf("pptxpatch: zip close: %w", err)
	}
	return buf.Bytes(), nil
}
