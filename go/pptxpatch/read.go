package pptxpatch

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"slices"
	"strconv"
	"strings"
)

// Slide READING: parse a real .pptx's slides back into the Deck model. Same
// token-based walk style as xlsxpatch's shaperead.go (a StartElement/
// EndElement/CharData state machine over encoding/xml, keyed on element
// LOCAL names so prefix choice in the source file — p:/a: is universal but
// nothing enforces it — never matters).
//
// Deliberately permissive on read, unlike the strict/curated write side:
// an unsupported prst is still surfaced as itself (mirrors shaperead.go's
// "an unknown kind stays visible, never silently dropped" rule) even though
// BuildPPTX wouldn't accept every possible prst back. Unrotated, unflipped
// p:grpSp trees are flattened into their supported p:sp/p:cxnSp children.
// A group with rotation/flips, a malformed transform, or an unsupported
// child is deliberately skipped rather than claiming false visual fidelity.
// A strict, plain p:graphicFrame/a:tbl subset is also surfaced: tables with
// merges, inherited/themed styling, or a transform this model cannot write
// back are omitted rather than approximated; pictures and charts remain
// unsupported here.
//
// Placeholder GEOMETRY and shape FILL/STROKE inherited from the slide
// layout/master (via <p:ph> position inheritance and <p:style>'s
// fillRef/lnRef + theme scheme colors) ARE resolved — this was NOT true in
// an earlier revision of this reader, and the gap was serious enough to be
// worth its own note: a real-world .pptx's title/body placeholders almost
// never carry their own <a:xfrm> at all (PowerPoint and python-pptx both
// leave position/size to the layout, which itself commonly leaves it to the
// MASTER), and a shape added without manually recoloring it (the common
// case — e.g. python-pptx's own add_shape default) gets its fill/line
// ENTIRELY through <p:style>'s theme references, not an explicit
// <a:solidFill>. Without resolving either, ParsePPTX read real geometry and
// colors as all-zero/empty for almost any real-world file — not
// approximately wrong, literally invisible on a canvas renderer (a
// zero-size shape paints nothing; an unset fill+stroke paints nothing
// either). Found live on staging (round 6 of the Slides OOXML pivot): a
// python-pptx-built title+bullets+one default-styled autoshape file parsed
// "successfully" (200, real shapes in the response) but rendered as a fully
// blank white canvas — confirmed by direct canvas pixel inspection (0 of
// 2,764,800 pixels non-white). Root-caused by reproducing the identical
// fixture locally and diffing its raw slide/layout/master XML: the
// placeholders' <p:spPr/> was empty at BOTH the slide and layout level
// (geometry lives on the MASTER), and the autoshape's fill/line came only
// from <p:style><a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef>
// — no <a:solidFill> anywhere on the shape itself. See resolveInheritance's
// own doc comment for exactly what is and isn't resolved.
//
// Still NOT resolved, an honest remaining gap: a run's TEXT color when it's
// only implied by <p:style>'s <a:fontRef> (not an explicit <a:solidFill>
// inside <a:rPr>) — falls back to the existing '#000000' default at the
// renderer, which stays legible on most fills but isn't necessarily the
// exact color PowerPoint would pick. And the resolved fillRef/lnRef color
// is the theme's RAW accent color, not the specific tint/shade PowerPoint's
// format-scheme fillStyleLst/lnStyleLst would actually apply at that
// index — a reasonable, visibly-correct-enough approximation, not a claim
// of exact color fidelity.

// readZipFile returns one part's raw bytes by exact path.
func readZipFile(zr *zip.Reader, name string) (string, bool) {
	for _, f := range zr.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return "", false
			}
			defer rc.Close() //nolint:errcheck
			b, err := io.ReadAll(rc)
			if err != nil {
				return "", false
			}
			return string(b), true
		}
	}
	return "", false
}

func attrValNS(el xml.StartElement, space, local string) string {
	for _, a := range el.Attr {
		if a.Name.Local == local && (space == "" || a.Name.Space == space) {
			return a.Value
		}
	}
	return ""
}

func attrVal(el xml.StartElement, local string) string {
	return attrValNS(el, "", local)
}

// dirOf returns a zip part's own directory ("" for a part at the archive
// root) — the baseDir joinPart needs to resolve a relationship Target found
// in that part's OWN .rels file.
func dirOf(part string) string {
	if idx := strings.LastIndex(part, "/"); idx >= 0 {
		return part[:idx]
	}
	return ""
}

// relsPathFor returns the .rels part for a given zip part, e.g.
// "ppt/slides/slide1.xml" -> "ppt/slides/_rels/slide1.xml.rels" — the
// standard OPC convention (a "_rels" sibling directory, same dir as the
// part itself).
func relsPathFor(part string) string {
	dir := dirOf(part)
	base := part
	if idx := strings.LastIndex(part, "/"); idx >= 0 {
		base = part[idx+1:]
	}
	if dir == "" {
		return "_rels/" + base + ".rels"
	}
	return dir + "/_rels/" + base + ".rels"
}

// joinPart resolves a relationship Target against the part it was declared
// in's directory (baseDir), handling the "../" segments real OOXML rels use.
func joinPart(baseDir, target string) string {
	for strings.HasPrefix(target, "../") {
		target = target[len("../"):]
		if idx := strings.LastIndex(baseDir, "/"); idx >= 0 {
			baseDir = baseDir[:idx]
		} else {
			baseDir = ""
		}
	}
	if baseDir == "" {
		return target
	}
	return baseDir + "/" + target
}

// parseRels returns id -> target for one .rels part's <Relationship> entries.
func parseRels(relsXML string) (map[string]string, error) {
	out := map[string]string{}
	dec := xml.NewDecoder(strings.NewReader(relsXML))
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("pptxpatch: parse rels: %w", err)
		}
		if se, ok := tok.(xml.StartElement); ok && se.Name.Local == "Relationship" {
			id := attrVal(se, "Id")
			target := attrVal(se, "Target")
			if id != "" {
				out[id] = target
			}
		}
	}
	return out, nil
}

type opcRelationship struct {
	ID, Type, Target, TargetMode string
}

// parseOPCRels keeps the fields opaque chart preservation needs. The
// presentation reader only needs id -> target; chart preservation must also
// prove that no relationship is external before carrying it forward.
func parseOPCRels(relsXML string) ([]opcRelationship, error) {
	dec := xml.NewDecoder(strings.NewReader(relsXML))
	var out []opcRelationship
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			return out, nil
		}
		if err != nil {
			return nil, fmt.Errorf("pptxpatch: parse rels: %w", err)
		}
		se, ok := tok.(xml.StartElement)
		if !ok || se.Name.Local != "Relationship" {
			continue
		}
		out = append(out, opcRelationship{
			ID: attrVal(se, "Id"), Type: attrVal(se, "Type"),
			Target: attrVal(se, "Target"), TargetMode: attrVal(se, "TargetMode"),
		})
	}
}

type contentTypeIndex struct {
	defaults  map[string]string
	overrides map[string]string
}

func parseContentTypes(raw string) (contentTypeIndex, error) {
	out := contentTypeIndex{defaults: map[string]string{}, overrides: map[string]string{}}
	dec := xml.NewDecoder(strings.NewReader(raw))
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			return out, nil
		}
		if err != nil {
			return out, fmt.Errorf("pptxpatch: parse content types: %w", err)
		}
		se, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		switch se.Name.Local {
		case "Default":
			if ext, ct := attrVal(se, "Extension"), attrVal(se, "ContentType"); ext != "" && ct != "" {
				out.defaults[strings.ToLower(ext)] = ct
			}
		case "Override":
			if part, ct := attrVal(se, "PartName"), attrVal(se, "ContentType"); strings.HasPrefix(part, "/") && ct != "" {
				out.overrides[strings.TrimPrefix(part, "/")] = ct
			}
		}
	}
}

func (c contentTypeIndex) forPart(part string) string {
	if ct := c.overrides[part]; ct != "" {
		return ct
	}
	if dot := strings.LastIndex(part, "."); dot >= 0 {
		return c.defaults[strings.ToLower(part[dot+1:])]
	}
	return ""
}

const relTypeChart = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"

// chartPartsForSlide constructs a lossless-but-bounded lookup for the chart
// graphicFrames in one slide. Every immediate chart relationship target must
// be local, present, content-typed and terminal (no unmodelled child .rels).
// Any broken graph is omitted, so ParsePPTX never returns a chart it could not
// later write without loss.
func chartPartsForSlide(zr *zip.Reader, slidePart string, types contentTypeIndex) map[string]ChartPart {
	out := map[string]ChartPart{}
	relsRaw, ok := readZipFile(zr, relsPathFor(slidePart))
	if !ok {
		return out
	}
	slideRels, err := parseOPCRels(relsRaw)
	if err != nil {
		return out
	}
	for _, sr := range slideRels {
		if sr.Type != relTypeChart || sr.ID == "" || sr.TargetMode != "" {
			continue
		}
		chartName, ok := resolveChartPart(slidePart, sr.Target)
		if !ok || types.forPart(chartName) != "application/vnd.openxmlformats-officedocument.drawingml.chart+xml" {
			continue
		}
		chartXML, ok := readZipFile(zr, chartName)
		if !ok {
			continue
		}
		chart := ChartPart{RelationshipID: sr.ID, PartName: chartName, XML: []byte(chartXML)}
		chartRelsRaw, hasRels := readZipFile(zr, relsPathFor(chartName))
		if hasRels {
			chartRels, err := parseOPCRels(chartRelsRaw)
			if err != nil {
				continue
			}
			valid := true
			for _, cr := range chartRels {
				if cr.ID == "" || cr.Type == "" || cr.TargetMode != "" {
					valid = false
					break
				}
				partName, ok := resolveChartPart(chartName, cr.Target)
				if !ok || types.forPart(partName) == "" {
					valid = false
					break
				}
				data, ok := readZipFile(zr, partName)
				if !ok {
					valid = false
					break
				}
				// A relationship part on an immediate dependency means this
				// would be a deeper graph than the model represents.
				if _, nested := readZipFile(zr, relsPathFor(partName)); nested {
					valid = false
					break
				}
				chart.Relationships = append(chart.Relationships, ChartRelationship{ID: cr.ID, Type: cr.Type, Target: cr.Target})
				chart.EmbeddedParts = append(chart.EmbeddedParts, ChartEmbeddedPart{Name: partName, Data: []byte(data), ContentType: types.forPart(partName)})
			}
			if !valid {
				continue
			}
		}
		if chart.Valid() {
			out[sr.ID] = chart
		}
	}
	return out
}

// findRelTargetByTypeSuffix scans one .rels part for a <Relationship> whose
// Type ends in the given suffix (e.g. "/slideLayout", "/slideMaster",
// "/theme") and returns its Target — unlike parseRels (keyed by Id, for
// callers that already know which Id they want, like presentation.xml's
// sldIdLst), this is for callers that only know what KIND of part they're
// looking for.
func findRelTargetByTypeSuffix(relsXML, typeSuffix string) (string, bool) {
	dec := xml.NewDecoder(strings.NewReader(relsXML))
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return "", false
		}
		if se, ok := tok.(xml.StartElement); ok && se.Name.Local == "Relationship" {
			if strings.HasSuffix(attrVal(se, "Type"), typeSuffix) {
				return attrVal(se, "Target"), true
			}
		}
	}
	return "", false
}

type presentationInfo struct {
	Cx, Cy    int
	SlideRIDs []string // in <p:sldIdLst> order
}

func parsePresentation(presXML string) (presentationInfo, error) {
	var info presentationInfo
	dec := xml.NewDecoder(strings.NewReader(presXML))
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return info, fmt.Errorf("pptxpatch: parse presentation.xml: %w", err)
		}
		se, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		switch se.Name.Local {
		case "sldSz":
			if v := attrVal(se, "cx"); v != "" {
				info.Cx, _ = strconv.Atoi(v)
			}
			if v := attrVal(se, "cy"); v != "" {
				info.Cy, _ = strconv.Atoi(v)
			}
		case "sldId":
			if rid := attrValNS(se, nsR, "id"); rid != "" {
				info.SlideRIDs = append(info.SlideRIDs, rid)
			}
		}
	}
	return info, nil
}

// ---- Layout/master inheritance: placeholder geometry + theme colors ----
// See this file's top doc comment for WHY this exists. Resolved once per
// unique slide (cached by slide part path in ParsePPTX, since many slides
// commonly share one layout/master) and handed to parseSlide as an optional
// (nilable) lookup — a slide with no matching layout/master (a malformed or
// unusual file) just resolves nothing, the same "degrade, don't fail"
// tolerance the rest of this reader uses.

type phKey struct {
	Type PlaceholderType
	Idx  string // "" when the placeholder (e.g. title) carries no idx at all
}

type phGeom struct{ X, Y, Cx, Cy int }

// parsePlaceholderGeoms walks a slideLayout or slideMaster's OWN shape tree
// (structurally identical to a slide's <p:cSld><p:spTree> for this
// purpose — the root element name (p:sldLayout/p:sldMaster vs p:sld)
// doesn't matter to a token walk that only looks at descendant elements) and
// extracts each placeholder's (type, idx) identity plus its OWN xfrm —
// zero/absent if that layout/master shape ALSO leaves geometry to whatever
// is above IT in the chain (a layout's placeholder frequently has none
// itself; ParsePPTX climbs slide -> layout -> master and keeps the first
// non-zero one it finds).
func parsePlaceholderGeoms(xmlStr string) map[phKey]phGeom {
	out := map[phKey]phGeom{}
	dec := xml.NewDecoder(strings.NewReader(xmlStr))
	var stack []string
	groupDepth := 0
	inside := func(name string) bool {
		return slices.Contains(stack, name)
	}

	inSp := false
	inBg := false
	var curType PlaceholderType
	var curIdx string
	var curX, curY, curCx, curCy int

	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return out
		}
		switch t := tok.(type) {
		case xml.StartElement:
			local := t.Name.Local
			stack = append(stack, local)
			switch local {
			case "grpSp":
				groupDepth++
			case "bg":
				inBg = true
			case "sp":
				if groupDepth == 0 {
					inSp = true
					curType, curIdx = "", ""
					curX, curY, curCx, curCy = 0, 0, 0, 0
				}
			case "ph":
				if inSp {
					pt := attrVal(t, "type")
					if pt == "" {
						pt = "body"
					}
					curType = PlaceholderType(pt)
					curIdx = attrVal(t, "idx")
				}
			case "off":
				if inSp && inside("spPr") && !inBg {
					if v := attrVal(t, "x"); v != "" {
						curX, _ = strconv.Atoi(v)
					}
					if v := attrVal(t, "y"); v != "" {
						curY, _ = strconv.Atoi(v)
					}
				}
			case "ext":
				if inSp && inside("spPr") && !inBg {
					if v := attrVal(t, "cx"); v != "" {
						curCx, _ = strconv.Atoi(v)
					}
					if v := attrVal(t, "cy"); v != "" {
						curCy, _ = strconv.Atoi(v)
					}
				}
			}
		case xml.EndElement:
			local := t.Name.Local
			switch local {
			case "grpSp":
				if groupDepth > 0 {
					groupDepth--
				}
			case "bg":
				inBg = false
			case "sp":
				if groupDepth == 0 && inSp {
					if curType != "" {
						out[phKey{curType, curIdx}] = phGeom{curX, curY, curCx, curCy}
					}
					inSp = false
				}
			}
			if len(stack) > 0 {
				stack = stack[:len(stack)-1]
			}
		}
	}
	return out
}

// themeSlots are a:clrScheme's 12 named color slots, in OOXML's own order.
var themeSlots = map[string]bool{
	"dk1": true, "lt1": true, "dk2": true, "lt2": true,
	"accent1": true, "accent2": true, "accent3": true,
	"accent4": true, "accent5": true, "accent6": true,
	"hlink": true, "folHlink": true,
}

// parseThemeColors reads a theme.xml's <a:clrScheme> into slot -> #rrggbb,
// honoring both explicit <a:srgbClr val=".."/> and <a:sysClr .. lastClr=".."/>
// (Office's default theme uses sysClr for dk1/lt1 — the resolved system
// window/text colors, with lastClr as the frozen fallback everything besides
// live rendering in an actual Office app uses).
func parseThemeColors(themeXML string) map[string]string {
	out := map[string]string{}
	dec := xml.NewDecoder(strings.NewReader(themeXML))
	inScheme := false
	curSlot := ""
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return out
		}
		switch t := tok.(type) {
		case xml.StartElement:
			local := t.Name.Local
			switch {
			case local == "clrScheme":
				inScheme = true
			case inScheme && themeSlots[local]:
				curSlot = local
			case inScheme && curSlot != "" && local == "srgbClr":
				out[curSlot] = "#" + strings.ToLower(attrVal(t, "val"))
			case inScheme && curSlot != "" && local == "sysClr":
				if lc := attrVal(t, "lastClr"); lc != "" {
					out[curSlot] = "#" + strings.ToLower(lc)
				}
			}
		case xml.EndElement:
			local := t.Name.Local
			if local == "clrScheme" {
				inScheme = false
			}
			if themeSlots[local] {
				curSlot = ""
			}
		}
	}
	return out
}

// parseClrMap reads a slide master's <p:clrMap .../> — the logical color
// name (bg1/tx1/bg2/tx2/accent1-6/hlink/folHlink, what a shape's own
// schemeClr actually references) -> theme clrScheme slot name indirection.
// Office's default theme maps these 1:1 (bg1->lt1, tx1->dk1, accentN->accentN,
// ...) but a custom theme can remap them (e.g. a dark master swapping
// bg1<->dk1), so this indirection is never safe to skip.
func parseClrMap(masterXML string) map[string]string {
	out := map[string]string{}
	dec := xml.NewDecoder(strings.NewReader(masterXML))
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return out
		}
		if se, ok := tok.(xml.StartElement); ok && se.Name.Local == "clrMap" {
			for _, a := range se.Attr {
				out[a.Name.Local] = a.Value
			}
			return out
		}
	}
	return out
}

// inheritance bundles everything ParsePPTX resolves ONCE per unique slide
// (cached by slide part path — see ParsePPTX) so parseSlide can fill in
// what a slide leaves to its layout/master: placeholder geometry (climbing
// slide -> layout -> master) and schemeClr/style-ref theme colors. Always
// non-nil from resolveInheritance (a lookup miss at any step just leaves the
// relevant map nil/empty — every method here is nil-map-safe, matching Go's
// own "reading a nil map returns the zero value" behavior), so parseSlide
// never needs to special-case "no inheritance data available" beyond what a
// plain failed map lookup already gives it.
type inheritance struct {
	layoutGeoms map[phKey]phGeom
	masterGeoms map[phKey]phGeom
	clrMap      map[string]string
	themeColors map[string]string
}

// resolveGeom returns the inherited (x, y, cx, cy) for a placeholder,
// preferring the layout's own entry over the master's, and — since a
// placeholder that carries no idx at all (the title) never has one to
// match — falling back to an idx-less entry when a specific idx lookup
// finds nothing. A zero-valued phGeom (both Cx and Cy 0) at any step is
// treated as "not actually set there either" and skipped, not returned as a
// false positive.
func (inh *inheritance) resolveGeom(phType PlaceholderType, idx string) (phGeom, bool) {
	if inh == nil {
		return phGeom{}, false
	}
	try := func(m map[phKey]phGeom, key phKey) (phGeom, bool) {
		g, ok := m[key]
		return g, ok && (g.Cx != 0 || g.Cy != 0)
	}
	if g, ok := try(inh.layoutGeoms, phKey{phType, idx}); ok {
		return g, true
	}
	if g, ok := try(inh.masterGeoms, phKey{phType, idx}); ok {
		return g, true
	}
	if idx != "" {
		if g, ok := try(inh.layoutGeoms, phKey{phType, ""}); ok {
			return g, true
		}
		if g, ok := try(inh.masterGeoms, phKey{phType, ""}); ok {
			return g, true
		}
	}
	return phGeom{}, false
}

// resolveSchemeColor resolves a schemeClr `val` (a logical name like
// "accent1"/"bg1"/"tx1", OR — some real files skip the clrMap indirection
// and reference a theme slot name directly — a slot name itself) into
// #rrggbb via this slide's own clrMap + theme. "phClr" (placeholder color:
// context-dependent, means "whatever color referenced ME") and any other
// unresolvable name return ok=false — the caller leaves the existing
// value (usually "", i.e. unset) rather than guessing.
func (inh *inheritance) resolveSchemeColor(name string) (string, bool) {
	if inh == nil || name == "" {
		return "", false
	}
	slot := name
	if mapped, ok := inh.clrMap[name]; ok {
		slot = mapped
	}
	hex, ok := inh.themeColors[slot]
	return hex, ok
}

// resolveInheritance climbs slide -> layout -> master -> theme via each
// part's own .rels file (the standard OPC relationship chain) and resolves
// everything parseSlide needs to fill in what the slide itself leaves
// unset. Never errors — any missing/malformed step along the way just
// leaves the corresponding map nil, same "degrade, don't fail the read"
// tolerance ParsePPTX already uses for a dangling sldId or a malformed
// slide.
func resolveInheritance(zr *zip.Reader, slidePart string) *inheritance {
	inh := &inheritance{}

	slideRelsXML, ok := readZipFile(zr, relsPathFor(slidePart))
	if !ok {
		return inh
	}
	layoutTarget, ok := findRelTargetByTypeSuffix(slideRelsXML, "/slideLayout")
	if !ok {
		return inh
	}
	layoutPart := joinPart(dirOf(slidePart), layoutTarget)
	if layoutXML, ok := readZipFile(zr, layoutPart); ok {
		inh.layoutGeoms = parsePlaceholderGeoms(layoutXML)
	}

	layoutRelsXML, ok := readZipFile(zr, relsPathFor(layoutPart))
	if !ok {
		return inh
	}
	masterTarget, ok := findRelTargetByTypeSuffix(layoutRelsXML, "/slideMaster")
	if !ok {
		return inh
	}
	masterPart := joinPart(dirOf(layoutPart), masterTarget)
	masterXML, ok := readZipFile(zr, masterPart)
	if !ok {
		return inh
	}
	inh.masterGeoms = parsePlaceholderGeoms(masterXML)
	inh.clrMap = parseClrMap(masterXML)

	masterRelsXML, ok := readZipFile(zr, relsPathFor(masterPart))
	if !ok {
		return inh
	}
	themeTarget, ok := findRelTargetByTypeSuffix(masterRelsXML, "/theme")
	if !ok {
		return inh
	}
	themePart := joinPart(dirOf(masterPart), themeTarget)
	if themeXML, ok := readZipFile(zr, themePart); ok {
		inh.themeColors = parseThemeColors(themeXML)
	}
	return inh
}

// ParsePPTX reads a real .pptx's slides into a Deck. Slide order follows
// presentation.xml's <p:sldIdLst> (the file's own authoritative order, not
// zip entry order).
func ParsePPTX(data []byte) (Deck, error) {
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return Deck{}, fmt.Errorf("pptxpatch: not a valid zip: %w", err)
	}
	presXML, ok := readZipFile(zr, "ppt/presentation.xml")
	if !ok {
		return Deck{}, fmt.Errorf("pptxpatch: missing ppt/presentation.xml")
	}
	contentTypesXML, ok := readZipFile(zr, "[Content_Types].xml")
	if !ok {
		return Deck{}, fmt.Errorf("pptxpatch: missing [Content_Types].xml")
	}
	contentTypes, err := parseContentTypes(contentTypesXML)
	if err != nil {
		return Deck{}, err
	}
	info, err := parsePresentation(presXML)
	if err != nil {
		return Deck{}, err
	}
	relsXML, ok := readZipFile(zr, "ppt/_rels/presentation.xml.rels")
	if !ok {
		return Deck{}, fmt.Errorf("pptxpatch: missing ppt/_rels/presentation.xml.rels")
	}
	rels, err := parseRels(relsXML)
	if err != nil {
		return Deck{}, err
	}

	// Cached by slide part path: many real decks share one layout (or a
	// handful) across every slide, so re-walking the same layout/master/
	// theme XML per-slide would be pure waste on anything but a one-slide
	// file.
	inhCache := map[string]*inheritance{}

	deck := Deck{Cx: info.Cx, Cy: info.Cy}
	for _, rid := range info.SlideRIDs {
		target, ok := rels[rid]
		if !ok {
			continue // dangling sldId — surfaced by omission, same tolerance as shaperead's "skip, don't fail the whole read"
		}
		part := joinPart("ppt", target)
		slideXML, ok := readZipFile(zr, part)
		if !ok {
			continue
		}
		inh, cached := inhCache[part]
		if !cached {
			inh = resolveInheritance(zr, part)
			inhCache[part] = inh
		}
		slide, err := parseSlide(slideXML, inh, slideParseAssets{
			images: readSlideImages(zr, part),
			charts: chartPartsForSlide(zr, part, contentTypes),
		})
		if err != nil {
			continue // one malformed slide doesn't sink the whole deck
		}
		parseSlideEffects(slideXML, &slide)
		deck.Slides = append(deck.Slides, slide)
	}
	return deck, nil
}

// parseSlideEffects reads only the compact S11 vocabulary emitted by this
// package: fade/push/wipe transitions and fade/fly-in entrance effects. It is
// intentionally separate from parseSlide's shape walk: p:timing follows the
// shape tree in OOXML, so every target shape is already present here. Unknown
// timing is ignored rather than guessed at. Our writer assigns shape ids in
// order starting at 2 (shapeID), which is the only target mapping accepted.
func parseSlideEffects(slideXML string, slide *Slide) {
	dec := xml.NewDecoder(strings.NewReader(slideXML))
	var stack []string
	type entrance struct {
		depth int
		anim  ShapeAnimation
		spid  int
	}
	var active *entrance

	for {
		tok, err := dec.Token()
		if err != nil {
			return
		}
		switch t := tok.(type) {
		case xml.StartElement:
			local := t.Name.Local
			stack = append(stack, local)
			switch local {
			case "fade":
				// p:fade occurs in a transition when directly below p:transition.
				if len(stack) >= 2 && stack[len(stack)-2] == "transition" {
					slide.Transition = SlideTransition{Type: TransitionFade}
				}
			case "push", "wipe":
				if len(stack) >= 2 && stack[len(stack)-2] == "transition" {
					typ := TransitionPush
					if local == "wipe" {
						typ = TransitionWipe
					}
					slide.Transition = SlideTransition{Type: typ, Direction: readEffectDirection(attrVal(t, "dir"))}
				}
			case "cTn":
				// The outer timing node is the writer's reliable admission signal:
				// presetClass=entr plus presetID 10 (fade) or 2 (fly-in).
				if attrVal(t, "presetClass") == "entr" {
					var effect AnimEffect
					switch attrVal(t, "presetID") {
					case "10":
						effect = AnimFade
					case "2":
						effect = AnimFlyIn
					}
					if effect != AnimNone {
						direction := Direction("")
						if effect == AnimFlyIn {
							switch attrVal(t, "presetSubtype") {
							case "1":
								direction = DirUp
							case "2":
								direction = DirRight
							case "4":
								direction = DirLeft
							default:
								direction = DirDown
							}
						}
						active = &entrance{depth: len(stack), anim: ShapeAnimation{Effect: effect, Direction: direction}}
					}
				}
			case "spTgt":
				if active != nil && active.spid == 0 {
					active.spid, _ = strconv.Atoi(attrVal(t, "spid"))
				}
			}
		case xml.EndElement:
			if t.Name.Local == "cTn" && active != nil && active.depth == len(stack) {
				idx := active.spid - 2
				if idx >= 0 && idx < len(slide.Shapes) {
					slide.Shapes[idx].Anim = active.anim
				}
				active = nil
			}
			if len(stack) > 0 {
				stack = stack[:len(stack)-1]
			}
		}
	}
}

func readEffectDirection(dir string) Direction {
	switch dir {
	case "r":
		return DirRight
	case "u":
		return DirUp
	case "d":
		return DirDown
	default:
		return DirLeft // OOXML default for omitted push/wipe direction
	}
}

// parseSlide walks one p:sld's XML into a Slide. `inh` (see resolveInheritance)
// fills in placeholder geometry and shape/style theme colors the slide
// itself leaves to its layout/master — nil is safe (every inheritance
// method tolerates it), so a direct/test caller with no inheritance context
// still gets exactly what the slide's own XML states, same as before this
// resolution existed.
// groupTransform is the subset of a:CT_GroupTransform2D which can be
// represented losslessly by this package's flat, axis-aligned Shape model.
// OOXML maps a child coordinate with:
//
//	parentOff + (child - chOff) * ext / chExt
//
// Rotation and either flip cannot be represented by the current Shape
// contract, so their enclosing group is not flattened.
type groupTransform struct {
	offX, offY     int
	extX, extY     int
	chOffX, chOffY int
	chExtX, chExtY int
}

type groupContext struct {
	xfrm     groupTransform
	xfrmSeen bool
	valid    bool
}

func (g groupTransform) apply(s *Shape) bool {
	if g.chExtX <= 0 || g.chExtY <= 0 || g.extX < 0 || g.extY < 0 {
		return false
	}
	// Use int64 for the intermediate product: slide coordinates are normally
	// below 13 million EMU, but producer-generated group transforms can be
	// considerably larger than a single slide.
	x := int64(g.offX) + (int64(s.X-g.chOffX)*int64(g.extX))/int64(g.chExtX)
	y := int64(g.offY) + (int64(s.Y-g.chOffY)*int64(g.extY))/int64(g.chExtY)
	cx := (int64(s.Cx) * int64(g.extX)) / int64(g.chExtX)
	cy := (int64(s.Cy) * int64(g.extY)) / int64(g.chExtY)
	if x < 0 || y < 0 || cx < 0 || cy < 0 ||
		x > int64(^uint(0)>>1) || y > int64(^uint(0)>>1) ||
		cx > int64(^uint(0)>>1) || cy > int64(^uint(0)>>1) {
		return false
	}
	s.X, s.Y, s.Cx, s.Cy = int(x), int(y), int(cx), int(cy)
	return true
}

func readSlideImages(zr *zip.Reader, slidePart string) map[string]Shape {
	out := map[string]Shape{}
	relsXML, ok := readZipFile(zr, relsPathFor(slidePart))
	if !ok {
		return out
	}
	rels, err := parseRels(relsXML)
	if err != nil {
		return out
	}
	for rid, target := range rels {
		part := joinPart(dirOf(slidePart), target)
		raw, ok := readZipFile(zr, part)
		if !ok {
			continue
		}
		data := []byte(raw)
		mime := ""
		if len(data) >= 8 && bytes.Equal(data[:8], []byte{137, 80, 78, 71, 13, 10, 26, 10}) {
			mime = "image/png"
		}
		if len(data) >= 3 && data[0] == 0xff && data[1] == 0xd8 && data[2] == 0xff {
			mime = "image/jpeg"
		}
		if mime != "" {
			out[rid] = Shape{Kind: KindImage, ImageData: data, ImageContentType: mime}
		}
	}
	return out
}

type slideParseAssets struct {
	images map[string]Shape
	charts map[string]ChartPart
}

// parseChartGraphicFrame accepts only the exact flat frame emitted by the
// writer. Tables are consumed by the sibling strict-table parser; all other
// graphicData payloads remain unsupported.  The chart part itself must already
// have passed chartPartsForSlide's complete local-graph validation.
func parseChartGraphicFrame(raw string, charts map[string]ChartPart) (Shape, bool) {
	dec := xml.NewDecoder(strings.NewReader(raw))
	shape := Shape{Kind: KindChart}
	insideFrame := false
	var rid string
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return Shape{}, false
		}
		start, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		switch start.Name.Local {
		case "graphicFrame":
			insideFrame = true
		case "cNvPr":
			if insideFrame && shape.Name == "" {
				shape.Name = attrVal(start, "name")
			}
		case "xfrm":
			if !insideFrame || attrVal(start, "rot") != "" || attrVal(start, "flipH") == "1" || attrVal(start, "flipV") == "1" {
				return Shape{}, false
			}
		case "off":
			if insideFrame {
				shape.X, _ = strconv.Atoi(attrVal(start, "x"))
				shape.Y, _ = strconv.Atoi(attrVal(start, "y"))
			}
		case "ext":
			if insideFrame {
				shape.Cx, _ = strconv.Atoi(attrVal(start, "cx"))
				shape.Cy, _ = strconv.Atoi(attrVal(start, "cy"))
			}
		case "chart":
			rid = attrValNS(start, nsR, "id")
		}
	}
	chart, ok := charts[rid]
	if !ok || !chart.Valid() || chart.RelationshipID != rid || shape.Cx <= 0 || shape.Cy <= 0 {
		return Shape{}, false
	}
	copy := chart
	shape.Chart = &copy
	return shape, true
}

func parseSlide(slideXML string, inh *inheritance, options ...slideParseAssets) (Slide, error) {
	images := map[string]Shape{}
	charts := map[string]ChartPart{}
	if len(options) > 0 {
		if options[0].images != nil {
			images = options[0].images
		}
		if options[0].charts != nil {
			charts = options[0].charts
		}
	}
	dec := xml.NewDecoder(strings.NewReader(slideXML))
	var slide Slide
	var stack []string
	var groups []groupContext

	inside := func(name string) bool {
		return slices.Contains(stack, name)
	}
	top := func() string {
		if len(stack) == 0 {
			return ""
		}
		return stack[len(stack)-1]
	}
	groupsSupported := func() bool {
		for _, g := range groups {
			if !g.valid || !g.xfrmSeen {
				return false
			}
		}
		return true
	}
	flattenGroups := func(s *Shape) bool {
		for i := len(groups) - 1; i >= 0; i-- {
			if !groups[i].xfrm.apply(s) {
				return false
			}
		}
		return true
	}

	var cur *Shape
	curValid := true
	var curPara *Paragraph
	var curRun *TextRun
	inBg := false

	// Per-shape placeholder idx (Shape itself only carries Placeholder TYPE —
	// idx is needed to resolve inheritance correctly when a slide layout has
	// more than one placeholder of the same type, e.g. a two-content layout,
	// so it's tracked here rather than added to the public Shape struct: it's
	// a read-time-only implementation detail, not part of this package's
	// wire/write contract).
	var curIdx string

	// <p:style>'s fillRef/lnRef — the theme-color fallback for a shape with
	// no explicit <a:solidFill>/<a:ln><a:solidFill> of its own (see this
	// file's top doc comment). idx="0" is treated as "no fill/line was
	// actually intended" (a common convention, e.g. python-pptx's own
	// default shapes never emit an idx="0" ref in the first place — they
	// simply omit fillRef/lnRef entirely when there's truly nothing to
	// reference) and skipped rather than resolved to a color.
	inStyle := false
	inFillRef := false
	inLnRef := false
	var styleFillIdx, styleFillClr string
	var styleLnIdx, styleLnClr string

	// applyColor mirrors the priority order the existing srgbClr handling
	// already used (background, then run text, then line stroke, then shape
	// fill) and additionally captures a <p:style> fillRef/lnRef color when
	// inside one of those — shared by both the literal-srgbClr and the
	// theme-schemeClr paths below so they can't drift apart on WHERE a
	// resolved color gets applied.
	applyColor := func(hex string) {
		switch {
		case inBg:
			slide.Background = hex
		case curRun != nil && inside("rPr"):
			curRun.Color = hex
		case inFillRef:
			styleFillClr = hex
		case inLnRef:
			styleLnClr = hex
		case cur != nil && inside("ln"):
			cur.Stroke = hex
		case cur != nil && inside("spPr") && !inBg:
			cur.Fill = hex
		}
	}

	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return slide, fmt.Errorf("pptxpatch: parse slide: %w", err)
		}
		switch t := tok.(type) {
		case xml.StartElement:
			local := t.Name.Local
			stack = append(stack, local)
			switch local {
			case "graphicFrame":
				// A graphicFrame has a substantially different shape tree from
				// p:sp. Consume it as one isolated XML fragment so the normal
				// p:sp token state cannot accidentally treat table text or an
				// a:xfrm as a regular Shape. This retains slide ordering because
				// the table is appended at the exact point the frame occurs.
				raw, err := graphicFrameXML(dec, t)
				if err != nil {
					return slide, fmt.Errorf("pptxpatch: parse graphicFrame: %w", err)
				}
				// A grouped graphicFrame would need the same coordinate-space
				// flattening as p:sp. Table has no group-transform carrier, so
				// skip it rather than retain an untransformed, visibly wrong
				// placement.
				if len(groups) == 0 {
					if tableShape, ok := parseTableGraphicFrame(raw); ok {
						slide.Shapes = append(slide.Shapes, tableShape)
					} else if chartShape, ok := parseChartGraphicFrame(raw, charts); ok {
						slide.Shapes = append(slide.Shapes, chartShape)
					}
				}
				// graphicFrameXML already consumed its EndElement, which the
				// outer loop normally uses to pop this stack entry.
				stack = stack[:len(stack)-1]
				continue
			case "grpSp":
				// A group's own transform is emitted before its children. Start
				// pessimistically: until a valid group xfrm is completely read,
				// no child can be flattened into our absolute Shape model.
				groups = append(groups, groupContext{valid: true})
			case "bg":
				inBg = true
			case "sp":
				if groupsSupported() {
					curValid = true
					// Kind starts unresolved: cNvSpPr (seen next, before
					// spPr/prstGeom in document order) may confirm it's a
					// text box via txBox="1"; otherwise prstGeom's prst
					// resolves it. An unresolved "" surfaces honestly on a
					// malformed/unusual sp rather than guessing.
					cur = &Shape{}
					curIdx = ""
					styleFillIdx, styleFillClr = "", ""
					styleLnIdx, styleLnClr = "", ""
				}
			case "cxnSp":
				if groupsSupported() {
					curValid = true
					cur = &Shape{Kind: KindLine}
					curIdx = ""
					styleFillIdx, styleFillClr = "", ""
					styleLnIdx, styleLnClr = "", ""
				}
			case "pic":
				if groupsSupported() {
					cur = &Shape{Kind: KindImage}
					curValid = true
				}
			case "blip":
				if cur != nil && cur.Kind == KindImage {
					if image, ok := images[attrValNS(t, nsR, "embed")]; ok {
						cur.ImageData, cur.ImageContentType = image.ImageData, image.ImageContentType
					} else {
						curValid = false
					}
				}
			case "srcRect", "effectLst":
				if cur != nil && cur.Kind == KindImage {
					curValid = false
				}
			case "cNvPr":
				// Pre-existing gap closed alongside S12: the shape's display
				// name (p:cNvPr@name) was written by BuildPPTX but never read
				// back — harmless for content/geometry fidelity, but the new
				// diagram round-trip test (and any future caller identifying
				// a specific connector post-parse) needs it.
				if cur != nil && cur.Name == "" {
					cur.Name = attrVal(t, "name")
				}
			case "cNvSpPr":
				if cur != nil && attrVal(t, "txBox") == "1" {
					cur.Kind = KindTextBox
				}
			case "ph":
				if cur != nil {
					pt := attrVal(t, "type")
					if pt == "" {
						pt = "body"
					}
					cur.Placeholder = PlaceholderType(pt)
					curIdx = attrVal(t, "idx")
				}
			case "prstGeom":
				if cur != nil && cur.Kind != KindLine && cur.Kind != KindTextBox && cur.Kind != KindImage && cur.Kind != KindChart {
					// cNvSpPr (seen earlier in document order) already
					// confirmed KindTextBox when txBox="1" was present; any
					// other case (including a genuine prst="rect" with no
					// txBox marker — a real rect autoshape) resolves from
					// prst directly.
					cur.Kind = ShapeKind(attrVal(t, "prst"))
				}
			case "xfrm":
				if inside("grpSpPr") && len(groups) > 0 {
					g := &groups[len(groups)-1]
					g.xfrmSeen = true
					// Rotation is measured in 1/60000 degrees; either flip is
					// equally outside the flat Shape representation. Mark the
					// complete group unsupported instead of approximating it.
					if attrVal(t, "rot") != "" && attrVal(t, "rot") != "0" {
						g.valid = false
					}
					if attrVal(t, "flipH") == "1" || attrVal(t, "flipV") == "1" {
						g.valid = false
					}
				}
				if cur != nil && inside("spPr") && !inBg {
					if cur.Kind == KindImage && ((attrVal(t, "rot") != "" && attrVal(t, "rot") != "0") || attrVal(t, "flipH") == "1" || attrVal(t, "flipV") == "1") {
						curValid = false
					}
					if attrVal(t, "flipH") == "1" {
						cur.FlipH = true
					}
				}
				if cur != nil && cur.Kind == KindChart && inside("graphicFrame") {
					if attrVal(t, "flipH") == "1" || attrVal(t, "flipV") == "1" || attrVal(t, "rot") != "" && attrVal(t, "rot") != "0" {
						// The opaque chart graph is valid, but this flat Shape
						// contract cannot faithfully keep a transformed frame.
						cur = nil
					}
				}
			case "off":
				if inside("grpSpPr") && len(groups) > 0 {
					g := &groups[len(groups)-1].xfrm
					g.offX, _ = strconv.Atoi(attrVal(t, "x"))
					g.offY, _ = strconv.Atoi(attrVal(t, "y"))
				}
				if cur != nil && inside("spPr") && !inBg {
					if v := attrVal(t, "x"); v != "" {
						cur.X, _ = strconv.Atoi(v)
					}
					if v := attrVal(t, "y"); v != "" {
						cur.Y, _ = strconv.Atoi(v)
					}
				}
				if cur != nil && cur.Kind == KindChart && inside("graphicFrame") {
					if v := attrVal(t, "x"); v != "" {
						cur.X, _ = strconv.Atoi(v)
					}
					if v := attrVal(t, "y"); v != "" {
						cur.Y, _ = strconv.Atoi(v)
					}
				}
			case "ext":
				if inside("grpSpPr") && len(groups) > 0 {
					g := &groups[len(groups)-1].xfrm
					g.extX, _ = strconv.Atoi(attrVal(t, "cx"))
					g.extY, _ = strconv.Atoi(attrVal(t, "cy"))
				}
				if cur != nil && inside("spPr") && !inBg {
					if v := attrVal(t, "cx"); v != "" {
						cur.Cx, _ = strconv.Atoi(v)
					}
					if v := attrVal(t, "cy"); v != "" {
						cur.Cy, _ = strconv.Atoi(v)
					}
				}
				if cur != nil && cur.Kind == KindChart && inside("graphicFrame") {
					if v := attrVal(t, "cx"); v != "" {
						cur.Cx, _ = strconv.Atoi(v)
					}
					if v := attrVal(t, "cy"); v != "" {
						cur.Cy, _ = strconv.Atoi(v)
					}
				}
			case "chOff":
				if inside("grpSpPr") && len(groups) > 0 {
					g := &groups[len(groups)-1].xfrm
					g.chOffX, _ = strconv.Atoi(attrVal(t, "x"))
					g.chOffY, _ = strconv.Atoi(attrVal(t, "y"))
				}
			case "chExt":
				if inside("grpSpPr") && len(groups) > 0 {
					g := &groups[len(groups)-1].xfrm
					g.chExtX, _ = strconv.Atoi(attrVal(t, "cx"))
					g.chExtY, _ = strconv.Atoi(attrVal(t, "cy"))
				}
			case "headEnd":
				if cur != nil && inside("ln") && attrVal(t, "type") != "none" {
					cur.HeadArrow = true
				}
			case "tailEnd":
				if cur != nil && inside("ln") && attrVal(t, "type") != "none" {
					cur.TailArrow = true
				}
			case "srgbClr":
				applyColor("#" + strings.ToLower(attrVal(t, "val")))
			case "sysClr":
				if lc := attrVal(t, "lastClr"); lc != "" {
					applyColor("#" + strings.ToLower(lc))
				}
			case "schemeClr":
				if hex, ok := inh.resolveSchemeColor(attrVal(t, "val")); ok {
					applyColor(hex)
				}
			case "style":
				if cur != nil {
					inStyle = true
				}
			case "fillRef":
				if inStyle {
					inFillRef = true
					styleFillIdx = attrVal(t, "idx")
				}
			case "lnRef":
				if inStyle {
					inLnRef = true
					styleLnIdx = attrVal(t, "idx")
				}
			case "ln":
				if cur != nil {
					if w := attrVal(t, "w"); w != "" {
						if n, err := strconv.Atoi(w); err == nil {
							cur.StrokeWidthPt = float64(n) / EMUPerPt
						}
					}
				}
			case "p":
				if inside("txBody") {
					curPara = &Paragraph{}
				}
			case "pPr":
				if curPara != nil {
					if a := attrVal(t, "algn"); a != "" {
						curPara.Align = TextAlign(a)
					}
					if l := attrVal(t, "lvl"); l != "" {
						curPara.Level, _ = strconv.Atoi(l)
					}
				}
			case "buChar":
				if curPara != nil {
					curPara.Bullet = true
				}
			case "buNone":
				if curPara != nil {
					curPara.Bullet = false
				}
			case "r":
				if curPara != nil {
					curRun = &TextRun{}
				}
			case "rPr":
				if curRun != nil {
					if sz := attrVal(t, "sz"); sz != "" {
						if n, err := strconv.Atoi(sz); err == nil {
							curRun.SizePt = float64(n) / 100
						}
					}
					if attrVal(t, "b") == "1" {
						curRun.Bold = true
					}
					if attrVal(t, "i") == "1" {
						curRun.Italic = true
					}
				}
			case "latin":
				if curRun != nil && inside("rPr") {
					if tf := attrVal(t, "typeface"); tf != "" {
						curRun.Font = tf
					}
				}
			}
		case xml.EndElement:
			local := t.Name.Local
			switch local {
			case "xfrm":
				if inside("grpSpPr") && len(groups) > 0 {
					g := &groups[len(groups)-1]
					// ext/chExt are required to map the group coordinate system.
					// Reject zero/negative dimensions rather than divide by zero
					// or silently collapse children.
					if g.xfrm.extX <= 0 || g.xfrm.extY <= 0 ||
						g.xfrm.chExtX <= 0 || g.xfrm.chExtY <= 0 {
						g.valid = false
					}
				}
			case "grpSp":
				if len(groups) > 0 {
					groups = groups[:len(groups)-1]
				}
			case "bg":
				inBg = false
			case "style":
				inStyle = false
			case "fillRef":
				inFillRef = false
			case "lnRef":
				inLnRef = false
			case "r":
				if curPara != nil && curRun != nil {
					curPara.Runs = append(curPara.Runs, *curRun)
					curRun = nil
				}
			case "p":
				if inside("txBody") && curPara != nil {
					cur.Paragraphs = append(cur.Paragraphs, *curPara)
					curPara = nil
				}
			case "sp", "cxnSp", "pic":
				if cur != nil {
					// Placeholder geometry inheritance: a real slide's
					// placeholder almost never carries its own <a:xfrm> (see
					// this file's top doc comment) — climb layout -> master
					// for the same (type, idx) identity. All-zero is this
					// reader's "not actually set" signal throughout (a real
					// shape at literal (0,0) with zero size is not something
					// any producer writes), so it's safe to treat as
					// "unset, try inheriting" without a separate sentinel.
					if cur.Placeholder != "" && cur.X == 0 && cur.Y == 0 && cur.Cx == 0 && cur.Cy == 0 {
						if g, ok := inh.resolveGeom(cur.Placeholder, curIdx); ok {
							cur.X, cur.Y, cur.Cx, cur.Cy = g.X, g.Y, g.Cx, g.Cy
						}
					}
					// <p:style> fallback fill/stroke — only when the shape
					// itself never resolved one from an explicit
					// <a:solidFill>/<a:ln><a:solidFill> (literal or
					// schemeClr, both already handled above by applyColor).
					if cur.Fill == "" && styleFillIdx != "" && styleFillIdx != "0" && styleFillClr != "" {
						cur.Fill = styleFillClr
					}
					if cur.Stroke == "" && styleLnIdx != "" && styleLnIdx != "0" && styleLnClr != "" {
						cur.Stroke = styleLnClr
					}
					// For a supported group tree, flatten this child's local
					// coordinates through every enclosing group. The conditions
					// are repeated here because malformed XML could otherwise
					// change group validity after this shape was opened.
					if curValid && groupsSupported() && flattenGroups(cur) {
						slide.Shapes = append(slide.Shapes, *cur)
					}
					cur = nil
				}
			}
			if len(stack) > 0 {
				stack = stack[:len(stack)-1]
			}
		case xml.CharData:
			if curRun != nil && top() == "t" && inside("r") {
				curRun.Text += string(t)
			}
		}
	}
	return slide, nil
}
