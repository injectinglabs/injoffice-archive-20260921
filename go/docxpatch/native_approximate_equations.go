package docxpatch

import (
	"encoding/xml"
	"fmt"
	"strings"
	"unicode/utf16"
	"unicode/utf8"
)

// Read-only OMML sidecar for the explicitly labeled approximate page preview.
//
// The strict extractor keeps refusing m:oMath / m:oMathPara paragraph content
// (UNMODELED_PARAGRAPH_CONTENT). This inspection joins those very diagnostics
// back to their source nodes and describes a bounded OMML subset as a layout
// tree with resolved text runs: rows, text runs (m:sty / m:scr / m:nor),
// fractions (bar / noBar), scripts (sSup, sSub, sSubSup), n-ary operators with
// upper/lower or side limits, delimiters, radicals, functions, bars, accents
// and lower/upper limits, plus m:oMathPara justification. Everything outside
// the subset omits the whole equation with a declared reason. Source bytes,
// strict extraction, editing authority and pagination are unchanged.
const NativeApproximateEquationsProtocol = "injoffice.docx.approximate-equations"
const NativeApproximateEquationPolicy = "docx.approximate-equation-preview-v1"
const nativeApproximateEquationLimit = 64
const nativeApproximateEquationNodeLimit = 512
const nativeApproximateEquationDepthLimit = 32
const nativeApproximateEquationTextLimit = 4096
const nativeApproximateEquationFontLimit = 32
const nativeApproximateEquationOmittedReason = "Unsupported equation structure or preview budget"

type NativeApproximateEquationFontV1 struct {
	Family string `json:"family"`
	Weight int    `json:"weight"`
	Style  string `json:"style"`
}

// NativeApproximateMathRunV1 is the resolved formatting of one math text run or
// of a construct's control characters (m:ctrlPr), after the paragraph style
// chain and the math style override.
type NativeApproximateMathRunV1 struct {
	FontFamily         string `json:"font_family"`
	FontSizeHalfPoints int    `json:"font_size_half_points"`
	Bold               bool   `json:"bold"`
	Italic             bool   `json:"italic"`
	Color              string `json:"color,omitempty"`
	// Authored m:sty value ("" when absent: letters lean per math convention).
	Style string `json:"style,omitempty"`
	// m:nor: text is normal (upright) prose inside the equation.
	Normal bool `json:"normal,omitempty"`
}

type NativeApproximateMathNodeV1 struct {
	Kind     string                        `json:"kind"`
	Text     string                        `json:"text,omitempty"`
	Run      *NativeApproximateMathRunV1   `json:"run,omitempty"`
	Children []NativeApproximateMathNodeV1 `json:"children,omitempty"`
	// nary / accent operator character.
	Chr string `json:"chr,omitempty"`
	// delimiter characters; nil keeps the OMML default, "" hides the character.
	BegChr *string `json:"beg_chr,omitempty"`
	EndChr *string `json:"end_chr,omitempty"`
	SepChr *string `json:"sep_chr,omitempty"`
	// fraction: false for m:type noBar (binomials).
	Bar *bool `json:"bar,omitempty"`
	// nary: undOvr (limits above/below) or subSup (limits at the side).
	LimitLocation string `json:"limit_location,omitempty"`
	SubHide       bool   `json:"sub_hide,omitempty"`
	SupHide       bool   `json:"sup_hide,omitempty"`
	DegreeHide    bool   `json:"degree_hide,omitempty"`
	// bar: top or bot.
	Position string `json:"position,omitempty"`
	Grow     bool   `json:"grow,omitempty"`
}

type NativeApproximateEquationV1 struct {
	ID            string                        `json:"id"`
	ParagraphID   string                        `json:"paragraph_id"`
	DiagnosticIDs []string                      `json:"diagnostic_ids"`
	Anchor        NativeSourceAnchorV1          `json:"anchor"`
	Status        string                        `json:"status"`
	Reason        string                        `json:"reason,omitempty"`
	Display       bool                          `json:"display"`
	Justification string                        `json:"justification,omitempty"`
	Lines         []NativeApproximateMathNodeV1 `json:"lines,omitempty"`
	Notes         []string                      `json:"notes,omitempty"`
}

type NativeApproximateEquationsV1 struct {
	Protocol      string                            `json:"protocol"`
	Version       int                               `json:"version"`
	Policy        string                            `json:"policy"`
	PackageSHA256 string                            `json:"package_sha256"`
	PartSHA256    string                            `json:"part_sha256"`
	Items         []NativeApproximateEquationV1     `json:"items"`
	OmittedCount  int                               `json:"omitted_count"`
	FontRequests  []NativeApproximateEquationFontV1 `json:"font_requests"`
}

type nativeApproximateEquationContext struct {
	resolver  *nativeLayoutResolver
	extractor *nativeExtractor
	ns, m     string
	main      string
	paragraph *NativeParagraphV1
	runs      []NativeRunV1
	styles    []string
	normal    []bool
	nodes     int
	units     int
	reason    string
	resolved  []NativeResolvedRunPropertiesV1
}

type nativeApproximateMathBuild struct {
	node     NativeApproximateMathNodeV1
	runIndex int
	children []*nativeApproximateMathBuild
}

// InspectNativeApproximateEquationsV1 returns nil when the body has no refused
// equation roots. Callers must treat every item as approximate evidence.
func InspectNativeApproximateEquationsV1(data []byte) (*NativeApproximateEquationsV1, error) {
	if len(data) == 0 || len(data) > NativeDOCXMaxPackageBytes {
		return nil, fmt.Errorf("approximate equations package size must be 1..%d bytes", NativeDOCXMaxPackageBytes)
	}
	data = append([]byte(nil), data...)
	resolver, err := newNativeLayoutResolver(data, NativeExtractionOptions{})
	if err != nil {
		return nil, err
	}
	doc := resolver.doc
	ns, main := resolver.wordNS, resolver.mainPart
	mathNS := nativePartialMathNamespace(ns)
	raw := resolver.pkg.files[main]
	nodesByPath := map[string]*nativeXMLNode{}
	var visit func(*nativeXMLNode)
	visit = func(n *nativeXMLNode) {
		if nativeDirectMathRoot(n, ns) {
			nodesByPath[n.Path] = n
			return
		}
		for _, c := range n.Children {
			visit(c)
		}
	}
	visit(resolver.mainRoot)
	if len(nodesByPath) == 0 {
		return nil, nil
	}
	rootDiagnostics := map[string]NativeUnsupportedCapabilityV1{}
	for _, d := range doc.Unsupported {
		if d.Code != "UNMODELED_PARAGRAPH_CONTENT" || d.Anchor == nil || d.Anchor.PartName != main {
			continue
		}
		if nodesByPath[d.Anchor.Path] != nil {
			rootDiagnostics[d.Anchor.Path] = d
		}
	}
	out := &NativeApproximateEquationsV1{Protocol: NativeApproximateEquationsProtocol, Version: 1, Policy: NativeApproximateEquationPolicy, PackageSHA256: doc.Source.PackageSHA256, PartSHA256: nativeSHA(raw), Items: []NativeApproximateEquationV1{}, FontRequests: []NativeApproximateEquationFontV1{}}
	fonts := map[string]NativeApproximateEquationFontV1{}
	fontOrder := []string{}
	for blockIndex := range doc.Body.Blocks {
		p := doc.Body.Blocks[blockIndex].Paragraph
		if p == nil {
			continue
		}
		owner := resolver.nodeForAnchor(p.Anchor)
		if owner == nil || owner.Name != (xml.Name{Space: ns, Local: "p"}) {
			continue
		}
		ordinal := 0
		for _, child := range owner.Children {
			root := nodesByPath[child.Path]
			if root == nil {
				continue
			}
			diagnostic, ok := rootDiagnostics[root.Path]
			if !ok || diagnostic.ScopeID != p.ID {
				// Without its retained strict refusal the root is not evidence.
				continue
			}
			ordinal++
			if len(out.Items) >= nativeApproximateEquationLimit {
				out.OmittedCount++
				continue
			}
			ids := []string{}
			for _, d := range doc.Unsupported {
				if d.ScopeID == p.ID && d.Anchor != nil && d.Anchor.PartName == main && d.Anchor.StartByte != nil && d.Anchor.EndByte != nil && *d.Anchor.StartByte >= root.Start && *d.Anchor.EndByte <= root.End {
					ids = append(ids, d.ID)
				}
			}
			item := NativeApproximateEquationV1{ID: fmt.Sprintf("approximate-equation:%s:%d", strings.TrimPrefix(p.ID, "paragraph:"), ordinal), ParagraphID: p.ID, DiagnosticIDs: ids, Anchor: *diagnostic.Anchor, Status: "omitted", Reason: nativeApproximateEquationOmittedReason, Display: root.Name.Local == "oMathPara"}
			context := &nativeApproximateEquationContext{resolver: resolver, extractor: nativeApproximateEquationExtractor(resolver), ns: ns, m: mathNS, main: main, paragraph: p}
			lines, justification, notes := context.describe(root)
			if context.reason != "" {
				item.Reason = context.reason
				item.Notes = notes
				out.Items = append(out.Items, item)
				continue
			}
			item.Status = "supported"
			item.Reason = ""
			item.Lines = lines
			item.Justification = justification
			item.Notes = notes
			for _, f := range context.fontRequests(lines) {
				key := f.Family + "\x00" + fmt.Sprint(f.Weight) + "\x00" + f.Style
				if _, seen := fonts[key]; seen {
					continue
				}
				if len(fontOrder) >= nativeApproximateEquationFontLimit {
					item.Notes = append(item.Notes, "font request beyond the preview budget")
					break
				}
				fonts[key] = f
				fontOrder = append(fontOrder, key)
			}
			out.Items = append(out.Items, item)
		}
	}
	for _, key := range fontOrder {
		out.FontRequests = append(out.FontRequests, fonts[key])
	}
	return out, nil
}

// nativeApproximateEquationExtractor is a private extractor over the same
// package. Its diagnostics only decide refusal and are never merged.
func nativeApproximateEquationExtractor(resolver *nativeLayoutResolver) *nativeExtractor {
	relNS := relNSTransitional
	if resolver.wordNS == wordMLStrict {
		relNS = relNSStrict
	}
	return &nativeExtractor{
		pkg: resolver.pkg, wordNS: resolver.wordNS, relNS: relNS, mainPart: resolver.mainPart, mainRoot: resolver.mainRoot,
		modeledParts: map[string]bool{resolver.mainPart: true}, unsupportedSet: map[string]bool{}, storyByRel: map[string]string{}, storyByNative: map[string]string{}, commentByNative: map[string]string{},
		unsupported: []NativeUnsupportedCapabilityV1{}, previousIDs: map[string][]string{}, previousUsed: map[string]int{}, previousNative: map[string]string{}, previousPath: map[string]string{}, reservedIDs: map[string]bool{}, allocatedIDs: map[string]bool{}, seenParaIDs: map[string]string{},
	}
}

func (context *nativeApproximateEquationContext) refuse(reason string) *nativeApproximateMathBuild {
	if context.reason == "" {
		context.reason = reason
	}
	return nil
}

func (context *nativeApproximateEquationContext) mathName(local string) xml.Name {
	return xml.Name{Space: context.m, Local: local}
}

// describe returns one layout row per m:oMath line, the oMathPara
// justification, and approximation notes; context.reason names a refusal.
func (context *nativeApproximateEquationContext) describe(root *nativeXMLNode) ([]NativeApproximateMathNodeV1, string, []string) {
	notes := []string{}
	justification := ""
	roots := []*nativeXMLNode{root}
	if root.Name.Local == "oMathPara" {
		justification = "centerGroup"
		roots = nil
		for _, child := range root.Children {
			switch {
			case child.Name == context.mathName("oMathParaPr"):
				for _, property := range child.Children {
					switch {
					case property.Name == context.mathName("jc"):
						value, _ := nativeAttr(property, context.m, "val")
						if value != "left" && value != "right" && value != "center" && value != "centerGroup" {
							context.refuse("unsupported equation paragraph justification")
							return nil, "", notes
						}
						justification = value
					case property.Name == context.mathName("brkBin"), property.Name == context.mathName("brkBinSub"):
						notes = append(notes, "binary operator break preferences are not applied")
					default:
						context.refuse("unsupported equation paragraph property " + nativeApproximateEquationName(property))
						return nil, "", notes
					}
				}
			case child.Name == context.mathName("oMath"):
				roots = append(roots, child)
			default:
				context.refuse("unsupported equation paragraph child " + nativeApproximateEquationName(child))
				return nil, "", notes
			}
		}
		if len(roots) == 0 {
			context.refuse("equation paragraph without equations")
			return nil, "", notes
		}
	}
	builds := []*nativeApproximateMathBuild{}
	for _, line := range roots {
		build := context.row(line, 0)
		if build == nil {
			return nil, "", notes
		}
		builds = append(builds, build)
	}
	resolved := context.resolveRuns()
	if resolved == nil {
		return nil, "", notes
	}
	context.resolved = resolved
	lines := []NativeApproximateMathNodeV1{}
	for _, build := range builds {
		lines = append(lines, context.finish(build, resolved))
	}
	if context.reason != "" {
		return nil, "", notes
	}
	return lines, justification, notes
}

func nativeApproximateEquationName(n *nativeXMLNode) string {
	return "m:" + n.Name.Local
}

func (context *nativeApproximateEquationContext) budget(depth int) bool {
	context.nodes++
	if depth > nativeApproximateEquationDepthLimit || context.nodes > nativeApproximateEquationNodeLimit {
		context.refuse("equation exceeds the preview depth or node budget")
		return false
	}
	return true
}

// row reads an argument container (oMath, e, num, den, sub, sup, deg, fName,
// lim) into a row of constructs.
func (context *nativeApproximateEquationContext) row(n *nativeXMLNode, depth int) *nativeApproximateMathBuild {
	if !context.budget(depth) {
		return nil
	}
	if !nativeExactContainer(n) {
		return context.refuse("equation container carries unsupported attributes")
	}
	result := &nativeApproximateMathBuild{node: NativeApproximateMathNodeV1{Kind: "row"}, runIndex: -1}
	for _, child := range n.Children {
		if child.Name.Space != context.m {
			return context.refuse("foreign equation content " + child.Name.Local)
		}
		var item *nativeApproximateMathBuild
		switch child.Name.Local {
		case "argPr":
			return context.refuse("unsupported equation argument property")
		case "r":
			item = context.run(child, depth+1)
		default:
			item = context.construct(child, depth+1)
		}
		if item == nil {
			return nil
		}
		result.children = append(result.children, item)
	}
	return result
}

// run reads m:r: optional m:rPr, optional w:rPr, then m:t text.
func (context *nativeApproximateEquationContext) run(n *nativeXMLNode, depth int) *nativeApproximateMathBuild {
	if !context.budget(depth) {
		return nil
	}
	if !nativeExactContainer(n) {
		return context.refuse("equation run carries unsupported attributes")
	}
	style, normal := "", false
	var wordProperties *nativeXMLNode
	text := strings.Builder{}
	sawText := false
	for _, child := range n.Children {
		switch {
		case child.Name == context.mathName("rPr"):
			for _, property := range child.Children {
				switch {
				case property.Name == context.mathName("sty"):
					value, _ := nativeAttr(property, context.m, "val")
					if value != "p" && value != "b" && value != "i" && value != "bi" {
						return context.refuse("unsupported math text style")
					}
					style = value
				case property.Name == context.mathName("scr"):
					value, _ := nativeAttr(property, context.m, "val")
					if value != "roman" {
						return context.refuse("unsupported math script alphabet " + value)
					}
				case property.Name == context.mathName("nor"):
					value, ok := nativeApproximateMathOnOff(property, context.m)
					if !ok {
						return context.refuse("invalid normal-text flag")
					}
					normal = value
				case property.Name == context.mathName("lit"), property.Name == context.mathName("aln"):
					// Literal / alignment hints do not change glyph placement here.
				default:
					return context.refuse("unsupported math run property " + nativeApproximateEquationName(property))
				}
			}
		case child.Name == (xml.Name{Space: context.ns, Local: "rPr"}):
			if wordProperties != nil {
				return context.refuse("duplicate run properties in equation run")
			}
			wordProperties = child
		case child.Name == context.mathName("t"):
			if len(child.Children) != 0 {
				return context.refuse("equation text contains markup")
			}
			for _, a := range child.Attrs {
				if a.Name.Space == "xmlns" || a.Name.Local == "xmlns" && a.Name.Space == "" {
					continue
				}
				if a.Name != (xml.Name{Space: "http://www.w3.org/XML/1998/namespace", Local: "space"}) || a.Value != "preserve" {
					return context.refuse("equation text carries unsupported attributes")
				}
			}
			if !utf8.ValidString(child.Text) {
				return context.refuse("equation text is not valid UTF-8")
			}
			context.units += len(utf16.Encode([]rune(child.Text)))
			if context.units > nativeApproximateEquationTextLimit {
				return context.refuse("equation text exceeds the preview budget")
			}
			text.WriteString(child.Text)
			sawText = true
		default:
			return context.refuse("unsupported equation run content " + child.Name.Local)
		}
	}
	if !sawText {
		return context.refuse("equation run without text")
	}
	index := context.registerRun(n, wordProperties, text.String(), style, normal)
	if index < 0 {
		return nil
	}
	return &nativeApproximateMathBuild{node: NativeApproximateMathNodeV1{Kind: "text", Text: text.String()}, runIndex: index}
}

// registerRun queues one synthetic run for the ordinary style resolver. The
// anchor is the math node itself, so no w:r owner is found and the extracted
// direct properties layer over the paragraph style chain exactly like body
// text without a raw owner.
func (context *nativeApproximateEquationContext) registerRun(n, wordProperties *nativeXMLNode, text, style string, normal bool) int {
	var properties *NativeRunPropertiesV1
	if wordProperties != nil {
		var unsafe bool
		properties, unsafe = context.extractor.extractRunProperties(context.main, context.paragraph.ID, wordProperties)
		if unsafe {
			context.refuse("unsupported run properties in equation")
			return -1
		}
	}
	index := len(context.runs)
	run := NativeRunV1{Kind: "text", ID: fmt.Sprintf("%s:math%d", context.paragraph.ID, index), Anchor: context.extractor.anchor(context.main, n), Properties: properties, Text: nativeString(text)}
	context.runs = append(context.runs, run)
	context.styles = append(context.styles, style)
	context.normal = append(context.normal, normal)
	return index
}

// control reads a construct property container: m:ctrlPr supplies the operator
// formatting; the remaining known properties are returned by local name.
func (context *nativeApproximateEquationContext) control(pr *nativeXMLNode, allowed map[string]bool) (int, map[string]*nativeXMLNode, bool) {
	properties := map[string]*nativeXMLNode{}
	runIndex := -1
	if pr == nil {
		return runIndex, properties, true
	}
	for _, child := range pr.Children {
		if child.Name.Space != context.m {
			context.refuse("foreign equation property " + child.Name.Local)
			return -1, nil, false
		}
		if child.Name.Local == "ctrlPr" {
			var wordProperties *nativeXMLNode
			for _, inner := range child.Children {
				if inner.Name != (xml.Name{Space: context.ns, Local: "rPr"}) || wordProperties != nil {
					context.refuse("unsupported control character properties")
					return -1, nil, false
				}
				wordProperties = inner
			}
			runIndex = context.registerRun(child, wordProperties, "", "", false)
			if runIndex < 0 {
				return -1, nil, false
			}
			continue
		}
		if !allowed[child.Name.Local] {
			context.refuse("unsupported equation property " + nativeApproximateEquationName(child))
			return -1, nil, false
		}
		if _, duplicate := properties[child.Name.Local]; duplicate {
			context.refuse("duplicate equation property " + nativeApproximateEquationName(child))
			return -1, nil, false
		}
		properties[child.Name.Local] = child
	}
	return runIndex, properties, true
}

func nativeApproximateMathOnOff(n *nativeXMLNode, mathNS string) (bool, bool) {
	value, ok := nativeAttr(n, mathNS, "val")
	if !ok {
		return true, true
	}
	switch value {
	case "1", "true", "on":
		return true, true
	case "0", "false", "off":
		return false, true
	}
	return false, false
}

// chr reads one operator character; an empty value is an explicit "none".
func (context *nativeApproximateEquationContext) chr(n *nativeXMLNode, allowEmpty bool) (string, bool) {
	value, ok := nativeAttr(n, context.m, "val")
	if !ok {
		context.refuse("equation operator character without a value")
		return "", false
	}
	if value == "" {
		if !allowEmpty {
			context.refuse("empty equation operator character")
			return "", false
		}
		return "", true
	}
	if !utf8.ValidString(value) || utf8.RuneCountInString(value) != 1 {
		context.refuse("equation operator character must be exactly one character")
		return "", false
	}
	return value, true
}

// construct reads one OMML construct element with its fixed argument order.
func (context *nativeApproximateEquationContext) construct(n *nativeXMLNode, depth int) *nativeApproximateMathBuild {
	if !context.budget(depth) {
		return nil
	}
	if !nativeExactContainer(n) {
		return context.refuse("equation construct carries unsupported attributes")
	}
	local := n.Name.Local
	var pr *nativeXMLNode
	arguments := []*nativeXMLNode{}
	for _, child := range n.Children {
		if child.Name.Space != context.m {
			return context.refuse("foreign equation content " + child.Name.Local)
		}
		if child.Name.Local == local+"Pr" {
			if pr != nil {
				return context.refuse("duplicate equation construct properties")
			}
			pr = child
			continue
		}
		arguments = append(arguments, child)
	}
	names := func(want ...string) bool {
		if len(arguments) != len(want) {
			return false
		}
		for i, w := range want {
			if arguments[i].Name != context.mathName(w) {
				return false
			}
		}
		return true
	}
	result := &nativeApproximateMathBuild{runIndex: -1}
	var allowed map[string]bool
	switch local {
	case "f":
		result.node.Kind = "fraction"
		allowed = map[string]bool{"type": true}
		if !names("num", "den") {
			return context.refuse("fraction requires numerator and denominator")
		}
	case "sSup":
		result.node.Kind = "superscript"
		allowed = map[string]bool{"alnScr": true}
		if !names("e", "sup") {
			return context.refuse("superscript requires base and script")
		}
	case "sSub":
		result.node.Kind = "subscript"
		allowed = map[string]bool{"alnScr": true}
		if !names("e", "sub") {
			return context.refuse("subscript requires base and script")
		}
	case "sSubSup":
		result.node.Kind = "subsuperscript"
		allowed = map[string]bool{"alnScr": true}
		if !names("e", "sub", "sup") {
			return context.refuse("sub-superscript requires base and both scripts")
		}
	case "nary":
		result.node.Kind = "nary"
		allowed = map[string]bool{"chr": true, "limLoc": true, "grow": true, "subHide": true, "supHide": true}
		if !names("sub", "sup", "e") {
			return context.refuse("n-ary operator requires lower limit, upper limit and body")
		}
	case "d":
		result.node.Kind = "delimiter"
		allowed = map[string]bool{"begChr": true, "endChr": true, "sepChr": true, "grow": true, "shp": true}
		if len(arguments) == 0 {
			return context.refuse("delimiter without content")
		}
		for _, a := range arguments {
			if a.Name != context.mathName("e") {
				return context.refuse("delimiter content must be arguments")
			}
		}
	case "rad":
		result.node.Kind = "radical"
		allowed = map[string]bool{"degHide": true}
		if !names("deg", "e") {
			return context.refuse("radical requires degree and radicand")
		}
	case "func":
		result.node.Kind = "function"
		allowed = map[string]bool{}
		if !names("fName", "e") {
			return context.refuse("function requires name and argument")
		}
	case "bar":
		result.node.Kind = "bar"
		allowed = map[string]bool{"pos": true}
		if !names("e") {
			return context.refuse("bar requires one argument")
		}
	case "acc":
		result.node.Kind = "accent"
		allowed = map[string]bool{"chr": true}
		if !names("e") {
			return context.refuse("accent requires one argument")
		}
	case "limLow":
		result.node.Kind = "limit-lower"
		allowed = map[string]bool{}
		if !names("e", "lim") {
			return context.refuse("lower limit requires base and limit")
		}
	case "limUpp":
		result.node.Kind = "limit-upper"
		allowed = map[string]bool{}
		if !names("e", "lim") {
			return context.refuse("upper limit requires base and limit")
		}
	default:
		return context.refuse("unsupported equation element " + nativeApproximateEquationName(n))
	}
	runIndex, properties, ok := context.control(pr, allowed)
	if !ok {
		return nil
	}
	result.runIndex = runIndex
	for name, property := range properties {
		switch name {
		case "type":
			value, _ := nativeAttr(property, context.m, "val")
			switch value {
			case "bar", "":
				result.node.Bar = nativeBool(true)
			case "noBar":
				result.node.Bar = nativeBool(false)
			default:
				return context.refuse("unsupported fraction type " + value)
			}
		case "chr":
			value, ok := context.chr(property, false)
			if !ok {
				return nil
			}
			result.node.Chr = value
		case "begChr", "endChr", "sepChr":
			value, ok := context.chr(property, true)
			if !ok {
				return nil
			}
			switch name {
			case "begChr":
				result.node.BegChr = nativeString(value)
			case "endChr":
				result.node.EndChr = nativeString(value)
			default:
				result.node.SepChr = nativeString(value)
			}
		case "limLoc":
			value, _ := nativeAttr(property, context.m, "val")
			if value != "undOvr" && value != "subSup" {
				return context.refuse("unsupported n-ary limit location")
			}
			result.node.LimitLocation = value
		case "pos":
			value, _ := nativeAttr(property, context.m, "val")
			if value != "top" && value != "bot" {
				return context.refuse("unsupported bar position")
			}
			result.node.Position = value
		case "grow", "subHide", "supHide", "degHide":
			value, ok := nativeApproximateMathOnOff(property, context.m)
			if !ok {
				return context.refuse("invalid equation flag " + name)
			}
			switch name {
			case "grow":
				result.node.Grow = value
			case "subHide":
				result.node.SubHide = value
			case "supHide":
				result.node.SupHide = value
			default:
				result.node.DegreeHide = value
			}
		case "alnScr", "shp":
			// Script alignment and delimiter shape hints do not change this preview.
		}
	}
	if result.node.Kind == "fraction" && result.node.Bar == nil {
		result.node.Bar = nativeBool(true)
	}
	for _, argument := range arguments {
		child := context.row(argument, depth+1)
		if child == nil {
			return nil
		}
		result.children = append(result.children, child)
	}
	return result
}

// resolveRuns resolves every queued run through the ordinary paragraph style
// resolver of the owning paragraph and returns the exported properties by run
// index, with the paragraph mark properties last.
func (context *nativeApproximateEquationContext) resolveRuns() []NativeResolvedRunPropertiesV1 {
	if context.reason != "" {
		return nil
	}
	local := *context.resolver
	local.diagnostics = append([]NativeResolutionDiagnosticV1{}, context.resolver.diagnostics...)
	local.diagnosticSet = map[string]bool{}
	for key, value := range context.resolver.diagnosticSet {
		local.diagnosticSet[key] = value
	}
	paragraph := *context.paragraph
	paragraph.Runs = append([]NativeRunV1{}, context.runs...)
	result := &NativeResolvedLayoutInputV1{}
	local.resolveParagraph(&paragraph, result, newNativeNumberingState(), nil)
	if local.diagnosticOverflow || len(result.Paragraphs) != 1 || len(result.Runs) != len(context.runs) {
		context.refuse("equation run formatting could not be resolved")
		return nil
	}
	out := make([]NativeResolvedRunPropertiesV1, 0, len(context.runs)+1)
	for _, run := range result.Runs {
		out = append(out, run.Properties)
	}
	return append(out, result.Paragraphs[0].ParagraphMarkProperties)
}

func (context *nativeApproximateEquationContext) mathRun(index int, resolved []NativeResolvedRunPropertiesV1) *NativeApproximateMathRunV1 {
	properties := resolved[index]
	mark := resolved[len(resolved)-1]
	run := &NativeApproximateMathRunV1{}
	switch {
	case properties.FontFamily != nil:
		run.FontFamily = *properties.FontFamily
	case mark.FontFamily != nil:
		run.FontFamily = *mark.FontFamily
	default:
		context.refuse("equation font family is unresolved")
		return nil
	}
	switch {
	case properties.FontSizeHalfPoint != nil && *properties.FontSizeHalfPoint > 0:
		run.FontSizeHalfPoints = *properties.FontSizeHalfPoint
	case mark.FontSizeHalfPoint != nil && *mark.FontSizeHalfPoint > 0:
		run.FontSizeHalfPoints = *mark.FontSizeHalfPoint
	default:
		context.refuse("equation font size is unresolved")
		return nil
	}
	if run.FontSizeHalfPoints > 3276 {
		context.refuse("equation font size exceeds the preview bound")
		return nil
	}
	run.Bold = properties.Bold != nil && *properties.Bold
	run.Italic = properties.Italic != nil && *properties.Italic
	if properties.Hidden != nil && *properties.Hidden {
		context.refuse("hidden equation text is not previewed")
		return nil
	}
	if properties.Color != nil {
		run.Color = *properties.Color
	}
	if index < len(context.styles) {
		run.Style = context.styles[index]
		run.Normal = context.normal[index]
		switch run.Style {
		case "p":
			run.Bold, run.Italic = false, false
		case "b":
			run.Bold, run.Italic = true, false
		case "i":
			run.Bold, run.Italic = false, true
		case "bi":
			run.Bold, run.Italic = true, true
		}
	}
	return run
}

func (context *nativeApproximateEquationContext) finish(build *nativeApproximateMathBuild, resolved []NativeResolvedRunPropertiesV1) NativeApproximateMathNodeV1 {
	node := build.node
	if build.runIndex >= 0 {
		node.Run = context.mathRun(build.runIndex, resolved)
	}
	for _, child := range build.children {
		node.Children = append(node.Children, context.finish(child, resolved))
	}
	return node
}

// fontRequests lists the distinct face identities the preview will ask the
// host for: every resolved math run plus the paragraph text face in the same
// weights and styles, which is the declared fallback when no math face exists.
func (context *nativeApproximateEquationContext) fontRequests(lines []NativeApproximateMathNodeV1) []NativeApproximateEquationFontV1 {
	requests := []NativeApproximateEquationFontV1{}
	seen := map[string]bool{}
	add := func(family string, bold, italic bool) {
		if family == "" {
			return
		}
		weight, style := 400, "normal"
		if bold {
			weight = 700
		}
		if italic {
			style = "italic"
		}
		key := family + "\x00" + fmt.Sprint(weight) + "\x00" + style
		if seen[key] {
			return
		}
		seen[key] = true
		requests = append(requests, NativeApproximateEquationFontV1{Family: family, Weight: weight, Style: style})
	}
	mark := ""
	if len(context.resolved) > 0 && context.resolved[len(context.resolved)-1].FontFamily != nil {
		mark = *context.resolved[len(context.resolved)-1].FontFamily
	}
	var visit func(n *NativeApproximateMathNodeV1)
	visit = func(n *NativeApproximateMathNodeV1) {
		if n.Run != nil {
			add(n.Run.FontFamily, n.Run.Bold, n.Run.Italic)
			add(mark, n.Run.Bold, n.Run.Italic)
			// Math conventions lean letters and keep operators upright.
			add(n.Run.FontFamily, n.Run.Bold, !n.Run.Italic)
			add(mark, n.Run.Bold, !n.Run.Italic)
		}
		for i := range n.Children {
			visit(&n.Children[i])
		}
	}
	for i := range lines {
		visit(&lines[i])
	}
	return requests
}
