package docxpatch

import (
	"encoding/json"
	"encoding/xml"
	"fmt"
	"path"
	"sort"
	"strconv"
	"strings"
	"unicode/utf8"
)

const (
	NativeDOCXResolvedLayoutProtocol       = "injoffice.docx.resolved-layout"
	NativeDOCXResolvedLayoutVersion        = 1
	NativeDOCXMaxResolvedDiagnostics       = 1_000
	nativeMaxTwipsForMilliPoints     int64 = 180143985094819
)

// NativeResolvedLayoutInputV1 is a deterministic, read-only, versioned
// projection of the native DOCX contract for a future text shaper and
// paginator. It deliberately contains no pagination decisions and does not
// replace NativeDocumentV1.
type NativeResolvedLayoutInputV1 struct {
	Protocol        string                           `json:"protocol"`
	Version         int                              `json:"version"`
	DocumentID      string                           `json:"document_id"`
	Revision        string                           `json:"revision"`
	SourceParts     NativeResolvedSourcePartsV1      `json:"source_parts"`
	NumberingSource *NativeResolvedNumberingSourceV1 `json:"numbering_source,omitempty"`
	Paragraphs      []NativeResolvedParagraphV1      `json:"paragraphs"`
	Runs            []NativeResolvedRunV1            `json:"runs"`
	Tables          []NativeResolvedTableV1          `json:"tables"`
	Fonts           []NativeResolvedFontV1           `json:"fonts"`
	Diagnostics     []NativeResolutionDiagnosticV1   `json:"diagnostics"`
}

type NativeResolvedSourcePartsV1 struct {
	MainPart      string  `json:"main_part"`
	StylesPart    *string `json:"styles_part,omitempty"`
	NumberingPart *string `json:"numbering_part,omitempty"`
	ThemePart     *string `json:"theme_part,omitempty"`
	FontTablePart *string `json:"font_table_part,omitempty"`
}

// NativeResolvedNumberingSourceV1 attests the exact OPC relationship closure
// and raw numbering bytes used to resolve native list markers. The relationship
// and numbering parts remain preserved in NativeDocumentV1.PassthroughParts;
// this projection makes their hashes explicit at every layout boundary.
type NativeResolvedNumberingSourceV1 struct {
	RelationshipsPart   string `json:"relationships_part"`
	RelationshipsSHA256 string `json:"relationships_sha256"`
	RelationshipID      string `json:"relationship_id"`
	RelationshipType    string `json:"relationship_type"`
	RelationshipTarget  string `json:"relationship_target"`
	PartName            string `json:"part_name"`
	ContentType         string `json:"content_type"`
	PartSHA256          string `json:"part_sha256"`
	ModelSHA256         string `json:"model_sha256"`
}

type NativeResolvedParagraphV1 struct {
	ParagraphID             string                              `json:"paragraph_id"`
	StyleID                 *string                             `json:"style_id,omitempty"`
	AppliedStyles           []string                            `json:"applied_styles"`
	Properties              NativeResolvedParagraphPropertiesV1 `json:"properties"`
	ParagraphMarkProperties NativeResolvedRunPropertiesV1       `json:"paragraph_mark_properties"`
	Numbering               *NativeResolvedNumberingV1          `json:"numbering,omitempty"`
}

type NativeResolvedRunV1 struct {
	RunID                  string                        `json:"run_id"`
	ParagraphID            string                        `json:"paragraph_id"`
	CharacterStyle         *string                       `json:"character_style_id,omitempty"`
	AppliedParagraphStyles []string                      `json:"applied_paragraph_styles"`
	AppliedCharacterStyles []string                      `json:"applied_character_styles"`
	Properties             NativeResolvedRunPropertiesV1 `json:"properties"`
}

type NativeResolvedTableV1 struct {
	TableID        string                `json:"table_id"`
	StyleID        *string               `json:"style_id,omitempty"`
	Borders        *NativeTableBordersV1 `json:"borders,omitempty"`
	CellShadingRGB *string               `json:"cell_shading_rgb,omitempty"`
}

type NativeResolvedParagraphPropertiesV1 struct {
	Alignment          *string `json:"alignment,omitempty"`
	SpacingBeforeTwips *int64  `json:"spacing_before_twips,omitempty"`
	SpacingAfterTwips  *int64  `json:"spacing_after_twips,omitempty"`
	Line               *int64  `json:"line,omitempty"`
	LineRule           *string `json:"line_rule,omitempty"`
	IndentLeftTwips    *int64  `json:"indent_left_twips,omitempty"`
	IndentRightTwips   *int64  `json:"indent_right_twips,omitempty"`
	IndentStartTwips   *int64  `json:"indent_start_twips,omitempty"`
	IndentEndTwips     *int64  `json:"indent_end_twips,omitempty"`
	FirstLineTwips     *int64  `json:"first_line_twips,omitempty"`
	HangingTwips       *int64  `json:"hanging_twips,omitempty"`
	Bidi               *bool   `json:"bidi,omitempty"`
	KeepNext           *bool   `json:"keep_next,omitempty"`
	KeepLines          *bool   `json:"keep_lines,omitempty"`
	PageBreakBefore    *bool   `json:"page_break_before,omitempty"`
	WidowControl       *bool   `json:"widow_control,omitempty"`
}

type NativeResolvedRunPropertiesV1 struct {
	FontFamily        *string `json:"font_family,omitempty"`
	FontSizeHalfPoint *int    `json:"font_size_half_points,omitempty"`
	Bold              *bool   `json:"bold,omitempty"`
	Italic            *bool   `json:"italic,omitempty"`
	Underline         *string `json:"underline,omitempty"`
	VerticalAlignment *string `json:"vertical_alignment,omitempty"`
	Color             *string `json:"color,omitempty"`
	Highlight         *string `json:"highlight,omitempty"`
	Language          *string `json:"language,omitempty"`
	RTL               *bool   `json:"rtl,omitempty"`
	Hidden            *bool   `json:"hidden,omitempty"`
}

type NativeResolvedNumberingV1 struct {
	MarkerID           string                         `json:"marker_id"`
	DefinitionSHA256   string                         `json:"definition_sha256"`
	NumID              string                         `json:"num_id"`
	AbstractNumID      string                         `json:"abstract_num_id"`
	Level              int                            `json:"level"`
	LevelStyleID       *string                        `json:"level_style_id,omitempty"`
	Start              int                            `json:"start"`
	Format             string                         `json:"format"`
	Text               string                         `json:"text"`
	Suffix             string                         `json:"suffix"`
	Alignment          string                         `json:"alignment"`
	RestartAfterLevel  *int                           `json:"restart_after_level,omitempty"`
	NeverRestart       bool                           `json:"never_restart"`
	CounterValue       int                            `json:"counter_value"`
	CounterValues      []NativeResolvedCounterValueV1 `json:"counter_values"`
	ResolvedText       string                         `json:"resolved_text"`
	LabelStartTwips    int64                          `json:"label_start_twips"`
	LabelEndTwips      int64                          `json:"label_end_twips"`
	TextStartTwips     int64                          `json:"text_start_twips"`
	NumberingTabTwips  *int64                         `json:"numbering_tab_twips,omitempty"`
	Marker             NativeResolvedRunPropertiesV1  `json:"marker_properties"`
	alignmentDefaulted bool
}

type NativeResolvedCounterValueV1 struct {
	Level  int    `json:"level"`
	Value  int    `json:"value"`
	Format string `json:"format"`
}

type NativeResolvedFontV1 struct {
	Name    string  `json:"name"`
	AltName *string `json:"alt_name,omitempty"`
}

type NativeResolutionDiagnosticV1 struct {
	Code         string  `json:"code"`
	Severity     string  `json:"severity"`
	ScopeID      string  `json:"scope_id"`
	PartName     *string `json:"part_name,omitempty"`
	Path         *string `json:"path,omitempty"`
	Preservation string  `json:"preservation"`
	Message      string  `json:"message"`
}

// ResolveNativeDocumentLayoutV1 extracts the native document and resolves the
// conservative style/numbering subset required by a future shaper. It never
// renders HTML and never mutates or regenerates the DOCX package.
func ResolveNativeDocumentLayoutV1(data []byte) (*NativeResolvedLayoutInputV1, error) {
	return ResolveNativeDocumentLayoutV1WithOptions(data, NativeExtractionOptions{})
}

// ResolveNativeDocumentLayout is the concise API alias for the versioned v1
// resolved-layout projection.
func ResolveNativeDocumentLayout(data []byte) (*NativeResolvedLayoutInputV1, error) {
	return ResolveNativeDocumentLayoutV1(data)
}

// ResolveNativeDocumentLayoutV1WithOptions preserves the extractor's durable
// identity options while returning a separate, derived layout-input model.
func ResolveNativeDocumentLayoutV1WithOptions(data []byte, options NativeExtractionOptions) (*NativeResolvedLayoutInputV1, error) {
	doc, err := ExtractNativeDocumentV1WithOptions(data, options)
	if err != nil {
		return nil, err
	}
	pkg, err := openNativeDOCXPackage(data)
	if err != nil {
		return nil, err
	}
	mainPart, strict, err := pkg.officeDocumentPart()
	if err != nil {
		return nil, err
	}
	resolver := &nativeLayoutResolver{
		pkg: pkg, doc: doc, mainPart: mainPart,
		wordNS: wordMLTransitional, relBase: relBaseTransitional,
		styles:        map[string]*nativeStyleDefinition{},
		abstractNums:  map[string]*nativeAbstractNumbering{},
		nums:          map[string]*nativeNumberingInstance{},
		diagnosticSet: map[string]bool{},
		nodeByAnchor:  map[string]*nativeXMLNode{},
	}
	if strict {
		resolver.wordNS, resolver.relBase = wordMLStrict, relBaseStrict
	}
	if err := resolver.loadParts(); err != nil {
		return nil, err
	}
	result, err := resolver.resolve()
	if err != nil {
		return nil, err
	}
	if _, err := EncodeNativeResolvedLayoutInputV1(result); err != nil {
		return nil, fmt.Errorf("docxpatch: native style resolution output is invalid: %w", err)
	}
	return result, nil
}

// EncodeNativeResolvedLayoutInputV1 validates and deterministically encodes a
// resolved layout input. Struct field order and source-order arrays are stable.
func EncodeNativeResolvedLayoutInputV1(input *NativeResolvedLayoutInputV1) ([]byte, error) {
	if err := ValidateNativeResolvedLayoutInputV1(input); err != nil {
		return nil, err
	}
	encoded, err := json.Marshal(input)
	if err != nil {
		return nil, err
	}
	if len(encoded) > NativeDOCXMaxJSONBytes {
		return nil, fmt.Errorf("resolved layout JSON exceeds %d bytes", NativeDOCXMaxJSONBytes)
	}
	return encoded, nil
}

type nativeLayoutResolver struct {
	pkg                     *nativePackage
	doc                     *NativeDocumentV1
	mainPart                string
	wordNS                  string
	relBase                 string
	parts                   NativeResolvedSourcePartsV1
	numberingSource         *NativeResolvedNumberingSourceV1
	styles                  map[string]*nativeStyleDefinition
	abstractNums            map[string]*nativeAbstractNumbering
	nums                    map[string]*nativeNumberingInstance
	docP                    nativeParagraphProperties
	docR                    nativeRunProperties
	defaultP                string
	defaultC                string
	fonts                   []NativeResolvedFontV1
	unusedFontDescriptors   []nativeUnusedFontDescriptor
	diagnostics             []NativeResolutionDiagnosticV1
	diagnosticSet           map[string]bool
	diagnosticOverflow      bool
	deferredNumbering       *[]nativeDeferredNumberingDiagnostic
	deferredDiagnosticCount int
	numberingRootDeferred   []nativeDeferredNumberingDiagnostic
	nodeByAnchor            map[string]*nativeXMLNode
	themeSrgb               map[string]string
	themeLatinFonts         nativeThemeLatinFonts
}

type nativeDeferredNumberingDiagnostic struct {
	code, partName, message string
	node                    *nativeXMLNode
}

type nativeStyleDefinition struct {
	id       string
	kind     string
	basedOn  string
	partName string
	node     *nativeXMLNode
	p        nativeParagraphProperties
	r        nativeRunProperties
	deferred []nativeDeferredNumberingDiagnostic
}

type nativeAbstractNumbering struct {
	id        string
	levels    map[int]*nativeNumberingLevel
	multiSeen bool
	deferred  []nativeDeferredNumberingDiagnostic
}

type nativeNumberingInstance struct {
	id         string
	abstractID string
	overrides  map[int]*nativeNumberingOverride
	deferred   []nativeDeferredNumberingDiagnostic
}

type nativeNumberingOverride struct {
	start    *int
	level    *nativeNumberingLevel
	deferred []nativeDeferredNumberingDiagnostic
}

type nativeNumberingLevel struct {
	level          int
	start          *int
	format         *string
	text           *string
	suffix         *string
	alignment      *string
	p              nativeParagraphProperties
	r              nativeRunProperties
	picture        bool
	restart        *int
	numTab         *int64
	styleLinks     []string
	styleLinkCount int
	deferred       []nativeDeferredNumberingDiagnostic
	partName       string
	node           *nativeXMLNode
}

type nativeBoolProperty struct {
	present bool
	value   bool
}

type nativeRunProperties struct {
	scriptProperties  map[string]nativeDeferredNumberingDiagnostic
	fontFamily        *string
	asciiFamily       *string
	hAnsiFamily       *string
	fontSize          *int
	bold              nativeBoolProperty
	italic            nativeBoolProperty
	underline         *string
	verticalAlignment *string
	color             *string
	highlight         *string
	language          *string
	rtl               nativeBoolProperty
	hidden            nativeBoolProperty
}

type nativeParagraphProperties struct {
	numbering       nativeNumberingProperties
	alignment       *string
	spacingBefore   *int64
	spacingAfter    *int64
	line            *int64
	lineRule        *string
	indentLeft      *int64
	indentRight     *int64
	indentStart     *int64
	indentEnd       *int64
	firstLine       *int64
	hanging         *int64
	bidi            nativeBoolProperty
	keepNext        nativeBoolProperty
	keepLines       nativeBoolProperty
	pageBreakBefore nativeBoolProperty
	widowControl    nativeBoolProperty
}

type nativeNumberingProperties struct {
	present      bool
	numID        *string
	level        *int
	abstractHint *string
}

func (resolver *nativeLayoutResolver) loadParts() error {
	resolver.parts.MainPart = resolver.mainPart
	partKinds := []struct {
		kind   string
		target **string
	}{
		{"styles", &resolver.parts.StylesPart},
		{"numbering", &resolver.parts.NumberingPart},
		{"theme", &resolver.parts.ThemePart},
		{"fontTable", &resolver.parts.FontTablePart},
	}
	for _, partKind := range partKinds {
		part, err := resolver.singletonRelatedPart(partKind.kind)
		if err != nil {
			return err
		}
		if part != "" {
			if err := resolver.validateRelatedPartContentType(partKind.kind, part); err != nil {
				return err
			}
			*partKind.target = nativeString(part)
			if partKind.kind == "numbering" {
				rel, relErr := resolver.singletonRelatedRelationship("numbering")
				if relErr != nil {
					return relErr
				}
				relPart := resolver.pkg.relsPart[resolver.mainPart]
				if relPart == "" {
					return fmt.Errorf("docxpatch: native style resolution: numbering relationship has no owning relationship part")
				}
				resolver.numberingSource = &NativeResolvedNumberingSourceV1{
					RelationshipsPart: relPart, RelationshipsSHA256: nativeSHA(resolver.pkg.files[relPart]),
					RelationshipID: rel.ID, RelationshipType: rel.Type, RelationshipTarget: rel.Target,
					PartName: part, ContentType: resolver.pkg.contentTypes[part], PartSHA256: nativeSHA(resolver.pkg.files[part]),
					ModelSHA256: "sha256:" + strings.Repeat("0", 64),
				}
			}
		}
	}
	// Style/default and numbering properties resolve theme references while
	// parsing, so the validated related theme must be available first.
	if resolver.parts.ThemePart != nil {
		if err := resolver.loadTheme(*resolver.parts.ThemePart); err != nil {
			return err
		}
	}
	if resolver.parts.StylesPart != nil {
		if err := resolver.loadStyles(*resolver.parts.StylesPart); err != nil {
			return err
		}
	}
	if resolver.parts.NumberingPart != nil {
		if err := resolver.loadNumbering(*resolver.parts.NumberingPart); err != nil {
			return err
		}
	}
	if resolver.parts.FontTablePart != nil {
		if err := resolver.loadFonts(*resolver.parts.FontTablePart); err != nil {
			return err
		}
	}
	return resolver.indexStoryNodes()
}

func (resolver *nativeLayoutResolver) singletonRelatedRelationship(kind string) (nativeRelationship, error) {
	want := resolver.relBase + kind
	var found *nativeRelationship
	for index := range resolver.pkg.rels[resolver.mainPart] {
		rel := &resolver.pkg.rels[resolver.mainPart][index]
		if rel.Type != want {
			continue
		}
		if found != nil {
			return nativeRelationship{}, fmt.Errorf("docxpatch: native style resolution: multiple %s relationships", kind)
		}
		found = rel
	}
	if found == nil {
		return nativeRelationship{}, fmt.Errorf("docxpatch: native style resolution: missing %s relationship", kind)
	}
	return *found, nil
}

func (resolver *nativeLayoutResolver) validateRelatedPartContentType(kind, partName string) error {
	want := map[string]string{
		"styles":    "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
		"numbering": "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
		"theme":     "application/vnd.openxmlformats-officedocument.theme+xml",
		"fontTable": "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
	}[kind]
	if got := resolver.pkg.contentTypes[partName]; !nativeASCIIEqual(got, want) {
		return fmt.Errorf("docxpatch: native style resolution: %s part %q has content type %q; expected %q", kind, partName, got, want)
	}
	return nil
}

func (resolver *nativeLayoutResolver) indexStoryNodes() error {
	parts := map[string]bool{resolver.mainPart: true}
	for _, story := range resolver.allStories() {
		parts[story.PartName] = true
	}
	names := make([]string, 0, len(parts))
	for partName := range parts {
		names = append(names, partName)
	}
	sort.Strings(names)
	for _, partName := range names {
		root, err := parseNativeXML(partName, resolver.pkg.files[partName])
		if err != nil {
			return err
		}
		var visit func(*nativeXMLNode)
		visit = func(node *nativeXMLNode) {
			resolver.nodeByAnchor[nativeAnchorKey(partName, node.Start, node.End)] = node
			for _, child := range node.Children {
				visit(child)
			}
		}
		visit(root)
	}
	return nil
}

func nativeAnchorKey(partName string, start, end int64) string {
	if canonical, err := nativeDecodedPartKey(partName); err == nil {
		partName = canonical
	}
	return partName + "\x00" + strconv.FormatInt(start, 10) + ":" + strconv.FormatInt(end, 10)
}

func (resolver *nativeLayoutResolver) nodeForAnchor(anchor NativeSourceAnchorV1) *nativeXMLNode {
	if anchor.StartByte == nil || anchor.EndByte == nil {
		return nil
	}
	return resolver.nodeByAnchor[nativeAnchorKey(anchor.PartName, *anchor.StartByte, *anchor.EndByte)]
}

func (resolver *nativeLayoutResolver) singletonRelatedPart(kind string) (string, error) {
	want := resolver.relBase + kind
	wrongBase := relBaseStrict
	if resolver.relBase == relBaseStrict {
		wrongBase = relBaseTransitional
	}
	var part string
	for _, rel := range resolver.pkg.rels[resolver.mainPart] {
		if rel.Type == wrongBase+kind {
			return "", fmt.Errorf("docxpatch: native style resolution: %s relationship %q uses the wrong Strict/Transitional namespace", kind, rel.ID)
		}
		if rel.Type != want {
			continue
		}
		if rel.External || rel.PartName == "" {
			return "", fmt.Errorf("docxpatch: native style resolution: %s relationship %q must be internal", kind, rel.ID)
		}
		if part != "" {
			return "", fmt.Errorf("docxpatch: native style resolution: multiple %s relationships", kind)
		}
		part = rel.PartName
	}
	return part, nil
}

func (resolver *nativeLayoutResolver) loadStyles(partName string) error {
	root, err := parseNativeXML(partName, resolver.pkg.files[partName])
	if err != nil {
		return err
	}
	if root.Name != (xml.Name{Space: resolver.wordNS, Local: "styles"}) {
		return fmt.Errorf("docxpatch: native style resolution: styles part %q has spoofed or invalid root", partName)
	}
	if err := rejectNativeStyleNamespaceSpoofing(root, resolver.wordNS); err != nil {
		return fmt.Errorf("docxpatch: native style resolution: styles part %q: %w", partName, err)
	}
	if len(root.Children) > NativeDOCXMaxCollectionItems {
		return fmt.Errorf("docxpatch: native style resolution: styles part exceeds %d top-level entries", NativeDOCXMaxCollectionItems)
	}
	seenDocDefaults := false
	seenLatentStyles := false
	for _, child := range root.Children {
		if child.Name.Space != resolver.wordNS {
			resolver.addDiagnostic("FOREIGN_STYLES_MARKUP", resolver.doc.DocumentID, partName, child, "Foreign styles metadata is preserved and not interpreted")
			continue
		}
		switch child.Name.Local {
		case "docDefaults":
			if seenDocDefaults {
				return fmt.Errorf("docxpatch: native style resolution: duplicate docDefaults")
			}
			seenDocDefaults = true
			if len(directNativeChildren(child, resolver.wordNS, "rPrDefault")) > 1 || len(directNativeChildren(child, resolver.wordNS, "pPrDefault")) > 1 {
				return fmt.Errorf("docxpatch: native style resolution: ambiguous duplicate docDefaults properties")
			}
			if rDefault := firstDirectNativeChild(child, resolver.wordNS, "rPrDefault"); rDefault != nil {
				if len(directNativeChildren(rDefault, resolver.wordNS, "rPr")) > 1 {
					return fmt.Errorf("docxpatch: native style resolution: duplicate default run properties")
				}
				if rPr := firstDirectNativeChild(rDefault, resolver.wordNS, "rPr"); rPr != nil {
					resolver.docR = resolver.parseRunProperties(partName, rPr, resolver.doc.DocumentID)
				}
			}
			if pDefault := firstDirectNativeChild(child, resolver.wordNS, "pPrDefault"); pDefault != nil {
				if len(directNativeChildren(pDefault, resolver.wordNS, "pPr")) > 1 {
					return fmt.Errorf("docxpatch: native style resolution: duplicate default paragraph properties")
				}
				if pPr := firstDirectNativeChild(pDefault, resolver.wordNS, "pPr"); pPr != nil {
					resolver.docP = resolver.parseParagraphProperties(partName, pPr, resolver.doc.DocumentID)
				}
			}
		case "style":
			id, okID := nativeAttr(child, resolver.wordNS, "styleId")
			kind, okKind := nativeAttr(child, resolver.wordNS, "type")
			if !okID || !nativeIDPattern.MatchString(id) || !okKind {
				return fmt.Errorf("docxpatch: native style resolution: invalid style at %s", child.Path)
			}
			if kind == "numbering" {
				if nativeEmptyDefaultNumberingStyle(child, root, resolver.wordNS) {
					resolver.addDiagnostic("EMPTY_NUMBERING_STYLE_PRESERVED", resolver.doc.DocumentID, partName, child, "Default numbering style has only exact UI metadata and no layout properties; source remains preserved")
					continue
				}
				resolver.addDiagnostic("NUMBERING_STYLE_PRESERVED", resolver.doc.DocumentID, partName, child, "Numbering-style linking is preserved and not guessed")
				continue
			}
			if kind != "paragraph" && kind != "character" && kind != "table" {
				resolver.addDiagnostic("UNSUPPORTED_STYLE_TYPE", resolver.doc.DocumentID, partName, child, "This valid extension style type is preserved and not interpreted")
				continue
			}
			key := kind + "\x00" + id
			if _, duplicate := resolver.styles[key]; duplicate {
				return fmt.Errorf("docxpatch: native style resolution: duplicate %s style %q", kind, id)
			}
			definition := &nativeStyleDefinition{id: id, kind: kind, partName: partName, node: child}
			if len(directNativeChildren(child, resolver.wordNS, "basedOn")) > 1 || len(directNativeChildren(child, resolver.wordNS, "pPr")) > 1 || len(directNativeChildren(child, resolver.wordNS, "rPr")) > 1 {
				return fmt.Errorf("docxpatch: native style resolution: style %q has ambiguous duplicate properties", id)
			}
			if basedOn := firstDirectNativeChild(child, resolver.wordNS, "basedOn"); basedOn != nil {
				definition.basedOn, _ = nativeAttr(basedOn, resolver.wordNS, "val")
			}
			if pPr := firstDirectNativeChild(child, resolver.wordNS, "pPr"); pPr != nil {
				resolver.deferredNumbering = &definition.deferred
				definition.p = resolver.parseParagraphProperties(partName, pPr, resolver.doc.DocumentID)
				resolver.deferredNumbering = nil
			}
			if rPr := firstDirectNativeChild(child, resolver.wordNS, "rPr"); rPr != nil {
				resolver.deferredNumbering = &definition.deferred
				definition.r = resolver.parseRunProperties(partName, rPr, resolver.doc.DocumentID)
				resolver.deferredNumbering = nil
			}
			resolver.styles[key] = definition
			if len(resolver.styles) > NativeDOCXMaxCollectionItems {
				return fmt.Errorf("docxpatch: native style resolution: styles exceed %d entries", NativeDOCXMaxCollectionItems)
			}
			if value, present := nativeAttr(child, resolver.wordNS, "default"); present {
				isDefault, valid := nativeLexicalOnOff(value)
				if !valid {
					return fmt.Errorf("docxpatch: native style resolution: invalid default flag on style %q", id)
				}
				if isDefault {
					switch kind {
					case "paragraph":
						if resolver.defaultP != "" {
							return fmt.Errorf("docxpatch: native style resolution: multiple default paragraph styles")
						}
						resolver.defaultP = id
					case "character":
						if resolver.defaultC != "" {
							return fmt.Errorf("docxpatch: native style resolution: multiple default character styles")
						}
						resolver.defaultC = id
					}
				}
			}
		case "latentStyles":
			if seenLatentStyles {
				return fmt.Errorf("docxpatch: native style resolution: duplicate latentStyles")
			}
			seenLatentStyles = true
			if nativeLatentStyleBehaviorOnly(child, resolver.wordNS) {
				resolver.addDiagnostic("LATENT_STYLE_BEHAVIOR_PRESERVED", resolver.doc.DocumentID, partName, child, "Exact latent style UI/locking metadata is preserved; it supplies no formatting and does not authorize edits")
			} else {
				resolver.addDiagnostic("LATENT_STYLES_PRESERVED", resolver.doc.DocumentID, partName, child, "Unqualified latent style metadata is preserved and not interpreted")
			}
		default:
			resolver.addDiagnostic("UNMODELED_STYLES_MARKUP", resolver.doc.DocumentID, partName, child, "Styles metadata outside the conservative cascade is preserved verbatim")
		}
	}
	return nil
}

// ISO/IEC 29500-1 17.7.4.5: latent styles contain behavior, not formatting.
// Recognize only the exact metadata grammar; unknown extensions stay blocking.
// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.latentstyles
func nativeLatentStyleBehaviorOnly(node *nativeXMLNode, wordNS string) bool {
	check := func(n *nativeXMLNode, root bool) bool {
		if !nativeXMLWhitespaceOnly(n.Text) {
			return false
		}
		seen := map[xml.Name]bool{}
		for _, attr := range n.Attrs {
			if nativeSettingsNamespaceDeclaration(attr) {
				continue
			}
			if attr.Name.Space != wordNS || seen[attr.Name] {
				return false
			}
			seen[attr.Name] = true
			key := attr.Name.Local
			if root {
				switch key {
				case "defLockedState", "defSemiHidden", "defUnhideWhenUsed", "defQFormat":
					if _, valid := nativeLexicalOnOff(attr.Value); !valid {
						return false
					}
				case "count", "defUIPriority":
					if _, err := strconv.ParseUint(attr.Value, 10, 32); err != nil {
						return false
					}
				default:
					return false
				}
			} else {
				switch key {
				case "locked", "semiHidden", "unhideWhenUsed", "qFormat":
					if _, valid := nativeLexicalOnOff(attr.Value); !valid {
						return false
					}
				case "uiPriority":
					if _, err := strconv.ParseUint(attr.Value, 10, 32); err != nil {
						return false
					}
				case "name":
					if strings.TrimSpace(attr.Value) == "" {
						return false
					}
				default:
					return false
				}
			}
		}
		return root || seen[xml.Name{Space: wordNS, Local: "name"}]
	}
	if !check(node, true) || len(node.Children) > NativeDOCXMaxCollectionItems {
		return false
	}
	for _, child := range node.Children {
		if child.Name != (xml.Name{Space: wordNS, Local: "lsdException"}) || len(child.Children) != 0 || !check(child, false) {
			return false
		}
	}
	return true
}

func (resolver *nativeLayoutResolver) loadFonts(partName string) error {
	root, err := parseNativeXML(partName, resolver.pkg.files[partName])
	if err != nil {
		return err
	}
	if root.Name != (xml.Name{Space: resolver.wordNS, Local: "fonts"}) {
		return fmt.Errorf("docxpatch: native style resolution: font table %q has spoofed or invalid root", partName)
	}
	if err := rejectNativeKnownLocalSpoofing(root, resolver.wordNS, map[string]bool{
		"fonts": true, "font": true, "altName": true,
		"embedRegular": true, "embedBold": true, "embedItalic": true, "embedBoldItalic": true,
	}); err != nil {
		return fmt.Errorf("docxpatch: native style resolution: font table %q: %w", partName, err)
	}
	if len(root.Children) > NativeDOCXMaxCollectionItems {
		return fmt.Errorf("docxpatch: native style resolution: font table exceeds %d top-level entries", NativeDOCXMaxCollectionItems)
	}
	attestedEmbedPaths := resolver.attestedEmbeddedFontPaths(partName)
	seen := map[string]bool{}
	for _, child := range root.Children {
		if child.Name.Space != resolver.wordNS {
			if child.Name.Local == "font" {
				return fmt.Errorf("docxpatch: native style resolution: namespace spoofing at %s", child.Path)
			}
			resolver.addDiagnostic("FOREIGN_FONT_TABLE_MARKUP", resolver.doc.DocumentID, partName, child, "Foreign font-table metadata is preserved and not interpreted")
			continue
		}
		if child.Name.Local != "font" {
			resolver.addDiagnostic("UNMODELED_FONT_TABLE_MARKUP", resolver.doc.DocumentID, partName, child, "Font-table metadata outside w:font is preserved")
			continue
		}
		name, ok := nativeAttr(child, resolver.wordNS, "name")
		if !ok || name == "" || len(name) > 256 {
			return fmt.Errorf("docxpatch: native style resolution: invalid font at %s", child.Path)
		}
		key := strings.ToLower(name)
		if seen[key] {
			return fmt.Errorf("docxpatch: native style resolution: duplicate font %q", name)
		}
		seen[key] = true
		font := NativeResolvedFontV1{Name: name}
		altNames := directNativeChildren(child, resolver.wordNS, "altName")
		if len(altNames) > 1 {
			resolver.addDiagnostic("DUPLICATE_FONT_ALT_NAME", resolver.doc.DocumentID, partName, child, "Duplicate font aliases are ambiguous and were not resolved")
		} else if len(altNames) == 1 {
			alt := altNames[0]
			if value, present := nativeAttr(alt, resolver.wordNS, "val"); present {
				if value != "" && nativeBoundedResolvedString(value, 256) {
					font.AltName = nativeString(value)
				} else {
					resolver.addDiagnostic("INVALID_FONT_ALT_NAME", resolver.doc.DocumentID, partName, alt, "Invalid font alias is preserved and ignored")
				}
			}
		}
		for _, property := range child.Children {
			if nativeQualifiedFontTableOwner(root) && nativeQualifiedFontDescriptor(property, child, resolver.wordNS) {
				resolver.addDiagnostic("FONT_MATCHING_METADATA_PRESERVED", resolver.doc.DocumentID, partName, property, "Validated font matching metadata is preserved; native painting requires exact supplied faces, not metadata-driven substitution")
				continue
			}
			if nativeQualifiedFontTableOwner(root) && nativePotentiallyUnusedFontDescriptor(property, child, resolver.wordNS) {
				if len(resolver.unusedFontDescriptors) >= NativeDOCXMaxResolvedDiagnostics {
					return fmt.Errorf("docxpatch: font descriptor qualification exceeds bounded diagnostic budget")
				}
				resolver.unusedFontDescriptors = append(resolver.unusedFontDescriptors, nativeUnusedFontDescriptor{name: name, alias: font.AltName, part: partName, node: property})
				continue
			}
			if property.Name.Space == resolver.wordNS && property.Name.Local == "altName" {
				continue
			}
			if property.Name.Space == resolver.wordNS && (property.Name.Local == "embedRegular" || property.Name.Local == "embedBold" || property.Name.Local == "embedItalic" || property.Name.Local == "embedBoldItalic") {
				// Embedded-resource bindings do not alter the resolved family/style
				// cascade, but suppressing the diagnostic is safe only after the same
				// package graph/license parser used by the canonical inventory proves
				// this exact element path.
				if attestedEmbedPaths[property.Path] {
					continue
				}
				resolver.addDiagnostic("UNATTESTED_EMBEDDED_FONT_BINDING", resolver.doc.DocumentID, partName, property, "Embedded font metadata lacks an exact package relationship, resource, digest, face-slot, and license attestation")
				continue
			}
			if property.Name != (xml.Name{Space: resolver.wordNS, Local: "altName"}) {
				resolver.addDiagnostic("UNMODELED_FONT_METADATA", resolver.doc.DocumentID, partName, property, "Font metadata is preserved for future font matching")
			}
		}
		resolver.fonts = append(resolver.fonts, font)
		if len(resolver.fonts) > NativeDOCXMaxCollectionItems {
			return fmt.Errorf("docxpatch: native style resolution: font table exceeds %d entries", NativeDOCXMaxCollectionItems)
		}
	}
	return nil
}

func (resolver *nativeLayoutResolver) attestedEmbeddedFontPaths(partName string) map[string]bool {
	paths := map[string]bool{}
	fontTableRel, err := nativeDOCXSingletonFontTable(resolver.pkg, resolver.mainPart, resolver.relBase)
	if err != nil || fontTableRel == nil || fontTableRel.PartName != partName {
		return paths
	}
	mainRelsPart := resolver.pkg.relsPart[resolver.mainPart]
	if mainRelsPart == "" {
		return paths
	}
	binding := &NativeDOCXFontTableBindingV1{
		PartName: partName, SHA256: nativeSHA(resolver.pkg.files[partName]),
		MainRelationshipsPart: mainRelsPart, MainRelationshipsSHA256: nativeSHA(resolver.pkg.files[mainRelsPart]),
		RelationshipID: fontTableRel.ID, RelationshipType: fontTableRel.Type, RelationshipTarget: fontTableRel.Target,
	}
	if relsPart := resolver.pkg.relsPart[partName]; relsPart != "" {
		binding.FontRelationshipsPart = nativeString(relsPart)
		binding.FontRelationshipsSHA256 = nativeString(nativeSHA(resolver.pkg.files[relsPart]))
	}
	relNS := relNSTransitional
	if resolver.wordNS == wordMLStrict {
		relNS = relNSStrict
	}
	families, err := nativeDOCXInventoryFamilies(resolver.pkg, partName, resolver.wordNS, relNS, resolver.relBase, binding)
	if err != nil {
		return paths
	}
	resources, assetParts, relationshipIDs, faceIDs := map[string]bool{}, map[string]bool{}, map[string]bool{}, map[string]bool{}
	for _, family := range families {
		for _, face := range family.Faces {
			assetKey := nativeASCIIFold(face.Source.AssetPart)
			if resources[face.Source.ResourceID] || assetParts[assetKey] || relationshipIDs[face.Source.RelationshipID] || faceIDs[face.FaceID] {
				return map[string]bool{}
			}
			resources[face.Source.ResourceID], assetParts[assetKey], relationshipIDs[face.Source.RelationshipID], faceIDs[face.FaceID] = true, true, true, true
			paths[face.Source.FontTablePath] = true
		}
	}
	return paths
}

func (resolver *nativeLayoutResolver) loadTheme(partName string) error {
	root, err := parseNativeXML(partName, resolver.pkg.files[partName])
	if err != nil {
		return err
	}
	drawingNS := drawingMLTransitional
	if resolver.wordNS == wordMLStrict {
		drawingNS = drawingMLStrict
	}
	if root.Name != (xml.Name{Space: drawingNS, Local: "theme"}) {
		return fmt.Errorf("docxpatch: native style resolution: theme part %q has spoofed or invalid root", partName)
	}
	// Exact a:srgbClr slots and sysClr lastClr snapshots project into run RGB.
	// schemeClr, tint/shade, and other DrawingML transforms stay unresolved
	// and are diagnosed at the referencing w:color.
	resolver.themeSrgb = nativeParseThemeSrgbColors(root, drawingNS)
	resolver.themeLatinFonts = nativeParseThemeLatinFonts(root, drawingNS)
	return nil
}

func (resolver *nativeLayoutResolver) resolveThemeSrgb(themeColor string) (string, bool) {
	slot, ok := nativeThemeColorSlot(themeColor)
	if !ok {
		return "", false
	}
	rgb, ok := resolver.themeSrgb[slot]
	return rgb, ok
}

func (resolver *nativeLayoutResolver) resolveExactRunColor(node *nativeXMLNode, scopeID, partName string) (*string, bool) {
	if !nativeWordColorLeaf(node, resolver.wordNS) {
		resolver.addDiagnostic("THEME_COLOR_PRESERVED", scopeID, partName, node, "Run color markup outside the exact RGB or theme-srgb subset is preserved and not guessed")
		return nil, false
	}
	if _, tint := nativeAttr(node, resolver.wordNS, "themeTint"); tint {
		resolver.addDiagnostic("THEME_COLOR_PRESERVED", scopeID, partName, node, "Theme tint/shade transforms require presentation context and are not guessed")
		return nil, false
	}
	if _, shade := nativeAttr(node, resolver.wordNS, "themeShade"); shade {
		resolver.addDiagnostic("THEME_COLOR_PRESERVED", scopeID, partName, node, "Theme tint/shade transforms require presentation context and are not guessed")
		return nil, false
	}
	theme, hasTheme := nativeAttr(node, resolver.wordNS, "themeColor")
	if hasTheme {
		rgb, ok := resolver.resolveThemeSrgb(theme)
		if !ok {
			resolver.addDiagnostic("THEME_COLOR_PRESERVED", scopeID, partName, node, "Theme/automatic color requires an exact a:srgbClr or sysClr lastClr snapshot and is not guessed")
			return nil, false
		}
		return nativeString(rgb), true
	}
	value, ok := nativeAttr(node, resolver.wordNS, "val")
	if !ok {
		resolver.addDiagnostic("INVALID_COLOR", scopeID, partName, node, "Invalid color is preserved and ignored")
		return nil, false
	}
	if strings.EqualFold(value, "auto") {
		resolver.addDiagnostic("THEME_COLOR_PRESERVED", scopeID, partName, node, "Automatic color requires presentation context and is not guessed")
		return nil, false
	}
	if rgb, ok := nativeExactRGB(value); ok {
		return nativeString(rgb), true
	}
	resolver.addDiagnostic("INVALID_COLOR", scopeID, partName, node, "Invalid color is preserved and ignored")
	return nil, false
}

func (resolver *nativeLayoutResolver) loadNumbering(partName string) error {
	root, err := parseNativeXML(partName, resolver.pkg.files[partName])
	if err != nil {
		return err
	}
	if root.Name != (xml.Name{Space: resolver.wordNS, Local: "numbering"}) {
		return fmt.Errorf("docxpatch: native style resolution: numbering part %q has spoofed or invalid root", partName)
	}
	numberingKnownLocals := map[string]bool{
		"numbering": true, "numPicBullet": true, "abstractNum": true, "num": true,
		"lvl": true, "abstractNumId": true, "lvlOverride": true, "startOverride": true,
		"start": true, "numFmt": true, "lvlText": true, "suff": true, "lvlJc": true,
		"pPr": true, "rPr": true, "lvlPicBulletId": true, "lvlRestart": true,
		"nsid": true, "multiLevelType": true, "tmpl": true, "pStyle": true, "isLgl": true, "legacy": true,
		"tabs": true, "tab": true,
	}
	for _, child := range root.Children {
		if child.Name.Space != resolver.wordNS {
			continue
		}
		if err := rejectNativeKnownLocalSpoofing(child, resolver.wordNS, numberingKnownLocals); err != nil {
			return fmt.Errorf("docxpatch: native style resolution: numbering part %q: %w", partName, err)
		}
	}
	if len(root.Children) > NativeDOCXMaxCollectionItems {
		return fmt.Errorf("docxpatch: native style resolution: numbering part exceeds %d top-level entries", NativeDOCXMaxCollectionItems)
	}
	for _, child := range root.Children {
		if child.Name.Space != resolver.wordNS {
			resolver.numberingRootDeferred = append(resolver.numberingRootDeferred, nativeDeferredNumberingDiagnostic{code: "FOREIGN_NUMBERING_ROOT", partName: partName, node: child, message: "Foreign numbering-root semantics are preserved and deferred until a concrete numId is referenced"})
			continue
		}
		switch child.Name.Local {
		case "numPicBullet":
			// The concrete lvlPicBulletId reference carries the refusal. Merely
			// declaring an unused picture-bullet resource is not a document semantic.
		case "abstractNum":
			id, ok := nativeDecimalIDAttr(child, resolver.wordNS, "abstractNumId")
			if !ok {
				return fmt.Errorf("docxpatch: native style resolution: invalid abstractNum id at %s", child.Path)
			}
			if _, duplicate := resolver.abstractNums[id]; duplicate {
				return fmt.Errorf("docxpatch: native style resolution: duplicate abstractNum id %q", id)
			}
			abstract := &nativeAbstractNumbering{id: id, levels: map[int]*nativeNumberingLevel{}}
			previousDeferred := resolver.deferredNumbering
			resolver.deferredNumbering = &abstract.deferred
			for _, property := range child.Children {
				if property.Name.Space != resolver.wordNS {
					resolver.addDiagnostic("FOREIGN_NUMBERING_MARKUP", resolver.doc.DocumentID, partName, property, "Foreign abstract numbering metadata is preserved")
					continue
				}
				if property.Name.Local == "nsid" || property.Name.Local == "tmpl" {
					// Stable authoring/template metadata does not alter counter or marker semantics.
					continue
				}
				if property.Name.Local == "multiLevelType" {
					if abstract.multiSeen {
						return fmt.Errorf("docxpatch: native style resolution: duplicate multiLevelType in abstractNum %q", id)
					}
					abstract.multiSeen = true
					value, present := nativeAttr(property, resolver.wordNS, "val")
					if !present || (value != "singleLevel" && value != "multilevel" && value != "hybridMultilevel") {
						return fmt.Errorf("docxpatch: native style resolution: invalid multiLevelType in abstractNum %q", id)
					}
					continue
				}
				if property.Name.Local != "lvl" {
					resolver.addDiagnostic("UNMODELED_ABSTRACT_NUMBERING", resolver.doc.DocumentID, partName, property, "Abstract numbering metadata outside ordinary levels is preserved")
					continue
				}
				level, err := resolver.parseNumberingLevel(partName, property, resolver.doc.DocumentID)
				if err != nil {
					return err
				}
				if _, duplicate := abstract.levels[level.level]; duplicate {
					return fmt.Errorf("docxpatch: native style resolution: duplicate level %d in abstractNum %q", level.level, id)
				}
				abstract.levels[level.level] = level
			}
			resolver.deferredNumbering = previousDeferred
			resolver.abstractNums[id] = abstract
		case "num":
			id, ok := nativeDecimalIDAttr(child, resolver.wordNS, "numId")
			if !ok {
				return fmt.Errorf("docxpatch: native style resolution: invalid num id at %s", child.Path)
			}
			if _, duplicate := resolver.nums[id]; duplicate {
				return fmt.Errorf("docxpatch: native style resolution: duplicate num id %q", id)
			}
			instance := &nativeNumberingInstance{id: id, overrides: map[int]*nativeNumberingOverride{}}
			previousDeferred := resolver.deferredNumbering
			resolver.deferredNumbering = &instance.deferred
			for _, property := range child.Children {
				if property.Name.Space != resolver.wordNS {
					resolver.addDiagnostic("FOREIGN_NUMBERING_MARKUP", resolver.doc.DocumentID, partName, property, "Foreign numbering-instance metadata is preserved")
					continue
				}
				switch property.Name.Local {
				case "abstractNumId":
					if instance.abstractID != "" {
						return fmt.Errorf("docxpatch: native style resolution: duplicate abstractNumId in num %q", id)
					}
					value, valid := nativeDecimalIDAttr(property, resolver.wordNS, "val")
					if !valid {
						return fmt.Errorf("docxpatch: native style resolution: invalid abstractNumId in num %q", id)
					}
					instance.abstractID = value
				case "lvlOverride":
					level, valid := nativeNumberingLevelAttr(property, resolver.wordNS, "ilvl")
					if !valid {
						return fmt.Errorf("docxpatch: native style resolution: invalid level override in num %q", id)
					}
					if _, duplicate := instance.overrides[level]; duplicate {
						return fmt.Errorf("docxpatch: native style resolution: duplicate level override %d in num %q", level, id)
					}
					override := &nativeNumberingOverride{}
					instanceDeferred := resolver.deferredNumbering
					resolver.deferredNumbering = &override.deferred
					seenStart, seenLevel := false, false
					for _, overrideChild := range property.Children {
						switch {
						case overrideChild.Name == (xml.Name{Space: resolver.wordNS, Local: "startOverride"}):
							if seenStart {
								return fmt.Errorf("docxpatch: native style resolution: duplicate startOverride in num %q", id)
							}
							seenStart = true
							value, ok := nativeNonnegativeIntAttr(overrideChild, resolver.wordNS, "val")
							if !ok {
								return fmt.Errorf("docxpatch: native style resolution: invalid startOverride in num %q", id)
							}
							override.start = nativeInt(value)
						case overrideChild.Name == (xml.Name{Space: resolver.wordNS, Local: "lvl"}):
							if seenLevel {
								return fmt.Errorf("docxpatch: native style resolution: duplicate replacement level in num %q", id)
							}
							seenLevel = true
							replacement, err := resolver.parseNumberingLevel(partName, overrideChild, resolver.doc.DocumentID)
							if err != nil || replacement.level != level {
								return fmt.Errorf("docxpatch: native style resolution: invalid replacement level in num %q", id)
							}
							override.level = replacement
						default:
							resolver.addDiagnostic("UNMODELED_NUMBERING_OVERRIDE", resolver.doc.DocumentID, partName, overrideChild, "Numbering override metadata is preserved and not interpreted")
						}
					}
					resolver.deferredNumbering = instanceDeferred
					instance.overrides[level] = override
				default:
					resolver.addDiagnostic("UNMODELED_NUMBERING_INSTANCE", resolver.doc.DocumentID, partName, property, "Numbering-instance metadata is preserved and not interpreted")
				}
			}
			if instance.abstractID == "" {
				return fmt.Errorf("docxpatch: native style resolution: num %q has no abstractNumId", id)
			}
			resolver.deferredNumbering = previousDeferred
			resolver.nums[id] = instance
		default:
			resolver.numberingRootDeferred = append(resolver.numberingRootDeferred, nativeDeferredNumberingDiagnostic{code: "UNMODELED_NUMBERING_ROOT", partName: partName, node: child, message: "Unknown numbering-root semantics are preserved and deferred until a concrete numId is referenced"})
		}
		if len(resolver.abstractNums)+len(resolver.nums) > NativeDOCXMaxCollectionItems {
			return fmt.Errorf("docxpatch: native style resolution: numbering definitions exceed %d entries", NativeDOCXMaxCollectionItems)
		}
	}
	return nil
}

func (resolver *nativeLayoutResolver) parseNumberingLevel(partName string, node *nativeXMLNode, scopeID string) (*nativeNumberingLevel, error) {
	levelValue, ok := nativeNumberingLevelAttr(node, resolver.wordNS, "ilvl")
	if !ok {
		return nil, fmt.Errorf("docxpatch: native style resolution: invalid numbering level at %s", node.Path)
	}
	level := &nativeNumberingLevel{level: levelValue, partName: partName, node: node}
	previousDeferred := resolver.deferredNumbering
	resolver.deferredNumbering = &level.deferred
	defer func() { resolver.deferredNumbering = previousDeferred }()
	seen := map[string]bool{}
	for _, child := range node.Children {
		if child.Name.Space != resolver.wordNS {
			resolver.addDiagnostic("FOREIGN_NUMBERING_LEVEL", scopeID, partName, child, "Foreign numbering-level metadata is preserved")
			continue
		}
		if child.Name.Local == "start" || child.Name.Local == "numFmt" || child.Name.Local == "lvlText" || child.Name.Local == "suff" || child.Name.Local == "lvlJc" || child.Name.Local == "pPr" || child.Name.Local == "rPr" || child.Name.Local == "lvlPicBulletId" || child.Name.Local == "lvlRestart" {
			if seen[child.Name.Local] {
				return nil, fmt.Errorf("docxpatch: native style resolution: duplicate %s in numbering level %d", child.Name.Local, levelValue)
			}
			seen[child.Name.Local] = true
		}
		switch child.Name.Local {
		case "start":
			if value, valid := nativeNonnegativeIntAttr(child, resolver.wordNS, "val"); valid {
				level.start = nativeInt(value)
			} else {
				return nil, fmt.Errorf("docxpatch: native style resolution: invalid level start at %s", child.Path)
			}
		case "numFmt":
			if _, custom := nativeAttr(child, resolver.wordNS, "format"); custom {
				resolver.addDiagnostic("CUSTOM_NUMBER_FORMAT", scopeID, partName, child, "Custom XSLT numbering formats are preserved and refused")
			}
			if value, valid := nativeAttr(child, resolver.wordNS, "val"); valid && value != "" {
				level.format = nativeString(value)
			} else {
				return nil, fmt.Errorf("docxpatch: native style resolution: invalid numFmt at %s", child.Path)
			}
		case "lvlText":
			if raw, present := nativeAttr(child, resolver.wordNS, "null"); present {
				isNull, valid := nativeLexicalOnOff(raw)
				if !valid || isNull {
					return nil, fmt.Errorf("docxpatch: native style resolution: null or invalid lvlText at %s", child.Path)
				}
			}
			if value, valid := nativeAttr(child, resolver.wordNS, "val"); valid && len(value) <= 1024 {
				level.text = nativeString(value)
			} else {
				return nil, fmt.Errorf("docxpatch: native style resolution: invalid level text at %s", child.Path)
			}
		case "suff":
			value, valid := nativeAttr(child, resolver.wordNS, "val")
			if valid && (value == "tab" || value == "space" || value == "nothing") {
				level.suffix = nativeString(value)
			} else {
				resolver.addDiagnostic("UNSUPPORTED_NUMBER_SUFFIX", scopeID, partName, child, "This numbering suffix is preserved and not guessed")
			}
		case "lvlJc":
			value, valid := nativeAttr(child, resolver.wordNS, "val")
			if valid && nativeParagraphAlignment(value) {
				level.alignment = nativeString(value)
			} else {
				level.alignment = nativeString("invalid")
				resolver.addDiagnostic("UNSUPPORTED_NUMBER_ALIGNMENT", scopeID, partName, child, "This numbering alignment is preserved and not guessed")
			}
		case "pPr":
			properties, numTab, err := resolver.parseNumberingParagraphProperties(partName, child, scopeID)
			if err != nil {
				return nil, err
			}
			level.p, level.numTab = properties, numTab
		case "rPr":
			level.r = resolver.parseRunProperties(partName, child, scopeID)
		case "lvlPicBulletId":
			level.picture = true
		case "lvlRestart":
			value, valid := nativeNonnegativeIntAttr(child, resolver.wordNS, "val")
			if !valid || value > 7 {
				return nil, fmt.Errorf("docxpatch: native style resolution: invalid lvlRestart at %s", child.Path)
			}
			level.restart = nativeInt(value)
		case "pStyle":
			level.styleLinkCount++
			value, valid := nativeAttr(child, resolver.wordNS, "val")
			if !valid || !nativeIDPattern.MatchString(value) {
				resolver.addDiagnostic("INVALID_NUMBERING_STYLE_LINK", scopeID, partName, child, "Invalid numbering-level paragraph-style link is preserved and not guessed")
				continue
			}
			level.styleLinks = append(level.styleLinks, value)
		case "isLgl", "legacy":
			resolver.addDiagnostic("UNMODELED_NUMBERING_LEVEL", scopeID, partName, child, "This numbering-level semantic is preserved and not guessed")
		default:
			resolver.addDiagnostic("UNMODELED_NUMBERING_LEVEL", scopeID, partName, child, "This numbering-level metadata is preserved and not interpreted")
		}
	}
	return level, nil
}

func (resolver *nativeLayoutResolver) parseNumberingParagraphProperties(partName string, node *nativeXMLNode, scopeID string) (nativeParagraphProperties, *int64, error) {
	copyNode := *node
	copyNode.Children = make([]*nativeXMLNode, 0, len(node.Children))
	var tabs *nativeXMLNode
	for _, child := range node.Children {
		if child.Name == (xml.Name{Space: resolver.wordNS, Local: "tabs"}) {
			if tabs != nil {
				return nativeParagraphProperties{}, nil, fmt.Errorf("docxpatch: native style resolution: duplicate numbering tabs at %s", node.Path)
			}
			tabs = child
			continue
		}
		copyNode.Children = append(copyNode.Children, child)
	}
	properties := resolver.parseParagraphProperties(partName, &copyNode, scopeID)
	if tabs == nil {
		return properties, nil, nil
	}
	if len(tabs.Children) != 1 || tabs.Children[0].Name != (xml.Name{Space: resolver.wordNS, Local: "tab"}) {
		return nativeParagraphProperties{}, nil, fmt.Errorf("docxpatch: native style resolution: numbering tabs must contain exactly one Word num tab at %s", tabs.Path)
	}
	tab := tabs.Children[0]
	value, present := nativeAttr(tab, resolver.wordNS, "val")
	position, valid := nativeNonnegativeInt64Attr(tab, resolver.wordNS, "pos")
	if !present || value != "num" || !valid || position > nativeMaxTwipsForMilliPoints {
		return nativeParagraphProperties{}, nil, fmt.Errorf("docxpatch: native style resolution: invalid numbering num tab at %s", tab.Path)
	}
	return properties, nativeInt64(position), nil
}

func nativeDecimalIDAttr(node *nativeXMLNode, namespace, local string) (string, bool) {
	raw, ok := nativeAttr(node, namespace, local)
	if !ok || raw == "" {
		return "", false
	}
	return nativeCanonicalDecimalID(raw)
}

func nativeNumberingLevelAttr(node *nativeXMLNode, namespace, local string) (int, bool) {
	raw, ok := nativeAttr(node, namespace, local)
	if !ok {
		return 0, false
	}
	value, err := strconv.Atoi(raw)
	return value, err == nil && value >= 0 && value <= 8
}

func nativeNonnegativeIntAttr(node *nativeXMLNode, namespace, local string) (int, bool) {
	raw, ok := nativeAttr(node, namespace, local)
	if !ok {
		return 0, false
	}
	value, err := strconv.ParseInt(raw, 10, 32)
	return int(value), err == nil && value >= 0
}

func nativeSignedInt64Attr(node *nativeXMLNode, namespace, local string) (int64, bool) {
	raw, ok := nativeAttr(node, namespace, local)
	if !ok {
		return 0, false
	}
	value, err := strconv.ParseInt(raw, 10, 64)
	return value, err == nil && value >= -9007199254740991 && value <= 9007199254740991
}

func (resolver *nativeLayoutResolver) resolve() (*NativeResolvedLayoutInputV1, error) {
	result := &NativeResolvedLayoutInputV1{
		Protocol:        NativeDOCXResolvedLayoutProtocol,
		Version:         NativeDOCXResolvedLayoutVersion,
		DocumentID:      resolver.doc.DocumentID,
		Revision:        resolver.doc.Revision,
		SourceParts:     resolver.parts,
		NumberingSource: resolver.numberingSource,
		Paragraphs:      []NativeResolvedParagraphV1{}, Runs: []NativeResolvedRunV1{}, Tables: []NativeResolvedTableV1{},
		Fonts: append([]NativeResolvedFontV1{}, resolver.fonts...), Diagnostics: []NativeResolutionDiagnosticV1{},
	}
	for _, story := range resolver.allStories() {
		numberingState := newNativeNumberingState()
		for index := range story.Blocks {
			resolver.resolveBlock(&story.Blocks[index], result, numberingState)
		}
	}
	resolver.resolveUnusedFontDescriptors(result)
	if resolver.diagnosticOverflow {
		return nil, fmt.Errorf("docxpatch: native style resolution: diagnostics exceed %d entries", NativeDOCXMaxResolvedDiagnostics)
	}
	result.Diagnostics = append(result.Diagnostics, resolver.diagnostics...)
	if result.NumberingSource != nil {
		result.NumberingSource.ModelSHA256 = nativeResolvedNumberingModelSHA256(result)
	}
	return result, nil
}

func (resolver *nativeLayoutResolver) allStories() []NativeStoryV1 {
	stories := []NativeStoryV1{resolver.doc.Body}
	stories = append(stories, resolver.doc.Headers...)
	stories = append(stories, resolver.doc.Footers...)
	stories = append(stories, resolver.doc.Notes...)
	stories = append(stories, resolver.doc.CommentStories...)
	return stories
}

func (resolver *nativeLayoutResolver) resolveBlock(block *NativeBlockV1, result *NativeResolvedLayoutInputV1, numberingState *nativeNumberingState) {
	if block.Paragraph != nil {
		resolver.resolveParagraph(block.Paragraph, result, numberingState, nil)
		return
	}
	if block.Table == nil {
		return
	}
	table := block.Table
	resolvedTable, tableStyles := resolver.resolveTableStyle(table)
	result.Tables = append(result.Tables, resolvedTable)
	for rowIndex := range table.Rows {
		for cellIndex := range table.Rows[rowIndex].Cells {
			for paragraphIndex := range table.Rows[rowIndex].Cells[cellIndex].Paragraphs {
				resolver.resolveParagraph(&table.Rows[rowIndex].Cells[cellIndex].Paragraphs[paragraphIndex], result, numberingState, tableStyles)
			}
		}
	}
}

func mergeNativeTableBorders(base, overlay *NativeTableBordersV1) *NativeTableBordersV1 {
	if overlay == nil {
		return base
	}
	if base == nil {
		copied := *overlay
		return &copied
	}
	out := *base
	if overlay.Top != nil {
		out.Top = overlay.Top
	}
	if overlay.Right != nil {
		out.Right = overlay.Right
	}
	if overlay.Bottom != nil {
		out.Bottom = overlay.Bottom
	}
	if overlay.Left != nil {
		out.Left = overlay.Left
	}
	if overlay.InsideHorizontal != nil {
		out.InsideHorizontal = overlay.InsideHorizontal
	}
	if overlay.InsideVertical != nil {
		out.InsideVertical = overlay.InsideVertical
	}
	return &out
}

func nativeTableStyleAuthoringLocal(local string) bool {
	switch local {
	case "name", "basedOn", "next", "link", "autoRedefine", "hidden", "uiPriority", "semiHidden", "unhideWhenUsed", "qFormat", "locked", "personal", "personalCompose", "personalReply", "rsid":
		return true
	default:
		return false
	}
}

func (resolver *nativeLayoutResolver) parseSimpleTableStyleBorders(node *nativeXMLNode) (*NativeTableBordersV1, bool) {
	if node == nil {
		return nil, true
	}
	if !nativeExactContainer(node) {
		return nil, false
	}
	var borders *NativeTableBordersV1
	for _, child := range node.Children {
		if child.Name.Space != resolver.wordNS {
			return nil, false
		}
		switch child.Name.Local {
		case "tblBorders":
			parsed, ok := nativeExtractTableBorders(child, resolver.wordNS, resolver.resolveThemeSrgb)
			if !ok {
				return nil, false
			}
			borders = parsed
		case "tblW", "tblLayout", "jc", "tblInd", "tblCellMar", "tblLook", "shd":
			// Geometry and look metadata do not map onto the existing border/fill commands.
		default:
			return nil, false
		}
	}
	return borders, true
}

func (resolver *nativeLayoutResolver) parseSimpleTableStyleCellFill(node *nativeXMLNode) (*string, bool) {
	if node == nil {
		return nil, true
	}
	if !nativeExactContainer(node) {
		return nil, false
	}
	var fill *string
	for _, child := range node.Children {
		if child.Name.Space != resolver.wordNS {
			return nil, false
		}
		if child.Name.Local != "shd" {
			return nil, false
		}
		parsed, ok := nativeExtractCellShading(child, resolver.wordNS, resolver.resolveThemeSrgb)
		if !ok {
			return nil, false
		}
		fill = parsed
	}
	return fill, true
}

func (resolver *nativeLayoutResolver) resolveTableStyle(table *NativeTableV1) (NativeResolvedTableV1, []*nativeStyleDefinition) {
	resolved := NativeResolvedTableV1{TableID: table.ID, StyleID: table.TableStyleID}
	if table.TableStyleID == nil {
		return resolved, nil
	}
	definition := resolver.styles["table\x00"+*table.TableStyleID]
	if definition == nil {
		resolver.addDiagnostic("MISSING_TABLE_STYLE", table.ID, resolver.partsValue(resolver.parts.StylesPart), nil, "The referenced table style is missing and was not guessed")
		return resolved, nil
	}
	chain := resolver.styleChain("table", *table.TableStyleID, table.ID)
	if len(chain) == 0 {
		return resolved, nil
	}
	simple := true
	for _, layer := range chain {
		if firstDirectNativeChild(layer.node, resolver.wordNS, "tblStylePr") != nil {
			resolver.addDiagnostic("CONDITIONAL_TABLE_STYLE_PRESERVED", table.ID, layer.partName, layer.node, "Conditional table-style semantics require table-region evaluation and are not guessed")
			resolver.addDiagnostic("TABLE_STYLE_EFFECTS_PRESERVED", table.ID, layer.partName, layer.node, "Table-style effects are preserved until table-region cascade support is implemented")
			simple = false
		}
		for _, child := range layer.node.Children {
			if child.Name.Space != resolver.wordNS {
				continue
			}
			switch child.Name.Local {
			case "tblPr":
				borders, ok := resolver.parseSimpleTableStyleBorders(child)
				if !ok {
					resolver.addDiagnostic("TABLE_STYLE_EFFECTS_PRESERVED", table.ID, layer.partName, child, "Table-style properties outside exact border/fill commands are preserved and not guessed")
					simple = false
					continue
				}
				if simple {
					resolved.Borders = mergeNativeTableBorders(resolved.Borders, borders)
				}
			case "tcPr":
				fill, ok := resolver.parseSimpleTableStyleCellFill(child)
				if !ok {
					resolver.addDiagnostic("TABLE_STYLE_EFFECTS_PRESERVED", table.ID, layer.partName, child, "Table-style cell properties outside exact clear fills are preserved and not guessed")
					simple = false
					continue
				}
				if simple && fill != nil {
					resolved.CellShadingRGB = fill
				}
			case "pPr", "rPr":
				// Parsed property layers apply after document defaults, before
				// paragraph/character/direct formatting in each cell paragraph.
			case "trPr", "tblStylePr":
				if child.Name.Local != "tblStylePr" {
					resolver.addDiagnostic("TABLE_STYLE_EFFECTS_PRESERVED", table.ID, layer.partName, child, "Table-style row effects are preserved until table-region cascade support is implemented")
					simple = false
				}
			default:
				if !nativeTableStyleAuthoringLocal(child.Name.Local) {
					resolver.addDiagnostic("TABLE_STYLE_EFFECTS_PRESERVED", table.ID, layer.partName, child, "This table-style layer is preserved and not guessed")
					simple = false
				}
			}
		}
	}
	if !simple {
		resolved.Borders = nil
		resolved.CellShadingRGB = nil
		return resolved, nil
	}
	return resolved, chain
}

func (resolver *nativeLayoutResolver) resolveParagraph(paragraph *NativeParagraphV1, result *NativeResolvedLayoutInputV1, numberingState *nativeNumberingState, tableStyles []*nativeStyleDefinition) {
	paragraphNode := resolver.nodeForAnchor(paragraph.Anchor)
	styleID := ""
	var directPPr *nativeXMLNode
	if paragraphNode != nil {
		pPrNodes := directNativeChildren(paragraphNode, resolver.wordNS, "pPr")
		switch len(pPrNodes) {
		case 0:
			styleID = resolver.defaultP
		case 1:
			directPPr = pPrNodes[0]
			styleNodes := directNativeChildren(directPPr, resolver.wordNS, "pStyle")
			switch len(styleNodes) {
			case 0:
				styleID = resolver.defaultP
			case 1:
				if value, ok := nativeAttr(styleNodes[0], resolver.wordNS, "val"); ok && nativeIDPattern.MatchString(value) {
					styleID = value
				} else {
					resolver.addDiagnostic("INVALID_PARAGRAPH_STYLE", paragraph.ID, paragraph.Anchor.PartName, styleNodes[0], "Invalid paragraph style reference is preserved and not resolved")
				}
			default:
				resolver.addDiagnostic("DUPLICATE_PARAGRAPH_PROPERTY", paragraph.ID, paragraph.Anchor.PartName, directPPr, "Duplicate pStyle values make the paragraph style ambiguous; no paragraph style was applied")
			}
		default:
			resolver.addDiagnostic("DUPLICATE_PARAGRAPH_PROPERTIES", paragraph.ID, paragraph.Anchor.PartName, paragraphNode, "Multiple pPr children make direct paragraph formatting ambiguous; the direct layer was not resolved")
		}
	} else {
		styleID = resolver.defaultP
		if paragraph.Properties != nil && paragraph.Properties.ParagraphStyleID != nil {
			styleID = *paragraph.Properties.ParagraphStyleID
		}
	}
	p := nativeParagraphProperties{}
	applyNativeParagraphProperties(&p, resolver.docP)
	runBase := nativeRunProperties{}
	applyNativeRunProperties(&runBase, resolver.docR, false)
	applied := []string{}
	var tableFontSize *int
	var tableAlignment *string
	for _, definition := range tableStyles {
		applyNativeParagraphProperties(&p, definition.p)
		applyNativeRunProperties(&runBase, definition.r, true)
		if definition.r.fontSize != nil {
			tableFontSize = definition.r.fontSize
		}
		if definition.p.alignment != nil {
			tableAlignment = definition.p.alignment
		}
		// Table provenance remains on resolved.Tables. The paragraph-style
		// chain has its own bounded namespace and must not mix table IDs.
	}
	numberingStyleID := ""
	if styleID != "" {
		chain := resolver.styleChain("paragraph", styleID, paragraph.ID)
		defaultIndex := -1
		var defaultSize *int
		var defaultAlignment *string
		for index, definition := range chain {
			if definition.r.fontSize != nil {
				defaultSize = definition.r.fontSize
			}
			if definition.p.alignment != nil {
				defaultAlignment = definition.p.alignment
			}
			if definition.id == resolver.defaultP {
				defaultIndex = index
				break
			}
		}
		for index, definition := range chain {
			paragraphLayer, runLayer := definition.p, definition.r
			// MS-DOCX 2.3.1: with the default (false) compatibility flag,
			// default paragraph 11/12pt and left alignment do not override
			// table styles. Explicit compatibility flags remain refused by
			// pagination settings until their policy is modeled end to end.
			// Resolve the default style's basedOn ancestors as one layer:
			// its inherited 11/12pt or left alignment has the same exception.
			if index <= defaultIndex {
				if tableFontSize != nil && defaultSize != nil && (*defaultSize == 22 || *defaultSize == 24) {
					runLayer.fontSize = nil
				}
				if tableAlignment != nil && defaultAlignment != nil && *defaultAlignment == "left" {
					paragraphLayer.alignment = nil
				}
			}
			applyNativeParagraphProperties(&p, paragraphLayer)
			applyNativeRunProperties(&runBase, runLayer, true)
			applied = append(applied, definition.id)
			if definition.p.numbering.present {
				numberingStyleID = definition.id
			}
		}
	}
	directP := nativeParagraphProperties{}
	if directPPr != nil {
		directP = resolver.parseParagraphProperties(paragraph.Anchor.PartName, directPPr, paragraph.ID)
	} else if paragraph.Properties != nil {
		if paragraphNode == nil {
			directP = nativeParagraphPropertiesFromContract(paragraph.Properties)
		}
	}
	numberingReference := p.numbering
	applyNativeNumberingProperties(&numberingReference, directP.numbering)
	var levelStyleID *string
	if !directP.numbering.present && p.numbering.present && numberingStyleID != "" {
		levelStyleID = nativeString(numberingStyleID)
	}
	resolvedNumbering, numberingP, markerR := resolver.resolveNumbering(numberingReference, levelStyleID, paragraph.ID, numberingState)
	applyNativeParagraphProperties(&p, numberingP)
	applyNativeParagraphProperties(&p, directP)
	paragraphMark := runBase
	var directParagraphMark nativeRunProperties
	if directPPr != nil {
		markRPrNodes := directNativeChildren(directPPr, resolver.wordNS, "rPr")
		switch len(markRPrNodes) {
		case 0:
		case 1:
			directParagraphMark = resolver.parseRunProperties(paragraph.Anchor.PartName, markRPrNodes[0], paragraph.ID)
			applyNativeRunProperties(&paragraphMark, directParagraphMark, false)
		default:
			resolver.addDiagnostic("DUPLICATE_PARAGRAPH_MARK_PROPERTIES", paragraph.ID, paragraph.Anchor.PartName, directPPr, "Multiple paragraph-mark rPr children make mark formatting ambiguous; the direct mark layer was not resolved")
		}
	}
	if resolvedNumbering != nil {
		marker := runBase
		applyNativeRunProperties(&marker, markerR, false)
		applyNativeRunProperties(&marker, directParagraphMark, false)
		resolver.resolveLatinRunFont(&marker, resolvedNumbering.ResolvedText, paragraph.ID, paragraph.Anchor.PartName)
		resolvedNumbering.Marker = nativeExportRunProperties(marker)
		if !resolver.resolveNumberingGeometry(resolvedNumbering, p, paragraph.ID) {
			resolvedNumbering = nil
		}
	}
	markScriptUncertain := p.bidi.value
	for _, run := range paragraph.Runs {
		if run.Text != nil {
			for _, character := range *run.Text {
				if character > 0x7f {
					markScriptUncertain = true
					break
				}
			}
		}
	}
	resolver.resolveLatinRunFont(&paragraphMark, "\r", paragraph.ID, paragraph.Anchor.PartName, markScriptUncertain)
	resolvedParagraph := NativeResolvedParagraphV1{
		ParagraphID: paragraph.ID, AppliedStyles: applied,
		Properties:              nativeExportParagraphProperties(p),
		ParagraphMarkProperties: nativeExportRunProperties(paragraphMark),
		Numbering:               resolvedNumbering,
	}
	if styleID != "" {
		resolvedParagraph.StyleID = nativeString(styleID)
	}
	result.Paragraphs = append(result.Paragraphs, resolvedParagraph)
	for runIndex := range paragraph.Runs {
		run := &paragraph.Runs[runIndex]
		r := runBase
		characterApplied := []string{}
		var characterStyle *string
		var directRPr *nativeXMLNode
		rawOwnerFound := false
		if runNode := resolver.nodeForAnchor(run.Anchor); runNode != nil {
			for owner := runNode; owner != nil; owner = owner.parent {
				if owner.Name != (xml.Name{Space: resolver.wordNS, Local: "r"}) {
					continue
				}
				rawOwnerFound = true
				rPrNodes := directNativeChildren(owner, resolver.wordNS, "rPr")
				switch len(rPrNodes) {
				case 0:
					if resolver.defaultC != "" {
						characterStyle = nativeString(resolver.defaultC)
					}
				case 1:
					directRPr = rPrNodes[0]
					styleNodes := directNativeChildren(directRPr, resolver.wordNS, "rStyle")
					switch len(styleNodes) {
					case 0:
						if resolver.defaultC != "" {
							characterStyle = nativeString(resolver.defaultC)
						}
					case 1:
						if value, ok := nativeAttr(styleNodes[0], resolver.wordNS, "val"); ok && nativeIDPattern.MatchString(value) {
							characterStyle = nativeString(value)
						} else {
							resolver.addDiagnostic("INVALID_CHARACTER_STYLE", run.ID, run.Anchor.PartName, styleNodes[0], "Invalid character style reference is preserved and not resolved")
						}
					default:
						resolver.addDiagnostic("DUPLICATE_RUN_PROPERTY", run.ID, run.Anchor.PartName, directRPr, "Duplicate rStyle values make the character style ambiguous; no character style was applied")
					}
				default:
					resolver.addDiagnostic("DUPLICATE_RUN_PROPERTIES", run.ID, run.Anchor.PartName, owner, "Multiple rPr children make direct run formatting ambiguous; the direct layer was not resolved")
				}
				break
			}
		}
		if !rawOwnerFound {
			if run.Properties != nil && run.Properties.CharacterStyleID != nil {
				characterStyle = run.Properties.CharacterStyleID
			} else if resolver.defaultC != "" {
				characterStyle = nativeString(resolver.defaultC)
			}
		}
		if characterStyle != nil {
			for _, definition := range resolver.styleChain("character", *characterStyle, run.ID) {
				applyNativeRunProperties(&r, definition.r, true)
				characterApplied = append(characterApplied, definition.id)
			}
		}
		if directRPr != nil {
			applyNativeRunProperties(&r, resolver.parseRunProperties(run.Anchor.PartName, directRPr, run.ID), false)
		}
		if !rawOwnerFound && run.Properties != nil {
			applyNativeRunProperties(&r, nativeRunPropertiesFromContract(run.Properties), false)
		}
		text := ""
		if run.Text != nil {
			text = *run.Text
		}
		resolver.resolveLatinRunFont(&r, text, run.ID, run.Anchor.PartName)
		if run.Kind != "text" && r.verticalAlignment != nil && *r.verticalAlignment != "baseline" {
			resolver.addDiagnostic("VERTICAL_ALIGNMENT_UNSUPPORTED", run.ID, run.Anchor.PartName, nil, "Script transforms on note markers and controls remain unqualified")
		}
		result.Runs = append(result.Runs, NativeResolvedRunV1{
			RunID: run.ID, ParagraphID: paragraph.ID, CharacterStyle: characterStyle,
			AppliedParagraphStyles: append([]string{}, applied...), AppliedCharacterStyles: characterApplied,
			Properties: nativeExportRunProperties(r),
		})
	}
}

func (resolver *nativeLayoutResolver) resolveNumbering(reference nativeNumberingProperties, levelStyleID *string, scopeID string, state *nativeNumberingState) (*NativeResolvedNumberingV1, nativeParagraphProperties, nativeRunProperties) {
	if !reference.present || reference.numID == nil {
		if reference.present && reference.level != nil {
			resolver.addDiagnostic("INCOMPLETE_NUMBERING_REFERENCE", scopeID, resolver.partsValue(resolver.parts.NumberingPart), nil, "A numbering level without an inherited numId was preserved and ignored")
		}
		return nil, nativeParagraphProperties{}, nativeRunProperties{}
	}
	numID, ok := nativeCanonicalDecimalID(*reference.numID)
	if !ok || numID == "0" {
		if numID != "0" {
			resolver.addDiagnostic("INVALID_NUMBERING_REFERENCE", scopeID, resolver.partsValue(resolver.parts.NumberingPart), nil, "The numbering instance id is invalid and was ignored")
		}
		return nil, nativeParagraphProperties{}, nativeRunProperties{}
	}
	resolver.emitNumberingDiagnostics(scopeID, resolver.numberingRootDeferred)
	instance := resolver.nums[numID]
	if instance == nil {
		resolver.addDiagnostic("MISSING_NUMBERING_INSTANCE", scopeID, resolver.partsValue(resolver.parts.NumberingPart), nil, "The referenced numbering instance is missing and was not guessed")
		return nil, nativeParagraphProperties{}, nativeRunProperties{}
	}
	abstract := resolver.abstractNums[instance.abstractID]
	if abstract == nil {
		resolver.addDiagnostic("MISSING_ABSTRACT_NUMBERING", scopeID, resolver.partsValue(resolver.parts.NumberingPart), nil, "The referenced abstract numbering definition is missing and was not guessed")
		return nil, nativeParagraphProperties{}, nativeRunProperties{}
	}
	if reference.abstractHint != nil {
		if expected, valid := nativeCanonicalDecimalID(*reference.abstractHint); !valid || expected != instance.abstractID {
			resolver.addDiagnostic("ABSTRACT_NUMBERING_MISMATCH", scopeID, resolver.partsValue(resolver.parts.NumberingPart), nil, "The native reference abstractNum hint disagrees with numbering.xml and was not trusted")
		}
	}
	levelIndex := 0
	if levelStyleID != nil {
		selected, ok := resolver.numberingLevelForParagraphStyle(abstract, *levelStyleID, scopeID)
		if !ok {
			return nil, nativeParagraphProperties{}, nativeRunProperties{}
		}
		levelIndex = selected
	} else if reference.level != nil {
		levelIndex = *reference.level
	}
	override := instance.overrides[levelIndex]
	abstractLevel := abstract.levels[levelIndex]
	effective := nativeEffectiveNumberingLevel(abstractLevel, override)
	if effective == nil {
		resolver.addDiagnostic("MISSING_NUMBERING_LEVEL", scopeID, resolver.partsValue(resolver.parts.NumberingPart), nil, "The referenced numbering level is missing and was not guessed")
		return nil, nativeParagraphProperties{}, nativeRunProperties{}
	}
	resolver.emitNumberingDiagnostics(scopeID, instance.deferred, abstract.deferred)
	if abstractLevel != nil {
		resolver.emitNumberingDiagnostics(scopeID, abstractLevel.deferred)
	}
	if override != nil {
		resolver.emitNumberingDiagnostics(scopeID, override.deferred)
		if override.level != nil {
			resolver.emitNumberingDiagnostics(scopeID, override.level.deferred)
		}
	}
	resolved := resolver.materializeNativeNumbering(instance, abstract, effective, override, levelIndex, scopeID, state)
	if resolved != nil && levelStyleID != nil {
		resolved.LevelStyleID = nativeString(*levelStyleID)
	}
	return resolved, effective.p, effective.r
}

func (resolver *nativeLayoutResolver) numberingLevelForParagraphStyle(abstract *nativeAbstractNumbering, styleID, scopeID string) (int, bool) {
	matches := make([]int, 0, 2)
	duplicate := false
	for levelIndex := 0; levelIndex <= 8; levelIndex++ {
		level := abstract.levels[levelIndex]
		if level == nil {
			continue
		}
		matchingLinks := 0
		for _, linkedStyleID := range level.styleLinks {
			if linkedStyleID == styleID {
				matchingLinks++
			}
		}
		if matchingLinks == 0 {
			continue
		}
		matches = append(matches, levelIndex)
		if matchingLinks != 1 || level.styleLinkCount != 1 {
			duplicate = true
		}
	}
	partName := resolver.partsValue(resolver.parts.NumberingPart)
	if duplicate {
		resolver.addDiagnostic("DUPLICATE_NUMBERING_STYLE_LEVEL", scopeID, partName, nil, "Duplicate paragraph-style links make the abstract numbering level ambiguous")
		return 0, false
	}
	if len(matches) == 0 {
		resolver.addDiagnostic("MISSING_NUMBERING_STYLE_LEVEL", scopeID, partName, nil, "The paragraph style has no exact w:lvl/w:pStyle mapping in its abstract numbering definition")
		return 0, false
	}
	if len(matches) != 1 {
		resolver.addDiagnostic("AMBIGUOUS_NUMBERING_STYLE_LEVEL", scopeID, partName, nil, "The paragraph style maps to more than one abstract numbering level")
		return 0, false
	}
	return matches[0], true
}

func nativeCanonicalDecimalID(raw string) (string, bool) {
	value, err := strconv.ParseUint(raw, 10, 31)
	if err != nil {
		return "", false
	}
	return strconv.FormatUint(value, 10), true
}

func nativeOrdinaryNumberFormat(value string) bool {
	switch value {
	case "decimal", "bullet", "lowerLetter", "upperLetter", "lowerRoman", "upperRoman":
		return true
	default:
		return false
	}
}

func (resolver *nativeLayoutResolver) styleChain(kind, id, scopeID string) []*nativeStyleDefinition {
	chain := []*nativeStyleDefinition{}
	seenAt := map[string]int{}
	current := id
	for current != "" {
		key := kind + "\x00" + current
		if cycleStart, seen := seenAt[key]; seen {
			resolver.addDiagnostic("STYLE_BASED_ON_CYCLE", scopeID, resolver.partsValue(resolver.parts.StylesPart), nil, "Cyclic basedOn layers were ignored; only acyclic descendant layers were retained")
			chain = chain[:cycleStart]
			break
		}
		seenAt[key] = len(chain)
		if len(seenAt) > NativeDOCXMaxDepth {
			resolver.addDiagnostic("STYLE_BASED_ON_DEPTH", scopeID, resolver.partsValue(resolver.parts.StylesPart), nil, "The over-depth basedOn ancestors were ignored; bounded descendant layers were retained")
			break
		}
		definition := resolver.styles[key]
		if definition == nil {
			resolver.addDiagnostic("MISSING_STYLE_REFERENCE", scopeID, resolver.partsValue(resolver.parts.StylesPart), nil, "The missing basedOn ancestor was ignored; available descendant layers were retained")
			break
		}
		chain = append(chain, definition)
		current = definition.basedOn
	}
	for left, right := 0, len(chain)-1; left < right; left, right = left+1, right-1 {
		chain[left], chain[right] = chain[right], chain[left]
	}
	// Style properties affect layout only through an actual cascade consumer.
	// Keep their exact source paths, but bind failures to that consumer rather
	// than letting an unused style block every paragraph in the document.
	for _, definition := range chain {
		for _, diagnostic := range definition.deferred {
			resolver.addDiagnostic(diagnostic.code, scopeID, diagnostic.partName, diagnostic.node, diagnostic.message)
		}
	}
	return chain
}

func (resolver *nativeLayoutResolver) resolveLatinRunFont(properties *nativeRunProperties, text, scopeID, partName string, scriptContextUncertain ...bool) {
	// MS-OI29500 17.3.2.26 assigns Basic Latin to ascii regardless of
	// inactive East-Asia/complex-script slots. Forced cs remains an unmodeled
	// run property; rtl and non-Basic-Latin text still require script shaping.
	basicLatin := !properties.rtl.value
	if len(scriptContextUncertain) > 0 && scriptContextUncertain[0] {
		basicLatin = false
	}
	for _, character := range text {
		if character > 0x7f {
			basicLatin = false
			break
		}
	}
	if !basicLatin {
		keys := make([]string, 0, len(properties.scriptProperties))
		for key := range properties.scriptProperties {
			keys = append(keys, key)
		}
		sort.Strings(keys)
		for _, key := range keys {
			diagnostic := properties.scriptProperties[key]
			resolver.addDiagnostic(diagnostic.code, scopeID, diagnostic.partName, diagnostic.node, diagnostic.message)
		}
	}
	ascii, hAnsi := properties.asciiFamily, properties.hAnsiFamily
	if ascii == nil && hAnsi == nil {
		return
	}
	if ascii == nil {
		ascii = hAnsi
	}
	if hAnsi == nil {
		hAnsi = ascii
	}
	if *ascii == *hAnsi {
		properties.fontFamily = nativeString(*ascii)
		return
	}
	// OOXML rFonts assigns U+0000..U+007F to ascii/asciiTheme.
	// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.runfonts
	// Only Basic Latin is qualified here. Complex-script overrides, high ANSI,
	// and mixed ranges still require slot-aware shaping and are not guessed.
	if !properties.rtl.value {
		basicLatin := true
		for _, character := range text {
			if character > 0x7f {
				basicLatin = false
				break
			}
		}
		if basicLatin {
			properties.fontFamily = nativeString(*ascii)
			return
		}
	}
	properties.fontFamily = nil
	resolver.addDiagnostic("SCRIPT_DEPENDENT_LATIN_FONT", scopeID, partName, nil, "Distinct ascii/hAnsi fonts are qualified only for non-RTL Basic Latin text; other script ranges require slot-aware shaping")
}

func (resolver *nativeLayoutResolver) parseRunProperties(partName string, node *nativeXMLNode, scopeID string) nativeRunProperties {
	properties := nativeRunProperties{}
	counts := nativeDirectWordChildCounts(node, resolver.wordNS)
	reportedDuplicate := map[string]bool{}
	modeledSingleton := map[string]bool{
		"rStyle": true, "rFonts": true, "sz": true, "szCs": true, "b": true, "i": true,
		"rtl": true, "vanish": true, "bCs": true, "iCs": true, "u": true, "color": true,
		"highlight": true, "lang": true, "vertAlign": true,
	}
	for _, child := range node.Children {
		if child.Name.Space != resolver.wordNS {
			resolver.addDiagnostic("FOREIGN_RUN_PROPERTY", scopeID, partName, child, "Foreign run-property markup is preserved verbatim")
			continue
		}
		if modeledSingleton[child.Name.Local] && counts[child.Name.Local] > 1 {
			if !reportedDuplicate[child.Name.Local] {
				resolver.addDiagnostic("DUPLICATE_RUN_PROPERTY", scopeID, partName, node, "Duplicate "+child.Name.Local+" values make this run-property layer ambiguous; that property was not resolved")
				reportedDuplicate[child.Name.Local] = true
			}
			continue
		}
		switch child.Name.Local {
		case "rStyle":
			// The style reference is consumed from the native run contract.
		case "rFonts":
			allowed := []xml.Name{}
			for _, local := range []string{"ascii", "hAnsi", "asciiTheme", "hAnsiTheme", "eastAsia", "eastAsiaTheme", "cs", "cstheme", "hint"} {
				allowed = append(allowed, xml.Name{Space: resolver.wordNS, Local: local})
			}
			if !nativeExactLeaf(child, allowed...) {
				resolver.addDiagnostic("UNMODELED_FONT_SELECTION", scopeID, partName, child, "Font selection contains unknown attributes or nested markup and is not resolved")
				continue
			}
			ascii, hasASCII := nativeAttr(child, resolver.wordNS, "ascii")
			hAnsi, hasHAnsi := nativeAttr(child, resolver.wordNS, "hAnsi")
			asciiTheme, hasAsciiTheme := nativeAttr(child, resolver.wordNS, "asciiTheme")
			hAnsiTheme, hasHAnsiTheme := nativeAttr(child, resolver.wordNS, "hAnsiTheme")
			_, eastAsia := nativeAttr(child, resolver.wordNS, "eastAsia")
			_, eastAsiaTheme := nativeAttr(child, resolver.wordNS, "eastAsiaTheme")
			_, cs := nativeAttr(child, resolver.wordNS, "cs")
			_, csTheme := nativeAttr(child, resolver.wordNS, "cstheme")
			hintValue, hint := nativeAttr(child, resolver.wordNS, "hint")
			validScript := true
			for _, local := range []string{"eastAsia", "cs"} {
				if value, present := nativeAttr(child, resolver.wordNS, local); present && !nativeBoundedResolvedString(value, 256) {
					validScript = false
				}
			}
			for _, local := range []string{"eastAsiaTheme", "cstheme"} {
				if value, present := nativeAttr(child, resolver.wordNS, local); present {
					switch value {
					case "majorAscii", "majorHAnsi", "majorEastAsia", "majorBidi", "minorAscii", "minorHAnsi", "minorEastAsia", "minorBidi":
					default:
						validScript = false
					}
				}
			}
			if hint && hintValue != "default" && hintValue != "eastAsia" && hintValue != "cs" {
				validScript = false
			}
			if !validScript {
				resolver.addDiagnostic("UNMODELED_FONT_SELECTION", scopeID, partName, child, "Invalid script font slot or hint is preserved and not resolved")
				continue
			}
			if eastAsia || eastAsiaTheme || cs || csTheme {
				properties.deferScriptProperty("fonts", "SCRIPT_FONT_PRESERVED", partName, child, "East-Asia/complex-script font selection requires script shaping and is not guessed")
			}
			if hint {
				properties.deferScriptProperty("hint", "FONT_HINT_PRESERVED", partName, child, "Font hint selection is preserved for a future script-aware shaper")
			}
			validASCII := !hasASCII || nativeBoundedResolvedString(ascii, 256)
			validHAnsi := !hasHAnsi || nativeBoundedResolvedString(hAnsi, 256)
			asciiFace := ""
			hAnsiFace := ""
			if hasAsciiTheme {
				face, ok := nativeThemeLatinTypeface(asciiTheme, resolver.themeLatinFonts)
				if !ok {
					resolver.addDiagnostic("THEME_FONT_PRESERVED", scopeID, partName, child, "Theme font slot has no exact latin typeface and is not guessed")
				} else {
					asciiFace = face
				}
			} else if hasASCII && validASCII {
				asciiFace = ascii
			}
			if hasHAnsiTheme {
				face, ok := nativeThemeLatinTypeface(hAnsiTheme, resolver.themeLatinFonts)
				if !ok {
					resolver.addDiagnostic("THEME_FONT_PRESERVED", scopeID, partName, child, "Theme font slot has no exact latin typeface and is not guessed")
				} else {
					hAnsiFace = face
				}
			} else if hasHAnsi && validHAnsi {
				hAnsiFace = hAnsi
			}
			if !validASCII || !validHAnsi {
				resolver.addDiagnostic("INVALID_FONT_FAMILY", scopeID, partName, child, "Invalid explicit font family is preserved and ignored")
			} else {
				if asciiFace != "" {
					properties.asciiFamily = nativeString(asciiFace)
				}
				if hAnsiFace != "" {
					properties.hAnsiFamily = nativeString(hAnsiFace)
				}
			}
		case "sz":
			if value, ok := nativePositiveIntAttr(child, resolver.wordNS, "val"); ok && value <= 3276 {
				properties.fontSize = nativeInt(value)
			} else {
				resolver.addDiagnostic("INVALID_FONT_SIZE", scopeID, partName, child, "Invalid font size is preserved and ignored")
			}
		case "szCs":
			if value, ok := nativePositiveIntAttr(child, resolver.wordNS, "val"); !ok || value > 3276 || !nativeExactLeaf(child, xml.Name{Space: resolver.wordNS, Local: "val"}) {
				resolver.addDiagnostic("INVALID_FONT_SIZE", scopeID, partName, child, "Invalid complex-script size is preserved and not resolved")
			} else {
				properties.deferScriptProperty("size", "COMPLEX_SCRIPT_SIZE_PRESERVED", partName, child, "Complex-script font size is preserved for a future shaper")
			}
		case "b", "i", "rtl", "vanish":
			value, ok := nativeOnOff(child, resolver.wordNS)
			if !ok || !nativeExactLeaf(child, xml.Name{Space: resolver.wordNS, Local: "val"}) {
				resolver.addDiagnostic("INVALID_ON_OFF_PROPERTY", scopeID, partName, child, "Invalid on/off property is preserved and ignored")
				continue
			}
			property := nativeBoolProperty{present: true, value: value}
			switch child.Name.Local {
			case "b":
				properties.bold = property
			case "i":
				properties.italic = property
			case "rtl":
				properties.rtl = property
			case "vanish":
				properties.hidden = property
			}
		case "bCs", "iCs":
			if _, ok := nativeOnOff(child, resolver.wordNS); !ok || !nativeExactLeaf(child, xml.Name{Space: resolver.wordNS, Local: "val"}) {
				resolver.addDiagnostic("INVALID_ON_OFF_PROPERTY", scopeID, partName, child, "Invalid complex-script toggle is preserved and not resolved")
			} else {
				properties.deferScriptProperty(child.Name.Local, "COMPLEX_SCRIPT_TOGGLE_PRESERVED", partName, child, "Complex-script toggles are preserved for a future shaper")
			}
		case "u":
			value, ok := nativeAttr(child, resolver.wordNS, "val")
			if !ok {
				value = "single"
			}
			if value == "none" || value == "single" || value == "double" || value == "words" {
				properties.underline = nativeString(value)
			} else {
				resolver.addDiagnostic("UNSUPPORTED_UNDERLINE", scopeID, partName, child, "This underline variant is preserved but not resolved")
			}
			if _, color := nativeAttr(child, resolver.wordNS, "color"); color {
				resolver.addDiagnostic("UNDERLINE_COLOR_PRESERVED", scopeID, partName, child, "Underline color is preserved for a future text painter")
			}
			if _, themeColor := nativeAttr(child, resolver.wordNS, "themeColor"); themeColor {
				resolver.addDiagnostic("THEME_UNDERLINE_COLOR_PRESERVED", scopeID, partName, child, "Theme-dependent underline color is preserved and not guessed")
			}
		case "color":
			if rgb, ok := resolver.resolveExactRunColor(child, scopeID, partName); ok {
				properties.color = rgb
			}
		case "highlight":
			if value, ok := nativeAttr(child, resolver.wordNS, "val"); ok && nativeResolvedHighlight(value) {
				properties.highlight = nativeString(value)
			} else {
				resolver.addDiagnostic("UNSUPPORTED_HIGHLIGHT", scopeID, partName, child, "This highlight value is preserved and not resolved")
			}
		case "lang":
			if !nativeExactLeaf(child, xml.Name{Space: resolver.wordNS, Local: "val"}, xml.Name{Space: resolver.wordNS, Local: "eastAsia"}, xml.Name{Space: resolver.wordNS, Local: "bidi"}) {
				resolver.addDiagnostic("INVALID_LANGUAGE", scopeID, partName, child, "Unknown language markup is preserved and not resolved")
				continue
			}
			value, ok := nativeAttr(child, resolver.wordNS, "val")
			_, eastAsia := nativeAttr(child, resolver.wordNS, "eastAsia")
			_, bidi := nativeAttr(child, resolver.wordNS, "bidi")
			if ok && nativeBoundedResolvedString(value, 256) {
				properties.language = nativeString(value)
			} else if ok {
				resolver.addDiagnostic("INVALID_LANGUAGE", scopeID, partName, child, "Invalid language metadata is preserved and ignored")
			}
			if eastAsia || bidi {
				valid := true
				for _, local := range []string{"eastAsia", "bidi"} {
					if value, present := nativeAttr(child, resolver.wordNS, local); present && !nativeScriptLanguageTag(value) {
						valid = false
					}
				}
				if !valid {
					resolver.addDiagnostic("INVALID_LANGUAGE", scopeID, partName, child, "Invalid script language is preserved and not resolved")
				} else {
					properties.deferScriptProperty("language", "SCRIPT_LANGUAGE_PRESERVED", partName, child, "East-Asia/bidi language metadata is preserved for script shaping")
				}
			}
		case "noProof":
			if !nativeNeutralSourceProperty(child, node, resolver.wordNS) {
				resolver.addDiagnostic("UNMODELED_RUN_PROPERTY", scopeID, partName, child, "Proofing metadata has malformed, duplicate or unknown source structure")
			}
		case "vertAlign":
			value, ok := nativeVerticalAlignmentValue(child, resolver.wordNS)
			if ok {
				properties.verticalAlignment = nativeString(value)
			} else {
				resolver.addDiagnostic("VERTICAL_ALIGNMENT_UNSUPPORTED", scopeID, partName, child, "Vertical alignment requires an exact baseline, subscript or superscript value")
			}
		default:
			resolver.addDiagnostic("UNMODELED_RUN_PROPERTY", scopeID, partName, child, "This run property is preserved and not guessed")
		}
	}
	return properties
}

func (resolver *nativeLayoutResolver) parseParagraphProperties(partName string, node *nativeXMLNode, scopeID string) nativeParagraphProperties {
	properties := nativeParagraphProperties{}
	counts := nativeDirectWordChildCounts(node, resolver.wordNS)
	reportedDuplicate := map[string]bool{}
	modeledSingleton := map[string]bool{
		"pStyle": true, "numPr": true, "rPr": true, "sectPr": true, "jc": true,
		"spacing": true, "ind": true, "keepNext": true, "keepLines": true,
		"pageBreakBefore": true, "widowControl": true, "bidi": true,
	}
	for _, child := range node.Children {
		if child.Name.Space != resolver.wordNS {
			resolver.addDiagnostic("FOREIGN_PARAGRAPH_PROPERTY", scopeID, partName, child, "Foreign paragraph-property markup is preserved verbatim")
			continue
		}
		if modeledSingleton[child.Name.Local] && counts[child.Name.Local] > 1 {
			if !reportedDuplicate[child.Name.Local] {
				resolver.addDiagnostic("DUPLICATE_PARAGRAPH_PROPERTY", scopeID, partName, node, "Duplicate "+child.Name.Local+" values make this paragraph-property layer ambiguous; that property was not resolved")
				reportedDuplicate[child.Name.Local] = true
			}
			continue
		}
		switch child.Name.Local {
		case "pStyle", "rPr", "sectPr":
			// Consumed elsewhere by the style/section cascade.
		case "numPr":
			properties.numbering = resolver.parseNumberingProperties(partName, child, scopeID)
		case "jc":
			if value, ok := nativeAttr(child, resolver.wordNS, "val"); ok && nativeParagraphAlignment(value) {
				properties.alignment = nativeString(value)
			} else {
				resolver.addDiagnostic("UNSUPPORTED_PARAGRAPH_ALIGNMENT", scopeID, partName, child, "This alignment is preserved but not resolved")
			}
		case "spacing":
			resolver.parseSpacing(child, scopeID, partName, &properties)
		case "ind":
			resolver.parseIndent(child, scopeID, partName, &properties)
		case "autoSpaceDE", "autoSpaceDN":
			if !nativeNeutralSourceProperty(child, node, resolver.wordNS) {
				resolver.addDiagnostic("UNMODELED_PARAGRAPH_PROPERTY", scopeID, partName, child, "Automatic East Asian spacing is supported only as an exact explicit disabled setting")
			}
		case "keepNext", "keepLines", "pageBreakBefore", "widowControl", "bidi":
			value, ok := nativeOnOff(child, resolver.wordNS)
			if !ok {
				resolver.addDiagnostic("INVALID_ON_OFF_PROPERTY", scopeID, partName, child, "Invalid on/off property is preserved and ignored")
				continue
			}
			property := nativeBoolProperty{present: true, value: value}
			switch child.Name.Local {
			case "keepNext":
				properties.keepNext = property
			case "keepLines":
				properties.keepLines = property
			case "pageBreakBefore":
				properties.pageBreakBefore = property
			case "widowControl":
				properties.widowControl = property
			case "bidi":
				properties.bidi = property
			}
		default:
			resolver.addDiagnostic("UNMODELED_PARAGRAPH_PROPERTY", scopeID, partName, child, "This paragraph property is preserved and not guessed")
		}
	}
	return properties
}

func nativeDirectWordChildCounts(node *nativeXMLNode, wordNS string) map[string]int {
	counts := map[string]int{}
	if node == nil {
		return counts
	}
	for _, child := range node.Children {
		if child.Name.Space == wordNS {
			counts[child.Name.Local]++
		}
	}
	return counts
}

func (resolver *nativeLayoutResolver) parseSpacing(node *nativeXMLNode, scopeID, partName string, properties *nativeParagraphProperties) {
	parseAuto := func(name, message string) (bool, bool) {
		raw, present := nativeAttr(node, resolver.wordNS, name)
		if !present {
			return false, false
		}
		value, valid := nativeLexicalOnOff(raw)
		if !valid {
			resolver.addDiagnostic("INVALID_AUTO_PARAGRAPH_SPACING", scopeID, partName, node, "Invalid automatic paragraph spacing is preserved and ignored")
			return false, true
		}
		if value {
			resolver.addDiagnostic("AUTO_PARAGRAPH_SPACING_PRESERVED", scopeID, partName, node, message)
		}
		return value, true
	}
	beforeAuto, _ := parseAuto("beforeAutospacing", "Automatic before spacing is not guessed")
	afterAuto, _ := parseAuto("afterAutospacing", "Automatic after spacing is not guessed")
	if _, present := nativeAttr(node, resolver.wordNS, "before"); present && !beforeAuto {
		if value, ok := nativeNonnegativeInt64Attr(node, resolver.wordNS, "before"); ok {
			properties.spacingBefore = nativeInt64(value)
		} else {
			resolver.addDiagnostic("INVALID_PARAGRAPH_SPACING", scopeID, partName, node, "Invalid before spacing is preserved and ignored")
		}
	}
	if _, present := nativeAttr(node, resolver.wordNS, "after"); present && !afterAuto {
		if value, ok := nativeNonnegativeInt64Attr(node, resolver.wordNS, "after"); ok {
			properties.spacingAfter = nativeInt64(value)
		} else {
			resolver.addDiagnostic("INVALID_PARAGRAPH_SPACING", scopeID, partName, node, "Invalid after spacing is preserved and ignored")
		}
	}
	linePresent := false
	if _, present := nativeAttr(node, resolver.wordNS, "line"); present {
		linePresent = true
		if value, ok := nativeNonnegativeInt64Attr(node, resolver.wordNS, "line"); ok {
			properties.line = nativeInt64(value)
			properties.lineRule = nativeString("auto")
		} else {
			resolver.addDiagnostic("INVALID_LINE_SPACING", scopeID, partName, node, "Invalid line spacing is preserved and ignored")
		}
	}
	if value, present := nativeAttr(node, resolver.wordNS, "lineRule"); present {
		if value == "auto" || value == "exact" || value == "atLeast" {
			if properties.line != nil {
				properties.lineRule = nativeString(value)
			} else {
				resolver.addDiagnostic("INCOMPLETE_LINE_SPACING", scopeID, partName, node, "A line rule without a line measurement is preserved and ignored")
			}
		} else {
			properties.line, properties.lineRule = nil, nil
			resolver.addDiagnostic("INVALID_LINE_SPACING", scopeID, partName, node, "Invalid line-spacing rule makes the line measurement unsafe to resolve")
		}
	} else if linePresent && properties.line != nil {
		properties.lineRule = nativeString("auto")
	}
	for _, name := range []string{"beforeLines", "afterLines"} {
		if _, present := nativeAttr(node, resolver.wordNS, name); present {
			resolver.addDiagnostic("LINE_UNIT_PARAGRAPH_SPACING_PRESERVED", scopeID, partName, node, "Line-unit paragraph spacing is preserved for a future line layouter")
		}
	}
}

func (resolver *nativeLayoutResolver) parseIndent(node *nativeXMLNode, scopeID, partName string, properties *nativeParagraphProperties) {
	for _, field := range []struct {
		name   string
		target **int64
		signed bool
	}{
		{"left", &properties.indentLeft, true},
		{"right", &properties.indentRight, true},
		{"start", &properties.indentStart, true},
		{"end", &properties.indentEnd, true},
		{"firstLine", &properties.firstLine, false},
		{"hanging", &properties.hanging, false},
	} {
		if _, present := nativeAttr(node, resolver.wordNS, field.name); !present {
			continue
		}
		value, valid := int64(0), false
		if field.signed {
			value, valid = nativeSignedInt64Attr(node, resolver.wordNS, field.name)
		} else {
			value, valid = nativeNonnegativeInt64Attr(node, resolver.wordNS, field.name)
		}
		if valid {
			*field.target = nativeInt64(value)
		} else {
			resolver.addDiagnostic("INVALID_PARAGRAPH_INDENT", scopeID, partName, node, "Invalid paragraph indentation is preserved and ignored")
		}
	}
	if properties.firstLine != nil && properties.hanging != nil {
		properties.firstLine, properties.hanging = nil, nil
		resolver.addDiagnostic("CONFLICTING_PARAGRAPH_INDENT", scopeID, partName, node, "Conflicting first-line and hanging indents are preserved and neither is guessed")
	}
	for _, name := range []string{"leftChars", "rightChars", "startChars", "endChars", "firstLineChars", "hangingChars"} {
		if _, present := nativeAttr(node, resolver.wordNS, name); present {
			resolver.addDiagnostic("CHARACTER_INDENT_PRESERVED", scopeID, partName, node, "Character-unit indentation requires font metrics and is not guessed")
		}
	}
}

func nativeResolvedHighlight(value string) bool {
	switch value {
	case "none", "black", "blue", "cyan", "green", "magenta", "red", "yellow", "white",
		"darkBlue", "darkCyan", "darkGreen", "darkMagenta", "darkRed", "darkYellow", "darkGray", "lightGray":
		return true
	default:
		return false
	}
}

func (resolver *nativeLayoutResolver) parseNumberingProperties(partName string, node *nativeXMLNode, scopeID string) nativeNumberingProperties {
	counts := nativeDirectWordChildCounts(node, resolver.wordNS)
	if counts["numId"] > 1 || counts["ilvl"] > 1 {
		resolver.addDiagnostic("DUPLICATE_NUMBERING_REFERENCE", scopeID, partName, node, "Duplicate numId/ilvl children make this numPr layer ambiguous; the layer was not resolved")
		return nativeNumberingProperties{}
	}
	properties := nativeNumberingProperties{present: true}
	for _, child := range node.Children {
		if child.Name.Space != resolver.wordNS {
			resolver.addDiagnostic("FOREIGN_NUMBERING_REFERENCE", scopeID, partName, child, "Foreign numbering-reference metadata is preserved")
			continue
		}
		switch child.Name.Local {
		case "numId":
			if value, ok := nativeDecimalIDAttr(child, resolver.wordNS, "val"); ok {
				properties.numID = nativeString(value)
			} else {
				resolver.addDiagnostic("INVALID_NUMBERING_REFERENCE", scopeID, partName, child, "Invalid numId is preserved and ignored")
			}
		case "ilvl":
			if value, ok := nativeNumberingLevelAttr(child, resolver.wordNS, "val"); ok {
				properties.level = nativeInt(value)
			} else {
				resolver.addDiagnostic("INVALID_NUMBERING_REFERENCE", scopeID, partName, child, "Invalid ilvl is preserved and ignored")
			}
		case "numberingChange", "ins":
			resolver.addDiagnostic("TRACKED_NUMBERING_CHANGE_PRESERVED", scopeID, partName, child, "Tracked numbering changes are preserved and not applied")
		default:
			resolver.addDiagnostic("UNMODELED_NUMBERING_REFERENCE", scopeID, partName, child, "Numbering-reference metadata is preserved and not interpreted")
		}
	}
	return properties
}

func nativeParagraphPropertiesFromContract(properties *NativeParagraphPropertiesV1) nativeParagraphProperties {
	result := nativeParagraphProperties{alignment: properties.Alignment}
	if properties.Numbering != nil {
		result.numbering = nativeNumberingProperties{present: true, numID: nativeString(properties.Numbering.NumID), level: nativeInt(*properties.Numbering.Level), abstractHint: properties.Numbering.AbstractNumID}
	}
	if properties.KeepNext != nil {
		result.keepNext = nativeBoolProperty{true, *properties.KeepNext}
	}
	if properties.KeepLines != nil {
		result.keepLines = nativeBoolProperty{true, *properties.KeepLines}
	}
	if properties.PageBreakBefore != nil {
		result.pageBreakBefore = nativeBoolProperty{true, *properties.PageBreakBefore}
	}
	if properties.WidowControl != nil {
		result.widowControl = nativeBoolProperty{true, *properties.WidowControl}
	}
	return result
}

func nativeRunPropertiesFromContract(properties *NativeRunPropertiesV1) nativeRunProperties {
	result := nativeRunProperties{
		fontFamily: properties.FontFamily, fontSize: properties.FontSizeHalfPoint,
		underline: properties.Underline, verticalAlignment: properties.VerticalAlignment, color: properties.Color, highlight: properties.Highlight,
		language: properties.Language,
	}
	if properties.Bold != nil {
		result.bold = nativeBoolProperty{true, *properties.Bold}
	}
	if properties.Italic != nil {
		result.italic = nativeBoolProperty{true, *properties.Italic}
	}
	if properties.RTL != nil {
		result.rtl = nativeBoolProperty{true, *properties.RTL}
	}
	if properties.Hidden != nil {
		result.hidden = nativeBoolProperty{true, *properties.Hidden}
	}
	return result
}

func applyNativeParagraphProperties(target *nativeParagraphProperties, layer nativeParagraphProperties) {
	applyNativeNumberingProperties(&target.numbering, layer.numbering)
	if layer.alignment != nil {
		target.alignment = nativeString(*layer.alignment)
	}
	if layer.spacingBefore != nil {
		target.spacingBefore = nativeInt64(*layer.spacingBefore)
	}
	if layer.spacingAfter != nil {
		target.spacingAfter = nativeInt64(*layer.spacingAfter)
	}
	if layer.line != nil {
		target.line = nativeInt64(*layer.line)
	}
	if layer.lineRule != nil {
		target.lineRule = nativeString(*layer.lineRule)
	}
	if layer.indentLeft != nil {
		target.indentLeft = nativeInt64(*layer.indentLeft)
	}
	if layer.indentRight != nil {
		target.indentRight = nativeInt64(*layer.indentRight)
	}
	if layer.indentStart != nil {
		target.indentStart = nativeInt64(*layer.indentStart)
	}
	if layer.indentEnd != nil {
		target.indentEnd = nativeInt64(*layer.indentEnd)
	}
	if layer.firstLine != nil {
		target.firstLine = nativeInt64(*layer.firstLine)
		target.hanging = nil
	}
	if layer.hanging != nil {
		target.hanging = nativeInt64(*layer.hanging)
		target.firstLine = nil
	}
	if layer.keepNext.present {
		target.keepNext = layer.keepNext
	}
	if layer.keepLines.present {
		target.keepLines = layer.keepLines
	}
	if layer.pageBreakBefore.present {
		target.pageBreakBefore = layer.pageBreakBefore
	}
	if layer.widowControl.present {
		target.widowControl = layer.widowControl
	}
	if layer.bidi.present {
		target.bidi = layer.bidi
	}
}

func applyNativeNumberingProperties(target *nativeNumberingProperties, layer nativeNumberingProperties) {
	if !layer.present {
		return
	}
	target.present = true
	if layer.numID != nil {
		target.numID = nativeString(*layer.numID)
	}
	if layer.level != nil {
		target.level = nativeInt(*layer.level)
	}
	if layer.abstractHint != nil {
		target.abstractHint = nativeString(*layer.abstractHint)
	}
}

func (properties *nativeRunProperties) deferScriptProperty(key, code, partName string, node *nativeXMLNode, message string) {
	if properties.scriptProperties == nil {
		properties.scriptProperties = map[string]nativeDeferredNumberingDiagnostic{}
	}
	properties.scriptProperties[key] = nativeDeferredNumberingDiagnostic{code: code, partName: partName, node: node, message: message}
}

// Conservative language-tag shape for inactive script attributes. This is not
// a registry lookup; malformed or non-tag values remain explicitly refused.
func nativeScriptLanguageTag(value string) bool {
	if !nativeBoundedResolvedString(value, 256) {
		return false
	}
	for index, part := range strings.Split(value, "-") {
		if len(part) == 0 || len(part) > 8 {
			return false
		}
		for _, character := range part {
			if (character < 'A' || character > 'Z') && (character < 'a' || character > 'z') && (index == 0 || character < '0' || character > '9') {
				return false
			}
		}
	}
	return true
}

func applyNativeRunProperties(target *nativeRunProperties, layer nativeRunProperties, styleToggle bool) {
	if len(layer.scriptProperties) > 0 {
		merged := make(map[string]nativeDeferredNumberingDiagnostic, len(target.scriptProperties)+len(layer.scriptProperties))
		for key, value := range target.scriptProperties {
			merged[key] = value
		}
		for key, value := range layer.scriptProperties {
			merged[key] = value
		}
		target.scriptProperties = merged
	}
	if layer.verticalAlignment != nil {
		target.verticalAlignment = nativeString(*layer.verticalAlignment)
	}
	if layer.fontFamily != nil {
		target.fontFamily = nativeString(*layer.fontFamily)
		target.asciiFamily, target.hAnsiFamily = nativeString(*layer.fontFamily), nativeString(*layer.fontFamily)
	}
	if layer.asciiFamily != nil {
		target.asciiFamily = nativeString(*layer.asciiFamily)
	}
	if layer.hAnsiFamily != nil {
		target.hAnsiFamily = nativeString(*layer.hAnsiFamily)
	}
	if layer.fontSize != nil {
		target.fontSize = nativeInt(*layer.fontSize)
	}
	if layer.underline != nil {
		target.underline = nativeString(*layer.underline)
	}
	if layer.color != nil {
		target.color = nativeString(*layer.color)
	}
	if layer.highlight != nil {
		target.highlight = nativeString(*layer.highlight)
	}
	if layer.language != nil {
		target.language = nativeString(*layer.language)
	}
	applyBool := func(targetProperty *nativeBoolProperty, layerProperty nativeBoolProperty, toggle bool) {
		if !layerProperty.present {
			return
		}
		if toggle {
			if layerProperty.value {
				targetProperty.present = true
				targetProperty.value = !targetProperty.value
			}
			return
		}
		*targetProperty = layerProperty
	}
	applyBool(&target.bold, layer.bold, styleToggle)
	applyBool(&target.italic, layer.italic, styleToggle)
	applyBool(&target.hidden, layer.hidden, styleToggle)
	applyBool(&target.rtl, layer.rtl, false)
}

func nativeExportParagraphProperties(properties nativeParagraphProperties) NativeResolvedParagraphPropertiesV1 {
	result := NativeResolvedParagraphPropertiesV1{
		Alignment: properties.alignment, SpacingBeforeTwips: properties.spacingBefore,
		SpacingAfterTwips: properties.spacingAfter, Line: properties.line, LineRule: properties.lineRule,
		IndentLeftTwips: properties.indentLeft, IndentRightTwips: properties.indentRight,
		IndentStartTwips: properties.indentStart, IndentEndTwips: properties.indentEnd,
		FirstLineTwips: properties.firstLine, HangingTwips: properties.hanging,
	}
	if properties.keepNext.present {
		result.KeepNext = nativeBool(properties.keepNext.value)
	}
	if properties.keepLines.present {
		result.KeepLines = nativeBool(properties.keepLines.value)
	}
	if properties.pageBreakBefore.present {
		result.PageBreakBefore = nativeBool(properties.pageBreakBefore.value)
	}
	if properties.widowControl.present {
		result.WidowControl = nativeBool(properties.widowControl.value)
	}
	if properties.bidi.present {
		result.Bidi = nativeBool(properties.bidi.value)
	}
	return result
}

func nativeExportRunProperties(properties nativeRunProperties) NativeResolvedRunPropertiesV1 {
	result := NativeResolvedRunPropertiesV1{
		FontFamily: properties.fontFamily, FontSizeHalfPoint: properties.fontSize,
		Underline: properties.underline, VerticalAlignment: properties.verticalAlignment, Color: properties.color, Highlight: properties.highlight,
		Language: properties.language,
	}
	if properties.bold.present {
		result.Bold = nativeBool(properties.bold.value)
	}
	if properties.italic.present {
		result.Italic = nativeBool(properties.italic.value)
	}
	if properties.rtl.present {
		result.RTL = nativeBool(properties.rtl.value)
	}
	if properties.hidden.present {
		result.Hidden = nativeBool(properties.hidden.value)
	}
	return result
}

func (resolver *nativeLayoutResolver) addDiagnostic(code, scopeID, partName string, node *nativeXMLNode, message string) {
	if resolver.deferredNumbering != nil {
		resolver.deferredDiagnosticCount++
		if resolver.deferredDiagnosticCount > NativeDOCXMaxResolvedDiagnostics {
			resolver.diagnosticOverflow = true
			return
		}
		*resolver.deferredNumbering = append(*resolver.deferredNumbering, nativeDeferredNumberingDiagnostic{code: code, partName: partName, node: node, message: message})
		return
	}
	path := ""
	if node != nil {
		path = node.Path
	}
	key := code + "\x00" + scopeID + "\x00" + partName + "\x00" + path + "\x00" + message
	if resolver.diagnosticSet[key] {
		return
	}
	resolver.diagnosticSet[key] = true
	if len(resolver.diagnostics) >= NativeDOCXMaxResolvedDiagnostics {
		resolver.diagnosticOverflow = true
		return
	}
	diagnostic := NativeResolutionDiagnosticV1{
		Code: code, Severity: "unsupported", ScopeID: scopeID,
		Preservation: "preserve-verbatim", Message: message,
	}
	if partName != "" {
		diagnostic.PartName = nativeString(partName)
	}
	if path != "" {
		diagnostic.Path = nativeString(path)
	}
	resolver.diagnostics = append(resolver.diagnostics, diagnostic)
}

func (resolver *nativeLayoutResolver) emitNumberingDiagnostics(scopeID string, diagnostics ...[]nativeDeferredNumberingDiagnostic) {
	for _, collection := range diagnostics {
		for _, diagnostic := range collection {
			resolver.addDiagnostic(diagnostic.code, scopeID, diagnostic.partName, diagnostic.node, diagnostic.message)
		}
	}
}

func (resolver *nativeLayoutResolver) partsValue(value *string) string {
	if value == nil {
		return ""
	}
	return *value
}

func nativeParagraphAlignment(value string) bool {
	switch value {
	case "left", "right", "center", "both", "distribute", "start", "end":
		return true
	default:
		return false
	}
}

func nativeLexicalOnOff(value string) (bool, bool) {
	switch strings.ToLower(value) {
	case "true", "1", "on":
		return true, true
	case "false", "0", "off":
		return false, true
	default:
		return false, false
	}
}

func rejectNativeStyleNamespaceSpoofing(root *nativeXMLNode, wordNS string) error {
	known := map[string]bool{
		"styles": true, "docDefaults": true, "rPrDefault": true, "pPrDefault": true,
		"style": true, "basedOn": true, "pPr": true, "rPr": true, "name": true,
		"pStyle": true, "numPr": true, "rStyle": true, "rFonts": true, "sz": true,
		"szCs": true, "b": true, "i": true, "bCs": true, "iCs": true, "rtl": true,
		"vanish": true, "u": true, "color": true, "highlight": true, "lang": true,
		"vertAlign": true,
		"jc":        true, "spacing": true, "ind": true, "keepNext": true, "keepLines": true,
		"pageBreakBefore": true, "widowControl": true, "bidi": true, "tblStylePr": true,
		"tabs": true, "tab": true,
		"tblPr": true, "tcPr": true, "trPr": true, "tblBorders": true, "tcBorders": true,
		"tblCellMar": true, "tblW": true, "tblLayout": true, "tblInd": true, "tblLook": true,
		"top": true, "right": true, "bottom": true, "left": true, "insideH": true, "insideV": true,
		"shd": true,
	}
	return rejectNativeKnownLocalSpoofing(root, wordNS, known)
}

func rejectNativeKnownLocalSpoofing(root *nativeXMLNode, wordNS string, known map[string]bool) error {
	var visit func(*nativeXMLNode) error
	visit = func(node *nativeXMLNode) error {
		if known[node.Name.Local] && node.Name.Space != wordNS {
			return fmt.Errorf("namespace spoofing at %s: {%s}%s", node.Path, node.Name.Space, node.Name.Local)
		}
		for _, child := range node.Children {
			if err := visit(child); err != nil {
				return err
			}
		}
		return nil
	}
	return visit(root)
}

// ValidateNativeResolvedLayoutInputV1 validates identity references, bounded
// collections and strings, source part names, and the conservative property
// subset exposed to a future shaper/paginator.
func ValidateNativeResolvedLayoutInputV1(input *NativeResolvedLayoutInputV1) error {
	if input == nil {
		return fmt.Errorf("resolved layout input is nil")
	}
	if input.Protocol != NativeDOCXResolvedLayoutProtocol || input.Version != NativeDOCXResolvedLayoutVersion {
		return fmt.Errorf("unsupported resolved layout protocol/version")
	}
	if !nativeIDPattern.MatchString(input.DocumentID) || !nativeIDPattern.MatchString(input.Revision) || input.SourceParts.MainPart == "" {
		return fmt.Errorf("invalid resolved layout identity/source")
	}
	parts := []*string{nativeString(input.SourceParts.MainPart), input.SourceParts.StylesPart, input.SourceParts.NumberingPart, input.SourceParts.ThemePart, input.SourceParts.FontTablePart}
	seenParts := map[string]bool{}
	for _, partName := range parts {
		if partName == nil {
			continue
		}
		if err := validateNativePartName(*partName); err != nil {
			return fmt.Errorf("invalid resolved source part %q: %w", *partName, err)
		}
		key, _ := nativeDecodedPartKey(*partName)
		if seenParts[key] {
			return fmt.Errorf("duplicate resolved source part %q", *partName)
		}
		seenParts[key] = true
	}
	if input.NumberingSource != nil {
		source := input.NumberingSource
		if input.SourceParts.NumberingPart == nil || *input.SourceParts.NumberingPart != source.PartName {
			return fmt.Errorf("numbering source does not match source_parts.numbering_part")
		}
		if err := validateNativePartName(source.RelationshipsPart); err != nil {
			return fmt.Errorf("invalid numbering relationships part: %w", err)
		}
		if err := validateNativePartName(source.PartName); err != nil {
			return fmt.Errorf("invalid numbering part: %w", err)
		}
		if !nativeSHA256.MatchString(source.RelationshipsSHA256) || !nativeSHA256.MatchString(source.PartSHA256) || !nativeSHA256.MatchString(source.ModelSHA256) {
			return fmt.Errorf("invalid numbering source digest")
		}
		if !nativeIDPattern.MatchString(source.RelationshipID) || !nativeSafeAbsoluteURI(source.RelationshipType) || !nativeBoundedResolvedString(source.RelationshipTarget, 4096) {
			return fmt.Errorf("invalid numbering relationship identity")
		}
		if source.RelationshipType != relBaseTransitional+"numbering" && source.RelationshipType != relBaseStrict+"numbering" {
			return fmt.Errorf("invalid numbering relationship type")
		}
		if source.ContentType != "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml" {
			return fmt.Errorf("invalid numbering content type")
		}
		expectedRelationshipsPart := path.Join(path.Dir(input.SourceParts.MainPart), "_rels", path.Base(input.SourceParts.MainPart)+".rels")
		expectedRelationshipsKey, expectedErr := nativeDecodedPartKey(expectedRelationshipsPart)
		actualRelationshipsKey, actualErr := nativeDecodedPartKey(source.RelationshipsPart)
		if expectedErr != nil || actualErr != nil || expectedRelationshipsKey != actualRelationshipsKey {
			return fmt.Errorf("numbering relationship part does not belong to source_parts.main_part")
		}
		resolvedTarget, targetErr := resolveNativeRelationshipTarget(input.SourceParts.MainPart, source.RelationshipTarget)
		resolvedTargetKey, resolvedErr := nativeDecodedPartKey(resolvedTarget)
		partKey, partErr := nativeDecodedPartKey(source.PartName)
		if targetErr != nil || resolvedErr != nil || partErr != nil || resolvedTargetKey != partKey {
			return fmt.Errorf("numbering relationship target does not resolve to numbering part")
		}
	} else if input.SourceParts.NumberingPart != nil {
		return fmt.Errorf("numbering part is missing its exact relationship/hash attestation")
	}
	collections := []struct {
		name   string
		length int
	}{
		{"paragraphs", len(input.Paragraphs)}, {"runs", len(input.Runs)}, {"tables", len(input.Tables)},
		{"fonts", len(input.Fonts)}, {"diagnostics", len(input.Diagnostics)},
	}
	for _, collection := range collections {
		if collection.length > NativeDOCXMaxCollectionItems {
			return fmt.Errorf("%s exceeds %d entries", collection.name, NativeDOCXMaxCollectionItems)
		}
	}
	if len(input.Diagnostics) > NativeDOCXMaxResolvedDiagnostics {
		return fmt.Errorf("diagnostics exceeds %d entries", NativeDOCXMaxResolvedDiagnostics)
	}
	paragraphs := map[string]bool{}
	for _, paragraph := range input.Paragraphs {
		if !nativeIDPattern.MatchString(paragraph.ParagraphID) || paragraphs[paragraph.ParagraphID] {
			return fmt.Errorf("invalid or duplicate paragraph id %q", paragraph.ParagraphID)
		}
		if paragraph.StyleID != nil && !nativeIDPattern.MatchString(*paragraph.StyleID) {
			return fmt.Errorf("invalid paragraph style id %q", *paragraph.StyleID)
		}
		if err := validateNativeResolvedStyleChain(paragraph.AppliedStyles, NativeDOCXMaxDepth); err != nil {
			return fmt.Errorf("paragraph %q: %w", paragraph.ParagraphID, err)
		}
		if err := validateNativeResolvedParagraphProperties(paragraph.Properties); err != nil {
			return fmt.Errorf("paragraph %q: %w", paragraph.ParagraphID, err)
		}
		if err := validateNativeResolvedRunProperties(paragraph.ParagraphMarkProperties); err != nil {
			return fmt.Errorf("paragraph %q mark: %w", paragraph.ParagraphID, err)
		}
		if paragraph.Numbering != nil {
			partSHA256 := ""
			if input.NumberingSource != nil {
				partSHA256 = input.NumberingSource.PartSHA256
			}
			if err := validateNativeResolvedNumbering(*paragraph.Numbering, partSHA256); err != nil {
				return fmt.Errorf("paragraph %q numbering: %w", paragraph.ParagraphID, err)
			}
			if paragraph.Numbering.LevelStyleID != nil {
				found := false
				for _, appliedStyleID := range paragraph.AppliedStyles {
					found = found || appliedStyleID == *paragraph.Numbering.LevelStyleID
				}
				if !found {
					return fmt.Errorf("paragraph %q numbering level style does not attest its style cascade", paragraph.ParagraphID)
				}
			}
		}
		paragraphs[paragraph.ParagraphID] = true
	}
	runs := map[string]bool{}
	for _, run := range input.Runs {
		if !nativeIDPattern.MatchString(run.RunID) || runs[run.RunID] || !paragraphs[run.ParagraphID] {
			return fmt.Errorf("invalid run reference %q", run.RunID)
		}
		if run.CharacterStyle != nil && !nativeIDPattern.MatchString(*run.CharacterStyle) {
			return fmt.Errorf("invalid character style id %q", *run.CharacterStyle)
		}
		if err := validateNativeResolvedStyleChain(run.AppliedParagraphStyles, NativeDOCXMaxDepth); err != nil {
			return fmt.Errorf("run %q paragraph styles: %w", run.RunID, err)
		}
		if err := validateNativeResolvedStyleChain(run.AppliedCharacterStyles, NativeDOCXMaxDepth); err != nil {
			return fmt.Errorf("run %q character styles: %w", run.RunID, err)
		}
		if err := validateNativeResolvedRunProperties(run.Properties); err != nil {
			return fmt.Errorf("run %q: %w", run.RunID, err)
		}
		runs[run.RunID] = true
	}
	tables := map[string]bool{}
	for _, table := range input.Tables {
		if !nativeIDPattern.MatchString(table.TableID) || tables[table.TableID] {
			return fmt.Errorf("invalid or duplicate table id %q", table.TableID)
		}
		if table.StyleID != nil && !nativeIDPattern.MatchString(*table.StyleID) {
			return fmt.Errorf("invalid table style id %q", *table.StyleID)
		}
		if table.CellShadingRGB != nil {
			if _, ok := nativeExactRGB(*table.CellShadingRGB); !ok {
				return fmt.Errorf("invalid table style cell shading %q", *table.CellShadingRGB)
			}
		}
		tables[table.TableID] = true
	}
	fontNames := map[string]bool{}
	for _, font := range input.Fonts {
		if !nativeBoundedResolvedString(font.Name, 256) {
			return fmt.Errorf("invalid resolved font name")
		}
		key := strings.ToLower(font.Name)
		if fontNames[key] {
			return fmt.Errorf("duplicate resolved font %q", font.Name)
		}
		fontNames[key] = true
		if font.AltName != nil && !nativeBoundedResolvedString(*font.AltName, 256) {
			return fmt.Errorf("invalid resolved alternate font name")
		}
	}
	for _, diagnostic := range input.Diagnostics {
		if !nativeIDPattern.MatchString(diagnostic.Code) || !nativeBoundedResolvedString(diagnostic.Message, NativeDOCXMaxTextLength) || diagnostic.Preservation != "preserve-verbatim" || diagnostic.Severity != "unsupported" {
			return fmt.Errorf("invalid resolution diagnostic")
		}
		if diagnostic.ScopeID != input.DocumentID && !paragraphs[diagnostic.ScopeID] && !runs[diagnostic.ScopeID] && !tables[diagnostic.ScopeID] {
			return fmt.Errorf("diagnostic has unknown scope %q", diagnostic.ScopeID)
		}
		if diagnostic.PartName != nil {
			if err := validateNativePartName(*diagnostic.PartName); err != nil {
				return fmt.Errorf("diagnostic has invalid source part")
			}
		}
		if diagnostic.Path != nil && (diagnostic.PartName == nil || !nativeBoundedResolvedString(*diagnostic.Path, 8192)) {
			return fmt.Errorf("diagnostic has invalid source path")
		}
	}
	if input.NumberingSource != nil && input.NumberingSource.ModelSHA256 != nativeResolvedNumberingModelSHA256(input) {
		return fmt.Errorf("numbering model digest does not attest resolved markers")
	}
	return nil
}

func validateNativeResolvedStyleChain(chain []string, limit int) error {
	if len(chain) > limit {
		return fmt.Errorf("style chain exceeds %d layers", limit)
	}
	seen := map[string]bool{}
	for _, styleID := range chain {
		if !nativeIDPattern.MatchString(styleID) {
			return fmt.Errorf("invalid applied style id %q", styleID)
		}
		if seen[styleID] {
			return fmt.Errorf("duplicate applied style id %q", styleID)
		}
		seen[styleID] = true
	}
	return nil
}

func validateNativeResolvedParagraphProperties(properties NativeResolvedParagraphPropertiesV1) error {
	if properties.Alignment != nil && !nativeParagraphAlignment(*properties.Alignment) {
		return fmt.Errorf("invalid alignment")
	}
	for _, value := range []*int64{properties.SpacingBeforeTwips, properties.SpacingAfterTwips, properties.Line, properties.FirstLineTwips, properties.HangingTwips} {
		if value != nil && (*value < 0 || *value > nativeMaxTwipsForMilliPoints) {
			return fmt.Errorf("paragraph measurement is outside the safe range")
		}
	}
	for _, value := range []*int64{properties.IndentLeftTwips, properties.IndentRightTwips, properties.IndentStartTwips, properties.IndentEndTwips} {
		if value != nil && (*value < -nativeMaxTwipsForMilliPoints || *value > nativeMaxTwipsForMilliPoints) {
			return fmt.Errorf("paragraph indent is outside the safe range")
		}
	}
	if properties.FirstLineTwips != nil && properties.HangingTwips != nil {
		return fmt.Errorf("first-line and hanging indents are mutually exclusive")
	}
	if properties.LineRule != nil {
		if properties.Line == nil || (*properties.LineRule != "auto" && *properties.LineRule != "exact" && *properties.LineRule != "atLeast") {
			return fmt.Errorf("invalid line rule")
		}
	} else if properties.Line != nil {
		return fmt.Errorf("line measurement has no unit rule")
	}
	return nil
}

func validateNativeResolvedRunProperties(properties NativeResolvedRunPropertiesV1) error {
	if properties.FontFamily != nil && !nativeBoundedResolvedString(*properties.FontFamily, 256) {
		return fmt.Errorf("invalid font family")
	}
	if properties.FontSizeHalfPoint != nil && (*properties.FontSizeHalfPoint <= 0 || *properties.FontSizeHalfPoint > 3276) {
		return fmt.Errorf("invalid font size")
	}
	if properties.Underline != nil && *properties.Underline != "none" && *properties.Underline != "single" && *properties.Underline != "double" && *properties.Underline != "words" {
		return fmt.Errorf("invalid underline")
	}
	if properties.VerticalAlignment != nil && *properties.VerticalAlignment != "baseline" && *properties.VerticalAlignment != "subscript" && *properties.VerticalAlignment != "superscript" {
		return fmt.Errorf("invalid vertical alignment")
	}
	if properties.Color != nil && (*properties.Color == "auto" || !nativeColor.MatchString(*properties.Color)) {
		return fmt.Errorf("invalid explicit color")
	}
	if properties.Highlight != nil && !nativeResolvedHighlight(*properties.Highlight) {
		return fmt.Errorf("invalid highlight")
	}
	if properties.Language != nil && !nativeBoundedResolvedString(*properties.Language, 256) {
		return fmt.Errorf("invalid language")
	}
	return nil
}

func validateNativeResolvedNumbering(numbering NativeResolvedNumberingV1, partSHA256 string) error {
	if !nativeIDPattern.MatchString(numbering.MarkerID) || !nativeSHA256.MatchString(numbering.DefinitionSHA256) {
		return fmt.Errorf("invalid marker identity/digest")
	}
	if canonical, ok := nativeCanonicalDecimalID(numbering.NumID); !ok || canonical != numbering.NumID || numbering.NumID == "0" {
		return fmt.Errorf("invalid num id")
	}
	if canonical, ok := nativeCanonicalDecimalID(numbering.AbstractNumID); !ok || canonical != numbering.AbstractNumID {
		return fmt.Errorf("invalid abstract num id")
	}
	if numbering.Level < 0 || numbering.Level > 8 || numbering.Start < 0 || numbering.Start > 2147483647 || numbering.CounterValue < 0 || numbering.CounterValue > 2147483647 {
		return fmt.Errorf("invalid level/start")
	}
	if numbering.LevelStyleID != nil && !nativeIDPattern.MatchString(*numbering.LevelStyleID) {
		return fmt.Errorf("invalid numbering level style id")
	}
	if !nativeOrdinaryNumberFormat(numbering.Format) {
		return fmt.Errorf("invalid ordinary number format")
	}
	if !nativeBoundedNumberingString(numbering.Text, 1024) || !nativeBoundedNumberingString(numbering.ResolvedText, 1024) {
		return fmt.Errorf("invalid level text")
	}
	if utf8.RuneCountInString(numbering.ResolvedText) > 31 {
		return fmt.Errorf("resolved numbering text exceeds 31 Unicode scalars")
	}
	if numbering.Suffix != "tab" && numbering.Suffix != "space" && numbering.Suffix != "nothing" {
		return fmt.Errorf("invalid numbering suffix")
	}
	if numbering.Alignment != "left" && numbering.Alignment != "right" && numbering.Alignment != "center" && numbering.Alignment != "start" && numbering.Alignment != "end" {
		return fmt.Errorf("invalid numbering alignment")
	}
	if numbering.RestartAfterLevel != nil && (*numbering.RestartAfterLevel < 0 || *numbering.RestartAfterLevel >= numbering.Level || numbering.NeverRestart) {
		return fmt.Errorf("invalid numbering restart policy")
	}
	if numbering.RestartAfterLevel == nil && !numbering.NeverRestart && numbering.Level > 0 {
		return fmt.Errorf("incomplete numbering restart policy")
	}
	if len(numbering.CounterValues) > 9 {
		return fmt.Errorf("too many numbering counter values")
	}
	seen := map[int]bool{}
	for _, counter := range numbering.CounterValues {
		if counter.Level < 0 || counter.Level > numbering.Level || seen[counter.Level] || counter.Value < 0 || counter.Value > 2147483647 || !nativeOrdinaryNumberFormat(counter.Format) || counter.Format == "bullet" {
			return fmt.Errorf("invalid numbering counter vector")
		}
		seen[counter.Level] = true
	}
	if numbering.LabelStartTwips < 0 || numbering.LabelEndTwips <= numbering.LabelStartTwips || numbering.TextStartTwips != numbering.LabelEndTwips || numbering.LabelStartTwips > nativeMaxTwipsForMilliPoints || numbering.LabelEndTwips > nativeMaxTwipsForMilliPoints {
		return fmt.Errorf("invalid numbering label geometry")
	}
	if numbering.NumberingTabTwips != nil && (*numbering.NumberingTabTwips < 0 || *numbering.NumberingTabTwips > nativeMaxTwipsForMilliPoints) {
		return fmt.Errorf("invalid numbering tab stop")
	}
	if !nativeSHA256.MatchString(partSHA256) || numbering.DefinitionSHA256 != nativeResolvedNumberingDefinitionSHA256(&numbering, partSHA256) {
		return fmt.Errorf("numbering definition digest does not attest exported semantics")
	}
	return validateNativeResolvedRunProperties(numbering.Marker)
}

func nativeBoundedNumberingString(value string, limit int) bool {
	return value != "" && len(value) <= limit && utf8.ValidString(value) && !strings.ContainsAny(value, "\x00\r\n")
}

func nativeBoundedResolvedString(value string, limit int) bool {
	return value != "" && len(value) <= limit && strings.TrimSpace(value) == value && !strings.ContainsAny(value, "\x00\r\n")
}
