package docxpatch

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"net/url"
	"regexp"
	"sort"
	"strings"
	"unicode/utf8"
)

const (
	NativeDOCXProtocol           = "injoffice.docx.native"
	NativeDOCXVersion            = 1
	NativeDOCXV1MediaType        = "application/vnd.injoffice.docx-native.v1+json"
	NativeDOCXMaxJSONBytes       = 8 * 1024 * 1024
	NativeDOCXMaxDepth           = 64
	NativeDOCXMaxNodes           = 100_000
	NativeDOCXMaxCollectionItems = 10_000
	NativeDOCXMaxTextLength      = 1_048_576
	NativeDOCXMaxIssues          = 100
)

type NativeSourcePackageV1 struct {
	PackageSHA256 string `json:"package_sha256"`
	MainPart      string `json:"main_part"`
}

type NativeSourceAnchorV1 struct {
	PartName  string `json:"part_name"`
	Path      string `json:"path"`
	StartByte *int64 `json:"start_byte"`
	EndByte   *int64 `json:"end_byte"`
	XMLSHA256 string `json:"xml_sha256"`
}

type NativeRefusalV1 struct {
	Code         string `json:"code"`
	Message      string `json:"message"`
	Preservation string `json:"preservation"`
}

type NativeEditPolicyV1 struct {
	Mode              string           `json:"mode"`
	AllowedOperations []string         `json:"allowed_operations"`
	Refusal           *NativeRefusalV1 `json:"refusal,omitempty"`
}

type NativeCapabilityV1 struct {
	Name   string  `json:"name"`
	Level  string  `json:"level"`
	Detail *string `json:"detail,omitempty"`
}

type NativePassthroughPartV1 struct {
	PartName    string `json:"part_name"`
	ContentType string `json:"content_type"`
	ByteLength  *int64 `json:"byte_length"`
	SHA256      string `json:"sha256"`
	Policy      string `json:"policy"`
}

type NativeRunPropertiesV1 struct {
	CharacterStyleID  *string `json:"character_style_id,omitempty"`
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

type NativeDrawingCropV1 struct {
	Left   *int64 `json:"left"`
	Top    *int64 `json:"top"`
	Right  *int64 `json:"right"`
	Bottom *int64 `json:"bottom"`
}

type NativeDrawingV1 struct {
	ID                     string               `json:"id"`
	Anchor                 NativeSourceAnchorV1 `json:"anchor"`
	RelationshipID         *string              `json:"relationship_id,omitempty"`
	MediaPart              *string              `json:"media_part,omitempty"`
	ContentType            *string              `json:"content_type,omitempty"`
	Name                   *string              `json:"name,omitempty"`
	AltText                *string              `json:"alt_text,omitempty"`
	Placement              string               `json:"placement"`
	WidthEMU               *int64               `json:"width_emu"`
	HeightEMU              *int64               `json:"height_emu"`
	RotationDegrees        *int64               `json:"rotation_degrees,omitempty"`
	FlipHorizontal         *bool                `json:"flip_horizontal,omitempty"`
	FlipVertical           *bool                `json:"flip_vertical,omitempty"`
	SourceCrop             *NativeDrawingCropV1 `json:"source_crop,omitempty"`
	XEMU                   *int64               `json:"x_emu,omitempty"`
	YEMU                   *int64               `json:"y_emu,omitempty"`
	HorizontalRelativeFrom *string              `json:"horizontal_relative_from,omitempty"`
	VerticalRelativeFrom   *string              `json:"vertical_relative_from,omitempty"`
	Wrap                   *string              `json:"wrap,omitempty"`
	EditPolicy             NativeEditPolicyV1   `json:"edit_policy"`
}

type NativeReferenceV1 struct {
	Kind     string `json:"kind"`
	TargetID string `json:"target_id"`
	Role     string `json:"role,omitempty"`
}

type NativeRunV1 struct {
	Kind       string                 `json:"kind"`
	ID         string                 `json:"id"`
	Anchor     NativeSourceAnchorV1   `json:"anchor"`
	Properties *NativeRunPropertiesV1 `json:"properties,omitempty"`
	Text       *string                `json:"text,omitempty"`
	PageField  string                 `json:"page_field,omitempty"`
	Control    string                 `json:"control,omitempty"`
	Reference  *NativeReferenceV1     `json:"reference,omitempty"`
	Drawing    *NativeDrawingV1       `json:"drawing,omitempty"`
}

type NativeNumberingReferenceV1 struct {
	NumID         string  `json:"num_id"`
	Level         *int    `json:"level"`
	AbstractNumID *string `json:"abstract_num_id,omitempty"`
}

type NativeParagraphPropertiesV1 struct {
	ParagraphStyleID *string                     `json:"paragraph_style_id,omitempty"`
	Numbering        *NativeNumberingReferenceV1 `json:"numbering,omitempty"`
	Alignment        *string                     `json:"alignment,omitempty"`
	KeepNext         *bool                       `json:"keep_next,omitempty"`
	KeepLines        *bool                       `json:"keep_lines,omitempty"`
	PageBreakBefore  *bool                       `json:"page_break_before,omitempty"`
	WidowControl     *bool                       `json:"widow_control,omitempty"`
}

type NativeParagraphV1 struct {
	ID         string                       `json:"id"`
	Anchor     NativeSourceAnchorV1         `json:"anchor"`
	EditPolicy NativeEditPolicyV1           `json:"edit_policy"`
	Properties *NativeParagraphPropertiesV1 `json:"properties"`
	Runs       []NativeRunV1                `json:"runs"`
}

type NativeTableCellV1 struct {
	ID            string                `json:"id"`
	Anchor        NativeSourceAnchorV1  `json:"anchor"`
	WidthTwips    *int64                `json:"width_twips,omitempty"`
	GridSpan      *int                  `json:"grid_span"`
	VerticalMerge string                `json:"vertical_merge"`
	Borders       *NativeTableBordersV1 `json:"borders,omitempty"`
	ShadingRGB    *string               `json:"shading_rgb,omitempty"`
	Paragraphs    []NativeParagraphV1   `json:"paragraphs"`
}

type NativeTableRowV1 struct {
	ID           string               `json:"id"`
	Anchor       NativeSourceAnchorV1 `json:"anchor"`
	HeightTwips  *int64               `json:"height_twips,omitempty"`
	HeightRule   *string              `json:"height_rule,omitempty"`
	RepeatHeader *bool                `json:"repeat_header"`
	CantSplit    *bool                `json:"cant_split,omitempty"`
	Cells        []NativeTableCellV1  `json:"cells"`
}

type NativeTableBorderV1 struct {
	Style            string  `json:"style"`
	SizeEighthPoints int64   `json:"size_eighth_points"`
	ColorRGB         *string `json:"color_rgb,omitempty"`
}

type NativeTableBordersV1 struct {
	Top              *NativeTableBorderV1 `json:"top,omitempty"`
	Right            *NativeTableBorderV1 `json:"right,omitempty"`
	Bottom           *NativeTableBorderV1 `json:"bottom,omitempty"`
	Left             *NativeTableBorderV1 `json:"left,omitempty"`
	InsideHorizontal *NativeTableBorderV1 `json:"inside_horizontal,omitempty"`
	InsideVertical   *NativeTableBorderV1 `json:"inside_vertical,omitempty"`
}

type NativeTableCellMarginsV1 struct {
	TopTwips    int64 `json:"top_twips"`
	RightTwips  int64 `json:"right_twips"`
	BottomTwips int64 `json:"bottom_twips"`
	LeftTwips   int64 `json:"left_twips"`
}

type NativeTableV1 struct {
	ID              string                    `json:"id"`
	Anchor          NativeSourceAnchorV1      `json:"anchor"`
	EditPolicy      NativeEditPolicyV1        `json:"edit_policy"`
	TableStyleID    *string                   `json:"table_style_id,omitempty"`
	WidthTwips      *int64                    `json:"width_twips,omitempty"`
	Layout          *string                   `json:"layout,omitempty"`
	Alignment       *string                   `json:"alignment,omitempty"`
	IndentTwips     *int64                    `json:"indent_twips,omitempty"`
	GridWidthsTwips []int64                   `json:"grid_widths_twips,omitempty"`
	CellMargins     *NativeTableCellMarginsV1 `json:"cell_margins,omitempty"`
	Borders         *NativeTableBordersV1     `json:"borders,omitempty"`
	Rows            []NativeTableRowV1        `json:"rows"`
}

type NativeBlockV1 struct {
	Kind      string             `json:"kind"`
	ID        string             `json:"id"`
	Paragraph *NativeParagraphV1 `json:"paragraph,omitempty"`
	Table     *NativeTableV1     `json:"table,omitempty"`
}

type NativeStoryV1 struct {
	ID             string                `json:"id"`
	Kind           string                `json:"kind"`
	PartName       string                `json:"part_name"`
	NativeStoryID  *string               `json:"native_story_id,omitempty"`
	RelationshipID *string               `json:"relationship_id,omitempty"`
	NoteRole       string                `json:"note_role,omitempty"`
	Anchor         *NativeSourceAnchorV1 `json:"anchor"`
	Blocks         []NativeBlockV1       `json:"blocks"`
}

type NativeHeaderFooterReferenceV1 struct {
	Kind           string `json:"kind"`
	StoryID        string `json:"story_id"`
	RelationshipID string `json:"relationship_id"`
}

type NativePageMarginsV1 struct {
	TopTwips    *int64 `json:"top_twips"`
	RightTwips  *int64 `json:"right_twips"`
	BottomTwips *int64 `json:"bottom_twips"`
	LeftTwips   *int64 `json:"left_twips"`
	HeaderTwips *int64 `json:"header_twips"`
	FooterTwips *int64 `json:"footer_twips"`
	GutterTwips *int64 `json:"gutter_twips"`
}

type NativePageGeometryV1 struct {
	WidthTwips         *int64              `json:"width_twips"`
	HeightTwips        *int64              `json:"height_twips"`
	Orientation        string              `json:"orientation"`
	Margins            NativePageMarginsV1 `json:"margins"`
	Columns            *int                `json:"columns"`
	ColumnSpacingTwips *int64              `json:"column_spacing_twips"`
	ColumnLayout       string              `json:"column_layout"`
	ColumnDefinitions  []NativeColumnV1    `json:"column_definitions"`
}

type NativeColumnV1 struct {
	ID              string `json:"id"`
	Ordinal         *int   `json:"ordinal"`
	WidthTwips      *int64 `json:"width_twips,omitempty"`
	SpaceAfterTwips *int64 `json:"space_after_twips,omitempty"`
}

func nativeIntValue(value *int) int {
	if value == nil {
		return -1
	}
	return *value
}

func nativeInt64Value(value *int64) int64 {
	if value == nil {
		return -1
	}
	return *value
}

func nativeBoundedLength(length, maximum int) int {
	if length < maximum {
		return length
	}
	return maximum
}

type NativeSectionV1 struct {
	ID              string                          `json:"id"`
	Anchor          NativeSourceAnchorV1            `json:"anchor"`
	StartsAtBlockID string                          `json:"starts_at_block_id"`
	BreakType       string                          `json:"break_type"`
	TitlePage       *bool                           `json:"title_page"`
	Page            NativePageGeometryV1            `json:"page"`
	HeaderRefs      []NativeHeaderFooterReferenceV1 `json:"header_refs"`
	FooterRefs      []NativeHeaderFooterReferenceV1 `json:"footer_refs"`
}

type NativeCommentV1 struct {
	ID              string                `json:"id"`
	NativeCommentID string                `json:"native_comment_id"`
	Author          string                `json:"author"`
	Initials        *string               `json:"initials,omitempty"`
	CreatedAt       *string               `json:"created_at,omitempty"`
	Anchor          *NativeSourceAnchorV1 `json:"anchor"`
	BodyStoryID     string                `json:"body_story_id"`
}

type NativeUnsupportedCapabilityV1 struct {
	ID           string                `json:"id"`
	Code         string                `json:"code"`
	Capability   string                `json:"capability"`
	ScopeID      string                `json:"scope_id"`
	Anchor       *NativeSourceAnchorV1 `json:"anchor,omitempty"`
	Preservation string                `json:"preservation"`
	Message      string                `json:"message"`
}

type NativeDocumentV1 struct {
	Protocol         string                          `json:"protocol"`
	Version          *int                            `json:"version"`
	DocumentID       string                          `json:"document_id"`
	Revision         string                          `json:"revision"`
	Source           NativeSourcePackageV1           `json:"source"`
	Body             NativeStoryV1                   `json:"body"`
	Sections         []NativeSectionV1               `json:"sections"`
	Headers          []NativeStoryV1                 `json:"headers"`
	Footers          []NativeStoryV1                 `json:"footers"`
	Notes            []NativeStoryV1                 `json:"notes"`
	CommentStories   []NativeStoryV1                 `json:"comment_stories"`
	Comments         []NativeCommentV1               `json:"comments"`
	Capabilities     []NativeCapabilityV1            `json:"capabilities"`
	PassthroughParts []NativePassthroughPartV1       `json:"passthrough_parts"`
	Unsupported      []NativeUnsupportedCapabilityV1 `json:"unsupported"`
}

type NativeValidationIssue struct {
	Code    string `json:"code"`
	Path    string `json:"path"`
	Message string `json:"message"`
}

type NativeValidationError struct {
	Issues []NativeValidationIssue `json:"issues"`
}

func (e *NativeValidationError) Error() string {
	if len(e.Issues) == 0 {
		return "invalid native DOCX document"
	}
	return fmt.Sprintf("invalid native DOCX document at %s: %s", e.Issues[0].Path, e.Issues[0].Message)
}

var (
	nativeIDPattern     = regexp.MustCompile(`^[A-Za-z0-9][A-Za-z0-9._:-]{0,255}$`)
	nativeNoteIDPattern = regexp.MustCompile(`^[1-9][0-9]{0,18}$`)
	nativeSHA256        = regexp.MustCompile(`^sha256:[0-9a-f]{64}$`)
	nativeColor         = regexp.MustCompile(`^(?:auto|[0-9A-F]{6})$`)
	nativePartSegment   = regexp.MustCompile(`^(?:[A-Za-z0-9._~!$&'()*+,;=@-]|%[0-9A-F]{2})+$`)
	nativeNegativeZero  = regexp.MustCompile(`^-0(?:\.0*)?(?:[eE][+-]?[0-9]+)?$`)
	nativeOperations    = map[string]bool{
		"text.replace": true, "properties.patch": true, "block.insert_after": true,
		"block.delete": true, "drawing.replace": true,
	}
	nativeParagraphOperations = map[string]bool{"text.replace": true, "properties.patch": true, "block.insert_after": true, "block.delete": true}
	nativeTableOperations     = map[string]bool{"properties.patch": true, "block.insert_after": true, "block.delete": true}
	nativeDrawingOperations   = map[string]bool{"drawing.replace": true}
)

// DecodeNativeDocumentV1 rejects unknown JSON fields and validates identities,
// unions, source anchors, references, and fail-closed edit metadata.
func DecodeNativeDocumentV1(data []byte) (*NativeDocumentV1, error) {
	if len(data) > NativeDOCXMaxJSONBytes {
		return nil, &NativeValidationError{Issues: []NativeValidationIssue{{Code: "LIMIT_EXCEEDED", Path: "", Message: fmt.Sprintf("JSON payload exceeds %d bytes", NativeDOCXMaxJSONBytes)}}}
	}
	if issues, err := preflightNativeJSON(data); err != nil {
		return nil, fmt.Errorf("docxpatch: decode native DOCX v1: %w", err)
	} else if len(issues) > 0 {
		return nil, &NativeValidationError{Issues: issues}
	}
	decoder := json.NewDecoder(bytes.NewReader(data))
	decoder.DisallowUnknownFields()
	var doc NativeDocumentV1
	if err := decoder.Decode(&doc); err != nil {
		return nil, fmt.Errorf("docxpatch: decode native DOCX v1: %w", err)
	}
	if err := requireJSONEOF(decoder); err != nil {
		return nil, err
	}
	if issues := ValidateNativeDocumentV1(&doc); len(issues) > 0 {
		return nil, &NativeValidationError{Issues: issues}
	}
	return &doc, nil
}

type nativeJSONScanner struct {
	issues []NativeValidationIssue
	nodes  int
	limit  bool
}

func preflightNativeJSON(data []byte) ([]NativeValidationIssue, error) {
	decoder := json.NewDecoder(bytes.NewReader(data))
	decoder.UseNumber()
	scanner := &nativeJSONScanner{}
	if err := scanner.value(decoder, "", 0); err != nil {
		return nil, err
	}
	if scanner.limit {
		sortNativeIssues(scanner.issues)
		return scanner.issues, nil
	}
	if token, err := decoder.Token(); err != io.EOF {
		if err != nil {
			return nil, err
		}
		return nil, fmt.Errorf("multiple JSON values after %v", token)
	}
	sortNativeIssues(scanner.issues)
	return scanner.issues, nil
}

func (s *nativeJSONScanner) value(decoder *json.Decoder, path string, depth int) error {
	if depth > NativeDOCXMaxDepth {
		s.add("LIMIT_EXCEEDED", path, fmt.Sprintf("document nesting exceeds %d levels", NativeDOCXMaxDepth))
		s.limit = true
		return nil
	}
	s.nodes++
	if s.nodes > NativeDOCXMaxNodes {
		s.add("LIMIT_EXCEEDED", path, fmt.Sprintf("document traversal exceeds %d values", NativeDOCXMaxNodes))
		s.limit = true
		return nil
	}
	token, err := decoder.Token()
	if err != nil {
		return err
	}
	switch value := token.(type) {
	case nil:
		s.add("INVALID_VALUE", path, "JSON null is not permitted anywhere in the native contract")
	case json.Number:
		if nativeNegativeZero.MatchString(value.String()) {
			s.add("INVALID_VALUE", path, "negative zero is not permitted")
		}
	case json.Delim:
		switch value {
		case '{':
			seen := map[string]bool{}
			for decoder.More() {
				keyToken, err := decoder.Token()
				if err != nil {
					return err
				}
				key, ok := keyToken.(string)
				if !ok {
					return fmt.Errorf("object key is not a string")
				}
				childPath := path + "/" + nativeEscapePointer(key)
				if seen[key] {
					s.add("INVALID_VALUE", childPath, "object field is duplicated")
				}
				seen[key] = true
				if err := s.value(decoder, childPath, depth+1); err != nil {
					return err
				}
				if s.limit {
					return nil
				}
			}
			_, err = decoder.Token()
			return err
		case '[':
			count := 0
			for decoder.More() {
				if count >= NativeDOCXMaxCollectionItems {
					s.add("LIMIT_EXCEEDED", path, fmt.Sprintf("must contain at most %d items", NativeDOCXMaxCollectionItems))
					s.limit = true
					return nil
				}
				if err := s.value(decoder, fmt.Sprintf("%s/%d", path, count), depth+1); err != nil {
					return err
				}
				if s.limit {
					return nil
				}
				count++
			}
			_, err = decoder.Token()
			return err
		}
	}
	return nil
}

func (s *nativeJSONScanner) add(code, path, message string) {
	if len(s.issues) >= NativeDOCXMaxIssues {
		return
	}
	for _, issue := range s.issues {
		if issue.Code == code && issue.Path == path && issue.Message == message {
			return
		}
	}
	s.issues = append(s.issues, NativeValidationIssue{Code: code, Path: path, Message: message})
}

func nativeEscapePointer(value string) string {
	return strings.ReplaceAll(strings.ReplaceAll(value, "~", "~0"), "/", "~1")
}

func requireJSONEOF(decoder *json.Decoder) error {
	var extra any
	if err := decoder.Decode(&extra); err == io.EOF {
		return nil
	} else if err != nil {
		return fmt.Errorf("docxpatch: decode native DOCX v1 trailer: %w", err)
	}
	return fmt.Errorf("docxpatch: decode native DOCX v1: multiple JSON values")
}

// EncodeNativeDocumentV1 validates then emits deterministic JSON. Object keys
// are lexicographically sorted; source-order arrays are never reordered.
func EncodeNativeDocumentV1(doc *NativeDocumentV1) ([]byte, error) {
	if issues := ValidateNativeDocumentV1(doc); len(issues) > 0 {
		return nil, &NativeValidationError{Issues: issues}
	}
	raw, err := json.Marshal(doc)
	if err != nil {
		return nil, fmt.Errorf("docxpatch: encode native DOCX v1: %w", err)
	}
	if len(raw) > NativeDOCXMaxJSONBytes {
		return nil, &NativeValidationError{Issues: []NativeValidationIssue{{Code: "LIMIT_EXCEEDED", Path: "", Message: fmt.Sprintf("JSON payload exceeds %d bytes", NativeDOCXMaxJSONBytes)}}}
	}
	var value any
	decoder := json.NewDecoder(bytes.NewReader(raw))
	decoder.UseNumber()
	if err := decoder.Decode(&value); err != nil {
		return nil, fmt.Errorf("docxpatch: canonicalize native DOCX v1: %w", err)
	}
	var out bytes.Buffer
	encoder := json.NewEncoder(&out)
	encoder.SetEscapeHTML(false)
	if err := encoder.Encode(value); err != nil {
		return nil, fmt.Errorf("docxpatch: canonicalize native DOCX v1: %w", err)
	}
	return bytes.TrimSuffix(out.Bytes(), []byte("\n")), nil
}

type nativeValidator struct {
	issues            []NativeValidationIssue
	ids               map[string]string
	references        []nativeReference
	passthrough       map[string]bool
	mediaParts        map[string]bool
	bodyBlocks        map[string]bool
	headerStories     map[string]bool
	footerStories     map[string]bool
	footnoteStories   map[string]bool
	endnoteStories    map[string]bool
	commentStories    map[string]bool
	comments          map[string]bool
	anchors           map[string]*nativeAnchorBounds
	storyNativeIDs    map[string]string
	noteNativeIDs     map[string]bool
	noteRelationships map[string]string
	noteRelByKind     map[string]string
	notePartByKind    map[string]string
	modeledParts      map[string]bool
	nodes             int
}

type nativeReference struct{ path, id, target string }

func nativeIsMediaContentType(value string) bool {
	folded := nativeASCIIFold(value)
	return strings.HasPrefix(folded, "image/") || strings.HasPrefix(folded, "audio/") || strings.HasPrefix(folded, "video/")
}

func ValidateNativeDocumentV1(doc *NativeDocumentV1) []NativeValidationIssue {
	v := &nativeValidator{
		ids: map[string]string{}, passthrough: map[string]bool{}, mediaParts: map[string]bool{}, bodyBlocks: map[string]bool{},
		headerStories: map[string]bool{}, footerStories: map[string]bool{}, footnoteStories: map[string]bool{},
		endnoteStories: map[string]bool{}, commentStories: map[string]bool{}, comments: map[string]bool{},
		anchors: map[string]*nativeAnchorBounds{}, storyNativeIDs: map[string]string{}, noteNativeIDs: map[string]bool{}, noteRelationships: map[string]string{}, noteRelByKind: map[string]string{}, notePartByKind: map[string]string{}, modeledParts: map[string]bool{},
	}
	if doc == nil {
		v.add("REQUIRED", "", "document is required")
		return v.issues
	}
	if doc.Protocol != NativeDOCXProtocol {
		v.add("UNSUPPORTED_PROTOCOL", "/protocol", "must equal "+NativeDOCXProtocol)
	}
	if doc.Version == nil || *doc.Version != NativeDOCXVersion {
		v.add("UNSUPPORTED_VERSION", "/version", fmt.Sprintf("must equal %d", NativeDOCXVersion))
	}
	v.id(doc.DocumentID, "/document_id")
	v.requiredID(doc.Revision, "/revision")
	v.sha(doc.Source.PackageSHA256, "/source/package_sha256")
	v.part(doc.Source.MainPart, "/source/main_part")
	if doc.Source.MainPart != "" {
		v.modeledParts[doc.Source.MainPart] = true
	}
	bodyAnchor := v.story(&doc.Body, "/body", map[string]bool{"body": true})
	if doc.Body.PartName != doc.Source.MainPart {
		v.add("INVALID_VALUE", "/body/part_name", "body story part must equal source.main_part")
	}
	if doc.Sections == nil || len(doc.Sections) == 0 {
		v.add("REQUIRED", "/sections", "at least one section is required")
	}
	for i := 0; i < v.collection(len(doc.Sections), "/sections", NativeDOCXMaxCollectionItems); i++ {
		v.section(&doc.Sections[i], fmt.Sprintf("/sections/%d", i), doc.Source.MainPart, bodyAnchor)
	}
	v.storyList(doc.Headers, "/headers", "header")
	v.storyList(doc.Footers, "/footers", "footer")
	if doc.Notes == nil {
		v.add("REQUIRED", "/notes", "field is required")
	}
	for i := 0; i < v.collection(len(doc.Notes), "/notes", NativeDOCXMaxCollectionItems); i++ {
		v.story(&doc.Notes[i], fmt.Sprintf("/notes/%d", i), map[string]bool{"footnote": true, "endnote": true})
	}
	if doc.CommentStories == nil {
		v.add("REQUIRED", "/comment_stories", "field is required")
	}
	for i := 0; i < v.collection(len(doc.CommentStories), "/comment_stories", NativeDOCXMaxCollectionItems); i++ {
		v.story(&doc.CommentStories[i], fmt.Sprintf("/comment_stories/%d", i), map[string]bool{"comment": true})
	}
	if doc.Comments == nil {
		v.add("REQUIRED", "/comments", "field is required")
	}
	for i := 0; i < v.collection(len(doc.Comments), "/comments", NativeDOCXMaxCollectionItems); i++ {
		path := fmt.Sprintf("/comments/%d", i)
		comment := &doc.Comments[i]
		v.id(comment.ID, path+"/id")
		if comment.ID != "" {
			v.comments[comment.ID] = true
		}
		v.requiredID(comment.NativeCommentID, path+"/native_comment_id")
		v.required(comment.Author, path+"/author")
		v.optionalString(comment.Initials, path+"/initials")
		v.optionalString(comment.CreatedAt, path+"/created_at")
		commentAnchor := v.anchor(comment.Anchor, path+"/anchor", "", nil)
		v.requiredID(comment.BodyStoryID, path+"/body_story_id")
		v.ref(comment.BodyStoryID, path+"/body_story_id", "comment-story")
		if storyAnchor := v.anchorByID(comment.BodyStoryID); commentAnchor != nil && storyAnchor != nil && !storyAnchor.within(commentAnchor) {
			v.add("OUT_OF_RANGE", path+"/body_story_id", "comment body story anchor must be contained by the owning comment anchor")
		}
		if nativeStoryID := v.storyNativeIDs[comment.BodyStoryID]; nativeStoryID != "" && nativeStoryID != comment.NativeCommentID {
			v.add("INVALID_VALUE", path+"/body_story_id", "comment body story native_story_id must equal native_comment_id")
		}
	}
	if doc.Capabilities == nil {
		v.add("REQUIRED", "/capabilities", "field is required")
	}
	capabilityNames := map[string]bool{}
	for i := 0; i < v.collection(len(doc.Capabilities), "/capabilities", NativeDOCXMaxCollectionItems); i++ {
		path := fmt.Sprintf("/capabilities/%d", i)
		capability := &doc.Capabilities[i]
		v.requiredID(capability.Name, path+"/name")
		if capabilityNames[capability.Name] {
			v.add("INVALID_VALUE", path+"/name", "capability name is duplicated")
		}
		capabilityNames[capability.Name] = true
		v.oneOf(capability.Level, path+"/level", "read-write", "read-only", "passthrough", "unsupported")
		v.optionalString(capability.Detail, path+"/detail")
	}
	if doc.PassthroughParts == nil {
		v.add("REQUIRED", "/passthrough_parts", "field is required")
	}
	for i := 0; i < v.collection(len(doc.PassthroughParts), "/passthrough_parts", NativeDOCXMaxCollectionItems); i++ {
		path := fmt.Sprintf("/passthrough_parts/%d", i)
		part := &doc.PassthroughParts[i]
		v.part(part.PartName, path+"/part_name")
		if v.passthrough[part.PartName] {
			v.add("INVALID_VALUE", path+"/part_name", "passthrough part is duplicated")
		}
		if v.modeledParts[part.PartName] {
			v.add("INVALID_VALUE", path+"/part_name", "modeled story parts cannot also be passthrough parts")
		}
		v.passthrough[part.PartName] = true
		v.required(part.ContentType, path+"/content_type")
		if nativeIsMediaContentType(part.ContentType) {
			v.mediaParts[part.PartName] = true
		}
		v.nonnegative(part.ByteLength, path+"/byte_length")
		v.sha(part.SHA256, path+"/sha256")
		v.oneOf(part.Policy, path+"/policy", "preserve-verbatim")
	}
	if doc.Unsupported == nil {
		v.add("REQUIRED", "/unsupported", "field is required")
	}
	for i := 0; i < v.collection(len(doc.Unsupported), "/unsupported", NativeDOCXMaxCollectionItems); i++ {
		path := fmt.Sprintf("/unsupported/%d", i)
		entry := &doc.Unsupported[i]
		v.id(entry.ID, path+"/id")
		v.requiredID(entry.Code, path+"/code")
		v.requiredID(entry.Capability, path+"/capability")
		v.requiredID(entry.ScopeID, path+"/scope_id")
		v.ref(entry.ScopeID, path+"/scope_id", "modeled-id")
		if entry.Anchor != nil {
			v.anchor(entry.Anchor, path+"/anchor", "", nil)
		}
		v.oneOf(entry.Preservation, path+"/preservation", "preserve-verbatim", "refuse-mutation")
		v.required(entry.Message, path+"/message")
	}
	for _, reference := range v.references {
		if !v.targetExists(reference.target, reference.id) {
			v.add("BROKEN_REFERENCE", reference.path, fmt.Sprintf("reference %q does not resolve to a %s", reference.id, reference.target))
		}
	}
	sortNativeIssues(v.issues)
	return v.issues
}

func sortNativeIssues(issues []NativeValidationIssue) {
	sort.SliceStable(issues, func(i, j int) bool {
		if issues[i].Path != issues[j].Path {
			return issues[i].Path < issues[j].Path
		}
		if issues[i].Code != issues[j].Code {
			return issues[i].Code < issues[j].Code
		}
		return issues[i].Message < issues[j].Message
	})
}

type nativeAnchorBounds struct {
	part       string
	start, end int64
}

func (anchor *nativeAnchorBounds) within(parent *nativeAnchorBounds) bool {
	return anchor != nil && parent != nil && anchor.part == parent.part && anchor.start >= parent.start && anchor.end <= parent.end
}

func (v *nativeValidator) anchorByID(id string) *nativeAnchorBounds {
	return v.anchors[id]
}

func (v *nativeValidator) storyList(stories []NativeStoryV1, base, kind string) {
	if stories == nil {
		v.add("REQUIRED", base, "field is required")
		return
	}
	for i := 0; i < v.collection(len(stories), base, NativeDOCXMaxCollectionItems); i++ {
		v.story(&stories[i], fmt.Sprintf("%s/%d", base, i), map[string]bool{kind: true})
	}
}

func (v *nativeValidator) story(story *NativeStoryV1, path string, kinds map[string]bool) *nativeAnchorBounds {
	v.id(story.ID, path+"/id")
	if !kinds[story.Kind] {
		v.add("INVALID_VALUE", path+"/kind", "story kind is not valid in this collection")
	}
	v.part(story.PartName, path+"/part_name")
	if story.PartName != "" {
		v.modeledParts[story.PartName] = true
	}
	if story.Kind == "footnote" || story.Kind == "endnote" {
		if known := v.notePartByKind[story.Kind]; known != "" && known != story.PartName {
			v.add("INVALID_VALUE", path+"/part_name", "one note kind must resolve to exactly one package part")
		}
		v.notePartByKind[story.Kind] = story.PartName
		v.oneOf(story.NoteRole, path+"/note_role", "content", "separator", "continuation-separator")
		if story.RelationshipID == nil {
			v.add("REQUIRED", path+"/relationship_id", "note stories require their resolved main-document relationship")
		} else {
			v.optionalID(story.RelationshipID, path+"/relationship_id")
			if known := v.noteRelByKind[story.Kind]; known != "" && known != *story.RelationshipID {
				v.add("INVALID_VALUE", path+"/relationship_id", "one note kind must resolve through exactly one relationship identity")
			}
			v.noteRelByKind[story.Kind] = *story.RelationshipID
			if known := v.noteRelationships[*story.RelationshipID]; known != "" && known != story.Kind {
				v.add("DUPLICATE_ID", path+"/relationship_id", "one relationship identity cannot resolve both footnotes and endnotes")
			}
			v.noteRelationships[*story.RelationshipID] = story.Kind
		}
		if story.NativeStoryID == nil {
			v.add("REQUIRED", path+"/native_story_id", story.Kind+" stories require their native OOXML id")
		} else {
			valid := story.NoteRole == "content" && nativeNoteIDPattern.MatchString(*story.NativeStoryID) || story.NoteRole == "separator" && *story.NativeStoryID == "-1" || story.NoteRole == "continuation-separator" && *story.NativeStoryID == "0"
			if !valid {
				v.add("INVALID_VALUE", path+"/native_story_id", "native note id must exactly match its content or separator role")
			}
			key := story.Kind + ":" + *story.NativeStoryID
			if v.noteNativeIDs[key] {
				v.add("DUPLICATE_ID", path+"/native_story_id", "native note id is duplicated within its note kind")
			}
			v.noteNativeIDs[key] = true
		}
	} else if story.RelationshipID != nil || story.NoteRole != "" {
		v.add("INVALID_UNION", path, "relationship_id and note_role are reserved for note stories")
	} else {
		v.optionalID(story.NativeStoryID, path+"/native_story_id")
		if story.Kind == "comment" && story.NativeStoryID == nil {
			v.add("REQUIRED", path+"/native_story_id", "comment stories require their native OOXML id")
		}
	}
	if story.ID != "" && story.NativeStoryID != nil {
		v.storyNativeIDs[story.ID] = *story.NativeStoryID
	}
	storyAnchor := v.anchor(story.Anchor, path+"/anchor", story.PartName, nil)
	if story.ID != "" && storyAnchor != nil {
		v.anchors[story.ID] = storyAnchor
	}
	switch story.Kind {
	case "header":
		v.headerStories[story.ID] = true
	case "footer":
		v.footerStories[story.ID] = true
	case "footnote":
		v.footnoteStories[story.ID] = true
	case "endnote":
		v.endnoteStories[story.ID] = true
	case "comment":
		v.commentStories[story.ID] = true
	}
	if story.Blocks == nil {
		v.add("REQUIRED", path+"/blocks", "field is required")
	}
	for i := 0; i < v.collection(len(story.Blocks), path+"/blocks", NativeDOCXMaxCollectionItems); i++ {
		v.block(&story.Blocks[i], fmt.Sprintf("%s/blocks/%d", path, i), story.PartName, storyAnchor, story.Kind == "body")
	}
	return storyAnchor
}

func (v *nativeValidator) block(block *NativeBlockV1, path, ownerPart string, parentAnchor *nativeAnchorBounds, bodyBlock bool) {
	v.id(block.ID, path+"/id")
	if bodyBlock && block.ID != "" {
		v.bodyBlocks[block.ID] = true
	}
	switch block.Kind {
	case "paragraph":
		if block.Paragraph == nil || block.Table != nil {
			v.add("INVALID_UNION", path, "paragraph block must contain only paragraph payload")
			return
		}
		if block.Paragraph.ID != block.ID {
			v.add("INVALID_VALUE", path+"/paragraph/id", "payload id must equal block id")
		}
		v.paragraph(block.Paragraph, path+"/paragraph", false, ownerPart, parentAnchor)
	case "table":
		if block.Table == nil || block.Paragraph != nil {
			v.add("INVALID_UNION", path, "table block must contain only table payload")
			return
		}
		if block.Table.ID != block.ID {
			v.add("INVALID_VALUE", path+"/table/id", "payload id must equal block id")
		}
		v.table(block.Table, path+"/table", false, ownerPart, parentAnchor)
	default:
		v.add("INVALID_UNION", path+"/kind", "unknown block kind")
	}
}

func (v *nativeValidator) paragraph(paragraph *NativeParagraphV1, path string, track bool, ownerPart string, parentAnchor *nativeAnchorBounds) {
	if track {
		v.id(paragraph.ID, path+"/id")
	} else {
		v.requiredID(paragraph.ID, path+"/id")
	}
	paragraphAnchor := v.anchor(&paragraph.Anchor, path+"/anchor", ownerPart, parentAnchor)
	v.editPolicy(&paragraph.EditPolicy, path+"/edit_policy", nativeParagraphOperations)
	properties := paragraph.Properties
	if properties == nil {
		v.add("REQUIRED", path+"/properties", "field is required")
		properties = &NativeParagraphPropertiesV1{}
	}
	v.optionalID(properties.ParagraphStyleID, path+"/properties/paragraph_style_id")
	if properties.Numbering != nil {
		v.requiredID(properties.Numbering.NumID, path+"/properties/numbering/num_id")
		v.nonnegativeInt(properties.Numbering.Level, path+"/properties/numbering/level")
		v.optionalID(properties.Numbering.AbstractNumID, path+"/properties/numbering/abstract_num_id")
	}
	if properties.Alignment != nil {
		v.oneOf(*properties.Alignment, path+"/properties/alignment", "left", "center", "right", "both", "distribute")
	}
	if paragraph.Runs == nil {
		v.add("REQUIRED", path+"/runs", "field is required")
	}
	for i := 0; i < v.collection(len(paragraph.Runs), path+"/runs", NativeDOCXMaxCollectionItems); i++ {
		v.run(&paragraph.Runs[i], fmt.Sprintf("%s/runs/%d", path, i), ownerPart, paragraphAnchor)
	}
}

func (v *nativeValidator) table(table *NativeTableV1, path string, track bool, ownerPart string, parentAnchor *nativeAnchorBounds) {
	if track {
		v.id(table.ID, path+"/id")
	} else {
		v.requiredID(table.ID, path+"/id")
	}
	tableAnchor := v.anchor(&table.Anchor, path+"/anchor", ownerPart, parentAnchor)
	v.editPolicy(&table.EditPolicy, path+"/edit_policy", nativeTableOperations)
	v.optionalID(table.TableStyleID, path+"/table_style_id")
	v.optionalTwips(table.WidthTwips, path+"/width_twips", 1)
	v.optionalTwips(table.IndentTwips, path+"/indent_twips", 0)
	for i := range table.GridWidthsTwips {
		v.twips(&table.GridWidthsTwips[i], fmt.Sprintf("%s/grid_widths_twips/%d", path, i), 1)
	}
	if table.CellMargins != nil {
		v.twips(&table.CellMargins.TopTwips, path+"/cell_margins/top_twips", 0)
		v.twips(&table.CellMargins.RightTwips, path+"/cell_margins/right_twips", 0)
		v.twips(&table.CellMargins.BottomTwips, path+"/cell_margins/bottom_twips", 0)
		v.twips(&table.CellMargins.LeftTwips, path+"/cell_margins/left_twips", 0)
	}
	if table.Rows == nil {
		v.add("REQUIRED", path+"/rows", "field is required")
	}
	for i := 0; i < v.collection(len(table.Rows), path+"/rows", NativeDOCXMaxCollectionItems); i++ {
		rowPath := fmt.Sprintf("%s/rows/%d", path, i)
		row := &table.Rows[i]
		v.id(row.ID, rowPath+"/id")
		rowAnchor := v.anchor(&row.Anchor, rowPath+"/anchor", ownerPart, tableAnchor)
		v.optionalTwips(row.HeightTwips, rowPath+"/height_twips", 0)
		if row.HeightRule != nil {
			v.oneOf(*row.HeightRule, rowPath+"/height_rule", "atLeast", "exact")
		}
		if (row.HeightTwips == nil) != (row.HeightRule == nil) {
			v.add("INVALID_VALUE", rowPath+"/height_rule", "row height requires both height_twips and height_rule")
		}
		if row.RepeatHeader == nil {
			v.add("REQUIRED", rowPath+"/repeat_header", "field is required")
		}
		if row.Cells == nil {
			v.add("REQUIRED", rowPath+"/cells", "field is required")
		}
		for j := 0; j < v.collection(len(row.Cells), rowPath+"/cells", NativeDOCXMaxCollectionItems); j++ {
			cellPath := fmt.Sprintf("%s/cells/%d", rowPath, j)
			cell := &row.Cells[j]
			v.id(cell.ID, cellPath+"/id")
			cellAnchor := v.anchor(&cell.Anchor, cellPath+"/anchor", ownerPart, rowAnchor)
			v.optionalTwips(cell.WidthTwips, cellPath+"/width_twips", 0)
			v.positiveInt(cell.GridSpan, cellPath+"/grid_span")
			v.oneOf(cell.VerticalMerge, cellPath+"/vertical_merge", "none", "restart", "continue")
			if cell.Paragraphs == nil {
				v.add("REQUIRED", cellPath+"/paragraphs", "field is required")
			}
			for k := 0; k < v.collection(len(cell.Paragraphs), cellPath+"/paragraphs", NativeDOCXMaxCollectionItems); k++ {
				v.paragraph(&cell.Paragraphs[k], fmt.Sprintf("%s/paragraphs/%d", cellPath, k), true, ownerPart, cellAnchor)
			}
		}
	}
}

func (v *nativeValidator) run(run *NativeRunV1, path, ownerPart string, parentAnchor *nativeAnchorBounds) {
	if run.PageField != "" {
		v.oneOf(run.PageField, path+"/page_field", "PAGE", "NUMPAGES")
		if run.Kind != "text" {
			v.add("INVALID_UNION", path+"/page_field", "page field requires a text run")
		}
		if run.Text == nil || *run.Text != "" {
			v.add("INVALID_VALUE", path+"/text", "page-field source text must be empty; cached text is not authoritative")
		}
	}
	v.id(run.ID, path+"/id")
	runAnchor := v.anchor(&run.Anchor, path+"/anchor", ownerPart, parentAnchor)
	payloads := 0
	if run.Text != nil {
		payloads++
	}
	if run.Control != "" {
		payloads++
	}
	if run.Reference != nil {
		payloads++
	}
	if run.Drawing != nil {
		payloads++
	}
	if payloads != 1 {
		v.add("INVALID_UNION", path, "run kind must match exactly one payload")
	}
	switch run.Kind {
	case "text":
		if run.Text == nil {
			v.add("REQUIRED", path+"/text", "field is required")
		} else if utf8.RuneCountInString(*run.Text) > NativeDOCXMaxTextLength {
			v.add("LIMIT_EXCEEDED", path+"/text", fmt.Sprintf("must contain at most %d characters", NativeDOCXMaxTextLength))
		}
	case "control":
		v.oneOf(run.Control, path+"/control", "tab", "line-break", "page-break", "column-break", "soft-hyphen")
	case "reference":
		if run.Reference == nil {
			v.add("REQUIRED", path+"/reference", "field is required")
		} else {
			v.oneOf(run.Reference.Kind, path+"/reference/kind", "footnote", "endnote", "comment", "comment-range-start", "comment-range-end")
			v.oneOf(run.Reference.Role, path+"/reference/role", "", "anchor", "label")
			v.requiredID(run.Reference.TargetID, path+"/reference/target_id")
			target := "comment"
			if run.Reference.Kind == "footnote" {
				target = "footnote-story"
			} else if run.Reference.Kind == "endnote" {
				target = "endnote-story"
			}
			v.ref(run.Reference.TargetID, path+"/reference/target_id", target)
		}
	case "drawing":
		if run.Drawing == nil {
			v.add("REQUIRED", path+"/drawing", "field is required")
		} else {
			v.drawing(run.Drawing, path+"/drawing", ownerPart, runAnchor)
		}
	default:
		v.add("INVALID_UNION", path+"/kind", "unknown run kind")
	}
	if run.Properties != nil {
		properties := run.Properties
		v.optionalID(properties.CharacterStyleID, path+"/properties/character_style_id")
		v.optionalString(properties.FontFamily, path+"/properties/font_family")
		if properties.FontSizeHalfPoint != nil {
			v.positiveInt(properties.FontSizeHalfPoint, path+"/properties/font_size_half_points")
		}
		if properties.Underline != nil {
			v.oneOf(*properties.Underline, path+"/properties/underline", "none", "single", "double", "words")
		}
		if properties.VerticalAlignment != nil {
			v.oneOf(*properties.VerticalAlignment, path+"/properties/vertical_alignment", "baseline", "subscript", "superscript")
		}
		if properties.Color != nil && !nativeColor.MatchString(*properties.Color) {
			v.add("INVALID_VALUE", path+"/properties/color", "must be auto or uppercase RRGGBB")
		}
		v.optionalString(properties.Highlight, path+"/properties/highlight")
		v.optionalString(properties.Language, path+"/properties/language")
	}
}

func (v *nativeValidator) drawing(drawing *NativeDrawingV1, path, ownerPart string, parentAnchor *nativeAnchorBounds) {
	v.id(drawing.ID, path+"/id")
	v.anchor(&drawing.Anchor, path+"/anchor", ownerPart, parentAnchor)
	v.optionalID(drawing.RelationshipID, path+"/relationship_id")
	v.optionalString(drawing.ContentType, path+"/content_type")
	v.optionalString(drawing.Name, path+"/name")
	v.optionalString(drawing.AltText, path+"/alt_text")
	v.oneOf(drawing.Placement, path+"/placement", "inline", "floating")
	v.positive(drawing.WidthEMU, path+"/width_emu")
	v.positive(drawing.HeightEMU, path+"/height_emu")
	if drawing.RotationDegrees != nil && *drawing.RotationDegrees != 0 && *drawing.RotationDegrees != 90 && *drawing.RotationDegrees != 180 && *drawing.RotationDegrees != 270 {
		v.add("INVALID_VALUE", path+"/rotation_degrees", "bounded inline transforms support only quarter turns")
	}
	if crop := drawing.SourceCrop; crop != nil {
		for _, field := range []struct {
			name  string
			value *int64
		}{{"left", crop.Left}, {"top", crop.Top}, {"right", crop.Right}, {"bottom", crop.Bottom}} {
			v.nonnegative(field.value, path+"/source_crop/"+field.name)
			if field.value != nil && *field.value > 99000 {
				v.add("OUT_OF_RANGE", path+"/source_crop/"+field.name, "crop must retain at least one percent per axis")
			}
		}
		if crop.Left != nil && crop.Right != nil && *crop.Left >= 0 && *crop.Left <= 99000 && *crop.Right >= 0 && *crop.Right <= 99000 && *crop.Left+*crop.Right > 99000 {
			v.add("OUT_OF_RANGE", path+"/source_crop", "horizontal crop must retain at least one percent")
		}
		if crop.Top != nil && crop.Bottom != nil && *crop.Top >= 0 && *crop.Top <= 99000 && *crop.Bottom >= 0 && *crop.Bottom <= 99000 && *crop.Top+*crop.Bottom > 99000 {
			v.add("OUT_OF_RANGE", path+"/source_crop", "vertical crop must retain at least one percent")
		}
	}
	v.optionalSafe(drawing.XEMU, path+"/x_emu")
	v.optionalSafe(drawing.YEMU, path+"/y_emu")
	if drawing.Wrap != nil {
		v.oneOf(*drawing.Wrap, path+"/wrap", "none", "square", "tight", "through", "top-and-bottom")
	}
	v.optionalString(drawing.HorizontalRelativeFrom, path+"/horizontal_relative_from")
	v.optionalString(drawing.VerticalRelativeFrom, path+"/vertical_relative_from")
	if drawing.Placement == "inline" && (drawing.XEMU != nil || drawing.YEMU != nil) {
		v.add("INVALID_VALUE", path, "inline drawings cannot carry floating offsets")
	}
	if drawing.MediaPart != nil {
		v.part(*drawing.MediaPart, path+"/media_part")
		v.ref(*drawing.MediaPart, path+"/media_part", "media-part")
	}
	v.editPolicy(&drawing.EditPolicy, path+"/edit_policy", nativeDrawingOperations)
}

func (v *nativeValidator) section(section *NativeSectionV1, path, mainPart string, bodyAnchor *nativeAnchorBounds) {
	v.id(section.ID, path+"/id")
	v.anchor(&section.Anchor, path+"/anchor", mainPart, bodyAnchor)
	v.requiredID(section.StartsAtBlockID, path+"/starts_at_block_id")
	v.ref(section.StartsAtBlockID, path+"/starts_at_block_id", "body-block")
	v.oneOf(section.BreakType, path+"/break_type", "continuous", "next-page", "even-page", "odd-page", "next-column")
	if section.TitlePage == nil {
		v.add("REQUIRED", path+"/title_page", "field is required")
	}
	v.twips(section.Page.WidthTwips, path+"/page/width_twips", 1)
	v.twips(section.Page.HeightTwips, path+"/page/height_twips", 1)
	v.oneOf(section.Page.Orientation, path+"/page/orientation", "portrait", "landscape")
	v.positiveInt(section.Page.Columns, path+"/page/columns")
	if nativeIntValue(section.Page.Columns) > 45 {
		v.add("OUT_OF_RANGE", path+"/page/columns", "must be from 1 through 45")
	}
	v.twips(section.Page.ColumnSpacingTwips, path+"/page/column_spacing_twips", 0)
	v.oneOf(section.Page.ColumnLayout, path+"/page/column_layout", "equal-width", "explicit")
	if len(section.Page.ColumnDefinitions) != nativeIntValue(section.Page.Columns) {
		v.add("INVALID_VALUE", path+"/page/column_definitions", "must contain exactly one identity per declared column")
	}
	if len(section.Page.ColumnDefinitions) > 45 {
		v.add("LIMIT_EXCEEDED", path+"/page/column_definitions", "must contain at most 45 columns")
	}
	for index := range section.Page.ColumnDefinitions[:nativeBoundedLength(len(section.Page.ColumnDefinitions), 45)] {
		column := &section.Page.ColumnDefinitions[index]
		columnPath := fmt.Sprintf("%s/page/column_definitions/%d", path, index)
		v.id(column.ID, columnPath+"/id")
		v.nonnegativeInt(column.Ordinal, columnPath+"/ordinal")
		if nativeIntValue(column.Ordinal) != index {
			v.add("INVALID_VALUE", columnPath+"/ordinal", "must equal the source-order column index")
		}
		if section.Page.ColumnLayout == "equal-width" {
			if column.WidthTwips != nil || column.SpaceAfterTwips != nil {
				v.add("INVALID_VALUE", columnPath, "equal-width columns derive width and spacing from section geometry")
			}
		} else if section.Page.ColumnLayout == "explicit" {
			v.twips(column.WidthTwips, columnPath+"/width_twips", 1)
			v.twips(column.SpaceAfterTwips, columnPath+"/space_after_twips", 0)
			if index == len(section.Page.ColumnDefinitions)-1 && nativeInt64Value(column.SpaceAfterTwips) != 0 {
				v.add("INVALID_VALUE", columnPath+"/space_after_twips", "the final explicit column cannot have trailing inter-column space")
			}
		}
	}
	margins := &section.Page.Margins
	v.twips(margins.TopTwips, path+"/page/margins/top_twips", 0)
	v.twips(margins.RightTwips, path+"/page/margins/right_twips", 0)
	v.twips(margins.BottomTwips, path+"/page/margins/bottom_twips", 0)
	v.twips(margins.LeftTwips, path+"/page/margins/left_twips", 0)
	v.twips(margins.HeaderTwips, path+"/page/margins/header_twips", 0)
	v.twips(margins.FooterTwips, path+"/page/margins/footer_twips", 0)
	v.twips(margins.GutterTwips, path+"/page/margins/gutter_twips", 0)
	v.headerFooterRefs(section.HeaderRefs, path+"/header_refs", "header-story")
	v.headerFooterRefs(section.FooterRefs, path+"/footer_refs", "footer-story")
}

func (v *nativeValidator) headerFooterRefs(refs []NativeHeaderFooterReferenceV1, path, target string) {
	if refs == nil {
		v.add("REQUIRED", path, "field is required")
		return
	}
	for i := 0; i < v.collection(len(refs), path, NativeDOCXMaxCollectionItems); i++ {
		refPath := fmt.Sprintf("%s/%d", path, i)
		v.oneOf(refs[i].Kind, refPath+"/kind", "default", "first", "even")
		v.requiredID(refs[i].StoryID, refPath+"/story_id")
		v.ref(refs[i].StoryID, refPath+"/story_id", target)
		v.requiredID(refs[i].RelationshipID, refPath+"/relationship_id")
	}
}

func (v *nativeValidator) editPolicy(policy *NativeEditPolicyV1, path string, supported map[string]bool) {
	v.oneOf(policy.Mode, path+"/mode", "read-write", "read-only")
	if policy.AllowedOperations == nil {
		v.add("REQUIRED", path+"/allowed_operations", "field is required")
	}
	seen := map[string]bool{}
	for i := 0; i < v.collection(len(policy.AllowedOperations), path+"/allowed_operations", len(nativeOperations)); i++ {
		operation := policy.AllowedOperations[i]
		opPath := fmt.Sprintf("%s/allowed_operations/%d", path, i)
		if !nativeOperations[operation] {
			v.add("INVALID_VALUE", opPath, "unsupported edit operation")
		}
		if nativeOperations[operation] && !supported[operation] {
			v.add("INVALID_VALUE", opPath, "operation is not valid for this object type")
		}
		if seen[operation] {
			v.add("INVALID_VALUE", opPath, "operation is duplicated")
		}
		seen[operation] = true
	}
	if policy.Mode == "read-only" {
		if policy.Refusal == nil {
			v.add("REQUIRED", path+"/refusal", "read-only content requires an explicit refusal")
		}
		if len(policy.AllowedOperations) > 0 {
			v.add("INVALID_VALUE", path+"/allowed_operations", "read-only content cannot advertise write operations")
		}
	} else if policy.Mode == "read-write" && policy.Refusal != nil {
		v.add("INVALID_VALUE", path+"/refusal", "read-write content cannot carry a refusal")
	}
	if policy.Refusal != nil {
		v.requiredID(policy.Refusal.Code, path+"/refusal/code")
		v.required(policy.Refusal.Message, path+"/refusal/message")
		v.oneOf(policy.Refusal.Preservation, path+"/refusal/preservation", "preserve-verbatim", "refuse-mutation")
	}
}

func (v *nativeValidator) anchor(anchor *NativeSourceAnchorV1, path, expectedPart string, parent *nativeAnchorBounds) *nativeAnchorBounds {
	if anchor == nil {
		v.add("REQUIRED", path, "field is required")
		return nil
	}
	v.part(anchor.PartName, path+"/part_name")
	v.required(anchor.Path, path+"/path")
	v.nonnegative(anchor.StartByte, path+"/start_byte")
	v.positive(anchor.EndByte, path+"/end_byte")
	if anchor.StartByte != nil && anchor.EndByte != nil && *anchor.EndByte <= *anchor.StartByte {
		v.add("OUT_OF_RANGE", path+"/end_byte", "must be greater than start_byte")
	}
	if expectedPart != "" && anchor.PartName != expectedPart {
		v.add("INVALID_VALUE", path+"/part_name", fmt.Sprintf("must equal owning part %q", expectedPart))
	}
	v.sha(anchor.XMLSHA256, path+"/xml_sha256")
	if anchor.StartByte == nil || anchor.EndByte == nil || *anchor.EndByte <= *anchor.StartByte || anchor.PartName == "" {
		return nil
	}
	bounds := &nativeAnchorBounds{part: anchor.PartName, start: *anchor.StartByte, end: *anchor.EndByte}
	if parent != nil && !bounds.within(parent) {
		v.add("OUT_OF_RANGE", path, "anchor must be contained by its parent anchor in the same part")
	}
	return bounds
}

func (v *nativeValidator) add(code, path, message string) {
	if len(v.issues) >= NativeDOCXMaxIssues {
		return
	}
	for _, issue := range v.issues {
		if issue.Code == code && issue.Path == path && issue.Message == message {
			return
		}
	}
	v.issues = append(v.issues, NativeValidationIssue{Code: code, Path: path, Message: message})
}
func (v *nativeValidator) required(value, path string) {
	if value == "" || utf8.RuneCountInString(value) > 4096 {
		v.add("INVALID_VALUE", path, "must contain 1 to 4096 characters")
	}
}
func (v *nativeValidator) optionalString(value *string, path string) {
	if value != nil {
		v.required(*value, path)
	}
}
func (v *nativeValidator) requiredID(value, path string) {
	if !nativeIDPattern.MatchString(value) {
		v.add("INVALID_VALUE", path, "contains an invalid identifier")
	}
}
func (v *nativeValidator) optionalID(value *string, path string) {
	if value != nil {
		v.requiredID(*value, path)
	}
}
func (v *nativeValidator) id(value, path string) {
	v.requiredID(value, path)
	if original, exists := v.ids[value]; exists {
		v.add("DUPLICATE_ID", path, "duplicate native id first used at "+original)
	} else if value != "" {
		v.ids[value] = path
	}
}
func (v *nativeValidator) ref(value, path, target string) {
	if value != "" {
		v.references = append(v.references, nativeReference{path: path, id: value, target: target})
	}
}
func (v *nativeValidator) targetExists(target, id string) bool {
	switch target {
	case "modeled-id":
		_, ok := v.ids[id]
		return ok
	case "body-block":
		return v.bodyBlocks[id]
	case "header-story":
		return v.headerStories[id]
	case "footer-story":
		return v.footerStories[id]
	case "footnote-story":
		return v.footnoteStories[id]
	case "endnote-story":
		return v.endnoteStories[id]
	case "comment-story":
		return v.commentStories[id]
	case "comment":
		return v.comments[id]
	case "media-part":
		return v.mediaParts[id]
	default:
		return false
	}
}
func (v *nativeValidator) sha(value, path string) {
	if !nativeSHA256.MatchString(value) {
		v.add("INVALID_VALUE", path, "must be lowercase sha256:<64 hex>")
	}
}
func (v *nativeValidator) part(value, path string) {
	if value == "" || utf8.RuneCountInString(value) > 4096 || strings.HasPrefix(value, "/") || strings.Contains(value, "\\") || strings.Contains(value, "//") {
		v.add("INVALID_VALUE", path, "contains an invalid OPC part name")
		return
	}
	for _, component := range strings.Split(value, "/") {
		if component == "." || component == ".." || !nativePartSegment.MatchString(component) {
			v.add("INVALID_VALUE", path, "contains a non-canonical OPC part name")
			return
		}
		decoded, err := url.PathUnescape(component)
		if err != nil || !utf8.ValidString(decoded) || decoded == "." || decoded == ".." || strings.HasSuffix(decoded, ".") || strings.ContainsAny(decoded, `\/?#%`) {
			v.add("INVALID_VALUE", path, "contains an unsafe percent-encoded OPC part segment")
			return
		}
		for _, character := range decoded {
			if character < 0x20 || character == 0x7f {
				v.add("INVALID_VALUE", path, "contains an encoded control character")
				return
			}
		}
	}
}
func (v *nativeValidator) oneOf(value, path string, allowed ...string) {
	for _, candidate := range allowed {
		if value == candidate {
			return
		}
	}
	v.add("INVALID_VALUE", path, "must be one of "+strings.Join(allowed, ", "))
}
func (v *nativeValidator) nonnegative(value *int64, path string) {
	if value == nil {
		v.add("REQUIRED", path, "field is required")
	} else if *value < 0 || *value > 9007199254740991 {
		v.add("OUT_OF_RANGE", path, "must be a nonnegative safe integer")
	}
}
func (v *nativeValidator) optionalNonnegative(value *int64, path string) {
	if value != nil && (*value < 0 || *value > 9007199254740991) {
		v.add("OUT_OF_RANGE", path, "must be a nonnegative safe integer")
	}
}
func (v *nativeValidator) optionalSafe(value *int64, path string) {
	if value != nil && (*value < -9007199254740991 || *value > 9007199254740991) {
		v.add("OUT_OF_RANGE", path, "must be a safe integer")
	}
}
func (v *nativeValidator) optionalTwips(value *int64, path string, minimum int64) {
	if value != nil && (*value < minimum || *value > nativeMaxTwipsForMilliPoints) {
		v.add("OUT_OF_RANGE", path, fmt.Sprintf("must be an integer from %d through %d before milli-point conversion", minimum, nativeMaxTwipsForMilliPoints))
	}
}
func (v *nativeValidator) twips(value *int64, path string, minimum int64) {
	if value == nil {
		v.add("REQUIRED", path, "field is required")
		return
	}
	v.optionalTwips(value, path, minimum)
}
func (v *nativeValidator) positive(value *int64, path string) {
	if value == nil {
		v.add("REQUIRED", path, "field is required")
	} else if *value < 1 || *value > 9007199254740991 {
		v.add("OUT_OF_RANGE", path, "must be a positive safe integer")
	}
}
func (v *nativeValidator) nonnegativeInt(value *int, path string) {
	if value == nil {
		v.add("REQUIRED", path, "field is required")
	} else if *value < 0 || int64(*value) > 9007199254740991 {
		v.add("OUT_OF_RANGE", path, "must be a nonnegative safe integer")
	}
}
func (v *nativeValidator) positiveInt(value *int, path string) {
	if value == nil {
		v.add("REQUIRED", path, "field is required")
	} else if *value < 1 || int64(*value) > 9007199254740991 {
		v.add("OUT_OF_RANGE", path, "must be a positive safe integer")
	}
}

func (v *nativeValidator) collection(length int, path string, maximum int) int {
	limit := length
	if limit > maximum {
		v.add("LIMIT_EXCEEDED", path, fmt.Sprintf("must contain at most %d items", maximum))
		limit = maximum
	}
	remaining := NativeDOCXMaxNodes - v.nodes
	if remaining <= 0 {
		v.add("LIMIT_EXCEEDED", path, fmt.Sprintf("document traversal exceeds %d values", NativeDOCXMaxNodes))
		return 0
	}
	if limit > remaining {
		v.add("LIMIT_EXCEEDED", path, fmt.Sprintf("document traversal exceeds %d values", NativeDOCXMaxNodes))
		limit = remaining
	}
	v.nodes += limit
	return limit
}
