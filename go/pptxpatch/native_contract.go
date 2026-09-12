package pptxpatch

const NativePPTXV1MediaType = "application/vnd.injoffice.pptx-native.v1+json"

// NativePPTXDeck is the stable JSON boundary between the authoritative Go
// OPC engine and render/edit consumers. It intentionally does not replace
// Deck yet: Deck remains the bounded build model while this contract gains
// anchored parse and patch support in later PRs.
//
// All geometry is integer EMU. Point-like values use OOXML-native integer
// units (hundredths of a point or EMU), so JSON never relies on float
// formatting. Optional values are omitted rather than encoded as null.
type NativePPTXDeck struct {
	ContractVersion string              `json:"contractVersion"`
	DocumentID      string              `json:"documentId"`
	Origin          NativeOrigin        `json:"origin"`
	SourceRevision  *string             `json:"sourceRevision,omitempty"`
	Size            NativeSize          `json:"size"`
	Assets          []NativeAsset       `json:"assets"`
	Slides          []NativeSlide       `json:"slides"`
	Compatibility   NativeCompatibility `json:"compatibility"`
}

type NativeSize struct {
	Cx *int64 `json:"cx"`
	Cy *int64 `json:"cy"`
}

type NativeTransform struct {
	X            *int64 `json:"x"`
	Y            *int64 `json:"y"`
	Cx           *int64 `json:"cx"`
	Cy           *int64 `json:"cy"`
	QuarterTurns *int64 `json:"quarterTurns,omitempty"`
}

// NativeSourceAnchor identifies the original OOXML object without exposing
// mutable raw XML to an untrusted browser. FingerprintSHA256 binds the anchor
// to the source revision; ObjectID is the stable native identity within the
// part (normally cNvPr id for a slide object).
type NativeSourceAnchor struct {
	PartName          string  `json:"partName"`
	ObjectID          string  `json:"objectId"`
	RelationshipID    *string `json:"relationshipId,omitempty"`
	FingerprintSHA256 string  `json:"fingerprintSha256"`
}

// NativePassthroughRef is an opaque, server-issued capability for an XML
// fragment, relationship closure, or source-asset read that must survive a
// surgical edit or be served by a trusted host. A host must resolve it only
// against the matching source revision, owner part, fingerprint, byte length,
// and issuance reason; client-supplied tokens are never package paths, raw XML,
// or permission inferred from Source metadata.
type NativePassthroughRef struct {
	Token             string                       `json:"token"`
	OwnerPart         string                       `json:"ownerPart"`
	FingerprintSHA256 string                       `json:"fingerprintSha256"`
	Disposition       NativePassthroughDisposition `json:"disposition"`
}

type NativeDiagnosticScope struct {
	SlideID   *string `json:"slideId,omitempty"`
	ElementID *string `json:"elementId,omitempty"`
	PartName  *string `json:"partName,omitempty"`
}

type NativeDiagnostic struct {
	Severity NativeDiagnosticSeverity `json:"severity"`
	Code     string                   `json:"code"`
	Message  string                   `json:"message"`
	Scope    *NativeDiagnosticScope   `json:"scope,omitempty"`
}

type NativeCompatibility struct {
	Status      NativeCompatibilityStatus `json:"status"`
	Diagnostics []NativeDiagnostic        `json:"diagnostics"`
}

// NativeAsset describes immutable bytes without granting access to them.
// Parsed source-only assets omit DataBase64 and require an asset passthrough
// capability; Source.PartName is an integrity anchor, not authorization for a
// gateway to read or serve that package part.
type NativeAsset struct {
	ID          string                 `json:"id"`
	Provenance  NativeProvenance       `json:"provenance"`
	ContentType string                 `json:"contentType"`
	SHA256      string                 `json:"sha256"`
	ByteLength  *int64                 `json:"byteLength"`
	DataBase64  *string                `json:"dataBase64,omitempty"`
	Source      *NativeSourceAnchor    `json:"source,omitempty"`
	Passthrough []NativePassthroughRef `json:"passthrough"`
}

type NativeTextRun struct {
	Text                *string `json:"text"`
	Bold                *bool   `json:"bold,omitempty"`
	Italic              *bool   `json:"italic,omitempty"`
	FontSizeHundredthPt *int64  `json:"fontSizeHundredthPt,omitempty"`
	Color               *string `json:"color,omitempty"`
	FontFamily          *string `json:"fontFamily,omitempty"`
	Language            *string `json:"language,omitempty"`
}

type NativeParagraph struct {
	Runs            []NativeTextRun  `json:"runs"`
	Align           *NativeTextAlign `json:"align,omitempty"`
	Level           *int64           `json:"level,omitempty"`
	Bullet          *bool            `json:"bullet,omitempty"`
	BulletCharacter *string          `json:"bulletCharacter,omitempty"`
	MarginLeftEmu   *int64           `json:"marginLeftEmu,omitempty"`
	IndentEmu       *int64           `json:"indentEmu,omitempty"`
}

// NativeTextBodyLayout is the exact v1 horizontal text-frame slice. Extraction
// materializes OOXML defaults. autoFit=shape-source-frame marks an explicit
// read-only approximation using the saved frame; it does not model resizing.
type NativeTextBodyLayout struct {
	LeftInsetEMU       *int64                   `json:"leftInsetEmu"`
	RightInsetEMU      *int64                   `json:"rightInsetEmu"`
	TopInsetEMU        *int64                   `json:"topInsetEmu"`
	BottomInsetEMU     *int64                   `json:"bottomInsetEmu"`
	Wrap               NativeTextWrap           `json:"wrap"`
	VerticalAnchor     NativeTextVerticalAnchor `json:"verticalAnchor"`
	AutoFit            string                   `json:"autoFit"`
	HorizontalOverflow string                   `json:"horizontalOverflow"`
	VerticalOverflow   string                   `json:"verticalOverflow"`
}

type NativeArrowEnd struct {
	Type string  `json:"type"`
	W    *string `json:"w,omitempty"`
	Len  *string `json:"len,omitempty"`
}

type NativeStroke struct {
	Color      string            `json:"color"`
	WidthEMU   *int64            `json:"widthEmu"`
	Cap        *NativeStrokeCap  `json:"cap,omitempty"`
	Join       *NativeStrokeJoin `json:"join,omitempty"`
	Dash       *NativeStrokeDash `json:"dash,omitempty"`
	MiterLimit *int64            `json:"miterLimit,omitempty"`
}

type NativeAnimation struct {
	Effect      NativeAnimationEffect `json:"effect"`
	Direction   *NativeDirection      `json:"direction,omitempty"`
	DelayMS     *int64                `json:"delayMs,omitempty"`
	DurationMS  *int64                `json:"durationMs,omitempty"`
	DistancePPM *int64                `json:"distancePpm,omitempty"`
}

type NativeTransition struct {
	Type      NativeTransitionType `json:"type"`
	Direction *NativeDirection     `json:"direction,omitempty"`
}

type NativeTableBorder struct {
	Color    string `json:"color"`
	WidthEMU *int64 `json:"widthEmu"`
}

type NativeTableCell struct {
	Text       *string               `json:"text"`
	Paragraphs *[]NativeParagraph    `json:"paragraphs,omitempty"`
	TextBody   *NativeTextBodyLayout `json:"textBody,omitempty"`
	Fill       *string               `json:"fill,omitempty"`
	Border     *NativeTableBorder    `json:"border,omitempty"`
	Align      *NativeTextAlign      `json:"align,omitempty"`
}

type NativeTable struct {
	ColumnWidths []int64             `json:"columnWidths"`
	RowHeights   []int64             `json:"rowHeights"`
	Rows         [][]NativeTableCell `json:"rows"`
}

type NativeOpaqueChart struct {
	ChartPart      string               `json:"chartPart"`
	RelationshipID string               `json:"relationshipId"`
	OpaqueRef      NativePassthroughRef `json:"opaqueRef"`
	PreviewAssetID *string              `json:"previewAssetId,omitempty"`
}

// NativeElement is the Go binding for the schema's discriminated union.
// Validate enforces which fields are required and permitted for each Kind;
// DecodeNativePPTXJSON additionally rejects every unknown JSON property.
// NativePictureCrop retains DrawingML source-edge insets (100000 = full image).
// Required pointers distinguish authored zero insets from missing JSON fields.
type NativePictureCrop struct {
	Left   *int64 `json:"left"`
	Top    *int64 `json:"top"`
	Right  *int64 `json:"right"`
	Bottom *int64 `json:"bottom"`
}

type NativeElement struct {
	Kind           NativeElementKind      `json:"kind"`
	ID             string                 `json:"id"`
	Provenance     NativeProvenance       `json:"provenance"`
	Name           *string                `json:"name,omitempty"`
	Transform      NativeTransform        `json:"transform"`
	Placeholder    *NativePlaceholderType `json:"placeholder,omitempty"`
	Paragraphs     *[]NativeParagraph     `json:"paragraphs,omitempty"`
	TextBody       *NativeTextBodyLayout  `json:"textBody,omitempty"`
	Preset         *NativeShapePreset     `json:"preset,omitempty"`
	Fill           *string                `json:"fill,omitempty"`
	Stroke         *NativeStroke          `json:"stroke,omitempty"`
	HeadArrow      *bool                  `json:"headArrow,omitempty"`
	HeadEnd        *NativeArrowEnd        `json:"headEnd,omitempty"`
	TailEnd        *NativeArrowEnd        `json:"tailEnd,omitempty"`
	TailArrow      *bool                  `json:"tailArrow,omitempty"`
	FlipH          *bool                  `json:"flipH,omitempty"`
	AssetID        *string                `json:"assetId,omitempty"`
	Crop           *NativePictureCrop     `json:"crop,omitempty"`
	Table          *NativeTable           `json:"table,omitempty"`
	Chart          *NativeOpaqueChart     `json:"chart,omitempty"`
	ChildTransform *NativeTransform       `json:"childTransform,omitempty"`
	Children       []NativeElement        `json:"children,omitempty"`
	Animation      *NativeAnimation       `json:"animation,omitempty"`
	Source         *NativeSourceAnchor    `json:"source,omitempty"`
	Passthrough    []NativePassthroughRef `json:"passthrough"`
	Compatibility  NativeCompatibility    `json:"compatibility"`
}

type NativeSlide struct {
	ID            string                 `json:"id"`
	Provenance    NativeProvenance       `json:"provenance"`
	Background    *string                `json:"background,omitempty"`
	Transition    *NativeTransition      `json:"transition,omitempty"`
	Elements      []NativeElement        `json:"elements"`
	Source        *NativeSourceAnchor    `json:"source,omitempty"`
	Passthrough   []NativePassthroughRef `json:"passthrough"`
	Compatibility NativeCompatibility    `json:"compatibility"`
}
