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
	Text                        *string `json:"text"`
	Bold                        *bool   `json:"bold,omitempty"`
	Italic                      *bool   `json:"italic,omitempty"`
	FontSizeHundredthPt         *int64  `json:"fontSizeHundredthPt,omitempty"`
	KerningThresholdHundredthPt *int64  `json:"kerningThresholdHundredthPt,omitempty"`
	Color                       *string `json:"color,omitempty"`
	FontFamily                  *string `json:"fontFamily,omitempty"`
	Language                    *string `json:"language,omitempty"`
}

type NativeParagraph struct {
	Runs               []NativeTextRun  `json:"runs"`
	Align              *NativeTextAlign `json:"align,omitempty"`
	Level              *int64           `json:"level,omitempty"`
	Bullet             *bool            `json:"bullet,omitempty"`
	BulletCharacter    *string          `json:"bulletCharacter,omitempty"`
	BulletFontFamily   *string          `json:"bulletFontFamily,omitempty"`
	BulletFontEncoding *string          `json:"bulletFontEncoding,omitempty"`
	MarginLeftEmu      *int64           `json:"marginLeftEmu,omitempty"`
	IndentEmu          *int64           `json:"indentEmu,omitempty"`
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
	WritingMode        *string                  `json:"writingMode,omitempty"`
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

type NativeLiteralPie struct {
	Profile         string   `json:"profile"`
	FirstSliceAngle int64    `json:"firstSliceAngle"`
	Values          []int64  `json:"values"`
	Colors          []string `json:"colors"`
}

type NativeLiteralDoughnut struct {
	Profile         string   `json:"profile"`
	FirstSliceAngle int64    `json:"firstSliceAngle"`
	HoleSize        int64    `json:"holeSize"`
	Values          []int64  `json:"values"`
	Colors          []string `json:"colors"`
}

// NativeChartAxisLabels retains complete local source text and tick policy.
// Layout margins and outward tick length are supplied preview policies.
type NativeChartAxisLabelStyle struct {
	FontFamily string `json:"fontFamily"`
	FontSize   int64  `json:"fontSize"`
	Color      string `json:"color"`
	Bold       bool   `json:"bold"`
	Italic     bool   `json:"italic"`
	Language   string `json:"language"`
}
type NativeChartAxisLabels struct {
	Profile       string                    `json:"profile"`
	Position      string                    `json:"position"`
	MajorTickMark string                    `json:"majorTickMark"`
	Style         NativeChartAxisLabelStyle `json:"style"`
	MajorUnit     *string                   `json:"majorUnit,omitempty"`
	NumberFormat  *string                   `json:"numberFormat,omitempty"`
}

type NativeLiteralBarAxis struct {
	Labels      *NativeChartAxisLabels `json:"labels,omitempty"`
	ID          int64                  `json:"id"`
	CrossAxisID int64                  `json:"crossAxisId"`
	Orientation string                 `json:"orientation"`
	Position    string                 `json:"position"`
	Deleted     bool                   `json:"deleted"`
	Color       *string                `json:"color,omitempty"`
	WidthEMU    *int64                 `json:"widthEmu,omitempty"`
	Min         *string                `json:"min,omitempty"`
	Max         *string                `json:"max,omitempty"`
	CrossesAt   *string                `json:"crossesAt,omitempty"`
}
type NativeLiteralBarSeries struct {
	Index  int64    `json:"index"`
	Order  int64    `json:"order"`
	Title  *string  `json:"title,omitempty"`
	Values []string `json:"values"`
	Colors []string `json:"colors"`
}
type NativeLiteralBar struct {
	Profile      string                   `json:"profile"`
	BarDirection string                   `json:"barDirection"`
	Grouping     string                   `json:"grouping"`
	DataOrigin   string                   `json:"dataOrigin"`
	GapWidth     int64                    `json:"gapWidth"`
	Overlap      int64                    `json:"overlap"`
	Categories   []string                 `json:"categories"`
	Series       []NativeLiteralBarSeries `json:"series"`
	CategoryAxis NativeLiteralBarAxis     `json:"categoryAxis"`
	ValueAxis    NativeLiteralBarAxis     `json:"valueAxis"`
}
type NativeLiteralConnectedSeries struct {
	Index    int64    `json:"index"`
	Order    int64    `json:"order"`
	Title    *string  `json:"title,omitempty"`
	Values   []string `json:"values"`
	XValues  []string `json:"xValues,omitempty"`
	Color    string   `json:"color"`
	WidthEMU int64    `json:"widthEmu"`
}
type NativeLiteralConnected struct {
	Profile    string                         `json:"profile"`
	DataOrigin string                         `json:"dataOrigin"`
	Categories []string                       `json:"categories"`
	Series     []NativeLiteralConnectedSeries `json:"series"`
	XAxis      NativeLiteralBarAxis           `json:"xAxis"`
	YAxis      NativeLiteralBarAxis           `json:"yAxis"`
}
type NativeOpaqueChart struct {
	LiteralConnected *NativeLiteralConnected `json:"literalConnected,omitempty"`
	LiteralBar       *NativeLiteralBar       `json:"literalBar,omitempty"`
	LiteralDoughnut  *NativeLiteralDoughnut  `json:"literalDoughnut,omitempty"`
	LiteralPie       *NativeLiteralPie       `json:"literalPie,omitempty"`
	ChartPart        string                  `json:"chartPart"`
	RelationshipID   string                  `json:"relationshipId"`
	OpaqueRef        NativePassthroughRef    `json:"opaqueRef"`
	PreviewAssetID   *string                 `json:"previewAssetId,omitempty"`
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
	Geometry       *NativeEvaluatedGeometry `json:"geometry,omitempty"`
	Kind           NativeElementKind        `json:"kind"`
	ID             string                   `json:"id"`
	Provenance     NativeProvenance         `json:"provenance"`
	Name           *string                  `json:"name,omitempty"`
	Transform      NativeTransform          `json:"transform"`
	Placeholder    *NativePlaceholderType   `json:"placeholder,omitempty"`
	Paragraphs     *[]NativeParagraph       `json:"paragraphs,omitempty"`
	TextBody       *NativeTextBodyLayout    `json:"textBody,omitempty"`
	Preset         *NativeShapePreset       `json:"preset,omitempty"`
	Fill           *string                  `json:"fill,omitempty"`
	Stroke         *NativeStroke            `json:"stroke,omitempty"`
	HeadArrow      *bool                    `json:"headArrow,omitempty"`
	HeadEnd        *NativeArrowEnd          `json:"headEnd,omitempty"`
	TailEnd        *NativeArrowEnd          `json:"tailEnd,omitempty"`
	TailArrow      *bool                    `json:"tailArrow,omitempty"`
	FlipH          *bool                    `json:"flipH,omitempty"`
	AssetID        *string                  `json:"assetId,omitempty"`
	Crop           *NativePictureCrop       `json:"crop,omitempty"`
	Clip           *string                  `json:"clip,omitempty"`
	Table          *NativeTable             `json:"table,omitempty"`
	Chart          *NativeOpaqueChart       `json:"chart,omitempty"`
	ChildTransform *NativeTransform         `json:"childTransform,omitempty"`
	Children       []NativeElement          `json:"children,omitempty"`
	Animation      *NativeAnimation         `json:"animation,omitempty"`
	Source         *NativeSourceAnchor      `json:"source,omitempty"`
	Passthrough    []NativePassthroughRef   `json:"passthrough"`
	Compatibility  NativeCompatibility      `json:"compatibility"`
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
