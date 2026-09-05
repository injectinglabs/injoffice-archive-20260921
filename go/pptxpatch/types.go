// Package pptxpatch is the pptx-native slide model + real OOXML read/write
// engine for the Slides OOXML pivot: same charter as go/xlsxpatch (build a
// real, valid file — not a template render — validated against an
// independent OOXML consumer, never trusted on self-consistency alone) but
// for p:sld presentation parts instead of xlsx worksheets.
//
// Scope for this foundation slice, deliberately narrow (see docs/ROADMAP.md
// "Slides OOXML pivot" for the fuller plan): title/body text placeholders, a
// handful of basic prstGeom shapes + a line connector, strict plain tables,
// and bounded PNG/JPEG pictures, one theme, one slideLayout/slideMaster
// shared by every slide. No charts, groups, animations, or multi-layout/
// master decks yet — those are
// explicit follow-on phases once this narrow slice is proven correct.
//
// DeckSpec (packages/slides) is NOT replaced by this — it stays the
// authoring/AI-generation contract and compiles INTO a Deck here, the same
// "spec as agent-facing contract, real file underneath" split as
// ChartSpec→xlsxpatch and PivotSpec→xlsxpatch. The DeckSpec→Deck compiler
// and the Deck→Konva canvas renderer are follow-up work; this package is
// the file-format core they both sit on top of.
package pptxpatch

import (
	"encoding/xml"
	"io"
	"path"
	"strings"
)

// ---- EMU (English Metric Unit) conversions — ECMA-376's native unit ----
// 1 inch = 914400 EMU; 1 point = 12700 EMU. Every position/size in this
// package is EMU, matching the file format directly (no px/pt conversion
// drift between write and read).
const (
	EMUPerInch = 914400
	EMUPerPt   = 12700
)

// Standard 16:9 slide size (PowerPoint's own default since 2013) — the only
// size this foundation slice writes; ParsePPTX reads whatever size a real
// file declares.
const (
	DefaultSlideCx = 12192000
	DefaultSlideCy = 6858000
)

// PlaceholderType marks a shape as a title/body placeholder rather than a
// free-floating shape — round-trips DeckSpec's title/body fields onto real
// OOXML placeholder semantics (ph type=, the same idiom PowerPoint's own
// "Insert Placeholder" uses) instead of losing that identity as a generic
// text box. "" = not a placeholder.
type PlaceholderType string

const (
	PlaceholderNone     PlaceholderType = ""
	PlaceholderTitle    PlaceholderType = "title"
	PlaceholderCtrTitle PlaceholderType = "ctrTitle"
	PlaceholderSubTitle PlaceholderType = "subTitle"
	PlaceholderBody     PlaceholderType = "body"
)

// ShapeKind: for a preset-geometry shape, the kind IS the OOXML prst value
// directly — the same "no translation table" convention shapewrite.go/
// shaperead.go already established for xlsx shapes. "textBox" and "line" are
// the two structural specials (a plain rect marked txBox="1", and a
// cxnSp connector) that aren't prst names themselves.
type ShapeKind string

const (
	KindTextBox ShapeKind = "textBox"
	KindLine    ShapeKind = "line"
	// KindChart is an opaque p:graphicFrame containing an OOXML chart. It is
	// deliberately not a semantic chart model: Chart carries the original
	// chart part and its immediate local relationship closure unchanged.
	KindChart      ShapeKind = "chart"
	KindRect       ShapeKind = "rect"
	KindRoundRect  ShapeKind = "roundRect"
	KindEllipse    ShapeKind = "ellipse"
	KindTriangle   ShapeKind = "triangle"
	KindDiamond    ShapeKind = "diamond"
	KindRightArrow ShapeKind = "rightArrow"
	KindPentagon   ShapeKind = "pentagon"
	KindHexagon    ShapeKind = "hexagon"
	KindStar5      ShapeKind = "star5"
	// KindImage is an embedded PNG/JPEG p:pic. ImageData is the exact media
	// part payload; BuildPPTX recreates a slide-local image relationship.
	KindImage ShapeKind = "image"
)

// presetGeomKinds are the prstGeom-backed shapes this writer/reader knows —
// narrow on purpose (see package doc); AddShape's xlsxpatch precedent is
// ~119 presets after several rounds, this is the deliberately small first
// slice. shapeWritable/shapeReadable both key off this one map so writer and
// reader can never silently drift apart on what "supported" means.
var presetGeomKinds = map[ShapeKind]bool{
	KindRect: true, KindRoundRect: true, KindEllipse: true, KindTriangle: true,
	KindDiamond: true, KindRightArrow: true, KindPentagon: true, KindHexagon: true,
	KindStar5: true,
}

func shapeKindSupported(k ShapeKind) bool {
	return k == KindTextBox || k == KindLine || k == KindImage || k == KindChart || presetGeomKinds[k]
}

// TextRun is one run of uniformly-styled text within a paragraph.
type TextRun struct {
	Text   string
	Bold   bool
	Italic bool
	// SizePt is the run's font size in points; 0 means "let PowerPoint use
	// its own default" (omits sz entirely rather than writing a fake value).
	SizePt float64
	// Color is #rrggbb, or "" for the theme/inherited default.
	Color string
	// Font is a typeface name (a:latin typeface), or "" to inherit the
	// theme/master's default (Calibri/Calibri Light). Applies to the latin
	// script run only — this slice doesn't carry ea/cs typeface overrides.
	Font string
}

// Paragraph is one paragraph of a text body: a run list plus paragraph-level
// formatting (alignment, bullet/indent level).
type Paragraph struct {
	Runs  []TextRun
	Align TextAlign
	// Level is the bullet/indent level, 0-based (OOXML's own lvl attribute).
	Level int
	// Bullet: emit a real bullet character (a:buChar) at Level, vs a:buNone
	// for a plain (non-bulleted) paragraph — title/body text is buNone,
	// list items are Bullet=true.
	Bullet bool
}

// TextAlign mirrors OOXML's a:pPr algn values directly.
type TextAlign string

const (
	AlignLeft   TextAlign = "l"
	AlignCenter TextAlign = "ctr"
	AlignRight  TextAlign = "r"
)

// Table is the conservative in-memory contract for a p:graphicFrame/a:tbl.
// It deliberately models only whole rectangular grids: no merge/span,
// banding, formulas, rich runs, effects, or inherited styles. Rows may carry
// their exact OOXML height; an empty RowHeights asks the writer to divide the
// containing Shape's height evenly for a newly-authored table.
type Table struct {
	Columns    []int // EMU widths, one per cell in every row
	RowHeights []int // optional EMU heights, one per row
	Rows       [][]TableCell
}

type TableCell struct {
	Text   string
	Fill   string // #rrggbb or empty
	Border TableBorder
	Align  TextAlign
}

type TableBorder struct {
	Color   string // #rrggbb or empty
	WidthPt float64
}

// Valid is intentionally strict so the future parser can fail closed rather
// than normalize an OOXML table it cannot preserve. Empty tables, zero-width
// columns, ragged rows, unsupported alignment values, and negative borders
// are rejected.
func (t Table) Valid() bool {
	if len(t.Columns) == 0 || len(t.Rows) == 0 {
		return false
	}
	if len(t.RowHeights) != 0 && len(t.RowHeights) != len(t.Rows) {
		return false
	}
	for _, width := range t.Columns {
		if width <= 0 {
			return false
		}
	}
	for rowIndex, row := range t.Rows {
		if len(t.RowHeights) != 0 && t.RowHeights[rowIndex] <= 0 {
			return false
		}
		if len(row) != len(t.Columns) {
			return false
		}
		for _, cell := range row {
			if !validTableColor(cell.Fill) || !validTableColor(cell.Border.Color) || cell.Border.WidthPt < 0 {
				return false
			}
			if (cell.Border.Color == "") != (cell.Border.WidthPt == 0) {
				return false
			}
			if cell.Align != "" && cell.Align != AlignLeft && cell.Align != AlignCenter && cell.Align != AlignRight {
				return false
			}
		}
	}
	return true
}

// ChartPart is opaque OOXML chart-part metadata for a future lossless
// graphicFrame/chart preservation path. It deliberately stores bytes and the
// relationship graph rather than interpreting chart series or styles.
type ChartPart struct {
	RelationshipID string
	PartName       string
	XML            []byte
	Relationships  []ChartRelationship
	EmbeddedParts  []ChartEmbeddedPart
}
type ChartRelationship struct{ ID, Type, Target string }
type ChartEmbeddedPart struct {
	Name        string
	Data        []byte
	ContentType string
}

// Valid admits only an opaque chart whose complete, immediate local
// relationship closure is present. Relationship targets retain their OPC
// spelling (for example "../embeddings/workbook.xlsx"), but resolvePart
// proves that they remain below ppt/ and point at exactly one supplied part.
// Anything external, escaping the package, incomplete, duplicate, or not a
// chartSpace XML document is refused rather than being silently damaged.
func (c ChartPart) Valid() bool {
	if c.RelationshipID == "" || !strings.HasPrefix(c.PartName, "ppt/charts/") || !safePPTPart(c.PartName) || len(c.XML) == 0 {
		return false
	}
	dec := xml.NewDecoder(strings.NewReader(string(c.XML)))
	var root xml.StartElement
	foundRoot := false
	for {
		tok, err := dec.Token()
		if err != nil {
			return false
		}
		if start, ok := tok.(xml.StartElement); ok {
			root, foundRoot = start, true
			break
		}
	}
	if !foundRoot || root.Name.Local != "chartSpace" {
		return false
	}
	for {
		_, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return false
		}
	}
	seen := map[string]bool{}
	targets := map[string]bool{}
	for _, r := range c.Relationships {
		resolved, ok := resolveChartPart(c.PartName, r.Target)
		if r.ID == "" || r.Type == "" || !ok || seen[r.ID] {
			return false
		}
		seen[r.ID] = true
		targets[resolved] = true
	}
	parts := map[string]bool{}
	for _, p := range c.EmbeddedParts {
		if !safeChartDependencyPart(p.Name) || len(p.Data) == 0 || p.ContentType == "" || parts[p.Name] || !targets[p.Name] {
			return false
		}
		parts[p.Name] = true
	}
	// A local relationship without a copied target is an incomplete graph;
	// an unreferenced copied part is equally suspect and is refused above.
	return len(parts) == len(targets)
}

func safePPTPart(name string) bool {
	return strings.HasPrefix(name, "ppt/") && !strings.Contains(name, `\\`) &&
		!strings.Contains(name, "//") && path.Clean(name) == name && name != "ppt"
}

// The opaque graph must not be able to overwrite presentation scaffolding on
// a subsequent BuildPPTX. These are the only locations chart parts normally
// own: chart/style/color XML, embedded workbooks, and chart-local images.
func safeChartDependencyPart(name string) bool {
	return safePPTPart(name) && (strings.HasPrefix(name, "ppt/charts/") ||
		strings.HasPrefix(name, "ppt/embeddings/") || strings.HasPrefix(name, "ppt/media/"))
}

// resolveChartPart resolves an OPC relationship target without allowing it
// to escape ppt/. Normal Office chart rels commonly use ../embeddings/; that
// is valid as long as its canonical result stays under ppt/.
func resolveChartPart(fromPart, target string) (string, bool) {
	if target == "" || strings.HasPrefix(target, "/") || strings.Contains(target, `\\`) ||
		strings.Contains(target, ":") || strings.Contains(target, "#") || strings.Contains(target, "?") {
		return "", false
	}
	resolved := path.Clean(path.Join(path.Dir(fromPart), target))
	if !safePPTPart(resolved) || resolved == fromPart {
		return "", false
	}
	return resolved, true
}

func validTableColor(color string) bool {
	if color == "" {
		return true
	}
	if len(color) != 7 || color[0] != '#' {
		return false
	}
	for _, r := range color[1:] {
		if !((r >= '0' && r <= '9') || (r >= 'a' && r <= 'f') || (r >= 'A' && r <= 'F')) {
			return false
		}
	}
	return true
}

// Shape is one object on a slide: a text placeholder, a free text box, a
// preset-geometry autoshape, or a line connector. Position/size are EMU,
// absolute within the slide (pptx has no per-shape anchor grid the way xlsx
// drawings do — a:xfrm off/ext IS the placement).
type Shape struct {
	Kind        ShapeKind
	Placeholder PlaceholderType
	// Name is the shape's display name (p:cNvPr@name) — cosmetic, but real
	// files always carry one, and it's useful for round-trip identification.
	Name   string
	X, Y   int // EMU
	Cx, Cy int // EMU
	// Fill/Stroke are #rrggbb, or "" for noFill/no outline. Text boxes and
	// placeholders default to no fill/no outline unless set explicitly.
	Fill          string
	Stroke        string
	StrokeWidthPt float64
	Paragraphs    []Paragraph

	// Anim is this shape's optional entrance animation (S11) — "" (the zero
	// value) means no animation at all, and BuildPPTX writes nothing extra
	// for this shape (byte-identical to a pre-S11 deck). See AnimEffect.
	Anim ShapeAnimation

	// HeadArrow/TailArrow: draw a triangle arrowhead at the connector's
	// first/second a:xfrm point respectively (off, and off+ext) — only
	// meaningful for KindLine; ignored (never written) for every other
	// kind, since an arrowhead on a shape's own border isn't a real OOXML
	// idiom this package needs. Added for S12's process-flow diagram slice
	// (boxes + directional arrows), see packages/slides' diagram.ts.
	HeadArrow bool
	TailArrow bool
	// FlipH mirrors a:xfrm's own flipH attribute — only meaningful for
	// KindLine. A straight-line connector always draws its OWN bounding
	// box's top-left-to-bottom-right diagonal (off -> off+ext) unless
	// flipped; when the two real endpoints you want to connect are instead
	// the box's OTHER diagonal (bottom-left-to-top-right — e.g. a parent
	// box connecting down-and-LEFT to a child that sits left of the
	// parent's center), FlipH=true draws that anti-diagonal instead. See
	// packages/slides' diagram.ts orgChart layout for the derivation.
	FlipH bool
	// ImageData/ImageContentType are meaningful only for KindImage. The
	// intentionally narrow reader accepts embedded image/png and image/jpeg
	// with no crop, effects, rotation, flips, or external relationship.
	ImageData        []byte
	ImageContentType string
	// Table is the conservative whole-table model above. A table Shape has no
	// Kind: its p:graphicFrame is structurally distinct from p:sp. Its position
	// and size are X/Y/Cx/Cy as for every other slide object.
	Table *Table
	// Chart is opaque preservation metadata only. No current parser/writer
	// reads or emits it; use Valid before a future chart-part round-trip.
	// KindChart is emitted as p:graphicFrame only after Chart.Valid proves its
	// complete local relationship closure.
	Chart *ChartPart
}

// AnimEffect names one of this package's supported p:timing entrance
// effects — deliberately narrow (see write.go's timing doc comment): the
// two effects real decks reach for most, not an attempt at PowerPoint's
// full animation-effect catalogue.
type AnimEffect string

const (
	AnimNone  AnimEffect = ""
	AnimFade  AnimEffect = "fade"
	AnimFlyIn AnimEffect = "flyIn"
)

// Direction is a side-of-slide direction, shared by AnimFlyIn ("which edge
// does the shape fly in FROM") and slide transitions push/wipe ("which edge
// does the incoming slide arrive FROM") — same four values OOXML's own
// ST_TransitionSideDirectionType uses (l/r/u/d), spelled out here for
// callers.
type Direction string

const (
	DirNone  Direction = ""
	DirLeft  Direction = "left"
	DirRight Direction = "right"
	DirUp    Direction = "up"
	DirDown  Direction = "down"
)

// ShapeAnimation is one shape's entrance animation. The zero value (Effect
// == AnimNone) means "no animation" — every other field is meaningless in
// that case and BuildPPTX ignores them.
type ShapeAnimation struct {
	Effect AnimEffect
	// Direction: AnimFlyIn only. "" defaults to DirDown (fly in from below
	// — the single most common PowerPoint entrance direction). Ignored for
	// AnimFade.
	Direction Direction
	// DelayMs/DurationMs: 0 means "use the default" (0ms delay — animations
	// play automatically when the slide loads, no click needed; 500ms
	// duration).
	DelayMs, DurationMs int
	// Distance: AnimFlyIn only, 0..1 as a fraction of slide width (left/
	// right) or height (up/down) — how far off-slide the shape starts. 0
	// means "use the default" (0.25).
	Distance float64
}

// TransitionType names one of this package's supported p:transition slide
// transitions. "" means no transition element is written at all (PowerPoint
// then falls back to its own default: an instant cut).
type TransitionType string

const (
	TransitionNone TransitionType = ""
	TransitionFade TransitionType = "fade"
	TransitionPush TransitionType = "push"
	TransitionWipe TransitionType = "wipe"
)

// SlideTransition is how a slide transitions IN when advancing to it from
// the previous slide. The zero value (Type == TransitionNone) writes no
// p:transition element at all.
type SlideTransition struct {
	Type TransitionType
	// Direction: Push/Wipe only, which edge the incoming slide arrives
	// from. "" defers to OOXML's own default ("left"). Ignored for Fade.
	Direction Direction
}

// Slide is one p:sld: its shape tree plus an optional solid background.
type Slide struct {
	Shapes []Shape
	// Background is #rrggbb for a solid slide background, or "" to inherit
	// the layout/master background (nothing written — the honest default).
	Background string
	// Transition is this slide's optional entrance transition (S11).
	Transition SlideTransition
}

// Deck is the whole presentation: ordered slides plus slide size. Cx/Cy of 0
// mean "use DefaultSlideCx/Cy" (BuildPPTX fills them in) — ParsePPTX always
// returns the real size found in the file, even if nonstandard.
type Deck struct {
	Slides []Slide
	Cx, Cy int
}
