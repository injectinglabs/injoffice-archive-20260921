package xlsxpatch

import (
	"bytes"
	"encoding/json"
	"sort"
	"strings"
	"testing"
)

func rectSpec() ShapeWriteSpec {
	return ShapeWriteSpec{
		SheetName:   "Data",
		Kind:        "rect",
		Text:        "Q3 target",
		Fill:        "#EEF4F2",
		Stroke:      "#0FA98F",
		StrokeWidth: 1.5,
		TextColor:   "#1D2427",
		FontSize:    14,
		Anchor:      ShapeAnchor{FromCol: 5, FromRow: 2, ToCol: 9, ToRow: 6},
	}
}

func TestAddShape_FreshWorkbook_RoundTripsThroughReader(t *testing.T) {
	orig := buildZip(t, fixtureWorkbook(false))
	out, err := AddShape(orig, rectSpec())
	if err != nil {
		t.Fatal(err)
	}

	shapes, err := ReadShapes(out)
	if err != nil {
		t.Fatal(err)
	}
	if len(shapes) != 1 {
		t.Fatalf("expected 1 shape, got %d: %+v", len(shapes), shapes)
	}
	s := shapes[0]
	if s.Identity != (ShapeIdentity{DrawingPart: "xl/drawings/drawing1.xml", ObjectID: 1}) {
		t.Errorf("identity: %+v", s.Identity)
	}
	if s.SheetName != "Data" || s.Kind != "rect" || s.Text != "Q3 target" {
		t.Errorf("round-trip mismatch: %+v", s)
	}
	if s.Fill != "#eef4f2" || s.Stroke != "#0fa98f" {
		t.Errorf("colors: fill=%q stroke=%q", s.Fill, s.Stroke)
	}
	if s.StrokeWidth < 1.4 || s.StrokeWidth > 1.6 {
		t.Errorf("stroke width: %v", s.StrokeWidth)
	}
	if s.TextColor != "#1d2427" || s.FontSize != 14 {
		t.Errorf("text style: color=%q size=%v", s.TextColor, s.FontSize)
	}
	if s.Anchor != (ShapeAnchor{FromCol: 5, FromRow: 2, ToCol: 9, ToRow: 6}) {
		t.Errorf("anchor: %+v", s.Anchor)
	}
}

func TestReadShapesJSONCarriesStableNativeIdentity(t *testing.T) {
	out, err := AddShape(buildZip(t, fixtureWorkbook(false)), rectSpec())
	if err != nil {
		t.Fatal(err)
	}
	shapes, err := ReadShapes(out)
	if err != nil {
		t.Fatal(err)
	}
	encoded, err := json.Marshal(shapes[0])
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Contains(encoded, []byte(`"identity":{"drawingPart":"xl/drawings/drawing1.xml","objectId":1}`)) {
		t.Fatalf("hydrated JSON lacks stable identity: %s", encoded)
	}
}

func TestAddShape_AllKinds_RoundTrip(t *testing.T) {
	kinds := []struct {
		kind string
		text string
	}{
		{"rect", "Box"},
		{"ellipse", "Circle"},
		{"rightArrow", ""},
		{"line", ""},
		{"text", "Just text"},
		{"wedgeRectCallout", "Note"},
	}
	orig := buildZip(t, fixtureWorkbook(false))
	for i, k := range kinds {
		spec := ShapeWriteSpec{
			SheetName: "Data", Kind: k.kind, Text: k.text,
			Anchor: ShapeAnchor{FromCol: i, FromRow: 0, ToCol: i + 1, ToRow: 1},
		}
		// Mirror @injoffice/shapes' SHAPE_DEFAULTS: "text" carries no
		// fill/stroke at all — that absence is exactly what lets the reader
		// tell a bare "text" apart from a stroked/filled "rect" (both write
		// prst="rect"). Every other kind needs a stroke to render.
		if k.kind != "text" {
			spec.Stroke = "#59636A"
			spec.StrokeWidth = 2
		}
		if k.kind == "rect" || k.kind == "ellipse" || k.kind == "wedgeRectCallout" {
			spec.Fill = "#FFF8E6"
		}
		out, err := AddShape(orig, spec)
		if err != nil {
			t.Fatalf("%s: %v", k.kind, err)
		}
		orig = out
	}
	shapes, err := ReadShapes(orig)
	if err != nil {
		t.Fatal(err)
	}
	if len(shapes) != len(kinds) {
		t.Fatalf("expected %d shapes, got %d: %+v", len(kinds), len(shapes), shapes)
	}
	for i, k := range kinds {
		if shapes[i].Kind != k.kind {
			t.Errorf("shape %d: kind = %q, want %q", i, shapes[i].Kind, k.kind)
		}
		if shapes[i].Text != k.text {
			t.Errorf("shape %d (%s): text = %q, want %q", i, k.kind, shapes[i].Text, k.text)
		}
	}
}

// The full curated catalogue (S6.2, ~80 presets): every shapePresets entry
// must write and read back as ITSELF — this is the sweep that actually
// proves "80 shapes" rather than the handful of hand-picked kinds above.
func TestAddShape_FullCatalogue_EveryPresetRoundTrips(t *testing.T) {
	orig := buildZip(t, fixtureWorkbook(false))
	var kinds []string
	for k := range shapePresets {
		kinds = append(kinds, k)
	}
	sort.Strings(kinds)
	for i, kind := range kinds {
		spec := ShapeWriteSpec{
			SheetName: "Data", Kind: kind, Text: "x",
			Stroke: "#59636A", StrokeWidth: 1.5,
			Anchor: ShapeAnchor{FromCol: i % 20, FromRow: i / 20, ToCol: i%20 + 1, ToRow: i/20 + 1},
		}
		if !arrowFillKinds[kind] {
			spec.Fill = "#EEF4F2"
		}
		out, err := AddShape(orig, spec)
		if err != nil {
			t.Fatalf("%s: write failed: %v", kind, err)
		}
		orig = out
	}
	shapes, err := ReadShapes(orig)
	if err != nil {
		t.Fatal(err)
	}
	if len(shapes) != len(kinds) {
		t.Fatalf("wrote %d presets, read back %d", len(kinds), len(shapes))
	}
	for i, kind := range kinds {
		if shapes[i].Kind != kind {
			t.Errorf("preset %q round-tripped as %q", kind, shapes[i].Kind)
		}
	}
	t.Logf("verified %d presets round-trip correctly", len(kinds))
}

func TestAddShape_ExtendsExistingDrawing_ChartAndShapeCoexist(t *testing.T) {
	orig := buildZip(t, fixtureWorkbook(false))
	withChart, err := AddChart(orig, writeSpec("column"))
	if err != nil {
		t.Fatal(err)
	}
	withBoth, err := AddShape(withChart, rectSpec())
	if err != nil {
		t.Fatal(err)
	}
	charts, err := ReadCharts(withBoth)
	if err != nil {
		t.Fatal(err)
	}
	if len(charts) != 1 {
		t.Fatalf("chart lost when shape was added: %d charts", len(charts))
	}
	shapes, err := ReadShapes(withBoth)
	if err != nil {
		t.Fatal(err)
	}
	if len(shapes) != 1 {
		t.Fatalf("expected 1 shape alongside the chart, got %d", len(shapes))
	}
}

func TestShapeAndChartInsertsAllocateGloballyUniqueDrawingIDs(t *testing.T) {
	orig := buildZip(t, fixtureWorkbook(false))
	withShape, err := AddShape(orig, rectSpec())
	if err != nil {
		t.Fatal(err)
	}
	withChart, err := AddChart(withShape, writeSpec("column"))
	if err != nil {
		t.Fatal(err)
	}
	withSecondShape, err := AddShape(withChart, ShapeWriteSpec{
		SheetName: "Data", Kind: "ellipse", Anchor: ShapeAnchor{FromCol: 10, FromRow: 2, ToCol: 13, ToRow: 6},
	})
	if err != nil {
		t.Fatal(err)
	}
	index, err := indexShapeDrawing([]byte(readEntry(t, withSecondShape, "xl/drawings/drawing1.xml")))
	if err != nil {
		t.Fatal(err)
	}
	if index.maxObjectID != 3 {
		t.Fatalf("max object id = %d, want 3", index.maxObjectID)
	}
	shapes, err := ReadShapes(withSecondShape)
	if err != nil {
		t.Fatal(err)
	}
	if len(shapes) != 2 || shapes[0].Identity.ObjectID != 1 || shapes[1].Identity.ObjectID != 3 {
		t.Fatalf("shape identities = %+v", shapes)
	}
}

func TestAddShape_RejectsUnknownKindAndDegenerateAnchor(t *testing.T) {
	orig := buildZip(t, fixtureWorkbook(false))
	if _, err := AddShape(orig, ShapeWriteSpec{SheetName: "Data", Kind: "notARealPreset", Anchor: ShapeAnchor{FromCol: 0, FromRow: 0, ToCol: 1, ToRow: 1}}); err == nil {
		t.Fatal("unknown kind should error")
	}
	if _, err := AddShape(orig, ShapeWriteSpec{SheetName: "Data", Kind: "rect", Anchor: ShapeAnchor{FromCol: 2, FromRow: 0, ToCol: 1, ToRow: 1}}); err == nil {
		t.Fatal("degenerate anchor should error")
	}
}

// Fidelity: adding a shape must not touch any part it doesn't need to, and
// MUST register the fresh drawing part's content type (a common OOXML bug —
// an unregistered part opens as a repair prompt in real Excel).
func TestAddShape_UntouchedPartsStayByteIdentical(t *testing.T) {
	orig := buildZip(t, fixtureWorkbook(false))
	out, err := AddShape(orig, rectSpec())
	if err != nil {
		t.Fatal(err)
	}
	// Never touched by a fresh-drawing AddShape: styles and doc properties.
	for _, part := range []string{"xl/styles.xml", "docProps/core.xml"} {
		orig := fixtureWorkbook(false)[part]
		if got := readEntry(t, out, part); got != orig {
			t.Errorf("%s unexpectedly changed", part)
		}
	}
	ct := readEntry(t, out, "[Content_Types].xml")
	if !strings.Contains(ct, "drawings/drawing1.xml") {
		t.Errorf("fresh drawing part not registered in [Content_Types].xml: %s", ct)
	}
}

// The old disambiguation (rect vs text by absence of fill/stroke) was
// fragile — this proves the real fix: a "text" shape with fill/stroke SET
// still reads back as "text", because it carries the actual OOXML marker
// (cNvSpPr txBox="1") rather than being inferred from styling.
func TestAddShape_TextWithFillStillReadsAsText(t *testing.T) {
	orig := buildZip(t, fixtureWorkbook(false))
	out, err := AddShape(orig, ShapeWriteSpec{
		SheetName: "Data", Kind: "text", Text: "Styled note",
		Fill: "#FFF8E6", Stroke: "#F2A13C", StrokeWidth: 1.5,
		Anchor: ShapeAnchor{FromCol: 0, FromRow: 0, ToCol: 2, ToRow: 2},
	})
	if err != nil {
		t.Fatal(err)
	}
	shapes, err := ReadShapes(out)
	if err != nil {
		t.Fatal(err)
	}
	if len(shapes) != 1 || shapes[0].Kind != "text" {
		t.Fatalf("expected 1 text shape, got %+v", shapes)
	}
	if shapes[0].Fill != "#fff8e6" || shapes[0].Stroke != "#f2a13c" {
		t.Errorf("styling should still round-trip: %+v", shapes[0])
	}
}
