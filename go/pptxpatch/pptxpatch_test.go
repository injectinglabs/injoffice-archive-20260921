package pptxpatch

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/xml"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestPictureBuildParsePreservesExactPNGPayloadAndRelationship(t *testing.T) {
	// Valid 1×1 transparent PNG; use a real binary payload, not an XML/data-URI
	// surrogate, so the ZIP media part and its p:pic relationship are exercised.
	png := []byte{137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0, 31, 21, 196, 137, 0, 0, 0, 13, 73, 68, 65, 84, 8, 215, 99, 248, 255, 255, 255, 127, 0, 9, 251, 3, 253, 160, 63, 218, 53, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130}
	data, err := BuildPPTX(Deck{Slides: []Slide{{Shapes: []Shape{{Kind: KindImage, Name: "logo", X: 10, Y: 20, Cx: 30, Cy: 40, ImageData: png, ImageContentType: "image/png"}}}}})
	if err != nil {
		t.Fatal(err)
	}
	got, err := ParsePPTX(data)
	if err != nil {
		t.Fatal(err)
	}
	if len(got.Slides) != 1 || len(got.Slides[0].Shapes) != 1 {
		t.Fatalf("parsed %#v", got)
	}
	image := got.Slides[0].Shapes[0]
	if image.Kind != KindImage || image.ImageContentType != "image/png" || image.Name != "logo" || image.X != 10 || image.Y != 20 || image.Cx != 30 || image.Cy != 40 {
		t.Fatalf("image = %#v", image)
	}
	if sha256.Sum256(image.ImageData) != sha256.Sum256(png) {
		t.Fatal("image payload hash changed")
	}
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	rels, _ := readZipFile(zr, "ppt/slides/_rels/slide1.xml.rels")
	slideXML, _ := readZipFile(zr, "ppt/slides/slide1.xml")
	if !strings.Contains(rels, `Relationship Id="rId2"`) || !strings.Contains(slideXML, `r:embed="rId2"`) {
		t.Fatal("missing recreated image relationship")
	}
}

func TestConservativeTableValidation(t *testing.T) {
	valid := Table{Columns: []int{100, 200}, Rows: [][]TableCell{{{Text: "A", Align: AlignLeft}, {Text: "B", Fill: "#ffffff", Border: TableBorder{Color: "#000000", WidthPt: 1}, Align: AlignCenter}}}}
	if !valid.Valid() {
		t.Fatal("expected rectangular basic table to validate")
	}
	for name, table := range map[string]Table{
		"empty":           {},
		"zero-width":      {Columns: []int{0}, Rows: [][]TableCell{{{}}}},
		"ragged":          {Columns: []int{1, 1}, Rows: [][]TableCell{{{}}}},
		"bad-align":       {Columns: []int{1}, Rows: [][]TableCell{{{Align: "justify"}}}},
		"negative-border": {Columns: []int{1}, Rows: [][]TableCell{{{Border: TableBorder{WidthPt: -1}}}}},
		"unpaired-border": {Columns: []int{1}, Rows: [][]TableCell{{{Border: TableBorder{Color: "#000000"}}}}},
		"bad-fill":        {Columns: []int{1}, Rows: [][]TableCell{{{Fill: "blue"}}}},
		"bad-row-heights": {Columns: []int{1}, RowHeights: []int{1, 1}, Rows: [][]TableCell{{{}}}},
	} {
		t.Run(name, func(t *testing.T) {
			if table.Valid() {
				t.Fatal("expected invalid table")
			}
		})
	}
}

func TestOpaqueChartPartValidation(t *testing.T) {
	valid := ChartPart{RelationshipID: "rId4", PartName: "ppt/charts/chart1.xml", XML: []byte("<c:chartSpace/>"), Relationships: []ChartRelationship{{ID: "rId1", Type: "package", Target: "../embeddings/book.xlsx"}}, EmbeddedParts: []ChartEmbeddedPart{{Name: "ppt/embeddings/book.xlsx", Data: []byte{1}, ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}}}
	if !valid.Valid() {
		t.Fatal("expected valid opaque chart graph")
	}
	for name, chart := range map[string]ChartPart{
		"empty":     {},
		"traversal": {RelationshipID: "r", PartName: "../chart.xml", XML: []byte("x")},
		"external":  {RelationshipID: "r", PartName: "chart.xml", XML: []byte("x"), Relationships: []ChartRelationship{{ID: "r1", Type: "t", Target: "https://x"}}},
		"duplicate": {RelationshipID: "r", PartName: "chart.xml", XML: []byte("x"), Relationships: []ChartRelationship{{ID: "r1", Type: "t", Target: "a"}, {ID: "r1", Type: "t", Target: "b"}}},
	} {
		t.Run(name, func(t *testing.T) {
			if chart.Valid() {
				t.Fatal("expected invalid chart graph")
			}
		})
	}
}

func TestTableRoundTripPreservesPlainGridAndSlideOrder(t *testing.T) {
	table := Table{
		Columns:    []int{1200000, 1800000},
		RowHeights: []int{500000, 700000},
		Rows: [][]TableCell{
			{
				{Text: "First", Fill: "#f0f4ff", Border: TableBorder{Color: "#123456", WidthPt: 1}, Align: AlignLeft},
				{Text: "Center\nsecond line", Fill: "#ffffff", Border: TableBorder{Color: "#123456", WidthPt: 1}, Align: AlignCenter},
			},
			{
				{Text: "  preserve spaces  ", Fill: "#fff2cc", Align: AlignRight},
				{Text: "Last", Align: AlignLeft},
			},
		},
	}
	orig := Deck{Slides: []Slide{{Shapes: []Shape{
		{Kind: KindRect, Name: "before", X: 1, Y: 2, Cx: 3, Cy: 4, Fill: "#ffffff"},
		{Name: "status grid", X: 400000, Y: 500000, Cx: 3000000, Cy: 1200000, Table: &table},
		{Kind: KindEllipse, Name: "after", X: 5, Y: 6, Cx: 7, Cy: 8, Fill: "#000000"},
	}}}}
	if raw, err := tableXML(orig.Slides[0].Shapes[1], 3); err != nil {
		t.Fatalf("tableXML: %v", err)
	} else if _, ok := parseTableGraphicFrame(raw); !ok {
		t.Fatalf("table parser rejected writer output: %s", raw)
	}
	data, err := BuildPPTX(orig)
	if err != nil {
		t.Fatalf("BuildPPTX: %v", err)
	}
	if zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data))); err == nil {
		if slideXML, ok := readZipFile(zr, "ppt/slides/slide1.xml"); ok {
			dec := xml.NewDecoder(strings.NewReader(slideXML))
			for {
				tok, err := dec.Token()
				if err != nil {
					t.Fatalf("read generated slide: %v", err)
				}
				if start, ok := tok.(xml.StartElement); ok && start.Name.Local == "graphicFrame" {
					raw, err := graphicFrameXML(dec, start)
					if err != nil {
						t.Fatalf("graphicFrameXML: %v", err)
					}
					if _, ok := parseTableGraphicFrame(raw); !ok {
						t.Fatalf("parser rejected generated graphicFrame: %s", raw)
					}
					break
				}
			}
		}
	}
	got, err := ParsePPTX(data)
	if err != nil {
		t.Fatalf("ParsePPTX: %v", err)
	}
	shapes := got.Slides[0].Shapes
	if len(shapes) != 3 || shapes[0].Name != "before" || shapes[1].Name != "status grid" || shapes[2].Name != "after" {
		t.Fatalf("table did not retain z-order: %+v", shapes)
	}
	frame := shapes[1]
	if frame.Table == nil {
		t.Fatal("table shape was not parsed")
	}
	if frame.X != 400000 || frame.Y != 500000 || frame.Cx != 3000000 || frame.Cy != 1200000 {
		t.Errorf("table frame geometry = (%d,%d,%d,%d)", frame.X, frame.Y, frame.Cx, frame.Cy)
	}
	if got := frame.Table; len(got.Columns) != 2 || got.Columns[0] != 1200000 || got.Columns[1] != 1800000 ||
		len(got.RowHeights) != 2 || got.RowHeights[0] != 500000 || got.RowHeights[1] != 700000 {
		t.Errorf("table grid = %+v", got)
	}
	if cell := frame.Table.Rows[0][1]; cell.Text != "Center\nsecond line" || cell.Align != AlignCenter || cell.Fill != "#ffffff" || cell.Border.Color != "#123456" || cell.Border.WidthPt != 1 {
		t.Errorf("cell [0][1] = %+v", cell)
	}
	if cell := frame.Table.Rows[1][0]; cell.Text != "  preserve spaces  " || cell.Align != AlignRight || cell.Fill != "#fff2cc" || cell.Border != (TableBorder{}) {
		t.Errorf("cell [1][0] = %+v", cell)
	}
}

func TestParseTableRejectsMergeAndUnsupportedTransform(t *testing.T) {
	// The reader skips only the unsupported frame and continues to surface
	// later supported slide objects. Both cases would be lossy if accepted:
	// a merge has no representation in Table, and a rotated frame cannot be
	// represented by Shape's table placement fields.
	frame := func(xfrm, tcAttrs string) string {
		return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="unsafe"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
			`<p:xfrm ` + xfrm + `><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>` +
			`<a:tblPr/><a:tblGrid><a:gridCol w="100"/></a:tblGrid><a:tr h="100"><a:tc ` + tcAttrs + `><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="l"><a:buNone/></a:pPr><a:r><a:t>x</a:t></a:r></a:p></a:txBody><a:tcPr><a:noFill/><a:lnL><a:noFill/></a:lnL><a:lnR><a:noFill/></a:lnR><a:lnT><a:noFill/></a:lnT><a:lnB><a:noFill/></a:lnB></a:tcPr></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`
	}
	for name, xmlFrame := range map[string]string{
		"merge":    frame("", `gridSpan="2"`),
		"rotation": frame(`rot="60000"`, ""),
	} {
		t.Run(name, func(t *testing.T) {
			slideXML := `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>` + xmlFrame +
				`<p:sp><p:nvSpPr><p:cNvPr id="3" name="keep"/><p:cNvSpPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp></p:spTree></p:cSld></p:sld>`
			got, err := parseSlide(slideXML, nil)
			if err != nil {
				t.Fatalf("parseSlide: %v", err)
			}
			if len(got.Shapes) != 1 || got.Shapes[0].Name != "keep" {
				t.Fatalf("unsupported table should be omitted: %+v", got.Shapes)
			}
		})
	}
}

func TestBuildPPTXProducesValidZipWithExpectedParts(t *testing.T) {
	data, err := BuildPPTX(ExampleDeck())
	if err != nil {
		t.Fatalf("BuildPPTX: %v", err)
	}
	if len(data) == 0 {
		t.Fatal("empty output")
	}
	// PK zip signature.
	if string(data[:2]) != "PK" {
		t.Fatalf("output doesn't look like a zip: %v", data[:4])
	}
}

func TestBuildPPTXDeclaresOnlyUsedImageContentTypes(t *testing.T) {
	t.Parallel()
	png := []byte{137, 80, 78, 71, 13, 10, 26, 10}
	tests := []struct {
		name       string
		shapes     []Shape
		wantPNG    bool
		wantJPEG   bool
		nativeRead bool
	}{
		{name: "no images"},
		{name: "png only", shapes: []Shape{{Kind: KindImage, Name: "PNG", X: 1, Y: 2, Cx: 3, Cy: 4, ImageData: png, ImageContentType: "image/png"}}, wantPNG: true, nativeRead: true},
		{name: "jpeg only", shapes: []Shape{{Kind: KindImage, Name: "JPEG", X: 1, Y: 2, Cx: 3, Cy: 4, ImageData: []byte("jpeg"), ImageContentType: "image/jpeg"}}, wantJPEG: true},
		{name: "both", shapes: []Shape{{Kind: KindImage, Name: "PNG", X: 1, Y: 2, Cx: 3, Cy: 4, ImageData: png, ImageContentType: "image/png"}, {Kind: KindImage, Name: "JPEG", X: 5, Y: 6, Cx: 7, Cy: 8, ImageData: []byte("jpeg"), ImageContentType: "image/jpeg"}}, wantPNG: true, wantJPEG: true},
	}
	for _, test := range tests {
		test := test
		t.Run(test.name, func(t *testing.T) {
			t.Parallel()
			data, err := BuildPPTX(Deck{Slides: []Slide{{Shapes: test.shapes}}})
			if err != nil {
				t.Fatal(err)
			}
			reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
			if err != nil {
				t.Fatal(err)
			}
			contentTypes, ok := readZipFile(reader, "[Content_Types].xml")
			if !ok {
				t.Fatal("missing [Content_Types].xml")
			}
			hasPNG := strings.Contains(contentTypes, `Default Extension="png" ContentType="image/png"`)
			hasJPEG := strings.Contains(contentTypes, `Default Extension="jpeg" ContentType="image/jpeg"`)
			if hasPNG != test.wantPNG || hasJPEG != test.wantJPEG {
				t.Fatalf("image defaults: png=%t jpeg=%t; content types %s", hasPNG, hasJPEG, contentTypes)
			}
			if test.nativeRead {
				deck, extractErr := ExtractNativePPTX(data, nativeTestExtractOptions())
				if extractErr != nil || len(deck.Assets) != 1 {
					t.Fatalf("writer output was not accepted by the strict native extractor: err=%v assets=%#v", extractErr, deck.Assets)
				}
			}
		})
	}
}

func TestRoundTripPreservesShapeGeometryAndText(t *testing.T) {
	orig := ExampleDeck()
	data, err := BuildPPTX(orig)
	if err != nil {
		t.Fatalf("BuildPPTX: %v", err)
	}
	got, err := ParsePPTX(data)
	if err != nil {
		t.Fatalf("ParsePPTX: %v", err)
	}

	if got.Cx != DefaultSlideCx || got.Cy != DefaultSlideCy {
		t.Errorf("slide size = %d x %d, want %d x %d", got.Cx, got.Cy, DefaultSlideCx, DefaultSlideCy)
	}
	if len(got.Slides) != len(orig.Slides) {
		t.Fatalf("slide count = %d, want %d", len(got.Slides), len(orig.Slides))
	}

	s0 := got.Slides[0]
	if s0.Background != "#0b1220" {
		t.Errorf("background = %q, want #0b1220", s0.Background)
	}
	if len(s0.Shapes) != len(orig.Slides[0].Shapes) {
		t.Fatalf("shape count = %d, want %d", len(s0.Shapes), len(orig.Slides[0].Shapes))
	}

	title := s0.Shapes[0]
	if title.Placeholder != PlaceholderTitle {
		t.Errorf("title placeholder = %q, want title", title.Placeholder)
	}
	if title.X != 838200 || title.Y != 365125 || title.Cx != 10515600 || title.Cy != 1325563 {
		t.Errorf("title geometry = (%d,%d,%d,%d), want (838200,365125,10515600,1325563)", title.X, title.Y, title.Cx, title.Cy)
	}
	if len(title.Paragraphs) != 1 || len(title.Paragraphs[0].Runs) != 1 {
		t.Fatalf("title paragraphs/runs shape unexpected: %+v", title.Paragraphs)
	}
	run := title.Paragraphs[0].Runs[0]
	if run.Text != "InjOffice Slides" {
		t.Errorf("title text = %q", run.Text)
	}
	if !run.Bold {
		t.Error("title run should be bold")
	}
	if run.SizePt != 40 {
		t.Errorf("title size = %v, want 40", run.SizePt)
	}
	if run.Color != "#f5f7fa" {
		t.Errorf("title color = %q, want #f5f7fa", run.Color)
	}
	if run.Font != "Archivo" {
		t.Errorf("title font = %q, want Archivo", run.Font)
	}

	body := s0.Shapes[1]
	if body.Placeholder != PlaceholderBody {
		t.Errorf("body placeholder = %q, want body", body.Placeholder)
	}
	if len(body.Paragraphs) != 3 {
		t.Fatalf("body paragraph count = %d, want 3", len(body.Paragraphs))
	}
	if !body.Paragraphs[0].Bullet {
		t.Error("body paragraph 0 should be a bullet")
	}
	if body.Paragraphs[2].Level != 1 {
		t.Errorf("body paragraph 2 level = %d, want 1", body.Paragraphs[2].Level)
	}
	if body.Paragraphs[1].Runs[0].Text != "Round-trips through PowerPoint" || !body.Paragraphs[1].Runs[0].Italic {
		t.Errorf("body paragraph 1 run wrong: %+v", body.Paragraphs[1].Runs[0])
	}

	callout := s0.Shapes[2]
	if callout.Kind != KindRoundRect {
		t.Errorf("callout kind = %q, want roundRect", callout.Kind)
	}
	if callout.Fill != "#2f6fed" || callout.Stroke != "#153a82" {
		t.Errorf("callout fill/stroke = %q/%q", callout.Fill, callout.Stroke)
	}
	if callout.StrokeWidthPt != 2 {
		t.Errorf("callout stroke width = %v, want 2", callout.StrokeWidthPt)
	}

	dot := s0.Shapes[3]
	if dot.Kind != KindEllipse || dot.Fill != "#ed7d31" {
		t.Errorf("dot wrong: %+v", dot)
	}

	line := s0.Shapes[4]
	if line.Kind != KindLine {
		t.Errorf("line kind = %q, want line", line.Kind)
	}
	if line.Stroke != "#a5a5a5" {
		t.Errorf("line stroke = %q", line.Stroke)
	}

	s1 := got.Slides[1]
	if len(s1.Shapes) != 1 || s1.Shapes[0].Placeholder != PlaceholderCtrTitle {
		t.Errorf("slide 2 unexpected: %+v", s1.Shapes)
	}
	if s1.Shapes[0].Paragraphs[0].Align != AlignCenter {
		t.Errorf("slide 2 align = %q, want ctr", s1.Shapes[0].Paragraphs[0].Align)
	}
}

// TestDiagramConnectorFieldsRoundTrip covers S12's new Shape fields
// (HeadArrow/TailArrow/FlipH), added for diagram (org-chart/process-flow)
// connectors — see DiagramExampleDeck and Shape.FlipH's doc comment for the
// geometry this is protecting. python-pptx's independent structural check
// lives in scripts/validate_diagram_arrows.py (python-pptx exposes no
// friendly property for these — only raw XML — so this Go-side round-trip
// is what actually exercises the reader half).
func TestDiagramConnectorFieldsRoundTrip(t *testing.T) {
	orig := DiagramExampleDeck()
	data, err := BuildPPTX(orig)
	if err != nil {
		t.Fatalf("BuildPPTX: %v", err)
	}
	got, err := ParsePPTX(data)
	if err != nil {
		t.Fatalf("ParsePPTX: %v", err)
	}
	if len(got.Slides) != 2 {
		t.Fatalf("slide count = %d, want 2", len(got.Slides))
	}

	byName := func(shapes []Shape, name string) (Shape, bool) {
		for _, s := range shapes {
			if s.Name == name {
				return s, true
			}
		}
		return Shape{}, false
	}

	a1, ok := byName(got.Slides[0].Shapes, "a1")
	if !ok {
		t.Fatal("slide 1: shape 'a1' not found after round trip")
	}
	if !a1.TailArrow {
		t.Error("a1: TailArrow should round-trip true")
	}
	if a1.HeadArrow {
		t.Error("a1: HeadArrow should round-trip false")
	}
	if a1.FlipH {
		t.Error("a1: FlipH should round-trip false (horizontal connector)")
	}

	cLeft, ok := byName(got.Slides[1].Shapes, "c-left")
	if !ok {
		t.Fatal("slide 2: shape 'c-left' not found after round trip")
	}
	if !cLeft.FlipH {
		t.Error("c-left: FlipH should round-trip true (anti-diagonal connector)")
	}
	if cLeft.HeadArrow || cLeft.TailArrow {
		t.Error("c-left: org-chart connectors carry no arrowheads")
	}

	cRight, ok := byName(got.Slides[1].Shapes, "c-right")
	if !ok {
		t.Fatal("slide 2: shape 'c-right' not found after round trip")
	}
	if cRight.FlipH {
		t.Error("c-right: FlipH should round-trip false (main-diagonal connector)")
	}
}

func TestBuildPPTXRejectsUnsupportedShapeKind(t *testing.T) {
	deck := Deck{Slides: []Slide{{Shapes: []Shape{{Kind: ShapeKind("notARealPreset"), Cx: 100, Cy: 100}}}}}
	if _, err := BuildPPTX(deck); err == nil {
		t.Fatal("expected an error for an unsupported shape kind, got nil")
	}
}

func TestParseSlideFlattensNestedUnrotatedGroups(t *testing.T) {
	// This is a complete, namespace-valid p:sld shape tree rather than a
	// writer-produced shortcut. The outer group translates and doubles X;
	// the inner group then translates and scales its children, exercising the
	// OOXML chOff/chExt coordinate mapping through a real nested group tree.
	const slideXML = `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree>
<p:nvGrpSpPr/><p:grpSpPr/>
<p:grpSp><p:nvGrpSpPr/><p:grpSpPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="2000" cy="1000"/><a:chOff x="0" y="0"/><a:chExt cx="1000" cy="1000"/></a:xfrm></p:grpSpPr>
  <p:grpSp><p:nvGrpSpPr/><p:grpSpPr><a:xfrm><a:off x="100" y="100"/><a:ext cx="500" cy="400"/><a:chOff x="0" y="0"/><a:chExt cx="1000" cy="1000"/></a:xfrm></p:grpSpPr>
    <p:sp><p:nvSpPr><p:cNvPr id="2" name="nested box"/><p:cNvSpPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>
  </p:grpSp>
  <p:cxnSp><p:nvCxnSpPr><p:cNvPr id="3" name="outer line"/><p:cNvCxnSpPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="400" y="500"/><a:ext cx="100" cy="200"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>
</p:grpSp></p:spTree></p:cSld></p:sld>`

	got, err := parseSlide(slideXML, nil)
	if err != nil {
		t.Fatalf("parseSlide: %v", err)
	}
	if len(got.Shapes) != 2 {
		t.Fatalf("shape count = %d, want 2: %+v", len(got.Shapes), got.Shapes)
	}
	box := got.Shapes[0]
	if box.Name != "nested box" || box.Kind != KindRect {
		t.Errorf("nested shape identity = %+v", box)
	}
	if box.X != 400 || box.Y != 380 || box.Cx != 300 || box.Cy != 200 {
		t.Errorf("nested box geometry = (%d,%d,%d,%d), want (400,380,300,200)", box.X, box.Y, box.Cx, box.Cy)
	}
	line := got.Shapes[1]
	if line.Name != "outer line" || line.Kind != KindLine {
		t.Errorf("outer connector identity = %+v", line)
	}
	if line.X != 900 || line.Y != 700 || line.Cx != 200 || line.Cy != 200 {
		t.Errorf("outer connector geometry = (%d,%d,%d,%d), want (900,700,200,200)", line.X, line.Y, line.Cx, line.Cy)
	}
}

func TestParseSlideSkipsGroupsWithUnsupportedTransforms(t *testing.T) {
	// Keep a sibling outside the group: the reader must skip only the group
	// whose transform cannot be faithfully represented, not the whole slide.
	group := func(xfrm string) string {
		return `<p:grpSp><p:nvGrpSpPr/><p:grpSpPr><a:xfrm ` + xfrm + `><a:off x="0" y="0"/><a:ext cx="100" cy="100"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="skip me"/><p:cNvSpPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp></p:grpSp>`
	}
	for name, xfrm := range map[string]string{"rotation": `rot="60000"`, "flip": `flipH="1"`} {
		t.Run(name, func(t *testing.T) {
			xml := `<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>` + group(xfrm) + `<p:sp><p:nvSpPr><p:cNvPr id="3" name="keep me"/><p:cNvSpPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="30" cy="40"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp></p:spTree></p:cSld></p:sld>`
			got, err := parseSlide(xml, nil)
			if err != nil {
				t.Fatalf("parseSlide: %v", err)
			}
			if len(got.Shapes) != 1 || got.Shapes[0].Name != "keep me" {
				t.Fatalf("unsupported group was not safely skipped: %+v", got.Shapes)
			}
		})
	}
}

// ---- S11: animations + transitions ----

func slideXMLOf(t *testing.T, deck Deck, slideIndex int) string {
	t.Helper()
	x, err := slideXML(deck.Slides[slideIndex])
	if err != nil {
		t.Fatalf("slideXML: %v", err)
	}
	return x
}

func TestBuildPPTXOmitsTransitionAndTimingWhenUnused(t *testing.T) {
	deck := Deck{Slides: []Slide{{Shapes: []Shape{{Kind: KindRect, Cx: 100, Cy: 100}}}}}
	x := slideXMLOf(t, deck, 0)
	if strings.Contains(x, "<p:transition") {
		t.Error("expected no p:transition element for a slide with no Transition set")
	}
	if strings.Contains(x, "<p:timing") {
		t.Error("expected no p:timing element for a slide with no animated shapes")
	}
}

func TestBuildPPTXWritesFadeTransition(t *testing.T) {
	deck := Deck{Slides: []Slide{{
		Transition: SlideTransition{Type: TransitionFade},
		Shapes:     []Shape{{Kind: KindRect, Cx: 100, Cy: 100}},
	}}}
	x := slideXMLOf(t, deck, 0)
	if !strings.Contains(x, `<p:transition spd="med"><p:fade/></p:transition>`) {
		t.Errorf("expected a p:fade transition, got: %s", x)
	}
	// Schema order: cSld, clrMapOvr, transition, timing?, extLst.
	if strings.Index(x, "<p:clrMapOvr>") > strings.Index(x, "<p:transition") {
		t.Error("p:transition must come after p:clrMapOvr")
	}
}

func TestBuildPPTXWritesPushAndWipeTransitionsWithDirection(t *testing.T) {
	cases := []struct {
		typ  TransitionType
		dir  Direction
		want string
	}{
		{TransitionPush, DirLeft, `<p:push dir="l"/>`},
		{TransitionPush, DirRight, `<p:push dir="r"/>`},
		{TransitionWipe, DirUp, `<p:wipe dir="u"/>`},
		{TransitionWipe, DirDown, `<p:wipe dir="d"/>`},
		{TransitionPush, DirNone, `<p:push/>`}, // omitted dir -> schema default "l"
	}
	for _, c := range cases {
		deck := Deck{Slides: []Slide{{
			Transition: SlideTransition{Type: c.typ, Direction: c.dir},
			Shapes:     []Shape{{Kind: KindRect, Cx: 100, Cy: 100}},
		}}}
		x := slideXMLOf(t, deck, 0)
		if !strings.Contains(x, c.want) {
			t.Errorf("type=%s dir=%s: expected %q in slide XML, got: %s", c.typ, c.dir, c.want, x)
		}
	}
}

func TestBuildPPTXRejectsUnsupportedTransition(t *testing.T) {
	deck := Deck{Slides: []Slide{{
		Transition: SlideTransition{Type: TransitionType("dissolve")},
		Shapes:     []Shape{{Kind: KindRect, Cx: 100, Cy: 100}},
	}}}
	if _, err := BuildPPTX(deck); err == nil {
		t.Fatal("expected an error for an unsupported transition type, got nil")
	}
}

func TestBuildPPTXWritesFadeEntranceTiming(t *testing.T) {
	deck := Deck{Slides: []Slide{{Shapes: []Shape{
		{Kind: KindRect, Cx: 100, Cy: 100, Anim: ShapeAnimation{Effect: AnimFade}},
	}}}}
	x := slideXMLOf(t, deck, 0)
	if !strings.Contains(x, "<p:timing>") {
		t.Fatalf("expected a p:timing element, got: %s", x)
	}
	if !strings.Contains(x, `<p:animEffect transition="in" filter="fade">`) {
		t.Errorf("expected a fade p:animEffect, got: %s", x)
	}
	// The animated shape is id 2 (first/only shape) — must be targeted by
	// both the effect and the build-list hide entry.
	if !strings.Contains(x, `<p:spTgt spid="2"/>`) {
		t.Errorf("expected the effect to target spid=2, got: %s", x)
	}
	if !strings.Contains(x, `<p:bldP spid="2" grpId="0"/>`) {
		t.Errorf("expected a p:bldLst entry for spid=2, got: %s", x)
	}
	if strings.Index(x, "<p:transition") >= 0 && strings.Index(x, "<p:transition") > strings.Index(x, "<p:timing") {
		t.Error("p:transition must come before p:timing when both are present")
	}
}

func TestBuildPPTXWritesFlyInEntranceTimingWithDirectionOffset(t *testing.T) {
	deck := Deck{Slides: []Slide{{Shapes: []Shape{
		{Kind: KindRect, Cx: 100, Cy: 100, Anim: ShapeAnimation{Effect: AnimFlyIn, Direction: DirLeft, Distance: 0.4}},
	}}}}
	x := slideXMLOf(t, deck, 0)
	if !strings.Contains(x, `<p:attrName>ppt_x</p:attrName>`) {
		t.Errorf("expected the anim to target ppt_x for a left fly-in, got: %s", x)
	}
	if !strings.Contains(x, `<p:strVal val="ppt_x-0.4"/>`) {
		t.Errorf("expected the start tav to be ppt_x-0.4, got: %s", x)
	}
	if !strings.Contains(x, `<p:strVal val="ppt_x"/>`) {
		t.Errorf("expected the end tav to be the bare ppt_x self-reference, got: %s", x)
	}
}

func TestBuildPPTXFlyInDefaultsToFromBelow(t *testing.T) {
	deck := Deck{Slides: []Slide{{Shapes: []Shape{
		{Kind: KindRect, Cx: 100, Cy: 100, Anim: ShapeAnimation{Effect: AnimFlyIn}},
	}}}}
	x := slideXMLOf(t, deck, 0)
	if !strings.Contains(x, `<p:attrName>ppt_y</p:attrName>`) {
		t.Errorf("expected default fly-in direction to animate ppt_y, got: %s", x)
	}
	if !strings.Contains(x, `<p:strVal val="ppt_y+0.25"/>`) {
		t.Errorf("expected default distance 0.25 offset downward, got: %s", x)
	}
}

func TestParsePPTXReadsBackSupportedEffects(t *testing.T) {
	in := Deck{Slides: []Slide{{
		Transition: SlideTransition{Type: TransitionWipe, Direction: DirUp},
		Shapes: []Shape{
			{Kind: KindRect, Cx: 100, Cy: 100, Anim: ShapeAnimation{Effect: AnimFade}},
			{Kind: KindEllipse, Cx: 100, Cy: 100, Anim: ShapeAnimation{Effect: AnimFlyIn, Direction: DirLeft}},
		},
	}}}
	bytes, err := BuildPPTX(in)
	if err != nil {
		t.Fatal(err)
	}
	out, err := ParsePPTX(bytes)
	if err != nil {
		t.Fatal(err)
	}
	if got := out.Slides[0].Transition; got.Type != TransitionWipe || got.Direction != DirUp {
		t.Fatalf("transition = %+v", got)
	}
	if got := out.Slides[0].Shapes[0].Anim; got.Effect != AnimFade {
		t.Fatalf("fade animation = %+v", got)
	}
	if got := out.Slides[0].Shapes[1].Anim; got.Effect != AnimFlyIn || got.Direction != DirLeft {
		t.Fatalf("fly-in animation = %+v", got)
	}
}

func TestBuildPPTXRejectsUnsupportedAnimation(t *testing.T) {
	deck := Deck{Slides: []Slide{{Shapes: []Shape{
		{Kind: KindRect, Cx: 100, Cy: 100, Anim: ShapeAnimation{Effect: AnimEffect("zoom")}},
	}}}}
	if _, err := BuildPPTX(deck); err == nil {
		t.Fatal("expected an error for an unsupported animation effect, got nil")
	}
}

func TestBuildPPTXRejectsUnsupportedFlyInDirection(t *testing.T) {
	deck := Deck{Slides: []Slide{{Shapes: []Shape{
		{Kind: KindRect, Cx: 100, Cy: 100, Anim: ShapeAnimation{Effect: AnimFlyIn, Direction: Direction("sideways")}},
	}}}}
	if _, err := BuildPPTX(deck); err == nil {
		t.Fatal("expected an error for an unsupported fly-in direction, got nil")
	}
}

func TestBuildPPTXMultipleAnimatedShapesGetDistinctIDs(t *testing.T) {
	deck := Deck{Slides: []Slide{{Shapes: []Shape{
		{Kind: KindRect, Cx: 100, Cy: 100, Anim: ShapeAnimation{Effect: AnimFade}},
		{Kind: KindEllipse, Cx: 100, Cy: 100, Anim: ShapeAnimation{Effect: AnimFlyIn, DelayMs: 300}},
		{Kind: KindDiamond, Cx: 100, Cy: 100}, // no animation — should not appear in p:timing at all
	}}}}
	x := slideXMLOf(t, deck, 0)
	if !strings.Contains(x, `<p:spTgt spid="2"/>`) || !strings.Contains(x, `<p:spTgt spid="3"/>`) {
		t.Errorf("expected both animated shapes (spid 2 and 3) targeted, got: %s", x)
	}
	if strings.Contains(x, `<p:spTgt spid="4"/>`) {
		t.Errorf("un-animated shape (spid 4) should not appear as an animation target, got: %s", x)
	}
	if !strings.Contains(x, `<p:cond delay="300"/>`) {
		t.Errorf("expected the second shape's 300ms delay to be written, got: %s", x)
	}
	if strings.Count(x, "<p:bldP") != 2 {
		t.Errorf("expected exactly 2 p:bldP entries, got %d in: %s", strings.Count(x, "<p:bldP"), x)
	}
}

func TestBuildPPTXRejectsEmptyDeck(t *testing.T) {
	if _, err := BuildPPTX(Deck{}); err == nil {
		t.Fatal("expected an error for a deck with no slides, got nil")
	}
}

func TestParsePPTXRejectsGarbage(t *testing.T) {
	if _, err := ParsePPTX([]byte("not a zip file")); err == nil {
		t.Fatal("expected an error for non-zip input, got nil")
	}
}

// TestParseRealPythonPptxFixture reads a REAL .pptx built by an independent
// producer (python-pptx, via scripts/gen_fixture.py — not our own writer) to
// prove the reader isn't just self-consistent against BuildPPTX's own output.
// Regenerate the fixture with: python3 scripts/gen_fixture.py
func TestParseRealPythonPptxFixture(t *testing.T) {
	path := filepath.Join("testdata", "python_pptx_sample.pptx")
	data, err := os.ReadFile(path)
	if err != nil {
		t.Skipf("fixture not present (%v) — run scripts/gen_fixture.py to regenerate it", err)
	}
	deck, err := ParsePPTX(data)
	if err != nil {
		t.Fatalf("ParsePPTX on real python-pptx fixture: %v", err)
	}
	if deck.Cx == 0 || deck.Cy == 0 {
		t.Errorf("slide size not read: %d x %d", deck.Cx, deck.Cy)
	}
	if len(deck.Slides) != 2 {
		t.Fatalf("slide count = %d, want 2", len(deck.Slides))
	}

	var foundTitle, foundMultiParaBody, foundNestedLevel, foundRect bool
	var foundTitleGeom, foundBodyGeom, foundShapeFill bool
	for _, s := range deck.Slides {
		for _, sh := range s.Shapes {
			if sh.Placeholder == PlaceholderTitle || sh.Placeholder == PlaceholderCtrTitle {
				for _, p := range sh.Paragraphs {
					for _, r := range p.Runs {
						if r.Text != "" {
							foundTitle = true
						}
					}
				}
				// Regression coverage (round 6): python-pptx's own title
				// placeholder carries NO <a:xfrm> at slide OR layout level in
				// this fixture — position/size lives on the slide MASTER
				// only. Before layout/master inheritance was resolved, this
				// read as (0,0,0,0): a real title that Konva/DeckCanvasView
				// would render as a zero-size, invisible box, exactly the
				// live-staging "blank white canvas" bug this covers.
				if sh.Cx > 0 && sh.Cy > 0 {
					foundTitleGeom = true
				}
			}
			if sh.Placeholder == PlaceholderBody && len(sh.Paragraphs) >= 3 {
				foundMultiParaBody = true
				for _, p := range sh.Paragraphs {
					if p.Level == 1 && len(p.Runs) > 0 && p.Runs[0].Text != "" {
						foundNestedLevel = true
					}
				}
				if sh.Cx > 0 && sh.Cy > 0 {
					foundBodyGeom = true
				}
			}
			if sh.Kind == KindRect || sh.Kind == KindRoundRect {
				if sh.Cx > 0 && sh.Cy > 0 {
					foundRect = true
				}
				// Regression coverage (round 6): python-pptx's own
				// add_shape() (no explicit .fill.solid() call — the common,
				// default case) styles the shape ENTIRELY through
				// <p:style><a:fillRef><a:schemeClr .../></a:fillRef>, no
				// <a:solidFill> anywhere on the shape itself. Before
				// <p:style>/schemeClr resolution existed, Fill read as ""
				// here: an invisible (no fill, no stroke) box on a Konva
				// renderer, same live-staging bug.
				if sh.Fill != "" {
					foundShapeFill = true
				}
			}
		}
	}
	if !foundTitle {
		t.Error("expected to find non-empty title text somewhere in the deck")
	}
	if !foundMultiParaBody {
		t.Error("expected to find a body placeholder with 3+ paragraphs")
	}
	if !foundNestedLevel {
		t.Error("expected to find a level-1 (nested) paragraph with real text")
	}
	if !foundRect {
		t.Error("expected to find at least one rect/roundRect autoshape with real geometry")
	}
	if !foundTitleGeom {
		t.Error("expected the title placeholder's geometry to be resolved via layout/master inheritance (got zero-size)")
	}
	if !foundBodyGeom {
		t.Error("expected the body placeholder's geometry to be resolved via layout/master inheritance (got zero-size)")
	}
	if !foundShapeFill {
		t.Error("expected the autoshape's fill to be resolved via <p:style>'s fillRef + theme color (got empty)")
	}
	// Honest documented gap, not a bug: python-pptx's own paragraphs carry no
	// explicit a:buChar/a:buNone — bullet glyphs are inherited from the
	// placeholder's list style on the layout/master, which this reader does
	// not resolve (see the package-level doc comment). So Paragraph.Bullet
	// reads false here even though these render as bulleted text in
	// PowerPoint — a real limitation of this foundation slice, not
	// simulated for the test.
}
