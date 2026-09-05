package pptxpatch

// ExampleDeck is a small hand-built Deck exercising every shape kind and
// text feature this foundation slice supports: a title+bulleted-body slide
// (with a nested bullet level, bold/italic/colored runs) plus a few
// autoshapes and a line connector, and a second, simpler centered-title
// slide. Also exercises S11's animations/transitions (slide 1 fades in with
// a fade-entrance callout and a fly-in-from-the-bottom dot; slide 2 pushes
// in from the left) so both the round-trip test and cmd/sample's
// python-pptx validation cover them. Used by the round-trip test AND by
// cmd/sample (which writes it to testdata/generated_sample.pptx for
// scripts/validate_with_python_pptx.py — the independent-consumer check on
// the WRITE side, mirroring gen_fixture.py's independent-producer check on
// the READ side).
func ExampleDeck() Deck {
	return Deck{
		Slides: []Slide{
			{
				Background: "#0B1220",
				Transition: SlideTransition{Type: TransitionFade},
				Shapes: []Shape{
					{
						Kind:        KindTextBox,
						Placeholder: PlaceholderTitle,
						Name:        "Title",
						X:           838200, Y: 365125, Cx: 10515600, Cy: 1325563,
						Paragraphs: []Paragraph{
							{Align: AlignLeft, Runs: []TextRun{{Text: "InjOffice Slides", Bold: true, SizePt: 40, Color: "#F5F7FA", Font: "Archivo"}}},
						},
					},
					{
						Kind:        KindTextBox,
						Placeholder: PlaceholderBody,
						Name:        "Body",
						X:           838200, Y: 1825625, Cx: 10515600, Cy: 4351338,
						Paragraphs: []Paragraph{
							{Bullet: true, Level: 0, Runs: []TextRun{{Text: "Real OOXML, not a template render", SizePt: 20}}},
							{Bullet: true, Level: 0, Runs: []TextRun{{Text: "Round-trips through PowerPoint", SizePt: 20, Italic: true}}},
							{Bullet: true, Level: 1, Runs: []TextRun{{Text: "Nested bullet", SizePt: 18}}},
						},
					},
					{
						Kind: KindRoundRect, Name: "Callout",
						X: 838200, Y: 5800725, Cx: 3200400, Cy: 838200,
						Fill: "#2F6FED", Stroke: "#153A82", StrokeWidthPt: 2,
						Paragraphs: []Paragraph{{Align: AlignCenter, Runs: []TextRun{{Text: "Shipped", Color: "#FFFFFF", SizePt: 16}}}},
						Anim:       ShapeAnimation{Effect: AnimFade},
					},
					{
						Kind: KindEllipse, Name: "Dot",
						X: 9500000, Y: 5800725, Cx: 500000, Cy: 500000,
						Fill: "#ED7D31",
						Anim: ShapeAnimation{Effect: AnimFlyIn, Direction: DirDown, DelayMs: 300},
					},
					{
						Kind: KindLine, Name: "Divider",
						X: 838200, Y: 1700000, Cx: 10515600, Cy: 0,
						Stroke: "#A5A5A5", StrokeWidthPt: 1,
					},
				},
			},
			{
				Transition: SlideTransition{Type: TransitionPush, Direction: DirLeft},
				Shapes: []Shape{
					{
						Kind: KindTextBox, Placeholder: PlaceholderCtrTitle, Name: "Section",
						X: 838200, Y: 2500000, Cx: 10515600, Cy: 1200000,
						Paragraphs: []Paragraph{{Align: AlignCenter, Runs: []TextRun{{Text: "Thank you", SizePt: 32}}}},
					},
				},
			},
		},
	}
}

// DiagramExampleDeck exercises S12's diagram slice: a process-flow row
// (boxes + arrowed connectors, HeadArrow/TailArrow) and an org-chart tree
// (boxes + plain connectors, exercising both the flipped and unflipped
// a:xfrm diagonal — see Shape.FlipH's doc comment). These are composed
// entirely from existing primitives (roundRect/rect boxes, KindLine
// connectors); there is no new shape kind, just the two new Shape fields
// (HeadArrow/TailArrow/FlipH) diagram layouts need to place connectors
// correctly. Used by cmd/diagram_sample for
// scripts/validate_with_python_pptx.py's independent check.
func DiagramExampleDeck() Deck {
	box := func(name string, x, y, cx, cy int, text string) Shape {
		return Shape{
			Kind: KindRoundRect, Name: name,
			X: x, Y: y, Cx: cx, Cy: cy,
			Fill: "#2F6FED", Stroke: "#153A82", StrokeWidthPt: 1,
			Paragraphs: []Paragraph{{Align: AlignCenter, Runs: []TextRun{{Text: text, Color: "#FFFFFF", SizePt: 14}}}},
		}
	}
	arrow := func(name string, x1, y1, x2, y2 int) Shape {
		return Shape{
			Kind: KindLine, Name: name,
			X: x1, Y: y1, Cx: x2 - x1, Cy: y2 - y1,
			Stroke: "#5C6B6D", StrokeWidthPt: 1.5, TailArrow: true,
		}
	}
	// connector draws a plain (no arrowhead) line between two arbitrary
	// points, normalizing to a non-negative a:xfrm bounding box and setting
	// FlipH when the points fall on the box's OTHER diagonal — see
	// Shape.FlipH's doc comment for the derivation this mirrors.
	connector := func(name string, x1, y1, x2, y2 int) Shape {
		offX, offY := x1, y1
		if x2 < offX {
			offX = x2
		}
		if y2 < offY {
			offY = y2
		}
		cx, cy := x1-x2, y1-y2
		if cx < 0 {
			cx = -cx
		}
		if cy < 0 {
			cy = -cy
		}
		// Anti-diagonal iff x and y orderings disagree; a degenerate
		// horizontal (cy=0) or vertical (cx=0) line has no "diagonal" at
		// all, so flip is forced false there — meaningless either way, but
		// leaving it unset is the honest value.
		flip := cx > 0 && cy > 0 && (x1 < x2) != (y1 < y2)
		return Shape{Kind: KindLine, Name: name, X: offX, Y: offY, Cx: cx, Cy: cy, Stroke: "#5C6B6D", StrokeWidthPt: 1, FlipH: flip}
	}

	const boxW, boxH = 2200000, 900000
	const rowY = 2500000
	step := func(i int) int { return 838200 + i*(boxW+400000) }

	root := box("CEO", 838200*5, 900000, boxW, boxH, "CEO")
	left := box("CTO", 838200, rowY, boxW, boxH, "CTO")
	mid := box("COO", 838200*5, rowY, boxW, boxH, "COO")
	right := box("CFO", 838200*9, rowY, boxW, boxH, "CFO")
	rootCx := root.X + root.Cx/2
	rootBottom := root.Y + root.Cy

	return Deck{
		Slides: []Slide{
			{
				Shapes: []Shape{
					box("s1", step(0), rowY, boxW, boxH, "Discover"),
					box("s2", step(1), rowY, boxW, boxH, "Design"),
					box("s3", step(2), rowY, boxW, boxH, "Build"),
					box("s4", step(3), rowY, boxW, boxH, "Ship"),
					arrow("a1", step(0)+boxW, rowY+boxH/2, step(1), rowY+boxH/2),
					arrow("a2", step(1)+boxW, rowY+boxH/2, step(2), rowY+boxH/2),
					arrow("a3", step(2)+boxW, rowY+boxH/2, step(3), rowY+boxH/2),
				},
			},
			{
				Shapes: []Shape{
					root, left, mid, right,
					connector("c-left", rootCx, rootBottom, left.X+left.Cx/2, left.Y),     // left child: anti-diagonal, needs FlipH
					connector("c-mid", rootCx, rootBottom, mid.X+mid.Cx/2, mid.Y),         // aligned: dx=0, degenerate vertical line
					connector("c-right", rootCx, rootBottom, right.X+right.Cx/2, right.Y), // right child: main diagonal, no flip
				},
			},
		},
	}
}
