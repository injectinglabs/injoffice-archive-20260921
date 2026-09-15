package docxpatch

import (
	"bytes"
	"encoding/json"
	"strings"
	"testing"
)

const testMC = "http://schemas.openxmlformats.org/markup-compatibility/2006"

// nativeApproximateShapeFixture wraps a wps shape in the markup-compatibility
// alternate Word writes, with a VML fallback that must never be read.
func nativeApproximateShapeFixture(container, shapeProperties, extra string) string {
	return `<mc:AlternateContent xmlns:mc="` + testMC + `"><mc:Choice Requires="wps"><w:drawing xmlns:wp="` + wordDrawingTransitional + `" xmlns:a="` + drawingMLTransitional + `" xmlns:wps="` + nativeTextboxWPS + `">` + container +
		`<a:graphic><a:graphicData uri="` + nativeTextboxWPS + `"><wps:wsp><wps:cNvSpPr/><wps:spPr>` + shapeProperties + `</wps:spPr>` + extra + `<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic>` +
		strings.TrimSuffix(strings.TrimPrefix(container, `<wp:anchor`), "") + `</w:drawing></mc:Choice><mc:Fallback><w:pict><v:rect xmlns:v="` + nativeTextboxVML + `" style="position:absolute;width:100pt;height:50pt" fillcolor="#ff0000"/></w:pict></mc:Fallback></mc:AlternateContent>`
}

func nativeApproximateAnchor(positionH, positionV string) string {
	return `<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>` + positionH + positionV + `<wp:extent cx="2998800" cy="2829600"/><wp:effectExtent l="0" t="0" r="11430" b="27940"/><wp:wrapNone/><wp:docPr id="2" name="Rectangle 2"/><wp:cNvGraphicFramePr/>`
}

func nativeApproximateShapeSource(t *testing.T, body string, parts func(map[string]string)) []byte {
	entries := nativeMutationParts(nativeMutationMain(body))
	if parts != nil {
		parts(entries)
	}
	return buildNativeDOCX(t, nativeEntries(entries))
}

func TestApproximateDrawingShapesAnchoredRect(t *testing.T) {
	// The anchor container closes after the graphic; build it explicitly.
	drawing := `<mc:AlternateContent xmlns:mc="` + testMC + `"><mc:Choice Requires="wps"><w:drawing xmlns:wp="` + wordDrawingTransitional + `" xmlns:a="` + drawingMLTransitional + `" xmlns:wps="` + nativeTextboxWPS + `">` +
		nativeApproximateAnchor(`<wp:positionH relativeFrom="page"><wp:align>center</wp:align></wp:positionH>`, `<wp:positionV relativeFrom="paragraph"><wp:posOffset>609600</wp:posOffset></wp:positionV>`) +
		`<a:graphic><a:graphicData uri="` + nativeTextboxWPS + `"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2998800" cy="2829600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4f81bd"/></a:solidFill><a:ln w="25400"><a:solidFill><a:srgbClr val="243F60"/></a:solidFill><a:prstDash val="dash"/></a:ln></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>` +
		`<mc:Fallback><w:pict><v:rect xmlns:v="` + nativeTextboxVML + `" style="position:absolute;width:100pt;height:50pt" fillcolor="#ff0000"/></w:pict></mc:Fallback></mc:AlternateContent>`
	source := nativeApproximateShapeSource(t, `<w:p><w:r><w:t>before</w:t></w:r><w:r>`+drawing+`</w:r><w:r><w:t>after</w:t></w:r></w:p>`, nil)
	before := append([]byte(nil), source...)
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	out, err := InspectNativeApproximateDrawingShapesV1(source)
	if err != nil {
		t.Fatal(err)
	}
	if out == nil || len(out.Items) != 1 || out.OmittedCount != 0 || out.Protocol != NativeApproximateDrawingShapesProtocol || out.Policy != NativeApproximateDrawingShapePolicy || out.PackageSHA256 != doc.Source.PackageSHA256 {
		t.Fatalf("unexpected sidecar: %#v", out)
	}
	item := out.Items[0]
	if item.Status != "supported" || item.Placement != "anchored" || item.Preset != "rect" || item.WidthEMU != 2998800 || item.HeightEMU != 2829600 || item.FillRGB == nil || *item.FillRGB != "4F81BD" || item.Line == nil || item.Line.RGB != "243F60" || item.Line.WidthEMU != 25400 || item.Line.Dash != "dash" || item.Wrap != "none" {
		t.Fatalf("unexpected shape: %#v", item)
	}
	anchor := item.PageAnchor
	if anchor == nil || anchor.Policy != "relative-position-no-wrap-v2" || anchor.HorizontalRelative != "page" || anchor.HorizontalAlign != "center" || anchor.XEMU != 0 || anchor.VerticalRelative != "paragraph" || anchor.YEMU != 609600 || anchor.VerticalAlign != "" || anchor.Stacking == nil || !anchor.Stacking.BehindDoc || anchor.Stacking.RelativeHeight != 251659264 {
		t.Fatalf("unexpected anchor: %#v", anchor)
	}
	paragraph := doc.Body.Blocks[0].Paragraph
	if item.ParagraphID != paragraph.ID || len(item.DiagnosticIDs) != 1 || !hasUnsupportedCode(doc, "UNMODELED_RUN_CONTENT") {
		t.Fatalf("shape must join the retained source refusal: %#v", item)
	}
	joined := false
	for _, d := range doc.Unsupported {
		if d.ID == item.DiagnosticIDs[0] && d.ScopeID == paragraph.ID && d.Anchor != nil && *d.Anchor.StartByte >= *item.RunAnchor.StartByte && *d.Anchor.EndByte <= *item.RunAnchor.EndByte {
			joined = true
		}
	}
	if !joined || *item.Anchor.StartByte < *item.RunAnchor.StartByte || *item.Anchor.EndByte > *item.RunAnchor.EndByte {
		t.Fatalf("diagnostic/run anchors do not nest: %#v", item)
	}
	for _, run := range paragraph.Runs {
		if *run.Anchor.StartByte >= *item.RunAnchor.StartByte && *run.Anchor.EndByte <= *item.RunAnchor.EndByte {
			t.Fatal("shape run overlaps a modeled run")
		}
	}
	if !bytes.Equal(source, before) {
		t.Fatal("source bytes changed")
	}
	if _, err := json.Marshal(out); err != nil {
		t.Fatal(err)
	}
}

func TestApproximateDrawingShapesInlineRectAndTheme(t *testing.T) {
	inline := `<w:drawing xmlns:wp="` + wordDrawingTransitional + `" xmlns:a="` + drawingMLTransitional + `" xmlns:wps="` + nativeTextboxWPS + `"><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="5940423" cy="25398"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="4" name="Rectangle 4"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="` + nativeTextboxWPS + `"><wps:wsp><wps:cNvSpPr/><wps:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="5940425" cy="25400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></wps:style><wps:bodyPr rot="0"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>`
	source := nativeApproximateShapeSource(t, `<w:p><w:r><w:t>text</w:t></w:r><w:r>`+inline+`</w:r></w:p>`, func(parts map[string]string) {
		parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>`, 1)
		parts["word/_rels/document.xml.rels"] = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="theme" Type="` + relBaseTransitional + `theme" Target="theme/theme1.xml"/></Relationships>`
		parts["word/theme/theme1.xml"] = `<a:theme xmlns:a="` + drawingMLTransitional + `" name="Office"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill/><a:gradFill/></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme></a:themeElements></a:theme>`
	})
	out, err := InspectNativeApproximateDrawingShapesV1(source)
	if err != nil {
		t.Fatal(err)
	}
	if out == nil || len(out.Items) != 1 {
		t.Fatalf("missing inline shape: %#v", out)
	}
	item := out.Items[0]
	if item.Status != "supported" || item.Placement != "inline" || item.PageAnchor != nil || item.WidthEMU != 5940423 || item.HeightEMU != 25398 || item.FillRGB == nil || *item.FillRGB != "4F81BD" || item.Line == nil || item.Line.WidthEMU != 25400 || item.Line.RGB != "28415F" || item.Line.Dash != "solid" {
		t.Fatalf("theme fill/line not resolved: %#v %v", item, item.Notes)
	}
	notes := strings.Join(item.Notes, "|")
	if !strings.Contains(notes, "theme fill style") || !strings.Contains(notes, "shade transform approximated") {
		t.Fatalf("theme approximations must be disclosed: %v", item.Notes)
	}
}

func TestApproximateDrawingShapesTextbox(t *testing.T) {
	content := `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="32"/></w:rPr><w:t xml:space="preserve">Heading </w:t></w:r><w:r><w:t>text</w:t></w:r></w:p><w:tbl/><w:p><w:r><w:t>second</w:t></w:r><w:r><w:footnoteReference w:id="7"/></w:r></w:p>`
	drawing := `<w:drawing xmlns:wp="` + wordDrawingTransitional + `" xmlns:a="` + drawingMLTransitional + `" xmlns:wps="` + nativeTextboxWPS + `">` +
		nativeApproximateAnchor(`<wp:positionH relativeFrom="page"><wp:posOffset>549275</wp:posOffset></wp:positionH>`, `<wp:positionV relativeFrom="page"><wp:posOffset>1983105</wp:posOffset></wp:positionV>`) +
		`<a:graphic><a:graphicData uri="` + nativeTextboxWPS + `"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2998800" cy="2829600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr><wps:txbx id="3"><w:txbxContent>` + content + `</w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>`
	linked := strings.Replace(strings.Replace(drawing, `<wps:txbx id="3"><w:txbxContent>`+content+`</w:txbxContent></wps:txbx>`, `<wps:linkedTxbx id="3" seq="1"/>`, 1), `<wp:docPr id="2" name="Rectangle 2"/>`, `<wp:docPr id="3" name="Rectangle 3"/>`, 1)
	source := nativeApproximateShapeSource(t, `<w:p><w:r>`+drawing+`</w:r><w:r>`+linked+`</w:r></w:p>`, nil)
	out, err := InspectNativeApproximateDrawingShapesV1(source)
	if err != nil {
		t.Fatal(err)
	}
	if out == nil || len(out.Items) != 2 {
		t.Fatalf("expected two shapes: %#v", out)
	}
	head, tail := out.Items[0], out.Items[1]
	if head.Status != "supported" || head.FillRGB != nil || head.Line != nil || head.Textbox == nil || tail.Textbox == nil {
		t.Fatalf("textbox shapes not admitted: %#v %#v", head, tail)
	}
	box := head.Textbox
	if box.LinkID != "3" || box.LinkSeq != 0 || box.VerticalAnchor != "ctr" || box.InsetsEMU != [4]int64{0, 0, 0, 0} || len(box.Paragraphs) != 2 || box.OmittedBlocks != 1 || box.OmittedRuns != 1 || len(box.ResolvedParagraphs) != 2 || len(box.ResolvedRuns) != 3 {
		t.Fatalf("unexpected textbox content: %+v", box)
	}
	if !strings.HasPrefix(box.Paragraphs[0].ID, head.ID) || box.ResolvedRuns[0].RunID != box.Paragraphs[0].Runs[0].ID || box.ResolvedRuns[0].Properties.Bold == nil || !*box.ResolvedRuns[0].Properties.Bold || box.ResolvedRuns[0].Properties.FontSizeHalfPoint == nil || *box.ResolvedRuns[0].Properties.FontSizeHalfPoint != 32 || box.ResolvedParagraphs[0].Properties.Alignment == nil || *box.ResolvedParagraphs[0].Properties.Alignment != "center" {
		t.Fatalf("textbox runs must resolve through the ordinary resolver: %+v", box.ResolvedRuns[0])
	}
	if tail.Textbox.LinkID != "3" || tail.Textbox.LinkSeq != 1 || len(tail.Textbox.Paragraphs) != 0 {
		t.Fatalf("linked continuation not described: %+v", tail.Textbox)
	}
	if head.ID == tail.ID {
		t.Fatal("shape ids must be unique")
	}
}

func TestApproximateDrawingShapesOmissions(t *testing.T) {
	base := `<w:drawing xmlns:wp="` + wordDrawingTransitional + `" xmlns:a="` + drawingMLTransitional + `" xmlns:wps="` + nativeTextboxWPS + `">` +
		nativeApproximateAnchor(`<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>`, `<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>`) +
		`<a:graphic><a:graphicData uri="` + nativeTextboxWPS + `"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm rot="0"><a:off x="0" y="0"/><a:ext cx="2998800" cy="2829600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="112233"/></a:solidFill></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>`
	for _, test := range []struct{ name, from, to, reason string }{
		{"ellipse", `prst="rect"`, `prst="ellipse"`, "unsupported-preset:ellipse"},
		{"adjust values", `<a:avLst/>`, `<a:avLst><a:gd name="adj" fmla="val 1"/></a:avLst>`, "adjust-values"},
		{"rotation", `<a:xfrm rot="0">`, `<a:xfrm rot="2700000">`, "rotation-unsupported"},
		{"group", `uri="` + nativeTextboxWPS + `"`, `uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"`, "unsupported-graphic:wsp"},
		{"simple position", `simplePos="0"`, `simplePos="1"`, "simple-position-unsupported"},
		{"line relative", `<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>`, `<wp:positionV relativeFrom="line"><wp:align>top</wp:align></wp:positionV>`, "unsupported-vertical-position"},
		{"scheme color without theme", `<a:srgbClr val="112233"/>`, `<a:schemeClr val="accent1"/>`, "unsupported-fill"},
	} {
		t.Run(test.name, func(t *testing.T) {
			drawing := strings.Replace(base, test.from, test.to, 1)
			if drawing == base {
				t.Fatal("fixture replacement did not apply")
			}
			source := nativeApproximateShapeSource(t, `<w:p><w:r>`+drawing+`</w:r></w:p>`, nil)
			out, err := InspectNativeApproximateDrawingShapesV1(source)
			if err != nil {
				t.Fatal(err)
			}
			if out == nil || len(out.Items) != 1 || out.Items[0].Status != "omitted" || out.Items[0].Reason != test.reason || out.Items[0].Textbox != nil || out.Items[0].PageAnchor != nil {
				t.Fatalf("expected omission %q: %#v", test.reason, out)
			}
		})
	}
	t.Run("gradient fill omits only the fill", func(t *testing.T) {
		drawing := strings.Replace(base, `<a:solidFill><a:srgbClr val="112233"/></a:solidFill>`, `<a:gradFill><a:gsLst/></a:gradFill><a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>`, 1)
		out, err := InspectNativeApproximateDrawingShapesV1(nativeApproximateShapeSource(t, `<w:p><w:r>`+drawing+`</w:r></w:p>`, nil))
		if err != nil {
			t.Fatal(err)
		}
		if out == nil || len(out.Items) != 1 || out.Items[0].Status != "supported" || out.Items[0].FillRGB != nil || out.Items[0].Line == nil || !strings.Contains(strings.Join(out.Items[0].Notes, "|"), "gradFill") {
			t.Fatalf("gradient fill handling: %#v", out)
		}
	})
	t.Run("line preset keeps flips", func(t *testing.T) {
		drawing := strings.Replace(strings.Replace(base, `prst="rect"`, `prst="line"`, 1), `<a:xfrm rot="0">`, `<a:xfrm rot="5400000" flipH="1">`, 1)
		drawing = strings.Replace(drawing, `<a:solidFill><a:srgbClr val="112233"/></a:solidFill>`, `<a:ln w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>`, 1)
		out, err := InspectNativeApproximateDrawingShapesV1(nativeApproximateShapeSource(t, `<w:p><w:r>`+drawing+`</w:r></w:p>`, nil))
		if err != nil {
			t.Fatal(err)
		}
		item := out.Items[0]
		if item.Status != "supported" || item.Preset != "line" || item.RotationDegrees != 90 || !item.FlipHorizontal || item.FlipVertical || item.Line == nil || item.Line.WidthEMU != 19050 || item.FillRGB != nil {
			t.Fatalf("line preset: %#v", item)
		}
	})
	t.Run("shared run with text", func(t *testing.T) {
		// Modeled text runs anchor inside the same w:r; the shape must not claim it.
		source := nativeApproximateShapeSource(t, `<w:p><w:r><w:t>a</w:t>`+base+`<w:t>b</w:t></w:r></w:p>`, nil)
		doc, err := ExtractNativeDocumentV1(source)
		if err != nil {
			t.Fatal(err)
		}
		out, err := InspectNativeApproximateDrawingShapesV1(source)
		if err != nil {
			t.Fatal(err)
		}
		if out == nil || len(out.Items) != 1 || out.Items[0].Status != "omitted" || out.Items[0].Reason != "shared-run" {
			t.Fatalf("shared run must be omitted: %#v", out)
		}
		if runs := doc.Body.Blocks[0].Paragraph.Runs; len(runs) != 2 || *runs[0].Text != "a" || *runs[1].Text != "b" {
			t.Fatalf("text runs must stay modeled: %#v", runs)
		}
	})
	t.Run("vml only stays refused", func(t *testing.T) {
		source := nativeApproximateShapeSource(t, `<w:p><w:r><w:pict><v:rect xmlns:v="`+nativeTextboxVML+`" style="width:10pt;height:10pt" fillcolor="#ff0000"/></w:pict></w:r></w:p>`, nil)
		out, err := InspectNativeApproximateDrawingShapesV1(source)
		if err != nil {
			t.Fatal(err)
		}
		if out != nil {
			t.Fatalf("VML-only drawings are not described: %#v", out)
		}
	})
	t.Run("no drawings", func(t *testing.T) {
		out, err := InspectNativeApproximateDrawingShapesV1(nativeApproximateShapeSource(t, `<w:p><w:r><w:t>plain</w:t></w:r></w:p>`, nil))
		if err != nil || out != nil {
			t.Fatalf("expected nil sidecar: %#v %v", out, err)
		}
	})
	t.Run("shape budget", func(t *testing.T) {
		run := `<w:r>` + base + `</w:r>`
		out, err := InspectNativeApproximateDrawingShapesV1(nativeApproximateShapeSource(t, `<w:p>`+strings.Repeat(run, nativeApproximateDrawingShapeLimit+3)+`</w:p>`, nil))
		if err != nil {
			t.Fatal(err)
		}
		if out == nil || len(out.Items) != nativeApproximateDrawingShapeLimit || out.OmittedCount != 3 {
			t.Fatalf("shape budget: %d items, %d omitted", len(out.Items), out.OmittedCount)
		}
	})
}
