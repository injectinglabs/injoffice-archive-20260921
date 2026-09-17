package docxpatch

import (
	"strings"
	"testing"
)

func firstNativeBodyDrawing(doc *NativeDocumentV1) *NativeDrawingV1 {
	for _, block := range doc.Body.Blocks {
		if block.Paragraph == nil {
			continue
		}
		for _, run := range block.Paragraph.Runs {
			if run.Drawing != nil {
				return run.Drawing
			}
		}
	}
	return nil
}

// Word writes a wp14 revision-identity pair on every drawing container and a
// a14:useLocalDpi resampling hint inside the blip of every picture it inserts.
// Neither states extent, position or wrap, so neither may turn the picture into
// preserve-only markup. Malformed identity and every other blip extension keep
// refusing by name.
func TestExtractNativePictureInertWordMarkup(t *testing.T) {
	const wp14 = `xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"`
	const a14 = `xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"`
	for _, test := range []struct {
		name    string
		from    string
		to      string
		refusal string
	}{
		{name: "revision identity on the inline container", from: `<wp:inline distT="0"`, to: `<wp:inline ` + wp14 + ` wp14:anchorId="02B88582" wp14:editId="7BE67881" distT="0"`},
		{name: "unqualified revision identity", from: `<wp:inline distT="0"`, to: `<wp:inline ` + wp14 + ` wp14:anchorId="not-hex" distT="0"`, refusal: "INLINE_DRAWING_SEMANTICS_PRESERVED"},
		{name: "local dpi hint on the blip", from: `<a:blip r:embed="rImage"/>`, to: `<a:blip r:embed="rImage"><a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"><a14:useLocalDpi ` + a14 + ` val="0"/></a:ext></a:extLst></a:blip>`},
		{name: "unknown blip extension", from: `<a:blip r:embed="rImage"/>`, to: `<a:blip r:embed="rImage"><a:extLst><a:ext uri="{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}"><a14:imgProps ` + a14 + `/></a:ext></a:extLst></a:blip>`, refusal: "PICTURE_TRANSFORM_PRESERVED"},
	} {
		t.Run(test.name, func(t *testing.T) {
			parts := transitionalNativeParts()
			replaced := strings.Replace(parts["Custom/Main.XML"], test.from, test.to, 1)
			if replaced == parts["Custom/Main.XML"] {
				t.Fatalf("fixture no longer contains %q", test.from)
			}
			parts["Custom/Main.XML"] = replaced
			doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			drawing := firstNativeBodyDrawing(doc)
			if test.refusal != "" {
				if drawing != nil || !hasUnsupportedCode(doc, test.refusal) {
					t.Fatalf("expected %s, drawing = %#v, unsupported = %#v", test.refusal, drawing, doc.Unsupported)
				}
				return
			}
			if drawing == nil || *drawing.WidthEMU != 914400 || *drawing.HeightEMU != 457200 {
				t.Fatalf("inert Word picture markup must stay projected: %#v", drawing)
			}
			if !hasNativePassthrough(doc, "Custom/Media/image.PNG") {
				t.Fatal("picture bytes must remain passthrough")
			}
		})
	}
}

// wp:extent states the drawing object's final size in the document, so an
// unrotated DrawingML shape extent adds no layout fact and Word's own writer
// rounds the two apart. A quarter turn still needs the shape extent to be the
// transpose of the painted box.
func TestExtractNativePictureExtentIsTheInlineExtent(t *testing.T) {
	for _, test := range []struct {
		name     string
		xfrm     string
		rotation int64
		refused  bool
	}{
		{name: "unrotated shape extent disagrees", xfrm: `<a:xfrm><a:off x="0" y="0"/><a:ext cx="914401" cy="457199"/></a:xfrm>`},
		{name: "flipped shape extent disagrees", xfrm: `<a:xfrm flipH="1"><a:off x="0" y="0"/><a:ext cx="912000" cy="456000"/></a:xfrm>`},
		{name: "quarter turn transposes the painted box", xfrm: `<a:xfrm rot="5400000"><a:off x="0" y="0"/><a:ext cx="457200" cy="914400"/></a:xfrm>`, rotation: 90},
		{name: "quarter turn without the transpose", xfrm: `<a:xfrm rot="5400000"><a:off x="0" y="0"/><a:ext cx="914401" cy="457200"/></a:xfrm>`, refused: true},
	} {
		t.Run(test.name, func(t *testing.T) {
			parts := transitionalNativeParts()
			parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `<a:xfrm/>`, test.xfrm, 1)
			doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			drawing := firstNativeBodyDrawing(doc)
			if test.refused {
				if drawing != nil || !hasUnsupportedCode(doc, "PICTURE_TRANSFORM_PRESERVED") {
					t.Fatalf("a rotated picture without the transposed shape extent must stay preserve-only: %#v", drawing)
				}
				return
			}
			if drawing == nil || *drawing.WidthEMU != 914400 || *drawing.HeightEMU != 457200 {
				t.Fatalf("the painted box must come from wp:extent: %#v", drawing)
			}
			rotation := int64(0)
			if drawing.RotationDegrees != nil {
				rotation = *drawing.RotationDegrees
			}
			if rotation != test.rotation {
				t.Fatalf("rotation = %d, want %d", rotation, test.rotation)
			}
		})
	}
}
