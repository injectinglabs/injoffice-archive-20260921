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
