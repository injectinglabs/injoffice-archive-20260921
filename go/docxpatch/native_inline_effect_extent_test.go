package docxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func TestNativeInlineEffectExtents(t *testing.T) {
	for _, attrs := range []string{`l="19050" t="0" r="9525" b="0"`, `l="127" t="254" r="381" b="508"`, `l="-1" t="0" r="0" b="0"`, `l="91440001" t="0" r="0" b="0"`, `l="x" t="0" r="0" b="0"`, `l="127" t="0" r="0"`, `l="127" t="0" r="0" b="0" extra="x"`} {
		parts := transitionalNativeParts()
		parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `l="0" t="0" r="0" b="0"`, attrs, 1)
		data := buildNativeDOCX(t, nativeEntries(parts))
		before := bytes.Clone(data)
		doc, err := ExtractNativeDocumentV1(data)
		if err != nil {
			t.Fatal(err)
		}
		var drawing *NativeDrawingV1
		for _, run := range doc.Body.Blocks[0].Paragraph.Runs {
			if run.Drawing != nil {
				drawing = run.Drawing
			}
		}
		want := strings.HasPrefix(attrs, `l="19050"`) || strings.HasPrefix(attrs, `l="127" t="254"`)
		if (drawing != nil) != want {
			t.Fatalf("%s drawing=%#v", attrs, drawing)
		}
		if drawing != nil && (drawing.InlineEffectExtentEMU == nil || *drawing.WidthEMU != 914400 || *drawing.HeightEMU != 457200 || drawing.EditPolicy.Mode != "read-only") {
			t.Fatal("image source authority changed")
		}
		if !bytes.Equal(data, before) {
			t.Fatal("source bytes changed")
		}
	}
}

func TestNativeEffectExtentZeroCannotHideMarkup(t *testing.T) {
	for _, effect := range []string{`<wp:effectExtent l="0" t="0" r="0" b="0">text</wp:effectExtent>`, `<wp:effectExtent l="0" t="0" r="0" b="0"><extra/></wp:effectExtent>`, `<wp:effectExtent l="0" t="0" r="0" b="0" extra="1"/>`, `<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:effectExtent l="0" t="0" r="0" b="0"/>`} {
		parts := transitionalNativeParts()
		parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `<wp:effectExtent l="0" t="0" r="0" b="0"/>`, effect, 1)
		doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
		if err != nil {
			t.Fatal(err)
		}
		for _, run := range doc.Body.Blocks[0].Paragraph.Runs {
			if run.Drawing != nil {
				t.Fatalf("accepted %s", effect)
			}
		}
	}
}
