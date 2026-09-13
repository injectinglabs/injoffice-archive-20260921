package docxpatch

import (
	"strings"
	"testing"
)

func nativeWrapFixture() string {
	return strings.ReplaceAll(strings.ReplaceAll(nativeGeometryFixture(), `wrap="none"`, `wrap="square"`), `w:line="240" w:lineRule="auto"`, `w:line="360" w:lineRule="exact"`)
}
func TestNativeTextboxWrappingSource(t *testing.T) {
	item, raw := inspectHardBreakFixture(t, nativeWrapFixture())
	w := item.WrapLayout
	if item.Geometry == nil || item.Geometry.TextWrap != "square" || w == nil || item.HardBreakLayout != nil || w.Policy != nativeTextboxWrapPolicy || w.LineStepTwips != 360 || item.Owner.Paragraphs[0] != "Rectangle source" {
		t.Fatalf("missing wrap evidence: %#v", item)
	}
	for _, a := range []NativeSourceAnchorV1{w.BodyPropertiesAnchor, w.ParagraphAnchor, w.RunAnchor, w.TextAnchor, w.SpacingAnchor} {
		if nativeSHA([]byte(raw[*a.StartByte:*a.EndByte])) != a.XMLSHA256 {
			t.Fatal("wrong source slice hash")
		}
	}
}
func TestNativeTextboxWrappingRefusals(t *testing.T) {
	for _, test := range []struct{ name, from, to string }{
		{"default wrap", ` wrap="square"`, ``}, {"unknown wrap", `wrap="square"`, `wrap="tight"`}, {"auto height", `w:lineRule="exact"`, `w:lineRule="auto"`}, {"zero height", `w:line="360"`, `w:line="0"`}, {"height cap", `w:line="360"`, `w:line="25601"`}, {"punctuation", `Rectangle source`, `Rectangle source.`}, {"unicode", `Rectangle source`, `Réctangle`}, {"double space", `Rectangle source`, `Rectangle  source`}, {"leading space", `Rectangle source`, ` Rectangle source`}, {"trailing space", `Rectangle source`, `Rectangle source `}, {"newline", `Rectangle source`, "Rectangle\nsource"}, {"missing preserve", ` xml:space="preserve"`, ``}, {"autofit", `a:noAutofit`, `a:spAutoFit`}, {"extra text attr", `xml:space="preserve"`, `xml:space="preserve" bogus="1"`}, {"hard break", `Rectangle source</w:t>`, `Rectangle</w:t><w:br w:type="textWrapping" w:clear="none"/><w:t xml:space="preserve">source</w:t>`},
	} {
		t.Run(test.name, func(t *testing.T) {
			item, _ := inspectHardBreakFixture(t, strings.ReplaceAll(nativeWrapFixture(), test.from, test.to))
			if item.Geometry != nil || item.WrapLayout != nil || len(item.Owner.Paragraphs) != 0 {
				t.Fatal("unsafe wrapping admitted")
			}
		})
	}
	item, _ := inspectHardBreakFixture(t, strings.Replace(nativeWrapFixture(), "Rectangle source", strings.Repeat("a", 4097), 1))
	if item.Geometry != nil {
		t.Fatal("source budget ignored")
	}
}
