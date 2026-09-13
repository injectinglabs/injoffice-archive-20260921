package docxpatch

import (
	"encoding/json"
	"strings"
	"testing"
)

func nativeHardBreakFixture(lines int) string {
	text := []string{}
	for i := 0; i < lines; i++ {
		text = append(text, `<w:t xml:space="preserve">Authored line</w:t>`)
	}
	s := strings.Replace(nativeGeometryFixture(), `<w:t xml:space="preserve">Rectangle source</w:t>`, strings.Join(text, `<w:br w:type="textWrapping" w:clear="none"/>`), 1)
	return strings.Replace(s, `w:line="240" w:lineRule="auto"`, `w:line="360" w:lineRule="exact"`, 1)
}
func inspectHardBreakFixture(t *testing.T, drawing string) (NativeTextboxGeometryItemV1, string) {
	t.Helper()
	main := nativeMutationMain(`<w:p><w:r>` + drawing + `</w:r></w:p>`)
	source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(main)))
	data, err := InspectNativePartialSourceV1(source)
	if err != nil {
		t.Fatal(err)
	}
	var result struct {
		Geometry NativeTextboxGeometryEvidenceV1 `json:"textbox_geometry"`
	}
	if err = json.Unmarshal(data, &result); err != nil {
		t.Fatal(err)
	}
	if len(result.Geometry.Items) != 1 {
		t.Fatalf("missing evidence: %s", data)
	}
	return result.Geometry.Items[0], main
}
func TestNativeTextboxHardBreakSource(t *testing.T) {
	for _, count := range []int{2, 3, 16} {
		item, raw := inspectHardBreakFixture(t, nativeHardBreakFixture(count))
		h := item.HardBreakLayout
		if item.Geometry == nil || h == nil || h.LineStepTwips != 360 || len(h.Lines) != count || len(item.Owner.Paragraphs) != 1 {
			t.Fatalf("missing layout: %#v", item)
		}
		offset := 0
		for i, line := range h.Lines {
			if i > 0 {
				offset++
			}
			if line.Ordinal != i || line.StartUTF16 != offset || line.EndUTF16 != offset+len("Authored line") {
				t.Fatal("wrong text coverage")
			}
			offset = line.EndUTF16
			if nativeSHA([]byte(raw[*line.TextAnchor.StartByte:*line.TextAnchor.EndByte])) != line.TextAnchor.XMLSHA256 {
				t.Fatal("wrong source hash")
			}
			if i == 0 {
				if line.BreakBeforeAnchor != nil {
					t.Fatal("leading break")
				}
			} else {
				a := line.BreakBeforeAnchor
				if a == nil || nativeSHA([]byte(raw[*a.StartByte:*a.EndByte])) != a.XMLSHA256 || *a.EndByte > *line.TextAnchor.StartByte {
					t.Fatal("wrong break join")
				}
			}
		}
	}
}
func TestNativeTextboxHardBreakRefusals(t *testing.T) {
	for _, test := range []struct{ name, from, to string }{
		{"page", `w:type="textWrapping"`, `w:type="page"`}, {"column", `w:type="textWrapping"`, `w:type="column"`}, {"clear", `w:clear="none"`, `w:clear="all"`}, {"absent clear", ` w:clear="none"`, ``}, {"unknown attr", `w:clear="none"`, `w:clear="none" bogus="1"`},
		{"auto spacing", `w:lineRule="exact"`, `w:lineRule="auto"`}, {"zero spacing", `w:line="360"`, `w:line="0"`}, {"spacing cap", `w:line="360"`, `w:line="25601"`}, {"empty text", `>Authored line</w:t>`, `></w:t>`}, {"raw newline", `Authored line`, "Authored\nline"}, {"unknown text", `xml:space="preserve"`, `xml:space="preserve" unknown="1"`},
		{"extra run", `<w:br w:type="textWrapping" w:clear="none"/>`, `</w:r><w:r><w:br w:type="textWrapping" w:clear="none"/>`},
	} {
		t.Run(test.name, func(t *testing.T) {
			item, _ := inspectHardBreakFixture(t, strings.ReplaceAll(nativeHardBreakFixture(2), test.from, test.to))
			if item.Geometry != nil || item.HardBreakLayout != nil || len(item.Owner.Paragraphs) != 0 {
				t.Fatal("unsafe source admitted")
			}
		})
	}
	for _, drawing := range []string{nativeHardBreakFixture(17), strings.ReplaceAll(nativeHardBreakFixture(2), "Authored line", strings.Repeat("x", 2048))} {
		item, _ := inspectHardBreakFixture(t, drawing)
		if item.Geometry != nil || item.HardBreakLayout != nil {
			t.Fatal("aggregate limit ignored")
		}
	}
}
