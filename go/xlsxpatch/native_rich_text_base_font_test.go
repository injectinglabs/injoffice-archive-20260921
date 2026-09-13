package xlsxpatch

import (
	"fmt"
	"strings"
	"testing"
)

func TestNativeRichBaseDecorationRefusesFallback(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, decoration := range []string{`<u/>`, `<u val="single"/>`, `<u val="none"/>`, `<strike/>`, `<strike val="false"/>`, `<outline/>`, `<shadow/>`, `<condense/>`, `<extend/>`, `<vertAlign val="baseline"/>`, `<u xmlns="urn:foreign"/>`, `<unknown/>`} {
			for _, runs := range []string{`<r><t>inherited</t></r>`, `<r><rPr><b val="false"/><i val="false"/><color rgb="FF000000"/><rFont val="Calibri"/><sz val="11"/></rPr><t>direct</t></r>`} {
				p := richFixture(strict, runs)
				p["Meta/Styles.style"] = strings.Replace(p["Meta/Styles.style"], `</font>`, decoration+`</font>`, 1)
				r := richPreviewFixture(t, p)
				if len(r.Cells) != 1 || r.Cells[0].Status != "omitted" || r.Cells[0].Runs != nil {
					t.Fatalf("base underline falsely qualified: %s %#v", decoration, r)
				}
			}
		}
	}
}

func TestNativeRichBaseFontEffectiveSelection(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, tt := range []struct {
			name                         string
			parentFont, childFont, apply int
			omitted                      bool
		}{
			{"unselected", 0, 0, 0, false}, {"parent-applied", 1, 1, 0, true}, {"child-applied", 0, 1, 1, true}, {"child-overrides-decorated-parent", 1, 0, 1, false},
		} {
			t.Run(tt.name, func(t *testing.T) {
				p := richFixture(strict, richPositiveRuns)
				styles := p["Meta/Styles.style"]
				start, end := strings.Index(styles, "<cellStyleXfs"), strings.Index(styles, "</cellXfs>")+len("</cellXfs>")
				styles = styles[:start] + fmt.Sprintf(`<cellStyleXfs count="1"><xf numFmtId="0" fontId="%d" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="%d" fillId="0" borderId="0" xfId="0" applyFont="%d"/></cellXfs>`, tt.parentFont, tt.childFont, tt.apply) + styles[end:]
				first := strings.Index(styles, "</font>")
				second := first + len("</font>") + strings.Index(styles[first+len("</font>"):], "</font>")
				styles = styles[:second] + `<strike/>` + styles[second:]
				p["Meta/Styles.style"] = styles
				r := richPreviewFixture(t, p)
				if len(r.Cells) != 1 || (r.Cells[0].Status == "omitted") != tt.omitted {
					t.Fatalf("effective font selection: %#v", r)
				}
			})
		}
	}
}
