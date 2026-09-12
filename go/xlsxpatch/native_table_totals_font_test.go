package xlsxpatch

import (
	"strings"
	"testing"
)

func TestTableTotalsBoldQualification(t *testing.T) {
	for _, variant := range []string{"default", "explicit-font", "nondefault-font", "no-totals", "font-dxf", "named-style"} {
		parts := tableNumberFixture()
		parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/theme/theme1.xml" ContentType="`+themePartContentType+`"/></Types>`, 1)
		parts["Book/_rels/Workbook.xml.rels"] = strings.Replace(parts["Book/_rels/Workbook.xml.rels"], `</Relationships>`, `<Relationship Id="theme" Type="`+relTypeThemeTransitional+`" Target="../theme/theme1.xml"/></Relationships>`, 1)
		parts["theme/theme1.xml"] = `<a:theme xmlns:a="` + drawingMLNamespace + `"><a:themeElements><a:clrScheme><a:accent1><a:srgbClr val="156082"/></a:accent1></a:clrScheme></a:themeElements></a:theme>`
		switch variant {
		case "explicit-font":
			parts["Meta/Styles.style"] = strings.Replace(parts["Meta/Styles.style"], `xfId="0"/>`, `xfId="0" applyFont="1"/>`, 1)
		case "nondefault-font":
			parts["Meta/Styles.style"] = strings.Replace(parts["Meta/Styles.style"], `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`, `<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>`, 1)
		case "no-totals":
			parts["Tables/table.xml"] = strings.Replace(parts["Tables/table.xml"], `totalsRowCount="1"`, `totalsRowCount="0"`, 1)
		case "font-dxf":
			parts["Meta/Styles.style"] = strings.Replace(parts["Meta/Styles.style"], `<dxf>`, `<dxf><font><b val="0"/></font>`, 1)
		case "named-style":
			parts["Tables/table.xml"] = strings.Replace(parts["Tables/table.xml"], `totalsRowCount="1"`, `totalsRowCount="1" totalsRowCellStyle="Custom"`, 1)
		}
		got, err := InspectNativeWorkbookObjectsV1(buildZip(t, parts))
		if err != nil {
			t.Fatal(err)
		}
		palette := got.Tables[0].FillPreview
		if palette == nil {
			t.Fatal("missing qualified base palette")
		}
		want := variant == "default" || variant == "explicit-font" || variant == "nondefault-font"
		if palette.TotalsBold != want {
			t.Fatalf("%s totals=%v", variant, palette.TotalsBold)
		}
		if variant == "explicit-font" || variant == "nondefault-font" {
			for _, id := range palette.HeaderFontStyleIDs {
				if id == 0 {
					t.Fatal("explicit font became totals authority")
				}
			}
		}
	}
}
