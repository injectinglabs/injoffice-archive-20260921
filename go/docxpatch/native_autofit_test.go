package docxpatch

import (
	"strings"
	"testing"
)

func TestNativeAutofitSourceProjection(t *testing.T) {
	for _, width := range []string{`<w:tblW w:w="2400" w:type="dxa"/>`, `<w:tblW w:w="0" w:type="auto"/>`} {
		parts := transitionalNativeParts()
		main := strings.Replace(parts["Custom/Main.XML"], `<w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>`, `<w:tblPr>`+width+`<w:tblLayout w:type="autofit"/></w:tblPr>`, 1)
		main = strings.Replace(main, `<w:tcW w:w="2400"/>`, `<w:tcW w:w="0" w:type="auto"/>`, 1)
		parts["Custom/Main.XML"] = main
		document, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
		if err != nil {
			t.Fatal(err)
		}
		table := document.Body.Blocks[1].Table
		if table == nil || table.Layout == nil || *table.Layout != "autofit" || table.Rows[0].Cells[0].WidthTwips != nil {
			t.Fatalf("autofit source projection missing: %#v", table)
		}
		if strings.Contains(width, `type="auto"`) && table.WidthTwips != nil {
			t.Fatal("auto width invented an absolute extent")
		}
	}
}
