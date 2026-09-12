package xlsxpatch

import (
	"strings"
	"testing"
)

func TestTableBordersKeepExplicitAndDifferentialOverrides(t *testing.T) {
	for _, variant := range []string{"default", "explicit-none", "totals-dxf"} {
		parts := nativeWorkbookFixture(false)
		styleXML := parts["Meta/Styles.style"]
		if variant == "explicit-none" {
			styleXML = strings.Replace(styleXML, `xfId="0"/>`, `xfId="0" applyBorder="1"/>`, 1)
		}
		registry, err := newStyleRegistry([]byte(styleXML))
		if err != nil {
			t.Fatal(err)
		}
		styles, err := parsePreviewXML([]byte(styleXML))
		if err != nil {
			t.Fatal(err)
		}
		tableXML := `<table xmlns="` + spreadsheetMLTransitional + `"/>`
		if variant == "totals-dxf" {
			tableXML = `<table xmlns="` + spreadsheetMLTransitional + `" totalsRowDxfId="0"/>`
		}
		root, _ := parsePreviewXML([]byte(tableXML))
		table := NativeTablePreviewV1{}
		qualifyNativeTableBorders(root, &table, registry, styles, "#156082")
		if variant == "totals-dxf" {
			if table.BorderPreview != nil {
				t.Fatal("unknown DXF border painted")
			}
			continue
		}
		if table.BorderPreview == nil || table.BorderPreview.Color != "#44B3E1" || table.BorderPreview.WidthPoints != 1 || table.BorderPreview.TotalsWidthPoints != 3 {
			t.Fatalf("wrong measured border: %+v", table.BorderPreview)
		}
		if (len(table.BorderPreview.StyleIDs) > 0) != (variant == "default") {
			t.Fatalf("direct border override lost: %+v", table.BorderPreview)
		}
	}
}
