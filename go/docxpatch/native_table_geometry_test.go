package docxpatch

import (
	"strings"
	"testing"
)

func TestResolvedTableGeometryCascade(t *testing.T) {
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		for _, tc := range []struct {
			name, base, child, direct string
			want                      *NativeResolvedTableGeometryV1
		}{
			{"documented defaults", "", "", "", &NativeResolvedTableGeometryV1{Layout: "autofit", Alignment: "left", WidthType: "auto", CellMargins: NativeTableCellMarginsV1{LeftTwips: 115, RightTwips: 115}}},
			{"cascade with auto reset", `<w:tblW w:type="dxa" w:w="4000"/><w:tblInd w:type="dxa" w:w="100"/><w:tblCellMar><w:left w:type="dxa" w:w="108"/><w:right w:type="dxa" w:w="108"/></w:tblCellMar>`, `<w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:type="dxa" w:w="40"/></w:tblCellMar>`, `<w:tblW w:type="auto" w:w="0"/><w:tblLayout w:type="autofit"/><w:tblInd w:type="dxa" w:w="200"/>`, &NativeResolvedTableGeometryV1{Layout: "autofit", Alignment: "left", WidthType: "auto", IndentTwips: 200, CellMargins: NativeTableCellMarginsV1{TopTwips: 40, LeftTwips: 108, RightTwips: 108}}},
			{"inherited fixed width", `<w:tblW w:type="dxa" w:w="4000"/><w:tblLayout w:type="fixed"/>`, "", "", &NativeResolvedTableGeometryV1{Layout: "fixed", Alignment: "left", WidthType: "dxa", WidthValue: 4000, CellMargins: NativeTableCellMarginsV1{LeftTwips: 115, RightTwips: 115}}},
			{"unknown geometry", `<w:futureTable/>`, "", "", nil},
			{"missing basedOn value", "", "", "", nil},
			{"empty basedOn value", "", "", "", nil},
			{"conditional ancestor", "", "", "", nil},
			{"missing ancestor", "", "", "", nil},
			{"cycle", "", "", "", nil},
			{"unmodeled table fill", `<w:shd w:val="clear" w:fill="FF0000"/>`, "", "", nil},
			{"malformed table fill", `<w:shd w:fill="oops"/>`, "", "", nil},
			{"duplicate width", `<w:tblW w:type="dxa" w:w="4000"/><w:tblW w:type="dxa" w:w="5000"/>`, "", "", nil},
			{"malformed width", `<w:tblW w:type="dxa" w:w="abc"/>`, "", "", nil},
			{"unsupported alignment", `<w:jc w:val="center"/>`, "", "", nil},
			{"duplicate margin", `<w:tblCellMar><w:left w:type="dxa" w:w="108"/><w:left w:type="dxa" w:w="108"/></w:tblCellMar>`, "", "", nil},
		} {
			t.Run(tc.name+ns, func(t *testing.T) {
				styles := `<w:styles xmlns:w="` + ns + `"><w:style w:type="table" w:styleId="Base"><w:tblPr>` + tc.base + `</w:tblPr></w:style><w:style w:type="table" w:styleId="Child"><w:basedOn w:val="Base"/><w:tblPr>` + tc.child + `</w:tblPr></w:style></w:styles>`
				switch tc.name {
				case "missing basedOn value":
					styles = strings.Replace(styles, `<w:basedOn w:val="Base"/>`, `<w:basedOn/>`, 1)
				case "empty basedOn value":
					styles = strings.Replace(styles, `<w:basedOn w:val="Base"/>`, `<w:basedOn w:val=""/>`, 1)
				case "conditional ancestor":
					styles = strings.Replace(styles, `w:styleId="Base">`, `w:styleId="Base"><w:tblStylePr w:type="firstRow"/>`, 1)
				case "missing ancestor":
					styles = strings.Replace(styles, `<w:basedOn w:val="Base"/>`, `<w:basedOn w:val="Missing"/>`, 1)
				case "cycle":
					styles = strings.Replace(styles, `w:styleId="Base">`, `w:styleId="Base"><w:basedOn w:val="Child"/>`, 1)
				}
				parts := resolvedStylesTestParts(styles)
				parts["word/document.xml"] = `<w:document xmlns:w="` + ns + `"><w:body><w:tbl><w:tblPr><w:tblStyle w:val="Child"/>` + tc.direct + `</w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="4000"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl><w:p/></w:body></w:document>`
				if ns == wordMLStrict {
					for k, v := range parts {
						parts[k] = strings.ReplaceAll(strings.ReplaceAll(v, wordMLTransitional, wordMLStrict), relBaseTransitional, relBaseStrict)
					}
				}
				data := buildNativeDOCX(t, nativeEntries(parts))
				before := string(data)
				layout, err := ResolveNativeDocumentLayoutV1(data)
				if err != nil {
					t.Fatal(err)
				}
				got := layout.Tables[0].Geometry
				if tc.want == nil {
					if got != nil {
						t.Fatalf("unexpected geometry %#v", got)
					}
				} else if got == nil || *got != *tc.want {
					t.Fatalf("geometry %#v, want %#v", got, tc.want)
				}
				if string(data) != before {
					t.Fatal("source changed")
				}
			})
		}
	}
}
