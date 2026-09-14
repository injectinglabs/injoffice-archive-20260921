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

func TestResolvedTableGeometryTableGridAutoWidth(t *testing.T) {
	borders := `<w:tblBorders>`
	for _, edge := range []string{"top", "left", "bottom", "right", "insideH", "insideV"} {
		borders += `<w:` + edge + ` w:val="single" w:sz="4" w:space="0" w:color="auto"/>`
	}
	borders += `</w:tblBorders>`
	margins := `<w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>`
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		for _, tc := range []struct {
			name, extra, look string
			wantGeometry      bool
		}{
			{name: "table-grid auto look", look: `<w:tblLook w:val="04A0"/>`, wantGeometry: true},
			{name: "unknown geometry child", extra: `<w:futureTable/>`},
			{name: "malformed look", look: `<w:tblLook w:val="FFFF"/>`},
		} {
			t.Run(tc.name+ns, func(t *testing.T) {
				styles := `<w:styles xmlns:w="` + ns + `"><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr>` + margins + `</w:tblPr></w:style><w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="39"/><w:rsid w:val="005E46CA"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr>` + margins + borders + tc.extra + `</w:tblPr></w:style></w:styles>`
				parts := resolvedStylesTestParts(styles)
				parts["word/document.xml"] = `<w:document xmlns:w="` + ns + `"><w:body><w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/>` + tc.look + `</w:tblPr><w:tblGrid><w:gridCol w:w="1510"/><w:gridCol w:w="1511"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="1510" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>1</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="1511" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl><w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/></w:sectPr></w:body></w:document>`
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
				if !tc.wantGeometry {
					if got != nil {
						t.Fatalf("unexpected geometry %#v", got)
					}
				} else {
					want := NativeResolvedTableGeometryV1{Layout: "autofit", Alignment: "left", WidthType: "auto", CellMargins: NativeTableCellMarginsV1{LeftTwips: 108, RightTwips: 108}}
					if got == nil || *got != want {
						t.Fatalf("geometry %#v, want %#v", got, want)
					}
					if !hasResolutionDiagnostic(layout, "TABLE_STYLE_EFFECTS_PRESERVED") {
						t.Fatalf("missing auto-border diagnostic: %#v", layout.Diagnostics)
					}
					if layout.Tables[0].AutomaticBorderPreview == nil {
						t.Fatal("missing automatic-border evidence")
					}
				}
				if string(data) != before {
					t.Fatal("source changed")
				}
			})
		}
	}
}
