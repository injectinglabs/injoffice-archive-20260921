package pptxpatch

import (
	"bytes"
	"fmt"
	"strings"
	"testing"
)

func paintTestStyle() string {
	borders := ""
	for _, name := range []string{"left", "right", "top", "bottom", "insideH", "insideV"} {
		borders += "<a:" + name + "><a:ln><a:noFill/></a:ln></a:" + name + ">"
	}
	return `<a:tblStyle styleId="{01234567-89AB-CDEF-0123-456789ABCDEF}" styleName="Synthetic unstyled"><a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"><a:scrgbClr r="0" g="0" b="0"/></a:fontRef><a:schemeClr val="tx1"/></a:tcTxStyle><a:tcStyle><a:tcBdr>` + borders + `</a:tcBdr><a:fill><a:noFill/></a:fill></a:tcStyle></a:wholeTbl></a:tblStyle>`
}
func paintTestEdge(name string) string {
	return `<a:` + name + ` w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:prstDash val="solid"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:` + name + `>`
}
func paintTestFixture(t *testing.T, strict, solid bool, change func(map[string]string)) []byte {
	t.Helper()
	return nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, extraParts: []nativeExtractZipPart{{name: "relocated/styles/table.xml", data: "placeholder"}}, mutate: func(parts map[string]string) {
		ns := nsDrawingTransitional
		relNS := nsOfficeRelsTransitional
		if strict {
			ns = nsDrawingStrict
			relNS = nsOfficeRelsStrict
		}
		frame := inspectionTestFrame()
		if solid {
			for _, name := range []string{"lnL", "lnR", "lnT", "lnB"} {
				frame = strings.Replace(frame, "<a:"+name+"><a:noFill/></a:"+name+">", paintTestEdge(name), 1)
			}
		}
		parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, frame+`</p:spTree>`, 1)
		parts["relocated/styles/table.xml"] = `<a:tblStyleLst xmlns:a="` + ns + `">` + paintTestStyle() + `</a:tblStyleLst>`
		parts["relocated/_rels/deck.xml.rels"] = strings.Replace(parts["relocated/_rels/deck.xml.rels"], `</Relationships>`, `<Relationship Id="rIdStyles" Type="`+relNS+`/tableStyles" Target="styles/table.xml"/></Relationships>`, 1)
		parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/relocated/styles/table.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/></Types>`, 1)
		if change != nil {
			change(parts)
		}
	}})
}
func TestInspectedSourceTablePaintKeepsSourceAndStrictDefaults(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, solid := range []bool{false, true} {
			t.Run(fmt.Sprintf("strict-%t-solid-%t", strict, solid), func(t *testing.T) {
				input := paintTestFixture(t, strict, solid, nil)
				before := append([]byte(nil), input...)
				result, err := InspectNativePPTXTables(input)
				if err != nil {
					t.Fatal(err)
				}
				if len(result.Tables) != 1 || result.Tables[0].Paint == nil {
					t.Fatal("source paint not qualified")
				}
				paint := result.Tables[0].Paint
				if paint.Fill != "none" || len(paint.Sources) != 4 || paint.Sources[0].PartName != "relocated/styles/table.xml" || paint.Sources[0].SHA256 != nativeSHA256(chartZipEntry(t, input, "relocated/styles/table.xml")) {
					t.Fatalf("paint provenance missing: %+v", paint)
				}
				if solid && (paint.Border == nil || paint.Border.Color != "112233" || paint.Border.WidthEMU != 12700) || !solid && paint.Border != nil {
					t.Fatal("border altered")
				}
				deck, err := ExtractNativePPTX(input, nativeTestExtractOptions())
				if err != nil {
					t.Fatal(err)
				}
				for _, element := range deck.Slides[0].Elements {
					if element.Table != nil {
						t.Fatal("paint inventory promoted strict native table")
					}
				}
				if !bytes.Equal(input, before) {
					t.Fatal("source mutation")
				}
			})
		}
	}
}
func TestInspectedTablePaintRequiresExactStyleAndBorders(t *testing.T) {
	changes := map[string]func(map[string]string){
		"guid-alone": func(p map[string]string) {
			p["relocated/styles/table.xml"] = strings.Replace(p["relocated/styles/table.xml"], `<a:noFill/>`, `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>`, 1)
		},
		"duplicate-style": func(p map[string]string) {
			p["relocated/styles/table.xml"] = strings.Replace(p["relocated/styles/table.xml"], `</a:tblStyleLst>`, paintTestStyle()+`</a:tblStyleLst>`, 1)
		},
		"duplicate-style-case": func(p map[string]string) {
			lower := strings.Replace(paintTestStyle(), "{01234567-89AB-CDEF-0123-456789ABCDEF}", "{01234567-89ab-cdef-0123-456789abcdef}", 1)
			p["relocated/styles/table.xml"] = strings.Replace(p["relocated/styles/table.xml"], `</a:tblStyleLst>`, lower+`</a:tblStyleLst>`, 1)
		},
		"conditional-style": func(p map[string]string) {
			p["relocated/styles/table.xml"] = strings.Replace(p["relocated/styles/table.xml"], `</a:tblStyle>`, `<a:firstRow/></a:tblStyle>`, 1)
		},
		"style-fill-attrs": func(p map[string]string) {
			p["relocated/styles/table.xml"] = strings.Replace(p["relocated/styles/table.xml"], `<a:fill>`, `<a:fill unknown="1">`, 1)
		},
		"style-text": func(p map[string]string) {
			p["relocated/styles/table.xml"] = strings.Replace(p["relocated/styles/table.xml"], `<a:wholeTbl>`, `<a:wholeTbl>unknown`, 1)
		},
		"dash": func(p map[string]string) {
			p["relocated/slides/slide-a.xml"] = strings.Replace(p["relocated/slides/slide-a.xml"], `<a:prstDash val="solid"/>`, `<a:prstDash val="dash"/>`, 1)
		},
		"unequal-edges": func(p map[string]string) {
			p["relocated/slides/slide-a.xml"] = strings.Replace(p["relocated/slides/slide-a.xml"], `<a:lnR w="12700"`, `<a:lnR w="25400"`, 1)
		},
		"inside-aligned": func(p map[string]string) {
			p["relocated/slides/slide-a.xml"] = strings.Replace(p["relocated/slides/slide-a.xml"], `cap="flat" cmpd="sng" algn="ctr"`, `cap="flat" cmpd="sng" algn="in"`, 1)
		},
		"too-wide": func(p map[string]string) {
			p["relocated/slides/slide-a.xml"] = strings.ReplaceAll(p["relocated/slides/slide-a.xml"], `w="12700"`, `w="127001"`)
		},
		"missing-edge": func(p map[string]string) {
			p["relocated/slides/slide-a.xml"] = strings.Replace(p["relocated/slides/slide-a.xml"], paintTestEdge("lnB"), ``, 1)
		},
		"external-style": func(p map[string]string) {
			p["relocated/_rels/deck.xml.rels"] = strings.Replace(p["relocated/_rels/deck.xml.rels"], `Target="styles/table.xml"`, `Target="https://example.test/styles.xml" TargetMode="External"`, 1)
		},
	}
	for name, change := range changes {
		t.Run(name, func(t *testing.T) {
			input := paintTestFixture(t, false, true, change)
			result, err := InspectNativePPTXTables(input)
			if err != nil {
				t.Fatal(err)
			}
			if len(result.Tables) != 1 || result.Tables[0].Paint != nil {
				t.Fatal("unqualified paint admitted or readable text lost")
			}
		})
	}
}
