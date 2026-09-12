package pptxpatch

import (
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"strings"
	"testing"
)

func TestInspectionBudgetBeforeGrammarAndConcatenation(t *testing.T) {
	for name, change := range map[string]func(string) string{
		"empty-paragraphs": func(s string) string { return strings.Replace(s, `<a:p>`, strings.Repeat(`<a:p/>`, 256)+`<a:p>`, 1) },
		"empty-runs": func(s string) string {
			return strings.Replace(s, `<a:r>`, strings.Repeat(`<a:r><a:rPr/><a:t/></a:r>`, 256)+`<a:r>`, 1)
		},
		"text-before-concatenation": func(s string) string {
			return strings.Replace(s, `Readable &amp; safe`, strings.Repeat("😀", 32769), 1)
		},
		"unknown-empty-nodes": func(s string) string {
			return strings.Replace(s, `<a:tblPr firstRow`, strings.Repeat(`<a:unknown/>`, 20001)+`<a:tblPr firstRow`, 1)
		},
	} {
		t.Run(name, func(t *testing.T) {
			_, err := InspectNativePPTXTables(nativeTableFixture(t, false, change(inspectionTestFrame())))
			var budget *nativeInspectionBudgetError
			if !errors.As(err, &budget) {
				t.Fatalf("budget failure was swallowed as omission: %v", err)
			}
		})
	}
	d := nativeExtractDialect{drawing: nsDrawingTransitional}
	for _, test := range []struct {
		name   string
		budget nativeInspectionBudget
		node   string
	}{
		{"nodes", nativeInspectionBudget{nodes: 20000}, "p"},
		{"cells", nativeInspectionBudget{cells: 4096}, "tc"},
		{"paragraphs", nativeInspectionBudget{paragraphs: 4096}, "p"},
		{"runs", nativeInspectionBudget{runs: 16384}, "r"},
	} {
		t.Run(test.name, func(t *testing.T) {
			if test.budget.scan(&nativeXMLNode{Name: xml.Name{Space: d.drawing, Local: test.node}}, d) == nil {
				t.Fatal("shared cumulative budget not enforced")
			}
		})
	}
}

func TestInspectionSkipsNonTableBeforeTableNonVisualQualification(t *testing.T) {
	d := nativeExtractDialect{presentation: nsPresentationTransitional, drawing: nsDrawingTransitional}
	source := strings.Replace(inspectionTestFrame(), nativeDrawingTableURI, "urn:chart", 1)
	source = strings.Replace(source, `noGrp="1"`, `noMove="1"`, 1)
	root, err := parseNativeXML([]byte(`<root xmlns:p="`+d.presentation+`" xmlns:a="`+d.drawing+`">`+source+`</root>`), "test.xml")
	if err != nil {
		t.Fatal(err)
	}
	table, _, err := inspectNativeTableSource(root.Children[0], d, &nativeInspectionBudget{})
	if err != nil || table != nil {
		t.Fatalf("non-table frame mislabeled: %v %#v", err, table)
	}
	if _, err := InspectNativePPTXTables(nil); err == nil {
		t.Fatal("empty package accepted")
	}
}

func TestInspectionSlideOmissionsShareTableBudget(t *testing.T) {
	for _, marker := range []string{`<p:sld `, `<p:cSld>`} {
		t.Run(marker, func(t *testing.T) {
			var frames strings.Builder
			for i := 0; i < 256; i++ {
				frames.WriteString(strings.Replace(inspectionTestFrame(), `id="3"`, fmt.Sprintf(`id="%d"`, i+3), 1))
			}
			data := nativeExtractFixture(t, nativeExtractFixtureOptions{secondSlide: true, mutate: func(parts map[string]string) {
				parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, frames.String()+`</p:spTree>`, 1)
				source := parts["relocated/slides/slide-b.xml"]
				replacement := strings.TrimSuffix(marker, ">") + ` show="0">`
				if marker == `<p:sld ` {
					replacement = `<p:sld show="0" `
				}
				source = strings.Replace(source, marker, replacement, 1)
				parts["relocated/slides/slide-b.xml"] = source
			}})
			_, err := InspectNativePPTXTables(data)
			var budget *nativeInspectionBudgetError
			if !errors.As(err, &budget) {
				t.Fatalf("slide omissions bypassed shared budget: %v", err)
			}
		})
	}
}

func TestInspectionObservedZeroRootScaffoldAndTransformedRootRefusal(t *testing.T) {
	scaffold := `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>`
	for _, test := range []struct {
		name, source string
		accepted     bool
	}{
		{"observed-zero", scaffold, true},
		{"nonzero", strings.Replace(scaffold, `x="0"`, `x="1"`, 1), false},
		{"unknown-rotation", strings.Replace(scaffold, `<a:xfrm>`, `<a:xfrm rot="1">`, 1), false},
		{"missing-child", strings.Replace(scaffold, `<a:chExt cx="0" cy="0"/>`, ``, 1), false},
	} {
		t.Run(test.name, func(t *testing.T) {
			data := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
				s := strings.Replace(parts["relocated/slides/slide-a.xml"], `<p:grpSpPr/>`, test.source, 1)
				parts["relocated/slides/slide-a.xml"] = strings.Replace(s, `</p:spTree>`, inspectionTestFrame()+`</p:spTree>`, 1)
			}})
			result, err := InspectNativePPTXTables(data)
			if err != nil {
				t.Fatal(err)
			}
			if test.accepted && (len(result.Tables) != 1 || len(result.Omissions) != 0) {
				t.Fatal("observed scaffold refused")
			}
			if !test.accepted && (len(result.Tables) != 0 || len(result.Omissions) != 1) {
				t.Fatal("unknown root geometry inspected")
			}
		})
	}
}

// Original synthetic source, not copied from an external fixture. It models
// the independently observed profile without importing upstream bytes.
func inspectionTestFrame() string {
	cell := `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>Readable &amp; safe</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></a:txBody><a:tcPr><a:lnL><a:noFill/></a:lnL><a:lnR><a:noFill/></a:lnR><a:lnT><a:noFill/></a:lnT><a:lnB><a:noFill/></a:lnB></a:tcPr></a:tc>`
	frame := nativeExactTableGraphicFrameXML(3, "Inspection", []int64{500000}, []int64{300000}, [][]string{{cell}}, "")
	frame = strings.Replace(frame, `<a:tblPr/>`, `<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{01234567-89AB-CDEF-0123-456789ABCDEF}</a:tableStyleId></a:tblPr>`, 1)
	frame = strings.Replace(frame, `<p:cNvGraphicFramePr/>`, `<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>`, 1)
	return frame
}

func TestInspectNativeTablesPreservesSourceAndStrictRefusal(t *testing.T) {
	for _, strict := range []bool{false, true} {
		data := nativeTableFixture(t, strict, inspectionTestFrame())
		before := append([]byte(nil), data...)
		result, err := InspectNativePPTXTables(data)
		if err != nil {
			t.Fatal(err)
		}
		if len(result.Tables) != 1 || len(result.Omissions) != 0 {
			t.Fatalf("unexpected inspection: %#v", result)
		}
		table := result.Tables[0]
		if result.SourceRevision != "rev-"+nativeSHA256(data) || result.PackageSHA256 != nativeSHA256(data) || table.ObjectID != "cNvPr-3" || table.SourceSHA256 == "" || table.PartSHA256 == "" || len(table.Cells) != 1 || table.Cells[0].Paragraphs[0] != "Readable & safe" || table.Cells[0].Rect.Width != 500000 || table.Rect.X != 300000 || len(table.Warnings) != 3 {
			t.Fatalf("source projection mismatch: %#v", result)
		}
		deck, err := ExtractNativePPTX(data, nativeTestExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		for _, element := range deck.Slides[0].Elements {
			if element.Kind == NativeElementKindTable {
				t.Fatal("inspection granted strict table authority")
			}
		}
		if !bytes.Equal(data, before) {
			t.Fatal("source bytes changed")
		}
	}
}

func TestInspectNativeTablesRejectsAmbiguousHiddenAndUnknownSource(t *testing.T) {
	for name, change := range map[string]func(string) string{
		"hidden": func(s string) string {
			return strings.Replace(s, `name="Inspection"`, `name="Inspection" hidden="1"`, 1)
		},
		"duplicate-body": func(s string) string { return strings.Replace(s, `<a:bodyPr/>`, `<a:bodyPr/><a:bodyPr/>`, 1) },
		"missing-body":   func(s string) string { return strings.Replace(s, `<a:bodyPr/>`, ``, 1) },
		"foreign-text": func(s string) string {
			return strings.Replace(s, `<a:t>Readable &amp; safe</a:t>`, `<evil:t xmlns:evil="urn:evil">Hidden</evil:t>`, 1)
		},
		"unknown-format": func(s string) string {
			return strings.Replace(s, `<a:rPr lang="en-US" dirty="0"/>`, `<a:rPr lang="en-US" dirty="0" secret="1"/>`, 1)
		},
		"hyperlink": func(s string) string {
			return strings.Replace(s, `<a:rPr lang="en-US" dirty="0"/>`, `<a:rPr lang="en-US"><a:hlinkClick r:id="rId1"/></a:rPr>`, 1)
		},
		"merge":          func(s string) string { return strings.Replace(s, `<a:tc>`, `<a:tc gridSpan="2">`, 1) },
		"wrong-height":   func(s string) string { return strings.Replace(s, `<a:tr h="300000">`, `<a:tr h="300001">`, 1) },
		"oversized-text": func(s string) string { return strings.Replace(s, `Readable &amp; safe`, strings.Repeat("x", 65537), 1) },
		"unknown-lock":   func(s string) string { return strings.Replace(s, `noGrp="1"`, `noMove="1"`, 1) },
		"malformed-xml":  func(s string) string { return strings.Replace(s, `</a:tc>`, `</a:wrong>`, 1) },
	} {
		t.Run(name, func(t *testing.T) {
			result, err := InspectNativePPTXTables(nativeTableFixture(t, false, change(inspectionTestFrame())))
			if err == nil && len(result.Tables) != 0 {
				t.Fatalf("unsafe source became readable: %#v", result)
			}
		})
	}
}

func TestInspectNativeTablesKnownModificationMetadataAndBudget(t *testing.T) {
	frame := strings.Replace(inspectionTestFrame(), `<p:nvPr/>`, `<p:nvPr><p:extLst><p:ext uri="{D42A27DB-BD31-4B8C-83A1-F6EECF244321}"><p14:modId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="123"/></p:ext></p:extLst></p:nvPr>`, 1)
	for _, test := range []struct {
		name, source string
		want         bool
	}{
		{"known", frame, true},
		{"foreign-modification", strings.Replace(frame, `http://schemas.microsoft.com/office/powerpoint/2010/main`, `urn:foreign`, 1), false},
		{"unknown-extension", strings.Replace(frame, `{D42A27DB-BD31-4B8C-83A1-F6EECF244321}`, `{01234567-89AB-CDEF-0123-456789ABCDEF}`, 1), false},
		{"oversized-modification", strings.Replace(frame, `val="123"`, `val="4294967296"`, 1), false},
	} {
		t.Run(test.name, func(t *testing.T) {
			r, e := InspectNativePPTXTables(nativeTableFixture(t, false, test.source))
			if test.want && (e != nil || len(r.Tables) != 1) {
				t.Fatalf("known evidence refused: %#v %v", r, e)
			}
			if !test.want && e == nil && len(r.Tables) > 0 {
				t.Fatal("unknown extension accepted")
			}
		})
	}
	var many strings.Builder
	for i := 0; i < 257; i++ {
		many.WriteString(strings.Replace(inspectionTestFrame(), `id="3"`, `id="`+fmt.Sprint(i+3)+`"`, 1))
	}
	if _, err := InspectNativePPTXTables(nativeTableFixture(t, false, many.String())); err == nil {
		t.Fatal("table count budget not enforced")
	}
}
