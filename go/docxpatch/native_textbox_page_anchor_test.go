package docxpatch

import (
	"bytes"
	"encoding/json"
	"strings"
	"testing"
)

func nativePageTextboxFixture() string {
	drawing := strings.Replace(nativeGeometryFixture(), `<wp:inline distT="0" distB="0" distL="0" distR="0">`, `<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>1828800</wp:posOffset></wp:positionV>`, 1)
	drawing = strings.Replace(drawing, `<wp:docPr`, `<wp:wrapNone/><wp:docPr`, 1)
	return strings.Replace(drawing, `</wp:inline>`, `</wp:anchor>`, 1)
}

func TestNativeTextboxPageAnchor(t *testing.T) {
	for _, strict := range []bool{false, true} {
		drawing := nativePageTextboxFixture()
		main := nativeMutationMain(`<w:p><w:r>` + drawing + `</w:r></w:p>`)
		if strict {
			main = strings.NewReplacer(wordMLTransitional, wordMLStrict, wordDrawingTransitional, wordDrawingStrict, drawingMLTransitional, drawingMLStrict).Replace(main)
		}
		parts := nativeMutationParts(main)
		if strict {
			for k, v := range parts {
				parts[k] = strings.NewReplacer(wordMLTransitional, wordMLStrict, relBaseTransitional, relBaseStrict).Replace(v)
			}
		}
		source := buildNativeDOCX(t, nativeEntries(parts))
		before := bytes.Clone(source)
		encoded, err := InspectNativePartialSourceV1(source)
		if err != nil {
			t.Fatal(err)
		}
		var result struct {
			Geometry *NativeTextboxGeometryEvidenceV1 `json:"textbox_geometry"`
			Document NativeDocumentV1                 `json:"document"`
		}
		if err := json.Unmarshal(encoded, &result); err != nil {
			t.Fatal(err)
		}
		if result.Geometry == nil || len(result.Geometry.Items) != 1 {
			t.Fatalf("missing geometry: %s", encoded)
		}
		item := result.Geometry.Items[0]
		p := item.PageAnchor
		if item.Geometry == nil || item.Owner.Status != "supported" || p == nil || p.Policy != "page-offset-no-wrap-v1" || p.XEMU != 914400 || p.YEMU != 1828800 {
			t.Fatalf("bad page anchor: %+v", item)
		}
		for _, a := range []NativeSourceAnchorV1{p.SourceAnchor, p.HorizontalAnchor, p.VerticalAnchor} {
			if a.StartByte == nil || a.EndByte == nil || a.XMLSHA256 != nativeSHA([]byte(main)[*a.StartByte:*a.EndByte]) {
				t.Fatalf("source hash mismatch: %+v", a)
			}
		}
		if !bytes.Equal(before, source) || result.Document.Body.Blocks[0].Paragraph.EditPolicy.Mode != "read-only" || len(result.Document.Unsupported) == 0 {
			t.Fatal("source or drawing protections changed")
		}
	}
}

func TestNativeTextboxPageAnchorRefusals(t *testing.T) {
	for _, tc := range []struct{ name, from, to string }{
		{"margin", `relativeFrom="page"`, `relativeFrom="margin"`},
		{"simple position", `simplePos="0"`, `simplePos="1"`},
		{"behind text", `behindDoc="0"`, `behindDoc="1"`},
		{"overlap", `allowOverlap="1"`, `allowOverlap="0"`},
		{"layer", `relativeHeight="0"`, `relativeHeight="1"`},
		{"distance", `distT="0"`, `distT="127"`},
		{"wrapping", `<wp:wrapNone/>`, `<wp:wrapSquare wrapText="bothSides"/>`},
		{"negative", `>914400<`, `>-127<`},
		{"inexact", `>914400<`, `>914401<`},
		{"oversized", `>914400<`, `>127000127<`},
		{"overflow", `>914400<`, `>9999999999999999999999<`},
		{"nested", `>914400<`, `><wp:posOffset>914400</wp:posOffset><`},
		{"attribute", `<wp:posOffset>`, `<wp:posOffset unknown="1">`},
		{"spoofed namespace", `<wp:positionH`, `<wp:positionH xmlns:wp="urn:spoof"`},
		{"duplicate wrap", `<wp:wrapNone/>`, `<wp:wrapNone/><wp:wrapNone/>`},
		{"alignment", `<wp:posOffset>914400</wp:posOffset>`, `<wp:align>left</wp:align>`},
	} {
		t.Run(tc.name, func(t *testing.T) {
			drawing := strings.Replace(nativePageTextboxFixture(), tc.from, tc.to, 1)
			source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(nativeMutationMain(`<w:p><w:r>`+drawing+`</w:r></w:p>`))))
			doc, err := ExtractNativeDocumentV1(source)
			if err != nil {
				t.Fatal(err)
			}
			out, err := inspectNativeTextboxGeometry(source, doc)
			if err != nil {
				t.Fatal(err)
			}
			if out != nil {
				for _, item := range out.Items {
					if item.Geometry != nil || item.PageAnchor != nil || item.Owner.Status == "supported" {
						t.Fatalf("unsupported placement admitted: %+v", item)
					}
				}
			}
		})
	}
}
