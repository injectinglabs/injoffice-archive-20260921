package docxpatch

import (
	"fmt"
	"strings"
	"testing"
)

func TestNativeBoundedInlineImageTransforms(t *testing.T) {
	for _, degrees := range []int64{0, 90, 180, 270} {
		for _, flipH := range []bool{false, true} {
			for _, flipV := range []bool{false, true} {
				parts := transitionalNativeParts()
				transform := fmt.Sprintf(`<a:xfrm rot="%d" flipH="%t" flipV="%t"/>`, degrees*60000, flipH, flipV)
				if degrees == 90 || degrees == 270 {
					transform = fmt.Sprintf(`<a:xfrm rot="%d" flipH="%t" flipV="%t"><a:off x="0" y="0"/><a:ext cx="457200" cy="914400"/></a:xfrm>`, degrees*60000, flipH, flipV)
				}
				parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `<a:xfrm/>`, transform, 1)
				doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
				if err != nil {
					t.Fatal(err)
				}
				var drawing *NativeDrawingV1
				for _, run := range doc.Body.Blocks[0].Paragraph.Runs {
					if run.Drawing != nil {
						drawing = run.Drawing
					}
				}
				if drawing == nil || drawing.RotationDegrees == nil || *drawing.RotationDegrees != degrees || drawing.FlipHorizontal == nil || *drawing.FlipHorizontal != flipH || drawing.FlipVertical == nil || *drawing.FlipVertical != flipV {
					t.Fatalf("wrong transform projection: %#v", drawing)
				}
				if *drawing.WidthEMU != 914400 || *drawing.HeightEMU != 457200 {
					t.Fatal("orientation changed source extents")
				}
			}
		}
	}
	for _, attrs := range []string{`rot="5400000"`, `rot="16200000"`, `rot="-10800000"`, `rot="21600000"`, `rot="NaN"`, `flipH="yes"`, `flipV="2"`, `rot="10800000" bogus="1"`, `rot="0" rot="10800000"`, `flipH="true" flipH="false"`} {
		parts := transitionalNativeParts()
		parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `<a:xfrm/>`, `<a:xfrm `+attrs+`/>`, 1)
		doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
		if err != nil {
			if strings.Contains(err.Error(), "duplicate attribute") && (strings.Count(attrs, "rot=") > 1 || strings.Count(attrs, "flipH=") > 1) {
				continue
			}
			t.Fatal(err)
		}
		for _, run := range doc.Body.Blocks[0].Paragraph.Runs {
			if run.Drawing != nil {
				t.Fatalf("unsupported transform accepted: %s", attrs)
			}
		}
	}
}

func TestNativeBoundedImageSourceCrop(t *testing.T) {
	for _, attrs := range []string{`l="50000"`, `l="12345" t="2500" r="5000" b="100"`, `l="99000"`, ``} {
		parts := transitionalNativeParts()
		parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `<pic:blipFill><a:blip`, `<pic:blipFill><a:srcRect `+attrs+`/><a:blip`, 1)
		doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
		if err != nil {
			t.Fatal(err)
		}
		found := false
		for _, run := range doc.Body.Blocks[0].Paragraph.Runs {
			if run.Drawing != nil {
				crop := run.Drawing.SourceCrop
				if crop == nil || crop.Left == nil || crop.Right == nil || crop.Top == nil || crop.Bottom == nil {
					t.Fatal("missing exact crop projection")
				}
				found = true
			}
		}
		if !found {
			t.Fatalf("bounded crop refused: %s", attrs)
		}
	}
	for _, attrs := range []string{`l="-1"`, `l="99001"`, `l="60000" r="40000"`, `t="50000" b="49901"`, `l="NaN"`, `bogus="1"`, `l="1%"`} {
		parts := transitionalNativeParts()
		parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `<pic:blipFill><a:blip`, `<pic:blipFill><a:srcRect `+attrs+`/><a:blip`, 1)
		doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
		if err != nil {
			t.Fatal(err)
		}
		for _, run := range doc.Body.Blocks[0].Paragraph.Runs {
			if run.Drawing != nil {
				t.Fatalf("unsupported crop accepted: %s", attrs)
			}
		}
	}
}
