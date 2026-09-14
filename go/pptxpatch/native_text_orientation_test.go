package pptxpatch

import (
	"encoding/xml"
	"fmt"
	"testing"
)

func TestNativeTextOrientationSignedAnglesAndFlags(t *testing.T) {
	for _, angle := range []int64{0, 5400000, -5400000, 21600000, -21600000, -2147483648, 2147483647} {
		for _, flag := range []string{"0", "1", "false", "true"} {
			node, err := parseNativeXML([]byte(fmt.Sprintf(`<a:bodyPr xmlns:a="%s" rot="%d" upright="%s" lIns="10"/>`, nsDrawingTransitional, angle, flag)), "body.xml")
			if err != nil {
				t.Fatal(err)
			}
			got, err := parseNativeTextOrientation(node)
			if err != nil {
				t.Fatal(err)
			}
			if got.Rotation != angle || !got.HasRotation || !got.HasUpright || got.Upright != (flag == "1" || flag == "true") || got.normalizedRotation() != (angle%21600000+21600000)%21600000 {
				t.Fatal("raw signed orientation was not preserved")
			}
		}
	}
	empty, err := parseNativeTextOrientation(&nativeXMLNode{})
	if err != nil || empty != (nativeTextOrientation{}) {
		t.Fatal("absent orientation did not retain default/presence distinction")
	}
}
func TestNativeTextOrientationRejectsMalformedAttributes(t *testing.T) {
	for _, value := range []string{"-0", "+1", "01", "1.5", "NaN", "2147483648", "-2147483649", " 1"} {
		if _, err := parseNativeTextOrientation(&nativeXMLNode{Attrs: []xml.Attr{{Name: xml.Name{Local: "rot"}, Value: value}}}); err == nil {
			t.Fatalf("rotation %q accepted", value)
		}
	}
	for _, value := range []string{"TRUE", "yes", "2", "", "-0"} {
		if _, err := parseNativeTextOrientation(&nativeXMLNode{Attrs: []xml.Attr{{Name: xml.Name{Local: "upright"}, Value: value}}}); err == nil {
			t.Fatalf("upright %q accepted", value)
		}
	}
	for _, name := range []string{"rot", "upright"} {
		duplicate := []xml.Attr{{Name: xml.Name{Local: name}, Value: "0"}, {Name: xml.Name{Local: name}, Value: "1"}}
		if _, err := parseNativeTextOrientation(&nativeXMLNode{Attrs: duplicate}); err == nil {
			t.Fatal("duplicate orientation accepted")
		}
		if _, err := parseNativeTextOrientation(&nativeXMLNode{Attrs: []xml.Attr{{Name: xml.Name{Space: "urn:foreign", Local: name}, Value: "0"}}}); err == nil {
			t.Fatal("foreign orientation accepted")
		}
	}
	if _, err := parseNativeTextOrientation(nil); err == nil {
		t.Fatal("missing body accepted")
	}
}
