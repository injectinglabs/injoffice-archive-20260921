package pptxpatch

import "testing"

func TestNativeSourceAffineAttributes(t *testing.T) {
	for _, test := range []struct {
		attrs string
		want  nativeSourceAffine
	}{
		{``, nativeSourceAffine{}},
		{`rot="2700000" flipH="true"`, nativeSourceAffine{2700000, true, false}},
		{`rot="-5400000" flipV="1"`, nativeSourceAffine{16200000, false, true}},
		{`rot="21600000" flipH="0" flipV="false"`, nativeSourceAffine{}},
		{`rot="2147483647"`, nativeSourceAffine{9083647, false, false}},
	} {
		node, err := parseNativeXML([]byte(`<a:xfrm xmlns:a="`+nativeGeometryTestNS+`" `+test.attrs+`/>`), "affine.xml")
		if err != nil {
			t.Fatal(err)
		}
		got, err := parseNativeSourceAffine(node)
		if err != nil || got != test.want {
			t.Errorf("%s: %+v %v", test.attrs, got, err)
		}
	}
}
func TestNativeSourceAffineAttributeRefusals(t *testing.T) {
	for _, attrs := range []string{`rot="2147483648"`, `rot="1.5"`, `rot="-0"`, `rot="01"`, `flipH="yes"`, `flipV="TRUE"`, `angle="0"`, `rot="0" rot="1"`, `a:rot="1"`} {
		node, err := parseNativeXML([]byte(`<a:xfrm xmlns:a="`+nativeGeometryTestNS+`" `+attrs+`/>`), "affine.xml")
		if err == nil {
			_, err = parseNativeSourceAffine(node)
		}
		if err == nil {
			t.Errorf("accepted %s", attrs)
		}
	}
}
