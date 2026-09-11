package docxpatch

import "testing"

func TestNativeScriptAlignmentStrictValueAndReset(t *testing.T) {
	const ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	for _, value := range []string{"baseline", "subscript", "superscript"} {
		node, err := parseNativeXML("word/document.xml", []byte(`<w:vertAlign xmlns:w="`+ns+`" w:val="`+value+`"/>`))
		if err != nil {
			t.Fatal(err)
		}
		actual, ok := nativeVerticalAlignmentValue(node, ns)
		if !ok || actual != value {
			t.Fatalf("exact value refused: %s", value)
		}
	}
	for _, attributes := range []string{`w:val="other"`, `w:val="superscript" w:val="subscript"`, `w:val="superscript" extra="true"`, ""} {
		node, err := parseNativeXML("word/document.xml", []byte(`<w:vertAlign xmlns:w="`+ns+`" `+attributes+`/>`))
		if err != nil {
			continue
		}
		if _, ok := nativeVerticalAlignmentValue(node, ns); ok {
			t.Fatalf("ambiguous value admitted: %s", attributes)
		}
	}
	inherited := nativeRunProperties{verticalAlignment: nativeString("superscript")}
	applyNativeRunProperties(&inherited, nativeRunProperties{verticalAlignment: nativeString("baseline")}, true)
	if inherited.verticalAlignment == nil || *inherited.verticalAlignment != "baseline" {
		t.Fatal("baseline did not reset inherited superscript")
	}
}
