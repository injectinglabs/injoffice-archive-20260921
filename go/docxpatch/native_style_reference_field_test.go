package docxpatch

import (
	"strings"
	"testing"
)

// STYLEREF is refused for its own reason, not for the page-field reason. The
// generic sentence names PAGE and NUMPAGES, which reads as though a page-field
// fix would admit the field; STYLEREF's result is instead pagination-dependent
// arbitrary paragraph text named by a localized style name the package need not
// carry, so nothing follows from the package alone.
func TestNativeStyleReferenceFieldRefusalNamesItsOwnReason(t *testing.T) {
	for _, tc := range []struct {
		name        string
		instruction string
		styleRef    bool
	}{
		{"unquoted", ` STYLEREF  Untertitel  \* MERGEFORMAT `, true},
		{"quoted", ` STYLEREF  &quot;Überschrift 1&quot;  \* MERGEFORMAT `, true},
		{"lowercase", ` styleref  Untertitel `, true},
		{"other-field", ` DATE \@ &quot;dd.MM.yyyy&quot; `, false},
		{"unmodeled-page-switch", ` PAGE \* ArabicDash `, false},
	} {
		t.Run(tc.name, func(t *testing.T) {
			parts := transitionalNativeParts()
			parts["Custom/Stories/HeaderA.XML"] = `<w:hdr xmlns:w="` + testW + `"><w:p><w:fldSimple w:instr="` + tc.instruction + `"><w:r><w:t>cached</w:t></w:r></w:fldSimple></w:p></w:hdr>`
			doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			messages := []string{}
			for _, entry := range doc.Unsupported {
				if entry.Code == "FIELD_SEMANTICS" && entry.Anchor.PartName == "Custom/Stories/HeaderA.XML" {
					messages = append(messages, entry.Message)
				}
			}
			if len(messages) != 1 {
				t.Fatalf("expected one FIELD_SEMANTICS record, got %#v", messages)
			}
			if tc.styleRef {
				if !strings.HasPrefix(messages[0], "STYLEREF ") {
					t.Fatalf("STYLEREF refused with the page-field reason: %q", messages[0])
				}
				if strings.Contains(messages[0], "NUMPAGES") {
					t.Fatalf("STYLEREF refusal still points at page fields: %q", messages[0])
				}
				return
			}
			if messages[0] != nativeUnmodeledPageFieldMessage {
				t.Fatalf("non-STYLEREF field lost the page-field reason: %q", messages[0])
			}
		})
	}
}

// The keyword scan stops at the first token, ignores case and reports none for
// a token no ECMA-376 field keyword could be.
func TestNativeFieldKeyword(t *testing.T) {
	for _, tc := range []struct{ instruction, keyword string }{
		{` STYLEREF  Untertitel `, "STYLEREF"},
		{"\r\n\tstyleref x", "STYLEREF"},
		{"PAGE", "PAGE"},
		{"", ""},
		{"   ", ""},
		{strings.Repeat("A", 65), ""},
		{strings.Repeat("A", 64), strings.Repeat("A", 64)},
		{"STYLEREFX y", "STYLEREFX"},
	} {
		if got := nativeFieldKeyword(tc.instruction); got != tc.keyword {
			t.Fatalf("nativeFieldKeyword(%q) = %q, want %q", tc.instruction, got, tc.keyword)
		}
	}
}
