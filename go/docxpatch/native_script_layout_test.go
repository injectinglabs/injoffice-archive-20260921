package docxpatch

import (
	"strings"
	"testing"
)

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

func TestNativeScriptTransformEligibleRunBoundary(t *testing.T) {
	for _, eligible := range []*NativeRunV1{
		{Kind: "text"},
		{Kind: "reference", Reference: &NativeReferenceV1{Kind: "footnote"}},
		{Kind: "reference", Reference: &NativeReferenceV1{Kind: "endnote"}},
		{Kind: "reference", Reference: &NativeReferenceV1{Kind: "footnote", Role: "label"}},
		{Kind: "reference", Reference: &NativeReferenceV1{Kind: "endnote", Role: "label"}},
	} {
		if !nativeScriptTransformEligibleRun(eligible) {
			t.Fatalf("shaped-text run refused a modeled script transform: %#v", eligible)
		}
	}
	for _, refused := range []*NativeRunV1{
		{Kind: "control", Control: "tab"},
		{Kind: "control", Control: "soft-hyphen"},
		{Kind: "control", Control: "line-break"},
		{Kind: "drawing", Drawing: &NativeDrawingV1{}},
		{Kind: "reference", Reference: &NativeReferenceV1{Kind: "comment"}},
		{Kind: "reference", Reference: &NativeReferenceV1{Kind: "comment-range-start"}},
		{Kind: "reference", Reference: &NativeReferenceV1{Kind: "comment-range-end"}},
		{Kind: "reference"},
	} {
		if nativeScriptTransformEligibleRun(refused) {
			t.Fatalf("glyphless or image run admitted an unmodeled script transform: %#v", refused)
		}
	}
}

// The note marker shapes its placed decimal number through the same shaped-text
// path as a text run, so the authored superscript on it is the modeled OS/2
// transform. A control run emits a glyphless atom that no transform reaches, so
// its superscript stays refused.
func TestNativeNoteMarkerSuperscriptResolvesWhileControlSuperscriptRefuses(t *testing.T) {
	for _, kind := range []NoteKind{Footnote, Endnote} {
		t.Run(string(kind), func(t *testing.T) {
			withNote, _, err := InsertNote(buildNoteTestDocx(t), kind, 0, "Marker script fixture.")
			if err != nil {
				t.Fatal(err)
			}
			documentXML, ok := zipPart(t, withNote, docPart)
			if !ok {
				t.Fatal("note fixture document.xml missing")
			}
			patched := strings.Replace(documentXML, `<w:sectPr>`, `<w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:tab/></w:r></w:p><w:sectPr>`, 1)
			if patched == documentXML {
				t.Fatal("note fixture body shape changed")
			}
			source, err := ApplyPatch(withNote, Patch{
				Replace: map[string][]byte{docPart: []byte(patched)},
				Add:     map[string][]byte{},
				Delete:  map[string]bool{},
			})
			if err != nil {
				t.Fatal(err)
			}
			document, err := ExtractNativeDocumentV1(source)
			if err != nil {
				t.Fatal(err)
			}
			controlRuns, markerRuns := map[string]bool{}, map[string]bool{}
			for _, story := range append([]NativeStoryV1{document.Body}, document.Notes...) {
				for _, block := range story.Blocks {
					if block.Paragraph == nil {
						continue
					}
					for _, run := range block.Paragraph.Runs {
						if run.Kind == "control" {
							controlRuns[run.ID] = true
						}
						if run.Kind == "reference" && run.Reference != nil && run.Reference.Kind == string(kind) {
							markerRuns[run.ID] = true
						}
					}
				}
			}
			if len(controlRuns) != 1 || len(markerRuns) != 2 {
				t.Fatalf("fixture must hold one control run and the anchor plus label note marks: controls=%d marks=%d", len(controlRuns), len(markerRuns))
			}
			resolved, err := ResolveNativeDocumentLayoutV1(source)
			if err != nil {
				t.Fatal(err)
			}
			refused := map[string]bool{}
			for _, diagnostic := range resolved.Diagnostics {
				if diagnostic.Code == "VERTICAL_ALIGNMENT_UNSUPPORTED" {
					refused[diagnostic.ScopeID] = true
				}
			}
			for id := range markerRuns {
				if refused[id] {
					t.Fatalf("note marker superscript is the modeled shaped-text transform, not a refusal: %s", id)
				}
			}
			for id := range controlRuns {
				if !refused[id] {
					t.Fatalf("control superscript must stay refused: %s", id)
				}
			}
			if len(refused) != len(controlRuns) {
				t.Fatalf("unexpected vertical alignment refusals: %#v", refused)
			}
			superscripts := 0
			for _, run := range resolved.Runs {
				if markerRuns[run.RunID] {
					if run.Properties.VerticalAlignment == nil || *run.Properties.VerticalAlignment != "superscript" {
						t.Fatalf("note marker lost its resolved superscript: %#v", run.Properties)
					}
					superscripts++
				}
			}
			if superscripts != len(markerRuns) {
				t.Fatalf("resolved note markers = %d, want %d", superscripts, len(markerRuns))
			}
		})
	}
}
