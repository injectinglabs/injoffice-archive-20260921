package docxpatch

import (
	"encoding/json"
	"strings"
	"testing"
)

func inspectReviewTest(t *testing.T, body string) (*NativeDocumentV1, NativePartialReviewV1, []byte) {
	t.Helper()
	source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(nativeMutationMain(body))))
	doc, e := ExtractNativeDocumentV1(source)
	if e != nil {
		t.Fatal(e)
	}
	encoded, e := InspectNativePartialSourceV1(source)
	if e != nil {
		t.Fatal(e)
	}
	var out struct {
		Review NativePartialReviewV1 `json:"review_changes"`
	}
	if e = json.Unmarshal(encoded, &out); e != nil {
		t.Fatal(e)
	}
	return doc, out.Review, source
}
func TestPartialReviewMetadataAndStrictMutationRefusal(t *testing.T) {
	doc, review, source := inspectReviewTest(t, `<w:p><w:ins w:id="7" w:author="A &amp; B" w:date="2026-09-12T00:00:00Z"><w:r><w:t>Inserted</w:t></w:r></w:ins></w:p><w:p><w:del w:id="8" w:author="A"><w:r><w:delText>SECRET_DELETED</w:delText></w:r></w:del></w:p>`)
	if len(review.Items) != 2 || review.Items[0].Kind != "insertion" || review.Items[0].Author != "A & B" || review.Items[0].PackageSHA256 != doc.Source.PackageSHA256 || len(review.Items[0].RunIDs) != 1 || review.Items[1].Kind != "deletion" || len(review.Items[1].RunIDs) != 0 {
		t.Fatalf("bad metadata: %+v", review)
	}
	encoded, _ := json.Marshal(review)
	if strings.Contains(string(encoded), "SECRET_DELETED") || strings.Contains(string(encoded), "Inserted") {
		t.Fatal("sidecar copied text")
	}
	p := doc.Body.Blocks[0].Paragraph
	if p.EditPolicy.Mode != "read-only" {
		t.Fatal("revision became editable")
	}
	_, err := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXTextMutationV1{{TargetKind: "paragraph", TargetID: p.ID, ExpectedXMLSHA256: p.Anchor.XMLSHA256, Text: "changed"}})
	if err == nil {
		t.Fatal("revision mutation admitted")
	}
}
func TestPartialReviewDoesNotTraverseNestedOrOpaqueRuns(t *testing.T) {
	for _, inner := range []string{`<w:del w:id="9" w:author="A"><w:r><w:delText>SECRET</w:delText></w:r></w:del>`, `<w:r><w:fldChar w:fldCharType="begin"/></w:r>`, `<w:smartTag><w:r><w:t>SECRET</w:t></w:r></w:smartTag>`, `<w:r other="unknown"><w:t>SECRET</w:t></w:r>`, `<w:r><w:t xml:space="default">SECRET</w:t></w:r>`} {
		_, review, _ := inspectReviewTest(t, `<w:p><w:ins w:id="1" w:author="A">`+inner+`</w:ins></w:p>`)
		if len(review.Items) != 1 || len(review.Items[0].RunIDs) != 0 {
			t.Fatalf("opaque content qualified: %s %+v", inner, review)
		}
	}
	for _, kind := range []string{"moveTo", "moveFrom"} {
		_, review, _ := inspectReviewTest(t, `<w:p><w:`+kind+` w:id="1" w:author="A"><w:r><w:t>SECRET</w:t></w:r></w:`+kind+`></w:p>`)
		if len(review.Items) != 1 || len(review.Items[0].RunIDs) != 0 {
			t.Fatal("move text qualified")
		}
	}
}
func TestPartialReviewRejectsUnknownWrapperMetadataAndBoundsInventory(t *testing.T) {
	for _, attrs := range []string{`w:id="1"`, `w:id="1" w:author="A" unknown="x"`} {
		_, review, _ := inspectReviewTest(t, `<w:p><w:ins `+attrs+`><w:r><w:t>Text</w:t></w:r></w:ins></w:p>`)
		if len(review.Items) != 0 || review.OmittedCount != 1 {
			t.Fatal("unknown metadata exposed")
		}
	}
	body := strings.Repeat(`<w:p><w:ins w:id="1" w:author="A"><w:r><w:t>Text</w:t></w:r></w:ins></w:p>`, 129)
	_, review, _ := inspectReviewTest(t, body)
	if len(review.Items) != 128 || review.OmittedCount != 1 {
		t.Fatalf("budget failed: %d/%d", len(review.Items), review.OmittedCount)
	}
}
