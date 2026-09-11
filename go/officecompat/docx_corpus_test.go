package officecompat_test

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/injectinglabs/injoffice/go/docxpatch"
	"github.com/injectinglabs/injoffice/go/officecompat/corpus"
)

func TestDOCXCorpusNativeExpectations(t *testing.T) {
	manifestBytes := readDOCXCorpusFile(t, "manifest.json")
	var manifest corpus.Manifest
	if err := json.Unmarshal(manifestBytes, &manifest); err != nil {
		t.Fatal(err)
	}

	docxFixtures := 0
	accepted := 0
	refused := 0
	for _, fixture := range manifest.Fixtures {
		if fixture.Format != "docx" {
			continue
		}
		docxFixtures++
		fixture := fixture
		t.Run(fixture.ID, func(t *testing.T) {
			if fixture.Provenance.Kind != "generated" || fixture.Provenance.License != "CC0-1.0" || !strings.Contains(fixture.Provenance.Source, "CC0 hand-authored") {
				t.Fatalf("unstable or non-redistributable provenance: %+v", fixture.Provenance)
			}

			packageBytes := readDOCXCorpusFile(t, fixture.Package)
			expectedBytes := readDOCXCorpusFile(t, fixture.Expected)
			if got := docxCorpusSHA256(packageBytes); got != fixture.SHA256 {
				t.Fatalf("package SHA-256 = %s, manifest = %s", got, fixture.SHA256)
			}
			if got := docxCorpusSHA256(expectedBytes); got != fixture.ExpectedSHA256 {
				t.Fatalf("expectation SHA-256 = %s, manifest = %s", got, fixture.ExpectedSHA256)
			}

			var expectation corpus.Expectation
			if err := json.Unmarshal(expectedBytes, &expectation); err != nil {
				t.Fatal(err)
			}
			if expectation.Protocol != corpus.ExpectationProtocol || expectation.FixtureID != fixture.ID || expectation.Format != "docx" || expectation.Dialect != fixture.Dialect || expectation.Outcome != fixture.Outcome {
				t.Fatalf("expectation envelope does not match manifest: %+v", expectation)
			}

			doc, extractErr := docxpatch.ExtractNativeDocumentV1(packageBytes)
			switch fixture.Outcome {
			case "accepted":
				accepted++
				assertAcceptedDOCXCorpusFixture(t, fixture, packageBytes, expectation, doc, extractErr)
			case "refused":
				refused++
				assertRefusedDOCXCorpusFixture(t, expectation, doc, extractErr)
			default:
				t.Fatalf("unsupported fixture outcome %q", fixture.Outcome)
			}
		})
	}
	if docxFixtures != 9 || accepted != 6 || refused != 3 {
		t.Fatalf("DOCX fixture matrix = total %d, accepted %d, refused %d; want 9/6/3", docxFixtures, accepted, refused)
	}
}

func TestDOCXCorpusSpecsContainNoLegacyReconstructionAuthority(t *testing.T) {
	forbidden := []string{"mammoth", "luckyexcel", "domparser", "jszip", "innerhtml", "<html", "document.objectmodel", "browser zip"}
	paths, err := filepath.Glob(filepath.Join("corpus", "specs", "docx", "*.json"))
	if err != nil {
		t.Fatal(err)
	}
	if len(paths) != 9 {
		t.Fatalf("DOCX spec count = %d, want 9", len(paths))
	}
	for _, path := range paths {
		data, err := os.ReadFile(path)
		if err != nil {
			t.Fatal(err)
		}
		lower := strings.ToLower(string(data))
		for _, token := range forbidden {
			if strings.Contains(lower, token) {
				t.Errorf("%s contains forbidden legacy reconstruction token %q", path, token)
			}
		}
		if !strings.Contains(string(data), `"license": "CC0-1.0"`) {
			t.Errorf("%s does not carry stable redistributable provenance", path)
		}
	}
}

func TestDOCXCorpusRefusesDanglingRelatedStoryRelationship(t *testing.T) {
	manifestBytes := readDOCXCorpusFile(t, "manifest.json")
	var manifest corpus.Manifest
	if err := json.Unmarshal(manifestBytes, &manifest); err != nil {
		t.Fatal(err)
	}
	var fixture *corpus.FixtureRecord
	for index := range manifest.Fixtures {
		if manifest.Fixtures[index].ID == "docx-related-stories" {
			fixture = &manifest.Fixtures[index]
			break
		}
	}
	if fixture == nil {
		t.Fatal("docx-related-stories fixture is missing from the canonical manifest")
	}

	broken := rewriteDOCXCorpusPart(t, readDOCXCorpusFile(t, fixture.Package), "word/_rels/document.xml.rels", func(data []byte) []byte {
		const target = `Target="stories/header.xml"`
		if bytes.Count(data, []byte(target)) != 1 {
			t.Fatalf("fixture header target occurrence count = %d, want 1", bytes.Count(data, []byte(target)))
		}
		return bytes.Replace(data, []byte(target), []byte(`Target="stories/missing-header.xml"`), 1)
	})
	doc, err := docxpatch.ExtractNativeDocumentV1(broken)
	if err == nil || doc != nil || !strings.Contains(err.Error(), `relationship "rHeader"`) || !strings.Contains(err.Error(), `targets missing part "word/stories/missing-header.xml"`) {
		t.Fatalf("dangling related-story relationship did not fail closed: doc=%#v err=%v", doc, err)
	}
}

func assertAcceptedDOCXCorpusFixture(t *testing.T, fixture corpus.FixtureRecord, packageBytes []byte, expectation corpus.Expectation, doc *docxpatch.NativeDocumentV1, extractErr error) {
	t.Helper()
	if extractErr != nil {
		t.Fatal(extractErr)
	}
	if expectation.Refusal != nil {
		t.Fatalf("accepted fixture carries refusal: %+v", expectation.Refusal)
	}
	if got, want := doc.Source.PackageSHA256, "sha256:"+fixture.SHA256; got != want {
		t.Fatalf("native package provenance = %q, want %q", got, want)
	}
	if issues := docxpatch.ValidateNativeDocumentV1(doc); len(issues) != 0 {
		t.Fatalf("extracted native contract is invalid: %+v", issues)
	}

	encoded, err := docxpatch.EncodeNativeDocumentV1(doc)
	if err != nil {
		t.Fatal(err)
	}
	var expectedCompact bytes.Buffer
	if err := json.Compact(&expectedCompact, expectation.Native); err != nil {
		t.Fatalf("compact expected native JSON: %v", err)
	}
	if !bytes.Equal(encoded, expectedCompact.Bytes()) {
		t.Fatalf("native JSON differs from exact expectation\n got: %s\nwant: %s", encoded, expectedCompact.Bytes())
	}

	decoded, err := docxpatch.DecodeNativeDocumentV1(encoded)
	if err != nil {
		t.Fatalf("decode extracted native JSON: %v", err)
	}
	reencoded, err := docxpatch.EncodeNativeDocumentV1(decoded)
	if err != nil {
		t.Fatalf("re-encode extracted native JSON: %v", err)
	}
	if !bytes.Equal(encoded, reencoded) {
		t.Fatal("Extract -> Encode -> Decode -> Encode was not byte-identical")
	}

	parts := readDOCXCorpusParts(t, packageBytes)
	assertDOCXCorpusPassthroughParts(t, parts, doc)
	anchors := collectDOCXCorpusAnchors(doc)
	if len(anchors) == 0 {
		t.Fatal("accepted fixture exposed no exact source anchors")
	}
	for index := range anchors {
		assertDOCXCorpusAnchor(t, parts, &anchors[index])
	}
	assertDOCXCorpusFixtureSemantics(t, fixture.ID, doc)
}

func assertRefusedDOCXCorpusFixture(t *testing.T, expectation corpus.Expectation, doc *docxpatch.NativeDocumentV1, extractErr error) {
	t.Helper()
	if extractErr == nil || doc != nil {
		t.Fatalf("refused fixture returned a partial or accepted native document: doc=%#v err=%v", doc, extractErr)
	}
	if expectation.Refusal == nil {
		t.Fatal("refused fixture is missing its refusal expectation")
	}
	if got := classifyDOCXCorpusRefusal(extractErr); got != expectation.Refusal.Class {
		t.Fatalf("refusal class = %q, want %q; error: %v", got, expectation.Refusal.Class, extractErr)
	}
	if !strings.Contains(extractErr.Error(), expectation.Refusal.Contains) {
		t.Fatalf("refusal %q does not contain stable substring %q", extractErr, expectation.Refusal.Contains)
	}
}

func assertDOCXCorpusPassthroughParts(t *testing.T, parts map[string][]byte, doc *docxpatch.NativeDocumentV1) {
	t.Helper()
	for _, passthrough := range doc.PassthroughParts {
		payload, found := parts[passthrough.PartName]
		if !found || passthrough.ByteLength == nil || *passthrough.ByteLength != int64(len(payload)) || passthrough.Policy != "preserve-verbatim" {
			t.Fatalf("passthrough provenance mismatch: %+v", passthrough)
		}
		digest := sha256.Sum256(payload)
		if got, want := passthrough.SHA256, "sha256:"+hex.EncodeToString(digest[:]); got != want {
			t.Fatalf("passthrough digest = %q, source package = %q for %s", got, want, passthrough.PartName)
		}
	}
}

func classifyDOCXCorpusRefusal(err error) string {
	message := err.Error()
	switch {
	case strings.Contains(message, "directives/DOCTYPE are forbidden"):
		return "invalid-xml"
	case strings.Contains(message, "unsafe percent-encoded"):
		return "invalid-opc-path"
	case strings.Contains(message, "compression-ratio limit"):
		return "resource-limit"
	default:
		return "unknown"
	}
}

func assertDOCXCorpusFixtureSemantics(t *testing.T, fixtureID string, doc *docxpatch.NativeDocumentV1) {
	t.Helper()
	switch fixtureID {
	case "docx-transitional-common":
		if doc.Source.MainPart != "word/document.xml" || len(doc.Body.Blocks) != 2 || doc.Body.Blocks[1].Kind != "table" {
			t.Fatalf("common Transitional projection changed: %+v", doc.Body)
		}
		if got := *doc.Body.Blocks[0].Paragraph.Runs[0].Text; got != "  Café 👋 " {
			t.Fatalf("lexical text projection = %q", got)
		}
		if !hasDOCXCorpusUnsupported(doc, "DEFAULT_SECTION_INFERRED") {
			t.Fatal("schema-optional final section was not diagnosed")
		}
	case "docx-strict-relocated":
		if doc.Source.MainPart != "Odd/Main.XML" || len(doc.Headers) != 1 || doc.Headers[0].PartName != "Odd/Stories/Header.XML" {
			t.Fatalf("Strict case-preserved routing changed: source=%+v headers=%+v", doc.Source, doc.Headers)
		}
	case "docx-preserve-refuse":
		for _, code := range []string{"FIELD_SEMANTICS", "HYPERLINK_SEMANTICS", "PARTIAL_RUN_PROPERTIES", "PICTURE_NONVISUAL_PRESERVED", "UNMODELED_BODY_BLOCK", "WRAPPED_RUN_MARKUP"} {
			if !hasDOCXCorpusUnsupported(doc, code) {
				t.Fatalf("accepted preserve/refuse fixture is missing %s", code)
			}
		}
		if !hasDOCXCorpusPassthrough(doc, "customXml/opaque.bin") {
			t.Fatal("opaque CC0 payload is not preserve-verbatim passthrough")
		}
	case "docx-related-stories":
		assertDOCXCorpusRelatedStorySemantics(t, doc)
	case "docx-sections-columns":
		wantBreaks := []string{"continuous", "next-page", "even-page", "odd-page", "next-column"}
		if len(doc.Sections) != len(wantBreaks) {
			t.Fatalf("section/column fixture section count = %d, want %d", len(doc.Sections), len(wantBreaks))
		}
		for index, section := range doc.Sections {
			if section.BreakType != wantBreaks[index] || section.Page.Columns == nil || *section.Page.Columns != 2 || len(section.Page.ColumnDefinitions) != 2 {
				t.Fatalf("section %d transition/columns changed: %+v", index, section)
			}
			wantLayout := "equal-width"
			if index == 1 || index == 3 {
				wantLayout = "explicit"
			}
			if section.Page.ColumnLayout != wantLayout {
				t.Fatalf("section %d column layout = %q, want %q", index, section.Page.ColumnLayout, wantLayout)
			}
			for ordinal, column := range section.Page.ColumnDefinitions {
				if column.Ordinal == nil || *column.Ordinal != ordinal || column.ID == "" {
					t.Fatalf("section %d column %d identity changed: %+v", index, ordinal, column)
				}
			}
		}
	}
}

func assertDOCXCorpusRelatedStorySemantics(t *testing.T, doc *docxpatch.NativeDocumentV1) {
	t.Helper()
	if len(doc.Headers) != 1 || len(doc.Footers) != 1 || len(doc.Notes) != 6 || len(doc.Comments) != 1 || len(doc.CommentStories) != 1 || len(doc.Sections) != 1 {
		t.Fatalf("related-story cardinality changed: headers=%d footers=%d notes=%d comments=%d commentStories=%d sections=%d", len(doc.Headers), len(doc.Footers), len(doc.Notes), len(doc.Comments), len(doc.CommentStories), len(doc.Sections))
	}
	section := &doc.Sections[0]
	if len(section.HeaderRefs) != 1 || section.HeaderRefs[0].RelationshipID != "rHeader" || section.HeaderRefs[0].StoryID != doc.Headers[0].ID || doc.Headers[0].PartName != "word/stories/header.xml" {
		t.Fatalf("header relationship-to-story binding changed: section=%+v headers=%+v", section.HeaderRefs, doc.Headers)
	}
	if len(section.FooterRefs) != 1 || section.FooterRefs[0].RelationshipID != "rFooter" || section.FooterRefs[0].StoryID != doc.Footers[0].ID || doc.Footers[0].PartName != "word/stories/footer.xml" {
		t.Fatalf("footer relationship-to-story binding changed: section=%+v footers=%+v", section.FooterRefs, doc.Footers)
	}

	storyIDs := map[string]string{}
	for index := range doc.Notes {
		note := &doc.Notes[index]
		if note.NativeStoryID == nil {
			t.Fatalf("%s story lacks native identity: %+v", note.Kind, note)
		}
		storyIDs[note.Kind+":"+*note.NativeStoryID] = note.ID
	}
	if storyIDs["footnote:-1"] == "" || storyIDs["footnote:0"] == "" || storyIDs["footnote:1"] == "" || storyIDs["endnote:-1"] == "" || storyIDs["endnote:0"] == "" || storyIDs["endnote:2"] == "" {
		t.Fatalf("note identities or separator sentinel classification changed: %+v", storyIDs)
	}
	comment := &doc.Comments[0]
	commentStory := &doc.CommentStories[0]
	if comment.NativeCommentID != "3" || comment.BodyStoryID != commentStory.ID || commentStory.NativeStoryID == nil || *commentStory.NativeStoryID != "3" {
		t.Fatalf("comment metadata/body identity changed: comment=%+v story=%+v", comment, commentStory)
	}

	wantReferences := map[string]string{
		"comment-range-start": comment.ID,
		"comment":             comment.ID,
		"comment-range-end":   comment.ID,
		"footnote":            storyIDs["footnote:1"],
		"endnote":             storyIDs["endnote:2"],
	}
	gotReferences := map[string]string{}
	for _, run := range doc.Body.Blocks[0].Paragraph.Runs {
		if run.Reference != nil {
			gotReferences[run.Reference.Kind] = run.Reference.TargetID
		}
	}
	for kind, want := range wantReferences {
		if gotReferences[kind] != want {
			t.Fatalf("%s reference target = %q, want %q (all=%+v)", kind, gotReferences[kind], want, gotReferences)
		}
	}
	if hasDOCXCorpusUnsupported(doc, "SPECIAL_NOTE_STORY") || !hasDOCXCorpusPassthrough(doc, "word/_rels/document.xml.rels") {
		t.Fatalf("related-story preservation/refusal evidence changed: unsupported=%+v passthrough=%+v", doc.Unsupported, doc.PassthroughParts)
	}
}

func hasDOCXCorpusUnsupported(doc *docxpatch.NativeDocumentV1, code string) bool {
	for _, entry := range doc.Unsupported {
		if entry.Code == code && entry.Preservation == "refuse-mutation" {
			return true
		}
	}
	return false
}

func hasDOCXCorpusPassthrough(doc *docxpatch.NativeDocumentV1, partName string) bool {
	for _, part := range doc.PassthroughParts {
		if part.PartName == partName && part.Policy == "preserve-verbatim" {
			return true
		}
	}
	return false
}

func collectDOCXCorpusAnchors(doc *docxpatch.NativeDocumentV1) []docxpatch.NativeSourceAnchorV1 {
	var anchors []docxpatch.NativeSourceAnchorV1
	add := func(anchor *docxpatch.NativeSourceAnchorV1) {
		if anchor != nil {
			anchors = append(anchors, *anchor)
		}
	}
	var paragraph func(*docxpatch.NativeParagraphV1)
	paragraph = func(value *docxpatch.NativeParagraphV1) {
		add(&value.Anchor)
		for index := range value.Runs {
			run := &value.Runs[index]
			add(&run.Anchor)
			if run.Drawing != nil {
				add(&run.Drawing.Anchor)
			}
		}
	}
	var story func(*docxpatch.NativeStoryV1)
	story = func(value *docxpatch.NativeStoryV1) {
		add(value.Anchor)
		for index := range value.Blocks {
			block := &value.Blocks[index]
			if block.Paragraph != nil {
				paragraph(block.Paragraph)
			}
			if block.Table != nil {
				add(&block.Table.Anchor)
				for rowIndex := range block.Table.Rows {
					row := &block.Table.Rows[rowIndex]
					add(&row.Anchor)
					for cellIndex := range row.Cells {
						cell := &row.Cells[cellIndex]
						add(&cell.Anchor)
						for paragraphIndex := range cell.Paragraphs {
							paragraph(&cell.Paragraphs[paragraphIndex])
						}
					}
				}
			}
		}
	}
	story(&doc.Body)
	for _, stories := range [][]docxpatch.NativeStoryV1{doc.Headers, doc.Footers, doc.Notes, doc.CommentStories} {
		for index := range stories {
			story(&stories[index])
		}
	}
	for index := range doc.Sections {
		add(&doc.Sections[index].Anchor)
	}
	for index := range doc.Comments {
		add(doc.Comments[index].Anchor)
	}
	for index := range doc.Unsupported {
		add(doc.Unsupported[index].Anchor)
	}
	return anchors
}

func assertDOCXCorpusAnchor(t *testing.T, parts map[string][]byte, anchor *docxpatch.NativeSourceAnchorV1) {
	t.Helper()
	raw, ok := parts[anchor.PartName]
	if !ok {
		t.Fatalf("anchor targets missing part %q", anchor.PartName)
	}
	if anchor.StartByte == nil || anchor.EndByte == nil || *anchor.StartByte < 0 || *anchor.EndByte > int64(len(raw)) || *anchor.EndByte <= *anchor.StartByte {
		t.Fatalf("invalid anchor range: %+v for %d-byte part", anchor, len(raw))
	}
	anchored := raw[*anchor.StartByte:*anchor.EndByte]
	if len(anchored) == 0 || anchored[0] != '<' {
		t.Fatalf("anchor %s does not begin at an XML element", anchor.Path)
	}
	digest := sha256.Sum256(anchored)
	if got, want := "sha256:"+hex.EncodeToString(digest[:]), anchor.XMLSHA256; got != want {
		t.Fatalf("anchor digest = %q, native contract = %q", got, want)
	}
}

func readDOCXCorpusParts(t *testing.T, data []byte) map[string][]byte {
	t.Helper()
	reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	parts := make(map[string][]byte, len(reader.File))
	for _, file := range reader.File {
		if file.FileInfo().IsDir() {
			continue
		}
		handle, err := file.Open()
		if err != nil {
			t.Fatal(err)
		}
		content, readErr := io.ReadAll(handle)
		closeErr := handle.Close()
		if readErr != nil {
			t.Fatal(readErr)
		}
		if closeErr != nil {
			t.Fatal(closeErr)
		}
		parts[file.Name] = content
	}
	return parts
}

func rewriteDOCXCorpusPart(t *testing.T, data []byte, partName string, rewrite func([]byte) []byte) []byte {
	t.Helper()
	reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	var output bytes.Buffer
	writer := zip.NewWriter(&output)
	found := false
	for _, file := range reader.File {
		handle, err := file.Open()
		if err != nil {
			t.Fatal(err)
		}
		payload, readErr := io.ReadAll(handle)
		closeErr := handle.Close()
		if readErr != nil {
			t.Fatal(readErr)
		}
		if closeErr != nil {
			t.Fatal(closeErr)
		}
		if file.Name == partName {
			payload = rewrite(payload)
			found = true
		}
		header := file.FileHeader
		header.CRC32 = 0
		header.CompressedSize = 0
		header.CompressedSize64 = 0
		header.UncompressedSize = 0
		header.UncompressedSize64 = 0
		entry, err := writer.CreateHeader(&header)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := entry.Write(payload); err != nil {
			t.Fatal(err)
		}
	}
	if !found {
		t.Fatalf("DOCX fixture has no part %q", partName)
	}
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	return output.Bytes()
}

func readDOCXCorpusFile(t *testing.T, relative string) []byte {
	t.Helper()
	data, err := os.ReadFile(filepath.Join("corpus", filepath.FromSlash(relative)))
	if err != nil {
		t.Fatal(err)
	}
	return data
}

func docxCorpusSHA256(data []byte) string {
	digest := sha256.Sum256(data)
	return fmt.Sprintf("%x", digest)
}
