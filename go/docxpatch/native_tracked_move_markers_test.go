package docxpatch

import "testing"

const trackedMoveRow = `<w:tr><w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr>`

func trackedMoveTableDocument(t *testing.T, between string) *NativeDocumentV1 {
	t.Helper()
	body := `<w:tbl><w:tblPr><w:tblW w:w="4680" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4680"/></w:tblGrid>` +
		trackedMoveRow + between + trackedMoveRow + `</w:tbl><w:p><w:r><w:t>after</w:t></w:r></w:p>`
	doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(nativeMutationParts(nativeMutationMain(body)))))
	if err != nil {
		t.Fatal(err)
	}
	return doc
}

func trackedMoveUnsupportedCode(doc *NativeDocumentV1, capability string) string {
	for _, entry := range doc.Unsupported {
		if entry.Capability == capability {
			return entry.Code
		}
	}
	return ""
}

// A dragged table row leaves its move-range endpoints between the rows. They
// are empty delimiters, not table content, and must not take the whole table
// out of pagination.
func TestNativeTableRangeEndpointIsNotUnmodeledTableContent(t *testing.T) {
	for name, markup := range map[string]string{
		"move from": `<w:moveFromRangeEnd w:id="1"/>`,
		"move to":   `<w:moveToRangeEnd w:id="8"/>`,
		"bookmark":  `<w:bookmarkEnd w:id="1"/>`,
		"start":     `<w:moveToRangeStart w:id="8" w:author="Nick" w:date="2026-01-01T00:00:00Z" w:name="move1"/>`,
	} {
		t.Run(name, func(t *testing.T) {
			doc := trackedMoveTableDocument(t, markup)
			if code := trackedMoveUnsupportedCode(doc, "table-structure"); code != "NON_VISUAL_RANGE_MARKER" {
				t.Fatalf("table-structure record = %q, want NON_VISUAL_RANGE_MARKER", code)
			}
			if len(doc.Body.Blocks) == 0 || doc.Body.Blocks[0].Table == nil || len(doc.Body.Blocks[0].Table.Rows) != 2 {
				t.Fatalf("rows around the endpoint were dropped: %#v", doc.Body.Blocks)
			}
		})
	}
}

// Only a content-free, word-namespace endpoint qualifies. Anything that can
// carry meaning this layer has not read stays generic table content.
func TestNativeTableContentOutsideRowsKeepsRefusing(t *testing.T) {
	for name, markup := range map[string]string{
		"paragraph":         `<w:p><w:r><w:t>stray</w:t></w:r></w:p>`,
		"nested markup":     `<w:bookmarkEnd w:id="1"><w:p/></w:bookmarkEnd>`,
		"foreign attribute": `<w:moveFromRangeEnd xmlns:x="urn:foreign" w:id="1" x:mode="opaque"/>`,
		"unknown element":   `<w:customXml w:element="rows"/>`,
	} {
		t.Run(name, func(t *testing.T) {
			doc := trackedMoveTableDocument(t, markup)
			if code := trackedMoveUnsupportedCode(doc, "table-structure"); code != "UNMODELED_TABLE_CONTENT" {
				t.Fatalf("table-structure record = %q, want UNMODELED_TABLE_CONTENT", code)
			}
		})
	}
}

func trackedMarkRevisionCodes(t *testing.T, markup string) []string {
	t.Helper()
	body := `<w:p><w:pPr><w:rPr>` + markup + `</w:rPr></w:pPr><w:r><w:t>moved</w:t></w:r></w:p><w:sectPr/>`
	resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(nativeMutationParts(nativeMutationMain(body)))))
	if err != nil {
		t.Fatal(err)
	}
	codes := []string{}
	for _, diagnostic := range resolved.Diagnostics {
		codes = append(codes, diagnostic.Code)
	}
	return codes
}

func trackedMarkHasCode(codes []string, want string) bool {
	for _, code := range codes {
		if code == want {
			return true
		}
	}
	return false
}

// w:ins, w:del, w:moveFrom and w:moveTo inside a paragraph mark's run
// properties name the revision the mark belongs to. They state no formatting,
// so reporting them as an unresolved run property names the wrong fact.
func TestNativeTrackedMarkRevisionIsNotAnUnresolvedRunProperty(t *testing.T) {
	for name, markup := range map[string]string{
		"move from": `<w:moveFrom w:id="1" w:author="Nick" w:date="2026-01-01T00:00:00Z"/>`,
		"move to":   `<w:moveTo w:id="2" w:author="Nick"/>`,
		"insertion": `<w:ins w:id="3" w:author="Nick"/>`,
		"deletion":  `<w:del w:id="4" w:author="Nick"/>`,
	} {
		t.Run(name, func(t *testing.T) {
			codes := trackedMarkRevisionCodes(t, markup)
			if !trackedMarkHasCode(codes, "TRACKED_MARK_REVISION_PRESERVED") {
				t.Fatalf("diagnostics = %v, want TRACKED_MARK_REVISION_PRESERVED", codes)
			}
			if trackedMarkHasCode(codes, "UNMODELED_RUN_PROPERTY") {
				t.Fatalf("tracked revision still reported as an unresolved run property: %v", codes)
			}
		})
	}
}

// A revision annotation carrying markup or attributes outside CT_TrackChange is
// not one this layer has read, and keeps its generic refusal.
func TestNativeUnqualifiedTrackedMarkRevisionKeepsRefusing(t *testing.T) {
	for name, markup := range map[string]string{
		"nested markup":     `<w:moveFrom w:id="1"><w:r><w:t>inner</w:t></w:r></w:moveFrom>`,
		"unknown attribute": `<w:moveTo w:id="2" w:mystery="1"/>`,
	} {
		t.Run(name, func(t *testing.T) {
			codes := trackedMarkRevisionCodes(t, markup)
			if !trackedMarkHasCode(codes, "UNMODELED_RUN_PROPERTY") {
				t.Fatalf("diagnostics = %v, want UNMODELED_RUN_PROPERTY", codes)
			}
		})
	}
}
