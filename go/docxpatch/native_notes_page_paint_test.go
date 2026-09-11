package docxpatch

import (
	"slices"
	"strings"
	"testing"
)

func nativeNotePagePaintParts() map[string]string {
	parts := transitionalNativeParts()
	parts["Custom/Notes/Foot.XML"] = `<w:footnotes xmlns:w="` + testW + `"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:t>---</w:t></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:t>continued</w:t></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>Exact footnote</w:t></w:r></w:p></w:footnote></w:footnotes>`
	parts["Custom/Notes/End.XML"] = `<w:endnotes xmlns:w="` + testW + `"><w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:t>---</w:t></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:t>continued</w:t></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:endnoteRef/></w:r><w:r><w:t>Exact endnote</w:t></w:r></w:p></w:endnote></w:endnotes>`
	return parts
}

func TestExtractNativeNoteStoriesAllowsIgnorableRootMetadata(t *testing.T) {
	parts := nativeNotePagePaintParts()
	for _, part := range []string{"Custom/Notes/Foot.XML", "Custom/Notes/End.XML"} {
		parts[part] = strings.Replace(parts[part], `xmlns:w="`+testW+`"`, `xmlns:w="`+testW+`" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"`, 1)
	}
	doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	if len(doc.Notes) != 6 {
		t.Fatalf("lost note stories: %d", len(doc.Notes))
	}
	for _, note := range doc.Notes {
		anchor := note.Anchor
		if anchor == nil || anchor.StartByte == nil || anchor.EndByte == nil || anchor.XMLSHA256 != nativeSHA([]byte(parts[note.PartName])[*anchor.StartByte:*anchor.EndByte]) {
			t.Fatalf("note lost exact source anchor: %#v", note)
		}
	}
}

func TestExtractNativeNoteStoriesRejectsSpoofedIgnorableRootMetadata(t *testing.T) {
	for _, attrs := range []string{`Ignorable="w14"`, `xmlns:mc="urn:spoof" mc:Ignorable="w14"`, `xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:MustUnderstand="w14"`} {
		parts := nativeNotePagePaintParts()
		parts["Custom/Notes/Foot.XML"] = strings.Replace(parts["Custom/Notes/Foot.XML"], `<w:footnotes `, `<w:footnotes `+attrs+` `, 1)
		if _, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts))); err == nil {
			t.Fatalf("unsupported root metadata accepted: %s", attrs)
		}
	}
}

func nativeNotePagePaintPackage(t *testing.T) []byte {
	parts := nativeNotePagePaintParts()
	return buildNativeDOCX(t, nativeEntries(parts))
}

func TestExtractNativeNoteStoriesRetainsRelationshipRolesLabelsAndAnchors(t *testing.T) {
	doc, err := ExtractNativeDocumentV1(nativeNotePagePaintPackage(t))
	if err != nil {
		t.Fatal(err)
	}
	if len(doc.Notes) != 6 {
		t.Fatalf("notes = %d, want six content/separator stories", len(doc.Notes))
	}
	roles := map[string]bool{}
	for _, story := range doc.Notes {
		if story.Anchor == nil || story.Anchor.PartName != story.PartName || story.RelationshipID == nil {
			t.Fatalf("note lost anchor/relationship closure: %#v", story)
		}
		wantRelationship := "rFoot"
		if story.Kind == "endnote" {
			wantRelationship = "rEnd"
		}
		if *story.RelationshipID != wantRelationship {
			t.Fatalf("%s relationship = %q, want %q", story.ID, *story.RelationshipID, wantRelationship)
		}
		roles[story.Kind+":"+story.NoteRole] = true
		if story.NoteRole != "content" {
			continue
		}
		labels := 0
		for _, block := range story.Blocks {
			for _, run := range block.Paragraph.Runs {
				if run.Reference != nil && run.Reference.Role == "label" {
					labels++
					if run.Reference.Kind != story.Kind || run.Reference.TargetID != story.ID {
						t.Fatalf("label does not self-bind to owning story: %#v", run.Reference)
					}
				}
			}
		}
		if labels != 1 {
			t.Fatalf("%s labels = %d, want one", story.ID, labels)
		}
	}
	for _, kind := range []string{"footnote", "endnote"} {
		for _, role := range []string{"content", "separator", "continuation-separator"} {
			if !roles[kind+":"+role] {
				t.Fatalf("missing %s %s story: %v", kind, role, roles)
			}
		}
	}
	first, err := EncodeNativeDocumentV1(doc)
	if err != nil {
		t.Fatal(err)
	}
	again, err := ExtractNativeDocumentV1(nativeNotePagePaintPackage(t))
	if err != nil {
		t.Fatal(err)
	}
	second, _ := EncodeNativeDocumentV1(again)
	if string(first) != string(second) {
		t.Fatal("native note extraction is not byte-deterministic")
	}
}

func TestExtractNativeNoteStoriesRejectsDuplicateAndAmbiguousRelationships(t *testing.T) {
	parts := transitionalNativeParts()
	parts["Custom/Notes/Foot.XML"] = strings.Replace(parts["Custom/Notes/Foot.XML"], `</w:footnotes>`, `<w:footnote w:id="1"><w:p/></w:footnote></w:footnotes>`, 1)
	if _, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts))); err == nil || !strings.Contains(err.Error(), "duplicate footnote id") {
		t.Fatalf("duplicate native note id error = %v", err)
	}

	parts = transitionalNativeParts()
	parts["Custom/_RELS/Main.XML.RELS"] = strings.Replace(parts["Custom/_RELS/Main.XML.RELS"], `</Relationships>`, `<Relationship Id="rFootDuplicate" Type="`+relBaseTransitional+`footnotes" Target="notes/foot.xml"/></Relationships>`, 1)
	if _, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts))); err == nil || !strings.Contains(err.Error(), "multiple footnotes relationships") {
		t.Fatalf("ambiguous note relationship error = %v", err)
	}

	parts = transitionalNativeParts()
	parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml", "application/xml", 1)
	if _, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts))); err == nil || !strings.Contains(err.Error(), "footnotes part") || !strings.Contains(err.Error(), "content type") {
		t.Fatalf("spoofed note content type error = %v", err)
	}
}

func TestExtractNativeNoteStoriesRejectsNonExactOPCAndNoteContainers(t *testing.T) {
	for name, mutate := range map[string]func(map[string]string){
		"relationship-extra-attribute": func(parts map[string]string) {
			parts["Custom/_RELS/Main.XML.RELS"] = strings.Replace(parts["Custom/_RELS/Main.XML.RELS"], `Id="rFoot"`, `Id="rFoot" Extra="1"`, 1)
		},
		"content-type-extra-content": func(parts map[string]string) {
			parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>`, `ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml">rogue</Override>`, 1)
		},
		"note-extra-attribute": func(parts map[string]string) {
			parts["Custom/Notes/Foot.XML"] = strings.Replace(parts["Custom/Notes/Foot.XML"], `<w:footnote w:id="1">`, `<w:footnote w:id="1" bogus="1">`, 1)
		},
		"note-direct-text": func(parts map[string]string) {
			parts["Custom/Notes/Foot.XML"] = strings.Replace(parts["Custom/Notes/Foot.XML"], `<w:footnote w:id="1">`, `<w:footnote w:id="1">rogue`, 1)
		},
	} {
		t.Run(name, func(t *testing.T) {
			parts := nativeNotePagePaintParts()
			mutate(parts)
			if _, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts))); err == nil {
				t.Fatal("non-exact OPC/note input was accepted")
			}
		})
	}
}

func TestExtractNativeNoteStoriesRequiresExactTypedInstructionSentinels(t *testing.T) {
	parts := nativeNotePagePaintParts()
	for _, partName := range []string{"Custom/Notes/Foot.XML", "Custom/Notes/End.XML"} {
		parts[partName] = strings.Replace(parts[partName], `<w:t>---</w:t>`, `<w:separator/>`, 1)
		parts[partName] = strings.Replace(parts[partName], `<w:t>continued</w:t>`, `<w:continuationSeparator/>`, 1)
	}
	doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	for _, unsupported := range doc.Unsupported {
		if unsupported.Code == "UNMODELED_NOTE_MARKUP" || unsupported.Code == "SPECIAL_NOTE_STORY" {
			t.Fatalf("exact typed instruction sentinel produced unsupported evidence: %#v", doc.Unsupported)
		}
	}

	idOnly := nativeNotePagePaintParts()
	idOnly["Custom/Notes/Foot.XML"] = strings.Replace(idOnly["Custom/Notes/Foot.XML"], ` w:type="separator"`, ``, 1)
	idOnlyDoc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(idOnly)))
	if err != nil {
		t.Fatal(err)
	}
	if !slices.ContainsFunc(idOnlyDoc.Unsupported, func(entry NativeUnsupportedCapabilityV1) bool { return entry.Code == "SPECIAL_NOTE_STORY" }) {
		t.Fatalf("id-only sentinel was incorrectly attributed to an external writer: %#v", idOnlyDoc.Unsupported)
	}

	parts = nativeNotePagePaintParts()
	parts["Custom/Notes/Foot.XML"] = strings.Replace(parts["Custom/Notes/Foot.XML"], `<w:t>---</w:t>`, `<w:separator bogus="1"/>`, 1)
	doc, err = ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	if !slices.ContainsFunc(doc.Unsupported, func(entry NativeUnsupportedCapabilityV1) bool { return entry.Code == "UNMODELED_NOTE_MARKUP" }) {
		t.Fatalf("hostile separator instruction was not fail-closed: %#v", doc.Unsupported)
	}
}

func TestExtractNativeNoteStoriesRejectsCustomMarkersAndNonDecimalContentIDs(t *testing.T) {
	parts := nativeNotePagePaintParts()
	parts["Custom/Main.XML"] = strings.ReplaceAll(parts["Custom/Main.XML"], `<w:footnoteReference w:id="1"/>`, `<w:footnoteReference w:id="1" w:customMarkFollows="1"/>`)
	parts["Custom/Main.XML"] = strings.ReplaceAll(parts["Custom/Main.XML"], `<w:endnoteReference w:id="2"/>`, `<w:endnoteReference w:id="2" w:customMarkFollows="1"/>`)
	doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	for _, block := range doc.Body.Blocks {
		if block.Paragraph == nil {
			continue
		}
		for _, run := range block.Paragraph.Runs {
			if run.Reference != nil && (run.Reference.Kind == "footnote" || run.Reference.Kind == "endnote") {
				t.Fatalf("custom note marker was modeled as an ordinary anchor: %#v", run.Reference)
			}
		}
	}
	foundUnsupported := false
	for _, unsupported := range doc.Unsupported {
		foundUnsupported = foundUnsupported || unsupported.Code == "UNRESOLVED_NOTE_REFERENCE"
	}
	if !foundUnsupported {
		t.Fatal("custom note marker did not produce a fail-closed unsupported record")
	}

	parts = nativeNotePagePaintParts()
	parts["Custom/Notes/Foot.XML"] = strings.Replace(parts["Custom/Notes/Foot.XML"], `w:id="1"`, `w:id="abc"`, 1)
	doc, err = ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	for _, story := range doc.Notes {
		if story.Kind == "footnote" && story.NoteRole == "content" {
			t.Fatalf("non-decimal native note content id was modeled: %#v", story)
		}
	}
}

func TestNativeNoteContractRejectsDuplicateNativeIDs(t *testing.T) {
	doc, err := ExtractNativeDocumentV1(nativeNotePagePaintPackage(t))
	if err != nil {
		t.Fatal(err)
	}
	for _, story := range doc.Notes {
		if story.Kind != "footnote" || story.NoteRole != "content" {
			continue
		}
		duplicate := story
		duplicate.ID += ":duplicate"
		duplicate.Blocks = []NativeBlockV1{}
		doc.Notes = append(doc.Notes, duplicate)
		break
	}
	found := false
	for _, issue := range ValidateNativeDocumentV1(doc) {
		found = found || issue.Code == "DUPLICATE_ID" && strings.HasSuffix(issue.Path, "/native_story_id")
	}
	if !found {
		t.Fatal("canonical validator accepted duplicate native footnote ids")
	}

	doc, err = ExtractNativeDocumentV1(nativeNotePagePaintPackage(t))
	if err != nil {
		t.Fatal(err)
	}
	var footRelationship string
	for i := range doc.Notes {
		if doc.Notes[i].Kind == "footnote" {
			footRelationship = *doc.Notes[i].RelationshipID
		}
	}
	for i := range doc.Notes {
		if doc.Notes[i].Kind == "endnote" {
			doc.Notes[i].RelationshipID = nativeString(footRelationship)
		}
	}
	found = false
	for _, issue := range ValidateNativeDocumentV1(doc) {
		found = found || issue.Code == "DUPLICATE_ID" && strings.HasSuffix(issue.Path, "/relationship_id")
	}
	if !found {
		t.Fatal("canonical validator accepted one relationship identity for both note kinds")
	}

	doc, err = ExtractNativeDocumentV1(nativeNotePagePaintPackage(t))
	if err != nil {
		t.Fatal(err)
	}
	for i := range doc.Notes {
		if doc.Notes[i].Kind == "footnote" && doc.Notes[i].NoteRole == "content" {
			doc.Notes[i].RelationshipID = nativeString("rFootDrifted")
			doc.Notes[i].PartName = "Custom/Notes/OtherFoot.XML"
			break
		}
	}
	issues := ValidateNativeDocumentV1(doc)
	if !slices.ContainsFunc(issues, func(issue NativeValidationIssue) bool { return strings.HasSuffix(issue.Path, "/relationship_id") }) || !slices.ContainsFunc(issues, func(issue NativeValidationIssue) bool { return strings.HasSuffix(issue.Path, "/part_name") }) {
		t.Fatalf("canonical validator accepted same-kind relationship/part drift: %#v", issues)
	}
}
