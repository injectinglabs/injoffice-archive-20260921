package docxpatch

import (
	"strings"
	"testing"
)

// overrideOnlyNativeParts is the shape of a package whose writer declared every
// part with an Override and never emitted a single Default Extension. Real
// producers ship this: it is what both corpus packages behind this test look
// like, and it makes any stored item the writer forgot to declare undeclarable
// by extension alone.
func overrideOnlyNativeParts() map[string]string {
	parts := cloneNativeParts(transitionalNativeParts())
	parts["[Content_Types].xml"] = `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
		`<Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
		`<Override PartName="/custom/_rels/main.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
		`<Override PartName="/custom/main.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
		`<Override PartName="/custom/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
		`<Override PartName="/custom/stories/headera.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>` +
		`<Override PartName="/custom/stories/footera.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>` +
		`<Override PartName="/custom/notes/foot.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>` +
		`<Override PartName="/custom/notes/end.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>` +
		`<Override PartName="/custom/notes/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>` +
		`<Override PartName="/custom/media/image.png" ContentType="image/png"/>` +
		`</Types>`
	return parts
}

func nativePassthroughNames(doc *NativeDocumentV1) map[string]string {
	names := map[string]string{}
	for _, part := range doc.PassthroughParts {
		names[part.PartName] = part.ContentType
	}
	return names
}

func TestNativeExtractOverrideOnlyPackage(t *testing.T) {
	doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(overrideOnlyNativeParts())))
	if err != nil {
		t.Fatalf("a package declared entirely by Override must extract: %v", err)
	}
	if len(doc.Headers) != 1 || len(doc.Footers) != 1 || len(doc.Notes) != 3 {
		t.Fatalf("declared parts must resolve: headers=%d footers=%d notes=%d", len(doc.Headers), len(doc.Footers), len(doc.Notes))
	}
	resolved := ""
	for _, run := range doc.Body.Blocks[0].Paragraph.Runs {
		if run.Drawing != nil && run.Drawing.MediaPart != nil {
			resolved = *run.Drawing.MediaPart
		}
	}
	if resolved != "Custom/Media/image.PNG" {
		t.Fatalf("a declared image part must still resolve: %q", resolved)
	}
}

// A stored ZIP item that no Default and no Override maps has no media type, and
// ECMA-376 Part 2 6.2.3 makes a media type constitutive of a part, so the item
// is not a part of this package. Editor leftovers are exactly this: a crashed
// or scripted save leaves "styles.xml~" backups and stray files behind, Word
// and LibreOffice open the documents around them, and the declared parts are
// untouched.
func TestNativeExtractIgnoresUndeclaredPackageItems(t *testing.T) {
	parts := overrideOnlyNativeParts()
	parts["Custom/Main.XML~"] = parts["Custom/Main.XML"]
	parts["Custom/Styles.XML~"] = parts["Custom/Styles.XML"]
	parts["Custom/_RELS/Main.XML.RELS~"] = parts["Custom/_RELS/Main.XML.RELS"]
	parts["8980.xml"] = "======= a patch someone zipped in by accident\n--- a/file\n"
	doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatalf("an undeclared stored item must not refuse the package: %v", err)
	}
	if len(doc.Headers) != 1 || len(doc.Footers) != 1 || len(doc.Notes) != 3 || len(doc.Comments) != 1 {
		t.Fatalf("declared parts must survive: headers=%d footers=%d notes=%d comments=%d", len(doc.Headers), len(doc.Footers), len(doc.Notes), len(doc.Comments))
	}
	var text strings.Builder
	for _, run := range doc.Body.Blocks[0].Paragraph.Runs {
		if run.Text != nil {
			text.WriteString(*run.Text)
		}
	}
	if !strings.Contains(text.String(), "Hello") {
		t.Fatalf("the body must still read: %q", text.String())
	}
	passthrough := nativePassthroughNames(doc)
	for _, name := range []string{"Custom/Main.XML~", "Custom/Styles.XML~", "Custom/_RELS/Main.XML.RELS~", "8980.xml"} {
		if contentType, present := passthrough[name]; present {
			t.Fatalf("an undeclared item is not a part and must not be carried as one: %q = %q", name, contentType)
		}
	}
	if len(passthrough) == 0 {
		t.Fatal("declared passthrough parts must still be reported")
	}
}

// A relationship resolves to a part or to nothing. Its target being stored but
// undeclared changes neither: reading those bytes would require inventing the
// media type that decides how they are parsed, which is the one fact the
// package never stated. The extension and the magic bytes both say PNG here,
// and neither is allowed to stand in for a declaration.
func TestNativeExtractRefusesToInferAnUndeclaredMediaType(t *testing.T) {
	parts := overrideOnlyNativeParts()
	parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `<Override PartName="/custom/media/image.png" ContentType="image/png"/>`, "", 1)
	parts["Custom/Media/image.PNG"] = "\x89PNG\r\n\x1a\n" + parts["Custom/Media/image.PNG"]
	doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatalf("an undeclared image must not refuse the document around it: %v", err)
	}
	for _, run := range doc.Body.Blocks[0].Paragraph.Runs {
		if run.Drawing != nil && run.Drawing.MediaPart != nil {
			t.Fatalf("an undeclared media part must not resolve: %q", *run.Drawing.MediaPart)
		}
	}
	for name, contentType := range nativePassthroughNames(doc) {
		if strings.HasPrefix(strings.ToLower(contentType), "image/") || strings.Contains(strings.ToLower(name), "image.png") {
			t.Fatalf("no media type may be inferred from an extension or from magic bytes: %q = %q", name, contentType)
		}
	}
}

// The tolerance above is a rule about items the package never declared, never a
// licence to read an undeclared part because its name looks familiar. A main
// document part with no declared media type is not a Word main part, and the
// package must still fail closed rather than paint whatever the bytes contain.
func TestNativeExtractRefusesUndeclaredRequiredParts(t *testing.T) {
	for _, test := range []struct {
		name    string
		remove  string
		message string
	}{
		{
			name:    "main document part",
			remove:  `<Override PartName="/custom/main.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>`,
			message: "stores with no declared content type",
		},
		{
			name:    "root relationship part",
			remove:  `<Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`,
			message: "root relationship part _rels/.rels is missing",
		},
		{
			// Undeclaring the main part's relationships part does not quietly
			// drop the stories it bound: the section still names the header
			// relationship, and nothing resolves it.
			name:    "main relationship part",
			remove:  `<Override PartName="/custom/_rels/main.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`,
			message: `section relationship "rHeader" does not resolve to a modeled header`,
		},
	} {
		t.Run(test.name, func(t *testing.T) {
			parts := overrideOnlyNativeParts()
			parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], test.remove, "", 1)
			doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err == nil {
				t.Fatalf("an undeclared required part must refuse: %#v", doc.Source)
			}
			if !strings.Contains(err.Error(), test.message) {
				t.Fatalf("refusal must name the defect: %v", err)
			}
		})
	}
}
