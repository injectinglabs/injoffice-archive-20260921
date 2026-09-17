package docxpatch

import (
	"bytes"
	"encoding/xml"
)

// NativeDocxAbsentFontFamilyV1 records that a scope resolves no font family
// because the package states no font selection anywhere at all -- no w:rFonts
// element in any part, so no ascii/hAnsi/theme slot, no style layer and no
// w:docDefaults run face exists to be read. ECMA-376 17.3.2.26 makes w:rFonts
// optional and leaves the face to the application when it is omitted, so the
// omission is a source fact.
//
// Source absence is not a font recommendation. Strict resolution still leaves
// these scopes without a family; a read-only consumer may project its own
// declared host default over the proven omission.
type NativeDocxAbsentFontFamilyV1 struct {
	ScopeKind     string `json:"scope_kind"`
	ScopeID       string `json:"scope_id"`
	PartName      string `json:"part_name"`
	Path          string `json:"path"`
	PackageSHA256 string `json:"package_sha256"`
}

// nativeAbsentFontFamilies produces evidence only for a package that selects no
// font anywhere. The gate is deliberately whole-package: a package that states
// any font selection, even one strict resolution refuses to read (an unmodelled
// slot, a script-dependent Latin pair, an unresolved theme face), keeps using
// what it states and produces no fact here.
func nativeAbsentFontFamilies(data []byte) ([]NativeDocxAbsentFontFamilyV1, error) {
	r, err := newNativeLayoutResolver(data, NativeExtractionOptions{})
	if err != nil {
		return nil, err
	}
	layout, err := r.resolve()
	if err != nil {
		return nil, err
	}
	if !r.absentPackageFontSelection(layout) {
		return nil, nil
	}
	paragraphs := map[string]NativeResolvedParagraphV1{}
	for _, p := range layout.Paragraphs {
		paragraphs[p.ParagraphID] = p
	}
	runs := map[string]NativeResolvedRunV1{}
	for _, run := range layout.Runs {
		runs[run.RunID] = run
	}
	var facts []NativeDocxAbsentFontFamilyV1
	add := func(kind, id string, anchor NativeSourceAnchorV1) {
		facts = append(facts, NativeDocxAbsentFontFamilyV1{ScopeKind: kind, ScopeID: id, PartName: anchor.PartName, Path: anchor.Path, PackageSHA256: r.doc.Source.PackageSHA256})
	}
	// A numbered paragraph also needs a marker face this evidence does not cover,
	// so it is skipped whole: leaving it unshaped keeps its disclosure honest
	// rather than painting its text without its marker.
	consider := func(p *NativeParagraphV1) {
		resolved, ok := paragraphs[p.ID]
		if !ok || resolved.Numbering != nil {
			return
		}
		node := r.nodeForAnchor(p.Anchor)
		if node == nil || node.Name != (xml.Name{Space: r.wordNS, Local: "p"}) {
			return
		}
		add("paragraph-mark", p.ID, p.Anchor)
		for _, run := range p.Runs {
			if _, ok := runs[run.ID]; !ok {
				continue
			}
			add("run", run.ID, run.Anchor)
		}
	}
	stories := append([]NativeStoryV1{r.doc.Body}, r.doc.Headers...)
	stories = append(stories, r.doc.Footers...)
	stories = append(stories, r.doc.Notes...)
	stories = append(stories, r.doc.CommentStories...)
	for _, story := range stories {
		for _, block := range story.Blocks {
			if block.Paragraph != nil {
				consider(block.Paragraph)
				continue
			}
			if block.Table == nil {
				continue
			}
			for _, row := range block.Table.Rows {
				for _, cell := range row.Cells {
					for index := range cell.Paragraphs {
						consider(&cell.Paragraphs[index])
					}
				}
			}
		}
	}
	if len(facts) > 1000 {
		return nil, nil
	}
	return facts, nil
}

// absentPackageFontSelection proves the package selects no font at all: no part
// carries a w:rFonts element, and strict resolution consequently resolved no
// family in any run, paragraph mark or list marker. Both halves are required --
// the first is the source fact, the second refuses any package where some other
// path still produced a face.
func (r *nativeLayoutResolver) absentPackageFontSelection(layout *NativeResolvedLayoutInputV1) bool {
	for _, run := range layout.Runs {
		if run.Properties.FontFamily != nil {
			return false
		}
	}
	for _, paragraph := range layout.Paragraphs {
		if paragraph.ParagraphMarkProperties.FontFamily != nil {
			return false
		}
		if paragraph.Numbering != nil && paragraph.Numbering.Marker.FontFamily != nil {
			return false
		}
	}
	// Only the parts that can supply a face to the stories this evidence covers
	// are scanned: the main part, the style/numbering/theme/font-table parts it
	// cascades through, and every header, footer, note and comment story. The
	// glossary document (ECMA-376 17.12.6) is a separate document with its own
	// stories and styles, so its w:rFonts states nothing about this one and is
	// deliberately outside the scan. Within those parts the scan is a byte-level
	// substring, deliberately broader than an element walk: any occurrence at
	// all, including in markup this resolver never reads, disables the evidence.
	scanned := map[string]bool{r.mainPart: true}
	for _, part := range []*string{r.parts.StylesPart, r.parts.NumberingPart, r.parts.ThemePart, r.parts.FontTablePart} {
		if part != nil {
			scanned[*part] = true
		}
	}
	stories := append([]NativeStoryV1{r.doc.Body}, r.doc.Headers...)
	stories = append(stories, r.doc.Footers...)
	stories = append(stories, r.doc.Notes...)
	stories = append(stories, r.doc.CommentStories...)
	for _, story := range stories {
		scanned[story.PartName] = true
	}
	for name := range scanned {
		if bytes.Contains(r.pkg.files[name], []byte("rFonts")) {
			return false
		}
	}
	return true
}
