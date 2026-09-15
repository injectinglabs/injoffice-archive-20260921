package docxpatch

// NativeDocxLatinFontFallbackV1 is approximate-preview evidence only. Strict
// resolution leaves the run or paragraph mark without a font family because an
// empty w:eastAsia/w:cs slot keeps the whole w:rFonts unresolved; the fact
// records the authored ascii/hAnsi face that resolution would otherwise have
// selected so a read-only consumer may project it under a declared policy.
type NativeDocxLatinFontFallbackV1 struct {
	ScopeKind     string `json:"scope_kind"`
	ScopeID       string `json:"scope_id"`
	PartName      string `json:"part_name"`
	Path          string `json:"path"`
	FontFamily    string `json:"font_family"`
	PackageSHA256 string `json:"package_sha256"`
}

// nativeLatinFontFallbacks diffs the strict resolution against a second pass
// that only tolerates empty script slots. Every other resolver decision is
// shared, so a face that appears only in the tolerant pass is attributable to
// the empty slot alone. Strict output is never touched.
func nativeLatinFontFallbacks(data []byte) ([]NativeDocxLatinFontFallbackV1, error) {
	strict, err := newNativeLayoutResolver(data, NativeExtractionOptions{})
	if err != nil {
		return nil, err
	}
	strictLayout, err := strict.resolve()
	if err != nil {
		return nil, err
	}
	tolerant, err := newNativeLayoutResolver(data, NativeExtractionOptions{tolerateEmptyScriptSlots: true})
	if err != nil {
		return nil, err
	}
	tolerantLayout, err := tolerant.resolve()
	if err != nil {
		return nil, err
	}
	strictRuns := map[string]*string{}
	for _, run := range strictLayout.Runs {
		strictRuns[run.RunID] = run.Properties.FontFamily
	}
	strictMarks := map[string]*string{}
	for _, paragraph := range strictLayout.Paragraphs {
		strictMarks[paragraph.ParagraphID] = paragraph.ParagraphMarkProperties.FontFamily
	}
	anchors := map[string]NativeSourceAnchorV1{}
	stories := append([]NativeStoryV1{strict.doc.Body}, strict.doc.Headers...)
	stories = append(stories, strict.doc.Footers...)
	stories = append(stories, strict.doc.Notes...)
	stories = append(stories, strict.doc.CommentStories...)
	var index func(paragraphs []NativeParagraphV1)
	index = func(paragraphs []NativeParagraphV1) {
		for _, paragraph := range paragraphs {
			anchors[paragraph.ID] = paragraph.Anchor
			for _, run := range paragraph.Runs {
				anchors[run.ID] = run.Anchor
			}
		}
	}
	for _, story := range stories {
		for _, block := range story.Blocks {
			if block.Paragraph != nil {
				index([]NativeParagraphV1{*block.Paragraph})
			}
			if block.Table != nil {
				for _, row := range block.Table.Rows {
					for _, cell := range row.Cells {
						index(cell.Paragraphs)
					}
				}
			}
		}
	}
	var facts []NativeDocxLatinFontFallbackV1
	add := func(kind, id string, family string) {
		anchor, ok := anchors[id]
		if !ok || !nativeBoundedResolvedString(family, 256) {
			return
		}
		facts = append(facts, NativeDocxLatinFontFallbackV1{ScopeKind: kind, ScopeID: id, PartName: anchor.PartName, Path: anchor.Path, FontFamily: family, PackageSHA256: strict.doc.Source.PackageSHA256})
	}
	for _, run := range tolerantLayout.Runs {
		if strictFamily, ok := strictRuns[run.RunID]; ok && strictFamily == nil && run.Properties.FontFamily != nil {
			add("run", run.RunID, *run.Properties.FontFamily)
		}
	}
	for _, paragraph := range tolerantLayout.Paragraphs {
		if strictFamily, ok := strictMarks[paragraph.ParagraphID]; ok && strictFamily == nil && paragraph.ParagraphMarkProperties.FontFamily != nil {
			add("paragraph-mark", paragraph.ParagraphID, *paragraph.ParagraphMarkProperties.FontFamily)
		}
	}
	if len(facts) > 1000 {
		return nil, nil
	}
	return facts, nil
}
