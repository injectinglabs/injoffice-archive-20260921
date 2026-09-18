package docxpatch

import "encoding/xml"

// NativeDocxEnclosedMarkerFontV1 records that a numbered paragraph's generated
// list-marker text uses the Enclosed Alphanumerics block (U+2460..U+24FF) while
// the face its font slot names is the paragraph's ordinary Latin face.
//
// ECMA-376 17.18.59 decimalEnclosedCircle generates U+2460..U+2473, and
// MS-OI29500 17.3.2.26 assigns that block to the High ANSI slot, so the marker
// takes whatever ascii/hAnsi face the level, style and w:docDefaults cascade
// resolve -- a face the package authored for Latin text and never for these
// code points. The level states no face of its own.
//
// This is a source fact, not a font recommendation: strict resolution keeps the
// resolved family and this tier reads no cmap, so it can neither prove nor
// disprove that the face contains the glyph. A read-only consumer may project
// its own declared host family over the recorded scope.
type NativeDocxEnclosedMarkerFontV1 struct {
	ScopeKind     string `json:"scope_kind"`
	ScopeID       string `json:"scope_id"`
	PartName      string `json:"part_name"`
	Path          string `json:"path"`
	SourceFamily  string `json:"source_family"`
	PackageSHA256 string `json:"package_sha256"`
}

// nativeEnclosedMarkerFonts produces one fact per numbered paragraph whose
// resolved marker text contains an Enclosed Alphanumerics rune and whose marker
// resolved a font family. A marker that resolved no family is already refused
// upstream and is not re-pointed here.
func nativeEnclosedMarkerFonts(data []byte) ([]NativeDocxEnclosedMarkerFontV1, error) {
	r, err := newNativeLayoutResolver(data, NativeExtractionOptions{})
	if err != nil {
		return nil, err
	}
	layout, err := r.resolve()
	if err != nil {
		return nil, err
	}
	// The gate is the resolver's own diagnostic, not a second reading of the
	// marker: exactly the scopes it recorded are the scopes a consumer may
	// approximate, so a marker it refused for any other reason -- including the
	// w:hint="eastAsia" escape -- produces no fact here.
	recorded := map[string]bool{}
	for _, diagnostic := range layout.Diagnostics {
		if diagnostic.Code == "ENCLOSED_NUMBER_MARKER_FONT_PRESERVED" {
			recorded[diagnostic.ScopeID] = true
		}
	}
	eligible := map[string]string{}
	for _, paragraph := range layout.Paragraphs {
		if paragraph.Numbering == nil || paragraph.Numbering.Marker.FontFamily == nil || !recorded[paragraph.ParagraphID] {
			continue
		}
		eligible[paragraph.ParagraphID] = *paragraph.Numbering.Marker.FontFamily
	}
	if len(eligible) == 0 {
		return nil, nil
	}
	var facts []NativeDocxEnclosedMarkerFontV1
	consider := func(p *NativeParagraphV1) {
		family, ok := eligible[p.ID]
		if !ok {
			return
		}
		node := r.nodeForAnchor(p.Anchor)
		if node == nil || node.Name != (xml.Name{Space: r.wordNS, Local: "p"}) {
			return
		}
		facts = append(facts, NativeDocxEnclosedMarkerFontV1{ScopeKind: "numbering-marker", ScopeID: p.ID, PartName: p.Anchor.PartName, Path: p.Anchor.Path, SourceFamily: family, PackageSHA256: r.doc.Source.PackageSHA256})
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
