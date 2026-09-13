package docxpatch

import "encoding/xml"

// This sidecar contains no native table or mutation model. Only an exact
// one-cell nested table with direct unformatted paragraphs enters resolution.
type NativePartialNestedTextRunV1 struct {
	ID     string               `json:"id"`
	Anchor NativeSourceAnchorV1 `json:"anchor"`
	Text   string               `json:"text"`
}
type NativePartialNestedTextParagraphV1 struct {
	ID     string                         `json:"id"`
	Anchor NativeSourceAnchorV1           `json:"anchor"`
	Runs   []NativePartialNestedTextRunV1 `json:"runs"`
}
type NativePartialNestedTextV1 struct {
	Owner          NativePartialNestedTableV1           `json:"owner"`
	StyleID        string                               `json:"style_id"`
	Paragraphs     []NativePartialNestedTextParagraphV1 `json:"paragraphs"`
	ResolvedLayout *NativeResolvedLayoutInputV1         `json:"resolved_layout"`
}

func inspectNativePartialNestedText(data []byte, doc *NativeDocumentV1, nested *NativePartialNestedTablesV1, contexts []NativePartialTableTextContextV1) ([]NativePartialNestedTextV1, error) {
	if nested == nil || len(contexts) == 0 {
		return nil, nil
	}
	resolver, err := newNativeLayoutResolver(data, NativeExtractionOptions{})
	if err != nil {
		return nil, err
	}
	ns := resolver.wordNS
	raw := resolver.pkg.files[doc.Source.MainPart]
	anchor := func(n *nativeXMLNode) NativeSourceAnchorV1 {
		return NativeSourceAnchorV1{PartName: doc.Source.MainPart, Path: n.Path, StartByte: nativeInt64(n.Start), EndByte: nativeInt64(n.End), XMLSHA256: nativeSHA(raw[n.Start:n.End])}
	}
	out := []NativePartialNestedTextV1{}
	units := 0
	initialDiagnostics := len(resolver.diagnostics)
	for _, owner := range nested.Items {
		if len(out) >= 32 {
			break
		}
		var context *NativePartialTableTextContextV1
		for i := range contexts {
			if contexts[i].TableID == owner.TableID {
				context = &contexts[i]
			}
		}
		if context == nil || len(context.StyleChain) == 0 {
			continue
		}
		styleID := context.StyleChain[len(context.StyleChain)-1].StyleID
		n := resolver.nodeForAnchor(owner.Anchor)
		if n == nil || !nativeExactContainer(n) || len(n.Children) != 3 {
			continue
		}
		props, grid, row := n.Children[0], n.Children[1], n.Children[2]
		if props.Name != (xml.Name{Space: ns, Local: "tblPr"}) || !nativeExactContainer(props) || len(props.Children) != 3 || grid.Name != (xml.Name{Space: ns, Local: "tblGrid"}) || !nativeExactContainer(grid) || len(grid.Children) != 1 || row.Name != (xml.Name{Space: ns, Local: "tr"}) || !nativeExactContainer(row) || len(row.Children) != 1 {
			continue
		}
		style, width, look := props.Children[0], props.Children[1], props.Children[2]
		value, _ := nativeAttr(style, ns, "val")
		if style.Name != (xml.Name{Space: ns, Local: "tblStyle"}) || !nativeExactLeaf(style, xml.Name{Space: ns, Local: "val"}) || value != styleID || width.Name != (xml.Name{Space: ns, Local: "tblW"}) || !nativePartialTextWidth(width, ns) || look.Name != (xml.Name{Space: ns, Local: "tblLook"}) || !nativePartialTextLook(look, ns) {
			continue
		}
		col := grid.Children[0]
		colWidth, ok := nativePositiveInt64Attr(col, ns, "w")
		if col.Name != (xml.Name{Space: ns, Local: "gridCol"}) || !nativeExactLeaf(col, xml.Name{Space: ns, Local: "w"}) || !ok || colWidth > 31680 {
			continue
		}
		cell := row.Children[0]
		if cell.Name != (xml.Name{Space: ns, Local: "tc"}) || !nativeExactContainer(cell) || len(cell.Children) == 0 || len(cell.Children) > 16 {
			continue
		}
		ps := []NativeParagraphV1{}
		item := NativePartialNestedTextV1{Owner: owner, StyleID: styleID, Paragraphs: []NativePartialNestedTextParagraphV1{}}
		exact := true
		itemUnits := 0
		for _, p := range cell.Children {
			if p.Name != (xml.Name{Space: ns, Local: "p"}) || !nativeExactContainer(p) || len(p.Children) > 32 {
				exact = false
				break
			}
			id := nativeStableID("nested-paragraph", doc.Source.MainPart, p.Path, "")
			paragraph := NativeParagraphV1{ID: id, Anchor: anchor(p), Properties: &NativeParagraphPropertiesV1{}, Runs: []NativeRunV1{}}
			entry := NativePartialNestedTextParagraphV1{ID: id, Anchor: paragraph.Anchor, Runs: []NativePartialNestedTextRunV1{}}
			for _, r := range p.Children {
				if r.Name != (xml.Name{Space: ns, Local: "r"}) || !nativeExactContainer(r) || len(r.Children) != 1 {
					exact = false
					break
				}
				text := r.Children[0]
				// No run properties, fields, revisions, hidden wrappers or alternate content.
				if text.Name != (xml.Name{Space: ns, Local: "t"}) || len(text.Attrs) != 0 || len(text.Children) != 0 {
					exact = false
					break
				}
				itemUnits += len(text.Text)
				if itemUnits > 32768 {
					exact = false
					break
				}
				rid := nativeStableID("nested-run", doc.Source.MainPart, text.Path, "")
				a := anchor(text)
				paragraph.Runs = append(paragraph.Runs, NativeRunV1{Kind: "text", ID: rid, Anchor: a, Text: nativeString(text.Text)})
				entry.Runs = append(entry.Runs, NativePartialNestedTextRunV1{rid, a, text.Text})
			}
			if !exact {
				break
			}
			ps = append(ps, paragraph)
			item.Paragraphs = append(item.Paragraphs, entry)
		}
		if !exact || units+itemUnits > 32768 {
			continue
		}

		chain := resolver.styleChain("table", styleID, owner.TableID)
		resolved := &NativeResolvedLayoutInputV1{Protocol: NativeDOCXResolvedLayoutProtocol, Version: 1, DocumentID: doc.DocumentID, Revision: doc.Revision, SourceParts: resolver.parts, Paragraphs: []NativeResolvedParagraphV1{}, Runs: []NativeResolvedRunV1{}, Tables: []NativeResolvedTableV1{}, Fonts: []NativeResolvedFontV1{}, Diagnostics: []NativeResolutionDiagnosticV1{}}
		for i := range ps {
			resolver.resolveParagraph(&ps[i], resolved, newNativeNumberingState(), chain)
		}
		if len(resolver.diagnostics) != initialDiagnostics {
			continue
		}
		for _, p := range resolved.Paragraphs {
			if p.Numbering != nil || p.ParagraphMarkProperties.Hidden != nil && *p.ParagraphMarkProperties.Hidden {
				exact = false
			}
		}
		for _, r := range resolved.Runs {
			if r.Properties.Hidden != nil && *r.Properties.Hidden {
				exact = false
			}
		}
		if !exact {
			continue
		}
		if _, err := EncodeNativeResolvedLayoutInputV1(resolved); err != nil {
			continue
		}
		item.ResolvedLayout = resolved
		out = append(out, item)
		units += itemUnits
	}
	if len(out) == 0 {
		return nil, nil
	}
	return out, nil
}
