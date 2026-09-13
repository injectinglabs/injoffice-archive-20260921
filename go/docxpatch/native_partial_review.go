package docxpatch

import "encoding/xml"

// Review evidence carries source metadata and modeled run IDs, never deletion
// text. It does not accept/reject changes or remove original diagnostics.
type NativePartialReviewChangeV1 struct {
	PackageSHA256 string               `json:"package_sha256"`
	ParagraphID   string               `json:"paragraph_id"`
	DiagnosticID  string               `json:"diagnostic_id"`
	Anchor        NativeSourceAnchorV1 `json:"anchor"`
	Kind          string               `json:"kind"`
	RevisionID    string               `json:"revision_id"`
	Author        string               `json:"author"`
	CreatedAt     string               `json:"created_at"`
	RunIDs        []string             `json:"run_ids"`
}
type NativePartialReviewV1 struct {
	Items        []NativePartialReviewChangeV1 `json:"items"`
	OmittedCount int                           `json:"omitted_count"`
}

func inspectNativePartialReview(data []byte, doc *NativeDocumentV1) (*NativePartialReviewV1, error) {
	pkg, err := openNativeDOCXPackage(data)
	if err != nil {
		return nil, err
	}
	main, strict, err := pkg.officeDocumentPart()
	if err != nil {
		return nil, err
	}
	ns := wordMLTransitional
	if strict {
		ns = wordMLStrict
	}
	root, err := parseNativeXML(main, pkg.files[main])
	if err != nil {
		return nil, err
	}
	nodes := map[string]*nativeXMLNode{}
	var visit func(*nativeXMLNode)
	visit = func(n *nativeXMLNode) {
		nodes[n.Path] = n
		for _, c := range n.Children {
			visit(c)
		}
	}
	visit(root)
	paragraphs := map[string]*NativeParagraphV1{}
	for i := range doc.Body.Blocks {
		b := &doc.Body.Blocks[i]
		if b.Paragraph != nil {
			paragraphs[b.Paragraph.ID] = b.Paragraph
		}
		if b.Table != nil {
			for _, row := range b.Table.Rows {
				for _, cell := range row.Cells {
					for i := range cell.Paragraphs {
						p := &cell.Paragraphs[i]
						paragraphs[p.ID] = p
					}
				}
			}
		}
	}
	out := &NativePartialReviewV1{Items: []NativePartialReviewChangeV1{}}
	for _, d := range doc.Unsupported {
		if d.Anchor == nil || d.Anchor.PartName != main || (d.Code != "WRAPPED_RUN_MARKUP" && d.Code != "UNMODELED_PARAGRAPH_CONTENT") {
			continue
		}
		n := nodes[d.Anchor.Path]
		p := paragraphs[d.ScopeID]
		if n == nil || p == nil || n.Name.Space != ns || n.parent == nil || n.parent.Path != p.Anchor.Path {
			continue
		}
		kind := map[string]string{"ins": "insertion", "del": "deletion", "moveTo": "move-to", "moveFrom": "move-from"}[n.Name.Local]
		if kind == "" {
			continue
		}
		expected := "UNMODELED_PARAGRAPH_CONTENT"
		if kind == "insertion" || kind == "move-to" {
			expected = "WRAPPED_RUN_MARKUP"
		}
		if d.Code != expected {
			continue
		}
		id, idOK := nativeAttr(n, ns, "id")
		author, authorOK := nativeAttr(n, ns, "author")
		date, _ := nativeAttr(n, ns, "date")
		if len(out.Items) >= 128 || !nativeExactContainer(n, xml.Name{Space: ns, Local: "id"}, xml.Name{Space: ns, Local: "author"}, xml.Name{Space: ns, Local: "date"}) || !idOK || id == "" || !authorOK || len(id) > 128 || len(author) > 1024 || len(date) > 128 {
			out.OmittedCount++
			continue
		}
		item := NativePartialReviewChangeV1{doc.Source.PackageSHA256, p.ID, d.ID, *d.Anchor, kind, id, author, date, []string{}}
		// Only direct ordinary insertion runs can later be visibility-qualified.
		// No descendants of nested revisions, fields, drawings or controls enter it.
		if kind == "insertion" && len(n.Children) > 0 && len(n.Children) <= 128 {
			paths := map[string]bool{}
			exact := true
			for _, r := range n.Children {
				if r.Name != (xml.Name{Space: ns, Local: "r"}) || !nativeExactContainer(r) || len(r.Children) == 0 {
					exact = false
					break
				}
				for i, c := range r.Children {
					if c.Name == (xml.Name{Space: ns, Local: "rPr"}) && i == 0 {
						continue
					}
					if c.Name != (xml.Name{Space: ns, Local: "t"}) || len(c.Children) > 0 {
						exact = false
						break
					}
					for _, a := range c.Attrs {
						if a.Name != (xml.Name{Space: "http://www.w3.org/XML/1998/namespace", Local: "space"}) || a.Value != "preserve" {
							exact = false
						}
					}
					paths[c.Path] = true
				}
			}
			ids := []string{}
			for _, r := range p.Runs {
				if paths[r.Anchor.Path] {
					if r.Kind != "text" {
						exact = false
					}
					ids = append(ids, r.ID)
				}
			}
			if exact && len(ids) == len(paths) && len(ids) > 0 && len(ids) <= 128 {
				item.RunIDs = ids
			}
		}
		out.Items = append(out.Items, item)
	}
	if len(out.Items) == 0 && out.OmittedCount == 0 {
		return nil, nil
	}
	return out, nil
}
