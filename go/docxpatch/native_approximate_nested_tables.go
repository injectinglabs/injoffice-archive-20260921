package docxpatch

import (
	"encoding/xml"
	"fmt"
	"sort"
	"strconv"
	"strings"
	"unicode/utf16"
)

// Read-only sidecar for the explicitly labeled approximate page preview.
//
// The strict extractor keeps refusing tables nested inside table cells
// (NESTED_TABLE_OR_CELL_MARKUP) and models only the cell's direct paragraphs.
// This inspection joins those very diagnostics back to their source nodes and
// describes one level of nested tables through the ordinary table/paragraph
// extractor and style resolver, so the approximate preview can lay the inner
// table out inside the containing cell's content box. Deeper nesting, merged
// cells and structural refusals stay omitted with a declared reason. It never
// changes source bytes, native extraction, editing authority or pagination.
const NativeApproximateNestedTablesProtocol = "injoffice.docx.approximate-nested-tables"
const NativeApproximateNestedTablePolicy = "docx.approximate-nested-table-preview-v1"
const nativeApproximateNestedTableItemLimit = 64
const nativeApproximateNestedTableRowLimit = 256
const nativeApproximateNestedTableCellLimit = 4096
const nativeApproximateNestedTableTextLimit = 200000

type NativeApproximateNestedTableV1 struct {
	ID            string               `json:"id"`
	TableID       string               `json:"table_id"`
	CellID        string               `json:"cell_id"`
	DiagnosticIDs []string             `json:"diagnostic_ids"`
	Anchor        NativeSourceAnchorV1 `json:"anchor"`
	// Direct w:p siblings before the nested w:tbl inside the containing cell:
	// the index in the cell's modeled paragraphs at which the inner table sits.
	PrecedingParagraphs int    `json:"preceding_paragraphs"`
	Status              string `json:"status"`
	Reason              string `json:"reason,omitempty"`
	// Inner table with read-only ids derived from the item id; direct
	// properties only. Style-derived geometry and borders follow separately
	// so the preview merges them exactly like the resolver would.
	Table               *NativeTableV1                 `json:"table,omitempty"`
	Geometry            *NativeResolvedTableGeometryV1 `json:"geometry,omitempty"`
	StyleBorders        *NativeTableBordersV1          `json:"style_borders,omitempty"`
	StyleCellShadingRGB *string                        `json:"style_cell_shading_rgb,omitempty"`
	// firstRow conditional cell fill when the inner tblLook enables the region.
	FirstRowCellShadingRGB *string `json:"first_row_cell_shading_rgb,omitempty"`
	// Approximate geometry of the containing table from its own style cascade,
	// for outer tables whose strict geometry resolution refused (active look).
	OuterGeometry      *NativeResolvedTableGeometryV1 `json:"outer_geometry,omitempty"`
	ResolvedParagraphs  []NativeResolvedParagraphV1    `json:"resolved_paragraphs"`
	ResolvedRuns        []NativeResolvedRunV1          `json:"resolved_runs"`
	OmittedRuns         int                            `json:"omitted_runs"`
	Notes               []string                       `json:"notes,omitempty"`
}

type NativeApproximateNestedTablesV1 struct {
	Protocol      string                           `json:"protocol"`
	Version       int                              `json:"version"`
	Policy        string                           `json:"policy"`
	PackageSHA256 string                           `json:"package_sha256"`
	PartSHA256    string                           `json:"part_sha256"`
	Items         []NativeApproximateNestedTableV1 `json:"items"`
	OmittedCount  int                              `json:"omitted_count"`
}

type nativeApproximateNestedContext struct {
	resolver  *nativeLayoutResolver
	ns        string
	main      string
	raw       []byte
	nodes     map[string]*nativeXMLNode
	textUnits int
}

// InspectNativeApproximateNestedTablesV1 returns nil when no body table cell
// carries a refused nested table. Callers must treat every item as approximate
// evidence; the strict document and its diagnostics are unchanged.
func InspectNativeApproximateNestedTablesV1(data []byte) (*NativeApproximateNestedTablesV1, error) {
	if len(data) == 0 || len(data) > NativeDOCXMaxPackageBytes {
		return nil, fmt.Errorf("approximate nested tables package size must be 1..%d bytes", NativeDOCXMaxPackageBytes)
	}
	data = append([]byte(nil), data...)
	resolver, err := newNativeLayoutResolver(data, NativeExtractionOptions{})
	if err != nil {
		return nil, err
	}
	doc := resolver.doc
	ns, main := resolver.wordNS, resolver.mainPart
	raw := resolver.pkg.files[main]
	context := &nativeApproximateNestedContext{resolver: resolver, ns: ns, main: main, raw: raw, nodes: map[string]*nativeXMLNode{}}
	var visit func(*nativeXMLNode)
	visit = func(n *nativeXMLNode) {
		context.nodes[n.Path] = n
		for _, c := range n.Children {
			visit(c)
		}
	}
	visit(resolver.mainRoot)
	diagnostics := map[string][]NativeUnsupportedCapabilityV1{}
	for _, d := range doc.Unsupported {
		if d.Code == "NESTED_TABLE_OR_CELL_MARKUP" && d.Anchor != nil && d.Anchor.PartName == main {
			diagnostics[d.ScopeID+"\x00"+d.Anchor.Path] = append(diagnostics[d.ScopeID+"\x00"+d.Anchor.Path], d)
		}
	}
	out := &NativeApproximateNestedTablesV1{Protocol: NativeApproximateNestedTablesProtocol, Version: 1, Policy: NativeApproximateNestedTablePolicy, PackageSHA256: doc.Source.PackageSHA256, PartSHA256: nativeSHA(raw), Items: []NativeApproximateNestedTableV1{}}
	for _, block := range doc.Body.Blocks {
		table := block.Table
		if table == nil {
			continue
		}
		for _, row := range table.Rows {
			for _, cell := range row.Cells {
				owner := context.nodes[cell.Anchor.Path]
				if owner == nil || owner.Name != (xml.Name{Space: ns, Local: "tc"}) {
					continue
				}
				preceding, ordinal := 0, 0
				for _, child := range owner.Children {
					if child.Name == (xml.Name{Space: ns, Local: "p"}) {
						preceding++
						continue
					}
					if child.Name != (xml.Name{Space: ns, Local: "tbl"}) {
						continue
					}
					ordinal++
					anchor := NativeSourceAnchorV1{PartName: main, Path: child.Path, StartByte: nativeInt64(child.Start), EndByte: nativeInt64(child.End), XMLSHA256: nativeSHA(raw[child.Start:child.End])}
					joined := []string{}
					for _, d := range diagnostics[table.ID+"\x00"+child.Path] {
						if d.Anchor.XMLSHA256 == anchor.XMLSHA256 && d.Anchor.StartByte != nil && d.Anchor.EndByte != nil && *d.Anchor.StartByte == child.Start && *d.Anchor.EndByte == child.End {
							joined = append(joined, d.ID)
						}
					}
					if len(joined) == 0 {
						// Modeled or otherwise owned content is never re-described here.
						continue
					}
					if len(out.Items) >= nativeApproximateNestedTableItemLimit {
						out.OmittedCount++
						continue
					}
					out.Items = append(out.Items, context.describe(table, &cell, joined, anchor, preceding, ordinal, child))
				}
			}
		}
	}
	if len(out.Items) == 0 && out.OmittedCount == 0 {
		return nil, nil
	}
	return out, nil
}

func nativeApproximateNestedTableID(cellID string, ordinal int) string {
	safe := strings.Map(func(r rune) rune {
		if r >= 'A' && r <= 'Z' || r >= 'a' && r <= 'z' || r >= '0' && r <= '9' || r == '.' || r == '_' || r == ':' || r == '-' {
			return r
		}
		return '-'
	}, strings.TrimPrefix(cellID, "cell:"))
	if len(safe) > 160 {
		safe = safe[:160]
	}
	return fmt.Sprintf("approximate-nested-table:%s:%d", safe, ordinal)
}

// contentExtractor is a private extractor over the same package. Its
// diagnostics classify the nested table and are never merged into the document.
func (context *nativeApproximateNestedContext) contentExtractor() *nativeExtractor {
	relNS := relNSTransitional
	if context.ns == wordMLStrict {
		relNS = relNSStrict
	}
	return &nativeExtractor{
		pkg: context.resolver.pkg, wordNS: context.ns, relNS: relNS, mainPart: context.main, mainRoot: context.resolver.mainRoot,
		modeledParts: map[string]bool{context.main: true}, unsupportedSet: map[string]bool{}, storyByRel: map[string]string{}, storyByNative: map[string]string{}, commentByNative: map[string]string{},
		unsupported: []NativeUnsupportedCapabilityV1{}, previousIDs: map[string][]string{}, previousUsed: map[string]int{}, previousNative: map[string]string{}, previousPath: map[string]string{}, reservedIDs: map[string]bool{}, allocatedIDs: map[string]bool{}, seenParaIDs: map[string]string{},
	}
}

func (context *nativeApproximateNestedContext) describe(outer *NativeTableV1, cell *NativeTableCellV1, joined []string, anchor NativeSourceAnchorV1, preceding, ordinal int, node *nativeXMLNode) NativeApproximateNestedTableV1 {
	id := nativeApproximateNestedTableID(cell.ID, ordinal)
	item := NativeApproximateNestedTableV1{ID: id, TableID: outer.ID, CellID: cell.ID, DiagnosticIDs: joined, Anchor: anchor, PrecedingParagraphs: preceding, Status: "omitted", ResolvedParagraphs: []NativeResolvedParagraphV1{}, ResolvedRuns: []NativeResolvedRunV1{}}
	omit := func(reason string) NativeApproximateNestedTableV1 {
		item.Reason = reason
		item.Table, item.Geometry, item.StyleBorders, item.StyleCellShadingRGB, item.Notes = nil, nil, nil, nil, nil
		return item
	}
	extractor := context.contentExtractor()
	table, err := extractor.extractTable(context.main, node)
	if err != nil {
		return omit("unsupported-nested-content")
	}
	notes := map[string]bool{}
	for _, d := range extractor.unsupported {
		switch d.Code {
		case "NESTED_TABLE_OR_CELL_MARKUP":
			// Depth two and deeper keep the strict omission; only one level is laid out.
			return omit("nested-depth-limit")
		case "UNMODELED_TABLE_CONTENT", "UNMODELED_ROW_CONTENT":
			return omit("unsupported-table-structure")
		case "WRAPPED_ROW_CELLS":
			// The control's cells are read through and laid out; nothing is
			// dropped, so this disclosure omits neither the table nor a run.
		case "UNMODELED_TABLE_PROPERTY", "UNMODELED_ROW_PROPERTY", "UNMODELED_CELL_PROPERTY":
			local := "unknown"
			if d.Anchor != nil {
				local = d.Anchor.Path[strings.LastIndex(d.Anchor.Path, "/")+1:]
				if i := strings.IndexByte(local, '['); i >= 0 {
					local = local[:i]
				}
			}
			notes["table property "+local+" not applied"] = true
		default:
			// Run/paragraph content refusals (fields, drawings, references,
			// wrapped markup) drop that content from the preview and are counted.
			item.OmittedRuns++
		}
	}
	if extractor.unsupportedOverflow {
		return omit("unsupported-nested-content")
	}
	if len(table.GridWidthsTwips) == 0 || len(table.Rows) == 0 {
		return omit("missing-grid")
	}
	if len(table.Rows) > nativeApproximateNestedTableRowLimit {
		return omit("row-limit")
	}
	cells := 0
	for _, row := range table.Rows {
		if len(row.Cells) != len(table.GridWidthsTwips) {
			return omit("grid-mismatch")
		}
		for _, c := range row.Cells {
			cells++
			if cells > nativeApproximateNestedTableCellLimit {
				return omit("cell-limit")
			}
			if c.GridSpan == nil || *c.GridSpan != 1 || c.VerticalMerge != "none" {
				return omit("merged-cells")
			}
		}
	}
	table.ID = id
	table.EditPolicy = nativeReadOnlyPolicy("APPROXIMATE_NESTED_TABLE_PREVIEW", "Nested table is read-only approximate preview evidence")
	localResolver := *context.resolver
	localResolver.diagnostics = append([]NativeResolutionDiagnosticV1{}, context.resolver.diagnostics...)
	localResolver.diagnosticSet = map[string]bool{}
	for key, value := range context.resolver.diagnosticSet {
		localResolver.diagnosticSet[key] = value
	}
	var tableStyles []*nativeStyleDefinition
	if table.TableStyleID != nil {
		if localResolver.styles["table\x00"+*table.TableStyleID] == nil {
			notes["table style "+*table.TableStyleID+" missing"] = true
		} else {
			tableStyles = localResolver.styleChain("table", *table.TableStyleID, table.ID)
		}
	}
	geometry, geometryNotes := context.approximateGeometry(tableStyles, node)
	for _, note := range geometryNotes {
		notes[note] = true
	}
	item.Geometry = geometry
	item.StyleBorders, item.StyleCellShadingRGB = context.approximateStyleBorders(&localResolver, tableStyles)
	firstRow, firstRowShading := context.firstRowLayers(&localResolver, tableStyles, node, table.ID)
	item.FirstRowCellShadingRGB = firstRowShading
	for _, style := range tableStyles {
		for _, conditional := range directNativeChildren(style.node, context.ns, "tblStylePr") {
			if kind, _ := nativeAttr(conditional, context.ns, "type"); kind == "firstRow" && len(firstRow) > 0 {
				notes["firstRow conditional table-style region applied to the first row (paragraph, run and cell fill properties); its borders and other regions are not applied"] = true
				continue
			}
			notes["conditional table-style regions other than firstRow not applied"] = true
		}
	}
	if outerNode := context.nodes[outer.Anchor.Path]; outerNode != nil {
		var outerStyles []*nativeStyleDefinition
		if outer.TableStyleID != nil && localResolver.styles["table\x00"+*outer.TableStyleID] != nil {
			outerStyles = localResolver.styleChain("table", *outer.TableStyleID, outer.ID)
		}
		item.OuterGeometry, _ = context.approximateGeometry(outerStyles, outerNode)
	}
	result := &NativeResolvedLayoutInputV1{}
	numberingState := newNativeNumberingState()
	for r := range table.Rows {
		row := &table.Rows[r]
		row.ID = fmt.Sprintf("%s:r%d", id, r)
		for k := range row.Cells {
			c := &row.Cells[k]
			c.ID = fmt.Sprintf("%s:r%dc%d", id, r, k)
			for p := range c.Paragraphs {
				paragraph := &c.Paragraphs[p]
				paragraph.ID = fmt.Sprintf("%s:p%d", c.ID, p)
				paragraph.EditPolicy = nativeReadOnlyPolicy("APPROXIMATE_NESTED_TABLE_PREVIEW", "Nested table paragraphs are read-only approximate preview evidence")
				runs := []NativeRunV1{}
				for runIndex := range paragraph.Runs {
					run := paragraph.Runs[runIndex]
					if run.Kind != "text" && run.Kind != "control" {
						item.OmittedRuns++
						continue
					}
					if run.Kind == "control" && run.Control != "tab" && run.Control != "line-break" {
						item.OmittedRuns++
						continue
					}
					if run.Text != nil {
						context.textUnits += len(utf16.Encode([]rune(*run.Text)))
						if context.textUnits > nativeApproximateNestedTableTextLimit {
							return omit("text-limit")
						}
					}
					run.ID = fmt.Sprintf("%s:r%d", paragraph.ID, runIndex)
					runs = append(runs, run)
				}
				paragraph.Runs = runs
				layers := tableStyles
				if r == 0 && len(firstRow) > 0 {
					layers = append(append([]*nativeStyleDefinition{}, tableStyles...), firstRow...)
				}
				localResolver.resolveParagraph(paragraph, result, numberingState, layers)
			}
		}
	}
	if localResolver.diagnosticOverflow {
		return omit("unsupported-nested-content")
	}
	item.Table = &table
	item.Status = "supported"
	item.ResolvedParagraphs = result.Paragraphs
	item.ResolvedRuns = result.Runs
	if item.ResolvedParagraphs == nil {
		item.ResolvedParagraphs = []NativeResolvedParagraphV1{}
	}
	if item.ResolvedRuns == nil {
		item.ResolvedRuns = []NativeResolvedRunV1{}
	}
	if len(notes) > 0 {
		item.Notes = make([]string, 0, len(notes))
		for note := range notes {
			item.Notes = append(item.Notes, note)
		}
		sort.Strings(item.Notes)
	}
	return item
}

// approximateGeometry cascades table geometry from the style chain (root
// first) and the table's own tblPr the way the strict geometry resolver does,
// but tolerates conditional style layers and unqualified siblings: whatever it
// cannot read keeps the ECMA-376 default and is noted. Approximate only.
func (context *nativeApproximateNestedContext) approximateGeometry(styles []*nativeStyleDefinition, node *nativeXMLNode) (*NativeResolvedTableGeometryV1, []string) {
	ns := context.ns
	g := &NativeResolvedTableGeometryV1{Layout: "autofit", Alignment: "left", WidthType: "auto", CellMargins: NativeTableCellMarginsV1{LeftTwips: 115, RightTwips: 115}}
	notes := []string{}
	layers := []*nativeXMLNode{}
	for _, style := range styles {
		if tblPr := firstDirectNativeChild(style.node, ns, "tblPr"); tblPr != nil {
			layers = append(layers, tblPr)
		}
	}
	if tblPr := firstDirectNativeChild(node, ns, "tblPr"); tblPr != nil {
		layers = append(layers, tblPr)
	}
	for _, layer := range layers {
		for _, child := range layer.Children {
			if child.Name.Space != ns {
				continue
			}
			switch child.Name.Local {
			case "tblLayout":
				if v, ok := nativeAttr(child, ns, "type"); ok && (v == "fixed" || v == "autofit") {
					g.Layout = v
				}
			case "jc":
				if v, ok := nativeAttr(child, ns, "val"); ok && v != "left" && v != "start" {
					notes = append(notes, "table alignment "+v+" approximated as left")
				}
			case "tblInd":
				v, ok := nativeNonnegativeInt64Attr(child, ns, "w")
				kind, has := nativeAttr(child, ns, "type")
				if ok && (!has || kind == "dxa") && v <= nativeMaxTwipsForMilliPoints {
					g.IndentTwips = v
				} else {
					notes = append(notes, "table indent unit not applied")
				}
			case "tblW":
				v, ok := nativeNonnegativeInt64Attr(child, ns, "w")
				kind, has := nativeAttr(child, ns, "type")
				switch {
				case ok && has && kind == "dxa" && v > 0 && v <= nativeMaxTwipsForMilliPoints:
					g.WidthType, g.WidthValue = "dxa", v
				case ok && has && kind == "pct" && v > 0 && v <= 5000:
					g.WidthType, g.WidthValue = "pct", v
				case has && kind == "auto" || ok && v == 0:
					g.WidthType, g.WidthValue = "auto", 0
				default:
					notes = append(notes, "table width unit not applied")
				}
			case "tblCellMar":
				for _, side := range child.Children {
					if side.Name.Space != ns {
						continue
					}
					v, ok := nativeNonnegativeInt64Attr(side, ns, "w")
					kind, has := nativeAttr(side, ns, "type")
					if !ok || has && kind != "dxa" || v > nativeMaxTwipsForMilliPoints {
						notes = append(notes, "cell margin unit not applied")
						continue
					}
					switch side.Name.Local {
					case "top":
						g.CellMargins.TopTwips = v
					case "left", "start":
						g.CellMargins.LeftTwips = v
					case "bottom":
						g.CellMargins.BottomTwips = v
					case "right", "end":
						g.CellMargins.RightTwips = v
					}
				}
			}
		}
	}
	return g, notes
}

// approximateStyleBorders merges the base tblPr borders and tcPr fill of every
// style layer root first, ignoring conditional regions. Direct table and cell
// properties override these in the preview, as in the resolver.
func (context *nativeApproximateNestedContext) approximateStyleBorders(resolver *nativeLayoutResolver, styles []*nativeStyleDefinition) (*NativeTableBordersV1, *string) {
	ns := context.ns
	var borders *NativeTableBordersV1
	var fill *string
	for _, style := range styles {
		if tblPr := firstDirectNativeChild(style.node, ns, "tblPr"); tblPr != nil {
			if node := firstDirectNativeChild(tblPr, ns, "tblBorders"); node != nil {
				if parsed, ok := nativeExtractTableBorders(node, ns, resolver.resolveThemeSrgb); ok && parsed != nil {
					borders = mergeNativeTableBorders(borders, parsed)
				}
			}
		}
		if tcPr := firstDirectNativeChild(style.node, ns, "tcPr"); tcPr != nil {
			if node := firstDirectNativeChild(tcPr, ns, "shd"); node != nil {
				if parsed, ok := nativeExtractCellShading(node, ns, resolver.resolveThemeSrgb); ok && parsed != nil {
					fill = parsed
				}
			}
		}
	}
	return borders, fill
}

// firstRowLayers returns the firstRow conditional layers of the style chain
// (root first) as synthetic style definitions when the table's look enables
// that region, plus the region's clear cell fill. Word applies the region to
// the first row when tblLook selects it; an absent tblLook selects it too.
func (context *nativeApproximateNestedContext) firstRowLayers(resolver *nativeLayoutResolver, styles []*nativeStyleDefinition, node *nativeXMLNode, tableID string) ([]*nativeStyleDefinition, *string) {
	ns := context.ns
	enabled := true
	if tblPr := firstDirectNativeChild(node, ns, "tblPr"); tblPr != nil {
		if look := firstDirectNativeChild(tblPr, ns, "tblLook"); look != nil {
			enabled = false
			if value, ok := nativeAttr(look, ns, "firstRow"); ok {
				enabled = value == "1" || value == "true" || value == "on"
			} else if value, ok := nativeAttr(look, ns, "val"); ok {
				if mask, err := strconv.ParseUint(value, 16, 16); err == nil {
					enabled = mask&0x0020 != 0
				}
			}
		}
	}
	if !enabled {
		return nil, nil
	}
	// The chain's firstRow layers merge into one region layer (later layers
	// override earlier ones) so the region's toggle properties apply once,
	// which is how the rendered header of the benchmark corpus appears.
	var merged *nativeStyleDefinition
	var fill *string
	for _, style := range styles {
		for _, conditional := range directNativeChildren(style.node, ns, "tblStylePr") {
			if kind, _ := nativeAttr(conditional, ns, "type"); kind != "firstRow" {
				continue
			}
			if merged == nil {
				merged = &nativeStyleDefinition{id: style.id + ":firstRow", kind: "table", partName: style.partName, node: conditional}
			}
			if pPr := firstDirectNativeChild(conditional, ns, "pPr"); pPr != nil {
				applyNativeParagraphProperties(&merged.p, resolver.parseParagraphProperties(style.partName, pPr, tableID))
			}
			if rPr := firstDirectNativeChild(conditional, ns, "rPr"); rPr != nil {
				applyNativeRunProperties(&merged.r, resolver.parseRunProperties(style.partName, rPr, tableID), false)
			}
			if tcPr := firstDirectNativeChild(conditional, ns, "tcPr"); tcPr != nil {
				if shd := firstDirectNativeChild(tcPr, ns, "shd"); shd != nil {
					if parsed, ok := nativeExtractCellShading(shd, ns, resolver.resolveThemeSrgb); ok && parsed != nil {
						fill = parsed
					}
				}
			}
		}
	}
	if merged == nil {
		return nil, fill
	}
	return []*nativeStyleDefinition{merged}, fill
}
