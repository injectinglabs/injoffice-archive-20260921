package docxpatch

import (
	"encoding/xml"
	"strings"
)

const NativeAutomaticTableBorderPolicyV1 = "auto-border-on-qualified-white-v1"

type NativeAutomaticTableBorderDiagnosticV1 struct {
	Code     string `json:"code"`
	ScopeID  string `json:"scope_id"`
	PartName string `json:"part_name"`
	Path     string `json:"path"`
}

// Evidence for an explicit read-only consumer contrast policy, not source RGB.
// Strict borders and diagnostics remain untouched. The preview consumer must
// provide an opaque white page and retain an approximate-fidelity warning.
type NativeAutomaticTableBorderPreviewV1 struct {
	Policy            string                                   `json:"policy"`
	ReadOnly          bool                                     `json:"read_only"`
	PageBackground    string                                   `json:"page_background"`
	BackgroundRGB     string                                   `json:"background_rgb"`
	PackageSHA256     string                                   `json:"package_sha256"`
	SourcePart        string                                   `json:"source_part"`
	SourcePath        string                                   `json:"source_path"`
	SourceSHA256      string                                   `json:"source_sha256"`
	Borders           NativeTableBordersV1                     `json:"borders"`
	AutomaticEdges    []string                                 `json:"automatic_edges"`
	CellIDs           []string                                 `json:"cell_ids"`
	SourceDiagnostics []NativeAutomaticTableBorderDiagnosticV1 `json:"source_diagnostics"`
}

func (resolver *nativeLayoutResolver) automaticTableBorderPreview(table *NativeTableV1, chain []*nativeStyleDefinition) *NativeAutomaticTableBorderPreviewV1 {
	if len(chain) == 0 || chain[0].basedOn != "" || resolver.resolveTableGeometry(table) == nil || !resolver.automaticBorderWhitePage() {
		return nil
	}
	result := &NativeAutomaticTableBorderPreviewV1{Policy: NativeAutomaticTableBorderPolicyV1, ReadOnly: true, PageBackground: "absent-on-white-preview", BackgroundRGB: "FFFFFF", PackageSHA256: resolver.doc.Source.PackageSHA256, AutomaticEdges: []string{}, CellIDs: []string{}, SourceDiagnostics: []NativeAutomaticTableBorderDiagnosticV1{}}
	var owner *nativeXMLNode
	ownerPart := ""
	addOwner := func(pr *nativeXMLNode, part string) bool {
		if pr == nil {
			return true
		}
		if !nativeExactContainer(pr) {
			return false
		}
		borders := directNativeChildren(pr, resolver.wordNS, "tblBorders")
		if len(borders) > 1 || len(borders) == 1 && owner != nil {
			return false
		}
		if len(borders) == 1 {
			owner = borders[0]
			ownerPart = part
		}
		return true
	}
	for _, style := range chain {
		if len(directNativeChildren(style.node, resolver.wordNS, "tblStylePr")) != 0 || len(directNativeChildren(style.node, resolver.wordNS, "tcPr")) > 1 || len(directNativeChildren(style.node, resolver.wordNS, "trPr")) > 0 {
			return nil
		}
		for _, c := range style.node.Children {
			if c.Name.Space != resolver.wordNS || !nativeTableStyleAuthoringLocal(c.Name.Local) && c.Name.Local != "pPr" && c.Name.Local != "rPr" && c.Name.Local != "tblPr" && c.Name.Local != "tcPr" {
				return nil
			}
		}
		if !nativeWhiteTableCellProperties(firstDirectNativeChild(style.node, resolver.wordNS, "tcPr"), resolver.wordNS, true) || !addOwner(firstDirectNativeChild(style.node, resolver.wordNS, "tblPr"), style.partName) {
			return nil
		}
	}
	node := resolver.nodeForAnchor(table.Anchor)
	if node == nil || !addOwner(firstDirectNativeChild(node, resolver.wordNS, "tblPr"), table.Anchor.PartName) || owner == nil {
		return nil
	}
	if !nativeExactContainer(owner) || len(owner.Children) == 0 || len(owner.Children) > 6 {
		return nil
	}
	seen := map[string]bool{}
	for _, edge := range owner.Children {
		if edge.Name.Space != resolver.wordNS || seen[edge.Name.Local] {
			return nil
		}
		seen[edge.Name.Local] = true
		var border *NativeTableBorderV1
		color, _ := nativeAttr(edge, resolver.wordNS, "color")
		kind, _ := nativeAttr(edge, resolver.wordNS, "val")
		automatic := color == "auto" && kind == "single"
		if automatic {
			if !nativeExactLeaf(edge, xml.Name{Space: resolver.wordNS, Local: "val"}, xml.Name{Space: resolver.wordNS, Local: "sz"}, xml.Name{Space: resolver.wordNS, Local: "space"}, xml.Name{Space: resolver.wordNS, Local: "color"}) {
				return nil
			}
			kind, _ := nativeAttr(edge, resolver.wordNS, "val")
			size, ok := nativeNonnegativeInt64Attr(edge, resolver.wordNS, "sz")
			if kind != "single" || !ok || size <= 0 || size > 768 || !nativeZeroOrAbsentTwipAttr(edge, resolver.wordNS, "space") {
				return nil
			}
			border = &NativeTableBorderV1{Style: "single", SizeEighthPoints: size, ColorRGB: nativeString("000000")}
		} else {
			var ok bool
			border, ok = nativeExtractTableBorder(edge, resolver.wordNS, resolver.resolveThemeSrgb)
			if !ok || border.SizeEighthPoints > 768 {
				return nil
			}
		}
		key := ""
		switch edge.Name.Local {
		case "top":
			result.Borders.Top = border
			key = "top"
		case "right":
			result.Borders.Right = border
			key = "right"
		case "bottom":
			result.Borders.Bottom = border
			key = "bottom"
		case "left":
			result.Borders.Left = border
			key = "left"
		case "insideH":
			result.Borders.InsideHorizontal = border
			key = "inside_horizontal"
		case "insideV":
			result.Borders.InsideVertical = border
			key = "inside_vertical"
		default:
			return nil
		}
		if automatic {
			result.AutomaticEdges = append(result.AutomaticEdges, key)
		}
	}
	if len(result.AutomaticEdges) == 0 {
		return nil
	}
	for _, row := range table.Rows {
		rowNode := resolver.nodeForAnchor(row.Anchor)
		if rowNode == nil || len(directNativeChildren(rowNode, resolver.wordNS, "tblPrEx")) > 0 {
			return nil
		}
		for _, cell := range row.Cells {
			if cell.GridSpan == nil || *cell.GridSpan != 1 || cell.VerticalMerge != "none" || cell.Borders != nil || len(result.CellIDs) >= 10000 {
				return nil
			}
			cellNode := resolver.nodeForAnchor(cell.Anchor)
			if cellNode == nil || len(directNativeChildren(cellNode, resolver.wordNS, "tcPr")) > 1 || !nativeWhiteTableCellProperties(firstDirectNativeChild(cellNode, resolver.wordNS, "tcPr"), resolver.wordNS, false) {
				return nil
			}
			result.CellIDs = append(result.CellIDs, cell.ID)
		}
	}
	if len(result.CellIDs) == 0 {
		return nil
	}
	result.SourcePart = ownerPart
	result.SourcePath = owner.Path
	result.SourceSHA256 = nativeSHA(resolver.pkg.files[ownerPart])
	for _, diagnostic := range resolver.diagnostics {
		if diagnostic.Code == "TABLE_STYLE_EFFECTS_PRESERVED" && diagnostic.ScopeID == table.ID && diagnostic.PartName != nil && *diagnostic.PartName == ownerPart && diagnostic.Path != nil && *diagnostic.Path == owner.parent.Path {
			result.SourceDiagnostics = append(result.SourceDiagnostics, NativeAutomaticTableBorderDiagnosticV1{diagnostic.Code, diagnostic.ScopeID, *diagnostic.PartName, *diagnostic.Path})
		}
	}
	for _, diagnostic := range resolver.doc.Unsupported {
		if diagnostic.Code == "UNMODELED_TABLE_PROPERTY" && diagnostic.ScopeID == table.ID && diagnostic.Anchor.PartName == ownerPart && diagnostic.Anchor.Path == owner.Path {
			result.SourceDiagnostics = append(result.SourceDiagnostics, NativeAutomaticTableBorderDiagnosticV1{diagnostic.Code, diagnostic.ScopeID, diagnostic.Anchor.PartName, diagnostic.Anchor.Path})
		}
	}
	if len(result.SourceDiagnostics) == 0 {
		return nil
	}
	return result
}

// The resolver and its indexed source trees are private to one immutable
// package resolution. Qualify the page once, not once per table.
func (resolver *nativeLayoutResolver) automaticBorderWhitePage() bool {
	if resolver.autoBorderWhiteChecked {
		return resolver.autoBorderWhite
	}
	resolver.autoBorderWhiteChecked = true
	root := resolver.mainRoot
	if len(resolver.doc.Headers)+len(resolver.doc.Footers)+len(resolver.doc.Notes)+len(resolver.doc.CommentStories) > 0 || root == nil || root.Name != (xml.Name{Space: resolver.wordNS, Local: "document"}) || !nativeExactContainer(root) || len(root.Children) != 1 || root.Children[0].Name != (xml.Name{Space: resolver.wordNS, Local: "body"}) || !nativeExactContainer(root.Children[0]) {
		return false
	}
	var drawing func(*nativeXMLNode) bool
	drawing = func(n *nativeXMLNode) bool {
		if n.Name.Local == "drawing" || n.Name.Local == "pict" || n.Name.Local == "object" || n.Name.Local == "background" {
			return true
		}
		for _, c := range n.Children {
			if drawing(c) {
				return true
			}
		}
		return false
	}
	if drawing(root) {
		return false
	}
	for _, child := range root.Children[0].Children {
		if child.Name.Space != resolver.wordNS || child.Name.Local != "p" && child.Name.Local != "tbl" && child.Name.Local != "sectPr" {
			return false
		}
	}
	resolver.autoBorderWhite = true
	return true
}

func nativeValidAutomaticBorderPreview(f *NativeAutomaticTableBorderPreviewV1, tableID string) bool {
	if f.Policy != NativeAutomaticTableBorderPolicyV1 || !f.ReadOnly || f.PageBackground != "absent-on-white-preview" || f.BackgroundRGB != "FFFFFF" || !nativeSHA256.MatchString(f.PackageSHA256) || !nativeSHA256.MatchString(f.SourceSHA256) || !nativeBoundedResolvedString(f.SourcePart, 1024) || !nativeBoundedResolvedString(f.SourcePath, 4096) || len(f.CellIDs) == 0 || len(f.CellIDs) > 10000 || len(f.AutomaticEdges) == 0 || len(f.AutomaticEdges) > 6 || len(f.SourceDiagnostics) == 0 || len(f.SourceDiagnostics) > 2 {
		return false
	}
	edges := map[string]*NativeTableBorderV1{"top": f.Borders.Top, "right": f.Borders.Right, "bottom": f.Borders.Bottom, "left": f.Borders.Left, "inside_horizontal": f.Borders.InsideHorizontal, "inside_vertical": f.Borders.InsideVertical}
	for _, border := range edges {
		if border == nil {
			continue
		}
		if border.Style == "none" {
			if border.SizeEighthPoints != 0 || border.ColorRGB != nil {
				return false
			}
		} else if border.Style != "single" || border.SizeEighthPoints <= 0 || border.SizeEighthPoints > 768 || border.ColorRGB == nil {
			return false
		} else if _, ok := nativeExactRGB(*border.ColorRGB); !ok {
			return false
		}
	}
	seen := map[string]bool{}
	for _, edge := range f.AutomaticEdges {
		border := edges[edge]
		if seen[edge] || border == nil || border.Style != "single" || border.ColorRGB == nil || *border.ColorRGB != "000000" {
			return false
		}
		seen[edge] = true
	}
	seen = map[string]bool{}
	for _, id := range f.CellIDs {
		if !nativeIDPattern.MatchString(id) || seen[id] {
			return false
		}
		seen[id] = true
	}
	for _, d := range f.SourceDiagnostics {
		if d.Code != "TABLE_STYLE_EFFECTS_PRESERVED" && d.Code != "UNMODELED_TABLE_PROPERTY" || d.ScopeID != tableID || d.PartName != f.SourcePart || !nativeBoundedResolvedString(d.Path, 4096) {
			return false
		}
		key := d.Code + "\x00" + d.Path
		if seen[key] {
			return false
		}
		seen[key] = true
		if d.Code == "UNMODELED_TABLE_PROPERTY" && d.Path != f.SourcePath {
			return false
		}
		if d.Code == "TABLE_STYLE_EFFECTS_PRESERVED" && f.SourcePath != d.Path+"/w:tblBorders[1]" {
			return false
		}
	}
	if !strings.HasSuffix(f.SourcePath, "/w:tblBorders[1]") {
		return false
	}
	return true
}

func nativeWhiteTableCellProperties(node *nativeXMLNode, ns string, style bool) bool {
	if node == nil {
		return true
	}
	if !nativeExactContainer(node) {
		return false
	}
	seen := map[string]bool{}
	for _, child := range node.Children {
		if child.Name.Space != ns || seen[child.Name.Local] {
			return false
		}
		seen[child.Name.Local] = true
		if child.Name.Local != "shd" {
			if style || child.Name.Local != "tcW" && child.Name.Local != "gridSpan" && child.Name.Local != "vMerge" && child.Name.Local != "vAlign" {
				return false
			}
			continue
		}
		if !nativeExactLeaf(child, xml.Name{Space: ns, Local: "val"}, xml.Name{Space: ns, Local: "fill"}, xml.Name{Space: ns, Local: "color"}) {
			return false
		}
		kind, _ := nativeAttr(child, ns, "val")
		fill, _ := nativeAttr(child, ns, "fill")
		if kind != "clear" || fill != "FFFFFF" {
			return false
		}
		if color, present := nativeAttr(child, ns, "color"); present {
			if color != "auto" {
				if _, ok := nativeExactRGB(color); !ok {
					return false
				}
			}
		}
	}
	return true
}
