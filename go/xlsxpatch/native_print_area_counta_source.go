package xlsxpatch

import (
	"strconv"
	"strings"
)

const nativePrintCountaLimit = 100000

// This private context is created only from Inspect's successful, complete
// extraction and used synchronously. It is not a client-supplied workbook view.
type nativePrintCountaSourceContext struct {
	workbook  *NativeWorkbookV2
	certified map[*NativeWorkbookSheetV2]bool
}

func newNativePrintCountaSourceContext(workbook *NativeWorkbookV2, workbookXML []byte, packageSHA string) *nativePrintCountaSourceContext {
	if workbook == nil || len(packageSHA) != 71 || !strings.HasPrefix(packageSHA, "sha256:") || workbook.Source.PackageSHA256 != packageSHA {
		return nil
	}
	root, err := parsePreviewXML(workbookXML)
	if err != nil || root.name.Local != "workbook" || !isSpreadsheetMLNamespace(root.name.Space) || strings.TrimSpace(root.text) != "" {
		return nil
	}
	expected := spreadsheetMLTransitional
	if workbook.Source.Dialect == "strict" {
		expected = spreadsheetMLStrict
	} else if workbook.Source.Dialect != "transitional" {
		return nil
	}
	if root.name.Space != expected {
		return nil
	}
	seen := map[string]bool{}
	for _, child := range root.children {
		if child.name.Space != expected || (!nativePrintCountaWorkbookChild(child)) || seen[child.name.Local] {
			return nil
		}
		seen[child.name.Local] = true
	}
	if !seen["sheets"] {
		return nil
	}
	names, _, ok := collectNativePrintNames(workbookXML, workbook.Sheets)
	if !ok {
		return nil
	}
	needed := false
	for _, definitions := range names {
		for _, definition := range definitions {
			if strings.Contains(definition.text, "COUNTA(") {
				needed = true
			}
		}
	}
	if !needed {
		return nil
	} // Do not certify inventories for unrelated previews.
	ctx := &nativePrintCountaSourceContext{workbook: workbook, certified: map[*NativeWorkbookSheetV2]bool{}}
	ids, parts := map[string]bool{}, map[string]bool{}
	for i := range workbook.Sheets {
		sheet := &workbook.Sheets[i]
		if ids[sheet.ID] || parts[sheet.PartName] || sheet.Order != i {
			return nil
		}
		ids[sheet.ID], parts[sheet.PartName] = true, true
	}
	for i := range workbook.Sheets {
		sheet := &workbook.Sheets[i]
		ctx.certified[sheet] = nativePrintCountaCertifySheet(workbook, sheet)
	}
	return ctx
}

// Explicit source-neutral workbook shapes; a generic unsupported record alone
// cannot distinguish these from unknown or alternative-content elements.
func nativePrintCountaWorkbookChild(n *previewXML) bool {
	switch n.name.Local {
	case "sheets", "definedNames":
		return true
	case "bookViews":
		if strings.TrimSpace(n.text) != "" || len(n.children) != 1 {
			return false
		}
		for _, a := range n.attrs {
			if !isPreviewNamespaceDeclaration(a) {
				return false
			}
		}
		v := n.children[0]
		if v.name.Space != n.name.Space || v.name.Local != "workbookView" || len(v.children) != 0 || strings.TrimSpace(v.text) != "" {
			return false
		}
		for _, a := range v.attrs {
			if !isPreviewNamespaceDeclaration(a) {
				return false
			}
		}
		return true
	case "calcPr":
		if len(n.children) != 0 || strings.TrimSpace(n.text) != "" {
			return false
		}
		found := false
		for _, a := range n.attrs {
			if isPreviewNamespaceDeclaration(a) {
				continue
			}
			if found || a.Name.Space != "" || a.Name.Local != "calcId" || len(a.Value) == 0 || len(a.Value) > 10 {
				return false
			}
			for _, c := range a.Value {
				if c < '0' || c > '9' {
					return false
				}
			}
			if _, err := strconv.ParseUint(a.Value, 10, 32); err != nil {
				return false
			}
			found = true
		}
		return found
	default:
		return false
	}
}

func nativePrintCountaCertifySheet(workbook *NativeWorkbookV2, sheet *NativeWorkbookSheetV2) bool {
	if !sheet.Editable || sheet.RefusalCode != nil || sheet.Cells == nil || len(sheet.Cells) > nativePrintCountaLimit {
		return false
	}
	coords := map[string]bool{}
	for _, cell := range sheet.Cells {
		if cell.Row < 0 || cell.Row >= 1048576 || cell.Column < 0 || cell.Column >= 16384 || cell.Ref != cellReference(cell.Row, cell.Column) || coords[cell.Ref] {
			return false
		}
		coords[cell.Ref] = true
	}
	core := map[string]bool{workbook.Source.WorkbookPart: true}
	for _, candidate := range workbook.Sheets {
		core[candidate.PartName] = true
	}
	for _, item := range workbook.Unsupported {
		if item.PartName != nil && (item.Capability == "styles" || item.Capability == "rich-text") {
			core[*item.PartName] = true
		}
	}
	for _, item := range workbook.Unsupported {
		if item.PartName == nil {
			return false
		}
		part := *item.PartName
		if item.ScopeID == "sheet:"+sheet.ID || part == sheet.PartName {
			return false
		}
		otherSheet := false
		for _, other := range workbook.Sheets {
			if other.ID != sheet.ID && item.ScopeID == "sheet:"+other.ID && part == other.PartName {
				otherSheet = true
			}
		}
		if otherSheet {
			continue
		}
		if item.CellRef != nil || item.RangeRef != nil {
			return false
		}
		for _, other := range workbook.Sheets {
			if part == other.PartName {
				return false
			}
		}
		switch item.Code {
		case "UNMODELED_WORKBOOK_FEATURE", "WORKBOOK_VIEW_METADATA":
			if item.Capability != "workbook-features" || item.ScopeID != "workbook" || part != workbook.Source.WorkbookPart {
				return false
			}
			// Constructor's closed root gate disambiguates this deduplicated code.
		case "STYLE_FONT_OPAQUE_CONTENT", "STYLE_TABLE_OPAQUE_CONTENT":
			if item.Capability != "styles" || item.ScopeID != "workbook" || part == workbook.Source.WorkbookPart {
				return false
			}
		case "STYLE_FONT_COLOR":
			if item.Capability != "styles" || !strings.HasPrefix(item.ScopeID, "style:") || part == workbook.Source.WorkbookPart {
				return false
			}
		case "RICH_SHARED_STRING":
			if item.Capability != "rich-text" || item.ScopeID != "workbook" || part == workbook.Source.WorkbookPart {
				return false
			}
		case "CHART_CONTENT":
			if item.Capability != "charts" || item.ScopeID != "workbook" || core[part] {
				return false
			}
		case "EXTERNAL_RELATIONSHIP":
			if item.Capability != "external-links" || item.ScopeID != "workbook" || core[part] {
				return false
			}
		case "OPAQUE_PACKAGE_PART":
			if item.Capability != "opaque-parts" || item.ScopeID != "workbook" || core[part] {
				return false
			}
		default:
			return false
		}
	}
	return true
}

// Occupancy is based on validated saved literals, never display formatting or
// formula presence. Empty inline text is ambiguous in the current projection.
func nativePrintCountaCell(cell NativeWorkbookCellV2) (int, bool) {
	if !cell.Editable || cell.Formula != nil {
		return 0, false
	}
	storage := "n"
	if cell.OOXMLType != nil {
		storage = *cell.OOXMLType
	}
	v := cell.Value
	if v == nil {
		return 0, storage == "n"
	}
	if v.Rich || len(v.Runs) != 0 {
		return 0, false
	}
	switch storage {
	case "n":
		return 1, v.Kind == "number" && v.Storage == "number" && v.Lexical != nil && *v.Lexical != "" && v.Text == nil
	case "b":
		return 1, v.Kind == "boolean" && v.Storage == "boolean" && v.Lexical != nil && v.Text == nil
	case "e":
		return 1, v.Kind == "error" && v.Storage == "error" && v.Lexical != nil && *v.Lexical != "" && v.Text == nil
	case "d":
		return 1, v.Kind == "date" && v.Storage == "date" && v.Lexical != nil && v.Text == nil
	case "s":
		return 1, v.Kind == "string" && v.Storage == "shared" && v.Text != nil && v.Lexical != nil
	case "inlineStr":
		return 1, v.Kind == "string" && v.Storage == "inline" && v.Text != nil && *v.Text != "" && v.Lexical == nil
	default:
		return 0, false
	}
}

func (ctx *nativePrintCountaSourceContext) count(sheet *NativeWorkbookSheetV2, rect NativePrintAreaRectV1) (int, bool) {
	if ctx == nil || sheet == nil || !ctx.certified[sheet] {
		return 0, false
	}
	for _, merged := range sheet.MergedRanges {
		if rect.Row <= merged.EndRow && merged.Row <= rect.EndRow && rect.Column <= merged.EndColumn && merged.Column <= rect.EndColumn {
			return 0, false
		}
	}
	count := 0
	for _, cell := range sheet.Cells {
		if cell.Row < rect.Row || cell.Row > rect.EndRow || cell.Column < rect.Column || cell.Column > rect.EndColumn {
			continue
		}
		n, ok := nativePrintCountaCell(cell)
		if !ok {
			return 0, false
		}
		count += n
	}
	return count, true
}
