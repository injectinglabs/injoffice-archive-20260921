package xlsxpatch

import (
	"bytes"
	"encoding/json"
	"encoding/xml"
	"regexp"
	"strconv"
	"strings"
	"unicode/utf16"
)

// This supplement attests direct properties, not complete rich-text fidelity.
// Omitted properties use an explicitly disclosed host/cell fallback.
type NativeRichTextPreviewV1 struct {
	Cells    []NativeRichTextCellV1 `json:"cells"`
	Warnings []string               `json:"warnings"`
}
type NativeRichTextCellV1 struct {
	SheetID     string                `json:"sheet_id"`
	SheetPart   string                `json:"sheet_part"`
	Row         int                   `json:"row"`
	Column      int                   `json:"column"`
	Ref         string                `json:"ref"`
	StyleID     uint32                `json:"style_id"`
	Storage     string                `json:"storage"`
	SourcePart  string                `json:"source_part"`
	SharedIndex string                `json:"shared_index"`
	Text        string                `json:"text"`
	Status      string                `json:"status"`
	Warnings    []string              `json:"warnings"`
	Runs        []NativeRichTextRunV1 `json:"runs,omitempty"`
}
type NativeRichTextRunV1 struct {
	Text           string   `json:"text"`
	Properties     string   `json:"properties"` // direct or cell-inherited (absent rPr)
	FontName       *string  `json:"font_name,omitempty"`
	FontSizePoints *float64 `json:"font_size_points,omitempty"`
	FontColor      *string  `json:"font_color,omitempty"`
	Bold           *bool    `json:"bold,omitempty"`
	Italic         *bool    `json:"italic,omitempty"`
	Omitted        []string `json:"omitted"`
}

var nativeRichRGB = regexp.MustCompile(`(?i)^FF[0-9A-F]{6}$`)
var nativeRichSize = regexp.MustCompile(`^[0-9]{1,3}(\.[0-9]{1,2})?$`)
var nativeRichFont = regexp.MustCompile(`^[A-Za-z0-9][A-Za-z0-9 ._-]{0,127}$`)

func nativeRichUnits(s string) int { return len(utf16.Encode([]rune(s))) }
func nativeRichTextSafe(s string) bool {
	for _, r := range s {
		if r < 32 || r == 127 || r == 0x2028 || r == 0x2029 {
			return false
		}
	}
	return true
}
func nativeRichNode(n *previewXML, ns, name string, attrs ...xml.Name) bool {
	if n == nil || n.name.Space != ns || n.name.Local != name {
		return false
	}
	for _, a := range n.attrs {
		if isNamespaceDeclaration(a) {
			continue
		}
		found := false
		for _, k := range attrs {
			if a.Name == k {
				found = true
			}
		}
		if !found {
			return false
		}
	}
	return true
}
func nativeRichContainer(n *previewXML, ns, name string) bool {
	return nativeRichNode(n, ns, name) && strings.TrimSpace(n.text) == ""
}

// A rejected declaration omits all run styling for this cell. Explicit baseline,
// disabled effects and the numbered family hint are closed and disclosed.
func nativeRichRuns(item *previewXML, ns string) ([]NativeRichTextRunV1, string) {
	if item == nil || !nativeRichContainer(item, ns, item.name.Local) || (item.name.Local != "si" && item.name.Local != "is") || len(item.children) < 1 || len(item.children) > 64 {
		return nil, "Unsupported rich-string container or run count."
	}
	runs := []NativeRichTextRunV1{}
	total := 0
	for _, r := range item.children {
		if !nativeRichContainer(r, ns, "r") || len(r.children) < 1 || len(r.children) > 2 {
			return nil, "Unknown rich-string content; run styling omitted."
		}
		run := NativeRichTextRunV1{Properties: "cell-inherited", Omitted: []string{}}
		ti := 0
		if r.children[0].name.Local == "rPr" {
			p := r.children[0]
			ti = 1
			run.Properties = "direct"
			if !nativeRichContainer(p, ns, "rPr") {
				return nil, "Unknown run-property attributes."
			}
			seen := map[string]bool{}
			for _, v := range p.children {
				name := v.name.Local
				attribute := "val"
				if name == "color" {
					attribute = "rgb"
				}
				if !nativeRichNode(v, ns, name, xml.Name{Local: attribute}) || seen[name] || len(v.children) != 0 || strings.TrimSpace(v.text) != "" {
					return nil, "Unknown, duplicate or foreign run property."
				}
				seen[name] = true
				val := v.attr("val")
				switch name {
				case "b", "i":
					b := true
					for _, a := range v.attrs {
						if a.Name.Space == "" && a.Name.Local == "val" && a.Value == "" {
							return nil, "Invalid empty run boolean."
						}
					}
					if val != "" {
						if val != "0" && val != "1" && val != "true" && val != "false" {
							return nil, "Invalid run boolean."
						}
						b = val == "1" || val == "true"
					}
					if name == "b" {
						run.Bold = &b
					} else {
						run.Italic = &b
					}
				case "rFont":
					if !nativeRichFont.MatchString(val) {
						return nil, "Unsupported direct font name."
					}
					run.FontName = &val
				case "sz":
					size, err := strconv.ParseFloat(val, 64)
					if err != nil || !nativeRichSize.MatchString(val) || size < 1 || size > 409 {
						return nil, "Unsupported direct font size."
					}
					run.FontSizePoints = &size
				case "color":
					rgb := v.attr("rgb")
					if !nativeRichRGB.MatchString(rgb) {
						return nil, "Unsupported non-opaque or indirect run color."
					}
					color := "#" + strings.ToUpper(rgb[2:])
					run.FontColor = &color
				case "family":
					if len(val) != 1 || val < "0" || val > "5" {
						return nil, "Unknown font-family classification."
					}
					run.Omitted = append(run.Omitted, "font-family-hint")
				case "vertAlign":
					if val != "baseline" {
						return nil, "Superscript/subscript run styling omitted."
					}
					run.Omitted = append(run.Omitted, "baseline")
				case "u":
					if val != "none" {
						return nil, "Underline run styling omitted."
					}
					run.Omitted = append(run.Omitted, "underline-none")
				case "strike":
					if val != "0" && val != "false" {
						return nil, "Unsupported run effect."
					}
					run.Omitted = append(run.Omitted, name+"-false")
				default:
					return nil, "Unknown or theme-dependent run property; styling omitted."
				}
			}
		}
		if ti != len(r.children)-1 {
			return nil, "Run requires exactly one text element."
		}
		t := r.children[ti]
		if !nativeRichNode(t, ns, "t", xml.Name{Space: "http://www.w3.org/XML/1998/namespace", Local: "space"}) || len(t.children) != 0 {
			return nil, "Unknown run text markup."
		}
		for _, a := range t.attrs {
			if !isNamespaceDeclaration(a) && a.Value != "preserve" && a.Value != "default" {
				return nil, "Invalid text whitespace declaration."
			}
		}
		text, err := decodeSpreadsheetString(t.text)
		if err != nil || text == "" || !nativeRichTextSafe(text) {
			return nil, "Unsupported empty, multiline or control-containing run text."
		}
		run.Text = text
		total += nativeRichUnits(text)
		if total > 2048 {
			return nil, "Rich text exceeds the per-cell text bound."
		}
		runs = append(runs, run)
	}
	return runs, ""
}

func previewNativeRichText(pkg *nativeWorkbookPackage, workbook *NativeWorkbookV2) NativeRichTextPreviewV1 {
	result := NativeRichTextPreviewV1{Cells: []NativeRichTextCellV1{}, Warnings: []string{}}
	read := func(name string) ([]byte, bool) { v, ok := pkg.files[name]; return v, ok }
	location, err := locateWorkbookPartBytes(pkg.index, read)
	if err != nil {
		return result
	}
	extractor := nativeWorkbookExtractor{pkg: pkg, workbook: location}
	expected, opposing := relTypeSharedStringsTransitional, relTypeSharedStringsStrict
	if location.strict {
		expected, opposing = opposing, expected
	}
	sharedPart, sharedErr := extractor.relatedCorePart(expected, opposing, nativeSharedStringsType, "sharedStrings", false)
	var shared *previewXML
	if sharedErr == nil && sharedPart != "" {
		shared, _ = parsePreviewXML(pkg.files[sharedPart])
		if shared != nil {
			for _, item := range shared.children {
				if item.name.Space != shared.name.Space || item.name.Local != "si" {
					shared = nil
					break
				}
			}
		}
	}
	totalRuns, totalText := 0, 0
	for _, sheet := range workbook.Sheets {
		ns := spreadsheetMLTransitional
		if location.strict {
			ns = spreadsheetMLStrict
		}
		root, e := parsePreviewXML(pkg.files[sheet.PartName])
		rawCells := map[string]*previewXML{}
		sheetOK := e == nil
		if sheetOK {
			sheetOK = root.name.Local == "worksheet" && root.name.Space == ns
			directData := map[*previewXML]bool{}
			for _, child := range root.children {
				if child.name.Local == "sheetData" && child.name.Space == ns {
					directData[child] = true
				}
			}
			if len(directData) != 1 {
				sheetOK = false
			}
			var walk func(*previewXML, *previewXML, *previewXML)
			walk = func(n, p, g *previewXML) {
				if n.name.Local == "tableParts" && (p != root || !nativeRichNode(n, ns, "tableParts", xml.Name{Local: "count"}) || n.attr("count") != "0" || strings.TrimSpace(n.text) != "" || len(n.children) != 0) || n.name.Local == "extLst" || n.name.Local == "AlternateContent" {
					sheetOK = false
				}
				if n.name.Local == "c" {
					ref := n.attr("r")
					if p == nil || g == nil || p.name.Local != "row" || p.name.Space != ns || g.name.Local != "sheetData" || g.name.Space != ns || !directData[g] || n.name.Space != ns || rawCells[ref] != nil {
						sheetOK = false
					}
					rawCells[ref] = n
				}
				for _, c := range n.children {
					walk(c, n, p)
				}
			}
			walk(root, nil, nil)
		}
		ids := map[string]bool{}
		for _, c := range sheet.Cells {
			if ids[c.Ref] || c.Ref != cellReference(c.Row, c.Column) || c.Formula != nil && c.Formula.Type != "normal" {
				sheetOK = false
			}
			ids[c.Ref] = true
		}
		for _, c := range sheet.Cells {
			if c.Value == nil || c.Value.Kind != "string" || !c.Value.Rich || c.Value.Text == nil {
				continue
			}
			text := *c.Value.Text
			units := nativeRichUnits(text)
			if len(result.Cells) >= 256 || units > 2048 || totalText+units > 32768 {
				if len(result.Warnings) == 0 {
					result.Warnings = append(result.Warnings, "Some rich cells exceed preview bounds; plain source text is retained without run styling.")
				}
				continue
			}
			totalText += units
			entry := NativeRichTextCellV1{SheetID: sheet.ID, SheetPart: sheet.PartName, Row: c.Row, Column: c.Column, Ref: c.Ref, StyleID: c.StyleID, Storage: c.Value.Storage, SourcePart: sheet.PartName, Text: text, Status: "omitted", Warnings: []string{"Run styling omitted: unqualified source, layout, or cell ownership. Plain source text retained."}}
			if c.Value.Storage == "shared" {
				if sharedPart != "" {
					entry.SourcePart = sharedPart
				}
				if c.Value.Lexical != nil {
					entry.SharedIndex = *c.Value.Lexical
				}
			}
			source := rawCells[c.Ref]
			ok := sheetOK && c.Formula == nil && source != nil && nativeRichNode(source, ns, "c", xml.Name{Local: "r"}, xml.Name{Local: "s"}, xml.Name{Local: "t"}) && strings.TrimSpace(source.text) == "" && len(source.children) == 1 && nativeRichTextSafe(text)
			for _, u := range workbook.Unsupported {
				if u.PartName != nil && *u.PartName == sheet.PartName && u.CellRef != nil && *u.CellRef == c.Ref && u.Code != "RICH_CELL_STRING" {
					ok = false
				}
			}
			for _, m := range sheet.MergedRanges {
				if c.Row >= m.Row && c.Row <= m.EndRow && c.Column >= m.Column && c.Column <= m.EndColumn {
					ok = false
				}
			}
			if int(c.StyleID) >= len(workbook.Styles) {
				ok = false
			} else {
				s := workbook.Styles[c.StyleID].Effective
				if workbook.Styles[c.StyleID].ID != c.StyleID || s.Projection != "full" || len(s.Unsupported) != 0 || s.WrapText != nil && *s.WrapText || s.ShrinkToFit != nil && *s.ShrinkToFit || s.TextRotation != nil && *s.TextRotation != 0 {
					ok = false
				}
			}
			var item *previewXML
			if ok {
				switch c.Value.Storage {
				case "inline":
					if source.attr("t") != "inlineStr" {
						ok = false
					} else {
						item = source.children[0]
					}
				case "shared":
					index, err := strconv.Atoi(entry.SharedIndex)
					v := source.children[0]
					if err != nil || strconv.Itoa(index) != entry.SharedIndex || index < 0 || shared == nil || !nativeRichNode(shared, ns, "sst", xml.Name{Local: "count"}, xml.Name{Local: "uniqueCount"}) || strings.TrimSpace(shared.text) != "" || index >= len(shared.children) || source.attr("t") != "s" || !nativeRichNode(v, ns, "v") || len(v.children) != 0 || v.text != entry.SharedIndex {
						ok = false
					} else {
						item = shared.children[index]
					}
				default:
					ok = false
				}
			}
			if ok {
				runs, reason := nativeRichRuns(item, ns)
				joined := ""
				for _, r := range runs {
					joined += r.Text
				}
				if reason == "" && (joined != text || len(runs) != len(c.Value.Runs)) {
					reason = "Raw run boundaries do not join the opened source text."
				}
				if reason == "" {
					for i, r := range runs {
						if r.Text != c.Value.Runs[i].Text {
							reason = "Raw run boundaries do not join the opened source text."
						}
					}
				}
				if reason == "" && totalRuns+len(runs) > 1024 {
					reason = "Aggregate run budget exceeded."
				}
				if reason == "" {
					totalRuns += len(runs)
					entry.Status = "available"
					entry.Runs = runs
					entry.Warnings = []string{"Approximate direct-run preview: missing properties use the cell font as a host fallback for each run; font matching, shaping and omitted font-family hints are approximate. Explicit baseline and disabled effects are no-op declarations. No full rich-style or Excel fidelity claim."}
				} else {
					entry.Warnings = []string{"Run styling omitted: " + reason + " Plain source text retained."}
				}
			}
			result.Cells = append(result.Cells, entry)
		}
	}
	return result
}

// Match the public objects decoder's structural ceiling before attaching this
// optional supplement. A full older envelope must remain usable without it.
func nativeRichTextEnvelopeFits(value any) bool {
	encoded, err := json.Marshal(value)
	if err != nil || len(encoded) > 8*1024*1024 {
		return false
	}
	decoder := json.NewDecoder(bytes.NewReader(encoded))
	nodes := 0
	var read func(int) bool
	read = func(depth int) bool {
		nodes++
		if nodes > 200000 || depth > 12 {
			return false
		}
		token, err := decoder.Token()
		if err != nil {
			return false
		}
		delimiter, container := token.(json.Delim)
		if !container {
			return true
		}
		if delimiter != '{' && delimiter != '[' {
			return false
		}
		for decoder.More() {
			if delimiter == '{' {
				if _, err := decoder.Token(); err != nil {
					return false
				}
			}
			if !read(depth + 1) {
				return false
			}
		}
		_, err = decoder.Token()
		return err == nil
	}
	return read(0)
}
