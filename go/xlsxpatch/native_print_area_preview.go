package xlsxpatch

import (
	"regexp"
	"strconv"
	"strings"
)

// NativeSheetPrintAreaV1 is a bounded read-only projection of a saved local name.
// It grants no mutation authority and never evaluates a defined-name formula.
type NativeSheetPrintAreaV1 struct {
	SheetID   string                 `json:"sheet_id"`
	SheetPart string                 `json:"sheet_part"`
	Status    string                 `json:"status"`
	Warnings  []string               `json:"warnings"`
	Area      *NativePrintAreaRectV1 `json:"area,omitempty"`
}

// Coordinates are zero-based and inclusive.
type NativePrintAreaRectV1 struct {
	Row       int `json:"row"`
	Column    int `json:"column"`
	EndRow    int `json:"end_row"`
	EndColumn int `json:"end_column"`
}

var nativeAbsolutePrintRect = regexp.MustCompile(`^\$([A-Z]{1,3})\$([1-9][0-9]{0,6})(:\$([A-Z]{1,3})\$([1-9][0-9]{0,6}))?$`)
var nativeUnquotedPrintSheet = regexp.MustCompile(`^[A-Za-z_][A-Za-z0-9_.]*$`)

func previewNativePrintAreas(raw []byte, sheets []NativeWorkbookSheetV2) []NativeSheetPrintAreaV1 {
	count := len(sheets)
	if count > 64 {
		count = 64
	}
	result := make([]NativeSheetPrintAreaV1, count)
	for i, sheet := range sheets[:count] {
		result[i] = NativeSheetPrintAreaV1{SheetID: sheet.ID, SheetPart: sheet.PartName, Status: "unavailable", Warnings: []string{"Saved print area unavailable: requires exactly one worksheet-local absolute same-sheet A1 rectangle and no repeated print titles."}}
	}
	root, err := parsePreviewXML(raw)
	if err != nil || root.name.Local != "workbook" || !isSpreadsheetMLNamespace(root.name.Space) {
		return result
	}
	owner := root.child("definedNames")
	if owner != nil {
		for _, a := range owner.attrs {
			if !isPreviewNamespaceDeclaration(a) {
				return result
			}
		}
		for _, child := range owner.children {
			if child.name.Space != root.name.Space || child.name.Local != "definedName" {
				return result
			}
		}
	}
	areas := map[int][]*previewXML{}
	titles := map[int]bool{}
	ambiguous := false
	var visit func(*previewXML, *previewXML)
	visit = func(n, parent *previewXML) {
		if n.name.Local == "definedNames" && (parent != root || n != owner || n.name.Space != root.name.Space || strings.TrimSpace(n.text) != "") {
			ambiguous = true
		}
		if n.name.Local == "definedName" {
			name := n.attr("name")
			if (strings.EqualFold(name, "_xlnm.Print_Area") && name != "_xlnm.Print_Area") || (strings.EqualFold(name, "_xlnm.Print_Titles") && name != "_xlnm.Print_Titles") {
				ambiguous = true
			}
			// Foreign name/scope attributes must not hide a conflicting built-in.
			for _, a := range n.attrs {
				if a.Name.Local == "name" && (a.Value == "_xlnm.Print_Area" || a.Value == "_xlnm.Print_Titles") && a.Name.Space != "" {
					ambiguous = true
				}
			}
			if name == "_xlnm.Print_Area" || name == "_xlnm.Print_Titles" {
				if owner == nil || parent != owner || n.name.Space != root.name.Space {
					ambiguous = true
				}
				scope := n.attr("localSheetId")
				index, e := strconv.Atoi(scope)
				if e != nil || index < 0 || index >= len(sheets) || strconv.Itoa(index) != scope {
					ambiguous = true
				} else {
					if name == "_xlnm.Print_Titles" {
						titles[index] = true
					} else {
						areas[index] = append(areas[index], n)
					}
				}
			}
		}
		for _, child := range n.children {
			visit(child, n)
		}
	}
	visit(root, nil)
	if ambiguous {
		return result
	}
	for i, sheet := range sheets[:count] {
		// Order is the workbook ordinal, not the potentially sparse sheetId.
		defs := areas[sheet.Order]
		if len(defs) != 1 || titles[sheet.Order] {
			continue
		}
		n := defs[0]
		if len(n.children) != 0 {
			continue
		}
		valid := true
		for _, a := range n.attrs {
			if isPreviewNamespaceDeclaration(a) {
				continue
			}
			if a.Name.Space != "" || (a.Name.Local != "name" && a.Name.Local != "localSheetId") {
				valid = false
			}
		}
		if !valid {
			continue
		}
		area := parseNativePrintAreaRect(n.text, sheet.Name)
		if area == nil {
			continue
		}
		result[i].Status, result[i].Area = "available", area
		result[i].Warnings = []string{"Read-only saved print-area rectangle. Page geometry is approximate; formulas, repeated titles and printer behavior are not reproduced."}
	}
	return result
}

func parseNativePrintAreaRect(text, sheetName string) *NativePrintAreaRectV1 {
	if len(text) > 4096 {
		return nil
	}
	// A quoted sheet token may contain escaped apostrophes and literal ! characters.
	var sheet, ref string
	if strings.HasPrefix(text, "'") {
		var name strings.Builder
		closed := false
		for i := 1; i < len(text); i++ {
			if text[i] != '\'' {
				name.WriteByte(text[i])
				continue
			}
			if i+1 < len(text) && text[i+1] == '\'' {
				name.WriteByte('\'')
				i++
				continue
			}
			if i+1 >= len(text) || text[i+1] != '!' {
				return nil
			}
			sheet, ref, closed = name.String(), text[i+2:], true
			break
		}
		if !closed {
			return nil
		}
	} else {
		var found bool
		sheet, ref, found = strings.Cut(text, "!")
		if !found || !nativeUnquotedPrintSheet.MatchString(sheet) {
			return nil
		}
	}
	if sheet != sheetName || strings.ContainsAny(sheet, "[]:") {
		return nil
	}
	m := nativeAbsolutePrintRect.FindStringSubmatch(ref)
	if m == nil {
		return nil
	}
	column := func(s string) int {
		v := 0
		for _, c := range s {
			v = v*26 + int(c-'A'+1)
		}
		return v - 1
	}
	row, _ := strconv.Atoi(m[2])
	a := &NativePrintAreaRectV1{Row: row - 1, Column: column(m[1])}
	a.EndRow, a.EndColumn = a.Row, a.Column
	if m[3] != "" {
		end, _ := strconv.Atoi(m[5])
		a.EndRow, a.EndColumn = end-1, column(m[4])
	}
	if a.EndRow < a.Row || a.EndColumn < a.Column || a.EndRow >= 1048576 || a.EndColumn >= 16384 {
		return nil
	}
	return a
}
