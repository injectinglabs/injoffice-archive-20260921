package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"regexp"
	"strconv"
	"strings"
)

// PrintSetupInfo is the bounded, renderer-neutral projection of one native
// worksheet's current print configuration. Warnings identify native settings
// that were observed but cannot be represented without approximation.
type PrintSetupInfo struct {
	SheetName          string        `json:"sheetName"`
	PrintAreaRef       string        `json:"printAreaRef,omitempty"`
	RepeatRowsRef      string        `json:"repeatRowsRef,omitempty"`
	RepeatColumnsRef   string        `json:"repeatColumnsRef,omitempty"`
	Orientation        string        `json:"orientation,omitempty"`
	PaperSize          string        `json:"paperSize,omitempty"`
	FitToWidth         *int          `json:"fitToWidth,omitempty"`
	FitToHeight        *int          `json:"fitToHeight,omitempty"`
	Scale              *int          `json:"scale,omitempty"`
	Margins            *PrintMargins `json:"margins,omitempty"`
	OddHeader          string        `json:"oddHeader,omitempty"`
	OddFooter          string        `json:"oddFooter,omitempty"`
	HorizontalCentered bool          `json:"horizontalCentered"`
	VerticalCentered   bool          `json:"verticalCentered"`
	PrintGridlines     bool          `json:"printGridlines"`
	PrintHeadings      bool          `json:"printHeadings"`
	Warnings           []string      `json:"warnings,omitempty"`
}

type printDefinedNames struct {
	area, titles string
	warnings     []string
}

var printPaperNames = func() map[int]string {
	result := make(map[int]string, len(printPaperCodes))
	for name, code := range printPaperCodes {
		result[code] = name
	}
	return result
}()

// ReadPrintSetups hydrates the print settings that are safely representable
// from each worksheet. It does not render pages or infer application defaults.
func ReadPrintSetups(data []byte) ([]PrintSetupInfo, error) {
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: read print setup: %w", err)
	}
	if _, err := newOPCPackageIndex(zr); err != nil {
		return nil, fmt.Errorf("xlsxpatch: read print setup: %w", err)
	}
	files := make(map[string]*zip.File, len(zr.File))
	for _, file := range zr.File {
		files[file.Name] = file
	}
	read := func(name string) (string, bool) {
		file, ok := files[name]
		if !ok {
			return "", false
		}
		value, readErr := readZipFile(file)
		return string(value), readErr == nil
	}
	workbookXML, ok := read("xl/workbook.xml")
	if !ok {
		return nil, fmt.Errorf("xlsxpatch: read print setup: missing xl/workbook.xml")
	}
	if _, _, err := preflightNativeCoreXML([]byte(workbookXML)); err != nil {
		return nil, fmt.Errorf("xlsxpatch: read print setup: workbook XML: %w", err)
	}
	workbookRels, ok := read("xl/_rels/workbook.xml.rels")
	if !ok {
		return nil, fmt.Errorf("xlsxpatch: read print setup: missing workbook relationships")
	}
	if _, _, err := preflightNativeCoreXML([]byte(workbookRels)); err != nil {
		return nil, fmt.Errorf("xlsxpatch: read print setup: workbook relationships: %w", err)
	}
	sheets, err := readWorkbookSheets(read)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: read print setup: %w", err)
	}
	defined, err := readPrintDefinedNames(workbookXML)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: read print setup: %w", err)
	}
	result := make([]PrintSetupInfo, 0, len(sheets))
	for index, sheet := range sheets {
		worksheetXML, ok := read(sheet.Part)
		if !ok {
			return nil, fmt.Errorf("xlsxpatch: read print setup: worksheet part %q is missing", sheet.Part)
		}
		info, err := parseWorksheetPrintSetup([]byte(worksheetXML))
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: read print setup: worksheet %q: %w", sheet.Name, err)
		}
		info.SheetName = sheet.Name
		if names := defined[index]; names != nil {
			info.Warnings = append(info.Warnings, names.warnings...)
			if names.area != "" {
				if ref, warning := definedPrintAreaRef(names.area, sheet.Name); warning != "" {
					info.Warnings = append(info.Warnings, warning)
				} else {
					info.PrintAreaRef = ref
				}
			}
			if names.titles != "" {
				rows, columns, warnings := definedPrintTitleRefs(names.titles, sheet.Name)
				info.RepeatRowsRef, info.RepeatColumnsRef = rows, columns
				info.Warnings = append(info.Warnings, warnings...)
			}
		}
		result = append(result, info)
	}
	return result, nil
}

func readPrintDefinedNames(workbookXML string) (map[int]*printDefinedNames, error) {
	decoder := xml.NewDecoder(strings.NewReader(workbookXML))
	result := map[int]*printDefinedNames{}
	rootNamespace := ""
	depth := 0
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			return result, nil
		}
		if err != nil {
			return nil, fmt.Errorf("parse workbook defined names: %w", err)
		}
		if _, ok := token.(xml.EndElement); ok {
			depth--
			continue
		}
		start, ok := token.(xml.StartElement)
		if !ok {
			continue
		}
		depth++
		if depth == 1 {
			rootNamespace = start.Name.Space
		}
		if start.Name.Local != "definedName" || start.Name.Space != rootNamespace {
			continue
		}
		name := attrVal(start, "name")
		if name != "_xlnm.Print_Area" && name != "_xlnm.Print_Titles" {
			continue
		}
		index, parseErr := strconv.Atoi(attrVal(start, "localSheetId"))
		var value string
		if err := decoder.DecodeElement(&value, &start); err != nil {
			return nil, fmt.Errorf("parse %s: %w", name, err)
		}
		depth--
		if parseErr != nil || index < 0 {
			continue
		}
		entry := result[index]
		if entry == nil {
			entry = &printDefinedNames{}
			result[index] = entry
		}
		if name == "_xlnm.Print_Area" {
			if entry.area != "" {
				entry.warnings = append(entry.warnings, "multiple native print-area names are unsupported")
			} else {
				entry.area = strings.TrimSpace(value)
			}
		} else if entry.titles != "" {
			entry.warnings = append(entry.warnings, "multiple native print-title names are unsupported")
		} else {
			entry.titles = strings.TrimSpace(value)
		}
	}
}

func parseWorksheetPrintSetup(data []byte) (PrintSetupInfo, error) {
	if _, _, err := preflightNativeCoreXML(data); err != nil {
		return PrintSetupInfo{}, err
	}
	decoder := xml.NewDecoder(bytes.NewReader(data))
	result := PrintSetupInfo{}
	depth := 0
	rootNamespace := ""
	seen := map[string]bool{}
	collect := ""
	var text strings.Builder
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return PrintSetupInfo{}, err
		}
		switch token := token.(type) {
		case xml.StartElement:
			depth++
			if depth == 1 {
				if token.Name.Local != "worksheet" {
					return PrintSetupInfo{}, fmt.Errorf("root element is %q, expected worksheet", token.Name.Local)
				}
				rootNamespace = token.Name.Space
				if rootNamespace != "" && !isSpreadsheetMLNamespace(rootNamespace) {
					return PrintSetupInfo{}, fmt.Errorf("unsupported worksheet namespace %q", rootNamespace)
				}
			}
			if depth == 2 && token.Name.Space == rootNamespace {
				switch token.Name.Local {
				case "printOptions", "pageMargins", "pageSetup", "headerFooter":
					if seen[token.Name.Local] {
						return PrintSetupInfo{}, fmt.Errorf("duplicate %s element", token.Name.Local)
					}
					seen[token.Name.Local] = true
				}
				switch token.Name.Local {
				case "printOptions":
					parsePrintOptions(token, &result)
				case "pageMargins":
					parsePrintMargins(token, &result)
				case "pageSetup":
					parsePageSetup(token, &result)
				case "headerFooter":
					for _, attribute := range token.Attr {
						if !isNamespaceDeclaration(attribute) {
							result.Warnings = append(result.Warnings, fmt.Sprintf("header/footer option %s is not represented", attribute.Name.Local))
						}
					}
				}
			} else if depth == 3 && token.Name.Space == rootNamespace && (token.Name.Local == "oddHeader" || token.Name.Local == "oddFooter") {
				collect = token.Name.Local
				text.Reset()
			} else if depth > 3 && collect != "" {
				result.Warnings = append(result.Warnings, "nested native header/footer markup is unsupported")
			}
		case xml.CharData:
			if collect != "" {
				text.Write(token)
			}
		case xml.EndElement:
			if depth == 3 && token.Name.Local == collect {
				if collect == "oddHeader" {
					result.OddHeader = text.String()
				} else {
					result.OddFooter = text.String()
				}
				collect = ""
			}
			depth--
		}
	}
	return result, nil
}

func parsePrintOptions(start xml.StartElement, result *PrintSetupInfo) {
	known := map[string]bool{"horizontalCentered": true, "verticalCentered": true, "gridLines": true, "headings": true}
	for _, attribute := range start.Attr {
		if isNamespaceDeclaration(attribute) {
			continue
		}
		if !known[attribute.Name.Local] {
			result.Warnings = append(result.Warnings, fmt.Sprintf("print option %s is not represented", attribute.Name.Local))
			continue
		}
		value, ok := parsePrintBool(attribute.Value)
		if !ok {
			result.Warnings = append(result.Warnings, fmt.Sprintf("print option %s has invalid boolean %q", attribute.Name.Local, attribute.Value))
			continue
		}
		switch attribute.Name.Local {
		case "horizontalCentered":
			result.HorizontalCentered = value
		case "verticalCentered":
			result.VerticalCentered = value
		case "gridLines":
			result.PrintGridlines = value
		case "headings":
			result.PrintHeadings = value
		}
	}
}

func parsePrintMargins(start xml.StartElement, result *PrintSetupInfo) {
	values := map[string]float64{}
	for _, attribute := range start.Attr {
		if isNamespaceDeclaration(attribute) {
			continue
		}
		if attribute.Name.Local != "left" && attribute.Name.Local != "right" && attribute.Name.Local != "top" && attribute.Name.Local != "bottom" && attribute.Name.Local != "header" && attribute.Name.Local != "footer" {
			result.Warnings = append(result.Warnings, fmt.Sprintf("page margin %s is not represented", attribute.Name.Local))
			continue
		}
		value, err := strconv.ParseFloat(attribute.Value, 64)
		if err != nil || value < 0 || value > 20 {
			result.Warnings = append(result.Warnings, fmt.Sprintf("page margin %s is invalid", attribute.Name.Local))
			return
		}
		values[attribute.Name.Local] = value
	}
	if len(values) != 6 {
		result.Warnings = append(result.Warnings, "page margins are incomplete")
		return
	}
	result.Margins = &PrintMargins{Left: values["left"], Right: values["right"], Top: values["top"], Bottom: values["bottom"], Header: values["header"], Footer: values["footer"]}
}

func parsePageSetup(start xml.StartElement, result *PrintSetupInfo) {
	known := map[string]bool{"orientation": true, "paperSize": true, "fitToWidth": true, "fitToHeight": true, "scale": true}
	for _, attribute := range start.Attr {
		if isNamespaceDeclaration(attribute) {
			continue
		}
		if !known[attribute.Name.Local] {
			result.Warnings = append(result.Warnings, fmt.Sprintf("page setup option %s is not represented", attribute.Name.Local))
			continue
		}
		switch attribute.Name.Local {
		case "orientation":
			if attribute.Value == "portrait" || attribute.Value == "landscape" {
				result.Orientation = attribute.Value
			} else {
				result.Warnings = append(result.Warnings, fmt.Sprintf("page orientation %q is unsupported", attribute.Value))
			}
		case "paperSize":
			code, err := strconv.Atoi(attribute.Value)
			if err != nil || printPaperNames[code] == "" {
				result.Warnings = append(result.Warnings, fmt.Sprintf("paper size code %q is unsupported", attribute.Value))
			} else {
				result.PaperSize = printPaperNames[code]
			}
		case "fitToWidth", "fitToHeight", "scale":
			value, err := strconv.Atoi(attribute.Value)
			if err != nil || value < 0 || value > 100_000 {
				result.Warnings = append(result.Warnings, fmt.Sprintf("page setup option %s is invalid", attribute.Name.Local))
				continue
			}
			if attribute.Name.Local == "fitToWidth" {
				result.FitToWidth = &value
			} else if attribute.Name.Local == "fitToHeight" {
				result.FitToHeight = &value
			} else if value < 10 || value > 400 {
				result.Warnings = append(result.Warnings, "page scale is outside 10..400")
			} else {
				result.Scale = &value
			}
		}
	}
}

func parsePrintBool(value string) (bool, bool) {
	switch value {
	case "1", "true":
		return true, true
	case "0", "false":
		return false, true
	default:
		return false, false
	}
}

func splitPrintNameFormula(value string) []string {
	var result []string
	start, quoted := 0, false
	for index := 0; index < len(value); index++ {
		if value[index] == '\'' {
			if quoted && index+1 < len(value) && value[index+1] == '\'' {
				index++
				continue
			}
			quoted = !quoted
		} else if value[index] == ',' && !quoted {
			result = append(result, strings.TrimSpace(value[start:index]))
			start = index + 1
		}
	}
	return append(result, strings.TrimSpace(value[start:]))
}

func splitQualifiedPrintRef(value string) (sheet, ref string, ok bool) {
	index, quoted := -1, false
	for cursor := 0; cursor < len(value); cursor++ {
		if value[cursor] == '\'' {
			if quoted && cursor+1 < len(value) && value[cursor+1] == '\'' {
				cursor++
				continue
			}
			quoted = !quoted
		} else if value[cursor] == '!' && !quoted {
			index = cursor
		}
	}
	if index <= 0 || index == len(value)-1 {
		return "", "", false
	}
	sheet, ref = value[:index], value[index+1:]
	if len(sheet) >= 2 && sheet[0] == '\'' && sheet[len(sheet)-1] == '\'' {
		sheet = strings.ReplaceAll(sheet[1:len(sheet)-1], "''", "'")
	}
	return sheet, ref, true
}

func definedPrintAreaRef(value, sheetName string) (string, string) {
	parts := splitPrintNameFormula(value)
	if len(parts) != 1 {
		return "", "multiple native print areas are not representable as one range"
	}
	sheet, ref, ok := splitQualifiedPrintRef(parts[0])
	if !ok || sheet != sheetName || !printAreaRe.MatchString(ref) {
		return "", "native print area is malformed or targets another worksheet"
	}
	return ref, ""
}

var (
	printTitleRowsRe    = regexp.MustCompile(`^\$?([0-9]+):\$?([0-9]+)$`)
	printTitleColumnsRe = regexp.MustCompile(`^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$`)
)

func definedPrintTitleRefs(value, sheetName string) (rows, columns string, warnings []string) {
	for _, part := range splitPrintNameFormula(value) {
		sheet, ref, ok := splitQualifiedPrintRef(part)
		if !ok || sheet != sheetName {
			warnings = append(warnings, "native print title is malformed or targets another worksheet")
			continue
		}
		if printTitleRowsRe.MatchString(ref) {
			if rows != "" {
				warnings = append(warnings, "multiple repeated row ranges are unsupported")
			} else {
				rows = ref
			}
		} else if printTitleColumnsRe.MatchString(ref) {
			if columns != "" {
				warnings = append(warnings, "multiple repeated column ranges are unsupported")
			} else {
				columns = ref
			}
		} else {
			warnings = append(warnings, fmt.Sprintf("native print title range %q is unsupported", ref))
		}
	}
	return rows, columns, warnings
}
