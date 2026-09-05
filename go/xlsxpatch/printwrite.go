package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"regexp"
	"strings"
)

// Print setup (InjOffice Phase 8): written INTO the workbook itself —
// <pageSetup>, <pageMargins>, <headerFooter> on the worksheet plus the
// _xlnm.Print_Area defined name — so Excel/LibreOffice/Google Sheets all see
// the same settings, and the PDF render path (LibreOffice honors these parts)
// needs no side-channel. Composed as a Patch through Apply, same fidelity
// guarantees as charts and pivots.

// PrintSetup describes one sheet's print configuration. Zero-value fields
// mean "leave/omit" except Orientation which defaults to portrait when the
// element is written.
type PrintSetup struct {
	SheetName string
	// PrintAreaRef like "A1:F20"; empty removes any existing print area for
	// the sheet.
	PrintAreaRef string
	// Orientation: "portrait" | "landscape".
	Orientation string
	// PaperSize is a named ISO/US size supported by the public TypeScript
	// facade. Empty preserves the application default.
	PaperSize string
	// FitToWidth pages across (1 = classic fit-to-width). 0 omits fitting.
	FitToWidth int
	// FitToHeight pages down; 0 with FitToWidth>0 means "as many as needed".
	FitToHeight int
	// Scale is an explicit percentage within 10..400. It is mutually
	// exclusive with fit-to-page settings; zero omits explicit scaling.
	Scale int
	// Margins in inches; nil uses Excel's normal preset.
	Margins *PrintMargins
	// Header/Footer center text; supports the OOXML tokens (&P page, &N
	// pages, &D date, &F file). Empty omits the element half.
	HeaderCenter string
	FooterCenter string
	// Worksheet printOptions supported by SpreadsheetML.
	HorizontalCentered bool
	VerticalCentered   bool
	PrintGridlines     bool
	PrintHeadings      bool
}

// PrintMargins in inches (OOXML's unit for pageMargins).
type PrintMargins struct {
	Left   float64 `json:"left"`
	Right  float64 `json:"right"`
	Top    float64 `json:"top"`
	Bottom float64 `json:"bottom"`
	Header float64 `json:"header"`
	Footer float64 `json:"footer"`
}

var normalMargins = PrintMargins{Left: 0.7, Right: 0.7, Top: 0.75, Bottom: 0.75, Header: 0.3, Footer: 0.3}

var printPaperCodes = map[string]int{
	"Letter": 1, "Tabloid": 3, "Legal": 5, "Statement": 6,
	"Executive": 7, "A3": 8, "A4": 9, "A5": 11, "B4": 12,
	"B5": 13, "Folio": 14,
}

// SetPrintSetup returns new workbook bytes with the sheet's print
// configuration applied.
func SetPrintSetup(orig []byte, setup PrintSetup) ([]byte, error) {
	if setup.Orientation != "" && setup.Orientation != "portrait" && setup.Orientation != "landscape" {
		return nil, fmt.Errorf("xlsxpatch: orientation must be portrait|landscape")
	}
	if setup.PaperSize != "" {
		if _, ok := printPaperCodes[setup.PaperSize]; !ok {
			return nil, fmt.Errorf("xlsxpatch: unsupported paper size %q", setup.PaperSize)
		}
	}
	if setup.Scale != 0 && (setup.Scale < 10 || setup.Scale > 400) {
		return nil, fmt.Errorf("xlsxpatch: scale must be within 10..400")
	}
	if setup.FitToWidth < 0 || setup.FitToHeight < 0 {
		return nil, fmt.Errorf("xlsxpatch: fit dimensions must be non-negative")
	}
	if setup.Scale != 0 && (setup.FitToWidth != 0 || setup.FitToHeight != 0) {
		return nil, fmt.Errorf("xlsxpatch: scale and fit-to-page settings are mutually exclusive")
	}
	if setup.PrintAreaRef != "" && !printAreaRe.MatchString(setup.PrintAreaRef) {
		return nil, fmt.Errorf("xlsxpatch: invalid print area %q", setup.PrintAreaRef)
	}

	zr, err := zip.NewReader(bytes.NewReader(orig), int64(len(orig)))
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: print setup: %w", err)
	}
	read := func(name string) (string, bool) {
		for _, f := range zr.File {
			if f.Name == name {
				rc, err := f.Open()
				if err != nil {
					return "", false
				}
				defer rc.Close()
				b, err := io.ReadAll(rc)
				if err != nil {
					return "", false
				}
				return string(b), true
			}
		}
		return "", false
	}

	sheetPart, err := worksheetPartFor(read, setup.SheetName)
	if err != nil {
		return nil, err
	}
	sheetXML, ok := read(sheetPart)
	if !ok {
		return nil, fmt.Errorf("xlsxpatch: worksheet part %q unreadable", sheetPart)
	}
	wbXML, ok := read("xl/workbook.xml")
	if !ok {
		return nil, fmt.Errorf("xlsxpatch: missing xl/workbook.xml")
	}

	newSheet, err := worksheetWithPrintSetup(sheetXML, setup)
	if err != nil {
		return nil, err
	}
	sheetIdx, err := sheetIndexFor(wbXML, setup.SheetName)
	if err != nil {
		return nil, err
	}
	newWb, err := workbookWithPrintArea(wbXML, setup.SheetName, sheetIdx, setup.PrintAreaRef)
	if err != nil {
		return nil, err
	}

	patch := Patch{Replace: map[string][]byte{
		sheetPart:         []byte(newSheet),
		"xl/workbook.xml": []byte(newWb),
	}}
	return Apply(orig, patch)
}

var printAreaRe = regexp.MustCompile(`^\$?[A-Za-z]{1,3}\$?\d+:\$?[A-Za-z]{1,3}\$?\d+$`)

// element removal/insertion helpers: worksheet print elements have a strict
// schema position — pageMargins, then pageSetup, then headerFooter — sitting
// after the drawing-preceding content but BEFORE drawing/tableParts/extLst.
var (
	pageMarginsRe  = regexp.MustCompile(`<pageMargins\b[^>]*/>`)
	printOptionsRe = regexp.MustCompile(`<printOptions\b[^>]*/>`)
	pageSetupRe    = regexp.MustCompile(`<pageSetup\b[^>]*/>`)
	headerFooterRe = regexp.MustCompile(`<headerFooter\b[^>]*>.*?</headerFooter>|<headerFooter\b[^>]*/>`)
	pageSetUpPrRe  = regexp.MustCompile(`<pageSetUpPr\b[^>]*/>|<pageSetUpPr\b[^>]*>.*?</pageSetUpPr>`)
)

func worksheetWithPrintSetup(sheetXML string, setup PrintSetup) (string, error) {
	// Drop existing print elements; we re-emit the full trio in order.
	sheetXML = pageMarginsRe.ReplaceAllString(sheetXML, "")
	sheetXML = printOptionsRe.ReplaceAllString(sheetXML, "")
	sheetXML = pageSetupRe.ReplaceAllString(sheetXML, "")
	sheetXML = headerFooterRe.ReplaceAllString(sheetXML, "")

	m := setup.Margins
	if m == nil {
		m = &normalMargins
	}
	var b strings.Builder
	if setup.HorizontalCentered || setup.VerticalCentered || setup.PrintGridlines || setup.PrintHeadings {
		b.WriteString(`<printOptions`)
		if setup.HorizontalCentered {
			b.WriteString(` horizontalCentered="1"`)
		}
		if setup.VerticalCentered {
			b.WriteString(` verticalCentered="1"`)
		}
		if setup.PrintGridlines {
			b.WriteString(` gridLines="1"`)
		}
		if setup.PrintHeadings {
			b.WriteString(` headings="1"`)
		}
		b.WriteString(`/>`)
	}
	fmt.Fprintf(&b, `<pageMargins left="%g" right="%g" top="%g" bottom="%g" header="%g" footer="%g"/>`,
		m.Left, m.Right, m.Top, m.Bottom, m.Header, m.Footer)

	orientation := setup.Orientation
	if orientation == "" {
		orientation = "portrait"
	}
	pageAttributes := ""
	if setup.FitToWidth > 0 || setup.FitToHeight > 0 {
		h := setup.FitToHeight
		pageAttributes += fmt.Sprintf(` fitToWidth="%d" fitToHeight="%d"`, setup.FitToWidth, h)
	}
	if setup.Scale > 0 {
		pageAttributes += fmt.Sprintf(` scale="%d"`, setup.Scale)
	}
	if setup.PaperSize != "" {
		pageAttributes += fmt.Sprintf(` paperSize="%d"`, printPaperCodes[setup.PaperSize])
	}
	fmt.Fprintf(&b, `<pageSetup orientation=%q%s/>`, orientation, pageAttributes)

	if setup.HeaderCenter != "" || setup.FooterCenter != "" {
		b.WriteString(`<headerFooter>`)
		if setup.HeaderCenter != "" {
			fmt.Fprintf(&b, `<oddHeader>&amp;C%s</oddHeader>`, esc(setup.HeaderCenter))
		}
		if setup.FooterCenter != "" {
			fmt.Fprintf(&b, `<oddFooter>&amp;C%s</oddFooter>`, esc(setup.FooterCenter))
		}
		b.WriteString(`</headerFooter>`)
	}

	// FitToWidth requires <sheetPr><pageSetUpPr fitToPage="1"/></sheetPr> or
	// Excel ignores the fit attributes. An EXISTING pageSetUpPr (openpyxl
	// writes an empty one) must be REPLACED, not duplicated — readers take
	// the first occurrence, so a duplicate silently disables fitting (caught
	// by external openpyxl validation).
	if setup.FitToWidth > 0 || setup.FitToHeight > 0 {
		switch {
		case pageSetUpPrRe.MatchString(sheetXML):
			sheetXML = pageSetUpPrRe.ReplaceAllString(sheetXML, `<pageSetUpPr fitToPage="1"/>`)
		case strings.Contains(sheetXML, "<sheetPr"):
			i := strings.Index(sheetXML, "<sheetPr")
			j := strings.Index(sheetXML[i:], ">")
			if j < 0 {
				return "", fmt.Errorf("xlsxpatch: malformed sheetPr")
			}
			if sheetXML[i+j-1] == '/' {
				// self-closing <sheetPr/> → expand it.
				sheetXML = sheetXML[:i] + `<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>` + sheetXML[i+j+1:]
			} else {
				// pageSetUpPr is the LAST child inside sheetPr per schema.
				if end := strings.Index(sheetXML, "</sheetPr>"); end >= 0 {
					sheetXML = sheetXML[:end] + `<pageSetUpPr fitToPage="1"/>` + sheetXML[end:]
				}
			}
		default:
			rootEnd := strings.Index(sheetXML, ">")
			if rootEnd < 0 {
				return "", fmt.Errorf("xlsxpatch: malformed worksheet part")
			}
			sheetXML = sheetXML[:rootEnd+1] + `<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>` + sheetXML[rootEnd+1:]
		}
	}

	// Insert before the first element that must follow the print trio.
	insertAt := strings.LastIndex(sheetXML, "</worksheet>")
	if insertAt < 0 {
		return "", fmt.Errorf("xlsxpatch: malformed worksheet part")
	}
	for _, follower := range []string{"<headerFooter", "<rowBreaks", "<colBreaks", "<drawing", "<legacyDrawing", "<picture", "<oleObjects", "<tableParts", "<extLst"} {
		if i := strings.Index(sheetXML, follower); i >= 0 && i < insertAt {
			insertAt = i
		}
	}
	return sheetXML[:insertAt] + b.String() + sheetXML[insertAt:], nil
}

// sheetIndexFor returns the ZERO-based position of the sheet among <sheet>
// elements — the localSheetId a Print_Area defined name binds to.
func sheetIndexFor(wbXML, sheetName string) (int, error) {
	idx := 0
	for _, m := range regexp.MustCompile(`<sheet\b[^>]*name="([^"]*)"`).FindAllStringSubmatch(wbXML, -1) {
		if m[1] == sheetName {
			return idx, nil
		}
		idx++
	}
	return 0, fmt.Errorf("xlsxpatch: sheet %q not found for print area", sheetName)
}

func workbookWithPrintArea(wbXML, sheetName string, sheetIdx int, areaRef string) (string, error) {
	// Remove any existing Print_Area for this localSheetId.
	existing := regexp.MustCompile(
		fmt.Sprintf(`<definedName name="_xlnm.Print_Area" localSheetId="%d"[^>]*>[^<]*</definedName>`, sheetIdx))
	wbXML = existing.ReplaceAllString(wbXML, "")

	if areaRef == "" {
		return wbXML, nil
	}
	quoted := sheetName
	if strings.ContainsAny(sheetName, " '") {
		quoted = "'" + strings.ReplaceAll(sheetName, "'", "''") + "'"
	}
	// Absolute-ize the ref ($A$1:$F$20) — Excel writes Print_Area absolute.
	abs := regexp.MustCompile(`\$?([A-Za-z]{1,3})\$?(\d+)`).ReplaceAllString(areaRef, `$$$1$$$2`)
	entry := fmt.Sprintf(`<definedName name="_xlnm.Print_Area" localSheetId="%d">%s!%s</definedName>`,
		sheetIdx, esc(quoted), abs)

	if i := strings.Index(wbXML, "<definedNames>"); i >= 0 {
		at := i + len("<definedNames>")
		return wbXML[:at] + entry + wbXML[at:], nil
	}
	if i := strings.Index(wbXML, "<definedNames/>"); i >= 0 {
		return wbXML[:i] + "<definedNames>" + entry + "</definedNames>" + wbXML[i+len("<definedNames/>"):], nil
	}
	// No definedNames element: schema places it after <sheets>.
	if i := strings.Index(wbXML, "</sheets>"); i >= 0 {
		at := i + len("</sheets>")
		return wbXML[:at] + "<definedNames>" + entry + "</definedNames>" + wbXML[at:], nil
	}
	return "", fmt.Errorf("xlsxpatch: workbook.xml has no <sheets> element")
}
