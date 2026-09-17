package xlsxpatch

import (
	"strconv"
	"strings"
)

// Source settings, plus the ECMA-376 attribute defaults the source relies on.
// This is not a printer-default resolver.
//
// ECMA-376 Part 1 §18.3.1.63 declares CT_PageSetup/@paperSize default="1",
// @orientation default="default" and @scale default="100": an omitted attribute
// carries a stated value rather than no value, and Excel prints such a
// worksheet at US Letter portrait 100%. Those defaults are supplied here and
// named in Defaults so the preview tier labels the resulting geometry as a host
// default instead of an authored one. Nothing that changes pagination is
// defaulted: manual page breaks, printed headings, page centering, non-Letter
// and non-A4 paper, and fit-to-page without its dimensions still refuse.
//
// Status "margins-only" is still a source fact and not a default: it asserts
// that the worksheet carries no pageSetup element at all, that its authored
// margins are bounded and exact, and that nothing else on the sheet blocks
// pagination. Choosing paper for such a sheet is a host decision, made and
// disclosed by the preview tier, never here.
type NativeSheetPageSettingsV1 struct {
	SheetID   string                    `json:"sheet_id"`
	SheetPart string                    `json:"sheet_part"`
	Status    string                    `json:"status"`
	Settings  *NativeSheetPageConfigV1  `json:"settings,omitempty"`
	Margins   *NativeSheetPageMarginsV1 `json:"margins,omitempty"`
	// Facts in Settings taken from an ECMA-376 attribute default because the
	// source omitted the attribute. Present only with status "available".
	Defaults []string `json:"defaulted,omitempty"`
	Warnings []string `json:"warnings"`
}

// Authored page margins for a worksheet that declares no pageSetup.
//
// Header and footer are the authored distances from the paper edge to the
// header and footer bands, not additions to Top and Bottom. ECMA-376 Part 1
// §18.3.1.62 measures all six from the edge, and Excel prints the body between
// max(Top, Header) and max(Bottom, Footer), so a header margin deeper than the
// top margin takes body height away. They are carried here so a consumer can
// apply that rule instead of silently printing a taller body than Excel does.
type NativeSheetPageMarginsV1 struct {
	Left   float64 `json:"left_inches"`
	Right  float64 `json:"right_inches"`
	Top    float64 `json:"top_inches"`
	Bottom float64 `json:"bottom_inches"`
	Header float64 `json:"header_inches"`
	Footer float64 `json:"footer_inches"`
}
type NativeSheetPageConfigV1 struct {
	Paper       string  `json:"paper"`
	Orientation string  `json:"orientation"`
	Scale       int     `json:"scale"`
	Left        float64 `json:"left_inches"`
	Right       float64 `json:"right_inches"`
	Top         float64 `json:"top_inches"`
	Bottom      float64 `json:"bottom_inches"`
	// Authored header/footer margins, measured from the paper edge like the
	// other four. See NativeSheetPageMarginsV1 for the body-height rule.
	Header    float64 `json:"header_inches"`
	Footer    float64 `json:"footer_inches"`
	PageOrder string  `json:"page_order,omitempty"`
	// Scale remains a compatibility value; fit mode ignores it.
	FitToPage *NativeSheetFitToPageV1 `json:"fit_to_page,omitempty"`
}

type NativeSheetFitToPageV1 struct {
	Width  int `json:"width"`
	Height int `json:"height"`
}

const (
	nativePageDefaultsDisclosure     = "Paper, orientation or scale is an ECMA-376 CT_PageSetup attribute default rather than an authored workbook value; the source omits the attribute and Excel prints the same default. Authored facts and authored margins are used unchanged."
	nativePageHeaderFooterNotPainted = "Worksheet declares headerFooter content. Header and footer text is not painted. The authored header and footer margins are reserved as Excel reserves them, so the body starts below max(top, header) and ends above max(bottom, footer); an overlong header that Excel would grow that reservation for is not reproduced."
	nativePageGridlinesNotPainted    = "Worksheet declares printOptions. Printed gridlines are not painted; printed row and column headings and page centering are not defaulted and still refuse."
	nativePagePrinterFactsIgnored    = "Printer-directed page setup attributes are not resolved: there is no printer, so printer defaults, copies, draft, black-and-white and printer DPI are ignored and 96 CSS px/in is the preview raster."
)

func previewNativePageSettings(raw []byte, part, id string) NativeSheetPageSettingsV1 {
	result := NativeSheetPageSettingsV1{SheetID: id, SheetPart: part, Status: "unavailable", Warnings: []string{"Worksheet page preview requires explicit bounded paper, orientation, margins and percentage scale or fit-to-page settings; printer defaults and unsupported print options are not inferred."}}
	root, err := parsePreviewXML(raw)
	if err != nil || root.name.Local != "worksheet" || !isSpreadsheetMLNamespace(root.name.Space) {
		return result
	}
	setup, margins := root.child("pageSetup"), root.child("pageMargins")
	if margins == nil {
		return result
	}
	var fitPr *previewXML
	sheetPrCount, pageSetupCount, headerFooterCount, printOptionsCount := 0, 0, 0, 0
	// Do not silently ignore alternate/foreign settings or page-affecting data.
	for _, child := range root.children {
		switch child.name.Local {
		case "pageSetup", "pageMargins":
			if child.name.Space != root.name.Space {
				return result
			}
			if child.name.Local == "pageSetup" {
				pageSetupCount++
			}
		case "rowBreaks", "colBreaks", "pageSetUpPr":
			// Manual page breaks divide the sheet into pages by hand; there is
			// no default for where they fall, so they keep refusing. A
			// top-level pageSetUpPr is misplaced markup, not a setting.
			return result
		case "headerFooter":
			// Header and footer content sits inside the header/footer margin,
			// which ECMA-376 measures from the paper edge like the body
			// margins. The authored header and footer margins are reported
			// either way, so the presence of text does not change the body
			// rectangle beyond the reservation those margins already make.
			headerFooterCount++
			if headerFooterCount > 1 || child.name.Space != root.name.Space {
				return result
			}
		case "printOptions":
			printOptionsCount++
			if printOptionsCount > 1 || child.name.Space != root.name.Space || !printOptionsWithoutPageGeometry(child) {
				return result
			}
		case "sheetPr":
			sheetPrCount++
			if sheetPrCount > 1 || child.name.Space != root.name.Space || strings.TrimSpace(child.text) != "" {
				return result
			}
			for _, a := range child.attrs {
				if isPreviewNamespaceDeclaration(a) {
					continue
				}
				if a.Name.Space != "" || !nativeSheetPrNonPageAttr(a.Name.Local) {
					return result
				}
			}
			for _, p := range child.children {
				if p.name.Local != "pageSetUpPr" {
					continue
				}
				if fitPr != nil || p.name.Space != root.name.Space {
					return result
				}
				fitPr = p
			}
		}
	}
	// Reject a second or misplaced activation flag rather than treating it as
	// inactive percentage settings. Local-name matching also catches foreign XML.
	var misplacedFit func(*previewXML) bool
	misplacedFit = func(n *previewXML) bool {
		if n.name.Local == "pageSetUpPr" && n != fitPr {
			return true
		}
		for _, child := range n.children {
			if misplacedFit(child) {
				return true
			}
		}
		return false
	}
	if misplacedFit(root) {
		return result
	}
	// pageSetUpPr carries the fit-to-page switch and nothing else this tier can
	// honour: autoPageBreaks="false" suppresses automatic pagination, which has
	// no default rectangle, so only a lone fitToPage is read. An explicit
	// fitToPage="false" is a source fact stating percentage scale applies.
	fitActive := false
	if fitPr != nil {
		if !previewPageExactLeaf(fitPr, "fitToPage") {
			return result
		}
		active, ok := previewRowBool(fitPr.attr("fitToPage"))
		if !ok {
			return result
		}
		fitActive = active
	}
	readMargins := func() (map[string]float64, bool) {
		if !previewPageExactLeaf(margins, "left", "right", "top", "bottom", "header", "footer") {
			return nil, false
		}
		values := map[string]float64{}
		for _, name := range []string{"left", "right", "top", "bottom", "header", "footer"} {
			value, ok := boundedPreviewRowNumber(margins.attr(name), 20)
			if !ok {
				return nil, false
			}
			values[name] = value
		}
		return values, true
	}
	// A worksheet that declares no pageSetup at all has authored margins and no
	// authored paper. Report that precisely instead of collapsing it into the
	// same "unavailable" as a pageSetup this tier cannot support: the two differ
	// in whether a host may supply paper of its own. Fit-to-page activation
	// without the pageSetup that carries its dimensions stays unavailable.
	//
	// child() returns nil for a duplicated element as well as an absent one, so
	// the count is what separates "no paper is authored" from "paper is authored
	// ambiguously"; only the former may be defaulted.
	if setup == nil {
		if fitActive || pageSetupCount != 0 {
			return result
		}
		values, ok := readMargins()
		if !ok {
			return result
		}
		result.Status = "margins-only"
		result.Margins = &NativeSheetPageMarginsV1{Left: values["left"], Right: values["right"], Top: values["top"], Bottom: values["bottom"], Header: values["header"], Footer: values["footer"]}
		result.Warnings = []string{"Worksheet declares authored page margins and no pageSetup element, so no paper size, orientation or scale is authored. Page geometry requires a host paper choice; this tier does not select one."}
		if headerFooterCount == 1 {
			result.Warnings = append(result.Warnings, nativePageHeaderFooterNotPainted)
		}
		if printOptionsCount == 1 {
			result.Warnings = append(result.Warnings, nativePageGridlinesNotPainted)
		}
		return result
	}
	var required []string
	var fit *NativeSheetFitToPageV1
	if fitActive {
		width, ew := strconv.Atoi(setup.attr("fitToWidth"))
		height, eh := strconv.Atoi(setup.attr("fitToHeight"))
		if ew != nil || eh != nil || width < 0 || width > 100 || height < 0 || height > 100 || width+height == 0 || strconv.Itoa(width) != setup.attr("fitToWidth") || strconv.Itoa(height) != setup.attr("fitToHeight") {
			return result
		}
		fit = &NativeSheetFitToPageV1{Width: width, Height: height}
		required = []string{"fitToWidth", "fitToHeight"}
	}
	if !pageSetupSupportedAttrs(setup, required) {
		return result
	}
	// useFirstPageNumber="false" (the default) means Excel numbers from 1 and
	// ignores firstPageNumber. A stated start other than 1 is a real numbering
	// fact this tier cannot carry, so it refuses rather than renumbering.
	if useFirst, _ := previewRowBool(setup.attr("useFirstPageNumber")); useFirst && setup.attr("firstPageNumber") != "1" && setup.attr("firstPageNumber") != "" {
		return result
	}
	var defaults []string
	paper := "Letter"
	if raw := setup.attr("paperSize"); raw == "" {
		defaults = append(defaults, "paper")
	} else if paper = map[string]string{"1": "Letter", "9": "A4"}[raw]; paper == "" {
		return result
	}
	orientation := setup.attr("orientation")
	if orientation == "" || orientation == "default" {
		orientation, defaults = "portrait", append(defaults, "orientation")
	} else if orientation != "portrait" && orientation != "landscape" {
		return result
	}
	scaleText := setup.attr("scale")
	if scaleText == "" {
		// Fit mode ignores the percentage entirely, so the compatibility 100
		// it carries there is not a defaulted geometry fact.
		scaleText = "100"
		if !fitActive {
			defaults = append(defaults, "scale")
		}
	}
	scale, err := strconv.Atoi(scaleText)
	if err != nil || scale < 10 || scale > 400 || strconv.Itoa(scale) != scaleText {
		return result
	}
	order := setup.attr("pageOrder")
	if order != "" && order != "downThenOver" && order != "overThenDown" {
		return result
	}
	values, ok := readMargins()
	if !ok {
		return result
	}
	result.Status = "available"
	result.Defaults = defaults
	result.Settings = &NativeSheetPageConfigV1{Paper: paper, Orientation: orientation, Scale: scale, Left: values["left"], Right: values["right"], Top: values["top"], Bottom: values["bottom"], Header: values["header"], Footer: values["footer"], PageOrder: order, FitToPage: fit}
	result.Warnings = []string{"Read-only selected-range page geometry approximation. Page settings do not select a range; repeated titles, chart paint, headers and printer behavior are not reproduced by page settings alone. No Excel fidelity claim."}
	if fit != nil {
		result.Warnings = append(result.Warnings, "Explicit fit-to-page dimensions apply to the selected preview range. A zero dimension is unconstrained. Percentage scale is ignored in fit mode; absent source scale is represented as 100 for compatibility.")
	}
	if len(defaults) != 0 {
		result.Warnings = append(result.Warnings, nativePageDefaultsDisclosure)
	}
	if headerFooterCount == 1 {
		result.Warnings = append(result.Warnings, nativePageHeaderFooterNotPainted)
	}
	if printOptionsCount == 1 {
		result.Warnings = append(result.Warnings, nativePageGridlinesNotPainted)
	}
	for _, name := range []string{"usePrinterDefaults", "copies", "draft", "blackAndWhite", "horizontalDpi", "verticalDpi"} {
		if setup.attr(name) != "" {
			result.Warnings = append(result.Warnings, nativePagePrinterFactsIgnored)
			break
		}
	}
	return result
}

// CT_SheetPr attributes. None of them describes page geometry: they carry
// filtering, calculation, synchronisation and code-name state, so a sheetPr
// that also holds pageSetUpPr stays readable instead of refusing wholesale.
func nativeSheetPrNonPageAttr(name string) bool {
	switch name {
	case "syncHorizontal", "syncVertical", "syncRef", "transitionEvaluation", "transitionEntry",
		"published", "codeName", "filterMode", "enableFormatConditionsCalculation":
		return true
	}
	return false
}

// CT_PrintOptions is accepted only while every flag that would move the printed
// block is off. gridLines and gridLinesSet are paint-only.
func printOptionsWithoutPageGeometry(node *previewXML) bool {
	if len(node.children) != 0 || strings.TrimSpace(node.text) != "" {
		return false
	}
	seen := map[string]int{}
	for _, a := range node.attrs {
		if isPreviewNamespaceDeclaration(a) {
			continue
		}
		if a.Name.Space != "" {
			return false
		}
		value, ok := previewRowBool(a.Value)
		if !ok {
			return false
		}
		seen[a.Name.Local]++
		if seen[a.Name.Local] != 1 {
			return false
		}
		switch a.Name.Local {
		case "gridLines", "gridLinesSet":
		case "headings", "horizontalCentered", "verticalCentered":
			// Printed headings consume body area and centering repositions the
			// printed block. Neither has a default this tier could apply.
			if value {
				return false
			}
		default:
			return false
		}
	}
	return true
}

func previewPageExactLeaf(n *previewXML, names ...string) bool {
	if len(n.children) != 0 || strings.TrimSpace(n.text) != "" {
		return false
	}
	allowed := map[string]bool{}
	for _, name := range names {
		allowed[name] = true
	}
	count := 0
	for _, a := range n.attrs {
		if isPreviewNamespaceDeclaration(a) {
			continue
		}
		if a.Name.Space != "" || !allowed[a.Name.Local] {
			return false
		}
		count++
	}
	return count == len(names)
}

func excelSentinelPrintDPI(raw string) bool {
	switch raw {
	case "0", "4294967292", "4294967293", "4294967294", "4294967295":
		return true
	default:
		return false
	}
}

func previewPageUnsignedAttr(raw string) bool {
	value, err := strconv.ParseUint(raw, 10, 32)
	return err == nil && strconv.FormatUint(value, 10) == raw
}

func previewPageBooleanAttr(raw string) bool {
	_, ok := previewRowBool(raw)
	return ok
}

// CT_PageSetup attributes this tier reads or can ignore without moving the
// page. The geometry group is read; the rest select printer behaviour, ink or
// page numbering. cellComments and errors are accepted only at the value that
// leaves the printed content alone: "atEnd" adds pages and a non-default error
// display rewrites cell text, and neither is a default. Anything else, custom
// paper dimensions included, is unsupported rather than ignorable.
func nativePageSetupAttrSupported(name, value string) bool {
	switch name {
	case "paperSize", "scale", "fitToWidth", "fitToHeight", "firstPageNumber", "copies":
		return previewPageUnsignedAttr(value)
	case "useFirstPageNumber", "usePrinterDefaults", "blackAndWhite", "draft":
		return previewPageBooleanAttr(value)
	case "orientation":
		return value == "default" || value == "portrait" || value == "landscape"
	case "pageOrder":
		return value == "downThenOver" || value == "overThenDown"
	case "cellComments":
		return value == "none"
	case "errors":
		return value == "displayed"
	case "horizontalDpi", "verticalDpi":
		return excelSentinelPrintDPI(value) || previewPageUnsignedAttr(value)
	}
	return false
}

func pageSetupSupportedAttrs(setup *previewXML, required []string) bool {
	seen := map[string]int{}
	for _, a := range setup.attrs {
		if isPreviewNamespaceDeclaration(a) {
			continue
		}
		if a.Name.Space != "" {
			if a.Name.Local == "id" && (a.Name.Space == "http://schemas.openxmlformats.org/officeDocument/2006/relationships" || a.Name.Space == "http://purl.oclc.org/ooxml/officeDocument/relationships") {
				continue
			}
			return false
		}
		if !nativePageSetupAttrSupported(a.Name.Local, a.Value) {
			return false
		}
		seen[a.Name.Local]++
		if seen[a.Name.Local] != 1 {
			return false
		}
	}
	for _, name := range required {
		if seen[name] != 1 {
			return false
		}
	}
	return true
}
