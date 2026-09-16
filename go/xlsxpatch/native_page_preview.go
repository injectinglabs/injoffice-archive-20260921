package xlsxpatch

import (
	"strconv"
	"strings"
)

// Explicit source settings only. This is not a printer-default resolver.
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
	Warnings  []string                  `json:"warnings"`
}

// Authored page margins for a worksheet that declares no pageSetup.
type NativeSheetPageMarginsV1 struct {
	Left   float64 `json:"left_inches"`
	Right  float64 `json:"right_inches"`
	Top    float64 `json:"top_inches"`
	Bottom float64 `json:"bottom_inches"`
}
type NativeSheetPageConfigV1 struct {
	Paper       string  `json:"paper"`
	Orientation string  `json:"orientation"`
	Scale       int     `json:"scale"`
	Left        float64 `json:"left_inches"`
	Right       float64 `json:"right_inches"`
	Top         float64 `json:"top_inches"`
	Bottom      float64 `json:"bottom_inches"`
	PageOrder   string  `json:"page_order,omitempty"`
	// Scale remains a compatibility value; fit mode ignores it.
	FitToPage *NativeSheetFitToPageV1 `json:"fit_to_page,omitempty"`
}

type NativeSheetFitToPageV1 struct {
	Width  int `json:"width"`
	Height int `json:"height"`
}

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
	sheetPrCount, pageSetupCount := 0, 0
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
		case "rowBreaks", "colBreaks", "headerFooter", "printOptions", "pageSetUpPr":
			return result
		case "sheetPr":
			sheetPrCount++
			if sheetPrCount > 1 || child.name.Space != root.name.Space {
				return result
			}
			for _, p := range child.children {
				if p.name.Local == "pageSetUpPr" {
					if fitPr != nil || p.name.Space != root.name.Space || len(child.children) != 1 || strings.TrimSpace(child.text) != "" {
						return result
					}
					for _, a := range child.attrs {
						if !isPreviewNamespaceDeclaration(a) {
							return result
						}
					}
					fitPr = p
				}
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
	exactLeaf := func(n *previewXML, names ...string) bool {
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
	readMargins := func() (map[string]float64, bool) {
		if !exactLeaf(margins, "left", "right", "top", "bottom", "header", "footer") {
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
		if fitPr != nil || pageSetupCount != 0 {
			return result
		}
		values, ok := readMargins()
		if !ok {
			return result
		}
		result.Status = "margins-only"
		result.Margins = &NativeSheetPageMarginsV1{Left: values["left"], Right: values["right"], Top: values["top"], Bottom: values["bottom"]}
		result.Warnings = []string{"Worksheet declares authored page margins and no pageSetup element, so no paper size, orientation or scale is authored. Page geometry requires a host paper choice; this tier does not select one."}
		return result
	}
	setupFields := []string{"paperSize", "orientation"}
	var fit *NativeSheetFitToPageV1
	if fitPr != nil {
		if !exactLeaf(fitPr, "fitToPage") || (fitPr.attr("fitToPage") != "true" && fitPr.attr("fitToPage") != "1") {
			return result
		}
		width, ew := strconv.Atoi(setup.attr("fitToWidth"))
		height, eh := strconv.Atoi(setup.attr("fitToHeight"))
		if ew != nil || eh != nil || width < 0 || width > 100 || height < 0 || height > 100 || width+height == 0 || strconv.Itoa(width) != setup.attr("fitToWidth") || strconv.Itoa(height) != setup.attr("fitToHeight") {
			return result
		}
		fit = &NativeSheetFitToPageV1{Width: width, Height: height}
		setupFields = append(setupFields, "fitToWidth", "fitToHeight")
	}
	if setup.attr("scale") != "" {
		setupFields = append(setupFields, "scale")
	}
	order := setup.attr("pageOrder")
	if order != "" {
		if order != "downThenOver" && order != "overThenDown" {
			return result
		}
		setupFields = append(setupFields, "pageOrder")
	}
	if !pageSetupLayoutAttrs(setup, setupFields) {
		return result
	}
	paper := map[string]string{"1": "Letter", "9": "A4"}[setup.attr("paperSize")]
	orientation := setup.attr("orientation")
	scaleText := setup.attr("scale")
	if scaleText == "" {
		scaleText = "100"
	}
	scale, err := strconv.Atoi(scaleText)
	if paper == "" || (orientation != "portrait" && orientation != "landscape") || err != nil || scale < 10 || scale > 400 || strconv.Itoa(scale) != scaleText {
		return result
	}
	values, ok := readMargins()
	if !ok {
		return result
	}
	result.Status = "available"
	result.Settings = &NativeSheetPageConfigV1{Paper: paper, Orientation: orientation, Scale: scale, Left: values["left"], Right: values["right"], Top: values["top"], Bottom: values["bottom"], PageOrder: order, FitToPage: fit}
	result.Warnings = []string{"Read-only selected-range page geometry approximation. Page settings do not select a range; repeated titles, chart paint, headers and printer behavior are not reproduced by page settings alone. No Excel fidelity claim."}
	if fit != nil {
		result.Warnings = append(result.Warnings, "Explicit fit-to-page dimensions apply to the selected preview range. A zero dimension is unconstrained. Percentage scale is ignored in fit mode; absent source scale is represented as 100 for compatibility.")
	}
	if setup.attr("horizontalDpi") != "" || setup.attr("verticalDpi") != "" {
		result.Warnings = append(result.Warnings, "Excel sentinel printer DPI values are ignored; 96 CSS px/in is the preview raster, not a source printer.")
	}
	return result
}

func excelSentinelPrintDPI(raw string) bool {
	switch raw {
	case "0", "4294967292", "4294967293", "4294967294", "4294967295":
		return true
	default:
		return false
	}
}

func pageSetupLayoutAttrs(setup *previewXML, required []string) bool {
	allowed := map[string]bool{"horizontalDpi": true, "verticalDpi": true}
	for _, name := range required {
		allowed[name] = true
	}
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
		if !allowed[a.Name.Local] {
			return false
		}
		seen[a.Name.Local]++
		if seen[a.Name.Local] != 1 {
			return false
		}
		if (a.Name.Local == "horizontalDpi" || a.Name.Local == "verticalDpi") && !excelSentinelPrintDPI(a.Value) {
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
