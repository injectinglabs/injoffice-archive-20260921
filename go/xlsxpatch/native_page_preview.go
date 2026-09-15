package xlsxpatch

import (
	"strconv"
	"strings"
)

// Explicit source settings only. This is not a printer-default resolver.
type NativeSheetPageSettingsV1 struct {
	SheetID   string                   `json:"sheet_id"`
	SheetPart string                   `json:"sheet_part"`
	Status    string                   `json:"status"`
	Settings  *NativeSheetPageConfigV1 `json:"settings,omitempty"`
	Warnings  []string                 `json:"warnings"`
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
	if setup == nil || margins == nil {
		return result
	}
	var fitPr *previewXML
	sheetPrCount := 0
	// Do not silently ignore alternate/foreign settings or page-affecting data.
	for _, child := range root.children {
		switch child.name.Local {
		case "pageSetup", "pageMargins":
			if child.name.Space != root.name.Space {
				return result
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
	if !pageSetupLayoutAttrs(setup, setupFields) || !exactLeaf(margins, "left", "right", "top", "bottom", "header", "footer") {
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
	values := map[string]float64{}
	for _, name := range []string{"left", "right", "top", "bottom", "header", "footer"} {
		value, ok := boundedPreviewRowNumber(margins.attr(name), 20)
		if !ok {
			return result
		}
		values[name] = value
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
