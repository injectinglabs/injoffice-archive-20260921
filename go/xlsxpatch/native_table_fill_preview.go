package xlsxpatch

import (
	"fmt"
	"math"
	"strconv"
	"strings"
)

func nativeTableStyleIDsWithinBudget(tables []NativeTablePreviewV1) bool {
	total := 0
	regions := 0
	for _, table := range tables {
		if table.BorderPreview != nil {
			total += len(table.BorderPreview.StyleIDs)
			if total > 16384 {
				return false
			}
		}
		regions += len(table.NumberFormats)
		if regions > 1024 {
			return false
		}
		for _, format := range table.NumberFormats {
			total += len(format.StyleIDs)
			if total > 16384 {
				return false
			}
		}
		if p := table.FillPreview; p != nil {
			total += len(p.FillStyleIDs) + len(p.HeaderFontStyleIDs)
			if total > 16384 {
				return false
			}
		}
	}
	return true
}

// Only a measured Medium2 fill subset is qualified. Other table formatting is
// deliberately not projected into the mutation model or inferred from labels.
func qualifyNativeTableFillPreview(pkg *nativeWorkbookPackage, tableXML *previewXML, table *NativeTablePreviewV1) {
	if table.Style != "TableStyleMedium2" || table.SheetPart == "" || table.ColumnStripes {
		return
	}
	options := tableXML.child("tableStyleInfo")
	if options == nil || len(options.children) > 0 || strings.TrimSpace(options.text) != "" {
		return
	}
	for _, a := range options.attrs {
		if a.Name.Space == "http://www.w3.org/2000/xmlns/" || (a.Name.Space == "" && a.Name.Local == "xmlns") {
			continue
		}
		if a.Name.Space != "" {
			return
		}
		switch a.Name.Local {
		case "name":
		case "showFirstColumn", "showLastColumn", "showRowStripes", "showColumnStripes":
			if a.Value != "0" && a.Value != "1" && a.Value != "true" && a.Value != "false" {
				return
			}
		default:
			return
		}
	}
	if options := tableXML.child("tableStyleInfo"); options != nil {
		for _, name := range []string{"showFirstColumn", "showLastColumn"} {
			if value := options.attr(name); value != "" && value != "0" && value != "false" {
				return
			}
		}
	}
	// Header/data differential formatting can override the builtin. Totals are
	// never painted by this subset, including their differential number formats.
	var hasOverrides func(*previewXML) bool
	hasOverrides = func(n *previewXML) bool {
		for _, a := range n.attrs {
			if a.Name.Space == "" && (a.Name.Local == "headerRowDxfId" || a.Name.Local == "dataDxfId" || a.Name.Local == "headerRowBorderDxfId" || a.Name.Local == "tableBorderDxfId" || a.Name.Local == "headerRowCellStyle" || a.Name.Local == "dataCellStyle") {
				return true
			}
		}
		for _, child := range n.children {
			if hasOverrides(child) {
				return true
			}
		}
		return false
	}
	if hasOverrides(tableXML) {
		return
	}
	worksheet, err := parsePreviewXML(pkg.files[table.SheetPart])
	if err != nil {
		return
	}
	// Conditional rules are not evaluated: even disjoint rules conservatively
	// suppress this optional decoration rather than claim precedence correctness.
	for _, child := range worksheet.children {
		if child.name.Local == "conditionalFormatting" {
			return
		}
	}
	location, err := locateWorkbookPartBytes(pkg.index, func(name string) ([]byte, bool) { v, ok := pkg.files[name]; return v, ok })
	if err != nil {
		return
	}
	ns, relNS, themeRel, opposingTheme, stylesRel, opposingStyles := spreadsheetMLTransitional, officeRelNamespaceTransitional, relTypeThemeTransitional, relTypeThemeStrict, relTypeStylesTransitional, relTypeStylesStrict
	if location.strict {
		ns, relNS, themeRel, opposingTheme, stylesRel, opposingStyles = spreadsheetMLStrict, officeRelNamespaceStrict, relTypeThemeStrict, relTypeThemeTransitional, relTypeStylesStrict, relTypeStylesTransitional
	}
	extractor := &nativeWorkbookExtractor{pkg: pkg, workbook: location, namespace: ns, relNamespace: relNS, modeled: map[string]bool{}, unsupportedKeys: map[string]bool{}, claimedXML: map[string]bool{}}
	stylesPart, err := extractor.relatedCorePart(stylesRel, opposingStyles, stylesPartContentType, "styles", false)
	if err != nil {
		return
	}
	headerFonts := []int{}
	defaultFills := []int{}
	var borderRegistry *styleRegistry
	var borderStyles *previewXML
	if stylesPart == "" {
		return
	}
	if stylesPart != "" {
		registry, e := newStyleRegistry(pkg.files[stylesPart])
		if e != nil || len(registry.fills) == 0 || !registry.fills[0].supported || registry.fills[0].color != nil {
			return
		}
		borderRegistry = registry
		for id, xf := range registry.cellXfs {
			base := effectiveCellStyleXF(registry.styleXfs[xf.xfID])
			if len(headerFonts) < 4096 && effectiveStyleComponent(xf.fontID, base.fontID, xf.applyFont) == 0 && (xf.applyFont == nil || !*xf.applyFont) {
				headerFonts = append(headerFonts, id)
			}
			if len(defaultFills) < 4096 && effectiveStyleComponent(xf.fillID, base.fillID, xf.applyFill) == 0 && (xf.applyFill == nil || !*xf.applyFill) {
				defaultFills = append(defaultFills, id)
			}
		}
		styles, e := parsePreviewXML(pkg.files[stylesPart])
		if e != nil {
			return
		}
		borderStyles = styles
		if custom := styles.child("tableStyles"); custom != nil {
			for _, child := range custom.children {
				if child.attr("name") == table.Style {
					return
				}
			}
		}
	}
	theme, err := parseWorkbookThemeDisplay(extractor, themeRel, opposingTheme)
	if err != nil {
		return
	}
	themePart, err := extractor.relatedCorePart(themeRel, opposingTheme, themePartContentType, "theme", false)
	if err != nil || themePart == "" {
		return
	}
	themeXML, err := parsePreviewXML(pkg.files[themePart])
	if err != nil {
		return
	}
	accentNode := themeXML.child("themeElements").child("clrScheme").child("accent1").child("srgbClr")
	if accentNode == nil || len(accentNode.children) != 0 {
		return
	}
	accent, ok := theme.colors[4]
	if !ok || len(accent) != 7 {
		return
	}
	table.FillPreview = &NativeTableFillPreviewV1{Header: accent, Stripe: tableLightenHLS(accent, 0.8), Body: "#FFFFFF", HeaderFontStyleIDs: headerFonts, FillStyleIDs: defaultFills}
	table.Warnings = []string{"Medium2 header and alternating body fills are previewed from the source theme. Default-font header cells use white bold text; explicit cell formatting retains precedence.", "Table text metrics and unqualified style components are not reproduced. Additional format and border notices describe supported subsets."}
	qualifyNativeTableBorders(tableXML, table, borderRegistry, borderStyles, accent)
	// The border qualification already excludes named cell styles and every
	// differential component except number formatting, including totals fonts.
	// Reuse the default-font authority list; explicit cell font choices still win.
	if table.TotalRows == 1 && table.BorderPreview != nil {
		table.FillPreview.TotalsBold = true
		table.Warnings = append(table.Warnings, "Medium2 totals use bold text only for source-qualified default-font cells; explicit fonts remain unchanged.")
	}
}

// Spreadsheet color tint is a luminance adjustment in HLS (ISO29500 ColorType),
// not independent RGB-channel blending. Keep this distinct from DrawingML.
func tableLightenHLS(rgb string, tint float64) string {
	value, err := strconv.ParseUint(rgb[1:], 16, 24)
	if err != nil {
		return rgb
	}
	r, g, b := float64(value>>16)/255, float64((value>>8)&255)/255, float64(value&255)/255
	hi, lo := math.Max(r, math.Max(g, b)), math.Min(r, math.Min(g, b))
	lum := (hi + lo) / 2
	delta := hi - lo
	sat, hue := 0.0, 0.0
	if delta != 0 {
		sat = delta / (1 - math.Abs(2*lum-1))
		switch hi {
		case r:
			hue = math.Mod((g-b)/delta, 6)
		case g:
			hue = (b-r)/delta + 2
		default:
			hue = (r-g)/delta + 4
		}
		hue /= 6
		if hue < 0 {
			hue++
		}
	}
	// Excel's observed integer HLS palette uses 240 units. Quantize each HLS
	// component and the tinted luminance before converting back to RGB.
	hue = math.Round(hue*240) / 240
	sat = math.Round(sat*240) / 240
	lum = math.Floor(math.Round(lum*240)*(1-tint)+240*tint) / 240
	l, s, h := int(math.Round(lum*240)), int(math.Round(sat*240)), int(math.Round(hue*240))
	if delta != 0 {
		dr, dg, db := int(math.Round((hi-r)*40/delta)), int(math.Round((hi-g)*40/delta)), int(math.Round((hi-b)*40/delta))
		switch hi {
		case r:
			h = db - dg
		case g:
			h = 80 + dr - db
		default:
			h = 160 + dg - dr
		}
		h = (h + 240) % 240
	}
	upper := l + s - (l*s+120)/240
	if l <= 120 {
		upper = (l*(240+s) + 120) / 240
	}
	lower := 2*l - upper
	channel := func(offset int) int {
		angle := (h + offset + 240) % 240
		var level int
		switch {
		case angle < 40:
			level = lower + ((upper-lower)*angle+20)/40
		case angle < 120:
			level = upper
		case angle < 160:
			level = lower + ((upper-lower)*(160-angle)+20)/40
		default:
			level = lower
		}
		return (level*255 + 120) / 240
	}
	return fmt.Sprintf("#%02X%02X%02X", channel(80), channel(0), channel(-80))
}
