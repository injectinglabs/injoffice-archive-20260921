package xlsxpatch

import (
	"encoding/xml"
	"strings"
)

// A separate closed qualification gate for the supplemental preview. The general
// display extractor is not evidence that an arbitrary theme has one font owner.
type nativeRichTheme struct {
	part, hash string
	fonts      map[string]string
}

func nativeRichThemeFonts(extractor *nativeWorkbookExtractor) nativeRichTheme {
	empty := nativeRichTheme{}
	// Strict DrawingML resolution is not yet joined by the display extractor.
	if extractor.workbook.strict {
		return empty
	}
	part, err := extractor.relatedCorePart(relTypeThemeTransitional, relTypeThemeStrict, themePartContentType, "theme", false)
	if err != nil || part == "" {
		return empty
	}
	root, err := parsePreviewXML(extractor.pkg.files[part])
	if err != nil || !nativeRichNode(root, drawingMLNamespace, "theme", xml.Name{Local: "name"}) || strings.TrimSpace(root.text) != "" {
		return empty
	}
	var elements, scheme *previewXML
	rootSeen := map[string]bool{}
	for _, c := range root.children {
		if c.name.Space != drawingMLNamespace || rootSeen[c.name.Local] {
			return empty
		}
		rootSeen[c.name.Local] = true
		switch c.name.Local {
		case "themeElements":
			if elements != nil || !nativeRichContainer(c, drawingMLNamespace, "themeElements") {
				return empty
			}
			elements = c
		case "objectDefaults", "extraClrSchemeLst":
			if !nativeRichContainer(c, drawingMLNamespace, c.name.Local) || len(c.children) != 0 {
				return empty
			}
		default:
			return empty
		}
	}
	if elements == nil {
		return empty
	}
	seen := map[string]bool{}
	for _, c := range elements.children {
		if c.name.Space != drawingMLNamespace || seen[c.name.Local] {
			return empty
		}
		seen[c.name.Local] = true
		switch c.name.Local {
		case "fontScheme":
			scheme = c
		case "clrScheme", "fmtScheme": // Unrelated colors and effects do not supply run fonts.
		default:
			return empty
		}
	}
	if scheme == nil || !nativeRichNode(scheme, drawingMLNamespace, "fontScheme", xml.Name{Local: "name"}) || strings.TrimSpace(scheme.text) != "" || len(scheme.children) != 2 {
		return empty
	}
	fonts := map[string]string{}
	for _, collection := range scheme.children {
		key := ""
		switch collection.name.Local {
		case "majorFont":
			key = "major"
		case "minorFont":
			key = "minor"
		default:
			return empty
		}
		if _, ok := fonts[key]; ok {
			return empty
		}
		if !nativeRichContainer(collection, drawingMLNamespace, collection.name.Local) {
			return empty
		}
		latin := ""
		children := map[string]bool{}
		scripts := map[string]bool{}
		for _, c := range collection.children {
			if c.name.Space != drawingMLNamespace || len(c.children) != 0 || c.text != "" {
				return empty
			}
			if c.name.Local == "font" {
				if !nativeRichNode(c, drawingMLNamespace, "font", xml.Name{Local: "script"}, xml.Name{Local: "typeface"}) || c.attr("script") == "" || c.attr("typeface") == "" || scripts[c.attr("script")] {
					return empty
				}
				scripts[c.attr("script")] = true
				continue // This preview's ASCII policy selects only the Latin face.
			}
			if children[c.name.Local] {
				return empty
			}
			children[c.name.Local] = true
			if !nativeRichNode(c, drawingMLNamespace, c.name.Local, xml.Name{Local: "typeface"}) {
				return empty
			}
			switch c.name.Local {
			case "latin":
				latin = c.attr("typeface")
				if !nativeRichFont.MatchString(latin) {
					return empty
				}
			case "ea", "cs":
				if c.attr("typeface") != "" {
					return empty
				}
			default:
				return empty
			}
		}
		if latin == "" {
			return empty
		}
		fonts[key] = latin
	}
	return nativeRichTheme{part: part, hash: nativeWorkbookDigest(extractor.pkg.files[part]), fonts: fonts}
}
