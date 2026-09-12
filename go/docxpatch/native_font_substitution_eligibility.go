package docxpatch

import (
	"encoding/xml"
	"fmt"
	"strings"
)

// These source-derived facts qualify only explicit read-only operator font
// selection. Extraction, authored font identity and mutation rules are unchanged.
type NativeDOCXFontSubstitutionEligibilityV1 struct {
	Protocol      string                              `json:"protocol"`
	Version       int                                 `json:"version"`
	DocumentID    string                              `json:"document_id"`
	Revision      string                              `json:"revision"`
	PackageSHA256 string                              `json:"package_sha256"`
	FontTable     *NativeDOCXFontTableBindingV1       `json:"font_table"`
	Facts         []NativeDOCXFontDescriptorPreviewV1 `json:"facts"`
}
type NativeDOCXFontDescriptorPreviewV1 struct {
	Family     string                       `json:"family"`
	Use        string                       `json:"use"`
	Kind       string                       `json:"kind"`
	Path       string                       `json:"path"`
	Values     map[string]string            `json:"values"`
	Diagnostic NativeResolutionDiagnosticV1 `json:"diagnostic"`
}

func ExtractNativeDOCXFontSubstitutionEligibilityV1(data []byte) (*NativeDOCXFontSubstitutionEligibilityV1, error) {
	resolver, err := newNativeLayoutResolver(data, NativeExtractionOptions{})
	if err != nil {
		return nil, err
	}
	layout, err := resolver.resolve()
	if err != nil {
		return nil, err
	}
	inventory, err := ExtractNativeDOCXFontInventoryV1(data)
	if err != nil {
		return nil, err
	}
	result := &NativeDOCXFontSubstitutionEligibilityV1{Protocol: "injoffice.docx.font-substitution-eligibility", Version: 1, DocumentID: layout.DocumentID, Revision: layout.Revision, PackageSHA256: inventory.PackageSHA256, FontTable: inventory.FontTable, Facts: []NativeDOCXFontDescriptorPreviewV1{}}
	fail := func() (*NativeDOCXFontSubstitutionEligibilityV1, error) {
		return nil, fmt.Errorf("docxpatch: font descriptors are not source-qualified for operator preview")
	}
	if inventory.FontTable == nil {
		return result, nil
	}
	part := inventory.FontTable.PartName
	root, err := parseNativeXML(part, resolver.pkg.files[part])
	if err != nil {
		return nil, err
	}
	if root.Name != (xml.Name{Space: resolver.wordNS, Local: "fonts"}) || !nativeQualifiedFontTableOwner(root) {
		return fail()
	}
	used := map[string]bool{}
	// Descriptor ordinals follow loadFonts' original declaration order. Names
	// and aliases actually consumed must identify one owner. An unused alias
	// overlap does not participate in explicit family selection.
	owners := map[string]int{}
	for _, font := range layout.Fonts {
		names := []string{font.Name}
		if font.AltName != nil {
			names = append(names, *font.AltName)
		}
		for _, name := range names {
			key := strings.ToLower(name)
			owners[key]++
		}
	}
	for _, run := range layout.Runs {
		if run.Properties.FontFamily == nil {
			return fail()
		}
	}
	for _, paragraph := range layout.Paragraphs {
		if paragraph.ParagraphMarkProperties.FontFamily == nil || paragraph.Numbering != nil && paragraph.Numbering.Marker.FontFamily == nil {
			return fail()
		}
	}
	for _, r := range inventory.References {
		used[strings.ToLower(r.Family)] = true
	}
	for family := range used {
		if owners[family] > 1 {
			return fail()
		}
	}
	diagnostics := map[string]NativeResolutionDiagnosticV1{}
	for _, d := range layout.Diagnostics {
		if d.PartName != nil && *d.PartName == part && d.Code == "FONT_MATCHING_METADATA_PRESERVED" {
			if d.Path == nil || diagnostics[*d.Path].Code != "" {
				return fail()
			}
			diagnostics[*d.Path] = d
		}
	}
	attested := resolver.attestedEmbeddedFontPaths(part)
	for _, font := range root.Children {
		if font.Name != (xml.Name{Space: resolver.wordNS, Local: "font"}) || !nativeExactContainer(font, xml.Name{Space: resolver.wordNS, Local: "name"}) {
			return fail()
		}
		name, present := nativeAttr(font, resolver.wordNS, "name")
		if !present || name == "" {
			return fail()
		}
		active := used[strings.ToLower(name)]
		aliases := directNativeChildren(font, resolver.wordNS, "altName")
		if len(aliases) > 1 {
			return fail()
		}
		if len(aliases) == 1 {
			alias := aliases[0]
			value, ok := nativeAttr(alias, resolver.wordNS, "val")
			if !ok || value == "" || !nativeExactLeaf(alias, xml.Name{Space: resolver.wordNS, Local: "val"}) {
				return fail()
			}
			active = active || used[strings.ToLower(value)]
		}
		classification := "unused"
		if active && len(font.Children) > 0 {
			classification = "latin-matching"
			charset := directNativeChildren(font, resolver.wordNS, "charset")
			if len(charset) != 1 || !nativeQualifiedFontDescriptor(charset[0], font, resolver.wordNS) {
				return fail()
			}
			for _, panose := range directNativeChildren(font, resolver.wordNS, "panose1") {
				v, _ := nativeAttr(panose, resolver.wordNS, "val")
				if len(v) != 20 || !strings.HasPrefix(v, "02") {
					return fail()
				}
			}
		}
		for _, node := range font.Children {
			if node.Name.Space != resolver.wordNS {
				return fail()
			}
			if node.Name.Local == "altName" {
				continue
			}
			if strings.HasPrefix(node.Name.Local, "embed") && attested[node.Path] {
				continue
			}
			qualified := nativeQualifiedFontDescriptor(node, font, resolver.wordNS)
			if !qualified && (active || !nativePotentiallyUnusedFontDescriptor(node, font, resolver.wordNS)) {
				return fail()
			}
			diagnostic, ok := diagnostics[node.Path]
			if !ok {
				return fail()
			}
			delete(diagnostics, node.Path)
			values := map[string]string{}
			for _, attr := range node.Attrs {
				if attr.Name.Space == resolver.wordNS {
					values[attr.Name.Local] = attr.Value
				}
			}
			result.Facts = append(result.Facts, NativeDOCXFontDescriptorPreviewV1{Family: name, Use: classification, Kind: node.Name.Local, Path: node.Path, Values: values, Diagnostic: diagnostic})
			if len(result.Facts) > NativeDOCXMaxResolvedDiagnostics {
				return fail()
			}
		}
	}
	if len(diagnostics) != 0 {
		return fail()
	}
	return result, nil
}
