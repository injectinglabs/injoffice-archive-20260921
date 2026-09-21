package docxpatch

import (
	"fmt"
	"strconv"
	"strings"
)

// Page setup is a whole-section operation. The full seven-field geometry makes
// orientation/size changes atomic; header/footer distances and gutter are retained.
type NativeDOCXSectionPagePatchV1 struct {
	WidthTwips        int
	HeightTwips       int
	Orientation       string
	MarginTopTwips    int
	MarginRightTwips  int
	MarginBottomTwips int
	MarginLeftTwips   int
}

var nativeSectionOperations = map[string]bool{"section.page.patch": true}
var nativeSectionPropertyOrder = []string{"headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"}

func decodeNativeSectionPage(raw []byte) (*NativeDOCXSectionPagePatchV1, error) {
	members, err := nativeFlatJSONObject(raw)
	if err != nil {
		return nil, err
	}
	if len(members) != 7 {
		return nil, fmt.Errorf("page requires all seven geometry fields")
	}
	patch := &NativeDOCXSectionPagePatchV1{}
	fields := map[string]*int{"width_twips": &patch.WidthTwips, "height_twips": &patch.HeightTwips, "margin_top_twips": &patch.MarginTopTwips, "margin_right_twips": &patch.MarginRightTwips, "margin_bottom_twips": &patch.MarginBottomTwips, "margin_left_twips": &patch.MarginLeftTwips}
	for key, target := range fields {
		value, ok := nativeJSONInt(members[key])
		if !ok {
			return nil, fmt.Errorf("page %s must be a whole number", key)
		}
		*target = value
	}
	patch.Orientation, err = decodeNativeMutationJSONString(members["orientation"])
	if err != nil {
		return nil, err
	}
	if patch.Orientation != "portrait" && patch.Orientation != "landscape" {
		return nil, fmt.Errorf("orientation must be portrait or landscape")
	}
	for key, value := range fields {
		minimum := 0
		if key == "width_twips" || key == "height_twips" {
			minimum = 1
		}
		if *value < minimum || *value > 31680 {
			return nil, fmt.Errorf("page %s must be %d..31680", key, minimum)
		}
	}
	if patch.MarginLeftTwips+patch.MarginRightTwips >= patch.WidthTwips || patch.MarginTopTwips+patch.MarginBottomTwips >= patch.HeightTwips {
		return nil, fmt.Errorf("page margins must leave positive body dimensions")
	}
	return patch, nil
}

func nativeSetSectionPolicies(doc *NativeDocumentV1, body *nativeXMLNode, part []byte, signed bool) {
	for i := range doc.Sections {
		section := &doc.Sections[i]
		policy := nativeReadOnlyPolicy("UNSUPPORTED_SECTION_STRUCTURE", "Page setup requires a single explicit section with one column and known geometry")
		section.EditPolicy = &policy
		node := nativeNodeByPath(body, section.Anchor.Path)
		if len(doc.Sections) != 1 || node == nil || node.Name.Local != "sectPr" || node.parent != body || !strings.Contains(nativeFormatQName(part, node), ":") || nativeIntValue(section.Page.Columns) != 1 || section.Page.ColumnLayout != "equal-width" {
			continue
		}
		if signed {
			policy = nativeReadOnlyPolicy("SIGNED_PACKAGE", "Page setup cannot invalidate a package signature")
			continue
		}
		supported := true
		for _, issue := range doc.Unsupported {
			if issue.ScopeID == section.ID && issue.Code != "MISSING_PAGE_SIZE" && issue.Code != "MISSING_PAGE_MARGINS" {
				policy = nativeReadOnlyPolicy(issue.Code, issue.Message)
				supported = false
				break
			}
		}
		if supported {
			policy = NativeEditPolicyV1{Mode: "read-write", AllowedOperations: []string{"section.page.patch"}}
		}
	}
}

func applyNativeSectionPage(source []byte, revision string, mutation nativeDOCXMutationV1) (*NativeDOCXMutationResultV1, error) {
	refuse := func(code, message string) (*NativeDOCXMutationResultV1, error) {
		return nil, nativeMutationError(code, mutation.TargetID, message)
	}
	if revision != nativeSHA(source) {
		return refuse("STALE_REVISION", "page setup requires the current exact package revision")
	}
	if mutation.TargetKind != "section" || mutation.Page == nil {
		return refuse("INVALID_SELECTOR", "page setup requires a section target and page geometry")
	}
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		return nil, err
	}
	var section *NativeSectionV1
	for i := range doc.Sections {
		if doc.Sections[i].ID == mutation.TargetID {
			section = &doc.Sections[i]
			break
		}
	}
	if section == nil {
		return refuse("TARGET_NOT_FOUND", "section is absent from the extracted document")
	}
	if section.Anchor.XMLSHA256 != mutation.ExpectedXMLSHA256 {
		return refuse("STALE_TARGET", "section XML fingerprint changed")
	}
	if section.EditPolicy == nil || !nativePolicyAllows(*section.EditPolicy, "section.page.patch") {
		return refuse("UNSUPPORTED_SECTION_STRUCTURE", "the extracted section does not allow page setup")
	}
	patch := mutation.Page
	if int64(patch.MarginLeftTwips+patch.MarginRightTwips)+nativeInt64Value(section.Page.Margins.GutterTwips) >= int64(patch.WidthTwips) {
		return refuse("INVALID_PAGE_GEOMETRY", "margins and the preserved gutter leave no body width")
	}
	pkg, err := openNativeDOCXPackage(source)
	if err != nil {
		return nil, err
	}
	if nativeDOCXPackageHasDigitalSignature(pkg) {
		return refuse("SIGNED_PACKAGE", "page setup cannot invalidate a package signature")
	}
	part := pkg.files[section.Anchor.PartName]
	root, err := parseNativeXML(section.Anchor.PartName, part)
	if err != nil {
		return nil, err
	}
	node := nativeNodeByPath(root, section.Anchor.Path)
	if node == nil || section.Anchor.StartByte == nil || section.Anchor.EndByte == nil || node.Start != *section.Anchor.StartByte || node.End != *section.Anchor.EndByte || nativeSHA(part[node.Start:node.End]) != mutation.ExpectedXMLSHA256 {
		return refuse("STALE_TARGET", "section source bytes do not match the anchor")
	}
	number := func(n int) *string { value := strconv.Itoa(n); return &value }
	written := map[string]map[string]*string{
		"pgSz":  {"w": number(patch.WidthTwips), "h": number(patch.HeightTwips), "orient": &patch.Orientation},
		"pgMar": {"top": number(patch.MarginTopTwips), "right": number(patch.MarginRightTwips), "bottom": number(patch.MarginBottomTwips), "left": number(patch.MarginLeftTwips)},
	}
	if firstDirectNativeChild(node, node.Name.Space, "pgMar") == nil {
		written["pgMar"]["header"] = number(int(nativeInt64Value(section.Page.Margins.HeaderTwips)))
		written["pgMar"]["footer"] = number(int(nativeInt64Value(section.Page.Margins.FooterTwips)))
		written["pgMar"]["gutter"] = number(int(nativeInt64Value(section.Page.Margins.GutterTwips)))
	}
	qname := nativeFormatQName(part, node)
	prefix := qname[:strings.Index(qname, ":")+1]
	replacement, err := nativeMergeLayoutChildren(part, node, nativeSectionPropertyOrder, written, prefix)
	if err != nil {
		return refuse("UNSUPPORTED_LEXICAL_FORM", err.Error())
	}
	return writeNativeMutationSplices(source, pkg, doc, revision, map[string][]nativeTextSplice{section.Anchor.PartName: {{start: node.Start, end: node.End, text: replacement}}}, func(after *NativeDocumentV1) error {
		if len(after.Sections) != 1 {
			return nativeMutationError("POST_WRITE_MISMATCH", section.ID, "section count changed")
		}
		page := after.Sections[0].Page
		got := NativeDOCXSectionPagePatchV1{int(nativeInt64Value(page.WidthTwips)), int(nativeInt64Value(page.HeightTwips)), page.Orientation, int(nativeInt64Value(page.Margins.TopTwips)), int(nativeInt64Value(page.Margins.RightTwips)), int(nativeInt64Value(page.Margins.BottomTwips)), int(nativeInt64Value(page.Margins.LeftTwips))}
		if got != *patch {
			return nativeMutationError("POST_WRITE_MISMATCH", section.ID, "page setup did not round-trip")
		}
		return nil
	})
}
