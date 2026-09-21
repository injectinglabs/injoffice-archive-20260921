package docxpatch

import (
	"bytes"
	"encoding/json"
	"fmt"
	"sort"
	"strconv"
	"strings"
)

// Run formatting is a guarded native mutation, not a rewrite: the requested
// character properties are merged into the run properties the source already
// carries, the run is split only at the requested range boundaries, and every
// other byte of the package is preserved. The contract operation is
// "properties.patch" on a paragraph or run target; a range is expressed in
// UTF-16 code units over the target's own text, exactly as the contract
// reports it.

// NativeDOCXRunPropertyPatchV1 is the bounded direct-formatting subset this
// tier can write back. An absent field is left exactly as the source has it.
type NativeDOCXRunPropertyPatchV1 struct {
	Bold               *bool
	Italic             *bool
	Underline          *string
	FontFamily         *string
	FontSizeHalfPoints *int
	Color              *string
	Highlight          *string
}

// NativeDOCXParagraphPropertyPatchV1 is the bounded paragraph-level
// formatting this tier can write back. An absent field is left as the source
// has it.
type NativeDOCXParagraphPropertyPatchV1 struct {
	Alignment *string
}

// NativeDOCXTextRangeV1 selects part of the target's text in UTF-16 code
// units, the same unit the renderer's selection uses.
type NativeDOCXTextRangeV1 struct {
	StartUTF16 int
	EndUTF16   int
}

// NativeDOCXFormatMutationV1 patches run properties over a paragraph or run
// target. ExpectedXMLSHA256 is mandatory and fails the mutation closed when
// the anchored XML moved.
type NativeDOCXFormatMutationV1 struct {
	TargetKind          string
	TargetID            string
	ExpectedXMLSHA256   string
	Properties          NativeDOCXRunPropertyPatchV1
	ParagraphProperties *NativeDOCXParagraphPropertyPatchV1
	Range               *NativeDOCXTextRangeV1
}

var (
	// ECMA-376 17.3.2.1 CT_RPr child order. A property this patch writes is
	// inserted at its schema position; every other child keeps its own.
	nativeRunPropertyOrder = []string{"rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"}
	// The underline subset native extraction reads back. Writing a value it
	// refuses would make the paragraph read-only on the next extraction.
	nativeUnderlineValues = map[string]bool{"none": true, "single": true, "double": true, "words": true}
	// ECMA-376 17.3.1.26 CT_PPr child order, to the depth this tier writes.
	nativeParagraphPropertyOrder = []string{"pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"}
	nativeAlignmentValues        = map[string]bool{"left": true, "center": true, "right": true, "both": true, "distribute": true}
	nativeHighlightValues        = map[string]bool{"none": true, "black": true, "blue": true, "cyan": true, "darkBlue": true, "darkCyan": true, "darkGray": true, "darkGreen": true, "darkMagenta": true, "darkRed": true, "darkYellow": true, "green": true, "lightGray": true, "magenta": true, "red": true, "white": true, "yellow": true}
)

// ApplyNativeMutationPayloadV1 is the renderer-neutral Office mutation
// envelope entrypoint. It proves the envelope's exact-byte CAS, then applies
// the payload as one transaction: exact text replacements, or run-property
// patches when every mutation carries "properties". A batch may not mix the
// two shapes; both are decoded by the same strict payload boundary.
func ApplyNativeMutationPayloadV1(packageBytes, payload []byte, outerExpectedRevision string) (*NativeDOCXMutationResultV1, error) {
	decoded, err := decodeNativeDOCXMutationPayloadV1(payload)
	if err != nil {
		return nil, err
	}
	formatting := 0
	for _, mutation := range decoded {
		if mutation.Properties != nil || mutation.ParagraphProperties != nil {
			formatting++
		}
	}
	if formatting != 0 && formatting != len(decoded) {
		return nil, nativeMutationError("INVALID_PAYLOAD", "", "a transaction may not mix text replacements with property patches")
	}
	if formatting == 0 {
		return ApplyNativeTextMutationPayloadV1(packageBytes, payload, outerExpectedRevision)
	}
	exactRevision := nativeSHA(packageBytes)
	if outerExpectedRevision != exactRevision {
		return nil, nativeMutationError("STALE_REVISION", "", fmt.Sprintf("outer expected revision %q does not match exact package revision %q", outerExpectedRevision, exactRevision))
	}
	mutations, err := DecodeNativeDOCXFormatMutationPayloadV1(payload)
	if err != nil {
		return nil, err
	}
	return ApplyNativeFormatMutationsV1(packageBytes, outerExpectedRevision, mutations)
}

// decodeNativeDOCXStructuredMutationField decodes the two object-valued
// mutation fields the envelope carries beside its string fields. It walks
// JSON tokens, like every other native payload boundary, and refuses an
// unknown or repeated member outright.
type nativeDOCXStructuredFieldV1 struct {
	runs      *NativeDOCXRunPropertyPatchV1
	paragraph *NativeDOCXParagraphPropertyPatchV1
	span      *NativeDOCXTextRangeV1
}

func decodeNativeDOCXStructuredMutationField(field string, raw []byte) (nativeDOCXStructuredFieldV1, error) {
	decoded := nativeDOCXStructuredFieldV1{}
	members, err := nativeFlatJSONObject(raw)
	if err != nil {
		return decoded, err
	}
	if field == "range" {
		start, startOK := nativeJSONInt(members["start_utf16"])
		end, endOK := nativeJSONInt(members["end_utf16"])
		if len(members) != 2 || !startOK || !endOK {
			return decoded, fmt.Errorf("must carry whole start_utf16 and end_utf16 numbers")
		}
		decoded.span = &NativeDOCXTextRangeV1{StartUTF16: start, EndUTF16: end}
		return decoded, nil
	}
	patch := &NativeDOCXRunPropertyPatchV1{}
	for member, value := range members {
		switch member {
		case "bold", "italic":
			flag, ok := nativeJSONBool(value)
			if !ok {
				return decoded, fmt.Errorf("member %q must be a JSON boolean", member)
			}
			if member == "bold" {
				patch.Bold = &flag
			} else {
				patch.Italic = &flag
			}
		case "underline", "font_family", "color", "highlight", "alignment":
			text, decodeErr := decodeNativeMutationJSONString(value)
			if decodeErr != nil {
				return decoded, fmt.Errorf("member %q must be a JSON string", member)
			}
			switch member {
			case "underline":
				patch.Underline = &text
			case "font_family":
				patch.FontFamily = &text
			case "color":
				patch.Color = &text
			case "highlight":
				patch.Highlight = &text
			default:
				decoded.paragraph = &NativeDOCXParagraphPropertyPatchV1{Alignment: &text}
			}
		case "font_size_half_points":
			size, ok := nativeJSONInt(value)
			if !ok {
				return decoded, fmt.Errorf("member %q must be a whole JSON number", member)
			}
			patch.FontSizeHalfPoints = &size
		default:
			return decoded, fmt.Errorf("carries unknown member %q", member)
		}
	}
	if decoded.paragraph != nil && len(members) != 1 {
		return decoded, fmt.Errorf("must patch paragraph alignment on its own")
	}
	if decoded.paragraph == nil {
		decoded.runs = patch
	}
	return decoded, nil
}

// nativeFlatJSONObject reads one JSON object of scalar members, refusing a
// repeated member and any nesting.
func nativeFlatJSONObject(raw []byte) (map[string]json.RawMessage, error) {
	decoder := json.NewDecoder(bytes.NewReader(raw))
	token, err := decoder.Token()
	if delimiter, ok := token.(json.Delim); err != nil || !ok || delimiter != '{' {
		return nil, fmt.Errorf("must be a JSON object")
	}
	members := map[string]json.RawMessage{}
	for decoder.More() {
		name, nameErr := decoder.Token()
		member, ok := name.(string)
		if nameErr != nil || !ok {
			return nil, fmt.Errorf("contains invalid JSON")
		}
		if _, repeated := members[member]; repeated || len(members) > 16 {
			return nil, fmt.Errorf("repeats member %q", member)
		}
		var value json.RawMessage
		if valueErr := decoder.Decode(&value); valueErr != nil {
			return nil, fmt.Errorf("member %q contains invalid JSON", member)
		}
		if len(value) == 0 || value[0] == '{' || value[0] == '[' {
			return nil, fmt.Errorf("member %q must be a scalar", member)
		}
		members[member] = value
	}
	if _, closeErr := decoder.Token(); closeErr != nil {
		return nil, fmt.Errorf("is unterminated")
	}
	return members, nil
}

func nativeJSONBool(raw []byte) (bool, bool) {
	value := string(raw)
	return value == "true", value == "true" || value == "false"
}

func nativeJSONInt(raw []byte) (int, bool) {
	value, err := strconv.Atoi(string(raw))
	if err != nil {
		return 0, false
	}
	return value, true
}

func validateNativeDOCXFormatMutation(mutation *NativeDOCXFormatMutationV1) error {
	if mutation.TargetKind != "paragraph" && mutation.TargetKind != "run" {
		return fmt.Errorf("target_kind must be paragraph or run")
	}
	if !nativeIDPattern.MatchString(mutation.TargetID) {
		return fmt.Errorf("target_id must be a bounded native identifier")
	}
	if !nativeSHA256.MatchString(mutation.ExpectedXMLSHA256) {
		return fmt.Errorf("expected_xml_sha256 must contain a full lowercase SHA-256")
	}
	if mutation.ParagraphProperties != nil {
		if mutation.TargetKind != "paragraph" {
			return fmt.Errorf("alignment is a paragraph property and needs a paragraph target")
		}
		if mutation.Range != nil {
			return fmt.Errorf("alignment applies to the whole paragraph and takes no range")
		}
		if mutation.ParagraphProperties.Alignment == nil || !nativeAlignmentValues[*mutation.ParagraphProperties.Alignment] {
			return fmt.Errorf("alignment must be left, center, right, both or distribute")
		}
		return nil
	}
	patch := &mutation.Properties
	if patch.Bold == nil && patch.Italic == nil && patch.Underline == nil && patch.FontFamily == nil && patch.FontSizeHalfPoints == nil && patch.Color == nil && patch.Highlight == nil {
		return fmt.Errorf("properties must set at least one run property")
	}
	if patch.Underline != nil && !nativeUnderlineValues[*patch.Underline] {
		return fmt.Errorf("underline must be none, single, double or words")
	}
	if patch.Highlight != nil && !nativeHighlightValues[*patch.Highlight] {
		return fmt.Errorf("highlight must be one of the ECMA-376 highlight names")
	}
	if patch.FontFamily != nil && !nativeFormatFontName(*patch.FontFamily) {
		return fmt.Errorf("font_family must be 1..64 printable characters without XML markup")
	}
	if patch.FontSizeHalfPoints != nil && (*patch.FontSizeHalfPoints < 2 || *patch.FontSizeHalfPoints > 3276) {
		return fmt.Errorf("font_size_half_points must be 2..3276")
	}
	if patch.Color != nil {
		normalized, ok := nativeExactRGB(*patch.Color)
		if !ok {
			return fmt.Errorf("color must be a six-digit RGB value")
		}
		mutation.Properties.Color = &normalized
	}
	if mutation.Range != nil && (mutation.Range.StartUTF16 < 0 || mutation.Range.EndUTF16 <= mutation.Range.StartUTF16 || mutation.Range.EndUTF16 > NativeDOCXMaxTextLength) {
		return fmt.Errorf("range must be a non-empty 0..%d UTF-16 span", NativeDOCXMaxTextLength)
	}
	return nil
}

func nativeFormatFontName(value string) bool {
	if value == "" || len(value) > 64 {
		return false
	}
	for _, character := range value {
		if character < 0x20 || character == '<' || character == '>' || character == '&' || character == '"' || character == '\'' || character == 0x7F {
			return false
		}
	}
	return true
}

// nativeFormatRequest is one resolved target: a paragraph and the UTF-16 span
// of that paragraph's own text the patch applies to.
type nativeFormatRequest struct {
	paragraph *NativeParagraphV1
	start     int
	end       int
	patch     NativeDOCXRunPropertyPatchV1
	alignment *string
	text      string
}

// ApplyNativeFormatMutationsV1 atomically applies run-property patches to
// native WordprocessingML bytes. expectedRevision must be the full exact
// package digest from NativeDocumentV1.Source.PackageSHA256. Every target
// anchor is verified against the current bytes, runs are split only at the
// requested boundaries, existing w:rPr children are preserved in place, the
// OPC is reopened and re-extracted with identity continuity, and the
// requested properties are proved present on exactly the selected text before
// the package is returned. Any refusal returns no output package.
func ApplyNativeFormatMutationsV1(packageBytes []byte, expectedRevision string, mutations []NativeDOCXFormatMutationV1) (*NativeDOCXMutationResultV1, error) {
	if len(mutations) == 0 || len(mutations) > NativeDOCXMaxMutations {
		return nil, nativeMutationError("INVALID_MUTATION_COUNT", "", fmt.Sprintf("mutation count must be 1..%d", NativeDOCXMaxMutations))
	}
	if !nativeSHA256.MatchString(expectedRevision) {
		return nil, nativeMutationError("INVALID_REVISION", "", "expected revision must be sha256 followed by the full lowercase exact-byte digest")
	}
	doc, err := ExtractNativeDocumentV1(packageBytes)
	if err != nil {
		return nil, fmt.Errorf("docxpatch: native mutation source validation: %w", err)
	}
	exactRevision := nativeSHA(packageBytes)
	if doc.Source.PackageSHA256 != exactRevision || expectedRevision != exactRevision {
		return nil, nativeMutationError("STALE_REVISION", "", fmt.Sprintf("expected %q, current exact package revision is %q", expectedRevision, exactRevision))
	}
	pkg, err := openNativeDOCXPackage(packageBytes)
	if err != nil {
		return nil, fmt.Errorf("docxpatch: native mutation reopen source: %w", err)
	}
	if nativeDOCXPackageHasDigitalSignature(pkg) {
		return nil, nativeMutationError("SIGNED_PACKAGE", "", "a signed OPC package cannot be mutated without invalidating its digital signature")
	}
	paragraphs := map[string]*NativeParagraphV1{}
	runOwners := map[string]*NativeParagraphV1{}
	forEachNativeParagraph(doc, func(paragraph *NativeParagraphV1) {
		paragraphs[paragraph.ID] = paragraph
		for _, run := range paragraph.Runs {
			runOwners[run.ID] = paragraph
		}
	})
	splicesByPart := map[string][]nativeTextSplice{}
	roots := map[string]*nativeXMLNode{}
	seen := map[string]bool{}
	requests := make([]nativeFormatRequest, 0, len(mutations))
	for _, mutation := range mutations {
		request, err := resolveNativeFormatRequest(mutation, paragraphs, runOwners)
		if err != nil {
			return nil, err
		}
		if seen[request.paragraph.ID] {
			return nil, nativeMutationError("DUPLICATE_TARGET", mutation.TargetID, "a paragraph may be formatted only once in an atomic transaction")
		}
		seen[request.paragraph.ID] = true
		partName := request.paragraph.Anchor.PartName
		root, ok := roots[partName]
		if !ok {
			root, err = parseNativeXML(partName, pkg.files[partName])
			if err != nil {
				return nil, fmt.Errorf("docxpatch: native run formatting: %w", err)
			}
			roots[partName] = root
		}
		splices, err := nativeFormatParagraphSplices(pkg.files[partName], root, request, mutation.TargetID)
		if err != nil {
			return nil, err
		}
		splicesByPart[partName] = append(splicesByPart[partName], splices...)
		requests = append(requests, request)
	}
	for partName := range splicesByPart {
		sort.Slice(splicesByPart[partName], func(i, j int) bool {
			return splicesByPart[partName][i].start < splicesByPart[partName][j].start
		})
	}
	return writeNativeMutationSplices(packageBytes, pkg, doc, exactRevision, splicesByPart, func(after *NativeDocumentV1) error {
		return validateNativeFormatResults(after, requests)
	})
}

func resolveNativeFormatRequest(mutation NativeDOCXFormatMutationV1, paragraphs map[string]*NativeParagraphV1, runOwners map[string]*NativeParagraphV1) (nativeFormatRequest, error) {
	refuse := func(code, message string) (nativeFormatRequest, error) {
		return nativeFormatRequest{}, nativeMutationError(code, mutation.TargetID, message)
	}
	if err := validateNativeDOCXFormatMutation(&mutation); err != nil {
		return refuse("INVALID_SELECTOR", err.Error())
	}
	paragraph := paragraphs[mutation.TargetID]
	if mutation.TargetKind == "run" {
		paragraph = runOwners[mutation.TargetID]
	}
	if paragraph == nil {
		return refuse("TARGET_NOT_FOUND", "target is absent from the current native contract")
	}
	if paragraph.EditPolicy.Mode != "read-write" || !nativePolicyAllows(paragraph.EditPolicy, "properties.patch") {
		message := "owning paragraph does not allow run-property patches"
		if paragraph.EditPolicy.Refusal != nil {
			message += ": " + paragraph.EditPolicy.Refusal.Code
		}
		return refuse("UNSUPPORTED_CONSTRUCT", message)
	}
	if mutation.ParagraphProperties != nil {
		if mutation.ExpectedXMLSHA256 != paragraph.Anchor.XMLSHA256 {
			return refuse("STALE_TARGET", fmt.Sprintf("expected XML fingerprint %q, current fingerprint is %q", mutation.ExpectedXMLSHA256, paragraph.Anchor.XMLSHA256))
		}
		return nativeFormatRequest{paragraph: paragraph, alignment: mutation.ParagraphProperties.Alignment}, nil
	}
	anchor := paragraph.Anchor
	text := strings.Builder{}
	total, runStart, runLength := 0, -1, 0
	for _, run := range paragraph.Runs {
		length := 0
		if run.Kind == "text" && run.Text != nil {
			length = nativeMutationUTF16CodeUnits(*run.Text)
			text.WriteString(*run.Text)
		}
		if mutation.TargetKind == "run" && run.ID == mutation.TargetID {
			if run.Kind != "text" || run.Text == nil {
				return refuse("UNSUPPORTED_TARGET_KIND", "only a modeled text run carries run properties this tier can patch")
			}
			anchor, runStart, runLength = run.Anchor, total, length
		}
		total += length
	}
	start, end := 0, total
	if mutation.TargetKind == "run" {
		if runStart < 0 {
			return refuse("TARGET_NOT_FOUND", "target is absent from the current native contract")
		}
		start, end = runStart, runStart+runLength
	}
	if mutation.Range != nil {
		if mutation.Range.EndUTF16 > end-start {
			return refuse("INVALID_RANGE", fmt.Sprintf("range end %d is past the target's %d UTF-16 code units", mutation.Range.EndUTF16, end-start))
		}
		start, end = start+mutation.Range.StartUTF16, start+mutation.Range.EndUTF16
	}
	if start >= end {
		return refuse("INVALID_RANGE", "the selected target carries no text to format")
	}
	if mutation.ExpectedXMLSHA256 != anchor.XMLSHA256 {
		return refuse("STALE_TARGET", fmt.Sprintf("expected XML fingerprint %q, current fingerprint is %q", mutation.ExpectedXMLSHA256, anchor.XMLSHA256))
	}
	return nativeFormatRequest{paragraph: paragraph, start: start, end: end, patch: mutation.Properties, text: text.String()}, nil
}

// nativeFormatParagraphSplices rewrites every w:r the requested span touches.
// A run the span covers entirely keeps its element and gains merged run
// properties; a run the span covers partly is split at the boundary into
// otherwise identical runs, so the source's own start tag, run properties and
// text element survive each piece.
func nativeFormatParagraphSplices(part []byte, root *nativeXMLNode, request nativeFormatRequest, targetID string) ([]nativeTextSplice, error) {
	refuse := func(code, message string) ([]nativeTextSplice, error) {
		return nil, nativeMutationError(code, targetID, message)
	}
	paragraph := nativeNodeByPath(root, request.paragraph.Anchor.Path)
	if paragraph == nil || request.paragraph.Anchor.StartByte == nil || request.paragraph.Anchor.EndByte == nil ||
		paragraph.Start != *request.paragraph.Anchor.StartByte || paragraph.End != *request.paragraph.Anchor.EndByte ||
		nativeSHA(part[paragraph.Start:paragraph.End]) != request.paragraph.Anchor.XMLSHA256 {
		return refuse("STALE_TARGET", "paragraph source bytes no longer match the issued anchor")
	}
	if request.alignment != nil {
		splice, alignErr := nativeFormatParagraphProperties(part, paragraph, *request.alignment)
		if alignErr != nil {
			return refuse("UNSUPPORTED_LEXICAL_FORM", alignErr.Error())
		}
		return []nativeTextSplice{splice}, nil
	}
	splices := []nativeTextSplice{}
	offset := 0
	for _, run := range request.paragraph.Runs {
		if run.Kind != "text" || run.Text == nil {
			continue
		}
		start, length := offset, nativeMutationUTF16CodeUnits(*run.Text)
		offset += length
		if start >= request.end || offset <= request.start {
			continue
		}
		textNode := nativeNodeByPath(paragraph, run.Anchor.Path)
		if textNode == nil || run.Anchor.StartByte == nil || textNode.Start != *run.Anchor.StartByte {
			return refuse("STALE_TARGET", "run source bytes no longer match the issued anchor")
		}
		owner := textNode.parent
		if owner == nil || owner.parent != paragraph || owner.Name.Space != textNode.Name.Space || owner.Name.Local != "r" {
			return refuse("UNSUPPORTED_CONSTRUCT", "only a run that is a direct child of its paragraph can be formatted")
		}
		body := []*nativeXMLNode{}
		var properties *nativeXMLNode
		for _, child := range owner.Children {
			if child.Name.Space == owner.Name.Space && child.Name.Local == "rPr" && properties == nil && len(body) == 0 {
				properties = child
				continue
			}
			body = append(body, child)
		}
		if len(body) != 1 || body[0] != textNode || owner.Text != "" {
			return refuse("UNSUPPORTED_CONSTRUCT", "only a run whose single content child is its text element can be formatted")
		}
		replacement, err := nativeFormatRunElement(part, owner, properties, textNode, *run.Text, max(request.start-start, 0), min(request.end-start, length), request.patch)
		if err != nil {
			return refuse("UNSUPPORTED_LEXICAL_FORM", err.Error())
		}
		splices = append(splices, nativeTextSplice{start: owner.Start, end: owner.End, text: replacement})
	}
	if len(splices) == 0 {
		return refuse("SEMANTIC_NO_OP", "the requested range covers no formattable run")
	}
	return splices, nil
}

// nativeFormatRunElement emits the replacement bytes for one w:r: up to three
// runs built from the source's own start tag, run properties and text element.
func nativeFormatRunElement(part []byte, owner, properties, textNode *nativeXMLNode, text string, from, to int, patch NativeDOCXRunPropertyPatchV1) ([]byte, error) {
	tagEnd := nativeStartTagEnd(part[owner.Start:owner.End])
	if tagEnd <= 0 {
		return nil, fmt.Errorf("run element has an unterminated start tag")
	}
	startTag := part[owner.Start : owner.Start+int64(tagEnd)]
	endTag := []byte("</" + nativeFormatQName(part, owner) + ">")
	sourceProperties := []byte(nil)
	if properties != nil {
		sourceProperties = part[properties.Start:properties.End]
	}
	patched, err := nativeFormatRunProperties(part, owner, properties, patch)
	if err != nil {
		return nil, err
	}
	sourceText := part[textNode.Start:textNode.End]
	fromByte, fromOK := nativeUTF16ByteOffset(text, from)
	toByte, toOK := nativeUTF16ByteOffset(text, to)
	if !fromOK || !toOK {
		return nil, fmt.Errorf("a formatting range may not split a surrogate pair")
	}
	output := []byte(nil)
	segments := []struct {
		text       string
		properties []byte
	}{{text[:fromByte], sourceProperties}, {text[fromByte:toByte], patched}, {text[toByte:], sourceProperties}}
	for _, segment := range segments {
		if segment.text == "" {
			continue
		}
		element, err := nativeFormatTextElement(sourceText, segment.text)
		if err != nil {
			return nil, err
		}
		output = append(output, startTag...)
		output = append(output, segment.properties...)
		output = append(output, element...)
		output = append(output, endTag...)
	}
	if len(output) == 0 {
		return nil, fmt.Errorf("run formatting produced no run")
	}
	return output, nil
}

// nativeFormatParagraphProperties merges w:jc into the paragraph's own
// w:pPr, creating the element when the source has none. Every other child of
// w:pPr keeps its bytes and its position.
func nativeFormatParagraphProperties(part []byte, paragraph *nativeXMLNode, alignment string) (nativeTextSplice, error) {
	prefix := ""
	if name := nativeFormatQName(part, paragraph); strings.Contains(name, ":") {
		prefix = name[:strings.Index(name, ":")+1]
	}
	written := []byte(`<` + prefix + `jc ` + prefix + `val="` + nativeMutationEscapeAttribute(alignment) + `"/>`)
	properties := firstDirectNativeChild(paragraph, paragraph.Name.Space, "pPr")
	if properties != nil && paragraph.Children[0] != properties {
		return nativeTextSplice{}, fmt.Errorf("paragraph properties must be the paragraph's first child")
	}
	if properties == nil {
		tagEnd := nativeStartTagEnd(part[paragraph.Start:paragraph.End])
		if tagEnd <= 0 {
			return nativeTextSplice{}, fmt.Errorf("paragraph element has an unterminated start tag")
		}
		at := paragraph.Start + int64(tagEnd)
		return nativeTextSplice{start: at, end: at, text: []byte(`<` + prefix + `pPr>` + string(written) + `</` + prefix + `pPr>`)}, nil
	}
	output := []byte(`<` + prefix + `pPr>`)
	placed := false
	for _, child := range properties.Children {
		local := child.Name.Local
		if child.Name.Space != paragraph.Name.Space {
			local = ""
		}
		if local == "jc" {
			output, placed = append(output, written...), true
			continue
		}
		if !placed && nativeParagraphPropertyRank(local) > nativeParagraphPropertyRank("jc") {
			output, placed = append(output, written...), true
		}
		output = append(output, part[child.Start:child.End]...)
	}
	if !placed {
		output = append(output, written...)
	}
	output = append(output, []byte(`</`+prefix+`pPr>`)...)
	return nativeTextSplice{start: properties.Start, end: properties.End, text: output}, nil
}

func nativeParagraphPropertyRank(local string) int {
	for rank, name := range nativeParagraphPropertyOrder {
		if name == local {
			return rank
		}
	}
	return len(nativeParagraphPropertyOrder)
}

// nativeFormatRunProperties merges the patch into the source w:rPr, keeping
// every child the source already carries, in its own position, and inserting
// a newly written property at its ECMA-376 schema position.
func nativeFormatRunProperties(part []byte, owner, properties *nativeXMLNode, patch NativeDOCXRunPropertyPatchV1) ([]byte, error) {
	prefix := ""
	if name := nativeFormatQName(part, owner); strings.Contains(name, ":") {
		prefix = name[:strings.Index(name, ":")+1]
	}
	written := map[string][]byte{}
	attribute := func(value string) string { return ` ` + prefix + `val="` + nativeMutationEscapeAttribute(value) + `"` }
	if patch.Bold != nil {
		written["b"] = []byte(nativeFormatToggle(prefix, "b", *patch.Bold))
	}
	if patch.Italic != nil {
		written["i"] = []byte(nativeFormatToggle(prefix, "i", *patch.Italic))
	}
	if patch.Underline != nil {
		written["u"] = []byte(`<` + prefix + `u` + attribute(*patch.Underline) + `/>`)
	}
	if patch.Color != nil {
		written["color"] = []byte(`<` + prefix + `color` + attribute(*patch.Color) + `/>`)
	}
	if patch.Highlight != nil {
		written["highlight"] = []byte(`<` + prefix + `highlight` + attribute(*patch.Highlight) + `/>`)
	}
	if patch.FontSizeHalfPoints != nil {
		// Only w:sz: native extraction treats a w:szCs it did not model as a
		// preserved complex-script slot and would make the paragraph read-only.
		written["sz"] = []byte(`<` + prefix + `sz` + attribute(strconv.Itoa(*patch.FontSizeHalfPoints)) + `/>`)
	}
	if patch.FontFamily != nil {
		fonts, err := nativeFormatRunFonts(properties, prefix, *patch.FontFamily)
		if err != nil {
			return nil, err
		}
		written["rFonts"] = fonts
	}
	children := []struct {
		local string
		raw   []byte
	}{}
	if properties != nil {
		for _, child := range properties.Children {
			local := child.Name.Local
			if child.Name.Space != owner.Name.Space {
				local = ""
			}
			raw := part[child.Start:child.End]
			if replacement, ok := written[local]; ok && local != "" {
				raw = replacement
				delete(written, local)
			}
			children = append(children, struct {
				local string
				raw   []byte
			}{local, raw})
		}
	}
	pending := make([]string, 0, len(written))
	for local := range written {
		pending = append(pending, local)
	}
	sort.Slice(pending, func(i, j int) bool {
		return nativeRunPropertyRank(pending[i]) < nativeRunPropertyRank(pending[j])
	})
	for _, local := range pending {
		index := len(children)
		for position, child := range children {
			if nativeRunPropertyRank(child.local) > nativeRunPropertyRank(local) {
				index = position
				break
			}
		}
		children = append(children, struct {
			local string
			raw   []byte
		}{})
		copy(children[index+1:], children[index:])
		children[index] = struct {
			local string
			raw   []byte
		}{local, written[local]}
	}
	output := []byte(`<` + prefix + `rPr>`)
	for _, child := range children {
		output = append(output, child.raw...)
	}
	return append(output, []byte(`</`+prefix+`rPr>`)...), nil
}

// nativeFormatRunFonts rewrites the ascii and hAnsi slots of an existing
// w:rFonts, keeping every other slot the source declares.
func nativeFormatRunFonts(properties *nativeXMLNode, prefix, family string) ([]byte, error) {
	escaped := nativeMutationEscapeAttribute(family)
	slots := map[string]string{"ascii": escaped, "hAnsi": escaped}
	order := []string{"ascii", "hAnsi"}
	if properties != nil {
		for _, child := range properties.Children {
			if child.Name.Local != "rFonts" {
				continue
			}
			for _, attr := range child.Attrs {
				if attr.Name.Space != child.Name.Space {
					return nil, fmt.Errorf("run fonts carry an attribute outside the WordprocessingML namespace")
				}
				if attr.Name.Local != "ascii" && attr.Name.Local != "hAnsi" {
					if _, seen := slots[attr.Name.Local]; !seen {
						order = append(order, attr.Name.Local)
					}
					slots[attr.Name.Local] = nativeMutationEscapeAttribute(attr.Value)
				}
			}
			break
		}
	}
	output := `<` + prefix + `rFonts`
	for _, slot := range order {
		output += ` ` + prefix + slot + `="` + slots[slot] + `"`
	}
	return []byte(output + `/>`), nil
}

func nativeFormatToggle(prefix, local string, on bool) string {
	if on {
		return `<` + prefix + local + `/>`
	}
	return `<` + prefix + local + ` ` + prefix + `val="0"/>`
}

func nativeRunPropertyRank(local string) int {
	for rank, name := range nativeRunPropertyOrder {
		if name == local {
			return rank
		}
	}
	return len(nativeRunPropertyOrder)
}

// nativeFormatTextElement reuses the source w:t element for a new text value,
// adding xml:space="preserve" when the split gave the piece edge whitespace.
func nativeFormatTextElement(raw []byte, text string) ([]byte, error) {
	if nativeMutationNeedsPreservedSpace(text) {
		end := nativeStartTagEnd(raw)
		if end <= 0 {
			return nil, fmt.Errorf("text anchor has an unterminated start tag")
		}
		nameEnd := 1
		for nameEnd < end && !nativeMutationXMLSpace(raw[nameEnd]) && raw[nameEnd] != '/' && raw[nameEnd] != '>' {
			nameEnd++
		}
		if !nativeMutationHasSpacePreserve(raw[:end], nameEnd) {
			insert := end - 1
			if insert > 0 && raw[insert-1] == '/' {
				insert--
			}
			spaced := append([]byte(nil), raw[:insert]...)
			spaced = append(spaced, []byte(` xml:space="preserve"`)...)
			raw = append(spaced, raw[insert:]...)
		}
	}
	return replaceNativeTextElement(raw, text)
}

// nativeStartTagEnd returns the offset just past the element's start tag.
func nativeStartTagEnd(raw []byte) int {
	quote := byte(0)
	for index := 0; index < len(raw); index++ {
		character := raw[index]
		if quote != 0 {
			if character == quote {
				quote = 0
			}
			continue
		}
		if character == '\'' || character == '"' {
			quote = character
			continue
		}
		if character == '>' {
			return index + 1
		}
	}
	return -1
}

func nativeFormatQName(part []byte, node *nativeXMLNode) string {
	raw := part[node.Start:node.End]
	end := 1
	for end < len(raw) && !nativeMutationXMLSpace(raw[end]) && raw[end] != '/' && raw[end] != '>' {
		end++
	}
	return string(raw[1:end])
}

func nativeNodeByPath(node *nativeXMLNode, path string) *nativeXMLNode {
	if node.Path == path {
		return node
	}
	if !strings.HasPrefix(path, node.Path+"/") {
		return nil
	}
	for _, child := range node.Children {
		if found := nativeNodeByPath(child, path); found != nil {
			return found
		}
	}
	return nil
}

// nativeUTF16ByteOffset converts a UTF-16 code-unit offset into a byte offset,
// refusing an offset that would fall inside a surrogate pair.
func nativeUTF16ByteOffset(text string, units int) (int, bool) {
	count := 0
	for index, character := range text {
		if count == units {
			return index, true
		}
		count++
		if character > 0xFFFF {
			count++
		}
		if count > units {
			return 0, false
		}
	}
	return len(text), count == units
}

func nativeMutationEscapeAttribute(value string) string {
	return strings.NewReplacer("&", "&amp;", "<", "&lt;", ">", "&gt;", `"`, "&quot;").Replace(value)
}

// validateNativeFormatResults proves, on the re-extracted contract, that the
// selected text is unchanged and that exactly the runs covering the requested
// span carry the requested properties.
func validateNativeFormatResults(after *NativeDocumentV1, requests []nativeFormatRequest) error {
	byPath := map[string]*NativeParagraphV1{}
	forEachNativeParagraph(after, func(paragraph *NativeParagraphV1) {
		byPath[paragraph.Anchor.PartName+"\x00"+paragraph.Anchor.Path] = paragraph
	})
	for _, request := range requests {
		paragraph := byPath[request.paragraph.Anchor.PartName+"\x00"+request.paragraph.Anchor.Path]
		if paragraph == nil {
			return nativeMutationError("POST_WRITE_MISMATCH", request.paragraph.ID, "the formatted paragraph is absent from the re-extracted contract")
		}
		if request.alignment != nil {
			if paragraph.Properties == nil || paragraph.Properties.Alignment == nil || *paragraph.Properties.Alignment != *request.alignment {
				return nativeMutationError("POST_WRITE_MISMATCH", request.paragraph.ID, "the formatted paragraph does not carry the requested alignment")
			}
			continue
		}
		text, offset := strings.Builder{}, 0
		for _, run := range paragraph.Runs {
			if run.Kind != "text" || run.Text == nil {
				continue
			}
			text.WriteString(*run.Text)
			start := offset
			offset += nativeMutationUTF16CodeUnits(*run.Text)
			if start >= request.end || offset <= request.start {
				continue
			}
			if start < request.start || offset > request.end {
				return nativeMutationError("POST_WRITE_MISMATCH", request.paragraph.ID, "a formatted run extends past the requested range")
			}
			if err := nativeFormatPropertiesApplied(run.Properties, request.patch); err != nil {
				return nativeMutationError("POST_WRITE_MISMATCH", request.paragraph.ID, err.Error())
			}
		}
		if text.String() != request.text {
			return nativeMutationError("POST_WRITE_MISMATCH", request.paragraph.ID, "run formatting changed the paragraph's source-authoritative text")
		}
	}
	return nil
}

func nativeFormatPropertiesApplied(properties *NativeRunPropertiesV1, patch NativeDOCXRunPropertyPatchV1) error {
	if properties == nil {
		return fmt.Errorf("a formatted run carries no run properties after write-back")
	}
	mismatch := func(name string) error { return fmt.Errorf("a formatted run does not carry the requested %s", name) }
	if patch.Bold != nil && (properties.Bold == nil || *properties.Bold != *patch.Bold) {
		return mismatch("bold")
	}
	if patch.Italic != nil && (properties.Italic == nil || *properties.Italic != *patch.Italic) {
		return mismatch("italic")
	}
	if patch.Underline != nil && (properties.Underline == nil || *properties.Underline != *patch.Underline) {
		return mismatch("underline")
	}
	if patch.FontFamily != nil && (properties.FontFamily == nil || *properties.FontFamily != *patch.FontFamily) {
		return mismatch("font family")
	}
	if patch.FontSizeHalfPoints != nil && (properties.FontSizeHalfPoint == nil || *properties.FontSizeHalfPoint != *patch.FontSizeHalfPoints) {
		return mismatch("font size")
	}
	if patch.Color != nil && (properties.Color == nil || *properties.Color != *patch.Color) {
		return mismatch("color")
	}
	if patch.Highlight != nil && (properties.Highlight == nil || *properties.Highlight != *patch.Highlight) {
		return mismatch("highlight")
	}
	return nil
}
