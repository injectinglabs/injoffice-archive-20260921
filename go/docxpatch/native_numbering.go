package docxpatch

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"path"
	"sort"
	"strconv"
	"strings"
)

const (
	nativeNumberingContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
	nativeNumberingMaxLevel    = 8
)

// NativeNumberingLevelDefinitionV1 is one authored w:lvl the editor can name
// in its list catalog. It is not a layout projection: marker geometry stays
// on the resolved-layout contract.
type NativeNumberingLevelDefinitionV1 struct {
	Level  int    `json:"level"`
	Format string `json:"format"`
	Text   string `json:"text"`
	Suffix string `json:"suffix,omitempty"`
	Start  int    `json:"start"`
}

// NativeNumberingDefinitionV1 is one concrete w:num the extractor could read
// from the numbering part. numId "0" is Word's "no list" sentinel and is never
// catalogued.
type NativeNumberingDefinitionV1 struct {
	NumID  string                             `json:"num_id"`
	Levels []NativeNumberingLevelDefinitionV1 `json:"levels"`
}

func nativeNumberingPatch(properties *NativeDOCXParagraphPropertyPatchV1) bool {
	return properties != nil && (properties.NumID != nil || properties.Kind != nil)
}

func nativeNumberingRemove(properties *NativeDOCXParagraphPropertyPatchV1) bool {
	if properties == nil || properties.NumID == nil || properties.Kind != nil {
		return false
	}
	id := *properties.NumID
	return id == "" || id == "0"
}

func validateNativeNumberingPatch(properties *NativeDOCXParagraphPropertyPatchV1) error {
	if properties == nil || !nativeNumberingPatch(properties) {
		return fmt.Errorf("numbering must set numbering_num_id or numbering_kind")
	}
	if properties.Alignment != nil {
		return fmt.Errorf("must patch numbering on its own")
	}
	if properties.Kind != nil {
		if *properties.Kind != "bullet" && *properties.Kind != "decimal" {
			return fmt.Errorf("numbering_kind must be bullet or decimal")
		}
		if nativeNumberingRemove(properties) {
			return fmt.Errorf("numbering_kind cannot remove a list")
		}
	}
	if properties.NumID != nil && *properties.NumID != "" && *properties.NumID != "0" && !nativeIDPattern.MatchString(*properties.NumID) {
		return fmt.Errorf("numbering_num_id must be a bounded native identifier")
	}
	if properties.Level != nil && (*properties.Level < 0 || *properties.Level > nativeNumberingMaxLevel) {
		return fmt.Errorf("numbering_level must be 0..%d", nativeNumberingMaxLevel)
	}
	if properties.NumID == nil && properties.Kind == nil {
		return fmt.Errorf("numbering must set numbering_num_id or numbering_kind")
	}
	return nil
}

func nativeNumberingRequestedLevel(properties *NativeDOCXParagraphPropertyPatchV1) int {
	if properties != nil && properties.Level != nil {
		return *properties.Level
	}
	return 0
}

func nativeExtractNumberingDefinitions(pkg *nativePackage, mainPart, wordNS string) []NativeNumberingDefinitionV1 {
	rel, partName, ok := nativeNumberingRelatedPart(pkg, mainPart)
	if !ok || rel.External || rel.Dangling || partName == "" {
		return nil
	}
	if !nativeASCIIEqual(pkg.contentTypes[partName], nativeNumberingContentType) {
		return nil
	}
	root, err := parseNativeXML(partName, pkg.files[partName])
	if err != nil || root.Name != (xml.Name{Space: wordNS, Local: "numbering"}) {
		return nil
	}
	abstracts := map[string]map[int]NativeNumberingLevelDefinitionV1{}
	for _, child := range root.Children {
		if child.Name != (xml.Name{Space: wordNS, Local: "abstractNum"}) {
			continue
		}
		id, ok := nativeDecimalIDAttr(child, wordNS, "abstractNumId")
		if !ok {
			continue
		}
		levels := map[int]NativeNumberingLevelDefinitionV1{}
		for _, property := range child.Children {
			if property.Name != (xml.Name{Space: wordNS, Local: "lvl"}) {
				continue
			}
			level, ok := nativeCatalogNumberingLevel(property, wordNS)
			if !ok {
				continue
			}
			if _, duplicate := levels[level.Level]; duplicate {
				delete(levels, level.Level)
				continue
			}
			levels[level.Level] = level
		}
		if len(levels) > 0 {
			abstracts[id] = levels
		}
	}
	seen := map[string]bool{}
	definitions := []NativeNumberingDefinitionV1{}
	for _, child := range root.Children {
		if child.Name != (xml.Name{Space: wordNS, Local: "num"}) {
			continue
		}
		id, ok := nativeDecimalIDAttr(child, wordNS, "numId")
		if !ok || id == "0" || seen[id] {
			continue
		}
		abstractID := ""
		for _, property := range child.Children {
			if property.Name == (xml.Name{Space: wordNS, Local: "abstractNumId"}) {
				if value, valid := nativeDecimalIDAttr(property, wordNS, "val"); valid {
					abstractID = value
				}
			}
		}
		levels := abstracts[abstractID]
		if abstractID == "" || len(levels) == 0 {
			continue
		}
		seen[id] = true
		ordered := make([]NativeNumberingLevelDefinitionV1, 0, len(levels))
		keys := make([]int, 0, len(levels))
		for level := range levels {
			keys = append(keys, level)
		}
		sort.Ints(keys)
		for _, key := range keys {
			ordered = append(ordered, levels[key])
		}
		definitions = append(definitions, NativeNumberingDefinitionV1{NumID: id, Levels: ordered})
	}
	if len(definitions) == 0 {
		return nil
	}
	return definitions
}

func nativeCatalogNumberingLevel(node *nativeXMLNode, wordNS string) (NativeNumberingLevelDefinitionV1, bool) {
	level, ok := nativeNumberingLevelAttr(node, wordNS, "ilvl")
	if !ok || firstDirectNativeChild(node, wordNS, "lvlPicBulletId") != nil {
		return NativeNumberingLevelDefinitionV1{}, false
	}
	formatNode := firstDirectNativeChild(node, wordNS, "numFmt")
	textNode := firstDirectNativeChild(node, wordNS, "lvlText")
	if formatNode == nil || textNode == nil {
		return NativeNumberingLevelDefinitionV1{}, false
	}
	format, okFormat := nativeAttr(formatNode, wordNS, "val")
	text, okText := nativeAttr(textNode, wordNS, "val")
	if !okFormat || !okText || format == "" || text == "" || !nativeExactLeaf(formatNode, xml.Name{Space: wordNS, Local: "val"}) || !nativeExactLeaf(textNode, xml.Name{Space: wordNS, Local: "val"}) {
		return NativeNumberingLevelDefinitionV1{}, false
	}
	start := 1
	if startNode := firstDirectNativeChild(node, wordNS, "start"); startNode != nil {
		if raw, ok := nativeAttr(startNode, wordNS, "val"); ok {
			if parsed, err := strconv.Atoi(raw); err == nil && parsed >= 0 {
				start = parsed
			}
		}
	}
	suffix := ""
	if suffixNode := firstDirectNativeChild(node, wordNS, "suff"); suffixNode != nil {
		if raw, ok := nativeAttr(suffixNode, wordNS, "val"); ok {
			suffix = raw
		}
	}
	return NativeNumberingLevelDefinitionV1{Level: level, Format: format, Text: text, Suffix: suffix, Start: start}, true
}

func nativeValidateNumberingDefinitions(definitions []NativeNumberingDefinitionV1, add func(code, path, message string), collection func(length int, path string, maximum int) int) {
	seen := map[string]bool{}
	for i, limit := 0, collection(len(definitions), "/numbering_definitions", NativeDOCXMaxCollectionItems); i < limit; i++ {
		path := fmt.Sprintf("/numbering_definitions/%d", i)
		definition := definitions[i]
		if !nativeIDPattern.MatchString(definition.NumID) {
			add("INVALID_VALUE", path+"/num_id", "must be a bounded native identifier")
		}
		if definition.NumID == "0" {
			add("INVALID_VALUE", path+"/num_id", "numId 0 is Word's no-list sentinel and is not a catalogued instance")
		}
		if seen[definition.NumID] {
			add("DUPLICATE_ID", path+"/num_id", "numbering instance id is duplicated")
		}
		seen[definition.NumID] = true
		if len(definition.Levels) == 0 {
			add("REQUIRED", path+"/levels", "a numbering instance requires at least one level")
		}
		levels := map[int]bool{}
		for j, limit := 0, collection(len(definition.Levels), path+"/levels", nativeNumberingMaxLevel+1); j < limit; j++ {
			levelPath := fmt.Sprintf("%s/levels/%d", path, j)
			level := definition.Levels[j]
			if level.Level < 0 || level.Level > nativeNumberingMaxLevel {
				add("OUT_OF_RANGE", levelPath+"/level", fmt.Sprintf("must be 0..%d", nativeNumberingMaxLevel))
			}
			if levels[level.Level] {
				add("DUPLICATE_ID", levelPath+"/level", "level index is duplicated")
			}
			levels[level.Level] = true
			if level.Format == "" {
				add("REQUIRED", levelPath+"/format", "field is required")
			}
			if level.Text == "" {
				add("REQUIRED", levelPath+"/text", "field is required")
			}
			if level.Start < 0 {
				add("OUT_OF_RANGE", levelPath+"/start", "must be a nonnegative integer")
			}
		}
	}
}

func nativeNumberingRelatedPart(pkg *nativePackage, mainPart string) (nativeRelationship, string, bool) {
	var found *nativeRelationship
	for i := range pkg.rels[mainPart] {
		rel := &pkg.rels[mainPart][i]
		if rel.Type != relBaseTransitional+"numbering" && rel.Type != relBaseStrict+"numbering" {
			continue
		}
		if found != nil {
			return nativeRelationship{}, "", false
		}
		found = rel
	}
	if found == nil {
		return nativeRelationship{}, "", false
	}
	return *found, found.PartName, true
}

// ApplyNativeNumberingMutationsV1 attaches, indents, or removes a paragraph
// list by writing w:numPr in schema order and creating or reusing the
// numbering part. Every other package byte is preserved; the re-extracted
// contract must carry the requested numbering reference (or none).
func ApplyNativeNumberingMutationsV1(packageBytes []byte, expectedRevision string, mutations []NativeDOCXFormatMutationV1) (*NativeDOCXMutationResultV1, error) {
	if len(mutations) == 0 || len(mutations) > NativeDOCXMaxMutations {
		return nil, nativeMutationError("INVALID_MUTATION_COUNT", "", fmt.Sprintf("mutation count must be 1..%d", NativeDOCXMaxMutations))
	}
	if !nativeSHA256.MatchString(expectedRevision) {
		return nil, nativeMutationError("INVALID_REVISION", "", "expected revision must be sha256 followed by the full lowercase exact-byte digest")
	}
	doc, err := ExtractNativeDocumentV1(packageBytes)
	if err != nil {
		return nil, fmt.Errorf("docxpatch: native numbering mutation source validation: %w", err)
	}
	exactRevision := nativeSHA(packageBytes)
	if doc.Source.PackageSHA256 != exactRevision || expectedRevision != exactRevision {
		return nil, nativeMutationError("STALE_REVISION", "", fmt.Sprintf("expected %q, current exact package revision is %q", expectedRevision, exactRevision))
	}
	pkg, err := openNativeDOCXPackage(packageBytes)
	if err != nil {
		return nil, fmt.Errorf("docxpatch: native numbering mutation reopen source: %w", err)
	}
	if nativeDOCXPackageHasDigitalSignature(pkg) {
		return nil, nativeMutationError("SIGNED_PACKAGE", "", "a signed OPC package cannot be mutated without invalidating its digital signature")
	}
	paragraphs := map[string]*NativeParagraphV1{}
	forEachNativeParagraph(doc, func(paragraph *NativeParagraphV1) {
		paragraphs[paragraph.ID] = paragraph
	})
	requests := make([]nativeNumberingResultRequest, 0, len(mutations))
	seen := map[string]bool{}
	for _, mutation := range mutations {
		if err := validateNativeDOCXFormatMutation(&mutation); err != nil {
			return nil, nativeMutationError("INVALID_SELECTOR", mutation.TargetID, err.Error())
		}
		if mutation.TargetKind != "paragraph" || !nativeNumberingPatch(mutation.ParagraphProperties) {
			return nil, nativeMutationError("INVALID_SELECTOR", mutation.TargetID, "numbering is a paragraph property and needs a paragraph target")
		}
		paragraph := paragraphs[mutation.TargetID]
		if paragraph == nil {
			return nil, nativeMutationError("TARGET_NOT_FOUND", mutation.TargetID, "target is absent from the current native contract")
		}
		if paragraph.EditPolicy.Mode != "read-write" || !nativePolicyAllows(paragraph.EditPolicy, "properties.patch") {
			message := "owning paragraph does not allow property patches"
			if paragraph.EditPolicy.Refusal != nil {
				message += ": " + paragraph.EditPolicy.Refusal.Code
			}
			return nil, nativeMutationError("UNSUPPORTED_CONSTRUCT", mutation.TargetID, message)
		}
		if mutation.ExpectedXMLSHA256 != paragraph.Anchor.XMLSHA256 {
			return nil, nativeMutationError("STALE_TARGET", mutation.TargetID, fmt.Sprintf("expected XML fingerprint %q, current fingerprint is %q", mutation.ExpectedXMLSHA256, paragraph.Anchor.XMLSHA256))
		}
		if seen[paragraph.ID] {
			return nil, nativeMutationError("DUPLICATE_TARGET", mutation.TargetID, "a paragraph may be numbered only once in an atomic transaction")
		}
		seen[paragraph.ID] = true
		request := nativeNumberingResultRequest{paragraph: paragraph, remove: nativeNumberingRemove(mutation.ParagraphProperties), level: nativeNumberingRequestedLevel(mutation.ParagraphProperties)}
		if mutation.ParagraphProperties.Kind != nil {
			request.kind = *mutation.ParagraphProperties.Kind
		}
		if mutation.ParagraphProperties.NumID != nil {
			request.numID = *mutation.ParagraphProperties.NumID
		}
		if !request.remove && request.kind == "" && (request.numID == "" || request.numID == "0") {
			return nil, nativeMutationError("INVALID_SELECTOR", mutation.TargetID, "numbering must set numbering_num_id or numbering_kind")
		}
		requests = append(requests, request)
	}

	replace := map[string][]byte{}
	add := map[string][]byte{}
	resolvedIDs := map[string]string{}
	wordNS := nativeWordNS(pkg, doc.Source.MainPart)
	for i := range requests {
		request := &requests[i]
		if request.remove {
			continue
		}
		if request.kind != "" {
			if id, ok := resolvedIDs[request.kind]; ok {
				request.numID = id
				continue
			}
			overlay := nativeNumberingOverlayPackage(pkg, replace, add)
			numID, extraReplace, extraAdd, fresh, ensureErr := nativeEnsureNumberingInstance(overlay, doc.Source.MainPart, wordNS, request.kind)
			if ensureErr != nil {
				return nil, nativeMutationError("UNSUPPORTED_CONSTRUCT", request.paragraph.ID, ensureErr.Error())
			}
			for name, part := range extraReplace {
				replace[name] = part
			}
			for name, part := range extraAdd {
				add[name] = part
			}
			request.numID = numID
			resolvedIDs[request.kind] = numID
			if fresh {
				resolvedIDs["bullet"] = "1"
				resolvedIDs["decimal"] = "2"
			}
			continue
		}
		if !nativeNumberingInstanceHasLevel(pkg, doc.Source.MainPart, wordNS, replace, add, request.numID, request.level) {
			return nil, nativeMutationError("UNSUPPORTED_CONSTRUCT", request.paragraph.ID, fmt.Sprintf("numbering instance %q has no authored level %d", request.numID, request.level))
		}
	}

	roots := map[string]*nativeXMLNode{}
	splicesByPart := map[string][]nativeTextSplice{}
	for _, request := range requests {
		partName := request.paragraph.Anchor.PartName
		root, ok := roots[partName]
		if !ok {
			parsed, parseErr := parseNativeXML(partName, pkg.files[partName])
			if parseErr != nil {
				return nil, fmt.Errorf("docxpatch: native numbering: %w", parseErr)
			}
			root = parsed
			roots[partName] = root
		}
		splice, spliceErr := nativeNumberingParagraphSplice(pkg.files[partName], root, request.paragraph, request.remove, request.numID, request.level)
		if spliceErr != nil {
			return nil, nativeMutationError("UNSUPPORTED_LEXICAL_FORM", request.paragraph.ID, spliceErr.Error())
		}
		splicesByPart[partName] = append(splicesByPart[partName], splice)
	}
	for partName, splices := range splicesByPart {
		source := pkg.files[partName]
		if replacement, ok := replace[partName]; ok {
			source = replacement
		}
		part, spliceErr := applyNativeTextSplices(source, splices)
		if spliceErr != nil {
			return nil, spliceErr
		}
		if _, exists := pkg.files[partName]; exists {
			replace[partName] = part
		} else {
			add[partName] = part
		}
	}
	return writeNativeNumberingPackage(packageBytes, pkg, doc, exactRevision, replace, add, func(after *NativeDocumentV1) error {
		return validateNativeNumberingResults(after, requests)
	})
}

func nativeWordNS(pkg *nativePackage, mainPart string) string {
	root, err := parseNativeXML(mainPart, pkg.files[mainPart])
	if err != nil {
		return wordMLTransitional
	}
	if root.Name.Space == wordMLStrict {
		return wordMLStrict
	}
	return wordMLTransitional
}

func nativeNumberingRelBase(wordNS string) string {
	if wordNS == wordMLStrict {
		return relBaseStrict
	}
	return relBaseTransitional
}

func nativeNumberingInstanceHasLevel(pkg *nativePackage, mainPart, wordNS string, replace, add map[string][]byte, numID string, level int) bool {
	catalog := nativeExtractNumberingDefinitions(nativeNumberingOverlayPackage(pkg, replace, add), mainPart, wordNS)
	for _, definition := range catalog {
		if definition.NumID != numID {
			continue
		}
		for _, item := range definition.Levels {
			if item.Level == level {
				return true
			}
		}
	}
	return false
}

func nativeNumberingOverlayPackage(pkg *nativePackage, replace, add map[string][]byte) *nativePackage {
	overlay := *pkg
	overlay.files = map[string][]byte{}
	for name, part := range pkg.files {
		overlay.files[name] = part
	}
	for name, part := range replace {
		overlay.files[name] = part
	}
	for name, part := range add {
		overlay.files[name] = part
	}
	overlay.contentTypes = pkg.contentTypes
	overlay.rels = pkg.rels
	overlay.relsPart = pkg.relsPart
	return &overlay
}

func nativeEnsureNumberingInstance(pkg *nativePackage, mainPart, wordNS, kind string) (string, map[string][]byte, map[string][]byte, bool, error) {
	replace := map[string][]byte{}
	add := map[string][]byte{}
	if existing := nativeFindNumberingKind(pkg, mainPart, wordNS, kind); existing != "" {
		return existing, replace, add, false, nil
	}
	rel, partName, hasRel := nativeNumberingRelatedPart(pkg, mainPart)
	if hasRel {
		if rel.External || rel.Dangling || partName == "" {
			return "", nil, nil, false, fmt.Errorf("the numbering relationship is not an internal package part")
		}
		if !nativeASCIIEqual(pkg.contentTypes[partName], nativeNumberingContentType) {
			return "", nil, nil, false, fmt.Errorf("numbering part %q has content type %q", partName, pkg.contentTypes[partName])
		}
		updated, numID, err := nativeAppendNumberingKind(pkg.files[partName], wordNS, kind)
		if err != nil {
			return "", nil, nil, false, err
		}
		replace[partName] = updated
		return numID, replace, add, false, nil
	}
	partName = nativeSiblingPart(mainPart, "numbering.xml")
	if _, exists := pkg.files[partName]; exists {
		return "", nil, nil, false, fmt.Errorf("numbering part %q already exists without a numbering relationship", partName)
	}
	relsPart := pkg.relsPart[mainPart]
	relsXML := pkg.files[relsPart]
	hasRels := relsPart != "" && len(relsXML) > 0
	if !hasRels {
		relsPart = nativeRelsPartName(mainPart)
		if _, exists := pkg.files[relsPart]; exists {
			return "", nil, nil, false, fmt.Errorf("relationship part %q already exists", relsPart)
		}
		relsXML = []byte(emptyRelsXML)
	}
	relID := nextFreeRelID(string(relsXML))
	newRels, err := appendRelationship(string(relsXML), relID, nativeNumberingRelBase(wordNS)+"numbering", nativeOPCRelTarget(mainPart, partName))
	if err != nil {
		return "", nil, nil, false, err
	}
	ct, err := nativeOverrideContentType(string(pkg.files[contentTypes]), partName, nativeNumberingContentType)
	if err != nil {
		return "", nil, nil, false, err
	}
	numID := "1"
	if kind == "decimal" {
		numID = "2"
	}
	add[partName] = []byte(nativeFreshNumberingPart(wordNS))
	replace[contentTypes] = []byte(ct)
	if hasRels {
		replace[relsPart] = []byte(newRels)
	} else {
		add[relsPart] = []byte(newRels)
	}
	return numID, replace, add, true, nil
}

func nativeFindNumberingKind(pkg *nativePackage, mainPart, wordNS, kind string) string {
	for _, definition := range nativeExtractNumberingDefinitions(pkg, mainPart, wordNS) {
		if len(definition.Levels) == 0 {
			continue
		}
		format := definition.Levels[0].Format
		if kind == "bullet" && format == "bullet" {
			return definition.NumID
		}
		if kind == "decimal" && format == "decimal" {
			return definition.NumID
		}
	}
	return ""
}

func nativeAppendNumberingKind(part []byte, wordNS, kind string) ([]byte, string, error) {
	root, err := parseNativeXML("word/numbering.xml", part)
	if err != nil {
		return nil, "", err
	}
	if root.Name != (xml.Name{Space: wordNS, Local: "numbering"}) {
		return nil, "", fmt.Errorf("numbering part has spoofed or invalid root")
	}
	maxAbstract, maxNum := -1, 0
	var lastAbstract, lastNum *nativeXMLNode
	mac := (*nativeXMLNode)(nil)
	for _, child := range root.Children {
		if child.Name.Space != wordNS {
			continue
		}
		switch child.Name.Local {
		case "abstractNum":
			lastAbstract = child
			if id, ok := nativeDecimalIDAttr(child, wordNS, "abstractNumId"); ok {
				if n, err := strconv.Atoi(id); err == nil && n > maxAbstract {
					maxAbstract = n
				}
			}
		case "num":
			lastNum = child
			if id, ok := nativeDecimalIDAttr(child, wordNS, "numId"); ok {
				if n, err := strconv.Atoi(id); err == nil && n > maxNum {
					maxNum = n
				}
			}
		case "numIdMacAtCleanup":
			mac = child
		}
	}
	abstractID := strconv.Itoa(maxAbstract + 1)
	numID := strconv.Itoa(maxNum + 1)
	if numID == "0" {
		numID = "1"
	}
	prefix := nativeXMLPrefix(part, root)
	markup := nativeNumberingAbstractXML(prefix, abstractID, kind) + nativeNumberingNumXML(prefix, numID, abstractID)
	at := nativeNumberingInsertPoint(part, root, lastAbstract, lastNum, mac)
	spliced, err := applyNativeTextSplices(part, []nativeTextSplice{{start: at, end: at, text: []byte(markup)}})
	if err != nil {
		return nil, "", err
	}
	return spliced, numID, nil
}

func nativeNumberingInsertPoint(part []byte, root, lastAbstract, lastNum, mac *nativeXMLNode) int64 {
	if lastNum != nil {
		return lastNum.End
	}
	if lastAbstract != nil {
		return lastAbstract.End
	}
	if mac != nil {
		return mac.Start
	}
	tagEnd := nativeStartTagEnd(part[root.Start:root.End])
	if tagEnd <= 0 {
		return root.End
	}
	return root.Start + int64(tagEnd)
}

func nativeFreshNumberingPart(wordNS string) string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="` + wordNS + `">` +
		nativeNumberingAbstractXML("w:", "0", "bullet") +
		nativeNumberingAbstractXML("w:", "1", "decimal") +
		nativeNumberingNumXML("w:", "1", "0") +
		nativeNumberingNumXML("w:", "2", "1") +
		`</w:numbering>`
}

func nativeNumberingAbstractXML(prefix, id, kind string) string {
	format := kind
	if kind != "bullet" {
		format = "decimal"
	}
	var b strings.Builder
	b.WriteString(`<` + prefix + `abstractNum ` + prefix + `abstractNumId="` + nativeMutationEscapeAttribute(id) + `"><` + prefix + `multiLevelType ` + prefix + `val="hybridMultilevel"/>`)
	for level := 0; level <= nativeNumberingMaxLevel; level++ {
		text := "•"
		if format != "bullet" {
			text = "%" + strconv.Itoa(level+1) + "."
		}
		b.WriteString(`<` + prefix + `lvl ` + prefix + `ilvl="` + strconv.Itoa(level) + `"><` + prefix + `start ` + prefix + `val="1"/><` + prefix + `numFmt ` + prefix + `val="` + format + `"/><` + prefix + `suff ` + prefix + `val="space"/><` + prefix + `lvlText ` + prefix + `val="` + nativeMutationEscapeAttribute(text) + `"/><` + prefix + `lvlJc ` + prefix + `val="left"/><` + prefix + `pPr><` + prefix + `ind ` + prefix + `left="` + strconv.Itoa(720*(level+1)) + `" ` + prefix + `hanging="360"/></` + prefix + `pPr></` + prefix + `lvl>`)
	}
	b.WriteString(`</` + prefix + `abstractNum>`)
	return b.String()
}

func nativeNumberingNumXML(prefix, numID, abstractID string) string {
	return `<` + prefix + `num ` + prefix + `numId="` + nativeMutationEscapeAttribute(numID) + `"><` + prefix + `abstractNumId ` + prefix + `val="` + nativeMutationEscapeAttribute(abstractID) + `"/></` + prefix + `num>`
}

func nativeSiblingPart(mainPart, name string) string {
	dir := path.Dir(mainPart)
	if dir == "." {
		return name
	}
	return dir + "/" + name
}

func nativeRelsPartName(partName string) string {
	dir := path.Dir(partName)
	base := path.Base(partName) + ".rels"
	if dir == "." {
		return "_rels/" + base
	}
	return dir + "/_rels/" + base
}

func nativeOPCRelTarget(sourcePart, targetPart string) string {
	sourceDir := path.Dir(sourcePart)
	if sourceDir == "." {
		return targetPart
	}
	prefix := sourceDir + "/"
	if strings.HasPrefix(targetPart, prefix) {
		return targetPart[len(prefix):]
	}
	return path.Base(targetPart)
}

func nativeOverrideContentType(ct, partName, contentType string) (string, error) {
	partURI := "/" + strings.TrimPrefix(partName, "/")
	if strings.Contains(ct, `PartName="`+partURI+`"`) {
		return ct, nil
	}
	idx := strings.LastIndex(ct, "</Types>")
	if idx < 0 {
		return "", fmt.Errorf("malformed [Content_Types].xml")
	}
	entry := `<Override PartName="` + nativeMutationEscapeAttribute(partURI) + `" ContentType="` + nativeMutationEscapeAttribute(contentType) + `"/>`
	return ct[:idx] + entry + ct[idx:], nil
}

func nativeXMLPrefix(part []byte, node *nativeXMLNode) string {
	name := nativeFormatQName(part, node)
	if i := strings.Index(name, ":"); i >= 0 {
		return name[:i+1]
	}
	return ""
}

func nativeNumberingParagraphSplice(part []byte, root *nativeXMLNode, paragraph *NativeParagraphV1, remove bool, numID string, level int) (nativeTextSplice, error) {
	node := nativeNodeByPath(root, paragraph.Anchor.Path)
	if node == nil || paragraph.Anchor.StartByte == nil || paragraph.Anchor.EndByte == nil ||
		node.Start != *paragraph.Anchor.StartByte || node.End != *paragraph.Anchor.EndByte ||
		nativeSHA(part[node.Start:node.End]) != paragraph.Anchor.XMLSHA256 {
		return nativeTextSplice{}, fmt.Errorf("paragraph source bytes no longer match the issued anchor")
	}
	prefix := nativeXMLPrefix(part, node)
	written := []byte(nil)
	if !remove {
		written = []byte(`<` + prefix + `numPr><` + prefix + `ilvl ` + prefix + `val="` + strconv.Itoa(level) + `"/><` + prefix + `numId ` + prefix + `val="` + nativeMutationEscapeAttribute(numID) + `"/></` + prefix + `numPr>`)
	}
	properties := firstDirectNativeChild(node, node.Name.Space, "pPr")
	if properties != nil && node.Children[0] != properties {
		return nativeTextSplice{}, fmt.Errorf("paragraph properties must be the paragraph's first child")
	}
	if properties == nil {
		if remove {
			return nativeTextSplice{}, fmt.Errorf("paragraph has no numbering to remove")
		}
		tagEnd := nativeStartTagEnd(part[node.Start:node.End])
		if tagEnd <= 0 {
			return nativeTextSplice{}, fmt.Errorf("paragraph element has an unterminated start tag")
		}
		at := node.Start + int64(tagEnd)
		return nativeTextSplice{start: at, end: at, text: []byte(`<` + prefix + `pPr>` + string(written) + `</` + prefix + `pPr>`)}, nil
	}
	if existing := firstDirectNativeChild(properties, node.Name.Space, "numPr"); existing != nil {
		if err := nativeNumberingNumPrWritable(existing, node.Name.Space); err != nil {
			return nativeTextSplice{}, err
		}
	} else if remove {
		return nativeTextSplice{}, fmt.Errorf("paragraph has no numbering to remove")
	}
	output := []byte(`<` + prefix + `pPr>`)
	placed := remove
	children := 0
	for _, child := range properties.Children {
		local := child.Name.Local
		if child.Name.Space != node.Name.Space {
			local = ""
		}
		if local == "numPr" {
			if !remove {
				output = append(output, written...)
				placed = true
				children++
			}
			continue
		}
		if !placed && nativeParagraphPropertyRank(local) > nativeParagraphPropertyRank("numPr") {
			output = append(output, written...)
			placed = true
			children++
		}
		output = append(output, part[child.Start:child.End]...)
		children++
	}
	if !placed {
		output = append(output, written...)
		children++
	}
	if children == 0 {
		tagEnd := nativeStartTagEnd(part[node.Start:node.End])
		if tagEnd <= 0 {
			return nativeTextSplice{}, fmt.Errorf("paragraph element has an unterminated start tag")
		}
		return nativeTextSplice{start: properties.Start, end: properties.End, text: nil}, nil
	}
	output = append(output, []byte(`</`+prefix+`pPr>`)...)
	return nativeTextSplice{start: properties.Start, end: properties.End, text: output}, nil
}

func nativeNumberingNumPrWritable(node *nativeXMLNode, wordNS string) error {
	if len(node.Attrs) != 0 || !nativeXMLWhitespaceOnly(node.Text) {
		return fmt.Errorf("numbering properties carry attributes or text this tier cannot rewrite")
	}
	seen := map[string]bool{}
	for _, child := range node.Children {
		if child.Name.Space != wordNS || (child.Name.Local != "ilvl" && child.Name.Local != "numId") {
			return fmt.Errorf("numbering properties carry children this tier cannot rewrite")
		}
		if seen[child.Name.Local] || !nativeExactLeaf(child, xml.Name{Space: wordNS, Local: "val"}) {
			return fmt.Errorf("numbering properties are not an exact ilvl/numId pair")
		}
		seen[child.Name.Local] = true
	}
	return nil
}

func writeNativeNumberingPackage(packageBytes []byte, pkg *nativePackage, doc *NativeDocumentV1, exactRevision string, replace, add map[string][]byte, readback func(after *NativeDocumentV1) error) (*NativeDOCXMutationResultV1, error) {
	for name, part := range replace {
		if bytes.Equal(part, pkg.files[name]) {
			delete(replace, name)
		}
		if len(part) > NativeDOCXMaxXMLPartBytes && strings.HasSuffix(strings.ToLower(name), ".xml") {
			return nil, nativeMutationError("RESOURCE_LIMIT", "", fmt.Sprintf("mutated XML part %q exceeds %d bytes", name, NativeDOCXMaxXMLPartBytes))
		}
	}
	if len(replace) == 0 && len(add) == 0 {
		return nil, nativeMutationError("SEMANTIC_NO_OP", "", "mutation batch produced no native part payload change")
	}
	produced, err := ApplyPatch(packageBytes, Patch{Replace: replace, Add: add})
	if err != nil {
		return nil, fmt.Errorf("docxpatch: native numbering write verification: %w", err)
	}
	after, err := ExtractNativeDocumentV1WithOptions(produced, NativeExtractionOptions{Previous: doc, RetainPathIdentity: true})
	if err != nil {
		return nil, fmt.Errorf("docxpatch: native numbering post-write extraction: %w", err)
	}
	if issues := ValidateNativeDocumentV1(after); len(issues) > 0 {
		return nil, fmt.Errorf("docxpatch: native numbering post-write validation: %w", &NativeValidationError{Issues: issues})
	}
	if err := readback(after); err != nil {
		return nil, err
	}
	changedNames := make([]string, 0, len(replace)+len(add))
	for name := range replace {
		changedNames = append(changedNames, name)
	}
	for name := range add {
		changedNames = append(changedNames, name)
	}
	sort.Strings(changedNames)
	changed := make([]NativeDOCXChangedPartV1, 0, len(changedNames))
	for _, name := range changedNames {
		before := []byte(nil)
		if part, ok := pkg.files[name]; ok {
			before = part
		}
		afterPart := replace[name]
		if afterPart == nil {
			afterPart = add[name]
		}
		changed = append(changed, NativeDOCXChangedPartV1{PartName: name, BeforeSHA256: nativeSHA(before), AfterSHA256: nativeSHA(afterPart)})
	}
	resultRevision := nativeSHA(produced)
	if after.Source.PackageSHA256 != resultRevision {
		return nil, nativeMutationError("POST_WRITE_MISMATCH", "", "reopened source fingerprint does not match the produced package bytes")
	}
	return &NativeDOCXMutationResultV1{Package: produced, Document: after, Evidence: NativeDOCXMutationEvidenceV1{SourceRevision: exactRevision, ResultRevision: resultRevision, ChangedParts: changed, UntouchedPartsVerified: len(pkg.files) - len(replace)}}, nil
}

type nativeNumberingResultRequest struct {
	paragraph *NativeParagraphV1
	remove    bool
	numID     string
	level     int
	kind      string
}

func validateNativeNumberingResults(after *NativeDocumentV1, requests []nativeNumberingResultRequest) error {
	byPath := map[string]*NativeParagraphV1{}
	forEachNativeParagraph(after, func(paragraph *NativeParagraphV1) {
		byPath[paragraph.Anchor.PartName+"\x00"+paragraph.Anchor.Path] = paragraph
	})
	for _, request := range requests {
		paragraph := byPath[request.paragraph.Anchor.PartName+"\x00"+request.paragraph.Anchor.Path]
		if paragraph == nil {
			return nativeMutationError("POST_WRITE_MISMATCH", request.paragraph.ID, "the numbered paragraph is absent from the re-extracted contract")
		}
		if request.remove {
			if paragraph.Properties != nil && paragraph.Properties.Numbering != nil && paragraph.Properties.Numbering.NumID != "0" {
				return nativeMutationError("POST_WRITE_MISMATCH", request.paragraph.ID, "the paragraph still carries numbering after removal")
			}
			continue
		}
		if paragraph.Properties == nil || paragraph.Properties.Numbering == nil {
			return nativeMutationError("POST_WRITE_MISMATCH", request.paragraph.ID, "the formatted paragraph does not carry numbering")
		}
		got := paragraph.Properties.Numbering
		if got.NumID != request.numID || nativeIntValue(got.Level) != request.level {
			return nativeMutationError("POST_WRITE_MISMATCH", request.paragraph.ID, fmt.Sprintf("re-extracted numbering is numId %s ilvl %d, want numId %s ilvl %d", got.NumID, nativeIntValue(got.Level), request.numID, request.level))
		}
	}
	return nil
}
