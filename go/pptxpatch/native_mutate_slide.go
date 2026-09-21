package pptxpatch

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"net/url"
	"path"
	"sort"
	"strconv"
)

func isNativeSlideMutation(kind NativePPTXMutationKind) bool {
	return kind == NativePPTXInsertSlide || kind == NativePPTXSetSlideBackground
}

// Slide edits do not derive source XML from the preview. The anchor is resolved
// again from the exact input bytes; inherited layout/master parts are untouched.
func applyNativeSlideMutation(orig []byte, before NativePPTXDeck, operation NativePPTXMutation) ([]byte, error) {
	var target *NativeSlide
	var index int
	for i := range before.Slides {
		if before.Slides[i].ID == operation.SlideID {
			target = &before.Slides[i]
			index = i
			break
		}
	}
	if target == nil || target.Source == nil {
		return nil, fmt.Errorf("pptxpatch: native mutations: source slide not found")
	}
	if target.Source.FingerprintSHA256 != operation.ExpectedFingerprintSHA256 {
		return nil, fmt.Errorf("pptxpatch: native mutations: stale slide fingerprint")
	}
	pkg, err := openNativeExtractPackage(orig)
	if err != nil {
		return nil, err
	}
	changes := map[string][]byte{}
	newPart := ""
	if operation.Kind == NativePPTXSetSlideBackground {
		part := target.Source.PartName
		root, err := parseNativeXML(pkg.parts[part], part)
		if err != nil {
			return nil, err
		}
		dialect, err := nativeDialectForPresentation(xml.Name{Space: root.Name.Space, Local: "presentation"})
		if err != nil {
			return nil, err
		}
		common, err := nativeSingleton(root, dialect.presentation, "cSld", true)
		if err != nil {
			return nil, err
		}
		tree, err := nativeSingleton(common, dialect.presentation, "spTree", true)
		if err != nil {
			return nil, err
		}
		bg, err := nativeSingleton(common, dialect.presentation, "bg", false)
		if err != nil {
			return nil, err
		}
		data := []byte(fmt.Sprintf(`<p:bg xmlns:p="%s" xmlns:a="%s"><p:bgPr><a:solidFill><a:srgbClr val="%s"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`, dialect.presentation, dialect.drawing, *operation.Fill))
		edit := nativeXMLReplacement{start: tree.RawStart, end: tree.RawStart, data: data}
		if bg != nil {
			edit.start = bg.RawStart
			edit.end = bg.RawEnd
		}
		changes[part], err = applyNativeXMLReplacements(pkg.parts[part], []nativeXMLReplacement{edit})
		if err != nil {
			return nil, err
		}
		if bytes.Equal(changes[part], pkg.parts[part]) {
			return nil, fmt.Errorf("pptxpatch: native mutations: background is unchanged")
		}
	} else {
		newPart, err = nativeInsertSlideParts(pkg, *target, changes)
		if err != nil {
			return nil, err
		}
	}
	produced, err := applyNativeSlidePartChanges(orig, pkg, changes)
	if err != nil {
		return nil, err
	}
	options := nativeMutationExtractOptions()
	options.Previous = &before
	after, err := ExtractNativePPTX(produced, options)
	if err != nil {
		return nil, fmt.Errorf("pptxpatch: native mutations: reopen slide edit: %w", err)
	}
	expectedCount := len(before.Slides)
	if operation.Kind == NativePPTXInsertSlide {
		expectedCount++
	}
	if len(after.Slides) != expectedCount {
		return nil, fmt.Errorf("pptxpatch: native mutations: slide count did not round trip")
	}
	for i, slide := range before.Slides {
		j := i
		if operation.Kind == NativePPTXInsertSlide && i > index {
			j++
		}
		if after.Slides[j].ID != slide.ID || after.Slides[j].Source == nil || after.Slides[j].Source.PartName != slide.Source.PartName {
			return nil, fmt.Errorf("pptxpatch: native mutations: slide order/identity changed")
		}
	}
	if operation.Kind == NativePPTXSetSlideBackground {
		if nativeStringValue(after.Slides[index].Background) != *operation.Fill || after.Slides[index].BackgroundGradient != nil {
			return nil, fmt.Errorf("pptxpatch: native mutations: background did not round trip")
		}
	} else if after.Slides[index+1].Source == nil || after.Slides[index+1].Source.PartName != newPart {
		return nil, fmt.Errorf("pptxpatch: native mutations: inserted slide did not round trip")
	}
	return produced, nil
}

func nativeInsertSlideParts(pkg nativeExtractPackage, slide NativeSlide, changes map[string][]byte) (string, error) {
	extractor := nativeExtractor{pkg: pkg, relationshipCache: map[string][]nativeExtractRelationship{}}
	roots, err := extractor.parseRelationships("")
	if err != nil {
		return "", err
	}
	presentationPart := ""
	for _, rel := range roots {
		if rel.Type == relOfficeDocumentTransitional || rel.Type == relOfficeDocumentStrict {
			presentationPart = rel.Part
		}
	}
	root, err := parseNativeXML(pkg.parts[presentationPart], presentationPart)
	if err != nil {
		return "", err
	}
	dialect, err := nativeDialectForPresentation(root.Name)
	if err != nil {
		return "", err
	}
	list, err := nativeSingleton(root, dialect.presentation, "sldIdLst", true)
	if err != nil {
		return "", err
	}
	rels, err := extractor.parseRelationships(presentationPart)
	if err != nil {
		return "", err
	}
	slideRels, _, err := extractor.optionalRelationships(slide.Source.PartName)
	if err != nil {
		return "", err
	}
	var layout *nativeExtractRelationship
	for i := range slideRels {
		if slideRels[i].Type == dialect.relSlideLayout {
			if layout != nil || !slideRels[i].internal() || slideRels[i].Part == "" || !asciiEqualFoldNative(pkg.contentTypes.forPart(slideRels[i].Part), contentTypeSlideLayout) {
				return "", fmt.Errorf("pptxpatch: native mutations: source slide layout is ambiguous or unavailable")
			}
			layout = &slideRels[i]
		}
	}
	// A slide without its required layout cannot authorize creating another
	// PowerPoint slide. Never invent a master or silently change its design.
	if layout == nil {
		return "", fmt.Errorf("pptxpatch: native mutations: source slide has no usable layout")
	}
	newPart := ""
	for n := 1; n <= nativeExtractMaxParts; n++ {
		candidate := path.Join(path.Dir(slide.Source.PartName), fmt.Sprintf("injoffice-slide%d.xml", n))
		alias, err := nativePartAlias(candidate)
		if err != nil {
			return "", err
		}
		relAlias, err := nativePartAlias(nativeRelationshipsPart(candidate))
		if err != nil {
			return "", err
		}
		if pkg.aliases[alias] == "" && !pkg.untypedAliases[alias] && pkg.aliases[relAlias] == "" && !pkg.untypedAliases[relAlias] {
			newPart = candidate
			break
		}
	}
	if newPart == "" {
		return "", fmt.Errorf("pptxpatch: native mutations: no unused slide part")
	}
	used := map[string]bool{}
	sourceRID := ""
	for _, rel := range rels {
		used[rel.ID] = true
		if rel.Type == dialect.relSlide && rel.Part == slide.Source.PartName {
			sourceRID = rel.ID
		}
	}
	newRID := ""
	for n := 1; n <= len(rels)+1; n++ {
		candidate := fmt.Sprintf("rIdInjOffice%d", n)
		if !used[candidate] {
			newRID = candidate
			break
		}
	}
	maxID := uint64(255)
	var anchor *nativeXMLNode
	for _, child := range list.Children {
		id, _ := exactNativeAttr(child, "", "id")
		number, err := strconv.ParseUint(id, 10, 32)
		if err != nil || number < 256 || number >= 2147483648 {
			return "", fmt.Errorf("pptxpatch: native mutations: invalid presentation slide id")
		}
		if number > maxID {
			maxID = number
		}
		rid, _ := exactNativeAttr(child, dialect.rels, "id")
		if rid == sourceRID {
			anchor = child
		}
	}
	if anchor == nil || maxID >= 2147483647 {
		return "", fmt.Errorf("pptxpatch: native mutations: missing slide anchor or exhausted slide ids")
	}
	entry := []byte(fmt.Sprintf(`<p:sldId xmlns:p="%s" xmlns:r="%s" id="%d" r:id="%s"/>`, dialect.presentation, dialect.rels, maxID+1, newRID))
	changes[presentationPart], err = applyNativeXMLReplacements(pkg.parts[presentationPart], []nativeXMLReplacement{{start: anchor.RawEnd, end: anchor.RawEnd, data: entry}})
	if err != nil {
		return "", err
	}
	relPart, err := extractor.actualRelationshipsPart(presentationPart)
	if err != nil {
		return "", err
	}
	changes[relPart], err = nativeAppendXMLChild(pkg.parts[relPart], relPart, fmt.Sprintf(`<Relationship xmlns="%s" Id="%s" Type="%s" Target="%s"/>`, nsPackageRels, newRID, dialect.relSlide, nativeSlidePartURI(newPart)))
	if err != nil {
		return "", err
	}
	changes["[Content_Types].xml"], err = nativeAppendXMLChild(pkg.parts["[Content_Types].xml"], "[Content_Types].xml", fmt.Sprintf(`<Override xmlns="%s" PartName="%s" ContentType="%s"/>`, nsContentTypes, nativeSlidePartURI(newPart), contentTypeSlide))
	if err != nil {
		return "", err
	}
	changes[newPart] = []byte(fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:p="%s" xmlns:a="%s" xmlns:r="%s"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`, dialect.presentation, dialect.drawing, dialect.rels))
	{
		target, err := nativeEscapeXML(layout.Target)
		if err != nil {
			return "", err
		}
		changes[nativeRelationshipsPart(newPart)] = []byte(fmt.Sprintf(`<Relationships xmlns="%s"><Relationship Id="rId1" Type="%s" Target="%s"/></Relationships>`, nsPackageRels, dialect.relSlideLayout, target))
	}
	return newPart, nil
}

func nativeSlidePartURI(part string) string {
	uri := (&url.URL{Path: "/" + part}).EscapedPath()
	escaped, _ := nativeEscapeXML(uri)
	return escaped
}

// Append without serializing the existing XML, including alternate prefixes
// and a self-closing root. Each new fragment declares its own namespaces.
func nativeAppendXMLChild(data []byte, part, child string) ([]byte, error) {
	root, err := parseNativeXML(data, part)
	if err != nil {
		return nil, err
	}
	raw := data[root.RawStart:root.RawEnd]
	var edit nativeXMLReplacement
	if bytes.HasSuffix(raw, []byte("/>")) {
		endName := bytes.IndexAny(raw, " \t\r\n/>")
		if endName < 2 {
			return nil, fmt.Errorf("pptxpatch: native mutations: invalid root tag")
		}
		edit = nativeXMLReplacement{start: root.RawEnd - 2, end: root.RawEnd, data: []byte(">" + child + "</" + string(raw[1:endName]) + ">")}
	} else {
		offset := bytes.LastIndex(raw, []byte("</"))
		if offset < 0 {
			return nil, fmt.Errorf("pptxpatch: native mutations: missing root end tag")
		}
		edit = nativeXMLReplacement{start: root.RawStart + int64(offset), end: root.RawStart + int64(offset), data: []byte(child)}
	}
	return applyNativeXMLReplacements(data, []nativeXMLReplacement{edit})
}

// Raw-copy untouched entries, then verify every original and added part. Only
// this slide-specific writer permits additions; element mutation stays closed.
func applyNativeSlidePartChanges(orig []byte, pkg nativeExtractPackage, changes map[string][]byte) ([]byte, error) {
	zr, err := zip.NewReader(bytes.NewReader(orig), int64(len(orig)))
	if err != nil {
		return nil, err
	}
	var output bytes.Buffer
	zw := zip.NewWriter(&output)
	if err := zw.SetComment(zr.Comment); err != nil {
		return nil, err
	}
	for _, file := range zr.File {
		if _, changed := changes[file.Name]; !changed {
			if err := zw.Copy(file); err != nil {
				return nil, err
			}
		}
	}
	names := make([]string, 0, len(changes))
	for name := range changes {
		names = append(names, name)
	}
	sort.Strings(names)
	for _, name := range names {
		writer, err := zw.CreateHeader(&zip.FileHeader{Name: name, Method: zip.Deflate})
		if err != nil {
			return nil, err
		}
		if _, err := writer.Write(changes[name]); err != nil {
			return nil, err
		}
	}
	if err := zw.Close(); err != nil {
		return nil, err
	}
	result := output.Bytes()
	updated, err := openNativeExtractPackage(result)
	if err != nil {
		return nil, err
	}
	expectedCount := len(pkg.parts)
	for name, data := range changes {
		if _, exists := pkg.parts[name]; !exists {
			expectedCount++
		}
		if !bytes.Equal(updated.parts[name], data) {
			return nil, fmt.Errorf("pptxpatch: native mutations: slide part %q differs", name)
		}
	}
	if len(updated.parts) != expectedCount {
		return nil, fmt.Errorf("pptxpatch: native mutations: unexpected package part count")
	}
	for name, data := range pkg.parts {
		if _, changed := changes[name]; !changed && !bytes.Equal(updated.parts[name], data) {
			return nil, fmt.Errorf("pptxpatch: native mutations: untouched part %q changed", name)
		}
	}
	if err := verifyNativePPTXRawEntries(orig, result, changes); err != nil {
		return nil, err
	}
	return result, nil
}
