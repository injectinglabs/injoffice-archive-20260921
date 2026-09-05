package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"sort"
	"strconv"
	"strings"
)

// RemovePivot removes exactly one hydrated native pivot and the relationship
// edges that own it. A cache (and its records) is removed only when no other
// native pivot references it. Ambiguous or unsupported graphs fail before a
// Patch is applied.
func RemovePivot(orig []byte, identity PivotIdentity) ([]byte, error) {
	graph, err := resolvePivotLifecycleGraph(orig, identity)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: remove pivot: %w", err)
	}

	patch := Patch{Replace: map[string][]byte{}, Delete: map[string]bool{}}
	updatedSheetRels, err := relationshipDocumentWithout(graph.parts[graph.sheetRelsPart], graph.sheetRelID, relTypePivotTable, graph.sheetPart, graph.tablePart)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: remove pivot: %w", err)
	}
	updatedSheet, err := collectionWithoutReference(graph.parts[graph.sheetPart], "worksheet", "pivotTableParts", "pivotTablePart", "id", graph.sheetRelID)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: remove pivot: %w", err)
	}
	patch.Replace[graph.sheetRelsPart] = updatedSheetRels
	patch.Replace[graph.sheetPart] = updatedSheet
	patch.Delete[graph.tablePart] = true
	patch.Delete[graph.tableRelsPart] = true

	deletedParts := []string{graph.tablePart}
	if !graph.cacheShared {
		updatedWorkbook, err := collectionWithoutReference(graph.parts["xl/workbook.xml"], "workbook", "pivotCaches", "pivotCache", "id", graph.workbookRelID)
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: remove pivot: %w", err)
		}
		updatedWorkbookRels, err := relationshipDocumentWithout(graph.parts["xl/_rels/workbook.xml.rels"], graph.workbookRelID, relTypePivotCacheDef, "xl/workbook.xml", graph.cachePart)
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: remove pivot: %w", err)
		}
		patch.Replace["xl/workbook.xml"] = updatedWorkbook
		patch.Replace["xl/_rels/workbook.xml.rels"] = updatedWorkbookRels
		patch.Delete[graph.cachePart] = true
		deletedParts = append(deletedParts, graph.cachePart)
		if graph.cacheRelsPart != "" {
			patch.Delete[graph.cacheRelsPart] = true
		}
		if graph.recordsPart != "" && !graph.recordsShared {
			patch.Delete[graph.recordsPart] = true
			deletedParts = append(deletedParts, graph.recordsPart)
		}
	}

	updatedContentTypes, err := contentTypesWithoutParts(graph.parts["[Content_Types].xml"], deletedParts)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: remove pivot: %w", err)
	}
	patch.Replace["[Content_Types].xml"] = updatedContentTypes
	return Apply(orig, patch)
}

// UpdatePivot atomically validates a replacement, removes the identified
// native object, and recreates it at the same table-part identity. If its old
// cache was shared, the replacement receives a private cache while the shared
// cache remains byte-for-byte intact for the other pivots.
func UpdatePivot(orig []byte, identity PivotIdentity, spec PivotWriteSpec) ([]byte, error) {
	without, err := RemovePivot(orig, identity)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: update pivot: %w", err)
	}
	updated, err := addPivotAtPart(without, spec, identity.Part)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: update pivot: %w", err)
	}
	pivots, err := ReadPivots(updated)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: update pivot: verify identity: %w", err)
	}
	for _, pivot := range pivots {
		if pivot.Part == identity.Part {
			return updated, nil
		}
	}
	return nil, fmt.Errorf("xlsxpatch: update pivot: replacement did not retain part identity %q", identity.Part)
}

type pivotLifecycleGraph struct {
	parts map[string][]byte

	tablePart     string
	tableRelsPart string
	sheetPart     string
	sheetRelsPart string
	sheetRelID    string
	cachePart     string
	cacheShared   bool
	cacheRelsPart string
	recordsPart   string
	recordsShared bool
	workbookRelID string
}

func resolvePivotLifecycleGraph(data []byte, identity PivotIdentity) (pivotLifecycleGraph, error) {
	graph := pivotLifecycleGraph{}
	if identity.Part == "" || !strings.HasPrefix(identity.Part, "xl/pivotTables/") || !strings.HasSuffix(identity.Part, ".xml") || strings.Contains(identity.Part, "\\") {
		return graph, fmt.Errorf("invalid pivot identity %q", identity.Part)
	}
	if _, err := canonicalOPCPartKey(identity.Part); err != nil {
		return graph, fmt.Errorf("invalid pivot identity %q: %w", identity.Part, err)
	}
	parts, err := readPivotArchive(data)
	if err != nil {
		return graph, err
	}
	graph.parts = parts
	graph.tablePart = identity.Part
	if _, ok := parts[graph.tablePart]; !ok {
		return graph, fmt.Errorf("stale or missing pivot identity %q", graph.tablePart)
	}
	for _, required := range []string{"xl/workbook.xml", "xl/_rels/workbook.xml.rels", "[Content_Types].xml"} {
		if _, ok := parts[required]; !ok {
			return graph, fmt.Errorf("missing required part %q", required)
		}
	}

	read := func(name string) (string, bool) { value, ok := parts[name]; return string(value), ok }
	sheets, err := readWorkbookSheets(read)
	if err != nil {
		return graph, err
	}
	type owner struct{ part, rels, id string }
	var owners []owner
	for _, sheet := range sheets {
		relsPart := relsPartFor(sheet.Part)
		relsData, ok := parts[relsPart]
		if !ok {
			continue
		}
		rels, err := parsePackageRelationships(string(relsData))
		if err != nil {
			return graph, fmt.Errorf("malformed worksheet relationships %q: %w", relsPart, err)
		}
		base := sheet.Part[:strings.LastIndex(sheet.Part, "/")]
		for _, rel := range rels {
			if rel.Type != relTypePivotTable && !strings.HasSuffix(rel.Type, "/pivotTable") {
				continue
			}
			external, err := packageRelationshipIsExternal(rel)
			if err != nil {
				return graph, fmt.Errorf("pivot table relationship %q: %w", rel.ID, err)
			}
			if external {
				continue
			}
			target, err := resolveRelPath(base, rel.Target)
			if err != nil {
				return graph, fmt.Errorf("pivot table relationship %q: %w", rel.ID, err)
			}
			if target == graph.tablePart {
				owners = append(owners, owner{sheet.Part, relsPart, rel.ID})
			}
		}
	}
	if len(owners) != 1 {
		return graph, fmt.Errorf("pivot %q has %d worksheet owners, expected 1", graph.tablePart, len(owners))
	}
	graph.sheetPart, graph.sheetRelsPart, graph.sheetRelID = owners[0].part, owners[0].rels, owners[0].id
	if _, ok := parts[graph.sheetPart]; !ok {
		return graph, fmt.Errorf("missing owning worksheet %q", graph.sheetPart)
	}
	if matches, err := collectionReferenceCount(parts[graph.sheetPart], "worksheet", "pivotTableParts", "pivotTablePart", "id", graph.sheetRelID); err != nil || matches != 1 {
		if err != nil {
			return graph, fmt.Errorf("owning worksheet pivotTableParts: %w", err)
		}
		return graph, fmt.Errorf("owning worksheet has %d pivotTablePart references for %q, expected 1", matches, graph.sheetRelID)
	}

	graph.tableRelsPart = relsPartFor(graph.tablePart)
	tableRels, ok := parts[graph.tableRelsPart]
	if !ok {
		return graph, fmt.Errorf("pivot table relationships %q are missing", graph.tableRelsPart)
	}
	cachePart, err := soleInternalRelationshipTarget(tableRels, graph.tablePart, relTypePivotCacheDef)
	if err != nil {
		return graph, fmt.Errorf("pivot table relationships %q: %w", graph.tableRelsPart, err)
	}
	graph.cachePart = cachePart
	if _, ok := parts[cachePart]; !ok {
		return graph, fmt.Errorf("pivot cache definition %q is missing", cachePart)
	}
	table, err := parsePivotTableDefinition(string(parts[graph.tablePart]))
	if err != nil {
		return graph, err
	}

	cacheUsers := 0
	for part := range parts {
		if !strings.HasPrefix(part, "xl/pivotTables/") || !strings.HasSuffix(part, ".xml") || strings.Contains(part, "/_rels/") {
			continue
		}
		relsPart := relsPartFor(part)
		relsData, ok := parts[relsPart]
		if !ok {
			return graph, fmt.Errorf("pivot table relationships %q are missing", relsPart)
		}
		target, err := soleInternalRelationshipTarget(relsData, part, relTypePivotCacheDef)
		if err != nil {
			return graph, fmt.Errorf("pivot table relationships %q: %w", relsPart, err)
		}
		if target == cachePart {
			other, err := parsePivotTableDefinition(string(parts[part]))
			if err != nil {
				return graph, fmt.Errorf("pivot table %q: %w", part, err)
			}
			if other.CacheID != table.CacheID {
				return graph, fmt.Errorf("pivot tables sharing cache %q disagree on cacheId (%d and %d)", cachePart, table.CacheID, other.CacheID)
			}
			cacheUsers++
		}
	}
	if cacheUsers < 1 {
		return graph, fmt.Errorf("pivot cache %q has no table users", cachePart)
	}
	graph.cacheShared = cacheUsers > 1

	workbookRelID, err := resolveWorkbookPivotCache(parts["xl/workbook.xml"], parts["xl/_rels/workbook.xml.rels"], table.CacheID, cachePart)
	if err != nil {
		return graph, err
	}
	graph.workbookRelID = workbookRelID

	if !graph.cacheShared {
		graph.cacheRelsPart = relsPartFor(cachePart)
		if cacheRels, ok := parts[graph.cacheRelsPart]; ok {
			recordsPart, err := optionalSoleInternalRelationshipTarget(cacheRels, cachePart, relTypePivotRecords)
			if err != nil {
				return graph, fmt.Errorf("pivot cache relationships %q: %w", graph.cacheRelsPart, err)
			}
			graph.recordsPart = recordsPart
			if recordsPart != "" {
				if _, ok := parts[recordsPart]; !ok {
					return graph, fmt.Errorf("pivot cache records %q are missing", recordsPart)
				}
				users, err := pivotRecordsUsers(parts, recordsPart)
				if err != nil {
					return graph, err
				}
				graph.recordsShared = users > 1
			}
		}
	}
	return graph, nil
}

func readPivotArchive(data []byte) (map[string][]byte, error) {
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, fmt.Errorf("open workbook: %w", err)
	}
	parts := make(map[string][]byte, len(zr.File))
	for _, file := range zr.File {
		if _, exists := parts[file.Name]; exists {
			return nil, fmt.Errorf("duplicate package entry %q", file.Name)
		}
		rc, err := file.Open()
		if err != nil {
			return nil, fmt.Errorf("read %q: %w", file.Name, err)
		}
		value, readErr := io.ReadAll(rc)
		closeErr := rc.Close()
		if readErr != nil {
			return nil, fmt.Errorf("read %q: %w", file.Name, readErr)
		}
		if closeErr != nil {
			return nil, fmt.Errorf("close %q: %w", file.Name, closeErr)
		}
		parts[file.Name] = value
	}
	return parts, nil
}

func soleInternalRelationshipTarget(data []byte, ownerPart, expectedType string) (string, error) {
	target, err := optionalSoleInternalRelationshipTarget(data, ownerPart, expectedType)
	if err != nil {
		return "", err
	}
	if target == "" {
		return "", fmt.Errorf("missing %s relationship", expectedType)
	}
	return target, nil
}

func optionalSoleInternalRelationshipTarget(data []byte, ownerPart, expectedType string) (string, error) {
	rels, err := parsePackageRelationships(string(data))
	if err != nil {
		return "", err
	}
	var target string
	base := ownerPart[:strings.LastIndex(ownerPart, "/")]
	for _, rel := range rels {
		if rel.Type != expectedType && !strings.HasSuffix(rel.Type, "/"+strings.TrimPrefix(expectedType, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/")) {
			return "", fmt.Errorf("unsupported relationship type %q", rel.Type)
		}
		if target != "" {
			return "", fmt.Errorf("multiple %s relationships", expectedType)
		}
		external, err := packageRelationshipIsExternal(rel)
		if err != nil || external {
			if err != nil {
				return "", err
			}
			return "", fmt.Errorf("external relationship %q is unsupported", rel.ID)
		}
		target, err = resolveRelPath(base, rel.Target)
		if err != nil {
			return "", err
		}
	}
	return target, nil
}

func pivotRecordsUsers(parts map[string][]byte, recordsPart string) (int, error) {
	users := 0
	for part, value := range parts {
		if !strings.HasPrefix(part, "xl/pivotCache/_rels/pivotCacheDefinition") || !strings.HasSuffix(part, ".rels") {
			continue
		}
		owner := "xl/pivotCache/" + strings.TrimSuffix(strings.TrimPrefix(part, "xl/pivotCache/_rels/"), ".rels")
		target, err := optionalSoleInternalRelationshipTarget(value, owner, relTypePivotRecords)
		if err != nil {
			return 0, fmt.Errorf("pivot cache relationships %q: %w", part, err)
		}
		if target == recordsPart {
			users++
		}
	}
	return users, nil
}

func resolveWorkbookPivotCache(workbook, rels []byte, cacheID int, cachePart string) (string, error) {
	children, _, err := nestedCollectionChildren(workbook, "workbook", "pivotCaches", "pivotCache")
	if err != nil {
		return "", fmt.Errorf("workbook pivotCaches: %w", err)
	}
	relList, err := parsePackageRelationships(string(rels))
	if err != nil {
		return "", fmt.Errorf("workbook relationships: %w", err)
	}
	byID := map[string]packageRelationship{}
	for _, rel := range relList {
		if _, duplicate := byID[rel.ID]; duplicate {
			return "", fmt.Errorf("duplicate workbook relationship id %q", rel.ID)
		}
		byID[rel.ID] = rel
	}
	var matches []string
	candidates := 0
	for _, child := range children {
		value, err := strconv.Atoi(attribute(child.start, "cacheId"))
		if err != nil || value != cacheID {
			continue
		}
		candidates++
		relID := attribute(child.start, "id")
		rel, ok := byID[relID]
		if !ok {
			return "", fmt.Errorf("pivot cache %d references missing workbook relationship %q", cacheID, relID)
		}
		if rel.Type != relTypePivotCacheDef && !strings.HasSuffix(rel.Type, "/pivotCacheDefinition") {
			return "", fmt.Errorf("pivot cache %d relationship %q has type %q", cacheID, relID, rel.Type)
		}
		external, err := packageRelationshipIsExternal(rel)
		if err != nil || external {
			return "", fmt.Errorf("pivot cache %d relationship %q is not internal", cacheID, relID)
		}
		target, err := resolveRelPath("xl", rel.Target)
		if err != nil {
			return "", err
		}
		if target == cachePart {
			matches = append(matches, relID)
		}
	}
	if candidates != 1 {
		return "", fmt.Errorf("pivot cache id %d has %d workbook entries, expected 1", cacheID, candidates)
	}
	if len(matches) != 1 {
		return "", fmt.Errorf("pivot cache %d has %d workbook entries for %q, expected 1", cacheID, len(matches), cachePart)
	}
	return matches[0], nil
}

type nestedChild struct {
	span  xmlSpan
	start xml.StartElement
}

func nestedCollectionChildren(data []byte, rootLocal, collectionLocal, childLocal string) ([]nestedChild, xmlSpan, error) {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	depth := 0
	rootNS := ""
	collection := xmlSpan{}
	var children []nestedChild
	var active *nestedChild
	for {
		before := int(decoder.InputOffset())
		token, err := decoder.Token()
		after := int(decoder.InputOffset())
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, xmlSpan{}, err
		}
		switch token := token.(type) {
		case xml.StartElement:
			depth++
			if depth == 1 {
				if token.Name.Local != rootLocal {
					return nil, xmlSpan{}, fmt.Errorf("root element is %q, expected %q", token.Name.Local, rootLocal)
				}
				rootNS = token.Name.Space
			} else if depth == 2 && token.Name.Space == rootNS && token.Name.Local == collectionLocal {
				if collection.start != 0 || collection.startTagEnd != 0 {
					return nil, xmlSpan{}, fmt.Errorf("multiple %s collections", collectionLocal)
				}
				collection = xmlSpan{qname: rawQName(data, before, after), start: before, startTagEnd: after}
			} else if depth == 3 && collection.startTagEnd != 0 && token.Name.Space == rootNS && token.Name.Local == childLocal {
				active = &nestedChild{span: xmlSpan{qname: rawQName(data, before, after), start: before, startTagEnd: after}, start: token}
			}
		case xml.EndElement:
			if active != nil && depth == 3 && token.Name.Space == rootNS && token.Name.Local == childLocal {
				active.span.endStart, active.span.end = before, after
				children = append(children, *active)
				active = nil
			}
			if depth == 2 && token.Name.Space == rootNS && token.Name.Local == collectionLocal {
				collection.endStart, collection.end = before, after
			}
			depth--
		}
	}
	if collection.end == 0 {
		return nil, xmlSpan{}, fmt.Errorf("missing complete %s collection", collectionLocal)
	}
	return children, collection, nil
}

func collectionReferenceCount(data []byte, root, collection, child, attr, value string) (int, error) {
	children, _, err := nestedCollectionChildren(data, root, collection, child)
	if err != nil {
		return 0, err
	}
	count := 0
	for _, item := range children {
		if attribute(item.start, attr) == value {
			count++
		}
	}
	return count, nil
}

func collectionWithoutReference(data []byte, root, collection, child, attr, value string) ([]byte, error) {
	children, collectionSpan, err := nestedCollectionChildren(data, root, collection, child)
	if err != nil {
		return nil, err
	}
	var matches []xmlSpan
	for _, item := range children {
		if attribute(item.start, attr) == value {
			matches = append(matches, item.span)
		}
	}
	if len(matches) != 1 {
		return nil, fmt.Errorf("%s has %d references for %q, expected 1", collection, len(matches), value)
	}
	if len(children) == 1 {
		return removeElementSpans(data, []xmlSpan{collectionSpan}), nil
	}
	updated := removeElementSpans(data, matches)
	_, newCollection, err := nestedCollectionChildren(updated, root, collection, child)
	if err != nil {
		return nil, err
	}
	startTag := updated[newCollection.start:newCollection.startTagEnd]
	count := strconv.Itoa(len(children) - 1)
	rewritten, err := rewriteUnqualifiedAttribute(startTag, "count", &count)
	if err != nil {
		return nil, err
	}
	return append(bytes.Clone(updated[:newCollection.start]), append(rewritten, updated[newCollection.startTagEnd:]...)...), nil
}

func relationshipDocumentWithout(data []byte, id, expectedType, ownerPart, expectedTarget string) ([]byte, error) {
	rels, _, err := directChildElements(data, "Relationships", "Relationship")
	if err != nil {
		return nil, fmt.Errorf("parse relationships: %w", err)
	}
	var matches []xmlSpan
	for _, rel := range rels {
		if attribute(rel.start, "Id") != id {
			continue
		}
		if typ := attribute(rel.start, "Type"); typ != expectedType && !strings.HasSuffix(typ, "/"+strings.TrimPrefix(expectedType, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/")) {
			return nil, fmt.Errorf("relationship %q has unexpected type %q", id, typ)
		}
		target := attribute(rel.start, "Target")
		if target == "" {
			return nil, fmt.Errorf("relationship %q has no target", id)
		}
		external, err := packageRelationshipIsExternal(packageRelationship{ID: id, TargetMode: attribute(rel.start, "TargetMode")})
		if err != nil || external {
			return nil, fmt.Errorf("relationship %q is not internal", id)
		}
		base := ownerPart[:strings.LastIndex(ownerPart, "/")]
		resolved, err := resolveRelPath(base, target)
		if err != nil {
			return nil, fmt.Errorf("relationship %q target: %w", id, err)
		}
		if resolved != expectedTarget {
			return nil, fmt.Errorf("relationship %q targets %q, expected %q", id, resolved, expectedTarget)
		}
		matches = append(matches, rel.span)
	}
	if len(matches) != 1 {
		return nil, fmt.Errorf("relationship %q has %d entries, expected 1", id, len(matches))
	}
	return removeElementSpans(data, matches), nil
}

func contentTypesWithoutParts(data []byte, parts []string) ([]byte, error) {
	overrides, _, err := directChildElements(data, "Types", "Override")
	if err != nil {
		return nil, fmt.Errorf("parse content types: %w", err)
	}
	wanted := map[string]string{}
	for _, part := range parts {
		key, err := canonicalOPCPartKey(part)
		if err != nil {
			return nil, err
		}
		wanted[key] = part
	}
	matches := map[string][]xmlSpan{}
	for _, override := range overrides {
		name := attribute(override.start, "PartName")
		key, err := canonicalOPCPartKey(name)
		if err != nil {
			return nil, fmt.Errorf("invalid content-type PartName %q: %w", name, err)
		}
		if _, ok := wanted[key]; ok {
			matches[key] = append(matches[key], override.span)
		}
	}
	var spans []xmlSpan
	for key, part := range wanted {
		if len(matches[key]) != 1 {
			return nil, fmt.Errorf("part %q has %d content-type overrides, expected 1", part, len(matches[key]))
		}
		spans = append(spans, matches[key][0])
	}
	sort.Slice(spans, func(i, j int) bool { return spans[i].start > spans[j].start })
	return removeElementSpans(data, spans), nil
}
