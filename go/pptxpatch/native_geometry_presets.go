package pptxpatch

import (
	"bytes"
	"compress/gzip"
	"crypto/sha256"
	_ "embed"
	"encoding/xml"
	"fmt"
	"io"
	"sort"
	"strings"
	"sync"
)

// The unchanged Apache POI resource and its license/NOTICE are retained in
// presetdata. No guide formulas or XML are interpreted by browser code.
//
//go:embed presetdata/preset-shapes.xml.gz
var nativePresetCatalogGzip []byte

const nativePresetCatalogXMLSize = 538970

const nativePresetCatalogSHA256 = "4a762444d8d85876881c02a5b1dedf6f73006fcd8acb7b4e393435615b37c780"
const nativePresetDrawingNS = "http://schemas.openxmlformats.org/drawingml/2006/main"

var nativePresetCatalogOnce sync.Once
var nativePresetCatalog map[string]*nativeXMLNode
var nativePresetCatalogError error

func loadNativePresetCatalog() (map[string]*nativeXMLNode, error) {
	nativePresetCatalogOnce.Do(func() {
		nativePresetCatalogXML, err := decodeNativePresetCatalog(nativePresetCatalogGzip)
		if err != nil {
			nativePresetCatalogError = err
			return
		}
		root, err := parseNativeXML(bytes.TrimPrefix(nativePresetCatalogXML, []byte(`<?xml version="1.0" encoding="utf-8"?>`)), "embedded-preset-catalog.xml")
		if err != nil {
			nativePresetCatalogError = err
			return
		}
		if root.Name != (xml.Name{Local: "presetShapeDefinitons"}) || len(root.Children) != 187 {
			nativePresetCatalogError = fmt.Errorf("invalid preset catalog root")
			return
		}
		catalog := map[string]*nativeXMLNode{}
		for _, definition := range root.Children {
			if definition.Name.Space != "" || catalog[definition.Name.Local] != nil {
				nativePresetCatalogError = fmt.Errorf("duplicate or namespaced preset definition")
				return
			}
			catalog[definition.Name.Local] = definition
		}
		nativePresetCatalog = catalog
	})
	return nativePresetCatalog, nativePresetCatalogError
}

func nativePresetNames() ([]string, error) {
	catalog, err := loadNativePresetCatalog()
	if err != nil {
		return nil, err
	}
	names := make([]string, 0, len(catalog))
	for name := range catalog {
		names = append(names, name)
	}
	sort.Strings(names)
	return names, nil
}

func cloneNativePresetNode(node *nativeXMLNode) *nativeXMLNode {
	copy := *node
	copy.Attrs = append([]xml.Attr{}, node.Attrs...)
	copy.Children = make([]*nativeXMLNode, len(node.Children))
	for i, child := range node.Children {
		copy.Children[i] = cloneNativePresetNode(child)
	}
	return &copy
}

// Alpha-renaming preserves sequential assignment semantics: a formula reads
// previous bindings, and only then installs its result name for later formulas.
// This is a catalog-only normalization; original embedded bytes stay unchanged.
func normalizeNativePresetGuides(node *nativeXMLNode) error {
	aliases := map[string]string{}
	counts := map[string]int{}
	used := map[string]bool{}
	for _, listName := range []string{"avLst", "gdLst"} {
		for _, g := range nativeChildren(nativeChild(node, nativePresetDrawingNS, listName), nativePresetDrawingNS, "gd") {
			name, _ := exactNativeAttr(g, "", "name")
			used[name] = true
		}
	}
	for _, listName := range []string{"avLst", "gdLst"} {
		list := nativeChild(node, nativePresetDrawingNS, listName)
		if list == nil {
			continue
		}
		for _, guide := range list.Children {
			name, _ := exactNativeAttr(guide, "", "name")
			formula, _ := exactNativeAttr(guide, "", "fmla")
			fields := strings.Fields(formula)
			if len(fields) < 2 {
				return fmt.Errorf("invalid catalog guide formula")
			}
			for i := 1; i < len(fields); i++ {
				if alias, ok := aliases[fields[i]]; ok {
					fields[i] = alias
				}
			}
			for i, attr := range guide.Attrs {
				if attr.Name == (xml.Name{Local: "fmla"}) {
					guide.Attrs[i].Value = strings.Join(fields, " ")
				}
			}
			counts[name]++
			renamed := name
			if counts[name] > 1 {
				renamed = fmt.Sprintf("injCatalog_%s_%d", name, counts[name])
				for used[renamed] {
					renamed += "_"
				}
				used[renamed] = true
				for i, attr := range guide.Attrs {
					if attr.Name == (xml.Name{Local: "name"}) {
						guide.Attrs[i].Value = renamed
					}
				}
			}
			aliases[name] = renamed
		}
	}
	var rewrite func(*nativeXMLNode)
	rewrite = func(n *nativeXMLNode) {
		for i, a := range n.Attrs {
			keys := map[string]bool{"x": true, "y": true, "l": true, "t": true, "r": true, "b": true, "wR": true, "hR": true, "stAng": true, "swAng": true, "ang": true, "gdRefX": true, "gdRefY": true, "gdRefR": true, "gdRefAng": true, "minX": true, "maxX": true, "minY": true, "maxY": true, "minR": true, "maxR": true, "minAng": true, "maxAng": true}
			if alias, ok := aliases[a.Value]; ok && keys[a.Name.Local] {
				n.Attrs[i].Value = alias
			}
		}
		for _, child := range n.Children {
			rewrite(child)
		}
	}
	for _, child := range node.Children {
		if child.Name.Local != "avLst" && child.Name.Local != "gdLst" {
			rewrite(child)
		}
	}
	return nil
}

// Adjustment overrides are named integer values, not replacement programs.
// Only names declared by this exact catalog preset can be overridden.
func prepareNativePresetGeometry(name string, adjustments map[string]int64) (*nativeXMLNode, error) {
	catalog, err := loadNativePresetCatalog()
	if err != nil {
		return nil, err
	}
	original, ok := catalog[name]
	if !ok {
		return nil, fmt.Errorf("unknown DrawingML preset %q", name)
	}
	node := cloneNativePresetNode(original)
	node.Name = xml.Name{Space: nativePresetDrawingNS, Local: "custGeom"}
	pending := map[string]int64{}
	for key, value := range adjustments {
		if value > nativeMaxSafeInteger || value < -nativeMaxSafeInteger {
			return nil, fmt.Errorf("preset adjustment exceeds numeric bounds")
		}
		pending[key] = value
	}
	if list := nativeChild(node, nativePresetDrawingNS, "avLst"); list != nil {
		for _, guide := range list.Children {
			key, _ := exactNativeAttr(guide, "", "name")
			value, ok := pending[key]
			if !ok {
				continue
			}
			for i, attr := range guide.Attrs {
				if attr.Name == (xml.Name{Local: "fmla"}) {
					guide.Attrs[i].Value = fmt.Sprintf("val %d", value)
				}
			}
			delete(pending, key)
		}
	}
	if len(pending) > 0 {
		return nil, fmt.Errorf("unknown adjustment for DrawingML preset %q", name)
	}
	if err := applyNativePresetErrata(name, node); err != nil {
		return nil, err
	}
	if err := normalizeNativePresetGuides(node); err != nil {
		return nil, err
	}
	return node, nil
}

// Inflate a single bounded gzip member, then verify the exact retained bytes.
// Compression changes distribution size only, never the catalog's identity.
func decodeNativePresetCatalog(compressed []byte) ([]byte, error) {
	if len(compressed) > 65536 {
		return nil, fmt.Errorf("compressed preset catalog budget exceeded")
	}
	source := bytes.NewReader(compressed)
	reader, err := gzip.NewReader(source)
	if err != nil {
		return nil, fmt.Errorf("invalid compressed preset catalog: %w", err)
	}
	reader.Multistream(false)
	decoded, err := io.ReadAll(io.LimitReader(reader, nativePresetCatalogXMLSize+1))
	closeErr := reader.Close()
	if err != nil || closeErr != nil || len(decoded) != nativePresetCatalogXMLSize || source.Len() != 0 {
		return nil, fmt.Errorf("invalid preset catalog inflate length or checksum")
	}
	if fmt.Sprintf("%x", sha256.Sum256(decoded)) != nativePresetCatalogSHA256 {
		return nil, fmt.Errorf("preset catalog fingerprint mismatch")
	}
	return decoded, nil
}

// nativePresetAdjustmentNames lists a catalog preset's adjustment guides in
// document order, which is the order DrawingML diagram adjust indexes
// (dgm:adj@idx, 1-based) address them in.
func nativePresetAdjustmentNames(name string) ([]string, error) {
	catalog, err := loadNativePresetCatalog()
	if err != nil {
		return nil, err
	}
	preset, ok := catalog[name]
	if !ok {
		return nil, fmt.Errorf("unknown DrawingML preset %q", name)
	}
	list := nativeChild(preset, nativePresetDrawingNS, "avLst")
	if list == nil {
		return nil, nil
	}
	names := make([]string, 0, len(list.Children))
	for _, guide := range list.Children {
		key, _ := exactNativeAttr(guide, "", "name")
		names = append(names, key)
	}
	return names, nil
}
