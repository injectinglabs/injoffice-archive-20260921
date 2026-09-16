package pptxpatch

import (
	"fmt"
	"strings"
	"unicode/utf8"
)

// nativeShapeIdentityMaxNameBytes bounds the authored shape name copied into a
// diagnostic. Names are source-controlled, so the label is truncated rather
// than allowed to dominate a bounded diagnostic message.
const nativeShapeIdentityMaxNameBytes = 64

// nativeShapeIdentityLabel names one shape-tree element the way the source
// does: its element type plus the p:cNvPr id and name PowerPoint writes on
// every authored shape. A slide-level refusal that carries this label says
// which shape was dropped instead of only that something was.
//
// The label is best-effort: a shape whose nonvisual block is missing or
// unreadable still gets its element type, because the caller is already on a
// refusal path and must not fail again while explaining itself.
func nativeShapeIdentityLabel(node *nativeXMLNode, dialect nativeExtractDialect) string {
	if node == nil {
		return "shape"
	}
	label := node.Name.Local
	if label == "" {
		return "shape"
	}
	label = "p:" + label
	properties := nativeShapeNonVisualProperties(node, dialect)
	if properties == nil {
		return label
	}
	identity := nativeChild(properties, dialect.presentation, "cNvPr")
	if identity == nil {
		return label
	}
	if id, ok := exactNativeAttr(identity, "", "id"); ok && id != "" {
		label += " id=" + nativeShapeIdentityToken(id)
	}
	if name, ok := exactNativeAttr(identity, "", "name"); ok && strings.TrimSpace(name) != "" {
		label += fmt.Sprintf(" name=%q", nativeShapeIdentityToken(name))
	}
	return label
}

// nativeShapeNonVisualProperties returns the p:nv*Pr block of a shape-tree
// element. The element type decides the local name (nvSpPr, nvGrpSpPr,
// nvPicPr, nvCxnSpPr, nvGraphicFramePr), so the lookup is by shape kind rather
// than by scanning for a prefix.
func nativeShapeNonVisualProperties(node *nativeXMLNode, dialect nativeExtractDialect) *nativeXMLNode {
	local := map[string]string{
		"sp":           "nvSpPr",
		"grpSp":        "nvGrpSpPr",
		"pic":          "nvPicPr",
		"cxnSp":        "nvCxnSpPr",
		"graphicFrame": "nvGraphicFramePr",
	}[node.Name.Local]
	if local == "" || node.Name.Space != dialect.presentation {
		return nil
	}
	return nativeChild(node, dialect.presentation, local)
}

// nativeShapeIdentityToken bounds and sanitises one source-controlled token so
// a diagnostic stays single-line and bounded.
func nativeShapeIdentityToken(value string) string {
	value = strings.Map(func(r rune) rune {
		if r == '\n' || r == '\r' || r == '\t' {
			return ' '
		}
		return r
	}, value)
	if len(value) > nativeShapeIdentityMaxNameBytes {
		// Truncate on a rune boundary so the label stays valid UTF-8.
		cut := nativeShapeIdentityMaxNameBytes
		for cut > 0 && !utf8.RuneStart(value[cut]) {
			cut--
		}
		value = value[:cut] + "…"
	}
	return value
}

// nativeDiscloseShapeRefusal prefixes a shape refusal reason with the shape's
// source identity, so "pptx.unsupported-shape" names the shape and the
// construct that refused it rather than only that the slide lost something.
func nativeDiscloseShapeRefusal(node *nativeXMLNode, dialect nativeExtractDialect, reason string) string {
	return nativeShapeIdentityLabel(node, dialect) + ": " + reason
}

// nativeDiagramShapeLabel names a dgm:shape by its declared preset type so a
// diagram refusal says which layout shape refused, not only that one did.
func nativeDiagramShapeLabel(shapeType string) string {
	if strings.TrimSpace(shapeType) == "" {
		return "(untyped)"
	}
	return "type=" + nativeShapeIdentityToken(shapeType)
}

// nativeDiagramAdjustLabel lists the dgm:adj entries a layout shape declares,
// so the refusal names the exact adjust indices and values that are unmodeled.
func nativeDiagramAdjustLabel(adjLst *nativeXMLNode) string {
	if adjLst == nil || len(adjLst.Children) == 0 {
		return "dgm:adjLst"
	}
	parts := make([]string, 0, len(adjLst.Children))
	for _, adj := range adjLst.Children {
		idx, _ := exactNativeAttr(adj, "", "idx")
		val, _ := exactNativeAttr(adj, "", "val")
		parts = append(parts, fmt.Sprintf("dgm:adj idx=%s val=%s", nativeShapeIdentityToken(idx), nativeShapeIdentityToken(val)))
		if len(parts) == nativeDiagramAdjustLabelMax {
			break
		}
	}
	label := strings.Join(parts, ", ")
	if len(adjLst.Children) > nativeDiagramAdjustLabelMax {
		label += fmt.Sprintf(" (+%d more)", len(adjLst.Children)-nativeDiagramAdjustLabelMax)
	}
	return label
}

// nativeDiagramAdjustLabelMax bounds a source-controlled adjust list in a
// bounded diagnostic message.
const nativeDiagramAdjustLabelMax = 4
