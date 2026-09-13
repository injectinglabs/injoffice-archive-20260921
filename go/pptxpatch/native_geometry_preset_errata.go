package pptxpatch

import (
	"encoding/xml"
	"fmt"
	"strings"
)

// Corrections apply only to the fingerprinted embedded catalog, never arbitrary
// source formulas. See docs/PPTX-PRESET-CATALOG.md for original/independent
// evidence. A changed upstream definition must be reviewed rather than patched
// opportunistically by these guards.
func applyNativePresetErrata(name string, node *nativeXMLNode) error {
	switch name {
	case "circularArrow", "leftCircularArrow", "leftRightCircularArrow":
		wanted := map[string]string{"xB": "+- xH 0 dxB 0", "yB": "+- yH 0 dyB 0"}
		if name == "leftRightCircularArrow" {
			wanted["xJ"] = "+- xI 0 dxJ 0"
			wanted["yJ"] = "+- yI 0 dyJ 0"
		}
		found := 0
		for _, guide := range nativeChildren(nativeChild(node, nativePresetDrawingNS, "gdLst"), nativePresetDrawingNS, "gd") {
			key, _ := exactNativeAttr(guide, "", "name")
			expected, applies := wanted[key]
			if !applies {
				continue
			}
			formula, _ := exactNativeAttr(guide, "", "fmla")
			if formula != expected {
				return fmt.Errorf("catalog xB erratum no longer matches")
			}
			for i, attr := range guide.Attrs {
				if attr.Name == (xml.Name{Local: "fmla"}) {
					guide.Attrs[i].Value = strings.TrimSuffix(expected, " 0")
				}
			}
			found++
		}
		if found != len(wanted) {
			return fmt.Errorf("missing catalog xB erratum target")
		}
	case "pie":
		rect := nativeChild(node, nativePresetDrawingNS, "rect")
		original := map[string]string{"l": "il", "t": "ir", "r": "it", "b": "ib"}
		for key, want := range original {
			if value, ok := exactNativeAttr(rect, "", key); !ok || value != want {
				return fmt.Errorf("catalog pie rectangle erratum no longer matches")
			}
		}
		for i, attr := range rect.Attrs {
			if attr.Name == (xml.Name{Local: "t"}) {
				rect.Attrs[i].Value = "it"
			}
			if attr.Name == (xml.Name{Local: "r"}) {
				rect.Attrs[i].Value = "ir"
			}
		}
	}
	return nil
}
