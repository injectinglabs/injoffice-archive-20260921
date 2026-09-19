package pptxpatch

import (
	"testing"
)

// bgGradient wraps stops and an a:lin direction into a p:bg/p:bgPr/a:gradFill.
func bgGradient(d nativeExtractDialect, gradientAttrs map[string]string, direction *nativeXMLNode, stops ...*nativeXMLNode) *nativeXMLNode {
	children := []*nativeXMLNode{bgNode("gsLst", d.drawing, nil, stops...)}
	if direction != nil {
		children = append(children, direction)
	}
	fill := bgNode("gradFill", d.drawing, gradientAttrs, children...)
	return bgNode("bg", d.presentation, nil, bgNode("bgPr", d.presentation, nil, fill))
}

func TestResolveNativeGradientBackground(t *testing.T) {
	d := bgDialect()
	colors := map[string]string{"lt1": "FFFFFF", "dk1": "000000", "dk2": "44546A", "lt2": "E7E6E6"}
	theme := bgTheme(colors, map[string]string{"bg1": "lt1", "bg2": "dk2", "tx1": "dk1", "tx2": "lt2"})
	srgb := func(v string) *nativeXMLNode { return bgNode("srgbClr", d.drawing, map[string]string{"val": v}) }
	stop := func(pos string, color *nativeXMLNode) *nativeXMLNode {
		return bgNode("gs", d.drawing, map[string]string{"pos": pos}, color)
	}
	lin := func(attrs map[string]string) *nativeXMLNode { return bgNode("lin", d.drawing, attrs) }

	t.Run("a linear gradient background is projected with its stops in order", func(t *testing.T) {
		part := bgPart("sld", d, bgGradient(d, map[string]string{"rotWithShape": "1"},
			lin(map[string]string{"ang": "5400000", "scaled": "1"}),
			stop("0", srgb("112233")), stop("100000", srgb("AABBCC"))))
		got := resolveNativeInheritedBackground(nativeSlideDependencyGraph{}, part, d, theme)
		if got.gradient == nil {
			t.Fatalf("gradient background was not projected: %+v", got)
		}
		if got.color != "" {
			t.Fatalf("a gradient background must not also state a flat color: %+v", got)
		}
		if got.gradient.Angle != 5400000 || len(got.gradient.Stops) != 2 {
			t.Fatalf("gradient direction or stop count is wrong: %+v", *got.gradient)
		}
		if got.gradient.Stops[0] != (NativeLinearGradientStop{PositionPct: 0, Color: "112233"}) ||
			got.gradient.Stops[1] != (NativeLinearGradientStop{PositionPct: 100000, Color: "AABBCC"}) {
			t.Fatalf("gradient stops are wrong: %+v", got.gradient.Stops)
		}
	})

	t.Run("a gradient inherits through the layout and master chain", func(t *testing.T) {
		master := bgPart("sldMaster", d, bgGradient(d, nil, lin(map[string]string{"ang": "0"}),
			stop("0", srgb("000000")), stop("100000", srgb("FFFFFF"))))
		got := resolveNativeInheritedBackground(nativeSlideDependencyGraph{masterRoot: master}, bgPart("sld", d, nil), d, theme)
		if got.gradient == nil || got.gradient.Angle != 0 {
			t.Fatalf("master gradient did not cascade: %+v", got)
		}
	})

	// Every one of these changes the painted mapping in a way the linear
	// projection does not represent, so the background must stay unresolved
	// rather than be approximated by a straight interpolation.
	t.Run("gradient forms outside the linear subset are not approximated", func(t *testing.T) {
		for _, tc := range []struct {
			name string
			node *nativeXMLNode
		}{
			{"no direction at all", bgGradient(d, nil, nil, stop("0", srgb("000000")), stop("100000", srgb("FFFFFF")))},
			{"a path shading", bgGradient(d, nil, bgNode("path", d.drawing, map[string]string{"path": "circle"}), stop("0", srgb("000000")), stop("100000", srgb("FFFFFF")))},
			{"a flipped gradient", bgGradient(d, map[string]string{"flip": "x"}, lin(map[string]string{"ang": "0"}), stop("0", srgb("000000")), stop("100000", srgb("FFFFFF")))},
			{"a single stop", bgGradient(d, nil, lin(map[string]string{"ang": "0"}), stop("0", srgb("000000")))},
			{"stops out of order", bgGradient(d, nil, lin(map[string]string{"ang": "0"}), stop("100000", srgb("000000")), stop("0", srgb("FFFFFF")))},
			{"a stop with no position", bgGradient(d, nil, lin(map[string]string{"ang": "0"}), bgNode("gs", d.drawing, nil, srgb("000000")), stop("100000", srgb("FFFFFF")))},
			{"a direction with no angle", bgGradient(d, nil, lin(map[string]string{"scaled": "1"}), stop("0", srgb("000000")), stop("100000", srgb("FFFFFF")))},
		} {
			t.Run(tc.name, func(t *testing.T) {
				got := resolveNativeInheritedBackground(nativeSlideDependencyGraph{}, bgPart("sld", d, tc.node), d, theme)
				if got.gradient != nil {
					t.Fatalf("unmodeled gradient was projected: %+v", *got.gradient)
				}
			})
		}
	})

	// The two stops of the subtitle-animation-save benchmark deck, whose whole
	// page is this one gradient. They exercise tint/hueMod/satMod/lumMod and
	// shade/satMod/lumMod on a bg2 scheme reference.
	t.Run("scheme stops resolve their documented color transforms", func(t *testing.T) {
		scheme := func(children ...*nativeXMLNode) *nativeXMLNode {
			return bgNode("schemeClr", d.drawing, map[string]string{"val": "bg2"}, children...)
		}
		pct := func(name, value string) *nativeXMLNode {
			return bgNode(name, d.drawing, map[string]string{"val": value})
		}
		part := bgPart("sld", d, bgGradient(d, map[string]string{"rotWithShape": "1"},
			lin(map[string]string{"ang": "6120000", "scaled": "1"}),
			stop("10000", scheme(pct("tint", "97000"), pct("hueMod", "92000"), pct("satMod", "169000"), pct("lumMod", "164000"))),
			stop("100000", scheme(pct("shade", "96000"), pct("satMod", "120000"), pct("lumMod", "90000")))))
		got := resolveNativeInheritedBackground(nativeSlideDependencyGraph{}, part, d, theme)
		if got.gradient == nil {
			t.Fatalf("scheme-color gradient was not projected: %+v", got)
		}
		if got.gradient.Stops[0].Color != "75A4B9" || got.gradient.Stops[1].Color != "37485F" {
			t.Fatalf("color transforms resolved to %+v", got.gradient.Stops)
		}
	})
}
