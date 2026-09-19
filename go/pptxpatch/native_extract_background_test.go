package pptxpatch

import (
	"encoding/xml"
	"testing"
)

func bgDialect() nativeExtractDialect {
	return nativeExtractDialect{presentation: nsPresentationTransitional, drawing: nsDrawingTransitional}
}

func bgNode(name string, space string, attrs map[string]string, children ...*nativeXMLNode) *nativeXMLNode {
	node := &nativeXMLNode{Name: xml.Name{Space: space, Local: name}, Children: children}
	for key, value := range attrs {
		node.Attrs = append(node.Attrs, xml.Attr{Name: xml.Name{Local: key}, Value: value})
	}
	return node
}

// A part whose cSld carries the given p:bg.
func bgPart(root string, d nativeExtractDialect, background *nativeXMLNode) *nativeXMLNode {
	children := []*nativeXMLNode{}
	if background != nil {
		children = append(children, background)
	}
	return bgNode(root, d.presentation, nil, bgNode("cSld", d.presentation, nil, children...))
}

func bgSolid(d nativeExtractDialect, color *nativeXMLNode, extra ...*nativeXMLNode) *nativeXMLNode {
	properties := bgNode("bgPr", d.presentation, nil, append([]*nativeXMLNode{bgNode("solidFill", d.drawing, nil, color)}, extra...)...)
	return bgNode("bg", d.presentation, nil, properties)
}

func bgTheme(colors map[string]string, clrMap map[string]string, styles ...*nativeXMLNode) nativeResolvedTheme {
	return nativeResolvedTheme{colors: colors, clrMap: clrMap, bgFillStyles: styles}
}

// ECMA-376 makes p:bg optional at every level: slide, then layout, then master,
// then the implicit default schemeClr bg1. Every scheme reference resolves under
// the SLIDE's effective colour map, not the map of the part it was authored in.
func TestResolveNativeInheritedBackground(t *testing.T) {
	d := bgDialect()
	colors := map[string]string{"lt1": "FFFFFF", "dk1": "000000", "dk2": "009DF0", "lt2": "EEEEEE"}
	srgb := func(v string) *nativeXMLNode { return bgNode("srgbClr", d.drawing, map[string]string{"val": v}) }
	scheme := func(v string) *nativeXMLNode { return bgNode("schemeClr", d.drawing, map[string]string{"val": v}) }
	phClrStyle := bgNode("solidFill", d.drawing, nil, scheme("phClr"))

	for _, tc := range []struct {
		name   string
		slide  *nativeXMLNode
		layout *nativeXMLNode
		master *nativeXMLNode
		theme  nativeResolvedTheme
		want   string
	}{
		{
			name:  "slide srgb background is unchanged",
			slide: bgPart("sld", d, bgSolid(d, srgb("123456"))),
			theme: bgTheme(colors, map[string]string{"bg1": "lt1"}),
			want:  "123456",
		},
		{
			name:  "slide scheme background with bwMode and an empty effectLst",
			slide: bgPart("sld", d, bgNode("bg", d.presentation, nil, bgNode("bgPr", d.presentation, map[string]string{"bwMode": "auto"}, bgNode("solidFill", d.drawing, nil, scheme("bg1")), bgNode("effectLst", d.drawing, nil)))),
			theme: bgTheme(colors, map[string]string{"bg1": "dk1"}),
			want:  "000000",
		},
		{
			name:   "layout background wins over master",
			layout: bgPart("sldLayout", d, bgSolid(d, srgb("0C322C"))),
			master: bgPart("sldMaster", d, bgSolid(d, srgb("FFFFFF"))),
			theme:  bgTheme(colors, map[string]string{"bg1": "lt1"}),
			want:   "0C322C",
		},
		{
			name:   "master scheme background resolves through the effective map",
			master: bgPart("sldMaster", d, bgSolid(d, scheme("bg1"))),
			theme:  bgTheme(colors, map[string]string{"bg1": "dk2"}),
			want:   "009DF0",
		},
		{
			// The p:bg lives on the master, whose own map would send bg1 to white.
			// The slide overrides bg1 to dk1, and PowerPoint paints black.
			name:   "inherited background uses the slide's map, not the authoring part's",
			master: bgPart("sldMaster", d, bgSolid(d, scheme("bg1"))),
			theme:  bgTheme(colors, map[string]string{"bg1": "dk1"}),
			want:   "000000",
		},
		{
			name:   "bgRef dereferences bgFillStyleLst and substitutes its colour for phClr",
			layout: bgPart("sldLayout", d, bgNode("bg", d.presentation, nil, bgNode("bgRef", d.presentation, map[string]string{"idx": "1001"}, scheme("bg2")))),
			theme:  bgTheme(map[string]string{"dk2": "0E2841"}, map[string]string{"bg2": "dk2"}, phClrStyle),
			want:   "0E2841",
		},
		{
			name:  "no p:bg anywhere falls back to the implicit bg1 default",
			theme: bgTheme(colors, map[string]string{"bg1": "dk1"}),
			want:  "000000",
		},
		{
			name:   "bgRef idx 0 is no fill and is refused",
			layout: bgPart("sldLayout", d, bgNode("bg", d.presentation, nil, bgNode("bgRef", d.presentation, map[string]string{"idx": "0"}, scheme("bg1")))),
			theme:  bgTheme(colors, map[string]string{"bg1": "dk1"}, phClrStyle),
			want:   "",
		},
		{
			name:   "an unapproximated gradient background is refused, not inherited past",
			layout: bgPart("sldLayout", d, bgNode("bg", d.presentation, nil, bgNode("bgPr", d.presentation, nil, bgNode("gradFill", d.drawing, nil)))),
			master: bgPart("sldMaster", d, bgSolid(d, srgb("FFFFFF"))),
			theme:  bgTheme(colors, map[string]string{"bg1": "lt1"}),
			want:   "",
		},
		{
			name:   "a bgFillStyleLst slot this tier cannot model is refused by index",
			layout: bgPart("sldLayout", d, bgNode("bg", d.presentation, nil, bgNode("bgRef", d.presentation, map[string]string{"idx": "1001"}, scheme("bg1")))),
			theme:  bgTheme(colors, map[string]string{"bg1": "dk1"}, nil),
			want:   "",
		},
	} {
		t.Run(tc.name, func(t *testing.T) {
			graph := nativeSlideDependencyGraph{layoutRoot: tc.layout, masterRoot: tc.master}
			got := resolveNativeInheritedBackground(graph, tc.slide, d, tc.theme)
			if got.gradient != nil || got.color != tc.want {
				t.Fatalf("background = %+v, want %q", got, tc.want)
			}
		})
	}
}
