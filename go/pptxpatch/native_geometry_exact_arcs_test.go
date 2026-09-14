package pptxpatch

import (
	"bytes"
	"encoding/json"
	"math/big"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"
)

func TestNativeGeometryExactCardinalChains(t *testing.T) {
	arc := `<a:arcTo wR="10000000000" hR="10000000000" stAng="0" swAng="21600000"/>`
	near := strings.Replace(arc, `stAng="0"`, `stAng="1"`, 1)
	move := `<a:moveTo><a:pt x="20000000000" y="10000000000"/></a:moveTo>`
	for _, tc := range []struct {
		name, attrs, commands string
		pass                  bool
	}{
		{"exact", "", move + strings.Repeat(arc, 60), true},
		{"near-cardinal", "", move + strings.Repeat(near, 60), false},
		{"close-restores-exact-start", "", move + near + `<a:close/>` + strings.Repeat(arc, 60), true},
		{"non-unit", ` w="10000000000"`, move + strings.Repeat(arc, 60), false},
		// Exact explicit moves reestablish the pen but must not erase old error.
		{"prior-error", "", move + strings.Repeat(near+move+arc, 60), false},
		{"command-budget", "", move + strings.Repeat(arc, 128), false},
	} {
		t.Run(tc.name, func(t *testing.T) {
			xml := `<a:custGeom xmlns:a="` + nativeGeometryTestNS + `"><a:pathLst><a:path` + tc.attrs + `>` + tc.commands + `</a:path></a:pathLst></a:custGeom>`
			node, err := parseNativeXML([]byte(xml), "cardinal.xml")
			if err != nil {
				t.Fatal(err)
			}
			_, err = evaluateNativeCustomGeometry(node, nativeGeometryTestNS, 20000000000, 20000000000)
			if (err == nil) != tc.pass {
				t.Fatalf("pass=%v error=%v", tc.pass, err)
			}
		})
	}
}

func TestNativeGeometryExactCardinalWidePresets(t *testing.T) {
	for _, name := range []string{"can", "leftBrace", "rightBrace", "leftBracket", "rightBracket"} {
		t.Run(name, func(t *testing.T) {
			geometry, err := evaluateNativePresetGeometry(name, nil, 4000000, 400000)
			if err != nil {
				t.Fatal(err)
			}
			if len(geometry.Paths) == 0 {
				t.Fatal("missing paths")
			}
		})
	}
	for _, name := range []string{"cornerTabs", "plaqueTabs", "squareTabs"} {
		if _, err := evaluateNativePresetGeometry(name, nil, 4000000, 400000); err == nil {
			t.Fatalf("inverted text rectangle admitted: %s", name)
		}
	}
}

func TestNativeGeometryExactCardinalSourceOracle(t *testing.T) {
	data, err := os.ReadFile("testdata/exact-cardinal-arcs-oracle.json")
	if err != nil {
		t.Fatal(err)
	}
	var oracle struct {
		Cases []struct {
			Name     string                    `json:"name"`
			Paths    [][]NativeGeometryCommand `json:"paths"`
			TextRect NativeGeometryTextRect    `json:"textRect"`
		} `json:"cases"`
	}
	if err := json.Unmarshal(data, &oracle); err != nil {
		t.Fatal(err)
	}
	out := os.Getenv("PPTX_EXACT_CARDINAL_FIXTURES")
	if out != "" {
		if err := os.MkdirAll(out, 0755); err != nil {
			t.Fatal(err)
		}
	}
	emit := func(name, geometry string, admitted bool, expectedPaths [][]NativeGeometryCommand, rect NativeGeometryTextRect) {
		t.Helper()
		for _, strict := range []bool{false, true} {
			shape := nativeAutoShapeXMLWithGeometry(3, geometry, `<a:solidFill><a:srgbClr val="336699"/></a:solidFill>`, nativeAutoShapeNoLine("flat", `<a:round/>`), "")
			shape = strings.Replace(shape, `cx="1000000" cy="500000"`, `cx="4000000" cy="400000"`, 1)
			// A shape-only source fixture avoids introducing an unrelated title font.
			input := nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, mutate: func(parts map[string]string) {
				part := "relocated/slides/slide-a.xml"
				raw := parts[part]
				start := strings.Index(raw, "<p:sp>")
				end := strings.Index(raw[start:], "</p:sp>") + start + len("</p:sp>")
				parts[part] = raw[:start] + shape + raw[end:]
			}})
			before := bytes.Clone(input)
			deck, err := ExtractNativePPTX(input, nativeTestExtractOptions())
			if err != nil {
				t.Fatal(err)
			}
			e := nativeFixtureAutoShapes(deck.Slides[0])[0]
			if admitted {
				if e.Geometry == nil || e.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
					t.Fatalf("%s missing qualified geometry: %+v", name, e.Compatibility)
				}
				if len(e.Geometry.Paths) != len(expectedPaths) || e.Geometry.TextRect != rect {
					t.Fatalf("%s path count/text rectangle differs from oracle", name)
				}
				for i, path := range e.Geometry.Paths {
					if !reflect.DeepEqual(path.Commands, expectedPaths[i]) {
						t.Fatalf("%s path %d differs from independent oracle: %+v", name, i, path.Commands)
					}
				}
			} else if e.Geometry != nil || e.Preset != nil || e.Compatibility.Status != NativeCompatibilityStatusRefused {
				t.Fatalf("%s unexpectedly admitted", name)
			}
			if !bytes.Equal(input, before) {
				t.Fatal("source package changed")
			}
			if issues := ValidateNativePPTX(deck); len(issues) > 0 {
				t.Fatal(issues)
			}
			if out != "" && !strict {
				if err := os.WriteFile(filepath.Join(out, name+".pptx"), input, 0644); err != nil {
					t.Fatal(err)
				}
			}
		}
	}
	for _, c := range oracle.Cases {
		emit(c.Name, `<a:prstGeom prst="`+c.Name+`"><a:avLst/></a:prstGeom>`, true, c.Paths, c.TextRect)
	}
	move := `<a:moveTo><a:pt x="20000000000" y="10000000000"/></a:moveTo>`
	near := `<a:arcTo wR="10000000000" hR="10000000000" stAng="1" swAng="21600000"/>`
	exact := strings.Replace(near, `stAng="1"`, `stAng="0"`, 1)
	for _, c := range []struct{ name, attrs, commands string }{
		{"near-cardinal", "", move + strings.Repeat(near, 60)},
		{"non-unit", ` w="2000000" h="200000"`, move + strings.Repeat(exact, 60)},
		{"prior-error", "", move + strings.Repeat(near+move+exact, 60)},
	} {
		emit(c.name, `<a:custGeom><a:pathLst><a:path`+c.attrs+`>`+c.commands+`</a:path></a:pathLst></a:custGeom>`, false, nil, NativeGeometryTextRect{})
	}
}

func TestNativeGeometryExactCardinalRequiresProvenance(t *testing.T) {
	g, _ := newNativeGeometryGuides(4000000, 400000)
	node, err := parseNativeXML([]byte(`<a:arcTo xmlns:a="`+nativeGeometryTestNS+`" wR="100" hR="50" stAng="angle" swAng="5400000"/>`), "arc.xml")
	if err != nil {
		t.Fatal(err)
	}
	g.values["angle"] = 5400000
	g.exact["angle"] = big.NewRat(5400000, 1)
	p := nativeGeometryPathBuilder{sx: 1, sy: 1, hasPen: true, pen: nativeGeometryPoint{100, 50}}
	values := []float64{100, 50, 5400000, 5400000}
	if !nativeGeometryExactCardinalArc(g, node, p, true, values) {
		t.Fatal("exact cardinal control refused")
	}
	g.exact["angle"] = new(big.Rat).Add(big.NewRat(5400000, 1), big.NewRat(1, 1000000000000))
	if nativeGeometryExactCardinalArc(g, node, p, true, values) {
		t.Fatal("rounded integer-looking angle certified")
	}
	g.exact["angle"] = big.NewRat(5400000, 1)
	if nativeGeometryExactCardinalArc(g, node, p, false, values) {
		t.Fatal("unproven integer-looking pen certified")
	}
	p.sx = 2
	if nativeGeometryExactCardinalArc(g, node, p, true, values) {
		t.Fatal("nonunit scale certified")
	}
	p.sx = 1
	p.pen.x = -nativeGeometryMaxMagnitude
	if nativeGeometryExactCardinalArc(g, node, p, true, values) {
		t.Fatal("unsafe cardinal endpoint certified")
	}
}
