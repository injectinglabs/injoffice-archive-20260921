package pptxpatch

import (
	"encoding/json"
	"reflect"
	"testing"
)

func TestNativeGeometryBindingShapes(t *testing.T) {
	for name, typ := range map[string]reflect.Type{"NativeEvaluatedGeometry": reflect.TypeOf(NativeEvaluatedGeometry{}), "NativeGeometryTextRect": reflect.TypeOf(NativeGeometryTextRect{}), "NativeGeometryPath": reflect.TypeOf(NativeGeometryPath{}), "NativeGeometryCommand": reflect.TypeOf(NativeGeometryCommand{})} {
		if got, want := reflectedBindingShape(typ), nativePPTXBindingShapes[name]; !reflect.DeepEqual(got, want) {
			t.Fatalf("%s drift: %v %v", name, got, want)
		}
	}
}
func TestNativeGeometryRequiredJSONAndUnion(t *testing.T) {
	shape := nativeAutoShapeXMLWithGeometry(3, nativeGeometrySourceXML, `<a:noFill/>`, nativeAutoShapeNoLine("flat", `<a:round/>`), "")
	deck, err := ExtractNativePPTX(nativeShapeStyleFixture(t, false, shape, `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>`), nativeTestExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	data, err := MarshalNativePPTXJSON(deck)
	if err != nil {
		t.Fatal(err)
	}
	for _, key := range []string{"x", "y", "cx", "cy", "stroke", "fillMode", "commands", "kind", "textRect", "paths", "profile", "extra", "editable", "preset"} {
		t.Run(key, func(t *testing.T) {
			var raw map[string]any
			if err := json.Unmarshal(data, &raw); err != nil {
				t.Fatal(err)
			}
			var element map[string]any
			for _, e := range raw["slides"].([]any)[0].(map[string]any)["elements"].([]any) {
				m := e.(map[string]any)
				if m["geometry"] != nil {
					element = m
				}
			}
			g := element["geometry"].(map[string]any)
			path := g["paths"].([]any)[0].(map[string]any)
			switch key {
			case "x", "y", "cx", "cy":
				delete(g["textRect"].(map[string]any), key)
			case "stroke", "fillMode", "commands":
				delete(path, key)
			case "kind":
				delete(path["commands"].([]any)[0].(map[string]any), key)
			case "textRect", "paths", "profile":
				delete(g, key)
			case "extra":
				g["textRect"].(map[string]any)["unknown"] = 1
			case "editable":
				element["compatibility"].(map[string]any)["status"] = "editable"
			case "preset":
				element["preset"] = "rect"
			}
			invalid, err := json.Marshal(raw)
			if err != nil {
				t.Fatal(err)
			}
			if _, err := DecodeNativePPTXJSON(invalid); err == nil {
				t.Fatal("invalid geometry JSON accepted")
			}
		})
	}
}
