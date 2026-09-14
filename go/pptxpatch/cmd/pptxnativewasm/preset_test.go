package main

import (
	"encoding/json"
	"github.com/injectinglabs/injoffice/go/pptxpatch"
	"strings"
	"testing"
)

func TestPresetEvaluationBridge(t *testing.T) {
	names, err := pptxpatch.NativePPTXPresetNames()
	if err != nil {
		t.Fatal(err)
	}
	for _, name := range names {
		t.Run(name, func(t *testing.T) {
			request := presetRequest{name, 4000000, 3000000, map[string]int64{}}
			payload, _ := json.Marshal(request)
			result, err := evaluatePresetJSON(payload)
			if err != nil {
				t.Fatal(err)
			}
			var decoded struct {
				Protocol string
				Request  presetRequest
				Geometry pptxpatch.NativeEvaluatedGeometry
			}
			if err = json.Unmarshal(result, &decoded); err != nil {
				t.Fatal(err)
			}
			if decoded.Protocol != "pptx-preset-evaluation-v1" || decoded.Request.Name != name || len(decoded.Geometry.Paths) == 0 {
				t.Fatalf("invalid result: %s", result)
			}
		})
	}
}
func TestPresetEvaluationBridgeRefusals(t *testing.T) {
	for _, payload := range []string{
		`{}`, `null`, `[]`,
		`{"name":"rect","widthEmu":100,"heightEmu":100,"name":"ellipse"}`,
		`{"name":"rect","widthEmu":100,"heightEmu":100,"unknown":1}`,
		`{"Name":"rect","widthEmu":100,"heightEmu":100}`,
		`{"name":"rect","WidthEmu":100,"heightEmu":100}`,
		`{"name":"triangle","widthEmu":1,"WidthEmu":2,"heightEmu":1}`,
		`{"name":"rect","widthEmu":100,"heightEmu":100,"Adjustments":{}}`,
		`{"name":"rect","widthEmu":null,"heightEmu":100}`,
		`{"name":"rect","widthEmu":1e2,"heightEmu":100}`,
		`{"name":"rect","widthEmu":100,"heightEmu":100,"adjustments":null}`,
		`{"name":"star5","widthEmu":100,"heightEmu":100,"adjustments":{"adj":null}}`,
		`{"name":"star5","widthEmu":100,"heightEmu":100,"adjustments":{"adj":-0}}`,
		`{"name":"star5","widthEmu":100,"heightEmu":100,"adjustments":{"adj":9007199254740992}}`,
		`{"name":"rect","widthEmu":100,"heightEmu":100,"\u006eame":"ellipse"}`,
		`{"name":"star5","widthEmu":100,"heightEmu":100,"adjustments":{"adj":1,"adj":2}}`,
		`{"name":"unknown","widthEmu":100,"heightEmu":100}`,
		`{"name":"rect","widthEmu":9007199254740992,"heightEmu":100}`,
		`{"name":"rect","widthEmu":100,"heightEmu":100} {}`,
		strings.Repeat(" ", 65537),
	} {
		if _, err := evaluatePresetJSON([]byte(payload)); err == nil {
			t.Errorf("accepted %s", payload[:min(len(payload), 200)])
		}
	}
}
