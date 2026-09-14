//go:build js && wasm

// Command xlsxrichsourcewasm exposes only the read-only rich source profile.
package main

import (
	"encoding/json"
	"fmt"
	"syscall/js"

	"github.com/injectinglabs/injoffice/go/xlsxpatch"
)

func main() {
	obj := js.Global().Get("Object").New()
	obj.Set("previewRichSource", js.FuncOf(preview))
	js.Global().Set("xlsxnative", obj)
	if ready := js.Global().Get("xlsxnativeOnReady"); ready.Type() == js.TypeFunction {
		ready.Invoke()
	}
	select {}
}
func failure(message string, fatal bool) js.Value {
	out := js.Global().Get("Object").New()
	out.Set("ok", false)
	out.Set("error", message)
	out.Set("fatal", fatal)
	return out
}
func preview(_ js.Value, args []js.Value) (result any) {
	defer func() {
		if r := recover(); r != nil {
			result = failure(fmt.Sprint(r), true)
		}
	}()
	if len(args) != 1 || args[0].Type() != js.TypeObject || args[0].IsNull() || !args[0].InstanceOf(js.Global().Get("Uint8Array")) {
		return failure("previewRichSource requires one Uint8Array", false)
	}
	v := args[0]
	n := v.Get("byteLength").Float()
	if n < 1 || n > xlsxpatch.NativeXLSXMaxPackageBytes || n != float64(int(n)) {
		return failure("rich source package size exceeds its bound", false)
	}
	data := make([]byte, int(n))
	if js.CopyBytesToGo(data, v) != len(data) {
		return failure("incomplete source copy", true)
	}
	p, err := xlsxpatch.PreviewNativeRichSourceV1(data)
	if err != nil {
		return failure(err.Error(), false)
	}
	encoded, err := json.Marshal(p)
	if err != nil {
		return failure(err.Error(), true)
	}
	out := js.Global().Get("Object").New()
	out.Set("ok", true)
	out.Set("value", string(encoded))
	return out
}
