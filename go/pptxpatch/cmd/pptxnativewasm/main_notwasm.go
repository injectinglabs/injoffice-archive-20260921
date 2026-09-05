//go:build !js || !wasm

package main

import "fmt"

func main() {
	fmt.Println("pptxnativewasm must be built with GOOS=js GOARCH=wasm")
}
