// Command diagram_sample writes pptxpatch.DiagramExampleDeck(), built via
// BuildPPTX, to a real .pptx file — the S12 diagram-slice counterpart to
// cmd/sample (which covers ExampleDeck's non-diagram shapes). Opened by
// scripts/validate_with_python_pptx.py and scripts/validate_diagram_arrows.py
// for independent (python-pptx) validation of the new arrowhead/flipH
// connector fields.
package main

import (
	"fmt"
	"os"

	"github.com/injectinglabs/injoffice/go/pptxpatch"
)

func main() {
	out := "testdata/generated_diagram_sample.pptx"
	if len(os.Args) > 1 {
		out = os.Args[1]
	}
	data, err := pptxpatch.BuildPPTX(pptxpatch.DiagramExampleDeck())
	if err != nil {
		fmt.Fprintln(os.Stderr, "pptxpatch: BuildPPTX:", err)
		os.Exit(1)
	}
	if err := os.WriteFile(out, data, 0o644); err != nil {
		fmt.Fprintln(os.Stderr, "pptxpatch: write:", err)
		os.Exit(1)
	}
	fmt.Println("wrote", out)
}
