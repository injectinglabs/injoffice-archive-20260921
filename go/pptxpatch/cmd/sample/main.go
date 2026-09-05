// Command sample writes pptxpatch.ExampleDeck(), built via BuildPPTX, to a
// real .pptx file — the artifact scripts/validate_with_python_pptx.py opens
// with an independent OOXML consumer. Not part of the library; a thin CLI
// so the write-side validation doesn't need a throwaway test harness.
package main

import (
	"fmt"
	"os"

	"github.com/injectinglabs/injoffice/go/pptxpatch"
)

func main() {
	out := "testdata/generated_sample.pptx"
	if len(os.Args) > 1 {
		out = os.Args[1]
	}
	data, err := pptxpatch.BuildPPTX(pptxpatch.ExampleDeck())
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
