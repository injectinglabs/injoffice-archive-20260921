package main

import (
	"flag"
	"fmt"
	"os"

	"github.com/injectinglabs/injoffice/go/officecompat/corpus"
)

func main() {
	check := flag.Bool("check", false, "verify generated corpus outputs without writing")
	root := flag.String("root", "corpus", "corpus root containing specs and generated outputs")
	flag.Parse()
	var err error
	if *check {
		err = corpus.Check(*root)
	} else {
		err = corpus.Generate(*root)
	}
	if err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
}
