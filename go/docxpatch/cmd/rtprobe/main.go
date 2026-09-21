package main

import (
	"archive/zip"
	"bytes"
	"io"
	"fmt"
	"os"
	"regexp"
	"strings"

	"github.com/injectinglabs/injoffice/go/docxpatch"
)

func text(rs []docxpatch.NativeRunV1) string {
	var b strings.Builder
	for _, r := range rs { if r.Text != nil { b.WriteString(*r.Text) } }
	return b.String()
}

func main() {
	data, _ := os.ReadFile(os.Args[1])
	doc, err := docxpatch.ExtractNativeDocumentV1(data)
	if err != nil { fmt.Println("extract:", err); return }
	var target *docxpatch.NativeParagraphV1
	for i := range doc.Body.Blocks {
		p := doc.Body.Blocks[i].Paragraph
		if p != nil && strings.Contains(text(p.Runs), "Invoice:") { target = p; break }
	}
	if target == nil { fmt.Println("paragraph not found"); return }
	fmt.Printf("  target policy: %s  ops=%v\n", target.EditPolicy.Mode, target.EditPolicy.AllowedOperations)
	// replace the text of the run that holds the invoice number
	var run *docxpatch.NativeRunV1
	for i := range target.Runs {
		r := &target.Runs[i]
		if r.Text != nil && strings.Contains(*r.Text, "94") { run = r }
	}
	if run == nil { fmt.Println("no run with the invoice number"); return }
	res, err := docxpatch.ApplyNativeTextMutationsV1(data, doc.Source.PackageSHA256, []docxpatch.NativeDOCXTextMutationV1{{
		TargetKind: "run", TargetID: run.ID, ExpectedXMLSHA256: run.Anchor.XMLSHA256, Text: "77",
	}})
	if err != nil { fmt.Println("  APPLY REFUSED:", err); return }
	fmt.Println("  apply ok, bytes:", len(res.Package))
	before := readPart(data); after := readPart(res.Package)
	fmt.Println("  textFill preserved:", strings.Count(before, "textFill") == strings.Count(after, "textFill"), strings.Count(after, "textFill"))
	fmt.Println("  w:hint preserved:  ", strings.Count(before, `w:hint`) == strings.Count(after, `w:hint`), strings.Count(after, "w:hint"))
	fmt.Println("  kern preserved:    ", strings.Count(before, "<w:kern") == strings.Count(after, "<w:kern"), strings.Count(after, "<w:kern"))
	fmt.Println("  shd preserved:     ", strings.Count(before, "<w:shd") == strings.Count(after, "<w:shd"), strings.Count(after, "<w:shd"))
	// what changed, apart from the edited text?
	re := regexp.MustCompile(`>[^<>]*<`)
	fmt.Println("  markup identical except text nodes:", re.ReplaceAllString(before, "><") == re.ReplaceAllString(after, "><"))
	d2, err := docxpatch.ExtractNativeDocumentV1(res.Package)
	if err != nil { fmt.Println("  RE-EXTRACT FAILED:", err); return }
	for i := range d2.Body.Blocks {
		p := d2.Body.Blocks[i].Paragraph
		if p != nil && strings.Contains(text(p.Runs), "Invoice:") { fmt.Printf("  re-extracted text: %q\n", text(p.Runs)) }
	}
}

func readPart(pkg []byte) string {
	zr, err := zip.NewReader(bytes.NewReader(pkg), int64(len(pkg)))
	if err != nil { return "" }
	for _, f := range zr.File {
		if f.Name != "word/document.xml" { continue }
		rc, err := f.Open(); if err != nil { return "" }
		defer rc.Close()
		b, _ := io.ReadAll(rc)
		return string(b)
	}
	return ""
}

