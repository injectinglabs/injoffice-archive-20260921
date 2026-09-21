package main

import ("fmt"; "os"; "strings"; "github.com/injectinglabs/injoffice/go/docxpatch")

func text(rs []docxpatch.NativeRunV1) string { var b strings.Builder; for _, r := range rs { if r.Text != nil { b.WriteString(*r.Text) } }; return b.String() }

func main() {
	rw, ro := 0, 0
	data, _ := os.ReadFile(os.Args[1])
	doc, err := docxpatch.ExtractNativeDocumentV1(data)
	if err != nil { fmt.Println("extract error:", err); return }
	for i, b := range doc.Body.Blocks {
		if b.Paragraph == nil { continue }
		_ = i
		t := strings.TrimSpace(text(b.Paragraph.Runs)); if t == "" { continue }
		if b.Paragraph.EditPolicy.Mode == "read-write" { rw++ } else { ro++ }

	}
	fmt.Printf("  editable %d / read-only %d\n", rw, ro)
}
func min(a,b int) int { if a<b { return a }; return b }
