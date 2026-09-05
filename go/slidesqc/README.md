# slidesqc

`slidesqc` is the Go counterpart of `@injoffice/slides` quality checks. It deterministically flags likely overflow and content-density problems without an LLM or renderer.

```bash
go get github.com/injectinglabs/injoffice/go/slidesqc
```

```go
issues := slidesqc.AuditDeck(slidesqc.DeckSpec{
	ID: "review",
	Slides: []slidesqc.SlideSpec{{
		ID: "intro", Kind: slidesqc.KindTitle, Title: "Quarterly review",
	}},
})
fmt.Print(slidesqc.FormatDeckAudit(issues))
```

The package intentionally mirrors only the `DeckSpec` fields used by quality checks. It does not render or modify a presentation.
