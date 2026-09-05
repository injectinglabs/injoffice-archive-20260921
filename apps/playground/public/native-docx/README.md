# Native DOCX playground sample

`northstar-launch-brief.docx.b64` is a deterministic, repository-owned OOXML package used by the Documents playground. It is generated from `go/docxpatch/cmd/nativeplaygroundfixture` and contains an original fictional launch decision memo: styled headings, an executive decision table, priorities, and explicit Letter page geometry. No external text or artwork is included.

The browser decodes it before sending the exact DOCX bytes to the browser-local Go WASM engine by default. An explicitly selected server fallback uses `POST /v1/docx/extract` instead. It supports the guarded text write-back proof, but is not evidence of Word pagination or page-paint fidelity.

Regenerate from the repository root with:

```bash
go run ./go/docxpatch/cmd/nativeplaygroundfixture
```
