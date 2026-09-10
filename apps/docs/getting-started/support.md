# Support and limitations

Treat capability discovery and runtime validation as authoritative. The table below is a navigation aid, not a promise that every file in a format can be edited.

| Workflow | Supported scope | Do not assume |
| --- | --- | --- |
| XLSX native Worker | Cell values/formulas, supported styles, row height, column width | Sheet/row/column insertion, arbitrary structural edits, merge write-back through this Worker adapter |
| DOCX native Worker | Source-anchored complete run or eligible single-run paragraph text replacement | Arbitrary rich-text restructuring, tracked-change editing, Word pagination parity |
| PPTX native Worker | Exact parsed text and AutoShape property changes | Imported slide insertion/removal, general pictures/charts/tables/groups/animation editing |
| Authored presentations | Bounded DeckSpec/WireDeck compilation | Lossless round-trip of an arbitrary imported deck |
| PDF agent adapter | Bounded page operations | General text rewriting, OCR, or redaction through this adapter |
| PDF package | Additional format-specific tools, some Node-only | That every root export works in a browser or that parsing alone proves redaction |
| Collaboration/history | Host-integrated protocols and lifecycle primitives | A hosted identity, storage, sync, or backup service |

## Read support is not write support

The spreadsheet v1 mutation schema recognizes operations that a particular native adapter can still refuse. Native XLSX v2 extraction does not mean there is a v2 write protocol. Similarly, the ability to display a field or image does not grant the authority to rewrite it.

Honor `editable`, refusal codes, unsupported-content inventories, required capabilities, and source identities. Never delete unsupported nodes to make a mutation succeed.

## Original bytes and output checks

Keep the original until your host has accepted the replacement. Reopen output from the exact bytes returned by the writer, not from the preview model. Define checks for the requested changes and for the preservation guarantees your users need.

A successful commit with failed verification is a distinct state: a write happened, but the output is not verified. Withhold the verified download and retain diagnostic evidence. Do not label it an automatic rollback.

## Versions and guarantees

These pages follow `main`. Consult each package's reference, [release status](../reference/generated/contracts/public-release), and your installed `.d.ts` files before adopting an API. Qualification of the repository's bounded test corpus is not evidence of complete Office compatibility.

The public demo is a guided example, not a production document-storage service. Browser-local and optional server readiness are separate; a missing optional server does not mean a browser-only deployment has failed.
