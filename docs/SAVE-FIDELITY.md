# Save write-back fidelity

InjOffice's save write-back fidelity product is a CI-gated unit suite for
surgical OOXML persistence. It is the InjOffice analogue of GenOffice unit
fidelity tests: parse a smallest synthetic package, apply one mutation, and
prove that unmodeled neighbors and untouched OPC parts survive. Go engines
remain file authority. Writers start from original package bytes and fail
closed: they return an error rather than silently rebuild unsupported content.

This is not a raster benchmark, not an Office-PDF comparison, and not a port of
GenOffice regex tests.

## What this is and is not

| Product | Question it answers | Relation to this suite |
| --- | --- | --- |
| Save write-back fidelity | After one `Apply*`, did unmodeled XML and untouched ZIP parts survive? | This product. |
| Native paint | How does TypeScript preview compile extracted JSON to renderer-neutral commands? | Paint is not a save path. Page-paint, RenderTree, SVG, and screenshots never decide write-back pass/fail. |
| `officecompat` package inventory | What OPC parts exist, and did a declared allowlist contain every mutation? | Inventories fingerprint parts. They do not prove intra-XML neighbor survival. The Native Office completion matrix is a coverage inventory, not this suite. |
| Local Office raster benches | Do PNG pixels match an external Word/Excel/PowerPoint or LibreOffice render? | Those benches are local-only. They are not CI write-back gates and are not Microsoft Office paint parity. |

Do not claim Word, Excel, or PowerPoint GUI or paint parity.

## The extract–mutate–reopen loop

A case is one synthetic snippet and one surgical mutation:

1. Build the smallest valid OOXML package that contains the unmodeled neighbor
   under test (and, when needed, an opaque sibling OPC part).
2. `Extract*` the native JSON (`ExtractNativeDocumentV1`, `ExtractNativePPTX`,
   `ExtractNativeWorkbookV1` / `ExtractNativeWorkbookV2`).
3. Apply one mutation through the format writer below.
4. Reopen the produced bytes with the same `Extract*` (and, for OPC identity,
   `officecompat.Inspect`).
5. Assert lexically: the requested text or cell value changed; unmodeled
   neighbor markup is byte-identical; every ZIP part outside the writer's
   declared mutable set is unchanged.

Lexical XML and part-hash assertions are the gate. Re-extracted native JSON
alone is not enough: the native model may omit the neighbor the case exists to
protect.

## Current CI homes

Do not invent writers. The suite hangs off these existing APIs and runs as
ordinary `go test ./...` in `.github/workflows/test.yml` (`go-modules`):

- `go/docxpatch`: `ApplyNativeTextMutationsV1` — exact text-node splices at a
  paragraph or run anchor, gated by `source.package_sha256` and
  `expected_xml_sha256`.
- `go/pptxpatch`: `ApplyNativePPTXMutations` with `kind: "text.replace"`
  (`NativePPTXReplaceText`) — replaces the shape's `<a:p>` span, then reopens
  through `ExtractNativePPTX`.
- `go/xlsxpatch`: `ApplyCellMutations` — `cell.set_value`, `cell.clear_value`,
  `cell.set_formula`, `cell.clear_formula` on worksheet XML; unrelated parts
  are raw-copied through `Apply`.
- `go/officecompat`: `Inspect` inventories every OPC part by name, size, CRC32,
  and SHA-256; `RequireUntouchedParts` fails closed if any part outside the
  exact mutable allowlist is missing, changed, or newly added.

Existing mutate tests already cover portions of this contract. Dedicated cases
belong in each format module's `native_writeback_fidelity_test.go` (or an
existing mutate test file). Sister PRs add those tests; this document defines
the product they implement.

## What to lock

A passing case must prove all of the following that apply:

- **Unmodeled neighbors survive the splice.**
  - DOCX: `w:pPr` character-unit indents and other unmodeled paragraph
    properties; `w:rPr` wrappers around the spliced `w:t`.
  - PPTX: `a:bodyPr` and `a:lstStyle` that sit beside the replaced paragraph
    list (not inside the regenerated `<a:p>` nodes; see below).
  - XLSX: custom cell attributes and the cell style index (`s=`) plus in-cell
    `extLst` that `ApplyCellMutations` does not own.
- **Untouched OPC parts are a no-op.** Bytes, CRC32, and SHA-256 of every part
  outside the writer's mutable set match the source. `RequireUntouchedParts`
  is the format-neutral form of that check.
- **Unsafe edits refuse closed.** Stale package or target fingerprints,
  unsupported markup, signed packages, overlapping targets, and operations the
  writer cannot prove safe return an error and no output package. They must
  never flatten into a reconstructed document.

## What not to lock yet

PPTX `text.replace` currently regenerates the `<a:p>` list from native JSON
using display `a:srgbClr` and `a:buNone`. A case must not claim that
`a:schemeClr` on runs, or bullets inherited from `a:lstStyle`, survive a
paragraph replace. Neighbor `bodyPr` / `lstStyle` elements can still be locked;
the regenerated paragraphs cannot.

A later writer that splices run text the way `ApplyNativeTextMutationsV1`
splices `w:t` could lock those interiors. That follow-up is a run-level splice,
not a regex port of GenOffice.

## Adding a case

1. Add a test in the format module's `native_writeback_fidelity_test.go`, or
   extend an existing mutate test in the same module.
2. Inline the smallest synthetic OOXML that reproduces the neighbor. Do not
   commit private documents, Office-export PDFs, proprietary fonts, or corpus
   binaries.
3. Apply one mutation. Do not rebuild the package.
4. Assert the neighbor string or attribute still present, the requested value
   changed, and untouched parts identical (`RequireUntouchedParts` or raw
   part-byte equality).
5. For a refusal case, assert the error and that no package is returned.

Keep the test deterministic and free of Word, Excel, PowerPoint, or LibreOffice
as a runtime. Paint, PNG comparators, and the completion matrix are separate
gates.
