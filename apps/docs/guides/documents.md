# Documents / DOCX

Use `@injoffice/docs/native-docx` for browser-safe contracts and `@injoffice/docx-wasm` for guarded native extraction and text writes. Retain the original DOCX archive: the projection is not a replacement serialization format.

## Bridge an editor transaction

The public adapter consumes a complete native-aware transaction projection. It does not accept an arbitrary TipTap/ProseMirror transaction object or HTML string. Your host projector must retain the exact native IDs, source fingerprints, run properties, and before/after text required by the contract.

<<< @/examples/docx.ts

The `project` callback is the integration seam you implement for your editor. Use its inferred `Transaction` type and the [transaction contract](../reference/generated/packages/docs#native-prosemirror-transaction-adapter-v1) to build the complete projection. The example is typechecked, but intentionally does not pretend a generic editor projector can be inferred from a file.

## Source anchors

The source package SHA-256 binds the original archive. Each target also carries a run or eligible paragraph ID and XML fingerprint. Keep these with the editable projection; do not find the save target again by searching text alone.

The current transaction adapter accepts a closed text replacement wholly within one native run and coalesces supported changes into guarded mutations. The direct writer also has a narrow paragraph selector for paragraphs containing exactly one text run.

## Refusals are part of the integration

Structural/open slices, mark changes, ambiguous or overlapping targets, missing coverage, stale anchors, and unsupported markup are refused without an envelope. If the adapter refuses a transaction, keep the original file and explain the limitation. Do not fall back to rebuilding the archive from editor HTML.

## Layout is a different concern

The root Docs package includes shaping and page-paint contracts with explicit font/layout providers. Browser extract/apply clients should use the narrow `/native-docx` entry to avoid Node-only shaping imports.

Successful text readback does not prove Word pagination, font matching, or full-page visual fidelity. Treat preview compiler diagnostics separately from native write validation.

## References

- [DOCX Worker assets, limits, and apply behavior](../reference/generated/packages/docx-wasm)
- [DOCX contract and source identities](../reference/generated/contracts/docx-native-contract)
- [Go document engine](../reference/generated/go/docxpatch)
