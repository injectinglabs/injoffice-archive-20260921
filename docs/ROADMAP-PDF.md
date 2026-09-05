# PDF roadmap

The PDF package is a composable byte-editing and viewing library. Browser-safe page operations and viewing are exported from @injoffice/pdf/browser; Node.js PDFium, font-file, OCR, image, and redaction operations are exported from the root package.

## Available

- Document metadata and page geometry.
- Rotate, delete, reorder, insert blank, crop, resize, N-up, split, and merge.
- pdf.js viewing, canvas rendering, text extraction, and search.
- Surgical text matching/editing, block reflow, insertion, font reuse, font rebuilding, style runs, and fallback-font resolution.
- Text markup, drawings, notes and replies, AcroForm values, stamps, visual signatures, and annotation deletion.
- Content-stream image listing, insertion, movement, resize, rotation, replacement, deletion, and verification.
- Caller-provided OCR with page rasterization and an invisible text layer.
- Whole-object text, image, annotation, path, form-XObject, and form-field redaction with proof objects.
- A restricted collaboration-operation codec for review-safe annotation/form operations.

Attribution for third-party dependencies and Unicode-derived data is recorded in the root NOTICE and accompanying third-party license files.

## Safety model

PDF edits are not generally commutative. The library does not claim that arbitrary concurrent byte edits can be merged. A host should apply edits atomically against an expected version and create an explicit conflict when the version changed.

Redaction is admitted only for object classes the selected operation can prove it removed. Unsupported mixed-content or ambiguous cases must fail rather than receive a cosmetic overlay. See [PDF-REDACTION-DESIGN.md](PDF-REDACTION-DESIGN.md).

A visual signature stamp is not a cryptographic digital signature. Certificate trust, incremental signing, revocation, and long-term validation are separate capabilities.

## Current limits

- Advanced editing depends on Node.js, PDFium WASM, HarfBuzz, installed or caller-provided fonts, and bounded native-memory processing.
- OCR requires a caller-provided provider; the package does not send document content to a hosted model.
- Some malformed or encrypted PDFs are outside scope.
- Visual verification remains the host's responsibility for operations where an independent renderer is required.
- Generic collaboration is limited to explicitly decoded annotation/form operations; page, text, image, and structural mutations are rejected by the collaboration codec.

## Near-term work

1. Publish a larger independent fixture corpus for malformed, encrypted, rotated, CJK, and mixed-content files.
2. Add cancellation and resource-budget options consistently across PDFium operations.
3. Expand browser-safe exports where dependencies can operate without Node built-ins.
4. Strengthen independent render verification and fuzz parser boundaries.
5. Document stable 1.0 compatibility guarantees for edit and proof result types.

Authentication, file storage, OCR credentials, version history, and product rollout are host responsibilities.
