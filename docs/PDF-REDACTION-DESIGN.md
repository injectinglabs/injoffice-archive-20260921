# PDF redaction: proof-safe delivery plan

Redaction is destructive security work, not a black rectangle. A result is acceptable only when the selected content is absent from the saved PDF's page objects, text extraction, and rendered pixels.

## First safe mode: whole-object removal

The initial engine mode accepts a page rectangle and removes a PDFium page object only if its full bounds are contained by that rectangle. Any intersecting object is blocked. This intentionally refuses a selection through the middle of a text run, image, vector path, form XObject, clipping path, or annotation; deleting it would remove unrelated content and preserving it would leak selected content.

The admission gate must be pure and explicit. Host/UI work must not claim a redaction succeeded if any requested object is blocked.

## Required destructive transform and verification

1. Snapshot every page object before removal (object type, bounds, text where available).
2. Remove only admitted objects with `FPDFPage_RemoveObject`, regenerate page content, and save a new document atomically.
3. Reload the saved bytes in a fresh PDFium document and prove no admitted object remains in the target rectangle.
4. Extract text through an independent reader (`PdfViewerDocument`/pdf.js) and assert every caller-supplied sensitive token is absent.
5. Rasterize the page and require the target rectangle to be fully covered by an opaque generated redaction mark. This mark is added only after source-object removal succeeds.
6. Fail the entire request (leave source bytes untouched) on any blocked object or failed verification.

Partial-object redaction, nested form XObjects, annotations, and image-pixel redaction are separate modes. They must each have their own removal strategy and extraction/render proof; none may silently fall back to visual masking.

## Collaboration model

PDF byte edits are non-commutative. A safe host model is presence plus versioned file-change reload: a clean follower reloads after a save; a dirty follower retains its staged edits and receives a conflict. The host's atomic save/version history is the authority. Do not submit PDF byte edits into a generic op log or auto-merge generated output. A future operation log may carry only immutable review proposals with an expected version hash; the server applies one proposal atomically only if the hash still matches, otherwise it creates a version/conflict for explicit user resolution.
