# Rendering qualification corpus

Run after `npm ci`, `npm run build`, and a playground renderer build with
`--base=/injoffice-smoke/`:

```sh
node --test scripts/qualify-rendering-corpus.test.mjs
node scripts/qualify-rendering-corpus.mjs
```

The runner executes all six required cases and writes `artifacts/rendering-corpus/report.json`,
logs and screenshots. A failure returns nonzero. Use `--case pdf-geometry` for a
focused check; its report explicitly marks the generated suite incomplete.
Node 22+, Go and Chrome are required. Local servers use isolated ports and
temporary profiles; no running user application or external service is needed.

## What this proves

- PDF: 38,400 exact RGBA pixels from an original two-page rectangle document,
  including `/UserUnit` and rotation. Reference pixels are calculated from the
  authored coordinates, independently of PDF.js. The reference PNG is never
  copied from the candidate. Both are checked by the bounded Go `visualcheck`.
- PDF browser: real viewing, selectable text, editing, continuous navigation
  and bounded raster allocation.
- DOCX: a generated embedded-font file passes the actual extract, shape,
  paginate and SVG-display path; unsupported files visibly refuse native paint.
- PPTX: original quadrant-image cropping has analytical pixel expectations.
  The crop branch uses a native model; the styles branch extracts an original
  PPTX file and checks authored text styles and bullets in the browser.
  These layout checks do not establish PowerPoint typography equivalence.
- Native PPTX: an explicitly uploaded source passes Go extraction, the isolated
  exact-font helper and browser SVG glyph replay, including mixed-size vertical
  anchors and visible missing-font refusal. Operator fonts stay server-local.
- XLSX/DOCX/PPTX: bundled real files pass browser-WASM extraction, editing and
  re-extraction with source-bound readback checks.

Screenshot captures are diagnostics, **not independent Office reference images**.
The manifest and every report explicitly record zero external Office references.
Passing does not prove Word pagination, Excel chart parity, PowerPoint typography,
or general PDF fidelity. Renderer/version dependencies are pinned in `package-lock.json`;
the report also records the candidate commit, manifest and runner hashes.

## Add legitimate Office-export references

`go/officecompat/cmd/visualcheck` accepts version-2 manifests with explicit
`referenceKind: "external-office-export"`, `referenceLicense`, and
`referenceProvenance`. Keep the original DOCX/XLSX/PPTX, exported reference PNGs,
and InjOffice candidate PNGs beneath one manifest directory. Pin source and
reference SHA-256 values. Record the Office app/build, OS, fonts, page or print
settings, export resolution, ownership/license and export date. List every
expected page; the capture procedure must also check for unexpected extra pages.
Do not resize candidates or increase tolerances automatically to make them pass.

```sh
cd go/officecompat
go run ./cmd/visualcheck -manifest /path/to/reviewed-corpus/manifest.json
```

Version 1 remains supported but its evidence is labeled `unclassified`.
Version 2 distinguishes `analytical-oracle`, `self-regression` and
`external-office-export`. Metadata describes provenance; it cannot authenticate
that a human actually exported an image using the named application.
No private/customer material or unlicensed reference assets should be committed.

## Short manual comparison checklist

1. Open your source in Office and InjOffice. Verify page/sheet/slide counts.
2. Compare line wraps, lists, tables, merged cells, currency/date displays,
   charts/labels, image crops, theme fonts and page breaks.
3. Export the original from Office with recorded settings; capture InjOffice
   at matching dimensions and run the comparison above.
4. Make one supported edit, download, and reopen it in Office. Check both the
   intended change and untouched content; keep refused cases explicitly refused.

The generated fixture/oracle definitions are original repository code under
Apache-2.0. The DOCX fixture embeds the installed DejaVu font under its existing
license; generated artifacts remain CI/local diagnostics unless separately reviewed.
