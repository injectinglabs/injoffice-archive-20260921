# Reviewed rendering comparisons

Run from `go/officecompat`:

```sh
go run ./cmd/visualcheck -manifest /path/to/reviewed-corpus/manifest.json
```

The command compares externally rendered PNG pages with the existing bounded
`ComparePNG` engine. It prints a JSON report and exits 0 for a match, 1 for any
failed page, or 2 for invalid input. It never captures or updates a reference.
The ordinary `go test ./...` CI job covers its regression tests.

Keep source files and images under the manifest directory. Pin source bytes and
reviewed reference PNG bytes with lowercase SHA-256 digests. Supply explicit
renderer versions and image budgets; zero tolerances require identical pixels.

```json
{
  "version": 1,
  "cases": [{
    "id": "budget-chart",
    "format": "xlsx",
    "source": "sources/budget.xlsx",
    "sourceSha256": "REPLACE_WITH_REVIEWED_SOURCE_SHA256",
    "referenceRenderer": "Excel VERSION / OS / fonts / export settings",
    "candidateRenderer": "InjOffice COMMIT / browser / fonts / scale",
    "limits": {
      "maxEncodedBytes": 16777216,
      "maxWidth": 4096,
      "maxHeight": 4096,
      "maxPixels": 16000000
    },
    "tolerance": {
      "perChannelDelta": 0,
      "maxDifferentPixelsPpm": 0,
      "maxMeanAbsoluteChannelErrorPpm": 0
    },
    "pages": [{
      "reference": "references/budget-1.png",
      "referenceSha256": "REPLACE_WITH_REVIEWED_REFERENCE_SHA256",
      "candidate": "candidates/budget-1.png"
    }]
  }]
}
```

## Capture and review policy

- Use shareable fixtures without customer data. Record ownership/license before
  committing any new real-world document or reference screenshot.
- Cover DOCX paragraphs, lists, tables, images and page breaks; XLSX display
  formats, merged cells, formulas and chart labels; PPTX bullets, theme fonts,
  image crop and positioning; and PDF embedded fonts, rotation and long files.
- Record the exact renderer, OS, fonts, viewport, zoom, pixel ratio, theme, and
  print/export settings. The renderer strings are metadata, not attestations.
- Render reference and candidate from the same pinned source, at identical
  dimensions. Do not resize the candidate to hide pagination or geometry drift.
- Review reference images independently. A screenshot produced by InjOffice is
  a regression baseline, not proof of Microsoft Office compatibility.
- List every expected page/slide/sheet viewport. A missing listed candidate
  fails. This command does not enumerate document pages: the capture job must
  separately assert page count and reject unexpected extra pages.
- Pixel comparison does not verify text semantics, formula results, source
  preservation, links, reading order or edit safety. Keep the existing native
  round-trip and browser interaction tests alongside these comparisons.
- Never increase tolerances automatically. Review layout, missing objects,
  clipped text and changed chart labels before accepting a baseline change.

No externally reviewed Office reference images are bundled with this command.
The CLI and its synthetic tests establish the comparison workflow, not a claim
that the real-world visual corpus has already passed.
