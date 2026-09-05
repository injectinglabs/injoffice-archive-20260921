# Office compatibility baseline

`officecompat` is the format-neutral preservation test layer for InjOffice's
native XLSX, PPTX, and DOCX engines. It inventories every OPC part by name,
uncompressed size, CRC32, and SHA-256, and produces a deterministic whole-package
fingerprint that is independent of ZIP order, timestamps, and compression.

It also owns the renderer-neutral `injoffice.office.mutations` v1 envelope and
shared atomic save policy. The envelope carries a full SHA-256 revision of the
exact source bytes and an opaque format-native payload. `ApplyMutation` rejects
a stale revision before calling a handler, requires an exact mutable OPC-part
allowlist, and releases candidate bytes only after the selected DOCX, PPTX, or
XLSX handler reopens and validates its native contract. Unsupported formats,
incomplete policies, no-op candidates, globs, duplicate fields, and unknown
envelope fields fail closed.

An engine test declares the exact parts an operation may mutate and calls
`RequireUntouchedParts`. A missing, changed, or newly-added part outside that
allowlist fails closed. This makes the operation's native Office mutation
boundary explicit and detects silent loss of unknown content.

## Baseline coverage

The initial suite covers:

- deterministic inventories across ZIP order and compression changes;
- malformed, non-OPC, duplicate-part, and escaped-mutation failures;
- an XLSX worksheet mutation through `xlsxpatch`;
- a PPTX text mutation through `pptxpatch.BuildPPTX`;
- a DOCX paragraph mutation through `docxpatch`;
- preservation of opaque `customXml` parts in XLSX and DOCX fixtures.

This first slice verifies package structure and content preservation. It does
not claim complete format coverage, raw compressed-byte identity, or Microsoft
Office/LibreOffice acceptance.

## Visual reference comparison

`ComparePNG` and `RequirePNGMatch` compare externally rendered reference and
candidate page/slide images without making a renderer part of this module. The
comparison is deterministic and integer-only: it composites alpha over white,
classifies pixels with an explicit per-channel threshold, and reports differing
pixels and mean absolute channel error in parts per million. PNG container
metadata and compression do not affect normalized pixel digests.

Callers must provide positive encoded-byte, width, height, and pixel ceilings.
Invalid, truncated, oversized, or dimension-mismatched images fail closed. The
dedicated qualification caller/config owns each explicit tolerance, and the
resulting report records it alongside the measured evidence. This module has no
implicit or recommended tolerance; a zero-value tolerance requires exact visual
pixels. Microsoft Office and LibreOffice may produce reference artifacts in
dedicated qualification jobs, but are never runtime or ordinary-test dependencies
of the comparator.

`VisualCorpusFixtures` consumes the canonical `corpus.Manifest` protocol
directly, validates its pinned generator identity and bounded record metadata,
and selects accepted records in manifest order. `CompareFixturePNG` binds each
comparison report to the exact `corpus.FixtureRecord`; refused fixtures remain
native refusal cases and cannot silently become visual cases.

## Deterministic native corpus

`corpus/specs` is the authority for the bounded DOCX, PPTX, and XLSX native
compatibility corpus. From this module, run:

```sh
go run ./cmd/corpusgen
go run ./cmd/corpusgen --check
```

The generator fixes ZIP ordering, timestamps, permissions, and compression;
emits exact native-contract JSON or explicit refusal expectations; and records
package/expectation hashes, expanded sizes, coverage, and redistributable
provenance in `corpus/manifest.json`. A repeated run must produce identical
bytes. Specs, not binaries, define semantic authority.

The corpus exercises native Go extractors directly. It is not a renderer,
visual-diff harness, browser ZIP/XML reader, or approximate save path. Expected
JSON retains each native contract verbatim inside a shared envelope so format
specific lexical and provenance fields remain observable.

Root-level `npm run check:office-provenance` and `npm run check:office-architecture`
enforce Microsoft-authored fixture digests and production dependency/authority
boundaries, including the forbidden-dependency list. This module does not
treat screenshots as authority and does not claim Microsoft Office/LibreOffice
reopen or performance qualification.

## Bounded structural qualification

`InspectWithLimits` applies explicit package, part-count, expanded-byte,
compression-ratio, and per-part ceilings before it inventories an OPC package.
It rejects encrypted or unsupported ZIP entries, non-RFC OPC segment
characters (apart from the reserved `[Content_Types].xml` root), unsafe URI
paths, and case/percent-equivalent part aliases. `ComparePackages` keeps exact
part bytes authoritative while adding a deterministic, namespace-aware
structural digest for changed XML. UTF-8 BOMs and XML 1.0 literal attribute
whitespace normalization are supported; character-reference whitespace remains
distinct. Malformed XML, directives, undeclared namespaces, and resource limits
produce an explicit lexical-only fallback; they never become a false structural
match. Lexical equality and SHA-256 evidence remain truthful even when a part is
too large for structural comparison.

`QualifyCorpusManifest` consumes the checked-in `corpus.Manifest` and
`FixtureRecord` protocol directly. It validates paths, exact hashes, declared
sizes, part counts, and every XML structural fingerprint without changing specs
or generated artifacts. Package bytes, expanded bytes, part counts, expectation
bytes, XML tokens, and XML attributes are also bounded cumulatively across the
manifest; package and expectation paths must be globally unique. The
deterministic shard selector is:

```sh
go test ./... -run '^(TestStructural|TestRoundTrip|FuzzStructural)'
```

Qualification jobs can set a hard `GOMEMLIMIT` and `go test -timeout` around
that selector. The default library envelope matches the largest current native
Office extractor; smaller shards can pass tighter `Limits` values.

`QualifyOPC` adds deterministic package-native evidence beyond ZIP and XML
well-formedness. It resolves the complete internal relationship graph to exact
package part spellings, retains external targets without dereferencing them,
and records every part's effective content type. Duplicate relationship IDs,
dangling targets, orphan relationship parts, ambiguous content-type
declarations, encoded traversal, and undeclared parts fail closed under the
same cumulative XML limits. The checked-in accepted DOCX, PPTX, and XLSX corpus
is required to pass this graph qualification.
