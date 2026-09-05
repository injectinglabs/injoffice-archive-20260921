# Native Office completion matrix

The canonical current audit is
[`testdata/native-office-completion/v3/manifest.json`](../testdata/native-office-completion/v3/manifest.json).
The v1, v2, and v3 matrices are pinned to the `origin/main` snapshot
`0cdd7e9340e9469326f57e6175ec87e7e1807443`; their fixture digests bind the
files in that snapshot. This makes the claims reproducible and prevents files
outside the pinned source tree from satisfying a released gate.

The matrix is a completion and coverage inventory, not a renderer. It maps 15
DOCX, PPTX, and XLSX capability groups to native engine modules; supported,
preserved, and refused semantics; fixture provenance and SHA-256; structural,
layout, mutation, consumer, unsupported-loss, and production E2E gates. It also
binds the merged differential/provenance and deterministic resource
qualifications, including their exact limitations. Every
format must explicitly cover native, package-preservation, identity, security,
resource, and visual-production dimensions.

V1 permits exactly one atomic promise in each supported, preserved, and
refused class. Supported and preserved promises are bound to that capability's
structural evidence; refused promises are bound to its unsupported-object-loss
evidence. A later schema version may add multiple promises only with explicit
per-promise evidence references.

All v1 and v2 capability groups remain `partial`. v3 (baseline
`0cdd7e9340e9469326f57e6175ec87e7e1807443`) is the canonical matrix: every
capability is `complete` only because `npm run check:native-office-completion`
agrees. Each v3 capability binds a provenance-bound Microsoft `kind:
office-export` fixture (Macintosh Word numbering, Excel Online `happy-tree.xlsx`,
or Macintosh PowerPoint 16 `attendee-survey-qr.pptx`) plus the host-owned
production E2E kit. This is not Word, Excel, or PowerPoint GUI parity. PNG
comparators and screenshots never become semantic pass/fail. Live kit-v1
host `observations.json` remains a separate production-E2E responsibility and
is not a substitute for this matrix. Host integration behavior is outside the
library's Office-format claims.

## Deterministic validation

`npm run check:native-office-completion` is read-only, offline, and has no
third-party office-suite runtime dependency (the architecture scanner forbids
Electron and Mammoth as native authority). It rejects:

- schema/protocol/base drift, unknown dimension names, unsafe or missing
  repository paths, stale test selectors, duplicate or unordered records;
- capability records without structural and unsupported-object-loss tests, or
  layout/mutation/consumer claims without the corresponding test;
- evidence without a provenance-bound, exact-digest fixture and fixtures that
  are not used by evidence;
- `complete` claims while required gates are missing, including production E2E
  and a provenance-bound real `office-export` fixture;
- production E2E evidence that is not the host-owned production E2E kit;
- missing DOCX/PPTX/XLSX parity for required completion dimensions;
- DOM/HTML, Mammoth, legacy DeckView, or Konva authority in the declared native
  production entrypoints, authoritative modules, and their pinned-baseline
  source dependency closure (legacy public surfaces may still exist, but cannot
  be authority);
- `require`, side-effect import, dynamic import, TypeScript import-equals, and
  interstitial-comment dependency forms after lexical, string/template/regex-
  aware comment normalization;
- dependency paths that escape through lexical traversal, a final symlink, or a
  nested parent symlink, using realpath containment for filesystem audits;
- pending work whose canonical local branch ref, exact head, merge base, target
  ordering, PR number, or 40-character SHA has drifted.

The validator does not query GitHub. Pull-request data is a reviewed snapshot,
not a live CI input. Its adversarial self-tests run with
`npm run test:native-office-completion` and prove the major failure modes.

The JSON Schema at
[`schemas/native-office-completion-v1.schema.json`](../schemas/native-office-completion-v1.schema.json)
documents the portable record shape. The Node validator enforces that shape and
the repository-dependent invariants that JSON Schema cannot express, such as
file digests, selectors, source authority, and cross-format coverage.

## Fixture authority

The matrix references a bounded subset of the existing
`go/officecompat/corpus/manifest.json`; it does not duplicate the corpus
generator or expected-native artifacts. Corpus specs remain the semantic
authority, and `go run ./cmd/corpusgen --check` remains their determinism gate.
Inline unit vectors are also explicit fixtures: the test source is their
generator and its exact baseline bytes are digest-bound. The `kind: office-export`
records identify Microsoft-authored Word numbering evidence and the unmodified
Microsoft Excel Online happy-tree workbook; they do not generalize into a v1
PowerPoint office-export row or production-host coverage. The in-repo
PowerPoint-authored package is provenance for a later matrix version, not a
v1 complete claim.

## Updating the matrix

Do not edit v1 to include files outside its pinned baseline. When native
coverage changes:

1. Keep the historical `origin/main` baseline pin unless a new matrix version
   is copied. Evidence files and selectors must exist in that pinned snapshot.
   After a schema or baseline change, copy the directory/schema/validator
   contract to the next version and validate every retained version plus the new
   canonical version.
2. Add or change fixture provenance first. For corpus artifacts, regenerate and
   check them from `go/officecompat`. For inline or Office-export artifacts, run
   `node scripts/validate-native-office-completion.mjs --print-fixture-digests`
   and review the resulting path/digest rows before patching the manifest.
3. Add exact test selectors and gate owners. A test path without fixture
   provenance and a digest is invalid. A production E2E must name its owning
   host boundary; a unit test is not a substitute.
4. Remove a pending entry only when its implementation and tests are present in
   the new local baseline. Never make validation depend on remote PR state.
5. Run the Node 24 suite, Go native/officecompat suite, corpus generator
   check, package and local-tarball consumer checks, and no-legacy/architecture
   checks applicable to that baseline.

The current pending ledger is empty. v3 upgrades capability status to
`complete` only with office-export plus production-E2E kit evidence the
validator accepts. Local adapter-request and observation handling is documented
in [`NATIVE-OFFICE-PRODUCTION-E2E.md`](NATIVE-OFFICE-PRODUCTION-E2E.md).
Environment-specific requests and observations are ignored by Git and do not
form part of the public matrix.
