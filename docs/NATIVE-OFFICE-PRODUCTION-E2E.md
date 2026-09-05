# Native Office production E2E contract kit

The v1 kit is the executable, offline handoff between InjOffice and its
production consumers:

- [`testdata/native-office-production-e2e/v1/manifest.json`](../testdata/native-office-production-e2e/v1/manifest.json)
  selects exact DOCX, PPTX, XLSX, and authored-deck bytes and binds SHA-256,
  native protocol/version, engine identity, canonical semantic output,
  preservation/refusal evidence, replay, production entrypoint, and bounded
  resource observations.
- [`schemas/native-office-production-e2e-v1.schema.json`](../schemas/native-office-production-e2e-v1.schema.json)
  publishes the portable record shape.
- [`scripts/verify-native-office-production-e2e.mjs`](../scripts/verify-native-office-production-e2e.mjs)
  is the dependency-free Node 24 verifier and adapter-request producer.

The verifier reads only local files. It makes no network request, reads no
secret, and does not query GitHub. Corpus fixtures remain owned by
`go/officecompat/corpus`; the kit exact-joins those existing paths and manifest
records instead of copying or regenerating their bytes. It hashes the complete
selected corpus records in fixture order, so unrelated corpus additions do not
rewrite a released contract version while any selected-record drift fails.

## Boundary with the completion matrix

The completion matrix owns capability status, coverage dimensions,
authoritative module inventory, test selectors, pending work, and whether a
format is complete. This kit owns none of those claims. It answers a narrower
question: given these exact input bytes, did a host service or browser adapter
produce the required native observation without silent fallback or mutation?

For corpus cases, each fixture also carries the raw expected-file SHA-256 and
the verifier-recomputed canonical digest of its complete `native` or `refusal`
value. A runner that declares the format's native contract protocol must match
that full authoritative output digest in addition to the smaller case-specific
semantic projection. Page-paint and renderer-neutral derived protocols instead
bind their own case projection digest; they cannot mislabel a parser-output
digest as the digest of a different protocol.

The completion matrix binds this kit's verifier, tests, and case manifest as
host-owned `productionE2E` evidence. This v1 manifest still owns none of the
matrix records: capability status, selectors, and pending work remain
matrix-owned. Architecture, dependency, provenance, allocation, and synthetic
work qualification remain owned by their repository checks. The resource
profiles here are transport/harness safety ceilings, not performance claims.

## Authority rules

Only a canonical value derived from validated native protocol output may be
semantic authority. Canvas2D may replay already-authoritative integer paint
commands. The following can never decide pass/fail semantics:

- DOM or HTML layout/content;
- Mammoth;
- legacy `DeckView` or `DeckCanvasView`;
- Konva layout;
- Canvas text measurement, including `measureText`;
- screenshots, pixel diffs, or raster appearance.

Screenshots remain diagnostic failure artifacts only. Production browser
adapters should reuse the website's static/runtime authority probe and add
explicit traps for Canvas measurement and screenshot assertions in the native
entrypoint closure.

## Included v1 cases

The suite has positive DOCX table/font-provider and inline-image page-paint
cases, a parsed PPTX picture case, an authored-deck native compile case, and a
native XLSX sheet case. It also has one atomic native refusal per format and a
fixture-byte tamper case that must fail before an adapter is called. Every
case runs twice and requires identical engine, evidence, and semantic digest.

The DOCX table case requires exact provider identity in the observation, but
does not claim that its package embeds a redistributable font. Current main has
no checked-in font-bearing DOCX package: embedded-font extraction is covered by
Go-generated test bytes and page paint uses locally installed licensed test
fonts. A future version may add a positive embedded-font production case only
after it binds an offline font asset's exact bytes, SHA-256, and license.

## Adapter API

The verifier owns fixture reads and validates length/SHA-256 before returning an
adapter request. Generate one request for a case, runner, and replay index:

```sh
node scripts/verify-native-office-production-e2e.mjs \
  --request docx.image-page-paint application.production-service 0
```

For ordinary cases, the single self-contained JSON line has protocol
`injoffice.native-office-production-e2e-adapter/v1` and contains:

- case, runner, replay index, format, and workflow;
- exact fixture media type, byte length, SHA-256, and base64 bytes;
- exact entrypoint, engine/component identities, native protocol, required
  production signals, and expected semantic/evidence result;
- the case's immutable resource limits.

The tamper case is deliberately different: request creation applies its
manifest-pinned byte mutation and must terminate with
`FIXTURE_DIGEST_MISMATCH` before returning JSON or invoking an adapter. The
controlling harness records that verifier-owned atomic refusal observation.

An adapter resets its production runner for each replay, executes the named
entrypoint, and returns an observation with protocol
`injoffice.native-office-production-e2e-observation/v1`. The observation must
echo exact fixture identity and entrypoint, report exact engine/build identity
and optional component identities, report native protocol/version/outcome,
carry the request's canonical native-output and semantic digests, and normalize
preservation or atomic refusal evidence. It must also report an empty
`forbiddenAuthoritySignals` array and bounded input/output/items/duration.

The verifier recomputes the semantic digest and rejects unknown fields,
duplicate JSON keys, trailing JSON, path escapes, percent aliases, symlinks,
fixture drift, protocol or engine substitution, non-native authority, missing
preservation, partial refusal output, budget excess, incomplete replay, and
replay drift. Adapters never get to waive verifier checks.

Verify a complete observation array from a file or standard input:

```sh
node scripts/verify-native-office-production-e2e.mjs --observations observations.json
node scripts/verify-native-office-production-e2e.mjs --observations - < observations.json
```

For Go, invoke the same dependency-free executable with `os/exec`, decode the
request JSON into a strict struct, run the production HTTP path, and marshal the
strict observation array back to the verifier. Browser-client tests can import
the module functions directly or use the CLI around Playwright.

## Required host wiring

A host service may vendor this manifest, schema, verifier, authored input, and
the listed corpus paths at a pinned InjOffice revision. Its adapter should
exercise the application's authenticated native extraction, DOCX page-paint,
and PPTX asset paths. It should record source before/after SHA-256, response
kind/content digest, exact worker/compiler/runtime and provider identities,
atomic refusal, and the absence of partial success.

Browser clients should replace inline OOXML-authority fixtures in their DOCX,
PPTX, and XLSX browser gates with verifier requests while retaining their
production build and authority probe. They should record hashed production
chunks, exact native/asset requests, zero browser reads of raw source bytes,
the native root or explicit refusal, and the canonical semantic observation.
The authored-deck runner should target the native-only production entrypoint
when the host implements it; until then that case remains downstream wiring
rather than an InjOffice CI success claim.

## How hosts record observations

True `injoffice.native-office-production-e2e-observation/v1` records require a
live host-service or browser adapter. The verifier can emit adapter requests and
verify a complete observation array; it does not invent host observations or
treat screenshots as pass/fail.

Host teams record observations as follows:

1. Generate adapter-request JSON the downstream host must fill:

```sh
node scripts/record-native-office-production-e2e.mjs --write-requests
```

   Requests are generated locally under
   [`testdata/native-office-production-e2e/v1/host-observations/adapter-requests/`](../testdata/native-office-production-e2e/v1/host-observations/adapter-requests/).
   Generated requests and observations are ignored by Git because they can
   contain environment-specific implementation details. The tamper case is not
   a request: request creation must terminate with
   `FIXTURE_DIGEST_MISMATCH` before an adapter runs, and the harness records
   that verifier-owned refusal in `tamper-refusals.json`.

2. The host executes each request against its production entrypoint twice and
   writes the observation array locally to
   `testdata/native-office-production-e2e/v1/host-observations/observations.json`.

3. Verify the host-generated observations without network access:

```sh
node scripts/verify-native-office-production-e2e.mjs \
  --observations testdata/native-office-production-e2e/v1/host-observations/observations.json
```

InjOffice CI does not treat a missing observations file as a production pass.
Host credentials and deployment endpoints must remain in the host environment
and must never be recorded in this repository. Without a host-supplied
`observations.json`, the kit remains a contract and local request generator,
not live host evidence.

Run local contract checks with:

```sh
npm run check:native-office-production-e2e
npm run test:native-office-production-e2e
```

Version by copying the v1 manifest directory and schema to v2 whenever fixture
identity, canonical semantics, adapter/observation shape, engine identity, or
authority policy changes. Never rewrite a released version to follow a moving
host implementation.
