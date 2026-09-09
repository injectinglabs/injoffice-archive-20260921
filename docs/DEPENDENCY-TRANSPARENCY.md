# Dependency transparency and offline SBOM

InjOffice must disclose the software it installs and ships. Dependency
independence means that no undeclared or unlicensed implementation is needed; it
does not mean hiding ordinary open-source dependencies.

This audit is engineering evidence, not legal advice. An absent `license` field
is not proof that a package is unlicensed. It is a release blocker until a
reviewer records the package's actual terms or removes the dependency.

## Reproducible checks

The dependency gate reads only checked-in files and does not contact a registry:

```sh
npm run check:dependency-integrity
```

It verifies that every workspace's direct, peer, optional, and development
dependency agrees with `package-lock.json`; inventories every resolved npm
component; binds reviewed license assertions to exact versions and integrity
digests; traces production-reachable dependencies; and audits every `go.mod`.
The reviewed assertions used only when npm omitted lockfile metadata live in
[`provenance/npm-license-evidence.json`](provenance/npm-license-evidence.json).

Generate a deterministic CycloneDX 1.5 JSON SBOM without network access:

```sh
npm run --silent generate:sbom > injoffice.cdx.json
```

Components with no defensible license assertion remain in the SBOM with an
`injoffice:license-status=NOASSERTION` property. Generation succeeds so auditors
can inspect blockers; the release gate fails while a production blocker exists.

## Current inventory and blockers

At the audited lockfile, the workspace contains 28 npm workspaces and 370 resolved
npm components. The seven Go modules have no direct third-party Go module: all
`require` directives point to sibling InjOffice modules through local `replace`
directives.

The earlier npm lockfile resolved 27 `@univerjs-pro/*` packages without license
metadata. They were production-reachable from the direct
`@univerjs/presets@0.25.1` dependency, via its advanced and collaboration presets:

- `@univerjs-pro/collaboration`
- `@univerjs-pro/collaboration-client`
- `@univerjs-pro/collaboration-client-ui`
- `@univerjs-pro/docs-exchange-client`
- `@univerjs-pro/docs-print`
- `@univerjs-pro/edit-history-loader`
- `@univerjs-pro/edit-history-viewer`
- `@univerjs-pro/engine-chart`
- `@univerjs-pro/engine-formula`
- `@univerjs-pro/engine-pivot`
- `@univerjs-pro/engine-shape`
- `@univerjs-pro/exchange-client`
- `@univerjs-pro/license`
- `@univerjs-pro/print`
- `@univerjs-pro/sheets-chart`
- `@univerjs-pro/sheets-chart-ui`
- `@univerjs-pro/sheets-exchange-client`
- `@univerjs-pro/sheets-outline`
- `@univerjs-pro/sheets-outline-ui`
- `@univerjs-pro/sheets-pivot`
- `@univerjs-pro/sheets-pivot-ui`
- `@univerjs-pro/sheets-print`
- `@univerjs-pro/sheets-shape`
- `@univerjs-pro/sheets-shape-ui`
- `@univerjs-pro/sheets-sparkline`
- `@univerjs-pro/sheets-sparkline-ui`
- `@univerjs-pro/thread-comment-datasource`

That issue is remediated. The demo and formula audit now use a minimal bootstrap
built on `@univerjs/core` plus the explicitly declared OSS sheet presets
(`@univerjs/preset-sheets-core`, `@univerjs/preset-sheets-drawing`, and the Node
core preset). Unused `@univerjs/presets` peer declarations were removed from
charts, collaboration, connectors, formulas, pivots, and shapes. The current
manifest and lockfile contain neither `@univerjs/presets` nor any
`@univerjs-pro/*` package.

Twenty-four `@univerjs/*@0.25.1` artifacts used by the OSS Sheets presets,
including telemetry, omit a license field from their package and lock metadata.
Each exact integrity digest is bound in the reviewed evidence file to the
Apache-2.0 license at the upstream `dream-num/univer` v0.25.1 repository tag. A
version or artifact change invalidates that assertion and requires a new review.

## Security advisory snapshot

An earlier `npm audit --omit=dev` run on 2026-09-02 reported one underlying
high-severity advisory, `GHSA-28wg-ghj8-5hjv`, for NanoID 5.1.11 as pinned by
`@univerjs/core@0.25.1`; npm expanded that relationship to 23 affected Univer
packages. A root override now scopes Univer core to the patched NanoID 5.1.16.
The full build and test suite pass with that override. The direct TipTap advisory
reported by the earlier lockfile was also remediated by raising the Docs peer
floor to TipTap 3.31.0. A subsequent online audit reports zero known
vulnerabilities in the repository installation, not necessarily in consumers.

**Consumer mitigation:** npm root overrides are not inherited from installed
packages. A fresh `@injoffice/collab@0.1.0` consumer can therefore resolve the
vulnerable NanoID through its Univer peer. Add this to the consuming application's
root `package.json`, then run `npm install` and `npm audit`:

```json
{
  "overrides": {
    "@univerjs/core": { "nanoid": "5.1.16" }
  }
}
```

This addresses the identified NanoID advisory, not every future dependency issue.
The blank-consumer smoke test uses the same explicit mitigation. Published 0.1.0
tarballs are unchanged. Removing unnecessary editor peers or adopting a qualified
Univer upgrade belongs in a subsequent package release.

The online advisory snapshot is not part of the deterministic license/SBOM gate.
Rerun it before release because registry advisory state changes over time, and
remove the NanoID override when a qualified Univer upgrade no longer needs it.

## Direct dependency audit

The production dependency and peer surface currently comprises the following
external families. Exact ranges, resolved versions, integrity hashes, and all
transitive components come from the generated SBOM rather than this summary.

| Role | Packages | Declared or reviewed license |
| --- | --- | --- |
| Editor shell | `@univerjs/core` and explicit sheet presets | Apache-2.0; telemetry assertion separately integrity-bound |
| UI/runtime | React, React DOM, ECharts, Konva, React Konva | MIT or Apache-2.0 |
| Document editing | TipTap core and ProseMirror facade (`@tiptap/pm`) | MIT |
| PDF | EmbedPDF PDFium, pdf-lib, PDF.js, pngjs | MIT or Apache-2.0 |
| Text and hashing | `@noble/hashes`, bidi-js, harfbuzzjs, require-from-string | MIT |
| Fonts | `dejavu-fonts-ttf` | package-bundled DejaVu font terms (`LicenseRef-DejaVu-Fonts`) |

The development-only direct surface contains Node/React/pngjs type packages,
TypeScript 7.0.2, Vite, its React plugin, Vitest,
canvas, DejaVu test fonts, and the Univer Node sheets preset. These are included
in the SBOM but are not runtime dependencies of published InjOffice packages.

`apps/playground` directly imports `pdfjs-dist`; the audit found that it previously
relied on workspace hoisting through `@injoffice/pdf`. The manifest now declares
that dependency explicitly so a clean workspace install is reproducible.

## Existing license and NOTICE coverage

The root `NOTICE` records the Microsoft-authored fixture provenance, HarfBuzz,
bidi-js, require-from-string, Unicode data, pngjs, and the PDF OCR ownership
statement. It also consolidates the Go BSD and harfbuzzjs, bidi-js, and
require-from-string MIT license texts. The root `LICENSE` contains Apache-2.0,
while `LICENSE-UNICODE.txt` contains the Unicode terms. Published package
manifests identify which legal files their tarballs must carry.

That is not a complete substitute for the SBOM and must not be described as one.
In particular, the repository does not currently retain a standalone license copy
for every MIT/Apache dependency or the DejaVu npm package. Dependencies installed
as separate npm packages normally bring their own license files; a bundled browser
distribution may need a consolidated third-party-notice artifact. Before a public
binary or hosted bundle release, counsel or a designated license reviewer should
check upstream Apache `NOTICE` files, the DejaVu font terms, and every component
marked `NOASSERTION`. Nothing in this audit authorizes deleting existing notices.

## Distribution notes

- Runtime and peer dependencies belong in the SBOM even when the host supplies
  them.
- Development-only compiler, test, and type packages remain inventoried. Missing
  metadata there is visible, but it does not imply those tools ship in an npm
  package.
- Optional platform binaries remain visible; the SBOM marks them optional rather
  than pretending they do not exist on the current platform.
- `NOTICE`, `LICENSE`, `LICENSE-UNICODE.txt`, and package-specific legal files
  remain authoritative distribution inputs. This gate does not authorize
  removing any notice.
