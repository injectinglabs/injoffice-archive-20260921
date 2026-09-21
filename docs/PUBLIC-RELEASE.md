# Public-release checklist

InjOffice is designed as independently consumable TypeScript packages and Go modules. A hosted backend is optional. This checklist separates public-repository readiness from work owned by an embedding product.

## Completed in the repository

- All 26 TypeScript packages emit ESM and declarations; WASM packages additionally ship their declared runtime assets.
- Every package has exports, a files allowlist, public publish metadata, repository links, and a package README with an example.
- The package check imports built entries and inspects every npm pack payload.
- A blank-consumer smoke test installs the generated tarballs and imports every package.
- The PDF package has a browser-safe subpath and packages required attribution/license material.
- The seven Go modules have public module paths, README examples, and independent tests.
- Node 24 and all Go modules run in public CI without private secrets.
- A tag-driven npm release workflow validates versions, license metadata, tests, builds, package contents, and dependency order before handing SHA-512-bound tarballs to a separate protected OIDC publishing job.
- The Sheets playground has a GitHub Pages workflow.
- Contribution, conduct, security, collaboration-protocol, roadmap, NOTICE, Apache, and Unicode license documents exist.
- The project license is Apache-2.0 and every package ships its SPDX metadata and license text.
- XLSX pivot hydration and fail-closed inventory preservation are implemented in the canonical library.
- Native extract/apply contracts exist for XLSX, DOCX, and PPTX. Native completion remains partial and is not Microsoft Office parity.

Workspace apps under `apps/` are public source but excluded from npm publication
with `"private": true`. The in-repo `injoffice-server` is optional. Browser WASM packages can replace it for supported local workflows and must not become required by other packages.

## Charter boundaries

- **Native file API.** Go `Extract*` → native JSON (`@injoffice/sheets`, `@injoffice/docs`, `@injoffice/pptx-native`) → TypeScript paint compilers (preview) → mutation JSON → Go `Apply*`.
- **Editor shell.** Univer is an optional editor shell, not the file authority. Native paint is preview mode.
- **Optional runtime.** Hosts may inject their own backend, use a published WASM package, or run offline. The in-repo `injoffice-server` and WASM runtimes must not become required by format-model packages.
- **Product isolation.** The public repository must not contain consumer-specific gateways, dashboards, authentication configuration, tenant data, or credentials.
- **No Office parity claim.** Native coverage is a completion matrix, not an unrestricted round-trip guarantee.

## Release status and remaining account setup

Registry verification on 2026-09-21 confirms all 26 packages are published at
0.1.0 under the controlled `@injoffice` scope, including `@injoffice/xlsx-wasm`.
Existing npm versions are immutable. Source fixes after that release, including
the collaboration peer removal, are included in the `0.1.1` release source.
All 26 `0.1.1-rc.0` packages were verified against the CI artifact SHA-512 hashes
and passed the external consumer checks on 2026-09-21. Stable `0.1.1` publication
is pending; changing source versions alone does not publish them.

- Preserve the attribution in NOTICE, LICENSE, LICENSE-UNICODE.txt, and package-specific legal assets.
- Apply the [consumer dependency mitigation](DEPENDENCY-TRANSPARENCY.md#security-advisory-snapshot) for Univer 0.25.1. A clean root audit alone does not establish a clean consumer install.
- Configure and verify direct trusted publishing for every package after bootstrap; disallow traditional publishing tokens once OIDC publishing is verified.
- Protect the `npm` GitHub environment with owner approval. A solo maintainer may approve their own deployment. This is the one approval per release; direct trusted publishing does not create per-package staging approvals. Keep npm account 2FA enabled.
- Private vulnerability reporting is enabled for this public repository. Keep the reporting link in SECURITY.md working.
- Create and test Go submodule release tags such as `go/xlsxpatch/v0.1.0` when releasing the Go modules; npm publication does not release them.

Repository visibility, npm publication, release tags, and Pages deployment remain explicit owner actions. Merging readiness work does not perform them.

## npm publishing bootstrap and steady state

npm requires a package to exist before either staged publishing or a trusted-publisher relationship can be configured. The one-time bootstrap therefore starts with a manual `Release npm packages` workflow run for `v0.1.0`. That run validates the repository and uploads a one-day `npm-release-*` artifact without publishing anything. Download and inspect that artifact, then publish those exact CI-built tarballs through an explicit local owner action:

```sh
gh run download RUN_ID --name npm-release-COMMIT_SHA --dir release-tarballs
npm run release:packages -- v0.1.0 --bootstrap-from release-tarballs
```

The bootstrap mode validates every recorded SHA-512 digest, refuses to run when `CI` is set, and publishes in dependency order through the owner's authenticated npm session, where npm enforces 2FA. It can safely resume by skipping already-published package versions only when their registry integrity exactly matches the validated tarballs. Registry errors fail closed; publishing disables automatic fetch retries so rate-limit failures are not obscured by retries with expired authentication. Resolve the error before explicitly retrying. Do not put an npm write token in GitHub.

After all packages exist, configure each package's trusted publisher with these exact claims:

- Provider: GitHub Actions
- Repository: `injectinglabs/injoffice`
- Workflow file: `release-npm.yml`
- Environment: `npm`
- Allowed actions: direct publishing (`npm publish`) and staged publishing

The tag workflow builds and validates packages in a job without OIDC permission, records each tarball's SHA-512 digest, and transfers those tarballs to the protected `publish` job. That job has `id-token: write`, contains no install/build/test step, and runs `npm publish` with lifecycle scripts disabled. It checks every artifact and public registry integrity before publishing anything. A retry skips already-published versions only when their SHA-512 integrity matches; an interrupted direct release can therefore resume safely. Re-run the failed publishing job to reuse the same validated artifacts. A full rebuild that changes tarball bytes is correctly rejected for versions already published.

For an existing release tag, run the workflow from `main`, supply the tag and select the `publish` input. Leaving `publish` unchecked validates and packs only. Manual runs use release tooling from the selected `main` workflow commit, build the existing tag, require its commit to be an ancestor of `main`, and pin the publishing job's package manifests and artifact name to that resolved commit. Existing tags must never be moved. The workflow run records both its tooling revision and release source revision; inspect both when approving a manual recovery release.

Configuring npm trust is a one-time owner operation protected by npm 2FA, independently for every package. A maintainer can use package settings or `npm trust github PACKAGE --file release-npm.yml --repository injectinglabs/injoffice --environment npm --allow-publish --allow-stage-publish`. Verify the resulting claims and preserve any intentionally configured publishers. Once OIDC publishing is verified, set each package's publishing access to require 2FA and disallow traditional tokens. The legacy `--stage-from` mode remains available for deliberate staged releases; those still require an npm approval for each version.

This source repository is public. Eligible trusted-publishing releases can generate npm provenance; verify provenance on each newly published version. Making source public does not add provenance retroactively to existing package versions.

The playground is not an npm package: `apps/playground` is marked `"private": true`, and the release script only enumerates publishable manifests under `packages/`. Package allowlists contain compiled ESM, declarations, source maps, package documentation, and explicitly declared runtime/legal assets. Source maps embed library TypeScript sources, so they include the library implementation as part of the public distribution. CI rejects individual tarballs over 3 MiB compressed or 10 MiB unpacked, except `@injoffice/xlsx-wasm`, whose separate, lazily loaded read-only rich-source WASM module raises its package budget to 3.5 MiB compressed and 12 MiB + 704 KiB unpacked. The native XLSX WASM module has a 7.75 MiB binary ceiling read from `go/xlsxpatch/cmd/xlsxnativewasm/max-bytes.txt` by CI and the local build; the read-only rich-source module retains a 7 MiB bound. CI also rejects a lockstep release set over 15 MiB compressed.

## Release validation

- Run npm ci, typecheck, tests, build, and check:packages from a clean public clone on Node 24.
- Run build, vet, and tests in every Go module.
- Publish a prerelease under the next tag, install it into a blank external consumer, and exercise at least one import from every package.
- Run the Pages artifact and verify its /injoffice/ asset base.
- Confirm the npm tarballs include declarations, sourcemaps, READMEs, and required legal files but no TypeScript source, tests, fixtures, credentials, or internal URLs.
- Create and test the seven Go submodule tags independently.

## Embedding-product gates

These are not missing library features:

- A vendored-source or pinned-version parity check in each consumer repository.
- Authenticated two-user acceptance tests across the host's Sheets, Docs, Decks, and PDF surfaces.
- Authorization, tenant isolation, connector egress policy, and collaboration-server conformance.
- An explicit production decision for the host's InjOffice feature flag.

The public repository should not contain consumer-specific gateways, dashboards, authentication configuration, credentials, tenant data, or production rollout state.
