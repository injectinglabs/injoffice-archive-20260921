# Public-release checklist

InjOffice is designed as independently consumable TypeScript packages and Go modules. A hosted backend is optional. This checklist separates public-repository readiness from work owned by an embedding product.

## Completed in the repository

- All 24 TypeScript packages emit ESM and declarations; WASM packages additionally ship their declared runtime assets.
- Every package has exports, a files allowlist, public publish metadata, repository links, and a package README with an example.
- The package check imports built entries and inspects every npm pack payload.
- A blank-consumer smoke test installs the generated tarballs and imports every package.
- The PDF package has a browser-safe subpath and packages required attribution/license material.
- The seven Go modules have public module paths, README examples, and independent tests.
- Node 24 and all Go modules run in public CI without private secrets.
- A tag-driven npm release workflow validates versions, license metadata, tests, builds, package contents, and dependency order before handing SHA-512-bound tarballs to a separate stage-only OIDC job.
- The Sheets playground has a GitHub Pages workflow.
- Contribution, conduct, security, collaboration-protocol, roadmap, NOTICE, Apache, and Unicode license documents exist.
- The project license is Apache-2.0 and every package ships its SPDX metadata and license text.
- XLSX pivot hydration and fail-closed inventory preservation are implemented in the canonical library.
- Native extract/apply contracts exist for XLSX, DOCX, and PPTX. Native completion remains partial and is not Microsoft Office parity.

Workspace apps (`apps/playground`, `apps/docx-page-paint-worker`) are private and unpublished. The in-repo `injoffice-server` is optional. Browser WASM packages can replace it for supported local workflows and must not become required by other packages.

## Charter boundaries

- **Native file API.** Go `Extract*` → native JSON (`@injoffice/sheets`, `@injoffice/docs`, `@injoffice/pptx-native`) → TypeScript paint compilers (preview) → mutation JSON → Go `Apply*`.
- **Editor shell.** Univer is an optional editor shell, not the file authority. Native paint is preview mode.
- **Optional runtime.** Hosts may inject their own backend, use a published WASM package, or run offline. The in-repo `injoffice-server` and WASM runtimes must not become required by format-model packages.
- **Product isolation.** The public repository must not contain consumer-specific gateways, dashboards, authentication configuration, tenant data, or credentials.
- **No Office parity claim.** Native coverage is a completion matrix, not an unrestricted round-trip guarantee.

## Required decisions and account setup

1. **Approve third-party attribution.** Review NOTICE, LICENSE, and LICENSE-UNICODE.txt before distribution.
2. **Confirm ownership.** Create or confirm the injectinglabs/injoffice GitHub repository and reserve/control the @injoffice npm scope. Registry lookup currently shows no published @injoffice/charts package; that does not prove scope ownership.
3. **Choose initial version policy.** The repository is aligned at 0.1.0. Decide whether all npm packages stay lockstep and create matching Go submodule tags such as go/xlsxpatch/v0.1.0.
4. **Configure release accounts.** Add required reviewers to the protected `npm` GitHub environment, bootstrap the first package versions interactively with 2FA, configure stage-only trusted publishing for every package, disallow traditional publishing tokens, and configure Pages source, branch protection, and private vulnerability reporting.

Repository visibility, npm publication, release tags, and Pages deployment remain explicit owner actions. Merging readiness work does not perform them.

## npm publishing bootstrap and steady state

npm requires a package to exist before either staged publishing or a trusted-publisher relationship can be configured. The one-time bootstrap therefore starts with a manual `Release npm packages` workflow run for `v0.1.0`. That run validates the repository and uploads a one-day `npm-release-*` artifact without publishing anything. Download and inspect that artifact, then publish those exact CI-built tarballs through an explicit local owner action:

```sh
gh run download RUN_ID --name npm-release-COMMIT_SHA --dir release-tarballs
npm run release:packages -- v0.1.0 --bootstrap-from release-tarballs
```

The bootstrap mode validates every recorded SHA-512 digest, refuses to run when `CI` is set, and publishes in dependency order through the owner's authenticated npm session, where npm enforces 2FA. It can safely resume by skipping already-published package versions only when their registry integrity exactly matches the validated tarballs. Do not put an npm write token in GitHub.

After all packages exist, configure each package's trusted publisher with these exact claims:

- Provider: GitHub Actions
- Repository: `injectinglabs/injoffice`
- Workflow file: `release-npm.yml`
- Environment: `npm`
- Allowed action: staged publishing only

The tag workflow then builds and validates packages in a job without OIDC permission, records each tarball's SHA-512 digest, and transfers only those tarballs to the protected `stage` job. That job has `id-token: write`, contains no install/build/test step, and runs `npm stage publish`. If a tag is pushed after its bootstrap release, the job instead verifies that every public registry integrity matches and exits successfully. A maintainer must review and approve every newly staged version with 2FA before it becomes public. Once OIDC staging is verified, set each package's publishing access to require 2FA and disallow traditional tokens.

Trusted publishing can still operate while this GitHub repository is private, but npm provenance is unavailable for a private source repository. Provenance will be generated automatically for public packages after the repository becomes public; the release script deliberately does not force `--provenance` while the repository remains private.

The playground is not an npm package: `apps/playground` is private, and the release script only enumerates publishable manifests under `packages/`. Package allowlists contain compiled ESM, declarations, source maps, package documentation, and explicitly declared runtime/legal assets. Source maps embed library TypeScript sources, so public npm publication exposes the library implementation even if GitHub is still private. CI rejects any individual tarball over 3 MiB compressed or 10 MiB unpacked, and rejects a lockstep release set over 15 MiB compressed.

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
