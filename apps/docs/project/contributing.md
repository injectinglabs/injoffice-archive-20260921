# Contributing

Start with a focused issue or a small change to a reproducible workflow. Include the document format, runtime, current behavior, expected behavior, and a synthetic fixture where possible.

## Prepare a checkout

Use Node.js 22+ and the Go versions declared by the relevant modules. Run `npm ci`, then `npm run build`. Package exports refer to built artifacts, so typechecking consumers before building can fail for the wrong reason.

## Validate a change

```sh
npm run typecheck
npm test
npm run check:packages
npm run check:dependency-integrity
npm run check:docs-api
```

For a Go module, run `go test ./...` in that module. Native browser changes should also pass the installed-tarball and real-file browser tests. Preserve failing fixtures so future regressions are reproducible.

## Preserve the architecture

Do not move native file authority into editor HTML, screenshots, or renderer state. Add a capability only when the corresponding parser/writer/adapter can validate and execute it. Refuse unsupported work rather than silently weakening the contract.

New agent operations need exact schemas, bounded reads, source checks, explicit authorization boundaries, and verification against committed output.

## Documentation changes

Task guides belong in `apps/docs`; package-specific reference belongs in `packages/<name>/README.md` or the relevant technical contract under `docs/`. Add or update a source-imported TypeScript example when the public workflow changes. See [documentation maintenance](documentation).

The repository's root contribution and security policies remain authoritative. Report vulnerabilities through the documented private channel rather than posting sensitive reproduction material in a public issue.
