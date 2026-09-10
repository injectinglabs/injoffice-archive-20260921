# Maintaining these docs

The documentation is a separate private npm workspace built with VitePress. It has its own server, static output, search index, and navigation. It does not import the playground application or bundle its editors/WASM fixtures.

## Content ownership

| Content | Edit here |
| --- | --- |
| Task guides and tutorials | `apps/docs/getting-started`, `guides`, `agents`, `integration` |
| Checked TypeScript examples | `apps/docs/examples/*.ts` |
| Package reference prose | `packages/<package>/README.md` |
| Go module reference | `go/<module>/README.md` |
| Protocol contracts | `docs/*.md` |
| Exported-name inventory | Regenerate `docs/api-reference.json` after building |

`apps/docs/scripts/generate.mjs` republishes selected source docs with repository-relative links resolved for the separate site. Generated reference pages are ignored by Git and rebuilt from source.

## Preview and build

```sh
npm run docs:dev
npm run docs:build
npm run docs:preview
```

Development and preview use port 3200 with strict port selection. They do not replace the demo on port 3100. The static output is `apps/docs/.vitepress/dist`.

For subpath hosting, set the base at build time:

```sh
DOCS_BASE=/docs/ npm run docs:build
```

The default build uses `/` and is suitable for a dedicated documentation hostname. The site uses explicit `.html` routes so a static host does not need an SPA fallback for article URLs. Configure directory index handling for `/` and section indexes. Serve the generated `404.html` with a real 404 status for missing pages.

Building does not publish or change AWS resources. Hostname, DNS, cache policy, and deployment approval are separate concerns.

## Regression checks

VitePress checks page links during the static build. Documentation tests check generated coverage and internal guide targets. Source-imported TypeScript examples are typechecked; pure examples execute against synthetic documents. Browser smoke tests cover article navigation, local search, code copying, responsive navigation, and a direct `.html` deep link.

The existing `check:docs-api` gate prevents the declaration inventory from drifting after a package build. A generated symbol list is intentionally not described as a complete signature-level API reference.
