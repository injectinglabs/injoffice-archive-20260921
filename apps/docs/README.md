# Standalone InjOffice documentation

VitePress documentation, separate from `apps/playground`. The static site loads
no demo code, file fixtures, WASM engines, or model services. Local search needs
no external indexing service. Current main-branch docs are not version-pinned
npm release docs.

VitePress 1.6.4 is pinned, with its nested Vite dependency overridden to 6.4.3
to avoid the advisories affecting its default Vite 5 dependency. The override
is scoped to documentation and does not change the playground's Vite version.
Build, dev-server, and browser checks cover this combination; remove/revisit
the override when the stable VitePress dependency is updated.

From the repository root:

```sh
npm ci
npm run docs:dev      # http://localhost:3200
npm run docs:build    # apps/docs/.vitepress/dist
npm run docs:preview  # built site, also port 3200
```

Native packages must be built before checking the TypeScript examples:

```sh
npm run build
npm run typecheck -w apps/docs
npm test -w apps/docs
npm run test:docs-browser
```

Guide snippets import actual `examples/*.ts` files. Pure PDF/agent/authoring/
chart/pivot examples execute in Node tests; Worker examples are typechecked and
the underlying engines are covered by the existing installed-tarball browser
suite. The DOCX projector is explicitly host-supplied, not a complete editor.

`scripts/generate.mjs` generates package/Go/contract references from their
maintained source files and the tracked `docs/api-reference.json`. Generated
reference pages are ignored. The build fails on unresolved local page links.

Use `DOCS_BASE=/docs/ npm run docs:build` for path hosting; default `/` suits a
dedicated docs hostname. Routes use `.html`, without a catch-all SPA rewrite.
Deployment is independent and does not update the demo's AWS release.

Design: white (#ffffff) article surface, mist-blue (#f4f7fb) navigation, blue
(#245cc5) links, ink (#18263d) headings, slate (#53677f) secondary text, and
green (#28704b) notes. Aptos/Segoe UI typography reflects the document domain;
code uses the host monospace family. A quiet left navigation and right contents
rail keep the central column focused on reading. No marketing hero or editor
dashboard is embedded in documentation.
