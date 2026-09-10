# Quickstart

Start with a PDF transformation. It exercises a public package without an editor, browser Worker, model, or running InjOffice server.

## Prerequisites

Use Node.js 22 or newer and npm. For this example you need an existing PDF you are allowed to process. Keep the source file unchanged until you have checked the output.

## Install one package

In an empty project directory:

```sh
npm init -y
npm install @injoffice/pdf
```

The docs track `main`, not a frozen release. If npm reports that a package is unavailable, check the [release status](../reference/generated/contracts/public-release) rather than substituting a similarly named package. Use a source checkout below for unreleased packages.

## Transform a real file

Create `rotate.mjs`:

```js
import { readFile, writeFile } from 'node:fs/promises'
import { applyPageOps, readInfo } from '@injoffice/pdf/browser'

const source = new Uint8Array(await readFile('input.pdf'))
const before = await readInfo(source)
if (before.pageCount < 1) throw new Error('The PDF has no pages')
const output = await applyPageOps(source, [
  { type: 'rotate', pages: [1], degrees: 90 },
])
const after = await readInfo(output)
if (after.pageCount !== before.pageCount) throw new Error('Unexpected page count')
await writeFile('rotated.pdf', output)
console.log(`Wrote rotated.pdf with ${after.pageCount} pages`)
```

Run it beside `input.pdf`:

```sh
node rotate.mjs
```

Open `rotated.pdf` in a PDF viewer. Page 1 should have rotated 90 degrees. The original remains in `input.pdf`. Successful parsing and a matching page count are useful checks, not a complete visual or semantic compatibility proof.

The `/browser` entry is also usable for these pure byte operations in Node. It avoids the root entry's Node-specific PDFium, font, OCR, and image integrations.

## From source

The repository has independent npm workspaces and Go modules. Native WASM distributions need the Go toolchain as well as Node. Follow the Go versions in each `go.mod`; the existing build scripts select the browser target.

```sh
git clone https://github.com/injectinglabs/injoffice.git
cd injoffice
npm ci
npm run build
npm run typecheck
npm run check:packages
```

Do not install from an unbuilt `packages/*` directory: package exports point at generated `dist` files, and Worker packages include compiled assets. For distributable payloads, use the repository's packing checks and release instructions.

## Read docs or run the demo

From the repository root:

```sh
# Documentation only, on port 3200
npm run docs:dev

# In a separate terminal, optional interactive demo on port 3100
npm run dev
```

The two applications have separate builds. Documentation does not need the demo, a backend, or model credentials.

## Next steps

- [Add an explicit review step with the agent API](../agents/quickstart).
- [Use native XLSX, DOCX, or PPTX in a Worker](../integration/browser).
- [Choose the smallest package set for your application](packages).
