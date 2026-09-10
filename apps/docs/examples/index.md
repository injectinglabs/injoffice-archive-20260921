# Code examples

The examples below are real TypeScript source files, imported into the guides rather than copied into separate code blocks. They compile against the repository's public package exports.

| Example | Runtime | What it demonstrates |
| --- | --- | --- |
| [PDF rotation](../guides/pdf) | Node or browser | Read, transform, and reopen bytes |
| [PDF agent](../agents/quickstart) | Node or browser | A deterministic proposal with separate review and verified commit |
| [XLSX cell edit](../guides/spreadsheets) | Browser Worker | Extract, source-bound apply, and readback |
| [DOCX transaction](../guides/documents) | Browser Worker + host projector | Guard an editor transaction against the original document |
| [PPTX inspection](../guides/presentations) | Browser Worker | Discover source-anchored text targets |
| [Presentation authoring](../guides/presentations#author-a-new-deck) | Node or browser | Compile a DeckSpec into a native model |
| [Charts and pivots](../guides/spreadsheet-tools) | Pure functions | Build a chart option and aggregate a pivot |

## Run the checks

From a source checkout with Node.js 22+ and the repository's Go toolchain:

```sh
npm ci
npm run build
npm run typecheck -w apps/docs
npm test -w apps/docs
```

The documentation tests execute the PDF, agent, authoring, chart, and pivot examples. They also check rejection paths and the documentation's generated references. The WASM examples are typechecked; native engine execution is covered separately by the repository's installed-tarball browser suite. The DOCX example requires a host-supplied native transaction projection—it does not claim to implement an editor.

## Using the examples in your application

Each source file exports a function. Copy that function into your app, install the packages named in its imports, and supply the declared arguments. Byte-oriented examples deliberately leave file selection, persistence, and download to the host. None runs an LLM or contacts an InjOffice service.

For a Node PDF script, the host boundary can be as small as:

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { rotateLastPage } from './pdf.js'

const { output } = await rotateLastPage(new Uint8Array(await readFile('input.pdf')))
await writeFile('rotated.pdf', output)
```

Use an output filename distinct from the input until your own acceptance checks pass.
