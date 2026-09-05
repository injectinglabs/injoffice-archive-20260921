# @injoffice/font-metrics

Font discovery plus InjOffice's renderer-neutral native text-layout contract.

The package has two boundaries:

- `@injoffice/font-metrics` exposes Node.js system-font discovery for document pipelines.
- `@injoffice/font-metrics/layout` is pure TypeScript. It has no DOM, canvas, React,
  Konva, Node, Electron, storage, or asset dependency and is safe for browser workers
  and servers.

```bash
npm install @injoffice/font-metrics
```

```ts
import { findSystemFont, isFamilyInstalled } from '@injoffice/font-metrics'

if (isFamilyInstalled('Arial')) {
  const bytes = findSystemFont('ArialMT', 'Arial')
  if (bytes) console.log(`found ${bytes.length} font bytes`)
}
```

Results depend on fonts installed on the current operating system. Callers remain responsible for font licensing and for embedding permissions in generated documents.

## Native Office layout contract

PPTX RenderTree and native DOCX pagination share the same input, resolution,
shaping, cluster, and line-metric vocabulary:

```ts
import {
  NATIVE_TEXT_LAYOUT_VERSION,
  classifyNativeOfficeLineBreak,
  shapingCacheKey,
  validateTextRunInput,
  type NativeFontResolver,
  type NativeFontManifest,
  type NativeTextShaper,
  type TextRunInput,
} from '@injoffice/font-metrics/layout'

const run: TextRunInput = {
  version: NATIVE_TEXT_LAYOUT_VERSION,
  text: 'Quarterly results',
  fontSizeMilliPoints: 18_000,
  font: {
    families: ['Aptos', 'Carlito'],
    weight: 600,
    style: 'normal',
    stretch: 100,
    fallbackChainIds: ['latin.default'],
  },
  script: 'Latn',
  language: 'en-US',
  direction: 'ltr',
  features: [{ tag: 'kern', value: 1 }],
}

if (!validateTextRunInput(run).ok) throw new Error('invalid native text run')

declare const resolver: NativeFontResolver
declare const shaper: NativeTextShaper
declare const manifest: NativeFontManifest
const resolved = await resolver.resolve({ manifest, run })
if (resolved.status === 'resolved') {
  const key = shapingCacheKey(run, resolved.face, shaper)
  // Cache by content digest + provider revision + every layout-affecting input.
}

const boundary = classifyNativeOfficeLineBreak('word-', 'next')
// => 'allowed'; 'unsupported' requires preserve/refusal if wrapping needs it.
```

Contract rules:

- Sizes, advances, offsets, and line metrics use integer `milliPoints` (1/1000 pt),
  never device pixels.
- Text and cluster ranges use UTF-16 offsets, matching JavaScript and ProseMirror.
- Resolved faces are content-addressed by SHA-256; a system family name alone is
  not a deterministic identity.
- Fallback is ordered by manifest face ids and reported as an exact, substitute,
  fallback, or explicit refusal decision. Missing glyphs are never silently hidden.
- Cache keys include manifest/provider revisions, font bytes, collection face,
  script, language, direction, OpenType features/variations, spacing, size, and text.
- Manifest and run validators reject unknown fields and unbounded or malformed
  input before a host provider executes.
- `classifyNativeOfficeLineBreak` and its allocation-free UTF-16 range form are
  the shared, deterministic, conservative Office boundary for shaped cluster text.
  They model a bounded v1 subset of
  words, glue/NBSP, word joiner, ZWSP, hyphen/slash, Han/Hangul, and common
  East-Asian open/close punctuation. Japanese kana/kinsoku, U+3000, and unknown classes return `unsupported` until their exact
  soft-line-edge paint semantics are qualified; provider `whitespace`
  flags never create a line-break opportunity.

`NativeFontResolver` is implemented by each host: a browser worker may lazily load
licensed font assets or document-embedded bytes, while a server may resolve an
audited font directory. Resolver implementations own font inventory, selection,
loading, and metrics extraction; the shaper below independently verifies the
current v1 table projection but never discovers, selects, loads, or substitutes
a face.

## Canonical HarfBuzz provider

Node 22 consumers can use the genuine pinned-runtime qualification boundary at
`@injoffice/font-metrics/harfbuzz`. It uses the exact `harfbuzzjs@1.6.0`
runtime (HarfBuzz 14.3.0) and verifies the installed package manifest, executed
JavaScript entry/loader, and shaping WASM against pinned SHA-256 digests before
creating a provider.

```ts
import { createHarfBuzzTextShaperV1 } from '@injoffice/font-metrics/harfbuzz'

const shaper = createHarfBuzzTextShaperV1({
  sourceRevision: 'git:0123456789abcdef',
})
```

The required source revision, exact runtime artifacts, generated Unicode 13
classifier revision, shaping policy, cluster level, flags, and deterministic scaling rule
are hashed into `providerRevision`; the unhashed immutable manifest is available as
`shaper.provenance`. The ordinary `shapingCacheKey` additionally binds the
exact font digest/index/face, text, size, explicit script/language/direction,
features, variations, and spacing.

The qualified v1 slice accepts complete horizontal LTR/RTL Latin, Arabic,
Hebrew, or Common runs,
non-system fixed TrueType `glyf`/`loca` sfnt or TTC faces, zero letter/word spacing, and
explicit `kern`/`liga` settings when advertised by the exact face. It emits real
glyph IDs, integer milli-point advances/offsets,
monotone-grapheme UTF-16 clusters, unsafe-break flags, and line metrics that
must exactly match the digest-bound font tables. It rejects `.notdef`,
structurally malformed/oversized fonts, fabricated metrics, collection confusion,
unpaired surrogates, partial runs, authored bidi controls, unqualified
scripts/explicit feature settings, system faces, variations, vertical directions, and every resource
overflow. It never calls DOM, canvas, HTML, browser text measurement,
LibreOffice, or OS-font substitution.

Unspecified OpenType features use HarfBuzz 14.3.0's shaping defaults; that
policy and runtime version are part of `providerRevision`.

This entry is qualified as a bounded per-compilation provider, not as a shared
long-lived service. `harfbuzzjs@1.6.0` relies on JavaScript finalization, so the
native page-paint compiler constructs an invocation-scoped provider with hard
face/byte/call budgets and discards it after the atomic result. The boundary
preflights fixed TrueType tables, requires exact resolver/shaper metric
agreement, and is covered end to end through shaped-lines and page paint.
CFF/CFF2, variable fonts, scripts other than Latin,
Arabic, and Hebrew, leading inherited marks, default ignorables, and
unqualified controls refuse in v1.

HarfBuzz shapes an already itemized run; it is not a Unicode bidi algorithm or
script/language resolver. `@injoffice/font-metrics/bidi` supplies the separate
bounded UAX #9 boundary used by native DOCX shaping. It pins `bidi-js@1.0.3`
(Unicode 13.0.0), verifies the exact entry, manifest, and factory SHA-256, applies explicit
DOCX run direction as isolates, emits per-UTF-16 embedding levels, and applies
line-specific L1/L2 ordering to cluster-aligned inputs. Authored bidi controls,
malformed UTF-16, every supplementary scalar, scalars outside the assigned Unicode 13 repertoire,
overlapping explicit ranges, and every resource overflow fail closed. The
Unicode 13 classifier is regenerated from vendored, digest-verified UCD inputs
and never consults host Unicode properties. Its source, projection, generator,
and runtime encoding digests are bound into provider revisions and checked
before use. Full
bidi-js, require-from-string, and Unicode notices are packaged with the module.
