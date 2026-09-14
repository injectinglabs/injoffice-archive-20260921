# Exact cardinal geometry arcs

At a 4,000,000 × 400,000 EMU frame, the pinned `can`, `leftBrace`,
`rightBrace`, `leftBracket`, and `rightBracket` definitions have exactly
representable integer arc centers, radii, and quarter-turn endpoints. Their
text rectangles are positive. Charging the generic polar/libm construction
allowance incorrectly refused these paths at the existing 0.125 EMU cap.

The source evaluator now omits that construction allowance only when source
rational provenance proves integral radii, cardinal start/sweep angles, an
exact current pen, exact unit path scaling, and safe integer center/endpoint
arithmetic. Cardinal ellipse evaluation already bypasses libm. Input error and
all earlier path error still count; a later move or close cannot clear the
accumulated budget. Radius fit, command limits, coordinate limits, and the
non-cardinal fallback are unchanged.

The independent Python oracle uses Fraction arithmetic for every path endpoint,
center, radius and SVG chord fit, and 90-digit Decimal arithmetic for text
rectangles. It reads the pinned catalog directly and does not call the product
evaluator. The Go source test compares every command and text rectangle after
strict and transitional PPTX extraction. The numerical harness then checks
actual WASM extraction and full public compile/paint, source immutability,
determinism, finite coordinates and retained negative placeholders.

Run after `npm ci` and `npm run build` with Node 22 or newer:

```sh
python3 scripts/generate-pptx-exact-cardinal-oracle.py --check
node scripts/qualify-pptx-exact-cardinal-arcs.mjs
```

The harness prints its temporary evidence directory, containing eight actual
PPTX sources and `report.json`: five painted positives, three refusals
(near-cardinal, non-unit scaling, and prior accumulated error). Go regressions
also exercise integer-looking values without exact provenance, close restoration,
unsafe intermediates, and the unchanged command limit.

This is a bounded numerical qualification, not browser-pixel or Office parity.
The six wide/tall `cornerTabs`, `plaqueTabs`, and `squareTabs` cases remain
unsupported: their pinned text rectangles invert at aspect 10:1. The catalog
uses `hypot(w,h)/20` on each edge, so continuous positive text dimensions require
an aspect below `sqrt(99)`; integer rounding also has to preserve positivity.
Those definitions are not clamped or rewritten.
