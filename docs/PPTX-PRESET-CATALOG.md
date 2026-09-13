# DrawingML preset catalog

This catalog uses the unchanged 187-shape Apache POI resource at commit `338882ac8898df5c13a7d15f533204c5dd8607d6`, retained with its original license and NOTICE in `go/pptxpatch/presetdata`. The WASM embeds a deterministic gzip copy; bounded inflation verifies the exact original length and SHA-256 before XML parsing, and rejects corrupt, concatenated or trailing data. The original XML remains unchanged in the repository. Runtime fingerprint verification prevents unreviewed definition changes. Evaluated geometry uses the same bounded Go engine as custom source geometry; there is no browser guide interpreter.

## Sequential guides and adjustments

ECMA-376 Part 1 §20.1.9.11 specifies guide calculation in source order. Several supplied presets reuse a guide name. The catalog normalizer assigns a distinct internal name to each reassignment: a formula reads earlier bindings before its new result is installed, and subsequent guides/path metadata refer to the latest binding. The original data is unchanged. Tests check values both before and after reassignment, including self-reference to the prior binding.

Adjustment overrides replace only a preset's declared adjustment values. Unknown names and out-of-bound integers refuse. Each evaluation clones its definition, so one override cannot change later defaults. Custom-source sequential reassignment remains a separate qualification task while the original custom profile rejects duplicate names.

The electronic shape definitions also use `cd3`, `hd10`, `wd12`, and `wd32`, despite their absence from the prose built-in table. These have explicit circle/height/width divisor definitions, with tests; arbitrary inferred names remain unsupported.

## Arithmetic policy

Arbitrary custom programs retain the MAXSAFE intermediate-result guard. Only fingerprinted catalog definitions use a finite `1e100` internal bound for higher-order distance equations; literal inputs, periodic angles, frames, and final EMU outputs retain MAXSAFE bounds.

The Go evaluator carries exact rational shadows for algebraic guide values, including built-in fractions, conditional comparisons, selected pin/min/max operands, perfect-square roots, and provable cardinal angles. Numerators and denominators have a 4096-bit limit. Bounded expression identities preserve `x-x=0` when equivalent irrational subexpressions share their source calculation. An independent callout regression proves that `(w/100000)*h/w-h/100000` is exactly zero, so an ordinary frame cannot choose the wrong wedge edge through binary64 cancellation. Arbitrary custom programs still refuse unsafe intermediate magnitudes; general widened custom arithmetic remains separately tracked.

Non-rational values carry outward interval estimates. Elementary functions include conservative binary64 allowances, including angle conversion error near sine/cosine zeros. Uncertain discontinuous branch predicates and invalid/uncertain domains refuse. Source coordinates, radii and arc angles must fit an accumulated one-eighth-EMU uncertainty allowance per path after path-space scaling; arc qualification includes both polar offsets, prior path uncertainty, aspect-ratio sensitivity and a conservative construction allowance. Final integer quantization is separate. This is an explicitly bounded preview policy, not a formal proof of Go libm error or arbitrary transcendental identities. Very thin, large or ill-conditioned configurations can refuse rather than silently select unstable geometry.

Independent review compared exact rational and 90-digit transcendental reference equations across 5,506 configurations, then exercised 23 suspect configurations through Go: accepted results were corrected by exact algebra or met the reference expectations; unstable cases refused. Permanent tests retain the wedge coordinates, near-zero cosine enclosure, numerator/denominator budgets, uncertain predicates, final-coordinate limits and accumulated arc uncertainty. All 187 default definitions pass direct evaluation and public source extraction in both OOXML dialects at 4,000,000×3,000,000 EMU.

## Explicit catalog errata

The same anomalies occur in the ECMA-376 2016 electronic shape addendum and the pinned POI data. Corrections below apply only to that fingerprinted catalog, with guards checking exact original values. Arbitrary source formulas remain untouched.

- `circularArrow` and `leftCircularArrow`: `xB` and `yB` have a redundant fourth operand after the defined three-operand `+-` expression.
- `leftRightCircularArrow`: the same issue affects `xB`, `yB`, `xJ`, and `yJ`.
- `pie`: the text rectangle's top/right references are transposed (`t="ir"`, `r="it"`); corrected references are `t="it"`, `r="ir"`.

The independent [ONLYOFFICE geometry implementation](https://github.com/ONLYOFFICE/sdkjs/blob/master/common/Drawings/Format/Geometry.js) defines `AddGuide(name, formula, x, y, z)` and evaluates addition/subtraction as `x+y-z`; extra JavaScript arguments in its [preset definitions](https://github.com/ONLYOFFICE/sdkjs/blob/master/common/Drawings/Format/CreateGeometry.js) do not participate. Its pie rectangle explicitly uses `il,it,ir,ib`. Tests independently verify subtraction and the non-square inscribed rectangle coordinates. No ONLYOFFICE implementation code is incorporated.

## Public source and remaining geometry work

The existing native extractor accepts the catalog's default definitions and declared literal adjustment overrides. The eight previously qualified default presets keep their existing preset contract and editability; new catalog shapes and adjusted geometry use evaluated paths with preserve-only authority. Unknown names, undeclared or duplicate adjustments, ambiguous geometry sources, unsupported paint clauses and unstable numerical results refuse. The Go API also exposes catalog names and evaluated preset geometry; the browser receives evaluated paths rather than interpreting guides.

All six path fill modes reach the public compiler, paint worker and existing `NativePptxVector` demo component. The WASM package retains the upstream catalog license, NOTICE and provenance in `dist/preset-catalog`. Browser authoring convenience, arbitrary transforms, further custom-guide capability and legacy `PptxFilePreview` integration remain subsequent geometry milestones. Catalog completion does not close the full complex-shape assignment or claim Office visual parity.

ECMA's [standard download](https://ecma-international.org/publications-and-standards/standards/ecma-376/) supplies the electronic geometry addendum. All downloaded benchmark/reference artifacts remain outside the repository; the explicitly licensed POI resource is product data.

## Deterministic path shading policy

ECMA-376 §20.1.10.37 specifies only qualitative lighter/darker fill modes; searches of Microsoft implementation notes did not establish normative numeric strengths. The renderer policy `linear-srgb-path-tone-20-40-v1` blends 20% (`*Less`) or 40% (unqualified mode) white/black in linear sRGB, followed by standard sRGB encoding and final 8-bit rounding. These explicit preview choices are not presented as Office behavior. Independent fixed color tests distinguish linear-light blending from encoded-RGB multiplication. Normal and no-fill paths preserve their original semantics. Source and renderer diagnostics disclose this policy whenever shaded modes are painted.
