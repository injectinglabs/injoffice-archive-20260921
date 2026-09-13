# DrawingML preset catalog

This catalog uses the unchanged 187-shape Apache POI resource at commit `338882ac8898df5c13a7d15f533204c5dd8607d6`, retained with its original license and NOTICE in `go/pptxpatch/presetdata`. Runtime fingerprint verification prevents unreviewed definition changes. Evaluated geometry uses the same bounded Go engine as custom source geometry; there is no browser guide interpreter.

## Sequential guides and adjustments

ECMA-376 Part 1 §20.1.9.11 specifies guide calculation in source order. Several supplied presets reuse a guide name. The catalog normalizer assigns a distinct internal name to each reassignment: a formula reads earlier bindings before its new result is installed, and subsequent guides/path metadata refer to the latest binding. The original data is unchanged. Tests check values both before and after reassignment, including self-reference to the prior binding.

Adjustment overrides replace only a preset's declared adjustment values. Unknown names and out-of-bound integers refuse. Each evaluation clones its definition, so one override cannot change later defaults. Custom-source sequential reassignment remains a separate qualification task while the original custom profile rejects duplicate names.

The electronic shape definitions also use `cd3`, `hd10`, `wd12`, and `wd32`, despite their absence from the prose built-in table. These have explicit circle/height/width divisor definitions, with tests; arbitrary inferred names remain unsupported.

## Arithmetic policy

Arbitrary custom programs retain the MAXSAFE intermediate-result guard. Only fingerprinted catalog definitions use a finite `1e100` internal bound for higher-order distance equations; literal inputs, periodic angles, frames, and final EMU outputs retain MAXSAFE bounds. This is a scoped floating-point preview policy, not exact rational arithmetic. Independent extreme-dimension and adjustment conditioning checks remain required before catalog qualification; general exact-intermediate algebra remains tracked. A regression refuses custom cancellation that would lose `+1` above binary64 integer precision.

## Explicit catalog errata

The same anomalies occur in the ECMA-376 2016 electronic shape addendum and the pinned POI data. Corrections below apply only to that fingerprinted catalog, with guards checking exact original values. Arbitrary source formulas remain untouched.

- `circularArrow` and `leftCircularArrow`: `xB` and `yB` have a redundant fourth operand after the defined three-operand `+-` expression.
- `leftRightCircularArrow`: the same issue affects `xB`, `yB`, `xJ`, and `yJ`.
- `pie`: the text rectangle's top/right references are transposed (`t="ir"`, `r="it"`); corrected references are `t="it"`, `r="ir"`.

The independent [ONLYOFFICE geometry implementation](https://github.com/ONLYOFFICE/sdkjs/blob/master/common/Drawings/Format/Geometry.js) defines `AddGuide(name, formula, x, y, z)` and evaluates addition/subtraction as `x+y-z`; extra JavaScript arguments in its [preset definitions](https://github.com/ONLYOFFICE/sdkjs/blob/master/common/Drawings/Format/CreateGeometry.js) do not participate. Its pie rectangle explicitly uses `il,it,ir,ib`. Tests independently verify subtraction and the non-square inscribed rectangle coordinates. No ONLYOFFICE implementation code is incorporated.

## Remaining integration

The catalog milestone is under implementation. The public source preset hook, shading policy, installed-WASM attribution, complete default/adjustment sweep, and browser catalog gallery must pass before it ships. Default catalog coverage reports remaining reasons; it is not a claim of completion or Office visual parity.

ECMA's [standard download](https://ecma-international.org/publications-and-standards/standards/ecma-376/) supplies the electronic geometry addendum. All downloaded benchmark/reference artifacts remain outside the repository; the explicitly licensed POI resource is product data.

## Deterministic path shading policy

ECMA-376 §20.1.10.37 specifies only qualitative lighter/darker fill modes; searches of Microsoft implementation notes did not establish normative numeric strengths. The proposed renderer policy `linear-srgb-path-tone-20-40-v1` blends 20% (`*Less`) or 40% (unqualified mode) white/black in linear sRGB, followed by standard sRGB encoding and final 8-bit rounding. These explicit preview choices are not presented as Office behavior. Independent fixed color tests distinguish linear-light blending from encoded-RGB multiplication. Normal and no-fill paths preserve their original semantics. Source and renderer diagnostics must disclose this policy whenever shaded modes are painted.
