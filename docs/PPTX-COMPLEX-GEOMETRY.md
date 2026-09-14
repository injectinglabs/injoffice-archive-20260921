# DrawingML evaluated geometry

The custom geometry engine evaluates DrawingML in the Go source parser. The native boundary carries numeric paths and a text rectangle, not XML or an executable guide language. Source geometry remains preserve-only; shape mutation does not rewrite custom paths.

This describes the custom-path engine introduced in the first geometry milestone. The [preset catalog extension](PPTX-PRESET-CATALOG.md) adds catalog definitions, declared adjustment overrides, all six path fill modes and numerical qualification. Source shape/group transforms and the local file-preview integration are documented separately. Interactive geometry handles, upright/separate body text rotation and newly complex graphic-frame qualification remain outside this engine milestone; no milestone establishes Office visual parity.

## Evaluation

The engine supports the seventeen ECMA-376 guide operations: `*/`, `+-`, `+/`, `?:`, `abs`, `at2`, `cat2`, `cos`, `max`, `min`, `mod`, `pin`, `sat2`, `sin`, `sqrt`, `tan`, and `val`. It evaluates adjustment values followed by ordered shape guides. Built-in frame, divisor, and angle guides follow the finite list in §20.1.10.56. User guide names may be reassigned in source order, including adjustment-list assignments followed by shape-guide assignments. Each formula reads the preceding bindings; all value, rational, interval and symbolic shadows are calculated before replacing the named result. Previously captured values remain unchanged. Built-in names and numeric-operand ambiguity retain explicit refusal. Forward references, initial self-references and undefined numeric domains refuse the whole geometry.

Calculations retain exact rational shadows where algebra permits and conservative interval estimates for non-rational results; see the catalog extension for numerical budgets and refusal policy. Coordinates round at the final integer-EMU boundary. Path `w` and `h` scale each axis independently. A missing axis extent uses the shape frame. Text rectangle guides use the frame coordinate system, independently of path coordinate spaces; missing `rect` uses the full frame.

Paths support move, line, quadratic and cubic Bezier curves, elliptical arcs, and close. Close restores the current point to the active subpath origin. Multiple subpaths retain their original order and direction for nonzero winding, including holes. Multiple paths preserve paint order. Paths accept all six DrawingML fill modes and Boolean stroke. Shaded modes use the explicitly disclosed deterministic policy described in the catalog extension. Both Boolean `extrusionOk` values are nonvisual permission metadata; applied scene/shape 3D remains refused by source qualification. The specification prose and schema disagree about its default, which does not affect this two-dimensional paint policy.

## Elliptical arcs

DrawingML angles describe geometric rays, not SVG ellipse parameters. For radii `rx`, `ry` and DrawingML angle θ, the ellipse parameter is:

```
t = atan2(rx × sin(θ), ry × cos(θ))
point = (rx × cos(t), ry × sin(t))
```

The current point anchors the start of the ellipse. Endpoint calculation takes place before path-axis scaling. Signed sweeps preserve clockwise or counterclockwise winding. Sweeps up to one revolution split into at most four segments of at most a quarter revolution; this avoids the SVG coincident-endpoint full-circle omission.

After rounding, an exact integer predicate checks that the endpoint chord fits the transported radii. A shape that would require SVG to silently enlarge its radii refuses instead. A zero/zero radius arc is a no-op; a single collapsed radius refuses. These checks include tiny arcs and extreme aspect ratios.

## Limits and refusal

The engine accepts at most 1,024 adjustment/shape guides combined, 128 paths, 512 emitted commands per path, and 8,192 emitted commands overall. Numeric values must remain finite and within the native safe-integer envelope. The renderer must also honor its smaller configured coordinate and aggregate paint limits.

Unknown attributes or child elements, invalid point counts, commands before the first move, unresolved operands, ambiguous numeric guide names, built-in shadowing, nonempty unqualified handles/connections, and unsupported path paint clauses refuse the entire geometry. No successfully parsed prefix is exposed as a complete shape.

## Sources and evidence

The governing source is [ECMA-376, Part 1](https://ecma-international.org/publications-and-standards/standards/ecma-376/): §20.1.9 custom geometry/path elements, §20.1.9.11 guide formulas, and §20.1.10.56 built-in guides. Non-square angle conversion was independently checked against [Apache POI ArcToCommand](https://github.com/apache/poi/blob/trunk/poi/src/main/java/org/apache/poi/sl/draw/geom/ArcToCommand.java), whose conversion explicitly distinguishes geometric OOXML angles from ellipse parameter angles. These references are evidence, not copied implementation code.

Dedicated Go tests exercise all operators, order and budget guards, both Bezier controls, scaled path coordinates, non-square arcs, negative/full sweeps, close/current-point behavior, atomic XML refusal, and exact rounded-arc radius consistency. The connected source, JSON, mutation, renderer, and paint-worker tests cover the complete path through the public API. An external generated fixture was extracted by the actual browser Go/WASM runtime, compiled with DejaVu Sans, and displayed by the existing `NativePptxVector` component in Chrome: two shape paths, a winding hole, both Bezier commands, and six real glyph paths. A malformed source command produced no painted geometry. This is functional evidence, not an independent Office fidelity reference.

## Available surfaces and remaining UI work

This milestone serves the native source API, public `compileNativePptxSlide` renderer, native paint worker, and the existing `NativePptxVector` view of worker results. The legacy `PptxFilePreview` fallback does not interpret evaluated paths; its integration remains an explicit next task. The next geometry milestone adds the data-driven preset catalog and adjustment evaluation on top of this engine.

The renderer checks numeric path coordinates, controls, and evaluated text rectangles against both local and cumulative transform limits. Arc limit checks use a conservative whole-ellipse envelope around the start point, so near-limit shapes can be refused even when their visible arc occupies a smaller region. The browser transport additionally retains its existing ±1,000,000,000 coordinate ceiling and eight-megabyte path-data budget.

## Sequential custom guide qualification

ECMA-376 §20.1.9.11 defines ordered calculation and assignment to guide names; §20.1.10.28 defines the name as a token without a uniqueness constraint. Custom geometry now follows the same source-order reassignment semantics already used by the pinned preset catalog. For example, `a=1/3; old=a; a=a*3; a=a+2` leaves `old=1/3` and `a=3`, with exact rational shadows. Later path and text-rectangle operands use the final binding. This does not enable forward dependency resolution or overwrite built-in frame/angle inputs.

A failed individual assignment changes none of its stored shadows, and a failed source geometry still exposes no partial geometry. Repeated names count as separate operations toward the existing 1,024 combined adjustment/guide limit; rational bits, intermediate magnitude, interval uncertainty, arc qualification and final coordinate budgets are unchanged. The original custom XML and source ownership remain preserved and read-only.

Permanent tests cover earlier aliases, self-reassignment, rational and noncardinal interval results, atomic failed writes, reserved/numeric-name refusal, cross-list adjustment bindings, both source dialects and repeated-name budget exhaustion. A generated PPTX fixture captures successive edge values 250,000, 1,750,000 and 2,000,000 EMU; actual Go/WASM and local browser preview verify these distinct path coordinates. The negative fixture references an unavailable later name and remains refused.
