# DrawingML evaluated geometry

The custom geometry engine evaluates DrawingML in the Go source parser. The native boundary carries numeric paths and a text rectangle, not XML or an executable guide language. Source geometry remains preserve-only; shape mutation does not rewrite custom paths.

This is the first integration milestone in the geometry completion plan. It does not establish Office visual parity or complete the preset catalog, adjustable presets, arbitrary rotations, shaded path fills, or interactive geometry handles.

## Evaluation

The engine supports the seventeen ECMA-376 guide operations: `*/`, `+-`, `+/`, `?:`, `abs`, `at2`, `cat2`, `cos`, `max`, `min`, `mod`, `pin`, `sat2`, `sin`, `sqrt`, `tan`, and `val`. It evaluates adjustment values followed by ordered shape guides. Built-in frame, divisor, and angle guides follow the finite list in §20.1.10.56. Names cannot shadow built-ins, earlier guides, or numeric operands. Forward references and undefined numeric domains refuse the whole geometry.

Calculations retain floating-point precision until coordinates cross the native boundary, where they round once to integer EMU. Path `w` and `h` scale each axis independently. A missing axis extent uses the shape frame. Text rectangle guides use the frame coordinate system, independently of path coordinate spaces; missing `rect` uses the full frame.

Paths support move, line, quadratic and cubic Bezier curves, elliptical arcs, and close. Close restores the current point to the active subpath origin. Multiple subpaths retain their original order and direction for nonzero winding, including holes. Multiple paths preserve paint order. This milestone accepts `norm`/`none` fill and Boolean stroke, refusing other fill modes. Both Boolean `extrusionOk` values are nonvisual permission metadata; applied scene/shape 3D remains refused by source qualification. The specification prose and schema disagree about its default, which does not affect this two-dimensional paint policy.

## Elliptical arcs

DrawingML angles describe geometric rays, not SVG ellipse parameters. For radii `rx`, `ry` and DrawingML angle θ, the ellipse parameter is:

```
t = atan2(rx × sin(θ), ry × cos(θ))
point = (rx × cos(t), ry × sin(t))
```

The current point anchors the start of the ellipse. Endpoint calculation takes place before path-axis scaling. Signed sweeps preserve clockwise or counterclockwise winding. Sweeps up to one revolution split into at most two segments of at most half a revolution; this avoids the SVG coincident-endpoint full-circle omission.

After rounding, an exact integer predicate checks that the endpoint chord fits the transported radii. A shape that would require SVG to silently enlarge its radii refuses instead. Collapsed radii also refuse. These checks include tiny arcs and extreme aspect ratios.

## Limits and refusal

The engine accepts at most 1,024 adjustment/shape guides combined, 128 paths, 512 emitted commands per path, and 8,192 emitted commands overall. Numeric values must remain finite and within the native safe-integer envelope. The renderer must also honor its smaller configured coordinate and aggregate paint limits.

Unknown attributes or child elements, invalid point counts, commands before the first move, unresolved operands, duplicate guide names, nonempty unqualified handles/connections, and unsupported path paint clauses refuse the entire geometry. No successfully parsed prefix is exposed as a complete shape.

## Sources and evidence

The governing source is [ECMA-376, Part 1](https://ecma-international.org/publications-and-standards/standards/ecma-376/): §20.1.9 custom geometry/path elements, §20.1.9.11 guide formulas, and §20.1.10.56 built-in guides. Non-square angle conversion was independently checked against [Apache POI ArcToCommand](https://github.com/apache/poi/blob/trunk/poi/src/main/java/org/apache/poi/sl/draw/geom/ArcToCommand.java), whose conversion explicitly distinguishes geometric OOXML angles from ellipse parameter angles. These references are evidence, not copied implementation code.

Dedicated Go tests exercise all operators, order and budget guards, both Bezier controls, scaled path coordinates, non-square arcs, negative/full sweeps, close/current-point behavior, atomic XML refusal, and exact rounded-arc radius consistency. The connected source, JSON, mutation, renderer, and paint-worker tests cover the complete path through the public API. An external generated fixture was extracted by the actual browser Go/WASM runtime, compiled with DejaVu Sans, and displayed by the existing `NativePptxVector` component in Chrome: two shape paths, a winding hole, both Bezier commands, and six real glyph paths. A malformed source command produced no painted geometry. This is functional evidence, not an independent Office fidelity reference.

## Available surfaces and remaining UI work

This milestone serves the native source API, public `compileNativePptxSlide` renderer, native paint worker, and the existing `NativePptxVector` view of worker results. The legacy `PptxFilePreview` fallback does not interpret evaluated paths; its integration remains an explicit next task. The next geometry milestone adds the data-driven preset catalog and adjustment evaluation on top of this engine.

The renderer checks numeric path coordinates, controls, and evaluated text rectangles against both local and cumulative transform limits. Arc limit checks use a conservative whole-ellipse envelope around the start point, so near-limit shapes can be refused even when their visible arc occupies a smaller region. The browser transport additionally retains its existing ±1,000,000,000 coordinate ceiling and eight-megabyte path-data budget.
