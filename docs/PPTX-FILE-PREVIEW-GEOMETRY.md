# Local PPTX file preview geometry

`PptxFilePreview` now renders supported native groups, evaluated custom/preset paths, and source shape transforms through the same public renderer as the supplied-font native preview. The local view still labels its text as an approximation: browser fonts, wrapping, line heights and glyph ink are not qualified native text layout or a PowerPoint-equivalence claim. Editing authority is unchanged.

## Geometry and source ownership

The browser validates the original native deck, creates a private clone, and empties paragraph arrays before calling `compileNativePptxSlide`. Source geometry, text-body metadata, compatibility, identifiers, resource bindings and passthrough anchors are retained. The projection is never returned as a source-authoritative deck or passed to a mutator. Original paragraphs remain in the original deck and are used only for the labeled DOM text overlay.

The compiler currently requires a nonempty font manifest even for geometry. The adapter therefore declares one explicitly unavailable host resource with no bytes, digest or metrics. Its resolve/load/shape methods throw if invoked. Successful geometry compilation performs no font resolution and supplies no synthetic glyph metrics or font evidence.

`paintSlideRenderTree` supplies ordered path paint, including the existing deterministic relative-tone policy and per-path stroke flags. The component uses the compiled render tree hierarchy, rational affine conversion, text region and orientation transforms. No DrawingML guide evaluator, preset catalog, group-affine interpreter or separate shade equations are implemented in the UI. Both source leaf-axis hierarchy and the existing authored conventional-group semantics remain those of the public compiler.

All move/line/quadratic/cubic/elliptical-arc/close commands survive into SVG; primitive rectangles, rounded rectangles and ellipses use their corresponding SVG elements. Every length and matrix translation uses points rather than large EMU browser coordinates. Arc flags/angles, matrix linear terms, stroke miter ratios and image crop percentages retain their units. Original raster resources are selected by the compiled asset ID, with the existing bounded PNG/JPEG-only data URI policy.

## Clips and text boundaries

Compiled node clips remain active, and the slide SVG clips to its viewport. Pictures whose preset outline was evaluated through the catalog (`pptx.picture-geometry-preview`) clip to that outline path; the exact `roundRect` clip and the frame rectangle behave as before. Evaluated callout geometry is not clipped to its nominal shape frame. Legacy text without text-body metadata remains clipped by its foreign-object frame; text with metadata uses the compiled text rectangle and insets exactly once, followed by the existing source orientation transform. Its DOM glyph metrics remain approximate.

An empty projection cannot establish the original vertical text/script/list admission. The adapter checks original vertical content before removing runs and keeps unsupported content unavailable. Existing source refusals and invalid contracts remain unavailable; unavailable native text-body metadata is not promoted into a qualified text claim. New complex table/chart group qualification, upright text and separate text-body rotation remain governed by the native source/compiler boundary. Tables still have an explicit local-preview placeholder; charts use only an existing embedded raster preview in this local view.

## Bounds and lifecycle

The native validator and public renderer retain their normal coordinate, path, depth, authority and affine-error budgets. Local display traversal counts groups and descendants toward 500 objects, preserves order, and reports the exact omitted count. Original topology is unchanged. Invalid source data fails validation before projection. Async deck/slide generations are bound to the original object and selected index, so stale results cannot paint a newly selected source.

Permanent tests cover immutable source/text, independent Annex nonuniform-group coordinates, reflected text bounds/insets, Q/C/A and all six fill modes, stroke flags, unclipped out-of-frame callouts, vertical-content refusal, nested 500-object omission, escaped DOM text, raster crop percentages, and paragraph styles. External browser evidence uses actual Go/WASM extraction and the real async component, including source arbitrary transforms, a non-square ring/hole with quadratic/cubic paths, source bevel multipath tones with slide navigation, refused custom source, and visible-inside/hidden-outside viewport controls. Generated functional fixtures are not Office-authored visual references.
