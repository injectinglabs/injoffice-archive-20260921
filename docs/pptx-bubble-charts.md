# Source-bound 2D bubble charts

Native extraction and the supplied-font slide preview support qualified positive/zero-size 2D bubble charts from explicit literal arrays or authoritative embedded-workbook references. Literal rendering requires `literalBubblePreview:true`; workbook rendering uses the separately admitted `workbook-bubble-v1` data record and existing workbook preview mode. The playground source-chart and embedded-workbook options enable these paths explicitly. Source ownership stays read-only.

The first source profile requires explicit RGB/no-line paint, non-inverted colors, no 3D, no negative bubbles, explicit size representation and bubble scale, paired explicit numeric axes, transparent chart/plot, and plotVisOnly=false. Arrays are dense, indexed and bounded to 256 points per series and 16 series. Series retain XML sequence and original index/order metadata; order must form a complete contiguous permutation. Points retain index order. References retain their formula and cache-presence provenance; caches never populate values. A workbook value must pass the existing actual-XLSX decoder/resolver before rendering. Missing/error/formula/numeric-text values remain refused.

## Sizing and numerical policy

`plot-minor-radius-v1` deliberately defines the host's default radius: at source bubbleScale 100%, the largest size across all admitted nonnegative series has radius one tenth of the smaller plot dimension. The source percentage multiplies radius. In area mode, radius is proportional to the square root of size/globalMaximum; in width mode, it is proportional to size/globalMaximum. This preserves relative source area or width semantics, but does not claim PowerPoint's default-radius algorithm.

The ISO-derived [bubbleScale remarks](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.bubblescale?view=openxml-2.11.2) define percentage of a default size without specifying that default radius. The [size representation enum](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.sizerepresentsvalues?view=openxml-3.0.1) distinguishes area and width. Strict percentage syntax reuses the qualified existing chart percentage parser; Transitional also accepts its integer spelling.

Existing 4096-bit exact rational arithmetic maps XY centers and size ratios. Integer square root plus an exact midpoint comparison rounds an area-mode radius directly to nearest EMU, with at most 0.5 EMU error. Width radius and signed centers quantize once. Exact conservative envelope comparisons remove wholly off-plot points before any Number conversion, so huge source coordinates cannot overflow output. A positive radius below the EMU threshold may disappear; no minimum visible size is invented. Scale 0 and all-zero size datasets emit no bubbles.

Each bubble uses one closed path with two semicircular arc commands, four commands total. Maximum 4096 bubbles imply 16384 generated commands, without expanding the generic 512-command per-path budget. The compiler clips entire bubble paths to the plot rectangle: a center outside the plot can still have visible ink. XML series sequence then point index is an explicit host painter-order policy, not a universal Office overlap claim.

Required later coverage remains negative-bubble rendering, 3D, labels, source inherited/theme/default styles, mixed data origins, blank/formula behavior and any further host layout qualification. This connected profile does not close those remaining rows.

## Integration bounds and authority

The compiler qualifies the conservative union of all complete circle hulls once in world coordinates, including ink outside the plot before clipping. This preserves the existing aggregate affine-operation budget even at 4096 admitted points; generic path and node/depth budgets remain active. Visible axes survive zero scale and all-zero sizes. Both numeric axes use the existing exact supplied-font label profile and measured host margins; unsupported font/script/layout falls back rather than silently removing labels.

Workbook sources retain x/y/size references, raw numeric lexemes, cache-presence and exact package/resource identity. Negative resolved sizes are refused after the actual XLSX decoder. The existing reference/value/transport budgets are unchanged. The shared geometry uses a temporary structural validation record only; it never publishes or admits workbook values as literal data.
