# Native chart source series sequence

Literal clustered bar/column, line/scatter, area and bubble profiles now retain
`c:ser` XML sequence. Embedded-workbook bar/column, line/scatter and bubble
inspection preserves that same sequence through actual XLSX resolution. Original
series indices and order values remain unchanged; order values must be canonical
integers forming a unique complete permutation from zero through series count
minus one. Duplicate indices, duplicate/missing/out-of-range orders and sparse
series remain refused. Numeric source spellings and reference/cache provenance
remain associated with the original series. Point indices and point overrides
keep their existing independent authority.

Cluster slots, connected-series overpaint, bubble overpaint and area accumulation
follow XML sequence as an explicit host policy. Bubble sizing continues to use
the permutation-invariant global maximum. This correction does not claim Office
cluster/overlap parity. Existing stacked bar/line records retain their source
sequence and metadata unchanged. Chart previews remain opt-in; package ownership,
private workbook admission, resource limits and unsupported-profile refusal are
unchanged. Pie/doughnut remain the existing single-series contract.

`chartSeriesOrder.test.ts` in native/render covers semantic permutations, negative
zero, sparse arrays, XY point order, cluster slots, bubble sizing, area accumulation
and workbook reference joins for bar, line, scatter and bubble. The worker test
generates nine real PPTX fixture families, extracts their actual embedded XLSX
bytes, and compiles both unlabeled and supplied-font labeled previews while
checking source identity, XML painter sequence and default-off behavior. Go tests
exercise strict and transitional sources and keep existing area fixture bytes
unchanged.

After building the workspace, `node scripts/smoke-pptx-series-order-browser.mjs`
runs the nine source packages through the actual helper and browser and writes
source hashes, replay JSON and screenshots. Browser execution remains unverified
in this restart environment: Chrome startup was denied before a target existed.
No new Office reference exports were performed; host-policy disclosures remain
required even after browser smoke succeeds.
