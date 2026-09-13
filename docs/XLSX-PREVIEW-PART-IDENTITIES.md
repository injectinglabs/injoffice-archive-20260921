# Exact Unicode package-part identities in supplemental previews

Native XLSX extraction and inspection preserve routed package-part names, including Unicode and spaces. Supplemental decoders for page settings, print areas, print titles, conditional fills and drawing objects now accept those same identities instead of requiring ASCII. A worksheet such as `Sheets/预算.xml` can therefore pass through the public object decoder and retain its source join.

The internal validator preserves the input string exactly. It does not normalize Unicode, decode percent escapes, trim spaces, resolve paths or make network/filesystem requests. Composed and decomposed Unicode spellings remain distinct identities. Source hash, worksheet ID and exact source-part equality checks are unchanged.

Paths remain limited to 1,024 UTF-16 code units and reject empty values, leading slash, backslash, empty slash-separated segments, `.` and `..` segments, and C0/DEL controls. The explicitly allowed empty `drawing_part` sentinel for unsupported drawings remains supported; worksheet and chart parts cannot be empty. This retains the supplemental decoders' prior control-character refusals while aligning their Unicode path acceptance with the base object decoder.

This change only corrects transport validation. It does not broaden the supported spreadsheet features, establish formula-cache freshness or claim Excel print fidelity. Tests exercise safe and unsafe identities across all five supplemental decoders, drawing sentinels and the complete public envelope. Actual Go-produced source/inspection envelopes and browser evidence remain in the local fidelity corpus, outside git and CI.
