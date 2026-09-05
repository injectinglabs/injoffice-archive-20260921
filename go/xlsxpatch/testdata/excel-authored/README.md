# Excel-authored XLSX fixture

These are unmodified workbooks published by Microsoft in the OfficeDev
`office-scripts-docs` repository. Their extended properties name Microsoft
Excel as the authoring application. The fixtures are kept byte-for-byte so
tests can distinguish actual Office OOXML from synthetic packages.

## `happy-tree.xlsx`

- Source: <https://github.com/OfficeDev/office-scripts-docs/blob/ad12607e3fb4080fc161e72f4d2b93ce86c29e66/docs/resources/samples/happy-tree.xlsx>
- Upstream commit: `ad12607e3fb4080fc161e72f4d2b93ce86c29e66`
- Upstream Git blob: `c4b91c1b56212df87ae4fba063ab20a66327a203`
- Local SHA-256: `c08f0bb099770a475556af8a78d9d6e46296e9791bf9af56c70571c1fa3cb513`
- License: [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/)
- Changes: none

This Excel Online workbook supplies direct-RGB solid fills without tables,
drawings, formula groups, or external links. Its exact Go extraction is pinned
as `../native-xlsx-v1/valid/excel-authored-happy-tree.json`; TypeScript consumes
that bridge through projection, geometry, decorations, emission, and replay.

## `email-chart-table.xlsx`

- Source: <https://github.com/OfficeDev/office-scripts-docs/blob/ad12607e3fb4080fc161e72f4d2b93ce86c29e66/docs/resources/samples/email-chart-table.xlsx>
- Upstream commit: `ad12607e3fb4080fc161e72f4d2b93ce86c29e66`
- Upstream Git blob: `d934dc31638bf642fdeb389d2b953261ce111a5d`
- Local SHA-256: `1da5a2011f4f1e6f2e4407f148e814cae47e93d97335c6819d582b6cf433e1be`
- License: [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/)
- Changes: none

The workbook exercises shared direct-RGB thin borders and solid fills as well
as tables and other source-authoritative content. Its explicit built-in-range
number-format override is honored as authored. Tests extract and encode/decode
the unchanged workbook while retaining its table/formula appearance refusals.

## `conditional-formatting-samples.xlsx`

- Source: <https://github.com/OfficeDev/office-scripts-docs/blob/ad12607e3fb4080fc161e72f4d2b93ce86c29e66/docs/resources/samples/conditional-formatting-samples.xlsx>
- Upstream commit: `ad12607e3fb4080fc161e72f4d2b93ce86c29e66`
- Upstream Git blob: `5affc0ac28bd74b7877c492a6bb541040fbf2a88`
- Local SHA-256: `fa17f45f47e0766f13b9cbbf9e63b83962ea7852606bfe0b33807fbfbae5ec64`
- License: [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/)
- Changes: none

This workbook provides a second desktop-Excel package with namespace,
worksheet, and style-table features outside the native projection. Tests prove
those features remain explicitly inventoried and that contract encoding stays
deterministic.
