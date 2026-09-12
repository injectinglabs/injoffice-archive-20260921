import { expect, it } from "vitest";
import {
  decodeNativeWorkbookObjectsV1,
  type NativeWorkbookObjectsV1,
} from "./nativeObjectsPreviewV1.js";
import { nativeTableTotalsTextPreview } from "./nativeTableTotalsTextPreview.js";

const revision = `sha256:${"a".repeat(64)}`;
const objects: NativeWorkbookObjectsV1 = {
  protocol: "injoffice.xlsx.preview-objects",
  version: 1,
  package_sha256: revision,
  charts: [],
  tables: [
    {
      part: "xl/tables/t.xml",
      sheet_part: "xl/worksheets/s.xml",
      name: "Original",
      ref: "B2:D5",
      style: "TableStyleMedium2",
      header_rows: 1,
      total_rows: 1,
      row_stripes: true,
      column_stripes: false,
      warnings: [],
      fill_preview: {
        header: "#156082",
        stripe: "#C0E6F5",
        body: "#FFFFFF",
        header_font_style_ids: [0],
        fill_style_ids: [0],
        totals_bold: true,
      },
      border_preview: {
        color: "#44B3E1",
        totals_color: "#156082",
        width_points: 1,
        totals_width_points: 3,
        style_ids: [0],
      },
    },
  ],
};
const preview = (input = objects, row = 4, column = 1, style = 0) =>
  nativeTableTotalsTextPreview(
    input,
    revision,
    "xl/worksheets/s.xml",
    row,
    column,
    style,
  );

it("qualifies only the source-bound totals row and default font IDs", () => {
  expect(preview()).toBe(true);
  expect(preview(objects, 4, 3)).toBe(true);
  for (const [row, column, style] of [
    [3, 1, 0],
    [4, 0, 0],
    [4, 4, 0],
    [4, 1, 7],
    [-1, 1, 0],
    [NaN, 1, 0],
  ])
    expect(preview(objects, row, column, style)).toBe(false);
  expect(
    nativeTableTotalsTextPreview(
      objects,
      "stale",
      "xl/worksheets/s.xml",
      4,
      1,
      0,
    ),
  ).toBe(false);
  expect(
    nativeTableTotalsTextPreview(objects, revision, "other", 4, 1, 0),
  ).toBe(false);
  expect(
    preview({
      ...objects,
      tables: [
        ...objects.tables,
        { ...objects.tables[0]!, part: "xl/tables/u.xml" },
      ],
    }),
  ).toBe(false);
  expect(
    preview({ ...objects, tables: [{ ...objects.tables[0]!, total_rows: 0 }] }),
  ).toBe(false);
});

it("validates exact totals qualification without expanding the existing style budget", () => {
  expect(decodeNativeWorkbookObjectsV1(objects, revision)).toEqual(objects);
  for (const mutate of [
    (v: any) => (v.tables[0].fill_preview.totals_bold = false),
    (v: any) => (v.tables[0].total_rows = 0),
    (v: any) => delete v.tables[0].border_preview,
  ]) {
    const input = structuredClone(objects);
    mutate(input);
    expect(() => decodeNativeWorkbookObjectsV1(input, revision)).toThrow();
  }
});
