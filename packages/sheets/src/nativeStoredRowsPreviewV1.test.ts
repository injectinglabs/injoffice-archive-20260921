import { expect, it } from "vitest";
import {
  decodeNativeWorkbookObjectsV1,
  type NativeWorkbookObjectsV1,
} from "./nativeObjectsPreviewV1.js";
import { nativeStoredRowPreviewV1 } from "./nativeStoredRowsPreviewV1.js";

const revision = `sha256:${"a".repeat(64)}`;
const input: NativeWorkbookObjectsV1 = {
  protocol: "injoffice.xlsx.preview-objects",
  version: 1,
  package_sha256: revision,
  tables: [],
  charts: [],
  row_geometry: [
    {
      sheet_part: "xl/worksheets/s.xml",
      rows: Array.from({ length: 32 }, (_, row) => ({
        row,
        height_points: row === 1 ? 24 : 14.4,
        hidden: row === 2,
      })),
      warnings: ["Stored sizes, not automatic fitting."],
    },
  ],
};
it("decodes bounded source rows and preserves heights, visibility and revision joins", () => {
  expect(decodeNativeWorkbookObjectsV1(input, revision)).toEqual(input);
  expect(
    nativeStoredRowPreviewV1(input, revision, "xl/worksheets/s.xml", 0),
  ).toEqual({ row: 0, height_points: 14.4, hidden: false });
  expect(
    nativeStoredRowPreviewV1(input, revision, "xl/worksheets/s.xml", 2)?.hidden,
  ).toBe(true);
  for (const row of [-1, 32, NaN, 1.2])
    expect(
      nativeStoredRowPreviewV1(input, revision, "xl/worksheets/s.xml", row),
    ).toBeUndefined();
  expect(
    nativeStoredRowPreviewV1(input, "stale", "xl/worksheets/s.xml", 0),
  ).toBeUndefined();
  expect(nativeStoredRowPreviewV1(input, revision, "other", 0)).toBeUndefined();
  expect(
    nativeStoredRowPreviewV1(
      {
        ...input,
        row_geometry: [...input.row_geometry!, ...input.row_geometry!],
      },
      revision,
      "xl/worksheets/s.xml",
      0,
    ),
  ).toBeUndefined();
  const unavailable = {
    ...input,
    row_geometry: [{ ...input.row_geometry![0]!, rows: [] }],
  };
  expect(decodeNativeWorkbookObjectsV1(unavailable, revision)).toEqual(
    unavailable,
  );
  expect(
    nativeStoredRowPreviewV1(unavailable, revision, "xl/worksheets/s.xml", 0),
  ).toBeUndefined();
});
it("rejects unknown fields, unbounded sizes, duplicate sheets and missing warnings", () => {
  for (const mutate of [
    (v: any) => (v.row_geometry[0].rows[0].height_points = Infinity),
    (v: any) => (v.row_geometry[0].rows[0].height_points = 410),
    (v: any) => (v.row_geometry[0].rows[0].row = 2),
    (v: any) => v.row_geometry[0].rows.push(v.row_geometry[0].rows[0]),
    (v: any) => (v.row_geometry[0].rows[0].extra = 1),
    (v: any) => v.row_geometry.push(v.row_geometry[0]),
    (v: any) => (v.row_geometry[0].warnings = []),
    (v: any) => (v.row_geometry[0].rows[0].height_points = 0),
  ]) {
    const value = structuredClone(input);
    mutate(value);
    expect(() => decodeNativeWorkbookObjectsV1(value, revision)).toThrow();
  }
});
