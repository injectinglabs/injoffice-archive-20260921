import type { NativeWorkbookObjectsV1 } from "./nativeObjectsPreviewV1.js";

export interface NativeStoredRowV1 {
  row: number;
  height_points: number;
  hidden: boolean;
}
export interface NativeStoredRowGeometryV1 {
  root_policy?: 'x14ac-descent-only-v1';
  sheet_part: string;
  rows: NativeStoredRowV1[];
  warnings: string[];
}

/** Bounded decoder for optional source-stored row dimensions, not autofit. */
export function decodeNativeStoredRowGeometryV1(
  value: unknown,
): NativeStoredRowGeometryV1[] {
  const fail = (): never => {
    throw new TypeError("Invalid stored row geometry");
  };
  const exact = (input: unknown, keys: string[]) => {
    if (!input || typeof input !== "object" || Array.isArray(input))
      return fail();
    const object = input as Record<string, unknown>;
    if (
      Object.keys(object).length !== keys.length ||
      keys.some((key) => !Object.hasOwn(object, key))
    )
      return fail();
    return object;
  };
  if (!Array.isArray(value) || value.length > 64) return fail();
  const seen = new Set<string>();
  return value.map((input) => {
    const hasPolicy = !!input && typeof input === 'object' && Object.hasOwn(input,'root_policy');
    const sheet = exact(input, ["sheet_part", "rows", "warnings", ...(hasPolicy?['root_policy']:[])]);
    if(hasPolicy && sheet.root_policy !== 'x14ac-descent-only-v1') return fail();
    const part = sheet.sheet_part;
    if (
      typeof part !== "string" ||
      part.length > 1024 ||
      !part ||
      part.startsWith("/") ||
      part.includes("\\") ||
      part.split("/").some((v) => !v || v === "." || v === "..") ||
      /[\x00-\x1f]/.test(part) ||
      seen.has(part)
    )
      return fail();
    seen.add(part);
    if (
      !Array.isArray(sheet.rows) ||
      (sheet.rows.length !== 0 && sheet.rows.length !== 32) ||
      !Array.isArray(sheet.warnings) ||
      sheet.warnings.length < 1 ||
      sheet.warnings.length > 4
    )
      return fail();
    const rows = sheet.rows.map((input, index) => {
      const row = exact(input, ["row", "height_points", "hidden"]);
      if (
        row.row !== index ||
        typeof row.height_points !== "number" ||
        !Number.isFinite(row.height_points) ||
        row.height_points < 0 ||
        row.height_points > 409 ||
        typeof row.hidden !== "boolean" ||
        (row.height_points === 0 && !row.hidden)
      )
        return fail();
      return {
        row: index,
        height_points: row.height_points,
        hidden: row.hidden,
      };
    });
    const warnings = sheet.warnings.map((value) =>
      typeof value === "string" &&
      value.length > 0 &&
      value.length <= 4096 &&
      !/[\x00-\x1f]/.test(value)
        ? value
        : fail(),
    );
    return { sheet_part: part, rows, warnings, ...(hasPolicy?{root_policy:'x14ac-descent-only-v1' as const}:{}) };
  });
}

/** Package/sheet joined stored-height lookup. Unknown geometry keeps host sizing. */
export function nativeStoredRowPreviewV1(
  objects: NativeWorkbookObjectsV1,
  revision: string,
  sheetPart: string,
  row: number,
): NativeStoredRowV1 | undefined {
  if (
    objects.package_sha256 !== revision ||
    !Number.isSafeInteger(row) ||
    row < 0 ||
    row >= 32
  )
    return undefined;
  const matches = objects.row_geometry?.filter(
    (sheet) => sheet.sheet_part === sheetPart,
  );
  if (matches?.length !== 1) return undefined;
  const result = matches[0]!.rows[row];
  return result?.row === row ? { ...result } : undefined;
}
