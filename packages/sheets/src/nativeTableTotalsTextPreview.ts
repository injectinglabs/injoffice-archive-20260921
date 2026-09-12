import type { NativeWorkbookObjectsV1 } from "./nativeObjectsPreviewV1.js";

/** Read-only bold totals text, using source-qualified default-font authority.
 * Explicit cell fonts, stale packages and overlapping tables are never replaced.
 */
export function nativeTableTotalsTextPreview(
  objects: NativeWorkbookObjectsV1,
  revision: string,
  sheetPart: string,
  row: number,
  column: number,
  styleID: number,
): boolean {
  if (
    objects.package_sha256 !== revision ||
    !Number.isSafeInteger(row) ||
    !Number.isSafeInteger(column)
  )
    return false;
  const matches = objects.tables.flatMap((table) => {
    if (table.sheet_part !== sheetPart) return [];
    const range =
      /^([A-Z]{1,3})([1-9][0-9]{0,6}):([A-Z]{1,3})([1-9][0-9]{0,6})$/.exec(
        table.ref,
      );
    if (!range) return [];
    const col = (s: string) =>
      [...s].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
    const left = col(range[1]!),
      right = col(range[3]!);
    const top = Number(range[2]) - 1,
      bottom = Number(range[4]) - 1;
    if (
      left > right ||
      top > bottom ||
      right >= 16384 ||
      bottom >= 1048576 ||
      row < top ||
      row > bottom ||
      column < left ||
      column > right
    )
      return [];
    return [{ table, bottom }];
  });
  if (matches.length !== 1) return false;
  const { table, bottom } = matches[0]!;
  return (
    row === bottom &&
    table.total_rows === 1 &&
    table.fill_preview?.totals_bold === true &&
    table.fill_preview.header_font_style_ids.includes(styleID)
  );
}
