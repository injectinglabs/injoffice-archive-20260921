import { decodeNativeStoredRowGeometryV1, type NativeStoredRowGeometryV1 } from "./nativeStoredRowsPreviewV1.js";
import { snapshotNativePlainData } from "./nativePlainData.js";
/** Read-only source-derived metadata. Never a workbook mutation envelope. */
export interface NativeWorkbookObjectsV1 {
  protocol: "injoffice.xlsx.preview-objects";
  version: 1;
  package_sha256: string;
  tables: NativeTablePreviewV1[];
  charts: NativeChartPreviewV1[];
  row_geometry?: NativeStoredRowGeometryV1[];
}
export interface NativeTablePreviewV1 {
  part: string;
  sheet_part: string;
  name: string;
  ref: string;
  style: string;
  header_rows: number;
  total_rows: number;
  row_stripes: boolean;
  column_stripes: boolean;
  warnings: string[];
  number_formats?: NativeTableNumberFormatV1[];
  border_preview?: NativeTableBorderPreviewV1;
  fill_preview?: {
    totals_bold?: true;
    header: string;
    stripe: string;
    body: string;
    header_font_style_ids: number[];
    fill_style_ids: number[];
  };
}
export interface NativeTableBorderPreviewV1 {
  color: string;
  totals_color: string;
  width_points: 1;
  totals_width_points: 3;
  style_ids: number[];
}
export interface NativeTableNumberFormatV1 {
  ref: string;
  dxf_id: number;
  number_format: string;
  style_ids: number[];
}
export interface NativeChartPreviewV1 {
  part: string;
  type: "col" | "bar" | "unsupported";
  series: NativeChartSeriesPreviewV1[];
  warnings: string[];
}
export interface NativeChartSeriesPreviewV1 {
  name: string;
  values: (number | null)[];
  labels: string[];
}

/** Copy and validate the supplemental projection against the opened package. */
export function decodeNativeWorkbookObjectsV1(
  input: unknown,
  packageSHA256: string,
): NativeWorkbookObjectsV1 {
  input = snapshotNativePlainData(input, { maxDepth: 12, maxNodes: 200000 });
  const fail = (): never => {
    throw new TypeError("Invalid or stale native workbook object preview");
  };
  const obj = (value: unknown, keys: string[]) => {
    if (!value || typeof value !== "object" || Array.isArray(value))
      return fail();
    const o = value as Record<string, unknown>;
    if (
      Object.keys(o).length !== keys.length ||
      keys.some((k) => !Object.hasOwn(o, k))
    )
      return fail();
    return o;
  };
  let textUnits = 0;
  const text = (v: unknown, max = 4096) =>
    typeof v === "string" &&
    v.length <= max &&
    (textUnits += v.length) <= 4 * 1024 * 1024 &&
    !/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(v)
      ? v
      : fail();
  const part = (v: unknown, empty = false) => {
    const s = text(v, 1024);
    if (empty && s === "") return s;
    if (
      !s ||
      s.startsWith("/") ||
      s.includes("\\") ||
      s.split("/").some((p) => !p || p === "." || p === "..")
    )
      return fail();
    return s;
  };
  const list = (v: unknown, max: number) =>
    Array.isArray(v) && v.length <= max ? v : fail();
  const warnings = (v: unknown) => list(v, 32).map((x) => text(x));
  const bit = (v: unknown) => (typeof v === "boolean" ? v : fail());
  const count = (v: unknown) => (v === 0 || v === 1 ? v : fail());
  const hasRows = !!input && typeof input === "object" && Object.hasOwn(input, "row_geometry");
  const value = obj(input, [
    "protocol",
    "version",
    "package_sha256",
    "tables",
    "charts",
    ...(hasRows ? ["row_geometry"] : []),
  ]);
  if (
    value.protocol !== "injoffice.xlsx.preview-objects" ||
    value.version !== 1 ||
    value.package_sha256 !== packageSHA256 ||
    !/^sha256:[a-f0-9]{64}$/.test(packageSHA256)
  )
    return fail();
  let styleIDs = 0,
    formatRegions = 0;
  const tables = list(value.tables, 64).map((v) => {
    const hasBorders =
      !!v && typeof v === "object" && Object.hasOwn(v, "border_preview");
    const hasFormats =
      !!v && typeof v === "object" && Object.hasOwn(v, "number_formats");
    const hasFill =
      !!v && typeof v === "object" && Object.hasOwn(v, "fill_preview");
    const t = obj(v, [
      "part",
      "sheet_part",
      "name",
      "ref",
      "style",
      "header_rows",
      "total_rows",
      "row_stripes",
      "column_stripes",
      "warnings",
      ...(hasFill ? ["fill_preview"] : []),
      ...(hasFormats ? ["number_formats"] : []),
      ...(hasBorders ? ["border_preview"] : []),
    ]);
    let fill_preview: NativeTablePreviewV1["fill_preview"];
    if (hasFill) {
      const hasTotals =
        !!t.fill_preview &&
        typeof t.fill_preview === "object" &&
        Object.hasOwn(t.fill_preview, "totals_bold");
      const p = obj(t.fill_preview, [
        "header",
        "stripe",
        "body",
        "header_font_style_ids",
        "fill_style_ids",
        ...(hasTotals ? ["totals_bold"] : []),
      ]);
      if (
        hasTotals &&
        (p.totals_bold !== true || t.total_rows !== 1 || !hasBorders)
      )
        return fail();
      const color = (v: unknown) =>
        typeof v === "string" && /^#[0-9A-F]{6}$/.test(v) ? v : fail();
      const ids = list(p.header_font_style_ids, 4096).map((v) =>
        typeof v === "number" && Number.isSafeInteger(v) && v >= 0 && v < 65536
          ? v
          : fail(),
      );
      const fillIDs = list(p.fill_style_ids, 4096).map((v) =>
        typeof v === "number" && Number.isSafeInteger(v) && v >= 0 && v < 65536
          ? v
          : fail(),
      );
      styleIDs += ids.length + fillIDs.length;
      if (
        styleIDs > 16384 ||
        new Set(fillIDs).size !== fillIDs.length ||
        new Set(ids).size !== ids.length
      )
        return fail();
      fill_preview = {
        ...(hasTotals ? { totals_bold: true as const } : {}),
        header: color(p.header),
        stripe: color(p.stripe),
        body: color(p.body),
        header_font_style_ids: ids,
        fill_style_ids: fillIDs,
      };
      if (
        t.style !== "TableStyleMedium2" ||
        t.column_stripes !== false ||
        !t.sheet_part
      )
        return fail();
    }
    let border_preview: NativeTableBorderPreviewV1 | undefined;
    if (hasBorders) {
      const b = obj(t.border_preview, [
        "color",
        "totals_color",
        "width_points",
        "totals_width_points",
        "style_ids",
      ]);
      const ids = list(b.style_ids, 4096).map((v) =>
        typeof v === "number" && Number.isSafeInteger(v) && v >= 0 && v < 65536
          ? v
          : fail(),
      );
      styleIDs += ids.length;
      if (
        styleIDs > 16384 ||
        new Set(ids).size !== ids.length ||
        !fill_preview ||
        b.width_points !== 1 ||
        b.totals_width_points !== 3 ||
        typeof b.color !== "string" ||
        !/^#[0-9A-F]{6}$/.test(b.color) ||
        typeof b.totals_color !== "string" ||
        !/^#[0-9A-F]{6}$/.test(b.totals_color)
      )
        return fail();
      border_preview = {
        color: b.color,
        totals_color: b.totals_color,
        width_points: 1,
        totals_width_points: 3,
        style_ids: ids,
      };
    }
    const number_formats = hasFormats
      ? list(t.number_formats, 384).map((v) => {
          const f = obj(v, ["ref", "dxf_id", "number_format", "style_ids"]);
          const ids = list(f.style_ids, 4096).map((v) =>
            typeof v === "number" &&
            Number.isSafeInteger(v) &&
            v >= 0 &&
            v < 65536
              ? v
              : fail(),
          );
          styleIDs += ids.length;
          if (
            styleIDs > 16384 ||
            ++formatRegions > 1024 ||
            new Set(ids).size !== ids.length ||
            typeof f.dxf_id !== "number" ||
            !Number.isSafeInteger(f.dxf_id) ||
            f.dxf_id < 0 ||
            f.dxf_id > 65535
          )
            return fail();
          return {
            ref: text(f.ref, 64),
            dxf_id: f.dxf_id,
            number_format: text(f.number_format),
            style_ids: ids,
          };
        })
      : undefined;
    return {
      part: part(t.part),
      sheet_part: part(t.sheet_part, true),
      name: text(t.name),
      ref: text(t.ref, 64),
      style: text(t.style, 256),
      header_rows: count(t.header_rows),
      total_rows: count(t.total_rows),
      row_stripes: bit(t.row_stripes),
      column_stripes: bit(t.column_stripes),
      warnings: warnings(t.warnings),
      ...(fill_preview ? { fill_preview } : {}),
      ...(number_formats ? { number_formats } : {}),
      ...(border_preview ? { border_preview } : {}),
    };
  });
  let points = 0;
  const charts = list(value.charts, 64).map((v): NativeChartPreviewV1 => {
    const c = obj(v, ["part", "type", "series", "warnings"]);
    if (c.type !== "col" && c.type !== "bar" && c.type !== "unsupported")
      return fail();
    const series = list(c.series, 32).map((v) => {
      const s = obj(v, ["name", "values", "labels"]);
      const values = list(s.values, 1024).map((v) => {
        if (++points > 65536) return fail();
        return v === null || (typeof v === "number" && Number.isFinite(v))
          ? v
          : fail();
      });
      return {
        name: text(s.name),
        values,
        labels: list(s.labels, 1024).map((x) => text(x, 256)),
      };
    });
    return {
      part: part(c.part),
      type: c.type,
      series,
      warnings: warnings(c.warnings),
    };
  });
  if (
    tables.length + charts.length > 64 ||
    new Set([...tables, ...charts].map((v) => v.part)).size !==
      tables.length + charts.length
  )
    return fail();
  return {
    protocol: "injoffice.xlsx.preview-objects",
    version: 1,
    package_sha256: packageSHA256,
    tables,
    charts,
    ...(hasRows ? { row_geometry: decodeNativeStoredRowGeometryV1(value.row_geometry) } : {}),
  };
}

export interface NativeCachedChartMarkV1 {
  series: number;
  point: number;
  value: number;
  x: number;
  y: number;
  width: number;
  height: number;
}
/** Normalized chart data marks only: no Office layout/style equivalence claim. */
export function layoutNativeCachedChartV1(
  chart: NativeChartPreviewV1,
):
  | {
      marks: NativeCachedChartMarkV1[];
      minimum: number;
      maximum: number;
      baseline: number;
    }
  | undefined {
  if (
    chart.type === "unsupported" ||
    chart.series.length === 0 ||
    chart.series.length > 32 ||
    chart.series.some((s) => s.values.length > 1024) ||
    chart.series.reduce((sum, s) => sum + s.values.length, 0) > 1024
  )
    return undefined;
  const values = chart.series.flatMap((s) =>
    s.values.filter((v): v is number => v !== null),
  );
  if (!values.length || values.some((v) => !Number.isFinite(v)))
    return undefined;
  const minimum = Math.min(0, ...values),
    maximum = Math.max(0, ...values);
  const extent = maximum - minimum;
  if (!Number.isFinite(extent)) return undefined;
  const scale = extent || 1,
    baseline = maximum / scale;
  const count = Math.max(...chart.series.map((s) => s.values.length));
  if (!count) return undefined;
  const marks: NativeCachedChartMarkV1[] = [];
  chart.series.forEach((s, series) =>
    s.values.forEach((value, point) => {
      if (value === null) return;
      const width = 0.8 / count / chart.series.length,
        x = (point + 0.1) / count + series * width,
        y = (maximum - Math.max(0, value)) / scale,
        height = Math.abs(value) / scale;
      marks.push(
        chart.type === "col"
          ? { series, point, value, x, y, width, height }
          : {
              series,
              point,
              value,
              x: (Math.min(0, value) - minimum) / scale,
              y: x,
              width: height,
              height: width,
            },
      );
    }),
  );
  return {
    marks,
    minimum,
    maximum,
    baseline: chart.type === "col" ? baseline : -minimum / scale,
  };
}
