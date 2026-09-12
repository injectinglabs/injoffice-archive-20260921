import {DOCX_LEGACY_TABLE_ORIGIN_WARNING} from './nativeLegacyTableOriginV1.js'
import {
  decodeNativeDocxDocument,
  type NativeDocxDocumentV1,
} from "./nativeContract.js";
import {
  decodeNativeDocxResolvedLayout,
  type NativeDocxResolvedLayoutInputV1,
} from "./nativeResolvedLayout.js";
import {
  DOCX_AUTO_BORDER_POLICY,
  DOCX_AUTO_BORDER_WARNING,
  validNativeDocxAutomaticBorderEvidenceV1,
  type NativeDocxAutomaticBorderDiagnosticV1,
  type NativeDocxAutomaticBorderEvidenceV1,
} from "./nativeAutomaticBorderEvidenceV1.js";
import {
  decodeNativeDocxPagePaintV1,
  preflightWire,
} from "./nativePagePaintWireV1.js";
import type { NativeDocxPagePaintV1 } from "./nativePagePaintV1.js";
import {
  decodeNativeDocxApproximationEligibilityV1,
  DOCX_APPROXIMATE_LINE_BOX_WARNING,
  type NativeDocxApproximationEligibilityV1,
} from "./nativeApproximationV1.js";
import { DOCX_ABSENT_FONT_SIZE_WARNING, validNativeDocxApproximatedFontSizesV1, type NativeDocxApproximatedFontSizeV1 } from './nativeAbsentFontSizeV1.js'

export const DOCX_AUTO_BORDER_PREVIEW_PROTOCOL =
  "injoffice.docx.auto-border-preview" as const;
export interface NativeDocxAutomaticBorderPreviewV1 {
  protocol: typeof DOCX_AUTO_BORDER_PREVIEW_PROTOCOL;
  version: 1;
  fidelity: "approximate";
  read_only: true;
  policy: typeof DOCX_AUTO_BORDER_POLICY;
  page_background_rgb: "FFFFFF";
  source: { document_id: string; revision: string; package_sha256: string };
  approximated_render_properties: {
    table_id: string;
    evidence: NativeDocxAutomaticBorderEvidenceV1;
  }[];
  source_diagnostics: {
    document: NativeDocxAutomaticBorderDiagnosticV1[];
    resolved: NativeDocxAutomaticBorderDiagnosticV1[];
  };
  reasons: string[];
  legacy_eligibility?: NativeDocxApproximationEligibilityV1;
  approximated_font_sizes?: NativeDocxApproximatedFontSizeV1[];
  status: NativeDocxPagePaintV1["status"];
  pages: NativeDocxPagePaintV1["pages"];
  resources: NativeDocxPagePaintV1["resources"];
  diagnostics: NativeDocxPagePaintV1["diagnostics"];
  rendering_provenance: NativeDocxPagePaintV1["provenance"];
}

const identity = (value: NativeDocxAutomaticBorderDiagnosticV1) =>
  JSON.stringify([value.code, value.scope_id, value.part_name, value.path]);
function sourceDiagnostics(
  document: NativeDocxDocumentV1,
  resolved: NativeDocxResolvedLayoutInputV1,
): NativeDocxAutomaticBorderPreviewV1["source_diagnostics"] {
  return {
    document: document.unsupported.map((value) => ({
      code: value.code,
      scope_id: value.scope_id,
      part_name: value.anchor?.part_name ?? "",
      path: value.anchor?.path ?? "",
    })),
    resolved: resolved.diagnostics.map((value) => ({
      code: value.code,
      scope_id: value.scope_id,
      part_name: value.part_name ?? "",
      path: value.path ?? "",
    })),
  };
}

/** @internal Private rendering projection. Never expose as a strict prepared artifact. */
export function projectNativeDocxAutomaticBordersV1(
  documentValue: unknown,
  resolvedValue: unknown,
) {
  const decodedDocument = decodeNativeDocxDocument(documentValue);
  const decodedResolved = decodeNativeDocxResolvedLayout(resolvedValue);
  if (!decodedDocument.ok || !decodedResolved.ok)
    throw new TypeError(
      "Automatic-border preview requires valid original source models",
    );
  const original = decodedDocument.value,
    originalResolved = decodedResolved.value;
  if (
    original.document_id !== originalResolved.document_id ||
    original.revision !== originalResolved.revision ||
    original.source.main_part !== originalResolved.source_parts.main_part
  )
    throw new TypeError("Automatic-border source identity does not exact-join");
  const allParagraphs = original.body.blocks.flatMap((block) =>
    block.paragraph
      ? [block.paragraph]
      : (block.table?.rows.flatMap((row) =>
          row.cells.flatMap((cell) => cell.paragraphs),
        ) ?? []),
  );
  if (
    original.headers.length ||
    original.footers.length ||
    original.notes.length ||
    original.comment_stories.length ||
    allParagraphs.some((paragraph) =>
      paragraph.runs.some((run) => run.drawing !== undefined),
    )
  )
    throw new TypeError(
      "Automatic-border backgrounds exclude other stories and drawings in this policy",
    );
  const document = structuredClone(original),
    resolved = structuredClone(originalResolved);
  const diagnostics = sourceDiagnostics(original, originalResolved);
  const originalKeys = new Set(
    [...diagnostics.document, ...diagnostics.resolved].map(identity),
  );
  const covered = new Set<string>();
  const facts: NativeDocxAutomaticBorderPreviewV1["approximated_render_properties"] =
    [];
  let totalCells = 0;
  for (const entry of resolved.tables) {
    const evidence = entry.automatic_border_preview;
    if (!evidence) continue;
    const matches = document.body.blocks.flatMap((block) =>
      block.table?.id === entry.table_id ? [block.table] : [],
    );
    if (
      matches.length !== 1 ||
      facts.length >= 1000 ||
      evidence.package_sha256 !== document.source.package_sha256
    )
      throw new TypeError("Automatic-border table/package does not exact-join");
    const table = matches[0]!;
    if (table.borders || entry.borders)
      throw new TypeError(
        "Automatic-border evidence cannot overwrite explicit source borders",
      );
    if (
      entry.cell_shading_rgb !== undefined &&
      entry.cell_shading_rgb !== "FFFFFF"
    )
      throw new TypeError(
        "Automatic borders require qualified white table fill",
      );
    const cells = table.rows.flatMap((row) => row.cells);
    totalCells += cells.length;
    if (
      totalCells > 100_000 ||
      JSON.stringify(cells.map((cell) => cell.id)) !==
        JSON.stringify(evidence.cell_ids) ||
      cells.some(
        (cell) =>
          cell.borders !== undefined ||
          cell.grid_span !== 1 ||
          cell.vertical_merge !== "none" ||
          (cell.shading_rgb !== undefined && cell.shading_rgb !== "FFFFFF"),
      )
    )
      throw new TypeError(
        "Automatic-border cells/backgrounds do not exact-join",
      );
    const direct = evidence.source_part === document.source.main_part;
    const expectedPrefix = direct
      ? `${table.anchor.path}/w:tblPr[1]`
      : "/w:styles[1]/w:style[";
    if (
      direct
        ? evidence.source_path !== `${expectedPrefix}/w:tblBorders[1]`
        : evidence.source_part !== resolved.source_parts.styles_part ||
          !/^\/w:styles\[1\]\/w:style\[[1-9][0-9]*\]\/w:tblPr\[1\]\/w:tblBorders\[1\]$/.test(
            evidence.source_path,
          )
    )
      throw new TypeError(
        "Automatic-border source owner is not the qualified table/style part",
      );
    const sourceParts = document.passthrough_parts.filter(
      (part) => part.part_name === evidence.source_part,
    );
    if ((!direct || sourceParts.length > 0) &&
      (sourceParts.length !== 1 || sourceParts[0]!.sha256 !== evidence.source_sha256)) {
      throw new TypeError("Automatic-border source part digest does not exact-join");
    }
    // The main XML part is normally modeled, not a passthrough part. Its
    // whole-part digest then relies on trusted native extraction of the exact
    // joined package; a body/element anchor hash is NOT a whole-part hash.
    for (const diagnostic of evidence.source_diagnostics) {
      const key = identity(diagnostic);
      if (
        !originalKeys.has(key) ||
        diagnostic.scope_id !== table.id ||
        diagnostic.part_name !== evidence.source_part ||
        (diagnostic.code === "UNMODELED_TABLE_PROPERTY"
          ? !direct || diagnostic.path !== evidence.source_path
          : diagnostic.code !== "TABLE_STYLE_EFFECTS_PRESERVED" ||
            direct ||
            diagnostic.path !==
              evidence.source_path.replace(/\/w:tblBorders\[1\]$/, ""))
      )
        throw new TypeError(
          "Automatic-border source diagnostic is not individually qualified",
        );
      covered.add(key);
    }
    table.borders = structuredClone(evidence.borders);
    entry.borders = structuredClone(evidence.borders);
    delete entry.automatic_border_preview;
    facts.push({
      table_id: entry.table_id,
      evidence: structuredClone(evidence),
    });
  }
  if (facts.length === 0)
    throw new TypeError("No source-qualified automatic borders are available");
  document.unsupported = document.unsupported.filter(
    (_, index) => !covered.has(identity(diagnostics.document[index]!)),
  );
  resolved.diagnostics = resolved.diagnostics.filter(
    (_, index) => !covered.has(identity(diagnostics.resolved[index]!)),
  );
  return {
    document,
    resolved,
    facts,
    source_diagnostics: diagnostics,
    source: {
      document_id: original.document_id,
      revision: original.revision,
      package_sha256: original.source.package_sha256,
    },
  };
}

/** Browser-safe distinct approximate output decoder; never upgrades to strict. */
export function decodeNativeDocxAutomaticBorderPreviewV1(
  value: unknown,
): { ok: true; value: NativeDocxAutomaticBorderPreviewV1 } | { ok: false } {
  try {
    // Legacy eligibility intentionally contains null slots; its existing decoder
    // owns those slots after the enclosing bounded wire check.
    const candidate = value as NativeDocxAutomaticBorderPreviewV1;
    const scan =
      candidate && typeof candidate === "object" ? { ...candidate } : value;
    if (scan && typeof scan === "object")
      delete (scan as Record<string, unknown>).legacy_eligibility;
    if (preflightWire(scan, "automatic-border preview").length)
      return { ok: false };
    // Validate the separately nullable legacy subtree before allocating its
    // owned copy. It must not bypass the enclosing allocation budget.
    const legacy = candidate?.legacy_eligibility === undefined ? undefined :
      decodeNativeDocxApproximationEligibilityV1(candidate.legacy_eligibility, candidate.rendering_provenance.pagination_settings);
    const input = structuredClone(scan) as NativeDocxAutomaticBorderPreviewV1;
    if (legacy !== undefined) input.legacy_eligibility = legacy;
    if (
      Object.keys(input)
        .filter((key) => key !== "legacy_eligibility" && key !== 'approximated_font_sizes')
        .sort()
        .join(",") !==
        [
          "protocol",
          "version",
          "fidelity",
          "read_only",
          "policy",
          "page_background_rgb",
          "source",
          "approximated_render_properties",
          "source_diagnostics",
          "reasons",
          "status",
          "pages",
          "resources",
          "diagnostics",
          "rendering_provenance",
        ]
          .sort()
          .join(",") ||
      input.protocol !== DOCX_AUTO_BORDER_PREVIEW_PROTOCOL ||
      input.version !== 1 ||
      input.fidelity !== "approximate" ||
      input.read_only !== true ||
      input.policy !== DOCX_AUTO_BORDER_POLICY ||
      input.page_background_rgb !== "FFFFFF"
    )
      return { ok: false };
    const paint = decodeNativeDocxPagePaintV1({
      protocol: "injoffice.docx.page-paint",
      version: 1,
      status: input.status,
      pages: input.pages,
      resources: input.resources,
      diagnostics: input.diagnostics,
      provenance: input.rendering_provenance,
    });
    if (
      !paint.ok ||
      !input.source ||
      Object.keys(input.source).sort().join(",") !==
        "document_id,package_sha256,revision" ||
      input.source.document_id !== paint.value.provenance.document_id ||
      input.source.revision !== paint.value.provenance.revision ||
      input.source.package_sha256 !== paint.value.provenance.package_sha256
    )
      return { ok: false };
    if (
      !Array.isArray(input.reasons) ||
      input.reasons.length < 1 ||
      input.reasons.length > 300 ||
      !input.reasons.includes(DOCX_AUTO_BORDER_WARNING) ||
      input.reasons.some(
        (reason) => typeof reason !== "string" || reason.length > 8192,
      )
    )
      return { ok: false };
    if (
      !input.source_diagnostics ||
      Object.keys(input.source_diagnostics).sort().join(",") !==
        "document,resolved"
    )
      return { ok: false };
    for (const list of Object.values(input.source_diagnostics))
      if (
        !Array.isArray(list) ||
        list.length > 1000 ||
        list.some(
          (d) =>
            !d ||
            Object.keys(d).sort().join(",") !==
              "code,part_name,path,scope_id" ||
            Object.values(d).some(
              (v) => typeof v !== "string" || v.length > 4096,
            ),
        )
      )
        return { ok: false };
    const retained = new Set(
      [
        ...input.source_diagnostics.document,
        ...input.source_diagnostics.resolved,
      ].map(identity),
    );
    let cells = 0;
    const facts = input.approximated_render_properties;
    if (
      !Array.isArray(facts) ||
      facts.length < 1 ||
      facts.length > 1000 ||
      new Set(facts.map((f) => f?.table_id)).size !== facts.length
    )
      return { ok: false };
    for (const fact of facts) {
      if (
        !fact ||
        Object.keys(fact).sort().join(",") !== "evidence,table_id" ||
        typeof fact.table_id !== "string" ||
        !/^[A-Za-z0-9][A-Za-z0-9._:-]{0,255}$/.test(fact.table_id) ||
        !validNativeDocxAutomaticBorderEvidenceV1(fact.evidence) ||
        fact.evidence.package_sha256 !== input.source.package_sha256 ||
        fact.evidence.source_diagnostics.some(
          (d) => d.scope_id !== fact.table_id || !retained.has(identity(d)),
        )
      )
        return { ok: false };
      cells += fact.evidence.cell_ids.length;
      if (cells > 100_000) return { ok: false };
    }
    if (input.legacy_eligibility !== undefined) {
      const eligibility = decodeNativeDocxApproximationEligibilityV1(
        input.legacy_eligibility,
        paint.value.provenance.pagination_settings,
      );
      if (
        eligibility.status !== "eligible" || !input.reasons.includes(DOCX_APPROXIMATE_LINE_BOX_WARNING) ||
        eligibility.reasons.some((reason) => !input.reasons.includes(reason))
      )
        return { ok: false };
      const absent = eligibility.absent_font_sizes ?? []
      if(eligibility.legacy_table_origins?.length&&!input.reasons.includes(DOCX_LEGACY_TABLE_ORIGIN_WARNING))return {ok:false}
      if (input.approximated_font_sizes !== undefined) {
        if (!validNativeDocxApproximatedFontSizesV1(input.approximated_font_sizes, absent, input.source.package_sha256) || !input.reasons.includes(DOCX_ABSENT_FONT_SIZE_WARNING)) return { ok: false }
      } else if (input.status === 'painted' && absent.length > 0) return { ok: false }
    } else if (
      paint.value.provenance.pagination_settings.diagnostics.length !== 0 || input.approximated_font_sizes !== undefined
    )
      return { ok: false };
    return { ok: true, value: input };
  } catch {
    return { ok: false };
  }
}
