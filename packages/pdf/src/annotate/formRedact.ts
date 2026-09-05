/**
 * Destructive AcroForm redaction.
 *
 * This is deliberately **not** a convenience wrapper around annotation deletion:
 * a Widget is only the visible instance of a field. Removing just that annotation
 * leaves the field tree, /V value, and often the appearance stream recoverable.
 *
 * The only selector this primitive accepts is an exact fully-qualified field name
 * plus an explicit whole-field acknowledgement. It removes every widget for that
 * terminal field, unlinks the field from AcroForm, deletes its value/appearance
 * closure, and reloads the saved bytes to prove that the field and widgets are gone.
 * Anything we cannot prove isolated (actions, XFA, indirect/shared appearance
 * resources, missing page ownership) is rejected before bytes are returned.
 */
import { PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFObject, PDFRawStream, PDFRef, PDFStream, PDFString } from 'pdf-lib';
import { renderPageToPng } from '../ocr/renderPage.js';
import { PdfViewerDocument } from '../viewer.js';

const N = {
  A: PDFName.of('A'),
  AA: PDFName.of('AA'),
  AP: PDFName.of('AP'),
  Annots: PDFName.of('Annots'),
  Fields: PDFName.of('Fields'),
  Kids: PDFName.of('Kids'),
  MK: PDFName.of('MK'),
  Subtype: PDFName.of('Subtype'),
  Type: PDFName.of('Type'),
  Widget: PDFName.of('Widget'),
  Font: PDFName.of('Font'),
} as const;

/** A redaction target can only address a complete AcroForm terminal field. */
export interface FormFieldRedactionSpec {
  /** Exact fully-qualified AcroForm field name (for example `customer.ssn`). */
  name: string;
  /** Must be literal true. Widget/rectangle/partial-glyph selection is not supported. */
  wholeField: true;
}

export interface FormFieldsRedactionResult {
  bytes: Uint8Array;
  removed: string[];
}

/** A complete, currently enumerable AcroForm terminal-field identity. This is
 * deliberately the same narrow selector accepted by redactFormFields: field
 * values, widget appearances, and client-drawn rectangles are never exposed
 * as a substitute for whole-field deletion. */
export interface RedactableFormFieldRef {
  name: string;
  wholeField: true;
}

type RefMap = Map<string, PDFRef>;
const refKey = (ref: PDFRef): string => `${ref.objectNumber}:${ref.generationNumber}`;
const addRef = (refs: RefMap, ref: PDFRef): void => { refs.set(refKey(ref), ref); };

interface PlannedField {
  name: string;
  field: ReturnType<ReturnType<PDFDocument['getForm']>['getFields']>[number];
  widgets: PDFRef[];
  appearanceClosure: RefMap;
  textValues: string[];
}

const isWidget = (dict: PDFDict): boolean => dict.lookupMaybe(N.Subtype, PDFName)?.toString() === N.Widget.toString();

function indirectDict(doc: PDFDocument, ref: PDFRef, what: string): PDFDict {
  const dict = doc.context.lookupMaybe(ref, PDFDict);
  if (!dict) throw new Error(`${what} is not an indirect dictionary`);
  return dict;
}

/** Add an indirect-object closure. It is only used below for /AP and /MK, never
 * arbitrary field actions: action trees can point at pages or document-level state. */
function addClosure(doc: PDFDocument, object: PDFObject | undefined, into: RefMap, seen: Set<PDFObject>): void {
  if (!object || seen.has(object)) return;
  seen.add(object);
  if (object instanceof PDFRef) {
    if (into.has(refKey(object))) return;
    const resolved = doc.context.lookup(object);
    // Fonts carry glyph programs, not a field value or its visible appearance.
    // They are commonly shared with unrelated page content; deleting one would
    // corrupt that content, while retaining it cannot reconstruct this field.
    if (resolved instanceof PDFDict && resolved.lookupMaybe(N.Type, PDFName)?.toString() === N.Font.toString()) return;
    addRef(into, object);
    addClosure(doc, resolved, into, seen);
    return;
  }
  if (object instanceof PDFArray) {
    for (let i = 0; i < object.size(); i++) addClosure(doc, object.get(i), into, seen);
  } else if (object instanceof PDFDict) {
    for (const [, value] of object.entries()) addClosure(doc, value, into, seen);
  } else if (object instanceof PDFStream) {
    addClosure(doc, object.dict, into, seen);
  }
}

function widgetRefsForField(doc: PDFDocument, field: PlannedField['field']): PDFRef[] {
  const refs: PDFRef[] = [];
  for (const widget of field.acroField.getWidgets()) {
    const ref = doc.context.getObjectRef(widget.dict);
    if (!ref) throw new Error(`field ${JSON.stringify(field.getName())} has a direct widget; refusing ambiguous removal`);
    if (!isWidget(widget.dict)) throw new Error(`field ${JSON.stringify(field.getName())} has a non-widget /Kids entry`);
    if (!refs.some((existing) => refKey(existing) === refKey(ref))) refs.push(ref);
  }
  if (refs.length === 0) throw new Error(`field ${JSON.stringify(field.getName())} has no widget to remove`);
  return refs;
}

function validateFieldShape(field: PlannedField['field']): void {
  const ft = field.acroField.FT();
  if (ft?.toString() === '/Sig') throw new Error(`signature field ${JSON.stringify(field.getName())} is not redaction-safe`);
  for (const forbidden of [N.A, N.AA]) {
    if (field.acroField.dict.has(forbidden)) {
      throw new Error(`field ${JSON.stringify(field.getName())} has actions; refusing to leave action data orphaned`);
    }
  }
}

function extractTextValues(value: PDFObject | undefined): string[] {
  if (value instanceof PDFString || value instanceof PDFHexString) return [value.decodeText()];
  if (value instanceof PDFArray) {
    const values: string[] = [];
    for (let i = 0; i < value.size(); i++) values.push(...extractTextValues(value.get(i)));
    return values;
  }
  return [];
}

function plan(doc: PDFDocument, specs: FormFieldRedactionSpec[]): PlannedField[] {
  if (specs.length === 0) return [];
  const form = doc.getForm();
  if (form.hasXFA()) throw new Error('XFA forms are not supported: field data may be mirrored outside the AcroForm tree');
  const names = new Set<string>();
  for (const spec of specs) {
    if (spec.wholeField !== true) throw new Error(`field ${JSON.stringify(spec.name)} must be redacted as a complete field`);
    if (spec.name.trim() === '') throw new Error('field name must not be empty');
    if (names.has(spec.name)) throw new Error(`field ${JSON.stringify(spec.name)} was selected more than once`);
    names.add(spec.name);
  }
  const fields = form.getFields();
  return specs.map((spec) => {
    const matches = fields.filter((field) => field.getName() === spec.name);
    if (matches.length !== 1) {
      throw new Error(matches.length === 0 ? `field ${JSON.stringify(spec.name)} does not exist` : `field ${JSON.stringify(spec.name)} is ambiguous`);
    }
    const field = matches[0]!;
    validateFieldShape(field);
    const widgets = widgetRefsForField(doc, field);
    const appearanceClosure: RefMap = new Map();
    for (const widgetRef of widgets) {
      const widget = indirectDict(doc, widgetRef, `widget for ${JSON.stringify(spec.name)}`);
      for (const key of [N.AP, N.MK]) addClosure(doc, widget.get(key), appearanceClosure, new Set());
      if (widget.has(N.A) || widget.has(N.AA)) {
        throw new Error(`widget for ${JSON.stringify(spec.name)} has actions; refusing to leave action data orphaned`);
      }
    }
    // A field that is itself a widget is one object. It is already deleted as the
    // field, so it must not be treated as a stand-alone appearance closure member.
    appearanceClosure.delete(refKey(field.ref));
    return { name: spec.name, field, widgets, appearanceClosure, textValues: extractTextValues(field.acroField.V()) };
  });
}

/**
 * Lists only terminal fields whose complete name and widget shape can be
 * verified without changing the source bytes. A returned identity is suitable
 * for passing unchanged to redactFormFields; fields with XFA mirroring,
 * actions, signatures, direct widgets, or malformed widget ownership are
 * withheld rather than approximated with a visible rectangle.
 *
 * This is discovery, not an eligibility promise for every later destructive
 * preflight: redaction still repeats its complete isolation/appearance proof
 * immediately before serializing output, so a changed or shared resource fails
 * closed instead of widening a match.
 */
export async function listRedactableFormFields(bytes: Uint8Array): Promise<RedactableFormFieldRef[]> {
  const doc = await PDFDocument.load(bytes);
  const form = doc.getForm();
  if (form.hasXFA()) return [];
  const out: RedactableFormFieldRef[] = [];
  for (const field of form.getFields()) {
    const name = field.getName();
    try {
      // plan is read-only. Keeping its exact validation shared with mutation
      // prevents discovery from offering selectors that the redactor rejects
      // for field/widget shape before it can begin a proof-safe operation.
      plan(doc, [{ name, wholeField: true }]);
      out.push({ name, wholeField: true });
    } catch {
      // There is intentionally no partial/widget fallback for an unsupported
      // field. The caller receives only safely addressable whole fields.
    }
  }
  return out;
}

function removeWidgetsFromPages(doc: PDFDocument, widgets: RefMap): Set<number> {
  const renderedPages = new Set<number>();
  const found = new Set<string>();
  for (const [pageIndex, page] of doc.getPages().entries()) {
    const annots = page.node.lookupMaybe(N.Annots, PDFArray);
    if (!annots) continue;
    for (let i = annots.size() - 1; i >= 0; i--) {
      const ref = annots.get(i);
      if (ref instanceof PDFRef && widgets.has(refKey(ref))) {
        annots.remove(i);
        found.add(refKey(ref));
        renderedPages.add(pageIndex + 1);
      }
    }
  }
  if (found.size !== widgets.size) throw new Error('one or more field widgets are not owned by a page annotation array');
  return renderedPages;
}

function reachableFieldRefs(doc: PDFDocument): Set<string> {
  const seen = new Set<string>();
  const visit = (ref: PDFRef): void => {
    if (seen.has(refKey(ref))) return;
    seen.add(refKey(ref));
    const dict = doc.context.lookupMaybe(ref, PDFDict);
    if (!dict || isWidget(dict)) return;
    const kids = dict.lookupMaybe(N.Kids, PDFArray);
    if (!kids) return;
    for (let i = 0; i < kids.size(); i++) {
      const child = kids.get(i);
      if (child instanceof PDFRef) visit(child);
    }
  };
  const fields = doc.getForm().acroForm.dict.lookupMaybe(N.Fields, PDFArray);
  if (!fields) return seen;
  for (let i = 0; i < fields.size(); i++) {
    const ref = fields.get(i);
    if (ref instanceof PDFRef) visit(ref);
  }
  return seen;
}

function ancestors(field: PlannedField['field']): PDFRef[] {
  const refs: PDFRef[] = [];
  let current = field.acroField.getParent();
  while (current) {
    refs.push(current.ref);
    current = current.getParent();
  }
  return refs;
}

function refsIn(object: PDFObject | undefined, out: RefMap, seen: Set<PDFObject>): void {
  if (!object || seen.has(object)) return;
  seen.add(object);
  if (object instanceof PDFRef) {
    addRef(out, object);
    return;
  }
  if (object instanceof PDFArray) {
    for (let i = 0; i < object.size(); i++) refsIn(object.get(i), out, seen);
  } else if (object instanceof PDFDict) {
    for (const [, value] of object.entries()) refsIn(value, out, seen);
  } else if (object instanceof PDFStream) {
    refsIn(object.dict, out, seen);
  }
}

/** There must be no surviving pointer to an object we are about to delete. This
 * rejects shared appearance XObjects instead of corrupting another field or
 * retaining a recoverable copy of a redacted appearance. */
function assertIsolated(doc: PDFDocument, deleting: RefMap): void {
  for (const [source, object] of doc.context.enumerateIndirectObjects()) {
    if (deleting.has(refKey(source))) continue;
    const refs: RefMap = new Map();
    refsIn(object, refs, new Set());
    for (const [key, target] of deleting) {
      if (refs.has(key)) throw new Error(`redaction object ${target.toString()} is shared by ${source.toString()}`);
    }
  }
}

function survivingTargetWidgets(doc: PDFDocument, targetRefs: RefMap): boolean {
  for (const page of doc.getPages()) {
    const annots = page.node.lookupMaybe(N.Annots, PDFArray);
    if (!annots) continue;
    for (let i = 0; i < annots.size(); i++) {
      const ref = annots.get(i);
      if (ref instanceof PDFRef && targetRefs.has(refKey(ref))) return true;
    }
  }
  return false;
}

function rawStreamSignatures(doc: PDFDocument): Set<string> {
  const signatures = new Set<string>();
  for (const [, object] of doc.context.enumerateIndirectObjects()) {
    if (object instanceof PDFRawStream) signatures.add(Buffer.from(object.getContents()).toString('base64'));
  }
  return signatures;
}

async function verify(
  bytes: Uint8Array,
  names: string[],
  deletedRefs: RefMap,
  targetAppearanceSignatures: Set<string>,
  removedTextValues: string[],
  renderPages: Set<number>,
): Promise<void> {
  const reloaded = await PDFDocument.load(bytes);
  const remaining = new Set(reloaded.getForm().getFields().map((field) => field.getName()));
  for (const name of names) {
    if (remaining.has(name)) throw new Error(`verification failed: field ${JSON.stringify(name)} remains in AcroForm`);
  }
  const outputRefs = new Set(reloaded.context.enumerateIndirectObjects().map(([ref]) => refKey(ref)));
  for (const [key, ref] of deletedRefs) {
    if (outputRefs.has(key)) throw new Error(`verification failed: deleted object ${ref.toString()} remains in serialized bytes`);
  }
  if (survivingTargetWidgets(reloaded, deletedRefs)) throw new Error('verification failed: a removed widget remains on a page');
  const outputStreams = rawStreamSignatures(reloaded);
  for (const signature of targetAppearanceSignatures) {
    if (outputStreams.has(signature)) throw new Error('verification failed: a redacted appearance stream remains in serialized bytes');
  }
  const viewer = await PdfViewerDocument.load(bytes);
  try {
    for (const page of renderPages) {
      await viewer.getPage(page); // fresh pdf.js parse/page inspection
      const extracted = await viewer.getPageText(page);
      for (const value of removedTextValues) {
        // A matching string elsewhere on the page is intentionally treated as a
        // failed proof, not guessed away as a harmless duplicate.
        if (value.length > 0 && extracted.includes(value)) throw new Error(`verification failed: redacted value remains extractable on page ${page}`);
      }
      const rendered = await renderPageToPng(bytes, page, 1);
      if (rendered.width < 1 || rendered.height < 1 || rendered.png.length === 0) throw new Error(`verification failed: page ${page} did not rasterize`);
    }
  } finally {
    await viewer.destroy();
  }
}

/**
 * Physically removes complete AcroForm terminal fields. This function is atomic:
 * all targets are validated before mutating the in-memory document, and no bytes are
 * returned unless fresh AcroForm/widget/viewer/raster verification succeeds.
 */
export async function redactFormFields(bytes: Uint8Array, specs: FormFieldRedactionSpec[]): Promise<FormFieldsRedactionResult> {
  if (specs.length === 0) return { bytes, removed: [] };
  const doc = await PDFDocument.load(bytes);
  const planned = plan(doc, specs);
  const widgets: RefMap = new Map();
  for (const item of planned) for (const ref of item.widgets) addRef(widgets, ref);
  const renderPages = removeWidgetsFromPages(doc, widgets);

  // Unlink whole fields only after every selector has passed validation. Calling the
  // low-level AcroForm method avoids pdf-lib's visual-field helper, which requires a
  // normal appearance stream and is unsuitable for fail-closed deletion.
  const allAncestors = planned.flatMap((item) => ancestors(item.field));
  for (const item of planned) doc.getForm().acroForm.removeField(item.field.acroField);

  const reachable = reachableFieldRefs(doc);
  const deleting: RefMap = new Map(widgets);
  for (const item of planned) addRef(deleting, item.field.ref);
  const targetAppearanceSignatures = new Set<string>();
  for (const item of planned) {
    for (const ref of item.appearanceClosure.values()) {
      addRef(deleting, ref);
      const object = doc.context.lookup(ref);
      if (object instanceof PDFRawStream) targetAppearanceSignatures.add(Buffer.from(object.getContents()).toString('base64'));
    }
  }
  // pdf-lib (and several producers) can leave a superseded appearance stream as an
  // unreferenced indirect object after regenerating a widget. It is still serialized,
  // and can contain the field value, so remove every byte-identical orphan copy too.
  // A live shared copy is caught by assertIsolated below and makes the operation fail.
  for (const [ref, object] of doc.context.enumerateIndirectObjects()) {
    if (object instanceof PDFRawStream && targetAppearanceSignatures.has(Buffer.from(object.getContents()).toString('base64'))) {
      addRef(deleting, ref);
    }
  }
  // Empty non-terminal ancestors were unlinked recursively by removeField; delete
  // their indirect objects too so parent names cannot survive as orphan form data.
  for (const ref of allAncestors) if (!reachable.has(refKey(ref))) addRef(deleting, ref);
  assertIsolated(doc, deleting);
  for (const ref of deleting.values()) doc.context.delete(ref);

  const out = await doc.save({ updateFieldAppearances: false });
  await verify(
    out,
    planned.map((item) => item.name),
    deleting,
    targetAppearanceSignatures,
    planned.flatMap((item) => item.textValues),
    renderPages,
  );
  return { bytes: out, removed: planned.map((item) => item.name) };
}
