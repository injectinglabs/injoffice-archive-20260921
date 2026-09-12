export { applyMarkups } from './markup.js';
export { applyDrawings, applyNoteEdits } from './drawing.js';
export { applyFormValues, type FormValueFailure, type FormValuesResult, type FormValuesOptions, type FormValueAppearance, type TextAppearanceFont } from './forms.js';
export { applyStamps, applySignatureStamps, VISUAL_SIGNATURE_CONTENT_PREFIX } from './stamp.js';
export { applyAnnotDeletes } from './annotDelete.js';
export { listRedactableFormFields, redactFormFields, type FormFieldRedactionSpec, type FormFieldsRedactionResult, type RedactableFormFieldRef } from './formRedact.js';
export type {
  AnnotDeleteSpec,
  DrawingSpec,
  FormValueSpec,
  MarkupSpec,
  MarkupType,
  NoteEditSpec,
  NoteReplyTarget,
  SignatureStampSpec,
  StampSpec,
} from './types.js';
