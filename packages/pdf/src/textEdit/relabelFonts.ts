export interface FontRelabelResult {
  readonly bytes: Uint8Array;
  readonly changed: false;
}

export function relabelEmbeddedFontSlots(fontBytes: Uint8Array): FontRelabelResult {
  if (!(fontBytes instanceof Uint8Array)) throw new TypeError('fontBytes must be a Uint8Array');
  return { bytes: fontBytes.slice(), changed: false };
}
