import './open-error.css';

export type OpenErrorKind = 'unsupported' | 'encrypted' | 'extract';

export function classifyOpenError(message: string): OpenErrorKind | undefined {
  const text = message.toLowerCase();
  if (/encrypt|password/.test(text)) return 'encrypted';
  if (/choose a docx|drop a supported|not a supported|\.xlsb|\.ods|open.?document/.test(text)) return 'unsupported';
  if (/extract|could not be opened|invalid zip|damaged/.test(text)) return 'extract';
}

// Office Open XML packages are ZIP archives and PDFs start with %PDF. A file that fails this test
// can never reach an engine, so the workspace refuses it before a tab, ribbon and "Saved" status
// appear around an editor that only has a raw engine message to show.
const signatures: Record<string, { magic: number[]; container: string }> = {
  docx: { magic: [0x50, 0x4b], container: 'Word document' },
  xlsx: { magic: [0x50, 0x4b], container: 'Excel workbook' },
  pptx: { magic: [0x50, 0x4b], container: 'PowerPoint presentation' },
  pdf: { magic: [0x25, 0x50, 0x44, 0x46], container: 'PDF document' },
};

/** The reason this file cannot be a document of its own extension, or undefined when it can. */
export function containerFailure(name: string, bytes: Uint8Array): string | undefined {
  const extension = name.split('.').pop()?.toLowerCase() ?? '';
  const signature = signatures[extension];
  if (!signature) return;
  if (!bytes.length) return `${name} is empty (0 bytes), so it holds no ${signature.container}.`;
  if (signature.magic.some((byte, index) => bytes[index] !== byte)) {
    const header = [...bytes.slice(0, 4)].map(byte => byte.toString(16).padStart(2, '0')).join(' ');
    const expected = extension === 'pdf' ? '%PDF' : 'a ZIP package (PK)';
    return `${name} does not start with ${expected}: the first bytes are ${header || '(none)'}. A .${extension} file must be ${extension === 'pdf' ? 'a PDF' : 'an Office Open XML ZIP package'}.`;
  }
}

const copy: Record<OpenErrorKind, string> = {
  unsupported: 'InjOffice opens .docx, .xlsx, .pptx and .pdf files. Binary Word 97–2003, PowerPoint 97–2003, Excel .xlsb and OpenDocument files stay closed so their bytes are never rewritten.',
  encrypted: 'The file is password-protected. Remove the password in the app that created it, then open the unprotected copy here.',
  extract: 'The file may be damaged, or it may not be the kind of file its name says it is. The copy on disk was not changed.',
};

export default function OpenError({
  name,
  kind,
  detail,
  busy,
  onOpen,
  onHome,
}: {
  name?: string;
  kind: OpenErrorKind;
  detail?: string;
  busy?: boolean;
  onOpen(): void;
  onHome(): void;
}) {
  return (
    <main className="open-error" aria-labelledby="open-error-title">
      <div className="open-error-card">
        <svg className="open-error-glyph" viewBox="0 0 32 38" fill="none" aria-hidden="true"><path d="M5 1h15l8 8v27H5V1Z" stroke="currentColor" strokeWidth="1.4" /><path d="M20 1v8h8" stroke="currentColor" strokeWidth="1.4" /><path d="m11 18 10 10M21 18 11 28" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" /></svg>
        <h1 id="open-error-title">InjOffice can&rsquo;t open this file</h1>
        {name ? <p className="open-error-name">{name}</p> : null}
        <p className="open-error-reason">{copy[kind]}</p>
        <div className="open-error-actions">
          <button type="button" className="open-error-open" disabled={busy} onClick={onOpen}>Open another file</button>
          <button type="button" onClick={onHome}>Go to start page</button>
        </div>
        {detail ? <details className="open-error-details"><summary>Details</summary><p>{detail}</p></details> : null}
      </div>
    </main>
  );
}
