import './open-error.css';

export type OpenErrorKind = 'unsupported' | 'encrypted' | 'extract';

const copy: Record<OpenErrorKind, { title: string; body: string }> = {
  unsupported: {
    title: 'This file type is not supported',
    body: 'InjOffice opens .docx, .xlsx, .pptx, and .pdf. Binary Word 97–2003, PowerPoint 97–2003, Excel .xlsb, and OpenDocument files stay closed so prior bytes are never rewritten.',
  },
  encrypted: {
    title: 'This file is encrypted',
    body: 'Password-protected compound-file documents are refused. Decrypt them in the original app, then open the unencrypted Office Open XML or PDF here.',
  },
  extract: {
    title: 'This file could not be opened',
    body: 'Native extract refused the document. The file on disk is unchanged. If this is a supported type, it may be damaged or use a construct that is not modeled yet.',
  },
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
  const message = copy[kind];
  return (
    <main className="open-error" aria-labelledby="open-error-title">
      <h1 id="open-error-title">{message.title}</h1>
      {name ? <p className="open-error-name">{name}</p> : null}
      <p>{message.body}</p>
      {detail ? <p className="open-error-detail">{detail}</p> : null}
      <div className="open-error-actions">
        <button type="button" className="open-error-open" disabled={busy} onClick={onOpen}>Open another file</button>
        <button type="button" onClick={onHome}>Back to start</button>
      </div>
    </main>
  );
}
