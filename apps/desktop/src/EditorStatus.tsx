import { createContext, useContext, type ReactNode } from 'react';
import { createPortal } from 'react-dom';

// null suppresses background sessions; undefined supports standalone editor hosts.
export const EditorStatusContext = createContext<HTMLElement | null | undefined>(undefined);
export function EditorStatus({ children, label }: { children: ReactNode; label: string }) {
  const target = useContext(EditorStatusContext);
  const content = <div className="editor-status-segment" aria-label={label}>{children}</div>;
  return target === undefined ? content : target ? createPortal(content, target) : null;
}
