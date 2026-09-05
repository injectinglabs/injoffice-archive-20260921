import type { ReactNode } from 'react'

export function Callout({
  kind = 'note',
  title,
  children,
}: {
  kind?: 'note' | 'caution' | 'ok'
  title?: string
  children: ReactNode
}) {
  const label = title ?? (kind === 'caution' ? 'Caution' : kind === 'ok' ? 'In this package' : 'Note')
  return (
    <aside className={`callout callout--${kind}`}>
      <strong>{label}</strong>
      <div>{children}</div>
    </aside>
  )
}
