import type { ReactNode } from 'react'

export function LiveBench({
  title = 'Preview',
  hint = 'Runs the published package in this page',
  children,
}: {
  title?: string
  hint?: string
  children: ReactNode
}) {
  return (
    <section className="bench" aria-label={title}>
      <div className="bench-head">
        <strong>{title}</strong>
        <span>{hint}</span>
      </div>
      <div className="bench-body">{children}</div>
    </section>
  )
}
