import type { ButtonHTMLAttributes, InputHTMLAttributes, ReactNode, SelectHTMLAttributes, SVGProps, TextareaHTMLAttributes } from 'react'

function StrokeIcon(props: SVGProps<SVGSVGElement>) {
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true" {...props} />
  )
}

export function IconUndo() {
  return <StrokeIcon><path d="M3.5 7.5H11a2.5 2.5 0 0 1 0 5H9.5" /><path d="M6 4.5 3.5 7.5 6 10.5" /></StrokeIcon>
}
export function IconRedo() {
  return <StrokeIcon><path d="M12.5 7.5H5a2.5 2.5 0 0 0 0 5h1.5" /><path d="M10 4.5 12.5 7.5 10 10.5" /></StrokeIcon>
}
export function IconPrint() {
  return <StrokeIcon><path d="M4.5 6V3.5h7V6" /><path d="M4 9.5h8v3.5H4z" /><path d="M3.5 6h9v4H12v-1.5H4V10H3.5z" /></StrokeIcon>
}
export function IconFilter() {
  return <StrokeIcon><path d="M2.5 3.5h11l-4 5v4l-3-1.5v-2.5z" /></StrokeIcon>
}
export function IconBorder() {
  return <StrokeIcon><rect x="3" y="3" width="10" height="10" /></StrokeIcon>
}

export function DsMark() {
  return (
    <span className="ds-mark" aria-hidden="true">
      <svg width="20" height="20" viewBox="0 0 20 20" fill="none">
        <rect x="2.5" y="2.5" width="15" height="15" stroke="currentColor" />
        <path d="M2.5 7.5h15M2.5 12.5h15M7.5 2.5v15M12.5 2.5v15" stroke="currentColor" />
        <rect x="7.5" y="7.5" width="5" height="5" fill="var(--ds-green)" />
      </svg>
    </span>
  )
}

export function DsButton({
  variant = 'text',
  className,
  ...props
}: ButtonHTMLAttributes<HTMLButtonElement> & { variant?: 'text' | 'filled' | 'green' | 'outlined' | 'refuse' }) {
  const extensionClasses = className
    ?.split(/\s+/)
    .filter((token) => token && token !== 'workbench-button' && !token.startsWith('workbench-button--'))
    .join(' ')

  return <button type="button" {...props} className={`ds-btn ds-btn--${variant}${extensionClasses ? ` ${extensionClasses}` : ''}`} />
}

export function DsTool(props: ButtonHTMLAttributes<HTMLButtonElement>) {
  return <button type="button" {...props} className="ds-tool" />
}

export function DsSegment({
  label,
  value,
  options,
  onChange,
}: {
  label: string
  value: string
  options: readonly { id: string; label: string }[]
  onChange: (id: string) => void
}) {
  return (
    <div className="ds-segment" role="group" aria-label={label}>
      {options.map((option) => (
        <button
          type="button"
          key={option.id}
          aria-pressed={value === option.id}
          onClick={() => onChange(option.id)}
        >
          {option.label}
        </button>
      ))}
    </div>
  )
}

export function DsField({
  label,
  children,
  className,
}: {
  label: string
  children: ReactNode
  className?: string
}) {
  return <label className={`ds-field${className ? ` ${className}` : ''}`}>{label}{children}</label>
}

export function DsInput(props: InputHTMLAttributes<HTMLInputElement>) {
  return <input {...props} />
}

export function DsSelect(props: SelectHTMLAttributes<HTMLSelectElement>) {
  return <select {...props} />
}

export function DsTextarea(props: TextareaHTMLAttributes<HTMLTextAreaElement>) {
  return <textarea {...props} />
}

export function DsChip({
  tone = 'plain',
  children,
}: {
  tone?: 'plain' | 'green' | 'blue' | 'refuse'
  children: ReactNode
}) {
  return <span className={`ds-chip ds-chip--${tone}`}>{children}</span>
}

export function DsCallout({
  tone = 'note',
  title,
  children,
}: {
  tone?: 'note' | 'green' | 'refuse'
  title: string
  children: ReactNode
}) {
  return (
    <div className={`ds-callout ds-callout--${tone}`} role={tone === 'refuse' ? 'alert' : 'status'}>
      <strong>{title}</strong>
      <p>{children}</p>
    </div>
  )
}

export function DsAvatar({
  initials,
  tone = 1,
}: {
  initials: string
  tone?: 1 | 2 | 3
}) {
  return <span className={`ds-avatar${tone === 1 ? '' : ` ds-avatar--${tone}`}`}>{initials}</span>
}
