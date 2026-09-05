import { useEffect, useMemo, useRef, useState, type CSSProperties, type ReactNode } from 'react'
import { resolveTheme } from './themes'
import type { DeckSpec, DeckTheme, SlideSpec } from './types'

// DeckView — the HTML renderer. One 16:9 stage, keyboard navigation
// (←/→/Home/End), a slide rail, and speaker notes below. Styling comes
// entirely from the resolved DeckTheme tokens + a small fixed layout system,
// so a theme swap restyles every slide with no markup changes. Hosts embed
// it read-only today; the deck editor (S3) builds on the same component.

/** Minimal inline markup: **bold** and *italic*. Everything else is text. */
export function renderInline(text: string): ReactNode[] {
  const out: ReactNode[] = []
  const re = /\*\*([^*]+)\*\*|\*([^*]+)\*/g
  let last = 0
  let m: RegExpExecArray | null
  let k = 0
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) out.push(text.slice(last, m.index))
    if (m[1] !== undefined) out.push(<strong key={k++}>{m[1]}</strong>)
    else out.push(<em key={k++}>{m[2]}</em>)
    last = m.index + m[0].length
  }
  if (last < text.length) out.push(text.slice(last))
  return out
}

function Eyebrow({ text, theme }: { text?: string; theme: DeckTheme }) {
  if (!text) return null
  return (
    <div
      style={{
        fontFamily: theme.monoFont,
        fontSize: '1.4cqw',
        letterSpacing: '0.16em',
        textTransform: 'uppercase',
        color: theme.accent,
        marginBottom: '1.6cqw',
      }}
    >
      {text}
    </div>
  )
}

function SlideBody({ slide, theme }: { slide: SlideSpec; theme: DeckTheme }) {
  const titleStyle: CSSProperties = {
    fontFamily: theme.displayFont,
    fontWeight: 750,
    color: theme.ink,
    lineHeight: 1.05,
    letterSpacing: '-0.01em',
    margin: 0,
  }
  const bodyStyle: CSSProperties = {
    fontFamily: theme.bodyFont,
    color: theme.muted,
    lineHeight: 1.55,
    whiteSpace: 'pre-wrap',
  }
  const bulletList = (items: string[] | undefined) => (
    <ul style={{ ...bodyStyle, color: theme.ink, fontSize: '2.6cqw', margin: 0, paddingLeft: '2.4cqw', display: 'grid', gap: '1.5cqw' }}>
      {(items ?? []).map((b, i) => (
        <li key={i}>{renderInline(b)}</li>
      ))}
    </ul>
  )

  switch (slide.kind) {
    case 'title':
      return (
        <div style={{ display: 'flex', flexDirection: 'column', justifyContent: 'center', height: '100%' }}>
          <Eyebrow text={slide.eyebrow} theme={theme} />
          <h1 style={{ ...titleStyle, fontSize: '7cqw', maxWidth: '80%' }}>{renderInline(slide.title ?? '')}</h1>
          {slide.subtitle && <p style={{ ...bodyStyle, fontSize: '2.6cqw', marginTop: '2.4cqw', maxWidth: '62%' }}>{renderInline(slide.subtitle)}</p>}
          {slide.body && <p style={{ ...bodyStyle, fontSize: '2cqw', marginTop: '1.6cqw', maxWidth: '62%' }}>{renderInline(slide.body)}</p>}
          <div style={{ position: 'absolute', left: '6cqw', bottom: '6cqw', width: '9cqw', height: '0.55cqw', background: theme.accent }} />
        </div>
      )
    case 'section':
      return (
        <div style={{ display: 'flex', flexDirection: 'column', justifyContent: 'center', height: '100%' }}>
          <Eyebrow text={slide.eyebrow} theme={theme} />
          <h2 style={{ ...titleStyle, fontSize: '5.4cqw', maxWidth: '78%' }}>{renderInline(slide.title ?? '')}</h2>
          {(slide.subtitle || slide.body) && (
            <p style={{ ...bodyStyle, fontSize: '2.3cqw', marginTop: '2cqw', maxWidth: '64%' }}>{renderInline(slide.subtitle ?? slide.body ?? '')}</p>
          )}
        </div>
      )
    case 'quote':
      return (
        <div style={{ display: 'flex', flexDirection: 'column', justifyContent: 'center', height: '100%' }}>
          <div style={{ background: theme.surface, borderLeft: `0.6cqw solid ${theme.accent}`, padding: '4cqw 5cqw', maxWidth: '80%' }}>
            <p style={{ fontFamily: theme.bodyFont, fontStyle: 'italic', color: theme.ink, fontSize: '3.4cqw', lineHeight: 1.4, margin: 0 }}>
              “{renderInline(slide.quote ?? '')}”
            </p>
            {slide.subtitle && (
              <p style={{ fontFamily: theme.monoFont, color: theme.muted, fontSize: '1.6cqw', marginTop: '2.2cqw', letterSpacing: '0.08em', textTransform: 'uppercase' }}>
                — {slide.subtitle}
              </p>
            )}
          </div>
        </div>
      )
    case 'two-col':
      return (
        <div style={{ height: '100%', display: 'flex', flexDirection: 'column' }}>
          <Eyebrow text={slide.eyebrow} theme={theme} />
          <h2 style={{ ...titleStyle, fontSize: '4cqw', marginBottom: '3cqw' }}>{renderInline(slide.title ?? '')}</h2>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '5cqw', flex: 1 }}>
            {[0, 1].map((col) => (
              <div key={col}>
                {slide.colTitles?.[col] && (
                  <div style={{ fontFamily: theme.monoFont, fontSize: '1.5cqw', letterSpacing: '0.1em', textTransform: 'uppercase', color: theme.accent, marginBottom: '1.6cqw' }}>
                    {slide.colTitles[col]}
                  </div>
                )}
                {bulletList(col === 0 ? slide.bullets : slide.bulletsRight)}
              </div>
            ))}
          </div>
        </div>
      )
    case 'closing':
      return (
        <div style={{ display: 'flex', flexDirection: 'column', justifyContent: 'center', alignItems: 'center', height: '100%', textAlign: 'center' }}>
          <h2 style={{ ...titleStyle, fontSize: '5.6cqw' }}>{renderInline(slide.title ?? 'Thank you')}</h2>
          {slide.subtitle && <p style={{ ...bodyStyle, fontSize: '2.2cqw', marginTop: '2cqw' }}>{renderInline(slide.subtitle)}</p>}
          <div style={{ width: '9cqw', height: '0.55cqw', background: theme.accent, marginTop: '3.4cqw' }} />
        </div>
      )
    default: // 'bullets'
      return (
        <div style={{ height: '100%', display: 'flex', flexDirection: 'column' }}>
          <Eyebrow text={slide.eyebrow} theme={theme} />
          {slide.title && <h2 style={{ ...titleStyle, fontSize: '4cqw', marginBottom: '3cqw' }}>{renderInline(slide.title)}</h2>}
          {slide.body && <p style={{ ...bodyStyle, fontSize: '2.1cqw', marginBottom: '2.4cqw', maxWidth: '80%' }}>{renderInline(slide.body)}</p>}
          {bulletList(slide.bullets)}
        </div>
      )
  }
}

export function SlideView({ slide, theme, index, total }: { slide: SlideSpec; theme: DeckTheme; index?: number; total?: number }) {
  return (
    <div
      className="ios-slide"
      style={{
        position: 'relative',
        width: '100%',
        aspectRatio: '16 / 9',
        background: theme.background,
        color: theme.ink,
        overflow: 'hidden',
        containerType: 'inline-size',
        borderRadius: 8,
      }}
    >
      <div style={{ position: 'absolute', inset: '6cqw', overflow: 'hidden' }}>
        <SlideBody slide={slide} theme={theme} />
      </div>
      {index !== undefined && total !== undefined && slide.kind !== 'title' && (
        <div style={{ position: 'absolute', right: '3cqw', bottom: '2.4cqw', fontFamily: theme.monoFont, fontSize: '1.3cqw', color: theme.muted }}>
          {index + 1} / {total}
        </div>
      )}
    </div>
  )
}

export interface DeckViewProps {
  spec: DeckSpec
  /** Show speaker notes under the stage (default true when any exist). */
  showNotes?: boolean
  /** Controlled slide index (the editor shares selection with the stage). */
  at?: number
  onAtChange?: (index: number) => void
}

/**
 * @deprecated Legacy HTML preview only. It is not a production Office-layout
 * authority. Compile authored decks through @injoffice/pptx-authored and render
 * the resulting native v1 contract through @injoffice/pptx-render.
 */
export function DeckView({ spec, showNotes, at: atProp, onAtChange }: DeckViewProps) {
  const theme = useMemo(() => resolveTheme(spec.theme), [spec.theme])
  const [atState, setAtState] = useState(0)
  const at = atProp ?? atState
  const setAt = (next: number | ((i: number) => number)) => {
    const v = typeof next === 'function' ? next(at) : next
    if (onAtChange) onAtChange(v)
    if (atProp === undefined) setAtState(v)
  }
  const total = spec.slides.length
  const slide = spec.slides[Math.min(at, total - 1)]
  const notesOn = showNotes ?? spec.slides.some((s) => s.notes)

  const setAtRef = useRef(setAt)
  setAtRef.current = setAt
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      // Never hijack arrows while someone is typing (the deck editor panel).
      const t = e.target as HTMLElement | null
      if (t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.tagName === 'SELECT' || t.isContentEditable)) return
      if (e.key === 'ArrowRight') setAtRef.current((i) => Math.min(i + 1, total - 1))
      else if (e.key === 'ArrowLeft') setAtRef.current((i) => Math.max(i - 1, 0))
      else if (e.key === 'Home') setAtRef.current(0)
      else if (e.key === 'End') setAtRef.current(total - 1)
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [total])

  if (!slide) return null
  return (
    <div className="ios-deck" style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
      <SlideView slide={slide} theme={theme} index={at} total={total} />
      <div className="ios-deck-rail" style={{ display: 'flex', gap: 6, overflowX: 'auto', paddingBottom: 4 }}>
        {spec.slides.map((s, i) => (
          <button
            key={s.id}
            type="button"
            aria-label={`Slide ${i + 1}${s.title ? `: ${s.title}` : ''}`}
            aria-current={i === at}
            onClick={() => setAt(i)}
            style={{
              flex: '0 0 84px',
              padding: 0,
              border: i === at ? `2px solid ${theme.accent}` : '1px solid rgba(128,128,128,.35)',
              borderRadius: 4,
              background: 'none',
              cursor: 'pointer',
            }}
          >
            <div style={{ pointerEvents: 'none' }}>
              <SlideView slide={s} theme={theme} />
            </div>
          </button>
        ))}
      </div>
      {notesOn && slide.notes && (
        <div className="ios-deck-notes" style={{ fontFamily: theme.monoFont, fontSize: 12, color: theme.muted, whiteSpace: 'pre-wrap' }}>
          {slide.notes}
        </div>
      )}
    </div>
  )
}
