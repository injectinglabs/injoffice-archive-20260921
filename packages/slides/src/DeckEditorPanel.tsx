import { addSlide, bulletsToText, moveSlide, removeSlide, textToBullets, updateSlide } from './edit'
import type { DeckSpec, SlideKind, SlideSpec } from './types'

// DeckEditorPanel — the S3 MVP editor: a side panel editing the SELECTED
// slide's fields plus slide-level operations (add/delete/move). The
// rendered DeckView stays the single source of visual truth; every change
// flows through the pure operations in edit.ts and lands as a new spec via
// onChange — the host owns state, dirty tracking and saving. Structure-only
// styling (ioc-* hooks), same convention as ChartPanel/PivotPanel.

const KINDS: readonly SlideKind[] = ['title', 'section', 'bullets', 'two-col', 'quote', 'closing']

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <label className="ioc-field ioc-field--stack">
      <span>{label}</span>
      {children}
    </label>
  )
}

function TextInput({ value, onCommit, placeholder }: { value: string; onCommit: (v: string) => void; placeholder?: string }) {
  return (
    <input
      type="text"
      defaultValue={value}
      key={value}
      placeholder={placeholder}
      onBlur={(e) => {
        if (e.target.value !== value) onCommit(e.target.value)
      }}
      onKeyDown={(e) => {
        if (e.key === 'Enter') (e.target as HTMLInputElement).blur()
      }}
    />
  )
}

function TextArea({ value, onCommit, rows = 5, placeholder }: { value: string; onCommit: (v: string) => void; rows?: number; placeholder?: string }) {
  return (
    <textarea
      defaultValue={value}
      key={value}
      rows={rows}
      placeholder={placeholder}
      onBlur={(e) => {
        if (e.target.value !== value) onCommit(e.target.value)
      }}
    />
  )
}

export interface DeckEditorPanelProps {
  spec: DeckSpec
  /** Selected slide index (host-controlled, shared with DeckView). */
  at: number
  onChange: (next: DeckSpec) => void
  onSelect: (index: number) => void
}

export function DeckEditorPanel({ spec, at, onChange, onSelect }: DeckEditorPanelProps) {
  const slide: SlideSpec | undefined = spec.slides[at]
  if (!slide) return null
  const patch = (p: Partial<Omit<SlideSpec, 'id'>>) => onChange(updateSlide(spec, at, p))
  const str = (v: string) => (v.trim() === '' ? undefined : v)

  return (
    <div className="ioc-panel ioc-deck-editor" aria-label="Slide editor">
      <div className="ioc-row ioc-deck-ops">
        <select
          aria-label="Slide kind"
          value={slide.kind}
          onChange={(e) => patch({ kind: e.target.value as SlideKind })}
        >
          {KINDS.map((k) => (
            <option key={k} value={k}>{k}</option>
          ))}
        </select>
        <button type="button" className="ioc-add" title="Move slide left" disabled={at === 0} onClick={() => { onChange(moveSlide(spec, at, -1)); onSelect(at - 1) }}>
          ←
        </button>
        <button type="button" className="ioc-add" title="Move slide right" disabled={at === spec.slides.length - 1} onClick={() => { onChange(moveSlide(spec, at, 1)); onSelect(at + 1) }}>
          →
        </button>
        <button type="button" className="ioc-remove" title="Delete this slide" disabled={spec.slides.length <= 1} onClick={() => { onChange(removeSlide(spec, at)); onSelect(Math.max(0, at - 1)) }}>
          ✕
        </button>
      </div>

      {slide.kind !== 'quote' && (
        <Field label="Title">
          <TextInput value={slide.title ?? ''} onCommit={(v) => patch({ title: str(v) })} />
        </Field>
      )}
      {slide.kind === 'quote' ? (
        <>
          <Field label="Quote">
            <TextArea rows={3} value={slide.quote ?? ''} onCommit={(v) => patch({ quote: str(v) })} />
          </Field>
          <Field label="Attribution">
            <TextInput value={slide.subtitle ?? ''} onCommit={(v) => patch({ subtitle: str(v) })} />
          </Field>
        </>
      ) : (
        <Field label="Subtitle">
          <TextInput value={slide.subtitle ?? ''} onCommit={(v) => patch({ subtitle: str(v) })} />
        </Field>
      )}
      {(slide.kind === 'bullets' || slide.kind === 'section' || slide.kind === 'two-col') && (
        <Field label="Eyebrow">
          <TextInput value={slide.eyebrow ?? ''} onCommit={(v) => patch({ eyebrow: str(v) })} placeholder="small caps kicker" />
        </Field>
      )}
      {(slide.kind === 'title' || slide.kind === 'section' || slide.kind === 'bullets') && (
        <Field label="Body">
          <TextArea rows={3} value={slide.body ?? ''} onCommit={(v) => patch({ body: str(v) })} />
        </Field>
      )}
      {(slide.kind === 'bullets' || slide.kind === 'two-col') && (
        <Field label={slide.kind === 'two-col' ? 'Left bullets (one per line)' : 'Bullets (one per line)'}>
          <TextArea value={bulletsToText(slide.bullets)} onCommit={(v) => patch({ bullets: textToBullets(v) })} />
        </Field>
      )}
      {slide.kind === 'two-col' && (
        <>
          <Field label="Right bullets (one per line)">
            <TextArea value={bulletsToText(slide.bulletsRight)} onCommit={(v) => patch({ bulletsRight: textToBullets(v) })} />
          </Field>
          <Field label="Column titles">
            <TextInput
              value={(slide.colTitles ?? []).join(' | ')}
              placeholder="Left | Right"
              onCommit={(v) => {
                const parts = v.split('|').map((p) => p.trim()).filter(Boolean)
                patch({ colTitles: parts.length === 2 ? [parts[0], parts[1]] : undefined })
              }}
            />
          </Field>
        </>
      )}
      <Field label="Speaker notes">
        <TextArea rows={3} value={slide.notes ?? ''} onCommit={(v) => patch({ notes: str(v) })} />
      </Field>

      <div className="ioc-row">
        <button type="button" className="ioc-add" onClick={() => { onChange(addSlide(spec, at, 'bullets')); onSelect(at + 1) }}>
          + Slide after
        </button>
      </div>
    </div>
  )
}
