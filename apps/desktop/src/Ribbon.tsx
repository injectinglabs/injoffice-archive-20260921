import { Children, createContext, useContext, useId, useLayoutEffect, useRef, useState, type ButtonHTMLAttributes, type KeyboardEvent, type ReactNode, type Ref } from 'react'
import RibbonIcon, { type RibbonIconName } from './RibbonIcons'
import { shortcutKeys, shortcutTooltip, type ShortcutId } from './shortcuts'
// ribbon.css is imported by each editor next to its own stylesheet so that
// components using RibbonButton stay CSS-free for unit-test bundlers.

/** A labelled command group inside a ribbon tab. Groups without content are not rendered. */
export interface RibbonGroupSpec { id: string; label: string; children?: ReactNode }
/** A ribbon tab; tabs whose groups are all empty are not rendered, so no tab ever appears blank. */
export interface RibbonTabSpec { id: string; label: string; groups: RibbonGroupSpec[] }

const hasContent = (node: ReactNode) => Children.toArray(node).length > 0

/** Tabs that have at least one populated group, with their empty groups removed. */
export function visibleRibbonTabs(tabs: RibbonTabSpec[]): RibbonTabSpec[] {
  return tabs.map(tab => ({ ...tab, groups: tab.groups.filter(group => hasContent(group.children)) })).filter(tab => tab.groups.length > 0)
}

/** Shell File surface. Legacy groups are retained for standalone ribbon consumers. */
export interface WorkspaceFileGroups { before: RibbonGroupSpec[]; after: RibbonGroupSpec[]; backstage?: { open(): void; render(groups: RibbonGroupSpec[]): ReactNode } }
export const WorkspaceFileGroupsContext = createContext<WorkspaceFileGroups | null>(null)

/** The editor's tabs with the workspace File groups merged into (or added as) the File tab. */
export function withWorkspaceFileGroups(tabs: RibbonTabSpec[], workspace: WorkspaceFileGroups | null): RibbonTabSpec[] {
  if (!workspace) return tabs
  const file = tabs.find(tab => tab.id === 'File') ?? { id: 'File', label: 'File', groups: [] }
  const merged = { ...file, groups: [...workspace.before, ...file.groups, ...workspace.after] }
  return tabs.includes(file) ? tabs.map(tab => tab === file ? merged : tab) : [merged, ...tabs]
}

/** How the active panel is laid out at the current width. `visible` is -1 while every group fits. */
export interface RibbonPanelLayout { compact: boolean; visible: number }
/** Width reserved at the right edge for the overflow chevron. */
export const RIBBON_OVERFLOW_WIDTH = 44

/**
 * How many leading groups fit in `available` px once the overflow chevron has its
 * own room. At least one group stays on the ribbon, so the panel is never only a
 * chevron; the rest move into the overflow popover.
 */
export function ribbonGroupsThatFit(widths: number[], available: number, overflow = RIBBON_OVERFLOW_WIDTH): number {
  const total = widths.reduce((sum, width) => sum + width, 0)
  if (total <= available) return widths.length
  let used = 0, fit = 0
  for (const width of widths) {
    if (used + width > available - overflow) break
    used += width
    fit++
  }
  return Math.max(1, fit)
}

export interface RibbonProps {
  /** Accessible name of the ribbon, e.g. "Document tools". */
  label: string
  tabs: RibbonTabSpec[]
  /** Id of the selected tab. Falls back to the first visible tab when it is not visible. */
  active: string
  onChange(tab: string): void
  /** Quick Access commands (Undo/Redo…) shown at the top-left, before the tab strip. */
  quickAccess?: ReactNode
  /** Content shown at the end of the tab strip. */
  trailing?: ReactNode
  /** Test hook: use this panel layout instead of measuring the panel. */
  panelLayout?: RibbonPanelLayout
}

/**
 * Microsoft-Office-style ribbon: a tab strip with roving-tabindex keyboard
 * navigation, one panel per tab, and labelled `role="group"` command groups.
 * At narrow widths a panel collapses to icons and then moves its trailing groups
 * behind an overflow chevron, so it never overflows the window unannounced.
 */
export default function Ribbon({ label, tabs, active, onChange, quickAccess, trailing, panelLayout }: RibbonProps) {
  const id = useId()
  const workspace = useContext(WorkspaceFileGroupsContext)
  const fileGroups = tabs.flatMap(tab => tab.id === 'File' ? tab.groups : tab.groups.filter(group => group.id === 'export'))
  const visible = visibleRibbonTabs(workspace?.backstage ? tabs.filter(tab => tab.id !== 'File') : withWorkspaceFileGroups(tabs, workspace))
  const current = visible.find(tab => tab.id === active) ?? visible[0]
  const tabId = (tab: RibbonTabSpec) => `${id}-tab-${tab.id}`
  const panelId = (tab: RibbonTabSpec) => `${id}-panel-${tab.id}`
  const navigate = (event: KeyboardEvent<HTMLButtonElement>, tab: RibbonTabSpec) => {
    if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return
    event.preventDefault()
    const index = event.key === 'Home' ? 0 : event.key === 'End' ? visible.length - 1 : (visible.indexOf(tab) + (event.key === 'ArrowRight' ? 1 : -1) + visible.length) % visible.length
    onChange(visible[index].id)
    ;(event.currentTarget.parentElement?.children[index] as HTMLButtonElement | undefined)?.focus()
  }
  // Narrow windows collapse the panel the way Office does: group labels and button
  // text go first, then whole groups move behind an overflow chevron. Widths are
  // measured once per mode and reused, so a resize never thrashes the layout.
  const panel = useRef<HTMLDivElement>(null)
  const fullWidths = useRef<number[]>([])
  const compactWidths = useRef<number[]>([])
  const [measured, setMeasured] = useState<RibbonPanelLayout>({ compact: false, visible: -1 })
  const [overflowOpen, setOverflowOpen] = useState(false)
  const layout = panelLayout ?? measured
  const signature = current ? `${current.id}:${current.groups.map(group => group.id).join(',')}` : ''
  useLayoutEffect(() => {
    fullWidths.current = []
    compactWidths.current = []
    setMeasured({ compact: false, visible: -1 })
    setOverflowOpen(false)
  }, [signature])
  useLayoutEffect(() => {
    const node = panel.current
    if (panelLayout || !node || typeof ResizeObserver !== 'function') return
    const apply = () => {
      const groups = Array.from(node.children).filter(child => child.classList.contains('ribbon-group')) as HTMLElement[]
      if (measured.visible === -1 && groups.length) {
        const widths = groups.map(group => Math.ceil(group.getBoundingClientRect().width) + 4)
        if (measured.compact) compactWidths.current = widths; else fullWidths.current = widths
      }
      // A panel in a hidden editor has no layout; keep the last good decision instead of measuring zeroes.
      if (!node.clientWidth) return
      const style = node.ownerDocument.defaultView?.getComputedStyle(node)
      const available = node.clientWidth - (parseFloat(style?.paddingLeft ?? '0') + parseFloat(style?.paddingRight ?? '0'))
      const total = (widths: number[]) => widths.reduce((sum, width) => sum + width, 0)
      let next: RibbonPanelLayout
      if (!fullWidths.current.length) next = { compact: false, visible: -1 }
      else if (total(fullWidths.current) <= available) next = { compact: false, visible: -1 }
      else if (!compactWidths.current.length) next = { compact: true, visible: -1 }
      else if (total(compactWidths.current) <= available) next = { compact: true, visible: -1 }
      else next = { compact: true, visible: ribbonGroupsThatFit(compactWidths.current, available) }
      setMeasured(previous => previous.compact === next.compact && previous.visible === next.visible ? previous : next)
    }
    apply()
    const observer = new ResizeObserver(apply)
    observer.observe(node)
    return () => observer.disconnect()
  }, [measured, signature, panelLayout])
  const renderGroup = (group: RibbonGroupSpec) => <div key={group.id} className="ribbon-group" role="group" aria-label={group.label}>
    <div className="ribbon-group-body">{group.children}</div>
    <span className="ribbon-group-label" aria-hidden="true">{group.label}</span>
  </div>
  return <div className="ribbon" aria-label={label}>
    {workspace?.backstage && workspace.backstage.render(fileGroups)}
    <div className="ribbon-strip">
      {workspace?.backstage && <button type="button" className="ribbon-file" aria-haspopup="dialog" onClick={workspace.backstage.open}>File</button>}
      {quickAccess && <div className="ribbon-quick-access" role="toolbar" aria-label="Quick access">{quickAccess}</div>}
      <div className="ribbon-tabs" role="tablist" aria-label={`${label} tabs`}>
        {visible.map(tab => <button key={tab.id} type="button" role="tab" id={tabId(tab)} aria-selected={current === tab} aria-controls={panelId(tab)} tabIndex={current === tab ? 0 : -1} onClick={() => onChange(tab.id)} onKeyDown={event => navigate(event, tab)}>{tab.label}</button>)}
      </div>
      {trailing && <div className="ribbon-trailing">{trailing}</div>}
    </div>
    {visible.map(tab => {
      const active = current === tab
      const shown = active && layout.visible >= 0 ? tab.groups.slice(0, layout.visible) : tab.groups
      const hidden = active && layout.visible >= 0 ? tab.groups.slice(layout.visible) : []
      return <div key={tab.id} ref={active ? panel : undefined} className={['ribbon-panel', active && layout.compact ? 'ribbon-panel-compact' : ''].join(' ').trim()} role="tabpanel" id={panelId(tab)} aria-labelledby={tabId(tab)} hidden={!active}>
        {shown.map(renderGroup)}
        {hidden.length > 0 && <div className="ribbon-overflow">
          <button type="button" className="ribbon-button ribbon-overflow-button" aria-label={`More commands (${hidden.length} groups)`} title="More commands" aria-haspopup="true" aria-expanded={overflowOpen} onClick={() => setOverflowOpen(open => !open)}>
            <RibbonIcon name="chevronDown" />
          </button>
          {overflowOpen && <div className="ribbon-overflow-popover">{hidden.map(renderGroup)}</div>}
        </div>}
      </div>
    })}
  </div>
}

export interface RibbonButtonProps extends Omit<ButtonHTMLAttributes<HTMLButtonElement>, 'title'> {
  /** Forwarded to the host button (React 19 passes ref as a prop). */
  ref?: Ref<HTMLButtonElement>
  icon: RibbonIconName
  /** Short visible label. Also the tooltip unless `title` is given. */
  label: string
  /** Shortcut shown in the tooltip and exposed through aria-keyshortcuts. */
  shortcut?: ShortcutId
  /** Tooltip body; defaults to the label. The shortcut is appended when present. */
  title?: string
  /** Show the icon only; the label stays available to assistive technology. */
  labelHidden?: boolean
}

/** Icon + label command button whose tooltip always carries the shortcut where one exists. */
export function RibbonButton({ icon, label, shortcut, title, labelHidden, className, children, ...rest }: RibbonButtonProps) {
  return <button type="button" {...rest} className={['ribbon-button', labelHidden ? 'ribbon-button-icon-only' : '', className ?? ''].join(' ').trim()} title={shortcutTooltip(title ?? label, shortcut)} aria-label={labelHidden ? label : rest['aria-label']} aria-keyshortcuts={shortcut ? shortcutKeys(shortcut) : undefined}>
    <RibbonIcon name={icon} />
    {!labelHidden && <span className="ribbon-button-label">{label}</span>}
    {children}
  </button>
}

/** One entry of a ribbon combo box. */
export interface RibbonComboOption { value: string; label: string }

export interface RibbonComboProps {
  /** Accessible name; also the tooltip unless `title` is given. */
  label: string
  title?: string
  /** The resolved value at the caret. An empty string means "not known / mixed": Office shows a blank box, never a word. */
  value: string
  options: RibbonComboOption[]
  onChange(value: string): void
  disabled?: boolean
  /** Width class from ribbon.css (`ribbon-combo-font`, `-size`, `-style`, `-wide`). Widths are fixed so groups do not reflow. */
  className?: string
}

/**
 * Office's ribbon combo box: a fixed-width select that shows the value in effect
 * at the caret. When the value is unknown or mixed across the selection the box
 * is empty (just the chevron), exactly like Word — never a placeholder word.
 */
export function RibbonCombo({ label, title, value, options, onChange, disabled, className }: RibbonComboProps) {
  return <select className={['ribbon-combo', className ?? ''].join(' ').trim()} aria-label={label} title={title ?? label} disabled={disabled} value={value} onChange={event => onChange(event.target.value)}>
    {(value === '' || !options.some(option => option.value === value)) && <option value="" />}
    {options.map(option => <option key={option.value} value={option.value}>{option.label}</option>)}
  </select>
}

/** A palette entry: `[hex, name]`. */
export type RibbonColor = [string, string]

export interface ColorButtonProps {
  /** Accessible name, e.g. "Text color". */
  label: string
  icon: RibbonIconName
  /** The colour in effect, or undefined when it is unknown or mixed (the underline then stays empty). */
  value?: string
  colors: RibbonColor[]
  onChange(color: string): void
  disabled?: boolean
  title?: string
}

/**
 * Office's colour control: an icon button carrying the current colour as an
 * underline plus a chevron that opens the palette. Never a native
 * `<input type="color">` with a caption above it.
 */
export function ColorButton({ label, icon, value, colors, onChange, disabled, title }: ColorButtonProps) {
  const [open, setOpen] = useState(false)
  const swatches = value && !colors.some(([color]) => color.toUpperCase() === value.toUpperCase()) ? [[value.toUpperCase(), 'Current color'] as RibbonColor, ...colors] : colors
  return <span className="ribbon-color" onBlur={event => { if (!event.currentTarget.contains(event.relatedTarget as Node | null)) setOpen(false) }} onKeyDown={event => { if (event.key === 'Escape' && open) { event.stopPropagation(); setOpen(false) } }}>
    <button type="button" className="ribbon-button ribbon-color-button" aria-label={label} title={title ?? label} aria-haspopup="true" aria-expanded={open} disabled={disabled} onClick={() => setOpen(shown => !shown)}>
      <span className="ribbon-color-glyph"><RibbonIcon name={icon} /><span className="ribbon-color-bar" style={value ? { background: value } : undefined} /></span>
      <span className="ribbon-chevron" aria-hidden="true" />
    </button>
    {open && <span className="ribbon-color-palette" role="group" aria-label={`${label} palette`}>
      {swatches.map(([color, name]) => <button key={color} type="button" className="ribbon-swatch" aria-label={name} title={name} aria-pressed={!!value && color.toUpperCase() === value.toUpperCase()} style={{ background: color }} onClick={() => { onChange(color); setOpen(false) }} />)}
    </span>}
  </span>
}

/** Office's horizontal alignment commands per host, in ribbon order. Word adds justify and distribute. */
export const RIBBON_ALIGNMENTS: Record<'docx' | 'xlsx' | 'pptx', Array<{ value: string; label: string; icon: RibbonIconName }>> = {
  xlsx: [{ value: 'left', label: 'Align left', icon: 'alignLeft' }, { value: 'center', label: 'Center', icon: 'alignCenter' }, { value: 'right', label: 'Align right', icon: 'alignRight' }],
  pptx: [{ value: 'left', label: 'Align left', icon: 'alignLeft' }, { value: 'center', label: 'Center', icon: 'alignCenter' }, { value: 'right', label: 'Align right', icon: 'alignRight' }],
  docx: [{ value: 'left', label: 'Align left', icon: 'alignLeft' }, { value: 'center', label: 'Center', icon: 'alignCenter' }, { value: 'right', label: 'Align right', icon: 'alignRight' }, { value: 'both', label: 'Justify', icon: 'alignJustify' }, { value: 'distribute', label: 'Distribute', icon: 'alignDistribute' }],
}
/** Excel's vertical alignment commands, in ribbon order. */
export const RIBBON_VERTICAL_ALIGNMENTS: Array<{ value: string; label: string; icon: RibbonIconName }> = [
  { value: 'top', label: 'Top align', icon: 'alignTop' },
  { value: 'middle', label: 'Middle align', icon: 'alignMiddle' },
  { value: 'bottom', label: 'Bottom align', icon: 'alignBottom' },
]

export interface AlignmentTogglesProps {
  /** The horizontal alignment in effect; 'general' or undefined presses nothing, as Excel does. */
  horizontal?: string
  /** The vertical alignment in effect. Omit `onVertical` for a host that has no vertical alignment. */
  vertical?: string
  onHorizontal(value: string): void
  onVertical?(value: string): void
  disabled?: boolean
  /** Which host's horizontal set to offer. Defaults to the three Excel and PowerPoint share. */
  kind?: 'docx' | 'xlsx' | 'pptx'
  /** Extra class on the wrapper, for a host that positions the block itself. */
  className?: string
}

/**
 * Excel's alignment block: top/middle/bottom over left/centre/right, as toggle
 * icons that show the alignment in effect — never a pair of dropdowns.
 */
export function AlignmentToggles({ horizontal, vertical, onHorizontal, onVertical, disabled, kind = 'xlsx', className }: AlignmentTogglesProps) {
  const toggle = (option: { value: string; label: string; icon: RibbonIconName }, current: string | undefined, apply: (value: string) => void) =>
    <RibbonButton key={option.value} icon={option.icon} label={option.label} labelHidden aria-pressed={current === option.value} disabled={disabled} onClick={() => apply(option.value)} />
  return <div className={['ribbon-alignment-toggles', className ?? ''].join(' ').trim()}>
    {onVertical && <div className="ribbon-alignment-row">{RIBBON_VERTICAL_ALIGNMENTS.map(option => toggle(option, vertical, onVertical))}</div>}
    <div className="ribbon-alignment-row">{RIBBON_ALIGNMENTS[kind].map(option => toggle(option, horizontal, onHorizontal))}</div>
  </div>
}

/** Vertical stack of control rows inside one group (Office stacks Font as two rows). */
export function RibbonRows({ children }: { children: ReactNode }) {
  return <div className="ribbon-rows">{children}</div>
}
