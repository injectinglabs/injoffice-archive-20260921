import { Children, useId, type ButtonHTMLAttributes, type KeyboardEvent, type ReactNode, type Ref } from 'react'
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
}

/**
 * Microsoft-Office-style ribbon: a tab strip with roving-tabindex keyboard
 * navigation, one panel per tab, and labelled `role="group"` command groups.
 * Panels scroll horizontally at narrow widths so groups keep their labels.
 */
export default function Ribbon({ label, tabs, active, onChange, quickAccess, trailing }: RibbonProps) {
  const id = useId()
  const visible = visibleRibbonTabs(tabs)
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
  return <div className="ribbon" aria-label={label}>
    <div className="ribbon-strip">
      {quickAccess && <div className="ribbon-quick-access" role="toolbar" aria-label="Quick access">{quickAccess}</div>}
      <div className="ribbon-tabs" role="tablist" aria-label={`${label} tabs`}>
        {visible.map(tab => <button key={tab.id} type="button" role="tab" id={tabId(tab)} aria-selected={current === tab} aria-controls={panelId(tab)} tabIndex={current === tab ? 0 : -1} onClick={() => onChange(tab.id)} onKeyDown={event => navigate(event, tab)}>{tab.label}</button>)}
      </div>
      {trailing && <div className="ribbon-trailing">{trailing}</div>}
    </div>
    {visible.map(tab => <div key={tab.id} className="ribbon-panel" role="tabpanel" id={panelId(tab)} aria-labelledby={tabId(tab)} hidden={current !== tab}>
      {tab.groups.map(group => <div key={group.id} className="ribbon-group" role="group" aria-label={group.label}>
        <div className="ribbon-group-body">{group.children}</div>
        <span className="ribbon-group-label" aria-hidden="true">{group.label}</span>
      </div>)}
    </div>)}
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

/** Vertical stack of control rows inside one group (Office stacks Font as two rows). */
export function RibbonRows({ children }: { children: ReactNode }) {
  return <div className="ribbon-rows">{children}</div>
}
