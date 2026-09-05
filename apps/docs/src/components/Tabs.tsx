import { useState, type ReactNode } from 'react'

export function Tabs({
  tabs,
}: {
  tabs: { id: string; label: string; content: ReactNode }[]
}) {
  const [active, setActive] = useState(tabs[0]?.id)
  const current = tabs.find((tab) => tab.id === active) ?? tabs[0]
  return (
    <div>
      <div className="tabs" role="tablist">
        {tabs.map((tab) => (
          <button
            key={tab.id}
            type="button"
            role="tab"
            aria-selected={tab.id === current.id}
            onClick={() => setActive(tab.id)}
          >
            {tab.label}
          </button>
        ))}
      </div>
      <div role="tabpanel">{current?.content}</div>
    </div>
  )
}
