import { useEffect, useRef } from 'react'
import * as echarts from 'echarts'
import { subscribeChart, getChartOption } from './registry'

// The React component Univer mounts inside a float-dom. It owns exactly one
// echarts instance and nothing else: the option comes from the registry (the
// ChartManager writes there whenever specs or source cells change), so the
// component has no Univer knowledge and re-renders on registry pushes only.
//
// Sized by its float-dom container (width/height 100%); a ResizeObserver keeps
// echarts in sync when the user drags the float-dom's resize handles.

export function ChartFloat({ data }: { data?: { chartId?: string } }) {
  const chartId = data?.chartId
  const hostRef = useRef<HTMLDivElement | null>(null)

  useEffect(() => {
    if (!chartId || !hostRef.current) return
    const host = hostRef.current
    // Univer mounts float-dom components BEFORE the float's layout pass, so
    // the host is 0×0 on first render and echarts.init would warn and paint
    // nothing. Initialize lazily from the ResizeObserver, on the first tick
    // where the host has real dimensions; afterwards the same observer keeps
    // the instance sized through user drags of the float's resize handles.
    let instance: echarts.ECharts | null = null

    const hostIsDark = () => {
      const root = document.documentElement
      return root.classList.contains('univer-dark') || root.dataset.theme === 'dark'
    }

    const apply = () => {
      const option = getChartOption(chartId)
      // notMerge: a spec change (e.g. column→pie) must fully replace the
      // previous option, not merge series shapes into each other.
      if (instance && option) instance.setOption(option, { notMerge: true })
    }
    const unsubscribe = subscribeChart(chartId, apply)

    const ensureInstance = () => {
      if (host.clientWidth === 0 || host.clientHeight === 0) return
      const dark = hostIsDark()
      if (!instance) {
        instance = echarts.init(host, dark ? 'dark' : undefined)
        apply()
        return
      }
      instance.resize()
    }

    const ro = new ResizeObserver(ensureInstance)
    ro.observe(host)
    const themeObserver = new MutationObserver(() => {
      if (!instance) return
      instance.dispose()
      instance = null
      ensureInstance()
    })
    themeObserver.observe(document.documentElement, { attributes: true, attributeFilter: ['class', 'data-theme'] })

    return () => {
      unsubscribe()
      ro.disconnect()
      themeObserver.disconnect()
      instance?.dispose()
    }
  }, [chartId])

  return (
    <div
      ref={hostRef}
      className="ioc-chart-float"
      style={{
        width: '100%',
        height: '100%',
        background: 'var(--paper, #fff)',
        border: '1px solid var(--rule, #d9dee2)',
        borderRadius: 6,
        boxSizing: 'border-box',
        overflow: 'hidden',
      }}
    />
  )
}
