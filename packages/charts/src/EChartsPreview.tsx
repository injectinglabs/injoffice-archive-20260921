import { useEffect, useRef, type CSSProperties } from 'react'
import * as echarts from 'echarts'

export interface EChartsPreviewProps {
  option: Record<string, unknown>
  ariaLabel: string
  className?: string
  style?: CSSProperties
}

function usesDarkTheme(): boolean {
  const root = document.documentElement
  return root.classList.contains('univer-dark') || root.dataset.theme === 'dark'
}

/**
 * Renderer-only ECharts surface for hosts that already have an adapter option.
 * It replaces incompatible options, follows container and theme changes, and
 * disposes the ECharts instance and observers when it unmounts.
 */
export function EChartsPreview({ option, ariaLabel, className, style }: EChartsPreviewProps) {
  const hostRef = useRef<HTMLDivElement | null>(null)
  const instanceRef = useRef<echarts.ECharts | null>(null)
  const optionRef = useRef(option)
  optionRef.current = option

  useEffect(() => {
    const host = hostRef.current
    if (!host) return

    const mountOrResize = () => {
      if (host.clientWidth === 0 || host.clientHeight === 0) return
      if (!instanceRef.current) {
        instanceRef.current = echarts.init(host, usesDarkTheme() ? 'dark' : undefined)
        instanceRef.current.setOption(optionRef.current, { notMerge: true })
      } else {
        instanceRef.current.resize()
      }
    }

    const resizeObserver = new ResizeObserver(mountOrResize)
    resizeObserver.observe(host)
    const themeObserver = new MutationObserver(() => {
      instanceRef.current?.dispose()
      instanceRef.current = null
      mountOrResize()
    })
    themeObserver.observe(document.documentElement, { attributes: true, attributeFilter: ['class', 'data-theme'] })
    mountOrResize()

    return () => {
      resizeObserver.disconnect()
      themeObserver.disconnect()
      instanceRef.current?.dispose()
      instanceRef.current = null
    }
  }, [])

  useEffect(() => {
    instanceRef.current?.setOption(option, { notMerge: true })
  }, [option])

  return (
    <div
      ref={hostRef}
      className={className}
      style={style}
      role="img"
      aria-label={ariaLabel}
      data-chart-renderer="echarts"
    />
  )
}
