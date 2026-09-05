import { lazy, type ComponentType, type LazyExoticComponent } from 'react'

export type PreloadableComponent = {
  component: LazyExoticComponent<ComponentType>
  preload: () => Promise<void>
}

export function preloadableLazy<T extends ComponentType>(
  loader: () => Promise<{ default: T }>,
): PreloadableComponent {
  let modulePromise: Promise<{ default: T }> | undefined
  const load = () => {
    if (!modulePromise) {
      const attempt = loader()
      modulePromise = attempt
      void attempt.catch(() => {
        if (modulePromise === attempt) modulePromise = undefined
      })
    }
    return modulePromise
  }

  return {
    component: lazy(load),
    preload: () => load().then(() => undefined),
  }
}

export function preloadableLazyNamed<T extends Record<string, unknown>, K extends keyof T>(
  loader: () => Promise<T>,
  name: K,
): PreloadableComponent {
  return preloadableLazy(async () => ({ default: (await loader())[name] as ComponentType }))
}
