import { BehaviorSubject, map, type Observable } from 'rxjs'
import type {
  InjOfficeFeatureActivation,
  InjOfficeFeatureContext,
  InjOfficeFeatureProvider,
  InjOfficeFeatureKey,
} from './featureRegistry'

export interface UniverFeatureRegistrationContext extends InjOfficeFeatureContext {
  /** Bind this to Univer menu items' `hidden$`. Commands and models remain active. */
  menuHidden$: Observable<boolean>
}

export interface UniverFeatureRegistration {
  dispose(): void | Promise<void>
}

export interface UniverFeatureProviderOptions {
  dependencies?: InjOfficeFeatureKey[]
  requiredServerCapabilities?: string[]
  defaultEnabled?: boolean
  defaultVisible?: boolean
}

export type UniverFeatureRegistrar = (
  context: Readonly<UniverFeatureRegistrationContext>,
) => UniverFeatureRegistration | Promise<UniverFeatureRegistration>

/**
 * Adapt a real InjOffice/Univer command registrar to the lazy feature registry.
 * The observable keeps menu visibility independent from command/model lifetime;
 * the registrar still owns all concrete registrations and their disposal.
 */
export function createUniverFeatureProvider(
  register: UniverFeatureRegistrar,
  options: UniverFeatureProviderOptions = {},
): InjOfficeFeatureProvider {
  if (typeof register !== 'function') throw new TypeError('feature registrar must be a function')
  return {
    ...options,
    load: async () => ({
      activate: async (context): Promise<InjOfficeFeatureActivation> => {
        const visible = new BehaviorSubject(context.visible)
        let registration: UniverFeatureRegistration
        try {
          registration = await register(Object.freeze({
            ...context,
            menuHidden$: visible.pipe(map((value) => !value)),
          }))
        } catch (cause) {
          visible.complete()
          throw cause
        }
        if (!registration || typeof registration.dispose !== 'function') {
          visible.complete()
          throw new TypeError('feature registrar must return a disposable registration')
        }
        let disposed = false
        return {
          setVisible(next) {
            if (disposed) throw new Error('feature registration is disposed')
            visible.next(next)
          },
          async dispose() {
            if (disposed) return
            const previous = visible.value
            visible.next(false)
            try {
              await registration.dispose()
              disposed = true
              visible.complete()
            } catch (cause) {
              visible.next(previous)
              throw cause
            }
          },
        }
      },
    }),
  }
}
