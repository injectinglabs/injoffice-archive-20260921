# @injoffice/univer-sheets

Configurable composition of Univer's Apache-2.0 spreadsheet plugins for an
InjOffice editor. The default bundle enables the complete public Univer Sheets
surface; individual capabilities can be omitted without editing menu schemas.

```bash
npm install @injoffice/univer-sheets @univerjs/core react react-dom rxjs
```

```ts
import { LocaleType } from '@univerjs/core'
import {
  createInjOfficeSheetsFeatureComposition,
  createOssUniver,
  createUniverFeatureProvider,
} from '@injoffice/univer-sheets/browser'
import { registerUniverChartCommands } from '@injoffice/charts/browser'
import '@injoffice/univer-sheets/styles.css'

let editor: ReturnType<typeof createOssUniver>
const features = createInjOfficeSheetsFeatureComposition({
  container: 'app',
  providers: {
    charts: createUniverFeatureProvider(({ menuHidden$ }) =>
      registerUniverChartCommands(editor.univer, editor.univerAPI, {
        createInput: () => ({ type: 'Column' }),
        menuHidden$,
      })),
  },
  features: {
    threadComments: false,
    notes: false,
    charts: { enabled: true, visible: true },
  },
})

editor = createOssUniver({
  locale: LocaleType.EN_US,
  locales: { [LocaleType.EN_US]: features.bundle.locale },
  presets: features.bundle.presets,
})

editor.univerAPI.createWorkbook({})
await features.activate()
await features.configure({ charts: { enabled: true, visible: false } })
```

The preset composition APIs are browser-only and are exported from
`@injoffice/univer-sheets/browser`. The root entry keeps the renderer-neutral
feature registry, provider adapter, and Univer bootstrap importable in Node
hosts without eagerly loading Univer's canvas renderer.

`InjOfficeSheetsFeatureComposition` accepts one configuration vocabulary for
all optional OSS presets and advanced InjOffice packages. OSS presets are fixed
at construction because Univer exposes no safe plugin-unregistration API.
Advanced providers remain lazy and runtime configurable. The registry validates
dependencies and server capabilities before loading, serializes transitions,
rolls back failed activation/visibility changes, and disposes in reverse order.

`createUniverFeatureProvider` adapts real command registrars to that lifecycle.
Its `menuHidden$` works with the chart, pivot, shape, sparkline, connector,
outline, and print browser adapters, so their menu contribution can hide or
show without unloading commands or models. A provider without `setVisible`
fails an attempted live visibility change instead of reporting false state.

Core editing, formulas, and number formats are always registered. Optional
features default to enabled: conditional formatting, data validation, drawing
and image insertion, filtering, find/replace, hyperlinks, notes, sorting,
tables, and threaded comments. Setting a feature to `false` omits its model and
UI plugins, removing its ribbon and context-menu contributions.

The `/browser` convenience entry and `styles.css` include all optional feature
modules. Feature flags control registration and visibility; hosts that also
need per-feature download-size splitting can compose the individual upstream
preset packages instead.

This package only composes public `@univerjs/*` plugins. It neither imports nor
repackages `@univerjs-pro/*`. Advanced packages remain optional peer/host
integrations: the generic provider bridge does not import them or erase their
feature-specific configuration. `registerInjOfficeInsertMenu` remains available
as a smaller legacy bridge for host-owned chart, pivot, and shape actions.
