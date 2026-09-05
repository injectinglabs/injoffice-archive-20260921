import { Tabs } from './Tabs'
import { CodeBlock } from './CodeBlock'

export function Install({ packages }: { packages: string }) {
  return (
    <Tabs
      tabs={[
        { id: 'npm', label: 'npm', content: <CodeBlock language="bash" title="npm" code={`npm install ${packages}`} /> },
        { id: 'pnpm', label: 'pnpm', content: <CodeBlock language="bash" title="pnpm" code={`pnpm add ${packages}`} /> },
        { id: 'yarn', label: 'yarn', content: <CodeBlock language="bash" title="yarn" code={`yarn add ${packages}`} /> },
      ]}
    />
  )
}
