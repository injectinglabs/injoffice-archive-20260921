import { defineConfig } from 'vitepress'
import { readFileSync } from 'node:fs'

const inventory = JSON.parse(readFileSync(new URL('../../../docs/api-reference.json', import.meta.url), 'utf8'))

export default defineConfig({
  title: 'InjOffice Docs',
  description: 'Build document workflows with InjOffice: TypeScript and Go guides, tested code examples, native file contracts, and model-neutral agent integration.',
  lang: 'en-US',
  base: process.env.DOCS_BASE || '/',
  srcExclude: ['README.md', 'tests/**', 'scripts/**'],
  cleanUrls: false,
  lastUpdated: false,
  head: [['meta', { name: 'theme-color', content: '#245cc5' }]],
  markdown: { lineNumbers: false },
  themeConfig: {
    logo: '/logo.svg',
    siteTitle: 'InjOffice / Docs',
    nav: [
      { text: 'Guides', link: '/getting-started/quickstart' },
      { text: 'Reference', link: '/reference/' },
      { text: 'Try the demo', link: 'https://injoffice.com/' },
    ],
    socialLinks: [{ icon: 'github', link: 'https://github.com/injectinglabs/injoffice' }],
    search: { provider: 'local' },
    outline: { level: [2, 3], label: 'On this page' },
    editLink: { pattern: 'https://github.com/injectinglabs/injoffice/edit/main/apps/docs/:path', text: 'Edit this guide on GitHub' },
    sidebar: [
      { text: 'Start here', items: [
        { text: 'Introduction', link: '/' },
        { text: 'Quickstart', link: '/getting-started/quickstart' },
        { text: 'Choose your packages', link: '/getting-started/packages' },
        { text: 'Architecture & runtimes', link: '/getting-started/architecture' },
        { text: 'Support & limitations', link: '/getting-started/support' },
      ] },
      { text: 'Document guides', items: [
        { text: 'Spreadsheets / XLSX', link: '/guides/spreadsheets' },
        { text: 'Documents / DOCX', link: '/guides/documents' },
        { text: 'Presentations / PPTX', link: '/guides/presentations' },
        { text: 'PDFs', link: '/guides/pdf' },
        { text: 'Charts, pivots & formulas', link: '/guides/spreadsheet-tools' },
      ] },
      { text: 'Agents & integration', items: [
        { text: 'Your first agent workflow', link: '/agents/quickstart' },
        { text: 'Tool protocol & adapters', link: '/agents/integration' },
        { text: 'Approval & verification', link: '/agents/safety' },
        { text: 'Browser workers & assets', link: '/integration/browser' },
        { text: 'Go & optional HTTP server', link: '/integration/server' },
        { text: 'Collaboration & history', link: '/integration/collaboration' },
        { text: 'Security & production checklist', link: '/integration/security' },
        { text: 'Troubleshooting', link: '/integration/troubleshooting' },
      ] },
      { text: 'Reference', collapsed: true, items: [
        { text: 'Reference index', link: '/reference/' },
        ...inventory.packages.map((pkg: { name: string }) => ({ text: pkg.name.replace('@injoffice/', ''), link: `/reference/generated/packages/${pkg.name.replace('@injoffice/', '')}` })),
      ] },
      { text: 'Project', items: [
        { text: 'Examples & verification', link: '/examples/' },
        { text: 'Contributing', link: '/project/contributing' },
        { text: 'Maintaining these docs', link: '/project/documentation' },
      ] },
    ],
    footer: { message: 'Apache-2.0 project. Capability-scoped tools, not full Office compatibility.', copyright: 'InjOffice · Injecting Inc.' },
  },
})
