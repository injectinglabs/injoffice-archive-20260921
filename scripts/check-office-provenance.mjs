import { readFileSync } from 'node:fs'
import { createHash } from 'node:crypto'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const failures = []
const requireValue = (condition, message) => { if (!condition) failures.push(message) }
const digest = (data) => createHash('sha256').update(data).digest('hex')

const excelFixtureRevision = 'ad12607e3fb4080fc161e72f4d2b93ce86c29e66'
const excelFixtureRoot = 'go/xlsxpatch/testdata/excel-authored'
const excelFixtureSources = {
  'conditional-formatting-samples.xlsx': ['fa17f45f47e0766f13b9cbbf9e63b83962ea7852606bfe0b33807fbfbae5ec64', '5affc0ac28bd74b7877c492a6bb541040fbf2a88'],
  'email-chart-table.xlsx': ['1da5a2011f4f1e6f2e4407f148e814cae47e93d97335c6819d582b6cf433e1be', 'd934dc31638bf642fdeb389d2b953261ce111a5d'],
  'happy-tree.xlsx': ['c08f0bb099770a475556af8a78d9d6e46296e9791bf9af56c70571c1fa3cb513', 'c4b91c1b56212df87ae4fba063ab20a66327a203'],
}

const pptxFixtureRevision = '40bc19c8261b158ff92b0ef26b35d4ffcf570458'
const pptxFixtureRoot = 'go/pptxpatch/testdata/powerpoint-authored'
const pptxFixtureSources = {
  'attendee-survey-qr.pptx': ['b4a503d90634e117ca53fb62d6aaf0657b78bfe0269b7314a313077a4899e095', '79ff3f21150cee5f6edd146e403fd8cbd2495aba'],
}

const excelReadme = readFileSync(resolve(root, excelFixtureRoot, 'README.md'), 'utf8')
for (const [name, [sha256, blob]] of Object.entries(excelFixtureSources)) {
  const bytes = readFileSync(resolve(root, excelFixtureRoot, name))
  requireValue(digest(bytes) === sha256, `Excel-authored fixture ${name} differs from its pinned byte digest`)
  requireValue(excelReadme.includes(`## \`${name}\``) && excelReadme.includes(excelFixtureRevision) && excelReadme.includes(blob) && excelReadme.includes(sha256), `Excel-authored fixture ${name} lacks exact revision/blob/digest provenance`)
}
const pptxReadme = readFileSync(resolve(root, pptxFixtureRoot, 'README.md'), 'utf8')
for (const [name, [sha256, blob]] of Object.entries(pptxFixtureSources)) {
  const bytes = readFileSync(resolve(root, pptxFixtureRoot, name))
  requireValue(digest(bytes) === sha256, `PowerPoint-authored fixture ${name} differs from its pinned byte digest`)
  requireValue(pptxReadme.includes(`## \`${name}\``) && pptxReadme.includes(pptxFixtureRevision) && pptxReadme.includes(blob) && pptxReadme.includes(sha256) && pptxReadme.includes('Changes: none'), `PowerPoint-authored fixture ${name} lacks exact revision/blob/digest provenance`)
  requireValue(pptxReadme.includes('Microsoft Macintosh PowerPoint') && pptxReadme.includes('office-scripts-docs') && pptxReadme.includes('Excel-only'), `PowerPoint-authored fixture ${name} lacks Microsoft PowerPoint provenance and Excel-only sample-set attestation`)
}
const excelBridge = JSON.parse(readFileSync(resolve(root, 'go/xlsxpatch/testdata/native-xlsx-v1/valid/excel-authored-happy-tree.json'), 'utf8'))
requireValue(excelBridge.source?.package_sha256 === `sha256:${excelFixtureSources['happy-tree.xlsx'][0]}` && excelBridge.revision === `rev:${excelFixtureSources['happy-tree.xlsx'][0]}`, 'Excel-authored Go/TypeScript bridge is not bound to the unchanged package bytes')

const projectLicense = readFileSync(resolve(root, 'LICENSE'))
const thirdPartyApache = readFileSync(resolve(root, 'THIRD_PARTY_LICENSES/APACHE-2.0.txt'))
requireValue(digest(projectLicense) === 'c71d239df91726fc519c6eb72d318ec65820627232b2f796219e87dcf35d0ab4' && projectLicense.equals(thirdPartyApache), 'Apache-2.0 third-party license text is absent, modified, or differs from the project Apache text')

if (failures.length > 0) {
  console.error(failures.map((failure) => `- ${failure}`).join('\n'))
  process.exit(1)
}
console.log(`Validated ${Object.keys(excelFixtureSources).length} Excel-authored fixtures and ${Object.keys(pptxFixtureSources).length} PowerPoint-authored fixtures.`)
