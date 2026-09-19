const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const React = require('react');

const root = path.resolve(__dirname, '..');

test('index.html, main.tsx, and Vite outDir match the Electron renderer host', () => {
  const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
  assert.match(html, /id="root"/);
  assert.match(html, /src="\/src\/main\.tsx"/);
  const main = fs.readFileSync(path.join(root, 'src/main.tsx'), 'utf8');
  assert.match(main, /createRoot/);
  assert.match(main, /from '\.\/App'/);
  const config = fs.readFileSync(path.join(root, 'vite.config.ts'), 'utf8');
  assert.match(config, /outDir:\s*'renderer'/);
  assert.match(config, /base:\s*'\.\/'/);
  const host = fs.readFileSync(path.join(root, 'electron/main.cjs'), 'utf8');
  assert.match(host, /path\.resolve\(__dirname, '\.\.\/renderer'\)/);
  const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
  assert.equal(pkg.main, 'electron/main.cjs');
  assert.equal(pkg.scripts.build, 'vite build');
  assert.equal(pkg.scripts.start, 'electron .');
  assert.equal(pkg.devDependencies.vite, '8.2.2');
  assert.equal(pkg.devDependencies['@vitejs/plugin-react'], '6.1.0');
  const policy = fs.readFileSync(path.join(root, 'src/spreadsheetSheetPolicy.ts'), 'utf8');
  assert.doesNotMatch(policy, /import \{[^}]*editableDefinedName/);
});

test('main.tsx mounts App on #root', async () => {
  const { rolldown } = await import('rolldown');
  const bundle = await rolldown({
    input: path.join(root, 'src/main.tsx'),
    platform: 'node',
    external: id => /^react(?:-dom)?(?:\/|$)/.test(id),
    transform: { jsx: { runtime: 'automatic' } },
    plugins: [{
      name: 'mocks',
      resolveId(id) {
        if (id === './App') return '\0app';
        if (id.endsWith('.css')) return '\0asset';
      },
      load(id) {
        if (id === '\0app') return 'export default function App() { return null }';
        if (id === '\0asset') return 'export default ""';
      },
    }],
  });
  try {
    const { output } = await bundle.generate({ format: 'cjs', codeSplitting: false });
    const mounted = [];
    const reactDom = {
      createRoot(el) {
        assert.equal(el && el.id, 'root');
        return { render(node) { mounted.push(node); } };
      },
    };
    const document = { getElementById(id) { return { id }; } };
    const req = id => {
      if (id === 'react') return React;
      if (id === 'react/jsx-runtime') return require('react/jsx-runtime');
      if (id === 'react-dom/client') return reactDom;
      return require(id);
    };
    const mod = { exports: {} };
    new Function('require', 'module', 'exports', 'document', output[0].code)(req, mod, mod.exports, document);
    assert.equal(mounted.length, 1);
  } finally {
    await bundle.close();
  }
});
