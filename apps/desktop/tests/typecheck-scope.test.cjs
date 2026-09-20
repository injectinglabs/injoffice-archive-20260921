// `npm run typecheck` covers every renderer module except the ones tsconfig.json excludes because
// they import package exports that do not exist on main yet. This test keeps that list honest: an
// excluded file must still fail (itself or through a file it imports), so the list can only shrink.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

const root = path.resolve(__dirname, '..');
const stripComments = text => text.replace(/^\s*\/\/.*$/gm, '');
const config = JSON.parse(stripComments(fs.readFileSync(path.join(root, 'tsconfig.json'), 'utf8')));
const tsc = path.join(root, '../../node_modules/typescript/bin/tsc');

function localImports(file) {
  const source = fs.readFileSync(path.join(root, file), 'utf8');
  const imports = [];
  for (const match of source.matchAll(/(?:from\s+|import\()\s*['"](\.\/[^'"]+)['"]/g)) {
    const base = path.posix.join('src', match[1].slice(2).replace(/\.(?:tsx?|js)$/, ''));
    for (const extension of ['.ts', '.tsx']) if (fs.existsSync(path.join(root, base + extension))) imports.push(base + extension);
  }
  return imports;
}

test('typecheck runs tsc over src with the documented exclusions only', () => {
  const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
  assert.equal(pkg.scripts.typecheck, 'tsc --noEmit -p tsconfig.json');
  assert.deepEqual(config.include, ['src']);
  assert.equal(config.compilerOptions.strict, true);
  for (const file of config.exclude) {
    assert.match(file, /^src\/[^/]+\.tsx?$/);
    assert.ok(fs.existsSync(path.join(root, file)), `${file} is excluded but does not exist; drop it from tsconfig.json`);
  }
});

test('every excluded renderer module still fails the full typecheck, directly or through its imports', () => {
  const result = spawnSync(process.execPath, [tsc, '--noEmit', '-p', 'tsconfig.full.json', '--pretty', 'false'], { cwd: root, encoding: 'utf8' });
  assert.notEqual(result.status, 0, 'the full renderer typechecks: remove tsconfig.json "exclude" and this test');
  const failing = new Set([...result.stdout.matchAll(/^(src\/[^\s(]+)\(\d+,\d+\): error TS/gm)].map(match => match[1]));
  assert.ok(failing.size > 0, result.stdout.slice(0, 2000));
  const excluded = new Set(config.exclude);
  for (const file of failing) assert.ok(excluded.has(file), `${file} fails the full typecheck but is not excluded: npm run typecheck is red`);
  const tainted = new Set(failing);
  for (let changed = true; changed;) {
    changed = false;
    for (const file of excluded) {
      if (!tainted.has(file) && localImports(file).some(dependency => tainted.has(dependency))) { tainted.add(file); changed = true; }
    }
  }
  for (const file of excluded) assert.ok(tainted.has(file), `${file} typechecks now; remove it from tsconfig.json "exclude"`);
  console.log(`full renderer typecheck: ${failing.size} files fail, ${excluded.size - failing.size} more excluded for importing them`);
});
