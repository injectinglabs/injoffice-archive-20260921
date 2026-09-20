const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const root = path.resolve(__dirname, '..');

test('electron-builder packages the Vite renderer and Electron host, not dist/', () => {
  const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
  assert.equal(pkg.main, 'electron/main.cjs');
  assert.equal(pkg.scripts.build, 'vite build');
  assert.equal(pkg.scripts.start, 'electron .');
  assert.equal(pkg.scripts.dist, 'npm run build && electron-builder');
  assert.equal(pkg.devDependencies['electron-builder'], '^26.0.12');
  assert.equal(pkg.build.appId, 'com.injecting.injoffice');
  assert.equal(pkg.build.productName, 'InjOffice');
  assert.equal(pkg.build.directories.output, 'release');
  assert.equal(pkg.build.electronVersion, '43.7.0');
  assert.equal(pkg.build.publish, null);
  assert.deepEqual(pkg.build.files, ['electron/**/*.cjs', 'electron/icon.png', 'renderer/**/*', 'package.json']);
  assert.ok(!pkg.build.files.some(pattern => pattern.startsWith('dist/')));
  assert.equal(pkg.build.mac.icon, 'build/icons/icon.icns');
  assert.equal(pkg.build.win.icon, 'build/icons/icon.ico');
  assert.equal(pkg.build.linux.icon, 'build/icons/icon.png');
  // electron-builder derives the executable from the package name (@injoffice/desktop -> "@injofficedesktop"), which AppImage refuses.
  assert.equal(pkg.build.linux.executableName, 'injoffice');
  // deb/rpm default to ${name}_${version}_${arch}, and the scoped name puts a slash in the file name; fpm cannot write it.
  assert.equal(pkg.build.linux.artifactName, '${productName}-${version}-linux-${arch}.${ext}');
  const extensions = pkg.build.fileAssociations.map(item => item.ext).sort();
  assert.deepEqual(extensions, ['docx', 'pdf', 'pptx', 'xlsx']);

  const release = fs.readFileSync(path.join(root, 'electron-builder.release.cjs'), 'utf8');
  assert.match(release, /require\('\.\/package\.json'\)/);
  assert.match(release, /tagNamePrefix:\s*'desktop-v'/);
  assert.match(release, /releaseType:\s*'draft'/);
  assert.match(release, /forceCodeSigning:\s*true/);
});
