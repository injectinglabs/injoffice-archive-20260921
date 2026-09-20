// Locate electron-builder's unpacked applications under apps/desktop/release and
// read their app.asar without extracting it. Layouts per platform:
//   mac-arm64/InjOffice.app/Contents/Resources/app.asar   (--mac --arm64)
//   mac/InjOffice.app/Contents/Resources/app.asar         (--mac --x64)
//   win-unpacked/resources/app.asar                       (--win)
//   linux-unpacked/resources/app.asar                     (--linux)
const fs = require('node:fs');
const path = require('node:path');

const defaultReleaseDir = path.resolve(__dirname, '../release');

function findPackagedApps(releaseDir = defaultReleaseDir) {
  if (!fs.existsSync(releaseDir)) return [];
  const apps = [];
  for (const entry of fs.readdirSync(releaseDir, { withFileTypes: true })) {
    if (!entry.isDirectory()) continue;
    const directory = path.join(releaseDir, entry.name);
    const mac = /^mac(?:-(arm64|x64|universal))?$/.exec(entry.name);
    if (mac) {
      for (const bundle of fs.readdirSync(directory).filter(name => name.endsWith('.app'))) {
        const resources = path.join(directory, bundle, 'Contents', 'Resources');
        apps.push({ platform: 'mac', arch: mac[1] ?? 'x64', name: `${entry.name}/${bundle}`, resources, asar: path.join(resources, 'app.asar') });
      }
      continue;
    }
    const unpacked = /^(win|linux)(?:-([a-z0-9]+))?-unpacked$/.exec(entry.name);
    if (unpacked) {
      const resources = path.join(directory, 'resources');
      apps.push({ platform: unpacked[1], arch: unpacked[2] ?? 'x64', name: entry.name, resources, asar: path.join(resources, 'app.asar') });
    }
  }
  return apps.sort((a, b) => a.name.localeCompare(b.name));
}

// asar: [u32 4][u32 header pickle size][u32 payload size][u32 json length] json … then file data.
function readAsar(file) {
  const fd = fs.openSync(file, 'r');
  const head = Buffer.alloc(16);
  fs.readSync(fd, head, 0, 16, 0);
  if (head.readUInt32LE(0) !== 4) throw new Error(`${file} is not an asar archive`);
  const headerSize = head.readUInt32LE(4);
  const jsonLength = head.readUInt32LE(12);
  const json = Buffer.alloc(jsonLength);
  fs.readSync(fd, json, 0, jsonLength, 16);
  const header = JSON.parse(json.toString('utf8'));
  const dataStart = 8 + headerSize;
  const entries = new Map();
  (function walk(node, prefix) {
    for (const [name, child] of Object.entries(node.files ?? {})) {
      const entryPath = prefix + name;
      if (child.files) walk(child, `${entryPath}/`);
      else entries.set(entryPath, child);
    }
  })(header, '');
  return {
    file,
    list: () => [...entries.keys()],
    has: entryPath => entries.has(entryPath),
    size: entryPath => entries.get(entryPath)?.size,
    read(entryPath) {
      const entry = entries.get(entryPath);
      if (!entry) throw new Error(`${entryPath} is not in ${file}`);
      if (entry.unpacked) return fs.readFileSync(path.join(`${file}.unpacked`, entryPath));
      const bytes = Buffer.alloc(entry.size);
      fs.readSync(fd, bytes, 0, entry.size, dataStart + Number(entry.offset));
      return bytes;
    },
    close: () => fs.closeSync(fd),
  };
}

module.exports = { defaultReleaseDir, findPackagedApps, readAsar };
