const fs = require('node:fs');
const path = require('node:path');

const pdfRoot = path.dirname(require.resolve('pdfjs-dist/package.json'));
const PDF_FONT = 'pdf-assets/standard_fonts/LiberationSans-Regular.ttf';

// Keep the runtime URLs used by PdfEditor available in both Vite and app.asar.
// Include the upstream licenses alongside the fonts, CMaps and decoder WASM.
function pdfResourceFiles() {
  const files = new Map();
  for (const folder of ['standard_fonts', 'cmaps', 'wasm']) {
    for (const entry of fs.readdirSync(path.join(pdfRoot, folder), { withFileTypes: true })) {
      if (entry.isFile()) files.set(`pdf-assets/${folder}/${entry.name}`, fs.readFileSync(path.join(pdfRoot, folder, entry.name)));
    }
  }
  if (!files.get(PDF_FONT)?.length) throw new Error(`Missing bundled PDF font: ${PDF_FONT}`);
  return files;
}

function pdfAssets() {
  return {
    name: 'injoffice-desktop-pdf-assets',
    generateBundle() {
      for (const [fileName, source] of pdfResourceFiles()) this.emitFile({ type: 'asset', fileName, source });
    },
    configureServer(server) {
      const files = pdfResourceFiles();
      server.middlewares.use((req, res, next) => {
        if (req.method !== 'GET' && req.method !== 'HEAD') return next();
        const source = files.get((req.url ?? '').split('?')[0].replace(/^\//, ''));
        if (!source) return next();
        res.setHeader('Content-Type', req.url.split('?')[0].endsWith('.wasm') ? 'application/wasm' : 'application/octet-stream');
        res.setHeader('Content-Length', source.length);
        res.end(req.method === 'HEAD' ? undefined : source);
      });
    },
  };
}

module.exports = { PDF_FONT, pdfResourceFiles, pdfAssets };
