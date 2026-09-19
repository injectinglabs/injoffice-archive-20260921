const path = require('node:path');
const { fileURLToPath } = require('node:url');

function isInside(root, filename) {
  const relative = path.relative(root, filename);
  return relative !== '..' && !relative.startsWith(`..${path.sep}`) && !path.isAbsolute(relative);
}

function isAppURL(url) {
  return url.protocol === 'injoffice:' && url.hostname === 'app' && !url.port && !url.username && !url.password;
}

function resolveAsset(root, address) {
  const url = new URL(address);
  if (!isAppURL(url)) throw new Error('Forbidden origin.');
  const pathname = decodeURIComponent(url.pathname);
  if (pathname.includes('\\') || pathname.includes('\0')) throw new Error('Invalid asset path.');
  const filename = path.resolve(root, `.${pathname}`);
  if (!isInside(root, filename)) throw new Error('Asset path escapes the renderer.');
  return filename;
}

function allowRequest(root, address) {
  try {
    const url = new URL(address);
    if (url.protocol === 'file:') return isInside(root, fileURLToPath(url));
    if (url.protocol === 'injoffice:') { resolveAsset(root, address); return true; }
    return ['data:', 'blob:'].includes(url.protocol);
  } catch { return false; }
}

function isTrustedSender(event, window, entryURL) {
  return Boolean(window && !window.isDestroyed() && event.sender === window.webContents && event.senderFrame === window.webContents.mainFrame && event.senderFrame?.url === entryURL);
}

module.exports = { resolveAsset, allowRequest, isTrustedSender };
