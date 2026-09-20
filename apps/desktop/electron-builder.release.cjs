// Public releases opt in to signed publishing configuration. Local builds stay unsigned.
const {build} = require('./package.json');
module.exports = {
  ...build,
  extraMetadata: {injofficeRelease: true},
  artifactName: '${productName}-${version}-${os}-${arch}.${ext}',
  publish: {
    provider: 'github', owner: 'injectinglabs', repo: 'injoffice',
    releaseType: 'draft', tagNamePrefix: 'desktop-v',
  },
  mac: {...build.mac, forceCodeSigning: true, hardenedRuntime: true, notarize: true},
  win: {...build.win, forceCodeSigning: true},
};
